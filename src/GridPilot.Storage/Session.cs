using GridPilot.Core;
namespace GridPilot.Storage;

public sealed record SessionState(string Original,string OriginalHash,string Draft,string ValidCopy,WorkbookProfile Profile,string Revision,bool Blocked=false);
public sealed class Session : IDisposable
{
    readonly IWorkbook excel; readonly Journal journal; readonly string directory; readonly object gate=new();
    SessionState state; ApprovalRecord? approval; ChangeSet? last; bool disposed;
    public bool Editable {get;private set;}
    public string DirectoryPath=>directory;
    public WorkbookProfile Profile=>state.Profile;
    public Session(IWorkbook workbook,string root,string original,WorkbookProfile profile)
    {
        excel=workbook;directory=Path.Combine(root,Guid.NewGuid().ToString("N"));Directory.CreateDirectory(directory);
        journal=new(Path.Combine(directory,"session.db"));
        original=Path.GetFullPath(original);var baseline=Path.Combine(directory,"base.xlsx");var draft=Path.Combine(directory,"draft.xlsx");
        File.Copy(original,baseline,false);File.SetAttributes(baseline,FileAttributes.ReadOnly);File.Copy(baseline,draft,false);File.SetAttributes(draft,FileAttributes.Normal);
        state=new(original,Json.FileHash(baseline),draft,baseline,profile.Copy(),"");
        excel.Open(draft);excel.Recalculate();var snapshot=excel.Snapshot();state=state with{Revision=snapshot.Revision};
        try{Rules.Validate(snapshot,state.Profile);Editable=true;}catch(InvalidOperationException){Editable=false;}
        Persist();
    }
    public Session(IWorkbook workbook,string sessionDirectory)
    {
        excel=workbook;directory=Path.GetFullPath(sessionDirectory);journal=new(Path.Combine(directory,"session.db"));
        state=journal.Get<SessionState>("session") ?? throw new InvalidOperationException("Sesión no encontrada.");
        // Only explicit recovery opens the last committed checkpoint. Never replay pending writes.
        File.Copy(state.ValidCopy,state.Draft,true);File.SetAttributes(state.Draft,FileAttributes.Normal);excel.Open(state.Draft);excel.Recalculate();
        Rules.Validate(excel.Snapshot(),state.Profile);state=state with{Blocked=false,Revision=excel.Snapshot().Revision};Editable=true;journal.RecoverPending();Persist();
    }
    void Persist()=>journal.Put("session",state);
    public WorkbookSnapshot Inspect(){lock(gate){return excel.Snapshot();}}
    public WorkbookSnapshot RecalculateValidate(){lock(gate){Ready();using var lease=excel.Exclusive();excel.Recalculate();var snapshot=excel.Snapshot();Rules.Validate(snapshot,state.Profile);return snapshot;}}
    void Ready(){if(!Editable)throw new InvalidOperationException("Libro solo para inspección: necesita un perfil compatible validado.");if(state.Blocked)throw new InvalidOperationException("Sesión bloqueada: recupere el último estado válido.");}
    public ChangeSet Execute(ActionPlan plan)
    {
        lock(gate)
        {
            Ready(); if(plan.Operations.Count is <1 or >100 || !Guid.TryParse(plan.OperationId,out _))throw new InvalidOperationException("Lote o identificador de operación inválido.");
            var old=journal.Find(plan.OperationId);
            if(old is not null){if(old.Value.Plan!=Json.Hash(Json.Write(plan)))throw new InvalidOperationException("El identificador ya pertenece a otro lote.");if(old.Value.Status=="completed")return old.Value.Result!;throw new InvalidOperationException("Lote ya intentado: inspeccione y cree una nueva instrucción.");}
            using var lease=excel.Exclusive();var before=excel.Snapshot();if(before.Revision!=plan.SourceRevision)throw new InvalidOperationException("El borrador cambió. Revise y vuelva a interpretar la instrucción.");
            Rules.Validate(before,state.Profile);
            var checkpoint=Path.Combine(directory,$"before-{plan.OperationId}.xlsx");excel.SaveCopy(checkpoint);
            var priorProfile=state.Profile.Copy();journal.Begin(plan);approval=null;
            try
            {
                foreach(var operation in plan.Operations)excel.Apply(operation,state.Profile);
                excel.Recalculate();var after=excel.Snapshot();Rules.Validate(after,state.Profile);
                var committed=Path.Combine(directory,$"valid-{plan.OperationId}.xlsx");excel.SaveCopy(committed);
                var result=Diff.Compare(plan,before,after,checkpoint);journal.Put("undoProfile",priorProfile);
                state=state with{Revision=after.Revision,ValidCopy=committed};Persist();journal.Finish(plan.OperationId,"completed",result);last=result;journal.Put("last",result);return result;
            }
            catch
            {
                try{excel.Restore(checkpoint,state.Draft);excel.Recalculate();Rules.Validate(excel.Snapshot(),priorProfile);state=state with{Profile=priorProfile,Revision=excel.Snapshot().Revision};journal.Finish(plan.OperationId,"failed");Persist();}
                catch{state=state with{Blocked=true,Profile=priorProfile};Persist();throw new InvalidOperationException("Falló la recuperación. Sesión bloqueada; cierre y recupere el último estado válido.");}
                throw;
            }
        }
    }
    public WorkbookSnapshot Undo(string revision)
    {
        lock(gate){Ready();using var lease=excel.Exclusive();if(excel.Snapshot().Revision!=revision)throw new InvalidOperationException("El borrador cambió.");var change=last ?? journal.Get<ChangeSet>("last") ?? throw new InvalidOperationException("No hay lote para deshacer.");
            try{excel.Restore(change.Checkpoint,state.Draft);excel.Recalculate();var profile=journal.Get<WorkbookProfile>("undoProfile")!;var snapshot=excel.Snapshot();Rules.Validate(snapshot,profile);state=state with{Profile=profile,ValidCopy=change.Checkpoint,Revision=snapshot.Revision};Persist();last=null;journal.Put<ChangeSet?>("last",null);approval=null;return snapshot;}
            catch{state=state with{Blocked=true};Persist();throw;}}
    }
    // Invoked exclusively by the UI controller, never registered as an MCP tool.
    public ApprovalRecord Approve(string reviewedRevision)
    {
        lock(gate){Ready();using var lease=excel.Exclusive();excel.Recalculate();var snap=excel.Snapshot();if(snap.Revision!=reviewedRevision)throw new InvalidOperationException("Hay cambios posteriores a la revisión. Inspeccione antes de aprobar.");Rules.Validate(snap,state.Profile);
            var path=Path.Combine(directory,$"approved-{Guid.NewGuid():N}.xlsx");excel.SaveCopy(path);File.SetAttributes(path,FileAttributes.ReadOnly);approval=new(snap.Revision,DateTimeOffset.UtcNow,Json.FileHash(path),path);journal.Put("approval",approval);return approval;}
    }
    public ApprovalRecord Export(string destination)
    {
        lock(gate){Ready();using var lease=excel.Exclusive();var record=approval ?? throw new InvalidOperationException("Se requiere aprobación explícita.");
            if(excel.Snapshot().Revision!=record.Revision){approval=null;throw new InvalidOperationException("La aprobación quedó invalidada por cambios posteriores.");}
            destination=Path.GetFullPath(destination);
            if(!destination.EndsWith(".xlsx",StringComparison.OrdinalIgnoreCase) || destination.Equals(state.Original,StringComparison.OrdinalIgnoreCase) || destination.StartsWith(directory+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase) || File.Exists(destination))throw new InvalidOperationException("Elija un archivo .xlsx nuevo fuera de la sesión.");
            if(Json.FileHash(record.ImmutablePath)!=record.Hash)throw new InvalidOperationException("La copia aprobada fue alterada.");
            using(var input=File.OpenRead(record.ImmutablePath))using(var output=new FileStream(destination,FileMode.CreateNew,FileAccess.Write,FileShare.None))input.CopyTo(output);
            journal.Put("export",new{destination,record});return record;}
    }
    public void Dispose(){if(disposed)return;disposed=true;excel.Dispose();journal.Dispose();}
}
