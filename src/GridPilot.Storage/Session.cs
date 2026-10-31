using GridPilot.Core;
namespace GridPilot.Storage;

public sealed record SessionState(string Original, string OriginalHash, string Draft, string ValidCopy, WorkbookProfile Profile, string Revision, bool Blocked = false);
public sealed class Session : IDisposable
{
    readonly IWorkbook excel; readonly Journal journal; readonly string directory; readonly object gate = new();
    SessionState state; ApprovalRecord? approval; ChangeSet? last; bool disposed;
    public bool Editable { get; private set; }
    public string DirectoryPath => directory;
    public WorkbookProfile Profile => state.Profile;
    public Session(IWorkbook workbook, string root, string original, WorkbookProfile profile)
    {
        profile.ValidateDefinition();
        excel = workbook; directory = Path.Combine(root, Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
        journal = new(Path.Combine(directory, "session.db"));
        try
        {
            original = Path.GetFullPath(original); var baseline = Path.Combine(directory, "base.xlsx"); var draft = Path.Combine(directory, "draft.xlsx");
            File.Copy(original, baseline, false); File.SetAttributes(baseline, FileAttributes.ReadOnly); File.Copy(baseline, draft, false); File.SetAttributes(draft, FileAttributes.Normal);
            state = new(original, Json.FileHash(baseline), draft, baseline, profile.Copy(), "");
            excel.Open(draft); excel.Recalculate(); var snapshot = excel.Snapshot(); state = state with { Revision = snapshot.Revision };
            try { Rules.Validate(snapshot, state.Profile); Editable = true; } catch (InvalidOperationException) { Editable = false; }
            Persist();
        }
        catch { journal.Dispose(); throw; }
    }
    public Session(IWorkbook workbook, string sessionDirectory)
    {
        excel = workbook; directory = Path.GetFullPath(sessionDirectory);
        if (!File.Exists(Path.Combine(directory, "session.db"))) throw new InvalidOperationException("Session not found.");
        journal = new(Path.Combine(directory, "session.db"));
        try
        {
            state = journal.Get<SessionState>("session") ?? throw new InvalidOperationException("Session not found.");
            state.Profile.ValidateDefinition();
            ValidateLocalPath(state.Draft); ValidateLocalPath(state.ValidCopy);
            // Only explicit recovery opens the last committed checkpoint. Never replay pending writes.
            File.Copy(state.ValidCopy, state.Draft, true); File.SetAttributes(state.Draft, FileAttributes.Normal); excel.Open(state.Draft); excel.Recalculate();
            Rules.Validate(excel.Snapshot(), state.Profile); state = state with { Blocked = false, Revision = excel.Snapshot().Revision }; Editable = true; journal.RecoverPending(); Persist();
        }
        catch { journal.Dispose(); throw; }
    }
    void Persist() => journal.Put("session", state);
    void ValidateLocalPath(string path) { var full = Path.GetFullPath(path); if (!full.StartsWith(directory + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) || !full.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || full.Equals(state.Original, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Recovery path is outside the session."); }
    public WorkbookSnapshot Inspect() { lock (gate) { return excel.Snapshot(); } }
    public WorkbookSnapshot RecalculateValidate() { lock (gate) { Ready(); using var lease = excel.Exclusive(); excel.Recalculate(); var snapshot = excel.Snapshot(); Rules.Validate(snapshot, state.Profile); return snapshot; } }
    void Ready() { if (!Editable) throw new InvalidOperationException("Workbook is open for inspection only: a validated compatible profile is required."); if (state.Blocked) throw new InvalidOperationException("Session is blocked: recover the last valid state."); }
    public ChangeSet Execute(ActionPlan plan)
    {
        lock (gate)
        {
            Ready(); if (plan.Operations.Count is < 1 or > 100 || !Guid.TryParse(plan.OperationId, out _)) throw new InvalidOperationException("Invalid batch or operation identifier.");
            var old = journal.Find(plan.OperationId);
            if (old is not null) { if (old.Value.Plan != Json.Hash(Json.Write(plan))) throw new InvalidOperationException("Operation identifier already belongs to a different batch."); if (old.Value.Status == "completed") return old.Value.Result!; throw new InvalidOperationException("Batch was already attempted: inspect the draft and create a new instruction."); }
            using var lease = excel.Exclusive(); var before = excel.Snapshot(); if (before.Revision != plan.SourceRevision) throw new InvalidOperationException("The draft has changed. Review and re-interpret the instruction.");
            Rules.Validate(before, state.Profile);
            var checkpoint = Path.Combine(directory, $"before-{plan.OperationId}.xlsx"); excel.SaveCopy(checkpoint);
            var priorProfile = state.Profile.Copy(); var priorState = state with { Profile = priorProfile }; journal.Begin(plan); approval = null;
            try
            {
                foreach (var operation in plan.Operations) excel.Apply(operation, state.Profile);
                excel.Recalculate(); var after = excel.Snapshot(); Rules.Validate(after, state.Profile);
                var committed = Path.Combine(directory, $"valid-{plan.OperationId}.xlsx"); excel.SaveCopy(committed);
                var result = Diff.Compare(plan, before, after, checkpoint);
                state = state with { Revision = after.Revision, ValidCopy = committed };
                journal.Atomic(() => { journal.Put("undoProfile", priorProfile); Persist(); journal.Finish(plan.OperationId, "completed", result); journal.Put("last", result); }); last = result; return result;
            }
            catch
            {
                try { excel.Restore(checkpoint, state.Draft); excel.Recalculate(); Rules.Validate(excel.Snapshot(), priorProfile); state = priorState with { Revision = excel.Snapshot().Revision }; journal.Atomic(() => { journal.Finish(plan.OperationId, "failed"); Persist(); }); }
                catch { state = priorState with { Blocked = true }; Persist(); throw new InvalidOperationException("Rollback failed. Session is blocked; close and recover the last valid state."); }
                throw;
            }
        }
    }
    public WorkbookSnapshot Undo(string revision)
    {
        lock (gate)
        {
            Ready(); using var lease = excel.Exclusive(); if (excel.Snapshot().Revision != revision) throw new InvalidOperationException("The draft has changed."); var change = last ?? journal.Get<ChangeSet>("last") ?? throw new InvalidOperationException("No batch available to undo.");
            try { ValidateLocalPath(change.Checkpoint); excel.Restore(change.Checkpoint, state.Draft); excel.Recalculate(); var profile = journal.Get<WorkbookProfile>("undoProfile")!; var snapshot = excel.Snapshot(); Rules.Validate(snapshot, profile); state = state with { Profile = profile, ValidCopy = change.Checkpoint, Revision = snapshot.Revision }; journal.Atomic(() => { Persist(); journal.Put<ChangeSet?>("last", null); }); last = null; approval = null; return snapshot; }
            catch { state = state with { Blocked = true }; Persist(); throw; }
        }
    }
    // Invoked exclusively by the UI controller, never registered as an MCP tool.
    public ApprovalRecord Approve(string reviewedRevision)
    {
        lock (gate)
        {
            Ready(); using var lease = excel.Exclusive(); excel.Recalculate(); var snap = excel.Snapshot(); if (snap.Revision != reviewedRevision) throw new InvalidOperationException("Changes occurred after the review. Inspect before approving."); Rules.Validate(snap, state.Profile);
            var path = Path.Combine(directory, $"approved-{Guid.NewGuid():N}.xlsx"); excel.SaveCopy(path); File.SetAttributes(path, FileAttributes.ReadOnly); approval = new(snap.Revision, DateTimeOffset.UtcNow, Json.FileHash(path), path); journal.Put("approval", approval); return approval;
        }
    }
    public ApprovalRecord Export(string destination)
    {
        lock (gate)
        {
            Ready(); using var lease = excel.Exclusive(); var record = approval ?? throw new InvalidOperationException("Explicit approval is required.");
            if (excel.Snapshot().Revision != record.Revision) { approval = null; throw new InvalidOperationException("Approval was invalidated by subsequent changes."); }
            destination = Path.GetFullPath(destination);
            if (!destination.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || destination.Equals(state.Original, StringComparison.OrdinalIgnoreCase) || destination.StartsWith(directory + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) || File.Exists(destination))
                throw new InvalidOperationException("Choose a new .xlsx file outside the session directory.");
            if (Json.FileHash(record.ImmutablePath) != record.Hash) throw new InvalidOperationException("The approved copy was altered.");
            using (var input = File.OpenRead(record.ImmutablePath)) using (var output = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None)) input.CopyTo(output);
            journal.Put("export", new { destination, record }); return record;
        }
    }
    public void Dispose() { if (disposed) return; disposed = true; try { excel.Dispose(); } finally { journal.Dispose(); } }
}
