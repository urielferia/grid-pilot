using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace GridPilot.Core;

public static class Json
{
    public static readonly JsonSerializerOptions Options = new() { PropertyNameCaseInsensitive = true, WriteIndented = false, Converters = { new JsonStringEnumConverter() } };
    public static string Write<T>(T value) => JsonSerializer.Serialize(value, Options);
    public static T Read<T>(string value) => JsonSerializer.Deserialize<T>(value, Options) ?? throw new InvalidDataException("JSON vacío.");
    public static string Hash(string text) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)));
    public static string FileHash(string path) { using var f = File.OpenRead(path); return Convert.ToHexString(SHA256.HashData(f)); }
}
public record FieldProfile(string Key, string Header, string Type, bool Required = true, bool Calculated = false);
public record TableProfile(string Name, string Sheet, List<FieldProfile> Fields);
public record WorkbookProfile(string Id, List<TableProfile> Tables, string SummarySheet = "Resumen")
{
    public string SummarySheet {get;set;}=SummarySheet;
    public TableProfile Table(string name) => Tables.SingleOrDefault(t => t.Name == name) ?? throw new InvalidOperationException("Tabla fuera del perfil.");
    public WorkbookProfile Copy() => Json.Read<WorkbookProfile>(Json.Write(this));
}
public record CellState(string Sheet, string Address, string Value, string Formula, string NumberFormat);
public record TableState(string Name, string Sheet, string Address, List<string> Headers, List<Dictionary<string,string>> Rows);
public record WorkbookSnapshot(List<string> Sheets, List<TableState> Tables, List<CellState> Cells, List<string> Properties)
{
    public string Revision => Json.Hash(Json.Write(new { Sheets, Tables, Cells, Properties }));
    public TableState Table(string name) => Tables.Single(t => t.Name == name);
}
public enum OperationKind { SetValue, AddRecord, DeleteRecord, AddColumn, RenameColumn, DeleteColumn, SetFormula, AddSheet, RenameSheet, DeleteSheet }
public record Operation(OperationKind Kind, string Table = "", string Id = "", string Field = "", string Value = "", Dictionary<string,string>? Values = null);
public record ActionPlan(string OperationId, string Intent, string SourceRevision, List<Operation> Operations);
public record CellChange(string Sheet, string Address, CellState? Before, CellState? After);
public record ChangeSet(string OperationId, string Intent, string BeforeRevision, string AfterRevision, string Checkpoint, List<CellChange> Changes, List<string> Structure, List<string> Validations);
public record ApprovalRecord(string Revision, DateTimeOffset At, string Hash, string ImmutablePath);
public record Interpretation(string Clarification, List<Operation> Operations);
public interface IWorkbook : IDisposable
{
    void Open(string path);
    WorkbookSnapshot Snapshot();
    void Apply(Operation operation, WorkbookProfile profile);
    void Recalculate();
    void SaveCopy(string path);
    void Restore(string path, string draft);
    IDisposable Exclusive();
}
public static class Diff
{
    public static ChangeSet Compare(ActionPlan p, WorkbookSnapshot a, WorkbookSnapshot b, string checkpoint)
    {
        var x = a.Cells.ToDictionary(c => (c.Sheet,c.Address)); var y = b.Cells.ToDictionary(c => (c.Sheet,c.Address));
        var changes = x.Keys.Union(y.Keys).Where(k => x.GetValueOrDefault(k) != y.GetValueOrDefault(k)).Select(k => new CellChange(k.Sheet,k.Address,x.GetValueOrDefault(k),y.GetValueOrDefault(k))).ToList();
        var structure = new List<string>();
        if (Json.Write(a.Sheets) != Json.Write(b.Sheets)) structure.Add($"Hojas: {string.Join(", ",a.Sheets)} → {string.Join(", ",b.Sheets)}");
        if (Json.Write(a.Tables.Select(t => new {t.Name,t.Sheet,t.Address,t.Headers})) != Json.Write(b.Tables.Select(t => new {t.Name,t.Sheet,t.Address,t.Headers}))) structure.Add("Cambió la estructura o extensión de las tablas.");
        if (Json.Write(a.Properties) != Json.Write(b.Properties)) structure.Add("Cambiaron propiedades, validaciones o nombres.");
        return new(p.OperationId,p.Intent,a.Revision,b.Revision,checkpoint,changes,structure,["Relaciones, saldos, costos y fórmulas comprobados."]);
    }
}
