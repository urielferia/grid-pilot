using GridPilot.Core;
namespace GridPilot.Tests;

// Test double: only domain/session tests use it. Excel integration uses real COM.
public sealed class MemoryWorkbook(WorkbookProfile profile) : IWorkbook
{
    WorkbookSnapshot current=Fixture(profile);
    public bool FailWrite {get;set;} public bool FailRestore {get;set;}
    public void Open(string path) { }
    public WorkbookSnapshot Snapshot()=>Json.Read<WorkbookSnapshot>(Json.Write(current));
    public void Recalculate() { }
    public void Apply(Operation op,WorkbookProfile p)
    {
        var t=current.Table(op.Table);var fp=p.Table(op.Table).Fields.Single(f=>f.Key==op.Field);var row=t.Rows.Single(r=>r[p.Table(op.Table).Fields.Single(f=>f.Key=="Id").Header]==op.Id);row[fp.Header]=op.Value;
        if(FailWrite)throw new InvalidOperationException("Fallo inyectado después de escribir.");
    }
    public void SaveCopy(string path)=>File.WriteAllText(path,Json.Write(current));
    public void Restore(string path,string draft){if(FailRestore)throw new IOException("Restauración interrumpida.");current=Json.Read<WorkbookSnapshot>(File.ReadAllText(path));}
    public IDisposable Exclusive()=>new Noop();
    public void Dispose() { }
    sealed class Noop:IDisposable{public void Dispose() { }}
    public static WorkbookSnapshot Fixture(WorkbookProfile profile)
    {
        var values=new Dictionary<string,Dictionary<string,string>> {
            ["Inventario"]=new(){["Id"]="M001",["Descripcion"]="Demo",["Unidad"]="pieza",["SaldoInicial"]="20",["Minimo"]="10",["CostoReferencia"]="5",["Entradas"]="0",["Consumos"]="2",["Existencia"]="18"},
            ["Visitas"]=new(){["Id"]="V001",["Fecha"]="46296",["Cliente"]="Demo",["Sitio"]="Demo",["Tecnico"]="Ana",["Estado"]="Realizada",["Descripcion"]="Demo",["CostoMateriales"]="10"},
            ["Movimientos"]=new(){["Id"]="MOV001",["Fecha"]="46296",["Tipo"]="Consumo",["Material"]="M001",["Cantidad"]="2",["Visita"]="V001",["CostoUnitario"]="5",["Importe"]="10"}
        };
        var tables=new List<TableState>();var cells=new List<CellState>();
        foreach(var t in profile.Tables){var data=t.Fields.ToDictionary(f=>f.Header,f=>values[t.Name][f.Key]);tables.Add(new(t.Name,t.Sheet,$"A5:{(char)('A'+t.Fields.Count-1)}6",t.Fields.Select(f=>f.Header).ToList(),[data]));for(int i=0;i<t.Fields.Count;i++){var f=t.Fields[i];cells.Add(new(t.Sheet,$"{(char)('A'+i)}6",data[f.Header],f.Calculated?"=1":"","General"));}}
        foreach(var (address,value) in new[]{("B5","46296"),("B6","46326"),("B9","1"),("B10","10"),("B11","0")})cells.Add(new("Resumen",address,value,address is "B5" or "B6"?"":"=1","General"));
        return new(["Resumen","Visitas","Inventario","Movimientos"],tables,cells,[]);
    }
}
