using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using GridPilot.Core;

namespace GridPilot.Excel;

public sealed class ExcelWorkbook(bool visible = true) : IWorkbook
{
    readonly StaThread sta = new();
    dynamic? app, book;
    bool disposed;
    public static void CheckPackage(string path)
    {
        if(!string.Equals(Path.GetExtension(path),".xlsx",StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Solo se admiten archivos .xlsx.");
        using var zip = ZipFile.OpenRead(path);
        foreach(var e in zip.Entries)
        {
            var n=e.FullName.ToLowerInvariant();
            if(n.Contains("vbaproject") || n.Contains("externallink") || n.Contains("connections") || n.Contains("querytable") || n.Contains("pivot") || n.Contains("activex") || n.Contains("embeddings") || n.Contains("customxml")) throw new InvalidOperationException("El libro contiene capacidades excluidas.");
            if(n.EndsWith(".rels")) { using var reader=new StreamReader(e.Open()); if(reader.ReadToEnd().Contains("TargetMode=\"External\"")) throw new InvalidOperationException("No se permiten relaciones externas."); }
            if(n.StartsWith("xl/worksheets/")&&n.EndsWith(".xml"))
            {
                using var stream=e.Open();var xml=XDocument.Load(stream);
                foreach(var f in xml.Descendants().Where(x=>x.Name.LocalName=="f"))
                {
                    if(f.Value.Contains('|')||f.Value.Contains(".xlsx]",StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("Referencia externa no admitida.");
                    var clean=Regex.Replace(f.Value,"\"[^\"]*\"|'[^']*'!","");
                    foreach(Match m in Regex.Matches(clean,@"([A-Za-z_][A-Za-z0-9_.]*)\s*\("))
                        if(!new[]{"SUM","SUMIFS","COUNTIFS","IF","ROUND","INDEX","MATCH","SUMPRODUCT"}.Contains(m.Groups[1].Value.ToUpperInvariant()))throw new InvalidOperationException("El libro usa una función no compatible: "+m.Groups[1].Value);
                }
            }
        }
    }
    public void Open(string path)
    {
        CheckPackage(path);
        sta.Run(() => {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application") ?? throw new InvalidOperationException("Excel no está instalado."))!;
            app.Visible=visible; app.DisplayAlerts=false; app.AutomationSecurity=3; app.EnableEvents=false; app.AskToUpdateLinks=false;
            using var c=new ComScope(); dynamic books=c.Own(app.Workbooks);
            book=books.Open(Path.GetFullPath(path),UpdateLinks:0,ReadOnly:false,IgnoreReadOnlyRecommended:true,AddToMru:false);
            app.Iteration=false;
        });
    }
    public IDisposable Exclusive()
    {
        sta.Run(() => { if(app is null || book is null) throw new InvalidOperationException("No hay libro abierto."); app.Interactive=false; });
        return new Lease(() => sta.Run(() => { if(app is not null) app.Interactive=true; }));
    }
    sealed class Lease(Action action) : IDisposable { public void Dispose() => action(); }
    static string Text(object? v) => v is null ? "" : Convert.ToString(v,CultureInfo.InvariantCulture) ?? "";
    public WorkbookSnapshot Snapshot() => sta.Run(() => {
        using var scope=new ComScope();dynamic sheets=scope.Own(book!.Worksheets);
        var formulas=new Dictionary<(string Sheet,string Address),string>();
        for(int i=1;i<=sheets.Count;i++){
            dynamic sheet=scope.Own(sheets.Item[i]),range=scope.Own(sheet.UsedRange),rows=scope.Own(range.Rows),cols=scope.Own(range.Columns);
            int countRows=rows.Count,countCols=cols.Count,rowOffset=range.Row,colOffset=range.Column;
            if(countRows>2000||countCols>50)throw new InvalidOperationException("Libro excede límites del piloto.");
            object matrix=range.Formula;
            for(int r=1;r<=countRows;r++)for(int c=1;c<=countCols;c++){
                object? value=matrix is Array array?array.GetValue(r,c):matrix;
                if(value is string f&&f.StartsWith('=')){
                    int col=colOffset+c-1;string address="";while(col>0){col--;address=(char)('A'+col%26)+address;col/=26;}
                    formulas[((string)sheet.Name,address+(rowOffset+r-1))]=f;
                }
            }
        }
        var copy=Path.Combine(Path.GetTempPath(),"GridPilot-snapshot-"+Guid.NewGuid().ToString("N")+".xlsx");
        try{book.SaveCopyAs(copy);return OpenXmlSnapshot.Read(copy,formulas);}finally{if(File.Exists(copy))File.Delete(copy);}
    });
    public void Recalculate() => sta.Run(() => {
        app!.CalculateFullRebuild();
        using var c=new ComScope();dynamic sheets=c.Own(book!.Worksheets);
        for(int i=1;i<=sheets.Count;i++){dynamic sheet=c.Own(sheets.Item[i]);object? circular=sheet.CircularReference;if(circular is not null){c.Own(circular);throw new InvalidOperationException("Fórmula circular detectada.");}}
    });
    public void SaveCopy(string path) => sta.Run(() => { if(File.Exists(path)) throw new IOException("El destino ya existe."); book!.SaveCopyAs(Path.GetFullPath(path)); });
    public void Restore(string path,string draft) => sta.Run(() => {
        book!.Close(false); Marshal.ReleaseComObject(book); book=null;
        File.Copy(path,draft,true);
        using var c=new ComScope(); dynamic books=c.Own(app!.Workbooks);
        book=books.Open(Path.GetFullPath(draft),UpdateLinks:0,ReadOnly:false,IgnoreReadOnlyRecommended:true,AddToMru:false);
    });
    public void Apply(Operation o,WorkbookProfile p) => sta.Run(() => {
        using var c=new ComScope(); dynamic sheets=c.Own(book!.Worksheets);
        if(o.Kind is OperationKind.AddSheet or OperationKind.DeleteSheet or OperationKind.RenameSheet)
        {
            if(o.Kind==OperationKind.AddSheet) { dynamic added=c.Own(sheets.Add()); added.Name=o.Value; return; }
            dynamic sheet=c.Own(sheets.Item[o.Table]);
            if(o.Kind==OperationKind.DeleteSheet) { if(p.SummarySheet==o.Table || p.Tables.Any(t=>t.Sheet==o.Table)) throw new InvalidOperationException("Hoja indispensable."); sheet.Delete(); }
            else { sheet.Name=o.Value; for(int i=0;i<p.Tables.Count;i++) if(p.Tables[i].Sheet==o.Table) p.Tables[i]=p.Tables[i] with{Sheet=o.Value}; if(p.SummarySheet==o.Table)p.SummarySheet=o.Value; }
            return;
        }
        var tp=p.Table(o.Table); dynamic s=c.Own(sheets.Item[tp.Sheet]), tables=c.Own(s.ListObjects), t=c.Own(tables.Item[tp.Name]), columns=c.Own(t.ListColumns), rows=c.Own(t.ListRows);
        var fp=tp.Fields.SingleOrDefault(f=>f.Key==o.Field || f.Header==o.Field);
        if(o.Kind==OperationKind.AddColumn)
        {
            if(string.IsNullOrWhiteSpace(o.Field) || tp.Fields.Any(f=>f.Header==o.Field || f.Key==o.Field)) throw new InvalidOperationException("Nombre de campo inválido o duplicado.");
            dynamic col=c.Own(columns.Add()); col.Name=o.Field;
            if(o.Value!="") { Rules.Formula(o.Value); FillFormula(col,o.Value,c); }
            tp.Fields.Add(new(o.Field,o.Field,o.Value==""?"text":"number",false,o.Value!="")); return;
        }
        if(o.Kind is OperationKind.RenameColumn or OperationKind.DeleteColumn or OperationKind.SetFormula)
        {
            if(fp is null) throw new InvalidOperationException("Campo inexistente.");
            dynamic col=c.Own(columns.Item[fp.Header]);
            if(o.Kind==OperationKind.RenameColumn){col.Name=o.Value;tp.Fields[tp.Fields.IndexOf(fp)]=fp with{Header=o.Value};}
            else if(o.Kind==OperationKind.DeleteColumn){if(fp.Required || Essential.Contains(fp.Key)) throw new InvalidOperationException("Campo indispensable.");col.Delete();tp.Fields.Remove(fp);}
            else { if(Essential.Contains(fp.Key) || !fp.Calculated) throw new InvalidOperationException("Solo se modifican fórmulas de campos calculados adicionales."); Rules.Formula(o.Value); FillFormula(col,o.Value,c); }
            return;
        }
        dynamic row;
        if(o.Kind==OperationKind.AddRecord)
        {
            var values=o.Values ?? throw new InvalidOperationException("Faltan valores.");
            foreach(var field in tp.Fields.Where(f=>f.Required && !f.Calculated)) if(!values.ContainsKey(field.Key)) throw new InvalidOperationException("Falta campo " + field.Key);
            row=c.Own(rows.Add());
            foreach(var pair in values){var field=tp.Fields.SingleOrDefault(f=>f.Key==pair.Key) ?? throw new InvalidOperationException("Campo desconocido."); WriteCell(row,columns,field,pair.Value,c,true);}
            // Explicit fill guarantees formulas even if Excel AutoFillFormulasInLists is disabled.
            if(rows.Count>1) foreach(var field in tp.Fields.Where(f=>f.Calculated)) {
                dynamic col=c.Own(columns.Item[field.Header]), body=c.Own(col.DataBodyRange); body.FillDown();
            }
        }
        else
        {
            dynamic idCol=c.Own(columns.Item[tp.Fields.Single(f=>f.Key=="Id").Header]);
            int found=0; for(int r=1;r<=rows.Count;r++){dynamic rr=c.Own(rows.Item[r]), range=c.Own(rr.Range), cells=c.Own(range.Cells), cell=c.Own(cells.Item[1,idCol.Index]);if(Text(cell.Value2)==o.Id){if(found!=0) throw new InvalidOperationException("Identificador ambiguo.");found=r;}}
            if(found==0) throw new InvalidOperationException("Registro inexistente: " + o.Id);
            row=c.Own(rows.Item[found]);
            if(o.Kind==OperationKind.DeleteRecord) { if(rows.Count<=1) throw new InvalidOperationException("El piloto requiere al menos un registro por tabla para conservar fórmulas."); row.Delete(); }
            else if(o.Kind==OperationKind.SetValue) { if(fp is null) throw new InvalidOperationException("Campo inexistente."); WriteCell(row,columns,fp,o.Value,c,false); }
            else throw new InvalidOperationException("Operación no soportada.");
        }
    });
    static readonly HashSet<string> Essential=["Id","Fecha","Cliente","Sitio","Tecnico","Estado","Descripcion","CostoMateriales","Unidad","SaldoInicial","Minimo","CostoReferencia","Entradas","Consumos","Existencia","Tipo","Material","Cantidad","Visita","CostoUnitario","Importe","BajoMinimo"];
    static void FillFormula(dynamic col,string formula,ComScope c){dynamic body=c.Own(col.DataBodyRange),cells=c.Own(body.Cells),first=c.Own(cells.Item[1,1]);first.Formula=formula;body.FillDown();}
    static void WriteCell(dynamic row,dynamic columns,FieldProfile f,string value,ComScope c,bool adding)
    {
        if(f.Calculated || (!adding && f.Key=="Id")) throw new InvalidOperationException("Identificadores y cálculos están protegidos.");
        dynamic col=c.Own(columns.Item[f.Header]), range=c.Own(row.Range), cells=c.Own(range.Cells), cell=c.Own(cells.Item[1,col.Index]);
        if(f.Type=="date") { if(!DateTime.TryParseExact(value,"yyyy-MM-dd",CultureInfo.InvariantCulture,DateTimeStyles.None,out var date)) throw new InvalidOperationException("Use fecha ISO yyyy-MM-dd."); cell.Value2=date.ToOADate();cell.NumberFormat="dd/mm/yyyy"; }
        else if(f.Type=="number") cell.Value2=(double)Rules.Number(value);
        else {cell.NumberFormat="@";cell.Value2=value;}
    }
    public void Dispose()
    {
        if(disposed)return;disposed=true;
        try { sta.Run(() => { try{ if(book is not null){book.Close(false);Marshal.ReleaseComObject(book);book=null;} } finally {if(app is not null){app.Quit();Marshal.ReleaseComObject(app);app=null;}} }); } finally {sta.Dispose();}
    }
}
