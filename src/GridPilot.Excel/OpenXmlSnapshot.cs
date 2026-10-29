using System.IO.Compression;
using System.Xml.Linq;
using GridPilot.Core;
namespace GridPilot.Excel;

// Excel creates this transient copy, so values reflect native calculation and unsaved edits.
// ZIP metadata is intentionally excluded; workbook content and styles determine the revision.
internal static class OpenXmlSnapshot
{
    static readonly XNamespace X="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    static readonly XNamespace R="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public static WorkbookSnapshot Read(string path,Dictionary<(string Sheet,string Address),string> formulas)
    {
        using var zip=ZipFile.OpenRead(path);
        XDocument Read(string name){using var stream=zip.GetEntry(name)?.Open()??throw new InvalidDataException("Parte XLSX faltante.");return XDocument.Load(stream);}
        Dictionary<string,string> Relations(string source)
        {
            var folder=source[..(source.LastIndexOf('/')+1)];var file=source[(source.LastIndexOf('/')+1)..];var entry=folder+"_rels/"+file+".rels";
            if(zip.GetEntry(entry) is null)return [];
            return Read(entry).Root!.Elements().ToDictionary(e=>(string)e.Attribute("Id")!,e=>new Uri(new Uri("https://package/"+source),(string)e.Attribute("Target")!).AbsolutePath.TrimStart('/'));
        }
        var wb=Read("xl/workbook.xml");var links=Relations("xl/workbook.xml");
        var shared=zip.GetEntry("xl/sharedStrings.xml") is null?[]:Read("xl/sharedStrings.xml").Descendants(X+"si").Select(e=>string.Concat(e.Descendants(X+"t").Select(t=>t.Value))).ToArray();
        var sheets=new List<string>();var cells=new List<CellState>();var tables=new List<TableState>();var props=new List<string>();
        var styles=Read("xl/styles.xml");props.Add("styles:"+styles.ToString(SaveOptions.DisableFormatting));
        foreach(var part in wb.Descendants(X+"sheet"))
        {
            var name=(string)part.Attribute("name")!;sheets.Add(name);props.Add("sheet:"+name+":"+(string?)part.Attribute("state"));
            var source=links[(string)part.Attribute(R+"id")!];var doc=Read(source);var sheetLinks=Relations(source);var local=new Dictionary<string,CellState>();
            foreach(var cell in doc.Descendants(X+"c"))
            {
                var address=(string)cell.Attribute("r")!;var type=(string?)cell.Attribute("t");var raw=cell.Element(X+"v")?.Value??"";
                var value=type=="s"?shared[int.Parse(raw)]:type=="inlineStr"?string.Concat(cell.Descendants(X+"t").Select(t=>t.Value)):raw;
                var formula=cell.Element(X+"f");
                var state=new CellState(name,address,value,formula is null?"":formulas.GetValueOrDefault((name,address))??"="+formula.Value,(string?)cell.Attribute("s")??"0");local[address]=state;cells.Add(state);
            }
            foreach(var tablePart in doc.Descendants(X+"tablePart"))
            {
                var table=Read(sheetLinks[(string)tablePart.Attribute(R+"id")!]);var element=table.Root!;var address=(string)element.Attribute("ref")!;
                var headers=element.Descendants(X+"tableColumn").Select(e=>(string)e.Attribute("name")!).ToList();var bounds=address.Split(':');var start=Parse(bounds[0]);var end=Parse(bounds[^1]);
                if(end.Row>2000||end.Col>50)throw new InvalidOperationException("Tabla excede límites del piloto.");
                var rows=new List<Dictionary<string,string>>();int totals=(int?)element.Attribute("totalsRowCount")??0;
                for(int r=start.Row+1;r<=end.Row-totals;r++){var row=new Dictionary<string,string>();for(int c=0;c<headers.Count;c++)row[headers[c]]=local.GetValueOrDefault(Address(start.Col+c,r))?.Value??"";rows.Add(row);}
                tables.Add(new((string)element.Attribute("name")!,name,address,headers,rows));props.Add("table:"+table.ToString(SaveOptions.DisableFormatting));
            }
            doc.Descendants(X+"sheetViews").Remove(); // Selection and scroll position do not change workbook meaning.
            foreach(var cell in doc.Descendants(X+"c"))cell.RemoveNodes(); // Values/formulas are compared through CellState, not as formatting.
            props.Add(name+":"+doc.ToString(SaveOptions.DisableFormatting));
        }
        wb.Descendants(X+"bookViews").Remove();wb.Descendants(X+"fileVersion").Remove();
        // SaveCopyAs assigns a fresh coauthoring documentId to each file. It is not a content change.
        wb.Descendants().Where(e=>e.Name.LocalName is "revisionPtr" or "absPath").Remove();
        props.Add("workbook:"+wb.ToString(SaveOptions.DisableFormatting));
        return new(sheets,tables,cells,props);
    }
    static (int Col,int Row) Parse(string address){int c=0,i=0;while(i<address.Length&&char.IsLetter(address[i])){c=c*26+address[i]-'A'+1;i++;}return(c,int.Parse(address[i..]));}
    static string Address(int col,int row){string s="";while(col>0){col--;s=(char)('A'+col%26)+s;col/=26;}return s+row;}
}
