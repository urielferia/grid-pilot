using System.Text.Json;
using GridPilot.Core;
using GridPilot.Excel;
using GridPilot.Storage;

// Newline-delimited JSON-RPC over stdio. Stdout is reserved for protocol messages.
Session? session=null;
var initialized=false;
object RequireSession()=>session ?? throw new InvalidOperationException("Abra un libro desde GridPilot.");
var toolNames=new[]{"inspect_workbook","read_records","read_range","execute_batch","recalculate_validate","compare_revision","restore_checkpoint"};
object ToolSchema(string name)=>name switch {
 "execute_batch"=>JsonSerializer.Deserialize<object>("""
 {"type":"object","properties":{"plan":{"type":"object","properties":{"OperationId":{"type":"string","format":"uuid"},"Intent":{"type":"string"},"SourceRevision":{"type":"string"},"Operations":{"type":"array","minItems":1,"maxItems":100,"items":{"type":"object","properties":{"Kind":{"type":"string","enum":["SetValue","AddRecord","DeleteRecord","AddColumn","RenameColumn","DeleteColumn","SetFormula","AddSheet","RenameSheet","DeleteSheet"]},"Table":{"type":"string"},"Id":{"type":"string"},"Field":{"type":"string"},"Value":{"type":"string"},"Values":{"type":"object","additionalProperties":{"type":"string"}}},"required":["Kind"],"additionalProperties":false}}},"required":["OperationId","Intent","SourceRevision","Operations"],"additionalProperties":false}},"required":["plan"],"additionalProperties":false}
 """)!,
 "read_records"=>JsonSerializer.Deserialize<object>("""{"type":"object","properties":{"table":{"type":"string"},"query":{"type":"string"}},"required":["table","query"],"additionalProperties":false}""")!,
 "read_range"=>JsonSerializer.Deserialize<object>("""{"type":"object","properties":{"sheet":{"type":"string"},"addresses":{"type":"array","items":{"type":"string"},"maxItems":100}},"required":["sheet","addresses"],"additionalProperties":false}""")!,
 "restore_checkpoint" or "compare_revision"=>JsonSerializer.Deserialize<object>("""{"type":"object","properties":{"revision":{"type":"string"}},"required":["revision"],"additionalProperties":false}""")!,
 _=>new{type="object",properties=new{},additionalProperties=false}
};
try{
while(await Console.In.ReadLineAsync() is { } line)
{
    JsonElement id=default;
    try{
        if(line.Length>1_000_000)throw new InvalidOperationException("Mensaje demasiado grande.");
        using var doc=JsonDocument.Parse(line);var root=doc.RootElement;
        if(!root.TryGetProperty("id",out id)){if(root.TryGetProperty("method",out var n)&&n.GetString()=="notifications/initialized")initialized=true;continue;}
        id=id.Clone();var method=root.GetProperty("method").GetString();var p=root.TryGetProperty("params",out var param)?param:default;
        object result;
        if(method=="initialize"){initialized=true;result=new{protocolVersion="2025-06-18",capabilities=new{tools=new{}},serverInfo=new{name="GridPilot",version="0.1.0"}};}
        else if(!initialized)throw new InvalidOperationException("Inicialice el protocolo.");
        else if(method=="ping")result=new{};
        else if(method=="tools/list")result=new{tools=toolNames.Select(name=>new{name,description=name=="restore_checkpoint"?"Deshace únicamente el último lote de la sesión.":name,inputSchema=ToolSchema(name)})};
        else if(method=="gridpilot/open"){
            session?.Dispose();session=null;
            var profile=Json.Read<WorkbookProfile>(p.GetProperty("profile").GetRawText());var excel=new ExcelWorkbook();
            try{session=new(excel,p.GetProperty("root").GetString()!,p.GetProperty("path").GetString()!,profile);}catch{excel.Dispose();throw;}
            result=new{session.DirectoryPath,session.Editable,snapshot=session.Inspect()};
        }
        else if(method=="gridpilot/recover") {session?.Dispose();session=null;var excel=new ExcelWorkbook();try{session=new(excel,p.GetProperty("directory").GetString()!);}catch{excel.Dispose();throw;}result=new{session.DirectoryPath,session.Editable,snapshot=session.Inspect()};}
        else if(method=="gridpilot/approve"){RequireSession();result=session!.Approve(p.GetProperty("revision").GetString()!);}
        else if(method=="gridpilot/export"){RequireSession();result=session!.Export(p.GetProperty("path").GetString()!);}
        else if(method=="gridpilot/close"){session?.Dispose();session=null;result=new{};}
        else if(method=="tools/call"){
            RequireSession();var name=p.GetProperty("name").GetString();var a=p.GetProperty("arguments");object data;
            try{
                data=name switch {
                    "inspect_workbook"=>new{snapshot=session!.Inspect(),profile=session.Profile},
                    "execute_batch"=>session!.Execute(Json.Read<ActionPlan>(a.GetProperty("plan").GetRawText())),
                    "read_records"=>ReadRecords(session!,a),
                    "read_range"=>ReadRange(session!,a),
                    "restore_checkpoint"=>session!.Undo(a.GetProperty("revision").GetString()!),
                    "compare_revision"=>new{unchanged=session!.Inspect().Revision==a.GetProperty("revision").GetString()},
                    "recalculate_validate"=>Validate(session!),
                    _=>throw new InvalidOperationException("Herramienta desconocida.")};
                result=new{content=new[]{new{type="text",text=Json.Write(data)}},isError=false};
            }catch(Exception e){result=new{content=new[]{new{type="text",text=SafeError(e)}},isError=true};}
        }
        else throw new InvalidOperationException("Método desconocido.");
        await Console.Out.WriteLineAsync(Json.Write(new{jsonrpc="2.0",id,result}));
    }catch(Exception e){await Console.Out.WriteLineAsync(Json.Write(new{jsonrpc="2.0",id=id.ValueKind==JsonValueKind.Undefined?(object?)null:id,error=new{code=-32603,message=SafeError(e)}}));}
}
}finally{session?.Dispose();}
static string SafeError(Exception e)=>e is InvalidOperationException or InvalidDataException?e.Message:"La operación local falló ("+e.GetType().Name+"). Inspeccione el borrador o recupere la sesión; no reintente escrituras automáticamente.";
static object ReadRecords(Session s,JsonElement a){var name=a.GetProperty("table").GetString()!;s.Profile.Table(name);var q=a.GetProperty("query").GetString()!;if(q.Length is <1 or >100)throw new InvalidOperationException("Use una búsqueda de 1 a 100 caracteres.");var rows=s.Inspect().Table(name).Rows.Where(r=>r.Values.Any(v=>v.Contains(q,StringComparison.OrdinalIgnoreCase))).Take(21).ToList();return new{rows=rows.Take(20).Select(r=>r.ToDictionary(p=>p.Key,p=>p.Value.Length>500?p.Value[..500]:p.Value)),truncated=rows.Count>20||rows.Any(r=>r.Values.Any(v=>v.Length>500))};}
static object ReadRange(Session s,JsonElement a){var addresses=a.GetProperty("addresses").EnumerateArray().Select(v=>v.GetString()).ToHashSet();if(addresses.Count>100)throw new InvalidOperationException("Máximo 100 celdas.");return s.Inspect().Cells.Where(c=>c.Sheet==a.GetProperty("sheet").GetString()&&addresses.Contains(c.Address)).ToList();}
static object Validate(Session s){var snap=s.RecalculateValidate();return new{valid=true,revision=snap.Revision};}
