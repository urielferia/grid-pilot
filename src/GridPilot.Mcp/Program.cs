using System.Text.Json;
using GridPilot.Core;
using GridPilot.Excel;
using GridPilot.Storage;

// Newline-delimited JSON-RPC over stdio. Stdout is reserved for protocol messages.
Console.InputEncoding = new System.Text.UTF8Encoding(false);
Console.OutputEncoding = new System.Text.UTF8Encoding(false);
Session? session = null;
var initialized = false;
object RequireSession() => session ?? throw new InvalidOperationException("Open a workbook from GridPilot first.");
var toolNames = new[] { "inspect_workbook", "read_records", "read_range", "execute_batch", "recalculate_validate", "compare_revision", "restore_checkpoint" };
object ToolSchema(string name) => name switch
{
    "execute_batch" => JsonSerializer.Deserialize<object>("""
 {"type":"object","properties":{"plan":{"type":"object","properties":{"OperationId":{"type":"string","format":"uuid"},"Intent":{"type":"string"},"SourceRevision":{"type":"string"},"Operations":{"type":"array","minItems":1,"maxItems":100,"items":{"type":"object","properties":{"Kind":{"type":"string","enum":["SetValue","AddRecord","DeleteRecord","AddColumn","RenameColumn","DeleteColumn","SetFormula","AddSheet","RenameSheet","DeleteSheet"]},"Table":{"type":"string"},"Id":{"type":"string"},"Field":{"type":"string"},"Value":{"type":"string"},"Values":{"type":"object","additionalProperties":{"type":"string"}}},"required":["Kind"],"additionalProperties":false}}},"required":["OperationId","Intent","SourceRevision","Operations"],"additionalProperties":false}},"required":["plan"],"additionalProperties":false}
 """)!,
    "read_records" => JsonSerializer.Deserialize<object>("""{"type":"object","properties":{"table":{"type":"string"},"query":{"type":"string"}},"required":["table","query"],"additionalProperties":false}""")!,
    "read_range" => JsonSerializer.Deserialize<object>("""{"type":"object","properties":{"sheet":{"type":"string"},"addresses":{"type":"array","items":{"type":"string"},"maxItems":100}},"required":["sheet","addresses"],"additionalProperties":false}""")!,
    "restore_checkpoint" or "compare_revision" => JsonSerializer.Deserialize<object>("""{"type":"object","properties":{"revision":{"type":"string"}},"required":["revision"],"additionalProperties":false}""")!,
    _ => new { type = "object", properties = new { }, additionalProperties = false }
};
try
{
    while (await Console.In.ReadLineAsync() is { } line)
    {
        JsonElement id = default;
        try
        {
            if (line.Length > 1_000_000) throw new InvalidOperationException("Message is too large.");
            using var doc = JsonDocument.Parse(line); var root = doc.RootElement;
            if (!root.TryGetProperty("id", out id)) { if (root.TryGetProperty("method", out var n) && n.GetString() == "notifications/initialized") initialized = true; continue; }
            id = id.Clone(); var method = root.GetProperty("method").GetString(); var p = root.TryGetProperty("params", out var param) ? param : default;
            object result;
            if (method == "initialize") { initialized = true; result = new { protocolVersion = "2025-06-18", capabilities = new { tools = new { } }, serverInfo = new { name = "GridPilot", version = "0.1.0" } }; }
            else if (!initialized) throw new InvalidOperationException("Initialize the protocol first.");
            else if (method == "ping") result = new { };
            else if (method == "tools/list") result = new { tools = toolNames.Select(name => new { name, description = name == "restore_checkpoint" ? "Undoes only the last batch of the session." : name, inputSchema = ToolSchema(name) }) };
            else if (method == "gridpilot/open")
            {
                session?.Dispose(); session = null;
                var profile = Json.Read<WorkbookProfile>(p.GetProperty("profile").GetRawText()); var excel = new ExcelWorkbook();
                try { session = new(excel, p.GetProperty("root").GetString()!, p.GetProperty("path").GetString()!, profile); } catch { excel.Dispose(); throw; }
                Owner.Record(session.DirectoryPath, excel);
                result = new { session.DirectoryPath, session.Editable, snapshot = session.Inspect() };
            }
            else if (method == "gridpilot/recover")
            {
                session?.Dispose(); session = null; var directory = p.GetProperty("directory").GetString()!; Owner.ReleaseOrphan(directory); var excel = new ExcelWorkbook();
                try { session = new(excel, directory); } catch { excel.Dispose(); throw; }
                Owner.Record(session.DirectoryPath, excel); result = new { session.DirectoryPath, session.Editable, snapshot = session.Inspect() };
            }
            else if (method == "gridpilot/approve") { RequireSession(); result = session!.Approve(p.GetProperty("revision").GetString()!); }
            else if (method == "gridpilot/export") { RequireSession(); result = session!.Export(p.GetProperty("path").GetString()!); }
            else if (method == "gridpilot/close") { session?.Dispose(); session = null; result = new { }; }
            else if (method == "tools/call")
            {
                RequireSession(); var name = p.GetProperty("name").GetString(); var a = p.GetProperty("arguments"); object data;
                try
                {
                    data = name switch
                    {
                        "inspect_workbook" => new { snapshot = session!.Inspect(), profile = session.Profile },
                        "execute_batch" => session!.Execute(Json.Read<ActionPlan>(a.GetProperty("plan").GetRawText())),
                        "read_records" => ReadRecords(session!, a),
                        "read_range" => ReadRange(session!, a),
                        "restore_checkpoint" => session!.Undo(a.GetProperty("revision").GetString()!),
                        "compare_revision" => new { unchanged = session!.Inspect().Revision == a.GetProperty("revision").GetString() },
                        "recalculate_validate" => Validate(session!),
                        _ => throw new InvalidOperationException("Unknown tool.")
                    };
                    result = new { content = new[] { new { type = "text", text = Json.Write(data) } }, isError = false };
                }
                catch (Exception e) { result = new { content = new[] { new { type = "text", text = SafeError(e) } }, isError = true }; }
            }
            else throw new InvalidOperationException("Unknown method.");
            await Console.Out.WriteLineAsync(Json.Write(new { jsonrpc = "2.0", id, result }));
        }
        catch (Exception e) { await Console.Out.WriteLineAsync(Json.Write(new { jsonrpc = "2.0", id = id.ValueKind == JsonValueKind.Undefined ? (object?)null : id, error = new { code = -32603, message = SafeError(e) } })); }
    }
}
finally { session?.Dispose(); }
static string SafeError(Exception e) => e is System.Runtime.InteropServices.COMException com ? $"Excel could not complete the operation (COM 0x{com.HResult:X8}). Close the draft and recover the session from an interactive Windows session." : e is InvalidOperationException or InvalidDataException ? e.Message : "Local operation failed (" + e.GetType().Name + "). Inspect the draft or recover the session; do not retry writes automatically.";
static object ReadRecords(Session s, JsonElement a)
{
    var name = a.GetProperty("table").GetString()!;
    var tableProfile = s.Profile.Table(name);
    var dateHeaders = tableProfile.Fields.Where(f => f.Type == "date").Select(f => f.Header).ToHashSet(StringComparer.OrdinalIgnoreCase);
    var q = a.GetProperty("query").GetString()!;
    if (q.Length > 100) throw new InvalidOperationException("Use a search query up to 100 characters.");
    var all = s.Inspect().Table(name).Rows;
    static string FormatCell(string header, string value, HashSet<string> dates)
    {
        if (dates.Contains(header) && double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var oa) && oa >= 1 && oa <= 2958465)
            return DateTime.FromOADate(oa).ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
        return value.Length > 500 ? value[..500] : value;
    }
    var indexed = all.Select((r, idx) =>
    {
        var dict = new Dictionary<string, string> { ["_row"] = (idx + 1).ToString() };
        foreach (var p in r) dict[p.Key] = FormatCell(p.Key, p.Value, dateHeaders);
        return dict;
    });
    var trim = q.Trim();
    IEnumerable<Dictionary<string, string>> filtered;
    if (trim is "" or "*" || trim.Equals("all", StringComparison.OrdinalIgnoreCase) || trim.Equals("todos", StringComparison.OrdinalIgnoreCase) || trim.Equals("todo", StringComparison.OrdinalIgnoreCase))
        filtered = indexed;
    else if (int.TryParse(trim, out var rowNum) && rowNum >= 1 && rowNum <= all.Count)
        filtered = indexed.Where(r => r["_row"] == rowNum.ToString() || r.Where(kv => kv.Key != "_row").Any(kv => kv.Value.Contains(trim, StringComparison.OrdinalIgnoreCase)));
    else if (System.Text.RegularExpressions.Regex.Match(trim, @"^(?:fila|row|registro|#)\s*(\d+)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase) is { Success: true } m && int.TryParse(m.Groups[1].Value, out var n))
        filtered = indexed.Where(r => r["_row"] == n.ToString());
    else if (System.Text.RegularExpressions.Regex.Match(trim, @"^(\d+)\s*(?:-|\.\.|a|to)\s*(\d+)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase) is { Success: true } rng && int.TryParse(rng.Groups[1].Value, out var aRow) && int.TryParse(rng.Groups[2].Value, out var bRow))
        filtered = indexed.Where(r => { var k = int.Parse(r["_row"]); return k >= Math.Min(aRow, bRow) && k <= Math.Max(aRow, bRow); });
    else
        filtered = indexed.Where(r => r.Where(kv => kv.Key != "_row").Any(kv => kv.Value.Contains(trim, StringComparison.OrdinalIgnoreCase)));
    var rows = filtered.Take(21).ToList();
    return new { totalRows = all.Count, rows = rows.Take(20).ToList(), truncated = rows.Count > 20 || all.Any(r => r.Values.Any(v => v.Length > 500)) };
}
static object ReadRange(Session s, JsonElement a) { var addresses = a.GetProperty("addresses").EnumerateArray().Select(v => v.GetString()).ToHashSet(); if (addresses.Count > 100) throw new InvalidOperationException("Maximum 100 cells per request."); return s.Inspect().Cells.Where(c => c.Sheet == a.GetProperty("sheet").GetString() && addresses.Contains(c.Address)).ToList(); }
static object Validate(Session s) { var snap = s.RecalculateValidate(); return new { valid = true, revision = snap.Revision }; }
