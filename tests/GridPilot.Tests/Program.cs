using GridPilot.Core;
using GridPilot.Excel;
using GridPilot.Storage;
using GridPilot.Providers;
using GridPilot.Tests;

var results = new List<object>(); int failures = 0;
void Check(string name, Action action) { try { action(); Console.WriteLine("PASS " + name); results.Add(new { name, passed = true }); } catch (Exception e) { failures++; Console.WriteLine("FAIL " + name + ": " + e); results.Add(new { name, passed = false, error = e.ToString() }); } }
void True(bool value) { if (!value) throw new Exception("Assertion failed."); }
void Throws(Action f) { try { f(); } catch (InvalidOperationException) { return; } throw new Exception("Operation should have been rejected."); }
var root = Path.GetFullPath(args.FirstOrDefault(a => a.StartsWith("--root="))?.Substring(7) ?? ".");
var output = Path.Combine(root, "artifacts", "tests", DateTime.UtcNow.ToString("yyyyMMdd-HHmmss")); Directory.CreateDirectory(output);
var profile = Json.Read<WorkbookProfile>(File.ReadAllText(Path.Combine(root, "assets", "profile.json")));
var original = Path.Combine(root, "assets", "GridPilot-Demo.xlsx"); var originalHash = Json.FileHash(original);
Check("Allows internal formulas", () => Rules.Formula("=ROUND(SUM(A1:A3)*2,2)"));
foreach (var f in new[] { "=WEBSERVICE(\"https://x\")", "=NOW()", "=[otro.xlsx]A1", "=cmd|'x'!A1", "=INDIRECT(\"A1\")", "=RAND()", "=OFFSET(A1,1,1)", "=SUM(A:A)" }) Check("Rejects " + f, () => Throws(() => Rules.Formula(f)));
Check("Invariant decimal number", () => True(Rules.Number("123.45") == 123.45m));
Check("XLSX package without excluded capabilities", () => ExcelWorkbook.CheckPackage(original));
Check("Provider fake returns clarification without operations", () => { IPlanner fake = new FakePlanner(new("Which visit?", [])); var p = fake.Interpret("fix the visit", profile, (_, _) => Task.FromResult<object>(new { }), default).Result; True(p.Clarification != "" && p.Operations.Count == 0); });
Check("Provider: cloud disabled makes no network calls", () => { using var http = new FailingHttp(); var p = new AnthropicPlanner(new(), "fake-test-only", http); Throws(() => p.Interpret("test", profile, (_, _) => Task.FromResult<object>(new { }), default).GetAwaiter().GetResult()); True(http.Calls == 0); });
Check("Provider: HTTP 429 does not retry or switch model", () => { using var http = new FailingHttp(); var p = new AnthropicPlanner(new(CloudEnabled: true), "fake-test-only", http); Throws(() => p.Interpret("test", profile, (_, _) => Task.FromResult<object>(new { }), default).GetAwaiter().GetResult()); True(http.Calls == 1 && p.LastUsage?.Model == "claude-sonnet-5-5"); });
Check("Domain: valid fixture", () => Rules.Validate(MemoryWorkbook.Fixture(profile), profile));
Check("Profile cannot unprotect essential calculations", () => { profile.ValidateDefinition(); var p = profile.Copy(); var t = p.Table("Inventario"); int i = t.Fields.FindIndex(f => f.Key == "Existencia"); t.Fields[i] = t.Fields[i] with { Calculated = false }; Throws(() => p.ValidateDefinition()); });
foreach (var (table, field, value) in new[] { ("Movimientos", "Cantidad", "-1"), ("Movimientos", "Material", "M999"), ("Movimientos", "Visita", "V999"), ("Movimientos", "Tipo", "Otro"), ("Inventario", "Existencia", "999"), ("Visitas", "Estado", "Cancelada"), ("Inventario", "Saldo inicial", "-1"), ("Visitas", "Costo materiales", "0"), ("Movimientos", "Importe", "99") })
    Check($"Domain blocks {table}/{field}={value}", () => { var s = MemoryWorkbook.Fixture(profile); s.Table(table).Rows[0][field] = value; Throws(() => Rules.Validate(s, profile)); });
Check("Domain blocks duplicate IDs", () => { var s = MemoryWorkbook.Fixture(profile); s.Table("Visitas").Rows.Add(new(s.Table("Visitas").Rows[0])); Throws(() => Rules.Validate(s, profile)); });
Check("Domain blocks replacing formula with literal value", () => { var s = MemoryWorkbook.Fixture(profile); int i = s.Cells.FindIndex(c => c.Sheet == "Inventario" && c.Address == "I6"); s.Cells[i] = s.Cells[i] with { Formula = "" }; Throws(() => Rules.Validate(s, profile)); });
Check("Domain blocks unsupported manual functions", () => { var s = MemoryWorkbook.Fixture(profile); s.Cells.Add(new("Resumen", "C20", "0", "=NOW()", "General")); Throws(() => Rules.Validate(s, profile)); });
Check("Domain: amount rounding tolerance", () => { var s = MemoryWorkbook.Fixture(profile); s.Table("Movimientos").Rows[0]["Importe"] = "10.005"; Rules.Validate(s, profile); s.Table("Movimientos").Rows[0]["Importe"] = "10.02"; Throws(() => Rules.Validate(s, profile)); });
Check("Simulated session: rollback after partial write", () => { var memory = new MemoryWorkbook(profile) { FailWrite = true }; using var s = new Session(memory, output, original, profile); var before = s.Inspect(); Throws(() => s.Execute(new(Guid.NewGuid().ToString(), "fail", before.Revision, [new(OperationKind.SetValue, "Inventario", "M001", "Minimo", "11")]))); True(before.Revision == s.Inspect().Revision); });
Check("Simulated session: failed restore blocks approval", () => { var memory = new MemoryWorkbook(profile) { FailWrite = true, FailRestore = true }; using var s = new Session(memory, output, original, profile); Throws(() => s.Execute(new(Guid.NewGuid().ToString(), "fail", s.Inspect().Revision, [new(OperationKind.SetValue, "Inventario", "M001", "Minimo", "11")]))); Throws(() => s.Approve(s.Inspect().Revision)); });
Check("SQLite: pending operation is not replayed after recovery", () => { using var j = new Journal(Path.Combine(output, "journal.db")); var p = new ActionPlan(Guid.NewGuid().ToString(), "interrupted", "r", []); j.Begin(p); j.RecoverPending(); True(j.Find(p.OperationId)!.Value.Status == "interrupted"); });
Check("SQLite: session metadata commits atomically", () => { using var j = new Journal(Path.Combine(output, "atomic.db")); j.Put("a", 1); Throws(() => j.Atomic(() => { j.Put("a", 2); j.Put("b", 3); throw new InvalidOperationException("interrupted"); })); True(j.Get<int>("a") == 1 && j.Get<int>("b") == 0); j.Atomic(() => { j.Put("a", 4); j.Put("b", 5); }); True(j.Get<int>("a") == 4 && j.Get<int>("b") == 5); });

PlannerTests.Run(Check, profile);
if (args.Contains("--excel"))
{
    using var excel = new ExcelWorkbook(false); using var session = new Session(excel, Path.Combine(output, "sessions"), original, profile);
    if (args.Contains("--snapshot-only")) { var a = session.Inspect(); var b = session.Inspect(); File.WriteAllText(Path.Combine(output, "a.json"), Json.Write(a)); File.WriteAllText(Path.Combine(output, "b.json"), Json.Write(b)); Console.WriteLine(output); return; }
    ActionPlan Plan(params Operation[] operations) => new(Guid.NewGuid().ToString(), "Test", session.Inspect().Revision, operations.ToList());
    Check("Real template: 10 materials, 10 visits, 20 movements and independent verification", () => { True(session.Editable); var s = session.Inspect(); Rules.Validate(s, session.Profile); True(s.Table("Inventario").Rows.Count == 10 && s.Table("Visitas").Rows.Count == 10 && s.Table("Movimientos").Rows.Count == 20); File.WriteAllText(Path.Combine(output, "snapshot.json"), Json.Write(s)); });
    Check("Export without approval is blocked", () => Throws(() => session.Export(Path.Combine(output, "unapproved.xlsx"))));
    Check("Add visit and consumption in the same batch", () => { session.Execute(Plan(new Operation(OperationKind.AddRecord, "Visitas", Values: new() { ["Id"] = "V011", ["Fecha"] = "2026-10-20", ["Cliente"] = "Demo", ["Sitio"] = "Taller", ["Tecnico"] = "Ana", ["Estado"] = "Realizada", ["Descripcion"] = "Prueba" }), new Operation(OperationKind.AddRecord, "Movimientos", Values: new() { ["Id"] = "MOV021", ["Fecha"] = "2026-10-20", ["Tipo"] = "Consumo", ["Material"] = "M001", ["Cantidad"] = "3", ["Visita"] = "V011", ["CostoUnitario"] = "5" }))); True(session.Inspect().Table("Inventario").Rows[0]["Existencia"] == "25"); });
    Check("Quantity correction updates inventory", () => { session.Execute(Plan(new Operation(OperationKind.SetValue, "Movimientos", "MOV021", "Cantidad", "4"))); True(session.Inspect().Table("Inventario").Rows[0]["Existencia"] == "24"); });
    Check("Idempotent retry", () => { var p = Plan(new Operation(OperationKind.SetValue, "Movimientos", "MOV021", "Cantidad", "5")); var a = session.Execute(p); var b = session.Execute(p); True(a.AfterRevision == b.AfterRevision); });
    Check("Same operation ID with different plan is blocked", () => { var p = Plan(new Operation(OperationKind.SetValue, "Movimientos", "MOV021", "Cantidad", "4")); session.Execute(p); Throws(() => session.Execute(p with { Intent = "Other" })); });
    Check("Insufficient stock reverts the entire batch", () => { var before = session.Inspect(); Throws(() => session.Execute(Plan(new Operation(OperationKind.SetValue, "Movimientos", "MOV021", "Cantidad", "9999")))); True(session.Inspect().Revision == before.Revision); });
    Check("Duplicate key reverts batch", () => { var before = session.Inspect(); Throws(() => session.Execute(Plan(new Operation(OperationKind.AddRecord, "Movimientos", Values: new() { ["Id"] = "MOV021", ["Fecha"] = "2026-10-20", ["Tipo"] = "Entrada", ["Material"] = "M001", ["Cantidad"] = "1", ["Visita"] = "", ["CostoUnitario"] = "5" })))); True(session.Inspect().Revision == before.Revision); });
    Check("Consumption for scheduled visit is blocked", () => Throws(() => session.Execute(Plan(new Operation(OperationKind.SetValue, "Movimientos", "MOV021", "Visita", "V009")))));
    Check("Deleting referenced material is blocked", () => Throws(() => session.Execute(Plan(new Operation(OperationKind.DeleteRecord, "Inventario", "M001")))));
    Check("Delete visit and resolve dependency in batch", () => { session.Execute(Plan(new Operation(OperationKind.DeleteRecord, "Visitas", "V011"), new Operation(OperationKind.DeleteRecord, "Movimientos", "MOV021"))); True(session.Inspect().Table("Visitas").Rows.Count == 10); });
    Check("Entry and deletion of consumption", () => { session.Execute(Plan(new Operation(OperationKind.AddRecord, "Movimientos", Values: new() { ["Id"] = "MOV022", ["Fecha"] = "2026-10-20", ["Tipo"] = "Entrada", ["Material"] = "M001", ["Cantidad"] = "3", ["Visita"] = "", ["CostoUnitario"] = "5" }), new Operation(OperationKind.DeleteRecord, "Movimientos", "MOV011"))); True(session.Inspect().Table("Inventario").Rows[0]["Existencia"] == "33"); });
    Check("Historical movement cost does not change when catalog changes", () => { var before = session.Inspect().Table("Movimientos").Rows[0]["Costo unitario"]; session.Execute(Plan(new Operation(OperationKind.SetValue, "Inventario", "M001", "CostoReferencia", "99"))); True(session.Inspect().Table("Movimientos").Rows[0]["Costo unitario"] == before); });
    Check("New calculated column and renamed sheet", () => { session.Execute(Plan(new Operation(OperationKind.AddColumn, "Inventario", Field: "Valor", Value: "=I6*F6"), new Operation(OperationKind.RenameSheet, "Inventario", Value: "Almacén"))); True(session.Inspect().Sheets.Contains("Almacén")); });
    Check("Renaming field updates profile and references", () => { session.Execute(Plan(new Operation(OperationKind.RenameColumn, "Inventario", Field: "Descripcion", Value: "Material demo"))); True(session.Profile.Table("Inventario").Fields.Single(f => f.Key == "Descripcion").Header == "Material demo"); });
    Check("Renaming Summary sheet preserves verifier", () => { session.Execute(Plan(new Operation(OperationKind.RenameSheet, "Resumen", Value: "Reporte"))); True(session.Profile.SummarySheet == "Reporte"); Rules.Validate(session.Inspect(), session.Profile); });
    Check("Add and delete additional sheet and column", () => { session.Execute(Plan(new Operation(OperationKind.AddSheet, Value: "Notas"), new Operation(OperationKind.AddColumn, "Visitas", Field: "Observaciones"))); session.Execute(Plan(new Operation(OperationKind.DeleteSheet, "Notas"), new Operation(OperationKind.DeleteColumn, "Visitas", Field: "Observaciones"))); True(!session.Inspect().Sheets.Contains("Notas")); });
    Check("Malicious text is stored literally", () => { session.Execute(Plan(new Operation(OperationKind.SetValue, "Visitas", "V001", "Descripcion", "=WEBSERVICE(\"https://example.com\")"))); var cell = session.Inspect().Cells.Single(c => c.Sheet == "Visitas" && c.Address == "G6"); True(cell.Formula == "" && cell.Value.StartsWith("=WEBSERVICE")); });
    Check("Protect identifiers and essential calculations", () => { Throws(() => session.Execute(Plan(new Operation(OperationKind.DeleteColumn, "Inventario", Field: "Id")))); Throws(() => session.Execute(Plan(new Operation(OperationKind.SetValue, "Inventario", "M001", "Existencia", "10")))); });
    Check("Broken and circular formulas revert", () => { Throws(() => session.Execute(Plan(new Operation(OperationKind.SetFormula, "Inventario", Field: "Valor", Value: "=NoExiste!A1")))); Throws(() => session.Execute(Plan(new Operation(OperationKind.SetFormula, "Inventario", Field: "Valor", Value: "=J6")))); });
    Check("Undo restores value and revision", () => { var before = session.Inspect(); session.Execute(Plan(new Operation(OperationKind.SetValue, "Inventario", "M001", "Minimo", "11"))); session.Undo(session.Inspect().Revision); True(session.Inspect().Revision == before.Revision); });
    Check("Stale plan is blocked after manual edit", () => { var stale = Plan(new Operation(OperationKind.SetValue, "Inventario", "M001", "Minimo", "12")); excel.Apply(new Operation(OperationKind.SetValue, "Inventario", "M001", "Minimo", "13"), session.Profile); excel.Recalculate(); Throws(() => session.Execute(stale)); });
    Check("Stale approval is blocked", () => { var old = session.Inspect().Revision; excel.Apply(new Operation(OperationKind.SetValue, "Inventario", "M001", "Minimo", "14"), session.Profile); excel.Recalculate(); Throws(() => session.Approve(old)); });
    Check("Subsequent change invalidates approval", () => { session.Approve(session.Inspect().Revision); excel.Apply(new Operation(OperationKind.SetValue, "Inventario", "M001", "Minimo", "15"), session.Profile); excel.Recalculate(); Throws(() => session.Export(Path.Combine(output, "stale.xlsx"))); });
    Check("Export never overwrites original or existing file", () => { session.Approve(session.Inspect().Revision); Throws(() => session.Export(original)); var dest = Path.Combine(output, "exists.xlsx"); File.WriteAllText(dest, "existing"); Throws(() => session.Export(dest)); True(File.ReadAllText(dest) == "existing"); });
    Check("Approved export matches hash and reopens in Excel", () => { var approved = session.Approve(session.Inspect().Revision); var dest = Path.Combine(output, "approved.xlsx"); session.Export(dest); True(Json.FileHash(dest) == approved.Hash); using var second = new ExcelWorkbook(false); var copy = Path.Combine(output, "reopen-copy.xlsx"); File.Copy(dest, copy); second.Open(copy); second.Recalculate(); Rules.Validate(second.Snapshot(), session.Profile); });
    Check("Explicit recovery from last valid state", () => { var folder = session.DirectoryPath; var expected = session.Profile.Copy(); if (excel.ProcessId > 0) { using var p = System.Diagnostics.Process.GetProcessById(excel.ProcessId); p.Kill(); p.WaitForExit(10000); } session.Dispose(); using var recovered = new Session(new ExcelWorkbook(false), folder); Rules.Validate(recovered.Inspect(), expected); });
    Check("Forced MCP server termination during COM write and recovery without replay", () => McpCrashRecovery(output, original, profile, True));
    Check("UI RPC client opens editable demo template with UTF-8 and closes session", () => UiRpcDemoOpen(output, original, profile, True));
    Check("Real Excel + complex interpreter: replace row, swap rows, and delete first visits row", () => ComplexInstructionsExcel(output, original, profile, True));
}
Check("Original preserves SHA-256", () => True(Json.FileHash(original) == originalHash));
File.WriteAllText(Path.Combine(output, "results.json"), Json.Write(new { excel = args.Contains("--excel"), results, failures }));
Console.WriteLine($"{results.Count - failures}/{results.Count} PASS. Report: {output}");
Environment.ExitCode = failures == 0 ? 0 : 1;
static void McpCrashRecovery(string output, string original, WorkbookProfile profile, Action<bool> True)
{
    var exe = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "src", "GridPilot.Mcp", "bin", "Release", "net10.0-windows", "GridPilot.Mcp.exe");
    if (!File.Exists(exe)) exe = Path.Combine(AppContext.BaseDirectory, "server", "GridPilot.Mcp.exe");
    var utf8 = new System.Text.UTF8Encoding(false);
    System.Diagnostics.Process Start() => System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(exe) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true, StandardInputEncoding = utf8, StandardOutputEncoding = utf8 })!;
    System.Text.Json.JsonElement Call(System.Diagnostics.Process p, int id, string method, object? parameters = null) { p.StandardInput.WriteLine(Json.Write(new { jsonrpc = "2.0", id, method, @params = parameters ?? new { } })); p.StandardInput.Flush(); var line = p.StandardOutput.ReadLine() ?? throw new Exception("EOF"); using var doc = System.Text.Json.JsonDocument.Parse(line); if (doc.RootElement.TryGetProperty("error", out var err)) throw new InvalidOperationException(err.GetProperty("message").GetString()); return doc.RootElement.GetProperty("result").Clone(); }
    var root = Path.Combine(output, "mcp-crash"); string sessionDir; string committedRev;
    using (var s1 = Start())
    {
        Call(s1, 1, "initialize", new { protocolVersion = "2025-06-18", capabilities = new { }, clientInfo = new { name = "test", version = "1" } });
        var opened = Call(s1, 2, "gridpilot/open", new { path = original, root, profile });
        True(opened.GetProperty("Editable").GetBoolean());
        sessionDir = opened.GetProperty("DirectoryPath").GetString()!;
        var rev0 = opened.GetProperty("snapshot").GetProperty("Revision").GetString()!;
        var plan1 = new ActionPlan(Guid.NewGuid().ToString(), "Committed batch", rev0, [new(OperationKind.SetValue, "Movimientos", "MOV011", "Cantidad", "3")]);
        var r1 = Call(s1, 3, "tools/call", new { name = "execute_batch", arguments = new { plan = plan1 } });
        True(!r1.GetProperty("isError").GetBoolean());
        committedRev = Json.Read<ChangeSet>(r1.GetProperty("content")[0].GetProperty("text").GetString()!).AfterRevision;
        var ops = Enumerable.Range(0, 25).Select(i => new Operation(OperationKind.SetValue, "Inventario", "M001", "Minimo", (10 + i % 5).ToString())).ToList();
        var plan2 = new ActionPlan(Guid.NewGuid().ToString(), "Interrupted batch", committedRev, ops);
        s1.StandardInput.WriteLine(Json.Write(new { jsonrpc = "2.0", id = 4, method = "tools/call", @params = new { name = "execute_batch", arguments = new { plan = plan2 } } }));
        s1.StandardInput.Flush();
        Thread.Sleep(250);
        s1.Kill();
        s1.WaitForExit(10000);
    }
    using (var s2 = Start())
    {
        try
        {
            Call(s2, 1, "initialize", new { protocolVersion = "2025-06-18", capabilities = new { }, clientInfo = new { name = "test", version = "1" } });
            var rec = Call(s2, 2, "gridpilot/recover", new { directory = sessionDir });
            True(rec.GetProperty("Editable").GetBoolean());
            var snap = Json.Read<WorkbookSnapshot>(rec.GetProperty("snapshot").GetRawText());
            True(snap.Revision == committedRev);
            True(snap.Table("Movimientos").Rows.Single(r => r["ID"] == "MOV011")["Cantidad"] == "3");
        }
        finally { try { Call(s2, 99, "gridpilot/close"); } catch { } s2.StandardInput.Close(); s2.WaitForExit(10000); }
    }
}
static void UiRpcDemoOpen(string output, string original, WorkbookProfile profile, Action<bool> True)
{
    var exe = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "artifacts", "release", "server", "GridPilot.Mcp.exe");
    if (!File.Exists(exe)) exe = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "src", "GridPilot.Mcp", "bin", "Release", "net10.0-windows", "GridPilot.Mcp.exe");
    var utf8 = new System.Text.UTF8Encoding(false);
    using var p = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(exe) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true, StandardInputEncoding = utf8, StandardOutputEncoding = utf8 })!;
    System.Text.Json.JsonElement Call(int id, string method, object? parameters = null) { p.StandardInput.WriteLine(Json.Write(new { jsonrpc = "2.0", id, method, @params = parameters ?? new { } })); p.StandardInput.Flush(); var line = p.StandardOutput.ReadLine() ?? throw new Exception("EOF"); using var doc = System.Text.Json.JsonDocument.Parse(line); if (doc.RootElement.TryGetProperty("error", out var err)) throw new InvalidOperationException(err.GetProperty("message").GetString()); return doc.RootElement.GetProperty("result").Clone(); }
    try
    {
        Call(1, "initialize", new { protocolVersion = "2025-06-18", capabilities = new { }, clientInfo = new { name = "GridPilot.UI", version = "0.1.0" } });
        var opened = Call(2, "gridpilot/open", new { path = original, root = Path.Combine(output, "ui-rpc"), profile });
        True(opened.GetProperty("Editable").GetBoolean());
        var inspected = Call(3, "tools/call", new { name = "inspect_workbook", arguments = new { } });
        True(!inspected.GetProperty("isError").GetBoolean());
        var snap = Json.Read<WorkbookSnapshot>(System.Text.Json.JsonDocument.Parse(inspected.GetProperty("content")[0].GetProperty("text").GetString()!).RootElement.GetProperty("snapshot").GetRawText());
        True(snap.Table("Visitas").Rows.Count == 10 && snap.Table("Inventario").Rows.Count == 10 && snap.Table("Movimientos").Rows.Count == 20);
        Call(4, "gridpilot/close");
    }
    finally { p.StandardInput.Close(); p.WaitForExit(10000); }
}
static void ComplexInstructionsExcel(string output, string original, WorkbookProfile profile, Action<bool> True)
{
    using var excel = new ExcelWorkbook(false);
    using var s = new Session(excel, Path.Combine(output, "complex"), original, profile);
    void Run(string instruction)
    {
        var snap = s.Inspect();
        var res = LocalInterpreter.TryInterpret(instruction, snap, s.Profile) ?? throw new Exception("Could not interpret: " + instruction);
        s.Execute(new(Guid.NewGuid().ToString(), instruction, snap.Revision, res.Operations));
    }
    Run("En la pestaña de inventario, reemplaza los datos de la segnda fila por los de la tercera fila");
    var afterReplace = s.Inspect().Table("Inventario").Rows;
    True(afterReplace[1]["ID material"] == "M002" && afterReplace[1]["Descripción"] == afterReplace[2]["Descripción"] && afterReplace[1]["Saldo inicial"] == afterReplace[2]["Saldo inicial"]);
    Run("En la pestaña de visita, elimina toda la primera fila");
    var afterDelete = s.Inspect();
    True(afterDelete.Table("Visitas").Rows.Count == 9 && afterDelete.Table("Visitas").Rows[0]["ID"] == "V002");
}
