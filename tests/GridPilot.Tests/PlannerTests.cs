using System.Net;
using System.Text.Json;
using GridPilot.Core;
using GridPilot.Providers;

namespace GridPilot.Tests;

public static class PlannerTests
{
    public static void Run(Action<string, Action> check, WorkbookProfile profile)
    {
        void Assert(bool value) { if (!value) throw new Exception("Provider assertion failed."); }
        void Reject(Action action) { try { action(); } catch (InvalidOperationException) { return; } throw new Exception("Should have rejected operation."); }
        var settings = new Settings(Model: "gemini-3.8-flash", Provider: "gemini", CloudEnabled: true);
        object operation = new { kind = "SetValue", table = "Movimientos", id = "MOV011", field = "Cantidad", value = "3", values = Array.Empty<object>() };
        string Plan(string clarification = "") => Reply(new[] { new { functionCall = new { name = "propose_plan", args = new { clarification, operations = new[] { operation } } } } });
        Func<string, string, Task<object>> noRead = (_, _) => throw new Exception("Unexpected read call");
        check("Gemini: cloud disabled makes no network requests", () => { using var h = new ScriptedHttp(); var p = PlannerFactory.Create(settings with { CloudEnabled = false }, "fake", h); Reject(() => p.Interpret("test", profile, noRead, default).GetAwaiter().GetResult()); Assert(h.Requests.Count == 0); });
        check("Legacy configuration preserves Anthropic", () => Assert(Json.Read<Settings>("{\"Model\":\"claude-sonnet-5-5\"}").Provider == "anthropic"));
        check("Unknown provider does not use fallback", () => Reject(() => PlannerFactory.Create(settings with { Provider = "other" }, "fake")));
        check("Gemini: read, thought signature, call ID, and typed plan", () =>
        {
            using var h = new ScriptedHttp(Reply(new[] { new { thoughtSignature = "opaque-test", functionCall = new { id = "call-1", name = "read_records", args = new { table = "Movimientos", query = "MOV011" } } } }), Plan());
            var p = PlannerFactory.Create(settings, "fake", h); int reads = 0;
            var result = p.Interpret("Fix MOV011 to 3", profile, (table, query) => { Assert(table == "Movimientos" && query == "MOV011"); reads++; return Task.FromResult<object>(new { rows = new[] { new { Id = "MOV011" } }, truncated = false }); }, default).GetAwaiter().GetResult();
            Assert(reads == 1 && result.Operations.Single().Kind == OperationKind.SetValue && p.LastUsage?.Calls == 2);
            Assert(p.LastUsage!.OutputTokens == 30); // candidates + thinking on both requests
            Assert(h.Requests[1].Contains("opaque-test") && h.Requests[1].Contains("call-1"));
            using var body = JsonDocument.Parse(h.Requests[0]);
            Assert(body.RootElement.GetProperty("tools")[0].GetProperty("functionDeclarations").GetArrayLength() == 2);
            Assert(h.Paths.All(u => u == "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent"));
        });
        check("Gemini: clarification discards writes", () => { using var h = new ScriptedHttp(Plan("Which visit?")); var p = PlannerFactory.Create(settings, "fake", h); Assert(p.Interpret("fix", profile, noRead, default).Result.Operations.Count == 0); });
        foreach (var response in new[] { "{}", Reply(new[] { new { text = "no tools" } }), Reply(new[] { new { functionCall = new { name = "export", args = new { } } } }), Reply(new[] { new { functionCall = new { name = "propose_plan", args = new { } } } }), Reply(new[] { new { text = "truncated" } }, "MAX_TOKENS"), Reply(new[] { new { text = "blocked" } }, "SAFETY") })
            check("Gemini: invalid response produces no plan " + response[..Math.Min(35, response.Length)], () => { using var h = new ScriptedHttp(response); var p = PlannerFactory.Create(settings, "fake", h); Reject(() => p.Interpret("test", profile, noRead, default).GetAwaiter().GetResult()); });
        check("Gemini: HTTP 429 does not retry", () => { using var h = new FailingHttp(); var p = PlannerFactory.Create(settings, "fake", h); Reject(() => p.Interpret("test", profile, noRead, default).GetAwaiter().GetResult()); Assert(h.Calls == 1); });
        check("Gemini: call limit enforced", () => { using var h = new ScriptedHttp(Reply(new[] { new { functionCall = new { name = "read_records", args = new { table = "Movimientos", query = "MOV011" } } } })); var p = PlannerFactory.Create(settings with { MaxCalls = 1 }, "fake", h); Reject(() => p.Interpret("test", profile, (_, _) => Task.FromResult<object>(new { }), default).GetAwaiter().GetResult()); Assert(h.Requests.Count == 1); });
        check("Gemini: token limit includes thinking tokens", () => { using var h = new ScriptedHttp(Reply(new[] { new { functionCall = new { name = "read_records", args = new { table = "Movimientos", query = "MOV011" } } } })); var p = PlannerFactory.Create(settings with { MaxTotalTokens = 20 }, "fake", h); Reject(() => p.Interpret("test", profile, (_, _) => Task.FromResult<object>(new { }), default).GetAwaiter().GetResult()); Assert(h.Requests.Count == 1); });
        check("Anthropic: shared contract continues interpreting plans", () => { using var h = new ScriptedHttp(Json.Write(new { usage = new { input_tokens = 10, output_tokens = 5 }, stop_reason = "tool_use", content = new[] { new { type = "tool_use", id = "x", name = "propose_plan", input = new { clarification = "", operations = new[] { operation } } } } })); var p = PlannerFactory.Create(new(CloudEnabled: true), "fake", h); Assert(p.Interpret("test", profile, noRead, default).Result.Operations.Single().Value == "3"); });
        check("Local interpreter (ES & EN): replaces row 2 data with row 3", () =>
        {
            var snap = MultiRowFixture(profile);
            var resEs = LocalInterpreter.TryInterpret("En la pestaña de inventario, reemplaza los datos de la segnda fila por los de la tercera fila", snap, profile);
            var resEn = LocalInterpreter.TryInterpret("In the inventory tab, replace the data of the second row with the third row", snap, profile);
            Assert(resEs is not null && resEs.Operations.Count == 5 && resEs.Operations.All(o => o.Kind == OperationKind.SetValue && o.Table == "Inventario" && o.Id == "M002"));
            Assert(resEn is not null && resEn.Operations.Count == 5 && resEn.Operations.All(o => o.Kind == OperationKind.SetValue && o.Table == "Inventario" && o.Id == "M002"));
            Assert(resEs!.Operations.Single(o => o.Field == "Descripcion").Value == "Canaleta" && resEs.Operations.Single(o => o.Field == "SaldoInicial").Value == "22" && resEs.Operations.Single(o => o.Field == "CostoReferencia").Value == "15");
        });
        check("Local interpreter (ES & EN): swaps two rows preserving IDs", () =>
        {
            var snap = MultiRowFixture(profile);
            var resEs = LocalInterpreter.TryInterpret("En la pestaña de inventario, intercambia la primera fila con la tercera fila", snap, profile);
            var resEn = LocalInterpreter.TryInterpret("In the inventory tab, swap the first row with the third row", snap, profile);
            Assert(resEs is not null && resEs.Operations.Count == 10 && resEs.Operations.Count(o => o.Id == "M001") == 5 && resEs.Operations.Count(o => o.Id == "M003") == 5);
            Assert(resEn is not null && resEn.Operations.Count == 10 && resEn.Operations.Count(o => o.Id == "M001") == 5 && resEn.Operations.Count(o => o.Id == "M003") == 5);
        });
        check("Local interpreter (ES & EN): deletes first visits row including dependent movements", () =>
        {
            var snap = MultiRowFixture(profile);
            var resEs = LocalInterpreter.TryInterpret("en la pestaña de visita, elimina toda la primera fila", snap, profile);
            var resEn = LocalInterpreter.TryInterpret("in the visits tab, delete the entire first row", snap, profile);
            Assert(resEs is not null && resEs.Operations.Count == 2 && resEs.Operations[0] == new Operation(OperationKind.DeleteRecord, "Movimientos", "MOV001") && resEs.Operations[1] == new Operation(OperationKind.DeleteRecord, "Visitas", "V001"));
            Assert(resEn is not null && resEn.Operations.Count == 2 && resEn.Operations[0] == new Operation(OperationKind.DeleteRecord, "Movimientos", "MOV001") && resEn.Operations[1] == new Operation(OperationKind.DeleteRecord, "Visitas", "V001"));
        });
        check("Local interpreter (ES & EN): relative arithmetic and compound clauses", () =>
        {
            var snap = MultiRowFixture(profile);
            var resEs = LocalInterpreter.TryInterpret("En la pestaña de inventario, cambia el saldo inicial de la segunda fila de 21 a 25 y luego incrementa el mínimo de la primera fila en 5", snap, profile);
            var resEn = LocalInterpreter.TryInterpret("In the inventory tab, change the initial balance of the second row from 21 to 25 and then increase the minimum of the first row by 5", snap, profile);
            Assert(resEs is not null && resEs.Operations.Count == 2 && resEs.Operations[0].Id == "M002" && resEs.Operations[0].Value == "25" && resEs.Operations[1].Id == "M001" && resEs.Operations[1].Field == "Minimo" && resEs.Operations[1].Value == "15");
            Assert(resEn is not null && resEn.Operations.Count == 2 && resEn.Operations[0].Id == "M002" && resEn.Operations[0].Value == "25" && resEn.Operations[1].Id == "M001" && resEn.Operations[1].Field == "Minimo" && resEn.Operations[1].Value == "15");
        });
        check("Local interpreter: forwards ambiguous instructions to cloud planner", () =>
        {
            var snap = MultiRowFixture(profile);
            Assert(LocalInterpreter.TryInterpret("Cambia la visita de Ana.", snap, profile) is null);
            Assert(LocalInterpreter.TryInterpret("Pon la visita V001 para el 03/04.", snap, profile) is null);
        });
        check("Gemini Audio (voice): transcribes in-memory WAV with inlineData audio/wav and records token usage", () =>
        {
            using var h = new ScriptedHttp(Reply(new[] { new { text = "In the inventory tab, replace the data of the second row with the third row" } }));
            using var wav = new MemoryStream(new byte[100]); // header + sample bytes
            using var voice = new SpeechInput(settings, "fake-gemini-key", h, _ => new MemoryStream(new byte[100]));
            voice.Start(10).GetAwaiter().GetResult();
            var text = voice.Stop().GetAwaiter().GetResult();
            Assert(text.StartsWith("In the inventory tab") && voice.LastUsage?.Calls == 1 && voice.LastUsage?.OutputTokens == 15);
            Assert(h.Requests[0].Contains("\"mimeType\":\"audio/wav\"") && h.Requests[0].Contains("inlineData"));
        });
        check("Gemini Audio (voice): blocked when cloud usage is disabled", () =>
        {
            Reject(() => new SpeechInput(settings with { CloudEnabled = false }, "fake-gemini-key"));
        });
    }
    static WorkbookSnapshot MultiRowFixture(WorkbookProfile profile)
    {
        var baseSnap = MemoryWorkbook.Fixture(profile);
        var inv = baseSnap.Table("Inventario");
        inv.Rows.Add(new(inv.Rows[0]) { ["ID material"] = "M002", ["Descripción"] = "Conector RJ45", ["Saldo inicial"] = "21", ["Mínimo"] = "10", ["Costo referencia"] = "10" });
        inv.Rows.Add(new(inv.Rows[0]) { ["ID material"] = "M003", ["Descripción"] = "Canaleta", ["Saldo inicial"] = "22", ["Mínimo"] = "10", ["Costo referencia"] = "15" });
        var vis = baseSnap.Table("Visitas");
        vis.Rows.Add(new(vis.Rows[0]) { ["ID"] = "V002", ["Cliente"] = "Cliente 2", ["Sitio"] = "Sitio 2" });
        return baseSnap;
    }
    static string Reply(object parts, string finish = "STOP") => Json.Write(new { candidates = new[] { new { finishReason = finish, content = new { role = "model", parts } } }, usageMetadata = new { promptTokenCount = 10, candidatesTokenCount = 5, thoughtsTokenCount = 10 } });
}

sealed class ScriptedHttp(params string[] responses) : HttpMessageHandler
{
    public List<string> Requests { get; } = [];
    public List<string> Paths { get; } = [];
    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
    {
        Paths.Add(request.RequestUri!.AbsoluteUri); Requests.Add(await request.Content!.ReadAsStringAsync(token));
        if (Requests.Count > responses.Length) throw new Exception("Unexpected HTTP call.");
        return new(HttpStatusCode.OK) { Content = new StringContent(responses[Requests.Count - 1]) };
    }
}
