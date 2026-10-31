using System.Diagnostics;
using System.Net.Http.Json;
using System.Text.Json;
using GridPilot.Core;
namespace GridPilot.Providers;

public interface IPlanner
{
    Usage? LastUsage => null;
    Task<Interpretation> Interpret(string instruction, WorkbookProfile profile, Func<string, string, Task<object>> read, CancellationToken token);
}
public sealed record Usage(string Model, int Calls, int InputTokens, int OutputTokens, long Milliseconds);
public sealed class AnthropicPlanner(Settings settings, string key, HttpMessageHandler? handler = null) : IPlanner
{
    public Usage? LastUsage { get; private set; }
    bool IsEs => string.Equals(settings.Language, "es", StringComparison.OrdinalIgnoreCase);
    static object Tool(string name, string description, object input_schema) => new { name, description, input_schema, strict = true };
    public async Task<Interpretation> Interpret(string instruction, WorkbookProfile profile, Func<string, string, Task<object>> read, CancellationToken token)
    {
        if (!settings.CloudEnabled) throw new InvalidOperationException(IsEs ? "Active el uso de nube en Ajustes para interpretar texto." : "Enable cloud usage in Settings to interpret text.");
        if (instruction.Length > 4000 || string.IsNullOrWhiteSpace(instruction)) throw new InvalidOperationException(IsEs ? "Instrucción vacía o demasiado larga." : "Instruction is empty or too long.");
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(token); timeout.CancelAfter(TimeSpan.FromSeconds(Math.Clamp(settings.TimeoutSeconds, 5, 180)));
        using var http = handler is null ? new HttpClient() : new HttpClient(handler, false); http.DefaultRequestHeaders.Add("x-api-key", key); http.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");
        var messages = new List<object> { new { role = "user", content = instruction } };
        var clock = Stopwatch.StartNew(); int input = 0, output = 0, calls = 0;
        try
        {
            for (; calls < Math.Clamp(settings.MaxCalls, 1, 10);)
            {
                if (input + output >= settings.MaxTotalTokens) throw new InvalidOperationException(IsEs ? "Se alcanzó el límite de tokens." : "Token limit reached.");
                var body = new
                {
                    model = settings.Model,
                    max_tokens = Math.Min(Math.Clamp(settings.MaxOutputTokens, 256, 6000), settings.MaxTotalTokens - input - output),
                    system = PlannerContract.Prompt(profile, settings.Language),
                    messages,
                    tools = new[]
                    {
                        Tool("read_records", "Searches or lists records from a table (max 20). Pass query='*' to view rows in order with their _row number, a row number ('1', 'row 2'), a range ('1-5'), or text/ID.", PlannerContract.Obj(new() { ["table"] = PlannerContract.Str(), ["query"] = PlannerContract.Str() })),
                        Tool("propose_plan", "Finishes with the list of typed operations or a clarification if the instruction is ambiguous.", PlannerContract.PlanSchema())
                    }
                };
                calls++;
                using var response = await http.PostAsJsonAsync("https://api.anthropic.com/v1/messages", body, Json.Options, timeout.Token);
                if (!response.IsSuccessStatusCode)
                    throw new InvalidOperationException(IsEs
                        ? $"Anthropic respondió {(int)response.StatusCode}. Compruebe credencial, cuota y disponibilidad del modelo {settings.Model}. No se cambió de modelo."
                        : $"Anthropic responded {(int)response.StatusCode}. Check your API key, quota, and availability of model {settings.Model}. Model was not switched.");
                using var doc = JsonDocument.Parse(await response.Content.ReadAsStringAsync(timeout.Token)); var root = doc.RootElement;
                input += root.GetProperty("usage").GetProperty("input_tokens").GetInt32(); output += root.GetProperty("usage").GetProperty("output_tokens").GetInt32();
                if (root.GetProperty("stop_reason").GetString() == "max_tokens") throw new InvalidOperationException(IsEs ? "Respuesta truncada; reduzca la instrucción." : "Response truncated; shorten the instruction.");
                var content = root.GetProperty("content").Clone(); messages.Add(new { role = "assistant", content }); var results = new List<object>();
                foreach (var block in content.EnumerateArray().Where(b => b.GetProperty("type").GetString() == "tool_use"))
                {
                    var name = block.GetProperty("name").GetString(); var data = block.GetProperty("input");
                    if (name == "propose_plan")
                    {
                        return PlannerContract.Parse(data);
                    }
                    if (name != "read_records") throw new InvalidOperationException(IsEs ? "Herramienta desconocida." : "Unknown tool.");
                    var q = data.TryGetProperty("query", out var qProp) ? qProp.GetString() ?? "*" : "*";
                    if (string.IsNullOrWhiteSpace(q)) q = "*";
                    var result = await read(data.GetProperty("table").GetString()!, q);
                    results.Add(new { type = "tool_result", tool_use_id = block.GetProperty("id").GetString(), content = Json.Write(result) });
                }
                if (results.Count == 0) throw new InvalidOperationException(IsEs ? "El modelo no produjo un plan estructurado." : "The model did not produce a structured plan.");
                messages.Add(new { role = "user", content = results });
            }
            throw new InvalidOperationException(IsEs ? "Se alcanzó el límite de llamadas por instrucción." : "Maximum calls per instruction reached.");
        }
        finally { LastUsage = new(settings.Model, calls, input, output, clock.ElapsedMilliseconds); }
    }
}
public sealed class FakePlanner(Interpretation result) : IPlanner
{
    public Task<Interpretation> Interpret(string instruction, WorkbookProfile profile, Func<string, string, Task<object>> read, CancellationToken token) => Task.FromResult(result);
}
