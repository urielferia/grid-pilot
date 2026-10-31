using System.Diagnostics;
using System.Net.Http.Json;
using System.Text.Json;
using GridPilot.Core;

namespace GridPilot.Providers;

public sealed class GeminiPlanner(Settings settings, string key, HttpMessageHandler? handler = null) : IPlanner
{
    public Usage? LastUsage { get; private set; }
    bool IsEs => string.Equals(settings.Language, "es", StringComparison.OrdinalIgnoreCase);

    public async Task<Interpretation> Interpret(string instruction, WorkbookProfile profile, Func<string, string, Task<object>> read, CancellationToken token)
    {
        LastUsage = null;
        if (!settings.CloudEnabled) throw new InvalidOperationException(IsEs ? "Active el uso de nube en Ajustes para interpretar texto." : "Enable cloud usage in Settings to interpret text.");
        if (string.IsNullOrWhiteSpace(instruction) || instruction.Length > 4000) throw new InvalidOperationException(IsEs ? "Instrucción vacía o demasiado larga." : "Instruction is empty or too long.");
        if (!System.Text.RegularExpressions.Regex.IsMatch(settings.Model, @"^gemini-[a-zA-Z0-9.\-]+$")) throw new InvalidOperationException(IsEs ? "Configure un identificador de modelo Gemini válido." : "Configure a valid Gemini model identifier.");
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(token);
        timeout.CancelAfter(TimeSpan.FromSeconds(Math.Clamp(settings.TimeoutSeconds, 5, 180)));
        using var http = handler is null ? new HttpClient() : new HttpClient(handler, false);
        http.DefaultRequestHeaders.Add("x-goog-api-key", key);
        var contents = new List<object> { new { role = "user", parts = new[] { new { text = instruction } } } };
        var clock = Stopwatch.StartNew(); int input = 0, output = 0, calls = 0;
        // The JSON schema is shared with Anthropic; the transport uses Gemini's native tool protocol.
        object Tool(string name, string description, object schema) => new { name, description, parametersJsonSchema = schema };
        try
        {
            while (calls < Math.Clamp(settings.MaxCalls, 1, 10))
            {
                if (input + output >= settings.MaxTotalTokens) throw new InvalidOperationException(IsEs ? "Se alcanzó el límite de tokens." : "Token limit reached.");
                var body = new
                {
                    systemInstruction = new { parts = new[] { new { text = PlannerContract.Prompt(profile, settings.Language) } } },
                    contents,
                    tools = new[]
                    {
                        new
                        {
                            functionDeclarations = new[]
                            {
                                Tool("read_records", "Searches or lists records from a table (max 20). Pass query='*' to view rows in order with their _row number, a row number ('1', 'row 2'), a range ('1-5'), or text/ID.", PlannerContract.Obj(new() { ["table"] = PlannerContract.Str(), ["query"] = PlannerContract.Str() })),
                                Tool("propose_plan", "Finishes with the list of typed operations or a clarification if the instruction is ambiguous.", PlannerContract.PlanSchema())
                            }
                        }
                    },
                    toolConfig = new { functionCallingConfig = new { mode = "ANY", allowedFunctionNames = new[] { "read_records", "propose_plan" } } },
                    generationConfig = new { maxOutputTokens = Math.Min(Math.Clamp(settings.MaxOutputTokens, 256, 6000), settings.MaxTotalTokens - input - output) }
                };
                calls++;
                using var response = await http.PostAsJsonAsync($"https://generativelanguage.googleapis.com/v1beta/models/{settings.Model}:generateContent", body, Json.Options, timeout.Token);
                if (!response.IsSuccessStatusCode)
                    throw new InvalidOperationException(IsEs
                        ? $"Gemini respondió {(int)response.StatusCode}. Compruebe credencial, cuota y disponibilidad del modelo {settings.Model}. No se cambió de modelo."
                        : $"Gemini responded {(int)response.StatusCode}. Check your API key, quota, and availability of model {settings.Model}. Model was not switched.");
                using var doc = JsonDocument.Parse(await response.Content.ReadAsStringAsync(timeout.Token)); var root = doc.RootElement;
                if (root.TryGetProperty("usageMetadata", out var usage))
                {
                    int Count(string name) => usage.TryGetProperty(name, out var n) ? n.GetInt32() : 0;
                    input += Count("promptTokenCount"); output += Count("candidatesTokenCount") + Count("thoughtsTokenCount");
                }
                if (!root.TryGetProperty("candidates", out var candidates) || candidates.GetArrayLength() != 1)
                    throw new InvalidOperationException(IsEs ? "Gemini no devolvió un plan; respuesta bloqueada o vacía." : "Gemini did not return a plan; response was blocked or empty.");
                var candidate = candidates[0];
                if (!candidate.TryGetProperty("finishReason", out var reason) || reason.GetString() != "STOP")
                    throw new InvalidOperationException(IsEs ? "Gemini devolvió una respuesta incompleta o bloqueada. No se ejecutó ningún cambio." : "Gemini returned an incomplete or blocked response. No changes were executed.");
                var content = candidate.GetProperty("content").Clone();
                // Preserve opaque thought signatures and call IDs verbatim for subsequent turns.
                contents.Add(content);
                var functions = content.GetProperty("parts").EnumerateArray().Where(p => p.TryGetProperty("functionCall", out _)).Select(p => p.GetProperty("functionCall")).ToList();
                if (functions.Count > 20) throw new InvalidOperationException(IsEs ? "Respuesta con demasiadas llamadas." : "Response contained too many tool calls.");
                if (functions.Count == 0)
                {
                    var text = string.Concat(content.GetProperty("parts").EnumerateArray().Where(p => p.TryGetProperty("text", out _)).Select(p => p.GetProperty("text").GetString())).Trim();
                    if (text.StartsWith('{')) { try { using var embedded = JsonDocument.Parse(text); return PlannerContract.Parse(embedded.RootElement); } catch (JsonException) { } }
                    throw new InvalidOperationException(IsEs ? "Respuesta sin herramientas." : "Response contained no tool calls.");
                }
                if (functions.Any(f => f.GetProperty("name").GetString() == "propose_plan"))
                {
                    if (functions.Count != 1) throw new InvalidOperationException(IsEs ? "El plan final debe venir separado de las lecturas." : "The final plan must be separate from read calls.");
                    return PlannerContract.Parse(functions[0].GetProperty("args"));
                }
                var results = new List<object>();
                foreach (var function in functions)
                {
                    if (function.GetProperty("name").GetString() != "read_records") throw new InvalidOperationException(IsEs ? "Herramienta desconocida." : "Unknown tool.");
                    var args = function.GetProperty("args");
                    var table = args.GetProperty("table").GetString()!; var query = args.TryGetProperty("query", out var qProp) ? qProp.GetString() ?? "*" : "*";
                    if (string.IsNullOrWhiteSpace(query)) query = "*";
                    profile.Table(table);
                    if (query.Length > 100) throw new InvalidOperationException(IsEs ? "Use una búsqueda de hasta 100 caracteres." : "Use a search query up to 100 characters.");
                    timeout.Token.ThrowIfCancellationRequested();
                    var result = await read(table, query).WaitAsync(timeout.Token);
                    var answer = new Dictionary<string, object> { { "name", "read_records" }, { "response", new { result } } };
                    if (function.TryGetProperty("id", out var id)) answer["id"] = id.GetString()!;
                    results.Add(new { functionResponse = answer });
                }
                contents.Add(new { role = "user", parts = results });
            }
            throw new InvalidOperationException(IsEs ? "Se alcanzó el límite de llamadas por instrucción." : "Maximum calls per instruction reached.");
        }
        finally { LastUsage = new(settings.Model, calls, input, output, clock.ElapsedMilliseconds); }
    }
}

public static class PlannerFactory
{
    public static IPlanner Create(Settings settings, string key, HttpMessageHandler? handler = null) => settings.Provider switch
    {
        "gemini" => new GeminiPlanner(settings, key, handler),
        "anthropic" => new AnthropicPlanner(settings, key, handler),
        _ => throw new InvalidOperationException("Unsupported provider. Select Gemini or Anthropic.")
    };
}
