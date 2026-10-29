using System.Diagnostics;
using System.IO;
using System.Text.Json;
using GridPilot.Core;
namespace GridPilot.App;
public sealed class RpcClient : IDisposable
{
    readonly Process process;readonly SemaphoreSlim gate=new(1);int sequence;
    public RpcClient()
    {
        var executable=Path.Combine(AppContext.BaseDirectory,"server","GridPilot.Mcp.exe");
        if(!File.Exists(executable))throw new InvalidOperationException("Falta el servidor MCP. Ejecute scripts/build.ps1.");
        process=Process.Start(new ProcessStartInfo(executable){UseShellExecute=false,RedirectStandardInput=true,RedirectStandardOutput=true,RedirectStandardError=true,CreateNoWindow=true})!;
        process.ErrorDataReceived+=(_,_)=>{};process.BeginErrorReadLine();
    }
    public async Task<JsonElement> Call(string method,object? parameters=null)
    {
        await gate.WaitAsync();try{
            if(process.HasExited)throw new InvalidOperationException("El servidor se cerró. Reinicie y recupere la sesión.");
            var id=++sequence;await process.StandardInput.WriteLineAsync(Json.Write(new{jsonrpc="2.0",id,method,@params=parameters??new{}}));await process.StandardInput.FlushAsync();
            var line=await process.StandardOutput.ReadLineAsync()??throw new InvalidOperationException("El servidor se cerró inesperadamente.");
            using var doc=JsonDocument.Parse(line);var root=doc.RootElement;
            if(root.TryGetProperty("error",out var error))throw new InvalidOperationException(error.GetProperty("message").GetString());
            if(root.GetProperty("id").GetInt32()!=id)throw new InvalidOperationException("Respuesta fuera de secuencia.");
            return root.GetProperty("result").Clone();
        }finally{gate.Release();}
    }
    public async Task<JsonElement> Tool(string name,object arguments)
    {
        var r=await Call("tools/call",new{name,arguments});var text=r.GetProperty("content")[0].GetProperty("text").GetString()!;
        if(r.GetProperty("isError").GetBoolean())throw new InvalidOperationException(text);return JsonSerializer.Deserialize<JsonElement>(text);
    }
    public void Dispose(){process.StandardInput.Close();if(process.WaitForExit(8000))process.Dispose();gate.Dispose();}
}
