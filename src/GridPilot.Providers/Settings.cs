using System.Security.Cryptography;
using System.Text;
using GridPilot.Core;
namespace GridPilot.Providers;

public sealed record Settings(string Model="claude-sonnet-5-5",string SpeechRegion="",bool CloudEnabled=false,int MaxCalls=6,int MaxOutputTokens=3000,int MaxTotalTokens=16000,int TimeoutSeconds=90,int MaxVoiceSeconds=30);
public sealed class CredentialStore(string root)
{
    string PathFor(string name) {if(name is not ("anthropic" or "speech"))throw new ArgumentException("Proveedor inválido.");Directory.CreateDirectory(root);return Path.Combine(root,name+".dpapi");}
    public void Save(string name,string value)=>File.WriteAllBytes(PathFor(name),ProtectedData.Protect(Encoding.UTF8.GetBytes(value),null,DataProtectionScope.CurrentUser));
    public string Read(string name) {var path=PathFor(name);if(!File.Exists(path))throw new InvalidOperationException("Configure la credencial de " + name + " en Ajustes.");return Encoding.UTF8.GetString(ProtectedData.Unprotect(File.ReadAllBytes(path),null,DataProtectionScope.CurrentUser));}
}
