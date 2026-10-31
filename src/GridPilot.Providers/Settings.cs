using System.Security.Cryptography;
using System.Text;
using GridPilot.Core;
namespace GridPilot.Providers;

public sealed record Settings(string Model = "claude-sonnet-5-5", string SpeechRegion = "", bool CloudEnabled = false, int MaxCalls = 6, int MaxOutputTokens = 3000, int MaxTotalTokens = 16000, int TimeoutSeconds = 90, int MaxVoiceSeconds = 30, string Provider = "anthropic", string Language = "en");
public sealed class CredentialStore(string root, string language = "en")
{
    bool IsEs => string.Equals(language, "es", StringComparison.OrdinalIgnoreCase);
    string PathFor(string name)
    {
        if (name is not ("anthropic" or "gemini" or "speech"))
            throw new ArgumentException(IsEs ? "Proveedor inválido." : "Invalid provider.");
        Directory.CreateDirectory(root);
        return Path.Combine(root, name + ".dpapi");
    }
    public void Save(string name, string value) => File.WriteAllBytes(PathFor(name), ProtectedData.Protect(Encoding.UTF8.GetBytes(value), null, DataProtectionScope.CurrentUser));
    public string Read(string name)
    {
        var path = PathFor(name);
        if (!File.Exists(path))
            throw new InvalidOperationException(IsEs ? "Configure la credencial de " + name + " en Ajustes." : "Configure the " + name + " credential in Settings.");
        return Encoding.UTF8.GetString(ProtectedData.Unprotect(File.ReadAllBytes(path), null, DataProtectionScope.CurrentUser));
    }
}
