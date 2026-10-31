using System.Diagnostics;
using GridPilot.Core;
using GridPilot.Excel;

/// <summary>
/// Records which server process and dedicated Excel instance own a session, so that explicit
/// recovery after a server crash can close the orphaned Excel instance GridPilot itself created
/// (and which still locks draft.xlsx). Excel instances not created by GridPilot are never touched.
/// </summary>
static class Owner
{
    sealed record Info(int ServerPid, DateTime ServerStart, int ExcelPid, DateTime ExcelStart);
    const string FileName = "owner.json";

    static DateTime? StartOf(int pid, string name)
    {
        if (pid <= 0) return null;
        try
        {
            using var process = Process.GetProcessById(pid);
            if (!process.ProcessName.Equals(name, StringComparison.OrdinalIgnoreCase) || process.HasExited) return null;
            return process.StartTime.ToUniversalTime();
        }
        catch (Exception e) when (e is ArgumentException or InvalidOperationException or System.ComponentModel.Win32Exception) { return null; }
    }

    public static void Record(string directory, ExcelWorkbook excel)
    {
        using var self = Process.GetCurrentProcess();
        var info = new Info(self.Id, self.StartTime.ToUniversalTime(), excel.ProcessId, StartOf(excel.ProcessId, "EXCEL") ?? DateTime.MinValue);
        File.WriteAllText(Path.Combine(directory, FileName), Json.Write(info));
    }

    public static void ReleaseOrphan(string directory)
    {
        var file = Path.Combine(Path.GetFullPath(directory), FileName);
        if (!File.Exists(file)) return;
        Info info;
        try { info = Json.Read<Info>(File.ReadAllText(file)); } catch (Exception e) when (e is System.Text.Json.JsonException or InvalidDataException) { return; }
        using var self = Process.GetCurrentProcess();
        if (info.ServerPid == self.Id) return; // Already closed by this server before recovering.
        var serverName = Path.GetFileNameWithoutExtension(Environment.ProcessPath ?? "GridPilot.Mcp");
        var serverStart = StartOf(info.ServerPid, serverName);
        if (serverStart is not null && Math.Abs((serverStart.Value - info.ServerStart).TotalSeconds) < 1)
            throw new InvalidOperationException("The session is still open in another GridPilot window. Close it before recovering.");
        var excelStart = StartOf(info.ExcelPid, "EXCEL");
        if (excelStart is null || Math.Abs((excelStart.Value - info.ExcelStart).TotalSeconds) >= 1) return;
        // The owning server ended unexpectedly and its dedicated Excel is still running: close it.
        using var orphan = Process.GetProcessById(info.ExcelPid);
        orphan.Kill();
        if (!orphan.WaitForExit(15000)) throw new InvalidOperationException("Could not close the orphaned Excel instance for this session.");
    }
}
