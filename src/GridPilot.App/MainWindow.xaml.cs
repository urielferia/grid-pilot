using System.ComponentModel;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using GridPilot.Core;
using GridPilot.Providers;
using Microsoft.Win32;

namespace GridPilot.App;
public partial class MainWindow : Window
{
    RpcClient? rpc; WorkbookSnapshot? snapshot; WorkbookProfile? profile; bool busy, recording, initializingLanguage; SpeechInput? speech;
    readonly string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GridPilot");
    Settings settings = new(Model: "gemini-3.8-flash", Provider: "gemini", Language: "en");
    readonly List<string> history = [];
    string? selectedProfile;
    bool IsEs => string.Equals(settings.Language, "es", StringComparison.OrdinalIgnoreCase);
    string T(string en, string es) => IsEs ? es : en;
    CredentialStore Credentials => new(Path.Combine(root, "credentials"), settings.Language);

    public MainWindow()
    {
        InitializeComponent();
        MaxHeight = SystemParameters.WorkArea.Height;
        MinHeight = Math.Min(MinHeight, MaxHeight);
        Height = Math.Min(Height, MaxHeight);
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        Directory.CreateDirectory(root);
        var file = Path.Combine(root, "settings.json");
        if (File.Exists(file)) settings = Json.Read<Settings>(File.ReadAllText(file));
        if (string.IsNullOrWhiteSpace(settings.Language)) settings = settings with { Language = "en" };
        ApplyLanguage(initial: true);
    }

    void ApplyLanguage(bool initial = false)
    {
        initializingLanguage = true;
        LanguageSelector.SelectedValue = IsEs ? "es" : "en";
        initializingLanguage = false;

        Title = T("GridPilot · Excel by voice and text", "GridPilot · Excel por voz y texto");
        LanguageLabel.Text = T("Language:", "Idioma:");
        SubtitleText.Text = T("Work on a visible draft. Review before exporting.", "Trabaja en un borrador visible. Revisa antes de exportar.");
        OpenBtn.Content = T("_Open workbook", "_Abrir libro");
        DemoBtn.Content = T("_Demo template", "Plantilla _demo");
        ProfileBtn.Content = T("_Profile…", "_Perfil…");
        RecoverBtn.Content = T("_Recover session", "_Recuperar sesión");
        SettingsBtn.Content = T("_Settings", "A_justes");
        ReviewTab.Header = T("Review", "Revisión");
        StructureTab.Header = T("Structure", "Estructura");
        BatchTab.Header = T("Local batch", "Lote local");
        BatchHint.Text = T("Deterministic mode: paste a JSON list of typed operations.", "Modo determinista: pega una lista JSON de operaciones tipadas.");
        ExecuteBatchBtn.Content = T("Run batch on draft", "Ejecutar lote en borrador");
        InstructionLabel.Text = T("Your instruction", "Tu instrucción");
        Voice.Content = recording ? T("Stop", "Detener") : T("Speak", "Hablar");
        SendBtn.Content = T("Send", "Enviar");
        RefreshBtn.Content = T("Refresh review", "Actualizar revisión");
        UndoBtn.Content = T("Undo last batch", "Deshacer último lote");
        ApproveBtn.Content = T("Approve and export…", "Aprobar y exportar…");
        DiscardBtn.Content = T("Discard session", "Descartar sesión");

        if (initial || Status.Text is "Select an .xlsx file to begin." or "Selecciona un archivo .xlsx para comenzar.")
            Status.Text = T("Select an .xlsx file to begin.", "Selecciona un archivo .xlsx para comenzar.");

        RenderSnapshotViews();
    }

    void ChangeLanguage(object sender, SelectionChangedEventArgs e)
    {
        if (initializingLanguage || LanguageSelector.SelectedValue is not string lang) return;
        settings = settings with { Language = lang };
        File.WriteAllText(Path.Combine(root, "settings.json"), Json.Write(settings));
        ApplyLanguage();
    }

    async Task Run(Func<Task> action)
    {
        if (busy || recording) return;
        busy = true; Controls.IsEnabled = false;
        try { await action(); }
        catch (Exception e) { Status.Text = e.Message; }
        finally { busy = false; Controls.IsEnabled = true; }
    }

    async Task<RpcClient> Client()
    {
        // If the local server ended (crash or Excel failure), start a new one so the user can recover the session without restarting the app.
        if (rpc is not null && !rpc.Alive) { rpc.Dispose(); rpc = null; snapshot = null; }
        if (rpc is null)
        {
            var created = new RpcClient();
            try { await created.Call("initialize", new { protocolVersion = "2025-06-18", capabilities = new { }, clientInfo = new { name = "GridPilot.UI", version = "0.1.0" } }); }
            catch { created.Dispose(); throw; }
            rpc = created;
        }
        return rpc;
    }

    async Task OpenPath(string path)
    {
        Status.Text = T("Opening copy in Excel…", "Abriendo copia en Excel…");
        var client = await Client();
        profile = Json.Read<WorkbookProfile>(File.ReadAllText(selectedProfile ?? Path.Combine(AppContext.BaseDirectory, "assets", "profile.json")));
        var result = await client.Call("gridpilot/open", new { path, root = Path.Combine(root, "sessions"), profile });
        history.Clear();
        await LoadSnapshot();
        Status.Text = result.GetProperty("Editable").GetBoolean()
            ? T("Draft ready. The original file remains untouched by writes.", "Borrador listo. El original permanece fuera de las escrituras.")
            : T("Workbook available for inspection only: incompatible profile.", "Libro disponible solo para inspección: perfil incompatible.");
    }

    async void Open(object sender, RoutedEventArgs e) => await Run(async () =>
    {
        var dialog = new OpenFileDialog { Filter = T("Excel Workbook|*.xlsx", "Libro Excel|*.xlsx") };
        if (dialog.ShowDialog() == true) await OpenPath(dialog.FileName);
    });

    async void Demo(object sender, RoutedEventArgs e) => await Run(async () =>
    {
        selectedProfile = null;
        await OpenPath(Path.Combine(AppContext.BaseDirectory, "assets", "GridPilot-Demo.xlsx"));
    });

    void SelectProfile(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = T("Profile for the next workbook", "Perfil para el próximo libro"),
            Filter = T("GridPilot Profile|*.json", "Perfil de GridPilot|*.json")
        };
        if (dialog.ShowDialog() != true) return;
        try
        {
            var candidate = Json.Read<WorkbookProfile>(File.ReadAllText(dialog.FileName));
            candidate.ValidateDefinition();
            selectedProfile = dialog.FileName;
            Status.Text = T(
                "Profile selected for the next workbook. Compatibility will be validated upon opening.",
                "Perfil seleccionado para el próximo libro. Se validará su compatibilidad al abrirlo.");
        }
        catch (Exception ex) { Status.Text = ex.Message; }
    }

    async void Recover(object sender, RoutedEventArgs e) => await Run(async () =>
    {
        var dialog = new OpenFileDialog
        {
            Title = T("Select session.db of the session to recover", "Selecciona session.db de la sesión a recuperar"),
            InitialDirectory = Path.Combine(root, "sessions"),
            Filter = T("GridPilot Session|session.db", "Sesión GridPilot|session.db")
        };
        if (dialog.ShowDialog() == true)
        {
            await (await Client()).Call("gridpilot/recover", new { directory = Path.GetDirectoryName(dialog.FileName) });
            history.Clear();
            await LoadSnapshot();
            Status.Text = T(
                "Session recovered from the last valid state. No pending writes were replayed.",
                "Sesión recuperada desde el último estado válido. No se repitieron escrituras.");
        }
    });

    async Task LoadSnapshot()
    {
        var r = await (await Client()).Tool("inspect_workbook", new { });
        snapshot = Json.Read<WorkbookSnapshot>(r.GetProperty("snapshot").GetRawText());
        profile = Json.Read<WorkbookProfile>(r.GetProperty("profile").GetRawText());
        RenderSnapshotViews();
    }

    void RenderSnapshotViews()
    {
        if (snapshot is null || profile is null) return;
        Structure.Text = string.Join("\n\n", snapshot.Tables.Select(t =>
            IsEs
                ? $"{t.Name} · hoja {t.Sheet} · {t.Rows.Count} registros\n{string.Join(", ", t.Headers)}"
                : $"{t.Name} · sheet {t.Sheet} · {t.Rows.Count} records\n{string.Join(", ", t.Headers)}"))
            + T("\n\nRevision: ", "\n\nRevisión: ") + snapshot.Revision;

        var summary = snapshot.Cells.Where(c => c.Sheet == profile.SummarySheet && new[] { "B9", "B10", "B11" }.Contains(c.Address));
        var labels = IsEs
            ? new Dictionary<string, string> { { "B9", "Visitas realizadas" }, { "B10", "Materiales del período (MXN)" }, { "B11", "Materiales bajo mínimo" } }
            : new Dictionary<string, string> { { "B9", "Completed visits" }, { "B10", "Period materials (MXN)" }, { "B11", "Materials below minimum" } };

        Review.Text = string.Join("\n\n", history)
            + T("\n\nCurrent summary\n", "\n\nResumen actual\n")
            + string.Join("\n", summary.Select(c => $"{labels[c.Address]}: {c.Value}"))
            + T(
                "\n\nAlso review the sheets in the Excel window. Approval covers the full review.",
                "\n\nRevisa también las hojas en la ventana de Excel. La aprobación cubre la revisión completa.");
    }

    void RequireBook()
    {
        if (snapshot is null || profile is null)
            throw new InvalidOperationException(T("Open a workbook first.", "Primero abre un libro."));
    }

    async Task Execute(List<Operation> operations, string intent, string revision)
    {
        Status.Text = T("Modifying and verifying draft…", "Modificando y comprobando el borrador…");
        var plan = new ActionPlan(Guid.NewGuid().ToString(), intent, revision, operations);
        var r = await (await Client()).Tool("execute_batch", new { plan });
        var change = Json.Read<ChangeSet>(r.GetRawText());
        var text = new StringBuilder(intent + "\n");
        var formulaLabel = T("Formula", "Fórmula");
        foreach (var c in change.Changes)
            text.AppendLine($"{c.Sheet}!{c.Address}: {c.Before?.Value ?? "∅"} → {c.After?.Value ?? "∅"}" + (c.Before?.Formula != c.After?.Formula ? $"\n{formulaLabel}: {c.Before?.Formula} → {c.After?.Formula}" : ""));
        foreach (var s in change.Structure) text.AppendLine(s);
        history.Add(text.ToString());
        await LoadSnapshot();
        Status.Text = T("Ready for review.", "Listo para revisar.");
    }

    async void Send(object sender, RoutedEventArgs e) => await Run(async () =>
    {
        RequireBook();
        Status.Text = T("Interpreting…", "Interpretando…");
        await LoadSnapshot();
        var revision = snapshot!.Revision;
        var local = LocalInterpreter.TryInterpret(Instruction.Text, snapshot!, profile!);
        if (local is not null && local.Clarification == "" && local.Operations.Count > 0)
        {
            await Execute(local.Operations, Instruction.Text, revision);
            return;
        }
        var planner = PlannerFactory.Create(settings, Credentials.Read(settings.Provider));
        Interpretation result;
        try
        {
            result = await planner.Interpret(Instruction.Text, profile!, async (table, query) => (object)await (await Client()).Tool("read_records", new { table, query }), CancellationToken.None);
        }
        finally
        {
            if (planner.LastUsage is not null)
                await File.AppendAllTextAsync(Path.Combine(root, "usage.jsonl"), Json.Write(planner.LastUsage) + Environment.NewLine);
        }
        if (result.Clarification != "") { Status.Text = T("Clarification needed: ", "Necesito aclaración: ") + result.Clarification; return; }
        if (result.Operations.Count == 0) { Status.Text = T("No changes were proposed.", "No se propusieron cambios."); return; }
        await Execute(result.Operations, Instruction.Text, revision);
    });

    async void ExecuteLocal(object sender, RoutedEventArgs e) => await Run(async () =>
    {
        RequireBook();
        await Execute(Json.Read<List<Operation>>(Batch.Text), T("Local batch", "Lote local"), snapshot!.Revision);
    });

    async void Refresh(object sender, RoutedEventArgs e) => await Run(async () =>
    {
        RequireBook();
        var before = snapshot!;
        await LoadSnapshot();
        if (before.Revision != snapshot!.Revision)
        {
            var delta = Diff.Compare(new("", T("Manual edit", "Edición manual"), before.Revision, []), before, snapshot, "");
            history.Add(T("Manual edit detected\n", "Edición manual detectada\n") + string.Join("\n", delta.Changes.Select(c => $"{c.Sheet}!{c.Address}: {c.Before?.Value} → {c.After?.Value}")));
            await LoadSnapshot();
        }
        Status.Text = T("Review updated.", "Revisión actualizada.");
    });

    async void Undo(object sender, RoutedEventArgs e) => await Run(async () =>
    {
        RequireBook();
        await (await Client()).Tool("restore_checkpoint", new { revision = snapshot!.Revision });
        history.Add(T("Undid the last batch.", "Se deshizo el último lote."));
        await LoadSnapshot();
        Status.Text = T("Draft restored.", "Borrador restaurado.");
    });

    async void ApproveExport(object sender, RoutedEventArgs e) => await Run(async () =>
    {
        RequireBook();
        var destination = new SaveFileDialog
        {
            Filter = T("Excel Workbook|*.xlsx", "Libro Excel|*.xlsx"),
            FileName = T("GridPilot-approved.xlsx", "GridPilot-aprobado.xlsx"),
            OverwritePrompt = true
        };
        if (destination.ShowDialog() != true) return;
        if (MessageBox.Show(
                T("Do you approve the full review shown and its export to a new file?", "¿Apruebas la revisión completa mostrada y su exportación a un archivo nuevo?"),
                T("Approve review", "Aprobar revisión"),
                MessageBoxButton.YesNo,
                MessageBoxImage.Question) != MessageBoxResult.Yes) return;
        await (await Client()).Call("gridpilot/approve", new { revision = snapshot!.Revision });
        var r = await (await Client()).Call("gridpilot/export", new { path = destination.FileName });
        Status.Text = T("Approved file exported. SHA-256: ", "Archivo aprobado exportado. SHA-256: ") + r.GetProperty("Hash").GetString();
    });

    async void Discard(object sender, RoutedEventArgs e) => await Run(async () =>
    {
        await (await Client()).Call("gridpilot/close");
        snapshot = null; profile = null; history.Clear(); Review.Clear(); Structure.Clear();
        Status.Text = T(
            "Session closed without exporting. Local recovery points are preserved.",
            "Sesión cerrada sin exportar. Los puntos de recuperación locales se conservan.");
    });

    void Configure(object sender, RoutedEventArgs e)
    {
        var dialog = new SettingsWindow(settings) { Owner = this };
        if (dialog.ShowDialog() != true) return;
        settings = dialog.Result;
        File.WriteAllText(Path.Combine(root, "settings.json"), Json.Write(settings));
        if (dialog.AnthropicKey != "") Credentials.Save("anthropic", dialog.AnthropicKey);
        if (dialog.GeminiKey != "") Credentials.Save("gemini", dialog.GeminiKey);
        if (dialog.SpeechKey != "") Credentials.Save("speech", dialog.SpeechKey);
        ApplyLanguage();
        Status.Text = T(
            "Settings saved; credentials protected for this Windows user.",
            "Ajustes guardados; credenciales protegidas para este usuario de Windows.");
    }

    void ResetVoiceButton()
    {
        Voice.Content = T("Speak", "Hablar");
        Voice.ClearValue(Button.BackgroundProperty);
        Voice.ClearValue(Button.ForegroundProperty);
    }

    async void ToggleVoice(object sender, RoutedEventArgs e)
    {
        if (busy) return;
        if (!recording)
        {
            try
            {
                speech = new(settings, Credentials.Read("gemini"));
                await speech.Start(settings.MaxVoiceSeconds);
                recording = true;
                Voice.Content = T("Stop", "Detener");
                Voice.Background = new SolidColorBrush(Color.FromRgb(211, 47, 47));
                Voice.Foreground = Brushes.White;
                Status.Text = T(
                    "Listening… Speak now and click «Stop» when finished.",
                    "Escuchando… Habla ahora y haz clic en «Detener» cuando termines.");
            }
            catch (Exception ex) { ResetVoiceButton(); recording = false; speech?.Dispose(); speech = null; Status.Text = ex.Message; }
            return;
        }
        recording = false; ResetVoiceButton(); busy = true; Controls.IsEnabled = false;
        try
        {
            if (speech is not null)
            {
                Status.Text = T("Transcribing audio with Gemini…", "Transcribiendo audio con Gemini…");
                Instruction.Text = await speech.Stop();
                if (speech.LastUsage is not null)
                    await File.AppendAllTextAsync(Path.Combine(root, "usage.jsonl"), Json.Write(speech.LastUsage) + Environment.NewLine);
            }
            Status.Text = T(
                "Transcription ready. Review the text and click Send.",
                "Transcripción lista. Revisa el texto y pulsa Enviar.");
        }
        catch (Exception ex) { Status.Text = ex.Message; }
        finally { speech?.Dispose(); speech = null; busy = false; Controls.IsEnabled = true; }
    }

    void OnClosing(object? sender, CancelEventArgs e)
    {
        if (busy || recording)
        {
            e.Cancel = true;
            Status.Text = T(
                "Wait for the current operation to finish or stop voice dictation.",
                "Espera a que termine la operación o detén el dictado de voz.");
            return;
        }
        speech?.Dispose();
        rpc?.Dispose();
    }
}
