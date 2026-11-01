using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using GridPilot.Core;
using GridPilot.Providers;
using Microsoft.Win32;

namespace GridPilot.App;

public partial class MainWindow : Window
{
    RpcClient? rpc;
    WorkbookSnapshot? snapshot;
    WorkbookProfile? profile;
    bool busy, recording, initializingLanguage;
    SpeechInput? speech;
    readonly string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GridPilot");
    Settings settings = new(Model: "gemini-3.8-flash", Provider: "gemini", Language: "en");
    readonly List<string> history = [];
    readonly ObservableCollection<ChatTurn> turns = [];
    string? selectedProfile;

    bool IsEs => string.Equals(settings.Language, "es", StringComparison.OrdinalIgnoreCase);
    string T(string en, string es) => IsEs ? es : en;
    CredentialStore Credentials => new(Path.Combine(root, "credentials"), settings.Language);

    public MainWindow()
    {
        InitializeComponent();
        ChatList.ItemsSource = turns;
        MaxHeight = SystemParameters.WorkArea.Height;
        MinHeight = Math.Min(MinHeight, MaxHeight);
        Height = Math.Min(Height, MaxHeight);
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        Directory.CreateDirectory(root);
        var file = Path.Combine(root, "settings.json");
        if (File.Exists(file)) settings = Json.Read<Settings>(File.ReadAllText(file));
        if (string.IsNullOrWhiteSpace(settings.Language)) settings = settings with { Language = "en" };
        ApplyLanguage(initial: true);
        PostWelcomeTurn();
    }

    // =========================================================================
    // 1. LOCALIZATION & DYNAMIC HARNESS CHIPS
    // =========================================================================
    void ApplyLanguage(bool initial = false)
    {
        initializingLanguage = true;
        LanguageSelector.SelectedValue = IsEs ? "es" : "en";
        initializingLanguage = false;

        Title = T("GridPilot · Excel by voice and text", "GridPilot · Excel por voz y texto");
        LanguageLabel.Text = T("Language:", "Idioma:");
        ContextToggleBtn.Content = T("☷ Context", "☷ Contexto");
        DrawerTitle.Text = T("WORKBOOK CONTEXT", "ESTRUCTURA DEL LIBRO");
        ReviewDrawerTitle.Text = T("CURRENT SUMMARY & LOG", "RESUMEN Y REVISIÓN");
        BatchDrawerTitle.Text = T("DETERMINISTIC JSON BATCH", "LOTE LOCAL JSON");
        BatchHint.Text = T(
            "Deterministic mode: paste a JSON list of typed operations.",
            "Modo determinista: pega una lista JSON de operaciones tipadas.");
        ExecuteBatchBtn.Content = T("Run batch on draft", "Ejecutar lote en borrador");
        PromptPlaceholder.Text = T(
            "Ask GridPilot in natural language or type / for commands (/demo, /open, /undo, /export, /settings)…",
            "Escribe una instrucción o escribe / para comandos (/demo, /open, /undo, /export, /settings)…");
        Voice.Content = recording ? T("Detener", "Detener") : T("Speak", "Hablar");
        if (recording && !IsEs) Voice.Content = "Stop";
        SendBtn.Content = T("Send", "Enviar");

        if (initial || Status.Text is "Select an .xlsx file to begin." or "Selecciona un archivo .xlsx para comenzar.")
            Status.Text = T("Select an .xlsx file to begin.", "Selecciona un archivo .xlsx para comenzar.");

        UpdateHeaderBadges();
        UpdateComposerChips();
        RenderSnapshotViews();
    }

    void UpdateHeaderBadges()
    {
        if (snapshot is null || profile is null)
        {
            SessionBadge.Text = T("No active workbook", "Sin libro activo");
            SessionBadge.Foreground = new SolidColorBrush(Color.FromRgb(139, 148, 158));
            RevisionPill.Visibility = Visibility.Collapsed;
        }
        else
        {
            var sheets = string.Join(", ", snapshot.Sheets);
            SessionBadge.Text = T($"Draft active · {sheets}", $"Borrador activo · {sheets}");
            SessionBadge.Foreground = new SolidColorBrush(Color.FromRgb(126, 231, 135));
            RevisionPill.Visibility = Visibility.Visible;
            RevisionBadge.Text = $"rev: {snapshot.Revision[..Math.Min(10, snapshot.Revision.Length)]}";
        }
    }

    void UpdateComposerChips()
    {
        if (snapshot is null)
        {
            ComposerChips.ItemsSource = new List<ChatChip>
            {
                new(T("⚡ Demo template (/demo)", "⚡ Plantilla demo (/demo)"), "/demo", true),
                new(T("📂 Open workbook (/open)", "📂 Abrir libro (/open)"), "/open"),
                new(T("🧩 Profile… (/profile)", "🧩 Perfil… (/profile)"), "/profile"),
                new(T("♻ Recover session (/recover)", "♻ Recuperar sesión (/recover)"), "/recover"),
                new(T("⚙ Settings (/settings)", "⚙ Ajustes (/settings)"), "/settings")
            };
        }
        else
        {
            ComposerChips.ItemsSource = new List<ChatChip>
            {
                new(T("🔄 Refresh review (/refresh)", "🔄 Actualizar revisión (/refresh)"), "/refresh"),
                new(T("↩ Undo last batch (/undo)", "↩ Deshacer último lote (/undo)"), "/undo"),
                new(T("✅ Approve & export… (/export)", "✅ Aprobar y exportar… (/export)"), "/export", true),
                new(T("☷ Structure (/structure)", "☷ Estructura (/structure)"), "/structure"),
                new(T("🗑 Discard (/discard)", "🗑 Descartar sesión (/discard)"), "/discard"),
                new(T("⚙ Settings (/settings)", "⚙ Ajustes (/settings)"), "/settings")
            };
        }
    }

    List<SlashCommandItem> GetSlashCommands() =>
    [
        new("/demo", "", T("Open the built-in GridPilot-Demo.xlsx workbook in a draft session", "Abre la plantilla GridPilot-Demo.xlsx en un borrador visible")),
        new("/open", "", T("Choose an .xlsx file from disk to open as a safe draft", "Selecciona un archivo .xlsx para abrir en borrador")),
        new("/profile", "", T("Select a custom profile.json for the next workbook", "Selecciona un perfil JSON para el próximo libro")),
        new("/recover", "", T("Recover an interrupted session from its session.db", "Recupera una sesión desde su archivo session.db")),
        new("/refresh", "", T("Inspect the Excel window and detect manual edits", "Actualiza la revisión y detecta ediciones manuales en Excel")),
        new("/undo", "", T("Undo the last executed batch and restore the previous checkpoint", "Deshace el último lote y restaura el punto previo")),
        new("/export", "", T("Explicitly approve the full review and export to a new .xlsx file", "Aprueba la revisión completa y exporta a un archivo .xlsx nuevo")),
        new("/discard", "", T("Close the active draft session without exporting", "Descarta y cierra la sesión sin exportar")),
        new("/structure", "", T("Toggle the workbook context drawer and display table schemas", "Muestra u oculta el panel lateral de estructura y lotes")),
        new("/batch", "[json]", T("Execute a deterministic JSON list of operations or open the batch drawer", "Ejecuta un lote JSON de operaciones tipadas")),
        new("/settings", "", T("Configure AI provider (Gemini / Anthropic), API keys, and limits", "Abre Ajustes (proveedor, claves API y límites)")),
        new("/lang", "en | es", T("Switch interface language between English (en) and Español (es)", "Cambia el idioma de la interfaz entre inglés (en) y español (es)")),
        new("/clear", "", T("Clear the conversation timeline", "Limpia el historial visual del chat")),
        new("/help", "", T("List available slash commands and sample workbook instructions", "Muestra los comandos disponibles y ejemplos de instrucciones"))
    ];

    void PostWelcomeTurn()
    {
        turns.Add(new ChatTurn
        {
            Role = ChatRole.Agent,
            Author = "GRIDPILOT AGENT",
            ToolBadge = "agent_ready",
            Text = T(
                "Welcome to GridPilot Agent Harness. Work on a visible Excel draft using natural-language instructions, voice dictation, or slash commands (/demo, /open, /undo, /export, /settings).\n\nSelect an action below or type / in the prompt box to begin.",
                "Bienvenido al Agent Harness de GridPilot. Trabaja en un borrador visible de Excel mediante instrucciones en lenguaje natural, dictado por voz o comandos slash (/demo, /open, /undo, /export, /settings).\n\nSelecciona una opción debajo o escribe / en la caja de texto para comenzar."),
            Chips =
            [
                new(T("⚡ Open demo template", "⚡ Plantilla demo"), "/demo", true),
                new(T("📂 Open workbook", "📂 Abrir libro"), "/open"),
                new(T("♻ Recover session", "♻ Recuperar sesión"), "/recover"),
                new(T("⚙ Settings", "⚙ Ajustes"), "/settings"),
                new(T("❓ Commands & examples", "❓ Comandos y ejemplos"), "/help")
            ]
        });
        ScrollToBottom();
    }

    void AddUserTurn(string text)
    {
        turns.Add(new ChatTurn
        {
            Role = ChatRole.User,
            Author = T("YOU", "TÚ"),
            Text = text
        });
        ScrollToBottom();
    }

    void AddAgentTurn(string text, string badge = "", string diffHeader = "", List<DiffLine>? diffs = null, List<string>? metrics = null, List<ChatChip>? chips = null, bool isError = false)
    {
        turns.Add(new ChatTurn
        {
            Role = ChatRole.Agent,
            Author = "GRIDPILOT AGENT",
            ToolBadge = badge,
            Text = text,
            DiffHeader = diffHeader,
            DiffLines = diffs ?? [],
            SummaryMetrics = metrics ?? [],
            Chips = chips ?? [],
            IsError = isError
        });
        ScrollToBottom();
    }

    void ScrollToBottom()
    {
        Dispatcher.InvokeAsync(() => ChatScroll.ScrollToEnd(), System.Windows.Threading.DispatcherPriority.Background);
    }

    List<string> BuildSummaryMetrics()
    {
        if (snapshot is null || profile is null) return [];
        var summary = snapshot.Cells.Where(c => c.Sheet == profile.SummarySheet && new[] { "B9", "B10", "B11" }.Contains(c.Address));
        var labels = IsEs
            ? new Dictionary<string, string> { { "B9", "Visitas realizadas" }, { "B10", "Materiales del período (MXN)" }, { "B11", "Materiales bajo mínimo" } }
            : new Dictionary<string, string> { { "B9", "Completed visits" }, { "B10", "Period materials (MXN)" }, { "B11", "Materials below minimum" } };
        return summary.Select(c => $"{labels[c.Address]}: {c.Value}").ToList();
    }

    List<ChatChip> SampleWorkbookChips() =>
    [
        new(
            T("Replace row 2 with row 3 (Inventory)", "Reemplazar fila 2 por fila 3 (Inventario)"),
            T("In the inventory tab, replace the data of the second row with the third row", "En la pestaña de inventario, reemplaza los datos de la segunda fila por los de la tercera fila")),
        new(
            T("Swap rows 1 and 3 (Inventory)", "Intercambiar fila 1 y fila 3 (Inventario)"),
            T("In the inventory tab, swap the first row with the third row", "En la pestaña de inventario, intercambia la primera fila con la tercera fila")),
        new(
            T("Delete first row (Visits)", "Eliminar primera fila (Visitas)"),
            T("In the visits tab, delete the entire first row", "En la pestaña de visita, elimina toda la primera fila")),
        new(
            T("↩ Undo (/undo)", "↩ Deshacer (/undo)"),
            "/undo"),
        new(
            T("✅ Approve & export (/export)", "✅ Aprobar y exportar (/export)"),
            "/export")
    ];

    // =========================================================================
    // 2. SLASH AUTOCOMPLETE, KEYBOARD & CHIP INTERACTION
    // =========================================================================
    void OnInstructionTextChanged(object sender, TextChangedEventArgs e)
    {
        var text = Instruction.Text;
        PromptPlaceholder.Visibility = string.IsNullOrEmpty(text) ? Visibility.Visible : Visibility.Collapsed;

        var trimmed = text.TrimStart();
        if (trimmed.StartsWith('/') && !trimmed.Contains('\n'))
        {
            var token = trimmed.Split(' ')[0];
            var matches = GetSlashCommands()
                .Where(c => c.Command.StartsWith(token, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (matches.Count > 0)
            {
                SlashList.ItemsSource = matches;
                SlashList.SelectedIndex = 0;
                SlashPopup.IsOpen = true;
                return;
            }
        }
        SlashPopup.IsOpen = false;
    }

    void OnInstructionPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (SlashPopup.IsOpen)
        {
            if (e.Key == Key.Down)
            {
                SlashList.SelectedIndex = Math.Min(SlashList.Items.Count - 1, SlashList.SelectedIndex + 1);
                e.Handled = true;
                return;
            }
            if (e.Key == Key.Up)
            {
                SlashList.SelectedIndex = Math.Max(0, SlashList.SelectedIndex - 1);
                e.Handled = true;
                return;
            }
            if (e.Key == Key.Tab && SlashList.SelectedItem is SlashCommandItem item)
            {
                Instruction.Text = item.Command + (string.IsNullOrEmpty(item.Usage) ? "" : " ");
                Instruction.CaretIndex = Instruction.Text.Length;
                SlashPopup.IsOpen = false;
                e.Handled = true;
                return;
            }
            if (e.Key == Key.Escape)
            {
                SlashPopup.IsOpen = false;
                e.Handled = true;
                return;
            }
        }

        // Enter sends the message; Shift+Enter inserts a newline.
        if (e.Key == Key.Enter && (Keyboard.Modifiers & ModifierKeys.Shift) == 0)
        {
            SlashPopup.IsOpen = false;
            e.Handled = true;
            Send(this, new RoutedEventArgs());
        }
    }

    void OnSlashListClick(object sender, MouseButtonEventArgs e)
    {
        if (SlashList.SelectedItem is not SlashCommandItem item) return;
        SlashPopup.IsOpen = false;
        if (!string.IsNullOrEmpty(item.Usage))
        {
            Instruction.Text = item.Command + " ";
            Instruction.CaretIndex = Instruction.Text.Length;
            Instruction.Focus();
            return;
        }
        Instruction.Text = item.Command;
        Send(this, new RoutedEventArgs());
    }

    void OnChipClick(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: string cmd } || string.IsNullOrWhiteSpace(cmd)) return;
        Instruction.Text = cmd;
        Send(this, new RoutedEventArgs());
    }

    void ToggleContextDrawer(object sender, RoutedEventArgs e)
    {
        ContextDrawer.Visibility = ContextDrawer.Visibility == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;
    }

    void ChangeLanguage(object sender, SelectionChangedEventArgs e)
    {
        if (initializingLanguage || LanguageSelector.SelectedValue is not string lang) return;
        settings = settings with { Language = lang };
        File.WriteAllText(Path.Combine(root, "settings.json"), Json.Write(settings));
        ApplyLanguage();
    }

    // =========================================================================
    // 3. CORE WORKBOOK, MCP RPC & EXECUTION CONTROLLERS
    // =========================================================================
    async Task Run(Func<Task> action)
    {
        if (busy || recording) return;
        busy = true; Controls.IsEnabled = false;
        try { await action(); }
        catch (Exception e)
        {
            Status.Text = e.Message;
            AddAgentTurn(e.Message, badge: "error", isError: true);
        }
        finally { busy = false; Controls.IsEnabled = true; Instruction.Focus(); }
    }

    async Task<RpcClient> Client()
    {
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
        var editable = result.GetProperty("Editable").GetBoolean();
        Status.Text = editable
            ? T("Draft ready. The original file remains untouched by writes.", "Borrador listo. El original permanece fuera de las escrituras.")
            : T("Workbook available for inspection only: incompatible profile.", "Libro disponible solo para inspección: perfil incompatible.");

        var tableSummary = snapshot!.Tables.Select(t =>
            new DiffLine(t.Name, T($"sheet {t.Sheet}", $"hoja {t.Sheet}"), T($"{t.Rows.Count} records ({string.Join(", ", t.Headers)})", $"{t.Rows.Count} registros ({string.Join(", ", t.Headers)})"))).ToList();

        AddAgentTurn(
            $"{Status.Text}\n{T("File:", "Archivo:")} {Path.GetFileName(path)}",
            badge: "gridpilot/open",
            diffHeader: T($"Workbook loaded · {snapshot.Tables.Count} tables · rev {snapshot.Revision[..8]}", $"Libro cargado · {snapshot.Tables.Count} tablas · rev {snapshot.Revision[..8]}"),
            diffs: tableSummary,
            metrics: BuildSummaryMetrics(),
            chips: SampleWorkbookChips());
    }

    async void Open(object sender, RoutedEventArgs e) => await Run(OpenInteractiveAsync);
    async Task OpenInteractiveAsync()
    {
        var dialog = new OpenFileDialog { Filter = T("Excel Workbook|*.xlsx", "Libro Excel|*.xlsx") };
        if (dialog.ShowDialog() == true) await OpenPath(dialog.FileName);
    }

    async void Demo(object sender, RoutedEventArgs e) => await Run(DemoAsync);
    async Task DemoAsync()
    {
        selectedProfile = null;
        await OpenPath(Path.Combine(AppContext.BaseDirectory, "assets", "GridPilot-Demo.xlsx"));
    }

    void SelectProfile(object sender, RoutedEventArgs e) => SelectProfileInteractive();
    void SelectProfileInteractive()
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
            AddAgentTurn($"{Status.Text}\n{dialog.FileName}", badge: "profile");
        }
        catch (Exception ex)
        {
            Status.Text = ex.Message;
            AddAgentTurn(ex.Message, badge: "error", isError: true);
        }
    }

    async void Recover(object sender, RoutedEventArgs e) => await Run(RecoverInteractiveAsync);
    async Task RecoverInteractiveAsync()
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
            AddAgentTurn(
                Status.Text,
                badge: "gridpilot/recover",
                diffHeader: T($"Recovered state · rev {snapshot!.Revision[..8]}", $"Estado recuperado · rev {snapshot!.Revision[..8]}"),
                metrics: BuildSummaryMetrics(),
                chips: SampleWorkbookChips());
        }
    }

    async Task LoadSnapshot()
    {
        var r = await (await Client()).Tool("inspect_workbook", new { });
        snapshot = Json.Read<WorkbookSnapshot>(r.GetProperty("snapshot").GetRawText());
        profile = Json.Read<WorkbookProfile>(r.GetProperty("profile").GetRawText());
        UpdateHeaderBadges();
        UpdateComposerChips();
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

        var metrics = BuildSummaryMetrics();
        Review.Text = string.Join("\n\n", history)
            + T("\n\nCurrent summary\n", "\n\nResumen actual\n")
            + string.Join("\n", metrics)
            + T(
                "\n\nAlso review the sheets in the Excel window. Approval covers the full review.",
                "\n\nRevisa también las hojas en la ventana de Excel. La aprobación cubre la revisión completa.");
    }

    void RequireBook()
    {
        if (snapshot is null || profile is null)
            throw new InvalidOperationException(T("Open a workbook first.", "Primero abre un libro."));
    }

    async Task Execute(List<Operation> operations, string intent, string revision, string badge = "execute_batch")
    {
        Status.Text = T("Modifying and verifying draft…", "Modificando y comprobando el borrador…");
        var plan = new ActionPlan(Guid.NewGuid().ToString(), intent, revision, operations);
        var r = await (await Client()).Tool("execute_batch", new { plan });
        var change = Json.Read<ChangeSet>(r.GetRawText());

        var text = new StringBuilder(intent + "\n");
        var formulaLabel = T("Formula", "Fórmula");
        var diffLines = new List<DiffLine>();

        foreach (var c in change.Changes)
        {
            var beforeVal = c.Before?.Value ?? "∅";
            var afterVal = c.After?.Value ?? "∅";
            var fDiff = c.Before?.Formula != c.After?.Formula ? $"{formulaLabel}: {c.Before?.Formula} → {c.After?.Formula}" : "";
            text.AppendLine($"{c.Sheet}!{c.Address}: {beforeVal} → {afterVal}" + (fDiff != "" ? $"\n{fDiff}" : ""));
            diffLines.Add(new($"{c.Sheet}!{c.Address}", beforeVal, afterVal, fDiff));
        }
        foreach (var s in change.Structure)
        {
            text.AppendLine(s);
            diffLines.Add(new(T("Structure", "Estructura"), "—", s));
        }

        history.Add(text.ToString());
        await LoadSnapshot();
        Status.Text = T("Ready for review.", "Listo para revisar.");

        AddAgentTurn(
            T(
                $"Applied {operations.Count} operation(s) and verified formulas, balances, and relationships.",
                $"Se aplicaron {operations.Count} operación(es) y se comprobaron relaciones, saldos, costos y fórmulas."),
            badge: $"{badge} · {operations.Count} op(s)",
            diffHeader: T($"{change.Changes.Count} cell change(s) · rev {snapshot!.Revision[..8]}", $"{change.Changes.Count} celda(s) modificada(s) · rev {snapshot!.Revision[..8]}"),
            diffs: diffLines,
            metrics: BuildSummaryMetrics(),
            chips:
            [
                new(T("↩ Undo last batch (/undo)", "↩ Deshacer último lote (/undo)"), "/undo"),
                new(T("🔄 Refresh review (/refresh)", "🔄 Actualizar revisión (/refresh)"), "/refresh"),
                new(T("✅ Approve & export… (/export)", "✅ Aprobar y exportar… (/export)"), "/export", true)
            ]);
    }

    // =========================================================================
    // 4. CONVERSATIONAL ROUTER (SLASH COMMANDS + NATURAL LANGUAGE + AI PLANNER)
    // =========================================================================
    async void Send(object sender, RoutedEventArgs e)
    {
        var input = Instruction.Text.Trim();
        if (string.IsNullOrWhiteSpace(input) || busy || recording) return;
        Instruction.Clear();
        SlashPopup.IsOpen = false;
        AddUserTurn(input);

        await Run(async () =>
        {
            if (await TryHandleCommandAsync(input)) return;

            RequireBook();
            Status.Text = T("Interpreting…", "Interpretando…");
            await LoadSnapshot();
            var revision = snapshot!.Revision;

            var local = LocalInterpreter.TryInterpret(input, snapshot!, profile!);
            if (local is not null && local.Clarification == "" && local.Operations.Count > 0)
            {
                await Execute(local.Operations, input, revision, "local_interpreter");
                return;
            }

            var planner = PlannerFactory.Create(settings, Credentials.Read(settings.Provider));
            Interpretation result;
            try
            {
                result = await planner.Interpret(input, profile!, async (table, query) => (object)await (await Client()).Tool("read_records", new { table, query }), CancellationToken.None);
            }
            finally
            {
                if (planner.LastUsage is not null)
                    await File.AppendAllTextAsync(Path.Combine(root, "usage.jsonl"), Json.Write(planner.LastUsage) + Environment.NewLine);
            }

            if (result.Clarification != "")
            {
                Status.Text = T("Clarification needed: ", "Necesito aclaración: ") + result.Clarification;
                AddAgentTurn(Status.Text, badge: "clarification");
                return;
            }
            if (result.Operations.Count == 0)
            {
                Status.Text = T("No changes were proposed.", "No se propusieron cambios.");
                AddAgentTurn(Status.Text, badge: "no_op");
                return;
            }
            await Execute(result.Operations, input, revision, $"{settings.Provider}_planner");
        });
    }

    async Task<bool> TryHandleCommandAsync(string input)
    {
        var lower = input.Trim().ToLowerInvariant();

        // Direct natural-language meta-aliases when the user types conversational system actions
        if (lower is "/demo" or "plantilla demo" or "abrir plantilla demo" or "abre la plantilla demo" or "open demo" or "open demo template" or "demo template")
        {
            await DemoAsync();
            return true;
        }
        if (lower is "/open" or "abrir libro" or "abre un libro" or "open workbook")
        {
            await OpenInteractiveAsync();
            return true;
        }
        if (lower is "/profile" or "perfil" or "seleccionar perfil" or "select profile")
        {
            SelectProfileInteractive();
            return true;
        }
        if (lower is "/recover" or "recuperar sesion" or "recuperar sesión" or "recover session")
        {
            await RecoverInteractiveAsync();
            return true;
        }
        if (lower is "/refresh" or "actualizar revision" or "actualizar revisión" or "refresh review" or "refresh")
        {
            await RefreshAsync();
            return true;
        }
        if (lower is "/undo" or "deshacer" or "deshacer ultimo lote" or "deshacer último lote" or "undo" or "undo last batch")
        {
            await UndoAsync();
            return true;
        }
        if (lower is "/export" or "/approve" or "aprobar y exportar" or "exportar" or "approve and export" or "export")
        {
            await ApproveExportAsync();
            return true;
        }
        if (lower is "/discard" or "/close" or "descartar sesion" or "descartar sesión" or "discard session")
        {
            await DiscardAsync();
            return true;
        }
        if (lower is "/structure" or "/context" or "estructura" or "structure")
        {
            ContextDrawer.Visibility = ContextDrawer.Visibility == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;
            if (snapshot is not null)
            {
                var lines = snapshot.Tables.Select(t => new DiffLine(t.Name, t.Sheet, $"{t.Rows.Count} · {string.Join(", ", t.Headers)}")).ToList();
                AddAgentTurn(
                    T("Workbook structure and context drawer toggled.", "Estructura del libro y panel lateral actualizados."),
                    badge: "inspect_workbook",
                    diffHeader: T($"Tables · rev {snapshot.Revision[..8]}", $"Tablas · rev {snapshot.Revision[..8]}"),
                    diffs: lines,
                    metrics: BuildSummaryMetrics());
            }
            return true;
        }
        if (lower.StartsWith("/batch"))
        {
            var payload = input[6..].Trim();
            if (string.IsNullOrWhiteSpace(payload))
            {
                ContextDrawer.Visibility = Visibility.Visible;
                AddAgentTurn(
                    T("Opened the Deterministic JSON Batch drawer on the right. Paste your JSON operations there or pass `/batch [{\"Kind\":...}]` directly in chat.",
                      "Se abrió el panel de Lote Local JSON a la derecha. Pega ahí tu lista JSON o escribe `/batch [{\"Kind\":...}]` directamente en el chat."),
                    badge: "batch_drawer");
                return true;
            }
            RequireBook();
            await Execute(Json.Read<List<Operation>>(payload), T("Local batch", "Lote local"), snapshot!.Revision, "json_batch");
            return true;
        }
        if (lower is "/settings" or "ajustes" or "configuracion" or "configuración" or "settings")
        {
            ConfigureInteractive();
            return true;
        }
        if (lower.StartsWith("/lang"))
        {
            var arg = lower.Replace("/lang", "").Trim();
            var nextLang = arg switch
            {
                "es" or "español" or "espanol" or "spanish" => "es",
                "en" or "english" or "ingles" or "inglés" => "en",
                _ => IsEs ? "en" : "es"
            };
            settings = settings with { Language = nextLang };
            File.WriteAllText(Path.Combine(root, "settings.json"), Json.Write(settings));
            ApplyLanguage();
            AddAgentTurn(
                T("Interface language switched to English.", "Idioma de la interfaz cambiado a Español."),
                badge: $"lang:{nextLang}");
            return true;
        }
        if (lower is "/clear" or "limpiar chat" or "clear chat")
        {
            turns.Clear();
            PostWelcomeTurn();
            return true;
        }
        if (lower is "/help" or "ayuda" or "help")
        {
            var cmds = string.Join("\n", GetSlashCommands().Select(c => $"• {c.DisplayCommand} — {c.Description}"));
            AddAgentTurn(
                T("Available Agent Harness commands:\n\n", "Comandos disponibles del Agent Harness:\n\n") + cmds,
                badge: "help",
                chips: snapshot is null
                    ? [new(T("⚡ Demo template (/demo)", "⚡ Plantilla demo (/demo)"), "/demo", true), new(T("📂 Open workbook (/open)", "📂 Abrir libro (/open)"), "/open")]
                    : SampleWorkbookChips());
            return true;
        }

        return false;
    }

    async void ExecuteLocal(object sender, RoutedEventArgs e) => await Run(async () =>
    {
        RequireBook();
        AddUserTurn($"/batch {Batch.Text}");
        await Execute(Json.Read<List<Operation>>(Batch.Text), T("Local batch", "Lote local"), snapshot!.Revision, "json_batch");
    });

    async void Refresh(object sender, RoutedEventArgs e) => await Run(RefreshAsync);
    async Task RefreshAsync()
    {
        RequireBook();
        var before = snapshot!;
        await LoadSnapshot();
        var diffs = new List<DiffLine>();
        if (before.Revision != snapshot!.Revision)
        {
            var delta = Diff.Compare(new("", T("Manual edit", "Edición manual"), before.Revision, []), before, snapshot, "");
            history.Add(T("Manual edit detected\n", "Edición manual detectada\n") + string.Join("\n", delta.Changes.Select(c => $"{c.Sheet}!{c.Address}: {c.Before?.Value} → {c.After?.Value}")));
            diffs = delta.Changes.Select(c => new DiffLine($"{c.Sheet}!{c.Address}", c.Before?.Value ?? "∅", c.After?.Value ?? "∅")).ToList();
            await LoadSnapshot();
        }
        Status.Text = T("Review updated.", "Revisión actualizada.");
        AddAgentTurn(
            diffs.Count > 0
                ? T($"Manual edit detected in Excel ({diffs.Count} cell(s)).", $"Edición manual detectada en Excel ({diffs.Count} celda(s)).")
                : Status.Text,
            badge: "inspect_workbook",
            diffHeader: T($"Current state · rev {snapshot!.Revision[..8]}", $"Estado actual · rev {snapshot!.Revision[..8]}"),
            diffs: diffs,
            metrics: BuildSummaryMetrics());
    }

    async void Undo(object sender, RoutedEventArgs e) => await Run(UndoAsync);
    async Task UndoAsync()
    {
        RequireBook();
        await (await Client()).Tool("restore_checkpoint", new { revision = snapshot!.Revision });
        history.Add(T("Undid the last batch.", "Se deshizo el último lote."));
        await LoadSnapshot();
        Status.Text = T("Draft restored.", "Borrador restaurado.");
        AddAgentTurn(
            $"{T("Undid the last batch.", "Se deshizo el último lote.")} {Status.Text}",
            badge: "restore_checkpoint",
            diffHeader: T($"Restored checkpoint · rev {snapshot!.Revision[..8]}", $"Punto restaurado · rev {snapshot!.Revision[..8]}"),
            metrics: BuildSummaryMetrics());
    }

    async void ApproveExport(object sender, RoutedEventArgs e) => await Run(ApproveExportAsync);
    async Task ApproveExportAsync()
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
        var hash = r.GetProperty("Hash").GetString();
        Status.Text = T("Approved file exported. SHA-256: ", "Archivo aprobado exportado. SHA-256: ") + hash;
        AddAgentTurn(
            $"{Status.Text}\n{destination.FileName}",
            badge: "gridpilot/export",
            diffHeader: T("Verified immutable export", "Exportación aprobada y verificada"),
            diffs: [new("SHA-256", "—", hash ?? "")],
            metrics: BuildSummaryMetrics());
    }

    async void Discard(object sender, RoutedEventArgs e) => await Run(DiscardAsync);
    async Task DiscardAsync()
    {
        await (await Client()).Call("gridpilot/close");
        snapshot = null; profile = null; history.Clear(); Review.Clear(); Structure.Clear();
        UpdateHeaderBadges();
        UpdateComposerChips();
        Status.Text = T(
            "Session closed without exporting. Local recovery points are preserved.",
            "Sesión cerrada sin exportar. Los puntos de recuperación locales se conservan.");
        AddAgentTurn(
            Status.Text,
            badge: "gridpilot/close",
            chips:
            [
                new(T("⚡ Demo template (/demo)", "⚡ Plantilla demo (/demo)"), "/demo", true),
                new(T("📂 Open workbook (/open)", "📂 Abrir libro (/open)"), "/open"),
                new(T("♻ Recover session (/recover)", "♻ Recuperar sesión (/recover)"), "/recover")
            ]);
    }

    void Configure(object sender, RoutedEventArgs e) => ConfigureInteractive();
    void ConfigureInteractive()
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
        AddAgentTurn(
            $"{Status.Text}\n{T("Active provider:", "Proveedor activo:")} {settings.Provider} ({settings.Model})",
            badge: "settings_saved");
    }

    // =========================================================================
    // 5. CLICK-TO-SPEAK / RED CLICK-TO-STOP VOICE DICTATION
    // =========================================================================
    void ResetVoiceButton()
    {
        Voice.Content = T("Speak", "Hablar");
        Voice.ClearValue(Button.BackgroundProperty);
        Voice.ClearValue(Button.ForegroundProperty);
        Voice.ClearValue(Button.BorderBrushProperty);
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
                Voice.BorderBrush = new SolidColorBrush(Color.FromRgb(248, 81, 73));
                Voice.Foreground = Brushes.White;
                Status.Text = T(
                    "Listening… Speak now and click «Stop» when finished.",
                    "Escuchando… Habla ahora y haz clic en «Detener» cuando termines.");
            }
            catch (Exception ex)
            {
                ResetVoiceButton();
                recording = false;
                speech?.Dispose();
                speech = null;
                Status.Text = ex.Message;
                AddAgentTurn(ex.Message, badge: "voice_error", isError: true);
            }
            return;
        }
        recording = false; ResetVoiceButton(); busy = true; Controls.IsEnabled = false;
        try
        {
            if (speech is not null)
            {
                Status.Text = T("Transcribing audio with Gemini…", "Transcribiendo audio con Gemini…");
                Instruction.Text = await speech.Stop();
                Instruction.CaretIndex = Instruction.Text.Length;
                if (speech.LastUsage is not null)
                    await File.AppendAllTextAsync(Path.Combine(root, "usage.jsonl"), Json.Write(speech.LastUsage) + Environment.NewLine);
            }
            Status.Text = T(
                "Transcription ready. Review the text and click Send.",
                "Transcripción lista. Revisa el texto y pulsa Enviar.");
        }
        catch (Exception ex)
        {
            Status.Text = ex.Message;
            AddAgentTurn(ex.Message, badge: "voice_error", isError: true);
        }
        finally { speech?.Dispose(); speech = null; busy = false; Controls.IsEnabled = true; Instruction.Focus(); }
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
