using System.Windows;
using System.Windows.Controls;
using GridPilot.Providers;
namespace GridPilot.App;
public sealed class SettingsWindow : Window
{
    public Settings Result { get; private set; }
    public string GeminiKey => gemini.Password;
    public string AnthropicKey => anthropic.Password;
    public string SpeechKey => speech.Password;
    readonly PasswordBox anthropic = new(), gemini = new(), speech = new();

    public SettingsWindow(Settings current)
    {
        Result = current;
        var langCode = string.Equals(current.Language, "es", StringComparison.OrdinalIgnoreCase) ? "es" : "en";
        bool IsEs() => langCode == "es";
        string T(string en, string es) => IsEs() ? es : en;

        Width = 480;
        Height = Math.Min(720, SystemParameters.WorkArea.Height);
        MaxHeight = SystemParameters.WorkArea.Height;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;

        var panel = new StackPanel { Margin = new Thickness(24) };
        Content = new ScrollViewer { Content = panel, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };

        var langLabel = new TextBlock { Margin = new Thickness(0, 0, 0, 4) };
        var langCombo = new ComboBox
        {
            SelectedValuePath = "Tag",
            Margin = new Thickness(0, 0, 0, 16),
            Items =
            {
                new ComboBoxItem { Content = "English", Tag = "en" },
                new ComboBoxItem { Content = "Español", Tag = "es" }
            },
            SelectedValue = langCode
        };
        panel.Children.Add(langLabel);
        panel.Children.Add(langCombo);

        var cloud = new CheckBox { IsChecked = current.CloudEnabled, Margin = new Thickness(0, 0, 0, 20) };
        panel.Children.Add(cloud);

        var providerLabel = new TextBlock();
        panel.Children.Add(providerLabel);
        var provider = new ComboBox { ItemsSource = new[] { "gemini", "anthropic" }, SelectedItem = current.Provider };
        panel.Children.Add(provider);

        var modelLabel = new TextBlock { Margin = new Thickness(0, 8, 0, 4) };
        var model = new TextBox { Text = current.Model };
        panel.Children.Add(modelLabel);
        panel.Children.Add(model);
        provider.SelectionChanged += (_, _) => { model.Text = (string)provider.SelectedItem == "gemini" ? "gemini-3.8-flash" : "claude-sonnet-5-5"; };

        var geminiLabel = new TextBlock { TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 12, 0, 4) };
        panel.Children.Add(geminiLabel);
        panel.Children.Add(gemini);

        var anthropicLabel = new TextBlock { Margin = new Thickness(0, 12, 0, 4) };
        panel.Children.Add(anthropicLabel);
        panel.Children.Add(anthropic);

        var voiceNote = new TextBlock { Foreground = System.Windows.Media.Brushes.DimGray, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 12, 0, 4) };
        panel.Children.Add(voiceNote);

        var callsLabel = new TextBlock { Margin = new Thickness(0, 8, 0, 4) };
        var calls = new TextBox { Text = current.MaxCalls.ToString() };
        panel.Children.Add(callsLabel);
        panel.Children.Add(calls);

        var tokensLabel = new TextBlock { Margin = new Thickness(0, 8, 0, 4) };
        var tokens = new TextBox { Text = current.MaxTotalTokens.ToString() };
        panel.Children.Add(tokensLabel);
        panel.Children.Add(tokens);

        var save = new Button { Margin = new Thickness(0, 20, 0, 0) };
        panel.Children.Add(save);

        void RefreshTexts()
        {
            Title = T("GridPilot Settings", "Ajustes de GridPilot");
            langLabel.Text = T("Interface language", "Idioma de la interfaz");
            cloud.Content = T("Authorize cloud services usage (may incur costs)", "Autorizar uso de servicios de nube (puede generar costos)");
            providerLabel.Text = T("Interpretation provider", "Proveedor de interpretación");
            modelLabel.Text = T("Exact model (changing requires re-evaluating the pilot)", "Modelo exacto (un cambio requiere reevaluar el piloto)");
            geminiLabel.Text = T("Gemini API / AI Studio key (interpretation and voice dictation; leave blank to keep saved)", "Clave Gemini API / AI Studio (interpretación y dictado por voz; vacío conserva la guardada)");
            anthropicLabel.Text = T("Anthropic key (optional; leave blank to keep saved)", "Clave Anthropic (opcional; vacío conserva la guardada)");
            voiceNote.Text = T("Voice dictation: automatically uses Gemini Multimodal Audio with your Gemini key (no Azure Speech needed).", "Dictado por voz: usa automáticamente Gemini Multimodal Audio (es-MX) con tu clave Gemini (sin necesidad de Azure Speech).");
            callsLabel.Text = T("Max calls per instruction (1–10)", "Máximo de llamadas por instrucción (1–10)");
            tokensLabel.Text = T("Max tokens per instruction", "Máximo de tokens por instrucción");
            save.Content = T("Save", "Guardar");
        }

        langCombo.SelectionChanged += (_, _) =>
        {
            if (langCombo.SelectedValue is string selected)
            {
                langCode = selected;
                RefreshTexts();
            }
        };
        RefreshTexts();

        save.Click += (_, _) =>
        {
            if (!int.TryParse(calls.Text, out var n) || n < 1 || n > 10 || !int.TryParse(tokens.Text, out var t) || t < 1000 || t > 50000 || string.IsNullOrWhiteSpace(model.Text) || provider.SelectedItem is null)
            {
                MessageBox.Show(T("Check the limits and model.", "Revisa los límites y el modelo."));
                return;
            }
            Result = current with
            {
                Provider = (string)provider.SelectedItem,
                CloudEnabled = cloud.IsChecked == true,
                Model = model.Text.Trim(),
                MaxCalls = n,
                MaxTotalTokens = t,
                Language = langCode
            };
            DialogResult = true;
        };
    }
}
