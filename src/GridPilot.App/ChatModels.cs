using System.Windows;

namespace GridPilot.App;

/// <summary>
/// Identifies the author of a turn in the Agent Harness timeline.
/// </summary>
public enum ChatRole
{
    User,
    Agent
}

/// <summary>
/// Represents an interactive pill chip attached to an agent turn or the composer bar.
/// Clicking a chip either executes a slash command (e.g. "/demo") or fills/sends a sample prompt.
/// </summary>
public sealed record ChatChip(string Label, string Command, bool IsPrimary = false);

/// <summary>
/// Structured line item inside an expandable Agent Tool/Diff Card.
/// </summary>
public sealed record DiffLine(string Target, string Before, string After, string FormulaDiff = "")
{
    public Visibility FormulaVisibility => string.IsNullOrWhiteSpace(FormulaDiff) ? Visibility.Collapsed : Visibility.Visible;
}

/// <summary>
/// View-model for a single turn in the Agent Harness conversation stream.
/// Designed as a clean template so new agent badges, telemetry fields, or cards can be added easily.
/// </summary>
public sealed class ChatTurn
{
    public ChatRole Role { get; init; } = ChatRole.Agent;
    public string Author { get; init; } = "GRIDPILOT";
    public string Timestamp { get; init; } = DateTime.Now.ToString("HH:mm");
    public string ToolBadge { get; init; } = "";
    public string Text { get; init; } = "";
    public string DiffHeader { get; init; } = "";
    public List<DiffLine> DiffLines { get; init; } = [];
    public List<string> SummaryMetrics { get; init; } = [];
    public List<ChatChip> Chips { get; init; } = [];
    public bool IsError { get; init; }

    public Visibility UserVisibility => Role == ChatRole.User ? Visibility.Visible : Visibility.Collapsed;
    public Visibility AgentVisibility => Role == ChatRole.Agent ? Visibility.Visible : Visibility.Collapsed;
    public Visibility BadgeVisibility => string.IsNullOrWhiteSpace(ToolBadge) ? Visibility.Collapsed : Visibility.Visible;
    public Visibility DiffCardVisibility => DiffLines.Count > 0 || SummaryMetrics.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
    public Visibility DiffLinesVisibility => DiffLines.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
    public Visibility SummaryVisibility => SummaryMetrics.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
    public Visibility ChipsVisibility => Chips.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
}

/// <summary>
/// Entry displayed in the floating '/' slash-command autocomplete popup.
/// </summary>
public sealed record SlashCommandItem(string Command, string Usage, string Description)
{
    public string DisplayCommand => string.IsNullOrWhiteSpace(Usage) ? Command : $"{Command} {Usage}";
}
