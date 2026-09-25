using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

namespace TiffToPdf;

/// <summary>The preferences form, shown in a <see cref="DialogWindow"/>.</summary>
public sealed partial class PreferencesPanel : UserControl
{
    private readonly AppSettings _settings;
    private readonly List<int> _qualities;

    internal PreferencesPanel(AppSettings settings)
    {
        InitializeComponent();
        _settings = settings;

        ToolsBox.Text = AppSettings.ResolveAppPath(settings.ToolsPath);
        LogToggle.IsOn = settings.WriteLog;
        OutputToggle.IsOn = settings.ShowOutput;
        ThreadsBox.Items.Add($"Use all ({AppSettings.CoreCount})");
        for (var t = 1; t <= AppSettings.CoreCount; t++) ThreadsBox.Items.Add(t.ToString());
        ThreadsBox.SelectedIndex = settings.Threads;   // index 0 = use all, index n = n threads

        // Quality presets; a value set by hand in the settings file is kept as a "Custom" entry.
        _qualities = AppSettings.QualityProfiles.Select(p => p.Quality).ToList();
        foreach (var (name, _) in AppSettings.QualityProfiles) QualityBox.Items.Add(name);
        if (!_qualities.Contains(settings.JpegQuality))
        {
            _qualities.Add(settings.JpegQuality);
            QualityBox.Items.Add($"Custom (JPEG {settings.JpegQuality})");
        }
        QualityBox.SelectedIndex = _qualities.IndexOf(settings.JpegQuality);
    }

    private void BrowseTools_Click(object sender, RoutedEventArgs e)
    {
        var hwnd = Win32Interop.GetWindowFromWindowId(XamlRoot.ContentIslandEnvironment.AppWindowId);
        var path = ShellDialog.PickFolder(hwnd, "Select the portable tools folder", ToolsBox.Text.Trim().Trim('"'));
        if (path != null) ToolsBox.Text = path;
    }

    /// <summary>Copies the form into the settings. False (with the reason shown) if a value is invalid.</summary>
    public bool TryApply()
    {
        string toolsPath;
        try
        {
            if (string.IsNullOrWhiteSpace(ToolsBox.Text)) return ShowError("Please choose a tools folder.");
            toolsPath = AppSettings.ToAppPath(ToolsBox.Text);
        }
        catch (Exception ex)
        {
            return ShowError($"Invalid tools folder: {ex.Message}");
        }
        _settings.ToolsPath = toolsPath;
        _settings.WriteLog = LogToggle.IsOn;
        _settings.ShowOutput = OutputToggle.IsOn;
        _settings.Threads = Math.Max(0, ThreadsBox.SelectedIndex);
        if (QualityBox.SelectedIndex >= 0) _settings.JpegQuality = _qualities[QualityBox.SelectedIndex];
        return true;
    }

    private bool ShowError(string message)
    {
        ErrorBar.Message = message;
        ErrorBar.IsOpen = true;
        ToolsBox.Focus(FocusState.Programmatic);
        return false;
    }
}
