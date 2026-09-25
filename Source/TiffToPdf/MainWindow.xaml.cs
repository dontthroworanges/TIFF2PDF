using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Animation;
using Windows.ApplicationModel.DataTransfer;
using Windows.Graphics;
using WinRT.Interop;

namespace TiffToPdf;

public sealed class LogLine(string text, Brush foreground)
{
    public string Text { get; } = text;
    public Brush Foreground { get; } = foreground;
}

/// <summary>
/// Runs Convert-TiffToSearchablePdf.ps1 (next to the exe) in a hidden Windows PowerShell process
/// and shows its output. All conversion logic stays in the script.
/// </summary>
public sealed partial class MainWindow : Window
{
    private const string ScriptName = "Convert-TiffToSearchablePdf.ps1";
    internal const string AppTitle = "TIFF2PDF";

    // "Converting 3 file(s) using 8 thread(s) (3 file(s) at a time)..."
    private static readonly Regex ConvertingRx = new(@"^Converting (\d+) file\(s\) using \d+ thread\(s\) \((\d+) file", RegexOptions.Compiled);
    private static readonly Regex ProgressRx = new(@"^\[(\d+)/(\d+)\] ", RegexOptions.Compiled);
    private static readonly Regex SummaryRx = new(@"^(\S.*?)\s+(\d+)$", RegexOptions.Compiled);   // "Success      12"

    private readonly AppSettings _settings;
    private readonly ObservableCollection<LogLine> _lines = [];
    private readonly ConcurrentQueue<(string Text, bool IsError)> _pending = new();
    private readonly DispatcherQueueTimer _flushTimer;
    private readonly IntPtr _hwnd;

    private Process? _proc;
    private bool _cancelled;
    private bool _closeAfterStop;
    // Progress state for the current run
    private const string StatusPrefix = "##STATUS|";
    private readonly Dictionary<string, string> _active = new(StringComparer.OrdinalIgnoreCase);   // file -> current step
    private string _phase = "";
    private bool _converting;
    private int _completed;
    private int _total;
    private int _runWorkers = 1;   // files the script converts at a time in this run
    // Run results, for the "finished" notification
    private string _runInput = "";
    private bool _runIsFile;
    private bool _inSummary;
    private bool _noFilesFound;
    private readonly Dictionary<string, int> _statusCounts = new(StringComparer.OrdinalIgnoreCase);
    // Output show/hide animation (window height, in pixels)
    private static readonly TimeSpan OutputAnimDuration = TimeSpan.FromMilliseconds(500);
    private readonly ExponentialEase _outputEase = new() { Exponent = 7, EasingMode = EasingMode.EaseOut };   // decelerate
    private readonly Stopwatch _animClock = new();
    private bool _outputHidden;
    private bool _animating;
    private int _animFrom;
    private int _animTo;
    private int _expandedHeight;   // height to return to when the output is shown again

    public MainWindow()
    {
        InitializeComponent();

        _hwnd = WindowNative.GetWindowHandle(this);
        ExtendsContentIntoTitleBar = true;
        SetTitleBar(AppTitleBar);
        AppWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets", "app.ico"));
        SizeAndCenter(960, 700);
        AppWindow.Closing += AppWindow_Closing;
        AppWindow.Changed += AppWindow_Changed;

        LogList.ItemsSource = _lines;
        _flushTimer = DispatcherQueue.CreateTimer();
        _flushTimer.Interval = TimeSpan.FromMilliseconds(100);
        _flushTimer.Tick += (_, _) => FlushOutput();

        _settings = AppSettings.Load(out var loadError);
        InputBox.Text = _settings.InputPath;
        UpdateSummary();

        if (!_settings.ShowOutput)
        {
            // Start collapsed; the window is shrunk once the layout (and so the collapsed height) is known.
            _outputHidden = true;
            UpdateOutputToggle();
            OutputPanel.Visibility = Visibility.Collapsed;
            Root.Loaded += (_, _) => CollapseWindowAtStartup();
        }

        if (loadError != null)
            Root.Loaded += async (_, _) => await ShowMessageAsync("Preferences not loaded",
                $"Could not read the preferences file; defaults are being used.\n\n{AppSettings.FilePath}\n\n{loadError}");
    }

    #region Window

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hwnd);

    private void SizeAndCenter(int width, int height)
    {
        var scale = GetDpiForWindow(_hwnd) / 96.0;
        var size = new SizeInt32((int)(width * scale), (int)(height * scale));
        var area = DisplayArea.GetFromWindowId(AppWindow.Id, DisplayAreaFallback.Nearest).WorkArea;
        size.Width = Math.Min(size.Width, area.Width);
        size.Height = Math.Min(size.Height, area.Height);
        AppWindow.MoveAndResize(new RectInt32(
            area.X + (area.Width - size.Width) / 2, area.Y + (area.Height - size.Height) / 2, size.Width, size.Height));
    }

    private async void AppWindow_Closing(AppWindow sender, AppWindowClosingEventArgs args)
    {
        if (!IsRunning) return;
        args.Cancel = true;
        if (!await ConfirmAsync("Conversion still running", "Stop the conversion and exit?", "Stop and exit", "Keep running"))
            return;
        if (IsRunning)
        {
            _closeAfterStop = true;
            StopConversion();
        }
        else
        {
            Close();
        }
    }

    #endregion

    #region Dialogs

    // Dialogs are separate windows so they aren't clipped when the main window is small or collapsed.
    private static TextBlock MessageText(string message) =>
        new() { Text = message, TextWrapping = TextWrapping.Wrap, IsTextSelectionEnabled = true };

    private Task ShowMessageAsync(string title, string message) =>
        new DialogWindow(title, MessageText(message), null, "OK").ShowAsync(this);

    private Task<bool> ConfirmAsync(string title, string message, string yes, string no) =>
        new DialogWindow(title, MessageText(message), yes, no).ShowAsync(this);

    private async void SaveSettings()
    {
        var error = _settings.Save();
        if (error != null)
            await ShowMessageAsync("Preferences not saved", $"Could not save preferences to:\n{AppSettings.FilePath}\n\n{error}");
    }

    #endregion

    #region Inputs and preferences

    private void UpdateSummary()
    {
        // Tools path last so it's the part that gets trimmed in a narrow window.
        SummaryText.Text = $"Quality: {_settings.QualityText}   ·   CPU threads: {_settings.ThreadsText}   ·   " +
                           $"Log: {(_settings.WriteLog ? "On" : "Off")}   ·   " +
                           $"Tools: {AppSettings.ResolveAppPath(_settings.ToolsPath)}";
    }

    /// <summary>The folder the pickers open at: the input box's folder, or the folder of the file in it.</summary>
    private string? InputStartFolder()
    {
        var current = InputBox.Text.Trim().Trim('"');
        if (current.Length == 0) return null;
        try { return File.Exists(current) ? Path.GetDirectoryName(current) : current; }
        catch { return null; }
    }

    private void BrowseInput_Click(object sender, RoutedEventArgs e)
    {
        var path = ShellDialog.PickFileOrFolder(_hwnd, "Select a TIFF file, or a folder to process with all its subfolders",
                                                InputStartFolder(), "Select this folder",
                                                ("TIFF images", "*.tif;*.tiff"), ("All files", "*.*"));
        if (path != null) InputBox.Text = path;
    }

    private async void About_Click(object sender, RoutedEventArgs e) =>
        await new DialogWindow("About", new AboutPanel(), null, "Close", width: 420).ShowAsync(this);

    private async void Preferences_Click(object sender, RoutedEventArgs e)
    {
        var panel = new PreferencesPanel(_settings);
        var dialog = new DialogWindow("Preferences", panel, "Save", "Cancel", width: 520) { PrimaryButtonClick = panel.TryApply };
        if (!await dialog.ShowAsync(this)) return;
        SaveSettings();
        UpdateSummary();
    }

    /// <summary>Same lookup as the script: the tools folder (2 levels deep), then PATH.</summary>
    private static bool ToolsPresent(string toolsDir)
    {
        var options = new EnumerationOptions { RecurseSubdirectories = true, MaxRecursionDepth = 2, IgnoreInaccessible = true };
        var pathDirs = (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries);
        foreach (var exe in new[] { "magick.exe", "tesseract.exe" })
        {
            var found = Directory.Exists(toolsDir) && Directory.EnumerateFiles(toolsDir, exe, options).Any();
            found = found || pathDirs.Any(d =>
            {
                try { return File.Exists(Path.Combine(d.Trim('"'), exe)); } catch { return false; }
            });
            if (!found) return false;
        }
        return true;
    }

    #endregion

    #region Conversion

    private bool IsRunning => _proc != null;

    /// <summary>One button: Start when idle, Cancel while a conversion is running.</summary>
    private async void StartCancel_Click(object sender, RoutedEventArgs e)
    {
        if (IsRunning) StopConversion();
        else await StartConversionAsync();
    }

    private async Task StartConversionAsync()
    {
        var input = InputBox.Text.Trim().Trim('"');
        var isFile = input.Length > 0 && File.Exists(input);
        if (input.Length == 0 || !(isFile || Directory.Exists(input)))
        {
            await ShowMessageAsync("Choose a folder or file", "Please choose an existing folder or TIFF file to process.");
            return;
        }
        if (isFile && !(input.EndsWith(".tif", StringComparison.OrdinalIgnoreCase) ||
                        input.EndsWith(".tiff", StringComparison.OrdinalIgnoreCase)))
        {
            await ShowMessageAsync("Not a TIFF file", "Please choose a .tif or .tiff file.");
            return;
        }
        var script = Path.Combine(AppSettings.AppDir, ScriptName);
        if (!File.Exists(script))
        {
            await ShowMessageAsync("Script not found", $"{ScriptName} must be in the same folder as this program:\n{AppSettings.AppDir}");
            return;
        }
        input = Path.GetFullPath(input);
        var tools = AppSettings.ResolveAppPath(_settings.ToolsPath);

        var download = false;
        if (!ToolsPresent(tools))
        {
            if (!await ConfirmAsync("Missing tools",
                    $"ImageMagick and/or Tesseract OCR were not found in:\n{tools}\n\n" +
                    "Download the portable versions there now? (about 500 MB, no install or admin rights needed)",
                    "Download", "Cancel"))
                return;
            download = true;
        }

        _settings.InputPath = input;
        SaveSettings();

        var psi = new ProcessStartInfo
        {
            FileName = Path.Combine(Environment.SystemDirectory, @"WindowsPowerShell\v1.0\powershell.exe"),
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            WorkingDirectory = AppSettings.AppDir,
        };
        // -File with separate arguments: .NET handles the quoting. (-EncodedCommand would make Windows
        // PowerShell also write every Write-Host message to stderr as CLIXML.)
        foreach (var arg in new[] { "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", script,
                                    "-InputPath", input, "-ToolsPath", tools, "-Threads", _settings.Threads.ToString() })
            psi.ArgumentList.Add(arg);
        psi.ArgumentList.Add("-ReportStatus");
        if (_settings.JpegQuality == 0)
        {
            psi.ArgumentList.Add("-LosslessColor");
        }
        else
        {
            psi.ArgumentList.Add("-JpegQuality");
            psi.ArgumentList.Add(_settings.JpegQuality.ToString());
        }
        if (!_settings.WriteLog) psi.ArgumentList.Add("-NoLog");
        if (download) psi.ArgumentList.Add("-DownloadMissingTools");

        _lines.Clear();
        AddLine($"Started {DateTime.Now:g}", LineKind.Header);
        AddLine(isFile ? $"File:    {input}" : $"Folder:  {input} (including subfolders)", LineKind.Header);
        AddLine($"Threads: {_settings.ThreadsText}", LineKind.Header);
        AddLine($"Quality: {_settings.QualityText}", LineKind.Header);
        AddLine("", LineKind.Normal);

        var proc = new Process { StartInfo = psi, EnableRaisingEvents = true };
        proc.OutputDataReceived += (_, a) => { if (a.Data != null) _pending.Enqueue((a.Data, false)); };
        proc.ErrorDataReceived += (_, a) => { if (a.Data != null) _pending.Enqueue((a.Data, true)); };
        proc.Exited += (_, _) => DispatcherQueue.TryEnqueue(OnProcessExited);
        try
        {
            proc.Start();
        }
        catch (Exception ex)
        {
            proc.Dispose();
            await ShowMessageAsync("Could not start", $"Could not start Windows PowerShell:\n{ex.Message}");
            return;
        }
        proc.BeginOutputReadLine();
        proc.BeginErrorReadLine();

        _proc = proc;
        _cancelled = false;
        _runInput = input;
        _runIsFile = isFile;
        SetRunning(true);
        _flushTimer.Start();
    }

    private void StopConversion()
    {
        if (_proc == null) return;
        _cancelled = true;
        StatusText.Text = "Cancelling…";
        StartButton.IsEnabled = false;   // re-enabled by SetRunning(false) once the process has exited
        try
        {
            // Kills PowerShell and the ImageMagick/Tesseract processes it started.
            _proc.Kill(entireProcessTree: true);
        }
        catch (Exception)
        {
            // Already exited.
        }
    }

    private async void OnProcessExited()
    {
        var proc = _proc;
        if (proc == null) return;

        // Exited can fire before the last output lines arrive; this waits for the streams to close.
        await Task.Run(() => proc.WaitForExit());
        _flushTimer.Stop();
        FlushOutput();

        var code = proc.ExitCode;
        proc.Dispose();
        _proc = null;

        AddLine("", LineKind.Normal);
        if (_cancelled)
        {
            AddLine("---- Cancelled ----", LineKind.Warning);
            StatusText.Text = "Cancelled";
        }
        else if (code == 0)
        {
            AddLine($"---- Finished {DateTime.Now:g} ----", LineKind.Success);
            StatusText.Text = "Finished";
        }
        else
        {
            AddLine($"---- Stopped with an error (exit code {code}) ----", LineKind.Error);
            StatusText.Text = "Stopped with an error - see the output below";
        }
        SetRunning(false);

        if (_closeAfterStop) Close();
        else if (!_cancelled) ShowFinishedNotification(code);
    }

    private void ShowFinishedNotification(int exitCode)
    {
        int Count(string status) => _statusCounts.GetValueOrDefault(status);
        var ocrFailed = Count("Converted - OCR failed");
        var converted = Count("Success") + ocrFailed;
        var skipped = Count("Skipped");
        var failed = Count("Failed");

        // Button target: the processed folder, or the new PDF selected in its folder.
        var openPath = _runInput;
        if (_runIsFile)
        {
            var pdf = Path.ChangeExtension(_runInput, ".pdf");
            openPath = File.Exists(pdf) ? pdf : Path.GetDirectoryName(_runInput) ?? _runInput;
        }
        var name = Path.GetFileName(_runInput);

        string title, message;
        if (exitCode != 0)
        {
            title = "Conversion stopped with an error";
            message = $"See the output in {AppTitle} for details.";
        }
        else if (_noFilesFound)
        {
            Notifications.Show("No TIFF files found", $"No .tif or .tiff files were found in {name}.", null);
            return;
        }
        else if (failed > 0)
        {
            title = "Conversion finished with errors";
            message = _runIsFile
                ? $"{name} could not be converted. See the output in {AppTitle} for details."
                : $"{converted} file(s) converted, {failed} failed. See the output in {AppTitle} for details.";
        }
        else if (converted == 0 && skipped > 0)
        {
            title = "Nothing to convert";
            message = _runIsFile ? $"A PDF of {name} already exists." : $"All {skipped} file(s) already have PDFs.";
        }
        else
        {
            title = "Conversion complete";
            message = _runIsFile ? $"{name} has been converted." : $"All {converted} file(s) have been converted.";
            if (skipped > 0) message += $" {skipped} skipped (PDF already existed).";
            if (ocrFailed > 0) message += $" {ocrFailed} without searchable text (OCR failed).";
        }
        Notifications.Show(title, message, openPath);
    }

    private void SetRunning(bool running)
    {
        StartButton.IsEnabled = true;
        StartButton.Style = (Style)Application.Current.Resources[running ? "DefaultButtonStyle" : "AccentButtonStyle"];
        StartIcon.Glyph = running ? "" : "";   // Cancel / Play
        StartText.Text = running ? "Cancel" : "Start";
        PrefsButton.IsEnabled = !running;
        BrowseInputButton.IsEnabled = !running;
        InputBox.IsReadOnly = running;
        Progress.Visibility = running ? Visibility.Visible : Visibility.Collapsed;
        Notifications.SetBusyBadge(running);
        if (running)
        {
            _phase = "Checking tools…";
            _converting = false;
            _completed = _total = 0;
            _active.Clear();
            _inSummary = _noFilesFound = false;
            _statusCounts.Clear();
            UpdateProgress();
        }
        AppWindow.Title = running ? $"{AppTitle} - running" : AppTitle;
    }

    private void FlushOutput()
    {
        var added = false;
        var changed = false;
        while (_pending.TryDequeue(out var item))
        {
            var text = item.Text;

            // "##STATUS|<file>|<step>" (from -ReportStatus): which step each file is on. Not shown in the output.
            if (!item.IsError && text.StartsWith(StatusPrefix, StringComparison.Ordinal))
            {
                var parts = text.Split('|', 3);
                if (parts.Length == 3) _active[parts[1]] = parts[2];
                changed = true;
                continue;
            }

            added = true;
            AddLine(text, item.IsError ? LineKind.Error : Classify(text));

            if (!item.IsError)
            {
                // The script's summary: "<status>   <count>" lines after "===== Summary".
                if (_inSummary && SummaryRx.Match(text) is { Success: true } s)
                    _statusCounts[s.Groups[1].Value] = int.Parse(s.Groups[2].Value);
                else if (text.StartsWith("No .tif or .tiff files found", StringComparison.Ordinal))
                    _noFilesFound = true;
            }

            // Progress comes from the script's "Converting N file(s) ... (W file(s) at a time)" and "[k/N] file" lines.
            if (ConvertingRx.Match(text) is { Success: true } c)
            {
                _converting = true;
                _total = int.Parse(c.Groups[1].Value);
                _runWorkers = int.Parse(c.Groups[2].Value);
                changed = true;
            }
            else if (ProgressRx.Match(text) is { Success: true } p)
            {
                var done = int.Parse(p.Groups[1].Value);
                _total = int.Parse(p.Groups[2].Value);
                if (_runWorkers == 1)
                {
                    // Single-threaded runs print this line when a file starts.
                    _completed = done - 1;
                    _active.Clear();
                }
                else
                {
                    // Parallel runs print it when a file finishes.
                    _completed = done;
                    _active.Remove(text[p.Length..]);
                }
                changed = true;
            }
            else if (text.StartsWith("===== Summary", StringComparison.Ordinal))
            {
                _inSummary = true;
                _completed = _total;
                _active.Clear();
                changed = true;
            }
            else if (!_converting)
            {
                // Phases before conversion starts.
                var phase = text.StartsWith("Searching '", StringComparison.Ordinal) ? "Searching for TIFF files…"
                          : text.StartsWith("Missing tool", StringComparison.Ordinal) ? "Downloading tools…"
                          : text.StartsWith("Tesseract language data missing", StringComparison.Ordinal) ? "Downloading OCR language data…"
                          : text.StartsWith("Found ", StringComparison.Ordinal) ? "Checking for existing PDFs…"
                          : null;
                if (phase != null)
                {
                    _phase = phase;
                    changed = true;
                }
            }
        }
        if (changed) UpdateProgress();
        if (added && _lines.Count > 0) LogList.ScrollIntoView(_lines[^1]);
    }

    /// <summary>
    /// Status line and progress bar. The bar keeps its moving (indeterminate) animation until the first
    /// file is finished; the text shows what the files in progress are doing.
    /// </summary>
    private void UpdateProgress()
    {
        if (_cancelled) return;   // keep "Cancelling…"
        if (!_converting)
        {
            Progress.IsIndeterminate = true;
            StatusText.Text = _phase;
            return;
        }

        Progress.IsIndeterminate = _completed == 0;
        Progress.Maximum = Math.Max(1, _total);
        Progress.Value = _completed;

        var text = $"{_completed} of {_total} files done";
        if (_active.Count == 1)
        {
            var (file, step) = _active.First();
            text += $"  ·  {step}: {Path.GetFileName(file)}";
        }
        else if (_active.Count > 1)
        {
            // e.g. "8 in progress: 5 checking page orientation, 3 running OCR"
            static string BaseStep(string step) => step.Split(" (")[0];
            var groups = _active.Values.GroupBy(BaseStep).OrderByDescending(g => g.Count())
                .Select(g => $"{g.Count()} {char.ToLowerInvariant(g.Key[0])}{g.Key[1..]}");
            text += $"  ·  {_active.Count} in progress: {string.Join(", ", groups)}";
        }
        StatusText.Text = text;
    }

    #endregion

    #region Output

    private enum LineKind { Normal, Header, Secondary, Success, Warning, Error }

    private static LineKind Classify(string text)
    {
        var t = text.TrimStart();
        if (t.StartsWith("FAILED", StringComparison.Ordinal)) return LineKind.Error;
        if (t == "Done." || t.StartsWith("Tools ready", StringComparison.Ordinal)) return LineKind.Success;
        if (t.StartsWith("Skipped", StringComparison.Ordinal)) return LineKind.Secondary;
        if (t.StartsWith("Warning", StringComparison.Ordinal) || t.StartsWith("Rotated", StringComparison.Ordinal) ||
            t.Contains("OCR failed", StringComparison.Ordinal) || t.StartsWith("Missing", StringComparison.Ordinal) ||
            t.StartsWith("No .tif", StringComparison.Ordinal))
            return LineKind.Warning;
        if (t.StartsWith('[') || t.StartsWith("=====", StringComparison.Ordinal) ||
            t.StartsWith("Converting", StringComparison.Ordinal)) return LineKind.Header;
        return LineKind.Normal;
    }

    private void AddLine(string text, LineKind kind)
    {
        var key = kind switch
        {
            LineKind.Error => "SystemFillColorCriticalBrush",
            LineKind.Success => "SystemFillColorSuccessBrush",
            LineKind.Warning => "SystemFillColorCautionBrush",
            LineKind.Secondary => "TextFillColorTertiaryBrush",
            LineKind.Header => "TextFillColorPrimaryBrush",
            _ => "TextFillColorSecondaryBrush",
        };
        var brush = Application.Current.Resources.TryGetValue(key, out var b) && b is Brush br
            ? br
            : new SolidColorBrush(Microsoft.UI.Colors.Gray);
        _lines.Add(new LogLine(text, brush));
    }

    private void CopyOutput_Click(object sender, RoutedEventArgs e)
    {
        var package = new DataPackage();
        package.SetText(string.Join(Environment.NewLine, _lines.Select(l => l.Text)));
        Clipboard.SetContent(package);
    }

    private void ToggleOutput_Click(object sender, RoutedEventArgs e) => SetOutputHidden(!_outputHidden);

    /// <summary>
    /// Hides or shows the output section. The window shrinks to end just below the progress row, or
    /// grows back to its previous height, with a decelerating (exponential ease-out) animation.
    /// </summary>
    private void SetOutputHidden(bool hidden)
    {
        if (AppWindow.Presenter is OverlappedPresenter { State: OverlappedPresenterState.Maximized } presenter)
            presenter.Restore();

        _outputHidden = hidden;
        UpdateOutputToggle();
        var current = AppWindow.Size.Height;
        int target;
        if (hidden)
        {
            // Mid-animation the window isn't at its expanded height; keep the one recorded earlier.
            if (!_animating) _expandedHeight = current;
            target = CollapsedWindowHeight();
        }
        else
        {
            OutputPanel.Visibility = Visibility.Visible;
            var scale = Root.XamlRoot.RasterizationScale;
            var area = DisplayArea.GetFromWindowId(AppWindow.Id, DisplayAreaFallback.Nearest).WorkArea;
            target = _expandedHeight > 0 ? _expandedHeight : (int)(700 * scale);
            target = Math.Max(current, Math.Min(target, area.Y + area.Height - AppWindow.Position.Y));
        }

        _animFrom = current;
        _animTo = target;
        _animClock.Restart();
        if (!_animating)
        {
            _animating = true;
            CompositionTarget.Rendering += AnimateWindowHeight;
        }
    }

    private void AnimateWindowHeight(object? sender, object e)
    {
        var t = Math.Min(1.0, _animClock.Elapsed / OutputAnimDuration);
        var height = _animFrom + (_animTo - _animFrom) * _outputEase.Ease(t);
        AppWindow.Resize(new SizeInt32(AppWindow.Size.Width, (int)Math.Round(height)));
        if (t < 1) return;

        CompositionTarget.Rendering -= AnimateWindowHeight;
        _animating = false;
        if (_outputHidden) OutputPanel.Visibility = Visibility.Collapsed;
    }

    /// <summary>Window height (pixels) that ends just below the progress row, keeping the bottom padding.</summary>
    private int CollapsedWindowHeight()
    {
        var scale = Root.XamlRoot.RasterizationScale;
        var client = AppTitleBar.ActualHeight + ContentGrid.Padding.Top + ContentGrid.RowDefinitions[0].ActualHeight +
                     ContentGrid.RowSpacing + ContentGrid.RowDefinitions[1].ActualHeight + ContentGrid.Padding.Bottom;
        var frame = AppWindow.Size.Height - AppWindow.ClientSize.Height;
        return (int)Math.Ceiling(client * scale) + frame;
    }

    /// <summary>Output hidden by preference: open at the collapsed height, centered on the screen.</summary>
    private void CollapseWindowAtStartup()
    {
        _expandedHeight = AppWindow.Size.Height;
        var height = CollapsedWindowHeight();
        var area = DisplayArea.GetFromWindowId(AppWindow.Id, DisplayAreaFallback.Nearest).WorkArea;
        AppWindow.MoveAndResize(new RectInt32(AppWindow.Position.X, area.Y + (area.Height - height) / 2, AppWindow.Size.Width, height));
    }

    private void UpdateOutputToggle()
    {
        OutputToggleIcon.Glyph = _outputHidden ? "" : "";   // ChevronDown / ChevronUp
        OutputToggleText.Text = _outputHidden ? "Show output" : "Hide output";
    }

    private void AppWindow_Changed(AppWindow sender, AppWindowChangedEventArgs args)
    {
        // Resizing or maximizing the window taller while the output is hidden brings it back.
        if (!args.DidSizeChange || !_outputHidden || _animating) return;
        if (AppWindow.Size.Height <= CollapsedWindowHeight() + 8) return;
        _outputHidden = false;
        _expandedHeight = 0;
        OutputPanel.Visibility = Visibility.Visible;
        UpdateOutputToggle();
    }

    #endregion
}
