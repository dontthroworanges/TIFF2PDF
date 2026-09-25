using System.Text.Json;
using System.Text.Json.Serialization;

namespace TiffToPdf;

/// <summary>
/// Preferences stored in Convert-TiffToSearchablePdf.settings.json next to the exe (the same file the
/// PowerShell GUI uses), so the folder stays portable. Paths inside the app folder are stored relative.
/// </summary>
internal sealed class AppSettings
{
    public const string FileName = "Convert-TiffToSearchablePdf.settings.json";

    // ProcessPath is the real exe location, even for a single-file build that unpacks elsewhere.
    public static string AppDir { get; } = Path.GetDirectoryName(Environment.ProcessPath) ?? AppContext.BaseDirectory;
    public static string FilePath => Path.Combine(AppDir, FileName);
    public static int CoreCount => Environment.ProcessorCount;

    public string ToolsPath { get; set; } = "Tools";
    public string InputPath { get; set; } = "";
    public bool WriteLog { get; set; } = true;

    /// <summary>Whether the output section is shown when the app starts.</summary>
    public bool ShowOutput { get; set; } = true;

    /// <summary>Files converted at once. 0 = one per CPU core.</summary>
    public int Threads { get; set; }

    /// <summary>JPEG quality for grayscale/color pages (50-100). 0 = lossless (Flate) instead of JPEG.</summary>
    public int JpegQuality { get; set; } = DefaultJpegQuality;

    public const int DefaultJpegQuality = 90;

    /// <summary>The choices offered in Preferences.</summary>
    public static readonly (string Name, int Quality)[] QualityProfiles =
    [
        ("Maximum quality (JPEG 95)", 95),
        ("High quality (JPEG 90) - default", 90),
        ("Balanced (JPEG 80)", 80),
        ("Smallest files (JPEG 70)", 70),
        ("Lossless (no JPEG - much larger files)", 0),
    ];

    [JsonIgnore]
    public string ThreadsText =>Threads == 0 ? $"All ({CoreCount})" : Threads.ToString();

    [JsonIgnore]
    public string QualityText => JpegQuality == 0 ? "Lossless" : $"JPEG {JpegQuality}";

    public static AppSettings Load(out string? error)
    {
        error = null;
        var s = new AppSettings();
        if (File.Exists(FilePath))
        {
            try
            {
                // Read leniently so a hand-edited file doesn't stop the app starting.
                using var doc = JsonDocument.Parse(File.ReadAllText(FilePath));
                foreach (var p in doc.RootElement.EnumerateObject())
                {
                    var v = p.Value;
                    switch (p.Name.ToLowerInvariant())
                    {
                        case "toolspath" when v.ValueKind == JsonValueKind.String: s.ToolsPath = v.GetString()!; break;
                        case "inputpath" when v.ValueKind == JsonValueKind.String: s.InputPath = v.GetString()!; break;
                        case "writelog" when v.ValueKind is JsonValueKind.True or JsonValueKind.False: s.WriteLog = v.GetBoolean(); break;
                        case "showoutput" when v.ValueKind is JsonValueKind.True or JsonValueKind.False: s.ShowOutput = v.GetBoolean(); break;
                        case "threads" when v.ValueKind == JsonValueKind.Number && v.TryGetInt32(out var t): s.Threads = t; break;
                        case "jpegquality" when v.ValueKind == JsonValueKind.Number && v.TryGetInt32(out var q): s.JpegQuality = q; break;
                    }
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }
        }
        if (string.IsNullOrWhiteSpace(s.ToolsPath)) s.ToolsPath = "Tools";
        s.Threads = Math.Clamp(s.Threads, 0, CoreCount);
        if (s.JpegQuality != 0) s.JpegQuality = Math.Clamp(s.JpegQuality, 50, 100);   // the script's range
        return s;
    }

    /// <summary>Saves the settings; returns an error message, or null on success.</summary>
    public string? Save()
    {
        try
        {
            File.WriteAllText(FilePath, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
            return null;
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
    }

    /// <summary>Turns a stored path (possibly relative to the app folder) into a full path.</summary>
    public static string ResolveAppPath(string path)
    {
        path = Environment.ExpandEnvironmentVariables(path.Trim().Trim('"'));
        if (path.Length == 0) return "";
        return Path.GetFullPath(Path.IsPathRooted(path) ? path : Path.Combine(AppDir, path));
    }

    /// <summary>Stores paths inside the app folder as relative paths so the folder can be moved.</summary>
    public static string ToAppPath(string path)
    {
        var full = Path.GetFullPath(path.Trim().Trim('"')).TrimEnd('\\');
        var root = AppDir.TrimEnd('\\');
        if (full.Equals(root, StringComparison.OrdinalIgnoreCase)) return ".";
        if (full.StartsWith(root + "\\", StringComparison.OrdinalIgnoreCase)) return full[(root.Length + 1)..];
        return full;
    }
}
