using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

namespace TiffToPdf;

/// <summary>The About box contents, shown in a <see cref="DialogWindow"/>.</summary>
public sealed partial class AboutPanel : UserControl
{
    // Placeholder until the project is published; replace with the repository URL.
    private const string GitHubUrl = "https://github.com/";

    public AboutPanel()
    {
        InitializeComponent();

        // Version and release date come from TiffToPdf.csproj (<Version> and <ReleaseDate>).
        var assembly = Assembly.GetExecutingAssembly();
        var version = assembly.GetName().Version ?? new Version(1, 0);
        var releaseDate = assembly.GetCustomAttributes<AssemblyMetadataAttribute>()
            .FirstOrDefault(a => a.Key == "ReleaseDate")?.Value;

        NameText.Text = MainWindow.AppTitle;
        VersionText.Text = $"Version {version.ToString(version.Build > 0 ? 3 : 2)}";
        ReleaseText.Text = DateTime.TryParseExact(releaseDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)
            ? $"Released {date:MMMM d, yyyy}"
            : "";
    }

    private void GitHub_Click(object sender, RoutedEventArgs e)
    {
        try { Process.Start(new ProcessStartInfo(GitHubUrl) { UseShellExecute = true }); }
        catch (Exception) { /* no browser available */ }
    }
}
