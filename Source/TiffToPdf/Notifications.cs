using System.Diagnostics;
using Microsoft.Windows.AppNotifications;
using Microsoft.Windows.AppNotifications.Builder;
using Microsoft.Windows.BadgeNotifications;

namespace TiffToPdf;

/// <summary>
/// Taskbar badge while converting, and the "finished" notification with its "Open folder" button.
/// Failures here are ignored: they must never stop a conversion.
/// </summary>
internal static class Notifications
{
    private const string ActionKey = "action";
    private const string OpenFolderAction = "openFolder";
    private const string PathKey = "path";

    private static bool _registered;

    /// <summary>Registers the app for notifications. Call once at startup, after subscribing to NotificationInvoked.</summary>
    public static void Register()
    {
        try
        {
            AppNotificationManager.Default.Register();
            _registered = true;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"App notifications unavailable: {ex.Message}");
        }
    }

    public static void SetBusyBadge(bool busy)
    {
        try
        {
            if (busy) BadgeNotificationManager.Current.SetBadgeAsGlyph(BadgeNotificationGlyph.Activity);
            else BadgeNotificationManager.Current.ClearBadge();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Badge not set: {ex.Message}");
        }
    }

    /// <param name="openPath">Folder to open, or a file to show selected in its folder. Null = no button.</param>
    public static void Show(string title, string message, string? openPath)
    {
        if (!_registered) return;
        try
        {
            var builder = new AppNotificationBuilder().AddText(title).AddText(message);
            if (openPath != null)
                builder.AddButton(new AppNotificationButton("Open folder")
                    .AddArgument(ActionKey, OpenFolderAction)
                    .AddArgument(PathKey, openPath));
            AppNotificationManager.Default.Show(builder.BuildNotification());
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Notification not shown: {ex.Message}");
        }
    }

    /// <summary>Carries out a notification button's action. Returns false if it wasn't one of ours.</summary>
    public static bool HandleInvoked(AppNotificationActivatedEventArgs args)
    {
        if (!args.Arguments.TryGetValue(ActionKey, out var action) || action != OpenFolderAction ||
            !args.Arguments.TryGetValue(PathKey, out var path))
            return false;
        OpenInExplorer(path);
        return true;
    }

    private static void OpenInExplorer(string path)
    {
        try
        {
            var explorer = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "explorer.exe");
            if (File.Exists(path))
                Process.Start(explorer, $"/select,\"{path}\"");
            else if (Directory.Exists(path))
                Process.Start(explorer, $"\"{path}\"");
            else if (Path.GetDirectoryName(path) is { } dir && Directory.Exists(dir))
                Process.Start(explorer, $"\"{dir}\"");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Could not open Explorer: {ex.Message}");
        }
    }
}
