using Microsoft.UI.Xaml;
using Microsoft.Windows.AppLifecycle;
using Microsoft.Windows.AppNotifications;

namespace TiffToPdf;

public partial class App : Application
{
    private Window? _window;

    public App()
    {
        InitializeComponent();
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        // Notification clicks while the app is running arrive here (on a background thread).
        AppNotificationManager.Default.NotificationInvoked += (_, e) =>
        {
            if (!Notifications.HandleInvoked(e))
                _window?.DispatcherQueue.TryEnqueue(() => _window.Activate());   // body clicked: bring the app forward
        };
        Notifications.Register();

        // Started by clicking "Open folder" after the app was closed: open the folder and quit.
        if (AppInstance.GetCurrent().GetActivatedEventArgs() is { Kind: ExtendedActivationKind.AppNotification } activation &&
            Notifications.HandleInvoked((AppNotificationActivatedEventArgs)activation.Data))
        {
            Exit();
            return;
        }

        _window = new MainWindow();
        _window.Activate();
    }
}
