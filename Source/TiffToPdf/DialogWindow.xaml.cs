using System.Runtime.InteropServices;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Windows.Foundation;
using Windows.Graphics;
using WinRT.Interop;

namespace TiffToPdf;

/// <summary>
/// A modal dialog in its own window. Unlike a ContentDialog it isn't limited to the main window's
/// area, so it stays usable when the main window is small or its output is hidden.
/// </summary>
public sealed partial class DialogWindow : Window
{
    private const int GWLP_HWNDPARENT = -8;

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
    private static extern IntPtr SetWindowLongPtr(IntPtr hwnd, int index, IntPtr value);

    private readonly double _width;
    private readonly bool _primaryIsDefault;
    private readonly TaskCompletionSource<bool> _result = new();
    private Window? _owner;
    private bool _accepted;

    /// <summary>Called when the primary button is clicked; return false to keep the dialog open.</summary>
    public Func<bool>? PrimaryButtonClick { get; set; }

    /// <param name="primaryText">Accent button text, or null for a single close button.</param>
    /// <param name="width">Width in device-independent pixels; the height fits the content.</param>
    public DialogWindow(string title, UIElement content, string? primaryText, string closeText,
                        bool primaryIsDefault = true, double width = 440)
    {
        InitializeComponent();
        _width = width;
        _primaryIsDefault = primaryIsDefault && primaryText != null;

        Title = title;
        TitleText.Text = title;
        Body.Content = content;
        CloseButton.Content = closeText;
        if (primaryText == null) PrimaryButton.Visibility = Visibility.Collapsed;
        else PrimaryButton.Content = primaryText;

        ExtendsContentIntoTitleBar = true;
        SetTitleBar(DialogTitleBar);
        Root.Loaded += Root_Loaded;
        Closed += (_, _) =>
        {
            _result.TrySetResult(_accepted);
            _owner?.Activate();
        };
    }

    /// <summary>Shows the dialog over <paramref name="owner"/>, which is disabled until it closes. True if the primary button was used.</summary>
    public Task<bool> ShowAsync(Window owner)
    {
        _owner = owner;
        // The owner must be set before the presenter is made modal.
        SetWindowLongPtr(WindowNative.GetWindowHandle(this), GWLP_HWNDPARENT, WindowNative.GetWindowHandle(owner));
        var presenter = OverlappedPresenter.CreateForDialog();
        presenter.IsModal = true;
        AppWindow.SetPresenter(presenter);
        var scale = owner.Content?.XamlRoot?.RasterizationScale ?? 1.0;
        PlaceOverOwner(owner, (int)(_width * scale), (int)(300 * scale));   // provisional; resized to fit once loaded
        Activate();
        return _result.Task;
    }

    private void Root_Loaded(object sender, RoutedEventArgs e)
    {
        // Size the window to its content, then center it over the owner.
        var scale = Root.XamlRoot.RasterizationScale;
        Root.Measure(new Size(_width, double.PositiveInfinity));
        var frame = AppWindow.Size.Height - AppWindow.ClientSize.Height;
        if (_owner != null)
            PlaceOverOwner(_owner, (int)Math.Ceiling(_width * scale), (int)Math.Ceiling(Root.DesiredSize.Height * scale) + frame);

        (_primaryIsDefault ? PrimaryButton : CloseButton).Focus(FocusState.Programmatic);
    }

    private void PlaceOverOwner(Window owner, int width, int height)
    {
        var area = DisplayArea.GetFromWindowId(owner.AppWindow.Id, DisplayAreaFallback.Nearest).WorkArea;
        width = Math.Min(width, area.Width);
        height = Math.Min(height, area.Height);
        var pos = owner.AppWindow.Position;
        var size = owner.AppWindow.Size;
        var x = Math.Clamp(pos.X + (size.Width - width) / 2, area.X, area.X + area.Width - width);
        var y = Math.Clamp(pos.Y + (size.Height - height) / 2, area.Y, area.Y + area.Height - height);
        AppWindow.MoveAndResize(new RectInt32(x, y, width, height));
    }

    private void Primary_Click(object sender, RoutedEventArgs e)
    {
        if (PrimaryButtonClick?.Invoke() == false) return;
        _accepted = true;
        Close();
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
