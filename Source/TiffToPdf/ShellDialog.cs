using System.Runtime.InteropServices;

namespace TiffToPdf;

/// <summary>
/// The standard Windows "Open" / "Select Folder" dialog (IFileOpenDialog). Used instead of the WinRT
/// pickers because it can open at the current folder and also works when the app runs elevated.
/// </summary>
internal static class ShellDialog
{
    private const uint FOS_PICKFOLDERS = 0x20;
    private const uint FOS_FORCEFILESYSTEM = 0x40;
    private const uint FOS_FILEMUSTEXIST = 0x1000;
    private const uint SIGDN_FILESYSPATH = 0x80058000;
    private const int ERROR_CANCELLED = unchecked((int)0x800704C7);
    private const int E_NOTIMPL = unchecked((int)0x80004001);
    private const uint SelectFolderButtonId = 1;

    public static string? PickFolder(IntPtr owner, string title, string? initialFolder) =>
        Show(owner, title, initialFolder, FOS_PICKFOLDERS, null, null);

    /// <summary>
    /// A file dialog with an extra button that returns a folder instead: the subfolder highlighted in the
    /// list, or else the folder being shown. Windows has no built-in "file or folder" dialog.
    /// </summary>
    public static string? PickFileOrFolder(IntPtr owner, string title, string? initialFolder, string folderButtonText,
                                           params (string Name, string Spec)[] filters) =>
        Show(owner, title, initialFolder, FOS_FILEMUSTEXIST,
             filters.Select(f => new COMDLG_FILTERSPEC { pszName = f.Name, pszSpec = f.Spec }).ToArray(), folderButtonText);

    private static string? Show(IntPtr owner, string title, string? initialFolder, uint extraOptions,
                                COMDLG_FILTERSPEC[]? filters, string? folderButtonText)
    {
        var dialog = (IFileOpenDialog)new FileOpenDialogRcw();
        FolderButtonEvents? events = null;
        uint cookie = 0;
        try
        {
            dialog.GetOptions(out var options);
            dialog.SetOptions(options | FOS_FORCEFILESYSTEM | extraOptions);
            dialog.SetTitle(title);
            if (filters is { Length: > 0 }) dialog.SetFileTypes((uint)filters.Length, filters);

            if (!string.IsNullOrWhiteSpace(initialFolder) && Directory.Exists(initialFolder))
            {
                var iid = typeof(IShellItem).GUID;
                if (SHCreateItemFromParsingName(initialFolder, IntPtr.Zero, ref iid, out var folder) == 0)
                    dialog.SetFolder(folder);
            }

            if (folderButtonText != null)
            {
                ((IFileDialogCustomize)dialog).AddPushButton(SelectFolderButtonId, folderButtonText);
                events = new FolderButtonEvents(dialog);
                dialog.Advise(events, out cookie);
            }

            var hr = dialog.Show(owner);
            if (events?.SelectedFolder != null) return events.SelectedFolder;
            if (hr == ERROR_CANCELLED) return null;
            Marshal.ThrowExceptionForHR(hr);

            dialog.GetResult(out var item);
            item.GetDisplayName(SIGDN_FILESYSPATH, out var path);
            return path;
        }
        finally
        {
            if (events != null) dialog.Unadvise(cookie);
            Marshal.ReleaseComObject(dialog);
        }
    }

    private static string? FileSystemPath(IShellItem item)
    {
        try
        {
            item.GetDisplayName(SIGDN_FILESYSPATH, out var path);
            return path;
        }
        catch (COMException)
        {
            return null;   // not a file system item (e.g. "This PC")
        }
    }

    /// <summary>Handles the "select folder" button: records the folder and closes the dialog.</summary>
    [ComVisible(true)]
    private sealed class FolderButtonEvents(IFileOpenDialog dialog) : IFileDialogEvents, IFileDialogControlEvents
    {
        public string? SelectedFolder { get; private set; }

        public int OnButtonClicked(IntPtr pfdc, uint dwIDCtl)
        {
            if (dwIDCtl != SelectFolderButtonId) return 0;

            string? folder = null;
            try
            {
                // A highlighted subfolder wins; otherwise the folder being shown.
                dialog.GetCurrentSelection(out var selected);
                var path = FileSystemPath(selected);
                if (path != null && Directory.Exists(path)) folder = path;
            }
            catch (COMException)
            {
                // Nothing selected.
            }
            if (folder == null)
            {
                try
                {
                    dialog.GetFolder(out var current);
                    folder = FileSystemPath(current);
                }
                catch (COMException)
                {
                }
            }
            if (folder == null || !Directory.Exists(folder)) return 0;   // e.g. "This PC": keep the dialog open

            SelectedFolder = folder;
            dialog.Close(ERROR_CANCELLED);
            return 0;
        }

        public int OnItemSelected(IntPtr pfdc, uint dwIDCtl, uint dwIDItem) => E_NOTIMPL;
        public int OnCheckButtonToggled(IntPtr pfdc, uint dwIDCtl, int bChecked) => E_NOTIMPL;
        public int OnControlActivating(IntPtr pfdc, uint dwIDCtl) => E_NOTIMPL;

        public int OnFileOk(IntPtr pfd) => 0;
        public int OnFolderChanging(IntPtr pfd, IntPtr psiFolder) => 0;
        public int OnFolderChange(IntPtr pfd) => 0;
        public int OnSelectionChange(IntPtr pfd) => 0;
        public int OnShareViolation(IntPtr pfd, IntPtr psi, out int pResponse) { pResponse = 0; return E_NOTIMPL; }
        public int OnTypeChange(IntPtr pfd) => 0;
        public int OnOverwrite(IntPtr pfd, IntPtr psi, out int pResponse) { pResponse = 0; return E_NOTIMPL; }
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHCreateItemFromParsingName(string pszPath, IntPtr pbc, ref Guid riid,
        [MarshalAs(UnmanagedType.Interface)] out IShellItem ppv);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct COMDLG_FILTERSPEC
    {
        [MarshalAs(UnmanagedType.LPWStr)] public string pszName;
        [MarshalAs(UnmanagedType.LPWStr)] public string pszSpec;
    }

    [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
    private class FileOpenDialogRcw { }

    // IModalWindow + IFileDialog methods, in vtable order (only up to Close is used).
    [ComImport, Guid("d57c7288-d4ad-4768-be02-9d969532d960"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileOpenDialog
    {
        [PreserveSig] int Show(IntPtr hwndOwner);
        void SetFileTypes(uint cFileTypes, [MarshalAs(UnmanagedType.LPArray)] COMDLG_FILTERSPEC[] rgFilterSpec);
        void SetFileTypeIndex(uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise([MarshalAs(UnmanagedType.Interface)] IFileDialogEvents pfde, out uint pdwCookie);
        void Unadvise(uint dwCookie);
        void SetOptions(uint fos);
        void GetOptions(out uint pfos);
        void SetDefaultFolder(IShellItem psi);
        void SetFolder(IShellItem psi);
        void GetFolder(out IShellItem ppsi);
        void GetCurrentSelection(out IShellItem ppsi);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void GetResult(out IShellItem ppsi);
        void AddPlace(IShellItem psi, int fdap);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        void Close(int hr);
    }

    // Only the methods up to AddPushButton are used.
    [ComImport, Guid("e6fdd21a-163f-4975-9c8c-a69f1ba37034"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileDialogCustomize
    {
        void EnableOpenDropDown(uint dwIDCtl);
        void AddMenu(uint dwIDCtl, [MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void AddPushButton(uint dwIDCtl, [MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
    }

    [ComImport, Guid("973510DB-7D7F-452B-8975-74A85828D354"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileDialogEvents
    {
        [PreserveSig] int OnFileOk(IntPtr pfd);
        [PreserveSig] int OnFolderChanging(IntPtr pfd, IntPtr psiFolder);
        [PreserveSig] int OnFolderChange(IntPtr pfd);
        [PreserveSig] int OnSelectionChange(IntPtr pfd);
        [PreserveSig] int OnShareViolation(IntPtr pfd, IntPtr psi, out int pResponse);
        [PreserveSig] int OnTypeChange(IntPtr pfd);
        [PreserveSig] int OnOverwrite(IntPtr pfd, IntPtr psi, out int pResponse);
    }

    [ComImport, Guid("36116642-D713-4b97-9B83-7484A9D00433"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileDialogControlEvents
    {
        [PreserveSig] int OnItemSelected(IntPtr pfdc, uint dwIDCtl, uint dwIDItem);
        [PreserveSig] int OnButtonClicked(IntPtr pfdc, uint dwIDCtl);
        [PreserveSig] int OnCheckButtonToggled(IntPtr pfdc, uint dwIDCtl, int bChecked);
        [PreserveSig] int OnControlActivating(IntPtr pfdc, uint dwIDCtl);
    }

    [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem
    {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
    }
}
