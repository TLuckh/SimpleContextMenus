using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace SimpleContextMenus.Interop;

public abstract class ShellLink
{
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("000214F9-0000-0000-C000-000000000046")]          // GUID von IShellLink
    private interface IShellLink
    {
        void GetPath([Out] StringBuilder pszFile, int cch, IntPtr pfd, uint fFlags);
        void GetIDList(out IntPtr ppidl);
        void SetIDList(IntPtr pidl);
        void GetDescription([Out] StringBuilder pszName, int cch);
        void SetDescription(string pszName);
        void GetWorkingDirectory([Out] StringBuilder pszDir, int cch);
        void SetWorkingDirectory(string pszDir);
        void GetArguments([Out] StringBuilder pszArgs, int cch);
        void SetArguments(string pszArgs);
        void GetHotkey(out short pwHotkey);
        void SetHotkey(short wHotkey);
        void GetShowCmd(out int piShowCmd);
        void SetShowCmd(int iShowCmd);
        void GetIconLocation([Out] StringBuilder pszIconPath, int cch, out int piIcon);
        void SetIconLocation(string pszIconPath, int iIcon);
        void SetRelativePath(string pszPathRel, uint dwReserved);
        void Resolve(IntPtr hwnd, uint fFlags);
        void SetPath(string pszFile);
    }

    [ComImport]
    [Guid("00021401-0000-0000-C000-000000000046")]  // GUID von ShellLink
    private class _ShellLink
    {
    }

    public static string GetShortcutTarget(string shortcutPath)
    {
        IShellLink link = (IShellLink)new _ShellLink();
        ((IPersistFile)link).Load(shortcutPath, 0);

        StringBuilder targetPath = new StringBuilder(260);
        link.GetPath(targetPath, targetPath.Capacity, IntPtr.Zero, 0);
        return targetPath.ToString();
    }
}
