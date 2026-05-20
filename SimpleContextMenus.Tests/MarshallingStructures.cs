using System.Runtime.InteropServices;
using SharpShell.Interop;

namespace SimpleContextMenus.Tests;


/// <summary>
/// Some of the Marshalling methods are instead in <see cref="IContextMenu"/> (and less importantly for the tests, its subclass tree)
/// </summary>
public class MarshallingStructures
{
    /// <summary>
    /// Creates an empty windows context menu.
    /// 
    /// Has to be manually disposed of at the end of its lifetime, using <see cref="DestroyMenu"/>
    /// </summary>
    /// <returns></returns>
    [DllImport("User32.dll")]
    internal static extern IntPtr CreatePopupMenu();


    /// <summary>
    /// 
    /// </summary>
    /// <param name="hMenu">
    /// The hMenu-Handle is easiest to get using CreatePopupMenu();
    /// Most likely followed by
    /// contextMenuInterface.QueryContextMenu(menuHandle, 0, 0, 0x7FFF, 0);
    /// </param>
    /// <param name="item">Which item we want information of. Use GetMenuItemCount() if you're looping for the bounds </param>
    /// <param name="fByPosition">Most likely set to True.</param>
    /// <param name="lpmii">
    /// Has to be partially populated with specific values, because from this the method reads what it is supposed to write back (also into this struct).
    /// See the <see cref="MENUITEMINFO"/> class, which has documentation for all parameters. 
    /// Though for most cases you can do something like this:
    ///   var info = new MENUITEMINFO
    ///   {&#xA;
    ///       cbSize = (uint)Marshal.SizeOf&lt;MENUITEMINFO>(),
    ///       fMask = 64 // MIIM_STRING
    ///               | 256 // MIIM_FTYPE
    ///               | 4, // MIIM_SUBMENU
    ///       dwTypeData = new string('\0', 256),
    ///       cch = 255
    ///   };
    /// </param>
    /// <returns></returns>
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern bool GetMenuItemInfo(
        IntPtr hMenu,
        uint item,
        bool fByPosition,
        ref MENUITEMINFO lpmii);

    /// <summary>
    /// Given a handle to a windows context menu window, returns the number of elements on the main level.
    /// </summary>
    /// <param name="hMenu"></param>
    /// <returns></returns>
    [DllImport("user32.dll")]
    public static extern int GetMenuItemCount(IntPtr hMenu);

    /// <summary>
    /// Given a handle to a windows context menu window, destroys it.
    /// </summary>
    /// <param name="hMenu"></param>
    /// <returns></returns>
    [DllImport("user32.dll")]
    public static extern bool DestroyMenu(IntPtr hMenu);
    
    /// <summary>
    /// Given a file path, returns the PIDL (Pointer to an Item ID List) for that file, which is used in a lot of the Shell API calls.
    /// </summary>
    /// <param name="path"></param>
    /// <returns></returns>
    public static IntPtr GetPIDLFromPath(string path)
    {
        int hr = SHParseDisplayName(path, IntPtr.Zero, out IntPtr pidl, 0, out _);

        if (hr != 0)
            Marshal.ThrowExceptionForHR(hr);

        return pidl;
    }
    
        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHParseDisplayName(
            string pszName,
            IntPtr pbc,
            out IntPtr ppidl,
            uint sfgaoIn,
            out uint psfgaoOut);

        [DllImport("shell32.dll")]
        public static extern IntPtr ILFindLastID(IntPtr pidl);
}

