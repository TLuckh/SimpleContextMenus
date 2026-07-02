using System.Runtime.InteropServices;
using SharpShell.Interop;

namespace SimpleContextMenus;

public class MarshallingStructures
{
    

    /// <summary>
    /// 
    /// </summary>
    /// <param name="hMenu">
    /// Handle to the windows context menu. Should be provided by the calling COM object. Cached in SharpShell.handleWindowsContextMenu.
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
    private static extern bool GetMenuItemInfo(
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
    private static extern int GetMenuItemCount(IntPtr hMenu);

    /// <summary>
    /// Specialization of SimpleContextMenus.Tests::...::ReadMenu(IntPtr) for the case that we just want to check if we've already added our stuff to the context menu.
    /// Works by checking for a menuItem named "Extensions" which has the bitmap from "Resources/Extensions Menu.png"
    /// </summary>
    /// <param name="hMenu">The <see cref="IntPtr"/> Handle for the windows context menu. Get it from SharpContextMenu.handleWindowsContextMenu</param>
    /// <param name="instance">The instance of SimpleContextMenus which generates the menu entries.</param>
    /// <returns>True, if SimpleContextMenu already added stuff to the context menu, false otherwise.</returns>
    public static bool CheckIfSCMAlreadyCalled(IntPtr hMenu, SimpleContextMenu instance)
    {
        for (int i = 0; i < GetMenuItemCount(hMenu); i++)
        {
            var info = new MENUITEMINFO
            {
                cbSize = (uint)Marshal.SizeOf<MENUITEMINFO>(),
                fMask = 64 // MIIM_STRING
                        | 256 // MIIM_FTYPE
                        | 32 // MIIM_DAT 
                        | 4 // MIIM_SUBMENU
                        | 2, // MIIM_ID  ; Sagt GetMenuItemInfo, dass es die ID des Menuitems zurückliefern soll. Später für InvokeCommand notwendig, zum Ausführen des Kontextmenüeintrags
                dwTypeData = new string('\0', 256),
                cch = 255
            };

            if (!GetMenuItemInfo(hMenu, (uint)i, true, ref info))
                continue;

            uint menuItemId = info.wID;

            bool isSeparator = (info.fType & 0x00000800) != 0;
            if (info.hSubMenu != IntPtr.Zero) // case submenu
            {
                if (info.dwTypeData == "Extensions" && (ulong) info.dwItemData == instance.ApplicationGUID) // We pass, when constructing the menu items, 
                    return true;
                
                if (CheckIfSCMAlreadyCalled(info.hSubMenu,instance)) // recursive call
                    return true;
            }
        }

        return false;
    }

}