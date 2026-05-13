using System.Collections.Specialized;
using System.Runtime.InteropServices;
using ServerManager.ShellDebugger;
using SharpShell.Interop;

namespace SimpleContextMenus.Tests;

public class FromDebug
{
    // private void InitializeParentFolder()
    // {
    //     #region CreateRootFolder (Desktop) 
    //     //  Get the dekstop PDIL.
    //     var desktopPIDL = IntPtr.Zero;
    //     var result = Shell32.SHGetFolderLocation(IntPtr.Zero, CSIDL.CSIDL_DESKTOP, IntPtr.Zero, 0, out desktopPIDL);
    //
    //     //  Validate the result.
    //     if (result != 0)
    //     {
    //         //  Throw the failure as an exception.
    //         Marshal.ThrowExceptionForHR(result);
    //     }
    //     
    //     
    //     //  Get the file info.
    //     var fileInfo = new SHFILEINFO();
    //     Shell32.SHGetFileInfo(desktopPIDL, 0, out fileInfo, (uint)Marshal.SizeOf(fileInfo),
    //         SHGFI.SHGFI_DISPLAYNAME | SHGFI.SHGFI_PIDL | SHGFI.SHGFI_SMALLICON | SHGFI.SHGFI_SYSICONINDEX);
    //
    //     //  Return the Shell Folder.
    //     return new ShellItem
    //     {
    //         DisplayName = fileInfo.szDisplayName,
    //         IconIndex = fileInfo.iIcon,
    //         HasSubFolders = true,
    //         IsFolder = true,
    //         ShellFolderInterface = desktopShellFolderInterface,
    //         PIDL = desktopPIDL,
    //         RelativePIDL = desktopPIDL
    //     };
    // }
    // #endregion


    public static void Main()
    {
        // Der Pfad der zu testenden Dateien und Ordner. 
        // TODo: Build von SimpleContextMenus.Tests sollte SimpleContextMenus bauen, und anschließend alle davon erzeugten Dateien in einen neuen Ordner kopieren, wo wir auch die Tests reinhauen (die noch zu schreiben sind)
        // Die Tests selbst sollten eine relativ flache Dateistruktur darstellen, so dass wir diese für jeden Test manuell bauen können: 
        // Es handelt sich einfach um eine Menge von Ordnern mit Dateien drin. Wir simulieren Klicks auf Teilmengen (inkl. leere Teilmenge) dieser Dateien im Ordner, und unser Ziel ist es, jeweils die richtigen Kontextmenüs zu kriegen.
        // Der Test der Kontextmenüs wiederum kommt stattdessen in einen Unit Test (das hier sind Integration Tests i think?)

        // Definition des Testpfades
        using DisposableMap testItemMap = ItemsInTestPath("1FilterBasedOnExtension");
        // Definition der markierten Items
        ShellItem[] items = [testItemMap["Dummy.mp3"]];

        var TestContextMenu = new SimpleContextMenu();
        // TestContextMenu.DisplayName = "SimpleContextMenu";

        // Ab hier Copy & Paste aus ServerManager

        var shellExtInitInterface = (IShellExtInit)TestContextMenu;
        var contextMenuInterface = (SharpShell.Interop.IContextMenu)TestContextMenu;

        try
        {
            //  Create the file paths.
            var filePaths = new StringCollection();
            filePaths.AddRange(items.Select(i => i.Path).ToArray());

            //  Create the data object from the file paths.
            var dataObject = new DataObject();
            dataObject.SetFileDropList(filePaths);

            //  Get the IUnknown COM interface address. Jesus .NET makes this easy.
            var dataObjectInterfacePointer = Marshal.GetIUnknownForObject(dataObject);

            //  Pass the data to the shell extension, attempt to initialise it.
            //  We must provide the data object as well as the parent folder PIDL.
            if (items.Any())
            {
                var folderPIDL = items.First().ParentItem.PIDL;
                shellExtInitInterface.Initialize(folderPIDL, dataObjectInterfacePointer, IntPtr.Zero); // Notwendig?
            }
        }
        catch (Exception)
        {
            //  Not supported for the file
            return;
        }

        //  Create a native menu.
        var menuHandle = CreatePopupMenu();

        //  Build the menu.
        contextMenuInterface.QueryContextMenu(menuHandle, 0, 0, 0x7FFF, 0);

        // Anzahl der Einträge
        int count = GetMenuItemCount(menuHandle);
        
        ReadMenu(menuHandle, "  ");


        // for (int i = 0; i < count; i++)
        // {
        //     var info = new MENUITEMINFO
        //     {
        //         cbSize = (uint)Marshal.SizeOf<MENUITEMINFO>(),
        //         fMask = 64 | 256 | 4, // MIIM_STRING u. MIIM_FTYPE u. MIIM_SUBMENU
        //         dwTypeData = new string('\0', 256),
        //         cch = 255
        //     };
        //     if (GetMenuItemInfo(menuHandle, (uint)i, true, ref info))
        //     {
        //         bool isSeparator = (info.fType & 0x00000800) != 0; // MFT_SEPARATOR
        //         if (isSeparator)
        //             Console.WriteLine($"[{i}] ---SEPARATOR---");
        //         else
        //             Console.WriteLine($"[{i}] {info.dwTypeData}");
        //         if (info.hSubMenu != IntPtr.Zero)
        //
        //
        //         
        //     }
        //     
        // }

        DestroyMenu(menuHandle);
    }

    static void ReadMenu(IntPtr hMenu, string indent = "")
    {
        int count = GetMenuItemCount(hMenu);
    
        for (int i = 0; i < count; i++)
        {
            var info = new MENUITEMINFO
            {
                cbSize = (uint)Marshal.SizeOf<MENUITEMINFO>(),
                fMask = 64 // MIIM_STRING
                        | 256 // MIIM_FTYPE
                        | 4, // MIIM_SUBMENU
                dwTypeData = new string('\0', 256),
                cch = 255
            };

            if (!GetMenuItemInfo(hMenu, (uint)i, true, ref info))
                continue;

            bool isSeparator = (info.fType & 0x00000800) != 0;

            if (isSeparator)
            {
                Console.WriteLine($"{indent}---");
            }
            else if (info.hSubMenu != IntPtr.Zero)
            {
                Console.WriteLine($"{indent}[+] {info.dwTypeData}");
                ReadMenu(info.hSubMenu, indent + "  "); // Rekursion
            }
            else
            {
                Console.WriteLine($"{indent}[ ] {info.dwTypeData}");
            }
        }
    }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct MENUITEMINFO
    {
        public uint cbSize;
        public uint fMask;
        public uint fType;
        public uint fState;
        public uint wID;
        public IntPtr hSubMenu;
        public IntPtr hbmpChecked;
        public IntPtr hbmpUnchecked;
        public IntPtr dwItemData;
        public string dwTypeData;
        public uint cch;
        public IntPtr hbmpItem;
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern bool GetMenuItemInfo(
        IntPtr hMenu,
        uint item,
        bool fByPosition,
        ref MENUITEMINFO lpmii);

    [DllImport("user32.dll")]
    static extern int GetMenuItemCount(IntPtr hMenu);

    [DllImport("user32.dll")]
    static extern bool DestroyMenu(IntPtr hMenu);


    /// <summary>
    /// Als using DisposableMap ...; verwenden!
    /// </summary>
    /// <param name="testFolder">Der Name des Unterordners in IntegrationTests, in dem die Dateien liegen, mit denen wir testen wollen</param>
    /// <returns>Eine Map, die die Namen der Dateien in testFolder auf die zugehörigen ShellItems mappt, mit welchen wir wiederum SimpleContextMenu querien können(???)</returns>
    /// <exception cref="Exception"></exception>
    private static DisposableMap ItemsInTestPath(string testFolder)
    {
        string testPath = Path.Combine(AppContext.BaseDirectory, "IntegrationTests", testFolder);


        // Nutze die Vorarbeit in ServerManager, indem du wie in diesem Desktop als Root setzt:
        ShellItem baseRoot = ShellItem.DesktopShellFolder;
        if (baseRoot == null)
        {
            throw new Exception("Failed to get the desktop folder.");
        }

        // Eigentliche Root die wir wollen, nämlich unser "Tests"-Ordner
        var testRoot = new ShellItem();
        var testPathPIDL = GetPIDLFromPath(testPath);

        testRoot.Initialise(testPathPIDL, baseRoot);
        if (testRoot == null)
        {
            throw new Exception("Failed to get the testRoot folder.");
        }


        // Die Map, die wir weiterreichen, die alle Namen der Ordner in IntegrationTests mit ShellItems verknüpft. Sollte als using-Block aufgerufen werden in der Methode, and die wir weiterreichen
        var testItemMap = new DisposableMap(baseRoot, testRoot);

        // Die Tests sind allesamt Dateien in testPath
        List<string> itemsInTestPath = Directory.EnumerateFileSystemEntries(testPath).ToList();

        foreach (string filename in itemsInTestPath)
        {
            string fullPath = Path.Combine(testPath, filename);
            IntPtr pathPIDL = GetPIDLFromPath(fullPath);
            var shellItem = new ShellItem();
            shellItem.Initialise(pathPIDL, testRoot);
            testItemMap.Add(Path.GetFileName(filename), shellItem);
        }

        return testItemMap;
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHParseDisplayName(
        string pszName,
        IntPtr pbc,
        out IntPtr ppidl,
        uint sfgaoIn,
        out uint psfgaoOut);

    public static IntPtr GetPIDLFromPath(string path)
    {
        int hr = SHParseDisplayName(path, IntPtr.Zero, out IntPtr pidl, 0, out _);

        if (hr != 0)
            Marshal.ThrowExceptionForHR(hr);

        return pidl;
    }


    public sealed class SafePidl : SafeHandle
    {
        public SafePidl(IntPtr pidl) : base(IntPtr.Zero, true)
        {
            SetHandle(pidl);
        }

        public override bool IsInvalid => handle == IntPtr.Zero;

        protected override bool ReleaseHandle()
        {
            Marshal.FreeCoTaskMem(handle);
            return true;
        }
    }

    /// <summary>
    /// A regular <see cref="Dictionary{string,ShellItem}"/>
    /// containing types that require explicit resource lifetime management.
    /// Dispose of this map before it goes out of scope,
    /// either by calling <see cref="Dispose"/> directly or by
    /// constraining its lifetime to a <c>using</c> block.
    /// </summary>
    private sealed class DisposableMap : Dictionary<string, ShellItem>, IDisposable
    {
        /// <summary>
        /// A collection of additional objects for the user to add everything which
        /// he wants to have the same lifetime as the DisposableMap instance.
        /// </summary>
        public List<IDisposable> disposables;

        /// <summary>
        /// A regular <see cref="Dictionary{string,ShellItem}"/>
        /// containing types that require explicit resource lifetime management.
        /// Dispose of this map before it goes out of scope,
        /// either by calling <see cref="Dispose"/> directly or by
        /// constraining its lifetime to a <c>using</c> block.
        /// </summary>
        /// <param name="disposables">Accepts a list of disposables, which will be disposed when the DisposableMap is being disposed.</param>
        public DisposableMap(params IDisposable[] disposables)
        {
            this.disposables = new List<IDisposable>(disposables);
            
        }

        public void Dispose()
        {
            this.Values.ToList().ForEach(item => item?.Dispose());
            foreach (IDisposable disposable in disposables)
            {
                disposable?.Dispose();
            }
        }
    }


    [DllImport("User32.dll")]
    internal static extern IntPtr CreatePopupMenu();
}