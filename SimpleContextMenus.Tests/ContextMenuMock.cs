using System.Collections.Specialized;
using System.Runtime.InteropServices;
using ServerManager.ShellDebugger;
using SharpShell.Interop;
using static SimpleContextMenus.Tests.MarshallingStructures;

namespace SimpleContextMenus.Tests;

public class ContextMenuMock
{
    public Node<string> contextMenuTree;
    public ContextMenuMock(string testFolderPath, List<string> selectedTestItems)
    {
        this.contextMenuTree = GetContextMenuView(testFolderPath, selectedTestItems);
        
        
    }
    /// <summary>
    /// Simulates the act of:
    /// 1. installing the SimpleContextMenus.dll,
    /// 2. then navigating into testFolderPath,
    /// 3. then marking the items in testFolderPath which are listed in selectedItems, and
    /// 4. then right-clicking on one of the marked files, so that we get the context.
    /// If no file has been marked, the click is viewed as a right-click on the empty space within the explorer.
    /// 
    /// </summary>
    /// <param name="testFolderPath">A subpath starting from the output directory, i.e. where the SimpleContextMenus.Tests.dll is being built</param>
    /// <param name="selectedTestItems">Files & Folders in 'testFolderPath', which shall be viewed as marked for the building of the context menu.</param>
    /// <returns>A tree of [ToDo] which can be used to view or execute the items of the context menu.</returns>
    public static Node<string> GetContextMenuView(string testFolderPath, List<string> selectedTestItems)
    {
        // Der Pfad der zu testenden Dateien und Ordner. 
        // TODo: Build von SimpleContextMenus.Tests sollte SimpleContextMenus bauen, und anschließend alle davon erzeugten Dateien in einen neuen Ordner kopieren, wo wir auch die Tests reinhauen (die noch zu schreiben sind)
        // Die Tests selbst sollten eine relativ flache Dateistruktur darstellen, so dass wir diese für jeden Test manuell bauen können: 
        // Es handelt sich einfach um eine Menge von Ordnern mit Dateien drin. Wir simulieren Klicks auf Teilmengen (inkl. leere Teilmenge) dieser Dateien im Ordner, und unser Ziel ist es, jeweils die richtigen Kontextmenüs zu kriegen.
        // Der Test der Kontextmenüs wiederum kommt stattdessen in einen Unit Test (das hier sind Integration Tests i think?)

        // We let SharpShells ShellItem-class handle most of the marashalling necessary to build the context menu; 
        // All ItemsInTestPath does is to get PIDLs of the files in the testFolderPath, which we'll use indirectly via SharpShells ShellItem-class.
        //
        using DisposableMap testItemMap = ItemsInTestPath(testFolderPath);
        // Definition der markierten Items
        ShellItem[] selectedItems = selectedTestItems.Select(name => testItemMap[name]).ToArray();

        var testContextMenu = new SimpleContextMenu();
        // TestContextMenu.DisplayName = "SimpleContextMenu";

        #region Copy & Paste aus ServerManager

        var shellExtInitInterface = (IShellExtInit)testContextMenu;
        var contextMenuInterface = (SharpShell.Interop.IContextMenu)testContextMenu;

        try
        {
            //  Create the file paths.
            var filePaths = new StringCollection();
            filePaths.AddRange(selectedItems.Select(i => i.Path).ToArray());

            //  Create the data object from the file paths.
            var dataObject = new DataObject();
            dataObject.SetFileDropList(filePaths);

            //  Get the IUnknown COM interface address. Jesus .NET makes this easy.
            var dataObjectInterfacePointer = Marshal.GetIUnknownForObject(dataObject);

            //  Pass the data to the shell extension, attempt to initialise it.
            //  We must provide the data object as well as the parent folder PIDL.
            if (selectedItems.Any())
            {
                var folderPIDL = selectedItems.First().ParentItem.PIDL;
                shellExtInitInterface.Initialize(folderPIDL, dataObjectInterfacePointer, IntPtr.Zero); // Notwendig?
            }
        }
        catch (Exception)
        {
            throw new Exception("Not supported for the file");
        }

        //  Create a native menu.
        var menuHandle = CreatePopupMenu();

        #endregion Copy & Paste aus ServerManager

        //  Build the menu
        contextMenuInterface.QueryContextMenu(menuHandle, 0, 0, 0x7FFF, 0);

        Node<string> returnItem = ReadMenu(menuHandle);

        // The handle for the main context menu has to be disposed; The submenus should be collected automatically then
        DestroyMenu(menuHandle);
        return returnItem;
    }

    public static void Main()
    {
        // Der Pfad der zu testenden Dateien und Ordner. 
        // TODo: Build von SimpleContextMenus.Tests sollte SimpleContextMenus bauen, und anschließend alle davon erzeugten Dateien in einen neuen Ordner kopieren, wo wir auch die Tests reinhauen (die noch zu schreiben sind)
        // Die Tests selbst sollten eine relativ flache Dateistruktur darstellen, so dass wir diese für jeden Test manuell bauen können: 
        // Es handelt sich einfach um eine Menge von Ordnern mit Dateien drin. Wir simulieren Klicks auf Teilmengen (inkl. leere Teilmenge) dieser Dateien im Ordner, und unser Ziel ist es, jeweils die richtigen Kontextmenüs zu kriegen.
        // Der Test der Kontextmenüs wiederum kommt stattdessen in einen Unit Test (das hier sind Integration Tests i think?)

        var testFolderPath = Path.Combine("IntegrationTests", "1FilterBasedOnExtension");
        List<string> selectedTestItems = ["Dummy.mp3"];

        Node<string> contextMenuTree = GetContextMenuView(testFolderPath, selectedTestItems);

        Console.WriteLine(contextMenuTree);
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="hMenu">The <see cref="IntPtr"/> pointing towards the handle for the windows context menu, most likely from <see cref="CreatePopupMenu"/></param>
    /// <param name="rootNodeName">Freely choosable name for the root node representing the context menu itself </param>
    static Node<string> ReadMenu(IntPtr hMenu, string rootNodeName = "Context Menu")
    {
        Node<string> contextMenuTreeView = new Node<string>(rootNodeName);

        for (int i = 0; i < GetMenuItemCount(hMenu); i++)
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
            if (isSeparator) // case separator
            {
                contextMenuTreeView.AddChild("---");
            }
            else if (info.hSubMenu != IntPtr.Zero) // case submenu
            {
                contextMenuTreeView.AddChild(ReadMenu(info.hSubMenu, info.dwTypeData)); // Recursion-step
            }
            else // case end-item
            {
                contextMenuTreeView.AddChild(info.dwTypeData);
            }
        }

        return contextMenuTreeView;
    }


    /// <summary>
    /// Returns the number of items in the passed windows explorer context menu
    /// </summary>
    /// <param name="hMenu"></param>
    /// <returns></returns>



    /// <summary>
    /// Als using DisposableMap ...; verwenden!
    /// </summary>
    /// <param name="testFolder">Der Name des Unterordners in IntegrationTests, in dem die Dateien liegen, mit denen wir testen wollen</param>
    /// <returns>Eine Map, die die Namen der Dateien in testFolder auf die zugehörigen ShellItems mappt, mit welchen wir wiederum SimpleContextMenu querien können(???)</returns>
    /// <exception cref="Exception"></exception>
    private static DisposableMap ItemsInTestPath(string testFolder)
    {
        string testPath = Path.Combine(AppContext.BaseDirectory, testFolder);


        // Nutze die Vorarbeit in ServerManager, indem du wie in diesem Desktop als Root setzt: [vermutlich nicht notwendig]
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





    /// <summary>
    /// Class to wrap the IntPtr handles in to make it easy to Dispose of them later of, e.g. with the using keyword,
    /// or by adding the to a DisposableMap (of which one then still has to  dispose of  manually later-on).
    /// </summary>
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
        /// <param name="disposables">Accepts a list of IDisposable, each of which will be disposed when the DisposableMap is being disposed.</param>
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



}