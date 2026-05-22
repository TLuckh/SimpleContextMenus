using System.Collections.Specialized;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using ServerManager.ShellDebugger;
using SharpShell.Interop;
using SharpShell.SharpContextMenu;
using static SimpleContextMenus.Tests.MarshallingStructures;

namespace SimpleContextMenus.Tests;

public class ExampleTests
{
    public static void ExampleTest1()
    {

        var testFolderPath = Path.Combine("CopyToOutputTests", "IntegrationTests", "1FilterBasedOnExtension");
        List<string> selectedTestItems = ["Dummy.mp3"];

        var contextMenuMock = new ContextMenuMock(testFolderPath, selectedTestItems);

        Console.WriteLine(contextMenuMock.contextMenuTree);
    }

    public static void ExampleTest2()
    {
        var testFolderPath = Path.Combine("CopyToOutputTests", "IntegrationTests", "1FilterBasedOnExtension");
        List<string> selectedTestItems = ["Dummy.mp3"];

        var contextMenuMock = new ContextMenuMock(testFolderPath, selectedTestItems);

        var exampleScriptNode = contextMenuMock.contextMenuTree.Children[1].Children[0];

        exampleScriptNode.Invoke(contextMenuMock.contextMenuInterface);


        // Die Klasse hat hier eigentlich, was wir wollen:
        // ((SharpContextMenu) contextMenuMock.contextMenuInterface).NativeContextMenuWrapper
        // Ist aber private...

        Console.WriteLine(contextMenuMock.contextMenuTree);
    }
    
}


/// <summary>
/// See IntegrationTestBaseClass.cs's documentation for its semantic, and the test cases for its uasge.
/// </summary>
public class ContextMenuMock
{
    /// <summary>
    /// Root node of a tree datastructure representing the context menu. Can be used to view or execute the items of the context menu.
    /// </summary>
    public Node contextMenuTree;

    /// <summary>
    /// A wrapper over Map with convenience methods to group the lifetimes of IDisposable-items. See type definition for more info.
    /// </summary>
    private DisposableMap disposableMap;

    /// <summary>
    /// The SimpleContextMenu instance, cast to IContextMenu, which is needed in some Windows-COM-APIs (to reference an item of a context menu, you need its ID in the context menu and a reference to the context menu itself).    /// </summary>
    public IContextMenu contextMenuInterface;

    /// <summary>
    /// The files & folders in testFolderPath, which shall be viewed as marked for the building of the context menu.
    /// </summary>
    public List<string> SelectedTestItems { get; }

    /// <summary>
    /// The subpath to the folder in which we want to simulate the building of the context menu, starting from the output directory, i.e., where the SimpleContextMenus.Tests.dll is being built.
    /// </summary>
    public string TestFolderPath { get; }

    /// <summary>
    /// See <see cref="TestFolderPath"/>.
    /// </summary>
    public string TestFolderPathAbsolute => Path.Combine(AppContext.BaseDirectory, TestFolderPath);


    /// <summary>Returns the path to the folder (that is, what it was during compilation) in which the .cs file containing this method is located</summary>
    private static string getProjectDir([CallerFilePath] string path = "") =>
        Path.GetDirectoryName(path)!;

    /// <summary>Returns the path to the folder (that is, what it was during compilation) in which the .cs file containing this method is located</summary>
    public static string GetProjectDir([CallerFilePath] string path = "") => getProjectDir();


    /// <summary>
    /// Returns the first Node found representing the context menu entry with the given name. Returns null if no such entry exists.
    /// Searches all sebmenus in a DFS fashion, i.e. it returns the first matching node in the PreOrder representation of the context menu
    /// </summary>
    /// <param name="name"></param>
    /// <param name="ignoreCase"></param>
    /// <returns></returns>
    public Node? GetContextMenuEntry(string name, bool ignoreCase = false)
    {
        var comparisonOptions = ignoreCase
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;

        // Manual DFS
        Stack<Node> stack = [];
        stack.Push(contextMenuTree);
        while (stack.Count > 0)
        {
            Node currentNode = stack.Pop();
            // Test if name of currentNode has been 
            if (String.Equals(currentNode.ContextMenuEntryName.Trim(), name.Trim(), comparisonOptions))
            {
                return currentNode;
            }

            currentNode.Children.ForEach(stack.Push);
        }

        return null;
    }

    public void InvokeContextMenuEntry(Node contextMenuEntry)
    {
        contextMenuEntry.Invoke(contextMenuInterface);
    }


    /// <summary>
    /// 
    /// </summary>
    /// <param name="hMenu">The <see cref="IntPtr"/> pointing towards the handle for the windows context menu, most likely from <see cref="CreatePopupMenu"/></param>
    /// <param name="rootNodeName">Freely choosable name for the root node representing the context menu itself </param>
    static Node ReadMenu(IntPtr hMenu, string rootNodeName = "Context Menu")
    {
        Node contextMenuTreeView = new Node(rootNodeName);

        for (int i = 0; i < GetMenuItemCount(hMenu); i++)
        {
            var info = new MENUITEMINFO
            {
                cbSize = (uint)Marshal.SizeOf<MENUITEMINFO>(),
                fMask = 64 // MIIM_STRING
                        | 256 // MIIM_FTYPE
                        | 4 // MIIM_SUBMENU
                        | 2, // MIIM_ID  ; Sagt GetMenuItemInfo, dass es die ID des Menuitems zurückliefern soll. Später für InvokeCommand notwendig, zum Ausführen des Kontextmenüeintrags
                dwTypeData = new string('\0', 256),
                cch = 255
            };

            if (!GetMenuItemInfo(hMenu, (uint)i, true, ref info))
                continue;

            uint menuItemId = info.wID;

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

            contextMenuTreeView.Children.Last().ContextMenuEntryHandlerCommandId =
                menuItemId; // Setze die CommandId für alle Items, auch Submenus, damit wir sie später mit InvokeCommand ausführen können
        }

        return contextMenuTreeView;
    }


    /// <summary>
    /// Simulates the act of:
    /// 1. installing the SimpleContextMenus.dll,
    /// 2. then navigating into testFolderPath,
    /// 3. then marking the items in testFolderPath which are listed in selectedItems, and
    /// 4. then right-clicking on one of the marked files, so that we get the context.
    /// If no file has been marked, the click is viewed as a right-click on the empty space within the explorer.
    ///
    /// Should be disposed of after use, either by calling Dispose() directly or by constraining its lifetime to a using-block, to free up the resources used for the ShellItems within the context menu tree!
    /// </summary>
    /// <param name="testFolderPath">A subpath starting from the output directory, i.e. where the SimpleContextMenus.Tests.dll is being built</param>
    /// <param name="selectedTestItems">Files &amp; Folders in 'testFolderPath', which shall be viewed as marked for the building of the context menu.</param>
    /// <returns>A tree of Nodes representing a view of the context menu. Can be used to view or execute the items of the context menu.</returns>
    public ContextMenuMock(string testFolderPath, List<string> selectedTestItems)
    {
        (this.contextMenuTree, disposableMap, contextMenuInterface) =
            GetContextMenuView(testFolderPath, selectedTestItems);
        TestFolderPath = testFolderPath;
        SelectedTestItems = selectedTestItems;
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
    /// <param name="testFolderPath">A subpath starting from the output directory, i.e., where the SimpleContextMenus.Tests.dll is being built</param>
    /// <param name="selectedTestItems">Files & Folders in 'testFolderPath', which shall be viewed as marked for the building of the context menu.</param>
    /// <returns>Returns: 1. A tree of Nodes representing a view of the context menu. Can be used to view or execute the items of the context menu.
    /// 2. A disposable map which maps paths to their ShellItem-definitions, and is to be used to group their lifetimes (i.e. dispopse of the map at the end of its lifecycle to free all handles from ShellItems).
    /// 3. The SimpleContextMenu instance, cast to IContextMenu, which is needed in some Windows-COM-APIs (to reference an item of a context menu, you need its ID in the context menu and a reference to the context menu itself).
    /// </returns>
    private static (Node, DisposableMap, IContextMenu) GetContextMenuView(string testFolderPath,
        List<string> selectedTestItems)
    {

        // We let SharpShells ShellItem-class handle most of the marashalling necessary to build the context menu; 
        // All ItemsInTestPath does is to get PIDLs of the files in the testFolderPath, which we'll use indirectly via SharpShells ShellItem-class.
        //
        DisposableMap testItemMap = ItemsInTestPath(testFolderPath);
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
            var comDataObject = (System.Runtime.InteropServices.ComTypes.IDataObject)dataObject;
            var dataObjectInterfacePointer = Marshal.GetComInterfaceForObject(
                dataObject, 
                typeof(System.Runtime.InteropServices.ComTypes.IDataObject));

            //  Pass the data to the shell extension, attempt to initialise it.
            //  We must provide the data object as well as the parent folder PIDL.

            var folderPIDL = testItemMap.TestRoot.PIDL;
            shellExtInitInterface.Initialize(folderPIDL, dataObjectInterfacePointer, IntPtr.Zero); // Initializes (via a base class), FolderPath and selectedItems in SimpleContextMenu
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

        Node returnItem = ReadMenu(menuHandle);

        // The handle for the main context menu has to be disposed; The submenus should be collected automatically then
        DestroyMenu(menuHandle);
        return (returnItem, testItemMap, contextMenuInterface);
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
        var testItemMap = new DisposableMap(testRoot, baseRoot);

        // Die Tests sind allesamt Dateien in testPath
        List<string> itemsInTestPath = Directory.EnumerateFileSystemEntries(testPath).ToList();

        foreach (string filename in itemsInTestPath)
        {
            string fullPath = Path.Combine(testPath, filename);
            IntPtr pathPIDL = GetPIDLFromPath(fullPath);
            IntPtr relativePIDL = ILFindLastID(pathPIDL);

            var shellItem = new ShellItem();
            shellItem.Initialise(relativePIDL, testRoot);

            testItemMap.Add(Path.GetFileName(filename), shellItem);
            testItemMap.disposables.Add(new SafePidl(pathPIDL));
        }

        return testItemMap;
    }
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
sealed class DisposableMap : Dictionary<string, ShellItem>, IDisposable
{
    /// <summary>
    /// A collection of additional objects for the user to add everything which
    /// he wants to have the same lifetime as the DisposableMap instance.
    /// </summary>
    public List<IDisposable> disposables;

    /// <summary>
    /// The root folder of the items in the context menu (not the baseRoot, which isn't shown there).
    /// </summary>
    public ShellItem TestRoot { get; set; }

    /// <summary>
    /// A regular <see cref="Dictionary{string,ShellItem}"/>
    /// containing types that require explicit resource lifetime management.
    /// Dispose of this map before it goes out of scope,
    /// either by calling <see cref="Dispose"/> directly or by
    /// constraining its lifetime to a <c>using</c> block.
    /// </summary>
    /// <param name="testRoot">The root folder of the items in the context menu (not the baseRoot, which isn't shown there).</param>
    /// <param name="disposables">Accepts a list of IDisposable, each of which will be disposed when the DisposableMap is being disposed.</param>
    public DisposableMap(ShellItem testRoot, params IDisposable[] disposables)
    {
        this.disposables = new List<IDisposable>(disposables);
        this.disposables.Add(testRoot);
        this.TestRoot = testRoot;
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