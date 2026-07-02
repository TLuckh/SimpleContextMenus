using System.Reflection;

namespace SimpleContextMenus.Tests.IntegrationTests;

/*
 TODO:
 * # Regression Tests
Add files to the Debug configuration that make the following tests quick to do:
- Test that the filters for file extensions/MIME types work correctly, so only if an applicable item is in the folder/selected items,
the folders with the file extensions/MIME types are shown
- Test that folders without files, and folders with no applicable items aren't shown
- Test that if a folder/file has a file extensions/MIME type it filters for, that only matching selected items are passed to it
- Test that if you execute Example Script.py after it has moved by putting a shortcut to it into the TopLevelFolder/Extensions folder, it correctly shows the exectued file location.
- Test that if a shortcut has been orphaned (file deleted, folder deleted, drive unplugged), the menu item will be replaced by an error message, without any other menu item being missing

 */

/// <summary>
///  
/// The tests construct, using a ContextMenu instance at its core, each a windows context menu using the official API for it.
/// However, the only COM Server we make visible to it is SimpleContextMenus.dll (and context menu entries which should originate from registry aren't passed at all). 
///  
/// When using the context menu API, we pass both a folder path and a list of selected items.
/// The windows context menu then passes the folder path & selected items to the COM Servers, i.e. in our case, to SimpleContextMenus.dll.
/// SimpleContextMenus.dll returns, based on the passed information, a number of context menus submenus & context menu entries, which the windows context menu shall show.
/// Finally, a click on a context menu entry prompts the context menu API to execute whatever the SimpleContextMenus.dll setup for the context menu entry to launch.
///  
///     Aside from potential bugs introduced during the calling of the API, this should be as close as possible to constructing
/// a real context menu in windows explorer when SimpleContextMenus.dll is installed in the computer.  
/// </summary>
public abstract class IntegrationTestBaseClass : IDisposable
{
    private static readonly string TestsFolderBasePath = Path.Combine("CopyToOutputTests", "IntegrationTests");
    internal static string PathToAssemblyFolder { get; } = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!;

    protected static string PathTo_CopyToOutputTests { get; } =
        Path.Combine(ContextMenuMock.GetProjectDir(), "CopyToOutputTests");

    protected static string PathToTopLevelItems_Folder { get; } = Path.Combine(PathToAssemblyFolder, "TopLevelItems");
    protected static string PathToExtensions_Folder { get; } = Path.Combine(PathToAssemblyFolder, "Extensions");

    /// <summary>
    /// The folder in CopyToOutput for which we want to take files from
    /// </summary>
    protected abstract string FromFolder { get; }

    /// <summary>
    /// Copies the files from  CopyToOutputTests/<see cref="FromFolder"/>, which are named in <paramref name="filesToCopy"/>,
    /// into a new subfolder in the folder in which the assembly SimpleContextMenus.Tests.dll is located.
    /// </summary>
    /// <param name="newFolderToCopyInto">The name of a new folder within the folder in which the .dll is located.</param>
    /// <param name="filesToCopy">The files to copy into the new folder.</param>
    protected void CopyFilesFromIntegrationTestsToAssemblyFolder(string newFolderToCopyInto = "", params string[] filesToCopy)
    {
        var rootPath = Path.Combine(ContextMenuMock.GetProjectDir(), TestsFolderBasePath);
        Directory.CreateDirectory(Path.Combine(PathToAssemblyFolder, newFolderToCopyInto));
        foreach (string file in filesToCopy)
        {
            File.Copy(
                Path.Combine(rootPath, FromFolder, file),
                Path.Combine(PathToAssemblyFolder, newFolderToCopyInto ,new FileInfo(file).Name),
                true);
        }
    }
    
    /// <summary>
    /// Copies the files from  CopyToOutputTests/<see cref="FromFolder"/>, which are named in filesToCopy,
    /// into the folder TopLevelItems
    /// so they can appear as menu entries in the context menu.
    /// </summary>
    /// <param name="filesToCopy"></param>
    protected void CopyFilesFromIntegrationTestsToTopLevelItems(params string[] filesToCopy)
    {
        var rootPath = Path.Combine(ContextMenuMock.GetProjectDir(), TestsFolderBasePath);
        foreach (string file in filesToCopy)
        {
            File.Copy(
                Path.Combine(rootPath, FromFolder, file),
                Path.Combine(PathToTopLevelItems_Folder, new FileInfo(file).Name),
                true);
        }
    }

    /// <summary>
    /// Copies the files from  CopyToOutputTests/<see cref="FromFolder"/>, which are named in filesToCopy,
    /// into the folder Extensions
    /// so they can appear as menu entries in the context menu.
    /// </summary>
    /// <param name="filesToCopy"></param>
    protected void CopyFilesFromIntegrationTestsToExtensions(params string[] filesToCopy)
    {
        var rootPath = Path.Combine(ContextMenuMock.GetProjectDir(), TestsFolderBasePath);
        foreach (string file in filesToCopy)
        {
            File.Copy(
                Path.Combine(rootPath, FromFolder, file),
                Path.Combine(PathToExtensions_Folder, new FileInfo(file).Name),
                true);
        }
    }

    
    public virtual void Dispose()
    {
        var topLevelItemsFolder = new DirectoryInfo(PathToTopLevelItems_Folder);
        var extensionsFolder = new DirectoryInfo(PathToExtensions_Folder);

        topLevelItemsFolder.GetFiles().ToList().ForEach(file => file.Delete());
        extensionsFolder.GetFiles().ToList().ForEach(file => file.Delete());
    }
}