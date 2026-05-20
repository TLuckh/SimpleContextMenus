using System.Reflection;
using PathLib;

namespace SimpleContextMenus.Tests;

/*
 * # Regression Tests
Add files to the Debug configuration that make the following tests quick to do:
- Test that the filters for file extensions/MIME types work correctly, so only if an applicable item is in the folder/selected items,
the folders with the file extensions/MIME types are shown
- Test that folders without files, and folders with no applicable items aren't shown
- Test that if a folder/file has a file extensions/MIME type it filters for, that only matching selected items are passed to it
- Test that if you execute Example Script.py after it has moved by putting a shortcut to it into the TopLevelFolder/Extensions folder, it correctly shows the exectued file location.
- Test that if a shortcut has been orphaned (file deleted, folder deleted, drive unplugged), the menu item will be replaced by an error message, without any other menu item being missing





 */

public abstract class IntegrationTestBaseClass : IDisposable
{
    
    internal static string PathToAssembly { get; } = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!;

    protected static string PathTo_CopyToOutputTests { get; } =
        Path.Combine(ContextMenuMock.GetProjectDir(), "CopyToOutputTests");

    protected static string PathToTopLevelItems_Folder { get; } = Path.Combine(PathToAssembly, "TopLevelItems");
    protected static string PathToExtensions_Folder { get; } = Path.Combine(PathToAssembly, "Extensions");

    /// <summary>
    /// The folder in CopyToOutput for which we want to take files from
    /// </summary>
    protected abstract string FromFolder { get; }
    
    
    /// <summary>
    /// Copies the files from  CopyToOutputTests/<see cref="FromFolder"/>, which are named in filesToCopy,
    /// into the folder TopLevelItems
    /// so they can appear as menu entries in the context menu.
    /// </summary>
    /// <param name="filesToCopy"></param>
    protected void CopyFilesFromIntegrationTestsToTopLevelItems(params string[] filesToCopy)
    {
        var rootPath = Path.Combine(ContextMenuMock.GetProjectDir(), "CopyToOutputTests", "IntegrationTests");
        foreach (string file in filesToCopy)
        {
            File.Copy(
                Path.Combine(rootPath,FromFolder, file),
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
        var rootPath = Path.Combine(ContextMenuMock.GetProjectDir(), "CopyToOutputTests", "IntegrationTests");
        foreach (string file in filesToCopy)
        {
            File.Copy(
                Path.Combine(rootPath, FromFolder, file),
                Path.Combine(PathToExtensions_Folder, new FileInfo(file).Name),
                true);
        }
    }


    public void Dispose()
    {
        var topLevelItemsFolder = new DirectoryInfo(PathToTopLevelItems_Folder);
        var extensionsFolder = new DirectoryInfo(PathToExtensions_Folder);
        
        topLevelItemsFolder.GetFiles().ToList().ForEach(file => file.Delete());
        extensionsFolder.GetFiles().ToList().ForEach(file => file.Delete());
    }
}