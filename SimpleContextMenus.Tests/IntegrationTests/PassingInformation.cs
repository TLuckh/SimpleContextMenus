namespace SimpleContextMenus.Tests;


/// <summary>
/// See <see cref="IntegrationTestBaseClass"/>
/// </summary>
public class PassingInformation: IntegrationTestBaseClass
{
    protected override string FromFolder => "PassedInformation";


    /// <summary>
    /// Tests that the path to the COM server, i.e. to the SimpleContMenu.dll that has been installed,
    /// is correctly passed as first parameter to a script executed by it.
    /// </summary>
    [Fact]
    public void TestComPath()
    {
        CopyFilesFromIntegrationTestsToTopLevelItems("WriteComServerLocation.py","WriteComServerLocationResult.txt");
        
        
        List<string> selectedTestItems = [];


        var contextMenuMock = new ContextMenuMock("", selectedTestItems);


        Node? exampleScriptNode = contextMenuMock.contextMenuTree.GetChildByName("WriteComServerLocation");

        Assert.True(exampleScriptNode != null, "Couldn't find the script to execute from within the context menu");

        exampleScriptNode.Invoke(contextMenuMock.contextMenuInterface);

        Thread.Sleep(1000);      // Give the script executed via context menu enough time to lock the files first 

        string content;
        var fs = new FileStream(
            Path.Combine(PathToTopLevelItems_Folder, "WriteComServerLocationResult.txt"),
            FileMode.Open,
            FileAccess.Read,
            FileShare.None);  // ← We lock the same file

        using (var file = new StreamReader(fs))
        {
            content = file.ReadToEnd();
        }

        Assert.Equal(Path.Combine(PathToTopLevelItems_Folder, "WriteComServerLocation.py"), content);
        
    }

    /// <summary>
    /// Tests that the path to current working directy (cwd) is correctly set as the folder in which the context menu has been spawned.
    ///
    /// Only works correctly if TestComPath() works correctly.
    /// </summary>
    [Fact]
    public void TestCwdPath()
    {
        CopyFilesFromIntegrationTestsToTopLevelItems("WriteCwd.py","WriteCwdResult.txt");
        
        List<string> selectedTestItems = [];

        var assemblySubFolder = "Extensions";
        var contextMenuMock = new ContextMenuMock(assemblySubFolder, selectedTestItems); // testFolderPath: anything works, but it shouldn't be "", since if instead of the cwd SimpleContextMenu passes its dll-location, we get a false positive
        
        Node? exampleScriptNode = contextMenuMock.contextMenuTree.GetChildByName("WriteCwd");
        
        Assert.True(exampleScriptNode != null, "Couldn't find the script to execute from within the context menu");
        
        exampleScriptNode.Invoke(contextMenuMock.contextMenuInterface);
        
        Thread.Sleep(1000);
        
        string content;
        var fs = new FileStream(
            Path.Combine(PathToTopLevelItems_Folder, "WriteCwdResult.txt"),
            FileMode.Open,
            FileAccess.Read,
            FileShare.None);  // ← We lock the same file

        using (var file = new StreamReader(fs))
        {
            content = file.ReadToEnd();
        }

        Assert.Equal(Path.Combine(PathToAssembly,assemblySubFolder), content);
        
    }

    /// <summary>
    /// Tests that the selected items, for which the context menu has been spawned, 
    /// are correctly passed as second, third, ... parameters to a script executed by it.
    ///
    /// Only works correctly if TestComPath() works correctly.
    /// </summary>
    [Fact]
    public void TestSelectedItems()
    {
        CopyFilesFromIntegrationTestsToTopLevelItems("WriteSelectedFiles.py","WriteSelectedFilesResult.txt");
        
        List<string> selectedTestItems = ["#Install.bat","Apex.dll","SimpleContextMenus.dll"];
        
        var contextMenuMock = new ContextMenuMock("", selectedTestItems);
        
        Node? exampleScriptNode = contextMenuMock.contextMenuTree.GetChildByName("WriteSelectedFiles");
        
        Assert.True(exampleScriptNode != null, "Couldn't find the script to execute from within the context menu");
        
        exampleScriptNode.Invoke(contextMenuMock.contextMenuInterface);
        
        Thread.Sleep(1000);
        
        string content;
        var fs = new FileStream(
            Path.Combine(PathToTopLevelItems_Folder, "WriteSelectedFilesResult.txt"),
            FileMode.Open,
            FileAccess.Read,
            FileShare.None);  // ← We lock the same file

        using (var file = new StreamReader(fs))
        {
            Assert.Equal(file.ReadLine().Trim() , Path.Combine(PathToAssembly,selectedTestItems[0]).Trim());
            Assert.Equal(file.ReadLine().Trim() , Path.Combine(PathToAssembly,selectedTestItems[1]).Trim());
            Assert.Equal(file.ReadLine().Trim() , Path.Combine(PathToAssembly,selectedTestItems[2]).Trim());
            
            
        }

    }

}