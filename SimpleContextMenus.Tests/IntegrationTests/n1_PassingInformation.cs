namespace SimpleContextMenus.Tests.IntegrationTests;


/// <summary>
/// See <see cref="IntegrationTestBaseClass"/>
/// </summary>
public class PassingInformation: IntegrationTestBaseClass
{
    protected override string FromFolder => "PassedInformation";


    /// <summary>
    /// Tests that the path to the file executed by SimpleContMenu.dll (via clicking a context menu entry),
    /// is correctly passed as the first argument to the file (CLI argument).
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

        Assert.Equal(Path.Combine(PathToAssemblyFolder,assemblySubFolder), content);
        
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
            Assert.Equal(file.ReadLine().Trim() , Path.Combine(PathToAssemblyFolder,selectedTestItems[0]).Trim());
            Assert.Equal(file.ReadLine().Trim() , Path.Combine(PathToAssemblyFolder,selectedTestItems[1]).Trim());
            Assert.Equal(file.ReadLine().Trim() , Path.Combine(PathToAssemblyFolder,selectedTestItems[2]).Trim());
            
            
        }

    }
    
    
    /// <summary>
    /// If nothing is selected, SimpleContextMenus.dll interprets it as "everything has been selected" for purposes of filtering the context menu entries.
    /// Yet it still means that we pass no selection to the script.
    ///
    /// This test tests the latter, i.e. we want to see, if we mark nothing as selected, then the script gets no selected items either. 
    ///
    /// Only works correctly if TestComPath() and TestSelectedItems() work correctly.
    /// </summary>
    [Fact]
    public void TestNoSelectionPassesNothing()
    {
        const string file1 = "WriteSelectedFiles.py";
        const string file2 = "WriteSelectedFilesResult.txt";
        CopyFilesFromIntegrationTestsToTopLevelItems(file1,file2);
        
        List<string> selectedTestItems = [];
        var testFolderPath = "TopLevelItems";

        var contextMenuMock = new ContextMenuMock(testFolderPath, selectedTestItems);
        
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

        string fileContents;
        using (var reader = new StreamReader(fs))
        {
            fileContents = reader.ReadToEnd();
        }
        Assert.Equal("", fileContents);
    }

}