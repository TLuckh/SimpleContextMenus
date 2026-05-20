namespace SimpleContextMenus.Tests;

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
        
        
        List<string> selectedTestItems = ["Music.AUDIO.txt"];


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
    /// </summary>
    [Fact]
    public void TestCwdPath()
    {
    }

    /// <summary>
    /// Tests that the selected items, for which the context menu has been spawned, 
    /// are correctly passed as second, third, ... parameters to a script executed by it.
    /// </summary>
    [Fact]
    public void TestSelectedItems()
    {
    }

}