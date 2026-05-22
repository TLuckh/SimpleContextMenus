using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using IWshRuntimeLibrary;

namespace SimpleContextMenus.Tests.IntegrationTests;

/// <summary>
/// Tests that menu entries that are links get resolved, and the executed script gets the working 
/// </summary>
public class FollowLinks : IntegrationTestBaseClass
{
    protected override string FromFolder => "FollowLinks";
    private static string TestFolderPath { get; } = "LinkSource";

    public FollowLinks()
    {
        var fullTestFolderPath = Path.Combine(PathToAssemblyFolder, TestFolderPath);
        if (Directory.Exists(fullTestFolderPath))
            Directory.Delete(fullTestFolderPath, true);
        Directory.CreateDirectory(fullTestFolderPath);
    }

    
    [Fact]
    public void Test_lnk_LinkIsFollowed()
    {
        CopyFilesFromIntegrationTestsToAssemblyFolder(
            newFolderToCopyInto: TestFolderPath,
            "WriteComServerLocation.py","WriteComServerLocationResult.txt");

        string linkeFileName = "Link";
        string linkPointsTo = Path.Combine(PathToAssemblyFolder, TestFolderPath, "WriteComServerLocation.py");
        string linkIsFileAt = Path.Combine(PathToTopLevelItems_Folder, linkeFileName + ".lnk");
        var shell = new WshShell();
        var shortcut = (IWshShortcut)shell.CreateShortcut(linkIsFileAt); // linkPath endet auf .lnk
        shortcut.TargetPath = linkPointsTo;
        shortcut.Save();

        
        List<string> selectedTestItems = [];
        var contextMenuMock = new ContextMenuMock("", selectedTestItems);


        Node? exampleScriptNode = contextMenuMock.contextMenuTree.GetChildByName(linkeFileName);

        Assert.True(exampleScriptNode != null, "Couldn't find the script to execute from within the context menu");

        exampleScriptNode.Invoke(contextMenuMock.contextMenuInterface);
        Thread.Sleep(1000);      // Give the script executed via context menu enough time to lock the files first 


        string content;
        var fs = new FileStream(
            Path.Combine(PathToAssemblyFolder,TestFolderPath, "WriteComServerLocationResult.txt"),
            FileMode.Open,
            FileAccess.Read,
            FileShare.None);  // ← We lock the same file

        using (var file = new StreamReader(fs))
        {
            content = file.ReadToEnd();
        }

        Assert.Equal(linkPointsTo, content);
        
    }


    [Fact]
    public void TestSoftLinkIsFollowed()
    {
        Assert.Fail("TODO");

    }
    
    [Fact]
    public void TestHardLinkIsFollowed()
    {
        Assert.Fail("TODO");

    }

}