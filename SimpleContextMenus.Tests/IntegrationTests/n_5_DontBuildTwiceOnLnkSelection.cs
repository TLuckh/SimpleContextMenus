/*
 * Unluckily, the ServerManager, and with that, the Mock I built, don't have the same behavior as explorer in this case.
 * Whereas explorer calls SimpleContextMenus twice, once to build  for the selection, and once for the resolved link, the ServerManager only calls it once:
 *
 *   /*
//          From SimpleContextMenu.cs::CanShowMenu:
//          In the windows explorer, if there's a selection, and the first item selected (i.e. whatever item the mouse hovers over when RMB'ing)
//          is a link-file (.lnk), the context menu is built twice:
//          Once for the selection, and once for the single resolved .lnk path.
//          
//          In this case it is expected that the second, and only the second building of the context menu is skipped.
//          * /
 *
 * Since I rely on the IntPtr to the context menu to sidestep this problem (second call gets ignored), which is set by explorer, this can't really be tested.
 */

// using IWshRuntimeLibrary;
//
// namespace SimpleContextMenus.Tests.IntegrationTests;
//
// /// <summary>
// /// Tests that menu entries that are links get resolved, and the executed script gets the working 
// /// </summary>
// public class DontBuildTwiceOnLnkSelection : IntegrationTestBaseClass
// {
//     protected override string FromFolder => "DontBuildTwiceOnLnkSelection";
//     private static string TestFolderPath { get; } = "LinkSource";
//
//     public DontBuildTwiceOnLnkSelection()
//     {
//         var fullTestFolderPath = Path.Combine(PathToAssemblyFolder, TestFolderPath);
//         if (Directory.Exists(fullTestFolderPath))
//             Directory.Delete(fullTestFolderPath, true);
//         Directory.CreateDirectory(fullTestFolderPath);
//     }
//
//
//     [Fact]
//     public void Test_lnk_LinkIsFollowed()
//     {
//         CopyFilesFromIntegrationTestsToAssemblyFolder(
//             newFolderToCopyInto: TestFolderPath,
//             "LinkTarget.txt");
//
//         CopyFilesFromIntegrationTestsToTopLevelItems("WriteSelectedFiles.py", "WriteSelectedFilesResult.txt");
//
//         var fullTestFolderPath = Path.Combine(PathToAssemblyFolder, TestFolderPath);
//
//
//         string linkFileName = "Link.lnk";
//         string linkPointsTo = Path.Combine(PathToAssemblyFolder, TestFolderPath, "LinkTarget.txt");
//         string linkIsFileAt = Path.Combine(fullTestFolderPath, linkFileName);
//         var shell = new WshShell();
//         var shortcut = (IWshShortcut)shell.CreateShortcut(linkIsFileAt); // linkPath endet auf .lnk
//         shortcut.TargetPath = linkPointsTo;
//         shortcut.Save();
//
//
//         List<string> selectedTestItems = [linkFileName];
//         var contextMenuMock = new ContextMenuMock(TestFolderPath, selectedTestItems);
//
//
//         List<Node> exampleScriptNode = contextMenuMock.contextMenuTree.GetChildrenByName("WriteSelectedFiles.py");
//
//
//         /*
//          From SimpleContextMenu.cs::CanShowMenu:
//          In the windows explorer, if there's a selection, and the first item selected (i.e. whatever item the mouse hovers over when RMB'ing)
//          is a link-file (.lnk), the context menu is built twice:
//          Once for the selection, and once for the single resolved .lnk path.
//          
//          In this case it is expected that the second, and only the second building of the context menu is skipped.
//          */
//         Assert.False(exampleScriptNode.Count == 0,
//             "Even if the target is a link, at least one of the two builds of the context menu has to go through.");
//         Assert.False(exampleScriptNode.Count > 1,
//             "Only one of the two builds should go trhough.");
//         
//         
//         
//         // Now, let's make sure the right of the two builds goes through:
//         exampleScriptNode[0].Invoke(contextMenuMock.contextMenuInterface);
//         Thread.Sleep(1000);      // Give the script executed via context menu enough time to lock the files first 
//
//
//         string content;
//         var fs = new FileStream(
//             Path.Combine(PathToTopLevelItems_Folder, "WriteSelectedFilesResult.txt"),
//             FileMode.Open,
//             FileAccess.Read,
//             FileShare.None);  // ← We lock the same file
//
//         using (var file = new StreamReader(fs))
//         {
//             content = file.ReadToEnd();
//         }
//
//         Assert.Equal(shortcut.TargetPath, content);
//
//     }
// }