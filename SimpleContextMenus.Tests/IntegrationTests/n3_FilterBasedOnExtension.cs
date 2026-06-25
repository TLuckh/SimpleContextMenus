namespace SimpleContextMenus.Tests.IntegrationTests;

/// <summary>
/// Tests whether NamingConvention.md's conventions work correctly.
/// </summary>
public class FilterBasedOnExtension : IntegrationTestBaseClass
{
    protected override string FromFolder => "FilterBasedOnExtension";
    private static string TestFolderPath { get; } = "SelectedItemsExample";

    public FilterBasedOnExtension()
    {
        var fullTestFolderPath = Path.Combine(PathToAssemblyFolder, TestFolderPath);
        if (Directory.Exists(fullTestFolderPath))
        {
            CheckFileLocks(fullTestFolderPath);


            Directory.Delete(fullTestFolderPath, true);
        }

        Directory.CreateDirectory(fullTestFolderPath);
        return;


        void CheckFileLocks(string s)
        {
            foreach (string? file in Directory.EnumerateFiles(s, "*", SearchOption.AllDirectories))
            {
                var x = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.None, 4096,
                    FileOptions.Asynchronous);
                x.Dispose();
            }
        }
    }

    [Fact]
    public void FilterOnExtensions_mp3FilteredOut()
    {
        List<string> selectedTestItems = ["Dummy.txt", "Dummy1.mp4"];
        string testFolderPath = TestFolderPath;
        CopyFilesFromIntegrationTestsToAssemblyFolder(
            newFolderToCopyInto: testFolderPath,
            selectedTestItems.ToArray());
        CopyFilesFromIntegrationTestsToTopLevelItems("Only_mp3.mp3.txt");

        var contextMenuMock = new ContextMenuMock(testFolderPath, selectedTestItems);

        Assert.Null(contextMenuMock.contextMenuTree.GetChildByName("Only_mp3"));
    }


    [Fact]
    public void FilterOnExtensions_mp3NotFilteredOut()
    {
        List<string> selectedTestItems = ["Dummy.mp3", "Dummy.txt", "Dummy1.mp4"];
        string testFolderPath = TestFolderPath;
        CopyFilesFromIntegrationTestsToAssemblyFolder(
            newFolderToCopyInto: testFolderPath,
            selectedTestItems.ToArray());
        CopyFilesFromIntegrationTestsToTopLevelItems("Only_mp3.mp3.txt");

        var contextMenuMock = new ContextMenuMock(testFolderPath, selectedTestItems);

        Assert.NotNull(contextMenuMock.contextMenuTree.GetChildByName("Only_mp3"));
    }

    [Fact]
    public void FilterOnExtensions_mp4()
    {
        string testFolderPath = TestFolderPath;
        List<string> selectedTestItems = ["Dummy.mp3", "Dummy.txt", "Dummy1.mp4", "Dummy2.mp4"];
        CopyFilesFromIntegrationTestsToAssemblyFolder(
            newFolderToCopyInto: testFolderPath,
            selectedTestItems.ToArray());
        CopyFilesFromIntegrationTestsToTopLevelItems("Only_mp4.mp4.txt");

        var contextMenuMock = new ContextMenuMock(testFolderPath, selectedTestItems);
        Assert.NotNull(contextMenuMock.contextMenuTree.GetChildByName("Only_mp4"));
    }

    [Fact]
    public void FilterOnMime_VIDEONotFilteredOut()
    {
        string testFolderPath = TestFolderPath;
        List<string> selectedTestItems = ["Dummy.mp3", "Dummy1.mp4", "Dummy2.mp4"];
        CopyFilesFromIntegrationTestsToAssemblyFolder(
            newFolderToCopyInto: testFolderPath,
            selectedTestItems.ToArray());
        CopyFilesFromIntegrationTestsToTopLevelItems("Only_VIDEO.VIDEO.txt");

        var contextMenuMock = new ContextMenuMock(testFolderPath, selectedTestItems);
        Assert.NotNull(contextMenuMock.contextMenuTree.GetChildByName("Only_VIDEO"));
    }

    [Fact]
    public void FilterOnMime_VIDEOFilteredOut()
    {
        string testFolderPath = TestFolderPath;
        List<string> selectedTestItems = ["Dummy.mp3"];
        CopyFilesFromIntegrationTestsToAssemblyFolder(
            newFolderToCopyInto: testFolderPath,
            selectedTestItems.ToArray());
        CopyFilesFromIntegrationTestsToTopLevelItems("Only_VIDEO.VIDEO.txt");

        var contextMenuMock = new ContextMenuMock(testFolderPath, selectedTestItems);
        Assert.Null(contextMenuMock.contextMenuTree.GetChildByName("Only_VIDEO"));
    }

    [Fact]
    public void FilterOnMimeFineType_VIDEONotFilteredOut()
    {
        string testFolderPath = TestFolderPath;
        List<string> selectedTestItems = ["Dummy.mp3", "Dummy1.mp4", "Dummy2.mp4"];
        CopyFilesFromIntegrationTestsToAssemblyFolder(
            newFolderToCopyInto: testFolderPath,
            selectedTestItems.ToArray());
        CopyFilesFromIntegrationTestsToTopLevelItems("Only_VIDEO.VIDEO..MP4.txt");

        var contextMenuMock = new ContextMenuMock(testFolderPath, selectedTestItems);
        Assert.NotNull(contextMenuMock.contextMenuTree.GetChildByName("Only_VIDEO"));
    }

    [Fact]
    public void FilterOnMimeFineType_VIDEOFilteredOut()
    {
        string testFolderPath = TestFolderPath;
        List<string> selectedTestItems = ["Dummy.mkv"];
        CopyFilesFromIntegrationTestsToAssemblyFolder(
            newFolderToCopyInto: testFolderPath,
            selectedTestItems.ToArray());
        CopyFilesFromIntegrationTestsToTopLevelItems("Only_VIDEO.VIDEO..MP4.txt");

        var contextMenuMock = new ContextMenuMock(testFolderPath, selectedTestItems);
        Assert.Null(contextMenuMock.contextMenuTree.GetChildByName("Only_VIDEO"));
    }


    [Fact]
    public void FilterOnMime_VIDEOFiltersSelectedFiles()
    {
        string testFolderPath = TestFolderPath;
        ;
        List<string> selectedTestItems = ["Dummy.mp3", "Dummy1.mp4", "Dummy2.mp4", "Dummy.mkv"];
        CopyFilesFromIntegrationTestsToAssemblyFolder(
            newFolderToCopyInto: testFolderPath,
            selectedTestItems.ToArray());
        CopyFilesFromIntegrationTestsToTopLevelItems("Only_VIDEO.VIDEO.py", "WriteSelectedFilesResult.txt");

        var contextMenuMock = new ContextMenuMock(testFolderPath, selectedTestItems);
        Node? contextMenuEntry = contextMenuMock.contextMenuTree.GetChildByName("Only_VIDEO");
        Assert.NotNull(contextMenuEntry);

        contextMenuEntry.Invoke(contextMenuMock.contextMenuInterface);
        Thread.Sleep(1000);

        string content;
        var fs = new FileStream(
            Path.Combine(PathToTopLevelItems_Folder, "WriteSelectedFilesResult.txt"),
            FileMode.Open,
            FileAccess.Read,
            FileShare.None); // ← We lock the same file

        List<string> selectedTestItemsFiltered = ["Dummy1.mp4", "Dummy2.mp4", "Dummy.mkv"];

        using (var file = new StreamReader(fs))
        {
            Assert.Equal(file.ReadLine().Trim(),
                Path.Combine(PathToAssemblyFolder, testFolderPath, selectedTestItemsFiltered[0]).Trim());
            Assert.Equal(file.ReadLine().Trim(),
                Path.Combine(PathToAssemblyFolder, testFolderPath, selectedTestItemsFiltered[1]).Trim());
            Assert.Equal(file.ReadLine().Trim(),
                Path.Combine(PathToAssemblyFolder, testFolderPath, selectedTestItemsFiltered[2]).Trim());
        }
    }
}