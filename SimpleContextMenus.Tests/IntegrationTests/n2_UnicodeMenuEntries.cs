using System.Text;
using Xunit.Abstractions;

namespace SimpleContextMenus.Tests.IntegrationTests;

public class UnicodeMenuEntries : IntegrationTestBaseClass
{
    private readonly ITestOutputHelper testOutputHelper;
    private string weirdUnicodeFilenameWithExtension;

    public UnicodeMenuEntries(ITestOutputHelper testOutputHelper)
    {
        this.testOutputHelper = testOutputHelper;
    }

    protected override string FromFolder => "ExoticUnicodeFileNames";

    public override void Dispose()
    {
        base.Dispose();
    }

    /// <summary>
    /// Tests that unicode menu entries are displayed correctly, and that passing to the script doesn't mangle the Unicode.
    /// </summary>
    [Fact]
    public void TestUnicodeMenuEntries()
    {
        const string weirdUnicodeFilename = "Héllo_wörld_jap_日本語";
        weirdUnicodeFilenameWithExtension = weirdUnicodeFilename + ".py";
        const string resultUnicodeFilename = weirdUnicodeFilename + "Result" + ".txt";

        CopyFilesFromIntegrationTestsToTopLevelItems(weirdUnicodeFilenameWithExtension, resultUnicodeFilename);


        var testFolderPath = "TopLevelItems";
        List<string> selectedTestItems = [weirdUnicodeFilenameWithExtension, resultUnicodeFilename];

        var contextMenuMock = new ContextMenuMock(testFolderPath, selectedTestItems);


        Node? contextMenuItem = contextMenuMock.GetContextMenuEntry(weirdUnicodeFilename);

        Assert.NotNull(contextMenuItem);
        Assert.Equal(weirdUnicodeFilename, contextMenuItem.ContextMenuEntryName);

        // Make sure that the displayed name of the context menu entry is the same as what we passed:
        contextMenuItem.Invoke(contextMenuMock.contextMenuInterface);
        Thread.Sleep(1000); // Give the script executed via context menu enough time to lock the files first 


        // Make sure that the script executed when clicking the context menu entry also sees the passed Selected Items correctly:
        string content;
        var fs = new FileStream(
            Path.Combine(PathToTopLevelItems_Folder, resultUnicodeFilename),
            FileMode.Open,
            FileAccess.Read,
            FileShare.None); // ← We lock the same file

        using (var file = new StreamReader(fs, Encoding.UTF8))
        {
            Assert.Equal(
                Path.Combine(PathToTopLevelItems_Folder, weirdUnicodeFilenameWithExtension),
                file.ReadLine().Trim());
            
            Assert.Equal(
                Path.Combine(PathToTopLevelItems_Folder, resultUnicodeFilename), 
                file.ReadLine().Trim());
        }
    }


}