namespace SimpleContextMenus.Tests;


public static class Extensions
{
    public static void Copy(this DirectoryInfo self, DirectoryInfo destination, bool recursively)
    {
        foreach (var file in self.GetFiles())
        {
            file.CopyTo(Path.Combine(destination.FullName, file.Name));
        }

        if (recursively)
        {
            foreach (var directory in self.GetDirectories())
            {
                directory.Copy(destination.CreateSubdirectory(directory.Name), recursively);
            }
        }
    }
}
