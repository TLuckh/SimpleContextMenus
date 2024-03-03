using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using SimpleContextMenus.Properties;
using VLCShellExtension;

namespace SimpleContextMenus
{
    
    /// <summary>
    /// The CountLinesExtensions is an example shell context menu extension,
    /// implemented with SharpShell. It adds the command 'Count Lines' to text
    /// files.
    /// </summary>
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.Class, @"Directory\Background")]
    [COMServerAssociation(AssociationType.AllFiles)]
    public class SimpleContextMenu : SharpContextMenu
    {
        private List<string>? _selectedItemPaths;
        public string GetFolderPath() => FolderPath;

        public string GetExePath() => Assembly.GetExecutingAssembly().Location ??
                                      throw new Exception("Could not get the executing assembly location.");

        public List<string> GetSelectedItemPaths()
        {
            if (_selectedItemPaths == null)
            {
                if (SelectedItemPaths == null)
                    _selectedItemPaths = new List<string>();
                else
                    _selectedItemPaths = SelectedItemPaths.ToList();
            }

            return _selectedItemPaths;
        }


        /// <summary>
        /// Determines whether this instance can a shell
        /// context show menu, given the specified selected file list
        /// </summary>
        /// <returns>
        /// <c>true</c> if this instance should show a shell context
        /// menu for the specified file list; otherwise, <c>false</c>
        /// </returns>
        protected override bool CanShowMenu()
        {
            //  We always show the menu
            return true;
        }

        /// <summary>
        /// Creates the context menu. This can be a single menu item or a tree of them.
        /// </summary>
        /// <returns>
        /// The context menu for the shell context menu.
        /// </returns>
        protected override ContextMenuStrip CreateMenu()
        {
            var exe_directory = System.IO.Path.GetDirectoryName(GetExePath()) ??
                                throw new Exception("No directory found for the executing assembly.");
            //  Create the menu strip
            var menu = new ContextMenuStrip();
            //  Create the base menu item
            var extensionBaseItem = new ToolStripMenuItem
            {
                Text = Resources.SimpleContextMenu_CreateMenu_Extensions,
                Image = Resources.Extension_Menu
            };


            AddMenuItems(extensionBaseItem, exe_directory);
            menu.Items.Add(extensionBaseItem);
            // Clicking on the "Extensions" menu item opens the directory which mirrors the context menu structure
            // Neither option works if it the menu has DropDownItems...
            void OpenExplorer(object sender, EventArgs args) => Process.Start("explorer.exe" , $"{exe_directory}");
            extensionBaseItem.DropDownItemClicked += OpenExplorer;
            extensionBaseItem.Click += OpenExplorer;
             

            return menu;
        }

        private void AddMenuItems(ToolStripMenuItem menu, string currentDirectory)
        {
            // Note:
            // Directory.GetFileSystemEntries() returns directories without a trailing backslash.
            foreach (var filePathFull in Directory.GetFileSystemEntries(currentDirectory))
            {
                // Check if we should show the file/directory by checking it against the selection
                // (or all elements in the directory if nothing is selected)
                DataClass1 dataPoint =
                    NamingConventionParser(filePathFull);
                string displayName = dataPoint.DisplayName; 
                List<string> mimeTypes = dataPoint.MimeTypes;
                List<string> fileExtensions = dataPoint.FileExtensions;     
                bool applicable = IsAnyMimeTypeOrFileExtensionApplicableToSelectedItems(mimeTypes, fileExtensions);
                if (!applicable)
                    continue;


                var menuItem = new ToolStripMenuItem
                {
                    Text = displayName,
                };
                menu.DropDownItems.Add(menuItem);

                //  If it's a file, pressing it launches the corresponding (e.g. Python) script
                if (!File.GetAttributes(filePathFull).HasFlag(FileAttributes.Directory))
                {
                    menuItem.Click += (sender, args) =>
                    {
                        // Konversion in Argumentliste:
                        StringBuilder stringBuilder = new StringBuilder();

                        foreach (string fileToAdd in GetSelectedItemPaths())
                        {
                            stringBuilder.Append($"\"{fileToAdd}\" ");

                        }

                        System.Diagnostics.Process process = new System.Diagnostics.Process();
                        System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo
                        {
                            WindowStyle = ProcessWindowStyle.Normal,
                            FileName = filePathFull,
                            WorkingDirectory = GetFolderPath(),
                            Arguments = stringBuilder.ToString()
                        };
                        process.StartInfo = startInfo;
                        process.Start();
                    };
                }
                // And if it's a directory, we recursively add the items in the directory to the menu
                else
                {
                    AddMenuItems(menuItem, filePathFull);
                    // Don't show empty submenu's
                    if (!menuItem.HasDropDownItems)
                        menuItem.Visible = false;
                }

            }

        }

        /// <summary>
        ///  If we get no selection (GetSelectedItemPaths()), select all files & folders in the current directory.
        /// For the selection, check if any of the selected items has a  matching MIME types or file extension.
        /// Note that any selected folder is interpreted as a file with file extension 'folder'.
        /// 
        /// </summary>
        /// <param name="mimeTypes"></param>
        /// <param name="fileExtensions"></param>
        /// <returns>Whether or not there's an overlap between the given MIME types and file extensions and those of the selected items.</returns>
        private bool IsAnyMimeTypeOrFileExtensionApplicableToSelectedItems(List<string> mimeTypes,
            List<string> fileExtensions)
        {
            // If neither mimeTypes nor fileExtensions are given, we assume that the item is applicable.
            if (mimeTypes.Count == 0 && fileExtensions.Count == 0)
                return true;
            
            // Parsing the selected items to their MIME types and file extensions.

            // var x1 = GetFolderPath();
            // var x2 = Directory.GetFileSystemEntries(x1);
            // var x3 = x2.ToList();
            // var x4 = GetSelectedItemPaths();
            var itemPathsToMatch = GetSelectedItemPaths();
            if (itemPathsToMatch.Count == 0)
                itemPathsToMatch = Directory.GetFileSystemEntries(GetFolderPath()).ToList();

            List<string> mimeTypesOfSelection =
                itemPathsToMatch
                    .Where(x => !File.GetAttributes(x).HasFlag(FileAttributes.Directory))
                    // Maps the file path to the MIME type, of which we just want the coarse type.
                    .Select(x => MIMEAssistant.GetMIMEType(x).Split('/')[0]) 
                    .ToList();
            List<string> fileExtensionsOfSelection =
                itemPathsToMatch
                    .Where(x => !File.GetAttributes(x).HasFlag(FileAttributes.Directory))
                    .Select(x => Path.GetExtension(x).Remove(0, 1))
                    .ToList();

            if (GetSelectedItemPaths().Any(x => File.GetAttributes(x).HasFlag(FileAttributes.Directory)))
            {
                fileExtensionsOfSelection.Add("folder");
            }

            // Comparing the MIME types and file extensions of the selection to the MIME types and file extensions passed into the method.
            
            return 
                mimeTypesOfSelection.Intersect(mimeTypes).Any() 
                ||
                fileExtensionsOfSelection.Intersect(fileExtensions).Any();



        }

        /// <summary>
        /// Takes in a full file path and turns it, according to the naming convention, into a tuple:
        /// (displayName, MIMETypes, FileExtensions)
        /// 
        /// The naming convention is as follows:
        /// The file name is split by each dot. The first part is the display name,
        /// and each following part (except the last for files, i.e. non-folders) is either a MIME type or a file extension.
        /// The last part is ignored. Hint: If you want to call a Python script,
        /// it should be .py if the script should run in foreground, or .pyw if the script should run in background.
        ///
        /// Each middle part should have all its letters in lower case if it's a file extension, and in upper case if it's a MIME type.
        ///
        /// The returned MIME types and file extensions are in lower case and without dot.
        /// </summary>
        /// <param name="filePathFull"> A full path to the file.</param> 
        /// <exception cref="NotImplementedException"></exception>
        private DataClass1 NamingConventionParser(
            string filePathFull)
        {
            // Get rid of the file extension & directory prefixes
            if (!File.GetAttributes(filePathFull).HasFlag(FileAttributes.Directory))
                filePathFull = Path.GetFileNameWithoutExtension(filePathFull);
            else
                filePathFull = Path.GetFileName(filePathFull);
            
            
            
            
            var parts = filePathFull.Split('.').ToList();
            string displayName = parts[0];
            List<string> middleParts = parts.Skip(1).ToList();
            

            List<string> mimeTypes = new List<string>();
            List<string> fileExtensions = new List<string>();

            foreach (var part in middleParts)
            {
                if (part == part.ToUpper())
                    mimeTypes.Add(part.ToLower());
                else
                    fileExtensions.Add(part);
            }

            return new DataClass1(displayName, mimeTypes, fileExtensions);
        }


        internal class DataClass1(string displayName, List<string> mimeTypes, List<string> fileExtensions)
        {
            public string DisplayName { get; set; } = displayName;
            public List<string> MimeTypes { get; set; } = mimeTypes;
            public List<string> FileExtensions { get; set; } = fileExtensions;
        }



    }
}