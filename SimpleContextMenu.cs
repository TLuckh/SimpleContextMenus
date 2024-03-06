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
    /// 
    /// </summary>
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.Class, @"Directory\Background")]
    [COMServerAssociation(AssociationType.AllFilesAndFolders)]
    public class SimpleContextMenu : SharpContextMenu
    {
        private List<string>? _selectedItemPaths;
        public string GetFolderPath() => FolderPath;

        /// <summary>
        /// Returns the full path to the subfolder "Extensions", which lies in the same folder as the executing assembly.
        /// </summary>
        /// <returns>The path to the folder of the executing assembly, i.e. the folder containing the COM Server .dll'.</returns>
        public string GetExtensionsFolderPath() => 
            Path.Combine(
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Extensions");
        
        /// <summary>
        /// Returns the full path to the subfolder "TopLevelItems", which lies in the same folder as the executing assembly.
        /// </summary>
        /// <returns>The path to the folder of the executing assembly, i.e. the folder containing the COM Server .dll'.</returns>
        public string GetTopLevelItemsFolderPath() => 
            Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "TopLevelItems");

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
            //  Create the menu strip
            var menuStrip = new ContextMenuStrip();
            menuStrip.Items.Add(new ToolStripSeparator());
            
            //  Create the base menu item "Extensions"
            var extensionBaseItem = new ToolStripMenuItem
            {
                Text = Resources.SimpleContextMenu_CreateMenu_Extensions,
                Image = Resources.Extension_Menu
            };
            // Add all submenus to "Extensions"
            AddMenuItems(menuStrip, extensionBaseItem, GetExtensionsFolderPath());
            // Show extensions in the context menu
            menuStrip.Items.Add(extensionBaseItem);

            // Add all top-level items to the context menu
            AddMenuItems(menuStrip, null, GetTopLevelItemsFolderPath());
            //
            // foreach (var filePathFull in 
            //          Directory.GetDirectories(GetTopLevelItemsFolderPath())
            //              .Union(Directory.GetFiles(GetTopLevelItemsFolderPath())))
            // {
            //     AddMenuItems(menuStrip, null, GetTopLevelItemsFolderPath());
            // }



            
            
            // // Clicking on the "Extensions" menu item opens the directory which mirrors the context menu structure
            // // Neither option works if it the menu has DropDownItems...
            // void OpenExplorer(object sender, EventArgs args) => Process.Start("explorer.exe" , $"{exe_directory}");
            // extensionBaseItem.DropDownItemClicked += OpenExplorer;
            // extensionBaseItem.Click += OpenExplorer;
            // extensionBaseItem.DoubleClick += OpenExplorer;
             

            menuStrip.Items.Add(new ToolStripSeparator());
            return menuStrip;
        }

        /// <summary>
        /// Builds the recursive menu structure
        /// by mirroring the file &amp; folder structure in currentDirectory.
        /// Each folder is turned into a dropdown menu, and each file into a menu end item.
        /// </summary>
        /// <param name="menuStrip"> The context menu itself</param>
        /// <param name="menu"> The current menu item in the context menu, or null if we're in no menu item (i.e. in the top level)</param>
        /// <param name="currentDirectory"></param>
        /// <returns>Returns whether any end-item in the menu structure is visible.
        /// This is mostly used by the calling element to decide whether its menu item should be visible.</returns>
        private bool AddMenuItems(
            ContextMenuStrip menuStrip, 
            ToolStripMenuItem? menu, 
            string currentDirectory)
        {
            bool anyEndItemApplicable = false;
            // Note:
            // Directory.GetFileSystemEntries() returns directories without a trailing backslash.
            foreach (var filePathFull in 
                     Directory.GetDirectories(currentDirectory).
                         Union(Directory.GetFiles(currentDirectory))
                         .Where(x => !File.GetAttributes(x).HasFlag(FileAttributes.Hidden) )) 
                // First directories, then files. Only list stuff that's visible in Windows Explorer
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
                // Whether or not we add menuItem as a drop down item to menu depends on whether it's an end item or at least contains one.
                // I've tried doing this using visibility, but for some reason that teg gets ignored...

                //  If it's a file, pressing it launches the corresponding (e.g. Python) script
                if (!File.GetAttributes(filePathFull).HasFlag(FileAttributes.Directory))
                {
                    // If we're in TopLevelFolder, then we need to add directly the menuStrip.
                    // Otherwise we add to the menu as submenu.
                    if (menu != null)
                        menu.DropDownItems.Add(menuItem);
                    else
                        menuStrip.Items.Add(menuItem);
                    
                    anyEndItemApplicable = true;
                    menuItem.Click += launchScriptOnMenuItemOnClick(filePathFull);
                }
                // And if it's a directory, we recursively add the items in the directory to the menu
                else
                {
                    // We build the submenus of menuItem corresponding to filePathFull, but only 
                    // show it, if at least one end item somewhere in the submenu is applicable.
                    bool subItemIsApplicable = AddMenuItems(menuStrip,menuItem, filePathFull);
                    if (subItemIsApplicable)
                    {
                        // If we're in TopLevelFolder, then we need to add directly the menuStrip.
                        // Otherwise we add to the menu as submenu.
                        if (menu != null)
                            menu.DropDownItems.Add(menuItem);
                        else
                            menuStrip.Items.Add(menuItem);
                    }

                    anyEndItemApplicable |= subItemIsApplicable;

                    // // Leftover from when I tried to hide empty submenus instead of only adding them if they're not empty .
                    // // Don't show empty submenus
                    // if (!menuItem.HasDropDownItems || !subItemIsApplicable)
                    //     menuItem.Visible = false;
                    // // else if (!menuItem.DropDownItems.Cast<ToolStripItem>().Any(menuItemDropDownItem => menuItemDropDownItem.Visible))
                    // //         menuItem.Visible = false;
                }
            }
            
            return anyEndItemApplicable;
        }

        /// <summary>
        /// To be called when a menu item is clicked.
        /// Launches the script in the location corresponding to the menu item.
        ///
        /// Passes the selection as arguments to the script (a background click has no selected items).
        /// Sets the working directory of the script to be the folder in which the right click occured.
        /// </summary>
        /// <param name="filePathFull"> The full path to the working directory in which the script should be launched</param>
        /// <returns></returns>
        private EventHandler launchScriptOnMenuItemOnClick(string filePathFull)
        {
            return (sender, args) =>
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
            {
                itemPathsToMatch = Directory.GetFileSystemEntries(GetFolderPath()).ToList();

                // To not slow down the explorer or get timeout issues, we simply show everything if the selection is too big.
                // If the selection was given by the user, then timeout issues hopefully aren't a concern.
                if (itemPathsToMatch.Count >= 50)
                    return true;
            }


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
        /// Starting with a file name with extension, but without path to it in its name:
        /// The file name is split by each dot. The first part is the display name,
        /// and each following part (except the last for files, i.e. non-folders) is either a MIME type (if it's uppercase)
        /// or a file extension (if it's lowercase).
        /// The last part (if it is a file extension) is ignored.
        /// Hint: If you want to call a Python script,
        /// it should be .py if the script should run in foreground, or .pyw if the script should run in background.
        ///
        /// Each middle part should have all its letters in lower case if it's a file extension, and in upper case if it's a MIME type.
        ///
        /// The returned MIME types and file extensions are in lower case and without dot.
        /// </summary>
        /// <param name="filePathFull"> A path to the file. Both absolute and relative paths are accepted.</param> 
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