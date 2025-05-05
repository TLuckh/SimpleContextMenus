using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using MimeTypes;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using SimpleContextMenus.Properties;

namespace SimpleContextMenus;

/// <summary>
/// </summary>
[ComVisible(true)]
[COMServerAssociation(AssociationType.Class, @"Directory\Background")]
[COMServerAssociation(AssociationType.AllFilesAndFolders)]
public class SimpleContextMenu : SharpContextMenu
{
    private List<string>? _selectedItemPaths;

    public new string FolderPath
    {
        get
        {
            List<string> selectedItems = GetSelectedItemPaths();
            if (selectedItems.Count == 0) // Case: No selection. Here FolderPath works as intended in the base-class
                return base.FolderPath;
            // Case: Selection. Here, FolderPath from base-class is bugged and returns the path of system32.
            //If all the files are in the same folder, we return its path.
            //Otherwise, nothing is returned as we might not be in a folder at all (e.g. 'Recently Used', or 'My Computer'),
            //which is detected in CanShowMenu(). 
            List<string> pathsOfSelectedItems = selectedItems.Select(Path.GetDirectoryName).Distinct().ToList();
            if (pathsOfSelectedItems.Count == 1)
                return pathsOfSelectedItems[0];
            return "";
        }
    }

    public string GetFolderPath()
    {
        return FolderPath;
    }

    /// <summary>
    ///     Returns the full path to the subfolder "Extensions", which lies in the same folder as the executing assembly.
    /// </summary>
    /// <returns>The path to the folder of the executing assembly, i.e. the folder containing the COM Server .dll'.</returns>
    public string GetExtensionsFolderPath()
    {
        return Path.Combine(
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Extensions");
    }

    /// <summary>
    ///     Returns the full path to the subfolder "TopLevelItems", which lies in the same folder as the executing assembly.
    /// </summary>
    /// <returns>The path to the folder of the executing assembly, i.e. the folder containing the COM Server .dll'.</returns>
    public string GetTopLevelItemsFolderPath()
    {
        return Path.Combine(
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "TopLevelItems");
    }

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
    ///     Determines whether this instance can a shell
    ///     context show menu, given the specified selected file list
    /// </summary>
    /// <returns>
    ///     <c>true</c> if this instance should show a shell context
    ///     menu for the specified file list; otherwise, <c>false</c>
    /// </returns>
    protected override bool CanShowMenu()
    {
        return FolderPath.Length > 0; // See FolderPath for explanation.
    }

    /// <summary>
    ///     Creates the context menu. This can be a single menu item or a tree of them.
    /// </summary>
    /// <returns>
    ///     The context menu for the shell context menu.
    /// </returns>
    protected override ContextMenuStrip CreateMenu()
    {
        
        //  Create the menu strip
        ContextMenuStrip menuStrip = new();
        menuStrip.Items.Add(new ToolStripSeparator());

        //  Create the base menu item "Extensions"
        ToolStripMenuItem extensionBaseItem = new()
        {
            Text = Resources.SimpleContextMenu_CreateMenu_Extensions,
            Image = Resources.Extension_Menu
        };

        try         // ToDo: Add actual error handling with a log, so this ugly try-catch that I deactivate in no-debug anyway can go away
        {
            // Add all submenus to "Extensions"
            bool subItemIsApplicable = AddMenuItems(menuStrip, extensionBaseItem, GetExtensionsFolderPath());
            // Show extensions in the context menu if there's at least one applicable item in a submenu.
            if (subItemIsApplicable)
                menuStrip.Items.Add(extensionBaseItem);

            // Add all top-level items to the context menu
            AddMenuItems(menuStrip, null, GetTopLevelItemsFolderPath());
            menuStrip.Items.Add(new ToolStripSeparator());
            return menuStrip;
        }
        catch (Exception e) when (IsDebug())
        {
            menuStrip.Items.Add(new ToolStripMenuItem(e.Message));
            return menuStrip;
        }


        // // Clicking on the "Extensions" menu item opens the directory which mirrors the context menu structure
        // // Neither option works if it the menu has DropDownItems...
        // void OpenExplorer(object sender, EventArgs args) => Process.Start("explorer.exe" , $"{exe_directory}");
        // extensionBaseItem.DropDownItemClicked += OpenExplorer;
        // extensionBaseItem.Click += OpenExplorer;
        // extensionBaseItem.DoubleClick += OpenExplorer;
    }

    /// <summary>
    ///     Builds the recursive menu structure
    ///     by mirroring the file &amp; folder structure in currentDirectory.
    ///     Each folder is turned into a dropdown menu, and each file into a menu end item.
    /// </summary>
    /// <param name="menuStrip"> The context menu itself</param>
    /// <param name="menu"> The current menu item in the context menu, or null if we're in no menu item (i.e. in the top level)</param>
    /// <param name="currentDirectory"></param>
    /// <returns>
    ///     Returns whether any end-item in the menu structure is visible.
    ///     This is mostly used by the calling element to decide whether its menu item should be visible.
    /// </returns>
    private bool AddMenuItems(
        ContextMenuStrip menuStrip,
        ToolStripMenuItem? menu,
        string currentDirectory)
    {
        bool anyEndItemApplicable = false;
        // Note:
        // Directory.GetFileSystemEntries() returns directories without a trailing backslash.
        
        // Iterates over the elements which shall be turned into menu items:
        // It iterates within the current directory first over the visible folders, then over the visible files.
        // Each such object is mapped to its relevant attributes,
        // namely its display name, the mime types and file extensions for which it is applicable, and its full path.
        foreach (FileAttributes menuItemAttributes in
                 Directory.GetDirectories(currentDirectory).Union(Directory.GetFiles(currentDirectory))
                     .Where(x => !File.GetAttributes(x).HasFlag(System.IO.FileAttributes.Hidden))
                     .Select(FileAttributes.NamingConventionParser))
        {
            
            ToolStripMenuItem menuItem = new()
            {
                Text = menuItemAttributes.DisplayName
            };

            #region Error handling

            if (menuItemAttributes.IsError)
            {
                // So far only error caught is when a .lnk has invalid target.
                // ToDo: Make more general, and make the corresponding context menu clickable such that it returns the error.
                // Not really a fan of having it as field, but Union types aren't supported (yet). Maybe let try-catch return a Subtype of FileAttributes?
                // Or move the .Select out of the foreach and put a try-catch around the call to NamingConventionParser?
                if (menu != null)
                    menu.DropDownItems.Add(menuItem);
                else
                    menuStrip.Items.Add(menuItem);


                anyEndItemApplicable = true;
                // menuItem.Click += launchScriptOnMenuItemClick(menuItemAttributes); // Throws some error if added as is
                continue;
            }
            #endregion

            
            
            // Check if we should show the file/directory by checking it against the selection
            // (or all elements in the directory if nothing is selected)
            if (!IsAnyMimeTypeOrFileExtensionApplicableToSelectedItems(menuItemAttributes.MimeTypes, 
                                                                       menuItemAttributes.FileExtensions))
                continue;



            // Whether or not we add menuItem as a drop down item to menu depends on whether it's an end item or at least contains one.
            // I've tried doing this using visibility, but for some reason that tag gets ignored...

            //  If it's a file, pressing it launches the corresponding (e.g. Python) script
            if (!File.GetAttributes(menuItemAttributes.FilePathFull).HasFlag(System.IO.FileAttributes.Directory))
            {
                // If we're in TopLevelFolder, then we need to add directly the menuStrip.
                // Otherwise we add to the menu as submenu.
                if (menu != null)
                    menu.DropDownItems.Add(menuItem);
                else
                    menuStrip.Items.Add(menuItem);

                anyEndItemApplicable = true;
                menuItem.Click += launchScriptOnMenuItemClick(menuItemAttributes);
            }
            // And if it's a directory, we recursively add the items in the directory to the menu
            else
            {
                // We build the submenus of menuItem corresponding to filePathFull, but only 
                // show it, if at least one end item somewhere in the submenu is applicable.
                bool subItemIsApplicable = AddMenuItems(menuStrip, menuItem, menuItemAttributes.FilePathFull);
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
    ///     To be called when a menu item is clicked.
    ///     Launches the script in the location corresponding to the menu item.
    ///     Passes those elements of the selection, which fit at least one MIME type or file extension
    ///     as arguments to the script (a background click has no selected items!).
    ///     Sets the working directory of the script to be the folder in which the right click occured.
    /// </summary>
    /// <param name="fileAttributes"> The full path to the working directory in which the script should be launched as well as filter info for file selection</param>
    /// <returns></returns>
    private EventHandler launchScriptOnMenuItemClick(FileAttributes fileAttributes)
    {
        return (sender, args) =>
        {
            Directory.SetCurrentDirectory(GetFolderPath());

            StringBuilder argumentsToPass = new();
            // Add all selected items which match at least one of the given mime types or file extensions to the argument list
            #region CreateArgumentListForScript

            foreach (string fileToAdd in GetSelectedItemPaths())
            {
                 var (mimeTypesOfSelection, fileExtensionsOfSelection) = GetMimeTypesAndFileExtensions([fileToAdd]);

                 if (!fileAttributes.HasAnyMimeTypeOrFileExtension()
                     || fileAttributes.MimeTypes.Intersect(mimeTypesOfSelection).Any()
                     || fileAttributes.FileExtensions.Intersect(fileExtensionsOfSelection).Any())
                 {
                     argumentsToPass.Append($"\"{fileToAdd}\" ");
                 }

            }
            #endregion

            Process process = new();
            ProcessStartInfo startInfo = new()
            {
                WindowStyle = ProcessWindowStyle.Normal,
                FileName = fileAttributes.FilePathFull,
                WorkingDirectory = GetFolderPath(),
                Arguments = argumentsToPass.ToString()
            };
            process.StartInfo = startInfo;
            process.Start();
        };
    }



    /// <summary>
    ///     If we get no selection (GetSelectedItemPaths()), select all files & folders in the current directory.
    ///     For the selection, check if any of the selected items has a  matching MIME types or file extension.
    ///     Note that any selected folder is interpreted as a file with file extension 'folder'.
    /// </summary>
    /// <param name="mimeTypes"></param>
    /// <param name="fileExtensions"></param>
    /// <returns>
    ///     Whether or not there's an overlap between the given MIME types and file extensions and those of the selected
    ///     items.
    /// </returns>
    private bool IsAnyMimeTypeOrFileExtensionApplicableToSelectedItems(List<string> mimeTypes,
        List<string> fileExtensions)
    {
        // If neither mimeTypes nor fileExtensions are given, we assume that the item is applicable.
        if (mimeTypes.Count == 0 && fileExtensions.Count == 0)
            return true;

        // Parsing the selected items to their MIME types and file extensions.
        // // Debug only vars:
        // var x1 = GetFolderPath();
        // var x2 = Directory.GetFileSystemEntries(x1);
        // var x3 = x2.ToList();
        // var x4 = GetSelectedItemPaths();
        List<string> itemPathsToMatch = GetSelectedItemPaths();
        if (itemPathsToMatch.Count == 0)
        {
            itemPathsToMatch = Directory.GetFileSystemEntries(GetFolderPath()).ToList();

            // To not slow down the explorer or get timeout issues, we simply show everything if the selection is too big.
            // If the selection was given by the user, then timeout issues hopefully aren't a concern.
            if (itemPathsToMatch.Count >= 50)
                return true;
        }

        // Get the MIME types and file extension for each of the selected items:

        (List<string> mimeTypesOfSelection, List<string> fileExtensionsOfSelection) = GetMimeTypesAndFileExtensions(itemPathsToMatch);

        // Comparing the MIME types and file extensions of the selection to the MIME types and file extensions passed into the method.

        return
            mimeTypesOfSelection.Intersect(mimeTypes).Any()
            ||
            fileExtensionsOfSelection.Intersect(fileExtensions).Any();
    }

    private (List<string> mimeTypesOfSelection, List<string> fileExtensionsOfSelection) GetMimeTypesAndFileExtensions(List<string> itemPathsToMatch)
    {
        #region GetMimeTypesAndFileExtensions

        List<string> mimeTypesOfSelection = [];
        List<string> fileExtensionsOfSelection = [];
        
        if (GetSelectedItemPaths().Any(x => File.GetAttributes(x).HasFlag(System.IO.FileAttributes.Directory)))
            fileExtensionsOfSelection.Add("folder");
        
        itemPathsToMatch = itemPathsToMatch.Where(x => !File.GetAttributes(x).HasFlag(System.IO.FileAttributes.Directory)).ToList();
        foreach (string itemPath in itemPathsToMatch)
        {
            if (File.GetAttributes(itemPath).HasFlag(System.IO.FileAttributes.Directory))
            {
                continue;
            }

            string mimeTypeFull = MimeTypeMap.GetMimeType(itemPath).ToLower();
            string mimeTypeRoughType = MimeTypeMap.GetMimeType(itemPath).Split('/')[0].ToLower();
            mimeTypesOfSelection.Add(mimeTypeFull);
            mimeTypesOfSelection.Add(mimeTypeRoughType);

            string fileExtension = Path.GetExtension(itemPath).Remove(0, 1);
            fileExtensionsOfSelection.Add(fileExtension);
        }

        #endregion

        return (mimeTypesOfSelection, fileExtensionsOfSelection);
    }


    /// <summary>
    /// The underlying data model for the view of each menu item.
    /// Represents attributes for a file, including its display name, associated MIME types, and file extensions.
    /// </summary>
    class FileAttributes(
        string displayName,
        List<string> mimeTypes,
        List<string> fileExtensions,
        string filePathFull,
        bool isError)
    {
        /// <summary>
        ///     Takes in a full file path and turns it, according to the naming convention, into a tuple:
        ///     (displayName, MIMETypes, FileExtensions, resolved [for .lnk] full path) <br></br>
        ///     For the naming convention applied, see NamingConvention.md.
        ///     The returned MIME types and file extensions are in lower case and without dot.
        /// </summary>
        /// <param name="filePathFullOfMenuItem"> A path to the file. Both absolute and relative paths are accepted.</param>
        public static FileAttributes NamingConventionParser(
                string filePathFullOfMenuItem)
            {
                string filePath = filePathFullOfMenuItem;

                try
                {
                    while (Path.GetExtension(filePathFullOfMenuItem) == ".lnk")
                        filePathFullOfMenuItem = ShellLink.GetShortcutTarget(filePathFullOfMenuItem);
                    
                    if (!File.Exists(filePathFullOfMenuItem) && !Directory.Exists(filePathFullOfMenuItem))
                        throw new System.IO.FileNotFoundException($"File or directory not found: {filePathFullOfMenuItem}");
                }
                catch (Exception e) when (e is System.IO.DirectoryNotFoundException || e is System.IO.FileNotFoundException) 
                {
                    return new FileAttributes($"""Shortcut "{filePath}" not found""", [], [],"",true);
                }





                // Get rid of the file extension & directory prefixes
                if (!File.GetAttributes(filePath).HasFlag(System.IO.FileAttributes.Directory))
                    filePath = Path.GetFileNameWithoutExtension(filePath);
                else
                    filePath = Path.GetFileName(filePath);
        
        
                List<string> parts = filePath.Replace("..", "/").Split('.').ToList();
                string displayName = parts[0];
                List<string> middleParts = parts.Skip(1).ToList();
        
        
                List<string> mimeTypes = new();
                List<string> fileExtensions = new();
        
                foreach (string? part in middleParts)
                    if (part == part.ToUpper())
                        mimeTypes.Add(part.ToLower());
                    else
                        fileExtensions.Add(part);
        
                return new FileAttributes(displayName, mimeTypes, fileExtensions,filePathFullOfMenuItem,false);
            }
        public string DisplayName { get; set; } = displayName;
        public List<string> MimeTypes { get; set; } = mimeTypes;
        public List<string> FileExtensions { get; set; } = fileExtensions;
        
        public string FilePathFull { get; set; } = filePathFull;
        
        public bool IsError { get; set; } = isError;
        
        public bool HasAnyMimeTypeOrFileExtension()
        {
            return MimeTypes.Count > 0 || FileExtensions.Count > 0;
        }
    }

    private bool IsDebug()
    {
#if DEBUG
    return true;
#else
        return false;
#endif
    }
}

