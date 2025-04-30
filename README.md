# SimpleContextMenus

Delivers a COM Server which is associated with files, folders, and the folder background.

This Server allows you to create a local file structure which will be mimicked within the context menu of Windows
Explorer.

Each file in the local file structure then gets run when you press the corresponding
item of the context menu.

## System Requirements
- Windows 10 (not tested on Windows 11)
- .NET Framework 4.8.1 or a compatible version ()

## Installation

To install the server, take the build output of this project and
place it into a directory of your choosing. Then, run the '#Install.bat'.

If the installation was successful, you should see a
new context menu called "Extensions" when right clicking in Windows Explorer.

Possibly you need to restart your explorer.exe process.

## Deinstallation

To uninstall the server, run the '#Uninstall.bat' which is located in
the directory where you installed the server.
To be able to delete the files, one has to additionally restart explorer.exe once.

## Configuration

In the directory where you installed the server, there are two subfolders: 'Extensions' and 'TopLevelItems'.
Here, you can add any number of files and folders.
<br>
(Elements in 'TopLevelItems' each get an item in the context menu, while elements in
'Extensions' get an item in the submenu named 'Extensions')

For each folder, a drop down menu in the context menu will be created.
For each file, a simple context menu entry will be created.

For the naming convention of files and folders, see [NamingConvention.md](NamingConvention.md). <br>
Their name will determine to which of selected items 
(or the items in the folder for a background click) 
the context menu submenu/item is applicable.

A context menu item will only be visible if at least one selected item is applicable.
A context menu submenu will only show if it contains at least one context menu item that is visible.

Furthermore, if a folder has many files (&ge; 50), the context menus will always be shown
(this does not apply when right-clicking on a selection).


Clicking on an entry in the context menu will:
1. Open the corresponding file in the local file structure
(using the standard association for the program)
2. Pass it each of the matching files, if any as, arguments (each wrapped in quotes).
    - Note that the file extension restrictions you encoded into the file-name affect the passed selection:<br>
Only those of the selected items
which match at least one of the given file extensions or MIME types
will be passed to the file.

Furthermore, the current working directory (in Python e.g. os.getcwd())
will be set to the directory in which you right-clicked.

For examples see [NamingConvention.md](NamingConvention.md)

## Odd Behavior

Note that if you select items and right click, the items are passed in the order in which they are
shown in Windows Explorer, but starting
from the item you right clicked on, and wrapping around.

That is, if in the explorer you select items 1, 2, 3, 4, 5, which are shown in this order,
and you right click on item 3, then order in which the elements are passed is 3,4,5,1,2,
and the passed string is

`"1" "2" "3" "4" "5" `

In general, any passed parameter is wrapped in quotes, and separated by a space.

