# SimpleContextMenus
Delivers a COM Server which is associated with files, folders, and the folder background.

This Server allows you to create a local file structure which will be mimicked within the context menu of Windows Explorer.

Each file in the local file structure then gets run when you press the corresponding
item of the context menu.

## Installation
To install the server, take the build output of this project and
place it into a directory of your choosing. Then, run the '#Install.bat'.

If the installation was successful, you should see a 
new context menu called "Extensions" when right clicking in Windows Explorer.

Possibly you need to restart your explorer.exe process.

## Deinstallation
To uninstall the server, run the '#Uninstall.bat' which is located in
the directory where you installed the server.


## Configuration
In the directory where you installed the server, navigate to the subfolder 'Extensions'.
Here, you can add any number of files and folders.

For each folder, a drop down menu in the context menu will be created.
For each file, a simple context menu entry will be created.

Each file has the following naming convention:
See SimpleContextMenu.cs/NamingConventionParser documentation.

Note that the file extension restrictions do not affect the passed selection,
only whether or not the context menu is visible.
Furthermore, if a folder has many files (>=50), the context menus will
not be filtered when doing a right click on the background 
(this does not apply when right clicking on a selection).

Clicking on an entry in the context menu will open the 
corresponding file in the local file structure 
(using the standard association for the program), 
and passes each of the selected files, if any as arguments (each wrapped in quotes).

Furthermore, the current working directory (in Python e.g. os.getcwd()) 
will be set to the directory in which you right clicked.

## Odd Behavior
Note that if you select items and right click, the items are passed as selected, but starting 
from the item you right clicked on, and wrapping around.

That is, if in the explorer you select items 1, 2, 3, 4, 5, which are shown in this order,
and you right click on item 3, then the passed elements are 3,4,5,1,2,
and the passed string is

`"1" "2" "3" "4" "5" `

In general, any passed parameter is wrapped in quotes, and separated by a space.

