# SimpleContextMenus
Delivers a COM Server which is associated both with all files and the folder background.

This Server allows you to create context menu entries which
mimic the local file structure (of wherever you placed and installed the server).

## Installation
To install the server, take the build output of this project and
place it into a directory of your choosing. Then, run the following command 
in cmd as an administrator, in that directory:

`ServerRegistrationManager.exe install SimpleContextMenus.dll -codebase`

Possibly you need to restart your explorer.exe process.

## Deinstallation
To uninstall the server, run the following command in cmd as an administrator, in the directory where you installed the server:
`ServerRegistrationManager.exe uninstall SimpleContextMenus.dll -codebase`

## Configuration
In the directory where you installed the server, you 
now can add any number of files and folders.

For each folder, a drop down menu in the context menu will be created.
For each file, a simple context menu entry will be created.

Each file has the following naming convention:
See SimpleContextMenu.cs/NamingConventionParser documentation.

Note that the file extension restrictions do not affect the passed selection,
only whether or not the context menu is visible.

So clicking the file in the context menu will execute cmd, which in turn calls Python 
with the following arguments:
1. The full path to the Python script.
2. Each of the selected files, if any.

Furthermore, os.getcwd() will return the directory in which you right clicked.

