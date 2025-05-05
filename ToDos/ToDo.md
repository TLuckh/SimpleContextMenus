# Regression Tests
Add files to the Debug configuration that make the following tests quick to do:
- Test that the Example Script.py works correctly, i.e. it shows:
- Test that the filters for file extensions/MIME types work correctly, so only if an applicable item is in the folder/selected items, 
the folders with the file extensions/MIME types are shown
- Test that folders without files, and folders with no applicable items aren't shown
- Test that if a folder/file has a file extensions/MIME type it filters for, that only matching selected items are passed to it
- Test that if you execute Example Script.py after it has moved by putting a shortcut to it into the TopLevelFolder/Extensions folder, it correctly shows the exectued file location.
- Test that if a shortcut has been orphaned (file deleted, folder deleted, drive unplugged), the menu item will be replaced by an error message, without any other menu item being missing








# Features
- Add support for pictures for contex menus. Probably by excluding all files starting with a dot.
- Add a STRICT switch, so that an option is only shown if all selected files match the filters.



# Examples
- Add example for folder with switches
- Add mp3 converter example
- Add Diashow converter example