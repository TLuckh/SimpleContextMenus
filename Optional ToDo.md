Implement lazy loading of context menus in "Extensions" folder.

(by System.Lazy (see SharpContextMenu.cs Lazy<...>) or by only calling

// Add all submenus to "Extensions"
bool subItemIsApplicable = AddMenuItems(menuStrip, extensionBaseItem, GetExtensionsFolderPath());

when the "Extensions" submenu is opened, i.e. an event of type DropDownOpening is fired)

As far as I tried, this seems unnecessary, as it's fast enough even with ~100 items in the subfolders in Extensions.

## Tips
If you add many context menus, try placing most of them in the 'Extensions' folder.
Not only avoids this cluttering the context menu, but also stops right clicking from becoming
sluggish if you have very many items to add to the context menu.