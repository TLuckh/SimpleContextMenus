This project is copied from https://github.com/dwmkerr/apex, 
the project in Apex/Core/Apex.WinForms.

A slight bugfix has been applied since OneDrive folders 
cannot be enumerated anymore using IShellFolder.

Since this bug seems to be limited to the ServerManger,
the fix is simply excluding all folders which, when queried with IShellFolder,
return REGDB_E_CLASSNOTREG  (among which is the OneDrive folder). 