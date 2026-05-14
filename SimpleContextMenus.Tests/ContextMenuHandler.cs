// using SharpShell.Interop;
// using SharpShell.SharpContextMenu;
//
// namespace SimpleContextMenus.Tests;
//
// using System;
// using System.Collections.Generic;
// using System.IO;
// using System.Runtime.InteropServices;
// using System.Threading;
//
//
// public class ContextMenuHandlerTester : IDisposable
// {
//     private readonly SharpContextMenu _handler;
//     private IntPtr _hMenu;
//     private IContextMenu _contextMenu;
//     private uint _idCmdFirst = 1;
//     private uint _idCmdLast  = 0x7FFF;
//
//     public ContextMenuHandlerTester(string path)
//     {
//         if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
//             throw new InvalidOperationException("STA-Thread erforderlich!");
//
//         // 1. Handler direkt instanziieren – kein COM, keine Registry
//         _handler = new SimpleContextMenu();
//
//         // 2. DataObject erstellen (simuliert Explorer-Übergabe)
//         var ppDataObject = DataObjectHelper.CreateForFile(path);
//
//         // 3. IShellExtInit.Initialize aufrufen
//         //    SharpShell liest hier die Dateipfade aus dem DataObject
//         var shellInit = (IShellExtInit)_handler;
//
//         int hr = NativeMethods.SHParseDisplayName(
//             Path.GetDirectoryName(path),
//             IntPtr.Zero,
//             out var pidlFolder, 0, out _);
//
//         shellInit.Initialize(pidlFolder, ppDataObject, IntPtr.Zero);
//
//         NativeMethods.ILFree(pidlFolder);
//         Marshal.Release(ppDataObject);
//
//         // 4. IContextMenu.QueryContextMenu aufrufen
//         _contextMenu = (IContextMenu)_handler;
//         _hMenu = NativeMethods.CreatePopupMenu();
//
//         _contextMenu.QueryContextMenu(
//             _hMenu, 0, _idCmdFirst, _idCmdLast,
//             NativeMethods.CMF_NORMAL);
//     }
//
//     public List<ContextMenuItem> ReadAllItems()
//     {
//         return ReadMenuItems(_hMenu);
//     }
//
//     // ReadMenuItems, GetVerb, Invoke etc. – 
//     // identisch wie in ShellContextMenuReader
//     private List<ContextMenuItem> ReadMenuItems(IntPtr hMenu)
//     {
//         var items = new List<ContextMenuItem>();
//         int count = NativeMethods.GetMenuItemCount(hMenu);
//
//         for (int i = 0; i < count; i++)
//         {
//             var mii = new MENUITEMINFO
//             {
//                 cbSize    = (uint)Marshal.SizeOf<MENUITEMINFO>(),
//                 fMask     = NativeMethods.MIIM_STRING
//                           | NativeMethods.MIIM_FTYPE
//                           | NativeMethods.MIIM_ID
//                           | NativeMethods.MIIM_SUBMENU,
//                 dwTypeData = new string('\0', 256),
//                 cch        = 256
//             };
//
//             NativeMethods.GetMenuItemInfo(hMenu, (uint)i, true, ref mii);
//
//             var item = new ContextMenuItem();
//
//             if ((mii.fType & NativeMethods.MFT_SEPARATOR) != 0)
//             {
//                 item.IsSeparator = true;
//             }
//             else
//             {
//                 item.Label     = mii.dwTypeData?.TrimEnd('\0') ?? "";
//                 item.CommandId = mii.wID - _idCmdFirst;
//                 item.Verb      = GetVerb(mii.wID - _idCmdFirst);
//
//                 if (mii.hSubMenu != IntPtr.Zero)
//                     item.Children = ReadMenuItems(mii.hSubMenu);
//             }
//
//             items.Add(item);
//         }
//
//         return items;
//     }
//
//     private string GetVerb(uint cmdIdOffset)
//     {
//         try
//         {
//             var buffer = new byte[256];
//             int hr = _contextMenu.GetCommandString(
//                 new UIntPtr(cmdIdOffset),
//                 NativeMethods.GCS_VERBW,
//                 IntPtr.Zero, buffer, 128);
//
//             if (hr == 0)
//                 return System.Text.Encoding.Unicode
//                     .GetString(buffer).TrimEnd('\0');
//         }
//         catch { }
//         return string.Empty;
//     }
//
//     public void Dispose()
//     {
//         if (_hMenu != IntPtr.Zero)
//         {
//             NativeMethods.DestroyMenu(_hMenu);
//             _hMenu = IntPtr.Zero;
//         }
//     }
// }
//
//
//
//
//
// public static class DataObjectHelper
// {
//     // Erstellt ein natives IDataObject mit einem Dateipfad
//     // (so wie Explorer es tut)
//     public static IntPtr CreateForFile(string path)
//     {
//         int hr = NativeMethods.SHParseDisplayName(
//             Path.GetDirectoryName(path),
//             IntPtr.Zero,
//             out var pidlFolder,
//             0,
//             out _);
//         if (hr != 0) Marshal.ThrowExceptionForHR(hr);
//
//         var isfGuid = typeof(IShellFolder).GUID;
//         hr = NativeMethods.SHBindToObject(
//             null, pidlFolder, IntPtr.Zero,
//             ref isfGuid, out var ppFolder);
//         if (hr != 0) Marshal.ThrowExceptionForHR(hr);
//
//         var parentFolder = (IShellFolder)
//             Marshal.GetTypedObjectForIUnknown(ppFolder, typeof(IShellFolder));
//
//         uint eaten = 0, attrs = 0;
//         parentFolder.ParseDisplayName(
//             IntPtr.Zero, IntPtr.Zero,
//             Path.GetFileName(path),
//             ref eaten, out var pidlItem, ref attrs);
//
//         var idoGuid = new Guid("0000010e-0000-0000-C000-000000000046");
//         hr = NativeMethods.SHCreateDataObject(
//             pidlFolder, 1, new[] { pidlItem },
//             IntPtr.Zero, ref idoGuid,
//             out var ppDataObject);
//         if (hr != 0) Marshal.ThrowExceptionForHR(hr);
//
//         NativeMethods.ILFree(pidlItem);
//         NativeMethods.ILFree(pidlFolder);
//         Marshal.ReleaseComObject(parentFolder);
//
//         return ppDataObject;
//     }
// }
//
// // IDataObject Interface – wird von SharpShell intern verwendet
// [ComImport]
// [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
// [Guid("0000010e-0000-0000-C000-000000000046")]
// public interface IDataObject
// {
//     // Wir brauchen nur die vollständige Vtable
//     // SharpShell ruft intern CF_HDROP ab
//     [PreserveSig] int GetData(ref FORMATETC fmt, out STGMEDIUM medium);
//     [PreserveSig] int GetDataHere(ref FORMATETC fmt, ref STGMEDIUM medium);
//     [PreserveSig] int QueryGetData(ref FORMATETC fmt);
//     [PreserveSig] int GetCanonicalFormatEtc(ref FORMATETC fmtIn, out FORMATETC fmtOut);
//     [PreserveSig] int SetData(ref FORMATETC fmt, ref STGMEDIUM medium, bool release);
//     [PreserveSig] int EnumFormatEtc(uint direction, out IntPtr enumFmt);
//     [PreserveSig] int DAdvise(ref FORMATETC fmt, uint advf, IntPtr advSink, out uint conn);
//     [PreserveSig] int DUnadvise(uint conn);
//     [PreserveSig] int EnumDAdvise(out IntPtr enumAdvise);
// }
//
// [StructLayout(LayoutKind.Sequential)]
// public struct FORMATETC
// {
//     public ushort cfFormat;
//     public IntPtr ptd;
//     public uint dwAspect;
//     public int lindex;
//     public uint tymed;
// }
//
// [StructLayout(LayoutKind.Sequential)]
// public struct STGMEDIUM
// {
//     public uint tymed;
//     public IntPtr unionmember;
//     public IntPtr pUnkForRelease;
// }
//
//
// [ComImport]
// [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
// [Guid("000214E4-0000-0000-C000-000000000046")]
// public interface IContextMenu
// {
//     [PreserveSig]
//     int QueryContextMenu(
//         IntPtr hMenu,
//         uint indexMenu,
//         uint idCmdFirst,
//         uint idCmdLast,
//         uint uFlags);
//
//     [PreserveSig]
//     int InvokeCommand(ref CMINVOKECOMMANDINFOEX pici);
//
//     [PreserveSig]
//     int GetCommandString(
//         UIntPtr idCmd,
//         uint uType,
//         IntPtr pReserved,
//         [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
//         byte[] pszName,
//         uint cchMax);
// }
//
// [ComImport]
// [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
// [Guid("000214E6-0000-0000-C000-000000000046")]
// public interface IShellFolder
// {
//     // 1. Muss EXAKT diese Reihenfolge haben (COM Vtable)
//     [PreserveSig]
//     int ParseDisplayName(
//         IntPtr hwnd,
//         IntPtr pbc,
//         [MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName,
//         ref uint pchEaten,
//         out IntPtr ppidl,
//         ref uint pdwAttributes);
//
//     [PreserveSig]
//     int EnumObjects(
//         IntPtr hwnd,
//         uint grfFlags,
//         out IntPtr ppenumIDList);
//
//     [PreserveSig]
//     int BindToObject(
//         IntPtr pidl,
//         IntPtr pbc,
//         ref Guid riid,
//         out IntPtr ppv);
//
//     [PreserveSig]
//     int BindToStorage(
//         IntPtr pidl,
//         IntPtr pbc,
//         ref Guid riid,
//         out IntPtr ppv);
//
//     [PreserveSig]
//     int CompareIDs(
//         IntPtr lParam,
//         IntPtr pidl1,
//         IntPtr pidl2);
//
//     [PreserveSig]
//     int CreateViewObject(
//         IntPtr hwndOwner,
//         ref Guid riid,
//         out IntPtr ppv);
//
//     [PreserveSig]
//     int GetAttributesOf(
//         uint cidl,
//         [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
//         IntPtr[] apidl,
//         ref uint rgfInOut);
//
//     [PreserveSig]
//     int GetUIObjectOf(
//         IntPtr hwndOwner,
//         uint cidl,
//         [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)]
//         IntPtr[] apidl,
//         ref Guid riid,
//         IntPtr rgfReserved,
//         out IntPtr ppv);
//
//     [PreserveSig]
//     int GetDisplayNameOf(
//         IntPtr pidl,
//         uint uFlags,
//         out IntPtr pName);
//
//     [PreserveSig]
//     int SetNameOf(
//         IntPtr hwnd,
//         IntPtr pidl,
//         [MarshalAs(UnmanagedType.LPWStr)] string pszName,
//         uint uFlags,
//         out IntPtr ppidlOut);
// }
//
// [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
// public struct CMINVOKECOMMANDINFOEX
// {
//     public uint cbSize;
//     public uint fMask;
//     public IntPtr hwnd;
//     public IntPtr lpVerb; // Low-Word = Command-ID offset
//     public IntPtr lpParameters;
//     public IntPtr lpDirectory;
//     public int nShow;
//     public uint dwHotKey;
//     public IntPtr hIcon;
//     public IntPtr lpTitle;
//     public IntPtr lpVerbW;
//     public IntPtr lpParametersW;
//     public IntPtr lpDirectoryW;
//     public IntPtr lpTitleW;
//     public POINT ptInvoke;
// }
//
// [StructLayout(LayoutKind.Sequential)]
// public struct POINT
// {
//     public int x, y;
// }
//
// public static class NativeMethods
// {
//     [DllImport("user32.dll")]
//     public static extern IntPtr CreatePopupMenu();
//
//     [DllImport("user32.dll")]
//     public static extern bool DestroyMenu(IntPtr hMenu);
//
//     [DllImport("user32.dll")]
//     public static extern int GetMenuItemCount(IntPtr hMenu);
//
//     [DllImport("user32.dll", CharSet = CharSet.Unicode)]
//     public static extern bool GetMenuItemInfo(
//         IntPtr hMenu, uint uItem, bool fByPosition,
//         ref MENUITEMINFO lpmii);
//
//     [DllImport("user32.dll")]
//     public static extern IntPtr GetSubMenu(IntPtr hMenu, int nPos);
//
//     [DllImport("shell32.dll")]
//     public static extern int SHGetDesktopFolder(out IShellFolder ppshf);
//     
//     [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
//     public static extern int SHParseDisplayName(
//         string pszName,
//         IntPtr pbc,
//         out IntPtr ppidl,
//         uint sfgaoIn,
//         out uint psfgaoOut);
//
//     [DllImport("shell32.dll")]
//     public static extern int SHBindToObject(
//         IShellFolder psf,
//         IntPtr pidl,
//         IntPtr pbc,
//         ref Guid riid,
//         out IntPtr ppv);
//
//     [DllImport("shell32.dll")] // ← der neue
//     public static extern int SHCreateDataObject(
//         IntPtr pidlFolder,
//         uint cidl,
//         IntPtr[] apidl,
//         IntPtr pdtInner,
//         ref Guid riid,
//         out IntPtr ppv);
//
//
//     [DllImport("shell32.dll")]
//     public static extern void ILFree(IntPtr pidl);
//
//     public const uint MIIM_STRING = 0x00000040;
//     public const uint MIIM_FTYPE = 0x00000100;
//     public const uint MIIM_ID = 0x00000002;
//     public const uint MIIM_SUBMENU = 0x00000004;
//     public const uint MFT_SEPARATOR = 0x00000800;
//
//     public const uint CMF_NORMAL = 0x00000000;
//     public const uint CMF_EXPLORE = 0x00000001;
//     public const uint CMF_EXTENDEDVERBS = 0x00000100;
//     public const uint GCS_VERBW = 0x00000004;
//
//     public const uint CMIC_MASK_UNICODE = 0x00004000;
// }
//
// [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
// public struct MENUITEMINFO
// {
//     public uint cbSize;
//     public uint fMask;
//     public uint fType;
//     public uint fState;
//     public uint wID;
//     public IntPtr hSubMenu;
//     public IntPtr hbmpChecked;
//     public IntPtr hbmpUnchecked;
//     public UIntPtr dwItemData;
//     public string dwTypeData;
//     public uint cch;
//     public IntPtr hbmpItem;
// }
//
// public class ContextMenuItem
// {
//     public string Label { get; set; }
//     public string Verb { get; set; } // z.B. "open", "delete"
//     public uint CommandId { get; set; }
//     public bool IsSeparator { get; set; }
//     public List<ContextMenuItem> Children { get; set; } = new List<ContextMenuItem>();
//
//     public override string ToString() =>
//         IsSeparator ? "---" : $"[{CommandId}] {Label} ({Verb})";
// }
//
// public class ShellContextMenuReader : IDisposable
// {
//     private IContextMenu _contextMenu;
//     private IntPtr _hMenu;
//     private uint _idCmdFirst = 1;
//     private uint _idCmdLast = 0x7FFF;
//
//     // Einstiegspunkt: Pfad zu Datei/Ordner
//     public static ShellContextMenuReader ForPath(string path)
//     {
//         var reader = new ShellContextMenuReader();
//         reader.Initialize(path);
//         return reader;
//     }
//
//     [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
//     private static extern int SHParseDisplayName(
//         string pszName, IntPtr pbc,
//         out IntPtr ppidl, uint sfgaoIn,
//         out uint psfgaoOut);
//
//     [DllImport("shell32.dll")]
//     private static extern int SHBindToObject(
//         IShellFolder psf, // null = Desktop
//         IntPtr pidl,
//         IntPtr pbc,
//         ref Guid riid,
//         out IntPtr ppv);
//
//     [DllImport("shell32.dll")]
//     private static extern int SHGetIDListFromObject(
//         [MarshalAs(UnmanagedType.IUnknown)] object punk,
//         out IntPtr ppidl);
//
//     private void Initialize(string path)
//     {
//         if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
//             throw new InvalidOperationException("STA-Thread erforderlich!");
//
//         if (!File.Exists(path) && !Directory.Exists(path))
//             throw new FileNotFoundException("Pfad nicht gefunden", path);
//
//         string directory = Path.GetDirectoryName(path);
//         string filename = Path.GetFileName(path);
//
//         // 1. Parent-Folder-PIDL holen
//         int hr = SHParseDisplayName(directory, IntPtr.Zero,
//             out var pidlFolder, 0, out _);
//         if (hr != 0) Marshal.ThrowExceptionForHR(hr);
//
//         // 2. An Parent-IShellFolder binden
//         var isfGuid = typeof(IShellFolder).GUID;
//         hr = SHBindToObject(null, pidlFolder, IntPtr.Zero,
//             ref isfGuid, out var ppFolder);
//         if (hr != 0) Marshal.ThrowExceptionForHR(hr);
//
//         var parentFolder = (IShellFolder)
//             Marshal.GetTypedObjectForIUnknown(ppFolder, typeof(IShellFolder));
//
//         // 3. Datei-PIDL relativ zum Parent parsen
//         uint eaten = 0, attrs = 0;
//         hr = parentFolder.ParseDisplayName(
//             IntPtr.Zero, IntPtr.Zero, filename,
//             ref eaten, out var pidlItem, ref attrs);
//         if (hr != 0) Marshal.ThrowExceptionForHR(hr);
//
//         try
//         {
//             // 4. IContextMenu vom Parent-Folder für die Datei holen
//             var icmGuid = typeof(IContextMenu).GUID;
//             var pidlArray = new[] { pidlItem };
//
//             hr = parentFolder.GetUIObjectOf(
//                 IntPtr.Zero, 1, pidlArray,
//                 ref icmGuid, IntPtr.Zero, out var ppv);
//             if (hr != 0) Marshal.ThrowExceptionForHR(hr);
//
//             _contextMenu = (IContextMenu)
//                 Marshal.GetObjectForIUnknown(ppv);
//             Marshal.Release(ppv);
//
//             // 5. Menü befüllen – mit EXTENDEDVERBS für alle Handler
//             _hMenu = NativeMethods.CreatePopupMenu();
//
//             _contextMenu.QueryContextMenu(
//                 _hMenu, 0, _idCmdFirst, _idCmdLast,
//                 NativeMethods.CMF_NORMAL |
//                 NativeMethods.CMF_EXTENDEDVERBS);
//         }
//         finally
//         {
//             NativeMethods.ILFree(pidlItem);
//             NativeMethods.ILFree(pidlFolder);
//             Marshal.ReleaseComObject(parentFolder);
//         }
//     }
//
//
//     // Alle Einträge rekursiv auslesen
//     public List<ContextMenuItem> ReadAllItems()
//     {
//         return ReadMenuItems(_hMenu);
//     }
//
//     private List<ContextMenuItem> ReadMenuItems(IntPtr hMenu)
//     {
//         var items = new List<ContextMenuItem>();
//         int count = NativeMethods.GetMenuItemCount(hMenu);
//
//         for (int i = 0; i < count; i++)
//         {
//             var mii = new MENUITEMINFO
//             {
//                 cbSize = (uint)Marshal.SizeOf<MENUITEMINFO>(),
//                 fMask = NativeMethods.MIIM_STRING
//                         | NativeMethods.MIIM_FTYPE
//                         | NativeMethods.MIIM_ID
//                         | NativeMethods.MIIM_SUBMENU,
//                 dwTypeData = new string('\0', 256),
//                 cch = 256
//             };
//
//             NativeMethods.GetMenuItemInfo(hMenu, (uint)i, true, ref mii);
//
//             var item = new ContextMenuItem();
//
//             if ((mii.fType & NativeMethods.MFT_SEPARATOR) != 0)
//             {
//                 item.IsSeparator = true;
//             }
//             else
//             {
//                 item.Label = mii.dwTypeData?.TrimEnd('\0') ?? "";
//                 item.CommandId = mii.wID - _idCmdFirst; // Offset normalisieren
//                 item.Verb = GetVerb(mii.wID - _idCmdFirst);
//
//                 // Rekursiv: Untermenü vorhanden?
//                 if (mii.hSubMenu != IntPtr.Zero)
//                 {
//                     item.Children = ReadMenuItems(mii.hSubMenu);
//                 }
//             }
//
//             items.Add(item);
//         }
//
//         return items;
//     }
//
//     private string GetVerb(uint cmdIdOffset)
//     {
//         try
//         {
//             var buffer = new byte[256];
//             int hr = _contextMenu.GetCommandString(
//                 new UIntPtr(cmdIdOffset),
//                 NativeMethods.GCS_VERBW,
//                 IntPtr.Zero,
//                 buffer,
//                 128); // Unicode → halbe Byte-Anzahl
//
//             if (hr == 0)
//                 return System.Text.Encoding.Unicode.GetString(buffer)
//                     .TrimEnd('\0');
//         }
//         catch
//         {
//             /* Viele Handler werfen hier */
//         }
//
//         return string.Empty;
//     }
//
//     // Eintrag per Label ausführen
//     public void Invoke(string label)
//     {
//         var item = FindItem(ReadAllItems(), label)
//                    ?? throw new InvalidOperationException(
//                        $"Menüeintrag '{label}' nicht gefunden.");
//         InvokeById(item.CommandId);
//     }
//
//     // Eintrag per Verb ausführen (stabiler als Label)
//     public void InvokeVerb(string verb)
//     {
//         var info = new CMINVOKECOMMANDINFOEX
//         {
//             cbSize = (uint)Marshal.SizeOf<CMINVOKECOMMANDINFOEX>(),
//             fMask = NativeMethods.CMIC_MASK_UNICODE,
//             lpVerb = Marshal.StringToHGlobalAnsi(verb),
//             lpVerbW = Marshal.StringToHGlobalUni(verb),
//             nShow = 1
//         };
//
//         try
//         {
//             _contextMenu.InvokeCommand(ref info);
//         }
//         finally
//         {
//             Marshal.FreeHGlobal(info.lpVerb);
//             Marshal.FreeHGlobal(info.lpVerbW);
//         }
//     }
//
//     private void InvokeById(uint cmdIdOffset)
//     {
//         var info = new CMINVOKECOMMANDINFOEX
//         {
//             cbSize = (uint)Marshal.SizeOf<CMINVOKECOMMANDINFOEX>(),
//             fMask = 0,
//             lpVerb = (IntPtr)(int)cmdIdOffset, // MAKEINTRESOURCE
//             nShow = 1
//         };
//         _contextMenu.InvokeCommand(ref info);
//     }
//
//     private ContextMenuItem FindItem(
//         List<ContextMenuItem> items, string label)
//     {
//         foreach (var item in items)
//         {
//             if (item.Label == label) return item;
//             var found = FindItem(item.Children, label);
//             if (found != null) return found;
//         }
//
//         return null;
//     }
//
//     public void Dispose()
//     {
//         if (_hMenu != IntPtr.Zero)
//         {
//             NativeMethods.DestroyMenu(_hMenu);
//             _hMenu = IntPtr.Zero;
//         }
//
//         if (_contextMenu != null)
//         {
//             Marshal.ReleaseComObject(_contextMenu);
//             _contextMenu = null;
//         }
//     }
// }