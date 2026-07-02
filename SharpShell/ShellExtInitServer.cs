using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using SharpShell.Extensions;
using SharpShell.Interop;

namespace SharpShell
{
    /// <summary>
    ///     The ShellExtInitServer is the base class for Shell Extension Servers
    ///     that implement IShellExtInit.
    /// </summary>
    public abstract class ShellExtInitServer : SharpShellServer, IShellExtInit
    {
        /// <summary>
        ///     The selected item paths.
        /// </summary>
        private List<string> selectedItemPaths = new List<string>();

        /// <summary>
        ///     Gets the selected item paths.
        /// </summary>
        public IEnumerable<string> SelectedItemPaths => selectedItemPaths;

        /// <summary>
        ///     Gets the folder path, if passed by us from the explorer. This is exactly the case if there are no selected items when the context menu is built.
        /// </summary>
        /// <value>
        ///     The folder path.
        /// </value>
        public string? FolderPath { get; private set; }

        #region Implementation of IShellExtInit

        /// <summary>
        ///     Initializes the shell extension with the parent folder and data object.
        /// </summary>
        /// <param name="pidlFolder">The pidl of the parent folder.</param>
        /// <param name="pDataObj">The data object pointer.</param>
        /// <param name="hKeyProgID">The handle to the key prog id.</param>
        void IShellExtInit.Initialize(IntPtr pidlFolder, IntPtr pDataObj, IntPtr hKeyProgID)
        {
            Log("Initializing shell extension...");

            //  If we have the folder PIDL, we can get the parent folder. Which is only the case if there was no selection at the time the context menu was built.
            if (pidlFolder != IntPtr.Zero)
            {
                StringBuilder stringBuilder = new StringBuilder(260);
                if (User32.SHGetPathFromIDListW(pidlFolder, stringBuilder))
                    //  Set parent folder path.
                    FolderPath = stringBuilder.ToString();
            }

            //  If we have the data object, we can use it to get the selected item paths.
            if (pDataObj != IntPtr.Zero)
            {
                //  Create the IDataObject from the provided pDataObj.
                IDataObject dataObject = (IDataObject)Marshal.GetObjectForIUnknown(pDataObj);

                //  Add the set of files to the selected file paths.
                selectedItemPaths = dataObject.GetFileList();
            }

            Log(string.Format("Shell extension initialised.{0}Parent folder: {1}{0}Items: {0}{2}",
                Environment.NewLine, FolderPath ?? "<none>", string.Join(Environment.NewLine, selectedItemPaths)));
        }

        #endregion
    }
}