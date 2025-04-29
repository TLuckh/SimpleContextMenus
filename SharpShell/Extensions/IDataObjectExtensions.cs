using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using SharpShell.Interop;

namespace SharpShell.Extensions
{
    /// <summary>
    ///     Extensions for the IDataObject interface.
    /// </summary>
    public static class IDataObjectExtensions
    {
        private const int MaxPath = 260;

        /// <summary>
        ///     Gets the file list.
        /// </summary>
        /// <param name="this">The IDataObject instance.</param>
        /// <returns>The file list in the data object.</returns>
        public static List<string> GetFileList(this IDataObject @this)
        {
            //  Create the format object.
            FORMATETC formatEtc = new FORMATETC();

            //  Set up the format object, based on CF_HDROP.
            formatEtc.cfFormat = (short)CLIPFORMAT.CF_HDROP;
            formatEtc.ptd = IntPtr.Zero;
            formatEtc.dwAspect = DVASPECT.DVASPECT_CONTENT;
            formatEtc.lindex = -1;
            formatEtc.tymed = TYMED.TYMED_HGLOBAL;

            //  Get the data.
            // ReSharper disable RedundantAssignment
            STGMEDIUM storageMedium = new STGMEDIUM();
            // ReSharper restore RedundantAssignment
            @this.GetData(ref formatEtc, out storageMedium);

            List<string> fileList = new List<string>();

            //  Things can get risky now.
            try
            {
                //  Get the handle to the HDROP.
                IntPtr hDrop = storageMedium.unionmember;
                if (hDrop == IntPtr.Zero)
                    throw new ArgumentException("Failed to get the handle to the drop data.");

                //  Get the count of the files in the operation.
                uint fileCount = Shell32.DragQueryFile(hDrop, uint.MaxValue, null, 0);

                //  Go through each file.
                for (uint i = 0; i < fileCount; i++)
                {
                    //  Storage for the file name.
                    StringBuilder fileNameBuilder = new StringBuilder(MaxPath);

                    //  Get the file name.
                    uint result = Shell32.DragQueryFile(hDrop, i, fileNameBuilder, (uint)fileNameBuilder.Capacity);

                    //  If we have a valid result, add the file name to the list of files.
                    if (result != 0)
                        fileList.Add(fileNameBuilder.ToString());
                }
            }
            finally
            {
                Ole32.ReleaseStgMedium(ref storageMedium);
            }

            return fileList;
        }
    }
}