using System;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using SharpShell.Attributes;
using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;

namespace SharpShell.SharpDataHandler
{
    /// <summary>
    ///     The SharpDataHandler is the base class for SharpShell servers that provide
    ///     custom Icon Handlers.
    /// </summary>
    [ServerType(ServerType.ShellDataHandler)]
    public abstract class SharpDataHandler : PersistFileServer, IDataObject
    {
        /// <summary>
        ///     Gets the data for the selected item. The selected item's path is stored in the SelectedItemPath property.
        /// </summary>
        /// <returns>The data for the selected item, or null if there is none.</returns>
        protected abstract DataObject GetData();

        #region Implementation of IDataObject

        int IDataObject.DAdvise(ref FORMATETC pFormatetc, ADVF advf, IAdviseSink adviseSink, out int connection)
        {
            //  DebugLog key events.
            Log("IDataObject.DAdvise called.");

            //  Not needed for Shell Data Handlers.
            throw new NotImplementedException();
        }

        void IDataObject.DUnadvise(int connection)
        {
            //  DebugLog key events.
            Log("IDataObject.DUnadvise called.");

            //  Not needed for Shell Data Handlers.
            throw new NotImplementedException();
        }

        int IDataObject.EnumDAdvise(out IEnumSTATDATA enumAdvise)
        {
            //  DebugLog key events.
            Log("IDataObject.EnumDAdvise called.");

            //  Not needed for Shell Data Handlers.
            throw new NotImplementedException();
        }

        IEnumFORMATETC IDataObject.EnumFormatEtc(DATADIR direction)
        {
            //  DebugLog key events.
            Log("IDataObject.EnumFormatEtc called.");

            //  Not needed for Shell Data Handlers.
            throw new NotImplementedException();
        }

        int IDataObject.GetCanonicalFormatEtc(ref FORMATETC formatIn, out FORMATETC formatOut)
        {
            //  DebugLog key events.
            Log("IDataObject.GetCanonicalFormatEtc called.");

            //  Not needed for Shell Data Handlers.
            throw new NotImplementedException();
        }

        void IDataObject.GetData(ref FORMATETC format, out STGMEDIUM medium)
        {
            //  DebugLog key events.
            Log("IDataObject.GetData called.");

            //  Not needed for Shell Data Handlers.
            throw new NotImplementedException();
        }

        void IDataObject.GetDataHere(ref FORMATETC format, ref STGMEDIUM medium)
        {
            //  DebugLog key events.
            Log("IDataObject.GetDataHere called.");

            //  Not needed for Shell Data Handlers.
            throw new NotImplementedException();
        }

        int IDataObject.QueryGetData(ref FORMATETC format)
        {
            //  DebugLog key events.
            Log("IDataObject.QueryGetData called.");

            //  Not needed for Shell Data Handlers.
            throw new NotImplementedException();
        }

        void IDataObject.SetData(ref FORMATETC formatIn, ref STGMEDIUM medium, bool release)
        {
            //  DebugLog key events.
            Log("IDataObject.SetData called.");

            try
            {
                //  Let the derived class provide data.
                DataObject itemData = GetData();

                //  Set the data.
                ((IDataObject)itemData).SetData(ref formatIn, ref medium, release);
            }
            catch (Exception exception)
            {
                LogError("An exception occured getting data for the item " + SelectedItemPath, exception);
            }
        }

        #endregion
    }
}