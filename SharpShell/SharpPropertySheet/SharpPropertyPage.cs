using System;
using System.Drawing;
using System.Windows.Forms;
using SharpShell.Diagnostics;

namespace SharpShell.SharpPropertySheet
{
    /// <summary>
    ///     The SharpPropertyPage class is the base class for Property Pages used in
    ///     Shell Property Sheet Extension servers.
    /// </summary>
    public class SharpPropertyPage : UserControl
    {
        /// <summary>
        ///     Gets or sets the page title.
        /// </summary>
        /// <value>
        ///     The page title.
        /// </value>
        public string PageTitle { get; set; }

        /// <summary>
        ///     Gets the page icon.
        /// </summary>
        public Icon PageIcon { get; set; }

        /// <summary>
        ///     Gets or sets the proxy.
        /// </summary>
        /// <value>
        ///     The proxy.
        /// </value>
        internal PropertyPageProxy PropertyPageProxy { get; set; }

        /// <summary>
        ///     Called when the page is initialised.
        /// </summary>
        /// <param name="parent">The parent property sheet.</param>
        protected internal virtual void OnPropertyPageInitialised(SharpPropertySheet parent)
        {
            Log("Page initialised");
        }

        /// <summary>
        ///     Called when the property page is activated.
        /// </summary>
        protected internal virtual void OnPropertyPageSetActive()
        {
            Log("Page activated");
        }

        /// <summary>
        ///     Called when the property page is de-activated or about to be closed.
        /// </summary>
        protected internal virtual void OnPropertyPageKillActive()
        {
            Log("Page deactivated");
        }

        /// <summary>
        ///     Called when apply is pressed on the property sheet, or the property
        ///     sheet is dismissed with the OK button.
        /// </summary>
        protected internal virtual void OnPropertySheetApply()
        {
            Log("Page apply");
        }

        /// <summary>
        ///     Called when OK is pressed on the property sheet.
        /// </summary>
        protected internal virtual void OnPropertySheetOK()
        {
            Log("Page OK");
        }

        /// <summary>
        ///     Called when Cancel is pressed on the property sheet.
        /// </summary>
        protected internal virtual void OnPropertySheetCancel()
        {
            Log("Page cancel");
        }

        /// <summary>
        ///     Called when the cross button on the property sheet is pressed.
        /// </summary>
        protected internal virtual void OnPropertySheetClose()
        {
            Log("Page close");
        }

        /// <summary>
        ///     Called when the property page is being released. Use this to clean up managed and unmanaged resources.
        ///     Do *not* use it to close window handles - the window will be closed automatically as the sheet closes.
        /// </summary>
        protected internal virtual void OnRelease()
        {
            Log("Page Release");
        }

        /// <summary>
        ///     Sets the page data changed.
        /// </summary>
        /// <param name="changed">
        ///     if set to <c>true</c> the data has changed and the property sheet should enable the 'Apply'
        ///     button.
        /// </param>
        protected void SetPageDataChanged(bool changed)
        {
            //  Pass on to the proxy.
            PropertyPageProxy.SetDataChanged(changed);
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            // 
            // SharpPropertyPage
            // 
            BackColor = Color.Transparent;
            Name = "SharpPropertyPage";
            ResumeLayout(false);
        }

        #region Logging Helper Functions

        /// <summary>
        ///     Logs the specified message. Will include the Shell Extension name and page name if available.
        /// </summary>
        /// <param name="message">The message.</param>
        protected void Log(string message)
        {
            SharpPropertySheet parent = PropertyPageProxy?.Parent;
            string level1 = parent != null ? parent.DisplayName : "Unknown";
            Logging.Log($"{level1} ('{PageTitle}' Page): {message}");
        }

        /// <summary>
        ///     Logs the specified message as an error.  Will include the Shell Extension name and page name if available.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="exception">Optional exception details.</param>
        protected void LogError(string message, Exception exception = null)
        {
            SharpPropertySheet parent = PropertyPageProxy?.Parent;
            string level1 = parent != null ? parent.DisplayName : "Unknown";
            Logging.Error($"{level1} ('{PageTitle}' Page): {message}", exception);
        }

        #endregion
    }
}