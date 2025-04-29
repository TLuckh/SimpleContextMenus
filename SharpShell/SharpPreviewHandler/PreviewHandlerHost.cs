using System.Drawing;
using System.Windows.Forms;

namespace SharpShell.SharpPreviewHandler
{
    /// <summary>
    ///     The PreviewHandlerHost is the window created in the preview
    ///     pane which will hold the derived preview handlers UI.
    /// </summary>
    internal class PreviewHandlerHost : Form
    {
        /// <summary>
        ///     Initializes a new instance of the <see cref="PreviewHandlerHost" /> class.
        /// </summary>
        public PreviewHandlerHost()
        {
            //  Initialize the component.
            InitializeComponent();

            FormBorderStyle = FormBorderStyle.None;

            //  Invisible by default.
            Visible = false;
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x02000000; // Turn on WS_EX_COMPOSITED
                return cp;
            }
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            // 
            // PreviewHandlerHost
            // 
            Name = "PreviewHandlerHost";
            ResumeLayout(false);

            BackColor = Color.White;
        }
    }
}