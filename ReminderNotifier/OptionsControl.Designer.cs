using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ReminderNotifier
{
    partial class OptionsControl : System.Windows.Forms.UserControl, Outlook.PropertyPage
    {
        const int captionDispID = -518;
        bool isDirty = false;
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        void Outlook.PropertyPage.Apply()
        {

        }
        bool Outlook.PropertyPage.Dirty
        {
            get
            {
                return isDirty;
            }
        }
        void Outlook.PropertyPage.GetPageInfo(ref string helpFile, ref int helpContext)
        {

        }

        [DispId(captionDispID)]
        public string PageCaption
        {
            get
            {
                return "Test page";
            }
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        }

        #endregion
    }
}
