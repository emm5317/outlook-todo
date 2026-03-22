using System;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookToObsidian
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // First-run: prompt user to select Obsidian vault folder
            if (string.IsNullOrEmpty(Properties.Settings.Default.VaultPath))
            {
                using (var dlg = new FolderBrowserDialog())
                {
                    dlg.Description = "Select your Obsidian vault folder";
                    dlg.ShowNewFolderButton = true;
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        Properties.Settings.Default.VaultPath = dlg.SelectedPath;
                        Properties.Settings.Default.Save();
                    }
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new ContextMenuRibbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
