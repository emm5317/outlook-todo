using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookToObsidian
{
    [ComVisible(true)]
    public class ContextMenuRibbon : Office.IRibbonExtensibility
    {
        public string GetCustomUI(string ribbonID)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream("OutlookToObsidian.ContextMenuRibbon.xml"))
            {
                if (stream == null)
                    throw new InvalidOperationException("Could not find embedded ContextMenuRibbon.xml resource.");

                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        public void OnCreateTask(Office.IRibbonControl control)
        {
            Outlook.Explorer explorer = null;
            Outlook.Selection selection = null;

            try
            {
                explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer == null) return;

                selection = explorer.Selection;
                if (selection == null || selection.Count == 0) return;

                // Check vault is configured
                string vaultPath = Properties.Settings.Default.VaultPath;
                if (string.IsNullOrEmpty(vaultPath) || !Directory.Exists(vaultPath))
                {
                    MessageBox.Show(
                        "Obsidian vault path is not configured or does not exist.\nPlease restart Outlook to set it up.",
                        "OutlookToObsidian",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                int created = 0;
                int skipped = 0;
                string lastSubject = "";

                for (int i = 1; i <= selection.Count; i++)
                {
                    object item = selection[i];
                    try
                    {
                        if (item is Outlook.MailItem mail)
                        {
                            // Duplicate check
                            if (TaskCreator.IsDuplicate(mail.EntryID))
                            {
                                skipped++;
                                continue;
                            }

                            string markdown = TaskCreator.CreateTask(mail);
                            TaskCreator.AppendToVault(markdown);
                            lastSubject = mail.Subject ?? "(no subject)";
                            created++;
                        }
                    }
                    finally
                    {
                        if (item != null)
                            Marshal.ReleaseComObject(item);
                    }
                }

                if (created > 0 || skipped > 0)
                {
                    string msg = "";
                    if (created == 1)
                        msg = $"Task created: {lastSubject}";
                    else if (created > 1)
                        msg = $"{created} tasks created in Obsidian";

                    if (skipped > 0)
                        msg += (msg.Length > 0 ? " | " : "") + $"{skipped} duplicate(s) skipped";

                    MessageBox.Show(msg, "OutlookToObsidian",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Could not create task:\n{ex.Message}",
                    "OutlookToObsidian",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                if (selection != null) Marshal.ReleaseComObject(selection);
                if (explorer != null) Marshal.ReleaseComObject(explorer);
            }
        }
    }
}
