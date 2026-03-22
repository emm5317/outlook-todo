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

        /// <summary>
        /// Instant create — same as v1 but with toast instead of MessageBox.
        /// </summary>
        public void OnCreateTaskInstant(Office.IRibbonControl control)
        {
            ProcessEmails(detailed: false);
        }

        /// <summary>
        /// Detailed create — shows TaskEntryForm for each email before creating.
        /// </summary>
        public void OnCreateTaskDetailed(Office.IRibbonControl control)
        {
            ProcessEmails(detailed: true);
        }

        private void ProcessEmails(bool detailed)
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
                string lastFileName = "";

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

                            string markdown;

                            if (detailed)
                            {
                                // Show the detailed entry form
                                var defaults = TaskCreator.BuildDefaultOptions(mail);
                                using (var form = new TaskEntryForm(defaults))
                                {
                                    if (form.ShowDialog() != DialogResult.OK)
                                        continue; // User cancelled

                                    // Preserve attachment count from mail
                                    form.Result.AttachmentCount = defaults.AttachmentCount;
                                    markdown = TaskCreator.CreateTask(mail, form.Result);
                                }
                            }
                            else
                            {
                                markdown = TaskCreator.CreateTask(mail);
                            }

                            lastFileName = TaskCreator.AppendToVault(markdown);
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

                // Show toast notification
                if (created > 0 || skipped > 0)
                {
                    string msg;
                    if (created == 1)
                        msg = $"Task created: {Truncate(lastSubject, 40)}";
                    else if (created > 1)
                        msg = $"{created} tasks created in Obsidian";
                    else
                        msg = "";

                    if (skipped > 0)
                    {
                        string skipMsg = $"{skipped} duplicate(s) skipped";
                        msg = string.IsNullOrEmpty(msg) ? skipMsg : msg + " | " + skipMsg;
                    }

                    string vaultName = TaskCreator.GetVaultName();
                    var toast = new ToastForm(msg, vaultName, lastFileName);
                    toast.Show();
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

        private static string Truncate(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLength)
                return text;
            return text.Substring(0, maxLength) + "...";
        }
    }
}
