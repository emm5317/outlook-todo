using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookToObsidian
{
    internal static class TaskCreator
    {
        /// <summary>
        /// Builds an Obsidian Tasks-formatted markdown block from a MailItem.
        /// </summary>
        public static string CreateTask(Outlook.MailItem mail)
        {
            string subject = SanitizeForMarkdown(mail.Subject ?? "(no subject)");
            string senderName = mail.SenderName ?? "Unknown";
            string senderEmail = mail.SenderEmailAddress ?? "";
            string received = mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm");
            string entryId = mail.EntryID;
            string today = DateTime.Now.ToString("yyyy-MM-dd");

            string priority = MapPriority(mail.Importance);
            string tags = MapCategories(mail.Categories);
            string bodyPreview = GetBodyPreview(mail.Body, 200);

            var sb = new StringBuilder();
            sb.AppendLine($"- [ ] {subject} {priority}#follow-up {tags}\u2795 {today}".TrimEnd());
            sb.AppendLine($"  > From: {senderName} ({senderEmail}) | {received}");
            if (!string.IsNullOrEmpty(bodyPreview))
            {
                sb.AppendLine($"  > {bodyPreview}");
            }
            sb.AppendLine($"  > [entry-id:: {entryId}]");
            sb.AppendLine();

            return sb.ToString();
        }

        /// <summary>
        /// Appends markdown to the configured vault file.
        /// Returns the resolved file path on success.
        /// </summary>
        public static string AppendToVault(string markdown)
        {
            var settings = Properties.Settings.Default;
            string vaultPath = settings.VaultPath;

            if (string.IsNullOrEmpty(vaultPath) || !Directory.Exists(vaultPath))
            {
                throw new InvalidOperationException(
                    "Obsidian vault path is not configured or does not exist. Please restart Outlook to set it up.");
            }

            string fileName;
            if (settings.UseDailyNotes)
            {
                string format = string.IsNullOrEmpty(settings.DailyNotesFormat)
                    ? "yyyy-MM-dd"
                    : settings.DailyNotesFormat;
                fileName = DateTime.Now.ToString(format) + ".md";
            }
            else
            {
                fileName = string.IsNullOrEmpty(settings.TaskFileName)
                    ? "Inbox.md"
                    : settings.TaskFileName;
            }

            string targetPath = Path.Combine(vaultPath, fileName);

            // Create file with header if it doesn't exist
            if (!File.Exists(targetPath))
            {
                string header = settings.UseDailyNotes
                    ? $"# {DateTime.Now:yyyy-MM-dd}\n\n"
                    : $"# {Path.GetFileNameWithoutExtension(fileName)}\n\n";
                File.WriteAllText(targetPath, header, Encoding.UTF8);
            }

            File.AppendAllText(targetPath, markdown, Encoding.UTF8);
            return targetPath;
        }

        /// <summary>
        /// Checks whether a task with this EntryID already exists in the target file.
        /// </summary>
        public static bool IsDuplicate(string entryId)
        {
            var settings = Properties.Settings.Default;
            if (string.IsNullOrEmpty(settings.VaultPath))
                return false;

            string fileName;
            if (settings.UseDailyNotes)
            {
                string format = string.IsNullOrEmpty(settings.DailyNotesFormat)
                    ? "yyyy-MM-dd"
                    : settings.DailyNotesFormat;
                fileName = DateTime.Now.ToString(format) + ".md";
            }
            else
            {
                fileName = string.IsNullOrEmpty(settings.TaskFileName)
                    ? "Inbox.md"
                    : settings.TaskFileName;
            }

            string targetPath = Path.Combine(settings.VaultPath, fileName);

            if (!File.Exists(targetPath))
                return false;

            string content = File.ReadAllText(targetPath, Encoding.UTF8);
            return content.Contains($"[entry-id:: {entryId}]");
        }

        /// <summary>
        /// Maps Outlook importance to Obsidian Tasks priority emoji.
        /// </summary>
        private static string MapPriority(Outlook.OlImportance importance)
        {
            switch (importance)
            {
                case Outlook.OlImportance.olImportanceHigh:
                    return "\u23EB "; // ⏫
                case Outlook.OlImportance.olImportanceLow:
                    return "\uD83D\uDD3D "; // 🔽
                default:
                    return "";
            }
        }

        /// <summary>
        /// Converts Outlook categories (comma-separated) to Obsidian #tags.
        /// </summary>
        private static string MapCategories(string categories)
        {
            if (string.IsNullOrEmpty(categories))
                return "";

            var tags = categories
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(c => "#" + c.Trim().ToLowerInvariant().Replace(" ", "-"))
                .Where(t => t.Length > 1);

            string result = string.Join(" ", tags);
            return string.IsNullOrEmpty(result) ? "" : result + " ";
        }

        /// <summary>
        /// Extracts a plain-text body preview, collapsing whitespace.
        /// </summary>
        private static string GetBodyPreview(string body, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(body))
                return "";

            // Collapse all whitespace (newlines, tabs, multiple spaces) into single spaces
            string cleaned = Regex.Replace(body, @"\s+", " ").Trim();

            if (cleaned.Length <= maxLength)
                return cleaned;

            return cleaned.Substring(0, maxLength) + "...";
        }

        /// <summary>
        /// Removes characters that would break markdown task formatting.
        /// </summary>
        private static string SanitizeForMarkdown(string text)
        {
            return text
                .Replace("\r\n", " ")
                .Replace("\n", " ")
                .Replace("\r", " ")
                .Replace("[", "(")
                .Replace("]", ")");
        }
    }
}
