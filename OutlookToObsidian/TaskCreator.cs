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
        /// Uses TaskOptions overrides if provided (from the detailed dialog).
        /// </summary>
        public static string CreateTask(Outlook.MailItem mail, TaskOptions options = null)
        {
            string subject = options?.Subject ?? SanitizeForMarkdown(mail.Subject ?? "(no subject)");
            string senderName = mail.SenderName ?? "Unknown";
            string senderEmail = mail.SenderEmailAddress ?? "";
            string received = mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm");
            string entryId = mail.EntryID;
            string today = DateTime.Now.ToString("yyyy-MM-dd");

            // Priority: use options override or auto-detect from Outlook importance
            string priority = options?.Priority ?? MapPriority(mail.Importance);

            // Tags: use options override or auto-detect from categories + default
            string tags = options?.Tags != null
                ? (options.Tags.Trim() + " ")
                : ("#follow-up " + MapCategories(mail.Categories));

            // Due date: from options or empty
            string dueDate = !string.IsNullOrEmpty(options?.DueDate)
                ? "\uD83D\uDCC5 " + options.DueDate + " "  // 📅
                : "";

            // Attachment count: from options or detect from mail
            int attachments = options?.AttachmentCount ?? GetAttachmentCount(mail);

            string bodyPreview = GetBodyPreview(mail.Body, 140);

            var sb = new StringBuilder();

            // Task line: - [ ] subject priority tags due-date created-date
            string taskLine = $"- [ ] {subject} {priority}{tags}{dueDate}\u2795 {today}";
            sb.AppendLine(taskLine.TrimEnd());

            // Sender line: bold name only (no email), date, optional attachment count, dedup hash
            string idHash = GetShortHash(entryId);
            string senderLine = $"  > **{senderName}** | {received}";
            if (attachments > 0)
                senderLine += $" | \uD83D\uDCCE {attachments}";  // 📎
            senderLine += $" | ^{idHash}";
            sb.AppendLine(senderLine);

            // Body preview (cleaned of URLs, invisible chars, junk)
            if (!string.IsNullOrEmpty(bodyPreview))
            {
                sb.AppendLine($"  > {bodyPreview}");
            }

            // User notes (from detailed dialog)
            if (!string.IsNullOrEmpty(options?.Notes))
            {
                sb.AppendLine($"  > **Note:** {SanitizeForMarkdown(options.Notes)}");
            }

            sb.AppendLine();

            return sb.ToString();
        }

        /// <summary>
        /// Builds a TaskOptions pre-filled from a MailItem (for the detailed dialog).
        /// </summary>
        public static TaskOptions BuildDefaultOptions(Outlook.MailItem mail)
        {
            return new TaskOptions
            {
                Subject = SanitizeForMarkdown(mail.Subject ?? "(no subject)"),
                DueDate = "",
                Priority = MapPriority(mail.Importance),
                Tags = "#follow-up " + MapCategories(mail.Categories).TrimEnd(),
                Notes = "",
                AttachmentCount = GetAttachmentCount(mail)
            };
        }

        /// <summary>
        /// Appends markdown to the configured vault file.
        /// Returns the resolved file name (not full path) on success.
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

            string fileName = GetTargetFileName();
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
            return fileName;
        }

        /// <summary>
        /// Checks whether a task with this EntryID already exists in the target file.
        /// </summary>
        public static bool IsDuplicate(string entryId)
        {
            var settings = Properties.Settings.Default;
            if (string.IsNullOrEmpty(settings.VaultPath))
                return false;

            string fileName = GetTargetFileName();
            string targetPath = Path.Combine(settings.VaultPath, fileName);

            if (!File.Exists(targetPath))
                return false;

            string content = File.ReadAllText(targetPath, Encoding.UTF8);
            string hash = GetShortHash(entryId);
            // Check current format (^hash), HTML comment format, and old Dataview format
            return content.Contains($"^{hash}")
                || content.Contains($"<!-- entry-id: {entryId} -->")
                || content.Contains($"[entry-id:: {entryId}]");
        }

        /// <summary>
        /// Gets the configured vault name (folder name of vault path).
        /// </summary>
        public static string GetVaultName()
        {
            var settings = Properties.Settings.Default;
            if (!string.IsNullOrEmpty(settings.VaultName))
                return settings.VaultName;

            if (!string.IsNullOrEmpty(settings.VaultPath))
                return new DirectoryInfo(settings.VaultPath).Name;

            return "";
        }

        private static string GetTargetFileName()
        {
            var settings = Properties.Settings.Default;
            if (settings.UseDailyNotes)
            {
                string format = string.IsNullOrEmpty(settings.DailyNotesFormat)
                    ? "yyyy-MM-dd"
                    : settings.DailyNotesFormat;
                return DateTime.Now.ToString(format) + ".md";
            }
            else
            {
                return string.IsNullOrEmpty(settings.TaskFileName)
                    ? "Inbox.md"
                    : settings.TaskFileName;
            }
        }

        /// <summary>
        /// Maps Outlook importance to Obsidian Tasks priority emoji.
        /// </summary>
        internal static string MapPriority(Outlook.OlImportance importance)
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
        internal static string MapCategories(string categories)
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
        /// Creates an 8-char hash from the EntryID for compact duplicate detection.
        /// </summary>
        private static string GetShortHash(string entryId)
        {
            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(entryId));
                // Take first 4 bytes → 8 hex chars
                return BitConverter.ToString(bytes, 0, 4).Replace("-", "").ToLowerInvariant();
            }
        }

        private static int GetAttachmentCount(Outlook.MailItem mail)
        {
            try
            {
                return mail.Attachments?.Count ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        private static string GetBodyPreview(string body, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(body))
                return "";

            string cleaned = body;

            // Strip URLs
            cleaned = Regex.Replace(cleaned, @"https?://\S+", "");

            // Strip invisible Unicode characters (zero-width spaces, combining marks, etc.)
            cleaned = Regex.Replace(cleaned, @"[\u034F\u200B-\u200F\u2028-\u202F\uFEFF]", "");

            // Strip common email junk phrases
            string[] junkPhrases = {
                "View Web Version", "View in browser", "View online",
                "Unsubscribe", "Click here", "Learn more",
                "Having trouble viewing", "Add us to your address book"
            };
            foreach (var phrase in junkPhrases)
            {
                cleaned = Regex.Replace(cleaned, Regex.Escape(phrase), "", RegexOptions.IgnoreCase);
            }

            // Strip leading/trailing quotes
            cleaned = cleaned.Trim('"', '\u201C', '\u201D', '\'');

            // Collapse whitespace
            cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim();

            // Strip markdown link artifacts like <URL> patterns
            cleaned = Regex.Replace(cleaned, @"<[^>]*>", "");
            cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim();

            if (string.IsNullOrEmpty(cleaned))
                return "";

            if (cleaned.Length <= maxLength)
                return cleaned;

            return cleaned.Substring(0, maxLength) + "...";
        }

        internal static string SanitizeForMarkdown(string text)
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
