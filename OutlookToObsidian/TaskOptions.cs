namespace OutlookToObsidian
{
    internal class TaskOptions
    {
        public string Subject { get; set; }
        public string DueDate { get; set; }       // "yyyy-MM-dd" or "" for none
        public string Priority { get; set; }      // emoji string (⏫/🔼/🔽) or "" for normal
        public string Tags { get; set; }          // space-separated #tags
        public string Notes { get; set; }         // optional user note
        public int AttachmentCount { get; set; }  // from MailItem.Attachments.Count
    }
}
