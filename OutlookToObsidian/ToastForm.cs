using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace OutlookToObsidian
{
    internal class ToastForm : Form
    {
        private readonly Timer _closeTimer;
        private readonly string _obsidianUri;

        public ToastForm(string message, string vaultName, string fileName)
        {
            // Build Obsidian URI
            string fileWithoutExt = System.IO.Path.GetFileNameWithoutExtension(fileName);
            _obsidianUri = $"obsidian://open?vault={Uri.EscapeDataString(vaultName)}&file={Uri.EscapeDataString(fileWithoutExt)}";

            // Form settings
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            ShowInTaskbar = false;
            BackColor = Color.FromArgb(30, 30, 30);
            Size = new Size(360, 70);
            Opacity = 0.95;
            Padding = new Padding(12);

            // Position: bottom-right of primary screen
            var workArea = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(workArea.Right - Width - 16, workArea.Bottom - Height - 16);

            // Checkmark + message label
            var lblMessage = new Label
            {
                Text = "\u2713 " + message,  // ✓
                ForeColor = Color.FromArgb(200, 255, 200),
                Font = new Font("Segoe UI", 10f, FontStyle.Regular),
                AutoSize = false,
                Size = new Size(336, 22),
                Location = new Point(12, 10),
                Cursor = Cursors.Hand
            };
            lblMessage.Click += OnClick;

            // "Click to open in Obsidian" link
            var lblLink = new Label
            {
                Text = "Click to open in Obsidian",
                ForeColor = Color.FromArgb(120, 160, 255),
                Font = new Font("Segoe UI", 8.5f, FontStyle.Underline),
                AutoSize = false,
                Size = new Size(336, 18),
                Location = new Point(12, 36),
                Cursor = Cursors.Hand
            };
            lblLink.Click += OnClick;

            Controls.Add(lblMessage);
            Controls.Add(lblLink);

            // Click anywhere on form to open
            Click += OnClick;

            // Auto-close after 4 seconds
            _closeTimer = new Timer { Interval = 4000 };
            _closeTimer.Tick += (s, e) => Close();
            _closeTimer.Start();
        }

        private void OnClick(object sender, EventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = _obsidianUri,
                    UseShellExecute = true
                });
            }
            catch
            {
                // Obsidian not installed or URI scheme not registered — ignore silently
            }
            Close();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _closeTimer?.Stop();
            _closeTimer?.Dispose();
            base.OnFormClosed(e);
        }

        // Prevent stealing focus from Outlook
        protected override bool ShowWithoutActivation => true;

        private const int WS_EX_TOPMOST = 0x00000008;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int WS_EX_TOOLWINDOW = 0x00000080;

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= WS_EX_TOPMOST | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW;
                return cp;
            }
        }
    }
}
