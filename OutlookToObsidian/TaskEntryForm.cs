using System;
using System.Drawing;
using System.Windows.Forms;

namespace OutlookToObsidian
{
    internal class TaskEntryForm : Form
    {
        private TextBox txtSubject;
        private DateTimePicker dtpDueDate;
        private CheckBox chkNoDueDate;
        private ComboBox cmbPriority;
        private TextBox txtTags;
        private TextBox txtNotes;
        private Button btnCreate;
        private Button btnCancel;

        public TaskOptions Result { get; private set; }

        public TaskEntryForm(TaskOptions defaults)
        {
            InitializeComponents();
            PopulateDefaults(defaults);
        }

        private void InitializeComponents()
        {
            Text = "Create Task in Obsidian";
            Size = new Size(460, 340);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            Font = new Font("Segoe UI", 9f);
            BackColor = Color.White;

            int labelX = 16;
            int inputX = 110;
            int inputWidth = 320;
            int y = 16;
            int rowHeight = 32;

            // Subject
            Controls.Add(new Label { Text = "Subject:", Location = new Point(labelX, y + 3), AutoSize = true });
            txtSubject = new TextBox { Location = new Point(inputX, y), Size = new Size(inputWidth, 23) };
            Controls.Add(txtSubject);
            y += rowHeight + 4;

            // Due Date
            Controls.Add(new Label { Text = "Due Date:", Location = new Point(labelX, y + 3), AutoSize = true });
            dtpDueDate = new DateTimePicker
            {
                Location = new Point(inputX, y),
                Size = new Size(160, 23),
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Now.AddDays(7)
            };
            Controls.Add(dtpDueDate);

            chkNoDueDate = new CheckBox
            {
                Text = "No due date",
                Location = new Point(inputX + 170, y + 2),
                AutoSize = true,
                Checked = true
            };
            chkNoDueDate.CheckedChanged += (s, e) => dtpDueDate.Enabled = !chkNoDueDate.Checked;
            dtpDueDate.Enabled = false;
            Controls.Add(chkNoDueDate);
            y += rowHeight + 4;

            // Priority
            Controls.Add(new Label { Text = "Priority:", Location = new Point(labelX, y + 3), AutoSize = true });
            cmbPriority = new ComboBox
            {
                Location = new Point(inputX, y),
                Size = new Size(160, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbPriority.Items.AddRange(new object[] { "Normal", "High \u23EB", "Medium \uD83D\uDD3C", "Low \uD83D\uDD3D" });
            cmbPriority.SelectedIndex = 0;
            Controls.Add(cmbPriority);
            y += rowHeight + 4;

            // Tags
            Controls.Add(new Label { Text = "Tags:", Location = new Point(labelX, y + 3), AutoSize = true });
            txtTags = new TextBox { Location = new Point(inputX, y), Size = new Size(inputWidth, 23) };
            Controls.Add(txtTags);
            y += rowHeight + 4;

            // Notes
            Controls.Add(new Label { Text = "Notes:", Location = new Point(labelX, y + 3), AutoSize = true });
            txtNotes = new TextBox
            {
                Location = new Point(inputX, y),
                Size = new Size(inputWidth, 50),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
            };
            Controls.Add(txtNotes);
            y += 58;

            // Buttons
            btnCreate = new Button
            {
                Text = "Create Task",
                Size = new Size(100, 30),
                Location = new Point(230, y),
                DialogResult = DialogResult.OK,
                BackColor = Color.FromArgb(55, 120, 200),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnCreate.FlatAppearance.BorderSize = 0;
            btnCreate.Click += BtnCreate_Click;
            Controls.Add(btnCreate);

            btnCancel = new Button
            {
                Text = "Cancel",
                Size = new Size(80, 30),
                Location = new Point(340, y),
                DialogResult = DialogResult.Cancel
            };
            Controls.Add(btnCancel);

            AcceptButton = btnCreate;
            CancelButton = btnCancel;
        }

        private void PopulateDefaults(TaskOptions defaults)
        {
            if (defaults == null) return;

            txtSubject.Text = defaults.Subject ?? "";
            txtTags.Text = defaults.Tags ?? "#follow-up";
            txtNotes.Text = defaults.Notes ?? "";

            // Priority
            if (!string.IsNullOrEmpty(defaults.Priority))
            {
                if (defaults.Priority.Contains("\u23EB")) // ⏫
                    cmbPriority.SelectedIndex = 1;
                else if (defaults.Priority.Contains("\uD83D\uDD3C")) // 🔼
                    cmbPriority.SelectedIndex = 2;
                else if (defaults.Priority.Contains("\uD83D\uDD3D")) // 🔽
                    cmbPriority.SelectedIndex = 3;
            }
        }

        private void BtnCreate_Click(object sender, EventArgs e)
        {
            string priorityEmoji = "";
            switch (cmbPriority.SelectedIndex)
            {
                case 1: priorityEmoji = "\u23EB "; break;   // ⏫
                case 2: priorityEmoji = "\uD83D\uDD3C "; break; // 🔼
                case 3: priorityEmoji = "\uD83D\uDD3D "; break; // 🔽
            }

            Result = new TaskOptions
            {
                Subject = TaskCreator.SanitizeForMarkdown(txtSubject.Text.Trim()),
                DueDate = chkNoDueDate.Checked ? "" : dtpDueDate.Value.ToString("yyyy-MM-dd"),
                Priority = priorityEmoji,
                Tags = txtTags.Text.Trim(),
                Notes = txtNotes.Text.Trim()
            };

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
