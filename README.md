# OutlookToObsidian

A VSTO add-in for Classic Outlook on Windows that turns emails into tasks in your [Obsidian](https://obsidian.md) vault. Right-click any email, click "Create Task in Obsidian," and a formatted Markdown task appears in your vault — ready for the [Obsidian Tasks](https://publish.obsidian.md/tasks/) plugin.

No servers. No cloud sync. No accounts. Just a right-click menu and a local file write.

## What It Does

- **Two right-click options**: instant create (one click) or detailed dialog (set due date, priority, tags, notes)
- Appends a Markdown task to a file in your Obsidian vault using [Obsidian Tasks emoji format](https://publish.obsidian.md/tasks/Reference/Task+Formats/Tasks+Emoji+Format/):
  ```
  - [ ] Re: Q3 Budget Review ⏫ #follow-up #project-alpha 📅 2026-03-28 ➕ 2026-03-22
    > **Jane Smith** | 2026-03-20 14:30 | 📎 2 | ^a3f1b2c4
    > First 140 characters of the email body preview...
  ```
- Maps **Outlook importance** (High/Low) to Obsidian Tasks **priority** (⏫/🔽)
- Converts **Outlook categories** to **Obsidian #tags**
- Shows **attachment count** (📎) when email has attachments
- Detects **duplicates** — won't create the same task twice
- **Toast notification** after creation — auto-closes, click to open in Obsidian
- Prompts for vault folder on first run — no config files to edit
- Supports **daily notes** or a single target file (configurable)
- **Body preview** strips URLs, invisible characters, and email junk phrases

## Requirements

- Windows 10/11
- Classic Outlook (the COM/VSTO desktop version, not "New Outlook")
- [Visual Studio 2022](https://visualstudio.microsoft.com/) with Office/SharePoint development workload (for building)
- [VSTO Runtime 4.0](https://learn.microsoft.com/en-us/visualstudio/vsto/visual-studio-tools-for-office-runtime-installation-scenarios) (for running)
- .NET Framework 4.8
- [Obsidian](https://obsidian.md) with a local vault

## Recommended Obsidian Plugins

**Core (required):**
- [Tasks](https://publish.obsidian.md/tasks/) — parses due dates, priorities, recurrence, and enables task queries

**Recommended:**
- [Kanban](https://github.com/mgmeyers/obsidian-kanban) — drag tasks between columns (Inbox / In Progress / Waiting / Done)
- [Reminder](https://github.com/uphy/obsidian-reminder) — desktop notifications for tasks with due dates
- [Calendar](https://github.com/liamcain/obsidian-calendar-plugin) — visual monthly view of tasks and daily notes
- [Periodic Notes](https://github.com/liamcain/obsidian-periodic-notes) — weekly review templates with auto-queried task summaries
- [Checklist](https://github.com/delashum/obsidian-checklist-plugin) — sidebar panel aggregating all open tasks
- [Quick Add](https://github.com/chhoumann/quickadd) — hotkey to add manual tasks from within Obsidian
- [Style Settings](https://github.com/mgmeyers/obsidian-style-settings) — UI to tweak CSS priority colors and styling

## Build

1. Clone the repo
2. Open `OutlookToObsidian/OutlookToObsidian.slnx` in Visual Studio
3. Go to project Properties → Signing → check "Sign the ClickOnce manifests" → Create Test Certificate
4. Build (Ctrl+Shift+B)
5. F5 to debug — Outlook opens with the add-in loaded

## Install (without building)

1. Build in Release mode, then Publish → Folder
2. On the target machine, double-click the `.vsto` file
3. Open Outlook → right-click any email → "Create Task in Obsidian"
4. First run prompts you to select your Obsidian vault folder

## Vault Setup

Copy the template files from `OutlookToObsidian/ObsidianVaultTemplates/` into your vault:

- **Inbox.md** — landing file where tasks are appended
- **Task Dashboard.md** — queries for overdue, due this week, high priority, needs triage, and completed tasks

A CSS snippet for color-coded priority checkboxes and visual styling is included — copy it to your vault's `.obsidian/snippets/` folder and enable it in Settings → Appearance → CSS snippets.

## Project Structure

```
OutlookToObsidian/
├── ContextMenuRibbon.xml      # Ribbon XML — context menu buttons (instant + detailed)
├── ContextMenuRibbon.cs       # Ribbon callbacks — instant create, detailed dialog, toast
├── TaskCreator.cs             # Core logic — markdown formatting, body cleaning, dedup
├── TaskOptions.cs             # Data model for detailed dialog results
├── TaskEntryForm.cs           # WinForms dialog — due date, priority, tags, notes
├── ToastForm.cs               # Auto-closing toast notification with Obsidian link
├── ThisAddIn.cs               # Add-in entry point — wires ribbon, first-run config
├── Properties/
│   └── Settings.settings      # User settings — vault path, file name, daily notes
└── ObsidianVaultTemplates/
    ├── Inbox.md               # Template for task landing file
    └── Task Dashboard.md      # Template with Obsidian Tasks queries
```

## Settings

Stored automatically in `%AppData%` (no manual editing needed):

| Setting | Default | Description |
|---------|---------|-------------|
| VaultPath | *(set on first run)* | Path to your Obsidian vault |
| VaultName | *(auto-detected)* | Vault name for Obsidian URI links |
| TaskFileName | `Inbox.md` | File to append tasks to |
| UseDailyNotes | `false` | Use date-based filenames instead |
| DailyNotesFormat | `yyyy-MM-dd` | Filename format for daily notes |

## License

MIT
