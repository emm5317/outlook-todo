# OutlookToObsidian

A VSTO add-in for Classic Outlook on Windows that turns emails into tasks in your [Obsidian](https://obsidian.md) vault. Right-click any email, click "Create Task in Obsidian," and a formatted Markdown task appears in your vault — ready for the [Obsidian Tasks](https://publish.obsidian.md/tasks/) plugin.

No servers. No cloud sync. No accounts. Just a right-click menu and a local file write.

## What It Does

- Adds a **"Create Task in Obsidian"** item to the Outlook right-click context menu (single and multi-select)
- Appends a Markdown task to a file in your Obsidian vault using [Obsidian Tasks emoji format](https://publish.obsidian.md/tasks/Reference/Task+Formats/Tasks+Emoji+Format/):
  ```
  - [ ] Re: Q3 Budget Review ⏫ #follow-up #project-alpha ➕ 2026-03-22
    > From: Jane Smith (jane@example.com) | 2026-03-20 14:30
    > First 200 characters of the email body...
    > [entry-id:: 0000000012345ABCDE]
  ```
- Maps **Outlook importance** (High/Low) to Obsidian Tasks **priority** (⏫/🔽)
- Converts **Outlook categories** to **Obsidian #tags**
- Detects **duplicates** — won't create the same task twice
- Prompts for vault folder on first run — no config files to edit
- Supports **daily notes** or a single target file (configurable)

## Requirements

- Windows 10/11
- Classic Outlook (the COM/VSTO desktop version, not "New Outlook")
- [Visual Studio 2022](https://visualstudio.microsoft.com/) with Office/SharePoint development workload (for building)
- [VSTO Runtime 4.0](https://learn.microsoft.com/en-us/visualstudio/vsto/visual-studio-tools-for-office-runtime-installation-scenarios) (for running)
- .NET Framework 4.8
- [Obsidian](https://obsidian.md) with a local vault

## Recommended Obsidian Plugins

- [Tasks](https://publish.obsidian.md/tasks/) — parses due dates, priorities, recurrence, and enables task queries
- [Dataview](https://blacksmithgu.github.io/obsidian-dataview/) — query tasks by `entry-id` or other inline fields
- [Calendar](https://github.com/liamcain/obsidian-calendar-plugin) — visual monthly view if using daily notes

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
- **Task Dashboard.md** — pre-built queries for open tasks, due soon, follow-ups, and completed items

## Project Structure

```
OutlookToObsidian/
├── ContextMenuRibbon.xml      # Ribbon XML — context menu definition
├── ContextMenuRibbon.cs       # Ribbon callback — handles right-click action
├── TaskCreator.cs             # Core logic — markdown formatting, file I/O, dedup
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
| TaskFileName | `Inbox.md` | File to append tasks to |
| UseDailyNotes | `false` | Use date-based filenames instead |
| DailyNotesFormat | `yyyy-MM-dd` | Filename format for daily notes |

## License

MIT
