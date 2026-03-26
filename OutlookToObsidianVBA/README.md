# OutlookToObsidian — VBA Macro Version

A VBA macro port of OutlookToObsidian that requires **zero installation**. Creates Obsidian tasks from Outlook emails via a toolbar with two buttons.

## Why VBA?

The original VSTO add-in uses ClickOnce deployment, which fails on many corporate/managed devices with:
> "Deployment and application do not have matching security zones"

The VBA version eliminates all external dependencies — no VSTO Runtime, no .NET Framework, no ClickOnce, no MSI installer.

## Installation

### Step 1: Open the VBA Editor

1. In Outlook, press **Alt+F11** to open the VBA Editor
2. If prompted about macros, click **Enable Macros**

### Step 2: Import the Modules

1. In the VBA Editor, go to **File → Import File...**
2. Import these files (one at a time):
   - `Settings.bas`
   - `TaskCreator.bas`
   - `TaskEntryForm.bas`

### Step 3: Add ThisOutlookSession Code

1. In the Project Explorer (left panel), expand **Project1 → Microsoft Outlook Objects**
2. Double-click **ThisOutlookSession**
3. Open `ThisOutlookSession.cls` in a text editor and copy everything **below** the `Option Explicit` line
4. Paste it into the ThisOutlookSession code window (after any existing code)

### Step 4: Restart Outlook

1. Close the VBA Editor
2. **Restart Outlook** completely (not just close/reopen a window)
3. On first launch, you'll be prompted to select your Obsidian vault folder
4. An "OutlookToObsidian" toolbar will appear with two buttons

## Usage

1. Select one or more emails in your inbox
2. Click **"Create Task in Obsidian"** on the toolbar for instant task creation
3. Click **"Create Task (detailed...)"** to customize subject, due date, priority, tags, and notes before creating

After creating a task, you'll be asked if you want to open the file in Obsidian.

## Macro Security

If macros are disabled on your device:

1. Go to **File → Options → Trust Center → Trust Center Settings → Macro Settings**
2. Select **"Notifications for all macros"** (recommended) or **"Enable all macros"**
3. Restart Outlook

If your organization enforces macro policies via Group Policy, contact your IT administrator.

## Reconfiguring the Vault

To change your Obsidian vault path:
1. Press **Alt+F11** to open the VBA Editor
2. Press **Ctrl+G** to open the Immediate Window
3. Type `ConfigureVault` and press Enter
4. Select the new vault folder

## Settings

Settings are stored in the Windows Registry under `HKCU\Software\VB and VBA Program Settings\OutlookToObsidian\Settings`.

| Setting | Default | Description |
|---------|---------|-------------|
| VaultPath | *(prompted on first run)* | Full path to your Obsidian vault |
| TaskFileName | `Inbox.md` | File to append tasks to |
| UseDailyNotes | `False` | If True, appends to daily note files instead |
| DailyNotesFormat | `yyyy-mm-dd` | Date format for daily note filenames |
| VaultName | *(auto-detected)* | Name of the vault (folder name) |

To change settings programmatically, use the Immediate Window (Ctrl+G):
```vba
Settings.SetTaskFileName "Tasks.md"
Settings.SetUseDailyNotes True
Settings.SetDailyNotesFormat "yyyy-mm-dd"
```

## Task Format

Tasks are created in [Obsidian Tasks](https://publish.obsidian.md/tasks/) format:

```markdown
- [ ] Re: Q3 Budget Review ⏫ #follow-up #project-alpha 📅 2026-03-28 ➕ 2026-03-22
  > **Jane Smith** | 2026-03-20 14:30 | 📎 2 | ^a3f1b2c4
  > First 140 characters of the email body preview...
```

## Files

| File | Description |
|------|-------------|
| `Settings.bas` | Registry-based settings storage |
| `TaskCreator.bas` | Core markdown generation, vault I/O, duplicate detection |
| `TaskEntryForm.bas` | Detailed task entry dialog (InputBox-based) |
| `ThisOutlookSession.cls` | Startup, toolbar, button handlers, notification |

## Vault Templates

Copy the template files from the `OutlookToObsidian/ObsidianVaultTemplates/` directory into your vault:
- `Inbox.md` — landing file for tasks
- `Task Dashboard.md` — Obsidian Tasks query dashboard

## Differences from VSTO Version

- **Toolbar buttons** instead of right-click context menu (VBA limitation)
- **InputBox dialogs** instead of a single WinForms dialog for detailed mode
- **MsgBox notification** instead of toast popup (with option to open in Obsidian)
- Settings stored in registry instead of .NET user settings
