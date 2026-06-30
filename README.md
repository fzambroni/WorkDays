# Work Days

**Work Days** is a Windows desktop application for planning, categorizing, reporting, and synchronizing workday status across a yearly calendar. It is built with **AutoIt** and stores its data locally in the Windows Registry under the current user profile.

The project also includes an optional **Outlook Agent** that synchronizes Work Days entries with Microsoft Outlook calendar all-day events, including safety guards, backups, verbose diagnostics, and manual sync controls.

> Developed by **Fabricio Zambroni**.

---

## Table of contents

- [Overview](#overview)
- [Main features](#main-features)
- [Day statuses](#day-statuses)
- [Outlook Agent](#outlook-agent)
- [Sync safety and backup protection](#sync-safety-and-backup-protection)
- [Project structure](#project-structure)
- [Requirements](#requirements)
- [Build instructions](#build-instructions)
- [Installation and first run](#installation-and-first-run)
- [Data storage](#data-storage)
- [Backups and restore](#backups-and-restore)
- [Logs and diagnostics](#logs-and-diagnostics)
- [Recommended Outlook Agent settings](#recommended-outlook-agent-settings)
- [Troubleshooting](#troubleshooting)
- [Development notes](#development-notes)
- [Roadmap ideas](#roadmap-ideas)
- [License](#license)

---

## Overview

Work Days provides a visual calendar-based way to manage where and how each workday is categorized throughout the year. It supports yearly and monthly views, quarter summaries, markers/notes, reports, custom colors, holiday import, backup/restore, and Outlook calendar synchronization.

Typical use cases include:

- tracking **On Site**, **Remote**, **Holiday**, **PTO**, **Travel**, and **Sick** days;
- reviewing quarterly and yearly workday distribution;
- adding markers or notes to specific dates;
- publishing Work Days entries to Outlook as all-day calendar events;
- importing manual Outlook all-day entries back into Work Days;
- keeping a local backup of the Work Days registry database.

---

## Main features

### Calendar management

- Full-year workday planner.
- Monthly calendar and yearly grid views.
- Quarter tabs and yearly statistics.
- Selected-day actions for common statuses.
- Context menu actions for dates.
- Weekend handling.
- Holiday import.
- Batch import of day entries.

### Status tracking

- Categorize days as On Site, Remote, Holiday, PTO, Travel, Sick, Blank, or Weekend.
- Count and summarize statuses by month, quarter, and year.
- Optional synchronization of Blank and Weekend days when explicitly enabled.

### Markers and notes

- Add/edit marker text for individual dates.
- Highlight dates that need attention.
- Optionally show marker information in Outlook subjects.
- Optionally create a separate Outlook category for marker days.
- Optional Outlook reminders when markers exist.

### Reports

- Generate report-style output from the Work Days data.
- Use status totals and calendar data for yearly or quarterly review.

### Customization

- Custom colors for day types.
- Text contrast support.
- Layout and UI settings.
- Persistent local configuration.

---

## Day statuses

Work Days uses compact status codes internally:

| Code | Status |
|---|---|
| `O` | On Site |
| `R` | Remote |
| `H` | Holiday |
| `P` | PTO |
| `T` | Travel |
| `S` | Sick |
| `B` | Blank |
| `W` | Weekend |

Registry day values are stored as:

```text
<StatusCode><OptionalMarkerText>
```

Example:

```text
O
RNeed to confirm customer visit
P
```

---

## Outlook Agent

The **WorkDays Outlook Agent** is a companion executable that synchronizes Work Days data with Microsoft Outlook.

It runs locally, uses the desktop Outlook COM interface, and creates/updates Outlook calendar items as all-day events.

### What appears in Outlook

By default, the agent creates all-day calendar entries with subjects such as:

```text
WorkDays - On Site
WorkDays - Remote
WorkDays - Holiday
WorkDays - PTO
WorkDays - Travel
WorkDays - Sick
```

If marker display is enabled, subjects may include a marker suffix, for example:

```text
WorkDays - On Site [Marker]
```

The agent also uses Outlook categories such as:

```text
WorkDays - On Site
WorkDays - Remote
WorkDays - Marker
```

### Manual Outlook entries supported

The agent can read manually created Outlook all-day events back into Work Days, as long as they are recognized as Work Days items.

Supported subject examples:

```text
WorkDays - On Site
WorkDays - Remote
WorkDays - Travel
WorkDays - PTO
WorkDays - Sick
W - On Site
W - Remote
W: Travel
WD - PTO
WorkDay - Sick
[WD:O]
[WD:R]
```

For manual Outlook-to-Work Days import, keep this setting unchecked unless you only want items created by the agent to be read:

```text
Only read items created by the agent
```

### Main-screen Sync button

The main Work Days screen includes a **Sync** button near the status buttons. It requests an immediate Outlook Agent synchronization.

If the agent is already running, it detects the sync request and starts immediately. If the agent is not running, Work Days can launch a one-time sync execution.

### Refresh notification

When the Outlook Agent changes the Work Days database, the main Work Days window uses the existing update area to show a refresh notification such as:

```text
OUTLOOK CHANGE - Refresh
```

Clicking it reloads the Work Days screen so the UI reflects database changes made by the agent.

---

## Sync safety and backup protection

The Outlook Agent includes several protections to prevent accidental mass changes or data loss.

Before applying Outlook-to-Work Days changes, the agent builds a **sync plan** and validates it. If the plan looks unsafe, no changes are applied.

### Safety rules

Default safety behavior:

| Setting | Default | Purpose |
|---|---:|---|
| Create backup before Outlook changes Work Days | On | Creates a registry backup before database writes. |
| Block mass changes | On | Prevents unexpectedly large sync operations. |
| Max Work Days changes per sync | `20` | Blocks sync if too many Work Days records would change. |
| Max change percentage per sync | `15%` | Blocks sync if the change ratio is too high. |
| Max clears per sync | `0` | Prevents Outlook reads from clearing many Work Days records. |
| Block incomplete Outlook read | On | Blocks sync if Outlook returns suspiciously few items. |

### Safe deletion philosophy

By default, missing Outlook items should **not** erase Work Days data.

The safest behavior is:

```text
If Work Days has a syncable record and Outlook is missing the item,
recreate the Outlook item from Work Days.
```

The dangerous behavior is:

```text
If Outlook is missing an item,
clear the Work Days record.
```

For this reason, the following setting should normally remain off:

```text
Deleting the Outlook item clears WorkDays
```

Even if it is enabled, the mass-change guard and max-clear limit are designed to prevent accidental large deletes.

### Sync plan file

The agent writes the latest sync plan to:

```text
<AgentFolder>\Logs\LastSyncPlan.txt
```

This file shows what the agent planned to do before applying or blocking a sync.

Example blocked plan summary:

```text
SYNC PLAN - BLOCKED
Reason: Too many changes detected.
Current WorkDays records in range: 190
Outlook WorkDays items found: 4
Planned updates: 3
Planned clears: 79
Action: No changes were applied.
```

---

## Project structure

Recommended source layout:

```text
WorkDays/
├─ Workdays.au3
├─ Workdays_Outlook_Agent.au3
├─ Workdays_Backup.au3
├─ Help.html
├─ splash.jpg
├─ About.db
├─ Updater.exe
├─ FileUpdate.exe
├─ CalendarSync.ico
├─ Workdays_Report_HTML_UTF8.au3
├─ Workdays_HTML_TOX.au3
├─ Workdays_Monitor_UDF.au3
├─ Workdays_ColorChooser.au3
├─ Workdays_ColorPicker.au3
├─ Updater_lib2.au3
└─ README.md
```

Important files:

| File | Purpose |
|---|---|
| `Workdays.au3` | Main Work Days desktop application. |
| `Workdays_Outlook_Agent.au3` | Outlook synchronization agent. |
| `Workdays_Backup.au3` | Shared backup library used by both Work Days and the Outlook Agent. |
| `Help.html` | Local HTML help file embedded or distributed with the app. |
| `splash.jpg` | Splash screen image resource. |
| `Updater.exe` | Updater executable resource. |
| `FileUpdate.exe` | Post-build/update helper used by the wrapper flow. |
| `CalendarSync.ico` | Outlook Agent icon. |

---

## Requirements

### Runtime requirements

- Windows 10 or Windows 11.
- Microsoft Outlook desktop application installed and configured.
- A valid Outlook profile with access to the default calendar.
- Current-user registry access under `HKEY_CURRENT_USER`.

### Build requirements

- AutoIt v3.
- AutoIt3Wrapper / SciTE4AutoIt3 recommended.
- Source dependencies listed in the project structure.
- The Outlook Agent executable must be compiled before compiling the main Work Days app if the main app embeds the agent.

The source currently references local wrapper resource paths such as:

```text
.\WorkDays\Help.html
.\WorkDays\splash.jpg
.\WorkDays\Updater.exe
.\WorkDays\Workdays_Outlook_Agent.exe
```

If your repository is stored elsewhere, update the `#AutoIt3Wrapper_Res_File_Add` paths before compiling.

---

## Build instructions

### 1. Compile the Outlook Agent first

Compile:

```text
Workdays_Outlook_Agent.au3
```

Output it as:

```text
.\WorkDays\Workdays_Outlook_Agent.exe
```

The exact output path can be changed, but it must match the path embedded or referenced by the main Work Days build.

### 2. Compile the main Work Days application

Compile:

```text
Workdays.au3
```

The main build can embed or distribute:

```text
Workdays_Outlook_Agent.exe
Help.html
splash.jpg
Updater.exe
About.db
```

### 3. Keep the shared backup library available

Both source files depend on:

```text
Workdays_Backup.au3
```

Keep it in the same source folder as the main app and the Outlook Agent while compiling.

---

## Installation and first run

1. Build `Workdays_Outlook_Agent.exe`.
2. Build `WorkDays.exe` from `Workdays.au3`.
3. Launch `WorkDays.exe`.
4. Open:

```text
Settings > Outlook Agent
```

5. Click:

```text
Install / Update
```

6. Configure sync behavior and safety settings.
7. Click **Save**. The settings window saves and closes.
8. Use the main-screen **Sync** button to force an immediate sync.

---

## Data storage

Work Days stores its main database in the Windows Registry:

```text
HKEY_CURRENT_USER\Software\WorkDays
```

Day entries are organized by year, month, and day:

```text
HKEY_CURRENT_USER\Software\WorkDays\YYYY\MM
  DD = <StatusCode><OptionalMarkerText>
```

Example:

```text
HKEY_CURRENT_USER\Software\WorkDays\2026\07
  02 = O
  03 = RCustomer workshop
```

Outlook Agent settings are stored under:

```text
HKEY_CURRENT_USER\Software\WorkDays\OutlookAgent
```

The agent also uses local files in its executable folder:

```text
Workdays_Outlook_Agent_State.ini
Workdays_Outlook_Agent.log
Logs\LastSyncPlan.txt
Backup\Agent_PreSync_*.bkp
```

---

## Backups and restore

Manual backups are available from the Work Days UI.

The backup library exports the Work Days registry database into `.bkp` files. The Outlook Agent also uses this library to create automatic pre-sync backups before applying Outlook-to-Work Days changes.

Default backup locations include:

```text
<WorkDaysFolder>\Backup
<AgentFolder>\Backup
```

Example automatic backup name:

```text
Agent_PreSync_2026_06_30_224500.bkp
```

Recommended practice:

- Keep automatic backup enabled.
- Do not relax mass-change protection unless you intentionally need a large sync.
- Before major imports or cleanup operations, create a manual backup.

---

## Logs and diagnostics

The Outlook Agent log is stored next to the agent executable:

```text
Workdays_Outlook_Agent.log
```

Verbose diagnostic logging can be enabled in:

```text
Settings > Outlook Agent > Logging
```

Verbose logging records details such as:

- sync range;
- Outlook calendar folder being read;
- Outlook `Restrict` filter;
- detected Outlook candidates;
- subject/category/status parsing;
- date conversion;
- Work Days registry reads/writes;
- state file updates;
- sync plan summary;
- safety guard decisions.

Useful log searches:

```text
Sync safety plan built
Sync blocked by safety guard
Pushed WorkDays change into Outlook
Pulled Outlook change into WorkDays
Re-created missing Outlook item
Date parser failed
Outlook candidate map loaded
```

---

## Recommended Outlook Agent settings

For normal safe operation:

```text
Start with Windows: optional
Outlook wins on conflict: optional, based on preference
Deleting the Outlook item clears WorkDays: OFF
Sync Blank: OFF
Sync Weekend: OFF
Sync tagged Blank/Weekend: ON
Only read items created by the agent: OFF if you manually create Outlook entries
Verbose log: OFF normally, ON during troubleshooting
Create backup before Outlook changes WorkDays: ON
Block mass changes: ON
Max changes per sync: 20
Max change percentage per sync: 15
Max clears per sync: 0
Block incomplete Outlook read: ON
```

---

## Troubleshooting

### Work Days data is not going to Outlook

Check the agent log for day decisions and push messages:

```text
Day decision YYYY-MM-DD
Pushed WorkDays change into Outlook
Re-created missing Outlook item
Created Outlook item for existing WorkDays date
```

If the state file is stale, the agent should now recreate missing Outlook items from Work Days. The state file is:

```text
<AgentFolder>\Workdays_Outlook_Agent_State.ini
```

Renaming it can still be used as a manual recovery step, but the reconcile logic should normally avoid requiring that.

### Outlook changes are not coming back to Work Days

Confirm:

- the Outlook item is an all-day event;
- the date is within the sync range;
- the subject/category is recognized as a Work Days item;
- `Only read items created by the agent` is off if the item was created manually;
- verbose logging is enabled while troubleshooting.

Supported manual subject examples:

```text
W - On Site
W - Remote
W: Travel
WD - PTO
[WD:O]
```

### Sync was blocked

If the Work Days UI shows:

```text
SYNC BLOCKED - Review
```

Open the review message and inspect:

```text
<AgentFolder>\Logs\LastSyncPlan.txt
<AgentFolder>\Workdays_Outlook_Agent.log
```

The block is usually caused by:

- too many planned changes;
- too many planned clears;
- Outlook returning suspiciously few items;
- pre-sync backup failure.

### Popup windows appear behind the main UI

Outlook Agent settings dialogs and confirmation popups are configured to use top-most behavior so they should stay visible above the main Work Days window.

### Outlook calendar looks duplicated

Check whether old Work Days items exist with a different prefix/category. Use the Outlook cleanup tool carefully and keep the default confirmation phrase enabled.

---

## Development notes

### Code style

- UI labels, settings names, comments, and messages are primarily in English.
- Registry settings use flattened names under `HKEY_CURRENT_USER\Software\WorkDays\OutlookAgent`.
- The Outlook Agent runs independently from the main Work Days UI.
- Shared backup functionality belongs in `Workdays_Backup.au3` rather than being duplicated in each executable.

### Important design decisions

- Work Days is the safest source of truth for rebuilding missing Outlook items.
- Missing Outlook items should not clear Work Days records unless explicitly enabled.
- Outlook-to-Work Days changes require a pre-sync backup and pass safety validation.
- The agent state file is an optimization and reconciliation helper, not the only authority.

### Local-only behavior

Work Days does not require a cloud service for its core functionality. Data is stored locally in the Windows Registry and synchronized with the local Outlook desktop profile through COM automation.

---

## Roadmap ideas

Potential future improvements:

- Export/import settings to a portable config file.
- Add a dedicated backup manager UI.
- Add a sync history viewer inside Work Days.
- Add a side-by-side Outlook vs Work Days reconciliation preview.
- Add unit-testable helper modules for date parsing and status parsing.
- Add GitHub Actions or a documented release checklist for packaged builds.
- Add signed releases and checksums.

---

## License

No license has been declared yet.

Before publishing publicly, add a license file such as:

- `MIT License`, if you want permissive reuse;
- `Apache License 2.0`, if you want permissive reuse with explicit patent terms;
- `All rights reserved`, if you do not want to grant reuse rights.

---

## Disclaimer

This project automates Microsoft Outlook through the local desktop Outlook COM interface and writes user-specific data to the Windows Registry. Always keep backups enabled before running synchronization, cleanup, import, or restore operations.
