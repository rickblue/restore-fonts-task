# restore-fonts-task

Windows font restore helper for reapplying custom UI font settings after logon, unlock, or remote session reconnect.

## What It Does

This project contains a PowerShell script that uses `SystemParametersInfo` to immediately reapply selected Windows UI fonts.

Current behavior:

- Leaves title bar fonts unchanged, so the caption height is not stretched
- Applies `MiSans Medium` to menu, status, and message fonts
- Applies `MiSans Medium` to desktop icon title font
- Supports manual execution and scheduled automatic restore

## Files

- `restore-fonts.ps1`: main script that applies font settings and writes logs
- `manage-task.ps1`: elevated task manager script for installing or updating the scheduled task
- `restore-fonts-task.xml`: scheduled task template

## Requirements

- Windows
- PowerShell
- Target fonts installed locally, especially `MiSans Medium`

## Run Once

```powershell
powershell -ExecutionPolicy Bypass -File .\restore-fonts.ps1
```

## Install Scheduled Task

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\manage-task.ps1
```

Then use the menu to:

- Add task
- Update task
- Delete task
- Run now

The scheduled task is intended to run on:

- User logon
- Session unlock
- Remote connect

## Notes

- The script writes logs to `restore-fonts.log`
- Local-only files such as logs, shortcuts, and `.claude/` are ignored in git
