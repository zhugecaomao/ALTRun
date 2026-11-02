**English** · [中文](Sync-Tasks-zh)

# Sync Tasks

Keep a folder in step with another one, for example a server share copied into your OneDrive, with Windows' own `robocopy`. ALTRun runs it in the background, checks risky settings first and tells you how it went.

> The editor in Preferences comes in a later version; until then the tasks are set in the settings file (see [below](#in-the-settings-file)).

## Running a task
- Type the task's name, or `sync ` to list all tasks (`sync proj` searches only them). **Enter** runs the task; while it runs, Enter stops it.
- `→` opens the actions: Run, Preview (changes nothing), Open log, open the source or destination folder.
- The result shows what the task does and how it is doing: *Running since 10:05*, *Monitoring since 10:05*, *Last run 09:30: OK*.
- A task that ends says how it went (files copied and removed). If it fails, ALTRun offers to open its log.

## Safety
- **Mirror** (`/MIR`) makes the destination an exact copy of the source: files in the destination that aren't in the source are **deleted**. Before a mirror runs, ALTRun previews it (`/L`, which changes nothing); when the run would delete anything, it says how many files and folders and asks first.
- A mirror doesn't run when the source folder can't be found or is empty (for example a server share that didn't connect), since that would empty the destination.
- A task doesn't run when the source and destination are the same folder or one is inside the other, or when a mirror would go into a drive, a share, the OneDrive folder or the user folder (it would delete everything else there).
- **Monitoring** tasks (`/MON`, `/MOT`) keep running and sync again after changes. They stop when ALTRun quits; when ALTRun only restarts (after changing a setting, an update), they keep running.

## Logs
Each run is logged: to the task's own log file, or to `%TEMP%\ALTRun\SyncLogs\<task name>_<date>.log`. **Open log** opens today's log.
- The logs aren't kept in the Data folder: when that is in OneDrive, logs that grow all the time would keep uploading.
- Logs in `%TEMP%\ALTRun\SyncLogs` older than 30 days are deleted when a task runs (`Features.Sync.KeepLogDays`, `0` = keep them all). A task's own log file is never deleted.
- `ALTRun.log` gets one line per run: when it started and how it ended (files copied and deleted, or the exit code).

## In the settings file
```jsonc
"SyncTasks": [
  {
    "Name": "Projects",
    "Source": "\\\\fileserver\\Projects",
    "Destination": "%OneDrive%\\Work\\Projects",
    "Files": "*.*",
    "Mode": "Mirror",            // Mirror (/MIR), Update (/E /XO: new and newer files, never deletes), Copy (/E: all that differ, never deletes) or Top (top folder only)
    "Copy": "DAT", "DirCopy": "DA",
    "Retries": 1, "Wait": 1, "Threads": 16,
    "Monitor": 0, "MonitorChanges": 1, "MonitorMinutes": 5,
    "ExcludeFiles": ["thumbs.db", "desktop.ini", "~$*.*"],
    "ExcludeDirs": ["System Volume Information", "$Recycle.Bin"],
    "LogFile": "%OneDrive%\\Logs\\Projects_{date}.log", "AppendLog": 1,
    "Extra": "/NP"               // other robocopy options, as they are
  }
]
```
Paths may use `%environment variables%`; the log file name may use `{date}` (yyyyMMdd) and `{time}` (HHmmss).
