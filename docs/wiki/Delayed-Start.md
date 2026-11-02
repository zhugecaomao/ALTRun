**English** · [中文](Delayed-Start-zh)

# Delayed Start

Preferences → Delayed Start: start programs a while after you sign in to Windows. Programs that don't need to be there at once (OneDrive, sync tools, chat apps) then wait a few minutes, and the desktop is ready sooner.

| Field | Example | Description |
|---|---|---|
| Name | `OneDrive` | Shown in the list; also the name of the task in Task Scheduler |
| Program | `%LocalAppData%\Microsoft\OneDrive\OneDrive.exe` | The program or script to start. Environment variables, `A_` folders such as `A_AppData`, paths relative to the ALTRun folder and program names in `PATH` (`notepad.exe`) work |
| Arguments | `/background` | Optional |
| Delay (seconds) | `300` | Seconds after signing in; `300` = 5 minutes, `0` = at once |
| Start it after signing in | ticked | Untick to keep the entry without starting it |

Click **Apply** or **OK** to set the programs up.

![Preferences → Delayed Start](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-delayedstart.png)

## How it works
- Each entry becomes a task in Windows Task Scheduler (the **Task Scheduler** button on the page opens it), in the folder `ALTRun\<your user name>`, which starts when you sign in, after the delay. Windows starts the programs, so they start even when ALTRun isn't running, and the delay counts from signing in.
- The tasks run in your session with your normal rights, like programs you start yourself. No password is stored. Programs that need administrator rights can't be started this way.
- Unlike Task Scheduler's defaults, the tasks also start on battery, keep running when you unplug, have no time limit and run at normal priority.
- The folder belongs to ALTRun: ALTRun makes it match the list, so changes made to these tasks in Task Scheduler are replaced. An empty list deletes the folder.
- ALTRun checks the tasks 30 seconds after it starts, when the list isn't empty: with the same settings file on another PC (for example in OneDrive), the tasks are set up there too, with that PC's paths.
- Each user of the PC has their own tasks; other users' tasks can't be seen.

## Programs that start by themselves
Many programs start by themselves at sign-in (OneDrive, Teams...). Such a program would start at once, not after the delay. When you add one, ALTRun says where it starts from (`Run: OneDrive`, `Startup: Teams.lnk`); turn its own start off in Task Manager → Startup apps or in the program's settings.

Example, OneDrive after 5 minutes:
1. In OneDrive's settings, turn off "Start OneDrive automatically when I sign in to Windows" (or turn OneDrive off in Task Manager → Startup apps).
2. Add an entry: program `%LocalAppData%\Microsoft\OneDrive\OneDrive.exe`, arguments `/background` (starts without opening the OneDrive window), delay `300`.

## In the settings file
```jsonc
"DelayedStart": [
  { "Name": "OneDrive", "Target": "%LocalAppData%\\Microsoft\\OneDrive\\OneDrive.exe", "Arguments": "/background", "Delay": 300, "Enabled": 1 }
]
```
