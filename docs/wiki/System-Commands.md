**English** · [中文](System-Commands-zh)

# System commands and terminal

## System commands
Type the name in the search box (English, Chinese or the Id all work); they can also be used in custom hotkeys.

![System commands: lock](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/system.png)

| Id | Command |
|---|---|
| **ALTRun** | |
| `Preferences` | ALTRun Preferences |
| `Preferences.Appearance` etc. | ALTRun Preferences: Appearance (one per settings page, opening that page directly; type the page name to find it, for example `appearance` or `hotkeys`) |
| `Reload` | Reload ALTRun |
| `RebuildIndex` | Rebuild ALTRun Index |
| `CheckUpdate` | Check for ALTRun Updates |
| `About` | About ALTRun: opens the Advanced page of Preferences (version, project page, check for updates) |
| `Log` | Open ALTRun Log |
| `Quit` | Quit ALTRun |
| **System** | |
| `Lock` `Sleep` `Hibernate` | Lock Screen / Sleep / Hibernate |
| `Shutdown` `Restart` `Logoff` | Shut Down / Restart / Log Off (confirmed first) |
| `EmptyRecycle` | Empty Recycle Bin (confirmed first) |
| `MonitorOff` | Turn Off Monitor |
| `Mute` `VolumeUp` `VolumeDown` | Toggle Mute / Volume +10 / Volume -10 |
| `MediaPlayPause` `MediaNext` `MediaPrev` `MediaStop` | Play / Pause, Next Track, Previous Track, Stop Media (music and video players) |
| `ShowIP` | Show IP Address |
| `TerminalHere` | Open a terminal in the current folder (Explorer / Total Commander) |
| `ListProcesses` `ListServices` | List running processes / services |
| `PTTools` `SPF2M` | PT Tools / SPF2M |
| **Clipboard text** (converted and put back on the clipboard) | |
| `TextUpper` `TextLower` `TextTitle` | UPPERCASE / lowercase / Title Case |
| `TextSortAsc` `TextSortDesc` | Sort lines |
| `TextTrimLines` `TextRemoveBlank` `TextDedupe` | Trim each line / remove blank lines / remove duplicate lines |
| `TextReverse` | Reverse text |
| `TextToTraditional` `TextToSimplified` | Simplified → Traditional / Traditional → Simplified Chinese |
| `TextUrlEncode` | URL encode |
| **Windows tools** | |
| | Task Manager, Control Panel, Windows Settings, Device Manager, Services, Registry Editor, Event Viewer, Disk Management, Computer Management, Task Scheduler, Programs and Features, System Properties, Network Connections, Windows Defender Firewall, Resource Monitor, Disk Cleanup, System Configuration, Group Policy Editor, Command Prompt, PowerShell, File Explorer, Recycle Bin, This PC, Devices and Printers, Notepad, Calculator, Paint, About Windows |

The confirmation for shut down, restart and similar commands can be turned off with `Features.System.ConfirmActions = 0`.

## Terminal
Type `>command` to run it in a terminal that stays open, for example `>ipconfig /all` or `>ping 8.8.8.8`. "Open Terminal Here" in the action panel opens a terminal in the file's folder.

The terminal program `Features.Terminal.Shell`: `cmd` (default) / `powershell` / `pwsh` / `wt` (Windows Terminal); the prefix `Prefix` is `>` by default.
