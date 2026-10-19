**English** · [中文](Extensions)

# Extensions

This is the legacy combined guide, retained for compatibility. Features with their own Preferences pages have dedicated guides: [Quick Switch](Quick-Switch), [Dialog Panel](Dialog-Panel), [Date Stamp](Date-Stamp), [Custom Hotkeys](Hotkeys), [Calculator](Calculator), and [Scripts](Scripts).

## Quick Switch

See [Quick Switch](Quick-Switch) for the full guide.

## Date stamp

See [Date Stamp](Date-Stamp) for the full guide.

## Custom hotkeys

See [Custom Hotkeys](Hotkeys) for the full guide.

### Caps Lock switches the input method
See [Custom Hotkeys](Hotkeys#caps-lock-switches-the-input-method).

### Setting hotkeys
See [Custom Hotkeys](Hotkeys#setting-hotkeys).

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

## Calculator

See [Calculator](Calculator) for the full guide.

### Unit and currency conversion
See [Calculator](Calculator#unit-and-currency-conversion).

## Scripts

See [Scripts](Scripts) for the full guide.

## PT Tools
Small tools for post-tensioning design; type `PTTools` / `SPF2M` or open them with a custom hotkey:
- **PT Tools**: rebar / BRC mesh area calculator, plus an expression calculator
- **SPF2M Post-Tensioning Tendon Profile Calculator**: calculated directly in ALTRun, no DOSBox needed any more. Pick the profile (double parabola / parabola-straight-parabola / parabola-straight / straight-parabola) and the strand type, enter the start and end levels and the horizontal distance, and the table on the right immediately gives the height at every support (Actual, Beam rounded to 5 mm, Slab rounded to 10 mm), plus the radius of curvature and the inflection point position
  - The results match the original SPF2M.EXE exactly: 139 cases (4 profiles × 6 strand types, rising / falling, non-whole-meter spans, custom radius / inflection point / support spacing) were run with the original program in DOSBox, and more than 1,100 values are compared one by one as part of the unit tests
  - Leaving an optional field empty = SPF2M's default (shown in gray): **minimum radius of curvature** (Slab 5000, 7S 3200, 12S 4200, 19S 5300, 22S 5700, 31S 6700), **inflection point distance** (when given, the radius is worked back from it, with a warning if it's below the minimum radius), **support spacing** (at most 1000 mm by default, with the remainder on the high-point side; you can also write `500, 1500, 800 ...`, which must add up to the horizontal distance)
  - **At C.G.**: tick it when the levels are given at the tendon center; half the duct diameter is subtracted automatically. With the duct diameter empty, the default is used (Slab 25, 7S 70, 12S 90, 19S 100, 22S 120, 31S 130); a value you enter is remembered for each strand type
  - The table starts from the high point (like SPF2M); **Copy Table** copies it as a table to paste straight into Excel
  - SPF2M simply quits on some impossible profiles; here you get an explanation instead. When the inflection point is beyond half the span SPF2M shows nothing, and here you get a warning too
  - The original SPF2M.EXE and DOSBox are no longer shipped. After upgrading from an old version, `DOSBox.exe`, `SDL.dll`, `SDL_net.dll`, `SPF2M.exe` and `Run.bat` in `Resources\` can be deleted
