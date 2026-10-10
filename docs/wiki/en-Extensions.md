**English** · [中文](Extensions)

# Extensions

Features outside the search window. Quick Switch and date stamp each have their own settings pages (Preferences → Quick Switch / Dialog Panel / Date Stamp).

## Quick Switch
Jump an Open / Save dialog to a folder you already have open. The approach follows [Listary](https://www.listary.com/)'s Quick Switch.

**Folder panel** (on by default, like Listary's Quick Switch window): as soon as a dialog appears, a panel as wide as the dialog attaches right below it (above it when there's no room below). Click a folder and the dialog jumps there, no hotkey to remember:

![Quick Switch folder panel below an Open dialog](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/quickswitch.png)

- Lists the current panel and the other panel of every Total Commander window, open Explorer windows and recently used folders;
- The search box at the top: typed text first filters the list, then Everything (or the built-in index) finds folders and files with matching names (folders first); `↑` `↓` select, `Enter` confirms, `Esc` returns to the dialog;
- When a file is selected: jumps to the folder it's in (the same for Open and Save dialogs; it doesn't open or save for you);
- The file name already in the file name box (for example the name a program suggests in Save As) is kept when jumping;
- The panel's colors follow ALTRun's theme (Preferences → Appearance), so with a dark theme the panel is dark too; the panel follows the dialog when it's resized;
- With a Chinese input method in the search box, `Enter` commits the pinyin rather than jumping; press `Enter` again after committing to jump;
- What the search box finds: Preferences → Dialog Panel → "Search box finds": Folders only (default) / Folders and files (`PanelSearch`: `folders` / `all`). Folders are searched separately, so they aren't crowded out when many files have the same name;
- Hidden when the dialog isn't in the foreground;
- Switch to TC, change folders and come back to the dialog, and the list is refreshed;
- The panel's settings (on / off, what to search, number of recent folders, keyboard hotkey) are in Preferences → Dialog Panel (`ShowPanel`). Recent folders come from Windows' "Recent items" (for opened files, the folder they're in), 10 by default, 0 = none.

Keyboard, in a standard "Open / Save file" dialog:
- `Ctrl+G`: jump to the folder open in Total Commander
- `Ctrl+E`: jump to the folder open in Explorer
- `Ctrl+Shift+G`: move the cursor to the folder panel's search box and pick with the keyboard: type to search, `↑` `↓` select, `Enter` jumps, `Esc` returns to the dialog. When the panel isn't shown automatically, this hotkey shows it temporarily

With "Jump automatically when switching from Total Commander" (`AutoSwitch`) turned on, switching from Total Commander to a dialog jumps automatically.

The TC folder is asked from TC directly (TC's `WM_COPYDATA` interface, TC 8.0 and later), not through the clipboard, so no path ends up in the Windows clipboard history (`Win+V`).

Preferences → Quick Switch:
- **Works in**: "Standard Open / Save dialogs (ahk_class #32770)" (ticked by default) and "Other dialogs", for example WPS's `ahk_class Qt5QWindowIcon`, one per line
- **Never in**: no jumping in these windows
- **No auto jump in**: no automatic jump in these dialogs; the hotkeys still work

In the settings file this is `Extensions.QuickSwitch` → `TotalCmdHotkey` / `ExplorerHotkey` / `MenuHotkey` / `RecentFolders` / `ShowPanel` / `PanelSearch` / `AutoSwitch` / `DialogWindows` / `ExcludeWindows` / `AutoSwitchExclude` (window conditions separated by commas).

## Date stamp
Press the hotkey (`Ctrl+D` by default, can be changed):
- **While renaming a file** (Explorer, Total Commander, desktop, dialogs): adds ` - date` before the extension; an existing date is updated to today. For example `Report.docx` → `Report - 23.09.2026.docx`
- **In a text note box** (for example Total Commander's file comments): adds ` - date` at the end

Preferences → Date Stamp: each case has its own hotkey and "works in" window list (one per line, for example `ahk_class CabinetWClass` for Explorer, `ahk_class TTOTAL_CMD` for Total Commander). In the settings file this is `Extensions.AutoDate` → `RenameHotkey` / `RenameWindows` / `AppendHotkey` / `AppendWindows`.

The date format `DateFormat` is `dd.MM.yyyy` by default (syntax: [AutoHotkey FormatTime](https://www.autohotkey.com/docs/v2/lib/FormatTime.htm)); `{date}` in snippets uses the same format.

## Custom hotkeys
Preferences → Hotkeys: assign a [system command](#system-commands) to any hotkey, optionally only inside a certain window.

| Field | Example | Description |
|---|---|---|
| Hotkey (Key) | `Ctrl+Alt+P` | Click the box and press the keys (see [Setting hotkeys](#setting-hotkeys) below); mouse hotkeys: click the middle or a side button in the box. With "Keep the key's original function" ticked, the key still does what it normally does (a leading `~` in the settings file) |
| Action | `PTTools` | A system command Id, or `ToggleWindow` (show / hide ALTRun) |
| Only in window (WinTitle) | `ahk_exe notepad.exe` | Empty = everywhere; otherwise only in matching windows; `ALTRun` = only in ALTRun's search window |

The default one: the middle mouse button in RAPT (`RAPTW.exe`) opens PT Tools. `F1`–`F4` in the search window are built in, see [Shortcuts](en-Usage#shortcuts).

### Setting hotkeys
Every hotkey in Preferences (the ALTRun hotkey, clipboard history, Quick Switch, date stamp, custom hotkeys) uses the same kind of box, recording the keys directly like Alfred and PowerToys:
- The box shows `Alt+Space`, `Ctrl+Alt+C` and so on. Click the box, "Press a shortcut..." appears, and the keys you press are recorded
- `Esc` cancels (keeps the old hotkey), `Backspace` / `Delete` clears it (no hotkey)
- A single letter, digit, space, Enter and the like would get in the way of typing and need `Ctrl`, `Alt` or `Win`; `F1`–`F24`, `Pause` and similar keys work on their own
- While recording, ALTRun's own hotkeys are paused, and combinations such as `Alt+Space` or `Win+E` don't trigger Windows or other programs
- When saving, you're warned if two global hotkeys are the same (for example the ALTRun hotkey and clipboard history both `Ctrl+Alt+C`). Hotkeys that only work in some windows (Quick Switch, date stamp, custom hotkeys with a window) can repeat

The settings file still uses AutoHotkey syntax (`!` Alt, `^` Ctrl, `+` Shift, `#` Win, for example `!Space`), and you can edit it directly; special forms such as `CapsLock & J` are kept as they are until you record the hotkey again in Preferences.

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
Type an expression directly: `12*(3+4)`, `=2^10`; `Enter` copies the result. Supports `+ - * /`, powers `^` (or `**`) and parentheses; results keep at most two decimals (`10/3` shows `3.33`; numbers below 0.005 keep one digit after the first significant digit, for example `0.004`).

With "Add structural results (main bars / rebar area)" (`Features.Calculator.StructuralCalc`) turned on in Preferences → Calculator, two lines of structural results are added below when the result is more than 2 × the cover:
- The result as a **beam width (mm)**: number and spacing of main bars. Number = ⌈(width − 2 × cover) / max spacing⌉ + 1, spacing = (width − 2 × cover) / (number − 1)
- The result as a **rebar area As (mm²)**: bars needed for each size (⌈As / area of one bar⌉, area of one bar = π d² / 4)

For example `300*2` (600) shows `Beam width 600 mm: 3 main bars @ 260 c/c` and `As = 600 mm²: 5H13  3H16  2H20  2H25  1H32` by default. The parameters are on the same page:

| Setting | Key | Default | Description |
|---|---|---|---|
| Rebar cover (mm) | `RebarCover` | `40` | Each side, mm |
| Max spacing (mm) | `MaxBarSpacing` | `300` | mm; more main bars are added when the spacing would exceed it |
| Bar sizes (mm) | `BarSizes` | `[13, 16, 20, 25, 32]` | Diameters (mm) listed in the rebar area line |
| Bar prefix | `BarPrefix` | `H` | Bar designation, for example `T`, `Y`, `Φ` |

![Preferences → Calculator](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-calculator.png)

### Unit and currency conversion
Syntax: number + unit + `in` / `to` / `=` / `->` / `转` + target unit; units are case-insensitive, and `m²` can be written `m2`. For example `10 km in mi`, `5ft to cm`, `100 f to c`, `3 亩 in m2`, `1 GB to MB`, `300 kN to kip`, `20 MPa in psi`. `Enter` copies the number.

| Category | Units |
|---|---|
| Length | mm cm m km in ft yd mi nmi (Chinese names work too) |
| Mass | mg g kg t lb oz 斤 两 |
| Area | mm2 cm2 m2 km2 ha acre ft2 in2 亩 |
| Volume | ml cl l m3 cm3 ft3 gal qt pt cup floz |
| Speed | m/s km/h mph knot ft/s |
| Time | ms s min h day week year |
| Data | bit B KB MB GB TB (powers of 1024) |
| Pressure / stress | Pa kPa MPa (N/mm2) bar psi ksi atm psf |
| Force | N kN lbf kip kgf tf |
| Energy / power | J kJ cal kcal Wh kWh / W kW hp |
| Temperature | C F K |

**Currency conversion** (`100 usd to sgd`) is off by default: tick "Currency conversion" in Preferences → Calculator (`Features.Calculator.Currency`). Once on, the reference rates published by the European Central Bank and other central banks are downloaded once a day from [Frankfurter](https://frankfurter.dev) (free, no sign-up) and saved in `Data\Currency.json`; apart from GitHub, this is the only website ALTRun contacts on its own (a web search's site is visited only when you click "Download Site Icon"). Results show the date of the rates.

## Scripts
Put your own scripts into the `Scripts\` folder (in the program folder; Preferences → Scripts → "Open Scripts Folder", or search for "Open Scripts Folder"; a few examples are created the first time: a background script that shows a notification, one that takes an argument after its keyword, and one whose output opens in your text editor) and you can find and run them by name or keyword in the search window, like Raycast's Script Commands. Supported: `.ahk` (run with the AutoHotkey inside ALTRun, nothing else to install), `.ps1`, `.bat` / `.cmd`, `.py` (needs Python). Preferences → Scripts lists the scripts found; double-click one to edit it in Notepad:

![Preferences → Scripts](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-scripts.png)

Comments at the top of a script can hold these settings (the comment markers `;` `#` `REM` `::` `//` all work; all are optional):

| Setting | Description |
|---|---|
| `@altrun.title name` | The name shown; the file name by default |
| `@altrun.keyword keyword` | An exact match ranks first |
| `@altrun.argument hint` | Needs an argument: type "keyword text" and the text is passed to the script as its first argument; when only the name is found, `Enter` completes it to "keyword " |
| `@altrun.mode window` | Default: runs normally (with a window) |
| `@altrun.mode silent` | Runs in the background and shows the last line of its output as a notification when done |
| `@altrun.mode output` | Runs in the background and opens all of its output with the default program for `.txt` files (usually Notepad) when done |

Example (`Scripts\Hash.ps1`):
```powershell
# @altrun.title    File hash
# @altrun.keyword  hash
# @altrun.argument file path
# @altrun.mode     output
Get-FileHash -Algorithm SHA256 $args[0] | Format-List
```
Type `hash D:\Downloads\setup.exe` and press `Enter`; the SHA256 opens in your text editor. Pick a keyword that isn't used by a default command (for example `ping` already belongs to the built-in Ping command).

In `.ahk` scripts, write output with `FileAppend("text", "*")`. Adding, removing or changing scripts takes effect right away without reloading; select a script and press `F3` to edit it in Notepad. Background scripts are given at most 2 minutes.

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
