**English** · [中文](Extensions-zh)

# Extensions

Features outside the search window. These have their own pages, which `F1` on the matching Preferences page also opens:

| Page | What it covers |
|---|---|
| [Quick Switch](Quick-Switch) | Jump an Open / Save dialog to the folder open in Total Commander, Directory Opus or Explorer (`Ctrl+G`), auto switch |
| [Dialog Folder Panel](Quick-Switch-Panel) | The folder panel under dialogs: open and recent folders, the search box (`Ctrl+Shift+G`) |
| [Date Stamp](Date-Stamp) | Add a date while renaming a file or at the end of a line of text (`Ctrl+D`) |
| [Custom Hotkeys](Hotkeys) | Assign system commands to hotkeys; Caps Lock switches the input method; how to set hotkeys |
| [Calculator](Calculator) | Expressions, unit and currency conversion, structural calculations |
| [Scripts](Scripts) | Turn your own scripts into commands in the search results |
| [System commands and terminal](System-Commands) | The full list of system commands (lock, shut down, empty the Recycle Bin...); `>` runs a command in the terminal |

PT Tools has no page of its own, see below.

## System commands
See [System commands and terminal](System-Commands#system-commands).

## Terminal
See [System commands and terminal](System-Commands#terminal).

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
