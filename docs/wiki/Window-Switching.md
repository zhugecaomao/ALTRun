**English** · [中文](Window-Switching-zh)

# Window switching

Type a window title or program name and `Enter` switches to that open window (restoring it if it is minimized), like PowerToys Run's Window Walker or Raycast's Switch Windows.

## Finding a window
| You type | You get |
|---|---|
| `w ` (`w` and a space) | Every window, front to back (the one you used last first) |
| `w excel` | Windows of Excel, and windows with "excel" in the title |
| `w quote` | Windows whose title contains "quote", such as `Quote 2026-118.xlsx - Excel` |
| `w bjb` | Pinyin initials of the title work too: a window titled `笔记本` |
| `quote` (no keyword) | Apps and commands first, then at most 3 windows whose title matches well (from 2 characters on) |

A window is matched by its title, its program name (`excel` for `EXCEL.EXE`) and the pinyin initials of its title, the same way as other results ([how matching works](Usage#how-matching-works)). Every window is a result of its own: three documents open in Word are three results, each with its title and "WINWORD.EXE - switch to this window" below it.

## Switching and closing
- `Enter` (or a double-click) switches to the window, restoring it first if it is minimized
- `→` on a window: **Close Window**. It is closed the normal way, as with its × button, so a program with unsaved changes still asks first
- When the window was closed in the meantime, a notification says "That window has been closed"

![w and a space lists the open windows](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/windows.png)

## Which windows are listed
The same windows as on the taskbar:

| Listed | Not listed |
|---|---|
| Visible windows with a title, on the current virtual desktop | Hidden windows, windows without a title, tool windows (floating palettes), dialogs owned by another window, windows on other virtual desktops, suspended Store apps, the desktop, ALTRun's own windows |

The list is read again for each search (reused while you keep typing within a second), so a window opened a moment ago shows up. Windows are never added to the [recent items](Recent) and the learned ranking doesn't apply to them: a window is there only while it is open.

## Settings
Preferences → Features: "Window switching" turns the whole feature on or off; under "Feature options", "Show in normal search" and "Keyword" map to:

| Setting | Default | Description |
|---|---|---|
| `Features.Windows.Enabled` | `1` | Turn window switching on or off |
| `Features.Windows.Keyword` | `w` | Keyword that searches windows only |
| `Features.Windows.InDefaultResults` | `1` | Also show windows whose title matches well when typing without the keyword |

To keep windows out of your normal results, turn off "Show in normal search" and use `w` when you want one. To use another keyword (for example `ww` when `w` is one of your own keywords), change "Keyword".

## Related pages
- [Search and shortcuts](Usage): keywords and how results are ranked
- [Custom hotkeys](Hotkeys): a hotkey with the action "Search in ALTRun" and the search text `w ` lists the windows with one key
