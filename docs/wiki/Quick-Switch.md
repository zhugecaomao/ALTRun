**English** · [中文](Quick-Switch-zh)

# Quick Switch

Quick Switch jumps an Open / Save dialog to a folder you already have open, without browsing through the directory tree again. For the visual folder panel and its settings, see [Dialog Folder Panel](Quick-Switch-Panel).

## Shortcuts

In a standard Open / Save file dialog:

- `Ctrl+G`: jump to the folder open in the file manager you used last: Total Commander, Directory Opus or File Explorer. ALTRun looks at the windows from front to back and takes the first of the three, so with several open it is the one you used most recently; the [folder panel](Quick-Switch-Panel) lists all of them
- `Ctrl+Shift+G`: show the folder panel and focus its search box so you can pick a folder or file with the keyboard

The existing file name is preserved when jumping. For example, a name suggested by a Save As dialog is not cleared.

## Automatic switching

Enable **Auto switch** in Preferences → Quick Switch to jump automatically when you switch from Total Commander, Directory Opus or File Explorer to an Open / Save dialog: the dialog goes to the folder of the window you came from.

**Works in**: tick **Standard Open / Save dialogs** (`ahk_class #32770`, on by default; only dialogs with an Explorer file list count, so program windows that are dialogs themselves, such as System Information (msinfo32), are left alone); **Browse for Folder dialogs** (on by default) covers the older folder picker that only shows a folder tree (see below); dialogs that programs draw themselves go in **Other dialogs** (**Edit...**, one per line), for example WPS's `ahk_class Qt5QWindowIcon` (ALTRun types the folder into the dialog's file name box and presses Enter, which works in most Qt, GTK and Java file dialogs). Windows in **Never in** never switch (these lists are edited with **Edit...** too, one window per line); **No auto jump in** disables only automatic switching, and the hotkeys still work there.

## Browse for Folder dialogs

Many programs still pick a folder with the older "Browse for Folder" dialog: a folder tree with OK / Cancel and no file list. The folder panel and `Ctrl+G` work there too: the tree expands to the folder and selects it, and you still click OK yourself (pressing Enter in that dialog would mean OK). Untick **Browse for Folder dialogs** under **Works in** to leave these dialogs alone (`BrowseDialogs`). Like other dialogs, it doesn't work in programs running as administrator.

## Where the folders come from

None of them go through the clipboard, so the paths don't appear in Windows clipboard history (`Win+V`):
- **Total Commander**: asked directly through its `WM_COPYDATA` interface (TC 8.0 or later)
- **Directory Opus**: Opus's own `dopusrt.exe /info …,paths`, which lists the tabs of every Opus window (found next to the running `dopus.exe`, nothing to set up). It runs only while an Opus window is open, when a folder is actually needed (the hotkey, auto switch, or the panel when a dialog opens; the panel shows the other folders first and adds Opus's a moment later), and a result is reused for 1.5 s. With logging on, the log shows how long each read took (`Perf: Directory Opus paths … ms`)
- **File Explorer**: the Shell windows list

## Settings

![Preferences → Quick Switch](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-quickswitch.png)

Preferences → Quick Switch controls the hotkeys, eligible windows and automatic switching. The settings are stored under `Extensions.QuickSwitch`:

| Setting | Default / purpose |
|---|---|
| `Enabled` | `1` = enable Quick Switch |
| `Hotkey` | The jump hotkey, default `Ctrl+G`: the folder of the Total Commander / Directory Opus / File Explorer window used last (`TotalCmdHotkey` before settings version 6) |
| `AutoSwitch` | Jump automatically when switching from Total Commander, Directory Opus or File Explorer to a dialog, default `0` (off) |
| `StandardDialogs` | `1` = in the standard Open / Save dialogs (`ahk_class #32770`) |
| `DialogWindows` | Other dialogs where Quick Switch works (dialogs programs draw themselves), a list of window conditions, default `[]` |
| `BrowseDialogs` | `1` = also in "Browse for Folder" dialogs (folder tree only) |
| `ExcludeWindows` | Windows where Quick Switch is fully disabled (a list) |
| `AutoSwitchExclude` | Windows where automatic switching is disabled but manual hotkeys still work (a list) |

Window conditions are `ahk_exe program.exe`, `ahk_class ClassName` or part of the window title, one per list item, for example `"ExcludeWindows": ["ahk_exe 7zG.exe", "ahk_class SysListView32"]`.

The [Dialog Folder Panel](Quick-Switch-Panel) has its own Preferences page and can appear beside Open / Save dialogs.
