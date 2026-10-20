**English** · [中文](Quick-Switch)

# Quick Switch

Quick Switch jumps an Open / Save dialog to a folder you already have open, without browsing through the directory tree again. For the visual folder panel and its settings, see [Dialog Folder Panel](en-Quick-Switch-Panel).

## Shortcuts

In a standard Open / Save file dialog:

- `Ctrl+G`: jump to the folder currently open in Total Commander
- `Ctrl+E`: jump to the folder currently open in File Explorer
- `Ctrl+Shift+G`: show the folder panel and focus its search box so you can pick a folder or file with the keyboard

The existing file name is preserved when jumping. For example, a name suggested by a Save As dialog is not cleared.

## Automatic switching

Enable **Auto switch** in Preferences → Quick Switch to jump automatically to Total Commander's current folder when switching from Total Commander to an Open / Save dialog.

**Works in**: tick **Standard Open / Save dialogs** (`ahk_class #32770`, on by default; only dialogs with an Explorer file list count, so program windows that are dialogs themselves, such as System Information (msinfo32), are left alone); dialogs that programs draw themselves go in **Other dialogs**, one per line, for example WPS's `ahk_class Qt5QWindowIcon`. Windows in **Never in** never switch; **No auto jump in** disables only automatic switching, and the hotkeys still work there.

The folder is read directly from Total Commander through its `WM_COPYDATA` interface (TC 8.0 or later), not through the clipboard, so the path does not appear in Windows clipboard history (`Win+V`).

## Settings

![Preferences → Quick Switch](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-quickswitch.png)

Preferences → Quick Switch controls the hotkeys, eligible windows and automatic switching. The settings are stored under `Extensions.QuickSwitch`:

| Setting | Default / purpose |
|---|---|
| `Enabled` | `1` = enable Quick Switch |
| `ExplorerHotkey` | Jump to File Explorer's current folder; default `Ctrl+E` |
| `TotalCmdHotkey` | Jump to Total Commander's current folder; default `Ctrl+G` |
| `AutoSwitch` | Jump automatically when switching from Total Commander to a dialog, default `0` (off) |
| `DialogWindows` | Window conditions where Quick Switch is enabled |
| `ExcludeWindows` | Windows where Quick Switch is fully disabled |
| `AutoSwitchExclude` | Windows where automatic switching is disabled but manual hotkeys still work |

The [Dialog Folder Panel](en-Quick-Switch-Panel) has its own Preferences page and can appear beside Open / Save dialogs.
