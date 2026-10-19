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

The **Works in** list accepts one window condition per line, for example `ahk_class #32770` for standard Windows dialogs. Use **Never in** to exclude windows entirely. **No auto jump in** disables only automatic switching; manual hotkeys remain available.

The folder is read directly from Total Commander through its `WM_COPYDATA` interface (TC 8.0 or later), not through the clipboard, so the path does not appear in Windows clipboard history (`Win+V`).

## Settings

Preferences → Quick Switch controls the hotkeys, eligible windows and automatic switching. The settings are stored under `Extensions.QuickSwitch`:

| Setting | Default / purpose |
|---|---|
| `Enabled` | `1` = enable Quick Switch |
| `ExplorerHotkey` | Jump to File Explorer's current folder; default `Ctrl+E` |
| `TotalCmdHotkey` | Jump to Total Commander's current folder; default `Ctrl+G` |
| `AutoSwitch` | Automatically jump when switching from Total Commander to a dialog |
| `DialogWindows` | Window conditions where Quick Switch is enabled |
| `ExcludeWindows` | Windows where Quick Switch is fully disabled |
| `AutoSwitchExclude` | Windows where automatic switching is disabled but manual hotkeys still work |

The [Dialog Folder Panel](en-Quick-Switch-Panel) has its own Preferences page and can appear beside Open / Save dialogs.
