**English** · [中文](Quick-Switch)

# Quick Switch

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
