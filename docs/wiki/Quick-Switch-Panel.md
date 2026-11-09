**English** · [中文](Quick-Switch-Panel-zh)

# Dialog Folder Panel

The folder panel is the visual way to use Quick Switch. It is enabled by default. When an Open / Save dialog (or a [Browse for Folder](Quick-Switch#browse-for-folder-dialogs) dialog) appears, the panel attaches below it, or above it if there is not enough room. Click a folder to jump there without remembering a shortcut.

![Quick Switch folder panel below an Open dialog](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/quickswitch.png)

## Contents and controls

- Lists the current and opposite panels of each Total Commander window, the current tab (and the other side) of each Directory Opus window, open File Explorer windows and recently used folders. Your [favorite folders](#favorite-folders) are in the star menu at the top right.
- The search box filters the list first, then searches matching folder and file names with Everything (or ALTRun's built-in file index); folders are listed before files.
- Selecting a file jumps to its containing folder; it does not open or save the file.
- Use `↑` / `↓` to select, `Enter` to jump and `Esc` to return to the original dialog. With a Chinese input method, the first `Enter` may commit the composition; press it again to jump.
- The panel hides when the dialog is not in the foreground and uses the current ALTRun theme. It is as wide as the dialog (at least 420 pixels at 100% scaling, at most the width of the screen), and follows when the dialog is resized.

In a standard Open / Save dialog, `Ctrl+Shift+G` temporarily shows the panel and focuses its search box. This works even when automatic display is disabled.

## Favorite folders

Folders you often save to or open from are in the star menu at the top right of the panel:

![The panel's star menu with a favorite's subfolders](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/quickswitch-favorites.png)

- **Jump**: click the star (or press `Alt+↓` or `F4` in the panel's search box, as for a drop-down list), then a favorite. A favorite with subfolders opens a submenu with its subfolders (one level); the favorite itself is the first, bold item, so a click always jumps. The first 30 subfolders by name are listed, sorted as in Explorer ("2" before "10"); hidden and system folders are left out. At most 500 are read, so a favorite with thousands of subfolders opens at once ("470+ more folders").
- **Search**: the panel's search box also finds favorites, listed first and marked "Favorite".
- **Add**: star menu → **Add Current Folder "…" to Favorites** for the folder the dialog shows now, or right-click a folder in the panel → **Add to Favorites**. A short notice confirms it.
- **Keyboard**: `Ctrl+D` adds the dialog's current folder while the caret is in the panel (click the panel, or press `Ctrl+Shift+G`, first). In the dialog itself `Ctrl+D` is the [date stamp](Date-Stamp) hotkey and adds a date to the file name.
- **Remove**: right-click a favorite in the panel (open in a file manager, or found by the search) → **Remove from Favorites**.
- **Edit the list**: star menu → **Edit Favorites...**, or Preferences → Dialog Panel → "Favorite folders" → **Edit...**, one folder per line; environment variables such as `%USERPROFILE%\Documents` work.

Favorites are listed in the order they were added and are saved in `ALTRun.json`. Subfolders are read when the menu opens, and only on this PC's own drives: for a favorite on a network share, a mapped network drive or a CD, the menu shows the favorite without subfolders, so a disconnected drive can't make the menu hang.

## Search scope

Preferences → Dialog Panel → **Search box finds** controls the scope:

- **Folders only** (default): only matching folders are shown.
- **Folders and files**: search both; folders remain grouped first so files with the same name do not push them out.

Recent folders come from Windows' **Recent items**. For a file that was opened, its containing folder is used. The default is 10 recent folders; set it to 0 to hide them.

## Settings

![Preferences → Dialog Panel](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-dialogpanel.png)

Change these in Preferences → Dialog Panel. They are stored under `Extensions.QuickSwitch`:

| Setting | Default / values | Purpose |
|---|---|---|
| `ShowPanel` | `1` | Automatically show the folder panel |
| `PanelSearch` | `folders` / `all` | Search folders only, or both folders and files |
| `Favorites` | `[]` | Favorite folders, in the star menu (see [Favorite folders](#favorite-folders)) |
| `RecentFolders` | `10` | Number of recent folders; `0` = none |
| `MenuHotkey` | `Ctrl+Shift+G` | Show the panel and focus its search box |

For the other shortcuts and automatic switching settings, see [Quick Switch](Quick-Switch).
