**English** · [中文](Configuration)

# Settings reference (ALTRun.json)

All settings and user data are saved in `Data\ALTRun.json` in the program folder (UTF-8 text; versions up to 2026.09.26 kept it in the program folder itself, and it's moved into `Data\` the first time you start after upgrading). When the program folder isn't writable (for example under `C:\Program Files`), `%APPDATA%\ALTRun\Data\` is used instead; Preferences → Advanced → Settings and data → Change Location... moves it to another folder (for example in a synced folder, see [sharing settings between PCs](en-FAQ#using-the-same-settings-on-several-pcs)). Most settings can be changed in the Preferences window (each has a line of gray help text below it): **OK** saves and closes, **Apply** saves and stays on the current page, **Cancel** discards your changes, **Help** (`F1`) opens the documentation for the current page; settings take effect after ALTRun reloads, which happens automatically. You can also edit the file directly (press `F4` in the search window, or Preferences → Advanced → Edit ALTRun.json...), and **ALTRun reloads automatically when you save it**.

- Missing entries are filled in with their defaults automatically, so you only need to write what you change
- If the file has a syntax error (a missing comma, say), ALTRun tells you, keeps the original renamed to `ALTRun.json.bad` and starts with the default settings; fix it and rename it back
- `SchemaVersion` is the version of the settings format; don't change it by hand. When the format is upgraded, ALTRun converts it automatically and makes a backup

## Structure
```jsonc
{
  "SchemaVersion": 5,
  "General":    { ... },          // general
  "Appearance": { ... },          // appearance
  "Features":   { ... },          // the search features
  "Extensions": { ... },          // extensions
  "Hotkeys":        [ ... ],      // custom hotkeys
  "CustomCommands": [ ... ],      // custom commands
  "Snippets":       [ ... ],      // snippets
  "DataLocation":   "%OneDrive%\\ALTRun"   // optional: use this data folder instead
}
```

`DataLocation` only works in the file at the default location (`Data\ALTRun.json`, or `%APPDATA%\ALTRun\Data\ALTRun.json`): when it's there, ALTRun uses the settings and data in that folder, and the file at the default location only points the way. Normally you set it in Preferences → Advanced → Settings and data → Change Location... rather than by hand; environment variables work, and relative paths start from the program folder.

## General
Split over the "General", "Search Window" and "Hotkeys" (Caps Lock) pages in Preferences.

| Key | Default | Description |
|---|---|---|
| Hotkey | `!Space` | The ALTRun hotkey (AutoHotkey syntax: `!` Alt `^` Ctrl `+` Shift `#` Win; Preferences records the keys directly, see [Setting hotkeys](en-Extensions#setting-hotkeys)) |
| SecondaryHotkey | `!r` | Second hotkey (Alt+R); empty = none |
| DoubleTap | empty | `Ctrl` / `Shift` = pressing this key twice quickly also opens ALTRun (like Listary); empty = off |
| CapsLock | empty | Like on a Mac: `Layout` = pressing Caps Lock switches to the next input method (like `Win+Space`), `Mode` = it switches the Chinese / English mode of the input method; holding it for 0.3 s turns caps lock on or off. Empty = Caps Lock works as usual. See [Caps Lock switches the input method](en-Extensions#caps-lock-switches-the-input-method) |
| CapsLockNotify | 1 | Show "Caps Lock On / Off" when caps lock is turned on or off; 0 = no notice |
| CapsLockElevated | 0 | 1 = Caps Lock switching also works in windows running as administrator: a scheduled task runs a small helper with administrator rights that only handles Caps Lock, while ALTRun itself keeps normal rights. See [Caps Lock switches the input method](en-Extensions#caps-lock-switches-the-input-method) |
| SelectionHotkey | `^!\` | Press it after selecting text / files / a URL to open their actions directly (Ctrl+Alt+\), see [Actions on selected content](en-Usage#actions-on-selected-content); empty = off |
| Language | `auto` | `auto` follows the Windows display language, or a language code: `en` / `zh-CN` (Simplified Chinese) / `zh-TW` (Traditional Chinese) / `ja`; apart from English, each language is a file in `Resources\Lang\` (the old `zh` means `zh-CN`) |
| LaunchAtLogin | 1 | Launch ALTRun at login |
| ShowTrayIcon | 1 | Show the tray icon |
| HideOnDeactivate | 1 | Hide the search window when it loses focus |
| SwitchToEnglishInput | 0 | Switch to English when the search window opens: 0 = off, 1 = English (United States) keyboard (only switches between input methods that are already installed; no keyboard layout is added), 2 = English mode of the current input method (same input method, its Chinese mode turned off) |
| RestoreInput | 1 | With `SwitchToEnglishInput` on, switch back to the previous input method or mode when the window hides (with 1: not needed, and not done, when Windows is set to use a separate input method for each app window) |
| SpaceToRun | 0 | After typing, Space runs the selected result and `Shift+Space` types a space |
| KeepLastQuery | 0 | Keep the last search (input, file search mode and selected row) when the window opens, with the text selected: `Enter` runs it again, typing starts a new search. Taken over from `KeepInput` when upgrading from 2.x |
| ShowTips | 1 | Show a usage tip, a different one each time: in the status bar, or in the empty search box when the status bar is off (`Appearance.StatusBar = 0`) |
| FileManager | `explorer.exe` | Program that opens folders, optionally with parameters, for example `C:\Apps\TotalCMD64\TOTALCMD64.exe /O /T /S` (Total Commander: `/O` reuse the open window, `/T` new tab, `/S` active panel); quote a path with spaces |
| SendToMenu | 1 | Add to Explorer's "Send to" menu |
| StartMenuShortcut | 1 | Add to the Start menu |
| CheckForUpdates | 1 | Check GitHub for new versions in the background at startup and every 6 hours after; a new version is shown in the search window, see [One-key updates](en-Installation#one-key-updates) |
| SaveLog | 0 | Write a debug log (`%Temp%\ALTRun.log`), including how many milliseconds each startup phase and any search over 30 ms took (lines starting with `Perf:`; what you typed is not recorded) |
| HistorySize | 30 | How many recent searches to remember |

## Appearance
| Key | Default | Description |
|---|---|---|
| Theme | `System` | Theme name, see [Themes](en-Themes); `System` follows the Windows light / dark mode |
| Width | 700 | Search window width (pixels, scaled with the display) |
| VisibleRows | 8 | Maximum number of result rows shown (1–9) |
| Placeholder | "" | Gray text in the empty search box; empty = the default "ALTRun Search", only spaces = show nothing (usage tips show there when the status bar is off) |
| StatusBar | 1 | Status bar at the bottom of the search window (shown when there are results: ☰ menu, usage tip, what `Enter` / `→` do); 0 = off, the usage tip shows in the empty search box |
| ShowOn | `Mouse` | Which screen the search window opens on: `Mouse` the one with the mouse pointer / `Primary` the main screen / `Active` the one with the active window |
| Fade | 0 | Fade the search window in and out when showing and hiding (about 120 ms in, up to the theme opacity; about 100 ms out); the window still gets focus at once, so typing is never delayed |
| RememberPosition | 0 | Remember the position after dragging the search window (drag the empty space around the input box) |
| Position | `{"X": 500, "Y": 200}` | The remembered position, in thousandths of the screen's work area: X 0 = far left, 1000 = far right; Y is the window's top edge measured from the top. Default = centered horizontally, 20% from the top |

## Features
Every feature has `Enabled` (1 = on).

### Applications
| Key | Default | Description |
|---|---|---|
| Folders | Start menu (current user / all users), desktop (current user / public) | Indexed folders; [path variables](en-Commands-and-Snippets#variables-in-paths) work |
| FileTypes | `*.lnk` `*.exe` `*.url` `*.appref-ms` | Indexed file types |
| Depth | 3 | Subfolder depth |
| Exclude | `i)(uninstall\|卸载\|readme\|help\|documentation)` | Names matching this regular expression are left out |
| Hidden | `[]` | Paths of apps removed with `Ctrl+Del` in the search results; delete a line to restore one |
| StoreApps | 1 | Include Microsoft Store apps |
| MatchPinyin | 1 | Match Chinese names by their pinyin initials |
| RefreshMinutes | 120 | No longer used (since 2026.10.16.8). The index updates itself by watching the `Folders`: when a shortcut is added, removed or renamed, it is rebuilt silently in the background a few seconds later (at most once a minute), without a notification, pausing while the search window is open; it is also rebuilt after you change `Folders`, `FileTypes`, `Exclude` or `Depth` in Preferences and click Apply, and when the cache is more than an hour old (fallback). Microsoft Store apps aren't in those folders and are rescanned at most once an hour |

### Snippets
| Key | Default | Description |
|---|---|---|
| Keyword | `snip` | Keyword that searches snippets only |
| SearchText | 1 | Also search the text (at least 3 characters typed, every word must appear; matches only in the text come last) |
| PasteMode | `Auto` | `Auto` = type single-line text of up to 100 characters, paste the rest via the clipboard; `Clipboard` = clipboard + Ctrl+V; `Type` = type character by character |
| PasteDelay | 300 | Milliseconds to wait after pasting before restoring the clipboard |
| AutoExpand | 1 | Typing prefix + keyword in any program expands the snippet |
| ExpandPrefix | `;` | Prefix for snippet expansion |
| ExpandExclude | `ahk_exe mstsc.exe, ahk_exe KeePass.exe, ahk_exe KeePassXC.exe` | No expansion in these windows (comma-separated: `ahk_exe`, `ahk_class` or part of the title); never in password boxes |

### Clipboard
| Key | Default | Description |
|---|---|---|
| Keyword | `clip` | Keyword |
| Hotkey | `^!c` | Hotkey that opens clipboard history directly |
| MaxItems | 200 | How many items to keep |
| MaxItemLength | 100000 | Content longer than this many characters isn't recorded |
| Persist | 1 | Save to disk (`ClipboardHistory.json`; items over 4,000 characters and images are stored separately in a `Clipboard\` folder; where they go: see `LocalFiles`); 0 = memory only |
| LocalFiles | 1 | Where clipboard history is saved (Preferences → Clipboard → "Save in"); this PC = `%LOCALAPPDATA%\ALTRun`, not carried along with a synced or portable Data folder. **1** = images and very long items (the `Clipboard\` folder) stay on this PC only: with Data in OneDrive or another synced folder, copying an image doesn't trigger an upload. `ClipboardHistory.json` stays in Data and syncs as usual; images / long items whose files aren't on another PC aren't shown there. **2** = `ClipboardHistory.json` stays on this PC too, so clipboard history never leaves this computer. **0** = everything in the Data folder (for example when you carry it on a USB stick). After a change, the files are moved from the old location at the next start; when both locations have them (say, another PC's history in a synced folder), nothing is overwritten and the new location is used |
| Images | 1 | Also record images (saved as PNG); needs `Persist = 1` |
| MaxImages | 50 | Maximum number of images; the oldest are removed first |
| MergeDoubleCopy | 0 | 1 = pressing `Ctrl+C` twice quickly appends the copied text to the previous item |
| IgnoreApps | KeePass, KeePassXC, 1Password, Bitwarden | Don't record what these programs copy (process names) |

### Calculator
| Key | Default | Description |
|---|---|---|
| StructuralCalc | 0 | Add beam main bar / rebar area results below the result, see [Extensions](en-Extensions#calculator) |
| RebarCover | 40 | Structural results: rebar cover (mm, each side) |
| MaxBarSpacing | 300 | Structural results: maximum main bar spacing (mm) |
| BarSizes | `[13, 16, 20, 25, 32]` | Structural results: bar diameters (mm) listed in the rebar area line |
| BarPrefix | `H` | Structural results: bar designation prefix (`T`, `Y`, `Φ`...) |
| Currency | 0 | Currency conversion (`100 usd to sgd`); downloads exchange rates from frankfurter.dev once a day, see [Unit and currency conversion](en-Extensions#unit-and-currency-conversion) |

### Bookmarks
| Key | Default | Description |
|---|---|---|
| Keyword | `bm` | Keyword that searches bookmarks only |
| InDefaultResults | 1 | Also show matching bookmarks when you type a name directly (after apps and commands, at most 8) |

Also in Preferences → Features → Feature options. Reads the bookmarks of every profile of Chrome, Edge, Brave and Vivaldi; Firefox keeps its bookmarks in an SQLite database and isn't supported yet.

### Scripts
Only `Enabled`. Scripts go in `Scripts\` in the program folder (`%APPDATA%\ALTRun\Scripts\` when the program folder isn't writable); see [Scripts](en-Extensions#scripts) for how to write them.

### Recent (pinned and recent items)
| Key | Default | Description |
|---|---|---|
| RecentCount | 0 | How many recently opened items the empty search box shows; 0 = none (only pinned items by default) |
| Pinned | `[]` | Items pinned to the empty search box (pin / unpin them in the action panel, no need to edit by hand) |

Recently opened items are kept in `Data\Knowledge.json`, see [Pinned and recent items](en-Usage#pinned-and-recent-items).

### Windows (window switcher)
| Key | Default | Description |
|---|---|---|
| Keyword | `w` | Keyword that searches open windows only; `w ` alone lists all windows |
| InDefaultResults | 1 | Also show windows whose title matches well when you type directly (after apps and commands, at most 3) |

Also in Preferences → Features → Feature options. Only windows you can see on the taskbar are listed (no tool windows, dialogs, windows on other virtual desktops or ALTRun's own windows).

### WebSearch
| Key | Description |
|---|---|
| Engines | List of search engines, each `{ "Id", "Keyword", "Title", "Url", "Icon" }`; `Keyword` can hold several keywords separated by commas (for example `"g, ㅎ"`). `{query}` in `Url` is replaced by what you typed. `{query|8000}` gives a default value: typing only the keyword opens the URL with 8000 (for example `http://localhost:{query|8000}`). `Icon` (optional): an `.ico`, `.png` or `.exe` file for the result icon, a relative path starts from the Data folder (for example `Icons\jira.png`). In the editor, **Download Site Icon** visits the website in the URL once, saves its icon in `Data\Icons` and fills this in. Empty = the built-in icon (the default engines have their own) or the browser icon |
| Fallbacks | Fallback items shown when nothing matches: engine Ids, or `files` (file search). Default `["google", "files", "bing"]` |

### FileSearch
See [File search](en-File-Search#settings).

### Terminal
| Key | Default | Description |
|---|---|---|
| Prefix | `>` | Prefix |
| Shell | `cmd` | `cmd` / `powershell` / `pwsh` / `wt` |

### Help (cheat sheet)
| Key | Default | Description |
|---|---|---|
| Enabled | 1 | Typing `?` shows every search syntax and shortcut |

### System (system commands)
| Key | Default | Description |
|---|---|---|
| ConfirmActions | 1 | Confirm before shutting down, restarting, logging off or emptying the Recycle Bin |
| SettingsPages | 1 | Also search Windows Settings pages (more than 40, such as Display, Bluetooth, Wi-Fi, Default apps and Windows Update, opening the matching `ms-settings:` page) |
| Hidden | `[]` | Ids of built-in commands removed with `Ctrl+Del` in the search results (for example `"Printers"`); delete an entry to restore it |

## Extensions
See [Extensions](en-Extensions).

| Node | Keys |
|---|---|
| QuickSwitch | `Enabled` `ExplorerHotkey` (`^e`) `TotalCmdHotkey` (`^g`) `MenuHotkey` (`^+g`, moves the cursor to the panel's search box) `RecentFolders` (10) `ShowPanel` (1, the folder panel below dialogs) `PanelSearch` (`folders` = folders only, `all` = folders and files) `AutoSwitch` `DialogWindows` (`ahk_class #32770`) `ExcludeWindows` `AutoSwitchExclude` |
| AutoDate | `Enabled` `DateFormat` (`dd.MM.yyyy`) `RenameHotkey` (`^d`) `RenameWindows` `AppendHotkey` (`^d`) `AppendWindows` |
| PTTools | Inputs and window positions saved by PT Tools itself |

## Lists
```jsonc
"Hotkeys": [
  { "Key": "~MButton", "Action": "PTTools", "WinTitle": "ahk_exe RAPTW.exe" }
],
"CustomCommands": [
  { "Title": "Desktop", "Type": "Folder", "Target": "A_Desktop", "Arguments": "", "Keyword": "" },
  { "Title": "IP Configuration", "Type": "Command", "Target": "cmd.exe", "Arguments": "/k ipconfig /all", "Keyword": "ipconfig" }
],
"Snippets": [
  { "Name": "Today's date", "Keyword": "today", "Text": "{date}", "AutoExpand": 1 }
]
```
For the fields, see [Custom commands and snippets](en-Commands-and-Snippets) and [Extensions](en-Extensions#custom-hotkeys).

## The Data folder
Created at runtime; deleting files only makes them be rebuilt (the learned ranking and clipboard history are lost):

| File | Contents |
|---|---|
| `AppIndex.json` | App index |
| `FileIndex.json` | Built-in file index (without Everything) |
| `Knowledge.json` | Learned ranking and recent searches |
| `Usage.json` | Usage statistics (how often each feature was used each day) |
| `ClipboardHistory.json` | Clipboard history (very long items and images are stored as separate files, by default on this PC in `%LOCALAPPDATA%\ALTRun\Clipboard\`; in `Data\Clipboard\` with `LocalFiles = 0`; with `LocalFiles = 2` this file is on this PC too, not in Data) |
| `Currency.json` | Exchange rates for currency conversion (only once currency conversion is on) |
| `Update.json` | Time of the last update check, skipped versions |
| `Icons\` | Web search icons saved by "Download Site Icon" |

The icon cache is not in Data: result icons that have been read are kept on this PC in `%LOCALAPPDATA%\ALTRun\IconCache\Icons-<size>.bin`, so icons show at once after a restart; each PC has its own, and it is rebuilt automatically if deleted.
