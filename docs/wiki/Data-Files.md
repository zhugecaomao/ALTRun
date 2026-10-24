**English** · [中文](Data-Files-zh)

# Data files

## Where the data folder is
| Situation | Location |
|---|---|
| Default (the program folder is writable, portable) | `Data\` in the program folder |
| The program folder isn't writable (for example under `C:\Program Files`) | `%APPDATA%\ALTRun\Data\` |
| You picked another folder in Preferences → Advanced → Data → Change Location | That folder (the `ALTRun.json` in the default location only keeps a `DataLocation` entry pointing there), see [Using the same settings on several PCs](FAQ#using-the-same-settings-on-several-pcs) |

Preferences → Advanced → Data → "Open Data Folder" opens it.

## Files in the Data folder
| File | Contents | If you delete it |
|---|---|---|
| `ALTRun.json` | All settings, custom commands, snippets, hotkeys and web searches, see [Settings reference](Configuration) | Settings go back to defaults and your commands and snippets are gone (**this is the file to back up**) |
| `ALTRun.json.bak` | Backup taken by Preferences → Advanced → "Restore Defaults" | No effect |
| `ALTRun.v<version>.backup.json` | Backup taken before the settings format was upgraded (for example `ALTRun.v4.backup.json`) | No effect; use it to go back to an older version |
| `ALTRun.json.bad` | When the settings file has a format error, the original is kept under this name; fix it and rename it back to `ALTRun.json` | No effect |
| `AppIndex.json` | App index | Rescanned at startup |
| `FileIndex.json` | Built-in file index (used without Everything) | Rescanned in the background |
| `Knowledge.json` | Learned ranking, search history, recent items (pinned items are in `ALTRun.json`) | Ranking has to be learned again, recent items are cleared |
| `Usage.json` | Usage statistics (how often each feature was used each day) | Statistics start from zero |
| `ClipboardHistory.json` | Clipboard history (text) | History is cleared |
| `Clipboard\` | Images and very long items of the clipboard history (only here when "Save in" is set to keep everything in the Data folder) | Those items no longer show in the history |
| `Currency.json` | Exchange rates for currency conversion (only once it is turned on) | Downloaded again at the next start |
| `Update.json` | Time of the last update check, skipped versions | Checked again at the next start |
| `Icons\` | Web search icons saved by "Download Site Icon" | Those engines use the default icon |

Apart from `ALTRun.json`, every file is created at runtime and simply rebuilt if deleted.

## Files outside the Data folder
| Location | Contents |
|---|---|
| `%LOCALAPPDATA%\ALTRun\IconCache\` | Icon cache (`Icons-<size>.bin`) so icons show at once after a restart; one per PC, rebuilt automatically if deleted |
| `%LOCALAPPDATA%\ALTRun\Clipboard\` | Images and very long clipboard items (kept here by default so they don't follow a sync folder); with "everything on this PC only", `ClipboardHistory.json` is in `%LOCALAPPDATA%\ALTRun\` too |
| `Themes\` in the program folder (`%APPDATA%\ALTRun\Themes\` if it isn't writable) | Your own themes, see [Themes](Themes#custom-themes) |
| `Scripts\` in the program folder (`%APPDATA%\ALTRun\Scripts\` if it isn't writable) | Your own scripts, see [Scripts](Scripts) |
| `%Temp%\ALTRun.log` | Debug log, written only with Preferences → Advanced → "Write a debug log" on |

For the startup, Start menu and "Send to" shortcuts and the administrator helper's scheduled task, see [Uninstalling](Installation#uninstalling).

## Backups and moving to another PC
- **One file is enough**: `ALTRun.json` holds every setting, custom command and snippet. To take the learned ranking and clipboard history along too, back up the whole Data folder
- **New PC**: copy the whole program folder (including `Data\`); use [path variables](Custom-Commands#variables-in-paths) in paths so they still work with a different user name
- **Syncing several PCs**: move the data folder into a sync folder, see [Using the same settings on several PCs](FAQ#using-the-same-settings-on-several-pcs). Don't change settings on two PCs with ALTRun open at the same time, or the sync service may create conflict copies
- **Snippets only**: Preferences → Snippets → Import / Export saves them as CSV, see [Snippets](Snippets#importing-and-exporting)
