**English** · [中文](Getting-Started-zh)

# Getting started in 5 minutes

## 1. Install and start
Download `ALTRun_v<version>.zip` from [GitHub Releases](https://github.com/zhugecaomao/ALTRun/releases/latest), unzip it to any folder (for example `D:\Tools\ALTRun`) and run `ALTRun.exe`. Nothing to install and no administrator rights needed; for Scoop / winget see [Installation and upgrades](Installation).

The first run opens the search window and puts the ALTRun icon in the notification area. ALTRun starts with Windows by default (turn it off in Preferences → General).

## 2. Open the search window
Press `Alt+Space` (or `Alt+R`). Press it again, press `Esc` or click elsewhere to hide it. Change the hotkey in Preferences → General; double-tapping `Ctrl` works too if you turn it on.

## 3. Open programs and files
- Type part of a program's name, for example `note` for Notepad; word initials work too (`vsc` → Visual Studio Code), as do pinyin initials for Chinese names (`wx` → 微信)
- `Enter` opens the selected item, `↑` `↓` select, `Ctrl+1` to `Ctrl+9` open that item directly
- Press `→` (or right-click) on a result for the action panel: run as administrator, show in the file manager, copy the path...
- Start with `'` to search only files and folders, for example `'report`; with [Everything](https://www.voidtools.com/) running it searches the whole disk, otherwise Desktop, Documents and Downloads, see [File search](File-Search)

ALTRun remembers what you pick, so the next time you type the same thing it comes first.

## 4. Try a few everyday features
| Type | Result |
|---|---|
| `12*(3+4)` | Calculation; `Enter` copies the result ([Calculator](Calculator)) |
| `10 km in mi` | Unit conversion |
| `g weather` | Search Google ([web search](Usage#search-syntax)) |
| `clip` or `Ctrl+Alt+C` | Clipboard history ([clipboard history](Clipboard)) |
| `lock` | Lock the screen and other system commands ([system commands](System-Commands#system-commands)) |
| `>ipconfig` | Run a command in the terminal |
| `?` | Cheat sheet: every search syntax and shortcut |

## 5. Add your own command
Turn folders, files and web addresses you use often into custom commands, then open them by name or keyword:
- In Explorer, right-click → Send to → ALTRun, adjust the name and press `Enter`
- Or Preferences → Custom Commands → Add

Save text you type often (signatures, addresses) as snippets and type `;keyword` in any program to expand it, see [Snippets](Snippets).

## 6. Preferences
Right-click the notification area icon → Preferences, or press `F2` / `Ctrl+,` in the search window. `F1` (or Help) on any page opens its documentation. Appearance has 22 themes and three densities, see [Themes](Themes).

## Next steps
- [Search and shortcuts](Usage): every search syntax, shortcut and the action panel
- [FAQ](FAQ): hotkey does nothing, things can't be found, antivirus false positives
- [Privacy and network access](Privacy): when ALTRun goes online and where your data lives
