**English** · [中文](Home)

# ALTRun documentation

ALTRun is a launcher for Windows modeled on [Alfred](https://www.alfredapp.com/) for macOS. It is written in AutoHotkey v2, free, open source and portable. Press `Alt+Space`, type a name and press `Enter`. Website: https://zhugecaomao.github.io/ALTRun/

- **Small**: about 2.5 MB: one exe plus a few resource files, with no .NET, Electron or other runtime
- **Fast**: results as you type; with [Everything](https://www.voidtools.com/) running, every file on your disks is searched in milliseconds
- **Learns what you use**: remembers your picks and puts frequent ones first; matches word initials (`vsc`) and pinyin initials (`wx` → 微信)
- **One window for everyday tasks**: calculator, web search, clipboard history, snippet expansion, system commands, terminal
- **Portable, data is local by default**: no registry entries, no admin rights, no background service; settings default to `Data\`, use `%APPDATA%\ALTRun\Data\` when the program folder isn't writable, and can be moved in Preferences → Advanced; one-key updates keep your settings
- **English / Simplified Chinese / Traditional Chinese / Japanese interface**, 22 built-in themes (each in three densities), and you can make your own

> This documentation covers version **3.0** and later. 3.0 redesigned the search window and the settings file; 2.x settings are converted automatically the first time you start it, see [Installation and upgrades](en-Installation).

![ALTRun: type to search, open the action panel, convert units and search files](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/demo.png)

## Start here
| Page | Contents |
|---|---|
| [Getting started](en-Getting-Started) | First steps: open the window, launch programs, everyday features, your own commands |
| [Installation and upgrades](en-Installation) | Download, Scoop / winget, start with Windows, one-key updates, upgrading from 2.x, uninstalling |
| [Search and shortcuts](en-Usage) | Every search syntax, shortcuts, action panel, learned ranking, usage statistics |
| [Custom commands and snippets](en-Commands-and-Snippets) | Command types, path variables, snippet placeholders, snippet expansion, clipboard history |
| [File search](en-File-Search) | Everything integration, built-in index, exclusion rules |
| [Themes](en-Themes) | 22 built-in themes (or follow the Windows light / dark mode), density, custom themes, every available key |
| [Browser bookmarks](en-Bookmarks) | Search, open and copy bookmark URLs |
| [Quick Switch](en-Quick-Switch) | Jump to a folder open in File Explorer or Total Commander |
| [Dialog Folder Panel](en-Quick-Switch-Panel) | Search and pick folders in an Open / Save dialog |
| [Date Stamp](en-Date-Stamp) | Add dates when renaming files or at the end of text |
| [Calculator](en-Calculator) | Expressions, unit and currency conversion, structural calculations |
| [Scripts](en-Scripts) | Turn your own scripts into commands in the search results |
| [Custom hotkeys](en-Hotkeys) | Assign system commands to hotkeys, Caps Lock input switching, how to set hotkeys |
| [Extensions](en-Extensions) | System commands, terminal, PT Tools, and links to the extension pages |
| [Settings reference](en-Configuration) | Every ALTRun.json setting and its default value |
| [Data files](en-Data-Files) | Where the data folder is, what each file is, backups and moving to another PC |
| [Privacy and network access](en-Privacy) | When ALTRun goes online, where to, and what is stored on your PC |
| [FAQ](en-FAQ) | Hotkey conflicts, things that can't be found, antivirus false positives, privacy... |
| [Development guide](en-Development) | Architecture, adding a search feature, code style, tests |

## Get going in 30 seconds
| To do this | Type |
|---|---|
| Open a program | `word` → `Enter` |
| Open a file / folder | `Space`, then `report`; or `'report` |
| Calculate | `12*(3+4)` → `Enter` copies the result |
| Search the web | `g keywords` / `bd keywords` |
| Clipboard history | `Ctrl+Alt+C` or `clip` |
| Run a command | `>ipconfig /all` |
| Lock / shut down | `lock` / `shutdown` |
| Paste text you use often | type `;sig` in any program to expand your signature |
| Jump an Open / Save dialog to a folder | click it in the [Dialog Folder Panel](en-Quick-Switch-Panel) below the dialog, or `Ctrl+G` for the current Total Commander folder |
| Edit the selected command | `F3` |
| More actions | `→` or right-click |
| Forgot the syntax | `?` lists every search syntax and shortcut |

Questions or ideas: [Issues](https://github.com/zhugecaomao/ALTRun/issues) · [Discussions](https://github.com/zhugecaomao/ALTRun/discussions)
