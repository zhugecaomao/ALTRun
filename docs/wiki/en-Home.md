**English** · [中文](Home)

# ALTRun documentation

ALTRun is a launcher for Windows modeled on [Alfred](https://www.alfredapp.com/) for macOS. It is written in AutoHotkey v2, free, open source and portable. Press `Alt+Space`, type a name and press `Enter`. Website: https://zhugecaomao.github.io/ALTRun/

- **Small**: under 1 MB to download and about 2 MB unpacked: one exe plus a few resource files, with no .NET, Electron or other runtime
- **Fast**: results as you type; with [Everything](https://www.voidtools.com/) running, every file on your disks is searched in milliseconds
- **Learns what you use**: remembers your picks and puts frequent ones first; matches word initials (`vsc`) and pinyin initials (`wx` → 微信)
- **One window for everyday tasks**: calculator, web search, clipboard history, snippet expansion, system commands, terminal
- **Portable, data stays on your PC**: no registry entries, no admin rights, no background service; all settings live in the `Data\` folder; one-key updates keep your settings
- **English / Simplified Chinese / Traditional Chinese / Japanese interface**, 20 built-in themes, and you can make your own

> This documentation covers version **3.0** and later. 3.0 redesigned the search window and the settings file; 2.x settings are converted automatically the first time you start it, see [Installation and upgrades](en-Installation).

![ALTRun: type to search, open the action panel, convert units and search files](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/demo.png)

## Start here
| Page | Contents |
|---|---|
| [Installation and upgrades](en-Installation) | Download, Scoop / winget, start with Windows, one-key updates, upgrading from 2.x, uninstalling |
| [Search and shortcuts](en-Usage) | Every search syntax, shortcuts, action panel, learned ranking, usage statistics |
| [Custom commands and snippets](en-Commands-and-Snippets) | Command types, path variables, snippet placeholders, snippet expansion, clipboard history |
| [File search](en-File-Search) | Everything integration, built-in index, exclusion rules |
| [Themes](en-Themes) | 20 built-in themes (or follow the Windows light / dark mode), custom themes, every available key |
| [Extensions](en-Extensions) | Quick Switch, date stamp, custom hotkeys, list of system commands, PT Tools |
| [Settings reference](en-Configuration) | Every ALTRun.json setting and its default value |
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
| Jump an Open / Save dialog to a folder | click it in the [Quick Switch](en-Extensions#quick-switch) panel below the dialog, or `Ctrl+G` for the current Total Commander folder |
| Edit the selected command | `F3` |
| More actions | `→` or right-click |
| Forgot the syntax | `?` lists every search syntax and shortcut |

Questions or ideas: [Issues](https://github.com/zhugecaomao/ALTRun/issues) · [Discussions](https://github.com/zhugecaomao/ALTRun/discussions)
