<h1 align="center"><img width="48" alt="ALTRun" src="docs/images/logo.png" /> ALTRun</h1>

<p align="center"><b>English</b> | <a href="README.zh-CN.md">简体中文</a></p>

<p align="center">
  <b>A lightweight, open-source launcher for Windows, modeled on <a href="https://www.alfredapp.com/">Alfred</a> for macOS</b><br>
  Press <code>Alt+Space</code>, type a few letters, press <code>Enter</code>: apps, files, web search, calculations, clipboard and system commands, all from one window
</p>

<p align="center">
  <a href="https://github.com/zhugecaomao/ALTRun/releases/latest"><img alt="Release" src="https://img.shields.io/github/v/release/zhugecaomao/ALTRun?label=release"></a>
  <a href="https://github.com/zhugecaomao/ALTRun/releases"><img alt="Downloads" src="https://img.shields.io/github/downloads/zhugecaomao/ALTRun/total"></a>
  <a href="https://github.com/zhugecaomao/ALTRun/actions/workflows/tests.yml"><img alt="Tests" src="https://github.com/zhugecaomao/ALTRun/actions/workflows/tests.yml/badge.svg?branch=main"></a>
  <a href="https://www.autohotkey.com/"><img alt="AutoHotkey v2" src="https://img.shields.io/badge/AutoHotkey-v2.0-334455?logo=autohotkey"></a>
  <img alt="Windows" src="https://img.shields.io/badge/platform-Windows%2010%20%7C%2011-0078D6">
  <a href="LICENSE"><img alt="License: GPL-3.0" src="https://img.shields.io/github/license/zhugecaomao/ALTRun"></a>
</p>

<p align="center">
  <a href="https://zhugecaomao.github.io/ALTRun/">Website</a> ·
  <a href="#why-altrun">Why ALTRun</a> ·
  <a href="#getting-started">Getting started</a> ·
  <a href="#features">Features</a> ·
  <a href="#keyboard-shortcuts">Shortcuts</a> ·
  <a href="https://github.com/zhugecaomao/ALTRun/wiki/en-Home">Documentation (Wiki)</a> ·
  <a href="CHANGELOG.md">Changelog</a>
</p>

<p align="center">
  <img src="docs/images/screenshots/demo.png" width="700" alt="ALTRun: type to search, open the action panel, convert units and search files">
</p>


## Why ALTRun
- **Tiny and fast**: under 1 MB to download, no .NET, Electron or other runtime. Results appear from the first character; with [Everything](https://www.voidtools.com/) running it searches every file on your disks in milliseconds, otherwise a built-in index is used.
- **Smart ranking**: learns from what you pick; matches word initials (`vsc` → Visual Studio Code) and pinyin initials (`wx` → 微信).
- **More than a launcher**: calculator and unit conversion, clipboard history, text snippets that expand anywhere, web search, bookmarks, window switching and system commands, so you don't need a handful of separate little tools.
- **Keyboard first**: `→` opens the action panel for any result, `F3` edits in place, `Ctrl+1`–`Ctrl+9` open a result directly.
- **Portable and private**: no installer, no registry entries, no admin rights, no background service; all data stays in the `Data\` folder. No telemetry: it only goes online for updates (announced in the search window, installed with `Enter`, settings kept).
- **Yours to shape**: English, 简体中文, 繁體中文 and 日本語; 20 built-in themes or follow the Windows light / dark mode. Free and open source (GPL-3.0), with more than 7,000 tests run on Windows for every pull request.


## Getting started
1. Download the [latest release](https://github.com/zhugecaomao/ALTRun/releases/latest), unzip it to any folder and run `ALTRun.exe` (no installation, no AutoHotkey needed). If SmartScreen says "Windows protected your PC", click **More info** → **Run anyway**; for antivirus warnings see the [FAQ](https://github.com/zhugecaomao/ALTRun/wiki/en-FAQ#my-antivirus-flags-it).
2. Press `Alt+Space` (or `Alt+R`) to open the search window, type a name and press `Enter`.
3. Type `?` to see every syntax and shortcut; press `Ctrl+,` to open Preferences.

Or install with [Scoop](https://scoop.sh/): `scoop bucket add altrun https://github.com/zhugecaomao/ALTRun` then `scoop install altrun`. With [AutoHotkey v2](https://www.autohotkey.com/) installed, you can also run `ALTRun.ahk` from the source. Upgrading from 2.x, starting with Windows and more: [Installation and upgrades](https://github.com/zhugecaomao/ALTRun/wiki/en-Installation).


## Features
**Search**
- **Apps, files and folders**: Start menu, desktop and Microsoft Store apps are indexed automatically. Press `Space` in an empty search box to search files (Everything or the built-in index), `doc report` to search by type, or type a path such as `D:\Projects\` to browse.
- **Custom commands**: files, folders, programs with arguments and web addresses, each with an optional keyword; add them from Explorer with right-click → Send to → ALTRun.
- **Calculator**: unit conversion (`10 km in mi`), number bases (`255 in hex`), date arithmetic (`today + 30 days`) and optional currency conversion (`100 usd to sgd`).
- **Web search and bookmarks**: `g keywords` (Google), `bd keywords` (Baidu) and your own engines; Chrome, Edge, Brave and Vivaldi bookmarks.
- **Windows and system commands**: switch to an open window (`w keywords`); lock, sleep, shut down, volume and media control, 40+ Windows Settings pages, text conversions.
- **Pinned and recent items** show before you type; usage statistics count what you use (counts only).

**Productivity**
- **Action panel** (`→` or right-click): run as administrator, open file location, copy path, open a terminal here, properties and more.
- **Actions on selected content**: select text, files or a URL in any program and press `Ctrl+Alt+\` to search, save as a snippet, or convert and replace the text.
- **Clipboard history** (`Ctrl+Alt+C`): text, files and images; pin items, paste as plain text, password managers ignored.
- **Text snippets**: placeholders such as `{date}`, `{clipboard}` and `{cursor}`; type `;keyword` in any program to expand.
- **Also**: run terminal commands (`>ipconfig /all`), show a result in large type (`Ctrl+L`).

**Extensions**
- **Quick Switch**: in an Open / Save dialog, `Ctrl+G` / `Ctrl+E` jump to the folder open in Total Commander / Explorer; a folder panel lists every open and recent folder.
- **Date stamp**: while renaming a file, `Ctrl+D` adds or updates a date before the extension.
- **Custom hotkeys** for any system command, optionally only in specific programs.
- **Scripts**: drop `.ahk`, `.ps1`, `.bat` or `.py` files into `Scripts\` to run them from the search window, with arguments and output.
- **For structural engineers**: beam rebar calculations in the calculator (off by default) and PT Tools (rebar / BRC mesh areas, SPF2M tendon profiles).


## Screenshots
| Search window | Action panel (`→`) |
|:---:|:---:|
| <img src="docs/images/screenshots/search.png" alt="Search window"> | <img src="docs/images/screenshots/actions.png" alt="Action panel"> |
| **File search (`Space` + name)** | **Calculator** |
| <img src="docs/images/screenshots/files.png" alt="File search"> | <img src="docs/images/screenshots/calculator.png" alt="Calculator"> |
| **Pinyin initials (`jsb` finds 记事本, Notepad)** | **Clipboard history (`clip`)** |
| <img src="docs/images/screenshots/pinyin.png" alt="Pinyin initials"> | <img src="docs/images/screenshots/clipboard.png" alt="Clipboard history"> |
| **Quick Switch (folder panel below an Open / Save dialog)** | **Themes (Preferences → Appearance)** |
| <img src="docs/images/screenshots/quickswitch.png" alt="Quick Switch"> | <img src="docs/images/screenshots/prefs-appearance.png" alt="Themes"> |

<details>
<summary><b>Built-in themes</b> (20 in total, click to expand)</summary>

| Dark | Classic | Midnight |
|:---:|:---:|:---:|
| <img src="docs/images/screenshots/theme-dark.png" alt="Dark"> | <img src="docs/images/screenshots/theme-classic.png" alt="Classic"> | <img src="docs/images/screenshots/theme-midnight.png" alt="Midnight"> |
| **Frost** | **Graphite** | **Ocean** |
| <img src="docs/images/screenshots/theme-frost.png" alt="Frost"> | <img src="docs/images/screenshots/theme-graphite.png" alt="Graphite"> | <img src="docs/images/screenshots/theme-ocean.png" alt="Ocean"> |
| **Paper** | **DarkCompact** | **LightCompact** |
| <img src="docs/images/screenshots/theme-paper.png" alt="Paper"> | <img src="docs/images/screenshots/theme-darkcompact.png" alt="DarkCompact"> | <img src="docs/images/screenshots/theme-lightcompact.png" alt="LightCompact"> |
| **TokyoNight** | **Dracula** | **CatppuccinMocha** |
| <img src="docs/images/screenshots/theme-tokyonight.png" alt="TokyoNight"> | <img src="docs/images/screenshots/theme-dracula.png" alt="Dracula"> | <img src="docs/images/screenshots/theme-catppuccinmocha.png" alt="CatppuccinMocha"> |
| **GruvboxDark** | **SolarizedLight** | **MidnightCompact** |
| <img src="docs/images/screenshots/theme-gruvboxdark.png" alt="GruvboxDark"> | <img src="docs/images/screenshots/theme-solarizedlight.png" alt="SolarizedLight"> | <img src="docs/images/screenshots/theme-midnightcompact.png" alt="MidnightCompact"> |
| **DarkModern** | **LightModern** | **Monokai** |
| <img src="docs/images/screenshots/theme-darkmodern.png" alt="DarkModern"> | <img src="docs/images/screenshots/theme-lightmodern.png" alt="LightModern"> | <img src="docs/images/screenshots/theme-monokai.png" alt="Monokai"> |
| **OneDark** | | |
| <img src="docs/images/screenshots/theme-onedark.png" alt="OneDark"> | | |

</details>

The screenshots are generated automatically by [Tests/Screenshots](Tests/Screenshots/TakeScreenshots.ahk) on a Windows runner in GitHub Actions.


## Keyboard shortcuts
| Key | Action |
|---|---|
| `Alt+Space` | Show / hide the search window (configurable) |
| `Enter` / `Ctrl+1`–`Ctrl+9` | Open the selected / Nth result |
| `Ctrl+Enter` | Files: open the containing folder; text: paste into the active window |
| `Alt+Enter` / `Ctrl+C` | Copy the path, URL or text |
| `→` / right-click | Action panel / action menu |
| `Space` (empty search box) | Switch to file search |
| `Tab` | Autocomplete; folders: browse into them |
| `F3` / `Ctrl+Del` | Edit / delete (apps: hide) the selected result |
| `Ctrl+,` / `F1` | Preferences / About ALTRun |
| `?` | Cheat sheet: every syntax and shortcut |

More in [Search and shortcuts](https://github.com/zhugecaomao/ALTRun/wiki/en-Usage).


## Documentation
The full documentation is in the [wiki](https://github.com/zhugecaomao/ALTRun/wiki/en-Home):

| Page | Contents |
|---|---|
| [Installation and upgrades](https://github.com/zhugecaomao/ALTRun/wiki/en-Installation) | Download, Scoop / winget, start with Windows, automatic updates, upgrading from 2.x, uninstalling |
| [Search and shortcuts](https://github.com/zhugecaomao/ALTRun/wiki/en-Usage) | Search syntax, shortcuts, action panel, learned ranking, usage statistics |
| [Custom commands and snippets](https://github.com/zhugecaomao/ALTRun/wiki/en-Commands-and-Snippets) | Command types, path variables, snippet placeholders, snippet expansion |
| [File search](https://github.com/zhugecaomao/ALTRun/wiki/en-File-Search) | Everything integration, built-in index, exclusion rules |
| [Themes](https://github.com/zhugecaomao/ALTRun/wiki/en-Themes) | Built-in themes, custom themes, every available key |
| [Extensions](https://github.com/zhugecaomao/ALTRun/wiki/en-Extensions) | Quick Switch, date stamp, custom hotkeys, system commands, PT Tools |
| [Settings reference](https://github.com/zhugecaomao/ALTRun/wiki/en-Configuration) | Every ALTRun.json setting and its default value |
| [FAQ](https://github.com/zhugecaomao/ALTRun/wiki/en-FAQ) | Hotkey conflicts, things that can't be found, antivirus false positives and more |
| [Development guide](https://github.com/zhugecaomao/ALTRun/wiki/en-Development) | Architecture, adding a search feature, code style, tests |


## Feedback and contributing
- Report a problem: [open an issue](https://github.com/zhugecaomao/ALTRun/issues/new/choose) (please include your Windows version and the steps to reproduce)
- Ideas and discussion: [Discussions](https://github.com/zhugecaomao/ALTRun/discussions)
- Code contributions: please read [CONTRIBUTING.md](CONTRIBUTING.md) first; the [Development guide](https://github.com/zhugecaomao/ALTRun/wiki/en-Development) covers the project structure, running the tests and adding a search feature

Issues and pull requests are welcome in English or Chinese. If ALTRun is useful to you, a star ⭐ is much appreciated.


## Code signing policy
Free code signing provided by [SignPath.io](https://about.signpath.io/), certificate by [SignPath Foundation](https://signpath.org/).

- Only `ALTRun.exe` built from this repository's source by the [Release workflow](.github/workflows/release.yml) is signed.
- Committers and reviewers: [zhugecaomao](https://github.com/zhugecaomao)
- Approvers: [zhugecaomao](https://github.com/zhugecaomao)
- Privacy policy: [SECURITY.md](SECURITY.md#privacy-policy)


## License and credits
ALTRun is open source under the [GPL-3.0](LICENSE) license © zhugecaomao

Thanks to these projects for the inspiration:
- [ALTRun](https://github.com/etworker/ALTRun) (etworker, Delphi): the name and the original design
- [RunZ](https://github.com/goreliu/runz) (goreliu, AutoHotkey): the idea of writing a launcher in AutoHotkey
- [Alfred](https://www.alfredapp.com/): the interaction model
- [Listary](https://www.listary.com/): its Quick Switch is the model for ALTRun's
