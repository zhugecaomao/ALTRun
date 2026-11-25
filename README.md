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
  <a href="#highlights">Highlights</a> ·
  <a href="#getting-started">Getting started</a> ·
  <a href="https://github.com/zhugecaomao/ALTRun/wiki/Home">Documentation (Wiki)</a> ·
  <a href="CHANGELOG.md">Changelog</a>
</p>

<p align="center">
  <img src="docs/images/screenshots/demo.png" width="700" alt="ALTRun: type to search, open the action panel, convert units and search files">
</p>


## Why ALTRun
- **Tiny**: about 2.5 MB, one `ALTRun.exe` and a few resource files; no .NET, Electron or other runtime to install.
- **Fast**: results appear as you type; with [Everything](https://www.voidtools.com/) running, file searches cover all disks in milliseconds.
- **Smart**: learns which result you pick for what you type; matches word initials (`vsc` → Visual Studio Code) and pinyin initials (`wx` → 微信).
- **Keyboard first**: `→` for actions, `F3` to edit in place, `Ctrl+1`–`Ctrl+9` to open a result, `?` for every shortcut.
- **Portable and private**: no installer, registry entries, admin rights or background service; your data stays on your PC, with no telemetry.
- **Always up to date**: new versions show up in the search window and install with `Enter`, SHA256-verified, settings kept; winget and Scoop too.
- **Free and open source**: GPL-3.0, with thousands of automated tests run on Windows for every change.


## Highlights
| **Search everything** from one box: apps, files, custom commands, windows and system commands, as you type ([more](https://github.com/zhugecaomao/ALTRun/wiki/Usage)) | **Action panel**: press `→` to run as administrator, open the containing folder, copy the path, open a terminal there... ([more](https://github.com/zhugecaomao/ALTRun/wiki/Usage#action-panel)) |
|:---:|:---:|
| <img src="docs/images/screenshots/search.png" alt="Search window"> | <img src="docs/images/screenshots/actions.png" alt="Action panel"> |
| **Quick Switch**: jump any Open / Save dialog to a folder open in File Explorer, Total Commander, Directory Opus or Arke ([more](https://github.com/zhugecaomao/ALTRun/wiki/Quick-Switch)) | **Your actions for any content**: select text, files or a URL anywhere and press `Ctrl+Alt+\`; your commands for that kind of content come first ([more](https://github.com/zhugecaomao/ALTRun/wiki/Content-Types)) |
| <img src="docs/images/screenshots/demo-quickswitch.png" alt="Quick Switch: Ctrl+G and the folder panel jump the dialog to an open folder"> | <img src="docs/images/screenshots/actions-file-type.png" alt="Your own commands first for a Markdown file"> |
| **Clipboard history**: `Ctrl+Alt+C` brings back copied text, files and images; filter and paste ([more](https://github.com/zhugecaomao/ALTRun/wiki/Clipboard)) | **Snippets**: type `;thx` in any program to expand a reply, a signature or an address ([more](https://github.com/zhugecaomao/ALTRun/wiki/Snippets)) |
| <img src="docs/images/screenshots/demo-clipboard.png" alt="Clipboard history: Ctrl+Alt+C, filter, Enter pastes"> | <img src="docs/images/screenshots/demo-snippets.png" alt="Snippets: ;thx expands into an email reply"> |
| **Calculator**: `12*(3+4)`, `10 km in mi`, `255 in hex`, `today + 30 days`, `100 usd to sgd`, as you type ([more](https://github.com/zhugecaomao/ALTRun/wiki/Calculator)) | **File search**: `Space` + a name searches every disk with Everything, or the built-in index ([more](https://github.com/zhugecaomao/ALTRun/wiki/File-Search)) |
| <img src="docs/images/screenshots/demo-calculator.png" alt="Calculator: arithmetic and unit conversion as you type"> | <img src="docs/images/screenshots/files.png" alt="File search"> |
| **Browse folders**: type a path, `Tab` goes in, `Backspace` goes up ([more](https://github.com/zhugecaomao/ALTRun/wiki/File-Search#browsing-folders)) | **Pinyin initials**: `jsb` finds 记事本 (Notepad); file names too ([more](https://github.com/zhugecaomao/ALTRun/wiki/Usage#how-matching-works)) |
| <img src="docs/images/screenshots/demo-browse.png" alt="Browse folders: type a path, Tab, filter, Backspace"> | <img src="docs/images/screenshots/pinyin.png" alt="Pinyin initials"> |
| **Web search**: `g keywords`, `bd keywords` or your own engines; one keyword can open several sites ([more](https://github.com/zhugecaomao/ALTRun/wiki/Web-Search)) | **Window switching**: `w` + a title or program name switches to an open window ([more](https://github.com/zhugecaomao/ALTRun/wiki/Window-Switching)) |
| <img src="docs/images/screenshots/websearch.png" alt="Web search"> | <img src="docs/images/screenshots/windows.png" alt="Window switching"> |
| **Scripts with live results**: an `.ahk`, `.ps1` or `.py` script can list results while you type ([more](https://github.com/zhugecaomao/ALTRun/wiki/Scripts#live-results)) | **22 themes**, each in three densities, following the Windows light / dark mode if you like ([more](https://github.com/zhugecaomao/ALTRun/wiki/Themes)) |
| <img src="docs/images/screenshots/script-filter.png" alt="Live results of a script"> | <img src="docs/images/screenshots/prefs-appearance.png" alt="Themes"> |

The screenshots are made automatically by [Tests/Screenshots](Tests/Screenshots/TakeScreenshots.ahk) on a Windows runner in GitHub Actions.


## More features
- **Search**: browser bookmarks and history, recent files, pinned and recent items, 40+ Windows Settings pages and system commands, terminal (`>ipconfig`), large type
- **Your own**: custom commands with arguments, content types, custom hotkeys, Send to ALTRun from Explorer, a [script library](https://zhugecaomao.github.io/ALTRun/#scripts)
- **Tools**: date stamp while renaming (`Ctrl+D`), folder sync with robocopy, delayed start, Caps Lock to switch the input method, PT Tools for structural engineers
- **Interface**: English, Simplified Chinese, Traditional Chinese and Japanese

Every feature, with screenshots: [Features](https://github.com/zhugecaomao/ALTRun/wiki/Features).


## Getting started
1. [Download the latest release](https://github.com/zhugecaomao/ALTRun/releases/latest), unzip it to any folder and run `ALTRun.exe` (no installation or AutoHotkey needed). Or install it with a package manager:
   ```powershell
   winget install zhugecaomao.ALTRun
   # or
   scoop bucket add altrun https://github.com/zhugecaomao/ALTRun
   scoop install altrun
   ```
2. Press `Alt+Space` (or `Alt+R`), type a few letters and press `Enter`.
3. Type `?` for every search syntax and shortcut; press `Ctrl+,` for Preferences.

If Windows SmartScreen blocks it, choose **More info** → **Run anyway**; for antivirus warnings, see the [FAQ](https://github.com/zhugecaomao/ALTRun/wiki/FAQ#my-antivirus-flags-it). Upgrading, running from source with AutoHotkey v2 and uninstalling: [Installation and upgrades](https://github.com/zhugecaomao/ALTRun/wiki/Installation).

| To do this | Type |
|---|---|
| Open a program | `word` → `Enter` |
| Find a file or folder | `Space`, then `report` |
| Calculate | `12*(3+4)`; `Enter` copies the result |
| Search the web | `g keywords` |
| Paste from the clipboard history | `Ctrl+Alt+C` |
| Run a command in a terminal | `>ipconfig /all` |
| See what you can do with a result | `→` or right-click |
| Edit the selected command | `F3` |

Every shortcut: [Search and shortcuts](https://github.com/zhugecaomao/ALTRun/wiki/Usage#shortcuts).


## Documentation
The [Wiki](https://github.com/zhugecaomao/ALTRun/wiki/Home) covers every feature in detail, with examples. Good places to start:
[Features](https://github.com/zhugecaomao/ALTRun/wiki/Features) ·
[Getting started](https://github.com/zhugecaomao/ALTRun/wiki/Getting-Started) ·
[Search and shortcuts](https://github.com/zhugecaomao/ALTRun/wiki/Usage) ·
[Custom commands](https://github.com/zhugecaomao/ALTRun/wiki/Custom-Commands) ·
[Snippets](https://github.com/zhugecaomao/ALTRun/wiki/Snippets) ·
[Quick Switch](https://github.com/zhugecaomao/ALTRun/wiki/Quick-Switch) ·
[Scripts](https://github.com/zhugecaomao/ALTRun/wiki/Scripts) ·
[Settings reference](https://github.com/zhugecaomao/ALTRun/wiki/Configuration) ·
[FAQ](https://github.com/zhugecaomao/ALTRun/wiki/FAQ)


## Feedback and contributing
- [Report a bug](https://github.com/zhugecaomao/ALTRun/issues/new?template=bug_report.yml) (with your Windows version and the steps to reproduce), [suggest a feature](https://github.com/zhugecaomao/ALTRun/issues/new?template=feature_request.yml) or [start a discussion](https://github.com/zhugecaomao/ALTRun/issues/new?template=discussion.yml), in English or Chinese
- Code: read [CONTRIBUTING.md](CONTRIBUTING.md) first; the [Development guide](https://github.com/zhugecaomao/ALTRun/wiki/Development) covers the project structure, the tests and adding a search feature
- Scripts: contribute one to the [script library](docs/scripts/README.md#contributing-a-script--投稿脚本) with a pull request

If ALTRun is useful to you, a star ⭐ is much appreciated.


## Code signing policy
Free code signing provided by [SignPath.io](https://about.signpath.io/), certificate by [SignPath Foundation](https://signpath.org/).

- Only `ALTRun.exe` built from this repository's source by the [Release workflow](.github/workflows/release.yml) is signed.
- Committers and reviewers: [zhugecaomao](https://github.com/zhugecaomao)
- Approvers: [zhugecaomao](https://github.com/zhugecaomao)
- Privacy policy: [SECURITY.md](SECURITY.md#privacy-policy)


## License and credits
ALTRun is open source under the [GPL-3.0](LICENSE) license © zhugecaomao

Thanks to these projects:
- [ALTRun](https://github.com/etworker/ALTRun) (etworker, Delphi): the name and the original design
- [RunZ](https://github.com/goreliu/runz) (goreliu, AutoHotkey): the idea of writing a launcher in AutoHotkey
- [Alfred](https://www.alfredapp.com/): the interaction model
- [Listary](https://www.listary.com/): its Quick Switch is the model for ALTRun's
- [pypinyin](https://github.com/mozillazg/python-pinyin) (Copyright (c) 2016 mozillazg, 闲耘; [MIT](https://github.com/mozillazg/python-pinyin/blob/master/LICENSE.txt)): pinyin initials of rare and Traditional characters
