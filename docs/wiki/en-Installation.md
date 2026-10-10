**English** · [中文](Installation)

# Installation and upgrades

## System requirements
- Windows 10 or Windows 11 (64-bit)
- Running from source needs [AutoHotkey v2](https://www.autohotkey.com/) 2.0 or later; the released exe doesn't
- No .NET, Electron or other runtime; the zip is under 1 MB and about 2 MB unpacked

## Installing
ALTRun is portable: there is nothing to install and nothing is written to the registry.

1. Download the latest version from [Releases](https://github.com/zhugecaomao/ALTRun/releases/latest) and unzip it to any folder (for example `D:\Apps\ALTRun`)
   - Or clone the repository and double-click `ALTRun.ahk` to run it with AutoHotkey v2
2. The ALTRun icon appears in the notification area; press `Alt+Space` to open the search window
3. The first run builds the app index in the background; after a few seconds your Start menu programs can be found

If Windows SmartScreen says "Windows protected your PC" when you first run it, click **More info** → **Run anyway**; if your antivirus complains, see the [FAQ](en-FAQ#my-antivirus-flags-it).

Put it in a folder you can write to (for example `D:\Apps\ALTRun`): settings and data are saved in the `Data\` folder next to the program, so you can take the whole folder with you. In a folder that needs admin rights to write to (such as `C:\Program Files`), settings and data are saved in `%APPDATA%\ALTRun\` instead, but one-key updates can't replace the program files and you have to download new versions by hand.

### Installing with Scoop
[Scoop](https://scoop.sh/) is a command-line installer for Windows. The ALTRun repository is itself a Scoop bucket:

```powershell
scoop bucket add altrun https://github.com/zhugecaomao/ALTRun
scoop install altrun
```

- After installing, the Start menu has "Scoop Apps → ALTRun"; the program is in `scoop\apps\altrun\current`
- Upgrading: quit ALTRun first, then run `scoop update altrun`. Settings and data (`Data\`, including `Data\ALTRun.json`) and `Themes\` are kept
- Uninstalling: `scoop uninstall altrun`; your settings stay in `scoop\persist\altrun` and come back when you reinstall. To delete them too, use `scoop uninstall altrun --purge`
- When installed with Scoop, a new version found by "Check for Updates" tells you to upgrade with `scoop update altrun`; don't unzip over it by hand

### Installing with winget
The winget manifest is ready (`packaging\winget`). Once the [official winget repository](https://github.com/microsoft/winget-pkgs) accepts it, you can run:

```powershell
winget install zhugecaomao.ALTRun
winget upgrade zhugecaomao.ALTRun     # later upgrades (quit ALTRun first)
```

- The program is in `%LOCALAPPDATA%\Microsoft\WinGet\Packages\zhugecaomao.ALTRun_...`; type `altrun` on the command line to start it
- Upgrading and uninstalling only remove the files that came in the zip; settings and data (`Data\`, `Themes\`) are kept. Only `winget uninstall --purge` deletes them as well
- winget itself doesn't create a Start menu shortcut: after you start it once with `altrun`, ALTRun adds itself to the Start menu (Preferences → General → "Add to Start menu")
- Before uninstalling, turn off "Launch ALTRun at login", "Add to Start menu" and "Add to Explorer 'Send to' menu" in Preferences → General, otherwise those shortcuts are left behind
- The first time it runs, Windows may say it "couldn't verify the publisher" (the program was downloaded from the internet); choose "Run"
- Upgrading and uninstalling both keep your settings, `Data\` and `Themes\`; to remove everything after uninstalling, delete the folder above by hand

## Files in the program folder
| File / folder | What it is | On upgrade |
|---|---|---|
| `ALTRun.ahk` or `ALTRun.exe` | The program | Replaced |
| `Lib\` `Src\` | Program code (source version) | Replaced |
| `Resources\` | Data shipped with the program: Simplified / Traditional table, built-in themes, interface languages, icons | Overwritten |
| `Data\ALTRun.json` | All your settings, custom commands and snippets | Kept (format upgraded automatically) |
| Other files in `Data\` | App index, file index, learned ranking, clipboard history, usage statistics | Kept (rebuilt if deleted) |
| `Themes\` | Your own themes | Kept |
| `Scripts\` | Your own [scripts](en-Extensions#scripts) | Kept |

## Start with Windows, Start menu, "Send to"
In Preferences → General:
- **Launch ALTRun at login**: creates a shortcut in the Startup folder
- **Add to Start menu**: lets you open ALTRun from the Start menu
- **Add to Explorer 'Send to' menu**: right-click a file in Explorer → Send to → ALTRun adds it as a custom command (one file opens the edit dialog; several are all added at once)

## Upgrading
### One-key updates
ALTRun checks GitHub for new versions in the background (at startup and then every 6 hours) (Preferences → General → "Check for updates automatically"). Like Alfred, it doesn't pop up a window: once a new version is found, the empty search box shows **"Update Available: ALTRun x"** when you open the search window, with the current version and the keys on the line below (typing `update` finds it too):
- `Enter`: **Install Update**. Downloads the new version, checks its SHA256 checksum, replaces the program files and restarts automatically. `Data\` (settings and data) and `Themes\` are not touched
- `→`: **Release Notes** (opens this version's notes on GitHub) or **Skip This Version** (you'll be told again when there is a newer one)

After Windows starts, it waits 1 minute (the network may not be up yet) and checks if the last check was more than 1 hour ago, so a version released the day before is found when you start your PC; after that it looks every hour whether 6 hours have passed since the last check. The time of the last check and skipped versions are kept in `Data\Update.json`, so reloading ALTRun after changing settings doesn't check again and again. When there's no network it only writes a log entry and tries again an hour later.

To check right away: tray menu → "Check for Updates", or search for "Check for ALTRun Updates" (`CheckUpdate`). A manual check always shows the result: you're up to date, or the "ALTRun Update Available" dialog: Install Update / Release Notes / Later.

If anything goes wrong during an update (network, checksum, write permission), the old program is left as it was and you're offered the download page to update by hand. The command line `ALTRun.exe -Update` checks for and installs an update without asking.

One-key updates aren't possible in these cases; the item in the search window tells you what to do instead:
- Installed with Scoop / winget: `Enter` copies the upgrade command (`scoop update altrun` / `winget upgrade zhugecaomao.ALTRun`); quit ALTRun and run it in a terminal
- Running the source `ALTRun.ahk`: update with `git pull`, or press `Enter` to open the download page
- The program is in a folder you can't write to: `Enter` opens the download page

### Updating by hand
Copy the new version over the program folder; your settings and data are not affected. If the settings file format has changed, ALTRun upgrades it at startup and backs up the original first.

### Upgrading from 2.x (v2026.08.12 and earlier) to 3.0
1. Tray icon → quit the old version
2. Unzip the new version into your existing ALTRun folder, replacing `ALTRun.exe`
3. Run `ALTRun.exe`

On the first run, the old `ALTRun.ini` is imported automatically (settings, custom commands, hotkeys) and saved as the new `Data\ALTRun.json`. `ALTRun.ini` is left unchanged: to go back to the old version, put the old `ALTRun.exe` back. The Startup, Start menu and "Send to" shortcuts keep pointing to the same `ALTRun.exe`, so nothing needs to be set up again.

- **Kept**: the hotkey, start with Windows and other general settings; user commands (File / Dir / CMD / URL → custom commands, Clip → snippets); index folders, file types and depth; the structural calculation switch; Quick Switch; date stamp (Ctrl+D); PT Tools settings; conditional hotkeys
- **No longer kept**: the built-in command list (replaced by [system commands](en-Extensions#system-commands)), the old index (rebuilt), run history, usage statistics, old list appearance options

The old `Res\` folder has been renamed to `Resources\`; at startup, any files you put in `Res\` yourself are moved over automatically.

## Uninstalling
1. Tray menu → Quit
2. In Preferences → General, first turn off "Launch ALTRun at login", "Add to Start menu" and "Add to Explorer 'Send to' menu" (or delete those shortcuts by hand)
3. Delete the program folder
