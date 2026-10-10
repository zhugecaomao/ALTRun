**English** · [中文](FAQ)

# FAQ

### Nothing happens when I press Alt+Space
- Make sure the ALTRun icon is in the notification area (if not, run it again)
- Another program (an input method, PowerToys Run, another launcher) may use the same hotkey. Pick another one in Preferences → General, or set a "Second hotkey"
- Programs running as administrator don't pass hotkeys to an ALTRun running with normal rights. Run ALTRun as administrator if you need that

### A program can't be found
1. Make sure its shortcut is in the Start menu or on the desktop (or add its folder to Preferences → Applications → Folders)
2. Check whether its name is filtered out by "Exclude (regex)", or whether it's in the "Hidden apps" list
3. Search `rebuild` and run "Rebuild ALTRun Index"

Simpler: add it as a [custom command](en-Commands-and-Snippets), optionally with a keyword.

### Files can't be found
Normal searches don't show files by default. Press `Space` in an empty search box first (or use `'report` / `open report`) to start a file search. Without Everything running, ALTRun only looks in Desktop, Documents and Downloads; install and run [Everything](https://www.voidtools.com/) to search all disks. See [File search](en-File-Search).

### The result I use most isn't first
Pick it a few more times and ALTRun remembers your choice ([learned ranking](en-Usage#learned-ranking)). You can also add it as a custom command with a short keyword: typing exactly that keyword always puts it first.

### Results include apps I don't want
Select one and press `Ctrl+Del`. The program isn't uninstalled, and you can restore it in Preferences → Applications → "Hidden apps".

### My antivirus flags it
Programs compiled with AutoHotkey are sometimes reported by mistake. Add the ALTRun folder to your antivirus' trusted list, or install AutoHotkey v2 and run the source `ALTRun.ahk` directly (the code is fully public).

### How do I send a debug log?
1. Preferences → General → "Write a debug log"
2. Reproduce the problem
3. Type `log` in the search box and run "Open ALTRun Log": `%Temp%\ALTRun.log` opens in the default program for `.log` files (Notepad unless you changed it)
4. Copy the lines around the time of the problem into your [issue](https://github.com/zhugecaomao/ALTRun/issues/new/choose). The log never contains what you type in the search box, but it does contain file paths and program names: remove anything private first

### ALTRun is slow or doesn't respond right after Windows starts
Just after Windows starts, Explorer and OneDrive (or another sync tool) are often busy for a while. ALTRun asks Windows for icons and shortcuts, so it has to wait for them too: the search window may open slowly, or ignore typing for a few seconds. It recovers by itself once Windows has settled down. To check, turn on the debug log: lines starting with `Perf: startup:` show how long each startup step took, and `Perf: icon ... ms` lines show icons that took a long time to load. If it happens every time, look in Task Manager at what is using the CPU or disk after startup (OneDrive syncing many files is a common cause).

### Searching is awkward with a Chinese input method
Turn on "Switch to English input when shown" in Preferences → Search Window. Chinese names can also be searched by their pinyin initials ("wx" → 微信).

### Snippet expansion doesn't work
- Make sure the snippet has a keyword in Preferences → Snippets and "Expand automatically when typed" is ticked
- Type the prefix: `;keyword` by default
- Some programs (games, Remote Desktop, programs running as administrator) don't receive simulated input

### I messed up my settings; how do I go back to the defaults?
Quit ALTRun, rename `Data\ALTRun.json` as a backup, and run it again to get the default settings. To reset only the learned ranking: Preferences → Advanced → Reset Learned Ranking.

### Does ALTRun go online or collect data?
It doesn't collect or upload any data. By default it only goes online to check for updates (at startup and every 6 hours) and for one-key updates, both only to GitHub (with currency conversion on, it also downloads exchange rates from frankfurter.dev once a day); a one-key update checks the SHA256 checksum before replacing anything. Clipboard history, usage statistics and the learned ranking stay in the `Data\` folder on your PC. See the [security policy](https://github.com/zhugecaomao/ALTRun/blob/main/SECURITY.md).

### How do I upgrade to a new version?
ALTRun checks in the background at startup and every 6 hours after. When there's a new version, open the search window to see "Update Available: ALTRun x" and press `Enter` (or tray icon → Check for Updates); settings and data are kept. With Scoop, use `scoop update altrun`. See [Installation and upgrades](en-Installation#upgrading).

### Using the same settings on several PCs
Two ways:
- **Sync only the settings and data**: Preferences → Advanced → Settings and data → **Change Location...**, and pick a folder in a synced drive (OneDrive, Dropbox...). Your current settings and data are copied there, and ALTRun reloads and uses that folder. On the other PC, pick the same folder; when it says the folder already has ALTRun settings, choose **Yes** to use them. The location is kept in `"DataLocation"` in `Data\ALTRun.json` at the default location (environment variables such as `%OneDrive%\ALTRun` work, so it's fine if the OneDrive path differs between PCs); pick the original `Data\` folder again in Change Location..., or delete that entry, to go back to the default location
- **Keep the whole folder on a USB stick or a synced drive**: ALTRun is portable, so the program and `Data\` travel together

Avoid running ALTRun and changing settings on two PCs at the same time with data in a synced folder, or the sync service may create conflicting copies. Use [path variables](en-Commands-and-Snippets#variables-in-paths) (for example `A_Desktop`, `%OneDrive%`) where you can, so paths work on every PC.

### My commands are gone after upgrading from 2.x
User commands from 2.x are imported from `ALTRun.ini` as custom commands, and the built-in commands are replaced by [system commands](en-Extensions#system-commands). The import happens only once, when there's no `ALTRun.json` yet: to import again, quit ALTRun, rename `Data\ALTRun.json` and run it again. `ALTRun.ini` is never changed; see [Installation and upgrades](en-Installation#upgrading-from-2x-v20260812-and-earlier-to-30).

---
Didn't find an answer? [Open an issue](https://github.com/zhugecaomao/ALTRun/issues/new/choose) or ask in [Discussions](https://github.com/zhugecaomao/ALTRun/discussions).
