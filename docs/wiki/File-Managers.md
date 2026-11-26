**English** · [中文](File-Managers-zh)

# File managers
ALTRun opens folders in the file manager you choose, and reads the folder open in it for [Quick Switch](Quick-Switch), the [Dialog Folder Panel](Quick-Switch-Panel), "Open terminal here" and the action panel. File Explorer, Total Commander, Directory Opus and Arke (0.8.0 or later) are supported in full; any other file manager that takes a folder on its command line can open folders.

## Choosing the file manager
Preferences → General → **File manager**: the program, optionally followed by its parameters (setting `General.FileManager`, see [Settings reference](Configuration)). When the program isn't found, File Explorer opens the folder (noted in the log).

| File manager | Write in **File manager** | Notes |
|---|---|---|
| File Explorer | `explorer.exe` (the default) | |
| Total Commander | `C:\Apps\TotalCMD64\TOTALCMD64.exe /O /T /S` | `/O` the window already open, `/T` a new tab, `/S` the active panel; `/R={path}` always the right panel |
| Directory Opus | the path of `dopus.exe`, with its own parameters if you want | the folder is added after the parameters, or where `{path}` is |
| Arke | `C:\Apps\Arke\arke.exe` | the file name must be `arke.exe` for folders in a new tab (see [Arke](#arke)) |
| Others | the program, with its parameters | the folder is added after the parameters, or where `{path}` is |

`{path}` in the parameters places the path (quoted automatically when it has spaces). A program path with spaces needs quotes, for example `"C:\Program Files\Arke\arke.exe"`. Environment variables work: `%LocalAppData%\Programs\Arke\arke.exe`.

**Open folders in a new tab** (below it, `General.FolderNewTab`, off by default) opens folders in a new tab of a window already open instead of a new window.

## What works with each file manager

| | File Explorer | Total Commander | Directory Opus | Arke 0.8.0+ | Others |
|---|---|---|---|---|---|
| Open a folder (`Enter` on a folder, folder commands) | ✓ | ✓ | ✓ | ✓ | ✓ |
| Reveal in File Manager (`Ctrl+Enter`): the folder, with the item selected | ✓ | ✓ (ALTRun adds `/P`) | the folder only | with "new tab": ✓; else the folder only | the folder only |
| Open folders in a new tab | ✓ (Windows 11 22H2 or later) | ✓ (`/O /T` added) | its own parameters | ✓ (asked through its interface) | its own parameters |
| `Ctrl+G` in a dialog, auto switch, jump when a dialog opens | ✓ | ✓ | ✓ | ✓ | - |
| Folders listed in the Dialog Folder Panel | every window | both panels of every window | the source tab and the other side of every window | the active pane and the other pane of every window | - |
| "Open terminal here" (`TerminalHere`) | ✓ | ✓ | ✓ | ✓ | - |
| Action panel: Copy / Move to Current Folder in ... | ✓ | ✓ | ✓ | ✓ | - |

The reading columns (from `Ctrl+G` down) don't depend on the **File manager** setting: with Explorer as the file manager, `Ctrl+G` still jumps to the folder of a Total Commander or Arke window when you used it last. ALTRun looks at the windows from front to back and takes the first file manager window, so it is the one you used most recently.

How ALTRun reads the folder, without the clipboard (the paths don't appear in Windows clipboard history):

| File manager | How | Waits at most |
|---|---|---|
| File Explorer | the Shell windows list | - |
| Total Commander | its `WM_COPYDATA` interface (TC 8.0 or later) | 0.5 s |
| Directory Opus | its own `dopusrt.exe /info …,paths` (next to the running `dopus.exe`); a result is reused for 1.5 s | 3 s |
| Arke | its `WM_COPYDATA` interface (`state`), one request for all its windows; from Arke 0.9.1 without the selected files, so it stays fast with thousands selected | 0.5 s |

## Arke
[Arke](https://github.com/Koimsurai/Arke) 0.8.0 and later has an interface for launchers and scripts; ALTRun uses it, and there's nothing to turn on in Arke. Older versions still open folders (`arke.exe "D:\Folder"`), but ALTRun can't read their current folder.

### Setting it up, step by step
1. Preferences → General → **File manager**: `C:\Apps\Arke\arke.exe` (where your `arke.exe` is)
2. Tick **Open folders in a new tab** if you want ALTRun's folders in a new tab of the Arke window you used last. Click **OK**
3. Try it:

| You do | Result |
|---|---|
| Type a folder's name in ALTRun and press `Enter` | Arke opens it in a new tab of the Arke window used last and comes to the front |
| Select a file in the results and press `Ctrl+Enter` (Reveal in File Manager) | Arke opens the file's folder in a new tab, with the file selected |
| In Word press `Ctrl+O`, then `Ctrl+G` | the dialog jumps to the folder of Arke's active pane |
| Look at the folder panel below that dialog | each Arke window has a row "Arke" (the active pane) and, in split view, "Arke (other pane)" |
| Add a [custom hotkey](Hotkeys) with the action `TerminalHere` and the window `ahk_class ArkeWindow` (the default `Ctrl+Alt+T` is for Explorer only), then press it in Arke | a terminal opens in Arke's current folder |

### Which folder counts
- Only the **active tab** of each window: the active pane, and in split view the other pane too (in the panel). Other tabs aren't listed
- Several Arke windows: `Ctrl+G` takes the one used last; the panel lists them all, front to back
- Only folders, drives (`D:\`) and network shares (`\\server\share`) count. Home, search results, the Recycle Bin, SFTP / WebDAV locations and pages such as Settings are no folder: `Ctrl+G` can't jump there (it says no file manager window is open), and the panel has no row for that pane
- With **Open folders in a new tab**, ALTRun asks the Arke window that was in front before ALTRun was shown, else the Arke window used last. When no Arke window is open (only its notification area icon), Arke is older than 0.8.0 or doesn't answer, ALTRun runs `arke.exe "folder"` as without the option (in Arke's single window mode that opens a tab too)
- Without **Open folders in a new tab**: `arke.exe "folder"`; Reveal opens the folder that holds the item, without selecting it

### Limits
- Arke running **as administrator** while ALTRun isn't: Windows doesn't pass the messages up, so ALTRun can't read Arke's folders (`Ctrl+G` says no file manager window is open while Arke is in front of the others), and opening falls back to `arke.exe "folder"`. Run both with the same rights
- ALTRun only reads Arke's state and asks it to open a folder; it never changes files through Arke (Arke's interface doesn't allow it either)
- With [logging](Configuration) on, the log shows how long each read took (`Perf: Arke state … ms, … windows`) and why a folder wasn't opened through the interface (`Arke: didn't open …`)
