**English** · [中文](Commands-and-Snippets)

# Custom commands and snippets

## Custom commands
Add the files, folders, programs or web addresses you use often to ALTRun and open them by name or keyword.

![Custom commands in Preferences](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-commands.png)

### Adding commands
- In the search results, select an app / file / folder / URL and press `F3` (or `→` → "Add to Custom Commands...")
- When nothing is found, press `F3` to create one from the text you typed
- In Explorer, right-click → Send to → ALTRun. With **one** file / folder selected, the edit dialog opens with the name, type and target already filled in; adjust the name or add a keyword and press `Enter` (Cancel adds nothing). With **several** selected, they're all added at once (no dialog for each one), and commands that already exist aren't added twice. Sending a file / folder that already has a command opens that command for editing. If Preferences is open, the new commands are added to the list in Preferences → Custom Commands and saved when you press OK / Apply
- Preferences → Custom Commands → Add...

Existing commands: find one and press `F3` to edit it or `Ctrl+Del` to delete it, or double-click it in Preferences. With many commands, type a word from the name, path or keyword into the **Filter** box at the bottom right of the list to show only matching commands (the snippet, web search and custom hotkey lists have one too).

### Checking paths
When a folder is renamed or a file is moved, the commands pointing to it stop working. Preferences → Custom Commands → **Check Paths** checks whether every command's target still exists, shows the result in the "Status" column and selects the first command with a problem; double-click it to fix the path:

| Status | Meaning |
|---|---|
| OK | The target exists |
| Not found | The file or folder is gone (renamed, moved or deleted); fix the path or delete the command |
| Offline | The drive or network location can't be reached right now (for example a network drive isn't connected), so it can't be checked; check again once it's connected |
| - | URLs and locations such as `shell:` aren't checked |

For the `Command` type with only a program name (for example `cmd.exe`), the program is looked up in PATH and the programs registered with Windows (App Paths). The results only appear in Preferences; no command is changed or deleted.

### Fields
Each field in the edit dialog has a line of gray help text below it.

| Field | Description |
|---|---|
| Title | The name shown in results and used for searching |
| Type | `File` file / program, `Folder` folder, `Command` program + arguments, `Url` web address / link (a web page, or a link such as `ms-settings:windowsupdate` or `mailto:`) |
| Target | Path, program or web address |
| Arguments | Command-line arguments passed to the program, for example `/k ipconfig /all` |
| Keyword | Optional. Typing exactly this keyword puts the command first |

Besides the title and keyword, `File` / `Folder` commands also match the target's file name (without the extension) or folder name: with the Target `Q:\Projects\PT1931 - 24 NIR` and the title `CKR, EA, JIB`, typing `nir` or `1931` finds it too. At the same match quality, a title match ranks first.

Folders open in Preferences → General → File manager (Explorer by default; Total Commander and others work too).

### Commands with an argument ({query})
Put `{query}` in the target or arguments and give the command a keyword; typing "keyword text" replaces `{query}` with the text (URL-encoded in web addresses). It works just like a web search keyword:

| Title | Type | Target | Arguments | Keyword | Type this |
|---|---|---|---|---|---|
| Jira | Link | `https://jira.example.com/browse/{query}` | | `jira` | `jira ABC-123` |
| Ping | Command line | `cmd.exe` | `/k ping {query}` | `ping` | `ping 10.0.0.1` |
| Project folder | Folder | `D:\Projects\{query}` | | `pj` | `pj 2026-05` |
| Google Maps | Link | `https://www.google.com/maps/search/{query}` | | `maps` | `maps Changi Airport` |

A new installation already has three examples among the default commands: Ping, Google Maps and Documents subfolder (`docs folder-name` opens a subfolder of Documents). When you find such a command by its name, the line below it says what to type (for example "Type after 'ping ', then press Enter"), and `Enter` completes it to "keyword " so you can type the argument. Without a keyword, `{query}` is replaced by nothing. "Check Paths" skips paths containing `{query}`.

### Variables in paths
| Write | Meaning |
|---|---|
| `A_Desktop` `A_DesktopCommon` | Desktop / public desktop |
| `A_MyDocuments` | Documents |
| `A_AppData` | `%AppData%` |
| `A_Programs` `A_ProgramsCommon` | Start menu "Programs" (current user / all users) |
| `A_StartMenu` `A_StartMenuCommon` | Start menu |
| `A_Startup` `A_StartupCommon` | Startup folder |
| `A_ProgramFiles` `A_WinDir` `A_Temp` | Program Files / Windows / temporary folder |
| `A_ScriptDir` | The folder ALTRun is in |
| `%Temp%` `%AppData%` `%UserProfile%` `%OneDrive%` | Environment variables |

`A_` variables must come first, for example `A_Desktop\Projects`. A bare program name (for example `notepad`) is looked up in PATH.

## Snippets
Save text you use often (signatures, addresses, templates...) as snippets. Search by name, keyword or a word in the text, and `Enter` pastes it into the window you were in before opening ALTRun. Type `snip` to list all snippets, `snip xxx` to search snippets only.

| Field | Description |
|---|---|
| Name | Required. Shown in search results and used for searching. Use words you would type to find it, for example `PT quotation` |
| Keyword | Optional and short, for example `pq`. An exact match ranks first; to auto-expand, type `;pq`. The prefix plus the keyword can be at most 40 characters (39 for the keyword with the default prefix `;`), without spaces; you're told before saving if it's too long |
| Text | Required. What gets pasted; can be several lines and use the placeholders below. Words in the text are searchable too (type at least 3 characters; turn it off in Preferences if you don't want this) |
| AutoExpand | Type `;keyword` in any program and it's replaced by the text right away |

### Placeholders
Replaced when pasting:

| Placeholder | Replaced by |
|---|---|
| `{date}` | Today's date, in the DateFormat of the [date stamp](en-Extensions#date-stamp) (default `dd.MM.yyyy`) |
| `{time}` | The current time `HH:mm` |
| `{datetime}` | Date + time |
| `{clipboard}` | The text on the clipboard |
| `{date:yyyy-MM-dd}` `{time:HH:mm:ss}` | Your own format (as in [FormatTime](https://www.autohotkey.com/docs/v2/lib/FormatTime.htm), for example `dddd` for the day of the week) |
| `{date+7}` `{date-1:dd.MM}` | A date some days later / earlier, optionally with a format |
| `{clipboard:1}` `{clipboard:2}` | The 1st, 2nd previous text in the clipboard history (like Alfred; `{clipboard}` is the current clipboard; files and images are skipped) |
| `{uuid}` | A randomly generated UUID, different each time |
| `{cursor}` | The cursor ends up here after pasting |

### Snippet expansion
Type `;keyword` (for example `;sig`) in any program and what you typed is deleted and replaced by the snippet text, like Alfred's snippets. The `;` prefix keeps normal typing from triggering it by accident; change it in `Features.Snippets.ExpandPrefix`. To stop a single snippet from expanding, untick "Expand automatically when typed" in its edit window (`"AutoExpand": 0`); it can still be searched. Snippets don't expand in ALTRun's own search window, in Windows password boxes, or in the programs listed in Preferences → Snippets → "Don't expand in" (Remote Desktop, KeePass and KeePassXC by default; `Features.Snippets.ExpandExclude`).

### Paste mode
`Features.Snippets.PasteMode`:
- `Clipboard` (default): borrows the clipboard + `Ctrl+V`, then restores the clipboard
- `Type`: types the text character by character, for programs that don't accept pasting

## Clipboard history
Press `Ctrl+Alt+C` or type `clip` to list the text, files and images you copied; `clip keywords` filters, `Enter` pastes into the front window, and `Ctrl+Del` removes an item from the history.

![Clipboard history](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/clipboard.png)

| What was copied | Shown as | `Enter` | `F3` / `→` |
|---|---|---|---|
| Text | The beginning of the text | Pastes the text | Save as a snippet / copy, paste, large type |
| Files | The file names (separated by commas) | Pastes the files (in Explorer that copies them there) | Add to custom commands / for one file: open, reveal in the file manager, copy path and more |
| Image (screenshot, picture from a web page...) | "Image 1920 × 1080" with a thumbnail | Pastes the image | - / open in the image viewer, reveal in the file manager |

- Images are saved as PNG files in the `Clipboard\` folder (by default on this PC in `%LOCALAPPDATA%\ALTRun\Clipboard`, so they aren't uploaded every time when `Data\` is in a synced folder), at most 50 by default (`MaxImages`), and only recorded while "Keep history after ALTRun quits" is on; to stop recording images, turn off "Also record images" in Preferences → Clipboard
- Type `clip image` to see only images
- **Pin**: select an item, press `→` and choose "Pin". Pinned items always stay at the top (with a pin at the bottom right of the icon), are never pushed out when the history is full, and are kept by "Clear Clipboard History"; good for addresses, account names and the like
- **Paste as plain text**: the system command "Paste as Plain Text" pastes the clipboard content without formatting (fonts, colors, tables...) and restores the clipboard afterwards. It's handiest as `Ctrl+Shift+V` in [custom hotkeys](en-Extensions#custom-hotkeys)
- **Merge copies** (off by default, Preferences → Clipboard → "Press Ctrl+C twice to append to the previous item"): copy one piece, then select the next and press `Ctrl+C` twice quickly; both become one item (separated by a line break) and the clipboard holds the merged text, ready to paste. Like Alfred's Merging, handy for collecting bits from several places. Note: translators such as DeepL also use "press Ctrl+C twice", so with this on both react

Privacy:
- Content copied from password managers (KeePass, 1Password, Bitwarden...) is never recorded, nor is content marked "don't add to clipboard history"
- Add other programs you don't want recorded to `Features.Clipboard.IgnoreApps`
- With `Persist = 0` the history is kept in memory only and is gone when ALTRun quits (no images either); otherwise it's saved in `Data\ClipboardHistory.json` (very long items and images are stored separately in `%LOCALAPPDATA%\ALTRun\Clipboard\`; in `Data\Clipboard\` with `LocalFiles = 0`)
- Preferences → Clipboard can clear the history
