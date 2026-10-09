**English** · [中文](File-Search)

# File search

## How to search for files
Normal searches show apps, commands, snippets and so on, but **not** files from all over your disks (so unrelated files such as `NIRMALA.TTF` don't get mixed in). To search for files:

- **Space**: in an empty search box, press `Space` first; the input box shows a gray "Search files...", then type the name. If you've already typed something, move the cursor to the start (`Home`) and press `Space` to search files for it. `Backspace` in an empty input box or at the start goes back to a normal search. Like Alfred's Quick File Search
- **`'report`**, **`open report`** or **`find report`**: same effect
- **Folders only**: **`folder bk`** (change the keyword in Preferences → File Search → "Folder keywords"); `folder bk` also works in file search mode. With Everything this is the same as Everything's `folder:bk`

- **By type** (like Listary's filters): **`doc report`** searches documents only; there are also `pic` pictures, `video`, `audio`, `zip` archives, `exe` programs and `cad` (dwg / dxf / dgn / rvt / ifc / skp...). `doc report` also works in file search mode. Change the keywords and extensions in Preferences → File Search → "File types", one `keyword = ext ext ...` per line. With Everything this is the same as `ext:doc;docx;...`
- To find files by modification date, write Everything's syntax directly, for example `Space` + `dm:today report` or `dm:thisweek`

A keyword needs a space after it to start a file search: typing just `folder`, `open` or `doc` shows commands and apps with that word in their name as usual.

Results are sorted by how well the name matches: exact > start of the name > start of a word (for example `PT2310-BK`) > contains (for example `notebk`); at the same match quality, folders come before files, then by modification time (newest first). With Everything, ALTRun takes Everything's first 300 results and then sorts them, so the best-matching folder isn't missed just because it wasn't modified recently.

Everything's search syntax can be used as is and is passed straight to Everything: for example `Space` + `folder:bk 2024`, `ext:pdf report` or `path:Tender report`.

At most 30 results are shown; when nothing is found you can continue in Everything or Windows Search with one key. When a normal search finds nothing at all, a "Search files" fallback item is offered at the end as well.

To show files in normal searches too (the old behavior): in Preferences → File Search, tick "Also show matching files and folders in normal results"; at most 6 are shown, and only those whose name or a word in it starts with what you typed.

Folders you use often work better as [custom commands](en-Commands-and-Snippets): custom commands also match their target's folder name, so a normal search finds them.

![File search](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/files.png)

`Enter` opens, `Ctrl+Enter` reveals in the file manager, `Alt+Enter` copies the path, `→` shows more actions (Open With, copy / cut the file, copy / move to the current folder in TC, Open Terminal Here, Properties, Move to Recycle Bin...), see the [action panel](en-Usage#action-panel).

## Browsing folders
Type a path in the search box to list what's in that folder (folders first); the last part of the path filters the list:

![Browsing folders](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/browse.png)

| Input | Result |
|---|---|
| `C:\` `D:\Projects\` | Everything in that folder |
| `D:\Projects\rep` | Items whose name matches `rep` (the same fuzzy matching as a normal search) |
| `~\` | Your user folder (`C:\Users\your-name`) |
| `%OneDrive%\` | Environment variables |
| `\\server\share\` | Network shares |

`Tab` goes into the selected folder and `Backspace` removes one part of the path to go up a level; a folder in any result (custom command, file search...) can be entered with `Tab` too. `Enter` opens, and `→` has all the file actions. Hidden and system files aren't listed. `Insert` marks several files to copy, move or delete them together, see [shortcuts](en-Usage#shortcuts).

## Where results come from
ALTRun picks automatically:

| Situation | What is searched |
|---|---|
| [Everything](https://www.voidtools.com/) is running | Queries **all disks** directly through Everything's IPC interface, instantly; no `Everything64.dll` or `es.exe` needed |
| Everything isn't running | The built-in index: scans Desktop, Documents and Downloads in the background (4 levels deep, at most 30,000 items), cached in `Data\FileIndex.json` and rescanned at startup when older than 30 minutes. If Everything is running when ALTRun starts, the built-in index isn't read or scanned until Everything is closed and it's actually needed |

The top of Preferences → Search Sources shows whether Everything is running right now.

Installing Everything and letting it start with Windows (Everything's "Options → General → Start Everything on system startup") is recommended: wider coverage and faster.

## Settings
Preferences → File Search (how a search starts, file types, number of results) and Search Sources (Everything, built-in index), or `ALTRun.json` → `Features.FileSearch`:

| Setting | Default | Description |
|---|---|---|
| SpacePrefix | 1 | Pressing Space in an empty search box starts file search mode |
| InDefaultResults | 0 | Also show matching files and folders in normal searches |
| DefaultResultsLimit | 6 | Maximum number of files in normal results |
| MinQueryLength | 2 | Minimum number of letters before files are searched |
| Keywords | `open, find` | Keywords that search files only |
| FolderKeywords | `folder` | Keywords that search folders only |
| TypeFilters | doc / pic / video / audio / zip / exe / cad | File type filters, one `keyword = ext ...` per line |
| QuotePrefix | 1 | Input starting with `'` searches files and folders only |
| MaxResults | 30 | Maximum number of results when searching files only |
| UseEverything | 1 | Use Everything when it's running |
| EverythingFilter | `!C:\Windows\ !\AppData\ !\$Recycle.Bin\` | Conditions added to every Everything search; `!` excludes (Everything search syntax) |
| EverythingPath | empty | Location of Everything.exe, only used for "Search in Everything"; leave empty to find it automatically |
| ScopeFolders | Desktop, Documents, Downloads | Folders the built-in index scans; [path variables](en-Commands-and-Snippets#variables-in-paths) work |
| ScopeDepth | 4 | Subfolder depth of the built-in index |
| ScopeExclude | `node_modules`, `.git`, `__pycache__`, Recycle Bin | Folders the built-in index skips (regular expression) |
| MaxEntries | 30000 | Maximum number of items in the built-in index |
| RefreshMinutes | 30 | How often the built-in index is refreshed |

After changing what the built-in index covers, click "Rebuild File Index" in Preferences to apply it right away.

## Troubleshooting
- **Everything is installed but not used**: Everything must be running (its icon is in the notification area). If Everything runs as administrator and ALTRun doesn't, Windows may block the communication between them; run both with the same rights (installing Everything as a service and running it normally is recommended).
- **Files on network drives**: Everything doesn't index network drives by default; add them in Everything's "Options → Indexes → Folders". The built-in index can include network folders in ScopeFolders, but scanning is slower. Network folders you use often work better as [custom commands](en-Commands-and-Snippets).
