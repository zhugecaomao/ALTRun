**English** · [中文](File-Search-zh)

# File search

## How to search for files
Normal searches show apps, commands, snippets and so on, but **not** files from all over your disks (so unrelated files such as `NIRMALA.TTF` don't get mixed in). To search for files:

- **Space**: in an empty search box, press `Space` first; the input box shows a gray "Search files...", then type the name. If you've already typed something, move the cursor to the start (`Home`) and press `Space` to search files for it. `Backspace` in an empty input box or at the start goes back to a normal search. Like Alfred's Quick File Search
- **`file report`**, **`open report`**, **`find report`** or **`'report`**: same effect
- **Folders only**: **`folder budget`** (change the keyword in Preferences → File Search → "Folder keywords"); `folder budget` also works in file search mode. With Everything this is the same as Everything's `folder:budget`

- **Recently opened files**: **`recent report`** searches the files and folders you opened recently in any program (Windows' "Recent items"), newest first, on local disks and network shares alike; **`recent `** with nothing after it lists the newest. It reads Windows' Recent folder directly, so it works without Everything and without the built-in index (Everything doesn't list these by default: `\AppData\` is excluded in Preferences → Search Sources → "Filter"). Change the keyword in Preferences → File Search → "Recent keywords"; `recent report` also works in file search mode. Network files are listed without accessing the network, so a slow or disconnected share doesn't hold up the list
- **By type** (like Listary's filters): **`doc report`** searches documents only; there are also `pic` pictures, `video`, `audio`, `zip` archives, `exe` programs and `cad` (dwg / dxf / dgn / rvt / ifc / skp...). `doc report` also works in file search mode. Change the keywords and extensions in Preferences → File Search → "File types" → **Edit...**, one `keyword = ext ext ...` per line (the line shows just the keywords). With Everything this is the same as `ext:doc;docx;...`
- To find files by modification date, write Everything's syntax directly, for example `Space` + `dm:today report` or `dm:thisweek`

A keyword needs a space after it to start a file search: typing just `folder`, `open`, `recent` or `doc` shows commands and apps with that word in their name as usual, plus a hint for the keyword. Apps and commands whose name starts with that word (e.g. Docker for `doc`) stay above the hint, so `Enter` still opens them; otherwise the hint comes first. `Enter`, `Tab` or a double-click on the hint completes the keyword and the space (e.g. `recent `). While a keyword isn't typed in full yet (at least 3 letters, e.g. `fol`, `rec`, `vid`), its hint shows at the end of the list.

Results are sorted by how well the name matches: exact > start of the name > start of a word (for example `2026-Budget`) > contains (for example `NewBudget`); at the same match quality, folders come before files, then by modification time (newest first). With Everything, ALTRun takes Everything's first 300 results and then sorts them, so the best-matching folder isn't missed just because it wasn't modified recently.

Everything's search syntax can be used as is and is passed straight to Everything: for example `Space` + `folder:budget 2024`, `ext:pdf report` or `path:Projects report`.

At most 30 results are shown; when nothing is found you can continue in Everything or Windows Search with one key. When a normal search finds nothing at all, a "Search files" fallback item is offered at the end as well.

To show files in normal searches too (the old behavior): in Preferences → File Search, tick "Also show matching files and folders in normal results"; at most 6 are shown, and only those whose name or a word in it starts with what you typed.

Folders you use often work better as [custom commands](Custom-Commands): custom commands also match their target's folder name, so a normal search finds them.

![File search](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/files.png)

![Search by file type: doc report](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/filetype.png)

`Enter` opens, `Ctrl+Enter` reveals in the file manager, `Alt+Enter` copies the path, `→` shows more actions (Open With, copy / cut the file, copy / move to the current folder in TC, Open Terminal Here, Properties, Move to Recycle Bin...), see the [action panel](Usage#action-panel).

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

`Tab` goes into the selected folder and `Backspace` removes one part of the path to go up a level; a folder in any result (custom command, file search...) can be entered with `Tab` too. `Enter` opens, and `→` has all the file actions. Hidden and system files aren't listed. `Insert` marks several files to copy, move or delete them together, see [shortcuts](Usage#shortcuts).

## Where results come from
ALTRun picks automatically:

| Situation | What is searched |
|---|---|
| [Everything](https://www.voidtools.com/) is running | Queries **all disks** directly through Everything's IPC interface, instantly; no `Everything64.dll` or `es.exe` needed |
| Everything isn't running | The built-in index: scans Desktop, Documents and Downloads in the background (4 levels deep, at most 30,000 items), cached in `Data\FileIndex.json`. When files or folders are added, removed or renamed in those folders, it is rescanned in the background a few seconds later (at most once every 5 minutes; edits to file contents and browser downloads in progress don't count). Changing the indexed folders, depth, exclusions or max entries in Preferences also rescans in the background after you click Apply (the old index keeps working until then). There is also a fallback rescan every 60 minutes by default. If Everything is running when ALTRun starts, the built-in index isn't read or scanned until Everything is closed and it's actually needed |

The top of Preferences → Search Sources shows whether Everything is running right now.

Installing Everything and letting it start with Windows (Everything's "Options → General → Start Everything on system startup") is recommended: wider coverage and faster.

## Settings

![Preferences → Search Sources](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-searchsources.png)

Preferences → File Search (how a search starts, file types, number of results) and Search Sources (Everything, built-in index), or `ALTRun.json` → `Features.FileSearch`:

| Setting | Default | Description |
|---|---|---|
| SpacePrefix | 1 | Pressing Space in an empty search box starts file search mode |
| InDefaultResults | 0 | Also show matching files and folders in normal searches |
| DefaultResultsLimit | 6 | Maximum number of files in normal results |
| MinQueryLength | 2 | Minimum number of letters before files are searched |
| Keywords | `open, find, file` | Keywords that search files only |
| FolderKeywords | `folder` | Keywords that search folders only |
| RecentKeywords | `recent` | Keywords that search recently opened files and folders (Windows' Recent items) |
| TypeFilters | doc / pic / video / audio / zip / exe / cad | File type filters, one `keyword = ext ...` per line |
| QuotePrefix | 1 | Input starting with `'` searches files and folders only |
| MaxResults | 30 | Maximum number of results when searching files only |
| UseEverything | 1 | Use Everything when it's running |
| EverythingFilter | `!C:\Windows\ !\AppData\ !\$Recycle.Bin\` | Conditions added to every Everything search; `!` excludes (Everything search syntax) |
| EverythingPath | empty | Location of Everything.exe, only used for "Search in Everything"; leave empty to find it automatically |
| ScopeFolders | Desktop, Documents, Downloads | Folders the built-in index scans; [path variables](Custom-Commands#variables-in-paths) work |
| ScopeDepth | 4 | Subfolder depth of the built-in index |
| ScopeExclude | `node_modules`, `.git`, `.svn`, `.hg`, `__pycache__`, `.venv`, `venv`, `.tox`, `.cache`, `.idea`, `.vs`, Recycle Bin | Folders the built-in index skips (regular expression, matched against the full path). By default it skips code repositories, Python virtual environments and developer tool caches: they often hold tens of thousands of files that would fill up MaxEntries, and none of them are documents you look for |
| MaxEntries | 30000 | Maximum number of items in the built-in index |
| RefreshMinutes | 60 | Fallback: the built-in index is also rescanned in the background when its cache is older than this (at least 5 minutes; checked at startup and then every minute, rebuilt once the keyboard and mouse have been idle for 30 seconds and the search window is closed). Normally it updates itself by watching the folders; nothing is scanned while Everything is used |

After changing what the built-in index covers, click "Rebuild File Index" in Preferences to apply it right away.

## Troubleshooting
- **Everything is installed but not used**: Everything must be running (its icon is in the notification area). If Everything runs as administrator and ALTRun doesn't, Windows may block the communication between them; run both with the same rights (installing Everything as a service and running it normally is recommended).
- **Files on network drives**: Everything doesn't index network drives by default; add them in Everything's "Options → Indexes → Folders". The built-in index can include network folders in ScopeFolders, but scanning is slower. Network folders you use often work better as [custom commands](Custom-Commands).
- **Searching Chinese file names by pinyin** (for example `cs` finding "测试.bat" or `zjsm` finding "证件扫描文件.zip"): pinyin initials are supported; files found this way are highlighted too and sorted by how well the initials match (at the same match quality, names that match directly come first)
  - **Without Everything (built-in index)**: works out of the box. The initials are worked out in the background after the index loads (a fraction of a second for 10,000 Chinese file names); for that short moment after startup pinyin isn't matched yet. Only a single word is matched: `zj sm` written apart isn't found
  - **With Everything**: Everything does the finding, so make Everything itself match pinyin first (ALTRun passes your input to Everything as is, so nothing else needs to be set):
    - **Everything 1.5 (beta)**: Tools → Options → Advanced, search for `pinyin` and set it to `true` (initials only). If you're still on a 1.5 alpha, ALTRun may not find it (the alpha runs as a named "1.5a" instance by default); please update to the beta
    - **Everything 1.4**: install the open-source extension [IbEverythingExt](https://github.com/Chaoses-Ib/IbEverythingExt), put it in Everything's folder as its instructions say and restart Everything. It supports initials (`cs`), full pinyin (`ceshi`) and shuangpin
    - With several words (for example `zj sm`), files Everything finds are still highlighted but not sorted by pinyin
  - Only common characters (GB2312 level 1) are recognized and a character with several readings uses one of them; full pinyin is only found by Everything + IbEverythingExt, and ALTRun doesn't sort or highlight by full pinyin
