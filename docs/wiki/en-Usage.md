**English** · [中文](Usage)

# Search and shortcuts

## Search window
Press `Alt+Space` to open it and start typing. You can change the hotkey in Preferences → General: click the box and press the new keys. You can also set a "Second hotkey", or open it like Listary with "Double-tap Ctrl" / "Double-tap Shift". Each row shows an icon, a title and one line of details (path, arguments or type), with the `Ctrl+number` shortcut on the right. The window hides itself when it loses focus. If ALTRun is already running, double-clicking `ALTRun.exe` again also opens the search window (it doesn't restart).

Drag the empty space around the input box to move the search window. To have it open where you dragged it next time: Preferences → Appearance → "Remember the position after dragging the window" (the position is remembered relative to the screen, so it also works on another monitor; "Reset Window Position" goes back to centered horizontally, 20% from the top). With more than one screen, "Show the window on" can be the screen with the mouse pointer (default), the main screen, or the screen with the active window. Short notices after an action (such as "Copied") appear on the screen with the search window; when the search window isn't open, the same setting picks the screen.

To have the search window remember your last search (like "keep input" in 2.x): Preferences → Search Window → "Keep the last search when the window opens". With it on, the last input and results are still there when you open the window, with the text selected: press `Enter` to open it again, start typing for a new search, or press `Space` for a file search.

## Pinned and recent items
Before you type anything, the search window lists (pinned items have a pin at the bottom right of their icon):

![Empty search box: pinned and recent items](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/empty.png)
- **Pinned items**: on any result, press `→` to open the action panel and choose **Pin to Empty Search Box**. Pinned items are saved in `ALTRun.json` (`Features.Recent.Pinned`), so they travel with your settings to another PC
- **Recently opened items**: not shown by default (0); set Preferences → Search Window → Empty search box → "Recent items to show" to, for example, 5 to list them

Only results you can open again are remembered: apps, files, folders, URLs, custom commands and system commands; clipboard history, calculations, snippets, web searches and windows are not. Select an item and press `Ctrl+Del` to unpin it or remove it from the recent items; Preferences → Advanced → "Reset Learned Ranking" clears the recent items. The whole feature can be turned off in Preferences → Features.

## Usage tips
Each time you open the search window, the status bar at the bottom shows another usage tip (for example "Tip: folder bk  search folders only"), in turn; once you type, it shows the number of results instead. Once you know them, turn off "Show a usage tip" under "Empty search box" in Preferences → Search Window. If you forget how to write something, type `?` to see everything.

## Status bar
The line at the bottom of the search window (like Listary and Raycast): on the left, the usage tip, "Searching files" in file search mode or the number of results; on the right, what `Enter` does for the selected result (Open, Copy, Paste, Run...) and `Ctrl+K` for its action panel. Turn it off in Preferences → Search Window → "Show the status bar" (`Appearance.StatusBar`); the usage tip then shows in the empty search box as before.

## Search syntax
| Input | Result |
|---|---|
| Any text | Apps, custom commands, snippets, system commands, Windows tools, Windows Settings pages (for example `bluetooth`, `default apps`) |
| `Space`, then type (press Space in an empty search box) | Files and folders only; the input box shows "Search files..."; `Backspace` in an empty box goes back |
| `12*(3+4)` or `=2^10` | Calculator; `Enter` copies the result |
| `10 km in mi` `100 f to c` `20 mpa in psi` | Unit conversion; with currency conversion turned on also `100 usd to sgd`, see [Extensions](en-Extensions#unit-and-currency-conversion) |
| `255 in hex` `0xff to dec` `10 in bin` / `0b1010` | Number bases (hex / bin / oct / dec); a number starting with `0x`, `0b` or `0o` on its own lists decimal, hexadecimal and binary |
| `today + 30 days` `2026-12-25 - 2w` / `2026-12-25 - today` | Date arithmetic (units d w m y), with the day of the week; subtracting two dates gives the number of days between them |
| `bm xxx` | Browser bookmarks only (typing a bookmark name without `bm` finds it too, ranked after apps and commands) |
| `w xxx` / `w ` | Switch to an open window (by title or program name; `w ` alone lists them all); `→` can close the window. Windows whose title matches well also show up without `w` (at most 3) |
| `'report` / `open report` / `find report` | Files and folders only, see [File search](en-File-Search) |
| `doc report` `pic logo` `cad plan` | One kind of file only (documents / pictures / CAD...); the kinds can be changed |
| `folder bk` | Folders only |
| `?` / `? file` | Cheat sheet: every search syntax and shortcut (filtered); `Enter` opens the matching wiki page |
| `g xxx` `bing xxx` `bd xxx` `gh xxx` `wiki xxx` `yt xxx` `tb xxx` `jd xxx` `tr xxx` | Web search: Google / Bing / Baidu / GitHub / Wikipedia / YouTube / Taobao / JD.com / Google Translate |
| `>command` | Runs it in a terminal, for example `>ipconfig /all` |
| `clip` / `clip xxx` | Clipboard history: text, files and images (all / filtered) |
| `snip` / `snip xxx` | Snippets only |
| `;keyword` (in any program) | Snippet expansion, see [Custom commands and snippets](en-Commands-and-Snippets) |

When a keyword is followed by a space (for example `clip `, `g xxx`, `'xxx`, `>xxx`), only that feature's results are shown. When nothing matches, fallback items are shown: search Google, search files, search Bing (change them in `Features.WebSearch.Fallbacks`).

![Web search: g autohotkey v2 hotkeys](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/websearch.png)

### How matching works
Case-insensitive, scored in this order:

| Rule | Example |
|---|---|
| Exact match | `notepad` → Notepad |
| Same beginning | `note` → Notepad |
| Beginning of a word | `code` → Visual Studio Code |
| Word initials | `vsc` → Visual Studio Code |
| Pinyin initials | `wx` → 微信 (WeChat) |
| Contains | `pad` → Notepad |
| Several keywords | `st co` → Visual Studio Code |
| Letters in order (at least 3, same first letter) | `ntpd` → Notepad |

![Pinyin initials: jsb → 记事本 (Notepad)](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/pinyin.png)

The characters in a result title that match your input are highlighted (consecutive characters, word initials, several keywords, letters in order and pinyin initials are all marked); the colors come from the theme's `Highlight` / `SelectedHighlight`, see [Themes](en-Themes). Results matched on something other than the title (a keyword, the path, a snippet's text and so on) have no highlighted title.

Custom commands are also matched on the folder / file name of their target: if the Target is `Q:\Projects\PT1931 - 24 NIR`, typing `nir` finds it.

### Learned ranking
ALTRun remembers "what you typed → which result you finally picked": if you typed `no` and picked Notepad, Notepad comes first the next time you type `no` or `not`; items you use often move up in general. Press `Ctrl+↑` / `Ctrl+↓` to recall the previous / next search. The learned ranking is saved in `Data\Knowledge.json` and can be reset in Preferences → Advanced.

## Shortcuts
| Key | Action |
|---|---|
| `Enter` | Run the selected result (open a file / program / URL, copy a calculation, paste a snippet...) |
| `Ctrl+Enter` | Files / folders: reveal in the file manager; text: paste into the front window |
| `Alt+Enter` | Copy the path / URL / text |
| `Ctrl+1` – `Ctrl+9` | Run the Nth visible row |
| `↑` `↓` / `Ctrl+P` `Ctrl+N` | Select the previous / next row (`Ctrl+P` `Ctrl+N` work like in terminals and Emacs, so your hands stay on the letter keys) |
| `PgUp` `PgDn` | Previous / next page (as many rows as the window shows) |
| `Ctrl+↑` / `Ctrl+↓` | Previous / next search from history (like Alfred); going past the newest restores what you had typed |
| `Tab` | Autocomplete; on a folder, go into it (the input becomes `path\`) and browse its contents |
| `Insert` | Mark / unmark the selected file or folder and move to the next row (like Total Commander); marked rows have a bar on the left. Then press `→` to act on all marked items: open all, copy / cut (paste in TC or Explorer), copy paths, copy / move to the current folder in TC, add to custom commands, move to the Recycle Bin |
| `Space` (empty search box, or cursor at the start) | Switch to file search, keeping what you typed (typed `seismic` first? press `Home`, then `Space`); `Backspace` in an empty box or at the start goes back to a normal search |
| `Space` (after typing) | With Preferences → Search Window → "Space runs the selected result" turned on: runs the selected result; `Shift+Space` types a space (for several keywords) |
| `→` (cursor at the end) / `Ctrl+K` | Open the action panel; `←` / `Esc` (or `Ctrl+K` again) go back. Holding `→` to move the cursor doesn't open it when the cursor reaches the end; press `→` once more |
| `F3` | Edit the selected result, see below |
| `Ctrl+Del` (cursor at the end) | Delete the selected result, see below |
| `Ctrl+C` | Copy the selected result (copies text as usual when text is selected in the input box) |
| `Ctrl+L` | Large type |
| `Ctrl+Backspace` | Delete the previous word |
| `F2` / `Ctrl+,` | Preferences |
| `F4` | Edit ALTRun.json in Notepad (reloaded automatically when saved) |
| `F1` | About ALTRun: the Advanced page of Preferences (version, check for updates, project page) |
| `Esc` | Close the action panel / hide the window |
| `Ctrl+Alt+C` (global) | Clipboard history |

Mouse: click to select, double-click to run, the wheel moves the selection, right-click opens the action menu.

## Action panel
Select a result and press `→` (or right-click) to list everything you can do with it; keep typing to filter:

![Action panel](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/actions.png)

| Result type | Actions |
|---|---|
| Program / file | Open, Run as Administrator (exe / lnk / bat...), Open With..., Reveal in File Manager, Copy Path, Copy Name, Copy File / Cut File, Copy / Move to Current Folder in TC, Open Terminal Here, Properties, Move to Recycle Bin |
| Folder | Open, Open Terminal Here, Reveal in File Manager, Copy Path, copy / cut the folder, Copy / Move to Current Folder in TC, Properties, Move to Recycle Bin |
| URL | Open, Copy URL |
| Text | Copy to Clipboard, Paste to Front Window |
| All | Edit... / Add to Custom Commands... (`F3`), Delete (`Ctrl+Del`), Show in Large Type (`Ctrl+L`) |

File actions (like Listary):
- **Copy File / Cut File**: puts the file itself on the clipboard; press `Ctrl+V` in TC or Explorer to paste it (a cut file is moved when pasted).
- **Copy / Move to Current Folder in TC**: the target is the folder currently open in the frontmost Total Commander (or Explorer) window; not shown when the file is already in that folder. Uses Windows' own copy, with a progress window, a prompt when names clash, and undo.
- **Move to Recycle Bin**: Windows asks first, and you can restore it from the Recycle Bin.

## Actions on selected content
Like Alfred's Universal Actions: select text, files or a URL in any program and press `Ctrl+Alt+\` (change it in Preferences → General → "Selection hotkey"); ALTRun opens the action panel for that content directly, and you can type to filter:

| Selection | Actions |
|---|---|
| Text | Copy, search with each search engine (Google, Baidu, GitHub, Translate...), Save as Snippet, Show in Large Type; an expression shows its result directly; **Replace with** upper case / lower case / title case / sorted / without blank or duplicate lines / Simplified ↔ Traditional / URL-encoded (converted and pasted back into the original program, replacing the selected text) |
| URL | Open, Copy |
| One file or folder (or selected text that is a path) | Open, Run as Administrator, Reveal in File Manager, Copy Path, Open Terminal Here, Properties, Add to Custom Commands |
| Several files | Copy paths, Add All to Custom Commands |

The selection is read by simulating a copy (`Ctrl+C`); the clipboard is restored right after and the copy isn't recorded in the clipboard history. In console / terminal windows `Ctrl+C` would interrupt the running program, so `Ctrl+Insert` is used there. `Esc` or `←` closes the window.

## Editing and deleting in the results
| Selected result | `F3` | `Ctrl+Del` |
|---|---|---|
| Custom command / snippet / search engine | Opens the edit window and returns to your search after saving | Deleted after confirmation |
| Clipboard history | Text: save as a snippet; file: add to custom commands | Removed from the history |
| App | Add as a custom command (with a keyword if you like) | Removed from the search results (not uninstalled); restore it in Preferences → Applications → "Hidden apps" |
| Built-in command (system commands, Windows tools and so on) | - | Removed from the search results, still works from hotkeys; restore it in Preferences → Features → "Hidden commands" |
| File / folder / URL | Add as a custom command | - |
| No results | Create a custom command from the text you typed | - |

## Usage statistics
Preferences → Usage (like Alfred's Usage):
- How many times you used ALTRun today, in the last 7 days, the last 30 days and in total, and how often you opened the search window
- Uses per day for the last 30 days (bar chart)
- Count and share for each feature (apps, custom commands, files, web search, calculator, clipboard history, snippets, system commands, terminal commands, snippet expansion, Quick Switch, date stamp)

Only counts are kept, never what you typed or opened. They're saved in `Data\Usage.json` (the last 400 days). The "Clear Usage Statistics" button starts over.
