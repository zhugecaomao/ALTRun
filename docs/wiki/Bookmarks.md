**English** · [中文](Bookmarks-zh)

# Browser bookmarks

Search browser bookmarks directly in ALTRun, without opening the browser and browsing its bookmarks first.

## Search and open

- Type `bm keyword` to search bookmarks only. `bm` alone shows a hint; `Enter` on it completes `bm `.
- When you type a bookmark name directly, matching bookmarks can also appear in normal results, after apps and custom commands, with up to 8 bookmarks shown. The URL's domain is also matched, so typing a site domain can find its bookmarks.
- Press `Enter` to open a bookmark; press `Alt+Enter` to copy its URL.
- Supported browsers: Chrome, Edge, Brave and Vivaldi, across all user profiles. Firefox stores bookmarks in an SQLite database and is not supported yet.

ALTRun reads bookmarks at startup, then checks their files in the background every 30 seconds and reloads them when they change. This avoids extra reads while you type.

## Browser history

- Type `h keyword` to search the browser history: pages whose title or URL contains every word you type, most recently visited first. Chinese titles can also be found by their pinyin initials. `h ` alone (with the space) lists the most recently visited pages; `h` alone shows a hint.
- Press `Enter` to open the page in your default browser; press `Alt+Enter` to copy its URL.
- Only with the keyword: your history never shows up in normal search results.
- Same browsers as bookmarks: Chrome, Edge, Brave and Vivaldi, all profiles (the most recent 10,000 pages of each); Firefox is not supported yet.
- ALTRun reads each profile's `History` database only when you type `h `, and again when you come back to it more than a minute later; nothing is read at startup or in the background. It reads the file in place without locking it, through SQLite that comes with Windows 10 / 11 (`winsqlite3.dll`), so the browser can stay open. ALTRun keeps no copy of your history and sends nothing anywhere; a page you open from the results is remembered like any other result (learned ranking, recent items).

## Settings

Enable or disable **Browser bookmarks** and **Browser history** in Preferences → Features. Under **Feature options**, configure:

| Setting | Default | Description |
|---|---|---|
| `Features.Bookmarks.Enabled` | `1` | Whether bookmark search is enabled |
| `Features.Bookmarks.Keyword` | `bm` | Keyword that searches bookmarks only |
| `Features.Bookmarks.InDefaultResults` | `1` | Also show bookmarks in normal search results |
| `Features.BrowserHistory.Enabled` | `1` | Whether browser history search is enabled |
| `Features.BrowserHistory.Keyword` | `h` | Keyword that searches the browser history |

For more search syntax, see [Search and shortcuts](Usage#search-syntax).
