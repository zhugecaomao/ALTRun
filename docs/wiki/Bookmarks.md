**English** · [中文](Bookmarks-zh)

# Browser bookmarks

Search browser bookmarks directly in ALTRun, without opening the browser and browsing its bookmarks first.

## Search and open

- Type `bm keyword` to search bookmarks only. `bm` alone shows a hint; `Enter` on it completes `bm `.
- When you type a bookmark name directly, matching bookmarks can also appear in normal results, after apps and custom commands, with up to 8 bookmarks shown. The URL's domain is also matched, so typing a site domain can find its bookmarks.
- Press `Enter` to open a bookmark; press `Alt+Enter` to copy its URL.
- Supported browsers: Chrome, Edge, Brave and Vivaldi, across all user profiles. Firefox stores bookmarks in an SQLite database and is not supported yet.

ALTRun reads bookmarks at startup, then checks their files in the background every 30 seconds and reloads them when they change. This avoids extra reads while you type.

## Settings

Enable or disable **Browser bookmarks** in Preferences → Features. Under **Feature options**, configure:

| Setting | Default | Description |
|---|---|---|
| `Features.Bookmarks.Enabled` | `1` | Whether bookmark search is enabled |
| `Features.Bookmarks.Keyword` | `bm` | Keyword that searches bookmarks only |
| `Features.Bookmarks.InDefaultResults` | `1` | Also show bookmarks in normal search results |

For more search syntax, see [Search and shortcuts](Usage#search-syntax).
