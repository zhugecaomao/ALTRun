**English** · [中文](Web-Search)

# Web search

Type "keyword search terms" to search a website, for example `g autohotkey hotkeys` searches Google; `Enter` opens it in your default browser.

![Web search: g autohotkey v2 hotkeys](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/websearch.png)

## Built-in search engines
| Keyword | Site |
|---|---|
| `g` | Google |
| `bing` | Bing |
| `bd` | Baidu |
| `gh` | GitHub |
| `wiki` | Wikipedia (English) |
| `yt` | YouTube |
| `tb` | Taobao |
| `jd` | JD.com |
| `tr` | Google Translate (into Simplified Chinese) |

- Once the keyword is followed by a space (for example `g `), only that site shows in the results, not mixed with apps and files
- Typing only the keyword (`g`) shows "Type after 'g '" and `Enter` completes it to `g `; typing the site name (for example `google`) shows the hint too
- Typing only the keyword (`g`) searches the **text on the clipboard**: the built-in sites all have `{query|clipboard}` in their URL, the list shows "Search Google for clipboard: '...'" and `Enter` opens it; if the clipboard holds an image, files or no text it says "The clipboard has no text" and can't be opened. If you typed search words, those are searched
- Select text in any program and press `Ctrl+Alt+\`: the action panel can search the selection on every site, see [Actions on selected content](en-Usage#actions-on-selected-content)

## Adding and editing
Preferences → Web Search: double-click an engine in the list to edit it, or click "Add". Each engine has these fields:

| Field | Description |
|---|---|
| Title | Shown in the result, for example "Search Google for ..." |
| Keyword | Type it followed by a space to search this site. Several keywords: separate them with commas, for example `g, ㅎ, google` (handy when you switch keyboard layouts: any of them works); hints and completion use the one you typed |
| URL | The address of the results page; `{query}` is replaced by what you typed (encoded). With `{query|default}`, typing only the keyword opens it straight away, for example `http://localhost:{query|8000}`; with the default `clipboard` (`https://www.google.com/search?q={query|clipboard}`), typing only the keyword searches the text on the clipboard |
| Id | A name (Latin letters) that identifies the engine, used by fallback searches |
| Icon | Optional: an `.ico`, `.png` or `.exe` file; a relative path starts from the Data folder. **Download Site Icon** visits the website in the URL once, saves its icon in `Data\Icons` and fills this in. Empty = the built-in icon (the engines above have one) or the browser icon |

Examples:

| Title | Keyword | URL |
|---|---|---|
| Merriam-Webster | `dict` | `https://www.merriam-webster.com/dictionary/{query}` |
| Amazon | `amz` | `https://www.amazon.com/s?k={query}` |
| Local dev server | `lh` | `http://localhost:{query|8000}` |

### Searching the clipboard
- `{query|clipboard}` reads the clipboard only when you typed the engine's keyword and this item shows in the list; the completion hint when you type only the engine name (for example `google`) doesn't read it
- Only text is used: whitespace is merged into one space, at most 200 characters; images and files are not searched
- The full address is shown under the title before you press `Enter`, so you can see what will be sent; the clipboard may hold passwords or other sensitive content, so check before pressing `Enter`
- Engines that existed before the upgrade are not changed; to get this, change `{query}` to `{query|clipboard}` in the URL
- In the clipboard history, select a text item and press `→` to search it with each site

## Fallback searches
When nothing matches, the list shows a few fallback items, by default "Search Google for '...'", "Search files" and "Search Bing for '...'". Change them in Preferences → Web Search → "Fallback searches": engine Ids or `files` (file search), separated by commas.

## Settings file
`Features.WebSearch` → `Engines` (the engine list, each with `Id` / `Keyword` / `Title` / `Url` / `Icon`) and `Fallbacks`, see [Settings reference](en-Configuration#websearch). The whole feature can be turned off in Preferences → Features.
