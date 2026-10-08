**English** · [中文](Development)

# Development guide

## Environment
- [AutoHotkey v2](https://www.autohotkey.com/) 2.0 or later
- Any editor; VS Code with the AutoHotkey v2 Language Support extension is recommended

```
AutoHotkey64.exe ALTRun.ahk                                  run
AutoHotkey64.exe /ErrorStdOut /validate ALTRun.ahk           check syntax and warnings only, don't run
AutoHotkey64.exe /ErrorStdOut Tests\RunTests.ahk             unit tests; exit code = number of failures
```

## Project structure
```
ALTRun.ahk          Entry point: #Includes every module and calls App.Start()
Lib\                Shared libraries with nothing ALTRun-specific, usable in other projects as is
                    (JSON, Logger, Util: Path/Win/Url..., TextTools, Kanji, Dialogs, Everything IPC)
Src\Core\           App (startup), AppSettings + SchemaMigration (settings and version upgrades),
                    SearchQuery / ResultItem (search model), FuzzyMatcher (match scoring),
                    Knowledge (learned ranking), Usage (usage statistics), ActionCatalog (actions), ProviderRegistry, FileIndex
Src\UI\             SearchWindow, PreferencesWindow, ItemEditor, LargeType, ThemeManager, IconCache
Src\Providers\      Search features, one class per feature
Src\Extensions\     Features outside the search window (SnippetExpander, QuickSwitch, AutoDate, TendonProfile + PTToolsWindow, UpdateChecker)
Resources\          Data shipped with the program (Kanji.txt, Themes\*.json, Lang\*.json, Icons\*.ico)
Tests\RunTests.ahk  Unit tests
Tests\Fixtures\     Test data: an old ALTRun.ini, SPF2M reference data
Tests\Screenshots\  Automatic screenshots
Tests\Tools\SPF2M\  Python scripts that generate the SPF2M reference data (run the original SPF2M.EXE in DOSBox; see the README there)
bucket\ packaging\  Scoop / winget manifests (see packaging\README.md)
```

## How a search runs
1. The `SearchWindow` input changes → `ProviderRegistry.Search(text)`
2. The text is parsed into a `SearchQuery` (keyword, prefix, remaining text)
3. Each enabled provider's `Search(query)` returns an array of `ResultItem`s
4. `ProviderRegistry` merges the results, adds the learned bonus from `Knowledge`, sorts by score and keeps one result per `Uid`; if there are `Exclusive` results only those are shown; if there are none at all, WebSearch's fallback items are shown
5. `SearchWindow` draws every row itself with NM_CUSTOMDRAW
6. When a result runs, `ActionCatalog` decides how to open it from its `Kind`, and `Knowledge.Record()` remembers the choice

## Adding a search feature
1. Create `XxxProvider.ahk` in `Src\Providers\`:
   ```autohotkey
   class XxxProvider {
       static Id := "Xxx"                  ; also the key under Features in ALTRun.json

       static Init() {                     ; called once at startup (build an index and so on)
       }

       static Search(query) {              ; query: SearchQuery
           results := []
           score := FuzzyMatcher.Score(query.Text, "Some Title")
           if (score > 0)
               results.Push(ResultItem("Some Title", "subtitle", {Kind: "url", Arg: "https://...", Score: score}))
           return results
       }
   }
   ```
2. `#Include` it in `ALTRun.ahk` and call `ProviderRegistry.Register(XxxProvider)` in `App.Start()`
3. Add `"Xxx", Map("Enabled", 1, ...)` under `Features` in `AppSettings.Defaults()`
4. UI text: English in `I18n.ahk`, translations in every language file in `Resources\Lang\` (a test checks that every file has them)
5. Add tests in `Tests\RunTests.ahk`

Optional:
- `EditItem(item)` / `DeleteItem(item)`: support `F3` editing and `Ctrl+Del` deleting in the results; the result's `Source` must point to the matching entry in the settings
- `DeletePrompt(item)`: the confirmation text before deleting

### ResultItem fields
| Field | Description |
|---|---|
| Title / Subtitle | The two lines of text |
| Icon | A file path (its icon is used), `folder:`, `url:`, `res:imageres.dll,-5314` |
| Kind | `file` / `folder` / `url` / `text` / empty; decides the default action and the action panel |
| Arg / Arguments | Target and command-line arguments |
| OnRun | A custom run function, takes precedence over Kind |
| Actions | Extra action panel items |
| Uid | Unique id for learned ranking and de-duplication |
| Score | Score; higher comes first (match score 0–100 plus each feature's own offset) |
| Valid / AutoComplete | With `Valid = false`, Enter only puts AutoComplete into the input box |
| Exclusive | Keyword mode: only results with this flag are shown |
| Source / Provider | The matching entry in the settings / the Id of the provider that produced the result (filled in by ProviderRegistry) |

## Changing the settings structure
- Only adding new settings: add them to `AppSettings.Defaults()`; entries missing from older files are filled in automatically
- Renaming, moving or restructuring existing entries: `AppSettings.CurrentVersion + 1`, and add a `_FromN(data)` to `SchemaMigration` that converts version N to N+1. A backup `ALTRun.vN.backup.json` is made automatically before upgrading

## Documentation
The wiki sources are in `docs/wiki/` in the repository and are published to the wiki automatically after merging to `main` (`.github/workflows/wiki.yml`). Change `docs/wiki/` through pull requests; don't edit the wiki on the web.

The documentation is bilingual; when you change one language, change the other one too:
- README and contributing guide: `README.md` / `CONTRIBUTING.md` (English, shown by default on GitHub) and `README.zh-CN.md` / `CONTRIBUTING.zh-CN.md` (Chinese), linking to each other at the top
- Wiki: two files per page, Chinese `Usage.md` and English `en-Usage.md`; each page links to the other language at the top, and links between English pages use the `en-` page names. The sidebar `_Sidebar.md` has an English and a Chinese group. Help opened from the program (cheat sheet, `F1` in Preferences) follows the UI language: the Chinese page for a Chinese UI, the `en-` page otherwise (`HelpProvider.WikiPage`). The `WikiPages` test checks that every page has an English page and that links point to existing pages
- CHANGELOG is written in Chinese, with an English sentence after each version's summary (that section is the release notes)

The website (https://zhugecaomao.github.io/ALTRun/) sources are in `site/`: `index.html` is a Chinese / English template, and `build.py` fills in the latest version, download link and size, download count, theme screenshots and the latest release notes, writing to `_site/`. `.github/workflows/pages.yml` builds it when related files change on main, after a successful release and once a day, and deploys it with the official GitHub Pages actions (Settings → Pages → Source: "GitHub Actions"). Local preview: `python3 site/build.py && python3 -m http.server -d _site`.

Screenshots of the UI are in `docs/images/screenshots/` and are generated by `Tests\Screenshots\TakeScreenshots.ahk`: it prepares a demo ALTRun in a temporary folder (English UI; the sample settings, custom commands and files are made-up generic content with nothing personal or work-related), then starts it for each scene, types and takes a screenshot. After UI changes, run **Screenshots** in Actions (`.github/workflows/screenshots.yml`); it retakes the screenshots on Windows and commits them back to the current branch. Pushes that change `Tests/Screenshots/` run it automatically too. To run it locally:
```
AutoHotkey64.exe Tests\Screenshots\TakeScreenshots.ahk [output folder] [scene names...]
```

## Code style
See [CONTRIBUTING.md](https://github.com/zhugecaomao/ALTRun/blob/main/CONTRIBUTING.md) in the repository. A few AutoHotkey v2 pitfalls:
- Names are case-insensitive: don't give a local variable the name of a class (`pinyin` hides the `Pinyin` class), and don't let a method and a property in the same class differ only in case
- A very large `static X := Map(...)` fails with "Declaration too long"; build it inside a method instead
- Strings passed to functions by value are copied whole: passing a big string (a whole file) to a function over and over in a loop makes the time grow with the square of its length; pass it by reference (`&text`), see `Lib\JSON.ahk`
- Lines starting with `Perf:` in the debug log (`General.SaveLog`) are the timings of startup phases and slow searches; use `Logger.Ms()` / `Logger.Time(label, start)` to record new ones
- When a nested function changes an outer variable, pass the result back through an object (`state := {Result: ""}`)
- Give every `Edit` control a row count (`r1 -Multi`), otherwise long text turns it into a multi-line box
- Treat every `#Warn All` warning as an error
