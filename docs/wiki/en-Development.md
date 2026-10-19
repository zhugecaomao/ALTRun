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
Src\Extensions\     Features outside the search window (SnippetExpander, QuickSwitch, AutoDate, CapsLockSwitch, TendonProfile + PTToolsWindow, UpdateChecker)
Resources\          Data shipped with the program (Kanji.txt, Themes\*.json, Lang\*.json, Icons\*.ico)
Tests\RunTests.ahk  Unit tests
Tests\Fixtures\     Test data: an old ALTRun.ini, SPF2M reference data
Tests\Screenshots\  Automatic screenshots
Tests\Bench\       Benchmark: simulates daily use on Windows and compares two versions (see "Benchmark" below)
Tests\Tools\SPF2M\  Python scripts that generate the SPF2M reference data (run the original SPF2M.EXE in DOSBox; see the README there)
bucket\ packaging\  Scoop / winget manifests (see packaging\README.md)
docs\               Wiki sources (docs\wiki, published after merging) and screenshots (docs\images)
site\               Website template and build script (published to GitHub Pages)
.github\            GitHub Actions (tests, screenshots, release, wiki, website, Gitee sync), issue / PR templates
```

When it runs, ALTRun creates `Data\` in its folder (the settings file `ALTRun.json`, plus indexes and history that can be deleted) and `Themes\` (custom themes).

## How a search runs
1. The `SearchWindow` input changes → `ProviderRegistry.Search(text)`
2. The text is parsed into a `SearchQuery` (keyword, prefix, remaining text)
3. Each enabled provider's `Search(query)` returns an array of `ResultItem`s
4. `ProviderRegistry` merges the results, adds the learned bonus from `Knowledge`, sorts by score and keeps one result per `Uid`; if there are `Exclusive` results only those are shown; if there are none at all, WebSearch's fallback items are shown
5. `SearchWindow` draws every row itself with NM_CUSTOMDRAW
6. When a result runs, `ActionCatalog` decides how to open it from its `Kind`, and `Knowledge.Record()` remembers the choice

## Core modules and runtime lifecycle

- `App` owns startup, reload and shutdown. `AppSettings` reads and saves settings; `SchemaMigration` upgrades older settings formats.
- `SearchWindow` receives input; `SearchQuery` parses keywords and prefixes; `ProviderRegistry` calls enabled providers and combines their results.
- `ResultItem` defines the common result model; `FuzzyMatcher` scores matches; `Knowledge` adjusts learned ranking; `ActionCatalog` performs actions such as opening, copying and running.
- `PreferencesWindow` edits a working copy of the settings. Saving reloads the application so controls don't mutate the live settings object directly.
- `ApplicationProvider` and `FileIndex` share `Lib\FolderWatcher.ahk` to watch indexed folders. Index rebuilds run in the background, and the previous index remains usable until the new one is ready.

A typical lifecycle is: start and load settings → initialize providers and windows → receive search input → parse the query and collect results → rank and display them → execute the selected action. Settings changes are saved and then reloaded. Shutdown stops folder watchers and background work and handles pending cache saves.

Background work must respect window and resource lifetimes. Timer callbacks are wrapped with `Logger.Watch(label, fn)` for timing; network requests use asynchronous `HttpGet.Start` rather than blocking the main thread; index caches are normally saved after keyboard and mouse idle, with pending data handled immediately at shutdown.

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

## Translating the interface
The English texts live in `Src\Core\I18n.ahk` (`s["Prefs.Page.General"] := "General"`) and define every key; other languages are `Resources\Lang\<code>.json` (UTF-8) with the same keys and translated values. A key missing from a file is looked up in the `_fallback` language first (for example Traditional Chinese falls back to Simplified Chinese), then in English.

| Special key | Purpose |
|---|---|
| `_name` | Name shown in Preferences → General → Language, written in that language (for example `日本語`) |
| `_font` | Interface font (used when the theme's `FontName` is `auto`), for example `Microsoft YaHei UI`, `Yu Gothic UI` |
| `_fallback` | Optional: which language to try first for missing keys |
| `_comment` | A note for translators; the program ignores it |

When translating: keep the same number of `{1}` `{2}` placeholders as the English text (their order may change); `\n` is a line break; keep Preferences descriptions short, Japanese overflows the page width most easily.

**Adding a new interface text**: add the English text to `I18n.ahk` and the same key with its translation to all three JSON files.

**Adding a language** (for example Korean):
1. Copy `zh-CN.json` to `Resources\Lang\ko.json`, set `_name` (`한국어`) and `_font` (`Malgun Gothic`), and translate every entry
2. No code changes: the language appears in Preferences automatically; "Auto" follows the Windows display language, and region codes (for example `zh-HK` → `zh-TW`) are mapped in `I18n.Aliases`
3. Run the tests: `I18nLanguages` checks that every language file has no missing or extra keys, no empty values, matching placeholder counts and a `_font`; `PreferencesFit` / `PreferencesSpacing` check that every page fits and isn't crowded in that language; `I18nUnused` checks the English file has no unused keys
4. The Wiki is Chinese and English only: help opens the Chinese pages for Simplified and Traditional Chinese, the English pages for every other language

## Release process
Detailed steps are in the [contributing guide](https://github.com/zhugecaomao/ALTRun/blob/main/CONTRIBUTING.md#releasing-a-new-version); in short:

- **Test builds** (every PR that changes the program): a four-part version `YYYY.MM.DD.N` in `App.Version` and `;@Ahk2Exe-SetVersion` (the `ReleaseVersion` test checks they match) and a `CHANGELOG.md` section (the `ChangelogLinks` test checks the compare links); run the **Release** workflow on the PR branch with `publish = false`, and the build is under Artifacts at the bottom of the Actions page (`ALTRun_v<version>`)
- **Releases**: a three-part version `YYYY.MM.DD`, at most one per day; merge the round's test build sections of the CHANGELOG into one (Chinese and English summary + upgrade / new / improved / fixed), then after merging to main run **Release** on main with `publish = true`

On Windows the Release workflow: reads the version → downloads AutoHotkey and Ahk2Exe → runs the unit tests → compiles `ALTRun.exe` → signs it when SignPath is configured → checks the program icon → runs the upgrade test with a 2.x `ALTRun.ini` → runs the single instance test on the compiled program → packages `ALTRun_v<version>.zip`. With `publish = true` it then: takes that CHANGELOG section as the release notes → creates the GitHub release (GitHub provides the SHA256 that one-key updates verify) → updates the Scoop and winget manifests in `packaging/` and commits them to main → submits the new version to winget-pkgs when `WINGET_TOKEN` is configured.

## Benchmark
GitHub → Actions → **Bench** → Run workflow: on a GitHub Windows VM, simulates daily use with real keystrokes and compares
two versions (baseline: the latest release by default; candidate: the selected branch by default). Each version runs two rounds,
alternating (A B B A), with ALTRun restarted for every round:

- Random searches (start of app names, word initials, random letters, file search, custom commands): time to show the window, and from the last key to drawn results
- Running programs / opening files and folders, the action panel (→), the calculator (result + Enter to copy), Quick Switch (folder panel + Ctrl+E),
  opening and saving Preferences, snippet expansion
- ALTRun's CPU time, memory, GDI / USER objects and handles at the end of each round
- Background tasks over 50 ms (`Perf: slow`, in versions that log them) and errors from the ALTRun log

By default each round has 500 searches + 100 of each other operation (about 50 minutes); the results (median / p90 / p99 / max) are in the run summary, and the raw data is an artifact.
The test copies #Include one extra file, `Tests\Bench\BenchProbe.ahk`, which tells the driver (`BenchDriver.ahk`) whether the results for the current input have been drawn;
the released program doesn't contain it. Pushing changes to the benchmark scripts to a branch other than main runs a short smoke test automatically.

## Quality checks and limitations

- `Tests\RunTests.ahk` covers pure logic such as matching, query parsing, settings migration, ranking, clipboard and indexes; it is not a full manual UI test.
- Windows CI runs unit tests and syntax checks, and checks Preferences layout, Wiki page pairs and documentation links.
- Screenshot tests check major windows across languages and themes; they do not verify every interaction.
- Performance benchmarks simulate common operations on Windows and compare before / after timings; compare runs in the same environment.
- Wiki automation detects missing pages and broken links, but cannot tell whether explanations are accurate or cover every setting. When settings change, check the code and both language versions item by item.

## Code style
See [CONTRIBUTING.md](https://github.com/zhugecaomao/ALTRun/blob/main/CONTRIBUTING.md) in the repository. A few AutoHotkey v2 pitfalls:
- Names are case-insensitive: don't give a local variable the name of a class (`pinyin` hides the `Pinyin` class), and don't let a method and a property in the same class differ only in case
- A very large `static X := Map(...)` fails with "Declaration too long"; build it inside a method instead
- Strings passed to functions by value are copied whole: passing a big string (a whole file) to a function over and over in a loop makes the time grow with the square of its length; pass it by reference (`&text`), see `Lib\JSON.ahk`
- Lines starting with `Perf:` in the debug log (`General.SaveLog`) are the timings of startup phases, slow searches and slow background tasks; use `Logger.Ms()` / `Logger.Time(label, start)` to record new ones, and wrap timer callbacks in `Logger.Watch(label, fn)` (logged only over 50 ms)
- Don't download with `Download()` on the main thread (the whole program stops until it finishes); use `HttpGet.Start(url, done, fail)` to download in the background
- To learn about changes in a folder, use `FolderWatcher` (`Lib\FolderWatcher.ahk`, call `Poll()` on a timer, it doesn't wait) instead of rescanning the folder on a timer; when closing it waits for the pending read to finish before releasing the buffer
- When a nested function changes an outer variable, pass the result back through an object (`state := {Result: ""}`)
- Give every `Edit` control a row count (`r1 -Multi`), otherwise long text turns it into a multi-line box
- Treat every `#Warn All` warning as an error
