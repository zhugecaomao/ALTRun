**English** · [中文](Development-zh)

# Development guide

## Environment
- [AutoHotkey v2](https://www.autohotkey.com/) 2.0 or later. Release builds and the two benchmarks use a pinned version, currently v2.0.31 (`$ahkTag` in `.github/workflows/release.yml`, `bench.yml` and `file-index-bench.yml`; change all three together); the unit tests on CI always use the latest stable release, so a new AutoHotkey version is tested before it is pinned
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
                    Knowledge (learned ranking), Usage (usage statistics), ActionCatalog (running results: open, copy, paste...),
                    ProviderRegistry, FileIndex, BackgroundWork (when background work may run)
Src\UI\             SearchWindow (+ SearchKeys, ResultRows, SearchStatusBar, TypeAhead, SearchInputLanguage),
                    ResultActions (actions panel, edit, delete), PreferencesWindow (+ PrefsPages: the pages,
                    PrefsForm: form helpers and layout constants), ItemEditor, LargeType, ThemeManager, IconCache
Src\Providers\      Search features, one class per feature
Src\Extensions\     Features outside the search window (SnippetExpander, QuickSwitch, AutoDate, CapsLockSwitch, TendonProfile + PTToolsWindow, UpdateChecker)
Resources\          Data shipped with the program (Kanji.txt, Pinyin.txt, Themes\*.json, Lang\*.json, Icons\*.ico)
Tests\RunTests.ahk  Unit test runner: loads the code with in-memory settings and runs every suite
Tests\Suites\       Unit tests, one class per area (Core, Provider, SearchWindow, Preferences, Extension, Project)
Tests\Fixtures\     Test data: an old ALTRun.ini, SPF2M reference data
Tests\Screenshots\  Automatic screenshots
Tests\Bench\       Benchmark: simulates daily use on Windows and compares two versions (see "Benchmark" below)
Tests\Tools\SPF2M\  Python scripts that generate the SPF2M reference data (run the original SPF2M.EXE in DOSBox; see the README there)
bucket\ packaging\  Scoop / winget manifests (see packaging\README.md)
docs\               Wiki sources (docs\wiki, published after merging) and screenshots (docs\images)
site\               Website template and build script (published to GitHub Pages)
.github\            GitHub Actions (tests, screenshots, release, wiki, website, Gitee sync), issue / PR templates
```

Layers: `Lib` uses nothing from `Src`; `Src\Core` uses nothing from `UI`, `Providers` or `Extensions` (only `App.ahk`, which starts everything, does). Where Core needs to know something from above, the upper layer sets a callback or calls Core, for example `BackgroundWork.IsBusy` (set by the search window) and `IconCache.Prefetch(ProviderRegistry.IconSpecs())` (called by `App`). The `Layers` test checks this.

When it runs, ALTRun creates `Data\` in its folder (the settings file `ALTRun.json`, plus indexes and history that can be deleted) and `Themes\` (custom themes).

## How a search runs
1. The `SearchWindow` input changes → `ProviderRegistry.Search(text)`
2. The text is parsed into a `SearchQuery` (keyword, prefix, remaining text)
3. Each enabled provider's `Search(query)` returns an array of `ResultItem`s
4. `ProviderRegistry` merges the results, adds the learned bonus from `Knowledge`, sorts by score and keeps one result per `Uid`; if there are `Exclusive` results only those are shown; if there are none at all, the providers' fallback items are shown (`Fallbacks(query)`, WebSearch)
5. `SearchWindow` draws every row itself with NM_CUSTOMDRAW
6. When a result runs, `ActionCatalog` decides how to open it from its `Kind` (the actions panel, F3 and Ctrl+Del go through `ResultActions`), and `Knowledge.Record()` remembers the choice

## Core modules and runtime lifecycle

- `App` owns startup, reload and shutdown. `AppSettings` reads and saves settings; `SchemaMigration` upgrades older settings formats.
- `SearchWindow` receives input; `SearchQuery` parses keywords and prefixes; `ProviderRegistry` calls enabled providers and combines their results.
- `ResultItem` defines the common result model; `FuzzyMatcher` scores matches; `Knowledge` adjusts learned ranking; `ActionCatalog` performs actions such as opening, copying and running.
- `PreferencesWindow` edits a working copy of the settings. Saving reloads the application so controls don't mutate the live settings object directly.
- `FileSearchProvider` (`recent report`) and `QuickSwitch` (recent folders) read Windows' Recent items through `Lib\WindowsRecent.ahk`.
- `ApplicationProvider` and `FileIndex` share `Lib\FolderWatcher.ahk` to watch indexed folders. Index rebuilds run in the background, and the previous index remains usable until the new one is ready.

A typical lifecycle is: start and load settings → initialize providers and windows → receive search input → parse the query and collect results → rank and display them → execute the selected action. Settings changes are saved and then reloaded. Shutdown stops folder watchers and background work and handles pending cache saves.

Background work must respect window and resource lifetimes. Timer callbacks are wrapped with `Logger.Watch(label, fn)` for timing; network requests use asynchronous `HttpGet.Start` rather than blocking the main thread; index caches are normally saved after keyboard and mouse idle, with pending data handled immediately at shutdown.

### Timers and idle CPU
While ALTRun sits idle it should use no CPU. AutoHotkey keeps an internal 10 ms timer running as long as **any** `SetTimer` timer is set, even one due in an hour, so a single long `SetTimer` costs about 330 million CPU cycles a minute (the constant 0.1% in Task Manager of #290). Rules:

- **`WinTimer.Set(fn, ms)`** (`Lib\WinTimer.ahk`, same meaning as `SetTimer`) for every timer that stays set while idle: periodic checks, folder watch polls, hourly checks, delays of a minute or more. It is a Windows timer, so nothing runs until it is due
- **`SetTimer`** only for short-lived work that is set while there is something to do and turned off afterwards: batches of a background scan, delays while typing, fades, a one-shot `-1` to run something after the current event
- **Window state comes from events, not polling**: `SetWinEventHook` (see the Phases in `Src\Extensions\QuickSwitch.ahk`), `OnMessage`, folder changes through `FolderWatcher`
- Checked by `ProjectTests.LongTimers` (a `SetTimer` with a literal period of 2 s or more, or a one-shot of a minute or more, fails; periods held in variables aren't seen) and by the benchmark: the candidate's median idle CPU over the budget in `Tests\Bench\bench_stats.py` is flagged in the run summary

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
5. Add tests to the matching class in `Tests\Suites\`

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
- README: `README.md` (English, shown by default on GitHub) and `README.zh-CN.md` (Chinese), linking to each other at the top
- `CONTRIBUTING.md` and `SECURITY.md`: one file each, English first and then Chinese, with links to both languages at the top. `CONTRIBUTING.md` only covers reporting problems and the pull request essentials; the details are on this page
- Wiki: two files per page, English `Usage.md` (the default, like `README.md`) and Chinese `Usage-zh.md`; each page links to the other language at the top, and links between Chinese pages use the `-zh` page names. The sidebar `_Sidebar.md` has an English and a Chinese group. Help opened from the program (cheat sheet, `F1` in Preferences) follows the UI language: the `-zh` page for a Chinese UI (Simplified or Traditional), the English page otherwise (`HelpProvider.WikiPage`). The `WikiPages` test checks that every page has a Chinese page and that links point to existing pages
- Until 2026.10.20 the English pages were named `en-Usage` and the Chinese ones `Usage`. For older links and older ALTRun versions, `.github/workflows/wiki.yml` publishes one-line `en-Usage` pages pointing to the new names; it stops by itself once three releases after 2026.10.20 are out
- CHANGELOG: test build sections are written in Chinese; a release section is English first and then Chinese, and is also the release notes (see the format under "Release process" below)
- Links between wiki pages use the page name without `.md` (for example `[Themes](Themes)`; Chinese pages use `[主题](Themes-zh)`). If you're not comfortable with Chinese, change just the English page and say so in the pull request; a maintainer will add the other language

The website (https://zhugecaomao.github.io/ALTRun/) sources are in `site/`: `index.html` is a Chinese / English template, and `build.py` fills in the latest version, download link and size, download count, theme screenshots and the latest release notes, writing to `_site/`. `.github/workflows/pages.yml` builds it when related files change on main, after a successful release and once a day, and deploys it with the official GitHub Pages actions (Settings → Pages → Source: "GitHub Actions"). Local preview: `python3 site/build.py && python3 -m http.server -d _site`.

Screenshots of the UI are in `docs/images/screenshots/` and are generated by `Tests\Screenshots\TakeScreenshots.ahk`: it prepares a demo ALTRun in a temporary folder (English UI; the sample settings, custom commands and files are made-up generic content with nothing personal or work-related), then starts it for each scene, types and takes a screenshot. After UI changes, run **Screenshots** in Actions (`.github/workflows/screenshots.yml`); it retakes the screenshots on Windows and commits them back to the current branch: three runners in parallel (the theme scenes at 150%, the other search window scenes at 150%, Preferences and Quick Switch at 100%), and one job that commits what they changed. To check only a few, fill in "Scenes" when running it (scene names separated by spaces, for example `prefs-general shellmenu`; the 150% and 100% scenes are sorted out by themselves): a couple of minutes instead of about ten. Pushes that change `Tests/Screenshots/` run it automatically too (all scenes), but not when the change only came in by merging main. The animations (APNG) are scenes too: `demo` (the README front page) and `demo-quickswitch`, `demo-clipboard`, `demo-snippets`, `demo-browse`, `demo-calculator` (README, site and Wiki) save one frame per step, and `MakeDemos.ps1` / `MakeDemo.py` put them together on a blurred background. To run it locally:
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
Versions are dates. `App.Version` in `Src\Core\App.ahk` and `;@Ahk2Exe-SetVersion` in `ALTRun.ahk` must match (the `ReleaseVersion` test checks this); every version heading in `CHANGELOG.md` has a link at the end of the file, and `[未发布]` (Unreleased) compares from the newest version (the `ChangelogLinks` test checks this).

**Test builds** (every PR that changes the program):
1. A four-part version `YYYY.MM.DD.N` (for example `2026.10.20.1`); it must be lower than the next release, or installed test builds won't see that release as an update
2. In `CHANGELOG.md`, add a section `## [YYYY.MM.DD.N]` below `## [未发布]`: first paragraph "测试版。", second "Test build.", then 新增 / 改进 / 修复 (New / Improvements / Fixes, in Chinese); add the compare link `[YYYY.MM.DD.N]: .../compare/<previous>...<this>` at the end and make `[未发布]` compare from this version
3. Run the **Release** workflow on the PR branch with `publish = false`; the build is under Artifacts at the bottom of the Actions page (`ALTRun_v<version>`)

**Releases** (at most one per day):
1. Open a pull request that sets the three-part version `YYYY.MM.DD`
2. In `CHANGELOG.md`, merge the round's test build sections into one `## [YYYY.MM.DD]` section, English first and then Chinese, with the English a full translation of the Chinese:
   - `### Summary` (one English paragraph) / `### Upgrade` / `### New` / `### Improvements` / `### Fixes`
   - `---`
   - `### 中文说明` (one Chinese paragraph) / `#### 升级` / `#### 新增` / `#### 改进` / `#### 修复`
   - Don't use `##` headings inside the section (the Release workflow takes everything up to the next `##`); remove the test build sections and their compare links; at the end of the file, `[YYYY.MM.DD]: .../releases/tag/YYYY.MM.DD` and `[未发布]: .../compare/YYYY.MM.DD...HEAD`
3. After merging to main, run **Release** on main with `publish = true`

On Windows the Release workflow: reads the version → downloads AutoHotkey and Ahk2Exe → runs the unit tests → compiles `ALTRun.exe` → signs it when SignPath is configured → checks the program icon → runs the upgrade test with a 2.x `ALTRun.ini` → runs the single instance test on the compiled program → packages `ALTRun_v<version>.zip`. With `publish = true` it then: takes that CHANGELOG section as the release notes → creates the GitHub release (GitHub provides the SHA256 that one-key updates verify) → updates the Scoop and winget manifests in `packaging/` and commits them to main → submits the new version to winget-pkgs when `WINGET_TOKEN` is configured.

## Benchmark
GitHub → Actions → **Bench** → Run workflow: on a GitHub Windows VM, simulates daily use with real keystrokes and compares
two versions (baseline: the latest release by default; candidate: the selected branch by default). Each version runs two rounds,
alternating (A B B A), with ALTRun restarted for every round:

- Random searches (start of app names, word initials, random letters, file search, custom commands): time to show the window, and from the last key to drawn results
- Running programs / opening files and folders, the action panel (→), the calculator (result + Enter to copy), Quick Switch (folder panel + Ctrl+G),
  opening and saving Preferences, snippet expansion
- ALTRun's CPU time, memory, GDI / USER objects and handles at the end of each round
- Background tasks over 50 ms (`Perf: slow`, in versions that log them) and errors from the ALTRun log

By default each round has 500 searches + 100 of each other operation, split over 4 runners working at the same time (about 17 minutes; `runners` sets how many).
Each runner runs both versions alternating (A B B A, or B A A B on even runners so each version gets the first round as often) with its share of the operations, so the two are always compared on the same machine; before the operations,
each round measures the CPU used while idle for a minute. The results (median / p90 / p99 / max) of all runners are pooled in the run summary, followed by the
candidate / baseline median on each runner: they should agree, and a difference seen on only one runner is not to be trusted. The raw data is an artifact.
The test copies #Include one extra file, `Tests\Bench\BenchProbe.ahk`, which tells the driver (`BenchDriver.ahk`) whether the results for the current input have been drawn;
the released program doesn't contain it. Pushing changes to the benchmark scripts to a branch other than main runs a short smoke test automatically.

### File index benchmark
GitHub → Actions → **File index benchmark** runs `Tests\RunTests.ahk FileIndexBench` (`Tests\Bench\FileIndexBench.ahk`) on Windows: it creates 64,000 empty files in 400 folders (one in ten with a Chinese name) and times a full rescan as ALTRun does it, writing and reading the cache, and searches with few and many matches; the numbers are in the run summary. It also runs on pull requests that change the file index, file search or `FuzzyMatcher`, so compare those numbers with the last run on `main`.

AutoHotkey copies a string held in an object property every time an expression reads it, and `obj.prop .= x` builds a whole new string. With long strings (the file index keeps all names in one) that turns linear work into quadratic: build long strings in a local variable or an array, and keep text that is read in small pieces in a `Buffer` (`StrGet` at an offset).

## Quality checks and limitations

- `Tests\RunTests.ahk` covers pure logic such as matching, query parsing, settings migration, ranking, clipboard and indexes; it is not a full manual UI test.
- A test is a static method of a class in `Tests\Suites\` whose name doesn't start with `_` (those are helpers); it runs without being registered anywhere. A test must put back the settings, UI language and logging switches it changes, or it fails, so tests don't depend on their order. Every setting the code reads must have a default in `AppSettings.Defaults()` (checked by `SettingsKeys`).
- Windows CI runs unit tests and syntax checks, and checks Preferences layout, Wiki page pairs and documentation links.
- Screenshot tests check major windows across languages and themes; they do not verify every interaction.
- Performance benchmarks simulate common operations on Windows and compare before / after timings; compare runs in the same environment.
- Wiki automation detects missing pages and broken links, but cannot tell whether explanations are accurate or cover every setting. When settings change, check the code and both language versions item by item.

## Code style
- **Files and classes**: one class per file, file name = class name, PascalCase (`SearchWindow.ahk` → `class SearchWindow`)
- **Folders**: `Lib\` holds general-purpose libraries with nothing ALTRun-specific that could be reused in other projects; ALTRun's own code goes in the matching subfolder of `Src\`
- **Naming**: PascalCase for classes, methods and properties; camelCase for local variables; internal members start with `_` (`_Layout()`, `_cache`)
- **Avoid name clashes**: AutoHotkey names are case-insensitive, so don't give a local variable the name of a class (for example, don't call a variable `pinyin`, it hides the `Pinyin` class); methods and properties in the same class shouldn't differ only in case either
- **Comments**: every file starts with a description of what it's for and how to use it; comments in the code explain "why" and are written in English. `git blame` skips commits listed in `.git-blame-ignore-revs` (such as the one-time translation of the comments) once you run `git config blame.ignoreRevsFile .git-blame-ignore-revs`; GitHub uses it automatically
- **Settings**: add new settings to `AppSettings.Defaults()`; for changes to the structure of existing settings see "Changing the settings structure" above
- **UI text**: the English original goes in `I18n.ahk`, translations go in every language file in `Resources\Lang\`, see "Translating the interface" above
- **Encoding**: UTF-8; `ALTRun.ahk` stays UTF-8 with BOM + CRLF, every other file UTF-8 + LF

A few AutoHotkey v2 pitfalls:
- A very large `static X := Map(...)` fails with "Declaration too long"; build it inside a method instead
- Strings passed to functions by value are copied whole: passing a big string (a whole file) to a function over and over in a loop makes the time grow with the square of its length; pass it by reference (`&text`), see `Lib\JSON.ahk`
- Lines starting with `Perf:` in the debug log (`General.SaveLog`) are the timings of startup phases, slow searches and slow background tasks; use `Logger.Ms()` / `Logger.Time(label, start)` to record new ones, and wrap timer callbacks in `Logger.Watch(label, fn)` (logged only over 50 ms)
- Don't download with `Download()` on the main thread (the whole program stops until it finishes); use `HttpGet.Start(url, done, fail)` to download in the background
- To learn about changes in a folder, use `FolderWatcher` (`Lib\FolderWatcher.ahk`, call `Poll()` on a `WinTimer`, it doesn't wait) instead of rescanning the folder on a timer; when closing it waits for the pending read to finish before releasing the buffer
- When a nested function changes an outer variable, pass the result back through an object (`state := {Result: ""}`)
- Give every `Edit` control a row count (`r1 -Multi`), otherwise long text turns it into a multi-line box (a test checks this)
- Treat every `#Warn All` warning as an error

## Submitting a pull request
1. Create a branch from `main`; one pull request does one thing
2. Before submitting, make sure all unit tests pass and `/validate` shows no warnings (see "Environment" above)
3. Include tests with new features or fixes (`Tests\Suites\`)
4. For UI changes, attach screenshots; the screenshots in the README / wiki are generated by the Screenshots workflow (see "Documentation" above) with the English UI and made-up sample data; don't include anything personal or work-related
5. For user-visible changes, update the README (English and Chinese), `CHANGELOG.md` and, where needed, the wiki (the English page and the `-zh` Chinese page)

By submitting code, you agree to release your contribution under the [GPL-3.0](https://github.com/zhugecaomao/ALTRun/blob/main/LICENSE) license.
