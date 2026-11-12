[English](Development) · **中文**

# 开发指南

## 环境
- [AutoHotkey v2](https://www.autohotkey.com/) 2.0 或更新。所有 workflow (单元测试、集成检查、基准测试、发布构建) 使用同一个固定版本, 目前是 v2.0.31 和 Ahk2Exe 1.1.37.02a2, 版本和 SHA256 写在 `.github/actions/setup-ahk/action.yml` (版本和哈希一起改); Tests workflow 每周一还用最新的稳定版跑一次, 所以 AutoHotkey 新版本先经过测试再固定
- 任意编辑器; 推荐 VS Code + AutoHotkey v2 Language Support 插件

```
AutoHotkey64.exe ALTRun.ahk                                  运行
AutoHotkey64.exe /ErrorStdOut /validate ALTRun.ahk           只检查语法和警告, 不运行
AutoHotkey64.exe /ErrorStdOut Tests\RunTests.ahk             单元测试, 退出码 = 失败的数量
```

## 项目结构
```
ALTRun.ahk          入口: #Include 所有模块, 调用 App.Start()
Lib\                通用库, 与 ALTRun 无关, 可以直接拿到别的项目用
                    (JSON, SafeFile, Logger, Util: Path/Win/Url..., TextTools, DelimitedText, Kanji, Units, Dialogs,
                    Everything IPC, HttpGet, FolderWatcher, WinTimer, WindowsRecent, ShellContextMenu, ExplorerTabs, ClipboardData,
                    TaskScheduler, TaskbarProgress, Sqlite: 通过 Windows 自带的 winsqlite3.dll 只读查询)
Src\Core\           App (启动、单实例、退出), ErrorHandler (未捕获的错误), I18n (界面文字),
                    AppSettings + SchemaMigration (设置和版本升级),
                    SearchQuery / ResultItem (搜索模型), FuzzyMatcher (匹配打分),
                    Knowledge (学习排序), Usage (使用统计), ActionCatalog (执行结果: 打开、复制、粘贴...),
                    ProviderRegistry, FileIndex, BackgroundWork (后台任务什么时候可以运行)
Src\UI\             SearchWindow (+ SearchSession、SearchKeys 键盘、ResultRows 结果行绘制、SearchStatusBar 状态栏、TypeAhead、
                    SearchInputLanguage、WindowFade), ResultActions (操作面板、编辑、删除), PreferencesWindow (+ PrefsPages 各个页面、
                    PrefsForm 表单工具和布局尺寸), ItemEditor, HotkeyEditor (自定义热键窗口), CommandPicker (选择热键的命令), HotkeyBox, LargeType, Hud (简短提示),
                    ThemeManager + ThemePreview, IconCache, NavIcons, SyncEditor + SyncPreview (同步任务编辑和同步窗口)
Src\Providers\      搜索功能, 每个功能一个类
Src\Extensions\     搜索窗口以外的功能: SnippetExpander, QuickSwitch (文件对话框; + QuickSwitchPanel、QuickSwitchFavorites、
                    FileDialogs、FileManagers), AutoDate, CapsLockSwitch +
                    CapsLockHelper, SelectionActions, TendonProfile + PTToolsWindow, CurrencyRates, UpdateChecker,
                    DelayedStart, RoboSync + SyncJobs + SyncProgress (用 robocopy 同步文件夹)
Resources\          随程序发布的数据 (Kanji.txt, Pinyin.txt, Themes\*.json, Lang\*.json, Icons\*.ico)
Tests\RunTests.ahk  单元测试的入口: 载入程序代码 (设置只在内存里), 运行所有测试类
Tests\Suites\       单元测试, 每个方面一个类 (Core, Provider, SearchWindow, Preferences, Extension, Project)
Tests\Fixtures\     测试数据: 旧版本的 ALTRun.ini, SPF2M 对照数据
Tests\Screenshots\  自动截图
Tests\Bench\       性能测试: 在 Windows 上模拟日常操作, 比较两个版本 (见下面 "性能测试")
Tests\Tools\SPF2M\  生成 SPF2M 对照数据的 Python 脚本 (在 DOSBox 里运行原来的 SPF2M.EXE, 见其中的 README)
bucket\ packaging\  Scoop / winget 清单 (见 packaging\README.md)
docs\               Wiki 源文件 (docs\wiki, 合并后自动发布)、截图 (docs\images)
site\               官网模板和生成脚本 (发布到 GitHub Pages)
.github\            GitHub Actions (测试、截图、发布、Wiki、官网、同步到 Gitee)、Issue / PR 模板
```

分层: `Lib` 不用 `Src` 里的任何东西; `Src\Core` 不用 `UI`、`Providers`、`Extensions` (只有负责启动的 `App.ahk` 可以)。Core 需要上层的信息时, 由上层设置回调或调用 Core, 例如 `BackgroundWork.IsBusy` (搜索窗口设置) 和 `IconCache.Prefetch(ProviderRegistry.IconSpecs())` (`App` 调用)。测试 `Layers` 检查这一点。Core 以上的 `UI`、`Providers`、`Extensions` 之间直接互相调用 (Provider 打开 `ItemEditor`, `SyncJobs` 打开同步窗口); 把各部分接起来的只有 `App`, 所以新模块在 `App.Start()` 里启动, 有需要保存或停止的东西就在 `App._OnExit()` 里处理。

运行后程序目录下会生成 `Data\` (设置文件 `ALTRun.json`, 以及可以删除的索引和历史) 和 `Themes\` (自定义主题)。

## 一次搜索的流程
1. `SearchWindow` 输入框变化 → `ProviderRegistry.Search(text)`
2. 文字被解析成 `SearchQuery` (关键字、前缀、剩余文字)
3. 每个启用的 Provider 的 `Search(query)` 返回 `ResultItem` 数组
4. `ProviderRegistry` 合并结果, 加上 `Knowledge` 的学习加分, 按分数排序, 同一个 `Uid` 只保留一条; 有 `Exclusive` 结果时只显示这些; 一条都没有时显示 Provider 的兜底项 (`Fallbacks(query)`, 网页搜索)
5. `SearchWindow` 用 NM_CUSTOMDRAW 自己画每一行
6. 执行时 `ActionCatalog` 按 `Kind` 决定怎么打开 (操作面板、F3、Ctrl+Del 经过 `ResultActions`), `Knowledge.Record()` 记住这次选择

## 核心模块与运行生命周期

- `App` 负责启动、重新载入和退出; `AppSettings` 读取 / 保存设置, `SchemaMigration` 负责旧设置格式升级。
- `SearchWindow` 接收输入; `SearchQuery` 解析关键字和前缀; `ProviderRegistry` 调用已启用的 Provider 并合并结果。
- `ResultItem` 统一结果数据; `FuzzyMatcher` 计算匹配分数; `Knowledge` 调整学习排序; `ActionCatalog` 执行打开、复制、运行等操作。
- `PreferencesWindow` 在设置副本上编辑, 保存后由程序重新载入, 避免控件直接修改运行中的设置对象。
- `FileSearchProvider` (`recent 报告`) 和 `QuickSwitch` (最近的文件夹) 通过 `Lib\WindowsRecent.ahk` 读 Windows 的 "最近使用的项目"。
- `ApplicationProvider` 与 `FileIndex` 共用 `Lib\FolderWatcher.ahk` 监视索引目录; 索引重建在后台进行, 新索引完成前继续使用旧索引。

一次典型的生命周期是: 启动并加载设置 → 初始化 Provider 和窗口 → 接收搜索输入 → 解析查询并收集结果 → 排序和显示 → 执行选中项的操作。设置变更会保存后重新载入; 退出时停止文件夹监视和后台任务, 并完成待保存的缓存处理。

后台工作要注意窗口和资源生命周期: 计时器回调通过 `Logger.Watch(label, fn)` 记录耗时; 网络请求使用 `HttpGet.Start` 异步启动, 不要在主线程里同步等待下载; 索引缓存通常在键鼠空闲后保存, 退出时会立即处理待保存数据。

### 设置和数据
- **位置**: 便携运行时 `AppSettings.DataDir` 是程序旁边的 `Data\`, 否则在用户目录下; 设置文件里的 `DataLocation` 可以指到别处 (`AppSettings.ResolveDataDir`)。ALTRun 写的东西都在这里: `ALTRun.json`、索引、历史、缓存
- **读取**: `AppSettings.Load()` 读 `ALTRun.json`, 旧格式经 `SchemaMigration` 升级 (先备份 `ALTRun.vN.backup.json`), 缺少的项从 `AppSettings.Defaults()` 补上。解析不了的文件留成 `ALTRun.json.bad`, ALTRun 提议改用 `ALTRun.json.bak`; 打不开的文件 (被占用、离线) 既不改名也不覆盖: `SafeFile.ReadText` 先重试一会儿, 然后由用户选重试还是退出 (`_ReadFile`)。"没读出来" 不等于 "是空的": 凡是会写回去的数据文件都用 `SafeFile.ReadText` 读, 读失败时保留原文件 (`ClipboardProvider._Load`)。更新版本的 ALTRun 保存的文件保留它的 `SchemaVersion` 和不认识的项 (`AppSettings.NewerVersion`)
- **保存**: `AppSettings.Save()` 先把上一次正常的文件复制成 `ALTRun.json.bak`, 再经 `SafeFile` 写入 (每次写入自己的临时文件, 刷到磁盘, 再用带 write-through 的 `MoveFileEx` 替换), 崩溃或断电不会留下写了一半的文件。所有数据文件都经 `SafeFile` / `JSON.WriteFile` 写, 不要用 `FileAppend` 直接写正式文件
- **偏好设置** 在副本上编辑; 确定 / 应用时保存并重新载入 ALTRun, 没有窗口直接改运行中的设置
- **退出** (`App._OnExit`): 在这里写入后台重建的索引、学习排序、使用统计、剪贴板历史、图标缓存和日志。新加的、在内存里存数据的功能, 把保存加到这里

### 数据和安全的不变量
每次修改都要保持的规则; 有自动检查的写明是哪一个。
- **写数据**: 会整个读回的数据文件 (`ALTRun.json`、`Data\` 里其它 `.json`、`FileIndex.bin`) 都用 `SafeFile` / `JSON.WriteFile` 写 (`ProjectTests.DataWrites`)
- **读数据**: 读不出来的文件不当成空的, 也不覆盖 (`SafeFile.ReadText`、`AppSettings._ReadFile`、`ClipboardProvider._Load`); 读到的内容先检查再用, 只保留代码能用的条目 (`Knowledge.ValidPicks`、`Usage.ValidDays`、`ApplicationProvider._LoadCache`), 手动改坏或只写了一半的文件不会让每次搜索都报错
- **日志**: 不写搜索框里输入的文字, 只写长度 (`ProjectTests.NoInputTextInLog`)
- **剪贴板**: 借用剪贴板的代码 (`saved := ClipboardAll()`) 在 `finally` 里还原 (`ProjectTests.ClipboardRestore`)
- **其他程序发来的消息**: WM_COPYDATA 只在发送方给出的长度内读取 (`Win.CopyDataText`、`Everything.ParseList`); 另一个实例转交过来、会改动程序的命令 (`-Update`) 先询问
- **命令行**: 参数用 `Win.QuoteArg` 加引号; 用户输入的内容带双引号时不经过 `cmd.exe` (`ScriptProvider.ThroughCmd`)
- **以管理员身份运行的东西**: 用到的文件不放在普通权限程序能在检查和使用之间改动的地方: 更新包暂存在程序旁边 (`UpdateChecker.TempDir`), 任务 XML 在 schtasks 读取期间被占住 (`CapsLockHelper.RegisterXml`); DLL 用完整路径加载
- **网络**: 只有 [SECURITY.md](https://github.com/zhugecaomao/ALTRun/blob/main/SECURITY.md#联网) 里列出的请求; 新增的先写进那里

### 线程、消息和重入
AutoHotkey 同一时间只运行一个线程, 但热键、计时器、`OnMessage` 处理函数和界面事件会开一个新线程, **在当前线程等待时插进来**: `Sleep`、`MsgBox` 等对话框、`WinWait`、`Download()`, 以及会处理消息的 `DllCall` / `SendMessage`。等待过的代码继续时, 要假设共享的状态已经变了:
- 同一个函数可能同时运行两次。`SafeFile` 每次写入用自己的临时文件, 并以最新的内容为准; `Everything` 每个查询 id 一个回复位置 (等回复时可能又开始一个查询); `ExplorerTabs.Open` 只拿没有被其他正在运行的 Open 拿走的新标签
- `Critical("On")` 让一小段代码不被打断 (画结果行、图标缓存); 期间消息会排队, 所以要短, 结束后重绘 (`SearchWindow`)
- **共用的消息**: 同一个消息有几个处理函数, 只在是自己的消息时才返回值: `WM_COPYDATA` (0x4A) 有 `App` (后启动的 ALTRun 传来的命令行)、`Everything` (查询回复)、`QuickSwitch` (Total Commander 的回复) 和 `CapsLockHelper`, 各自检查 `dwData`; 按键消息 (`WM_KEYDOWN` / `WM_SYSKEYDOWN`) 每个窗口的处理函数都会收到, 各自检查窗口
- **菜单**: `Menu.Show` / `TrackPopupMenuEx` 会把 ALTRun 隐藏的主窗口设成前台; 没选菜单项就关掉时, 要把焦点还回去 (`QuickSwitchFavorites._DialogAfterMenu`)
- **网络路径**: 服务器离线时, `FileExist`、`DirExist` 和解析路径 (`SHParseDisplayName`) 会在调用的线程里等网络超时 (CI 上 21 秒), ALTRun 卡住。不要在界面线程里不限时地访问网络路径 (`\\server\share`、映射的网络驱动器): 直接当作存在 (`Path.ExistsOrRemote`), 或者在单独的线程里限时检查 (`ExplorerTabs.FolderWithin`)
- **按键状态**: 远程桌面软件注入所有按键, `GetKeyState(key, "P")` 永远看不到它们按下; 用 `App.IsKeyDown(key)` (逻辑或物理按下)

### 计时器和空闲 CPU
ALTRun 空闲时应该不占 CPU。只要设置了**任何一个** `SetTimer` 计时器 (哪怕一小时后才触发), AutoHotkey 内部就一直开着一个 10 毫秒的计时器, 一个长周期的 `SetTimer` 每分钟大约要 3.3 亿个 CPU 周期 (#290 里任务管理器一直显示的 0.1%)。规则:

- **`WinTimer.Set(fn, ms)`** (`Lib\WinTimer.ahk`, 用法和 `SetTimer` 一样) 用于空闲时也一直设着的计时器: 定期检查、文件夹监视的轮询、每小时的检查、一分钟以上的延迟。它是 Windows 计时器, 到时间之前什么都不运行
- **`SetTimer`** 只用于有事要做时才设、做完就关掉的短时间工作: 后台扫描的分批、打字时的延迟、淡入淡出、`-1` 表示当前事件之后再运行
- **窗口状态靠事件, 不靠轮询**: `SetWinEventHook` (见 `Src\Extensions\QuickSwitch.ahk` 的各个阶段)、`OnMessage`、经 `FolderWatcher` 的文件夹变化
- 由 `ProjectTests.LongTimers` 检查 (周期写成字面量 2 秒以上的 `SetTimer`, 或一分钟以上的一次性计时器会失败; 放在变量里的周期看不到), 性能测试也会检查: 候选版本空闲 CPU 的中位数超过 `Tests\Bench\bench_stats.py` 里的上限时, 会在运行摘要里标出

## 新增一个搜索功能
1. 在 `Src\Providers\` 新建 `XxxProvider.ahk`:
   ```autohotkey
   class XxxProvider {
       static Id := "Xxx"                  ; 同时是 ALTRun.json 里 Features 下的键名

       static Init() {                     ; 启动时调用一次 (建立索引等)
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
2. 在 `ALTRun.ahk` 里 `#Include`, 在 `App.Start()` 里 `ProviderRegistry.Register(XxxProvider)`
3. 在 `AppSettings.Defaults()` 的 `Features` 下加 `"Xxx", Map("Enabled", 1, ...)`
4. 界面文字: 英文加到 `I18n.ahk`, 翻译加到 `Resources\Lang\` 的每个语言文件 (测试会检查每个文件都有)
5. 在 `Tests\Suites\` 里对应的类加测试

可选:
- `EditItem(item)` / `DeleteItem(item)`: 支持在结果里 `F3` 编辑、`Ctrl+Del` 删除; 结果的 `Source` 要指向设置里对应的那一条
- `DeletePrompt(item)`: 删除前的确认文字

### ResultItem 字段
| 字段 | 说明 |
|---|---|
| Title / Subtitle | 两行文字 |
| Icon | 文件路径 (取它的图标)、`folder:`、`url:`、`res:imageres.dll,-5314` |
| Kind | `file` / `folder` / `url` / `text` / 空; 决定默认操作和操作面板 |
| Arg / Arguments | 目标和命令行参数 |
| OnRun | 自定义的执行函数, 优先于 Kind |
| Actions | 额外的操作面板项 |
| LeadActions | 排在内置操作前面的操作面板项 (选中内容的自定义命令, 见 `SelectionActions`) |
| Uid | 学习排序和去重用的唯一标识 |
| Score | 分数, 越大越靠前 (匹配分 0~100, 再加各功能自己的偏移) |
| Valid / AutoComplete | `Valid = false` 时 Enter 只把 AutoComplete 填进输入框 |
| Exclusive | 关键字模式: 只显示有这个标记的结果 |
| Source / Provider | 设置里对应的那一条 / 产生结果的 Provider Id (由 ProviderRegistry 填) |

## 修改设置结构
- 只是加新的设置项: 加到 `AppSettings.Defaults()` 即可, 旧文件缺少的项会自动补上
- 改名、移动或改变已有项的结构: `AppSettings.CurrentVersion + 1`, 在 `SchemaMigration` 里加一个 `_FromN(data)` 把版本 N 转换成 N+1。升级前会自动备份为 `ALTRun.vN.backup.json`
- 把设置项挪到另一组: 加到 `AppSettings.MovedKeys`。`AppSettings.MoveGroups` 每次加载都跑, 不只在升级时: 旧版 ALTRun 读这个文件时会在旧位置补上默认值, 新版下次加载时删掉 (以新位置为准)
- 自定义命令、文字片段、同步任务有 `Id` (`AppSettings.EnsureIds`, 每次加载和保存前都跑): 热键按它引用命令 (`custom:<Id>`), 学到的排序也按它记 (`Knowledge.json` 有自己的 `Version`; `Knowledge.ConvertUids`)。有效的 Id 不改; 复制出来的条目去掉 Id (保存时生成新的)

## 修改清单
一处改动往往牵涉不止那一行。动手前先搜出用到要改的东西的所有地方 (整个仓库, 包括 `Tests\`、`docs\`、`Resources\Lang\` 和 `.github\`), 再对照下面相应的一条:

| 改动 | 还要改 |
|---|---|
| **新的设置项** | `AppSettings.Defaults()` 里的默认值 (没有的话 `SettingsKeys` 会失败); 数值设置还要在 `AppSettings.NumberRanges` 里写上可用范围 (超出的值在加载和偏好设置保存时改回范围内, 并写日志); 偏好设置页面上的控件 (`PrefsPages`); 四种语言的界面文字; 介绍这个功能的 Wiki 页 (英文和 `-zh`); 测试 |
| **已有设置的结构** (改名、移动、改格式) | 见 "修改设置结构": `CurrentVersion + 1` 和 `SchemaMigration` 里的 `_FromN`, 加上从旧文件升级的测试; 攒到下次升级设置版本时一起做 |
| **界面文字** (新加、修改、删除) | `I18n.ahk` 和 `Resources\Lang\` 的三个文件 (`I18nLanguages`、`I18nUnused`); 偏好设置里检查日文是否还放得下 (`PreferencesFit`、`PreferencesSpacing`) |
| **偏好设置页面** | 排版规则: 短框两列对齐, 长内容伸到右边; 这一页的截图场景 (加新页面时, 后面场景的 `-Preferences N` 编号要跟着改); 要显示别的页面改动的页面, 在构建函数里设 `page.OnShow` (每次显示这一页时调用, 见 选中内容 页); 界面改动先用截图确认再合并 |
| **列表的默认条目** (自定义热键、命令、文字片段、网页搜索引擎) | 只写在 `AppSettings.Defaults()` 里: 列表的 "添加示例" 按钮从那里读。新的列表页要有默认条目时, 给 `PrefsForm.List` 传 `examples` 和 `Keys` 函数 (怎样算同一个条目, 见 `PrefsPages.HotkeyKeys`), 在 `ListExamples` 里测试。列表的额外命令多于一个时放进 "更多 ▾" 菜单 (`menu`), 按钮行给筛选框留出位置 (`PreferencesFit`) |
| **新的搜索功能** | 见 "新增一个搜索功能" |
| **搜索窗口以外的功能** (扩展、后台工作) | 在 `App.Start()` 里启动, 在 `App._OnExit()` 里停止 / 保存数据; 一直设着的计时器用 `WinTimer` (见 "计时器和空闲 CPU"); 会处理消息的等待按 "线程、消息和重入" 处理 |
| **联网** | 只用 `HttpGet.Start` (异步); 写进 Wiki 的隐私页 (英文和 `-zh`)、`SECURITY.md` 和 README |
| **依赖 Windows 或别的程序的行为** (API 的返回、命令的输出、注册表、任务计划程序、窗口类名) | 先在 Windows CI 上探测, 按实际返回写代码; 能自动化的写成集成检查 (`Tests\Integration\`) |
| **新的测试或检查脚本** | CI 要运行的, 加到 `tests.yml` 的语法检查列表 (`foreach ($script in ...)`), 经 `Tests\Tools\RunAhk.ps1` 有时间上限地运行, 并把相关文件列进 workflow 的 `paths` |
| **改名或删除函数、设置、文字** | 搜出所有用到的地方, 包括测试、Wiki、语言文件和 workflow; 跑 `NameCheck.ahk` (改了名的局部变量可能遮住某个类) |
| **用户能看到的变化** | README (中英文)、Wiki 页 (英文和 `-zh`)、`CHANGELOG.md` (测试版一节)、版本号 (`App.Version` 和 `;@Ahk2Exe-SetVersion`) |

## 文档
Wiki 的源文件在仓库的 `docs/wiki/`, 合并到 `main` 后自动发布到 Wiki (`.github/workflows/wiki.yml`)。请通过 PR 修改 `docs/wiki/`, 不要直接在网页上编辑。

文档是中英双语的, 改一种语言时另一种也要一起改:
- README: `README.md` (英文, GitHub 默认显示) 和 `README.zh-CN.md` (中文), 开头互相链接
- `CONTRIBUTING.md` 和 `SECURITY.md`: 一个文件里英文在前、中文在后, 开头有跳到两种语言的链接。`CONTRIBUTING.md` 只写反馈问题和提交 PR 的要点, 详细内容在这一页
- Wiki: 每页两个文件, 英文 `Usage.md` (默认, 和 `README.md` 一样) 和中文 `Usage-zh.md`; 每页开头链接到另一种语言, 中文页之间的链接写 `-zh` 的页面名。侧边栏 `_Sidebar.md` 分中英两组。程序里的帮助 (速查表、偏好设置的 `F1`) 按界面语言打开: 中文界面 (简体、繁体) 打开 `-zh` 页, 其它语言打开英文页 (`HelpProvider.WikiPage`)。`WikiPages` 测试检查每页都有中文页、链接指向存在的页面
- 2026.10.20 以前英文页叫 `en-Usage`、中文页叫 `Usage`。为了旧链接和旧版本的 ALTRun, `.github/workflows/wiki.yml` 曾发布一行的 `en-Usage` 页面指向新的名字, 2026.10.20 之后发布满三个正式版后已停止, 这些页面已经删除
- CHANGELOG: 测试版各节用中文写; 正式版那一节英文在前、中文在后, 也是 Release 的说明 (格式见下面的 "发布流程")
- Wiki 页面之间的链接写页面名, 不带 `.md` (例如 `[主题](Themes-zh)`, 英文页写 `[Themes](Themes)`)。英文不熟悉的话可以只改中文, 在 PR 里说明, 维护者会补上英文

官网 (https://zhugecaomao.github.io/ALTRun/) 的源文件在 `site/`: `index.html` 是中英文模板, `build.py` 填入最新版本号、下载链接和大小、下载次数、主题截图和最新版本的更新内容, 生成到 `_site/`。`.github/workflows/pages.yml` 在 main 上相关文件变化、Release 成功后和每天一次自动生成, 用 GitHub Pages 的官方 Actions 部署 (Settings → Pages → Source 选 "GitHub Actions")。本地预览: `python3 site/build.py && python3 -m http.server -d _site`。

界面截图在 `docs/images/screenshots/`, 由 `Tests\Screenshots\TakeScreenshots.ahk` 生成: 它在临时文件夹里准备一份演示用的 ALTRun (英文界面; 示例设置、自定义命令、示例文件都是虚构的通用内容, 不放个人或工作相关的信息), 逐个场景启动、输入、截图。界面改动后, 在 Actions 里运行 **Screenshots** (`.github/workflows/screenshots.yml`), 它在 Windows 上重新截图并提交回当前分支: 三台机器并行 (150% 的主题场景、150% 的其它搜索窗口场景、100% 的偏好设置和对话框跳转), 最后一个任务把它们改动的截图一起提交。只想看几张时, 运行时在 "Scenes" 里填场景名, 用空格分开 (例如 `prefs-general shellmenu`; 150% 和 100% 的场景会自动分开), 几分钟就好, 不用等十分钟左右。修改 `Tests/Screenshots/` 的推送也会自动运行 (全部场景), 但只是合并 main 带进来的改动不会。动图 (APNG) 也是场景: `demo` (README 首页) 和 `demo-quickswitch`、`demo-clipboard`、`demo-snippets`、`demo-browse`、`demo-calculator` (README、网站和 Wiki) 每一步存一帧, 由 `MakeDemos.ps1` / `MakeDemo.py` 合成到模糊的背景上。本地运行:
```
AutoHotkey64.exe Tests\Screenshots\TakeScreenshots.ahk [输出文件夹] [场景名...]
```

## 界面翻译
英文原文写在 `Src\Core\I18n.ahk` 里 (`s["Prefs.Page.General"] := "General"`), 是所有键的标准; 其它语言是 `Resources\Lang\<代码>.json` (UTF-8), 键和英文一样, 值是翻译。文件里没有的键先找 `_fallback` 指定的语言 (例如繁体中文缺的找简体中文), 最后用英文。

| 特殊的键 | 作用 |
|---|---|
| `_name` | 偏好设置 → 通用 → 语言 里显示的名称, 用这种语言自己的文字 (例如 `日本語`) |
| `_font` | 界面字体 (主题的 `FontName` 是 `auto` 时用), 例如 `Microsoft YaHei UI`、`Yu Gothic UI` |
| `_fallback` | 可选, 缺的键先找哪种语言 |
| `_comment` | 写给翻译者的说明, 程序不读 |

翻译时: `{1}` `{2}` 这些占位符的个数要和英文一样 (位置可以换); `\n` 是换行; 偏好设置的说明要短, 日文最容易超出页面宽度。

**新加一条界面文字**: 在 `I18n.ahk` 里加英文, 再在三个 JSON 里加同一个键的翻译。

**添加一种语言** (例如韩语):
1. 复制 `zh-CN.json` 为 `Resources\Lang\ko.json`, 改 `_name` (`한국어`) 和 `_font` (`Malgun Gothic`), 逐条翻译
2. 不用改代码: 偏好设置的语言列表自动出现; "自动" 按 Windows 的显示语言选, 地区代码 (例如 `zh-HK` → `zh-TW`) 在 `I18n.Aliases` 里对应
3. 跑测试: `I18nLanguages` 检查每个语言文件不缺键、没有多余的键、不是空的、占位符个数一致、写了 `_font`; `PreferencesFit` / `PreferencesSpacing` 检查每一页在这种语言下放得下、不挤在一起; `I18nUnused` 检查英文里没有不用的键
4. Wiki 只有中英文: 中文界面 (简体、繁体) 的帮助打开中文页, 其它语言打开英文页

## 发布流程
版本号用日期。`Src\Core\App.ahk` 的 `App.Version` 和 `ALTRun.ahk` 的 `;@Ahk2Exe-SetVersion` 要一致 (`ReleaseVersion` 测试检查); `CHANGELOG.md` 每个版本标题在文件末尾都有链接, `[未发布]` 从最新的版本比较 (`ChangelogLinks` 测试检查)。

**测试版** (每个改了程序的 PR):
1. 版本号四段 `YYYY.MM.DD.N` (例如 `2026.10.20.1`); 要比下一个正式版小, 否则装了测试版的检测不到正式版的更新
2. `CHANGELOG.md` 在 `## [未发布]` 下面加一节 `## [YYYY.MM.DD.N]`: 第一段 "测试版。", 第二段 "Test build.", 再写 新增 / 改进 / 修复 (用中文); 文件末尾加比较链接 `[YYYY.MM.DD.N]: .../compare/<上一版>...<这一版>`, `[未发布]` 改成从这一版比较
3. PR 合并到 main 后自动构建测试版 (`App.Version` 变了时, **Release** workflow 以 `publish = false` 运行); 构建好的程序在 Actions 页面底部的 Artifacts (`ALTRun_v<版本>`)。合并前就要测试版时, 在 PR 的分支上运行 **Release**, `publish = false`

**正式版** (每天最多一个):
1. 开一个 PR, 版本号改成三段 `YYYY.MM.DD`
2. `CHANGELOG.md` 把这一轮测试版的各节合并成一节 `## [YYYY.MM.DD]`, 英文在前、中文在后, 英文是中文的完整翻译:
   - `### Summary` (英文摘要一段) / `### Upgrade` / `### New` / `### Improvements` / `### Fixes`
   - `---`
   - `### 中文说明` (中文摘要一段) / `#### 升级` / `#### 新增` / `#### 改进` / `#### 修复`
   - 这一节里不要用 `##` 标题 (Release workflow 取到下一个 `##` 为止); 删掉测试版各节和它们的比较链接; 文件末尾 `[YYYY.MM.DD]: .../releases/tag/YYYY.MM.DD`, `[未发布]: .../compare/YYYY.MM.DD...HEAD`
3. 合并到 main 后会自动跑一次检查构建 (`publish = false`); 再在 main 上运行 **Release**, `publish = true`

Release workflow 在 Windows 上依次: 读版本号 → 下载 AutoHotkey 和 Ahk2Exe (固定版本, 核对 SHA256) → 单元测试 (只在发布正式版时; 测试版的同一个提交由 Tests workflow 测试) → 编译 `ALTRun.exe` → 配置了 SignPath 时签名 → 检查程序图标 → 用 2.x 的 `ALTRun.ini` 做升级测试 → 编译后的程序做单实例测试 → 打包 `ALTRun_v<版本>.zip`。`publish = true` 时再: 从 CHANGELOG 取这一节作为 Release 说明 → 创建 GitHub Release (一键更新核对的 SHA256 由 GitHub 给出) → 更新 `packaging/` 里的 Scoop 和 winget 清单并提交到 main → 配置了 `WINGET_TOKEN` 时向 winget-pkgs 提交新版本 (`packaging/Submit-Winget.ps1`: 先把 token 所属账号 fork 的 winget-pkgs 同步到最新, 再用 wingetcreate 开 PR; 失败时用 **Submit to winget** 工作流重新提交已发布的版本)。

## 性能测试
GitHub → Actions → **Bench** → Run workflow: 在 GitHub 的 Windows 虚拟机上用真实的按键模拟日常操作, 比较两个版本
(baseline 默认最新正式版, candidate 默认选中的分支)。每个版本两轮, 两个版本轮流 (A B B A), 每轮重新启动 ALTRun:

- 随机搜索 (应用名称开头、单词首字母、随便几个字母、文件搜索、自定义命令): 呼出窗口的时间, 最后一个字到结果画好的时间
- 运行程序 / 打开文件和文件夹、操作面板 (→)、计算器 (结果 + Enter 复制)、Quick Switch (文件夹面板 + Ctrl+G 跳转)、
  打开并保存偏好设置、片段展开
- 每轮结束时 ALTRun 的 CPU 时间、内存、GDI / USER 对象和句柄数
- ALTRun 日志里一次超过 50 ms 的后台任务 (`Perf: slow`, 有这种记录的版本) 和错误

默认每轮 500 次搜索 + 每种操作 100 次, 分给 4 台 runner 同时跑 (约 17 分钟; 台数用 `runners` 设置)。
每台都轮流跑两个版本 (A B B A, 双数的 runner 是 B A A B, 两个版本排在第一轮的次数一样), 做自己那一份操作, 所以两个版本总是在同一台机器上比较; 每轮开始操作前先测一分钟空闲时的 CPU。
所有 runner 的结果 (中位数 / p90 / p99 / 最慢) 合在一起放在运行摘要里, 后面是每台 runner 上 candidate / baseline 的中位数比值:
各台应该一致, 只在一台上出现的差别不可信。原始数据是 artifact。
测试副本里多 #Include 一个 `Tests\Bench\BenchProbe.ahk`, 驱动程序 (`BenchDriver.ahk`) 通过它知道 "当前输入的结果画好了没有";
发布的程序里没有它。改了测试脚本推到 main 以外的分支时会自动用很少的次数跑一遍。

### 文件索引性能测试
GitHub → Actions → **File index benchmark** 在 Windows 上运行 `Tests\RunTests.ahk FileIndexBench` (`Tests\Bench\FileIndexBench.ahk`): 创建 64000 个空文件 (400 个文件夹, 十分之一是中文名), 测量和 ALTRun 一样在后台完整扫描一遍、写入和读取缓存、匹配少和匹配多的搜索各要多久, 数字在运行摘要里。改到文件索引、文件搜索或 `FuzzyMatcher` 的 PR 也会自动运行, 和 `main` 上最近一次的数字比较。

### Everything
GitHub → Actions → **Everything** 在 Windows 上用真实的 Everything 检查 ALTRun 的 Everything IPC (`Lib\Everything.ahk`): 1.4 (稳定版) 和 1.5 (beta), 都从 voidtools.com 下载固定的版本并校验 SHA256 (更新时在 `.github/workflows/everything.yml` 里两个一起改)。`Tests\Integration\EverythingCheck.ahk` 搜索一个名字唯一的测试文件和文件夹, 检查路径和是否文件夹; 1.5 还会以命名实例 "1.5a" 运行一次, 和 1.5 alpha 一样。改到 `Lib\Everything.ahk` 或这个检查的 PR 会自动运行; 脚本也可以在自己的电脑上 Everything 运行时使用 (`EverythingCheck.ahk <文件夹> <名字>`)。单元测试 (`Tests\Suites\EverythingTests.ahk`) 用的是 `Tests\FakeEverything.ahk`: 一个和 Everything 同类名的隐藏窗口, 像 Everything 一样回复, 可以马上回复、晚一点回复、不回复, 或者先发一个过期的回复。

AutoHotkey 每次在表达式里读取对象属性里的字符串都会复制一份, `obj.prop .= x` 也会重新生成整个字符串。字符串很长时 (文件索引把所有文件名放在一个字符串里), 线性的工作会变成平方级: 长字符串在局部变量或数组里拼接, 需要一小段一小段读取的文字放在 `Buffer` 里 (按偏移用 `StrGet` 读)。

### Quick Switch
GitHub → Actions → **Quick Switch** 用 `Tests\Integration\QuickSwitchCheck.ahk` 在真实的 Windows 桌面上检查自动跳转 (`Extensions.QuickSwitch.AutoSwitch`)。它启动一份打开了自动跳转的 ALTRun, 用资源管理器打开一个文件夹、再打开一个 "打开" 对话框, 然后用三种方式从资源管理器切到对话框: 直接激活 (`SetForegroundWindow`)、模拟 Alt+Tab、点对话框的标题栏。每种方式试 5 次, 每次对话框都要在 5 秒内跳到资源管理器的文件夹; 失败的那一轮会打印 ALTRun 日志里 Quick Switch 的几行。然后打开文件夹面板的 ☆ 菜单 (点星号、在搜索框里按 Alt+↓), 按 Esc 关掉: 对话框要在 2 秒内回到前台; 失败时打印 Esc 之后到过前台的每个窗口。修改 Quick Switch (`Src\Extensions\QuickSwitch*.ahk`、`FileDialogs.ahk`、`FileManagers.ahk`) 或这个检查的 PR 会运行它, main 上也会, 每周一也运行一次 (运行环境的新镜像可能改变窗口到前台的方式; **Explorer tabs** 也是每周一运行); 想找偶尔出现的失败, 可以手动运行并加大次数。脚本也能在自己电脑上运行 (会占用键盘和屏幕一分钟左右): `QuickSwitchCheck.ahk <仓库的一份副本> [次数]`。

### 延迟启动
GitHub → Actions → **Delayed start** 用 `Tests\Integration\DelayedStartCheck.ahk` (以 `Tests\RunTests.ahk DelayedStartCheck` 运行) 对真实的 Windows 任务计划程序检查延迟启动: 先用一个新建的普通用户 (这才是关键: `schtasks /SC ONLOGON` 需要管理员, ALTRun 用的 COM 接口不需要), 再用 runner 本身的管理员。它通过 `DelayedStart.Sync` 注册任务, 读回来, 修改再删除, 并把为检查建的一个开机启动程序关掉再打开。修改延迟启动或这个检查的 PR 和 main 上会运行; 在自己的电脑上也能跑, 它只动自己的任务文件夹 (`\ALTRun\<用户名> check`), 最后会删掉。

## 质量验证与局限

- `Tests\RunTests.ahk`: 覆盖匹配、查询解析、设置迁移、排序、剪贴板和索引等纯逻辑; 不等于完整的手动 UI 测试。
- 测试就是 `Tests\Suites\` 里各个类的静态方法, 名字不以 `_` 开头 (`_` 开头的是辅助函数); 不用另外登记就会运行。测试改了设置、界面语言或日志开关要改回去, 否则算失败, 所以测试之间不依赖先后顺序。程序读到的每个设置都要在 `AppSettings.Defaults()` 里有默认值 (`SettingsKeys` 检查)。
- Windows CI: 运行单元测试、语法检查和重名检查 (`Tests\Tools\NameCheck.ahk`, 见下面的 "避免重名"), 并检查偏好设置布局、Wiki 页面配对和文档链接。
- CI 通过 `Tests\Tools\RunAhk.ps1` 运行 AutoHotkey 脚本, 每个脚本都有时间上限 (单元测试 10 分钟, 每个语法检查 2 分钟, 截图 25 分钟)。脚本卡在没人关的对话框上时 (例如类初始化时出错, 弹出了错误窗口) 会被停止, 日志里打出它的窗口标题和文字, 从对话框的内容就能看出哪里出错。每个任务另外还有 `timeout-minutes` 兜底。
- 截图测试: 检查主要窗口在不同语言 / 主题下的显示; 不验证所有交互流程。
- 性能基准: 在 Windows 上模拟常见操作比较改动前后的耗时, 应在相同环境下对比。
- Wiki 自动检查只能发现缺页和断链, 不能判断说明是否准确或是否覆盖每个设置项; 修改设置时仍需逐项对照代码和中英文文档。
- 文件已删除的 workflow (临时探测) 的运行记录会留在 Actions 侧栏。GitHub → Actions → **Clean up runs** (`.github/workflows/cleanup-runs.yml`) 列出文件在所有分支上都已不存在的 workflow 和它们的运行记录; 勾选 "delete" 再运行就删除这些记录。删掉临时 workflow 和它的分支后运行一次。

### 在用户电脑上排查按键问题

有些按键问题只在某一台电脑上、用真键盘才出现, CI 上用程序发键复现不了。`Tests\Tools\KeyDiag.ahk` 记录在那台电脑上哪些按键送到了哪个窗口。用 AutoHotkey v2 运行, 权限和出问题的程序一样; 它待在托盘里, ALTRun 退出、重启时也照样记录。

- 只在 ALTRun 或 `KeyDiagSettings.Watch` 里的程序在前面时记录: 每个按键的按下 / 松开 (字母、数字、符号和输入的文字只记成 `key`)、按键来自键盘还是程序、自动重复、到前面的窗口、打开的菜单, 以及 ALTRun 退出和启动。不记录窗口标题。
- 出问题的窗口 (`KeyDiagSettings.Trigger*`) 出现时, 响一声并保存 `KeyDiag-hit-<时间>.txt`: 之前 15 秒的记录、键盘设置和正在运行的程序 (名字和版本)。文件放在 `%LOCALAPPDATA%\ALTRun\KeyDiag`, 用户只要发这一个文件。
- `Ctrl+F9` 从 `KeyDiagSettings.TestWindow` 开始自动测试: 按 ALTRun 的热键、输入退出的文字、按 Enter (短按和按住带自动重复), 检查有没有出现那个窗口, 再重启 ALTRun。`Ctrl+F10` 保存最近 60 秒的报告。
- 默认设置针对写它时的问题: Total Commander 的 "同步目录" 窗口在 Alt+Space、对 "退出" 按 Enter 之后弹出属性窗口, 因为它把单独收到的 Enter 松开当成了 Alt+Enter: 窗口已经激活、焦点还没恢复时收到的按键松开会变成 `WM_SYSKEYUP` (2026.11.09 修复, 本机键盘和远程桌面软件都是: 它们的按键都是程序注入的, 从来不算 "物理按下", 要看 Enter 的逻辑状态, 不能用 `GetKeyState(..., "P")`)。按键由程序发出时显示为 `sent by a program (0x...)`: `0x43524430` ("CRD0") 是 Chrome 远程桌面。排查别的问题时改开头的设置。

## 代码规范
- **文件与类**: 一个文件一个类, 文件名 = 类名, PascalCase (`SearchWindow.ahk` → `class SearchWindow`)
- **文件夹**: `Lib\` 放和 ALTRun 无关、可以拿到别的项目用的通用库; ALTRun 自己的代码放 `Src\` 下对应的子文件夹
- **命名**: 类、方法、属性用 PascalCase; 局部变量用 camelCase; 内部使用的成员以 `_` 开头 (`_Layout()`、`_cache`)
- **避免重名**: AutoHotkey 的名字不区分大小写, `_tick` 和 `_Tick()` 是同一个名字
  - 存函数的属性 (计时器回调、传进来的回调) 加后缀, 就不会和方法重名: `_tickTimer`、`_pollTimer`、`onDoneFn`。类里写 `static _tick := ...`, 旁边又有方法 `_Tick()`, 会在类初始化时出错, 这时 ALTRun 还接不住错误, ALTRun (和单元测试) 启动时就停在一个对话框上; `this._tick := ...` 在运行到这一行时出错。CI 用 `Tests\Tools\NameCheck.ahk` 检查这两种 (也会看项目里从别的类继承来的方法); 自己运行: `AutoHotkey64.exe /ErrorStdOut Tests\Tools\NameCheck.ahk`
  - 局部变量不要和类或内置函数同名: `pinyin` 会遮住 `Pinyin` 类, `path` 遮住 `Path` 类, `file` 遮住 AutoHotkey 的 `File` 类, `log` 遮住它的 `Log()` 函数; 用 `filePath`、`handle`、`logFile`。这种情况 `#Warn` 会报
- **注释**: 每个文件开头说明用途和用法; 代码里的注释说明 "为什么", 一律用英文写。`.git-blame-ignore-revs` 里列出的提交 (例如注释的一次性翻译) 在 `git blame` 里跳过: 本地运行一次 `git config blame.ignoreRevsFile .git-blame-ignore-revs`, GitHub 会自动使用
- **设置**: 新的设置项加到 `AppSettings.Defaults()`; 改变已有设置的结构见上面的 "修改设置结构"
- **界面文字**: 英文原文写在 `I18n.ahk`, 翻译写在 `Resources\Lang\` 的每个语言文件里, 见上面的 "界面翻译"
- **编码**: UTF-8; `ALTRun.ahk` 保持 UTF-8 BOM + CRLF, 其它文件 UTF-8 + LF

几个 AutoHotkey v2 的坑:
- 很大的 `static X := Map(...)` 会报 "Declaration too long", 改成在方法里构造
- 字符串按值传给函数时会整段复制: 在循环里反复把一个大字符串 (整个文件) 传给函数, 耗时会随长度平方增长, 要按引用传 (`&text`), 见 `Lib\JSON.ahk`
- 调试日志 (`General.SaveLog`) 里 `Perf:` 开头的行是启动各阶段、慢搜索和慢的后台任务的耗时; 用 `Logger.Ms()` / `Logger.Time(label, start)` 记录新的, 计时器的回调用 `Logger.Watch(label, fn)` 包一层 (超过 50 ms 才记)
- 不要在主线程里用 `Download()` 下载 (下载完之前整个程序都停着), 用 `HttpGet.Start(url, done, fail)` 在后台下载
- 要知道文件夹里的变化, 用 `FolderWatcher` (`Lib\FolderWatcher.ahk`, 定时 `Poll()` 不等待), 不要定时扫描整个文件夹; 关闭时它会等异步读取结束再放掉缓冲区
- 嵌套函数修改外层变量时, 通过对象传回结果 (`state := {Result: ""}`)
- 所有 `Edit` 控件都写明行数 (`r1 -Multi`), 否则长文字会让它自动变成多行 (有测试检查)
- `#Warn All` 的警告都当作错误处理

## 提交 Pull Request
1. 从 `main` 新建分支, 一个 PR 只做一件事
2. 提交前确认单元测试全部通过、`/validate` 没有警告 (见上面的 "环境")
3. 新功能或修复请附带测试 (`Tests\Suites\`)
4. 涉及界面的改动请附上截图; README / Wiki 里的截图由 Screenshots workflow 生成 (见上面的 "文档"), 用英文界面和虚构的示例数据, 不要放个人或工作相关的信息
5. 用户可见的变化请更新 README (中英文)、`CHANGELOG.md`, 必要时更新 Wiki (英文页和 `-zh` 中文页)

提交代码即表示你同意以 [GPL-3.0](https://github.com/zhugecaomao/ALTRun/blob/main/LICENSE) 许可发布你的贡献。
