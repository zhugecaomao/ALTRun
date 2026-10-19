[English](en-Development) · **中文**

# 开发指南

## 环境
- [AutoHotkey v2](https://www.autohotkey.com/) 2.0 或更新
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
                    (JSON, Logger, Util: Path/Win/Url..., TextTools, Kanji, Dialogs, Everything IPC)
Src\Core\           App (启动流程), AppSettings + SchemaMigration (设置和版本升级),
                    SearchQuery / ResultItem (搜索模型), FuzzyMatcher (匹配打分),
                    Knowledge (学习排序), Usage (使用统计), ActionCatalog (操作), ProviderRegistry, FileIndex
Src\UI\             SearchWindow, PreferencesWindow, ItemEditor, LargeType, ThemeManager, IconCache
Src\Providers\      搜索功能, 每个功能一个类
Src\Extensions\     搜索窗口以外的功能 (SnippetExpander, QuickSwitch, AutoDate, CapsLockSwitch, TendonProfile + PTToolsWindow, UpdateChecker)
Resources\          随程序发布的数据 (Kanji.txt, Themes\*.json, Lang\*.json, Icons\*.ico)
Tests\RunTests.ahk  单元测试
Tests\Fixtures\     测试数据: 旧版本的 ALTRun.ini, SPF2M 对照数据
Tests\Screenshots\  自动截图
Tests\Bench\       性能测试: 在 Windows 上模拟日常操作, 比较两个版本 (见下面 "性能测试")
Tests\Tools\SPF2M\  生成 SPF2M 对照数据的 Python 脚本 (在 DOSBox 里运行原来的 SPF2M.EXE, 见其中的 README)
bucket\ packaging\  Scoop / winget 清单 (见 packaging\README.md)
docs\               Wiki 源文件 (docs\wiki, 合并后自动发布)、截图 (docs\images)
site\               官网模板和生成脚本 (发布到 GitHub Pages)
.github\            GitHub Actions (测试、截图、发布、Wiki、官网、同步到 Gitee)、Issue / PR 模板
```

运行后程序目录下会生成 `Data\` (设置文件 `ALTRun.json`, 以及可以删除的索引和历史) 和 `Themes\` (自定义主题)。

## 一次搜索的流程
1. `SearchWindow` 输入框变化 → `ProviderRegistry.Search(text)`
2. 文字被解析成 `SearchQuery` (关键字、前缀、剩余文字)
3. 每个启用的 Provider 的 `Search(query)` 返回 `ResultItem` 数组
4. `ProviderRegistry` 合并结果, 加上 `Knowledge` 的学习加分, 按分数排序, 同一个 `Uid` 只保留一条; 有 `Exclusive` 结果时只显示这些; 一条都没有时显示 WebSearch 的兜底项
5. `SearchWindow` 用 NM_CUSTOMDRAW 自己画每一行
6. 执行时 `ActionCatalog` 按 `Kind` 决定怎么打开, `Knowledge.Record()` 记住这次选择

## 模块边界与数据流

- App 负责启动、全局热键、托盘、重载和退出; AppSettings / SchemaMigration 负责设置默认值、保存与版本迁移。
- SearchWindow 接收输入并显示结果; SearchQuery 解析输入语法, ProviderRegistry 调用各个 Provider 并合并结果, FuzzyMatcher 计算匹配分数, Knowledge 提供学习排序加分。
- ResultItem 是搜索功能与界面之间的结果模型; ActionCatalog 根据结果类型生成默认操作和操作面板项目。
- Src/Providers 中的 Provider 负责各自功能的数据读取和搜索; 共用的文件系统监视放在 Lib/FolderWatcher.ahk, 避免应用索引和文件索引各自实现一套监视器。

~~~mermaid
flowchart TD
    Input[SearchWindow 输入] --> Query[SearchQuery]
    Query --> Registry[ProviderRegistry]
    Registry --> Providers[各 Provider]
    Providers --> Match[FuzzyMatcher 与 Knowledge]
    Match --> Results[ResultItem 列表]
    Results --> UI[SearchWindow 显示]
    UI --> Actions[ActionCatalog 执行操作]
~~~

## 后台任务与资源生命周期

- 启动时加载设置、主题、索引和各 Provider; 设置变更通过应用 / 确定保存后重载, 取消则丢弃工作副本。
- 文件夹监视器和定时器只负责发现变化并安排后台工作; 定时回调应通过 Logger.Watch(label, fn) 记录耗时。主线程里的网络请求使用 HttpGet.Start, 避免同步下载阻塞窗口。
- 索引与历史数据保存在数据目录, 可重建的缓存应允许丢失后恢复; 退出时应停止监视和定时任务, 并完成必要的缓存保存, 避免回调继续访问已关闭的窗口或已释放的状态。
- 排查问题时先看搜索结果是否只在特定 Provider 缺失, 再检查对应设置、索引目录与排除规则; 需要更多信息时开启高级页的调试日志。日志用于诊断, 不应记录用户输入的搜索文字。

## 测试能验证什么

- Tests/RunTests.ahk 覆盖可脱离界面的解析、匹配、设置迁移、计算器、文本处理、排序和部分 Provider 逻辑; 通过不代表 Windows 窗口交互一定正确。
- Windows CI 还负责语法检查和偏好设置布局检查; 截图工作流用于人工核对界面, 性能基准用于比较典型搜索工作负载。三者用途不同, 不能互相替代。

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
5. 在 `Tests\RunTests.ahk` 里加测试

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
| Uid | 学习排序和去重用的唯一标识 |
| Score | 分数, 越大越靠前 (匹配分 0~100, 再加各功能自己的偏移) |
| Valid / AutoComplete | `Valid = false` 时 Enter 只把 AutoComplete 填进输入框 |
| Exclusive | 关键字模式: 只显示有这个标记的结果 |
| Source / Provider | 设置里对应的那一条 / 产生结果的 Provider Id (由 ProviderRegistry 填) |

## 修改设置结构
- 只是加新的设置项: 加到 `AppSettings.Defaults()` 即可, 旧文件缺少的项会自动补上
- 改名、移动或改变已有项的结构: `AppSettings.CurrentVersion + 1`, 在 `SchemaMigration` 里加一个 `_FromN(data)` 把版本 N 转换成 N+1。升级前会自动备份为 `ALTRun.vN.backup.json`

## 文档
Wiki 的源文件在仓库的 `docs/wiki/`, 合并到 `main` 后自动发布到 Wiki (`.github/workflows/wiki.yml`)。请通过 PR 修改 `docs/wiki/`, 不要直接在网页上编辑。

文档是中英双语的, 改一种语言时另一种也要一起改:
- README 和贡献指南: `README.md` / `CONTRIBUTING.md` (英文, GitHub 默认显示) 和 `README.zh-CN.md` / `CONTRIBUTING.zh-CN.md` (中文), 开头互相链接
- Wiki: 每页两个文件, 中文 `Usage.md`, 英文 `en-Usage.md`; 每页开头链接到另一种语言, 英文页之间的链接写 `en-` 的页面名。侧边栏 `_Sidebar.md` 分中英两组。程序里的帮助 (速查表、偏好设置的 `F1`) 按界面语言打开: 中文界面打开中文页, 其它语言打开 `en-` 页 (`HelpProvider.WikiPage`)。`WikiPages` 测试检查每页都有英文页、链接指向存在的页面
- CHANGELOG 用中文写, 每个版本开头的概要后面加一句英文 (Release 的说明就是这一节)

官网 (https://zhugecaomao.github.io/ALTRun/) 的源文件在 `site/`: `index.html` 是中英文模板, `build.py` 填入最新版本号、下载链接和大小、下载次数、主题截图和最新版本的更新内容, 生成到 `_site/`。`.github/workflows/pages.yml` 在 main 上相关文件变化、Release 成功后和每天一次自动生成, 用 GitHub Pages 的官方 Actions 部署 (Settings → Pages → Source 选 "GitHub Actions")。本地预览: `python3 site/build.py && python3 -m http.server -d _site`。

界面截图在 `docs/images/screenshots/`, 由 `Tests\Screenshots\TakeScreenshots.ahk` 生成: 它在临时文件夹里准备一份演示用的 ALTRun (英文界面; 示例设置、自定义命令、示例文件都是虚构的通用内容, 不放个人或工作相关的信息), 逐个场景启动、输入、截图。界面改动后, 在 Actions 里运行 **Screenshots** (`.github/workflows/screenshots.yml`), 它在 Windows 上重新截图并提交回当前分支; 修改 `Tests/Screenshots/` 的推送也会自动运行。本地运行:
```
AutoHotkey64.exe Tests\Screenshots\TakeScreenshots.ahk [输出文件夹] [场景名...]
```

## 性能测试
GitHub → Actions → **Bench** → Run workflow: 在 GitHub 的 Windows 虚拟机上用真实的按键模拟日常操作, 比较两个版本
(baseline 默认最新正式版, candidate 默认选中的分支)。每个版本两轮, 两个版本轮流 (A B B A), 每轮重新启动 ALTRun:

- 随机搜索 (应用名称开头、单词首字母、随便几个字母、文件搜索、自定义命令): 呼出窗口的时间, 最后一个字到结果画好的时间
- 运行程序 / 打开文件和文件夹、操作面板 (→)、计算器 (结果 + Enter 复制)、Quick Switch (文件夹面板 + Ctrl+E 跳转)、
  打开并保存偏好设置、片段展开
- 每轮结束时 ALTRun 的 CPU 时间、内存、GDI / USER 对象和句柄数
- ALTRun 日志里一次超过 50 ms 的后台任务 (`Perf: slow`, 有这种记录的版本) 和错误

默认每轮 500 次搜索 + 每种操作 100 次 (约 50 分钟), 结果 (中位数 / p90 / p99 / 最慢) 在运行摘要里, 原始数据是 artifact。
测试副本里多 #Include 一个 `Tests\Bench\BenchProbe.ahk`, 驱动程序 (`BenchDriver.ahk`) 通过它知道 "当前输入的结果画好了没有";
发布的程序里没有它。改了测试脚本推到 main 以外的分支时会自动用很少的次数跑一遍。

## 代码规范
见仓库的 [贡献指南](https://github.com/zhugecaomao/ALTRun/blob/main/CONTRIBUTING.zh-CN.md)。几个 AutoHotkey v2 的坑:
- 名字不区分大小写: 局部变量不要和类同名 (`pinyin` 会遮住 `Pinyin` 类), 同一个类里方法和属性不要只差大小写
- 很大的 `static X := Map(...)` 会报 "Declaration too long", 改成在方法里构造
- 字符串按值传给函数时会整段复制: 在循环里反复把一个大字符串 (整个文件) 传给函数, 耗时会随长度平方增长, 要按引用传 (`&text`), 见 `Lib\JSON.ahk`
- 调试日志 (`General.SaveLog`) 里 `Perf:` 开头的行是启动各阶段、慢搜索和慢的后台任务的耗时; 用 `Logger.Ms()` / `Logger.Time(label, start)` 记录新的, 计时器的回调用 `Logger.Watch(label, fn)` 包一层 (超过 50 ms 才记)
- 不要在主线程里用 `Download()` 下载 (下载完之前整个程序都停着), 用 `HttpGet.Start(url, done, fail)` 在后台下载
- 要知道文件夹里的变化, 用 `FolderWatcher` (`Lib\FolderWatcher.ahk`, 定时 `Poll()` 不等待), 不要定时扫描整个文件夹; 关闭时它会等异步读取结束再放掉缓冲区
- 嵌套函数修改外层变量时, 通过对象传回结果 (`state := {Result: ""}`)
- 所有 `Edit` 控件都写明行数 (`r1 -Multi`), 否则长文字会让它自动变成多行
- `#Warn All` 的警告都当作错误处理
