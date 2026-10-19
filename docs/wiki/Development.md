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

## 核心模块与运行生命周期

- `App` 负责启动、重新载入和退出; `AppSettings` 读取 / 保存设置, `SchemaMigration` 负责旧设置格式升级。
- `SearchWindow` 接收输入; `SearchQuery` 解析关键字和前缀; `ProviderRegistry` 调用已启用的 Provider 并合并结果。
- `ResultItem` 统一结果数据; `FuzzyMatcher` 计算匹配分数; `Knowledge` 调整学习排序; `ActionCatalog` 执行打开、复制、运行等操作。
- `PreferencesWindow` 在设置副本上编辑, 保存后由程序重新载入, 避免控件直接修改运行中的设置对象。
- `ApplicationProvider` 与 `FileIndex` 共用 `Lib\FolderWatcher.ahk` 监视索引目录; 索引重建在后台进行, 新索引完成前继续使用旧索引。

一次典型的生命周期是: 启动并加载设置 → 初始化 Provider 和窗口 → 接收搜索输入 → 解析查询并收集结果 → 排序和显示 → 执行选中项的操作。设置变更会保存后重新载入; 退出时停止文件夹监视和后台任务, 并完成待保存的缓存处理。

后台工作要注意窗口和资源生命周期: 计时器回调通过 `Logger.Watch(label, fn)` 记录耗时; 网络请求使用 `HttpGet.Start` 异步启动, 不要在主线程里同步等待下载; 索引缓存通常在键鼠空闲后保存, 退出时会立即处理待保存数据。

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
详细步骤见 [贡献指南](https://github.com/zhugecaomao/ALTRun/blob/main/CONTRIBUTING.zh-CN.md#发布新版本), 这里是概要:

- **测试版** (每个改了程序的 PR): 版本号四段 `YYYY.MM.DD.N`, 改 `App.Version` 和 `;@Ahk2Exe-SetVersion` (`ReleaseVersion` 测试检查两者一致), `CHANGELOG.md` 加一节 (`ChangelogLinks` 测试检查比较链接); 在 PR 的分支上运行 **Release** workflow, `publish = false`, 构建好的程序在 Actions 页面底部的 Artifacts (`ALTRun_v<版本>`)
- **正式版**: 版本号三段 `YYYY.MM.DD`, 每天最多一个; 把这一轮测试版的 CHANGELOG 合成一节 (中英文摘要 + 升级 / 新增 / 改进 / 修复), 合并到 main 后在 main 上运行 **Release**, `publish = true`

Release workflow 在 Windows 上依次: 读版本号 → 下载 AutoHotkey 和 Ahk2Exe → 单元测试 → 编译 `ALTRun.exe` → 配置了 SignPath 时签名 → 检查程序图标 → 用 2.x 的 `ALTRun.ini` 做升级测试 → 编译后的程序做单实例测试 → 打包 `ALTRun_v<版本>.zip`。`publish = true` 时再: 从 CHANGELOG 取这一节作为 Release 说明 → 创建 GitHub Release (一键更新核对的 SHA256 由 GitHub 给出) → 更新 `packaging/` 里的 Scoop 和 winget 清单并提交到 main → 配置了 `WINGET_TOKEN` 时向 winget-pkgs 提交新版本。

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

## 质量验证与局限

- `Tests\RunTests.ahk`: 覆盖匹配、查询解析、设置迁移、排序、剪贴板和索引等纯逻辑; 不等于完整的手动 UI 测试。
- Windows CI: 运行单元测试和语法检查, 并检查偏好设置布局、Wiki 页面配对和文档链接。
- 截图测试: 检查主要窗口在不同语言 / 主题下的显示; 不验证所有交互流程。
- 性能基准: 在 Windows 上模拟常见操作比较改动前后的耗时, 应在相同环境下对比。
- Wiki 自动检查只能发现缺页和断链, 不能判断说明是否准确或是否覆盖每个设置项; 修改设置时仍需逐项对照代码和中英文文档。

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
