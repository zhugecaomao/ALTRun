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

## 代码规范
见仓库的 [贡献指南](https://github.com/zhugecaomao/ALTRun/blob/main/CONTRIBUTING.zh-CN.md)。几个 AutoHotkey v2 的坑:
- 名字不区分大小写: 局部变量不要和类同名 (`pinyin` 会遮住 `Pinyin` 类), 同一个类里方法和属性不要只差大小写
- 很大的 `static X := Map(...)` 会报 "Declaration too long", 改成在方法里构造
- 字符串按值传给函数时会整段复制: 在循环里反复把一个大字符串 (整个文件) 传给函数, 耗时会随长度平方增长, 要按引用传 (`&text`), 见 `Lib\JSON.ahk`
- 调试日志 (`General.SaveLog`) 里 `Perf:` 开头的行是启动各阶段和慢搜索的耗时; 用 `Logger.Ms()` / `Logger.Time(label, start)` 记录新的
- 嵌套函数修改外层变量时, 通过对象传回结果 (`state := {Result: ""}`)
- 所有 `Edit` 控件都写明行数 (`r1 -Multi`), 否则长文字会让它自动变成多行
- `#Warn All` 的警告都当作错误处理
