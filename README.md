<h1 align="Center"><img width="45" alt="ALTRun" src="https://github.com/user-attachments/assets/91f36c04-3dd3-4486-9e7e-f70c9fabd6b8" /> ALTRun </h1>

ALTRun - 基于 AutoHotkey v2、开源免费、轻量高效的 Windows 启动器, 操作习惯和设计参照 macOS 上的 [Alfred](https://www.alfredapp.com/)

> 3.0 版本重新设计了搜索窗口和整体架构, 界面截图会在 Windows 上实测后更新


## 特性
- **Alfred 式搜索窗口**: `Alt+Space` 呼出, 输入即搜, 结果两行显示 (标题 + 路径/说明), 窗口高度随结果伸缩
- **学习排序**: 记住 "输入了什么 → 选了哪一项", 越常用越靠前; 搜索框为空时按 `↑` 调出最近的搜索
- **操作面板**: 选中一项按 `→` 列出全部操作 (打开 / 以管理员运行 / 显示位置 / 复制路径 / 在此打开终端 / 属性 / 大字显示...)
- **修饰键**: `Ctrl+Enter` 在文件管理器中显示, `Alt+Enter` 复制路径, `Ctrl+1~9` 直接执行第 N 行
- **应用搜索**: 自动索引开始菜单、桌面和应用商店应用, 中文名称支持拼音首字母 ("wx" → 微信)
- **自定义命令 / 文字片段**: 文件、文件夹、程序+参数、网址; 片段支持 `{date}` `{clipboard}` `{cursor}` 等占位符
- **计算器**: 直接输入算式, 可选附带梁主筋 / 配筋面积的结构计算
- **网页搜索**: `g 关键词` (Google)、`bd 关键词` (百度) 等, 引擎可自行添加; 没有结果时显示兜底搜索
- **文件搜索**: `'报告` 或 `open 报告`, 通过 Everything 搜索文件
- **终端**: `>ipconfig /all` 在终端运行命令
- **系统命令**: 锁屏、睡眠、关机、清空回收站、音量、Windows 工具 (设备管理器、服务、注册表...)、剪贴板文字转换
- **大字显示**: `Ctrl+L` 全屏大字显示结果 (电话号码、计算结果...)
- **主题**: 内置浅色 / 深色, 可以用 JSON 写自己的主题
- **对话框快速跳转**: 打开/保存对话框里 `Ctrl+G` 跳到 Total Commander 目录, `Ctrl+E` 跳到资源管理器目录
- **Ctrl+D 加日期**: 重命名文件时在扩展名前加上日期, 备注框里在末尾加日期
- **PT 工具箱**: 钢筋 / BRC 面积计算器、SPF2M 束线型计算器
- 绿色便携, 不写注册表; 设置保存在程序目录下的 `ALTRun.json`
- 中英文界面自动切换


## 快速开始
1. [下载程序](https://github.com/zhugecaomao/ALTRun/releases), 或安装 [AutoHotkey v2](https://www.autohotkey.com/) 后直接运行 `ALTRun.ahk`
2. 按 `Alt+Space` 呼出搜索窗口, 输入关键词, `Enter` 执行


## 搜索窗口快捷键
| 按键 | 作用 |
|---|---|
| `Alt+Space` | 显示 / 隐藏搜索窗口 (可在设置里修改) |
| `Enter` | 执行选中项 (打开文件/程序/网址, 复制计算结果, 粘贴片段...) |
| `Ctrl+Enter` | 文件/文件夹: 在文件管理器中显示; 文字: 粘贴到前台窗口 |
| `Alt+Enter` | 复制路径 / 网址 / 文字 |
| `Ctrl+1` ~ `Ctrl+9` | 直接执行可见的第 N 行 |
| `↑` `↓` `PgUp` `PgDn` `Ctrl+P` `Ctrl+N` | 移动选择 |
| `↑` (搜索框为空时) | 调出最近的搜索 |
| `Tab` | 自动补全 |
| `→` (光标在末尾) | 打开操作面板, `←` / `Esc` 返回 |
| `Ctrl+C` | 复制选中项 (输入框里选中了文字时照常复制文字) |
| `Ctrl+L` | 大字显示 |
| `Ctrl+,` | 偏好设置 |
| `Esc` / 切换到其它窗口 | 隐藏 |


## 搜索功能与关键字
| 输入 | 功能 |
|---|---|
| 任意文字 | 应用、自定义命令、片段、系统命令、Windows 工具 |
| `12*(3+4)` 或 `=2^10` | 计算器 |
| `g xxx` `bing xxx` `bd xxx` `gh xxx` `wiki xxx` `yt xxx` `tb xxx` `jd xxx` `tr xxx` | 网页搜索 (Google / Bing / 百度 / GitHub / 维基百科 / YouTube / 淘宝 / 京东 / 翻译) |
| `'xxx` 或 `open xxx` / `find xxx` | 文件搜索 (Everything) |
| `>命令` | 在终端运行 |
| `snip xxx` | 只搜索片段 |

文件搜索需要安装 [Everything](https://www.voidtools.com/), 并把 `Everything64.dll` (Everything SDK) 或 `es.exe` (命令行版) 放在 ALTRun 目录; 都没有时会改为在 Everything 或 Windows 搜索里打开。


## 设置 ALTRun.json
所有设置和用户数据都在程序目录下的 `ALTRun.json` (UTF-8 文本)。`Ctrl+,` 或托盘菜单 "偏好设置" 用记事本打开它, **保存后 ALTRun 自动重新载入**。
(图形化的偏好设置窗口会在后续版本加入。)

```jsonc
{
  "SchemaVersion": 3,
  "General":    { "Hotkey": "!Space", "Language": "auto", "LaunchAtLogin": 1, "HideOnDeactivate": 1, ... },
  "Appearance": { "Theme": "Light", "Width": 700, "VisibleRows": 8 },
  "Features": {
    "Applications": { "Enabled": 1, "Folders": [...], "MatchPinyin": 1, ... },
    "WebSearch":    { "Engines": [ { "Id": "google", "Keyword": "g", "Title": "Google", "Url": "https://www.google.com/search?q={query}" } ],
                      "Fallbacks": ["google", "files", "bing"] },
    "Calculator":   { "StructuralCalc": 0 },
    "Terminal":     { "Prefix": ">", "Shell": "cmd" },            // cmd / powershell / pwsh / wt
    ...
  },
  "CustomCommands": [
    { "Title": "Desktop", "Type": "Folder", "Target": "A_Desktop", "Arguments": "", "Keyword": "" },
    { "Title": "IP Configuration", "Type": "Command", "Target": "cmd.exe", "Arguments": "/k ipconfig /all", "Keyword": "ipconfig" }
  ],
  "Snippets": [ { "Name": "Today's date", "Keyword": "today", "Text": "{date}" } ],
  "Hotkeys":  [ { "Key": "~MButton", "Action": "PTTools", "WinTitle": "ahk_exe RAPTW.exe" } ]
}
```

- **CustomCommands**: `Type` 可以是 `File` / `Folder` / `Command` / `Url`; `Target` 支持 `A_Desktop`、`A_ScriptDir` 等内置变量开头和 `%AppData%` 等环境变量。也可以在资源管理器里右键 "发送到 → ALTRun" 添加, 或在操作面板里选 "添加到自定义命令"。
- **Snippets**: 占位符 `{date}` `{time}` `{datetime}` `{clipboard}` `{cursor}` (粘贴后光标停在这里)。
- **Hotkeys**: 自定义热键执行一条系统命令, `WinTitle` 不为空时只在该窗口里生效。可用的命令 Id 见 `Src\Providers\SystemProvider.ahk` (例如 `Lock`、`PTTools`、`TextUpper`、`ToggleWindow`)。
- **主题**: 在 `Themes\<名称>.json` 里写出要修改的键 (颜色 `RRGGBB`、字号、行高...), 然后设置 `"Theme": "<名称>"`; 可用的键见 `Src\UI\ThemeManager.ahk`。
- 运行时生成的数据放在 `Data\` 目录 (应用索引、学习记录), 删掉只会重新生成。


## 从 2.x 升级
第一次启动 3.0 时会自动把旧的 `ALTRun.json` 升级到新格式, 原文件备份为 `ALTRun.v2.backup.json`:
- 保留: 呼出热键、开机启动等通用设置、用户命令 (File/Dir/CMD/URL → 自定义命令, Clip → 片段)、索引目录、结构计算开关、Listary 跳转、Ctrl+D 加日期、PT Tools 设置、条件热键
- 不再保留: 内置命令列表 (已由系统命令取代)、旧索引 (会重新建立)、执行历史、使用统计、旧列表外观选项

以后设置格式再有变化时, 同样会逐版本自动升级 (见 `Src\Core\SchemaMigration.ahk`)。


## 项目结构
```
ALTRun.ahk          入口: 列出所有模块并调用 App.Start()
Lib\                通用库, 与 ALTRun 无关 (JSON, Logger, Util, TextTools, Kanji, Dialogs)
Src\Core\           启动流程 (App), 设置与版本升级, 搜索模型 (SearchQuery / ResultItem), 匹配打分, 学习排序, 操作
Src\UI\             搜索窗口, 大字显示, 主题, 图标缓存
Src\Providers\      搜索功能: 应用 / 自定义命令 / 片段 / 系统命令 / 计算器 / 网页搜索 / 文件搜索 / 终端
Src\Extensions\     搜索窗口以外的功能: 对话框快速跳转, Ctrl+D 加日期, PT 工具箱, 检查更新
Res\                数据文件 (Kanji.txt 简繁对照表)
Tests\              单元测试
```

新增一个搜索功能只需要在 `Src\Providers\` 里加一个类 (`Id` / `Init()` / `Search(query)` 返回 `ResultItem` 数组), 在 `ALTRun.ahk` 里 `#Include`, 并在 `App.Start()` 里注册。


## 开发
单元测试不依赖界面, 运行后输出结果, 退出码为失败的数量:
```
AutoHotkey64.exe /ErrorStdOut Tests\RunTests.ahk
```


## 贡献
欢迎提交 Issue 和 PR, 或在 [Discussions](https://github.com/zhugecaomao/ALTRun/discussions) 交流建议

如果你喜欢这个项目, 请给它一个星标 ⭐


## 其他说明
更多详细用法、FAQ、进阶技巧请见 [Wiki](https://github.com/zhugecaomao/ALTRun/wiki)

特别感谢 [ALTRun by etworker](https://github.com/etworker/ALTRun) (Delphi)、[RunZ by goreliu](https://github.com/goreliu/runz) (AutoHotkey), 以及 [Alfred](https://www.alfredapp.com/) 的设计
