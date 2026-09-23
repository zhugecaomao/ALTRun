<h1 align="Center"><img width="45" alt="ALTRun" src="https://github.com/user-attachments/assets/91f36c04-3dd3-4486-9e7e-f70c9fabd6b8" /> ALTRun </h1>

ALTRun - 基于 AutoHotkey v2、开源免费、轻量高效的 Windows 启动器, 操作习惯和设计参照 macOS 上的 [Alfred](https://www.alfredapp.com/)

> 3.0 版本重新设计了搜索窗口和整体架构, 界面截图会在 Windows 上实测后更新


## 特性
- **Alfred 式搜索窗口**: `Alt+Space` 呼出, 输入即搜, 结果两行显示 (标题 + 路径/说明), 窗口高度随结果伸缩
- **学习排序**: 记住 "输入了什么 → 选了哪一项", 越常用越靠前; 搜索框为空时按 `↑` 调出最近的搜索
- **操作面板**: 选中一项按 `→` 列出全部操作 (打开 / 以管理员运行 / 显示位置 / 复制路径 / 在此打开终端 / 属性 / 大字显示...)
- **修饰键**: `Ctrl+Enter` 在文件管理器中显示, `Alt+Enter` 复制路径, `Ctrl+1~9` 直接执行第 N 行
- **应用搜索**: 自动索引开始菜单、桌面和应用商店应用, 中文名称支持拼音首字母 ("wx" → 微信)
- **剪贴板历史**: `Ctrl+Alt+C` 或输入 `clip` 列出复制过的文字, `clip 关键词` 过滤, Enter 粘贴; 密码管理器复制的内容自动忽略
- **自定义命令 / 文字片段**: 文件、文件夹、程序+参数、网址; 片段支持 `{date}` `{clipboard}` `{cursor}` 等占位符
- **片段自动展开**: 在任何程序里输入 `;关键字` (例如 `;sig`) 自动替换成片段正文
- **计算器**: 直接输入算式, 可选附带梁主筋 / 配筋面积的结构计算
- **网页搜索**: `g 关键词` (Google)、`bd 关键词` (百度) 等, 引擎可自行添加; 没有结果时显示兜底搜索
- **文件 / 文件夹搜索**: 直接输入名称, 匹配的文件和文件夹显示在应用下面; `'报告` 或 `open 报告` 只搜文件。Everything 在运行时直接查询全盘, 否则使用内置索引 (桌面、文档、下载)
- **在结果里直接编辑**: 选中一项按 `F3` 或右键 "编辑...", 修改自定义命令 / 片段 / 搜索引擎; 应用、文件、网址可一键加为自定义命令; `Ctrl+Del` 删除
- **终端**: `>ipconfig /all` 在终端运行命令
- **系统命令**: 锁屏、睡眠、关机、清空回收站、音量、Windows 工具 (设备管理器、服务、注册表...)、剪贴板文字转换
- **大字显示**: `Ctrl+L` 全屏大字显示结果 (电话号码、计算结果...)
- **主题**: 9 套内置主题 (跟随系统、浅色、深色、经典、午夜、霜白、石墨、海洋、纸张), 也可以用 JSON 写自己的主题
- **偏好设置窗口**: `Ctrl+,` 打开, 按分类修改设置, 列表式编辑自定义命令 / 片段 / 搜索引擎 / 自定义热键
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
| `Ctrl+Alt+C` | 打开剪贴板历史 |
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
| `F3` | 编辑选中项 (自定义命令 / 片段 / 搜索引擎); 应用、文件、网址: 添加为自定义命令; 没有结果时用输入的文字新建命令 |
| `Ctrl+Del` (光标在末尾) | 删除选中项 (自定义命令 / 片段 / 搜索引擎 / 剪贴板历史), 删除前确认 |
| 鼠标右键 | 选中项的操作菜单 (和操作面板相同) |
| `F2` / `Ctrl+,` | 偏好设置 |
| `F4` | 用记事本编辑 ALTRun.json |
| `Esc` / 切换到其它窗口 | 隐藏 |


## 搜索功能与关键字
| 输入 | 功能 |
|---|---|
| 任意文字 | 应用、自定义命令、片段、系统命令、Windows 工具, 以及名称匹配的文件和文件夹 |
| `12*(3+4)` 或 `=2^10` | 计算器 |
| `g xxx` `bing xxx` `bd xxx` `gh xxx` `wiki xxx` `yt xxx` `tb xxx` `jd xxx` `tr xxx` | 网页搜索 (Google / Bing / 百度 / GitHub / 维基百科 / YouTube / 淘宝 / 京东 / 翻译) |
| `'xxx` 或 `open xxx` / `find xxx` | 只搜索文件和文件夹 |
| `>命令` | 在终端运行 |
| `clip` / `clip xxx` | 剪贴板历史 |
| `snip xxx` | 只搜索片段 |
| `;关键字` (在任何程序里) | 片段自动展开 |

关键字后面加了空格 (例如 `clip `、`g xxx`、`'xxx`、`>xxx`) 就进入该功能的专属模式, 只显示这个功能的结果。

文件搜索会自动选择数据来源:
- [Everything](https://www.voidtools.com/) 在运行: 通过 Everything 的 IPC 接口直接查询全盘, 不需要 `Everything64.dll` 或 `es.exe`
- 否则: 使用内置索引, 在后台扫描 `ScopeFolders` 里的文件夹 (默认: 桌面、文档、下载, 深度 4 层), 缓存在 `Data\FileIndex.json`, 每 30 分钟更新

偏好设置的 "文件搜索" 页可以修改: 是否显示在默认结果里、显示几条、Everything 的排除条件、内置索引的文件夹和深度, 以及立即重建索引。


## 偏好设置
`Ctrl+,` (搜索窗口里) 或托盘菜单 "偏好设置" 打开偏好设置窗口, 左边选择分类:

| 分类 | 内容 |
|---|---|
| 通用 | 呼出热键、界面语言、开机启动、失焦隐藏、文件管理器... |
| 外观 | 主题、窗口宽度、显示行数 |
| 功能 | 启用 / 关闭各项搜索功能, 结构计算、终端选项 |
| 应用搜索 | 索引的文件夹、文件类型、拼音首字母、应用商店应用, 重建索引 |
| 文件搜索 | 默认结果里显示文件、Everything 状态和排除条件、内置索引的文件夹和深度, 重建文件索引 |
| 自定义命令 / 文字片段 / 网页搜索 / 自定义热键 | 列表, 添加 / 编辑 (双击) / 删除 |
| 剪贴板历史 | 热键、保存条数、不记录的程序、清空历史 |
| 扩展功能 | 对话框快速跳转、Ctrl+D 加日期 |
| 高级 | 直接编辑 ALTRun.json、数据文件夹、重置学习排序、版本和更新 |

点 "保存" 后设置立即生效 (ALTRun 自动重新载入, 并回到刚才的分类)。

## 设置文件 ALTRun.json
所有设置和用户数据都保存在程序目录下的 `ALTRun.json` (UTF-8 文本), 也可以在偏好设置的 "高级" 里直接编辑, **保存文件后 ALTRun 自动重新载入**。

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

- **CustomCommands**: `Type` 可以是 `File` / `Folder` / `Command` / `Url`; `Target` 支持 `A_Desktop`、`A_ScriptDir` 等内置变量开头和 `%AppData%` 等环境变量。 搜索时除了 `Title` 和 `Keyword`, `File` / `Folder` 类型还会匹配目标的文件名 (不含扩展名) 或文件夹名, 例如 Target 是 `Q:\Projects\PT1931 - 24 NIR` 时输入 `nir` 也能找到。也可以在资源管理器里右键 "发送到 → ALTRun" 添加, 或在操作面板里选 "添加到自定义命令"。
- **Snippets**: 占位符 `{date}` `{time}` `{datetime}` `{clipboard}` `{cursor}` (粘贴后光标停在这里)。有 `Keyword` 的片段可以在任何程序里输入 `;关键字` 自动展开 (前缀见 `Features.Snippets.ExpandPrefix`, 单个片段设 `"AutoExpand": 0` 可以关闭)。
- **Clipboard**: `Features.Clipboard` 里可以修改热键、保存条数、是否保存到磁盘 (`Persist`)、不记录的程序 (`IgnoreApps`)。历史保存在 `Data\ClipboardHistory.json`。
- **Hotkeys**: 自定义热键执行一条系统命令, `WinTitle` 不为空时只在该窗口里生效。可用的命令 Id 见 `Src\Providers\SystemProvider.ahk` (例如 `Lock`、`PTTools`、`TextUpper`、`ToggleWindow`)。
- **主题**: 内置 `System` (跟随 Windows 浅色 / 深色, 系统切换时自动更新)、`Light`、`Dark`、`Classic`、`Midnight`、`Frost` (半透明)、`Graphite`、`Ocean`、`Paper`, 配色文件在 `Resources\Themes\` (升级时会被替换, 不要直接修改)。自定义主题放在 `Themes\<名称>.json`: 在偏好设置 "外观" 里点 "复制为自定义主题" 会生成一份完整的文件, 改好后选中即可; 也可以只写要改的键, 用 `"Base"` 指定从哪个主题开始, 例如 `{ "Base": "Dark", "SelectedBackground": "1D4ED8", "SelectedRadius": 8, "Opacity": 240 }`。和内置主题同名的用户主题优先; 全部可用的键见 `Src\UI\ThemeManager.ahk`。
- 运行时生成的数据放在 `Data\` 目录 (应用索引、文件索引、学习记录、剪贴板历史), 删掉只会重新生成。


## 从 2.x 升级
第一次启动 3.0 时会自动把旧的 `ALTRun.json` 升级到新格式, 原文件备份为 `ALTRun.v2.backup.json`:
- 保留: 呼出热键、开机启动等通用设置、用户命令 (File/Dir/CMD/URL → 自定义命令, Clip → 片段)、索引目录、结构计算开关、Listary 跳转、Ctrl+D 加日期、PT Tools 设置、条件热键
- 不再保留: 内置命令列表 (已由系统命令取代)、旧索引 (会重新建立)、执行历史、使用统计、旧列表外观选项

以后设置格式再有变化时, 同样会逐版本自动升级 (见 `Src\Core\SchemaMigration.ahk`)。

旧版本的 `Res\` 文件夹已改名为 `Resources\`: 启动时会自动把 `Res\` 里你自己放的文件 (例如 SPF2M 用的 `DOSBox.exe`、`SPF2M.exe`) 移到 `Resources\`, 然后删除空的 `Res\`。


## 项目结构
```
ALTRun.ahk          入口: 列出所有模块并调用 App.Start()
Lib\                通用库, 与 ALTRun 无关 (JSON, Logger, Util, TextTools, Kanji, Dialogs, Everything IPC)
Src\Core\           启动流程 (App), 设置与版本升级, 搜索模型 (SearchQuery / ResultItem), 匹配打分, 学习排序, 操作, 文件索引
Src\UI\             搜索窗口, 偏好设置窗口, 通用编辑对话框, 大字显示, 主题, 图标缓存
Src\Providers\      搜索功能: 剪贴板历史 / 应用 / 自定义命令 / 片段 / 系统命令 / 计算器 / 网页搜索 / 文件搜索 / 终端
Src\Extensions\     搜索窗口以外的功能: 片段自动展开, 对话框快速跳转, Ctrl+D 加日期, PT 工具箱, 检查更新
Resources\          随程序发布的数据文件 (Kanji.txt 简繁对照表, Themes\ 内置主题)
Tests\              单元测试
```

运行后在程序目录下还会出现: `ALTRun.json` (设置)、`Data\` (索引、学习记录、剪贴板历史, 可以删除)、`Themes\` (你自己的主题)。升级时把新版本复制覆盖到程序目录即可, 不会影响这些文件 (不要先删除 `Resources\`, 里面可能有你自己放的 SPF2M 文件)。

新增一个搜索功能只需要在 `Src\Providers\` 里加一个类 (`Id` / `Init()` / `Search(query)` 返回 `ResultItem` 数组), 在 `ALTRun.ahk` 里 `#Include`, 并在 `App.Start()` 里注册。结果可以在搜索窗口里编辑 / 删除时, 再实现可选的 `EditItem(item)` / `DeleteItem(item)` (结果的 `Source` 指向设置里对应的那一条)。


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
