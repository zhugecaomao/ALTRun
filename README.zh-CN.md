<h1 align="center"><img width="48" alt="ALTRun" src="docs/images/logo.png" /> ALTRun</h1>

<p align="center"><a href="README.md">English</a> | <b>简体中文</b></p>

<p align="center">
  <b>轻量、开源的 Windows 启动器，操作方式参照 macOS 上的 <a href="https://www.alfredapp.com/">Alfred</a></b><br>
  按 <code>Alt+Space</code>，输入几个字母，按 <code>Enter</code>：程序、文件、网页、计算、剪贴板、系统命令，一个窗口完成
</p>

<p align="center">
  <a href="https://github.com/zhugecaomao/ALTRun/releases/latest"><img alt="Release" src="https://img.shields.io/github/v/release/zhugecaomao/ALTRun?label=release"></a>
  <a href="https://github.com/zhugecaomao/ALTRun/releases"><img alt="Downloads" src="https://img.shields.io/github/downloads/zhugecaomao/ALTRun/total"></a>
  <a href="https://github.com/zhugecaomao/ALTRun/actions/workflows/tests.yml"><img alt="Tests" src="https://github.com/zhugecaomao/ALTRun/actions/workflows/tests.yml/badge.svg?branch=main"></a>
  <a href="https://www.autohotkey.com/"><img alt="AutoHotkey v2" src="https://img.shields.io/badge/AutoHotkey-v2.0-334455?logo=autohotkey"></a>
  <img alt="Windows" src="https://img.shields.io/badge/platform-Windows%2010%20%7C%2011-0078D6">
  <a href="LICENSE"><img alt="License: GPL-3.0" src="https://img.shields.io/github/license/zhugecaomao/ALTRun"></a>
</p>

<p align="center">
  <a href="https://zhugecaomao.github.io/ALTRun/">官网</a> ·
  <a href="#为什么选择-altrun">为什么选择</a> ·
  <a href="#快速开始">快速开始</a> ·
  <a href="#功能">功能</a> ·
  <a href="#快捷键">快捷键</a> ·
  <a href="https://github.com/zhugecaomao/ALTRun/wiki">使用文档 (Wiki)</a> ·
  <a href="CHANGELOG.md">更新日志</a>
</p>

<p align="center">
  <img src="docs/images/screenshots/demo.png" width="700" alt="ALTRun：输入即搜、操作面板、单位换算、文件搜索">
</p>


## 为什么选择 ALTRun
- **小巧**：约 2.5 MB，由一个 `ALTRun.exe` 和少量资源文件组成，无需安装 .NET、Electron 或其他运行库。
- **快速响应**：输入即出结果；配合 [Everything](https://www.voidtools.com/) 可在毫秒内搜索全盘文件，否则使用内置索引。
- **智能排序**：根据使用习惯自动调整排名；支持单词首字母（`vsc` → Visual Studio Code）和拼音首字母（`wx` → 微信；文件也可以：`cs` → 测试文件.zip），并高亮匹配内容。
- **不只是启动器**：[对话框快速跳转（Quick Switch）](#功能)可将打开 / 保存对话框快速切换到 Total Commander、Directory Opus 或资源管理器中已打开的文件夹；可操作任意程序中选中的文字、文件或网址；输入 `;关键字` 可展开文字片段。
- **多合一**：集成计算与单位换算、网页搜索、浏览器书签、剪贴板历史、文字片段、系统命令和终端，无需为每项任务安装单独工具。
- **键盘优先**：按 `→` 打开操作面板，状态栏提示 `Enter` 的操作，按 `F3` 直接编辑，用 `Ctrl+1`～`Ctrl+9` 快速打开对应结果；按 `Ctrl+Enter` 可打开文件所在文件夹，或将文本粘贴到当前活动窗口。
- **便携与隐私**：免安装、不写注册表、无需管理员权限，无后台服务或驱动。数据默认保存在本机；部分剪贴板数据存于 `%LOCALAPPDATA%\ALTRun`。不收集遥测数据；联网用于更新、可选汇率换算和用户主动下载的网站图标。
- **自动更新**：新版本会在搜索窗口中提示，按 `Enter` 即可安装；更新包会校验 SHA256 并保留设置，也支持 Scoop。
- **个性化定制**：支持简体中文、繁體中文、English、日本語四种界面语言和 22 套内置主题（每套都有宽松 / 常规 / 紧凑三种密度）；可跟随 Windows 浅色 / 深色模式，也可通过 JSON 文件自定义主题和设置。
- **免费开源**：采用 GPL-3.0 许可证；每个 Pull Request 都会在 Windows 上自动运行 8,000 余项测试。

## 快速开始
1. [下载最新版本](https://github.com/zhugecaomao/ALTRun/releases/latest)，解压到任意文件夹，运行 `ALTRun.exe`——无需安装，也无需 AutoHotkey。
2. 按 `Alt+Space`（或 `Alt+R`），输入程序或文件名称，再按 `Enter`。
3. 输入 `?` 查看全部搜索语法和快捷键；按 `Ctrl+,` 打开偏好设置。

如果 Windows SmartScreen 拦截程序，点击 **更多信息** → **仍要运行**。杀毒软件误报请参阅[常见问题](https://github.com/zhugecaomao/ALTRun/wiki/FAQ-zh#杀毒软件报毒)。

也可以通过 Scoop 或 winget（审核收录后）安装，或使用 AutoHotkey v2 直接运行 `ALTRun.ahk`。详情请参阅[其他安装方式与升级](#其他安装方式与升级)。


## 功能
**搜索**
- **搜索窗口**：输入即搜，显示标题和路径；窗口高度随结果变化，可拖动并记住位置，多显示器时显示在鼠标所在屏幕。
- **应用**：自动索引开始菜单、桌面和 Microsoft Store 应用，新装的软件几秒后就能搜到；不需要的应用和内置命令可按 `Ctrl+Del` 隐藏，在偏好设置中恢复。
- **匹配高亮**：标题中与输入匹配的部分（连续字符、单词首字母、拼音首字母）以高亮色显示。
- **文件和文件夹**：在空白搜索框中先按 `空格` 再输入名称（或 `'报告`、`open 报告`）；Everything 运行时搜索全盘，否则使用内置索引；输入 `D:\Projects\` 这样的路径可以逐级浏览文件夹；`recent 报告` 搜最近在任何程序里打开过的文件。也可以让文件直接显示在普通结果里，按文件名开头或文件名包含匹配（偏好设置 → 文件搜索）。
- **自定义命令**：文件、文件夹、程序（可带参数）、网址，可设置关键字（可写多个，用逗号分隔）；在资源管理器中右键 → 发送到 → ALTRun 即可添加（可多选）；可检查路径已失效的命令。
- **学习排序**：记住每次输入所选的结果，常用项自动靠前。
- **置顶和最近使用**：呼出窗口还没输入时，直接列出置顶的项目和最近打开的项目。
- **使用统计**：按天、按功能统计使用次数（仅记录次数）。
- **计算器**：直接输入算式；支持单位换算（`10 km in mi`）、进制换算（`255 in hex`）、日期加减（`today + 30 days`）和可选的货币换算（`100 usd to sgd`）。
- **网页搜索**：`g 关键词`（Google）、`bd 关键词`（百度）等，可自定义搜索引擎（输入 `g` 加空格、不输入搜索词就搜索剪贴板里的文字）；无结果时提供网页搜索。
- **浏览器书签**：搜索 Chrome、Edge、Brave、Vivaldi 的书签；`bm 关键词` 仅搜索书签。
- **切换窗口**：输入窗口标题或程序名切换到已打开的窗口；`w 关键词` 仅搜索窗口，`w ` 列出全部。
- **系统命令**：锁屏、睡眠、关机、清空回收站、音量和媒体控制、Windows 工具、Windows 设置页面（显示、蓝牙、WLAN 等 40 多个）、文本转换（大小写、排序、简繁转换等）。

**效率**
- **操作面板**：选中结果后按 `→`（或右键）列出可用操作，如以管理员身份运行、打开所在位置、复制路径、在此处打开终端、属性。
- **选中内容操作**：在任意程序中选中文字、文件或网址，按 `Ctrl+Alt+\` 调出操作，如网页搜索、保存为片段、转换后替换原文。
- **直接编辑**：按 `F3` 编辑命令、片段或搜索引擎；应用和文件可一键添加为自定义命令。
- **剪贴板历史**：按 `Ctrl+Alt+C` 或输入 `clip`；同时记录文字、文件和图片；常用的可以置顶；可以粘贴为纯文本；自动忽略密码管理器复制的内容。
- **文字片段**：可按名称、关键字或正文搜索；支持 `{date}`、`{clipboard}`、`{cursor}` 等占位符；在任意程序中输入 `;关键字` 自动展开；可从 Excel / CSV 批量导入，也可导出。
- **终端**：输入 `>ipconfig /all` 直接在终端中运行。
- **大字显示**：按 `Ctrl+L` 全屏显示结果，便于查看电话号码、计算结果等（在偏好设置 → 功能里打开）。

**扩展**
- **对话框快速跳转（Quick Switch）**：在打开 / 保存对话框中按 `Ctrl+G` 跳转到最近使用的文件管理器（Total Commander、Directory Opus 或资源管理器）的当前目录；对话框旁的文件夹面板列出所有已打开和最近使用的文件夹。只有文件夹树的旧式“浏览文件夹”对话框也能用。
- **一键加日期**：重命名文件时按 `Ctrl+D`，在扩展名前添加或更新日期（`Report.docx` → `Report - 28.09.2026.docx`）。
- **自定义热键**：为任意系统命令设置热键，可限定在指定程序中生效。
- **脚本扩展**：把 `.ahk`、`.ps1`、`.bat`、`.py` 脚本放进 `Scripts\` 文件夹，就能在搜索窗口里运行，可以带参数、在后台运行并显示输出。

**结构工程师专用**
- **结构计算**（默认关闭）：计算器结果下方附带梁主筋和配筋面积计算。
- **PT 工具箱**：钢筋 / BRC 面积计算、SPF2M 后张预应力束线型计算，结果可复制到 Excel。


## 截图
| 搜索窗口 | 操作面板（`→`） |
|:---:|:---:|
| <img src="docs/images/screenshots/search.png" alt="搜索窗口"> | <img src="docs/images/screenshots/actions.png" alt="操作面板"> |
| **文件搜索（`空格` + 名称）** | **按类型搜索（`doc` + 名称）** |
| <img src="docs/images/screenshots/files.png" alt="文件搜索"> | <img src="docs/images/screenshots/filetype.png" alt="按类型搜索"> |
| **计算器** | **拼音首字母（`jsb` 找到记事本）** |
| <img src="docs/images/screenshots/calculator.png" alt="计算器"> | <img src="docs/images/screenshots/pinyin.png" alt="拼音首字母"> |
| **剪贴板历史（`clip`）** | **置顶和最近使用（空搜索框）** |
| <img src="docs/images/screenshots/clipboard.png" alt="剪贴板历史"> | <img src="docs/images/screenshots/empty.png" alt="置顶和最近使用"> |
| **浏览文件夹（输入路径）** | **网页搜索（`g` + 关键词）** |
| <img src="docs/images/screenshots/browse.png" alt="浏览文件夹"> | <img src="docs/images/screenshots/websearch.png" alt="网页搜索"> |
| **对话框快速跳转（打开 / 保存对话框下面的文件夹面板）** | **主题（偏好设置 → 外观）** |
| <img src="docs/images/screenshots/quickswitch.png" alt="对话框快速跳转"> | <img src="docs/images/screenshots/prefs-appearance.png" alt="主题"> |
| **窗口切换（`w` + 空格）** | **从资源管理器添加命令（右键 → 发送到 → ALTRun）** |
| <img src="docs/images/screenshots/windows.png" alt="窗口切换"> | <img src="docs/images/screenshots/sendto.png" alt="发送到 ALTRun"> |
| **偏好设置** | **自定义命令** |
| <img src="docs/images/screenshots/prefs-general.png" alt="偏好设置"> | <img src="docs/images/screenshots/prefs-commands.png" alt="自定义命令"> |
| **文字片段** | **脚本扩展** |
| <img src="docs/images/screenshots/prefs-snippets.png" alt="文字片段"> | <img src="docs/images/screenshots/prefs-scripts.png" alt="脚本扩展"> |

<details>
<summary><b>内置主题</b>（共 22 套，每套都有宽松 / 常规 / 紧凑三种密度，点击展开）</summary>

| Dark | Classic | Midnight |
|:---:|:---:|:---:|
| <img src="docs/images/screenshots/theme-dark.png" alt="Dark"> | <img src="docs/images/screenshots/theme-classic.png" alt="Classic"> | <img src="docs/images/screenshots/theme-midnight.png" alt="Midnight"> |
| **Frost** | **Graphite** | **Ocean** |
| <img src="docs/images/screenshots/theme-frost.png" alt="Frost"> | <img src="docs/images/screenshots/theme-graphite.png" alt="Graphite"> | <img src="docs/images/screenshots/theme-ocean.png" alt="Ocean"> |
| **Paper** | **Lavender** | **Rose** |
| <img src="docs/images/screenshots/theme-paper.png" alt="Paper"> | <img src="docs/images/screenshots/theme-lavender.png" alt="Lavender"> | <img src="docs/images/screenshots/theme-rose.png" alt="Rose"> |
| **Forest** | **Sunset** | **HighContrast** |
| <img src="docs/images/screenshots/theme-forest.png" alt="Forest"> | <img src="docs/images/screenshots/theme-sunset.png" alt="Sunset"> | <img src="docs/images/screenshots/theme-highcontrast.png" alt="HighContrast"> |
| **Dark · 紧凑** | **TokyoNight** | **Dracula** |
| <img src="docs/images/screenshots/theme-compact.png" alt="Dark · 紧凑"> | <img src="docs/images/screenshots/theme-tokyonight.png" alt="TokyoNight"> | <img src="docs/images/screenshots/theme-dracula.png" alt="Dracula"> |
| **CatppuccinMocha** | **GruvboxDark** | **SolarizedLight** |
| <img src="docs/images/screenshots/theme-catppuccinmocha.png" alt="CatppuccinMocha"> | <img src="docs/images/screenshots/theme-gruvboxdark.png" alt="GruvboxDark"> | <img src="docs/images/screenshots/theme-solarizedlight.png" alt="SolarizedLight"> |
| **DarkModern** | **LightModern** | **Monokai** |
| <img src="docs/images/screenshots/theme-darkmodern.png" alt="DarkModern"> | <img src="docs/images/screenshots/theme-lightmodern.png" alt="LightModern"> | <img src="docs/images/screenshots/theme-monokai.png" alt="Monokai"> |
| **OneDark** | |  |
| <img src="docs/images/screenshots/theme-onedark.png" alt="OneDark"> | |  |

</details>

截图由 [Tests/Screenshots](Tests/Screenshots/TakeScreenshots.ahk) 在 GitHub Actions 的 Windows 环境中自动生成。


## 快捷键
| 按键 | 功能 |
|---|---|
| `Alt+Space` | 显示 / 隐藏搜索窗口（可在设置中修改） |
| `Enter` | 打开选中项 |
| `Ctrl+Enter` | 文件 / 文件夹：打开所在位置；文字：粘贴到当前窗口 |
| `Alt+Enter` | 复制路径、网址或文字 |
| `Ctrl+1`～`Ctrl+9` | 打开第 N 项 |
| `↑` `↓` / `Ctrl+P` `Ctrl+N` | 上移 / 下移一行 |
| `PgUp` `PgDn` | 上翻 / 下翻一页 |
| `Ctrl+↑` / `Ctrl+↓` | 上一条 / 下一条搜索记录 |
| `Tab` | 自动补全；文件夹：进入浏览 (也可以设成选择下一个结果, 见 偏好设置 → 键盘) |
| `Backspace`（浏览文件夹时） | 返回上一级 |
| `Insert` | 标记多个文件 / 文件夹，再按 `→` 一起操作 |
| `空格`（搜索框为空时，或光标在最前面时） | 进入文件搜索模式，已输入的文字保留；搜索框为空或光标在最前面时按 `Backspace` 返回 |
| `folder 名称` | 仅搜索文件夹 |
| `?` | 速查表：全部语法和快捷键 |
| `→` / 右键 | 操作面板 / 操作菜单 |
| `F3` | 编辑选中项，或将应用、文件、网址添加为自定义命令 |
| `Ctrl+Del` | 删除选中项（需确认）；应用：从结果中隐藏 |
| `Ctrl+C` / `Ctrl+L` | 复制选中项 / 大字显示 |
| `Ctrl+E` | 用编辑程序打开选中的文件（例如 .ahk 脚本） |
| `F2` 或 `Ctrl+,` / `F4` | 偏好设置 / 用记事本编辑 ALTRun.json |
| `F1` | 关于 ALTRun（版本、检查更新、项目主页） |
| `Ctrl+Alt+C` | 剪贴板历史 |
| `Esc` | 关闭操作面板 / 隐藏窗口 |

完整语法（`'文件`、`>命令`、`clip`、`snip`、网页搜索关键字等）见 Wiki：[搜索与快捷键](https://github.com/zhugecaomao/ALTRun/wiki/Usage-zh)。


## 文档
完整文档见 [Wiki](https://github.com/zhugecaomao/ALTRun/wiki)：

| 页面 | 内容 |
|---|---|
| [安装与升级](https://github.com/zhugecaomao/ALTRun/wiki/Installation-zh) | 下载、Scoop / winget、开机启动、自动更新、从 2.x 升级、卸载 |
| [搜索与快捷键](https://github.com/zhugecaomao/ALTRun/wiki/Usage-zh) | 输入语法、快捷键、操作面板、学习排序、使用统计 |
| [自定义命令](https://github.com/zhugecaomao/ALTRun/wiki/Custom-Commands-zh) | 命令类型、带参数的命令、路径变量 |
| [文字片段](https://github.com/zhugecaomao/ALTRun/wiki/Snippets-zh) | 占位符、自动展开、不展开的窗口、导入 / 导出 |
| [文件搜索](https://github.com/zhugecaomao/ALTRun/wiki/File-Search-zh) | Everything 联动、内置索引、排除规则 |
| [主题](https://github.com/zhugecaomao/ALTRun/wiki/Themes-zh) | 内置主题、自定义主题、全部可用的键 |
| [扩展功能](https://github.com/zhugecaomao/ALTRun/wiki/Extensions-zh) | 对话框快速跳转、一键加日期、自定义热键、系统命令、PT 工具箱 |
| [设置文件参考](https://github.com/zhugecaomao/ALTRun/wiki/Configuration-zh) | ALTRun.json 各项的含义和默认值 |
| [常见问题](https://github.com/zhugecaomao/ALTRun/wiki/FAQ-zh) | 热键冲突、搜索不到、杀毒软件误报等 |
| [开发指南](https://github.com/zhugecaomao/ALTRun/wiki/Development-zh) | 架构、新增搜索功能、代码规范、测试 |


## 其他安装方式与升级
**通过 [Scoop](https://scoop.sh/) 安装**（自动创建开始菜单快捷方式，升级时保留设置）：
```powershell
scoop bucket add altrun https://github.com/zhugecaomao/ALTRun
scoop install altrun
scoop update altrun    # 升级（请先退出 ALTRun）
```

winget 清单已提交，等待审核收录；收录后可使用 `winget install zhugecaomao.ALTRun` 安装。

从源码运行：安装 [AutoHotkey v2](https://www.autohotkey.com/)，然后运行 `ALTRun.ahk`。

从较早版本或 2.x 升级，请参阅[安装与升级](https://github.com/zhugecaomao/ALTRun/wiki/Installation-zh)。


## 贡献与反馈
- 报告问题：[提交 Issue](https://github.com/zhugecaomao/ALTRun/issues/new/choose)（请附上 Windows 版本和复现步骤）
- 建议与讨论：[Discussions](https://github.com/zhugecaomao/ALTRun/discussions)
- 贡献代码：请先阅读 [贡献指南](CONTRIBUTING.md#参与贡献)；项目结构、运行测试、新增搜索功能见 Wiki 的[开发指南](https://github.com/zhugecaomao/ALTRun/wiki/Development-zh)

如果 ALTRun 对你有帮助，欢迎点亮星标 ⭐


## Code signing policy
Free code signing provided by [SignPath.io](https://about.signpath.io/), certificate by [SignPath Foundation](https://signpath.org/).

- Only `ALTRun.exe` built from this repository's source by the [Release workflow](.github/workflows/release.yml) is signed.
- Committers and reviewers: [zhugecaomao](https://github.com/zhugecaomao)
- Approvers: [zhugecaomao](https://github.com/zhugecaomao)
- Privacy policy: [SECURITY.md](SECURITY.md#privacy-policy)（中文：[隐私说明](SECURITY.md#隐私说明)）


## 许可证与致谢
ALTRun 以 [GPL-3.0](LICENSE) 许可开源 © zhugecaomao

感谢这些项目：
- [ALTRun](https://github.com/etworker/ALTRun)（etworker，Delphi）：名字和最初的设计
- [RunZ](https://github.com/goreliu/runz)（goreliu，AutoHotkey）：用 AutoHotkey 写启动器的思路
- [Alfred](https://www.alfredapp.com/)：操作方式
- [Listary](https://www.listary.com/)：对话框快速跳转参考了它的 Quick Switch
- [pypinyin](https://github.com/mozillazg/python-pinyin)（Copyright (c) 2016 mozillazg, 闲耘；[MIT](https://github.com/mozillazg/python-pinyin/blob/master/LICENSE.txt)）：生僻字和繁体字的拼音首字母
