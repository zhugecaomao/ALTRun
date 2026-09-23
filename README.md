<h1 align="center"><img width="45" alt="ALTRun" src="https://github.com/user-attachments/assets/91f36c04-3dd3-4486-9e7e-f70c9fabd6b8" /> ALTRun</h1>

<p align="center">
  <b>轻量、高效、开源的 Windows 启动器, 操作习惯参照 macOS 上的 <a href="https://www.alfredapp.com/">Alfred</a></b><br>
  A lightweight, Alfred-style launcher for Windows, written in AutoHotkey v2
</p>

<p align="center">
  <a href="https://github.com/zhugecaomao/ALTRun/releases/latest"><img alt="Release" src="https://img.shields.io/github/v/release/zhugecaomao/ALTRun?label=release"></a>
  <a href="https://github.com/zhugecaomao/ALTRun/releases"><img alt="Downloads" src="https://img.shields.io/github/downloads/zhugecaomao/ALTRun/total"></a>
  <a href="https://www.autohotkey.com/"><img alt="AutoHotkey v2" src="https://img.shields.io/badge/AutoHotkey-v2.0-334455?logo=autohotkey"></a>
  <img alt="Windows" src="https://img.shields.io/badge/platform-Windows%2010%20%7C%2011-0078D6">
  <a href="LICENSE"><img alt="License: GPL-3.0" src="https://img.shields.io/github/license/zhugecaomao/ALTRun"></a>
</p>

<p align="center">
  <a href="#快速开始">快速开始</a> ·
  <a href="#特性">特性</a> ·
  <a href="#快捷键">快捷键</a> ·
  <a href="https://github.com/zhugecaomao/ALTRun/wiki">使用文档 (Wiki)</a> ·
  <a href="CHANGELOG.md">更新日志</a> ·
  <a href="#english">English</a>
</p>

> 3.0 重新设计了搜索窗口和整体架构。界面截图会在 Windows 上实测后更新。


## 快速开始
1. 下载 [最新版本](https://github.com/zhugecaomao/ALTRun/releases/latest) 解压到任意文件夹; 或者安装 [AutoHotkey v2](https://www.autohotkey.com/) 后直接运行源码里的 `ALTRun.ahk`
2. 按 `Alt+Space` 呼出搜索窗口, 输入名称, `Enter` 打开
3. `Ctrl+,` 打开偏好设置, 修改热键、主题、索引范围等

绿色便携, 不写注册表: 设置保存在程序目录下的 `ALTRun.json`。从 2.x 升级时旧设置会自动转换 (见 [安装与升级](https://github.com/zhugecaomao/ALTRun/wiki/Installation))。


## 特性
**搜索**
- **Alfred 式搜索窗口**: 输入即搜, 每行显示标题 + 路径/说明, 窗口高度随结果伸缩, 9 套内置主题
- **应用**: 自动索引开始菜单、桌面和应用商店应用; 中文名称支持拼音首字母 ("wx" → 微信); 不需要的应用按 `Ctrl+Del` 从结果中删除
- **文件和文件夹**: 直接输入名称就能找到; [Everything](https://www.voidtools.com/) 在运行时查询全盘, 否则使用内置索引 (桌面、文档、下载)
- **自定义命令**: 文件、文件夹、程序+参数、网址, 可设关键字; 也按目标的文件夹名 / 文件名匹配
- **学习排序**: 记住 "输入了什么 → 选了哪一项", 常用的自动靠前
- **计算器**: 直接输入算式, 可选附带梁主筋 / 配筋面积的结构计算
- **网页搜索**: `g 关键词` (Google)、`bd 关键词` (百度) 等, 引擎可自行添加; 没有结果时给出兜底搜索
- **系统命令**: 锁屏、睡眠、关机、清空回收站、音量、Windows 工具、剪贴板文字转换 (大小写 / 排序 / 简繁转换...)

**效率**
- **操作面板**: 选中一项按 `→` 列出全部操作 (以管理员运行、显示位置、复制路径、在此打开终端、属性...), 右键也可以
- **在结果里直接编辑**: `F3` 修改命令 / 片段 / 搜索引擎, 应用和文件一键加为自定义命令
- **剪贴板历史**: `Ctrl+Alt+C` 或输入 `clip`, Enter 粘贴; 密码管理器复制的内容自动忽略
- **文字片段**: 支持 `{date}` `{clipboard}` `{cursor}` 等占位符; 在任何程序里输入 `;关键字` 自动展开
- **终端**: `>ipconfig /all` 直接在终端运行
- **大字显示**: `Ctrl+L` 全屏显示结果 (电话号码、计算结果...)

**扩展**
- **对话框快速跳转**: 打开 / 保存对话框里 `Ctrl+G` 跳到 Total Commander 当前目录, `Ctrl+E` 跳到资源管理器当前目录
- **Ctrl+D 加日期**: 重命名文件时在扩展名前加上日期
- **自定义热键**: 任意热键执行一条系统命令, 可限定在某个程序里生效
- **PT 工具箱**: 钢筋 / BRC 面积计算器、SPF2M 束线型计算器 (预应力设计)


## 快捷键
| 按键 | 作用 |
|---|---|
| `Alt+Space` | 显示 / 隐藏搜索窗口 (可在设置里修改) |
| `Enter` | 执行选中项 |
| `Ctrl+Enter` | 文件 / 文件夹: 在文件管理器中显示; 文字: 粘贴到前台窗口 |
| `Alt+Enter` | 复制路径 / 网址 / 文字 |
| `Ctrl+1` ~ `Ctrl+9` | 直接执行第 N 行 |
| `↑` `↓` `PgUp` `PgDn` `Ctrl+P` `Ctrl+N` | 移动选择; 搜索框为空时 `↑` 调出最近的搜索 |
| `Tab` | 自动补全 |
| `→` / 鼠标右键 | 操作面板 / 操作菜单 |
| `F3` | 编辑选中项; 应用、文件、网址: 添加为自定义命令; 没有结果时用输入的文字新建命令 |
| `Ctrl+Del` | 删除选中项 (删除前确认); 应用: 从搜索结果中删除, 可在偏好设置里恢复 |
| `Ctrl+C` / `Ctrl+L` | 复制选中项 / 大字显示 |
| `F2` 或 `Ctrl+,` / `F4` | 偏好设置 / 用记事本编辑 ALTRun.json |
| `Ctrl+Alt+C` | 剪贴板历史 |
| `Esc` | 关闭操作面板 / 隐藏窗口 |

完整的输入语法 (`'文件`、`>命令`、`clip`、`snip`、网页搜索关键字...) 见 Wiki: [搜索与快捷键](https://github.com/zhugecaomao/ALTRun/wiki/Usage)。


## 文档
使用文档都在 [Wiki](https://github.com/zhugecaomao/ALTRun/wiki):

| 页面 | 内容 |
|---|---|
| [安装与升级](https://github.com/zhugecaomao/ALTRun/wiki/Installation) | 下载、运行、开机启动、从 2.x 升级、卸载 |
| [搜索与快捷键](https://github.com/zhugecaomao/ALTRun/wiki/Usage) | 所有输入语法、快捷键、操作面板 |
| [自定义命令与片段](https://github.com/zhugecaomao/ALTRun/wiki/Commands-and-Snippets) | 命令类型、路径变量、片段占位符、自动展开 |
| [文件搜索](https://github.com/zhugecaomao/ALTRun/wiki/File-Search) | Everything 联动、内置索引、排除规则 |
| [主题](https://github.com/zhugecaomao/ALTRun/wiki/Themes) | 内置主题、自定义主题、全部可用的键 |
| [扩展功能](https://github.com/zhugecaomao/ALTRun/wiki/Extensions) | 对话框跳转、加日期、自定义热键、系统命令列表、PT 工具箱 |
| [设置文件参考](https://github.com/zhugecaomao/ALTRun/wiki/Configuration) | ALTRun.json 每一项的含义和默认值 |
| [常见问题](https://github.com/zhugecaomao/ALTRun/wiki/FAQ) | 热键冲突、搜不到、杀毒误报... |
| [开发指南](https://github.com/zhugecaomao/ALTRun/wiki/Development) | 架构、新增搜索功能、代码规范、测试 |


## 项目结构
```
ALTRun.ahk          入口: 列出所有模块并调用 App.Start()
Lib\                通用库, 与 ALTRun 无关 (JSON, Logger, Util, TextTools, Kanji, Dialogs, Everything IPC)
Src\Core\           启动流程, 设置与版本升级, 搜索模型, 匹配打分, 学习排序, 操作, 文件索引
Src\UI\             搜索窗口, 偏好设置窗口, 编辑对话框, 大字显示, 主题, 图标缓存
Src\Providers\      搜索功能: 应用 / 自定义命令 / 片段 / 剪贴板 / 系统命令 / 计算器 / 网页 / 文件 / 终端
Src\Extensions\     搜索窗口以外的功能: 片段自动展开, 对话框跳转, 加日期, PT 工具箱, 检查更新
Resources\          随程序发布的数据 (Kanji.txt 简繁对照表, Themes\ 内置主题)
Tests\              单元测试
```

运行后程序目录下还会出现 `ALTRun.json` (设置)、`Data\` (索引和历史, 可以删除)、`Themes\` (你自己的主题)。升级时把新版本复制覆盖到程序目录即可, 不要先删除 `Resources\` (里面可能有你自己放的 SPF2M 文件)。


## 开发
```
AutoHotkey64.exe /ErrorStdOut Tests\RunTests.ahk
```
单元测试不依赖界面, 退出码为失败的数量。新增搜索功能、代码规范和提交 PR 的流程见 [CONTRIBUTING.md](CONTRIBUTING.md) 和 Wiki 的 [开发指南](https://github.com/zhugecaomao/ALTRun/wiki/Development)。


## 贡献与反馈
- 发现问题: [提交 Issue](https://github.com/zhugecaomao/ALTRun/issues/new/choose) (请附上 Windows 版本和复现步骤)
- 想法和讨论: [Discussions](https://github.com/zhugecaomao/ALTRun/discussions)
- 提交代码: 请先阅读 [CONTRIBUTING.md](CONTRIBUTING.md)

如果 ALTRun 对你有帮助, 欢迎给它一个星标 ⭐


## English
ALTRun is a keyboard launcher for Windows modelled on Alfred for macOS. Press `Alt+Space`, type, press `Enter`.

- Finds apps (Start menu, desktop, Microsoft Store; pinyin initials for Chinese names), files and folders (through Everything when it is running, otherwise a built-in index), custom commands, snippets and system commands
- Learns which result you pick for each query, and ranks it first next time
- Action panel (`→` or right-click), in-place editing (`F3`), clipboard history, snippet auto-expansion, inline calculator, web search keywords, terminal commands, large type
- Nine built-in themes, plus custom themes as JSON files
- Portable: all settings live in `ALTRun.json` next to the program; older 2.x settings are converted automatically

The interface follows your Windows language (English or Chinese). Documentation is in the [Wiki](https://github.com/zhugecaomao/ALTRun/wiki) (Chinese). Issues and pull requests in English are welcome.


## 许可证与致谢
[GPL-3.0](LICENSE) © zhugecaomao

感谢 [ALTRun by etworker](https://github.com/etworker/ALTRun) (Delphi)、[RunZ by goreliu](https://github.com/goreliu/runz) (AutoHotkey), 以及 [Alfred](https://www.alfredapp.com/) 的设计。
