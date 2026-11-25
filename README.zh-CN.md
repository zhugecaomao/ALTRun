<h1 align="center"><img width="48" alt="ALTRun" src="docs/images/logo.png" /> ALTRun</h1>

<p align="center"><a href="README.md">English</a> | <b>简体中文</b></p>

<p align="center">
  <b>轻量、开源的 Windows 启动器, 操作方式参照 macOS 上的 <a href="https://www.alfredapp.com/">Alfred</a></b><br>
  按 <code>Alt+Space</code>, 输入几个字母, 按 <code>Enter</code>: 程序、文件、网页、计算、剪贴板、系统命令, 一个窗口完成
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
  <a href="#亮点">亮点</a> ·
  <a href="#快速开始">快速开始</a> ·
  <a href="https://github.com/zhugecaomao/ALTRun/wiki">使用文档 (Wiki)</a> ·
  <a href="CHANGELOG.md">更新日志</a>
</p>

<p align="center">
  <img src="docs/images/screenshots/demo.png" width="700" alt="ALTRun: 输入即搜、操作面板、单位换算、文件搜索">
</p>


## 为什么选择 ALTRun
- **小巧**: 约 2.5 MB, 一个 `ALTRun.exe` 加少量资源文件, 不用安装 .NET、Electron 等运行库。
- **快**: 输入即出结果; 开着 [Everything](https://www.voidtools.com/) 时毫秒级搜遍全盘文件。
- **越用越懂你**: 记住每次输入后选了哪一项, 常用的排在前面; 支持单词首字母 (`vsc` → Visual Studio Code) 和拼音首字母 (`wx` → 微信)。
- **键盘优先**: `→` 打开操作面板, `F3` 直接编辑, `Ctrl+1`–`Ctrl+9` 直接打开结果, `?` 查看全部快捷键。
- **便携、注重隐私**: 免安装、不写注册表、不需要管理员权限、没有后台服务; 数据留在你的电脑上, 不收集遥测数据。
- **随时更新**: 新版本在搜索窗口里提示, 按 `Enter` 安装, 校验 SHA256, 保留设置; 也支持 winget 和 Scoop。
- **免费开源**: GPL-3.0 许可, 每次改动都在 Windows 上自动跑几千项测试。


## 亮点
| **一个框搜索一切**: 程序、文件、自定义命令、窗口、系统命令, 输入即出 ([详情](https://github.com/zhugecaomao/ALTRun/wiki/Usage-zh)) | **操作面板**: 按 `→` 用管理员身份运行、打开所在文件夹、复制路径、在这里打开终端... ([详情](https://github.com/zhugecaomao/ALTRun/wiki/Usage-zh#操作面板)) |
|:---:|:---:|
| <img src="docs/images/screenshots/search.png" alt="搜索窗口"> | <img src="docs/images/screenshots/actions.png" alt="操作面板"> |
| **对话框快速跳转 (Quick Switch)**: 打开 / 保存对话框一键跳到资源管理器、Total Commander、Directory Opus 或 Arke 里打开的文件夹 ([详情](https://github.com/zhugecaomao/ALTRun/wiki/Quick-Switch-zh)) | **对任何内容运行自己的命令**: 在任何程序里选中文字、文件或网址, 按 `Ctrl+Alt+\`, 为这类内容设的命令排在最前 ([详情](https://github.com/zhugecaomao/ALTRun/wiki/Content-Types-zh)) |
| <img src="docs/images/screenshots/demo-quickswitch.png" alt="对话框快速跳转: Ctrl+G 和文件夹面板把对话框跳到打开的文件夹"> | <img src="docs/images/screenshots/actions-file-type.png" alt="Markdown 文件: 自己的命令排在最前"> |
| **剪贴板历史**: `Ctrl+Alt+C` 找回复制过的文字、文件和图片, 筛选后粘贴 ([详情](https://github.com/zhugecaomao/ALTRun/wiki/Clipboard-zh)) | **文字片段**: 在任何程序里输入 `;thx` 展开回复、签名或地址 ([详情](https://github.com/zhugecaomao/ALTRun/wiki/Snippets-zh)) |
| <img src="docs/images/screenshots/demo-clipboard.png" alt="剪贴板历史: Ctrl+Alt+C, 筛选, Enter 粘贴"> | <img src="docs/images/screenshots/demo-snippets.png" alt="文字片段: ;thx 展开成一封回复"> |
| **计算器**: `12*(3+4)`、`10 km in mi`、`255 in hex`、`today + 30 days`、`100 usd to cny`, 输入即出结果 ([详情](https://github.com/zhugecaomao/ALTRun/wiki/Calculator-zh)) | **文件搜索**: `空格` + 名称; 有 Everything 时搜遍全盘, 否则用内置索引 ([详情](https://github.com/zhugecaomao/ALTRun/wiki/File-Search-zh)) |
| <img src="docs/images/screenshots/demo-calculator.png" alt="计算器: 输入即出结果, 也能换算单位"> | <img src="docs/images/screenshots/files.png" alt="文件搜索"> |
| **浏览文件夹**: 输入路径, `Tab` 进入, `Backspace` 返回上一级 ([详情](https://github.com/zhugecaomao/ALTRun/wiki/File-Search-zh#浏览文件夹)) | **拼音首字母**: `jsb` 找到记事本; 文件名也可以 ([详情](https://github.com/zhugecaomao/ALTRun/wiki/Usage-zh#怎样匹配)) |
| <img src="docs/images/screenshots/demo-browse.png" alt="浏览文件夹: 输入路径, Tab, 筛选, Backspace"> | <img src="docs/images/screenshots/pinyin.png" alt="拼音首字母"> |
| **网页搜索**: `g 关键词`、`bd 关键词` 或自己加的搜索引擎; 一个关键字可以同时打开几个网站 ([详情](https://github.com/zhugecaomao/ALTRun/wiki/Web-Search-zh)) | **窗口切换**: `w` + 标题或程序名, 切换到已经打开的窗口 ([详情](https://github.com/zhugecaomao/ALTRun/wiki/Window-Switching-zh)) |
| <img src="docs/images/screenshots/websearch.png" alt="网页搜索"> | <img src="docs/images/screenshots/windows.png" alt="窗口切换"> |
| **脚本实时结果**: `.ahk`、`.ps1` 或 `.py` 脚本可以在你打字时列出结果 ([详情](https://github.com/zhugecaomao/ALTRun/wiki/Scripts-zh#实时结果)) | **22 套主题**, 每套三种密度, 可以跟随 Windows 浅色 / 深色模式 ([详情](https://github.com/zhugecaomao/ALTRun/wiki/Themes-zh)) |
| <img src="docs/images/screenshots/script-filter.png" alt="脚本的实时结果"> | <img src="docs/images/screenshots/prefs-appearance.png" alt="主题"> |

截图由 [Tests/Screenshots](Tests/Screenshots/TakeScreenshots.ahk) 在 GitHub Actions 的 Windows 机器上自动生成。


## 更多功能
- **搜索**: 浏览器书签和历史记录、最近的文件、置顶和最近使用、40 多个 Windows 设置页面和系统命令、终端 (`>ipconfig`)、大字显示
- **按自己的需要**: 带参数的自定义命令、内容类型、自定义热键、在资源管理器里 "发送到 ALTRun"、[脚本库](https://zhugecaomao.github.io/ALTRun/#scripts)
- **工具**: 重命名时一键加日期 (`Ctrl+D`)、用 robocopy 同步文件夹、延迟启动、CapsLock 切换输入法、给结构工程师的 PT 工具箱
- **界面**: 简体中文、繁體中文、English、日本語

全部功能和截图: [功能一览](https://github.com/zhugecaomao/ALTRun/wiki/Features-zh)。


## 快速开始
1. [下载最新版本](https://github.com/zhugecaomao/ALTRun/releases/latest), 解压到任意文件夹, 运行 `ALTRun.exe` (不用安装, 也不需要 AutoHotkey)。也可以用包管理器安装:
   ```powershell
   winget install zhugecaomao.ALTRun
   # 或者
   scoop bucket add altrun https://github.com/zhugecaomao/ALTRun
   scoop install altrun
   ```
2. 按 `Alt+Space` (或 `Alt+R`), 输入几个字母, 按 `Enter`。
3. 输入 `?` 查看全部搜索语法和快捷键; 按 `Ctrl+,` 打开偏好设置。

如果 Windows SmartScreen 拦截, 点 **更多信息** → **仍要运行**; 杀毒软件误报见 [常见问题](https://github.com/zhugecaomao/ALTRun/wiki/FAQ-zh#杀毒软件报毒)。升级、用 AutoHotkey v2 运行源码、卸载: 见 [安装与升级](https://github.com/zhugecaomao/ALTRun/wiki/Installation-zh)。

| 要做的事 | 输入 |
|---|---|
| 打开程序 | `word` → `Enter` |
| 找文件或文件夹 | `空格`, 然后 `report` |
| 计算 | `12*(3+4)`; `Enter` 复制结果 |
| 搜索网页 | `g 关键词` |
| 从剪贴板历史粘贴 | `Ctrl+Alt+C` |
| 在终端里运行命令 | `>ipconfig /all` |
| 看对结果能做什么 | `→` 或右键 |
| 编辑选中的命令 | `F3` |

全部快捷键: [搜索与快捷键](https://github.com/zhugecaomao/ALTRun/wiki/Usage-zh#快捷键)。


## 文档
[Wiki](https://github.com/zhugecaomao/ALTRun/wiki/Home-zh) 详细介绍每个功能, 附示例。可以从这几页开始:
[功能一览](https://github.com/zhugecaomao/ALTRun/wiki/Features-zh) ·
[5 分钟上手](https://github.com/zhugecaomao/ALTRun/wiki/Getting-Started-zh) ·
[搜索与快捷键](https://github.com/zhugecaomao/ALTRun/wiki/Usage-zh) ·
[自定义命令](https://github.com/zhugecaomao/ALTRun/wiki/Custom-Commands-zh) ·
[文字片段](https://github.com/zhugecaomao/ALTRun/wiki/Snippets-zh) ·
[对话框快速跳转](https://github.com/zhugecaomao/ALTRun/wiki/Quick-Switch-zh) ·
[脚本扩展](https://github.com/zhugecaomao/ALTRun/wiki/Scripts-zh) ·
[设置文件参考](https://github.com/zhugecaomao/ALTRun/wiki/Configuration-zh) ·
[常见问题](https://github.com/zhugecaomao/ALTRun/wiki/FAQ-zh)


## 反馈与参与
- [报告问题](https://github.com/zhugecaomao/ALTRun/issues/new?template=bug_report.yml) (请写上 Windows 版本和重现步骤)、[功能建议](https://github.com/zhugecaomao/ALTRun/issues/new?template=feature_request.yml) 或 [发起讨论](https://github.com/zhugecaomao/ALTRun/issues/new?template=discussion.yml), 中文、英文都可以
- 代码: 先读 [CONTRIBUTING.md](CONTRIBUTING.md#参与贡献); [开发指南](https://github.com/zhugecaomao/ALTRun/wiki/Development-zh) 介绍项目结构、测试和如何添加搜索功能
- 脚本: 通过 Pull Request 投稿到 [脚本库](docs/scripts/README.md#contributing-a-script--投稿脚本)

如果 ALTRun 对你有帮助, 欢迎点个 Star ⭐


## Code signing policy
Free code signing provided by [SignPath.io](https://about.signpath.io/), certificate by [SignPath Foundation](https://signpath.org/).

- Only `ALTRun.exe` built from this repository's source by the [Release workflow](.github/workflows/release.yml) is signed.
- Committers and reviewers: [zhugecaomao](https://github.com/zhugecaomao)
- Approvers: [zhugecaomao](https://github.com/zhugecaomao)
- Privacy policy: [SECURITY.md](SECURITY.md#privacy-policy) (中文: [隐私说明](SECURITY.md#隐私说明))


## 许可证与致谢
ALTRun 以 [GPL-3.0](LICENSE) 许可开源 © zhugecaomao

感谢这些项目:
- [ALTRun](https://github.com/etworker/ALTRun) (etworker, Delphi): 名字和最初的设计
- [RunZ](https://github.com/goreliu/runz) (goreliu, AutoHotkey): 用 AutoHotkey 写启动器的思路
- [Alfred](https://www.alfredapp.com/): 操作方式
- [Listary](https://www.listary.com/): 对话框快速跳转参考了它的 Quick Switch
- [pypinyin](https://github.com/mozillazg/python-pinyin) (Copyright (c) 2016 mozillazg, 闲耘; [MIT](https://github.com/mozillazg/python-pinyin/blob/master/LICENSE.txt)): 生僻字和繁体字的拼音首字母
