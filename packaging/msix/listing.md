# Microsoft Store listing / 商店介绍

The text for Partner Center (Apps and games → ALTRun → Store listings). Keep it in step with the README when features change;
the limits in brackets are Partner Center's. Screenshots: at least 1366 × 768 (see the last section).

商店介绍的文字 (Partner Center → 应用和游戏 → ALTRun → 商店一览)。功能变了要和 README 一起改; 方括号里是 Partner Center 的长度限制。

## Properties / 属性

| Field | Value |
|---|---|
| Category | Productivity |
| Privacy policy URL | https://github.com/zhugecaomao/ALTRun/blob/main/SECURITY.md#privacy-policy |
| Website | https://zhugecaomao.github.io/ALTRun/ |
| Support contact | https://github.com/zhugecaomao/ALTRun/issues |
| Price | Free, all markets |
| Copyright | Copyright (c) 2013-2026 zhugecaomao |
| Additional license terms | GPL-3.0: https://github.com/zhugecaomao/ALTRun/blob/main/LICENSE |
| Age rating (IARC) | No violence, no user interaction or shared content, no location, no purchases, no unrestricted web access (links open in the user's browser) |

## Restricted capabilities / 受限权限的理由 (Submission options)

> **runFullTrust**: ALTRun is a keyboard launcher for the Windows desktop. It starts the programs, files and folders the user searches for, runs system commands, registers global hotkeys (Alt+Space) with a low-level keyboard hook, reads the clipboard for its clipboard history, and moves Open / Save dialogs to folders open in file managers. This needs a full-trust desktop process.
>
> **unvirtualizedResources**: ALTRun keeps its settings in %APPDATA%\ALTRun, the same place as its other editions (portable zip, winget, Scoop), so users can move between them and back up or sync their settings. Its Delayed Start feature turns off other programs' start-up entries (HKCU\...\Explorer\StartupApproved) and starts them later through Task Scheduler; with registry write virtualization those changes would not reach Windows. Nothing is written outside the user's own profile and registry hive.

---

## English (en-us)

### Short description [1000]

Press Alt+Space, type a few letters, press Enter: apps, files, web search, calculations, clipboard history and system commands, all from one small, fast, keyboard-first window. Free and open source.

### Description [10000]

ALTRun is a lightweight, open-source launcher for Windows. Press Alt+Space, type a few letters and press Enter to open an app, a file or a website, calculate, convert units, paste from your clipboard history or run a system command, without reaching for the mouse.

Fast and small
• Results appear as you type. With Everything (voidtools) running, file searches cover all disks in milliseconds; otherwise ALTRun uses its own index.
• About 2.5 MB, no runtime to install, low memory use.

Smart search
• Learns from what you pick and moves frequent choices up.
• Matches word initials (vsc → Visual Studio Code) and pinyin initials (wx → 微信), and highlights what matched.
• Apps from the Start menu, the desktop and the Microsoft Store are indexed automatically.

More than a launcher
• Calculator with unit, number base, date and optional currency conversion.
• Clipboard history for text, files and images, with pinning; content from password managers is never recorded.
• Text snippets that expand anywhere when you type ;keyword.
• Web search with customizable engines, browser bookmarks, window switching, more than 40 Windows Settings pages and system commands.
• Quick Switch: jump Open / Save dialogs to the folder open in Explorer, Total Commander or Directory Opus.
• Actions on selected text, files or URLs in any program.
• Folder sync with a live progress window, delayed start of programs after sign-in, and scripts (.ahk, .ps1, .bat, .py).

Made for the keyboard
• → opens the action panel (run as administrator, open file location, copy path, terminal here, Windows context menu).
• F3 edits in place; Ctrl+1 to Ctrl+9 open a result directly; the status bar shows what Enter will do.

Yours to adjust
• English, Simplified Chinese, Traditional Chinese and Japanese.
• 22 themes in three densities, following Windows light and dark mode.
• Settings in plain JSON files.

Private
• No account, no telemetry, no ads. Your data stays on your computer.

ALTRun is free software under the GPL-3.0 license. Documentation, source code and issue tracker: https://github.com/zhugecaomao/ALTRun

### What's new in this version [1500]

First release in the Microsoft Store. The Store keeps ALTRun up to date; your settings stay in %APPDATA%\ALTRun.

### Product features [20 × 200]

1. Open apps, files, folders and websites from one window: Alt+Space, a few letters, Enter
2. Results as you type; learns from what you pick
3. Word-initial and pinyin-initial matching with highlighting
4. Instant file search across all disks with Everything, or the built-in index
5. Calculator with unit, number base, date and currency conversion
6. Clipboard history for text, files and images, with pinning
7. Text snippets that expand anywhere with ;keyword
8. Web search with customizable engines and browser bookmarks
9. Window switcher and more than 40 Windows Settings pages
10. Quick Switch: jump Open / Save dialogs to folders open in file managers
11. Actions on selected text, files or URLs in any program
12. Action panel: run as administrator, open file location, copy path, terminal here
13. Folder sync with live progress; delayed start of programs after sign-in
14. English, Simplified Chinese, Traditional Chinese and Japanese
15. 22 themes; follows Windows light and dark mode
16. No account, no telemetry, no ads; free and open source (GPL-3.0)

### Search terms [7 × 30, 21 words in all]

launcher, app launcher, file search, clipboard history, keyboard, productivity, quick search

---

## 简体中文 (zh-cn)

### 简短说明 [1000]

按 Alt+Space, 输入几个字母, 回车: 应用、文件、网页搜索、计算、剪贴板历史和系统命令, 都在一个小巧、快速、以键盘为主的窗口里完成。免费开源。

### 说明 [10000]

ALTRun 是一款轻量、开源的 Windows 启动器。按 Alt+Space, 输入几个字母再按回车, 就能打开应用、文件或网站, 计算、换算单位, 从剪贴板历史粘贴, 或者运行系统命令, 手不用离开键盘。

又快又小
• 边输入边出结果。运行 Everything (voidtools) 时, 几毫秒就能搜遍所有磁盘的文件; 没有它时使用 ALTRun 自己的索引。
• 约 2.5 MB, 不用安装任何运行库, 占用内存少。

聪明的搜索
• 记住你选了什么, 常用的排在前面。
• 支持单词首字母 (vsc → Visual Studio Code) 和拼音首字母 (wx → 微信), 并高亮匹配的部分。
• 自动收录开始菜单、桌面和 Microsoft Store 里的应用。

不只是启动器
• 计算器, 支持单位、进制、日期换算和可选的汇率换算。
• 剪贴板历史, 记录文字、文件和图片, 可以置顶; 密码管理器复制的内容不会记录。
• 文字片段, 在任何程序里输入 ;关键字 自动展开。
• 可自定义的网页搜索、浏览器书签、窗口切换、40 多个 Windows 设置页和系统命令。
• Quick Switch: 在打开 / 保存对话框里一键跳到资源管理器、Total Commander 或 Directory Opus 当前的文件夹。
• 对任何程序里选中的文字、文件或网址执行操作。
• 带实时进度窗口的文件夹同步、登录后延迟启动程序, 以及脚本 (.ahk、.ps1、.bat、.py)。

为键盘而设计
• → 打开操作面板 (以管理员身份运行、打开所在位置、复制路径、在此处打开终端、Windows 右键菜单)。
• F3 直接编辑; Ctrl+1 到 Ctrl+9 直接打开结果; 状态栏显示 Enter 会做什么。

随你调整
• 英文、简体中文、繁体中文和日文界面。
• 22 款主题, 三种密度, 跟随 Windows 浅色 / 深色模式。
• 设置是普通的 JSON 文件。

注重隐私
• 不需要账号, 没有遥测, 没有广告。数据都留在你的电脑上。

ALTRun 是 GPL-3.0 许可的自由软件。文档、源代码和问题反馈: https://github.com/zhugecaomao/ALTRun

### 此版本的新增内容 [1500]

首次在 Microsoft Store 发布。由 Store 负责更新; 设置仍保存在 %APPDATA%\ALTRun。

### 产品功能 [20 × 200]

1. 在一个窗口里打开应用、文件、文件夹和网站: Alt+Space, 几个字母, 回车
2. 边输入边出结果; 记住你的选择
3. 单词首字母和拼音首字母匹配, 并高亮
4. 配合 Everything 瞬间搜遍所有磁盘, 或使用内置索引
5. 计算器: 单位、进制、日期和汇率换算
6. 剪贴板历史: 文字、文件和图片, 可以置顶
7. 文字片段: 在任何程序里输入 ;关键字 展开
8. 可自定义的网页搜索和浏览器书签
9. 窗口切换和 40 多个 Windows 设置页
10. Quick Switch: 打开 / 保存对话框一键跳到文件管理器当前的文件夹
11. 对任何程序里选中的文字、文件或网址执行操作
12. 操作面板: 以管理员身份运行、打开所在位置、复制路径、在此处打开终端
13. 带实时进度的文件夹同步; 登录后延迟启动程序
14. 英文、简体中文、繁体中文和日文界面
15. 22 款主题, 跟随 Windows 浅色和深色模式
16. 不需要账号、没有遥测和广告; 免费开源 (GPL-3.0)

### 搜索词 [7 × 30, 共 21 个词]

启动器, 快速启动, 文件搜索, 剪贴板历史, 效率工具, 拼音搜索, launcher

---

## Screenshots / 截图

Partner Center wants desktop screenshots of at least 1366 × 768 (PNG, up to 10, the first ones show first). The README's
screenshots are window-sized (about 700–1000 px wide), so the Store set is made separately: the windows on a 1920 × 1080
background, English interface, made-up data (as all of ALTRun's screenshots). Planned order:

1. Search window with results (`search.png` / `demo.png`)
2. Action panel (`actions.png`)
3. Calculator and unit conversion (`demo-calculator.png`)
4. Clipboard history (`demo-clipboard.png`)
5. File search with Everything (`files.png`)
6. Quick Switch (`demo-quickswitch.png`)
7. Themes (`theme-dark.png` and others)
8. Preferences (`prefs-general.png`)

商店截图至少 1366 × 768。README 里的截图是窗口大小, 所以另外做一套: 窗口放在 1920 × 1080 的背景上, 英文界面、虚构数据。
