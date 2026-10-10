# Security Policy

[English](#security-policy) · [中文](#安全策略)

## Supported versions

| Version | Supported |
|---|---|
| [Latest release](https://github.com/zhugecaomao/ALTRun/releases/latest) | ✅ |
| Older releases | ❌ Please update to the latest release |

Security fixes are released as a new version. ALTRun checks for new versions in the background and installs them with one key press.

## Reporting a vulnerability

Please **do not open a public issue** for security problems. Report them privately through GitHub:
**[Report a vulnerability](https://github.com/zhugecaomao/ALTRun/security/advisories/new)**

Please include, if possible:
- the ALTRun version and Windows version
- steps to reproduce, or a proof of concept
- the impact you expect (for example: another program can make ALTRun run commands, clipboard history can be read by other users)

What to expect: ALTRun is a free project maintained by one person in their spare time, so there are no guaranteed response times. Security reports are handled before other work:
- you will get a reply as soon as possible, usually within a few weeks
- confirmed issues are fixed in the next release, with priority given to serious ones
- **Coordinated disclosure**: please keep the details private until a fixed version is available; a security advisory is then published, and reporters are credited in the advisory and the release notes unless they prefer to stay anonymous

### Scope

In scope:
- the source code in this repository and the release packages on [GitHub Releases](https://github.com/zhugecaomao/ALTRun/releases)
- the update mechanism (update check, download, SHA256 verification, file replacement)
- the Scoop and winget manifests and the project website in this repository

Out of scope:
- attacks that need administrator rights or an already compromised account on the same computer
- third-party software that ALTRun works with, such as Everything, Total Commander or AutoHotkey itself (please report to their authors)
- custom commands, snippets and scripts that a user adds and that behave as written

## Security measures

- **Builds**: every release is built from the public source code by [GitHub Actions](.github/workflows/release.yml); no binaries are built on personal computers. Each change runs 6,000+ automated tests on Windows, and the repository is scanned by CodeQL.
- **Updates**: one-click updates download only from GitHub Releases over HTTPS, verify the SHA256 checksum published by GitHub before replacing any file, and restore the previous version if any step fails.
- **Code signing**: see the [code signing policy](README.md#code-signing-policy).
- **Least privilege**: ALTRun runs as a normal user. It needs no administrator rights and installs no services or drivers.
- **Clipboard**: content marked by password managers and other programs as private (`ExcludeClipboardContentFromMonitorProcessing`, `CanIncludeInClipboardHistory = 0`) is never recorded; content copied from KeePass, KeePassXC, 1Password and Bitwarden is ignored by default.

## Privacy policy

ALTRun runs entirely on your computer. **It collects no personal data and has no telemetry, analytics, crash reporting, advertising or user accounts.** This program will not transfer any information to other networked systems unless specifically requested by the user or the person installing or operating it.

### Network connections

ALTRun itself connects to the internet only in these cases:

| Purpose | Destination | When | How to turn it off |
|---|---|---|---|
| Check for updates | api.github.com | At startup (if the last check was more than an hour ago), every 6 hours while running, or when you check manually | Preferences → General → "Check for updates automatically" |
| Download an update | github.com (GitHub Releases) | Only after you choose to install an update | Do not install updates from ALTRun |
| Exchange rates | api.frankfurter.dev | Once a day, only when currency conversion is turned on (off by default) | Preferences → Calculator → "Currency conversion" |
| Web search icon | The website in that web search's URL | Only when you click "Download Site Icon" while editing a web search | Do not click it |

These are plain HTTPS requests that send no personal data; like any web request they reveal your IP address to the server. Web searches, release notes and links that you open from ALTRun are opened in your web browser and are subject to the privacy policy of those websites. The Everything integration uses local inter-process communication, not the network.

### Data stored on your computer

All data stays on your computer, in the `Data` folder next to `ALTRun.exe` (or `%APPDATA%\ALTRun\Data` when the program folder is not writable, or a folder you choose in Preferences → Advanced); clipboard images and very long entries are kept in `%LOCALAPPDATA%\ALTRun\Clipboard`. If you choose a folder that is synchronized by a cloud service such as OneDrive, that service's privacy policy applies to the synchronized files.

| Data | File | Notes |
|---|---|---|
| Settings, custom commands, snippets, pinned items | `ALTRun.json` | |
| Learned ranking, search history, recent items | `Knowledge.json` | Reset in Preferences → Advanced |
| Usage statistics | `Usage.json` | Counts per feature only; no typed text or opened items |
| Clipboard history | `ClipboardHistory.json`; images and very long entries in `%LOCALAPPDATA%\ALTRun\Clipboard` | Images and long entries stay on this computer even when the `Data` folder is synced; can be kept in memory only; can be cleared at any time |
| Search indexes | `AppIndex.json`, `FileIndex.json` | Names and paths of applications and files, used for searching |
| Exchange rates, update status | `Currency.json`, `Update.json` | |
| Debug log (off by default) | `%TEMP%\ALTRun.log` | Timings and errors; no typed text |

To remove all data, quit ALTRun and delete the `Data` folder. No copies are kept anywhere else.

### Changes to this policy

Changes to this policy are recorded in the [changelog](CHANGELOG.md) and the [history of this file](https://github.com/zhugecaomao/ALTRun/commits/main/SECURITY.md). Questions can be asked in [GitHub Issues](https://github.com/zhugecaomao/ALTRun/issues).

---

# 安全策略

## 支持的版本

| 版本 | 是否支持 |
|---|---|
| [最新版本](https://github.com/zhugecaomao/ALTRun/releases/latest) | ✅ |
| 以前的版本 | ❌ 请更新到最新版本 |

安全修复以新版本发布。ALTRun 会在后台检查新版本, 按一下 `Enter` 即可安装。

## 报告安全问题

发现安全问题时 **请不要公开提交 Issue**, 而是通过 GitHub 私下报告:
**[私下报告漏洞](https://github.com/zhugecaomao/ALTRun/security/advisories/new)**

请尽量附上:
- ALTRun 版本和 Windows 版本
- 复现步骤或概念验证
- 可能的影响 (例如: 其它程序能让 ALTRun 执行命令、剪贴板历史能被其他用户读取)

处理方式: ALTRun 是一个人利用业余时间维护的免费项目, 不承诺固定的处理时限, 但安全问题会优先处理:
- 会尽快回复, 一般在几周之内
- 确认的问题在下一个版本里修复, 严重的优先
- **协调披露**: 请在修复版本发布之前不要公开细节; 修复后发布安全公告, 在公告和 Release 说明里致谢报告者 (不希望署名的除外)

### 范围

包括:
- 本仓库的源码和 [GitHub Releases](https://github.com/zhugecaomao/ALTRun/releases) 上的发布包
- 更新机制 (检查更新、下载、SHA256 校验、替换文件)
- 本仓库里的 Scoop / winget 清单和官网

不包括:
- 需要管理员权限或同一台电脑上已被攻破的账户才能进行的攻击
- ALTRun 配合使用的第三方软件, 例如 Everything、Total Commander、AutoHotkey 本身 (请报告给它们的作者)
- 用户自己添加的自定义命令、片段和脚本按其内容执行

## 安全措施

- **构建**: 每个发布版本都由 [GitHub Actions](.github/workflows/release.yml) 从公开源码编译, 不在个人电脑上构建。每次修改都在 Windows 上运行 6000 余项自动测试, 仓库开启了 CodeQL 扫描。
- **更新**: 一键更新只从 GitHub Releases 通过 HTTPS 下载, 替换文件前先核对 GitHub 给出的 SHA256, 任何一步出错都还原到原来的版本。
- **代码签名**: 见 [代码签名策略](README.md#code-signing-policy)。
- **最小权限**: ALTRun 以普通用户身份运行, 不需要管理员权限, 不安装服务或驱动。
- **剪贴板**: 带有 "不要记录" 标记的内容 (`ExcludeClipboardContentFromMonitorProcessing`、`CanIncludeInClipboardHistory = 0`, 密码管理器常用) 从不记录; 默认忽略 KeePass、KeePassXC、1Password、Bitwarden 复制的内容。

## 隐私说明

ALTRun 完全在本机运行, **不收集任何个人数据, 没有遥测、统计、崩溃报告、广告或用户账号。** 除非用户 (或安装、操作本程序的人) 明确要求, 本程序不会向其它网络系统传输任何信息。

### 联网

ALTRun 本身只在以下情况联网:

| 用途 | 地址 | 什么时候 | 怎样关闭 |
|---|---|---|---|
| 检查更新 | api.github.com | 启动时 (距上次检查满 1 小时)、运行期间每 6 小时, 或手动检查时 | 偏好设置 → 通用 → "自动检查更新" |
| 下载更新 | github.com (GitHub Releases) | 只在你选择安装更新之后 | 不在 ALTRun 里安装更新 |
| 汇率 | api.frankfurter.dev | 打开货币换算后每天一次 (默认关闭) | 偏好设置 → 计算器 → "货币换算" |
| 网页搜索的图标 | 这个网页搜索的网址所在的网站 | 只在编辑网页搜索时点了 "下载网站图标" | 不点这个按钮 |

这些都是普通的 HTTPS 请求, 不发送任何个人数据; 和所有网络请求一样, 服务器能看到你的 IP 地址。从 ALTRun 打开的网页搜索、更新说明和链接都在你的浏览器里打开, 适用对应网站的隐私政策。Everything 联动通过本机的进程间通信, 不经过网络。

### 保存在本机的数据

所有数据都只保存在你的电脑上, 位于 `ALTRun.exe` 旁边的 `Data` 文件夹 (程序目录不能写入时在 `%APPDATA%\ALTRun\Data`, 或者你在 偏好设置 → 高级 里选的文件夹); 剪贴板的图片和很长的条目在 `%LOCALAPPDATA%\ALTRun\Clipboard`。如果选了 OneDrive 等云同步的文件夹, 同步的文件适用该服务的隐私政策。

| 数据 | 文件 | 说明 |
|---|---|---|
| 设置、自定义命令、片段、置顶的项目 | `ALTRun.json` | |
| 学习排序、搜索历史、最近使用 | `Knowledge.json` | 可在 偏好设置 → 高级 里重置 |
| 使用统计 | `Usage.json` | 只记录每个功能用了几次, 不记录输入的文字和打开的内容 |
| 剪贴板历史 | `ClipboardHistory.json`; 图片和很长的条目在 `%LOCALAPPDATA%\ALTRun\Clipboard` | 图片和长条目只存在这台电脑上, Data 文件夹放在同步盘里也不会同步; 可以设为只保存在内存里; 随时可以清空 |
| 搜索索引 | `AppIndex.json`、`FileIndex.json` | 应用和文件的名称、路径, 用于搜索 |
| 汇率、更新状态 | `Currency.json`、`Update.json` | |
| 调试日志 (默认关闭) | `%TEMP%\ALTRun.log` | 耗时和错误信息, 不记录输入的文字 |

要删除全部数据, 退出 ALTRun 后删除 `Data` 文件夹即可, 其它地方没有副本。

### 政策变更

本政策的修改记录在 [更新日志](CHANGELOG.md) 和 [本文件的历史](https://github.com/zhugecaomao/ALTRun/commits/main/SECURITY.md) 里。有问题可以在 [GitHub Issues](https://github.com/zhugecaomao/ALTRun/issues) 里提出。
