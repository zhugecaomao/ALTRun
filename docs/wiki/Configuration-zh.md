[English](Configuration) · **中文**

# 设置文件参考 (ALTRun.json)

所有设置和用户数据都保存在程序目录下的 `Data\ALTRun.json` (UTF-8 文本; 2026.09.26 及更早的版本放在程序目录里, 升级后第一次启动时自动移到 `Data\`)。程序目录不能写入时 (例如放在 `C:\Program Files`) 改用 `%APPDATA%\ALTRun\Data\`; 在 偏好设置 → 高级 → 数据 → 更改位置 可以换到别的文件夹 (例如同步盘里的, 见 [多台电脑共用设置](FAQ-zh#在多台电脑上使用同一份设置))。大部分设置可以在偏好设置窗口里修改 (每个设置下面有一行灰色的说明): **确定** 保存并关闭, **应用** 保存并留在当前页, **取消** 放弃修改, **帮助** (`F1`) 打开当前页的说明; 设置在 ALTRun 重新载入后生效, 由程序自动完成。也可以直接编辑这个文件 (搜索窗口里按 `F4`, 或 偏好设置 → 高级 → 编辑 ALTRun.json), **保存后 ALTRun 自动重新载入**。

- 文件里缺少的项目会自动用默认值补上, 所以只需要写要修改的部分
- 文件格式有错误 (例如少了逗号) 时, ALTRun 会提示, 把原文件改名为 `ALTRun.json.bad` 保留下来, 然后用默认设置启动; 修好后改回原名即可
- `SchemaVersion` 是设置格式的版本号, 不要手动修改; 格式升级时 ALTRun 会自动转换并备份

## 偏好设置页面与文档

偏好设置底部的 **帮助** 按钮和 `F1` 会打开当前页面的说明。下面列出每个页面对应的 Wiki 入口, 方便查找具体设置:

| 偏好设置页面 | 文档 |
|---|---|
| 通用 / 高级 | [设置文件参考](Configuration-zh) |
| 搜索窗口 / 功能 / 使用统计 | [搜索与快捷键](Usage-zh); 功能页里的各项: [浏览器书签](Bookmarks-zh)、[窗口切换](Window-Switching-zh)、[置顶和最近使用](Recent-zh)、[系统命令和终端](System-Commands-zh) |
| 网页搜索 | [网页搜索](Web-Search-zh) |
| 外观 | [主题](Themes-zh) |
| 应用 | [设置文件参考: 应用索引](Configuration-zh#applications-应用) |
| 文件搜索 / 搜索来源 | [文件搜索](File-Search-zh) |
| 自定义命令 | [自定义命令](Custom-Commands-zh) |
| 文字片段 | [文字片段](Snippets-zh) |
| 剪贴板 | [剪贴板历史](Clipboard-zh) |
| 计算器 | [计算器](Calculator-zh) |
| 脚本扩展 | [脚本扩展](Scripts-zh) |
| 键盘 / 自定义热键 | [自定义热键](Hotkeys-zh) (CapsLock、切换到英文; 空格和 Tab 键另见 [快捷键](Usage-zh#快捷键)) |
| 对话框跳转 | [对话框快速跳转](Quick-Switch-zh) |
| 对话框面板 | [对话框文件夹面板](Quick-Switch-Panel-zh) |
| 一键加日期 | [一键加日期](Date-Stamp-zh) |

## 结构
```jsonc
{
  "SchemaVersion": 5,
  "General":    { ... },          // 通用
  "Appearance": { ... },          // 外观
  "Features":   { ... },          // 各个搜索功能
  "Extensions": { ... },          // 扩展功能
  "Hotkeys":        [ ... ],      // 自定义热键
  "CustomCommands": [ ... ],      // 自定义命令
  "Snippets":       [ ... ],      // 文字片段
  "DataLocation":   "%OneDrive%\\ALTRun"   // 可选: 改用这个数据文件夹
}
```

`DataLocation` 只在默认位置 (`Data\ALTRun.json`, 或 `%APPDATA%\ALTRun\Data\ALTRun.json`) 的文件里起作用: 有这一项时, ALTRun 改用那个文件夹里的设置和数据, 默认位置的文件只用来指路。一般在 偏好设置 → 高级 → 数据 → 更改位置 里设置, 不用手动写; 可以用环境变量, 相对路径从程序目录算起。

## General 通用
偏好设置里分在 "通用"、"搜索窗口" 和 "键盘" 三页。

| 键 | 默认 | 说明 |
|---|---|---|
| Hotkey | `!Space` | 呼出热键 (AutoHotkey 写法, `!` Alt `^` Ctrl `+` Shift `#` Win; 偏好设置里直接按键录制, 见 [设置热键](Hotkeys-zh#设置热键)) |
| SecondaryHotkey | `!r` | 第二个呼出热键 (Alt+R), 留空 = 不用 |
| DoubleTap | 空 | `Ctrl` / `Shift` = 快速按两下这个键也能呼出 (和 Listary 一样), 空 = 不用 |
| CapsLock | 空 | 和 Mac 一样: `Layout` = 按一下 CapsLock 切换到下一个输入法 (和 `Win+Space` 一样), `Mode` = 切换输入法的中 / 英模式; 按住 0.3 秒才开 / 关大写锁定。空 = CapsLock 照常。见 [CapsLock 切换输入法](Hotkeys-zh#capslock-切换输入法) |
| CapsLockNotify | 1 | 开 / 关大写锁定时显示 "大写锁定: 开 / 关" 的提示; 0 = 不显示 |
| CapsLockElevated | 0 | 1 = CapsLock 切换在管理员权限的窗口里也生效: 用计划任务以管理员身份运行一个只管 CapsLock 的小助手, ALTRun 本身还是普通权限。见 [CapsLock 切换输入法](Hotkeys-zh#capslock-切换输入法) |
| SelectionHotkey | `^!\` | 选中文字 / 文件 / 网址后按下, 直接打开它们的操作 (Ctrl+Alt+\), 见 [选中内容的操作](Usage-zh#选中内容的操作); 留空 = 不用 |
| Language | `auto` | `auto` 跟随 Windows 的显示语言, 或语言代码: `en` / `zh-CN` (简体中文) / `zh-TW` (繁體中文) / `ja`; 除了英文, 每种语言是 `Resources\Lang\` 里的一个文件 (以前的 `zh` 等于 `zh-CN`) |
| LaunchAtLogin | 1 | 开机自动启动 |
| ShowTrayIcon | 1 | 显示托盘图标 |
| HideOnDeactivate | 1 | 搜索窗口失去焦点时隐藏 |
| SwitchToEnglishInput | 0 | 显示搜索窗口时切换到英文: 0 = 不切换, 1 = 英语 (美国) 键盘 (只在已经装好的输入法之间切换, 不会添加新的键盘布局), 2 = 当前输入法的英文模式 (输入法不变, 关掉它的中文) |
| RestoreInput | 1 | 打开 `SwitchToEnglishInput` 时, 窗口隐藏后切回原来的输入法或中 / 英模式 (1 时: Windows 设置了每个应用窗口各用各的输入法时不需要, 也不会切换) |
| SpaceToRun | 0 | 已输入文字时按空格执行选中项, `Shift+空格` 输入空格 |
| TabKey | `"Complete"` | 搜索窗口里 `Tab` 的作用: `"Complete"` 自动补全 (文件夹: 进入浏览); `"Next"` 选下一个结果 (`Shift+Tab` 上一个, 到头后循环), `Ctrl+Tab` 自动补全 |
| KeepLastQuery | 0 | 呼出窗口时保留上一次的搜索 (输入、文件搜索模式和选中的行), 文字全选: `Enter` 再执行一次, 直接输入开始新的搜索。从 2.x 升级时沿用 `KeepInput` |
| ShowTips | 1 | 轮换显示使用提示, 用灰色文字显示在空搜索框里 |
| FileManager | `explorer.exe` | 打开文件夹用的程序, 后面可以带参数, 例如 `C:\Apps\TotalCMD64\TOTALCMD64.exe /O /T /S` (Total Commander: `/O` 用已打开的窗口, `/T` 新标签页, `/S` 当前面板); 路径有空格时加引号 |
| SendToMenu | 1 | 添加到资源管理器的 "发送到" 菜单 |
| StartMenuShortcut | 1 | 添加到开始菜单 |
| CheckForUpdates | 1 | 启动时和之后每 6 小时在后台检查一次 GitHub 上的新版本, 有新版本时显示在搜索窗口里, 见 [一键更新](Installation-zh#一键更新) |
| SaveLog | 0 | 写入调试日志 (`%Temp%\ALTRun.log`), 包括启动各阶段和超过 30 ms 的搜索用了多少毫秒 (`Perf:` 开头, 不记输入的文字) |
| HistorySize | 30 | 记住多少条最近的搜索 |

## Appearance 外观
| 键 | 默认 | 说明 |
|---|---|---|
| Theme | `System` | 主题名, 见 [主题](Themes-zh); `System` 跟随 Windows 的深 / 浅色 |
| Density | `Normal` | 显示密度, 对任何主题都有效: `Comfortable` 宽松 (主题本来的大小) / `Normal` 常规 / `Compact` 紧凑, 见 [主题](Themes-zh#显示密度)。从旧版本升级时是 `Comfortable` (保持原来的大小) |
| Width | 700 | 搜索窗口宽度 (像素, 按屏幕缩放; 最小 450) |
| VisibleRows | 8 | 最多显示几行结果 (1~9) |
| Placeholder | "" | 空搜索框里的灰色文字; 空 = 默认的 "ALTRun 搜索", 只写空格 = 什么都不显示 (有使用提示时不显示) |
| StatusBar | 1 | 搜索窗口底部的状态栏 (有结果时显示: ☰ 菜单, 以及 `Enter` / `Ctrl+Enter` / `F3` / `→` 做什么); 0 = 不显示 |
| ShowOn | `Mouse` | 搜索窗口显示在哪块屏幕: `Mouse` 鼠标所在 / `Primary` 主屏幕 / `Active` 当前窗口所在 |
| Fade | 0 | 显示和隐藏搜索窗口时淡入淡出 (淡入约 120 ms 到主题的透明度, 淡出约 100 ms); 窗口照样马上获得焦点, 不影响打字 |
| RememberPosition | 0 | 拖动搜索窗口 (按住输入框四周的空白处) 后记住位置 |
| Position | `{"X": 500, "Y": 200}` | 记住的位置, 在屏幕工作区里的千分比: X 0 = 最左, 1000 = 最右; Y 是窗口上沿离顶部的比例。默认 = 水平居中、离顶部 20% |

## Features 搜索功能
每个功能都有 `Enabled` (1 = 开启)。

### Applications 应用
| 键 | 默认 | 说明 |
|---|---|---|
| Folders | 开始菜单 (当前用户 / 所有用户)、桌面 (当前用户 / 公共) | 索引的文件夹, 可以用 [路径变量](Custom-Commands-zh#路径里可以用的变量) |
| FileTypes | `*.lnk` `*.exe` `*.url` `*.appref-ms` | 索引的文件类型 |
| Depth | 3 | 子文件夹深度 |
| Exclude | `i)(uninstall\|卸载\|readme\|help\|documentation)` | 名称匹配这个正则的不收录 |
| Hidden | `[]` | 在搜索结果里按 `Ctrl+Del` 删除的应用路径, 删掉一行即可恢复 |
| StoreApps | 1 | 包含应用商店应用 |
| MatchPinyin | 1 | 中文名称按拼音首字母匹配 |

索引靠监视 `Folders` 自动更新: 有快捷方式新增、删除、改名, 几秒后就在后台静默重建 (两次至少隔 1 分钟), 不弹通知, 搜索窗口打开时暂停; 偏好设置里改了 `Folders`、`FileTypes`、`Exclude` 或 `Depth`, 点 "应用" 后也会重建; 另外缓存超过 1 小时也重建 (兜底)。应用商店应用不在这些文件夹里, 最多每小时扫一次

### Snippets 文字片段
| 键 | 默认 | 说明 |
|---|---|---|
| Keyword | `snip` | 只搜索片段的关键字 |
| SearchText | 1 | 也搜索正文 (至少输入 3 个字符, 每个词都要出现; 只有正文匹配的排在后面) |
| PasteMode | `Auto` | `Auto` = 100 个字符以内的单行文字直接输入, 其它用剪贴板; `Clipboard` = 剪贴板 + Ctrl+V; `Type` = 逐字输入 |
| PasteDelay | 300 | 粘贴后等待多少毫秒再还原剪贴板 |
| AutoExpand | 1 | 在任何程序里输入 前缀 + 关键字 自动展开 |
| ExpandPrefix | `;` | 自动展开的前缀 |
| ExpandExclude | `["ahk_exe mstsc.exe", "ahk_exe KeePass.exe", "ahk_exe KeePassXC.exe"]` | 这些窗口里不自动展开 (每一项: `ahk_exe`、`ahk_class` 或标题的一部分); 密码输入框里总是不展开 |

### Clipboard 剪贴板历史
| 键 | 默认 | 说明 |
|---|---|---|
| Keyword | `clip` | 关键字 |
| Hotkey | `^!c` | 直接打开剪贴板历史的热键 |
| MaxItems | 200 | 保存多少条 |
| MaxItemLength | 100000 | 超过这么多字的内容不记录 |
| Persist | 1 | 保存到磁盘 (`ClipboardHistory.json`, 超过 4000 字的条目和图片单独存在 `Clipboard\` 文件夹; 存在哪里见 `LocalFiles`); 0 = 只在内存里 |
| LocalFiles | 1 | 剪贴板历史存在哪里 (偏好设置 → 剪贴板 → "保存位置"); 本机 = `%LOCALAPPDATA%\ALTRun`, 不跟着同步盘或 U 盘里的 Data 文件夹走。**1** = 图片和很长的条目 (`Clipboard\` 文件夹) 只存本机: Data 在 OneDrive 等同步盘里时, 复制图片不会触发上传; `ClipboardHistory.json` 仍在 Data 里照常同步, 另一台电脑上没有对应文件的图片 / 长条目不显示。**2** = `ClipboardHistory.json` 也只存本机, 剪贴板历史完全不离开这台电脑。**0** = 都放在 Data 文件夹 (例如放在 U 盘里随身带)。改了之后, 下次启动时自动把文件从原来的位置搬过去; 两边都有时 (例如同步盘里另一台电脑的历史) 不覆盖, 用新位置的 |
| Images | 1 | 也记录图片 (存成 PNG); 需要 `Persist = 1` |
| MaxImages | 50 | 最多保存多少张图片, 超过时先删最早的 |
| MergeDoubleCopy | 0 | 1 = 快速按两次 `Ctrl+C` 时把这次复制的文字接到上一条后面 |
| IgnoreApps | KeePass、KeePassXC、1Password、Bitwarden | 不记录这些程序复制的内容 (进程名) |

### Calculator 计算器
| 键 | 默认 | 说明 |
|---|---|---|
| StructuralCalc | 0 | 结果下方附带梁主筋 / 配筋面积计算, 见 [计算器](Calculator-zh) |
| RebarCover | 40 | 结构计算: 保护层 (mm, 每边) |
| MaxBarSpacing | 300 | 结构计算: 主筋最大间距 (mm) |
| BarSizes | `[13, 16, 20, 25, 32]` | 结构计算: 配筋面积那一行列出的钢筋直径 (mm) |
| BarPrefix | `H` | 结构计算: 钢筋代号前缀 (`T`、`Y`、`Φ`...) |
| Currency | 0 | 货币换算 (`100 usd to sgd`), 每天从 frankfurter.dev 下载汇率, 见 [单位和货币换算](Calculator-zh#单位和货币换算) |

### Bookmarks 浏览器书签
| 键 | 默认 | 说明 |
|---|---|---|
| Keyword | `bm` | 只搜书签的关键字 |
| InDefaultResults | 1 | 直接输入名称时也显示匹配的书签 (排在应用和命令后面, 最多 8 条) |

偏好设置 → 功能 → 功能选项里也能改。读 Chrome、Edge、Brave、Vivaldi 所有用户配置的书签; Firefox 的书签存在 SQLite 数据库里, 暂不支持。

### Scripts 脚本扩展
| 键 | 默认 | 说明 |
|---|---|---|
| Keyword | `script` | `script 备份` 只搜脚本, `script ` 列出全部脚本 (只输入 `script` 不带空格时, 照常显示名字里有这个词的命令) |

脚本放在程序目录的 `Scripts\` 里 (程序目录不能写入时在 `%APPDATA%\ALTRun\Scripts\`), 写法见 [脚本扩展](Scripts-zh)。

### Recent 置顶和最近使用
| 键 | 默认 | 说明 |
|---|---|---|
| RecentCount | 0 | 空搜索框里显示几个最近打开的项目, 0 = 不显示 (默认只显示置顶的项目) |
| Pinned | `[]` | 置顶到空搜索框的项目 (在操作面板里置顶 / 取消, 不用手动改) |

最近打开的项目记在 `Data\Knowledge.json` 里, 见 [置顶和最近使用](Recent-zh)。

### Windows 切换窗口
| 键 | 默认 | 说明 |
|---|---|---|
| Keyword | `w` | 只搜已打开的窗口的关键字; 只输入 `w ` 列出全部窗口 |
| InDefaultResults | 1 | 直接输入时也显示标题匹配得好的窗口 (排在应用和命令后面, 最多 3 个) |

偏好设置 → 功能 → 功能选项里也能改。只列出任务栏上能看到的窗口 (不包括工具窗口、对话框、别的虚拟桌面上的窗口和 ALTRun 自己的窗口)。

### WebSearch 网页搜索
| 键 | 说明 |
|---|---|
| Engines | 搜索引擎列表, 每一条 `{ "Id", "Keyword", "Title", "Url", "Icon" }`, `Keyword` 可以写几个, 用逗号分隔 (例如 `"g, ㅎ"`)。`Url` 里的 `{query}` 替换为输入的文字。写成 `{query|天气}` 时有默认值: 输入关键字加空格、不输入搜索词就搜索 "天气"; 默认值写成 `clipboard` (`{query|clipboard}`) 时搜索剪贴板里的文字, 自带的引擎都是这样写的 (从旧版本升级时, 网址还是原来默认值的自带引擎会自动换成这样)。`Icon` (可选): 结果的图标, `.ico`、`.png` 或 `.exe` 文件, 相对路径从 Data 文件夹算起 (例如 `Icons\jira.png`)。编辑框里的 **下载网站图标** 会访问网址所在的网站一次, 把它的图标存到 `Data\Icons` 并填好这一项。留空 = 自带的图标 (默认的几个引擎都有) 或浏览器的图标 |
| Fallbacks | 没有任何结果时显示的兜底项: 引擎的 Id, 或 `files` (文件搜索)。默认 `["google", "files", "bing"]` |

### FileSearch 文件搜索
见 [文件搜索](File-Search-zh#设置)。

### Terminal 终端
| 键 | 默认 | 说明 |
|---|---|---|
| Prefix | `>` | 前缀 |
| Shell | `cmd` | `cmd` / `powershell` / `pwsh` / `wt` |

### Help 速查表
| 键 | 默认 | 说明 |
|---|---|---|
| Enabled | 1 | 输入 `?` 显示所有输入语法和快捷键 |

### LargeType 大字显示
| 键 | 默认 | 说明 |
|---|---|---|
| Enabled | 0 | `Ctrl+L` 和操作面板里的 "大字显示"。新安装默认关; 从 2026.10.19.15 以前的版本升级的保持打开 |

### System 系统命令
| 键 | 默认 | 说明 |
|---|---|---|
| ConfirmActions | 1 | 关机、重启、注销、清空回收站前确认 |
| SettingsPages | 1 | 也搜索 Windows 设置的页面 (显示、蓝牙、WLAN、默认应用、Windows 更新等 40 多个, 打开 `ms-settings:` 对应的页面) |
| Hidden | `[]` | 在搜索结果里按 `Ctrl+Del` 删除的内置命令的 Id (例如 `"Printers"`), 删掉一项即可恢复 |

## Extensions 扩展功能
见 [扩展功能](Extensions-zh)。

| 节点 | 键 |
|---|---|
| QuickSwitch | 见 [对话框快速跳转](Quick-Switch-zh) 和 [对话框文件夹面板](Quick-Switch-Panel-zh) |
| AutoDate | 见 [一键加日期](Date-Stamp-zh) |
| PTTools | PT 工具箱自己保存的输入和窗口位置 |

## 列表
```jsonc
"Hotkeys": [
  { "Key": "^!#v", "Action": "PastePlain", "WinTitle": "" },
  { "Key": "^!t", "Action": "TerminalHere", "WinTitle": "ahk_class CabinetWClass" },
  { "Key": "^r", "Action": "RebuildIndex", "WinTitle": "ALTRun" }
],
"CustomCommands": [
  { "Title": "Desktop", "Type": "Folder", "Target": "A_Desktop", "Arguments": "", "Keyword": "" },
  { "Title": "IP Configuration", "Type": "Command", "Target": "cmd.exe", "Arguments": "/k ipconfig /all", "Keyword": "ipconfig" }
],
"Snippets": [
  { "Name": "Today's date", "Keyword": "today", "Text": "{date}", "AutoExpand": 1 }
]
```
字段说明见 [自定义命令](Custom-Commands-zh#字段)、[文字片段](Snippets-zh) 和 [自定义热键](Hotkeys-zh)。

## Data 文件夹
数据文件夹在哪里、里面每个文件是什么、能不能删、怎样备份, 见 [数据文件](Data-Files-zh)。
