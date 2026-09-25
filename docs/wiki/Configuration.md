# 设置文件参考 (ALTRun.json)

所有设置和用户数据都保存在程序目录下的 `ALTRun.json` (UTF-8 文本)。大部分设置可以在偏好设置窗口里修改 (每个设置下面有一行灰色的说明): **确定** 保存并关闭, **应用** 保存并留在当前页, **取消** 放弃修改, **帮助** (`F1`) 打开当前页的说明; 设置在 ALTRun 重新载入后生效, 由程序自动完成。也可以直接编辑这个文件 (搜索窗口里按 `F4`, 或 偏好设置 → 高级 → 编辑 ALTRun.json), **保存后 ALTRun 自动重新载入**。

- 文件里缺少的项目会自动用默认值补上, 所以只需要写要修改的部分
- 文件格式有错误 (例如少了逗号) 时, ALTRun 会提示, 把原文件改名为 `ALTRun.json.bad` 保留下来, 然后用默认设置启动; 修好后改回原名即可
- `SchemaVersion` 是设置格式的版本号, 不要手动修改; 格式升级时 ALTRun 会自动转换并备份

## 结构
```jsonc
{
  "SchemaVersion": 4,
  "General":    { ... },          // 通用
  "Appearance": { ... },          // 外观
  "Features":   { ... },          // 各个搜索功能
  "Extensions": { ... },          // 扩展功能
  "Hotkeys":        [ ... ],      // 自定义热键
  "CustomCommands": [ ... ],      // 自定义命令
  "Snippets":       [ ... ]       // 文字片段
}
```

## General 通用
偏好设置里分在 "通用" 和 "搜索窗口" 两页。

| 键 | 默认 | 说明 |
|---|---|---|
| Hotkey | `!Space` | 呼出热键 (AutoHotkey 写法, `!` Alt `^` Ctrl `+` Shift `#` Win) |
| SecondaryHotkey | 空 | 第二个呼出热键 |
| Language | `auto` | `auto` 跟随系统 / `en` / `zh` |
| LaunchAtLogin | 1 | 开机自动启动 |
| ShowTrayIcon | 1 | 显示托盘图标 |
| HideOnDeactivate | 1 | 搜索窗口失去焦点时隐藏 |
| SwitchToEnglishInput | 0 | 显示搜索窗口时切换到英文输入法 |
| SpaceToRun | 0 | 已输入文字时按空格执行选中项, `Shift+空格` 输入空格 |
| KeepLastQuery | 0 | 呼出窗口时保留上一次的搜索 (输入、文件搜索模式和选中的行), 文字全选: `Enter` 再执行一次, 直接输入开始新的搜索。从 2.x 升级时沿用 `KeepInput` |
| ShowTips | 1 | 空搜索框里轮换显示使用提示 |
| FileManager | `explorer.exe` | 打开文件夹用的程序, 例如 `C:\Apps\TotalCMD64.exe` |
| SendToMenu | 1 | 添加到资源管理器的 "发送到" 菜单 |
| StartMenuShortcut | 1 | 添加到开始菜单 |
| CheckForUpdates | 1 | 启动时检查 GitHub 上的新版本 |
| SaveLog | 0 | 写入调试日志 (`%Temp%\ALTRun.log`) |
| HistorySize | 30 | 记住多少条最近的搜索 |

## Appearance 外观
| 键 | 默认 | 说明 |
|---|---|---|
| Theme | `Light` | 主题名, 见 [主题](Themes) |
| Width | 700 | 搜索窗口宽度 (像素, 按屏幕缩放) |
| VisibleRows | 8 | 最多显示几行结果 (1~9) |
| ShowOn | `Mouse` | 搜索窗口显示在哪块屏幕: `Mouse` 鼠标所在 / `Primary` 主屏幕 / `Active` 当前窗口所在 |
| RememberPosition | 0 | 拖动搜索窗口 (按住输入框四周的空白处) 后记住位置 |
| Position | `{"X": 500, "Y": 200}` | 记住的位置, 在屏幕工作区里的千分比: X 0 = 最左, 1000 = 最右; Y 是窗口上沿离顶部的比例。默认 = 水平居中、离顶部 20% |

## Features 搜索功能
每个功能都有 `Enabled` (1 = 开启)。

### Applications 应用
| 键 | 默认 | 说明 |
|---|---|---|
| Folders | 开始菜单 (当前用户 / 所有用户)、桌面 (当前用户 / 公共) | 索引的文件夹, 可以用 [路径变量](Commands-and-Snippets#路径里可以用的变量) |
| FileTypes | `*.lnk` `*.exe` `*.url` `*.appref-ms` | 索引的文件类型 |
| Depth | 3 | 子文件夹深度 |
| Exclude | `i)(uninstall\|卸载\|readme\|help\|documentation)` | 名称匹配这个正则的不收录 |
| Hidden | `[]` | 在搜索结果里按 `Ctrl+Del` 删除的应用路径, 删掉一行即可恢复 |
| StoreApps | 1 | 包含应用商店应用 |
| MatchPinyin | 1 | 中文名称按拼音首字母匹配 |
| RefreshMinutes | 60 | 索引多久在后台更新一次 |

### Snippets 文字片段
| 键 | 默认 | 说明 |
|---|---|---|
| Keyword | `snip` | 只搜索片段的关键字 |
| PasteMode | `Clipboard` | `Clipboard` = 剪贴板 + Ctrl+V; `Type` = 逐字输入 |
| PasteDelay | 300 | 粘贴后等待多少毫秒再还原剪贴板 |
| AutoExpand | 1 | 在任何程序里输入 前缀 + 关键字 自动展开 |
| ExpandPrefix | `;` | 自动展开的前缀 |

### Clipboard 剪贴板历史
| 键 | 默认 | 说明 |
|---|---|---|
| Keyword | `clip` | 关键字 |
| Hotkey | `^!c` | 直接打开剪贴板历史的热键 |
| MaxItems | 200 | 保存多少条 |
| MaxItemLength | 100000 | 超过这么多字的内容不记录 |
| Persist | 1 | 保存到磁盘 (`Data\ClipboardHistory.json`); 0 = 只在内存里 |
| IgnoreApps | KeePass、KeePassXC、1Password、Bitwarden | 不记录这些程序复制的内容 (进程名) |

### Calculator 计算器
| 键 | 默认 | 说明 |
|---|---|---|
| StructuralCalc | 0 | 结果下方附带梁主筋 / 配筋面积计算, 见 [扩展功能](Extensions#计算器) |

### WebSearch 网页搜索
| 键 | 说明 |
|---|---|
| Engines | 搜索引擎列表, 每一条 `{ "Id", "Keyword", "Title", "Url" }`, `Url` 里的 `{query}` 替换为输入的文字 |
| Fallbacks | 没有任何结果时显示的兜底项: 引擎的 Id, 或 `files` (文件搜索)。默认 `["google", "files", "bing"]` |

### FileSearch 文件搜索
见 [文件搜索](File-Search#设置)。

### Terminal 终端
| 键 | 默认 | 说明 |
|---|---|---|
| Prefix | `>` | 前缀 |
| Shell | `cmd` | `cmd` / `powershell` / `pwsh` / `wt` |

### Help 速查表
| 键 | 默认 | 说明 |
|---|---|---|
| Enabled | 1 | 输入 `?` 显示所有输入语法和快捷键 |

### System 系统命令
| 键 | 默认 | 说明 |
|---|---|---|
| ConfirmActions | 1 | 关机、重启、注销、清空回收站前确认 |

## Extensions 扩展功能
见 [扩展功能](Extensions)。

| 节点 | 键 |
|---|---|
| QuickSwitch | `Enabled` `ExplorerHotkey` (`^e`) `TotalCmdHotkey` (`^g`) `AutoSwitch` `DialogWindows` `ExcludeWindows` |
| AutoDate | `Enabled` `DateFormat` (`dd.MM.yyyy`) `RenameHotkey` (`^d`) `RenameWindows` `AppendHotkey` (`^d`) `AppendWindows` |
| PTTools | PT 工具箱自己保存的输入和窗口位置 |

## 列表
```jsonc
"Hotkeys": [
  { "Key": "~MButton", "Action": "PTTools", "WinTitle": "ahk_exe RAPTW.exe" }
],
"CustomCommands": [
  { "Title": "Desktop", "Type": "Folder", "Target": "A_Desktop", "Arguments": "", "Keyword": "" },
  { "Title": "IP Configuration", "Type": "Command", "Target": "cmd.exe", "Arguments": "/k ipconfig /all", "Keyword": "ipconfig" }
],
"Snippets": [
  { "Name": "Today's date", "Keyword": "today", "Text": "{date}", "AutoExpand": 1 }
]
```
字段说明见 [自定义命令与片段](Commands-and-Snippets) 和 [扩展功能](Extensions#自定义热键)。

## Data 文件夹
运行时生成, 删掉只会重新生成 (学习记录和剪贴板历史会清空):

| 文件 | 内容 |
|---|---|
| `AppIndex.json` | 应用索引 |
| `FileIndex.json` | 内置文件索引 (没有 Everything 时) |
| `Knowledge.json` | 学习排序和最近的搜索 |
| `ClipboardHistory.json` | 剪贴板历史 |
