[English](Quick-Switch) · **中文**

# 对话框快速跳转 (Quick Switch)

在 Windows 的打开 / 保存对话框里, 快速跳到已经打开的文件夹, 不必重新浏览目录。文件夹面板的使用方式和设置见[对话框文件夹面板](Quick-Switch-Panel-zh)。

## 快捷键

在标准的打开 / 保存文件对话框里:

- `Ctrl+G`: 跳到最近用过的文件管理器当前的文件夹: Total Commander、Directory Opus 或资源管理器。ALTRun 按窗口从前到后找到这三种里的第一个, 所以同时开着几个时用的是最近用过的那个; [文件夹面板](Quick-Switch-Panel-zh) 里列出了所有窗口的文件夹
- `Ctrl+Shift+G`: 打开文件夹面板并把光标放到搜索框, 用键盘选择目标文件夹或文件

跳转时, 文件名框里原来的内容会保留。例如另存为对话框预填的文件名不会被清掉。

## 自动跳转

在偏好设置 → 对话框跳转里打开“自动跳转”后, 从 Total Commander、Directory Opus 或资源管理器切换到打开 / 保存对话框时, ALTRun 会自动跳到刚才那个窗口当前的文件夹。

“生效的窗口”: 勾选 “标准的打开 / 保存对话框” (`ahk_class #32770`, 默认勾选; 只认有资源管理器文件列表的对话框, "系统信息" (msinfo32) 这类本身是对话框的程序窗口不算); “浏览文件夹对话框” (默认勾选) 是只有文件夹树的旧式选择文件夹对话框 (见下文); 其它程序自己画的对话框写在 “其它对话框” 里 (点 **编辑...**, 每行一个), 例如 WPS 的 `ahk_class Qt5QWindowIcon` (ALTRun 把文件夹路径输入到对话框的文件名框里再按回车, 大多数 Qt、GTK、Java 的文件对话框都能这样跳转)。“不生效的窗口” 里的窗口不跳转 (这几个列表也是点 **编辑...**, 每行一个窗口); “不自动跳转的对话框” 只关闭自动跳转, 手动热键仍然可用。

## 浏览文件夹对话框

很多程序选择文件夹时还在用旧式的 "浏览文件夹" 对话框: 只有一棵文件夹树和 确定 / 取消, 没有文件列表。文件夹面板和 `Ctrl+G` 在这里也能用: 文件夹树会展开并选中目标文件夹, 确定还是由你自己点 (在这种对话框里按回车就等于确定)。不想在这种对话框里使用时, 在 “生效的窗口” 里取消 “浏览文件夹对话框” (`BrowseDialogs`)。和其它对话框一样, 在以管理员身份运行的程序里不起作用。

## 文件夹从哪里读取

都不经过剪贴板, 不会把路径写进 Windows 剪贴板历史 (`Win+V`):
- **Total Commander**: 通过其 `WM_COPYDATA` 接口直接读取 (需要 TC 8.0 或更新)
- **Directory Opus**: 用 Opus 自带的 `dopusrt.exe /info …,paths`, 它列出每个 Opus 窗口的标签页 (在正在运行的 `dopus.exe` 旁边找到, 不用设置)。只在开着 Opus 窗口、真正需要文件夹时运行 (按热键、自动跳转, 或对话框出现时的面板; 面板先显示其它文件夹, Opus 的稍后补上), 1.5 秒内的结果直接复用。打开调试日志时会记下每次读取的耗时 (`Perf: Directory Opus paths … ms`)
- **资源管理器**: Shell 窗口列表

## 相关设置

![偏好设置 → 对话框跳转](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-quickswitch.png)

在偏好设置 → 对话框跳转中可修改热键、适用窗口和自动跳转行为。对应设置文件为 `Extensions.QuickSwitch`:

| 设置 | 默认值 / 作用 |
|---|---|
| `Enabled` | `1` = 启用 Quick Switch |
| `Hotkey` | 跳转热键, 默认 `Ctrl+G`: 跳到最近用过的 Total Commander / Directory Opus / 资源管理器窗口当前的文件夹 (设置版本 6 以前叫 `TotalCmdHotkey`) |
| `AutoSwitch` | 从 Total Commander、Directory Opus 或资源管理器切换到对话框时自动跳转, 默认 `0` (关) |
| `StandardDialogs` | `1` = 在标准的打开 / 保存对话框 (`ahk_class #32770`) 里生效 |
| `DialogWindows` | 其它生效的对话框 (程序自己画的对话框), 窗口条件列表, 默认 `[]` |
| `BrowseDialogs` | `1` = 也在只有文件夹树的 "浏览文件夹" 对话框里生效 |
| `ExcludeWindows` | 完全排除的窗口 (列表) |
| `AutoSwitchExclude` | 不自动跳转、但仍可手动操作的窗口 (列表) |

窗口条件写 `ahk_exe 程序.exe`、`ahk_class 窗口类名` 或者标题的一部分, 列表里每项一个, 例如 `"ExcludeWindows": ["ahk_exe 7zG.exe", "ahk_class SysListView32"]`。

[对话框文件夹面板](Quick-Switch-Panel-zh)有独立设置页, 可自动显示在打开 / 保存对话框附近。
