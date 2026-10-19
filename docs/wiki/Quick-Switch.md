[English](en-Quick-Switch) · **中文**

# 对话框快速跳转 (Quick Switch)

在 Windows 的打开 / 保存对话框里, 快速跳到已经打开的文件夹, 不必重新浏览目录。文件夹面板的使用方式和设置见[对话框文件夹面板](Quick-Switch-Panel)。

## 快捷键

在标准的打开 / 保存文件对话框里:

- `Ctrl+G`: 跳到 Total Commander 当前打开的文件夹
- `Ctrl+E`: 跳到资源管理器当前打开的文件夹
- `Ctrl+Shift+G`: 打开文件夹面板并把光标放到搜索框, 用键盘选择目标文件夹或文件

跳转时, 文件名框里原来的内容会保留。例如另存为对话框预填的文件名不会被清掉。

## 自动跳转

在偏好设置 → 对话框跳转里打开“自动跳转”后, 从 Total Commander 切换到打开 / 保存对话框时, ALTRun 会自动跳到 Total Commander 当前面板的文件夹。

“生效的窗口”: 勾选 “标准的打开 / 保存对话框” (`ahk_class #32770`, 默认勾选); 其它程序自己画的对话框写在 “其它对话框” 里, 每行一个, 例如 WPS 的 `ahk_class Qt5QWindowIcon`。“不生效的窗口” 里的窗口不跳转; “不自动跳转的对话框” 只关闭自动跳转, 手动热键仍然可用。

Total Commander 的路径通过其 `WM_COPYDATA` 接口直接读取 (需要 TC 8.0 或更新), 不经过剪贴板, 不会把路径写进 Windows 剪贴板历史 (`Win+V`)。

## 相关设置

![偏好设置 → 对话框跳转](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-quickswitch.png)

在偏好设置 → 对话框跳转中可修改热键、适用窗口和自动跳转行为。对应设置文件为 `Extensions.QuickSwitch`:

| 设置 | 默认值 / 作用 |
|---|---|
| `Enabled` | `1` = 启用 Quick Switch |
| `ExplorerHotkey` | 跳到资源管理器当前文件夹, 默认 `Ctrl+E` |
| `TotalCmdHotkey` | 跳到 Total Commander 当前文件夹, 默认 `Ctrl+G` |
| `AutoSwitch` | 从 Total Commander 切换到对话框时自动跳转, 默认 `0` (关) |
| `DialogWindows` | 允许生效的对话框窗口条件 |
| `ExcludeWindows` | 完全排除的窗口条件 |
| `AutoSwitchExclude` | 不自动跳转、但仍可手动操作的窗口条件 |

[对话框文件夹面板](Quick-Switch-Panel)有独立设置页, 可自动显示在打开 / 保存对话框附近。
