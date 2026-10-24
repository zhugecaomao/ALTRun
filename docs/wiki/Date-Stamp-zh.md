[English](Date-Stamp) · **中文**

# 一键加日期

一键加日期用同一个热键, 在重命名文件或编辑文本时添加日期。默认热键为 `Ctrl+D`, 可在偏好设置 → 一键加日期里修改。

## 重命名文件

在资源管理器、Total Commander、桌面或打开 / 保存对话框的重命名输入框里按热键, 会在扩展名前添加 ` - 日期`。如果文件名里已经有日期, 会更新为今天。

例如 `Report.docx` 变成 `Report - 23.09.2026.docx`。只有重命名输入框会触发; 在文件列表里按 `Ctrl+D` 仍执行原程序的操作 (例如资源管理器中是删除)。

## 在文本末尾添加日期

在当前行末尾添加 ` - 日期`, 例如 `已审核` 变成 `已审核 - 23.09.2026`。适合 Total Commander 文件备注等文本输入场景。

## 相关设置

![偏好设置 → 一键加日期](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-datestamp.png)

偏好设置 → 一键加日期只有一个“热键”, 两种场景共用; 按下时做哪一种, 由各自的“生效的窗口”列表决定: 点 **编辑...**, 每行写一个窗口条件。例如 `ahk_class CabinetWClass` 表示资源管理器, `ahk_class TTOTAL_CMD` 表示 Total Commander。

设置保存在 `Extensions.AutoDate`:

| 设置 | 默认值 | 作用 |
|---|---|---|
| `Enabled` | `1` | 是否启用一键加日期 |
| `Hotkey` | `Ctrl+D` | 热键, 重命名和追加日期共用 (设置版本 6 以前叫 `RenameHotkey`) |
| `RenameWindows` | 按默认窗口规则 | 允许重命名操作生效的窗口 (窗口条件列表) |
| `AppendWindows` | 按默认窗口规则 | 允许文本追加操作生效的窗口 (窗口条件列表) |
| `DateFormat` | `dd.MM.yyyy` | 日期格式, 也用于文字片段的 `{date}` 占位符 |

日期格式写法见 [AutoHotkey FormatTime](https://www.autohotkey.com/docs/v2/lib/FormatTime.htm)。
