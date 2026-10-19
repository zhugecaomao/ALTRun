[English](en-Date-Stamp) · **中文**

# 一键加日期

按热键 (默认 `Ctrl+D`, 可以改):
- **重命名文件时** (资源管理器、Total Commander、桌面、对话框): 在扩展名前加上 ` - 日期`, 已经有日期的更新为今天。例如 `Report.docx` → `Report - 23.09.2026.docx`。只在改名的输入框里生效, 在文件列表里按 Ctrl+D 还是程序自己的功能 (资源管理器里是删除)
- **文字末尾** (例如 Total Commander 的文件备注): 在光标所在行的末尾加上 ` - 日期`, 例如 `已审核` → `已审核 - 23.09.2026`

偏好设置 → 一键加日期: 两种场景各有自己的热键和 "生效的窗口" 列表 (每行一个, 例如 `ahk_class CabinetWClass` 资源管理器、`ahk_class TTOTAL_CMD` Total Commander)。设置文件里是 `Extensions.AutoDate` → `RenameHotkey` / `RenameWindows` / `AppendHotkey` / `AppendWindows`。

日期格式 `DateFormat` 默认 `dd.MM.yyyy` (写法见 [AutoHotkey FormatTime](https://www.autohotkey.com/docs/v2/lib/FormatTime.htm)), 片段里的 `{date}` 也用这个格式。
