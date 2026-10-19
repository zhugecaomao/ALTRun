**English** · [中文](Date-Stamp)

# Date Stamp

Press the hotkey (`Ctrl+D` by default, can be changed):
- **While renaming a file** (Explorer, Total Commander, desktop, dialogs): adds ` - date` before the extension; an existing date is updated to today. For example `Report.docx` → `Report - 23.09.2026.docx`. It only works in the rename box; in the file list Ctrl+D keeps doing what the program does (delete in Explorer)
- **End of the text** (for example Total Commander's file comments): adds ` - date` at the end of the current line, for example `Checked` → `Checked - 23.09.2026`

Preferences → Date Stamp: each case has its own hotkey and "works in" window list (one per line, for example `ahk_class CabinetWClass` for Explorer, `ahk_class TTOTAL_CMD` for Total Commander). In the settings file this is `Extensions.AutoDate` → `RenameHotkey` / `RenameWindows` / `AppendHotkey` / `AppendWindows`.

The date format `DateFormat` is `dd.MM.yyyy` by default (syntax: [AutoHotkey FormatTime](https://www.autohotkey.com/docs/v2/lib/FormatTime.htm)); `{date}` in snippets uses the same format.
