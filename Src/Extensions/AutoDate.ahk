;===============================================================================
; AutoDate.ahk - Ctrl+D 自动添加日期 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 两个场景:
;   1. 重命名 (RenameWindows 里的窗口, 如资源管理器 / Total Commander): 在文件名
;      扩展名前加上 " - 日期", 已经有日期的更新为今天。
;   2. 文字备注 (AppendWindows 里的窗口, 如 TC 的文件备注框): 在文字末尾加 " - 日期"。
; 日期格式见 DateFormat (默认 dd.MM.yyyy)。
;
; 设置 (ALTRun.json -> Extensions.AutoDate):
;   Enabled / DateFormat / RenameHotkey / RenameWindows / AppendHotkey / AppendWindows
;
; 用法:
;   AutoDate.Init(AppSettings.Extension("AutoDate"))    启动时
;   AutoDate.AddDateToName("Report.docx")               -> "Report - 23.09.2026.docx"
;===============================================================================

class AutoDate {
    static Options := Map()

    static Init(options) {
        AutoDate.Options := options
        if !options["Enabled"]
            return
        Loop Parse, options["RenameWindows"], ","
            GroupAdd("ALTRunRenameWindows", Trim(A_LoopField))
        Loop Parse, options["AppendWindows"], ","
            GroupAdd("ALTRunAppendWindows", Trim(A_LoopField))
        ; 重命名: 只在焦点是改名的编辑框时才接管热键, 其它时候 (例如资源管理器的文件列表里 Ctrl+D = 删除)
        ; 按键原样交给程序, 不用再自己补发一次。以前是先接管再补发: 补发时如果已经松开了 Ctrl, 而 ALTRun
        ; 记的按键状态还是按着 (电脑睡眠 / 锁屏一夜之后可能这样), 只发出一个 "d", 把选中的文件名换成了 "d"
        try {
            HotIf((*) => WinActive("ahk_group ALTRunRenameWindows") && AutoDate.IsEditControl(AutoDate.FocusedControl()))
            Hotkey(options["RenameHotkey"], (*) => AutoDate._OnRename())
            HotIfWinActive("ahk_group ALTRunAppendWindows")
            Hotkey(options["AppendHotkey"], (*) => AutoDate._OnAppend())
        } catch as e {
            Logger.Error("AutoDate: cannot register hotkeys - " e.Message)
        }
        HotIf()
    }

    ; 当前窗口里有焦点的控件的 ClassNN, 取不到时 ""
    static FocusedControl() {
        try return ControlGetClassNN(ControlGetFocus("A"))
        return ""
    }

    ; 改名 / 输入用的编辑框 (资源管理器、对话框的 Edit, TC 的改名框 TInEdit, Scintilla)
    static IsEditControl(classNN) => (classNN != "" && (InStr(classNN, "Edit") || InStr(classNN, "Scintilla"))) ? true : false

    static Today() {
        return FormatTime(, AutoDate.Format())
    }

    static Format() {
        return (AutoDate.Options is Map && AutoDate.Options.Has("DateFormat") && AutoDate.Options["DateFormat"] != "")
            ? AutoDate.Options["DateFormat"] : "dd.MM.yyyy"
    }

    ; 日期格式 -> 匹配这种格式的日期的正则 ("dd.MM.yyyy" -> "\d{2}\.\d{2}\.\d{4}"), 用来找出名字末尾已有的日期
    static DatePattern(format) {
        static tokens := [["yyyy", "\d{4}"], ["dddd", "\p{L}+"], ["MMMM", "\p{L}+"], ["ddd", "\p{L}+"], ["MMM", "\p{L}+"]
            , ["yy", "\d{2}"], ["MM", "\d{2}"], ["dd", "\d{2}"], ["M", "\d{1,2}"], ["d", "\d{1,2}"], ["y", "\d{1,2}"]]
        pattern := "", i := 1
        while (i <= StrLen(format)) {
            for token in tokens {
                if (SubStr(format, i, StrLen(token[1])) == token[1]) {
                    pattern .= token[2], i += StrLen(token[1])
                    continue 2
                }
            }
            ch := SubStr(format, i, 1), i += 1
            pattern .= InStr("\.*?+[](){}|^$", ch) ? "\" ch : ch
        }
        return pattern
    }

    ; 只在重命名的编辑框里触发 (见 Init 的 HotIf)。万一触发后焦点已经离开编辑框, 什么都不做, 不补发按键
    static _OnRename() {
        focused := AutoDate.FocusedControl()
        if !AutoDate.IsEditControl(focused) {
            Logger.Debug("AutoDate: focus left the edit box (" focused "), nothing done")
            return
        }
        Usage.Count("AutoDate")
        oldName := ControlGetText(focused, "A")
        newName := AutoDate.AddDateToName(oldName)
        Logger.Debug("AutoDate: " focused " in " WinGetClass("A") ", " StrLen(oldName) " -> " StrLen(newName) " characters")
        ControlFocus(focused, "A")
        ControlSetText(newName, focused, "A")
        SendInput("{Blind}{End}")
    }

    static _OnAppend() {
        Usage.Count("AutoDate")
        SendInput("{End}")
        Sleep(10)
        SendInput("{Blind}{Text} - " AutoDate.Today())
    }

    ; 纯函数, 方便测试: 文件名 (可以带扩展名) -> 加上 / 更新日期后的文件名
    ; format: 日期格式 (默认用设置里的 DateFormat); 名字末尾已经是这种格式的日期时换成今天
    static AddDateToName(fileName, today := "", format := "") {
        today := (today != "") ? today : AutoDate.Today()
        datePattern := "\s?-\s?" AutoDate.DatePattern((format != "") ? format : AutoDate.Format()) "$"
        ; 只有 "点 + 1~4 个字母数字" 且不是纯数字时才算扩展名 ("1. DWG" 这种不算)
        if RegExMatch(fileName, "^(.*)\.([A-Za-z0-9]{1,4})$", &m) && !RegExMatch(m[2], "^\d+$")
            return RegExReplace(m[1], datePattern) " - " today "." m[2]
        if RegExMatch(fileName, datePattern)
            return RegExReplace(fileName, datePattern) " - " today
        return fileName " - " today
    }
}
