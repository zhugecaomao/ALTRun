;===============================================================================
; AutoDate.ahk - Ctrl+D adds a date (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Two cases:
;   1. renaming (windows in RenameWindows, e.g. Explorer / Total Commander): adds " - date" before the
;      file name's extension; an existing date is updated to today.
;   2. text comments (windows in AppendWindows, e.g. TC's file comment box): adds " - date" at the end of the text.
; Date format: see DateFormat (default dd.MM.yyyy).
;
; Settings (ALTRun.json -> Extensions.AutoDate):
;   Enabled / DateFormat / Hotkey (one hotkey for both) / RenameWindows / AppendWindows (arrays of window conditions)
;
; Usage:
;   AutoDate.Init(AppSettings.Extension("AutoDate"))    at start-up
;   AutoDate.AddDateToName("Report.docx")               -> "Report - 23.09.2026.docx"
;===============================================================================

class AutoDate {
    static Options := Map()

    static Init(options) {
        AutoDate.Options := options
        if !options["Enabled"]
            return
        for window in options["RenameWindows"]
            GroupAdd("ALTRunRenameWindows", window)
        for window in options["AppendWindows"]
            GroupAdd("ALTRunAppendWindows", window)
        ; Renaming: the hotkey is only taken while the focus is in the rename box; at other times (e.g. Ctrl+D = delete in Explorer's file list)
        ; the key goes to the program unchanged, with no need to send it again. It used to take the key and send it on: if Ctrl was already up when sending but ALTRun
        ; still thought it was down (possible after a night of sleep / lock), only a "d" was sent, replacing the selected file name with "d"
        try {
            HotIf((*) => WinActive("ahk_group ALTRunRenameWindows") && AutoDate.IsEditControl(AutoDate.FocusedControl()))
            Hotkey(options["Hotkey"], (*) => AutoDate._OnRename())
            HotIfWinActive("ahk_group ALTRunAppendWindows")                 ; the same hotkey: the window decides
            Hotkey(options["Hotkey"], (*) => AutoDate._OnAppend())
        } catch as e {
            Logger.Error("AutoDate: cannot register hotkeys - " e.Message)
        }
        HotIf()
    }

    ; ClassNN of the focused control in the current window, "" if it can't be read
    static FocusedControl() {
        try return ControlGetClassNN(ControlGetFocus("A"))
        return ""
    }

    ; edit boxes for renaming / typing (Edit in Explorer and dialogs, TC's rename box TInEdit, Scintilla)
    static IsEditControl(classNN) => (classNN != "" && (InStr(classNN, "Edit") || InStr(classNN, "Scintilla"))) ? true : false

    static Today() {
        return FormatTime(, AutoDate.Format())
    }

    static Format() {
        return (AutoDate.Options is Map && AutoDate.Options.Has("DateFormat") && AutoDate.Options["DateFormat"] != "")
            ? AutoDate.Options["DateFormat"] : "dd.MM.yyyy"
    }

    ; Date format -> a regex matching dates in that format ("dd.MM.yyyy" -> "\d{2}\.\d{2}\.\d{4}"), to find a date already at the end of the name
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

    ; Only fires in a rename box (see the HotIf in Init). If the focus has left the box after firing, do nothing and don't send the key on
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

    ; Pure function, easy to test: a file name (with or without extension) -> the name with the date added / updated
    ; format: the date format (default: DateFormat from the settings); a date in this format at the end of the name is replaced by today
    static AddDateToName(fileName, today := "", format := "") {
        today := (today != "") ? today : AutoDate.Today()
        datePattern := "\s?-\s?" AutoDate.DatePattern((format != "") ? format : AutoDate.Format()) "$"
        ; only "dot + 1~4 letters or digits" that aren't all digits count as an extension ("1. DWG" doesn't)
        if RegExMatch(fileName, "^(.*)\.([A-Za-z0-9]{1,4})$", &m) && !RegExMatch(m[2], "^\d+$")
            return RegExReplace(m[1], datePattern) " - " today "." m[2]
        if RegExMatch(fileName, datePattern)
            return RegExReplace(fileName, datePattern) " - " today
        return fileName " - " today
    }
}
