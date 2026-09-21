;===============================================================================
; Plugins.ahk - 自动日期插件 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 在设定好的窗口里按 Ctrl+D 自动追加/更新当前日期, 两个独立场景:
;   1. 文件/文件夹改名场景 (g_HOTKEY["AutoDateBefExt"] 指定的窗口, 如资源管理器/
;      Total Commander): 日期加在文件名里, 扩展名之前。
;   2. 纯文本备注场景 (g_HOTKEY["AutoDateAtEnd"] 指定的窗口, 如 TC 的文件备注框/
;      Notepad2): 日期直接加在文本末尾。
;
; 用法 (ALTRun.ahk 启动时调用一次):
;   Plugins.Init()
;===============================================================================

Class Plugins {

    ; 程序启动时调用一次: 按配置的窗口分组注册两个 Ctrl+D 场景的热键。
    static Init() {
        Loop Parse, g_HOTKEY["AutoDateBefExt"], ","
            GroupAdd("FileListMangr", A_LoopField)

        Loop Parse, g_HOTKEY["AutoDateAtEnd"], ","
            GroupAdd("TextBox", A_LoopField)

        HotIfWinActive("ahk_group FileListMangr")                           ; 针对所有设定好的程序 按Ctrl+D自动在文件(夹)名之后添加日期
        Hotkey(g_HOTKEY["AutoDateBEHKey"], Plugins.RenameWithDate)


        HotIfWinActive("ahk_group TextBox")
        Hotkey(g_HOTKEY["AutoDateAEHKey"], Plugins.LineEndAddDate)
        HotIfWinActive

        g_LOG.Debug("Plugins: Load AutoDate plugins...OK")
        return
    }

    static RenameWithDate(*) {                                             ; 针对所有设定好的程序 按Ctrl+D自动在文件(夹)名之后添加日期
        FocusedHwnd  := ControlGetFocus("A")                                ; 获取当前激活的窗口中的聚焦的控件名称
        FocusedClassNN := ControlGetClassNN(FocusedHwnd)

        if (InStr(FocusedClassNN, "Edit") or InStr(FocusedClassNN, "Scintilla")) ; 如果当前激活的控件为Edit类或者Scintilla1(Notepad2),则Ctrl+D功能生效
            Plugins.NameAddDate("FileListMangr", FocusedClassNN)
        else
            SendInput "^D"                                                  ; 如果不是,则发送原始的Ctrl+D

        g_LOG.Debug("RenameWithDate: Current control=" FocusedClassNN)
        Return
    }

    static LineEndAddDate(*) {                                             ; 针对TC File Comment对话框　按Ctrl+D自动在备注文字之后添加日期
        CurrentDate := FormatTime(, "dd.MM.yyyy")
        SendInput "{End}"
        Sleep 10
        SendInput "{Blind}{Text} - " CurrentDate
        g_LOG.Debug("LineEndAddDate: Add date at end= - " CurrentDate)
    }

    static NameAddDate(WinName, CurrCtrl) {                                 ; 在文件（夹）名编辑框中添加日期,CurrCtrl为当前控件(名称编辑框Edit)
        EditCtrlText := ControlGetText(CurrCtrl, "A")
        SplitPath(EditCtrlText, &fileName, &fileDir, &fileExt, &nameNoExt)
        CurrentDate := FormatTime(, "dd.MM.yyyy")

        ; 仅当扩展名不是空 & 最后是点后直接跟 1~4 个字母/数字时 & 字符 <5 & 不是纯数字, 才把后缀视为真实扩展名（避免像 "1. DWG" 这种点后有空格被误判）, 才加日期在后缀名之前
        if (fileExt != "" && RegExMatch(EditCtrlText, "\.[A-Za-z0-9]{1,4}$") && StrLen(fileExt) < 5 && !RegExMatch(fileExt,"^\d+$")) {
            if RegExMatch(nameNoExt, " - \d{2}\.\d{2}\.\d{4}$") {
                baseName := RegExReplace(nameNoExt, " - \d{2}\.\d{2}\.\d{4}$", "")
            }
            else if RegExMatch(nameNoExt, "-\d{2}\.\d{2}\.\d{4}$") {
                baseName := RegExReplace(nameNoExt, "-\d{2}\.\d{2}\.\d{4}$", "")
            } else {
                baseName := nameNoExt
            }
            NameWithDate := baseName " - " CurrentDate "." fileExt
        } else if (RegExMatch(fileName, " - \d{2}\.\d{2}\.\d{4}$")) {         ; 如果无后缀, 文件(夹)名最后有日期,则更新为当前日期
            NameWithDate := RegExReplace(fileName, " - \d{2}\.\d{2}\.\d{4}$", " - " CurrentDate)
        } else if (RegExMatch(nameNoExt, "-\d{2}\.\d{2}\.\d{4}$")) {
            NameWithDate := RegExReplace(fileName, "-\d{2}\.\d{2}\.\d{4}$", " - " CurrentDate)
        } else {
            NameWithDate := EditCtrlText " - " CurrentDate
        }
        ControlFocus(CurrCtrl, "A")
        ControlSetText(NameWithDate, CurrCtrl, "A")
        SendInput "{Blind}{End}"
        g_LOG.Debug("NameAddDate: Add date to filename= " NameWithDate)
    }
}
