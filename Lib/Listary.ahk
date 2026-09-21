;===============================================================================
; Listary.ahk - 打开/保存对话框路径快速同步 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 功能和 Listary/Directory Opus 里的"路径联动"类似: 在标准的打开/保存文件
; 对话框里, 按一下热键就能把对话框当前路径跳转到 Total Commander 或资源
; 管理器正在浏览的目录, 不用手动复制粘贴路径。
;
; 用法 (ALTRun.ahk 启动时调用一次):
;   Listary.Init()
;
; 内部依赖的全局变量 (ALTRun.ahk 顶部已定义, 这里只读不写):
;   g_CONFIG["FileMgrID"/"DialogWin"/"ExcludeWin"/"AutoSwitchDir"]
;   g_HOTKEY["ExplorerDir"/"TotalCMDDir"]
;   g_LNG[219/220/221], g_TITLE, g_LOG
;===============================================================================

Class Listary {

    ; 程序启动时调用一次: 注册窗口分组、快捷键, 并按需开启自动跳转监控线程。
    static Init() {
        Loop Parse, g_CONFIG["FileMgrID"], ","                              ; File Manager Class, default is Windows Explorer & Total Commander
            GroupAdd("FileMgrID", A_LoopField)

        Loop Parse, g_CONFIG["DialogWin"], ","                              ; 需要QuickSwith的窗口, 包括打开/保存对话框等
            GroupAdd("DialogBox", A_LoopField)

        Loop Parse, g_CONFIG["ExcludeWin"], ","                             ; 排除特定窗口,避免被 Auto-QuickSwitch 影响
            GroupAdd("ExcludeWin", A_LoopField)

        if (g_CONFIG["AutoSwitchDir"]) {
            g_LOG.Debug("Listary: Auto-QuickSwitch enabled, monitoring thread...")
            Loop {
                WinWaitActive("ahk_class TTOTAL_CMD")
                WinWaitNotActive()

                ; 检测当前窗口是否符合打开/保存对话框条件
                if (Listary.IsQuickSwitchDialog()) {
                    winTitle := WinGetTitle("A")
                    procName := WinGetProcessName("A")
                    g_LOG.Debug("Listary: Dialog detected, active window ahk_title=" winTitle ", ahk_exe=" procName)
                    Listary.SyncTCPath()                                    ; NO Return, as will terimate loop (AutoSwitchDir)
                }
                Sleep 100  ; Reduce CPU usage
            }
        }

        HotIf((p*) => Listary.IsQuickSwitchDialog(p*))                      ; 仅在真正的打开/保存对话框启用路径定位热键
        try {
            Hotkey(g_HOTKEY["ExplorerDir"], (p*) => Listary.SyncExplorerPath(p*)) ; Ctrl+E 把打开/保存对话框的路径定位到资源管理器当前浏览的目录
            Hotkey(g_HOTKEY["TotalCMDDir"], (p*) => Listary.SyncTCPath(p*)) ; Ctrl+G 把打开/保存对话框的路径定位到TC当前浏览的目录
            g_LOG.Debug("Listary: Set quickswitch hotkey " g_HOTKEY["ExplorerDir"] " for Explorer, " g_HOTKEY["TotalCMDDir"] " for Total Commander...OK")
        } catch as e {
            g_LOG.Debug("Listary: Failed to set quickswitch hotkey..." e.Message)
        }
        HotIf                                                                ; Turn off context, make subsequent hotkeys global again

        ; 在打开/保存对话框标题中显示快捷键信息
        SetTimer(() => Listary.ShowListaryHint(), 250)
        return
    }

    static ShowListaryHint() {
        static originalTitles := Map()
        static activeHwnd := 0

        if (Listary.IsQuickSwitchDialog()) {
            dlgHwnd := WinGetID("A")
            if (!dlgHwnd || !WinExist("ahk_id " dlgHwnd))
                return

            ; Restore the previous dialog before switching to another one.
            if (activeHwnd && activeHwnd != dlgHwnd && originalTitles.Has(activeHwnd)) {
                if WinExist("ahk_id " activeHwnd)
                    WinSetTitle(originalTitles[activeHwnd], "ahk_id " activeHwnd)
                originalTitles.Delete(activeHwnd)
            }

            if (!originalTitles.Has(dlgHwnd))
                originalTitles[dlgHwnd] := WinGetTitle("ahk_id " dlgHwnd)

            title := originalTitles[dlgHwnd] " / " Listary.GetListaryHintText()
            if (WinGetTitle("ahk_id " dlgHwnd) != title)
                WinSetTitle(title, "ahk_id " dlgHwnd)
            activeHwnd := dlgHwnd
            return
        }

        ; Restore the native title after leaving the file dialog.
        if (activeHwnd && originalTitles.Has(activeHwnd)) {
            if WinExist("ahk_id " activeHwnd)
                WinSetTitle(originalTitles[activeHwnd], "ahk_id " activeHwnd)
            originalTitles.Delete(activeHwnd)
            activeHwnd := 0
        }
    }

    static GetListaryHintText() {
        tcHotkey  := Win.HotkeyLabel(g_HOTKEY["TotalCMDDir"])
        expHotkey := Win.HotkeyLabel(g_HOTKEY["ExplorerDir"])
        msg := StrReplace(g_LNG[221], "{1}", tcHotkey)
        return StrReplace(msg, "{2}", expHotkey)
    }

    ; 更严格地识别 "打开/保存文件" 对话框，避免普通 #32770 对话框误触发
    static IsQuickSwitchDialog(*) {
        if (!WinActive("ahk_group DialogBox") || WinActive("ahk_group ExcludeWin"))
            return false

        winTitle := ""
        try winTitle := WinGetTitle("A")

        ctrlNames := []
        try ctrlNames := WinGetControls("A")
        catch
            ctrlNames := []

        hasNameEdit := Listary.HasAnyCtrlMatch(ctrlNames, "^Edit\d+$")              ; 文件名输入框
        hasFileView := Listary.HasAnyCtrlMatch(ctrlNames, "^DirectUIHWND\d+$", "^SHELLDLL_DefView\d*$", "^SysListView32\d*$")
        hasPathCtrl := Listary.HasAnyCtrlMatch(ctrlNames, "^ToolbarWindow32\d+$", "^ComboBoxEx32\d+$", "^ComboBox\d+$")
        hasMainBtn  := Listary.HasAnyCtrlMatch(ctrlNames, "^Button\d+$")

        ; DialogBox group already limits window classes via config (e.g. #32770 / Qt5QWindowIcon),
        ; so use a unified control rule here to avoid per-app special cases.
        if (ctrlNames.Length > 0) {
            ; Allow either main buttons or path bar to accommodate app-hosted dialog variations.
            return hasNameEdit && hasFileView && (hasMainBtn || hasPathCtrl)
        }

        ; Fallback when control enumeration fails (e.g. privilege boundary / owner-drawn dialogs).
        return Listary.IsLikelyFileDialogTitle(winTitle)
    }

    static HasAnyCtrlMatch(ctrlNames, patternList*) {
        for ctrlName in ctrlNames {
            for pattern in patternList {
                if RegExMatch(ctrlName, "i)" pattern)
                    return true
            }
        }
        return false
    }

    static IsLikelyFileDialogTitle(title) {
        if (!title)
            return false
        return RegExMatch(title, "i)(open|save|select|attach|import|export|reference|file|browse|打开|另存|选择|导入|导出|引用)")
    }

    ; Sync dialog box to Total Commander path (TC 7.x ~ 11.x)
    static SyncTCPath(*) {
        clipSaved   := ClipboardAll()
        A_Clipboard := ""
        ; Get the HWND of TC (WinGetID may occur error if TC not found)
        tcHwnd := WinExist("ahk_class TTOTAL_CMD")
        if (!tcHwnd) {
            MsgBox(g_LNG[219], g_TITLE, 48)
            g_LOG.Debug("SyncTCPath: No Total Commander window found")
            return
        }
        try {
            SendMessage(1075, 2029, 0, , "ahk_class TTOTAL_CMD")            ; TC: WM_USER + 75, TC_GETCURRENTPATH = 2029
        } catch as e {
            g_LOG.Debug("SyncTCPath: SendMessage failed, exception - " . e.Message)
            A_Clipboard := clipSaved
            return
        }
        ; Wait up to 0.1 seconds for the clipboard to contain data
        if (ClipWait(0.1) = 0) {
            A_Clipboard := clipSaved
            g_LOG.Debug("SyncTCPath: Clipboard wait timed out")
            return
        }
        ; 确保路径以反斜杠结尾, 解决AutoCAD不识别路径问题
        targetDir   := RTrim(A_Clipboard, "\") . "\"
        A_Clipboard := clipSaved
        Listary.SetDialogPath(targetDir)
    }

    ; Sync dialog box to Explorer path (Win7 ~ Win11)
    static SyncExplorerPath(*) {
        ; Get the HWND of Explorer (WinGetID may occur error if Explorer not found)
        expHwnd := WinExist("ahk_class CabinetWClass")
        if (!expHwnd) {
            MsgBox(g_LNG[220], g_TITLE, 48)
            g_LOG.Debug("SyncExplorerPath: No Explorer window found")
            return
        }
        try {
            for shellWin in ComObject("Shell.Application").Windows
                if (shellWin.HWND = expHwnd) {
                    targetDir := shellWin.Document.Folder.Self.Path
                    Listary.SetDialogPath(targetDir)
                    return
                }
            g_LOG.Debug("SyncExplorerPath: No matching Explorer window")
        } catch as e {
            g_LOG.Debug("SyncExplorerPath: COM error - " e.Message)
        }
    }

    ; Set dialog box path to specified directory
    static SetDialogPath(targetDir) {
        if (!targetDir || !FileExist(targetDir)) {
            g_LOG.Debug("SetDialogPath: Invalid directory :" targetDir)
            return
        }
        activeClass := WinGetClass("A")
        if (activeClass = "Qt5QWindowIcon") {
            ; WPS dialog: Its Edit control has no valid id, try simulate keyboard input (not fully reliable)
            SendText targetDir
            SendInput "{Enter}"
            g_LOG.Debug("SetDialogPath: Set path to " targetDir " (WPS dialog)")
        } else {
            ; Windows Standard dialog: Edit1 is the path input box
            editHwnd := ControlGetHwnd("Edit1", "A")
            if (editHwnd) {
                ControlFocus("Edit1", "A")
                ControlSetText(targetDir, "Edit1", "A")
                ControlSend("{Enter}", "Edit1", "A")
                g_LOG.Debug("SetDialogPath: Set dialog path to=" targetDir)
            } else {
                ; Fallback for dialogs without Edit1 (some owner-drawn or app-customized file dialogs).
                SendInput "!d"
                Sleep 60
                SendText targetDir
                SendInput "{Enter}"
                g_LOG.Debug("SetDialogPath: Edit1 not found, used address bar fallback path=" targetDir)
            }
        }
    }
}
