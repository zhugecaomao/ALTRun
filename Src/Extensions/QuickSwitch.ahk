;===============================================================================
; QuickSwitch.ahk - 打开/保存对话框路径快速跳转 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 Listary 的 "Quick Switch" 一样: 在标准的打开/保存文件对话框里按热键, 把对话框
; 跳到 Total Commander (默认 Ctrl+G) 或资源管理器 (默认 Ctrl+E) 当前打开的文件夹。
; 对话框标题上会提示这两个热键。AutoSwitch = 1 时, 从 TC 切换到对话框会自动跳转。
;
; 设置 (ALTRun.json -> Extensions.QuickSwitch):
;   Enabled / ExplorerHotkey / TotalCmdHotkey / AutoSwitch / DialogWindows / ExcludeWindows
;
; 用法:
;   QuickSwitch.Init(AppSettings.Extension("QuickSwitch"))    启动时
;   QuickSwitch.FolderOfWindow(hwnd)     TC / 资源管理器窗口当前的文件夹 (只读, 给 "在此处打开终端" 用)
;===============================================================================

class QuickSwitch {
    static Options := Map()
    static _titles := Map(), _hintedHwnd := 0, _lastWasTC := false

    static Init(options) {
        QuickSwitch.Options := options
        if !options["Enabled"]
            return
        Loop Parse, options["DialogWindows"], ","
            GroupAdd("ALTRunDialogs", Trim(A_LoopField))
        Loop Parse, options["ExcludeWindows"], ","
            GroupAdd("ALTRunDialogExclude", Trim(A_LoopField))

        HotIf((*) => QuickSwitch.IsFileDialog())
        try {
            if (options["ExplorerHotkey"] != "")
                Hotkey(options["ExplorerHotkey"], (*) => QuickSwitch.SyncExplorerPath())
            if (options["TotalCmdHotkey"] != "")
                Hotkey(options["TotalCmdHotkey"], (*) => QuickSwitch.SyncTotalCmdPath())
        } catch as e {
            Logger.Error("QuickSwitch: cannot register hotkeys - " e.Message)
        }
        HotIf()

        SetTimer(() => QuickSwitch._Watch(), 250)
    }

    ; 每 250 ms: 在对话框标题上显示热键提示; AutoSwitch 时从 TC 切到对话框自动跳转
    static _Watch() {
        isDialog := QuickSwitch.IsFileDialog()
        if (isDialog && QuickSwitch.Options["AutoSwitch"] && QuickSwitch._lastWasTC)
            QuickSwitch.SyncTotalCmdPath(true)
        QuickSwitch._lastWasTC := WinActive("ahk_class TTOTAL_CMD") ? true : false
        QuickSwitch._UpdateHint(isDialog)
    }

    static _UpdateHint(isDialog) {
        titles := QuickSwitch._titles
        if isDialog {
            hwnd := WinGetID("A")
            if (QuickSwitch._hintedHwnd && QuickSwitch._hintedHwnd != hwnd)
                QuickSwitch._RestoreTitle(QuickSwitch._hintedHwnd)
            if !titles.Has(hwnd)
                titles[hwnd] := WinGetTitle("ahk_id " hwnd)
            hint := I18n.T("QuickSwitch.Hint", Win.HotkeyLabel(QuickSwitch.Options["TotalCmdHotkey"]), Win.HotkeyLabel(QuickSwitch.Options["ExplorerHotkey"]))
            title := titles[hwnd] " / " hint
            if (WinGetTitle("ahk_id " hwnd) != title)
                WinSetTitle(title, "ahk_id " hwnd)
            QuickSwitch._hintedHwnd := hwnd
        } else if QuickSwitch._hintedHwnd {
            QuickSwitch._RestoreTitle(QuickSwitch._hintedHwnd)
            QuickSwitch._hintedHwnd := 0
        }
    }

    static _RestoreTitle(hwnd) {
        if QuickSwitch._titles.Has(hwnd) {
            try WinSetTitle(QuickSwitch._titles[hwnd], "ahk_id " hwnd)
            QuickSwitch._titles.Delete(hwnd)
        }
    }

    ; 只认真正的打开/保存文件对话框 (有文件名输入框 + 文件列表 + 按钮或路径栏)
    static IsFileDialog(*) {
        if (!WinActive("ahk_group ALTRunDialogs") || WinActive("ahk_group ALTRunDialogExclude"))
            return false
        controls := []
        try controls := WinGetControls("A")
        if !controls.Length {
            title := ""
            try title := WinGetTitle("A")
            return RegExMatch(title, "i)(open|save|select|import|export|browse|打开|另存|选择|导入|导出)") > 0
        }
        hasNameEdit := QuickSwitch._HasControl(controls, "^Edit\d+$")
        hasFileView := QuickSwitch._HasControl(controls, "^DirectUIHWND\d+$", "^SHELLDLL_DefView\d*$", "^SysListView32\d*$")
        hasButtons  := QuickSwitch._HasControl(controls, "^Button\d+$", "^ToolbarWindow32\d+$", "^ComboBoxEx32\d+$", "^ComboBox\d+$")
        return hasNameEdit && hasFileView && hasButtons
    }

    static _HasControl(controls, patterns*) {
        for name in controls
            for pattern in patterns
                if RegExMatch(name, "i)" pattern)
                    return true
        return false
    }

    static SyncTotalCmdPath(silent := false) {
        folder := QuickSwitch.TotalCmdFolder()
        if (folder = "")
            return silent ? "" : MsgBox(I18n.T("QuickSwitch.NoTC"), App.Name, 48)
        QuickSwitch.SetDialogPath(folder "\")                               ; 结尾带反斜杠, AutoCAD 等程序才认
    }

    static SyncExplorerPath() {
        folder := QuickSwitch.ExplorerFolder()
        if (folder = "")
            return MsgBox(I18n.T("QuickSwitch.NoExplorer"), App.Name, 48)
        QuickSwitch.SetDialogPath(folder)
    }

    ; Total Commander 当前面板的文件夹 (TC 7 ~ 11: WM_USER+75, 2029 = 当前路径放到剪贴板)
    static TotalCmdFolder() {
        if !WinExist("ahk_class TTOTAL_CMD")
            return ""
        savedClipboard := ClipboardAll()
        A_Clipboard := ""
        folder := ""
        try {
            SendMessage(1075, 2029, 0, , "ahk_class TTOTAL_CMD")
            if ClipWait(0.2)
                folder := RTrim(A_Clipboard, "\")
        }
        A_Clipboard := savedClipboard
        return folder
    }

    ; 资源管理器窗口当前的文件夹; hwnd 为 0 时取最前面的资源管理器窗口
    static ExplorerFolder(hwnd := 0) {
        if !hwnd
            hwnd := WinExist("ahk_class CabinetWClass")
        if !hwnd
            return ""
        try {
            for window in ComObject("Shell.Application").Windows
                if (window.HWND = hwnd)
                    return window.Document.Folder.Self.Path
        }
        return ""
    }

    static FolderOfWindow(hwnd) {
        if (!hwnd || !WinExist("ahk_id " hwnd))
            return ""
        switch WinGetClass("ahk_id " hwnd) {
            case "TTOTAL_CMD"   : return QuickSwitch.TotalCmdFolder()
            case "CabinetWClass": return QuickSwitch.ExplorerFolder(hwnd)
        }
        return ""
    }

    static SetDialogPath(folder) {
        if (folder = "" || !FileExist(folder))
            return
        if (WinGetClass("A") = "Qt5QWindowIcon") {                          ; WPS 的对话框没有可用的 Edit 控件, 只能模拟输入
            SendText(folder)
            SendInput("{Enter}")
            return
        }
        try {
            ControlFocus("Edit1", "A")
            ControlSetText(folder, "Edit1", "A")
            ControlSend("{Enter}", "Edit1", "A")
        } catch {
            SendInput("!d")                                                 ; 没有 Edit1 的自绘对话框: 用地址栏
            Sleep(60)
            SendText(folder)
            SendInput("{Enter}")
        }
    }
}
