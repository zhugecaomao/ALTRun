;===============================================================================
; CapsLockHelper.ahk - 管理员窗口里也能按 CapsLock 切换输入法 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Windows 不让普通权限的程序控制管理员权限的窗口 (UIPI): 前台是管理员窗口时, ALTRun 的键盘钩子收不到按键,
; 发给它的切换消息也会被挡掉 (#158)。整个 ALTRun 以管理员身份运行又会让打开的程序都变成管理员权限
; (看不到映射的网络盘、拖放被拒)。所以单独起一个以管理员身份运行的小助手, 只做 CapsLock 切换:
;   - 设置 General.CapsLockElevated = 1 时 (偏好设置 → 自定义热键), 建一个 "最高权限" 的计划任务 (只在这时
;     确认一次 UAC); ALTRun 每次启动都运行这个任务 (schtasks /run, 不弹 UAC) 启动小助手, 退出时让它也退出
;   - 小助手就是 ALTRun 本身, 带参数 -CapsLockHelper 启动: 只注册 CapsLock 热键 (CapsLockSwitch), 不建窗口、
;     不写设置、不运行任何程序。主窗口的标题改成 Title, 不会被当作 "已经在运行的 ALTRun"
;   - ALTRun 和小助手都注册 CapsLock 热键: 键盘钩子按安装的先后依次调用, 先收到的那个拦下按键, 后一个就收不到,
;     所以普通窗口里只会切换一次; 管理员窗口里只有小助手收得到
;   - 大写锁定开 / 关的提示由 ALTRun 显示 (小助手发 "-CapsLockState 1/0" 过去, 管理员权限的程序给普通程序发消息不受限制)
;   - 小助手只接受 "-Exit" (WM_COPYDATA, 放行普通程序发来的这一种消息); 找不到 ALTRun 时自己退出
;===============================================================================

class CapsLockHelper {
    static Arg      := "-CapsLockHelper"
    static TaskName := "ALTRun CapsLock Helper"
    static Title    := "ALTRun CapsLock Helper"

    ;---------------------------------------------------------------------------
    ; 小助手这边 (以管理员身份运行的进程)
    ;---------------------------------------------------------------------------
    static Run() {
        A_IconHidden := true
        DetectHiddenWindows(true)
        for hwnd in WinGetList(CapsLockHelper.Title " ahk_class AutoHotkey")   ; 已经有一个在运行
            if (hwnd != A_ScriptHwnd)
                ExitApp()
        DllCall("SetWindowTextW", "Ptr", A_ScriptHwnd, "Str", CapsLockHelper.Title)
        DllCall("ChangeWindowMessageFilterEx", "Ptr", A_ScriptHwnd, "UInt", 0x4A, "UInt", 1, "Ptr", 0)   ; WM_COPYDATA, MSGFLT_ALLOW
        OnMessage(0x4A, (p*) => CapsLockHelper._OnCopyData(p*))
        mode := CapsLockHelper.ModeFromSettings(AppSettings.File)
        if !CapsLockSwitch.IsMode(mode)
            ExitApp()
        CapsLockSwitch.Mode := mode
        CapsLockSwitch.Notify := (on) => CapsLockHelper._TellMain(on)
        CapsLockSwitch.Register((key, fn) => Hotkey(key, fn))
        SetTimer(() => App._RunningInstance() || ExitApp(), 3000)           ; ALTRun 不在了 (被结束、崩溃): 跟着退出
    }

    ; 设置文件里的 General.CapsLock (只读, 不写回); 读不到时 ""
    static ModeFromSettings(file) {
        try {
            general := JSON.Parse(FileRead(file, "UTF-8"))["General"]
            if (general["CapsLockElevated"] && general.Has("CapsLock"))
                return general["CapsLock"]
        }
        return ""
    }

    static _OnCopyData(wParam, lParam, msg, hwnd) {
        if (NumGet(lParam, 0, "UPtr") != App.CopyDataId)
            return
        text := NumGet(lParam, A_PtrSize, "UInt") ? StrGet(NumGet(lParam, A_PtrSize * 2, "Ptr"), "UTF-16") : ""
        if (Trim(text, "`n") = "-Exit")
            SetTimer(() => ExitApp(), -1)
        return 1
    }

    ; 大写锁定开 / 关的提示: 让 ALTRun 显示
    static _TellMain(on) {
        if (hwnd := App._RunningInstance())
            App.SendArgs(hwnd, ["-CapsLockState", on ? "1" : "0"])
    }

    ;---------------------------------------------------------------------------
    ; ALTRun 这边 (普通权限)
    ;---------------------------------------------------------------------------
    ; 打开了这个选项、CapsLock 设成了切换输入法、ALTRun 本身不是管理员时, 在后台运行计划任务启动小助手
    static Wanted(general) {
        return general.Has("CapsLockElevated") && general["CapsLockElevated"] && general.Has("CapsLock")
            && CapsLockSwitch.IsMode(general["CapsLock"])
    }

    static Start() {
        if (A_IsAdmin || !CapsLockHelper.Wanted(AppSettings.General) || CapsLockHelper.Window())
            return
        SetTimer(() => CapsLockHelper._RunTask(), -500)                     ; 不拖慢启动
    }

    static _RunTask() {
        code := -1
        try code := RunWait('schtasks.exe /run /tn "' CapsLockHelper.TaskName '"', , "Hide")
        Logger.Debug("CapsLockHelper: schtasks /run -> " code)
        if (code != 0)
            App.Toast(I18n.T("Caps.HelperFailed"), 8000)
    }

    ; ALTRun 退出 / 重新载入时让小助手也退出 (重新载入后按新的设置再启动)
    static Stop() {
        if (hwnd := CapsLockHelper.Window())
            try App.SendArgs(hwnd, ["-Exit"])
    }

    static Window() {
        saved := A_DetectHiddenWindows
        DetectHiddenWindows(true)
        hwnd := WinExist(CapsLockHelper.Title " ahk_class AutoHotkey")
        DetectHiddenWindows(saved)
        return hwnd
    }

    ; 建计划任务 (确认一次 UAC)。返回 true = 建好了; 取消 UAC 或失败时 false
    static Install() {
        if A_IsCompiled
            command := A_ScriptFullPath, arguments := CapsLockHelper.Arg
        else
            command := A_AhkPath, arguments := '"' A_ScriptFullPath '" ' CapsLockHelper.Arg
        xmlFile := A_Temp "\ALTRunCapsLockHelper.xml"
        try {
            try FileDelete(xmlFile)
            FileAppend(CapsLockHelper.TaskXml(command, arguments, A_ScriptDir, CapsLockHelper.UserId()), xmlFile, "UTF-16")
            code := RunWait('*RunAs schtasks.exe /create /tn "' CapsLockHelper.TaskName '" /xml "' xmlFile '" /f', , "Hide")
        } catch as e {
            Logger.Debug("CapsLockHelper: install failed - " e.Message)
            code := -1
        }
        try FileDelete(xmlFile)
        Logger.Debug("CapsLockHelper: schtasks /create -> " code)
        return code = 0
    }

    ; 删掉计划任务 (确认一次 UAC; 取消也没关系, 选项关了就不会再运行它)
    static Uninstall() {
        CapsLockHelper.Stop()
        try RunWait('*RunAs schtasks.exe /delete /tn "' CapsLockHelper.TaskName '" /f', , "Hide")
    }

    static UserId() {
        domain := EnvGet("USERDOMAIN")
        return (domain != "") ? domain "\" A_UserName : A_UserName
    }

    ; 计划任务的 XML: 最高权限、用户登录时才运行 (交互式, 能装键盘钩子)、只手动运行 (没有触发器)、
    ; 用电池时也运行、不限运行时长、正常优先级 (默认 7 是低优先级, 按键反应会慢)
    static TaskXml(command, arguments, workingDir, userId) {
        esc(s) => StrReplace(StrReplace(StrReplace(StrReplace(s, "&", "&amp;"), "<", "&lt;"), ">", "&gt;"), '"', "&quot;")
        return '<?xml version="1.0" encoding="UTF-16"?>`n'
            . '<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">`n'
            . '  <RegistrationInfo><Description>ALTRun: switch the input method with Caps Lock in windows running as administrator</Description></RegistrationInfo>`n'
            . '  <Principals><Principal id="Author"><UserId>' esc(userId) '</UserId><LogonType>InteractiveToken</LogonType><RunLevel>HighestAvailable</RunLevel></Principal></Principals>`n'
            . '  <Settings>`n'
            . '    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>`n'
            . '    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>`n'
            . '    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>`n'
            . '    <AllowStartOnDemand>true</AllowStartOnDemand>`n'
            . '    <AllowHardTerminate>true</AllowHardTerminate>`n'
            . '    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>`n'
            . '    <Priority>4</Priority>`n'
            . '    <Enabled>true</Enabled>`n'
            . '  </Settings>`n'
            . '  <Actions Context="Author"><Exec><Command>' esc(command) '</Command><Arguments>' esc(arguments) '</Arguments><WorkingDirectory>' esc(workingDir) '</WorkingDirectory></Exec></Actions>`n'
            . '</Task>`n'
    }
}
