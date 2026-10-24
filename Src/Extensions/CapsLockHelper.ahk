;===============================================================================
; CapsLockHelper.ahk - CapsLock switches the input language in administrator windows too (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Windows doesn't let programs with normal rights control windows with administrator rights (UIPI): with an administrator window in front, ALTRun's keyboard hook gets no keys,
; and switch messages sent to it are blocked (#158). Running all of ALTRun as administrator would make every program it opens elevated too
; (mapped network drives invisible, drag and drop refused). So a separate small helper runs as administrator and only does the CapsLock switch:
;   - with General.CapsLockElevated = 1 (Preferences → Custom hotkeys), a scheduled task with "highest privileges" is created (UAC is confirmed once,
;     only then); ALTRun runs this task at every start-up (schtasks /run, no UAC prompt) to start the helper, and makes it exit on exit
;   - the helper is ALTRun itself started with -CapsLockHelper: it only registers the CapsLock hotkey (CapsLockSwitch), creates no windows,
;     writes no settings and runs no programs. Its main window's title becomes Title, so it isn't taken for "an ALTRun already running"
;   - ALTRun and the helper both register the CapsLock hotkey: keyboard hooks are called in the order they were installed, and the first to get the key catches it so the other doesn't,
;     so a normal window switches only once; in an administrator window only the helper gets it
;   - ALTRun shows the Caps Lock on / off notice (the helper sends "-CapsLockState 1/0" to it; an elevated program may send messages to a normal one freely)
;   - the helper only accepts "-Exit" (WM_COPYDATA, the one message allowed from normal programs); it exits by itself when ALTRun can't be found
;===============================================================================

class CapsLockHelper {
    static Arg      := "-CapsLockHelper"
    static TaskName := "ALTRun CapsLock Helper"
    static Title    := "ALTRun CapsLock Helper"

    ;---------------------------------------------------------------------------
    ; The helper's side (the process running as administrator)
    ;---------------------------------------------------------------------------
    static Run() {
        A_IconHidden := true
        DetectHiddenWindows(true)
        for hwnd in WinGetList(CapsLockHelper.Title " ahk_class AutoHotkey")   ; one is already running
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
        SetTimer(() => App._RunningInstance() || ExitApp(), 3000)           ; ALTRun is gone (ended, crashed): exit too
    }

    ; General.CapsLock from the settings file (read only, never written back); "" if it can't be read
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

    ; The Caps Lock on / off notice: let ALTRun show it
    static _TellMain(on) {
        if (hwnd := App._RunningInstance())
            App.SendArgs(hwnd, ["-CapsLockState", on ? "1" : "0"])
    }

    ;---------------------------------------------------------------------------
    ; ALTRun's side (normal rights)
    ;---------------------------------------------------------------------------
    ; With the option on, CapsLock set to switch the input language and ALTRun itself not elevated, run the scheduled task in the background to start the helper
    static Wanted(general) {
        return general.Has("CapsLockElevated") && general["CapsLockElevated"] && general.Has("CapsLock")
            && CapsLockSwitch.IsMode(general["CapsLock"])
    }

    static Start() {
        if (A_IsAdmin || !CapsLockHelper.Wanted(AppSettings.General) || CapsLockHelper.Window())
            return
        SetTimer(() => CapsLockHelper._RunTask(), -500)                     ; don't slow down start-up
    }

    static _RunTask() {
        code := -1
        try code := RunWait('schtasks.exe /run /tn "' CapsLockHelper.TaskName '"', , "Hide")
        Logger.Debug("CapsLockHelper: schtasks /run -> " code)
        if (code != 0)
            App.Toast(I18n.T("Caps.HelperFailed"), 8000)
    }

    ; When ALTRun exits / reloads, the helper exits too (restarted with the new settings after reloading)
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

    ; Creates the scheduled task (UAC confirmed once). true = created; false when UAC is canceled or it fails
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

    ; Deletes the scheduled task (UAC confirmed once; canceling is fine, it won't run anymore with the option off)
    static Uninstall() {
        CapsLockHelper.Stop()
        try RunWait('*RunAs schtasks.exe /delete /tn "' CapsLockHelper.TaskName '" /f', , "Hide")
    }

    static UserId() {
        domain := EnvGet("USERDOMAIN")
        return (domain != "") ? domain "\" A_UserName : A_UserName
    }

    ; The scheduled task's XML: highest privileges, runs only while the user is logged on (interactive, so it can install a keyboard hook), only run manually (no triggers),
    ; runs on battery too, no run time limit, normal priority (the default 7 is low priority, which makes keys respond slowly)
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
