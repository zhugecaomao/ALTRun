;===============================================================================
; AdminHelper.ahk - a small helper with administrator rights for windows running as administrator (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Windows doesn't let programs with normal rights control windows with administrator rights (UIPI): with an administrator window in front, ALTRun's keyboard hook gets no keys,
; and switch messages sent to it are blocked (#158). Running all of ALTRun as administrator would make every program it opens elevated too
; (mapped network drives invisible, drag and drop refused). So a separate small helper runs as administrator and does only what needs those rights:
;   - when an option that needs it is on (Keyboard.CapsLockElevated, Preferences → Keyboard), a scheduled task with "highest privileges" is created
;     (UAC is confirmed once, only then); ALTRun runs this task at every start-up (schtasks /run, no UAC prompt) to start the helper, and makes it exit on exit.
;     One task and one helper for every such option; the task is deleted when the last of them is turned off
;   - the helper is ALTRun itself started with -AdminHelper: it creates no windows, writes no settings and runs no programs. Its main window's
;     title becomes Title, so it isn't taken for "an ALTRun already running"
;   - Caps Lock: ALTRun and the helper both register the CapsLock hotkey (CapsLockSwitch): keyboard hooks are called in the order they were installed,
;     and the first to get the key catches it so the other doesn't, so a normal window switches only once; in an administrator window only the helper gets it.
;     ALTRun shows the Caps Lock on / off notice (the helper sends "-CapsLockState 1/0" to it; an elevated program may send messages to a normal one freely)
;   - the helper only accepts "-Exit" (WM_COPYDATA, the one message allowed from normal programs); it exits by itself when ALTRun can't be found
;   - older versions named the task "ALTRun CapsLock Helper" and started the helper with -AdminHelper. A helper started that way moves the task to
;     the new name itself (it has the rights, so no UAC prompt): it registers TaskName and deletes OldTaskName. Until then ALTRun runs the old one
;===============================================================================

class AdminHelper {
    static Arg         := "-AdminHelper"
    static TaskName    := "ALTRun Admin Helper"
    static Title       := "ALTRun Admin Helper"
    static OldArg      := "-CapsLockHelper"                                 ; versions up to 2026.11.13: the task and the helper had these names
    static OldTaskName := "ALTRun CapsLock Helper"
    static OldTitle    := "ALTRun CapsLock Helper"

    ; ALTRun was started as the helper (by the task, new or old)
    static IsHelperArg(arg) => arg = AdminHelper.Arg || arg = AdminHelper.OldArg

    ;---------------------------------------------------------------------------
    ; The helper's side (the process running as administrator)
    ;---------------------------------------------------------------------------
    static Run(arg := AdminHelper.Arg) {
        A_IconHidden := true
        Logger.Enabled := false                                             ; a log file written as administrator: lines go to ALTRun instead (_TellMain)
        if AdminHelper.Window(A_ScriptHwnd)                                 ; one is already running
            ExitApp()
        DllCall("SetWindowTextW", "Ptr", A_ScriptHwnd, "Str", AdminHelper.Title)
        if (arg = AdminHelper.OldArg)                                       ; started by the task of an older version: move it to the new name
            SetTimer(() => AdminHelper.MoveTask(), -2000)                    ; after the hotkey is set up, which matters more
        DllCall("ChangeWindowMessageFilterEx", "Ptr", A_ScriptHwnd, "UInt", 0x4A, "UInt", 1, "Ptr", 0)   ; WM_COPYDATA, MSGFLT_ALLOW
        OnMessage(0x4A, (p*) => AdminHelper._OnCopyData(p*))
        mode := AdminHelper.ModeFromSettings(AppSettings.File)
        if !CapsLockSwitch.IsMode(mode)
            ExitApp()
        CapsLockSwitch.Mode := mode
        CapsLockSwitch.Notify := (on) => AdminHelper._TellMain("-CapsLockState", on ? "1" : "0")
        CapsLockSwitch.Register((key, fn) => Hotkey(key, fn))
        WinTimer.Set(() => App._RunningInstance() || ExitApp(), 3000)          ; ALTRun is gone (ended, crashed): exit too
    }

    ; Keyboard.CapsLock from the settings file (read only, never written back); "" if it can't be read or the option is off.
    ; A file not yet upgraded to settings version 7 (ALTRun was updated while this helper ran) has them in General
    static ModeFromSettings(file) {
        try {
            data := JSON.Parse(FileRead(file, "UTF-8"))
            keyboard := (data.Has("Keyboard") && data["Keyboard"] is Map && data["Keyboard"].Has("CapsLockElevated")) ? data["Keyboard"] : data["General"]
            if (keyboard["CapsLockElevated"] && keyboard.Has("CapsLock"))
                return keyboard["CapsLock"]
        }
        return ""
    }

    static _OnCopyData(wParam, lParam, msg, hwnd) {
        if (NumGet(lParam, 0, "UPtr") != App.CopyDataId)
            return
        if (Trim(Win.CopyDataText(lParam, 64), "`n") = "-Exit")
            SetTimer(() => ExitApp(), -1)
        return 1
    }

    ; To ALTRun: the Caps Lock on / off notice for it to show (-CapsLockState), a line for its log (-AdminHelperNote; the helper writes no files)
    static _TellMain(args*) {
        if (hwnd := App._RunningInstance())
            App.SendArgs(hwnd, args)
    }

    ;---------------------------------------------------------------------------
    ; ALTRun's side (normal rights)
    ;---------------------------------------------------------------------------
    ; With the option on, CapsLock set to switch the input language and ALTRun itself not elevated, run the scheduled task in the background to start the helper
    static Wanted(keyboard) {
        return keyboard.Has("CapsLockElevated") && keyboard["CapsLockElevated"] && keyboard.Has("CapsLock")
            && CapsLockSwitch.IsMode(keyboard["CapsLock"])
    }

    static Start() {
        if (A_IsAdmin || !AdminHelper.Wanted(AppSettings.Keyboard) || AdminHelper.Window())
            return
        SetTimer(() => AdminHelper._RunTask(), -500)                        ; don't slow down start-up
    }

    ; The task under its new name; the old one if that isn't there yet (the helper it starts moves it, see MoveTask)
    static _RunTask() {
        for name in [AdminHelper.TaskName, AdminHelper.OldTaskName] {
            code := -1
            try code := RunWait('schtasks.exe /run /tn "' name '"', , "Hide")
            Logger.Debug("AdminHelper: schtasks /run " name " -> " code)
            if (code = 0)
                return
        }
        App.Toast(I18n.T("Caps.HelperFailed"), 8000)
    }

    ; When ALTRun exits / reloads, the helper exits too (restarted with the new settings after reloading)
    static Stop() {
        if (hwnd := AdminHelper.Window())
            try App.SendArgs(hwnd, ["-Exit"])
    }

    ; The helper's main window (also one started by an older version, under its old title), 0 if none; except: one to leave out (the helper itself)
    static Window(except := 0) {
        saved := A_DetectHiddenWindows
        DetectHiddenWindows(true)
        found := 0
        for title in [AdminHelper.Title, AdminHelper.OldTitle]
            for hwnd in WinGetList(title " ahk_class AutoHotkey")
                if (hwnd != except && !found)
                    found := hwnd
        DetectHiddenWindows(saved)
        return found
    }

    ; Creates the scheduled task (UAC confirmed once). true = created; false when UAC is canceled or it fails
    static Install() {
        code := AdminHelper.RegisterXml(AdminHelper.TaskName, AdminHelper.OwnTaskXml())
        Logger.Debug("AdminHelper: schtasks /create -> " code)
        return code = 0
    }

    ; The task that starts this ALTRun (the exe, or the interpreter with this script) as the helper
    static OwnTaskXml() {
        if A_IsCompiled
            command := App.ExePath, arguments := AdminHelper.Arg
        else
            command := A_AhkPath, arguments := '"' A_ScriptFullPath '" ' AdminHelper.Arg
        return AdminHelper.TaskXml(command, arguments, App.ExeDir, AdminHelper.UserId())
    }

    ; In the helper started by the old task (so already with administrator rights, no UAC prompt): the same task under the new name, starting
    ; the same program, then the old one deleted. The helper runs from the old task's program, so the new task starts the program the user
    ; once agreed to. Deleting the task doesn't end the helper it started (checked on Windows CI). If creating fails, the old task stays and is used again
    static MoveTask() {
        code := AdminHelper.RegisterXml(AdminHelper.TaskName, AdminHelper.OwnTaskXml(), false)
        deleted := -1
        if (code = 0)
            try deleted := RunWait('schtasks.exe /delete /tn "' AdminHelper.OldTaskName '" /f', , "Hide")
        AdminHelper._TellMain("-AdminHelperNote", "moved the task to its new name: create " code ", delete the old one " deleted)
    }

    ; Registers a task from its XML with schtasks /create run as administrator (UAC confirmed once; elevate = false: in the helper, which has
    ; the rights already); schtasks' exit code, -1 when it didn't run.
    ; schtasks takes the XML only as a file, in %TEMP%, where any program with normal rights can write: rewritten while the UAC prompt is up,
    ; it would put that program's command into a task with highest privileges. So the file has a name nobody can guess, and from the moment
    ; it is written until schtasks is done it is held open with a handle that lets others read it but not change, rename or delete it;
    ; read back through that handle, it must be the XML written. Checked against schtasks on Windows by Tests\Integration\DelayedStartCheck.ahk
    static RegisterXml(taskName, xml, elevate := true) {
        xmlFile := A_Temp "\ALTRun-task-" Format("{:08x}{:08x}", Random(0, 0xFFFFFFFF), Random(0, 0xFFFFFFFF)) ".xml"
        held := ""
        try {
            FileAppend(xml, xmlFile, "UTF-16")
            held := FileOpen(xmlFile, "r-wd", "UTF-16")
            if (held.Read() !== xml)
                throw Error("the task file was changed before it could be locked")
            code := RunWait((elevate ? '*RunAs ' : '') 'schtasks.exe /create /tn "' taskName '" /xml "' xmlFile '" /f', , "Hide")
        } catch as e {
            Logger.Debug("AdminHelper: registering " taskName " failed - " e.Message)
            code := -1
        }
        if IsObject(held)
            held.Close()
        try FileDelete(xmlFile)
        return code
    }

    ; A folder holding what the task runs as administrator that programs with normal rights can change, "" if none: such a program could
    ; replace ALTRun there and start the task (schtasks /run needs no confirmation), getting administrator rights. Running as administrator
    ; every folder looks writable, so nothing is checked then; nor for the Store package, whose files are protected (the task starts it
    ; through its execution alias)
    static WritableFolder(isWritable := (dir) => Path.IsWritable(dir), admin := A_IsAdmin) {
        if (App.PackageFamily != "" || admin)
            return ""
        SplitPath(A_AhkPath, , &ahkDir)
        for dir in (A_IsCompiled ? [App.ExeDir] : [A_ScriptDir, ahkDir])
            if isWritable(dir)
                return dir
        return ""
    }

    ; Deletes the scheduled task, under both names (one UAC prompt; canceling is fine, it won't run anymore with the option off)
    static Uninstall() {
        AdminHelper.Stop()
        try RunWait('*RunAs ' A_ComSpec ' /c ' AdminHelper.DeleteCommand(), , "Hide")
    }

    ; Both names in one command line, so one UAC prompt is enough (fixed names, nothing the user typed)
    static DeleteCommand() => 'schtasks.exe /delete /tn "' AdminHelper.TaskName '" /f & schtasks.exe /delete /tn "' AdminHelper.OldTaskName '" /f'

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
            . '  <RegistrationInfo><Description>ALTRun: a small helper for windows running as administrator (Caps Lock switching)</Description></RegistrationInfo>`n'
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
