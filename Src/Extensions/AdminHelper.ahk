;===============================================================================
; AdminHelper.ahk - a small helper with administrator rights for windows running as administrator (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Windows doesn't let programs with normal rights control windows with administrator rights (UIPI): with an administrator window in front, ALTRun's keyboard hook gets no keys,
; and switch messages sent to it are blocked (#158). Running all of ALTRun as administrator would make every program it opens elevated too
; (mapped network drives invisible, drag and drop refused). So a separate small helper runs as administrator and does only what needs those rights:
;   - when an option that needs it is on (Keyboard.CapsLockElevated, Preferences → Keyboard; Extensions.QuickSwitch.AdminDialogs, Preferences → Quick Switch),
;     a scheduled task with "highest privileges" is created
;     (UAC is confirmed once, only then); ALTRun runs this task at every start-up (schtasks /run, no UAC prompt) to start the helper, and makes it exit on exit.
;     One task and one helper for every such option; the task is deleted when the last of them is turned off
;   - the helper is ALTRun itself started with -AdminHelper: it creates no windows, writes no settings and runs no programs. Its main window's
;     title becomes Title, so it isn't taken for "an ALTRun already running"
;   - Caps Lock: ALTRun and the helper both register the CapsLock hotkey (CapsLockSwitch): keyboard hooks are called in the order they were installed,
;     and the first to get the key catches it so the other doesn't, so a normal window switches only once; in an administrator window only the helper gets it.
;     ALTRun shows the Caps Lock on / off notice (the helper sends "-CapsLockState 1/0" to it; an elevated program may send messages to a normal one freely)
;   - Quick Switch: ALTRun can't jump a dialog of a program running as administrator (the messages and keys are blocked, its process can't be
;     opened), so it hands "-Jump <dialog> <folder>" to the helper (Jump). Any program with normal rights can send the helper that message,
;     so the helper checks it itself before doing anything (CheckJump): the window is the foreground window, a #32770 open / save or Browse for
;     Folder dialog (by its controls), of a program running as administrator, and the folder is a full path to a folder that exists. It then only
;     jumps the dialog there (the folder with a trailing backslash, so it can't be taken for a file to open or save) and presses nothing else
;   - the helper only accepts "-Exit" and "-Jump" (WM_COPYDATA, the one message allowed from normal programs); it exits by itself when ALTRun can't be found
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
        Persistent(true)                                                    ; keeps running without a hotkey too (only dialog jumps)
        if (arg = AdminHelper.OldArg)                                       ; started by the task of an older version: move it to the new name
            SetTimer(() => AdminHelper.MoveTask(), -2000)                    ; after the hotkey is set up, which matters more
        DllCall("ChangeWindowMessageFilterEx", "Ptr", A_ScriptHwnd, "UInt", 0x4A, "UInt", 1, "Ptr", 0)   ; WM_COPYDATA, MSGFLT_ALLOW
        OnMessage(0x4A, (p*) => AdminHelper._OnCopyData(p*))
        mode := AdminHelper.ModeFromSettings(AppSettings.File)
        AdminHelper._dialogs := AdminHelper.DialogsFromSettings(AppSettings.File)
        if !(CapsLockSwitch.IsMode(mode) || AdminHelper._dialogs)
            ExitApp()
        if CapsLockSwitch.IsMode(mode) {
            CapsLockSwitch.Mode := mode
            CapsLockSwitch.Notify := (on) => AdminHelper._TellMain("-CapsLockState", on ? "1" : "0")
            CapsLockSwitch.Register((key, fn) => Hotkey(key, fn))
        }
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

    ; Extensions.QuickSwitch: the jump for dialogs of administrator programs is on (and Quick Switch itself); read only, like ModeFromSettings
    static DialogsFromSettings(file) {
        try {
            options := JSON.Parse(FileRead(file, "UTF-8"))["Extensions"]["QuickSwitch"]
            return AdminHelper.DialogsWanted(options)
        }
        return false
    }

    static _dialogs := false

    ; -Exit, or -Jump <dialog> <folder> (see CheckJump; the reply is its code). Read within the length the sender gave, at most 1100 characters
    static _OnCopyData(wParam, lParam, msg, hwnd) {
        if (NumGet(lParam, 0, "UPtr") != App.CopyDataId)
            return
        lines := StrSplit(Trim(Win.CopyDataText(lParam, 1100), "`n"), "`n")
        if (lines[1] = "-Exit")
            SetTimer(() => ExitApp(), -1)
        else if (lines[1] = "-Jump")
            return AdminHelper._JumpRequest(lines)
        return 1
    }

    ; Checked here, in the message, so ALTRun hears why it was refused; the jump itself runs right after, outside the message
    static _JumpRequest(lines) {
        if !AdminHelper._dialogs
            return AdminHelper.Refused["option off"]
        if (lines.Length != 3 || !RegExMatch(lines[2], "^\d{1,20}$"))
            return AdminHelper.Refused["bad request"]
        dialog := Integer(lines[2]), folder := lines[3]
        code := AdminHelper.CheckJump(folder, AdminHelper.JumpFacts(dialog, folder))
        if (code = 1)
            SetTimer(() => AdminHelper._Jump(dialog, folder), -1)
        return code
    }

    static _Jump(dialog, folder) {
        folder := RTrim(folder, "\") "\"
        ok := FileDialogs.Jump(folder, dialog)
        AdminHelper._TellMain("-AdminHelperNote", "jump to [" folder "] " (ok = false ? "failed" : "done"))
    }

    ; Reply codes of -Jump: 1 = jumping; the others say why not (ALTRun logs the name)
    static Refused := Map("not in front", 2, "not a dialog", 3, "not elevated", 4, "not a folder path", 5, "folder not found", 6, "option off", 7, "bad request", 8)

    static ReplyText(code) {
        if (code = 1)
            return "jumping"
        for name, value in AdminHelper.Refused
            if (value = code)
                return "refused: " name
        return "no answer"
    }

    ; What the helper sees of the window and the folder (facts: see CheckJump). Nothing more is read once one of them fails
    static JumpFacts(dialog, folder) {
        facts := {Front: DllCall("GetForegroundWindow", "Ptr") = dialog, Class: "", Kind: "", Elevated: false, Exists: false}
        if !facts.Front
            return facts
        target := "ahk_id " dialog, controls := []
        try facts.Class := WinGetClass(target)
        try controls := WinGetControls(target)
        if FileDialogs.LooksLikeBrowseDialog(facts.Class, controls, FileDialogs._TreeId(controls, target))
            facts.Kind := "browse"
        else if (facts.Class = "#32770" && controls.Length && FileDialogs.LooksLikeFileDialog(facts.Class, controls))
            facts.Kind := "file"
        try facts.Elevated := Win.IsElevated(WinGetPID(target))
        if (facts.Kind != "" && facts.Elevated && AdminHelper.IsFolderPath(folder))
            facts.Exists := DirExist(folder) != ""                          ; seen as administrator: no mapped network drives (the program wouldn't see them either)
        return facts
    }

    ; 1 if the helper may jump the dialog to the folder, otherwise a code of Refused. A program with normal rights may have sent the request,
    ; so it only goes through for the dialog in front, a system Open / Save or Browse for Folder dialog, of a program running as administrator
    ; (the case ALTRun can't handle itself), and for a folder that exists
    static CheckJump(folder, facts) {
        if !AdminHelper.IsFolderPath(folder)
            return AdminHelper.Refused["not a folder path"]
        if !facts.Front
            return AdminHelper.Refused["not in front"]
        if (facts.Class != "#32770" || facts.Kind = "")
            return AdminHelper.Refused["not a dialog"]
        if !facts.Elevated
            return AdminHelper.Refused["not elevated"]
        if !facts.Exists
            return AdminHelper.Refused["folder not found"]
        return 1
    }

    ; A full path: a drive (C:\...) or a network share (\\server\share\...); no wildcards, quotes, <>|, control characters or a second colon
    ; (a stream), nothing a dialog would take as a filter or a command
    static IsFolderPath(folder) {
        return StrLen(folder) <= 1024 && RegExMatch(folder, "^(?:[A-Za-z]:\\|\\\\[^\\/]+\\[^\\/]+)") > 0
            && !RegExMatch(folder, '[\x00-\x1F*?"<>|]') && !InStr(SubStr(folder, 3), ":")
    }

    ; To ALTRun: the Caps Lock on / off notice for it to show (-CapsLockState), a line for its log (-AdminHelperNote; the helper writes no files)
    static _TellMain(args*) {
        if (hwnd := App._RunningInstance())
            App.SendArgs(hwnd, args)
    }

    ;---------------------------------------------------------------------------
    ; ALTRun's side (normal rights)
    ;---------------------------------------------------------------------------
    ; The helper has something to do: Caps Lock in administrator windows (the option on, CapsLock set to switch the input language), or
    ; jumps in dialogs of administrator programs (the option on, Quick Switch on)
    static Wanted(keyboard, options := Map()) => AdminHelper.CapsLockWanted(keyboard) || AdminHelper.DialogsWanted(options)

    static CapsLockWanted(keyboard) {
        return keyboard.Has("CapsLockElevated") && keyboard["CapsLockElevated"] && keyboard.Has("CapsLock")
            && CapsLockSwitch.IsMode(keyboard["CapsLock"])
    }

    static DialogsWanted(options) {
        return options is Map && options.Has("AdminDialogs") && options["AdminDialogs"] && options.Has("Enabled") && options["Enabled"]
    }

    ; The scheduled task is kept while one of the options is on, even with its feature off for now (no UAC prompt each time it is turned back on)
    static TaskWanted(keyboard, options) {
        return AdminHelper.CapsLockWanted(keyboard) || (options is Map && options.Has("AdminDialogs") && options["AdminDialogs"])
    }

    ; With ALTRun itself not elevated, run the scheduled task in the background to start the helper
    static Start() {
        if (A_IsAdmin || !AdminHelper.Wanted(AppSettings.Keyboard, AppSettings.Extension("QuickSwitch")) || AdminHelper.Window())
            return
        SetTimer(() => AdminHelper._RunTask(), -500)                        ; don't slow down start-up
    }

    ; The task under its new name; the old one if that isn't there yet (the helper it starts moves it, see MoveTask). true = started
    static _RunTask() {
        for name in [AdminHelper.TaskName, AdminHelper.OldTaskName] {
            code := -1
            try code := RunWait('schtasks.exe /run /tn "' name '"', , "Hide")
            Logger.Debug("AdminHelper: schtasks /run " name " -> " code)
            if (code = 0)
                return true
        }
        App.Toast(I18n.T("AdminHelper.Failed"), 8000)
        return false
    }

    ; ALTRun's side of a jump in a dialog of a program running as administrator (FileDialogs.SetDialogPath): the helper does it. With the
    ; option off, a hint says where to turn it on (once per dialog). A helper that isn't running (ended, or never started) is started first.
    ; If the helper finds the program isn't elevated after all (ALTRun was only refused access to it), ALTRun jumps itself
    static Jump(folder, dialog) {
        if !AdminHelper.DialogsWanted(QuickSwitch.Options) {
            Logger.Debug("QuickSwitch: the dialog's program runs as administrator, and jumping there is off")
            return FileDialogs.ElevatedHint(dialog)
        }
        if !(helper := AdminHelper.Window()) {
            if !AdminHelper._RunTask()
                return false
            deadline := A_TickCount + 3000
            while (!(helper := AdminHelper.Window()) && A_TickCount < deadline)
                Sleep(100)
            if !helper
                return (Logger.Error("QuickSwitch: the helper for administrator windows didn't start"), false)
        }
        try WinWaitActive("ahk_id " dialog, , 0.5)                         ; the helper only jumps the dialog in front
        reply := 0
        App.SendArgs(helper, ["-Jump", dialog, folder], &reply)
        Logger.Debug("QuickSwitch: the jump to [" folder "] is handed to the helper for administrator windows - " AdminHelper.ReplyText(reply))
        if (reply = AdminHelper.Refused["not elevated"])
            return FileDialogs.Jump(folder, dialog)
        return reply = 1
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
            . '  <RegistrationInfo><Description>ALTRun: a small helper for windows running as administrator (Caps Lock switching, folder jumps in dialogs)</Description></RegistrationInfo>`n'
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
