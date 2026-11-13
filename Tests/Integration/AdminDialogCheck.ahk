;===============================================================================
; AdminDialogCheck.ahk - Quick Switch in a dialog of a program running as administrator, through the helper (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Run by Tests\RunTests.ahk AdminDialogCheck on Windows CI (.github/workflows/quickswitch.yml), as the administrator; it starts the
; ALTRun side as a standard user, whose programs run at medium integrity, so Windows blocks its messages to the dialog (UIPI) as on a
; normal PC. Three processes besides this one:
;   the dialog   an Open dialog at "Start" (FileSelect), run as administrator
;   the client   RunTests.ahk AdminDialogCheck client, as the standard user: the ALTRun side. Its main window takes the title of
;                ALTRun's (the helper exits when it can't find ALTRun)
;   the helper   ALTRun.ahk -AdminHelper from <appDir>, as administrator, with Extensions.QuickSwitch.AdminDialogs on
; Steps (the two sides wait for each other's files in <io>):
;   1. without the helper, the client jumps the dialog itself (FileDialogs.Jump): it must stay at Start (Windows blocks it), or the
;      check would prove nothing
;   2. with the helper: FileDialogs.SetDialogPath sees the dialog's program is elevated and hands the jump over: it must reach "Target";
;      the helper refuses a folder that doesn't exist and a window that isn't a dialog in front
; Prints one line per check (PASS / FAIL / INFO); returns the number of failed checks.
;
;   RunTests.ahk AdminDialogCheck admin <appDir> <testsDir> <user> <password>     appDir: a copy of the repository the standard user can read
;===============================================================================
class AdminDialogCheck {
    static Root := "C:\AdminDialogCheck", Failures := 0

    static Run(args) {
        if (args.Length >= 2 && args[2] = "client")
            return AdminDialogCheck.Client(args[3], Integer(args[4]))
        appDir := args[3], testsDir := args[4], user := args[5], password := args[6]
        demo := AdminDialogCheck.Root "\demo", io := AdminDialogCheck.Root "\io"
        dialogPid := 0, helperPid := 0, clientPid := 0
        try {
            DirCreate(demo "\Start"), DirCreate(demo "\Target"), DirCreate(io)
            RunWait(A_ComSpec ' /c icacls "' io '" /grant *S-1-1-0:(OI)(CI)M', , "Hide")   ; the client (another user) writes its answers here
            loop files io "\*"
                FileDelete(A_LoopFileFullPath)
            settings := Map("SchemaVersion", AppSettings.CurrentVersion, "General", Map("Language", "en", "SaveLog", 0)
                , "Extensions", Map("QuickSwitch", Map("Enabled", 1, "AdminDialogs", 1)))
            try DirDelete(appDir "\Data", true)
            DirCreate(appDir "\Data")
            FileAppend(JSON.Stringify(settings, 4), appDir "\Data\ALTRun.json", "UTF-8")

            dialogScript := AdminDialogCheck.Root "\dialog.ahk"
            try FileDelete(dialogScript)
            FileAppend('#NoTrayIcon`nFileSelect(, "' demo '\Start", "Open")`n', dialogScript, "UTF-8")
            Run('"' A_AhkPath '" "' dialogScript '"', , , &dialogPid)
            if !(dialog := WinWait("Open ahk_class #32770 ahk_pid " dialogPid, , 10))
                throw Error("the Open dialog didn't open")
            AdminDialogCheck.Log("INFO dialog " Format("0x{:X}", dialog) ", administrator " (A_IsAdmin ? "yes" : "no"))

            RunAs(user, password, A_ComputerName)
            try Run('"' A_AhkPath '" /ErrorStdOut "' testsDir '\RunTests.ahk" AdminDialogCheck client "' appDir '" ' dialog, testsDir, , &clientPid)
            finally RunAs()
            if !AdminDialogCheck.WaitFor(io "\ready.txt", 30)
                throw Error("the client didn't start as the standard user")
            AdminDialogCheck.Log("INFO client: " Trim(FileRead(io "\ready.txt"), "`r`n"))

            ; 1. without the helper
            AdminDialogCheck.Front(dialog), AdminDialogCheck.Signal(io "\go1.txt")
            AdminDialogCheck.WaitFor(io "\done1.txt", 15)
            Sleep(1500)
            folder := FileDialogs.DialogFolder(dialog)
            AdminDialogCheck.Check("without the helper Windows blocks the jump", folder ~= "\\Start$", "the dialog shows [" folder "]; "
                . "without UIPI on this runner the check proves nothing")

            ; 2. with the helper
            Run('"' A_AhkPath '" "' appDir '\ALTRun.ahk" -AdminHelper', appDir, , &helperPid)
            deadline := A_TickCount + 10000
            while (!AdminHelper.Window() && A_TickCount < deadline)
                Sleep(100)
            AdminDialogCheck.Check("the helper started", AdminHelper.Window() != 0)
            AdminDialogCheck.Front(dialog), AdminDialogCheck.Signal(io "\go2.txt")
            AdminDialogCheck.WaitFor(io "\done2.txt", 20)
            deadline := A_TickCount + 5000
            while (!((folder := FileDialogs.DialogFolder(dialog)) ~= "\\Target$") && A_TickCount < deadline)
                Sleep(100)
            answers := FileExist(io "\done2.txt") ? FileRead(io "\done2.txt") : ""
            AdminDialogCheck.Check("handed over: the program counts as elevated", InStr(answers, "elevated 1") > 0, answers)
            AdminDialogCheck.Check("with the helper the dialog jumps", folder ~= "\\Target$", "the dialog shows [" folder "]; client: " answers)
            AdminDialogCheck.Check("refused: a folder that doesn't exist", InStr(answers, "missing refused: folder not found") > 0, answers)
            AdminDialogCheck.Check("refused: a window that isn't a dialog in front", InStr(answers, "own window refused: not in front") > 0, answers)
        } catch as e
            AdminDialogCheck.Check("no error", false, e.Message " (line " e.Line ")")
        try AdminHelper.Stop()
        for pid in [clientPid, dialogPid, helperPid]
            if pid
                try ProcessClose(pid)
        return AdminDialogCheck.Failures
    }

    ; The ALTRun side, as the standard user
    static Client(appDir, dialog) {
        io := AdminDialogCheck.Root "\io", target := AdminDialogCheck.Root "\demo\Target\"
        DllCall("SetWindowTextW", "Ptr", A_ScriptHwnd, "Str", appDir "\ALTRun.ahk - AutoHotkey v2 (AdminDialogCheck)")   ; App._RunningInstance, for the helper
        QuickSwitch.Options := Map("Enabled", 1, "AdminDialogs", 1, "BrowseDialogs", 1)
        AdminDialogCheck.Signal(io "\ready.txt", "user " A_UserName ", administrator " (A_IsAdmin ? "yes" : "no") ", dialog found " (WinExist("ahk_id " dialog) ? "yes" : "no"))
        if !AdminDialogCheck.WaitFor(io "\go1.txt", 30)
            return 1
        try FileDialogs.Jump(target, dialog)
        AdminDialogCheck.Signal(io "\done1.txt")
        if !AdminDialogCheck.WaitFor(io "\go2.txt", 30)
            return 1
        answers := "elevated " (FileDialogs.IsElevatedDialog(dialog) ? 1 : 0) "`n"
        try answers .= "jump " (FileDialogs.SetDialogPath(target, dialog) ? 1 : 0) "`n"
        catch as e
            answers .= "jump error " e.Message "`n"
        Sleep(1000)
        if (helper := AdminHelper.Window()) {
            reply := 0
            App.SendArgs(helper, ["-Jump", dialog, AdminDialogCheck.Root "\demo\Missing\"], &reply)
            answers .= "missing " AdminHelper.ReplyText(reply) "`n"
            App.SendArgs(helper, ["-Jump", A_ScriptHwnd, target], &reply)
            answers .= "own window " AdminHelper.ReplyText(reply) "`n"
        } else
            answers .= "no helper window`n"
        AdminDialogCheck.Signal(io "\done2.txt", answers)
        Sleep(30000)                                                        ; stays the "ALTRun" the helper looks for, until closed
        return 0
    }

    static Front(dialog) {
        try WinActivate("ahk_id " dialog)
        try WinWaitActive("ahk_id " dialog, , 2)
    }

    static Signal(file, text := "") {
        FileAppend(text, file ".part", "UTF-8")
        FileMove(file ".part", file, true)
    }

    static WaitFor(file, seconds) {
        deadline := A_TickCount + seconds * 1000
        while (!FileExist(file) && A_TickCount < deadline)
            Sleep(100)
        return FileExist(file) != ""
    }

    static Check(name, ok, detail := "") {
        if !ok
            AdminDialogCheck.Failures++
        AdminDialogCheck.Log((ok ? "PASS " : "FAIL ") name (ok || detail = "" ? "" : " - " detail))
    }

    static Log(text) => FileAppend(text "`n", "*")
}
