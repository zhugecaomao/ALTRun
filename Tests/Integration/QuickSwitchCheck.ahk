;===============================================================================
; QuickSwitchCheck.ahk - automatic switching (Extensions.QuickSwitch.AutoSwitch) on a real Windows desktop (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Run by .github/workflows/quickswitch.yml; also works on your own PC (it takes over the keyboard and the screen for a minute
; and starts its own copy of ALTRun from <appDir> with settings of its own). Opens Explorer at "Folder A" and an Open dialog at
; "Start", then switches from Explorer to the dialog in several ways and checks that the dialog jumps to Folder A:
;   activate      SetForegroundWindow (WinActivate), as a program bringing the dialog to the front
;   alt+tab       a simulated Alt+Tab (SendInput), as the keyboard does it
;   title click   a click on the dialog's title bar, as the mouse does it
; Then the folder panel's ☆ menu: opened with a click on the star, or with Alt+↓ in the panel's search box, and closed with Esc:
; the dialog must be in front again within 2 s (so typing goes on in the dialog); a failure prints the windows in front before
; and after Esc.
; Then Explorer with a second tab at "Folder B", shown (#521; Windows 11 22H2 and later; opened with Ctrl+T as a user does):
; switching to the dialog must jump to Folder B, not to Folder A of the first tab, and so must the jump hotkey (Ctrl+G, sent at
; send level 1 like a person's keys); the tab is closed again (Ctrl+W) after each round.
; Then AutoSwitchOnOpen: with a window of this check in front (not a file manager) and Explorer behind it at Folder A, a new Open
; dialog must jump to Folder A as it opens (a switch from Explorer isn't involved, so AutoSwitch alone wouldn't jump).
; Then the extract windows of 7-Zip and WinRAR (#609; each skipped when not installed, the workflow installs WinRAR): opened after
; Explorer, they must not jump by themselves (as they open, or on a switch from Explorer), and Ctrl+G must put Folder A into the
; destination box and nothing more (the window stays open, nothing is extracted).
; Each way is tried Rounds times; the dialog is sent back to "Start" in between. Prints one line per round (PASS / FAIL) and
; ALTRun's QuickSwitch log lines on failure; the exit code is the number of failed rounds.
;
;   AutoHotkey64.exe QuickSwitchCheck.ahk <appDir> [rounds]       appDir: a copy of the repository (ALTRun.ahk, Lib, Src, Resources)
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut
#Include ..\..\Lib\JSON.ahk
#Include ..\..\Lib\ExplorerTabs.ahk

global AppDir := A_Args.Length ? A_Args[1] : A_ScriptDir "\..\.."
global Rounds := A_Args.Length > 1 ? Integer(A_Args[2]) : 5
global Demo := "C:\QuickSwitchCheck\demo", LogFile := A_Temp "\ALTRun.log"
global Failures := 0, AppPid := 0, AppWindow := 0, DialogPid := 0, Explorer := 0, Dialog := 0
SetTitleMatchMode(3)
SetKeyDelay(30, 30)

try {
    Setup()
    for method in ["activate", "alt+tab", "title click"]
        Loop Rounds
            Round_(method, A_Index)
    if ExplorerTabs.Supported()
        Loop Rounds
            TabRound(A_Index)
    else
        Say("SKIP tab shown: no Explorer tabs on Windows " A_OSVersion)
    Loop Rounds
        OpenRound(A_Index)
    for how in ["click", "alt+down"]
        Loop Rounds
            MenuRound(how, A_Index)
    for program in ["7-Zip", "WinRAR"]
        ExtractRound(program)
} catch as setupError {
    Say("FAIL setup: " setupError.Message " (line " setupError.Line ")")
    Failures++
}
TearDown()
Say(Failures ? Failures " round(s) failed" : "all rounds passed")
ExitApp(Failures)

Say(text) => FileAppend(text "`n", "*")

Setup() {
    global AppPid, AppWindow, DialogPid, Explorer, Dialog
    try DirDelete(Demo, true)
    DirCreate(Demo "\Folder A"), DirCreate(Demo "\Folder B"), DirCreate(Demo "\Start")
    settings := Map("SchemaVersion", 6
        , "General", Map("Language", "en", "LaunchAtLogin", 0, "CheckForUpdates", 0, "ShowTips", 0, "SaveLog", 1, "SendToMenu", 0, "StartMenuShortcut", 0)
        , "Extensions", Map("QuickSwitch", Map("AutoSwitch", 1, "AutoSwitchOnOpen", 1)))
    try DirDelete(AppDir "\Data", true)
    DirCreate(AppDir "\Data")
    FileAppend(JSON.Stringify(settings, 4), AppDir "\Data\ALTRun.json", "UTF-8")
    try FileDelete(LogFile)

    Run('"' A_AhkPath '" "' AppDir '\ALTRun.ahk"', AppDir, , &AppPid)
    DetectHiddenWindows(true)
    SetTitleMatchMode(2)
    AppWindow := WinWait("ALTRun.ahk ahk_class AutoHotkey ahk_pid " AppPid, , 30)
    SetTitleMatchMode(3)
    DetectHiddenWindows(false)
    if !AppWindow
        throw Error("ALTRun didn't start")
    Sleep(5000)                                                             ; start-up done

    Run('explorer.exe "' Demo '\Folder A"')
    SetTitleMatchMode(2)                                                    ; "Folder A" or "Folder A - File Explorer"
    Explorer := WinWait("Folder A ahk_class CabinetWClass", , 30)           ; the first Explorer window on a fresh runner takes a while
    SetTitleMatchMode(3)
    if !Explorer {
        Windows()
        throw Error("Explorer didn't open Folder A")
    }
    helper := A_Temp "\QuickSwitchCheckDialog.ahk"
    try FileDelete(helper)
    FileAppend('#NoTrayIcon`nFileSelect(, "' Demo '\Start", "Open")`n', helper, "UTF-8")
    Run('"' A_AhkPath '" "' helper '"', , , &DialogPid)
    Dialog := WinWait("Open ahk_class #32770 ahk_pid " DialogPid, , 10)
    if !Dialog {
        Windows()
        throw Error("the Open dialog didn't open")
    }
    Say("INFO ALTRun pid " AppPid ", Explorer " Format("0x{:X}", Explorer) ", dialog " Format("0x{:X}", Dialog))
    ; With AutoSwitchOnOpen the new dialog jumps to Folder A by itself. Wait for that jump before the rounds: ALTRun's first look at
    ; Explorer's folder took 0.7 to 1 s on a fresh runner, and a switch made meanwhile (round 1) wasn't seen as one from Explorer
    start := A_TickCount
    while (A_TickCount - start < 10000 && !ShowsFolder("Folder A"))
        Sleep(20)
    if !ShowsFolder("Folder A")
        return Fail("setup", 1, "the new dialog didn't jump to Folder A as it opened (AutoSwitchOnOpen)")
    Say("PASS setup: the new dialog jumped to Folder A as it opened in " (A_TickCount - start) " ms")
}

; One switch from Explorer to the dialog; PASS when the dialog shows Folder A within 5 s
Round_(method, n) {
    global Failures
    if !BackToStart()
        return Fail(method, n, "couldn't send the dialog back to Start")
    WinActivate(Explorer)
    if !WinWaitActive(Explorer, , 3)
        return Fail(method, n, "Explorer didn't come to the front")
    Sleep(400)                                                              ; ALTRun notes Explorer as the window in front
    logSize := LogLength()
    start := A_TickCount
    switch method {
        case "activate":
            WinActivate(Dialog)
        case "alt+tab":
            SendInput("{Alt down}"), Sleep(80), SendInput("{Tab}"), Sleep(80), SendInput("{Alt up}")
        case "title click":
            WinGetPos(&x, &y, &w, , Dialog)
            CoordMode("Mouse", "Screen")
            Click(x + w // 2 - 60, y + 12)
    }
    if !WinWaitActive(Dialog, , 3)
        return Fail(method, n, "the dialog didn't come to the front (in front: " WinGetClass("A") ")", logSize)
    while (A_TickCount - start < 5000 && !ShowsFolder("Folder A"))
        Sleep(20)
    if !ShowsFolder("Folder A")
        return Fail(method, n, "the dialog didn't jump to Folder A", logSize)
    Say("PASS " method " " n ": jumped in " (A_TickCount - start) " ms")
}

; Explorer with a second tab at Folder B, shown (#521): switching to the dialog jumps to Folder B, not to Folder A of the first
; tab; then the jump hotkey (Ctrl+G) in the dialog does too. The tab is closed again
TabRound(n) {
    name := "tab shown"
    if !BackToStart()
        return TabFail(name, n, "couldn't send the dialog back to Start")
    folderB := Demo "\Folder B"
    ; As a user does it, with Explorer's own keys (this script makes no COM call into Explorer, which could change what is tested):
    ; Ctrl+T a new tab, Ctrl+L the address bar, the folder, Enter
    WinActivate(Explorer)
    if !WinWaitActive(Explorer, , 3)
        return TabFail(name, n, "Explorer didn't come to the front")
    Send("^t")
    Sleep(1000)
    Send("^l")
    Sleep(300)
    SendText(folderB)
    Send("{Enter}")
    try {
        SetTitleMatchMode(2)
        shown := WinWait("Folder B ahk_id " Explorer, , 5)                  ; the title is the name of the tab shown
        SetTitleMatchMode(3)
        if !shown
            return TabFail(name, n, "Explorer doesn't show Folder B (title '" WinGetTitle(Explorer) "')")
        WinActivate(Explorer)
        if !WinWaitActive(Explorer, , 3)
            return TabFail(name, n, "Explorer didn't come to the front")
        Sleep(400)                                                          ; ALTRun notes Explorer as the window in front
        logSize := LogLength()
        start := A_TickCount
        WinActivate(Dialog)
        if !WinWaitActive(Dialog, , 3)
            return TabFail(name, n, "the dialog didn't come to the front", logSize)
        while (A_TickCount - start < 5000 && !ShowsFolder("Folder B"))
            Sleep(20)
        if !ShowsFolder("Folder B")                                         ; Ctrl+G is tried all the same
            TabFail(name, n, "the dialog didn't jump to Folder B of the tab shown" (ShowsFolder("Folder A") ? " (it went to Folder A of the first tab)" : ""), logSize)
        else
            Say("PASS " name " " n ": jumped to the tab shown in " (A_TickCount - start) " ms")

        if !BackToStart()                                                   ; Explorer stays the file manager used last
            return TabFail(name " ctrl+g", n, "couldn't send the dialog back to Start")
        logSize := LogLength()
        start := A_TickCount
        SendLevel(1)                                                        ; hook hotkeys ignore other scripts' keys at send level 0
        SendEvent("^g")
        SendLevel(0)
        while (A_TickCount - start < 3000 && !ShowsFolder("Folder B"))
            Sleep(20)
        if !ShowsFolder("Folder B")
            return TabFail(name " ctrl+g", n, "Ctrl+G didn't jump to Folder B of the tab shown" (ShowsFolder("Folder A") ? " (it went to Folder A of the first tab)" : ""), logSize)
        Say("PASS " name " ctrl+g " n ": jumped to the tab shown in " (A_TickCount - start) " ms")
    } finally {
        WinActivate(Explorer)                                               ; close the Folder B tab (Ctrl+W): Explorer shows Folder A again
        if WinWaitActive(Explorer, , 3) {
            SetTitleMatchMode(2)
            if WinActive("Folder B ahk_id " Explorer)
                Send("^w")
            SetTitleMatchMode(3)
        }
        Sleep(800)
    }
}

; A tab round's failure with what it needs to be told apart: every new line of ALTRun's log (an error isn't a QuickSwitch line),
; the dialog's address, Explorer's title and tabs, and the windows front to back
TabFail(name, n, why, logSize := -1) {
    Fail(name, n, why)
    if (logSize >= 0)
        for line in StrSplit(Trim(LogSince(logSize), "`r`n"), "`n", "`r")
            Say("     log: " line)
    address := ""
    try {
        for ctl in WinGetControls(Dialog)
            if (InStr(ctl, "ToolbarWindow32") = 1 && InStr(ControlGetText(ctl, Dialog), ":"))
                address .= (address = "" ? "" : " | ") ControlGetText(ctl, Dialog)
    }
    Say("     dialog: " address)
    tabs := ""
    for each in ExplorerTabs._Tabs(Explorer)
        tabs .= (tabs = "" ? "" : " | ") ExplorerTabs._Path(Explorer, each.Tab)
    shown := ""
    try shown := ExplorerTabs.CurrentItem(Explorer).Document.Folder.Self.Path
    Say("     Explorer: '" WinGetTitle(Explorer) "', tabs " tabs ", shown " shown)
    answered := DllCall("SendMessageTimeoutW", "Ptr", AppWindow, "UInt", 0, "Ptr", 0, "Ptr", 0, "UInt", 2, "UInt", 2000, "Ptr*", 0, "Ptr")   ; WM_NULL, SMTO_ABORTIFHUNG
    Say("     ALTRun running: " (ProcessExist(AppPid) ? "yes" : "no") ", answers messages: " (answered ? "yes" : "no (busy or hung)")
        . ", hung: " (DllCall("IsHungAppWindow", "Ptr", AppWindow) ? "yes" : "no"))
    lines := StrSplit(Trim(LogSince(0), "`r`n"), "`n", "`r")                ; where ALTRun got to: the end of its log
    Loop Min(25, lines.Length)
        Say("     log end: " lines[lines.Length - Min(25, lines.Length) + A_Index])
    for hwnd in WinGetList()
        try Say("     window: " Describe(hwnd) " | " WinGetTitle(hwnd))
}

; AutoSwitchOnOpen: Explorer at Folder A behind a window of this check (not a file manager), then a new Open dialog opens at
; Start: it jumps to Folder A within 5 s, as it opens. The dialog is closed again
OpenRound(n) {
    static first := "", second := ""
    name := "dialog opens"
    if !IsObject(first)
        first := Gui("", "QuickSwitchCheck first"), second := Gui("", "QuickSwitchCheck second")
    WinActivate(Explorer)                                                   ; Explorer: the file manager used last
    if !WinWaitActive(Explorer, , 3)
        return Fail(name, n, "Explorer didn't come to the front")
    Sleep(300)
    ; Two windows of this check one after the other: the first comes from Explorer (ALTRun counts a switch from a file manager), the
    ; second from the first (not from a file manager any more), so AutoSwitch alone wouldn't jump the dialog that opens next
    for window in [first, second] {
        window.Show("w300 h100")
        if !WinWaitActive("ahk_id " window.Hwnd, , 3)
            return Fail(name, n, "a window of this check didn't come to the front")
        Sleep(400)
    }
    helper := A_Temp "\QuickSwitchCheckOpen.ahk"
    try FileDelete(helper)
    FileAppend('#NoTrayIcon`nFileSelect(, "' Demo '\Start", "Open on open")`n', helper, "UTF-8")
    logSize := LogLength()
    start := A_TickCount
    Run('"' A_AhkPath '" "' helper '"', , , &pid)
    opened := WinWait("Open on open ahk_class #32770 ahk_pid " pid, , 10)
    try {
        if !opened
            return Fail(name, n, "the dialog didn't open")
        while (A_TickCount - start < 5000 && !ShowsFolder("Folder A", opened))
            Sleep(20)
        if !ShowsFolder("Folder A", opened)
            return Fail(name, n, "the new dialog didn't jump to Folder A" (ShowsFolder("Start", opened) ? " (still at Start)" : ""), logSize)
        if !InStr(LogSince(logSize), "QuickSwitch: a dialog opened")             ; the jump came from AutoSwitchOnOpen, not from AutoSwitch
            return Fail(name, n, "it jumped, but not because the dialog opened (no 'a dialog opened' in ALTRun's log)", logSize)
        Say("PASS " name " " n ": jumped as it opened in " (A_TickCount - start) " ms")
    } finally {
        first.Hide(), second.Hide()
        if opened
            try WinClose(opened)
        if (pid && !ProcessWaitClose(pid, 3))
            try ProcessClose(pid)
        Sleep(300)
    }
}

; An extract window (#609) of 7-Zip (7zG x -ad) or WinRAR (Alt+E in its main window), opened with Explorer at Folder A used last:
; no jump by itself (AutoSwitch and AutoSwitchOnOpen are on), Ctrl+G fills Folder A into the destination box (Edit1 in both) without
; extracting. Skipped when the program isn't installed
ExtractRound(program) {
    name := "extract " program
    sevenZip := A_ProgramFiles "\7-Zip", winRar := A_ProgramFiles "\WinRAR\WinRAR.exe"
    if !FileExist(sevenZip "\7z.exe")
        return Say("SKIP " name ": no 7-Zip to make the archive")
    if (program = "WinRAR" && !FileExist(winRar))
        return Say("SKIP " name ": WinRAR isn't installed")
    archive := Demo "\test.zip", target := Demo "\Folder A\"
    if !FileExist(archive) {
        FileAppend("text", Demo "\Start\a.txt")
        RunWait('"' sevenZip '\7z.exe" a "' archive '" "' Demo '\Start\a.txt"', , "Hide")
    }
    WinActivate(Explorer)                                                   ; Explorer: the file manager used last
    if !WinWaitActive(Explorer, , 3)
        return Fail(name, 1, "Explorer didn't come to the front")
    Sleep(300)
    logSize := LogLength(), pids := []
    try {
        if (program = "7-Zip") {
            Run('"' sevenZip '\7zG.exe" x "' archive '" -o"' Demo '\Start" -ad', Demo, , &pid), pids.Push(pid)
            window := WinWait("ahk_class #32770 ahk_pid " pid, , 10)
        } else {
            Run('"' winRar '" "' archive '"', Demo, , &pid), pids.Push(pid)
            if !(main := WinWait("ahk_class WinRarWindow ahk_pid " pid, , 20))
                return Fail(name, 1, "WinRAR's main window didn't open")
            WinActivate(main)
            if !WinWaitActive(main, , 3)
                return Fail(name, 1, "WinRAR's main window didn't come to the front")
            Sleep(500)
            Send("!e")                                                      ; Extract To
            window := WinWait("ahk_class #32770 ahk_pid " pid, , 10)
        }
        if !window {
            Windows()
            return Fail(name, 1, "the extract window didn't open")
        }
        Sleep(300)
        before := box()
        Say("INFO " name ": " WinGetTitle(window) ", destination [" before "]")
        Sleep(1500)                                                         ; AutoSwitchOnOpen jumps a new dialog within a few hundred ms
        if (box() != before)
            return Fail(name, 1, "the destination changed by itself as the window opened: [" box() "]", logSize)
        WinActivate(Explorer), WinWaitActive(Explorer, , 3), Sleep(400)      ; a switch from Explorer: AutoSwitch must leave it alone too
        WinActivate(window)
        if !WinWaitActive(window, , 3)
            return Fail(name, 1, "the extract window didn't come back to the front")
        Sleep(1500)
        if (box() != before)
            return Fail(name, 1, "the destination changed on a switch from Explorer: [" box() "]", logSize)
        Say("PASS " name ": no jump by itself")
        start := A_TickCount
        SendLevel(1)                                                        ; hook hotkeys ignore other scripts' keys at send level 0
        SendEvent("^g")
        SendLevel(0)
        while (A_TickCount - start < 3000 && box() != target)
            Sleep(20)
        if (box() != target)
            return Fail(name " ctrl+g", 1, "the destination box has [" box() "], not [" target "]", logSize)
        Sleep(1000)
        if !WinExist(window)
            return Fail(name " ctrl+g", 1, "the extract window closed: Ctrl+G must not start extracting", logSize)
        if FileExist(target "a.txt")
            return Fail(name " ctrl+g", 1, "a.txt was extracted into Folder A", logSize)
        Say("PASS " name " ctrl+g: filled the destination in " (A_TickCount - start) " ms, still open, nothing extracted")
    } finally {
        for pid in pids
            try ProcessClose(pid)
        Sleep(300)
    }

    box() {                                                                 ; the destination box: Edit1 in both programs
        try return ControlGetText("Edit1", window)
        return ""
    }
}

; The panel's ☆ menu (how: "click" on the star, or "alt+down" in the search box), closed with Esc: the dialog is in front again
MenuRound(how, n) {
    name := "menu " how " + esc"
    WinActivate(Dialog)
    if !WinWaitActive(Dialog, , 3)
        return Fail(name, n, "the dialog didn't come to the front")
    panel := 0
    start := A_TickCount
    while (!panel && A_TickCount - start < 3000)
        panel := WinExist("ALTRun Quick Switch ahk_pid " AppPid), Sleep(50)
    if !panel
        return Fail(name, n, "the folder panel isn't shown")
    star := "", panelW := 0
    WinGetClientPos(&px, &py, &panelW, , panel)
    for ctl in WinGetControls(panel)                                        ; the star: the small picture (the other Static is the full-width line)
        if (InStr(ctl, "Static") = 1) {
            ControlGetPos(&cx, &cy, &cw, &ch, ctl, panel)
            if (cw < panelW // 4)
                star := {X: px + cx + cw // 2, Y: py + cy + ch // 2}
        }
    if !IsObject(star)
        return Fail(name, n, "no star button in the panel")
    CoordMode("Mouse", "Screen")
    if (how = "click")
        Click(star.X, star.Y)
    else {
        ControlGetPos(&ex, &ey, &ew, &eh, "Edit1", panel)                   ; click into the panel's search box, then Alt+↓
        Click(px + ex + ew // 2, py + ey + eh // 2)
        Sleep(200)
        Send("!{Down}")
    }
    popup := WinWait("ahk_class #32768", , 3)
    if !popup
        return Fail(name, n, "the menu didn't open")
    before := DllCall("GetForegroundWindow", "Ptr")                         ; with the menu open: ALTRun's hidden window, the menu's owner
    Send("{Esc}")
    start := A_TickCount
    ; Every window that comes to the front after Esc, with the time, until the dialog does (at most 2 s): when it doesn't,
    ; this shows which window kept the front and whether the dialog came back late or not at all
    front := -1, trace := "", closed := -1
    while (A_TickCount - start < 2000) {
        if (closed < 0 && !WinExist("ahk_id " popup))
            closed := A_TickCount - start
        now := DllCall("GetForegroundWindow", "Ptr")
        if (now != front)
            front := now, trace .= "`n       " (A_TickCount - start) " ms: " Describe(front)
        if (front = Dialog && closed >= 0)
            break
        Sleep(10)
    }
    if (front != Dialog)
        return Fail(name, n, "after Esc the window in front is " Describe(front) ", not the dialog`n     before Esc: " Describe(before)
            . "`n     menu closed: " (closed < 0 ? "no" : "after " closed " ms") "`n     in front after Esc:" trace)
    Say("PASS " name " " n ": the dialog is in front again after " (A_TickCount - start) " ms")
    Sleep(300)                                                              ; a click on the star right after the menu closed doesn't open it again
}

; A window for the failure messages: its handle, class, process and whose it is (ALTRun's hidden main window, this check, ...)
Describe(hwnd) {
    if !hwnd
        return "none"
    prev := DetectHiddenWindows(true)
    try {
        pid := WinGetPID(hwnd)
        whose := hwnd = Dialog ? "the dialog" : hwnd = Explorer ? "Explorer" : hwnd = AppWindow ? "ALTRun's hidden main window"
            : pid = AppPid ? "ALTRun" : pid = DialogPid ? "the dialog's script" : pid = ProcessExist() ? "this check" : WinGetProcessName(hwnd)
        text := Format("0x{:X} {} of pid {} ({}{})", hwnd, WinGetClass(hwnd), pid, whose, DllCall("IsWindowVisible", "Ptr", hwnd) ? "" : ", hidden")
    } catch
        text := Format("0x{:X} (gone)", hwnd)
    DetectHiddenWindows(prev)
    return text
}

Fail(method, n, why, logSize := -1) {
    global Failures
    Failures++
    Say("FAIL " method " " n ": " why)
    if (logSize >= 0)                                                       ; every line: what kept ALTRun busy shows too (Perf: slow ...)
        for line in StrSplit(Trim(LogSince(logSize), "`r`n"), "`n", "`r")
            Say("     " line)
}

; Types the Start folder into the file name box and presses Enter: the dialog goes there
BackToStart() {
    WinActivate(Dialog)
    if !WinWaitActive(Dialog, , 3)
        return false
    try ControlSetText(Demo "\Start", "Edit1", Dialog)
    catch
        return false
    ControlSend("{Enter}", "Edit1", Dialog)
    start := A_TickCount
    while (A_TickCount - start < 3000 && !ShowsFolder("Start"))
        Sleep(20)
    return ShowsFolder("Start")
}

; The dialog's address bar (a ToolbarWindow32 "Address: ...") shows the folder
ShowsFolder(name, window := 0) {      ; window: the dialog, the global Dialog when 0 (names don't differ by case alone: AutoHotkey ignores case)
    window := window ? window : Dialog
    try {
        for ctl in WinGetControls(window)
            if (InStr(ctl, "ToolbarWindow32") = 1 && RegExMatch(ControlGetText(ctl, window), "\Q" name "\E$"))
                return true
    }
    return false
}

; The visible top-level windows, to see what is there when something didn't appear
Windows() {
    for hwnd in WinGetList()
        try Say("     window: " WinGetClass(hwnd) " | " WinGetTitle(hwnd) " | " WinGetProcessName(hwnd))
}

LogLength() => FileExist(LogFile) ? FileGetSize(LogFile) : 0

LogSince(size) {
    try {
        f := FileOpen(LogFile, "r", "UTF-8")                                 ; shared: ALTRun keeps it open for writing ("-w" would refuse that)
        f.Pos := size
        text := f.Read()
        f.Close()
        return text
    }
    return ""
}

TearDown() {
    if Failures {                                                           ; what ALTRun did up to the end (a check that took long shows as "Perf: slow")
        lines := StrSplit(Trim(LogSince(0), "`r`n"), "`n", "`r")
        Loop Min(40, lines.Length)
            Say("     final log: " lines[lines.Length - Min(40, lines.Length) + A_Index])
    }
    try WinClose(Dialog)
    if DialogPid && !ProcessWaitClose(DialogPid, 3)
        try ProcessClose(DialogPid)
    try WinClose(Explorer)
    if AppPid {
        DetectHiddenWindows(true)
        try WinClose("ahk_class AutoHotkey ahk_pid " AppPid)
        if !ProcessWaitClose(AppPid, 10)
            try ProcessClose(AppPid)
    }
}
