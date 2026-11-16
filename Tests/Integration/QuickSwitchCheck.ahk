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
; Then Explorer with a second tab at "Folder B", shown (#521; Windows 11 22H2 and later): switching to the dialog must jump to
; Folder B, not to Folder A of the first tab, and so must the jump hotkey (Ctrl+G); the tab is closed again after each round.
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
    for how in ["click", "alt+down"]
        Loop Rounds
            MenuRound(how, A_Index)
    if ExplorerTabs.Supported()
        Loop Rounds
            TabRound(A_Index)
    else
        Say("SKIP tab shown: no Explorer tabs on Windows " A_OSVersion)
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
        , "Extensions", Map("QuickSwitch", Map("AutoSwitch", 1)))
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
    if !ExplorerTabs.Open(folderB)                                          ; a new tab of the Explorer window, shown and in front
        return TabFail(name, n, "no new tab at Folder B: " ExplorerTabs.LastError)
    try {
        SetTitleMatchMode(2)
        shown := WinWait("Folder B ahk_id " Explorer, , 3)                  ; the title is the name of the tab shown
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
        Send("^g")
        while (A_TickCount - start < 3000 && !ShowsFolder("Folder B"))
            Sleep(20)
        if !ShowsFolder("Folder B")
            return TabFail(name " ctrl+g", n, "Ctrl+G didn't jump to Folder B of the tab shown" (ShowsFolder("Folder A") ? " (it went to Folder A of the first tab)" : ""), logSize)
        Say("PASS " name " ctrl+g " n ": jumped to the tab shown in " (A_TickCount - start) " ms")
    } finally {
        for each in ExplorerTabs._Tabs(Explorer)                            ; close the Folder B tab: Explorer shows Folder A again
            if ExplorerTabs.SameFolder(ExplorerTabs._Path(Explorer, each.Tab), folderB)
                try each.Item.Quit()
        Sleep(500)
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
    Say("     ALTRun running: " (ProcessExist(AppPid) ? "yes" : "no"))
    for hwnd in WinGetList()
        try Say("     window: " Describe(hwnd) " | " WinGetTitle(hwnd))
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
    if (logSize >= 0)
        for line in StrSplit(LogSince(logSize), "`n", "`r")
            if InStr(line, "QuickSwitch")
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
ShowsFolder(name) {
    try {
        for ctl in WinGetControls(Dialog)
            if (InStr(ctl, "ToolbarWindow32") = 1 && RegExMatch(ControlGetText(ctl, Dialog), "\Q" name "\E$"))
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
        f := FileOpen(LogFile, "r -w", "UTF-8")                              ; ALTRun keeps it open for writing
        f.Pos := size
        text := f.Read()
        f.Close()
        return text
    }
    return ""
}

TearDown() {
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
