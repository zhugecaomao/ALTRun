;===============================================================================
; QuickSwitchCheck.ahk - automatic switching (Extensions.QuickSwitch.AutoSwitch) on a real Windows desktop (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Run by .github/workflows/quickswitch.yml; also works on your own PC (it takes over the keyboard and the screen for a minute
; and starts its own copy of ALTRun from <appDir> with settings of its own). Opens Explorer at "Folder A" and an Open dialog at
; "Start", then switches from Explorer to the dialog in several ways and checks that the dialog jumps to Folder A:
;   activate      SetForegroundWindow (WinActivate), as a program bringing the dialog to the front
;   alt+tab       a simulated Alt+Tab (SendInput), as the keyboard does it
;   title click   a click on the dialog's title bar, as the mouse does it
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

global AppDir := A_Args.Length ? A_Args[1] : A_ScriptDir "\..\.."
global Rounds := A_Args.Length > 1 ? Integer(A_Args[2]) : 5
global Demo := "C:\QuickSwitchCheck\demo", LogFile := A_Temp "\ALTRun.log"
global Failures := 0, AppPid := 0, DialogPid := 0, Explorer := 0, Dialog := 0
SetTitleMatchMode(3)
SetKeyDelay(30, 30)

try {
    Setup()
    for method in ["activate", "alt+tab", "title click"]
        Loop Rounds
            Round_(method, A_Index)
} catch as setupError {
    Say("FAIL setup: " setupError.Message " (line " setupError.Line ")")
    Failures++
}
TearDown()
Say(Failures ? Failures " round(s) failed" : "all rounds passed")
ExitApp(Failures)

Say(text) => FileAppend(text "`n", "*")

Setup() {
    global AppPid, DialogPid, Explorer, Dialog
    try DirDelete(Demo, true)
    DirCreate(Demo "\Folder A"), DirCreate(Demo "\Start")
    settings := Map("SchemaVersion", 6
        , "General", Map("Language", "en", "LaunchAtLogin", 0, "CheckForUpdates", 0, "ShowTips", 0, "SaveLog", 1, "SendToMenu", 0, "StartMenuShortcut", 0)
        , "Extensions", Map("QuickSwitch", Map("AutoSwitch", 0)))
    try DirDelete(AppDir "\Data", true)
    DirCreate(AppDir "\Data")
    FileAppend(JSON.Stringify(settings, 4), AppDir "\Data\ALTRun.json", "UTF-8")
    try FileDelete(LogFile)

    Run('"' A_AhkPath '" "' AppDir '\ALTRun.ahk"', AppDir, , &AppPid)
    DetectHiddenWindows(true)
    SetTitleMatchMode(2)
    main := WinWait("ALTRun.ahk ahk_class AutoHotkey ahk_pid " AppPid, , 30)
    SetTitleMatchMode(3)
    DetectHiddenWindows(false)
    if !main
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
