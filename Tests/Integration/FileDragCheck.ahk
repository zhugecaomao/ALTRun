;===============================================================================
; FileDragCheck.ahk - dragging files out of the search window with the real mouse on a Windows desktop (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Run by .github/workflows/file-drag.yml; also works on your own PC (it takes over the mouse and keyboard for about a minute and
; starts its own copy of ALTRun from <appDir> with settings of its own: one pinned file, so the empty search box lists it):
;   drag-out    Alt+Space, drag the pinned file (row 1) onto an Explorer window of another folder: the file is copied there, the
;               original stays (never moved), and the search window hides
;   cancel      the same drag, Esc before letting go: nothing is copied, the search window is still shown
; Each is tried Rounds times. Prints one line per round (PASS / FAIL) and ALTRun's log on failure; the exit code is the number of
; failed rounds.
;
;   AutoHotkey64.exe FileDragCheck.ahk <appDir> [rounds]       appDir: a copy of the repository (ALTRun.ahk, Lib, Src, Resources)
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut
#Include ..\..\Lib\JSON.ahk

global AppDir := A_Args.Length ? A_Args[1] : A_ScriptDir "\..\.."
global Rounds := A_Args.Length > 1 ? Integer(A_Args[2]) : 3
global Work := "C:\FileDragCheck\work", LogFile := A_Temp "\ALTRun.log"
global SourceDir := Work "\Source", TargetDir := Work "\Target", DragFile := Work "\Source\drag-me.txt"
global Failures := 0, AppPid := 0, Explorer := 0
CoordMode("Mouse", "Screen")
OnError(Crashed)

try {
    Setup()
    Loop Rounds
        DragOut(A_Index)
    Loop Rounds
        DragCancel(A_Index)
} catch as setupError {
    Say("FAIL setup: " setupError.Message " (line " setupError.Line ")")
    Failures++
}
TearDown()
Say(Failures ? Failures " round(s) failed" : "all rounds passed")
ExitApp(Failures)

Say(text) => FileAppend(text "`n", "*")

; An error anywhere: one FAIL line and out, never a message box that would hold the runner until its time limit
Crashed(e, *) {
    Say("FAIL error: " e.Message " (line " e.Line ")")
    TearDown()
    ExitApp(Failures + 1)
}

Setup() {
    global AppPid, Explorer
    try DirDelete(Work, true)
    DirCreate(SourceDir), DirCreate(TargetDir)
    FileAppend("drag me", DragFile, "UTF-8")
    entry := Map("Title", "drag-me.txt", "Subtitle", DragFile, "Kind", "file", "Arg", DragFile, "Arguments", "", "Icon", DragFile
        , "Uid", "file:" StrLower(DragFile))
    settings := Map("SchemaVersion", 7
        , "General", Map("Language", "en", "LaunchAtLogin", 0, "CheckForUpdates", 0, "SaveLog", 1, "SendToMenu", 0, "StartMenuShortcut", 0)
        , "Window", Map("ShowTips", 0, "ShowPinned", 1, "ShowRecent", 0)
        , "Features", Map("Recent", Map("Pinned", [entry])))
    try DirDelete(AppDir "\Data", true)
    DirCreate(AppDir "\Data")
    FileAppend(JSON.Stringify(settings, 4), AppDir "\Data\ALTRun.json", "UTF-8")
    try FileDelete(LogFile)

    Run('explorer.exe "' TargetDir '"')                                     ; the drop target of drag-out: the whole screen, behind the search window
    if !(Explorer := WinWait("Target ahk_class CabinetWClass", , 30))       ; the first Explorer window on a fresh runner takes a while
        throw Error("the Explorer window didn't open")
    WinMove(0, 0, A_ScreenWidth, A_ScreenHeight - 60, Explorer)

    Run('"' A_AhkPath '" "' AppDir '\ALTRun.ahk" -Startup', AppDir, , &AppPid)   ; -Startup: no search window
    DetectHiddenWindows(true)
    SetTitleMatchMode(2)
    window := WinWait("ALTRun.ahk ahk_class AutoHotkey ahk_pid " AppPid, , 30)
    DetectHiddenWindows(false)
    if !window
        throw Error("ALTRun didn't start")
    Sleep(5000)                                                             ; start-up done (hotkeys registered)
    Say("INFO ALTRun pid " AppPid ", screen " A_ScreenWidth "x" A_ScreenHeight ", " Rounds " round(s) each")
}

; Explorer in front, then Alt+Space: the search window with the pinned file in row 1; returns its hwnd (0 if it didn't show)
ShowSearch() {
    WinActivate(Explorer)
    WinWaitActive(Explorer, , 3)
    Sleep(200)
    SendLevel(1), SendEvent("!{Space}"), SendLevel(0)
    if !(win := WinWait(SearchTitle(), , 3))
        return 0
    WinWaitActive(win, , 2)
    Sleep(400)                                                              ; the list is filled and painted
    return win
}

SearchTitle() => "ALTRun ahk_class AutoHotkeyGUI ahk_pid " AppPid

; The screen position of row 1 of the result list -> {X, Y}
RowOne(win) {
    ControlGetPos(&cx, &cy, &cw, &ch, "SysListView321", win)
    WinGetClientPos(&wx, &wy, , , win)
    return {X: wx + cx + cw // 3, Y: wy + cy + Min(ch // 2, 20)}
}

; A point on Explorer's file list below the search window
DropPoint(win) {
    WinGetPos(&x, &y, &w, &h, win)
    return {X: x + w // 2, Y: Min(y + h + 120, A_ScreenHeight - 150)}
}

; From the current position to (x, y) in steps, as a hand would (one jump wouldn't start a drag or pass over the windows)
Glide(x, y, steps := 20) {
    MouseGetPos(&mx, &my)
    Loop steps {
        MouseMove(mx + (x - mx) * A_Index // steps, my + (y - my) * A_Index // steps, 0)
        Sleep(25)
    }
}

DragOut(n) {
    copied := TargetDir "\drag-me.txt"
    try FileDelete(copied)
    offset := LogLength()
    if !(win := ShowSearch())
        return Result("drag-out", n, false, "Alt+Space didn't show the search window in 3 s", offset)
    row := RowOne(win), drop := DropPoint(win)
    MouseMove(row.X, row.Y, 0), Sleep(150)
    Click("Down")
    Glide(row.X + 30, row.Y + 30, 6)                                        ; past the drag threshold: the list's LVN_BEGINDRAG
    Sleep(300)                                                              ; the drag starts from a timer
    Glide(drop.X, drop.Y)
    Sleep(300)
    Click("Up")
    start := A_TickCount
    while (!FileExist(copied) && A_TickCount - start < 5000)
        Sleep(100)
    there := FileExist(copied) != "", kept := FileExist(DragFile) != ""
    hidden := WinWaitClose(win, , 3)
    if !hidden
        SendEvent("{Esc}"), WinWaitClose(win, , 2)
    Result("drag-out", n, there && kept && hidden
        , !there ? "the file wasn't copied to the Explorer folder in 5 s" : !kept ? "the original is gone (moved)" : !hidden ? "copied, but the search window didn't hide" : "copied, original kept, window hidden"
        , offset)
}

DragCancel(n) {
    copied := TargetDir "\drag-me.txt"
    try FileDelete(copied)
    offset := LogLength()
    if !(win := ShowSearch())
        return Result("cancel", n, false, "Alt+Space didn't show the search window in 3 s", offset)
    row := RowOne(win), drop := DropPoint(win)
    MouseMove(row.X, row.Y, 0), Sleep(150)
    Click("Down")
    Glide(row.X + 30, row.Y + 30, 6)
    Sleep(300)
    Glide(drop.X, drop.Y)
    Sleep(200)
    SendEvent("{Esc}")                                                      ; cancels the drag
    Sleep(300)
    Click("Up")
    Sleep(1500)
    none := !FileExist(copied), shown := WinExist(win) && DllCall("IsWindowVisible", "Ptr", win)
    if shown
        SendEvent("{Esc}"), WinWaitClose(win, , 3)
    Result("cancel", n, none && shown, !none ? "the file was copied although Esc cancelled the drag" : !shown ? "nothing copied, but the search window hid" : "nothing copied, window still shown", offset)
}

Result(name, n, ok, detail, offset := 0) {
    global Failures
    Say((ok ? "PASS " : "FAIL ") name " " n ": " detail)
    if !ok {
        Failures++
        Say("--- ALTRun log ---`n" LogSince(offset) "`n---")
    }
}

LogLength() => FileExist(LogFile) ? FileGetSize(LogFile) : 0

LogSince(offset) {
    try {
        f := FileOpen(LogFile, "r", "UTF-8")                                ; shared: ALTRun keeps it open for writing
        f.Seek(offset)
        text := f.Read()
        f.Close()
        return text
    }
    return "(no log)"
}

TearDown() {
    global AppPid
    try WinClose(Explorer)
    if AppPid {
        DetectHiddenWindows(true)
        try WinClose("ahk_class AutoHotkey ahk_pid " AppPid)
        if !ProcessWaitClose(AppPid, 10)
            try ProcessClose(AppPid)
        AppPid := 0
    }
}
