;===============================================================================
; HotkeyCheck.ahk - custom hotkeys that run a custom command ("custom:<Id>") or a script ("script:<file name>"), pressed for real on a
; Windows desktop (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Run by .github/workflows/hotkeys.yml; also works on your own PC (it takes over the keyboard for half a minute and starts its own
; copy of ALTRun from <appDir> with settings of its own). The commands write marker files with cmd.exe, so the check only looks at
; what ALTRun is responsible for:
;   plain       Ctrl+Alt+F9 runs a command without {query}: its marker is written
;   selection   Ctrl+Alt+F10 runs a {query} command while text is selected in a window of this script: the marker holds that
;               text, and the clipboard is what it was before (ALTRun copies the selection with Ctrl+C and puts the clipboard back)
;   deleted     Ctrl+Alt+F11 points to a command that doesn't exist: nothing runs, ALTRun's log says so
;   script      Ctrl+Alt+F12 runs a script with "@altrun.selection Text" (in <appDir>\Scripts) while text is selected: the script writes
;               its first argument to its marker, so the marker holds that text, and the clipboard is what it was before
;   alt         double-tap left Alt (DoubleTap = LAlt, #501) over a window with a menu bar: the search window shows; after Esc the
;               window isn't left in menu mode (a lone Alt opens its menu bar, the second tap closes it again)
;   alt-once    one Alt tap doesn't show the search window
;   alt-space   Alt+Space (the main hotkey) shows it, a lone Alt tap right after doesn't toggle it again
; Each is tried Rounds times. Prints one line per round (PASS / FAIL) and ALTRun's log on failure; the exit code is the number of
; failed rounds.
;
;   AutoHotkey64.exe HotkeyCheck.ahk <appDir> [rounds]       appDir: a copy of the repository (ALTRun.ahk, Lib, Src, Resources)
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut
#Include ..\..\Lib\JSON.ahk

global AppDir := A_Args.Length ? A_Args[1] : A_ScriptDir "\..\.."
global Rounds := A_Args.Length > 1 ? Integer(A_Args[2]) : 3
global Work := "C:\HotkeyCheck\work", LogFile := A_Temp "\ALTRun.log"
global PlainMarker := Work "\plain.txt", QueryMarker := Work "\query.txt", ScriptMarker := Work "\script.txt"
global ScriptsDir := AppDir "\Scripts"                                    ; a portable copy (its folder is writable): the Scripts folder is in it
global Failures := 0, AppPid := 0, Source := "", SourceEdit := "", MenuPid := 0, MenuWin := 0
OnError(Crashed)

try {
    Setup()
    Loop Rounds
        Plain(A_Index)
    Loop Rounds
        Selection(A_Index)
    Loop Rounds
        Deleted(A_Index)
    Loop Rounds
        ScriptSelection(A_Index)
    Loop Rounds
        DoubleTapAlt(A_Index)
    Loop Rounds
        SingleAlt(A_Index)
    Loop Rounds
        AltSpaceThenAlt(A_Index)
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
    global AppPid, Source, SourceEdit, MenuPid, MenuWin
    try DirDelete(Work, true)
    DirCreate(Work)
    cmd := A_WinDir "\System32\cmd.exe"                                     ; full paths: nothing depends on the working folder
    settings := Map("SchemaVersion", 7
        , "General", Map("Language", "en", "LaunchAtLogin", 0, "CheckForUpdates", 0, "SaveLog", 1, "SendToMenu", 0, "StartMenuShortcut", 0
            , "DoubleTap", "LAlt")
        , "Window", Map("ShowTips", 0)
        , "CustomCommands", [
            Map("Id", "plain001", "Title", "Write plain marker", "Type", "Command", "Target", cmd, "Arguments", '/c echo plain> "' PlainMarker '"', "Keyword", "")
          , Map("Id", "query001", "Title", "Write query marker", "Type", "Command", "Target", cmd, "Arguments", '/c echo {query}> "' QueryMarker '"', "Keyword", "")]
        , "Hotkeys", [
            Map("Key", "^!F9", "Action", "custom:plain001", "WinTitle", "")
          , Map("Key", "^!F10", "Action", "custom:query001", "WinTitle", "")
          , Map("Key", "^!F11", "Action", "custom:gone0001", "WinTitle", "")
          , Map("Key", "^!F12", "Action", "script:marker.ahk", "WinTitle", "")])
    try DirDelete(ScriptsDir, true)
    DirCreate(ScriptsDir)
    FileAppend("; @altrun.title Write script marker`n; @altrun.selection Text`n"
        . 'FileAppend((A_Args.Length ? A_Args[1] : "") "``n", "' ScriptMarker '")`n', ScriptsDir "\marker.ahk", "UTF-8")
    try DirDelete(AppDir "\Data", true)
    DirCreate(AppDir "\Data")
    FileAppend(JSON.Stringify(settings, 4), AppDir "\Data\ALTRun.json", "UTF-8")
    try FileDelete(LogFile)

    Run('"' A_AhkPath '" "' AppDir '\ALTRun.ahk" -Startup', AppDir, , &AppPid)   ; -Startup: no search window
    DetectHiddenWindows(true)
    SetTitleMatchMode(2)
    window := WinWait("ALTRun.ahk ahk_class AutoHotkey ahk_pid " AppPid, , 30)
    DetectHiddenWindows(false)
    if !window
        throw Error("ALTRun didn't start")
    Sleep(5000)                                                             ; start-up done (hotkeys registered)

    Source := Gui("+AlwaysOnTop", "HotkeyCheck source")                     ; the program with the selection: an edit box of this script
    SourceEdit := Source.AddEdit("w300 r1", "hello123")
    Source.Show()
    ; a window with a menu bar, in a process of its own (a lone Alt opens the menu bar, like in Notepad: its menu loop would hold this script)
    FileAppend('#NoTrayIcon`ng := Gui("+AlwaysOnTop", "HotkeyCheck menu")`nm := Menu(), m.Add("Item", (*) => 0)`n'
        . 'g.MenuBar := MenuBar(), g.MenuBar.Add("&File", m)`ng.AddEdit("w300 r1")`ng.OnEvent("Close", (*) => ExitApp())`ng.Show()`n'
        , Work "\menu.ahk", "UTF-8")
    Run('"' A_AhkPath '" "' Work '\menu.ahk"', Work, , &MenuPid)
    if !(MenuWin := WinWait("HotkeyCheck menu ahk_pid " MenuPid, , 10))
        throw Error("the menu window didn't open")
    Say("INFO ALTRun pid " AppPid ", " Rounds " round(s) each")
}

Plain(n) {
    try FileDelete(PlainMarker)
    FrontSource()
    SendInput("^!{F9}")
    text := WaitFor(PlainMarker)
    Result("plain", n, InStr(text, "plain") > 0, "marker: " (text = "" ? "not written in 5 s" : Trim(text, "`r`n ")))
}

Selection(n) {
    try FileDelete(QueryMarker)
    A_Clipboard := "clipboard before " n
    SourceEdit.Value := "hello123"                                          ; a failed round may have typed into it
    FrontSource()
    SendMessage(0xB1, 0, -1, SourceEdit)                                    ; EM_SETSEL: all of it selected
    SendInput("^!{F10}")
    text := WaitFor(QueryMarker)
    Sleep(300)                                                              ; the clipboard is put back right after the command starts
    clip := A_Clipboard
    Result("selection", n, Trim(text, "`r`n ") == "hello123" && clip == "clipboard before " n
        , "marker: " (text = "" ? "not written in 5 s" : Trim(text, "`r`n ")) ", clipboard: " clip)
}

Deleted(n) {
    try FileDelete(PlainMarker)
    try FileDelete(QueryMarker)
    FrontSource()
    before := LogLength()
    SendInput("^!{F11}")
    found := false, start := A_TickCount
    while (!found && A_TickCount - start < 5000) {
        Sleep(200)
        found := InStr(LogSince(before), "custom:gone0001 no longer exists") > 0
    }
    ran := FileExist(PlainMarker) || FileExist(QueryMarker)
    Result("deleted", n, found && !ran, (found ? "logged" : "no log line in 5 s") (ran ? ", but a command ran" : ", nothing ran"))
}

ScriptSelection(n) {
    try FileDelete(ScriptMarker)
    A_Clipboard := "clipboard before script " n
    SourceEdit.Value := "hello123"
    FrontSource()
    SendMessage(0xB1, 0, -1, SourceEdit)                                    ; EM_SETSEL: all of it selected
    SendInput("^!{F12}")
    text := WaitFor(ScriptMarker)
    Sleep(300)
    clip := A_Clipboard
    Result("script", n, Trim(text, "`r`n ") == "hello123" && clip == "clipboard before script " n
        , "marker: " (text = "" ? "not written in 5 s" : Trim(text, "`r`n ")) ", clipboard: " clip)
}

; Double-tap left Alt (hook hotkeys ignore other scripts' keys at SendLevel 0: sent at level 1, like a person's)
DoubleTapAlt(n) {
    FrontMenu()
    TapAlt(), Sleep(120), TapAlt()
    shown := WinWait(SearchTitle(), , 3)
    if shown {
        SendEvent("{Esc}")
        WinWaitClose(shown, , 3)
    }
    Sleep(500)
    inMenu := InMenuMode(MenuWin), back := WinActive(MenuWin)
    Result("alt", n, shown && !inMenu && back, (shown ? "shown" : "not shown in 3 s") (inMenu ? ", the window was left in menu mode" : "")
        . (back ? "" : ", the window didn't get the focus back"))
    LeaveMenu()
}

SingleAlt(n) {
    FrontMenu()
    TapAlt()
    Sleep(900)                                                              ; longer than a double tap may take
    shown := WinExist(SearchTitle())
    if shown {
        SendEvent("{Esc}")
        WinWaitClose(shown, , 3)
    }
    Result("alt-once", n, !shown, shown ? "one tap showed the search window" : "not shown")
    LeaveMenu()
}

AltSpaceThenAlt(n) {
    FrontMenu()
    SendLevel(1), SendEvent("!{Space}"), SendLevel(0)
    shown := WinWait(SearchTitle(), , 3)
    still := false
    if shown {
        Sleep(150), TapAlt(), Sleep(700)
        still := WinExist(SearchTitle())
        SendEvent("{Esc}")
        WinWaitClose(shown, , 3)
    }
    Result("alt-space", n, shown && still, !shown ? "Alt+Space didn't show it in 3 s" : still ? "still shown after a lone Alt" : "a lone Alt after Alt+Space hid it")
    LeaveMenu()
}

TapAlt() {
    SendLevel(1)
    SendEvent("{LAlt}")
    SendLevel(0)
}

SearchTitle() => "ALTRun ahk_class AutoHotkeyGUI ahk_pid " AppPid

; Whether hwnd's thread is in a menu loop (GetGUIThreadInfo: GUI_INMENUMODE | GUI_SYSTEMMENUMODE | GUI_POPUPMENUMODE)
InMenuMode(hwnd) {
    info := Buffer(A_PtrSize = 8 ? 72 : 48, 0)
    NumPut("UInt", info.Size, info)
    tid := DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "Ptr", 0, "UInt")
    return tid && DllCall("GetGUIThreadInfo", "UInt", tid, "Ptr", info) && (NumGet(info, 4, "UInt") & 0x1C) != 0
}

LeaveMenu() {                                                               ; a failed round mustn't leave the menu open for the next
    if InMenuMode(MenuWin)
        SendEvent("{Esc}{Esc}")
}

FrontMenu() {
    WinActivate(MenuWin)
    if !WinWaitActive(MenuWin, , 3)
        throw Error("the menu window didn't come to the front")
    Sleep(200)
}

FrontSource() {
    WinActivate(Source.Hwnd)
    if !WinWaitActive(Source.Hwnd, , 3)
        throw Error("the source window didn't come to the front")
    SourceEdit.Focus()
    Sleep(200)
}

; The marker's text once it is written (cmd may still be writing: read until it has a line), "" after 5 s
WaitFor(file) {
    start := A_TickCount
    while (A_TickCount - start < 5000) {
        Sleep(100)
        text := ""
        try text := FileRead(file)
        if InStr(text, "`n")
            return text
    }
    return ""
}

Result(name, n, ok, detail) {
    global Failures
    Say((ok ? "PASS " : "FAIL ") name " " n ": " detail)
    if !ok {
        Failures++
        Say("--- ALTRun log ---`n" LogSince(0) "`n---")
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
    return ""
}

TearDown() {
    global AppPid
    try Source.Destroy()
    if MenuPid
        try ProcessClose(MenuPid)
    if AppPid {
        DetectHiddenWindows(true)
        try WinClose("ahk_class AutoHotkey ahk_pid " AppPid)
        if !ProcessWaitClose(AppPid, 10)
            try ProcessClose(AppPid)
        AppPid := 0
    }
    try DirDelete(ScriptsDir, true)
}
