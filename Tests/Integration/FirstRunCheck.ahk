;===============================================================================
; FirstRunCheck.ahk - the first start of a new installation asks whether to start at login, on a real Windows desktop (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Run by .github/workflows/first-run.yml; also works on your own PC (about a minute): it starts its own copy of ALTRun from <appDir>
; and puts back the Startup / Send to / Start menu shortcuts named ALTRun.lnk that were there before. Each case starts ALTRun as a
; user does (no arguments) and looks at the real Startup folder:
;   no        new installation (no Data folder): the question shows, and while it waits nothing is in the Startup folder and
;             ALTRun.json has LaunchAtLogin 0; answering No keeps it off and writes no start-up shortcut
;   again     the next start: not asked again, still no start-up shortcut
;   yes       new installation again: Enter (Yes is the default button) turns it on, saves it and writes the start-up shortcut
;   existing  an ALTRun.json that has LaunchAtLogin 1 (an upgrade) and no shortcut: not asked, the shortcut is written as before
;   ini       only the ALTRun.ini of v2026.08.12 (an upgrade from 2.x): not asked
; Prints one line per case (PASS / FAIL, with ALTRun's windows on failure); the exit code is the number of failed cases.
;
;   AutoHotkey64.exe FirstRunCheck.ahk <appDir>       appDir: a copy of the repository (ALTRun.ahk, Lib, Src, Resources)
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut
#Include ..\..\Lib\JSON.ahk

global AppDir := A_Args.Length ? A_Args[1] : A_ScriptDir "\..\.."
global Fixture := A_ScriptDir "\..\Fixtures\ALTRun.v2026.08.12.ini"
global StartupLink := A_Startup "\ALTRun.lnk"
global Shortcuts := [StartupLink, RegExReplace(A_StartMenu, "\\Start Menu$", "\SendTo") "\ALTRun.lnk", A_Programs "\ALTRun.lnk"]
global Kept := Map()                                                      ; shortcuts that were there before: path -> where they wait
global Failures := 0, AppPid := 0
OnError(Crashed)

try {
    Setup()
    RunCase("no", AnswerNo)
    RunCase("again", NotAskedAgain)
    RunCase("yes", AnswerYes)
    RunCase("existing", ExistingSettings)
    RunCase("ini", IniUpgrade)
} catch as setupError {
    Say("FAIL setup: " setupError.Message " (line " setupError.Line ")")
    Failures++
}
TearDown()
Say(Failures ? Failures " case(s) failed" : "all cases passed")
ExitApp(Failures)

Say(text) => FileAppend(text "`n", "*")

; An error anywhere: one FAIL line and out, never a message box that would hold the runner until its time limit
Crashed(e, *) {
    Say("FAIL error: " e.Message " (line " e.Line ")")
    TearDown()
    ExitApp(Failures + 1)
}

Setup() {
    if !FileExist(AppDir "\ALTRun.ahk")
        throw Error("no ALTRun.ahk in " AppDir)
    if !FileExist(Fixture)
        throw Error("no " Fixture)
    for path in Shortcuts {                                                 ; the user's own shortcuts wait aside until the end
        if FileExist(path) {
            FileMove(path, path ".firstrun-check", true)
            Kept[path] := path ".firstrun-check"
        }
    }
    DetectHiddenWindows(false)
    SetTitleMatchMode(3)
    Say("INFO Startup folder: " A_Startup)
}

; Runs one case with a clean ALTRun: the case returns "" when it passed, otherwise what went wrong
RunCase(name, check) {
    global Failures
    problem := ""
    try problem := check.Call()
    catch as e
        problem := "error: " e.Message " (line " e.Line ")"
    Say((problem = "" ? "PASS " : "FAIL ") name (problem = "" ? "" : ": " problem))
    if (problem != "") {
        Failures++
        DumpWindows()
    }
    StopApp()
}

AnswerNo() {
    Fresh()
    if !(dialog := StartAndWaitForQuestion())
        return "the question didn't show in 30 s"
    Say("INFO question: " StrReplace(Trim(WinGetText(dialog), "`r`n"), "`r`n", " | "))
    if FileExist(StartupLink)
        return "a start-up shortcut was written before the answer"
    if ((value := Stored()) != 0)
        return "LaunchAtLogin before the answer: " value " (expected 0)"
    ControlClick("Button2", dialog)                                         ; No (Button1 is Yes)
    if !WinWaitClose(dialog, , 5)
        return "the question didn't close after No"
    Sleep(1500)
    if FileExist(StartupLink)
        return "No wrote a start-up shortcut"
    if ((value := Stored()) != 0)
        return "LaunchAtLogin after No: " value " (expected 0)"
    return ""
}

NotAskedAgain() {
    if (problem := StartAndExpectNoQuestion())
        return problem
    if FileExist(StartupLink)
        return "a start-up shortcut was written"
    return ""
}

AnswerYes() {
    Fresh()
    if !(dialog := StartAndWaitForQuestion())
        return "the question didn't show in 30 s"
    WinActivate(dialog)
    if !WinWaitActive(dialog, , 5)
        return "the question didn't come to the front"
    Send("{Enter}")                                                         ; the default button: Yes
    if !WinWaitClose(dialog, , 5)
        return "the question didn't close after Enter"
    if !WaitForFile(StartupLink, 5000)
        return "Yes wrote no start-up shortcut in 5 s"
    if ((value := Stored()) != 1)
        return "LaunchAtLogin after Yes: " value " (expected 1)"
    FileGetShortcut(StartupLink, , , &arguments)
    if !InStr(arguments, "-Startup")
        return "the start-up shortcut's arguments: " arguments
    return ""
}

ExistingSettings() {
    Fresh()
    settings := Map("SchemaVersion", 7, "General", Map("Language", "en", "LaunchAtLogin", 1, "CheckForUpdates", 0))
    FileAppend(JSON.Stringify(settings, 4), AppDir "\Data\ALTRun.json", "UTF-8")
    if (problem := StartAndExpectNoQuestion())
        return problem
    if !WaitForFile(StartupLink, 5000)
        return "the start-up shortcut wasn't written (LaunchAtLogin 1)"
    return ""
}

IniUpgrade() {
    Fresh()
    FileCopy(Fixture, AppDir "\ALTRun.ini", true)
    try {
        if (problem := StartAndExpectNoQuestion())
            return problem
        if !FileExist(AppDir "\Data\ALTRun.json")
            return "ALTRun.ini wasn't imported"
    } finally {
        StopApp()
        try FileDelete(AppDir "\ALTRun.ini")
    }
    return ""
}

; A new installation: no settings and no shortcuts of an earlier case
Fresh() {
    StopApp()
    try DirDelete(AppDir "\Data", true)
    for path in Shortcuts
        try FileDelete(path)
}

LaunchApp() {
    global AppPid
    Run('"' A_AhkPath '" "' AppDir '\ALTRun.ahk"', AppDir, , &AppPid)      ; no arguments: as a user starts it
}

StartAndWaitForQuestion() {
    LaunchApp()
    return WinWait("ALTRun ahk_class #32770 ahk_pid " AppPid, , 30)
}

; Not asked: the search window shows (start-up is done; the question would come before it) or, after an upgrade that shows a
; notice instead, 15 s pass; no question appears meanwhile
StartAndExpectNoQuestion() {
    LaunchApp()
    began := A_TickCount
    while (A_TickCount - began < 15000) {
        if WinExist("ALTRun ahk_class #32770 ahk_pid " AppPid)
            return "asked: " StrReplace(Trim(WinGetText(), "`r`n"), "`r`n", " | ")
        if WinExist("ALTRun ahk_class AutoHotkeyGUI ahk_pid " AppPid) {
            Sleep(1000)
            if WinExist("ALTRun ahk_class #32770 ahk_pid " AppPid)
                return "asked after the search window showed"
            return ""
        }
        Sleep(200)
    }
    return ""
}

; LaunchAtLogin in the copy's ALTRun.json ("missing" when there is no file)
Stored() {
    settingsFile := AppDir "\Data\ALTRun.json"
    if !FileExist(settingsFile)
        return "missing"
    return JSON.Parse(FileRead(settingsFile, "UTF-8"))["General"]["LaunchAtLogin"]
}

WaitForFile(path, ms) {
    began := A_TickCount
    while (A_TickCount - began < ms) {
        if FileExist(path)
            return true
        Sleep(100)
    }
    return false
}

DumpWindows() {
    if !AppPid
        return
    DetectHiddenWindows(true)
    for w in WinGetList("ahk_pid " AppPid)
        try Say("  window [" WinGetClass(w) "] " WinGetTitle(w) (DllCall("IsWindowVisible", "Ptr", w) ? " (visible)" : ""))
    DetectHiddenWindows(false)
}

StopApp() {
    global AppPid
    if !AppPid
        return
    if WinExist("ALTRun ahk_class #32770 ahk_pid " AppPid)                  ; a question still waiting: the process is closed
        try ProcessClose(AppPid)
    DetectHiddenWindows(true)
    try WinClose("ahk_class AutoHotkey ahk_pid " AppPid)
    DetectHiddenWindows(false)
    if !ProcessWaitClose(AppPid, 10)
        try ProcessClose(AppPid)
    AppPid := 0
}

TearDown() {
    StopApp()
    for path in Shortcuts
        try FileDelete(path)
    for path, aside in Kept
        try FileMove(aside, path, true)
    try DirDelete(AppDir "\Data", true)
}
