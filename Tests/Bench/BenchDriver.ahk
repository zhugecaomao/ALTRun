; BenchDriver.ahk - simulates daily use on Windows to compare the performance of two versions (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Run by .github/workflows/bench.yml (GitHub → Actions → Bench → Run workflow); it can also run on your own PC
; (it takes over the keyboard and screen for tens of minutes; a virtual machine is recommended). Each version is copied to C:\Bench\<name>, and ALTRun.ahk gets
; #Include BenchProbe.ahk just before App.Start() at the end (only in the test copies; it answers "are the results drawn?").
;
;   BenchDriver.ahk demo                                               creates the sample files in C:\BenchDemo
;   BenchDriver.ahk prepare <appDir>                                   writes the test settings (the same for every version)
;   BenchDriver.ahk round <label> <appDir> <outFile> <seed> <searches> <others>
;
; One round: start ALTRun, randomly mix searches searches and others each of running programs / opening files and folders, the actions panel, calculator,
; Quick Switch, opening and saving Preferences, snippet expansion (before them, three idle minutes measure the CPU used while idle); each timing goes into outFile (one line label|item|ms), and at the end
; CPU, memory, GDI / USER objects and handle counts are noted before a normal exit. Input uses real key presses (SendLevel 1, so hotkeys and hotstrings respond),
; and a "BenchEdit" window is always in front (normally some program is in front when the hotkey is pressed). bench_stats.py summarizes it into a table.
#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Include ..\..\Lib\JSON.ahk

OnError(DriverError)
DriverError(e, mode) {
    try FileAppend("driver error: " e.Message " line " e.Line " " e.What "`n", A_Temp "\BenchDriverErrors.txt")
    return -1                                                               ; go on (this item counts as failed)
}
DllCall("winmm\timeBeginPeriod", "UInt", 1)
SendLevel(1)
SetKeyDelay(0, 0)
SetTitleMatchMode(3)
DetectHiddenWindows(false)

global Demo := "C:\BenchDemo", Ahk := A_AhkPath, LogFile := A_Temp "\ALTRun.log"   ; ALTRun's log (Logger.File)
switch A_Args[1] {
    case "prepare": Prepare(A_Args[2])
    case "demo":    MakeDemo()
    case "round":   Round_(A_Args[2], A_Args[3], A_Args[4], Integer(A_Args[5]), Integer(A_Args[6]), Integer(A_Args[7]))
}
ExitApp(0)

;-------------------------------------------------------------------------------
; Preparation
;-------------------------------------------------------------------------------
MakeDemo() {
    try DirDelete(Demo, true)
    for folder in ["Folder A", "Folder B", "Start", "Projects", "Reports 2026", "Photos"]
        DirCreate(Demo "\" folder)
    FileAppend("Bench notes`n", Demo "\notes.txt")
    words := ["Budget", "Report", "Invoice", "Meeting", "Design", "Plan", "Summary", "Contract", "Proposal", "Schedule"]
    exts := [".txt", ".docx", ".xlsx", ".pdf", ".png", ".csv", ".pptx", ".zip"]
    folders := ["Projects", "Reports 2026", "Photos", "Folder A", "Folder B"]
    Loop 400 {
        name := words[Mod(A_Index, words.Length) + 1] " " (2000 + A_Index) exts[Mod(A_Index * 7, exts.Length) + 1]
        FileAppend("x", Demo "\" folders[Mod(A_Index * 3, folders.Length) + 1] "\" name)
    }
}

Prepare(appDir) {
    commands := [
        Map("Title", "Bench Notepad", "Type", "File", "Target", A_WinDir "\System32\notepad.exe", "Arguments", "", "Keyword", "bnp"),
        Map("Title", "Bench Notes", "Type", "File", "Target", Demo "\notes.txt", "Arguments", "", "Keyword", "bnt"),
        Map("Title", "Bench Folder", "Type", "Folder", "Target", Demo "\Folder B", "Arguments", "", "Keyword", "bfo")
    ]
    Loop Files A_WinDir "\System32\*.exe" {                                 ; 40 more, so there are plenty of custom commands as usual
        if (A_Index > 40)
            break
        SplitPath(A_LoopFileName, , , , &name)
        commands.Push(Map("Title", "Tool " name, "Type", "File", "Target", A_LoopFileFullPath, "Arguments", "", "Keyword", ""))
    }
    snippets := [
        Map("Name", "Bench signature", "Keyword", "bsig", "Text", "Best regards, Bench", "AutoExpand", 1),
        Map("Name", "Bench address", "Keyword", "badr", "Text", "1 Bench Road, Test City", "AutoExpand", 1),
        Map("Name", "Bench thanks", "Keyword", "bthx", "Text", "Thank you very much!", "AutoExpand", 1)
    ]
    Loop 30
        snippets.Push(Map("Name", "Filler snippet " A_Index, "Keyword", "fill" A_Index, "Text", "Filler text number " A_Index, "AutoExpand", 1))
    settings := Map(
        "SchemaVersion", 4,
        "General", Map("Language", "en", "LaunchAtLogin", 0, "HideOnDeactivate", 1, "SendToMenu", 0, "StartMenuShortcut", 0,
                       "CheckForUpdates", 0, "Hotkey", "!Space", "ShowTips", 0, "SaveLog", 1),   ; Log: slow background tasks, errors (see HarvestLog)
        "Features", Map("FileSearch", Map("UseEverything", 0, "ScopeFolders", [Demo])),
        "CustomCommands", commands,
        "Snippets", snippets
    )
    try DirDelete(appDir "\Data", true)
    DirCreate(appDir "\Data")
    FileAppend(JSON.Stringify(settings, 4), appDir "\Data\ALTRun.json", "UTF-8")
}

;-------------------------------------------------------------------------------
; One round: start ALTRun, randomly mix all kinds of actions, note each timing and the final resource use, exit normally
;-------------------------------------------------------------------------------
class B {
    static Label := "", Out := "", Pid := 0, Main := 0, Search := 0, Seed := 1, Editor := "", Explorer := 0
    static Words := [], Files := []
}

NowMs() {
    static freq := 0
    if !freq
        DllCall("QueryPerformanceFrequency", "Int64*", &freq)
    DllCall("QueryPerformanceCounter", "Int64*", &now := 0)
    return now * 1000 / freq
}

Rand(n) {                                                                   ; 0 .. n-1, fixed seed, the same for both versions
    B.Seed := Mod(B.Seed * 1103515245 + 12345, 2147483648)
    return Mod(B.Seed // 65536, n)
}

Rec(op, value) => FileAppend(B.Label "|" op "|" (IsNumber(value) ? Round(value, 2) : value) "`n", B.Out, "UTF-8")

Probe(code, l := 0) {                                                       ; use the handle directly (ALTRun's main window is hidden; DetectHiddenWindows doesn't matter)
    result := 0
    if !DllCall("SendMessageTimeoutW", "Ptr", B.Main, "UInt", 0x8123, "Ptr", code, "Ptr", l, "UInt", 0, "UInt", 10000, "Ptr*", &result)
        return -1
    return result
}

WaitFor(fn, timeout := 5000) {                                              ; Waits until fn() is true and returns the milliseconds; "" on timeout
    start := NowMs()
    loop {
        if fn()
            return NowMs() - start
        if (NowMs() - start > timeout)
            return ""
        Sleep(5)                                                            ; asking too often keeps ALTRun busy answering, delaying its timers and repaints
    }
}

Visible() => DllCall("IsWindowVisible", "Ptr", B.Search)

Round_(label, appDir, outFile, seed, searches, others) {
    B.Label := label, B.Out := outFile, B.Seed := seed
    CollectWords()
    for f in [LogFile, LogFile ".old"]
        try FileDelete(f)
    start := NowMs()
    Run('"' Ahk '" "' appDir '\ALTRun.ahk" -Startup', appDir, , &pid)
    B.Pid := pid
    DetectHiddenWindows(true)
    SetTitleMatchMode(2)
    B.Main := WinWait("ALTRun.ahk ahk_class AutoHotkey ahk_pid " pid, , 30)
    SetTitleMatchMode(3)
    if !B.Main {
        Rec("fatal", "main window not found")
        return
    }
    WaitFor(() => Probe(2) = 1, 30000)
    Rec("startup until idle ms", NowMs() - start)
    WaitFor(() => B.Search := WinExist("ALTRun ahk_class AutoHotkeyGUI ahk_pid " pid), 10000)
    if !B.Search {
        Rec("fatal", "search window not found")
        return
    }
    DetectHiddenWindows(false)
    Sleep(6000)                                                             ; after start-up warm-up / prefetch is done (nobody searches right at start-up either)
    Resources("start")

    B.Editor := Gui("+AlwaysOnTop", "BenchEdit")                            ; imitates a program in use (also for testing snippet expansion)
    B.Editor.Add("Edit", "w420 h60 vText")
    B.Editor.Show("x0 y0 NoActivate")
    Run('explorer.exe "' Demo '\Folder A"')                                 ; the target of Quick Switch jumps
    B.Explorer := WinWait("Folder A ahk_class CabinetWClass", , 15)
    IdleCpu()

    ops := []
    Loop searches
        ops.Push("search")
    for op in ["run", "action", "calc", "quickswitch", "prefs", "snippet"]
        Loop others
            ops.Push(op)
    Loop ops.Length - 1 {                                                   ; shuffle (Fisher-Yates)
        i := ops.Length - A_Index + 1, j := Rand(i) + 1
        t := ops[i], ops[i] := ops[j], ops[j] := t
    }
    for op in ops {
        try {
            %"Op_" op%()
        } catch as e {
            Rec(op " fail", e.Message)
            Reset()
        }
    }
    Resources("end")
    try WinClose(B.Explorer)
    try B.Editor.Destroy()
    DetectHiddenWindows(true)
    WinClose("ahk_id " B.Main)                                              ; normal exit (saves the learned ranking and icon cache)
    if !ProcessWaitClose(pid, 15)
        ProcessClose(pid)
    HarvestLog()
}

; From ALTRun's log: background tasks over 50 ms ("Perf: slow <task> N ms", in versions that log it) are recorded as "bg <task>",
; errors as "log error", the icon cache save time and steps over 200 ms after start-up as "note"
HarvestLog() {
    try text := FileRead(LogFile, "UTF-8")
    catch
        return
    Loop Parse, text, "`n", "`r" {
        if RegExMatch(A_LoopField, "\[DBG\] Perf: slow (.+) (\d+) ms$", &m)
            Rec("bg " m[1], m[2])
        else if RegExMatch(A_LoopField, "\[ERR\] (.*)$", &m)
            Rec("log error", SubStr(StrReplace(m[1], "|", "/"), 1, 120))
        else if RegExMatch(A_LoopField, "\[DBG\] Perf: (icon cache saved .*)$", &m)
            Rec("note", m[1])
        else if (RegExMatch(A_LoopField, "^\S+ (\S+) \[DBG\] Perf: (.+ (\d+) ms)$", &m) && m[3] >= 200
              && !RegExMatch(m[2], "^(startup|warm-up|empty results|icons prefetched|icon cache loaded|font warm-up|prepare)"))
            Rec("note", m[1] " " m[2])                                       ; a very slow step after start-up (showing the window, searching...): compare the times to find the cause
    }
}

CollectWords() {
    names := ""
    for folder in [A_ProgramsCommon, A_Programs]
        Loop Files folder "\*.lnk", "R" {
            SplitPath(A_LoopFileName, , , , &name)
            names .= name "`n"
        }
    for name in StrSplit(Sort(RTrim(names, "`n"), "U"), "`n")
        if (name != "")
            B.Words.Push(name)
    Loop Files Demo "\*.*", "R"
        B.Files.Push(A_LoopFileName)
}

Resources(when) {
    h := DllCall("OpenProcess", "UInt", 0x1410, "Int", 0, "UInt", B.Pid, "Ptr")   ; QUERY_INFORMATION | VM_READ | QUERY_LIMITED
    if !h
        return
    c := 0, e := 0, k := 0, u := 0
    DllCall("GetProcessTimes", "Ptr", h, "Int64*", &c, "Int64*", &e, "Int64*", &k, "Int64*", &u)
    mem := Buffer(80, 0), NumPut("UInt", 80, mem)
    DllCall("psapi\GetProcessMemoryInfo", "Ptr", h, "Ptr", mem, "UInt", 80)
    handles := 0
    DllCall("GetProcessHandleCount", "Ptr", h, "UInt*", &handles)
    Rec("res " when " cpu ms", (k + u) / 10000)
    Rec("res " when " working set MB", NumGet(mem, 16, "UPtr") / 1048576)
    Rec("res " when " private MB", NumGet(mem, 72, "UPtr") / 1048576)
    Rec("res " when " GDI objects", DllCall("GetGuiResources", "Ptr", h, "UInt", 0))
    Rec("res " when " USER objects", DllCall("GetGuiResources", "Ptr", h, "UInt", 1))
    Rec("res " when " handles", handles)
    DllCall("CloseHandle", "Ptr", h)
}

; CPU ALTRun uses with another program in front and nothing typed (timers, polling): Samples one-minute samples in a row, each recorded,
; so background work that happens to fall into one (start-up prefetch, a cache write) shows as one outlier and the median stays meaningful.
; Counted in CPU cycles (QueryProcessCycleTime): process times only advance in 15.6 ms steps, too coarse for a few timers a minute
IdleCpu() {
    static Samples := 3, SampleMs := 60000
    try WinActivate("ahk_id " B.Editor.Hwnd)
    Sleep(2000)                                                             ; let the switch settle
    Loop Samples {
        before := CpuCycles()
        Sleep(SampleMs)
        if (before != "" && (after := CpuCycles()) != "")
            Rec("idle cpu Mcycles per minute", (after - before) / 1e6 * 60000 / SampleMs)
    }
}

CpuCycles() {
    if !(h := DllCall("OpenProcess", "UInt", 0x1000, "Int", 0, "UInt", B.Pid, "Ptr"))   ; PROCESS_QUERY_LIMITED_INFORMATION
        return ""
    cycles := 0
    ok := DllCall("QueryProcessCycleTime", "Ptr", h, "Int64*", &cycles)
    DllCall("CloseHandle", "Ptr", h)
    return ok ? cycles : ""
}

ActiveTitle() {
    try return WinGetTitle("A")
    return "(none)"
}

Work() {
    try {
        WinActivate("ahk_id " B.Editor.Hwnd)
        WinWaitActive("ahk_id " B.Editor.Hwnd, , 1)
    }
}

; Back to a clean state: close Preferences and the actions panel, hide the search window
Reset() {
    if (prefs := WinExist("ALTRun Preferences ahk_pid " B.Pid))
        try WinClose(prefs)
    Loop 3 {
        if !Visible()
            break
        try ControlSend("{Esc}", , "ahk_id " B.Search)
        Sleep(150)
    }
}

;-------------------------------------------------------------------------------
; Basic actions
;-------------------------------------------------------------------------------
Show() {
    if Visible()
        Hide()
    Work()                                                                  ; normally some program is in front when the hotkey is pressed
    start := NowMs()
    Send("!{Space}")
    if (WaitFor(Visible, 3000) = "")
        throw Error("window not shown")
    Rec("show: visible ms", NowMs() - start)
    if (WaitFor(() => Probe(1, 0) = 1, 3000) != "")
        Rec("show: ready ms", NowMs() - start)
    if (WaitFor(() => WinActive("ahk_id " B.Search), 2000) = "")            ; characters typed before it has focus would go into another window
        throw Error("window not active (active: " ActiveTitle() ")")
}

Hide() {
    Send("{Esc}")
    WaitFor(() => !Visible(), 2000)
}

Type(text) {                                                                ; typed character by character, with a short pause between them as usual
    Loop Parse text {
        Send("{Text}" A_LoopField)
        Sleep(30)
    }
}

; from the last character typed until the results are drawn
TypeAndWait(text, metric, length := "") {
    Type(text)
    start := NowMs()
    ms := WaitFor(() => Probe(1, (length = "") ? StrLen(text) : length) = 1, 5000)
    if (ms = "")
        throw Error("results not ready for [" text "]: visible " Visible() " input length " Probe(8) " expected " ((length = "") ? StrLen(text) : length)
            . " SetResults calls " Probe(9) " active [" ActiveTitle() "]")
    Rec(metric, ms)
    if (ms > 300)
        Rec("slow " metric, "[" text "] " Round(ms))
}

RandomQuery() {
    kind := Rand(100)
    if (kind < 50 && B.Words.Length) {                                      ; the start of an application name
        word := B.Words[Rand(B.Words.Length) + 1]
        return SubStr(StrLower(word), 1, 2 + Rand(4))
    }
    if (kind < 65 && B.Words.Length) {                                      ; word initials (vsc -> Visual Studio Code)
        initials := ""
        for part in StrSplit(B.Words[Rand(B.Words.Length) + 1], " ")
            if (part != "")
                initials .= StrLower(SubStr(part, 1, 1))
        if (StrLen(initials) >= 2)
            return SubStr(initials, 1, 4)
    }
    if (kind < 80) {                                                        ; a few random letters
        q := ""
        Loop 2 + Rand(2)
            q .= Chr(97 + Rand(26))
        return q
    }
    if (kind < 90 && B.Files.Length)                                        ; file search: Space + the start of a file name
        return " " SubStr(StrLower(B.Files[Rand(B.Files.Length) + 1]), 1, 3 + Rand(4))
    return ["tool", "bench", "bnp", "bnt", "bfo", "too"][Rand(6) + 1]       ; custom commands
}

;-------------------------------------------------------------------------------
; All kinds of actions
;-------------------------------------------------------------------------------
Op_search() {
    q := RandomQuery()
    Show()
    if (SubStr(q, 1, 1) = " ") {                                            ; file search mode
        Send("{Space}")
        Sleep(60)
        name := SubStr(q, 2)
        TypeAndWait(name, "search (files): last key to results ms", StrLen(name))
    } else {
        TypeAndWait(q, "search: last key to results ms")
    }
    Hide()
}

Op_run() {
    target := [["bnp", "ahk_exe notepad.exe"], ["bnt", "notes ahk_exe notepad.exe"], ["bfo", "Folder B ahk_class CabinetWClass"]][Rand(3) + 1]
    SetTitleMatchMode(2)
    before := Map()
    for hwnd in WinGetList(target[2])
        before[hwnd] := true
    Show()
    TypeAndWait(target[1], "run: search ms")
    start := NowMs(), found := 0
    Send("{Enter}")
    ms := WaitFor(() => NewWindow(target[2], before, &found), 10000)
    SetTitleMatchMode(3)
    if (ms = "")
        throw Error("window for " target[1] " not opened")
    Rec("run " target[1] ": enter to window ms", ms)
    Sleep(300)
    try {
        if InStr(target[2], "CabinetWClass")
            WinClose("ahk_id " found)
        else
            ProcessClose(WinGetPID("ahk_id " found))
    }
    Sleep(300)
}

NewWindow(title, before, &found) {
    for hwnd in WinGetList(title)
        if !before.Has(hwnd)
            return found := hwnd
    return false
}

Op_action() {
    Show()
    TypeAndWait("bench", "action: search ms")
    start := NowMs()
    Send("{Right}")
    ms := WaitFor(() => Probe(3) = 1, 3000)
    if (ms = "")
        throw Error("action panel not shown")
    Rec("action: right to panel ms", ms)
    Send("{Esc}")
    Sleep(50)
    Hide()
}

Op_calc() {
    x := 10 + Rand(990), y := 2 + Rand(98), z := Rand(1000)
    expr := x "*" y "+" z
    Show()
    TypeAndWait(expr, "calc: last key to result ms")
    A_Clipboard := "bench-sentinel"
    start := NowMs()
    Send("{Enter}")
    ms := WaitFor(() => A_Clipboard != "bench-sentinel", 3000)
    if (ms = "")
        throw Error("calc result not copied")
    Rec("calc: enter to copied ms", ms)
    WaitFor(() => !Visible(), 2000)
}

Op_quickswitch() {
    helper := A_Temp "\BenchOpenDialog.ahk"
    if !FileExist(helper)
        FileAppend('#NoTrayIcon`nFileSelect(, "' Demo '\Start", "Open")`n', helper, "UTF-8")
    Run('"' Ahk '" "' helper '"', , , &pid)
    dialog := WinWait("Open ahk_class #32770 ahk_pid " pid, , 10)
    if !dialog
        throw Error("open dialog not shown")
    WinActivate(dialog)
    start := NowMs()
    ms := WaitFor(() => WinExist("ALTRun Quick Switch ahk_pid " B.Pid), 5000)
    if (ms != "")
        Rec("quickswitch: panel ms", ms)
    else
        Rec("quickswitch fail", "panel not shown")
    WinActivate(dialog)
    start := NowMs()
    Send("^g")                                                         ; the jump hotkey (default Ctrl+G since settings version 6)
    ms := WaitFor(() => DialogFolder(dialog, "Folder A"), 5000)
    if (ms != "")
        Rec("quickswitch: ctrl+g to folder ms", ms)
    else
        Rec("quickswitch fail", "folder not changed")
    try WinClose(dialog)
    if !ProcessWaitClose(pid, 3)
        try ProcessClose(pid)
}

DialogFolder(dialog, name) {
    try {
        for ctl in WinGetControls(dialog)
            if (InStr(ctl, "ToolbarWindow32") = 1)
                if InStr(ControlGetText(ctl, dialog), name)
                    return true
    }
    return false
}

Op_prefs() {
    Show()
    start := NowMs()
    Send("^,")
    prefs := 0
    ms := WaitFor(() => (prefs := WinExist("ALTRun Preferences ahk_pid " B.Pid)) && DllCall("IsWindowVisible", "Ptr", prefs), 10000)
    if (ms = "")
        throw Error("preferences not opened")
    Rec("prefs: open ms", ms)
    WaitFor(() => Probe(2) = 1, 10000)
    Rec("prefs: open until idle ms", NowMs() - start)
    Sleep(200)
    start := NowMs()
    ok := 0
    for ctl in WinGetControls("ahk_id " prefs)
        if (InStr(ctl, "Button") = 1 && ControlGetText(ctl, "ahk_id " prefs) == "OK")
            ok := ctl
    if !ok
        throw Error("OK button not found")
    PostMessage(0xF5, 0, 0, ok, "ahk_id " prefs)                            ; BM_CLICK
    ms := WaitFor(() => !WinExist("ahk_id " prefs), 10000)
    if (ms = "")
        throw Error("preferences not closed")
    WaitFor(() => Probe(2) = 1, 10000)
    Rec("prefs: ok to saved and idle ms", NowMs() - start)
    Reset()
}

Op_snippet() {
    snippet := [["bsig", "Best regards, Bench"], ["badr", "1 Bench Road, Test City"], ["bthx", "Thank you very much!"]][Rand(3) + 1]
    if Visible()
        Hide()
    B.Editor["Text"].Value := ""
    Work()
    B.Editor["Text"].Focus()
    Sleep(50)
    start := NowMs()
    Send("{Text};" snippet[1])
    if (WaitFor(() => B.Editor["Text"].Value == snippet[2], 3000) = "")
        throw Error("snippet not expanded: [" B.Editor["Text"].Value "]")
    Rec("snippet: type prefix+keyword to expanded ms", NowMs() - start)
}
