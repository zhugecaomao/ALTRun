; Memory of ALTRun broken down by part (experiment, not part of the build): each variant is a fresh copy of ALTRun with its own settings;
; after it has settled, the private working set (what Task Manager shows), the working set and the commit charge are recorded
#Requires AutoHotkey v2.0
#SingleInstance Off
#Include ..\..\Lib\JSON.ahk
if A_Args.Length {
    DataMode(A_Args[1])
    Persistent()
    return
}
Root := A_ScriptDir "\..\..", Work := "C:\MemProbe", DataFolder := "C:\MemData"
Ahk := A_AhkPath
Out := ""
Variants := [
    {Name: "#289 (arrays) default settings", Kind: "app", Off: ["Clipboard"], Src: "v289"},
    {Name: "before #311 default settings", Kind: "app", Off: ["Clipboard"], Src: "base"},
    {Name: "after #311 default settings", Kind: "app", Off: ["Clipboard"], Src: "new"},
    {Name: "#289 (arrays) 30,000 files scanned and saved", Kind: "app", Off: ["Clipboard"], Files: true, Src: "v289"},
    {Name: "#289 (arrays) restarted, index from cache", Kind: "app", Off: ["Clipboard"], Files: true, Keep: true, Src: "v289"},
    {Name: "before #311 30,000 files scanned and saved", Kind: "app", Off: ["Clipboard"], Files: true, Src: "base"},
    {Name: "before #311 restarted, index from cache", Kind: "app", Off: ["Clipboard"], Files: true, Keep: true, Src: "base"},
    {Name: "after #311 30,000 files scanned and saved", Kind: "app", Off: ["Clipboard"], Files: true, Src: "new"},
    {Name: "after #311 restarted, index from cache", Kind: "app", Off: ["Clipboard"], Files: true, Keep: true, Src: "new"},
]
Loop 2 {
    pass := A_Index
    for v in Variants {
        m := RunVariant(v)
        line := Format("{}|round {}|private WS {:.1f} MB|WS {:.1f} MB|commit {:.1f} MB|peak WS {:.1f} MB|peak commit {:.1f} MB", v.Name, pass, m.Private, m.WS, m.Commit, m.PeakWS, m.PeakCommit)
        Out .= line "`n"
        FileAppend(line "`n", "*")
    }
}
ExitApp

RunVariant(v) {
    if (v.Kind = "data") {
        Run('"' Ahk '" "' A_ScriptFullPath '" ' v.Mode, , , &pid)
        Sleep(15000)
        return Finish(pid)
    }
    if (v.HasOwnProp("Keep") && IndexSaved()) {    ; restart the same copy: the index comes from its cache
        Run('"' Ahk '" "' Work '\ALTRun.ahk" -Startup', Work, , &pid)
        Sleep(60000)                                                        ; before #311 reading a 30,000-entry cache took about 20 s
        return Finish(pid)
    }
    try DirDelete(Work, true)
    try DirDelete(A_AppData "\ALTRun", true)
    try DirDelete(EnvGet("LOCALAPPDATA") "\ALTRun", true)
    if (v.Kind = "bare") {
        Run('"' Ahk '" "' A_ScriptFullPath '" bare', , , &pid)
        Sleep(10000)
        return Finish(pid)
    }
    DirCreate(Work)
    src := Root "\" v.Src
    FileCopy(src "\ALTRun.ahk", Work "\ALTRun.ahk")
    for d in ["Lib", "Src", "Resources"]
        DirCopy(src "\" d, Work "\" d)
    if (v.Kind = "code") {
        text := RTrim(FileRead(Work "\ALTRun.ahk", "UTF-8"), "`r`n `t")
        text := RegExReplace(text, "App\.Start\(\)$", "Persistent()")
        FileDelete(Work "\ALTRun.ahk"), FileAppend(text, Work "\ALTRun.ahk", "UTF-8-RAW")
        Run('"' Ahk '" "' Work '\ALTRun.ahk"', Work, , &pid)
        Sleep(15000)
        return Finish(pid)
    }
    features := Map()
    for name in v.Off
        features[name] := Map("Enabled", 0)
    if v.HasOwnProp("Files")
        features["FileSearch"] := Map("Enabled", 1, "UseEverything", 0, "ScopeFolders", [DataFolder])
    settings := Map("SchemaVersion", 6,
        "General", Map("Language", "en", "LaunchAtLogin", 0, "HideOnDeactivate", 1, "SendToMenu", 0, "StartMenuShortcut", 0,
                       "CheckForUpdates", 0, "Hotkey", "!Space", "ShowTips", 0),
        "Features", features)
    DirCreate(Work "\Data")
    FileAppend(JSON.Stringify(settings, 4), Work "\Data\ALTRun.json", "UTF-8")
    Run('"' Ahk '" "' Work '\ALTRun.ahk" -Startup', Work, , &pid)
    Sleep(20000)                                                            ; start-up, icon prefetch, application scan
    if v.HasOwnProp("Files") {                                              ; until the file index is scanned and saved
        start := A_TickCount
        while (!IndexSaved() && A_TickCount - start < 300000)
            Sleep(1000)
    }
    if v.HasOwnProp("Clips") {
        Loop v.Clips {
            A_Clipboard := "Clipboard entry " A_Index " " StrReplace(Format("{:80}", ""), " ", "x") " " A_TickCount
            Sleep(80)
        }
    }
    Sleep(15000)                                                            ; let saves and background work finish
    return Finish(pid)
}

Finish(pid) {
    m := Mem(pid)
    DetectHiddenWindows(true)
    try WinClose("ahk_class AutoHotkey ahk_pid " pid)
    if !ProcessWaitClose(pid, 15)
        ProcessClose(pid)
    Sleep(2000)
    return m
}

; CPU cycles the process uses in one idle minute (nothing typed, another window in front), in millions
IdleCycles(pid) {
    h := DllCall("OpenProcess", "UInt", 0x1000, "Int", 0, "UInt", pid, "Ptr")   ; PROCESS_QUERY_LIMITED_INFORMATION
    if !h
        return -1
    DllCall("QueryProcessCycleTime", "Ptr", h, "Int64*", &before := 0)
    Sleep(60000)
    DllCall("QueryProcessCycleTime", "Ptr", h, "Int64*", &after := 0)
    DllCall("CloseHandle", "Ptr", h)
    return (after - before) / 1e6
}

Mem(pid) {
    h := DllCall("OpenProcess", "UInt", 0x0410, "Int", 0, "UInt", pid, "Ptr")   ; PROCESS_QUERY_INFORMATION | PROCESS_VM_READ
    if !h
        return {Private: -1, WS: -1, Commit: -1, PeakWS: -1, PeakCommit: -1}
    pmc := Buffer(80, 0), NumPut("UInt", 80, pmc)
    DllCall("psapi\GetProcessMemoryInfo", "Ptr", h, "Ptr", pmc, "UInt", 80)
    ws := NumGet(pmc, 16, "UPtr"), commit := NumGet(pmc, 72, "UPtr"), peakWS := NumGet(pmc, 8, "UPtr"), peakCommit := NumGet(pmc, 64, "UPtr")
    size := 1 << 20                                                         ; private working set: working set pages that aren't shared (bit 8 of each entry)
    Loop 5 {
        buf := Buffer(size, 0)
        if DllCall("psapi\QueryWorkingSet", "Ptr", h, "Ptr", buf, "UInt", size)
            break
        size := (NumGet(buf, 0, "UPtr") + 8192) * A_PtrSize
    }
    n := NumGet(buf, 0, "UPtr"), private := 0
    Loop n
        if !(NumGet(buf, A_Index * A_PtrSize, "UPtr") & 0x100)
            private++
    DllCall("CloseHandle", "Ptr", h)
    return {Private: private * 4096 / 1048576, WS: ws / 1048576, Commit: commit / 1048576, PeakWS: peakWS / 1048576, PeakCommit: peakCommit / 1048576}
}

; The index data alone, built the way FileIndex does it, in a bare process
DataMode(mode) {
    global Paths, Names, Folders, Extra
    folder := "C:\MemData"
    if (mode = "bare")
        return
    if (mode = "load") {
        data := JSON.Parse(FileRead("C:\MemCache\FileIndex.json", "UTF-8"))
        Paths := data["Paths"], Folders := data["Folders"], Names := []
        for fullPath in Paths {
            SplitPath(fullPath, &fileName)
            Names.Push(StrLower(fileName))
        }
        data := ""
        return
    }
    if (mode = "compact") {                                                 ; each folder once; per entry: folder number (negative = a folder) and the name as it is
        dirIds := Map(), Dirs := [], Ids := [], Names := []
        Loop Files, folder "\*", "FDR" {
            if !dirIds.Has(A_LoopFileDir)
                Dirs.Push(A_LoopFileDir), dirIds[A_LoopFileDir] := Dirs.Length
            id := dirIds[A_LoopFileDir]
            Ids.Push(InStr(A_LoopFileAttrib, "D") ? -id : id), Names.Push(A_LoopFileName)
        }
        Paths := Dirs, Folders := Ids
        return
    }
    if (mode = "packed") {
        Packed(folder)
        return
    }
    Paths := [], Names := [], Folders := []
    Loop Files, folder "\*", "FDR"
        Paths.Push(A_LoopFileFullPath), Names.Push(StrLower(A_LoopFileName)), Folders.Push(InStr(A_LoopFileAttrib, "D") ? 1 : 0)
    if (mode = "scan") {                                                    ; the search as FileIndex does it now: every name scored in a loop
        t := Timed((needle) => ArraySearch(Names, needle))
        DirCreate("C:\MemCache"), FileAppend("search, arrays: " t "`n", "C:\MemCache\times.txt")
    }
    if (mode = "scansave")
        Extra := JSON.Stringify(Map("Paths", Paths, "Folders", Folders)), Extra := ""
}

; ms per search for each needle (20 times each) and its number of matches
Timed(search) {
    static Needles := ["doc", "document 15-", "-99", "xyz"]
    out := ""
    for needle in Needles {
        n := search(needle), start := QPC()
        Loop 20
            search(needle)
        out .= Format("{}={} in {:.2f} ms; ", needle, n, (QPC() - start) / 20)
    }
    return out
}

QPC() {
    static freq := 0
    if !freq
        DllCall("QueryPerformanceFrequency", "Int64*", &freq)
    DllCall("QueryPerformanceCounter", "Int64*", &now := 0)
    return now * 1000 / freq
}

ArraySearch(names, needle) {
    n := 0
    for index, name in names
        if ScoreName(needle, name)
            n++
    return n
}

ScoreName(needle, name) {
    if !(pos := InStr(name, needle))
        return 0
    penalty := Min(StrLen(name) - StrLen(needle), 40) * 0.1
    if (StrLen(name) = StrLen(needle))
        return 100
    if (pos = 1)
        return 90 - penalty
    if InStr(" -_.()[]", SubStr(name, pos - 1, 1))
        return 80 - penalty
    return 60 - Min(pos, 20) * 0.25 - penalty
}

; All lower-case names in one string ("`n" before each), the names as they are in another; a Buffer holds per entry its start in the
; strings (Int32) and its folder number (Int32, negative = a folder); each folder path once
Packed(folder) {
    global Lower, Shown, Info, Dirs
    lowerText := "", shownText := "", count := 0, dirIds := Map(), Dirs := [], starts := []
    Loop Files, folder "\\*", "FDR"
        count++
    Info := Buffer(count * 8), i := 0, pos := 1
    Loop Files, folder "\\*", "FDR" {
        if !dirIds.Has(A_LoopFileDir)
            Dirs.Push(A_LoopFileDir), dirIds[A_LoopFileDir] := Dirs.Length
        id := dirIds[A_LoopFileDir]
        NumPut("Int", pos, Info, i * 8), NumPut("Int", InStr(A_LoopFileAttrib, "D") ? -id : id, Info, i * 8 + 4)
        lowerText .= "`n" StrLower(A_LoopFileName), shownText .= "`n" A_LoopFileName
        pos += StrLen(A_LoopFileName) + 1, i++
    }
    Lower := lowerText "`n", Shown := shownText "`n", lowerText := "", shownText := ""
    t := Timed((needle) => PackedSearch(needle, count))
    DirCreate("C:\MemCache"), FileAppend("search, packed: " t "`n", "C:\MemCache\times.txt")
}

; every place the needle occurs, found with InStr in the long string. The entry: a few steps forward from the last one (matches come in
; order), a binary search when it is further. Scored from the offsets, without cutting the name out; the rest of that entry is skipped
PackedSearch(needle, count) {
    global Lower, Info
    n := 0, at := 1, len := StrLen(needle), entry := 0, total := StrLen(Lower)
    while (at := InStr(Lower, needle, , at)) {
        steps := 0
        while (entry + 1 < count && NumGet(Info, (entry + 1) * 8, "Int") < at && ++steps <= 4)
            entry++
        if (steps > 4) {
            lo := entry, hi := count - 1
            while (lo < hi) {
                mid := (lo + hi + 1) >> 1
                if (NumGet(Info, mid * 8, "Int") < at)
                    lo := mid
                else
                    hi := mid - 1
            }
            entry := lo
        }
        start := NumGet(Info, entry * 8, "Int") + 1                         ; the name's first character
        next := (entry + 1 < count) ? NumGet(Info, (entry + 1) * 8, "Int") : total
        nameLen := next - start, pos := at - start + 1
        if (pos = 1 && nameLen = len)
            score := 100
        else {
            penalty := Min(nameLen - len, 40) * 0.1
            score := (pos = 1) ? 90 - penalty : InStr(" -_.()[]", SubStr(Lower, at - 1, 1)) ? 80 - penalty : 60 - Min(pos, 20) * 0.25 - penalty
        }
        n++
        if (++entry >= count)
            break
        at := NumGet(Info, entry * 8, "Int") + 1                            ; on to the next entry
    }
    return n
}

IndexSaved() => FileExist(Work "\Data\FileIndex.json") || FileExist(Work "\Data\FileIndex.txt") || FileExist(Work "\Data\FileIndex.bin")
