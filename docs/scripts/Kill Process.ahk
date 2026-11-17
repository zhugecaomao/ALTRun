; @altrun.title          Kill Process
; @altrun.keyword        kill
; @altrun.mode           filter
; @altrun.filter         altrun
; @altrun.description    Live results: your running programs with their memory, the largest first; Enter ends every process of the chosen program
; @altrun.description_zh 实时结果: 正在运行的程序和占用的内存, 大的在前; 回车结束所选程序的全部进程
; @altrun.network        none
;
; Type "kill" and a space in ALTRun, then part of a program's name (pinyin initials and fuzzy matching work as everywhere in ALTRun).
; Enter ends all processes of that program at once, like Task Manager's "End task" on each: unsaved work in it is lost.
; Only programs in your own Windows session are listed, without the parts of Windows that the desktop needs (csrss, dwm...).
; A program running as administrator can only be ended when ALTRun runs as administrator too.
#Requires AutoHotkey v2.0
#NoTrayIcon

; Ending them would break the desktop or sign-in: never listed
Protected := Map("csrss.exe", 1, "winlogon.exe", 1, "dwm.exe", 1, "fontdrvhost.exe", 1, "sihost.exe", 1, "svchost.exe", 1
    , "ctfmon.exe", 1, "smss.exe", 1, "wininit.exe", 1, "lsass.exe", 1, "services.exe", 1, "lsaiso.exe", 1)
Chinese := RegExMatch(A_Language, "i)^(0804|0404|0C04|1004|1404)$") > 0

if (EnvGet("ALTRUN_ACTION") = "run")                                        ; Enter on a result: end that program
    EndProgram(A_Args.Length ? A_Args[1] : "")
ListPrograms()
ExitApp

; One result per program (the processes of one .exe added up), the most memory first
ListPrograms() {
    programs := Map(), programs.CaseSense := false
    for process in Processes() {
        if !programs.Has(process.Name)
            programs[process.Name] := {Name: process.Name, Count: 0, Memory: 0, Path: ""}
        p := programs[process.Name], p.Count += 1, p.Memory += process.Memory
        if (p.Path = "")
            p.Path := ImagePath(process.Pid)
    }
    sorted := []
    for , p in programs
        sorted.Push(p)
    SortBy(sorted, (a, b) => b.Memory - a.Memory)
    items := ""
    for p in sorted {
        count := Chinese ? p.Count " 个进程" : p.Count (p.Count = 1 ? " process" : " processes")
        items .= (items = "" ? "" : ",`n") '{"title": "' Json(p.Name) '", "subtitle": "' Json(count " · " Megabytes(p.Memory)
            . (Chinese ? " · 回车结束 (未保存的内容会丢失)" : " · Enter ends it (unsaved work is lost)")) '", "arg": "' Json(p.Name)
            . '", "type": "script", "uid": "' Json(StrLower(p.Name)) '"' (p.Path != "" ? ', "icon": "' Json(p.Path) '"' : "") '}'
    }
    FileAppend('{"items": [' items ']}', "*", "UTF-8-RAW")
}

; Ends every process named name in this session; the notification says how many
EndProgram(name) {
    ended := 0, failed := 0
    for process in Processes()
        if (process.Name = name) {
            handle := DllCall("OpenProcess", "UInt", 0x1, "Int", false, "UInt", process.Pid, "Ptr")   ; PROCESS_TERMINATE
            if (handle && DllCall("TerminateProcess", "Ptr", handle, "UInt", 1))
                ended += 1
            else
                failed += 1
            if handle
                DllCall("CloseHandle", "Ptr", handle)
        }
    if (ended = 0 && failed = 0)
        Fail(Chinese ? name " 已经不在运行" : name " isn't running any more")
    if (ended = 0)
        Fail(Chinese ? "不能结束 " name " (它以管理员身份运行?)" : "Can't end " name " (is it running as administrator?)")
    FileAppend(Chinese ? "已结束 " name " (" ended " 个进程" (failed ? ", " failed " 个不能结束" : "") ")"
        : "Ended " name " (" ended (ended = 1 ? " process" : " processes") (failed ? ", " failed " couldn't be ended" : "") ")", "*", "UTF-8-RAW")
    ExitApp
}

; The processes of this Windows session except this script and the protected ones: [{Pid, Name, Memory (working set, bytes)}]
Processes() {
    static TH32CS_SNAPPROCESS := 0x2
    entrySize := (A_PtrSize = 8) ? 568 : 556, nameAt := (A_PtrSize = 8) ? 44 : 36   ; PROCESSENTRY32W
    session := 0, DllCall("ProcessIdToSessionId", "UInt", DllCall("GetCurrentProcessId"), "UInt*", &session)
    snapshot := DllCall("CreateToolhelp32Snapshot", "UInt", TH32CS_SNAPPROCESS, "UInt", 0, "Ptr")
    list := [], entry := Buffer(entrySize, 0), NumPut("UInt", entrySize, entry)
    more := DllCall("Process32FirstW", "Ptr", snapshot, "Ptr", entry)
    while more {
        pid := NumGet(entry, 8, "UInt"), name := StrGet(entry.Ptr + nameAt, 260, "UTF-16")
        theirs := -1
        if (pid != DllCall("GetCurrentProcessId") && !Protected.Has(StrLower(name))
            && DllCall("ProcessIdToSessionId", "UInt", pid, "UInt*", &theirs) && theirs = session)
            list.Push({Pid: pid, Name: name, Memory: WorkingSet(pid)})
        more := DllCall("Process32NextW", "Ptr", snapshot, "Ptr", entry)
    }
    DllCall("CloseHandle", "Ptr", snapshot)
    return list
}

WorkingSet(pid) {                                   ; bytes, 0 when the process can't be opened
    handle := DllCall("OpenProcess", "UInt", 0x1000 | 0x10, "Int", false, "UInt", pid, "Ptr")   ; PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_VM_READ
    if !handle
        return 0
    counters := Buffer(A_PtrSize = 8 ? 72 : 40, 0)   ; PROCESS_MEMORY_COUNTERS
    ok := DllCall("K32GetProcessMemoryInfo", "Ptr", handle, "Ptr", counters, "UInt", counters.Size)
    DllCall("CloseHandle", "Ptr", handle)
    return ok ? NumGet(counters, A_PtrSize = 8 ? 16 : 12, "UPtr") : 0
}

ImagePath(pid) {                                    ; the program's file (for its icon), "" when it can't be read
    handle := DllCall("OpenProcess", "UInt", 0x1000, "Int", false, "UInt", pid, "Ptr")
    if !handle
        return ""
    size := 1024, path := Buffer(size * 2, 0)
    ok := DllCall("QueryFullProcessImageNameW", "Ptr", handle, "UInt", 0, "Ptr", path, "UInt*", &size)
    DllCall("CloseHandle", "Ptr", handle)
    return ok ? StrGet(path, size, "UTF-16") : ""
}

Megabytes(bytes) => (bytes >= 1073741824) ? Format("{:.1f} GB", bytes / 1073741824) : Round(bytes / 1048576) " MB"

SortBy(list, compare) {                             ; insertion sort: a few hundred programs at most
    Loop list.Length - 1 {
        i := A_Index + 1, item := list[i]
        while (i > 1 && compare(list[i - 1], item) > 0)
            list[i] := list[i - 1], i -= 1
        list[i] := item
    }
}

Json(s) {                                           ; text -> inside a JSON string
    s := StrReplace(StrReplace(s, "\", "\\"), '"', '\"')
    s := StrReplace(StrReplace(StrReplace(s, "`r", "\r"), "`n", "\n"), "`t", "\t")
    while RegExMatch(s, "[\x01-\x1F]", &c)
        s := StrReplace(s, c[0], Format("\u{:04X}", Ord(c[0])))
    return s
}

Fail(message) {                                     ; ALTRun shows the first line of standard error as the reason
    FileAppend(message "`n", "**", "UTF-8-RAW")
    ExitApp(1)
}
