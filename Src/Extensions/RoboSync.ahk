;===============================================================================
; RoboSync.ahk - folder sync tasks run with Windows' robocopy (AutoHotkey v2)
;-------------------------------------------------------------------------------
; A sync task copies or mirrors one folder into another, typically a server share into a local (OneDrive) folder, with
; robocopy, which comes with Windows. This file holds the tasks' model: the robocopy command line made from a task, a
; command line read back into a task (to import what was set up elsewhere, such as a Total Commander start menu entry),
; the safety checks and what robocopy's exit code means. Running the tasks, the editor and search come in later steps.
;
; ALTRun.json, settings key "SyncTasks" (paths may use %environment variables%; {date} / {time} in the log file name):
;   [ { "Name": "Projects", "Source": "\\\\server\\projects", "Destination": "%OneDrive%\\Projects", "Files": "*.*",
;       "Mode": "Mirror",                Mirror = /MIR (deletes in the destination what isn't in the source), Copy = /E, Top = top folder only
;       "Copy": "DAT", "DirCopy": "DA",  /COPY: /DCOPY: (data, attributes, times; no permissions, which OneDrive doesn't keep anyway)
;       "Retries": 1, "Wait": 1, "Threads": 16,                      /R: /W: /MT: (0 = no /MT)
;       "Monitor": 0, "MonitorChanges": 1, "MonitorMinutes": 5,      /MON: /MOT: (keeps running, syncs again after changes / minutes)
;       "ExcludeFiles": ["thumbs.db", "~$*.*"], "ExcludeDirs": ["$Recycle.Bin"],   /XF /XD, per task
;       "LogFile": "%OneDrive%\\Logs\\Projects_{date}.log", "AppendLog": 1,        /LOG+: (AppendLog 0: /LOG:)
;       "Extra": "/NP" } ]                                           other robocopy options, as they are
;
; Usage:
;   RoboSync.NewTask()                       a task with the defaults
;   RoboSync.Build(task, preview)            -> {Program, Args}; preview: /L (lists what it would do, changes nothing), no monitoring
;   RoboSync.Parse(commandLine)              -> {Task, Notes}: a robocopy command line (Total Commander variables converted)
;   RoboSync.Check(task, others)             -> [{Level: "Error" | "Warning", Code, Detail}...]: settings that are wrong or risky
;   RoboSync.CheckSource(task)               -> "" / "SourceMissing" / "SourceEmpty": before a mirror (an empty source would empty the destination)
;   RoboSync.Result(exitCode)                -> {Ok, Copied, Extras, Mismatched, Failed, Fatal}: robocopy's exit code is a bit mask, < 8 = success
;===============================================================================

class RoboSync {
    static Items => AppSettings.Data["SyncTasks"]
    static Program := A_WinDir "\System32\robocopy.exe"

    static NewTask() => Map("Name", "", "Source", "", "Destination", "", "Files", "*.*", "Mode", "Mirror", "Copy", "DAT", "DirCopy", "DA"
        , "Retries", 1, "Wait", 1, "Threads", 16, "Monitor", 0, "MonitorChanges", 1, "MonitorMinutes", 5
        , "ExcludeFiles", ["thumbs.db", "desktop.ini", "~$*.*"], "ExcludeDirs", ["System Volume Information", "$Recycle.Bin"]
        , "LogFile", "", "AppendLog", 1, "Extra", "/NP")

    ;---------------------------------------------------------------------------
    ; Task -> command line
    ;---------------------------------------------------------------------------
    ; The robocopy command line of a task. now: the time for {date} / {time} (YYYYMMDDHH24MISS; empty = now)
    static Build(task, preview := false, now := "") {
        get := (key) => RoboSync._Get(task, key)
        args := [RoboSync.Quote(RoboSync.Expand(get("Source"), now)), RoboSync.Quote(RoboSync.Expand(get("Destination"), now))]
        for pattern in RoboSync.Words(get("Files") = "" ? "*.*" : get("Files"))
            args.Push(RoboSync.Quote(pattern))
        if preview
            args.Push("/L")
        switch get("Mode") {
            case "Mirror": args.Push("/MIR")
            case "Copy": args.Push("/E")
        }
        if (get("Copy") != "")
            args.Push("/COPY:" get("Copy"))
        if (get("DirCopy") != "")
            args.Push("/DCOPY:" get("DirCopy"))
        args.Push("/R:" RoboSync._Number(get("Retries"), 0), "/W:" RoboSync._Number(get("Wait"), 0))
        if (threads := RoboSync._Number(get("Threads"), 0))
            args.Push("/MT:" threads)
        if (get("Monitor") && !preview)                                     ; a preview must end: no monitoring
            args.Push("/MON:" Max(1, RoboSync._Number(get("MonitorChanges"), 1)), "/MOT:" Max(1, RoboSync._Number(get("MonitorMinutes"), 5)))
        for option in [["ExcludeFiles", "/XF"], ["ExcludeDirs", "/XD"]] {
            list := get(option[1])
            if (list is Array && list.Length) {
                args.Push(option[2])
                for item in list
                    args.Push(RoboSync.Quote(item))
            }
        }
        if (get("LogFile") != "")
            args.Push((get("AppendLog") ? "/LOG+:" : "/LOG:") RoboSync.Quote(RoboSync.Expand(get("LogFile"), now)))
        if (Trim(get("Extra")) != "")
            args.Push(Trim(get("Extra")))
        text := ""
        for arg in args
            text .= (text = "" ? "" : " ") arg
        return {Program: RoboSync.Program, Args: text}
    }

    ; %environment variables% and {date} (yyyyMMdd) / {time} (HHmmss) in a path
    static Expand(text, now := "") {
        if InStr(text, "{") {
            text := StrReplace(text, "{date}", FormatTime(now, "yyyyMMdd"))
            text := StrReplace(text, "{time}", FormatTime(now, "HHmmss"))
        }
        return Path.ExpandEnv(text)
    }

    ; An argument as robocopy reads it: quoted when it has spaces (or is empty). A backslash before the closing quote would escape it
    ; ("C:\" becomes C:" and the rest of the line), so a path ending in a backslash gets a second one inside the quotes
    static Quote(arg) {
        if (arg != "" && !RegExMatch(arg, '[\s"]'))
            return arg
        return '"' (SubStr(arg, -1) = "\" ? arg "\" : arg) '"'
    }

    ;---------------------------------------------------------------------------
    ; Command line -> task
    ;---------------------------------------------------------------------------
    ; A robocopy command line (with or without "robocopy" in front) as a task. Total Commander's variables are converted:
    ; %|OneDrive| -> %OneDrive%, %|$DATE:YMD| -> {date}. /L (a preview) isn't kept: previewing is a separate command.
    ; Notes: what was left out or changed, for the import to show
    static Parse(commandLine) {
        task := RoboSync.NewTask(), notes := []
        task["ExcludeFiles"] := [], task["ExcludeDirs"] := [], task["Extra"] := "", task["Mode"] := "Top", task["Threads"] := 0
        task["Copy"] := "", task["DirCopy"] := "", task["Retries"] := 1000000, task["Wait"] := 30   ; robocopy's own defaults when not given
        tokens := RoboSync.Tokens(RoboSync.FromTotalCommander(commandLine, notes))
        if (tokens.Length && RegExMatch(tokens[1], "i)(^|\\)robocopy(\.exe)?$"))
            tokens.RemoveAt(1)
        positional := [], extra := [], list := ""
        for token in tokens {
            if (SubStr(token, 1, 1) != "/") {
                if IsObject(list)
                    list.Push(token)
                else
                    positional.Push(token)
                continue
            }
            list := ""
            option := StrUpper(token)
            if (option = "/XF" || option = "/XD") {
                list := task[option = "/XF" ? "ExcludeFiles" : "ExcludeDirs"]
            } else if (option = "/MIR") {
                task["Mode"] := "Mirror"
            } else if (option = "/E") {
                if (task["Mode"] != "Mirror")
                    task["Mode"] := "Copy"
            } else if (option = "/L") {
                notes.Push("Preview")
            } else if RegExMatch(token, "i)^/(R|W|MON|MOT):(\d+)$", &m) {
                key := Map("R", "Retries", "W", "Wait", "MON", "MonitorChanges", "MOT", "MonitorMinutes")[StrUpper(m[1])]
                task[key] := Integer(m[2])
                if (key = "MonitorChanges" || key = "MonitorMinutes")
                    task["Monitor"] := 1
            } else if RegExMatch(token, "i)^/MT(:(\d+))?$", &m) {
                task["Threads"] := (m[2] = "") ? 8 : Integer(m[2])
            } else if RegExMatch(token, "i)^/(D?COPY):(\w+)$", &m) {
                task[StrUpper(m[1]) = "COPY" ? "Copy" : "DirCopy"] := StrUpper(m[2])
            } else if RegExMatch(token, "i)^/LOG(\+?):(.+)$", &m) {
                task["LogFile"] := m[2], task["AppendLog"] := (m[1] = "+") ? 1 : 0
            } else {
                extra.Push(RoboSync.Quote(token))
            }
        }
        if (positional.Length >= 1)
            task["Source"] := RoboSync.TrimSlash(positional[1])
        if (positional.Length >= 2)
            task["Destination"] := RoboSync.TrimSlash(positional[2])
        if (positional.Length >= 3) {
            files := ""
            Loop positional.Length - 2
                files .= (files = "" ? "" : " ") RoboSync.Quote(positional[A_Index + 2])
            task["Files"] := files
        }
        for item in extra
            task["Extra"] .= (task["Extra"] = "" ? "" : " ") item
        if (positional.Length < 2)
            notes.Push("NoFolders")
        task["Name"] := RoboSync.SuggestName(task["Destination"] != "" ? task["Destination"] : task["Source"])
        return {Task: task, Notes: notes}
    }

    ; Total Commander's %|NAME| (an environment variable) and %|$DATE:YMD| (the date) in the ALTRun form; other date formats are noted
    static FromTotalCommander(text, notes := "") {
        text := RegExReplace(text, "i)%\|\$DATE:YMD\|", "{date}")
        if (RegExMatch(text, "i)%\|\$DATE:[^|]*\|") && IsObject(notes))
            notes.Push("DateFormat")
        return RegExReplace(text, "%\|([A-Za-z_][\w()]*)\|", "%$1%")
    }

    ; The arguments of a command line, as programs started with it read them (CommandLineToArgvW); a dummy program name goes first,
    ; since the first word is read by other rules
    static Tokens(commandLine) {
        tokens := []
        if (Trim(commandLine) = "")
            return tokens
        count := 0
        argv := DllCall("shell32\CommandLineToArgvW", "WStr", "x " commandLine, "Int*", &count, "Ptr")
        if !argv
            return tokens
        Loop count - 1
            tokens.Push(StrGet(NumGet(argv, A_Index * A_PtrSize, "Ptr"), "UTF-16"))
        DllCall("LocalFree", "Ptr", argv)
        return tokens
    }

    ; Space-separated words, with "quoted parts" kept together (the Files field: *.docx "my report.*")
    static Words(text) => RoboSync.Tokens(text)

    static TrimSlash(folder) => RegExMatch(folder, "^[A-Za-z]:\\$") ? folder : RTrim(folder, "\")

    ; A name for an imported task: the last folder of its path
    static SuggestName(folder) {
        folder := RTrim(RegExReplace(folder, "%(\w+)%", "$1"), "\")
        SplitPath(folder, &leaf)
        return leaf
    }

    ;---------------------------------------------------------------------------
    ; Checks
    ;---------------------------------------------------------------------------
    ; Settings that are wrong (Error: don't run) or risky (Warning), without touching the disk. others: the other tasks, for overlaps.
    ; Codes: NoSource, NoDestination, SamePath, Nested (one folder inside the other), MirrorRoot (a mirror into a drive / share root,
    ; the OneDrive folder or the user folder: it would delete everything else there), Threads (1 to 128), Overlap (another task writes
    ; into the same folders; Detail = its name), SameLog (another task writes the same log file; Detail = its name)
    static Check(task, others := "") {
        problems := []
        add(level, code, detail := "") => problems.Push({Level: level, Code: code, Detail: detail})
        source := RoboSync.Normal(RoboSync._Get(task, "Source")), target := RoboSync.Normal(RoboSync._Get(task, "Destination"))
        if (source = "")
            add("Error", "NoSource")
        if (target = "")
            add("Error", "NoDestination")
        if (source != "" && target != "") {
            if (source = target)
                add("Error", "SamePath")
            else if (RoboSync.Inside(source, target) || RoboSync.Inside(target, source))
                add("Error", "Nested")
        }
        if (RoboSync._Get(task, "Mode") = "Mirror" && target != "" && RoboSync.IsRootLike(target))
            add("Error", "MirrorRoot", target)
        threads := RoboSync._Get(task, "Threads")
        if (threads != "" && threads != 0 && (!IsInteger(threads) || threads < 1 || threads > 128))
            add("Error", "Threads", threads)
        if IsObject(others) {
            logFile := RoboSync.Normal(RoboSync.Expand(RoboSync._Get(task, "LogFile"), "20000101"))
            for other in others {
                if (other = task)
                    continue
                name := RoboSync._Get(other, "Name"), otherTarget := RoboSync.Normal(RoboSync._Get(other, "Destination"))
                if (target != "" && otherTarget != "" && (target = otherTarget || RoboSync.Inside(target, otherTarget) || RoboSync.Inside(otherTarget, target)))
                    add("Warning", "Overlap", name)
                if (logFile != "" && logFile = RoboSync.Normal(RoboSync.Expand(RoboSync._Get(other, "LogFile"), "20000101")))
                    add("Warning", "SameLog", name)
            }
        }
        return problems
    }

    ; Before a mirror runs: the source must be there and hold something, or the mirror would delete everything in the destination
    ; (a server share that didn't connect can look empty). Touches the disk: call it when running, not while editing
    static CheckSource(task) {
        source := RoboSync.Expand(RoboSync._Get(task, "Source"))
        if (source = "" || !DirExist(source))
            return "SourceMissing"
        Loop Files, RTrim(source, "\") "\*", "FD"
            return ""
        return "SourceEmpty"
    }

    ; A path for comparing: variables expanded, lower case, no trailing backslash
    static Normal(folder) {
        folder := Trim(RoboSync.Expand(folder))
        return (folder = "") ? "" : StrLower(RegExMatch(folder, "^[A-Za-z]:\\?$") ? SubStr(folder, 1, 2) : RTrim(folder, "\"))
    }

    static Inside(inner, outer) => (outer != "" && SubStr(inner, 1, StrLen(outer) + 1) = outer "\")

    ; A drive (C:), a share (\\server\share), the OneDrive folder or the user folder (lower case, from Normal)
    static IsRootLike(folder) {
        if RegExMatch(folder, "^[a-z]:$") || RegExMatch(folder, "^\\\\[^\\]+(\\[^\\]+)?$")
            return true
        for root in [EnvGet("OneDrive"), EnvGet("OneDriveCommercial"), EnvGet("OneDriveConsumer"), EnvGet("UserProfile")]
            if (root != "" && folder = RoboSync.Normal(root))
                return true
        return false
    }

    ;---------------------------------------------------------------------------
    ; Exit code
    ;---------------------------------------------------------------------------
    ; robocopy's exit code is a bit mask: 1 files copied, 2 extra files in the destination (deleted by a mirror), 4 mismatched files,
    ; 8 some copies failed, 16 a fatal error (nothing copied). Below 8 is success, 0 = nothing to do
    ; (a negative code: robocopy didn't run or was stopped: nothing set)
    static Result(code) {
        bits := (IsInteger(code) && code >= 0) ? code : 0
        return {Ok: IsInteger(code) && code >= 0 && code < 8, Copied: (bits & 1) != 0, Extras: (bits & 2) != 0, Mismatched: (bits & 4) != 0
            , Failed: (bits & 8) != 0, Fatal: (bits & 16) != 0}
    }

    static _Number(value, default) => IsInteger(value) ? Max(0, Integer(value)) : default

    static _Get(task, key) {
        if (task is Map && task.Has(key))
            return task[key]
        defaults := RoboSync.NewTask()
        return defaults.Has(key) ? defaults[key] : ""
    }
}
