;===============================================================================
; SyncJobs.ahk - runs the sync tasks (RoboSync) with robocopy in the background (AutoHotkey v2)
;-------------------------------------------------------------------------------
; robocopy runs hidden, one process per task; ALTRun keeps its process handle and looks once a second (WinTimer, only
; while something runs) whether it ended, then reads its exit code (RoboSync.Result: < 8 = success).
;
; Run (search: "sync <name>", Enter):
;   - the settings are checked first (RoboSync.Check); a mirror also needs a source that exists and isn't empty (CheckSource)
;   - the sync window (SyncPreview) opens at once and follows it: what robocopy finds and copies, read from its log as it is written
;     (SyncProgress), a progress bar, the result
;   - a mirror is previewed first (/L, which changes nothing): when it would delete anything in the destination, the window lists
;     it and asks before running. Previews that delete nothing go straight on, in the same window
;   - a task without a log file of its own logs to %TEMP%\ALTRun\SyncLogs\<name>_<date>.log, so there is always a log to open;
;     logs there older than KeepLogDays (Features.Sync, 30) are deleted when a task runs. Not in Data: it may be in OneDrive,
;     which would upload logs that grow all the time. Each run also leaves a line in ALTRun.log (started, how it ended)
;   - a task that ends tells how it went in its window, and with a notification when the window isn't in front (or closed);
;     a monitoring task (/MON /MOT) keeps running until stopped
; Monitoring tasks stop when ALTRun quits. When ALTRun only restarts (after Preferences, an update), they keep running and the
; new ALTRun takes them over (Data\SyncRunning.json), so changing a setting doesn't interrupt a sync.
;
; Usage:
;   SyncJobs.Run(task)            the user's Run (checks, preview of a mirror, asks before deleting)
;   SyncJobs.RunNow(task, totals) runs it without a preview, the window following it
;   SyncJobs.Preview(task)        /L, then what it would do
;   SyncJobs.Start(task, preview, onDone)   -> "" or a problem code: starts robocopy, nothing asked
;   SyncJobs.Stop(name) / StopAll()        SyncJobs.IsRunning(name)        SyncJobs.Last[key] = {Code, Ended, Preview}
;   SyncJobs.Summary(logText)     -> {Dirs, Files, Bytes} ({Total, Copied, Skipped, Mismatch, Failed, Extras}) of the last run in a log, "" if none
;   SyncJobs.Shutdown(exitBy) / Adopt()     ALTRun quits or restarts / at start-up
;===============================================================================

class SyncJobs {
    static Jobs := Map(), Last := Map()                                     ; Key(name) -> the running job / how the last run ended
    static PollMs := 1000, StopWaitMs := 2000
    static StateFile => AppSettings.DataDir "\SyncRunning.json"
    static LogDir := A_Temp "\ALTRun\SyncLogs"
    static _pollTimer := () => SyncJobs._Poll()
    static Changed := ""                                                    ; a function called when a job starts or ends (the Preferences page shows the states)

    static Key(name) => StrLower(Trim(name))
    static FileName(name) => RegExReplace(Trim(name), '[\\/:*?"<>|]', "_")
    static IsRunning(name) => SyncJobs.Jobs.Has(SyncJobs.Key(name))

    ;---------------------------------------------------------------------------
    ; What the user runs
    ;---------------------------------------------------------------------------
    ; The sync window (SyncPreview) opens at once and follows the task: the preview as robocopy finds the changes, the run as it copies
    static Run(task) {
        name := task["Name"]
        if (problem := SyncJobs._Problem(task, false))
            return SyncJobs._Tell(problem, name)
        if (task["Mode"] != "Mirror")
            return SyncJobs.RunNow(task)
        problem := SyncJobs.Start(task, true, (job, code) => SyncJobs._AfterCheck(task, job, code))
        if (problem != "")
            return SyncJobs._Tell(problem, name)
        SyncPreview.Open(task, "Check", {Run: true, Confirm: true})
    }

    ; Runs it without a preview (checked already), the sync window following it. totals: {Files, Bytes} a preview found, for the progress bar
    static RunNow(task, totals := "") {
        if (problem := SyncJobs.Start(task))
            return SyncJobs._Tell(problem, task["Name"])
        SyncPreview.Open(task, "Run", {Totals: totals})
    }

    ; The preview before a mirror ended: run it at once when it deletes nothing; otherwise the preview window lists what it would
    ; delete and runs it only when the user chooses to
    static _AfterCheck(task, job, code) {
        if !RoboSync.Result(code).Ok
            return SyncPreview.Finish(job, code) || SyncJobs._Failed(task["Name"], code, job.LogFile)
        result := SyncJobs.ReadPreview(job.LogFile)
        if (IsObject(result.Summary) && result.Summary.Files.Extras + result.Summary.Dirs.Extras > 0)
            return SyncPreview.Show(task, result, {Run: true, Confirm: true})
        SyncJobs.RunNow(task, SyncJobs.Totals(result))
    }

    ; What a preview would copy, for the progress bar of the run that follows: {Files, Bytes}
    static Totals(result) {
        summary := result.Summary
        if IsObject(summary)
            return {Files: summary.Files.Copied, Bytes: IsObject(summary.Bytes) ? summary.Bytes.Copied : 0}
        return ""
    }

    ; options.Run: the preview window offers to run the task (not for a task still being edited)
    static Preview(task, options := "") {
        name := task["Name"]
        if (problem := SyncJobs._Problem(task, true))
            return SyncJobs._Tell(problem, name)
        options := IsObject(options) ? options : {Run: true}
        problem := SyncJobs.Start(task, true, (job, code) => SyncJobs._ShowPreview(task, job, code, options))
        if (problem != "")
            return SyncJobs._Tell(problem, name)
        SyncPreview.Open(task, "Check", {Run: options.HasOwnProp("Run") && options.Run, Confirm: false})
    }

    static _ShowPreview(task, job, code, options) {
        if !RoboSync.Result(code).Ok
            return SyncPreview.Finish(job, code) || SyncJobs._Failed(task["Name"], code, job.LogFile)
        result := SyncJobs.ReadPreview(job.LogFile)
        if !IsObject(result.Summary)
            return Path.OpenText(job.LogFile)
        SyncPreview.Show(task, result, {Run: options.HasOwnProp("Run") && options.Run, Confirm: false})
    }

    ; A robocopy log's text. Logs are written with /UNILOG, UTF-16 with a byte order mark (Start writes the mark for /UNILOG+):
    ; /LOG doesn't write Unicode, and names outside its code page come out wrong (Chinese names become ??). FileRead goes by
    ; the mark; an older /LOG log has none and is read in the system code page, as before
    static ReadLog(logFile) => FileRead(logFile)

    ; A preview's log -> {Summary, Changes, LogFile} (Summary "" when the log has none)
    static ReadPreview(logFile) {
        text := ""
        try text := SyncJobs.ReadLog(logFile)
        return {Summary: SyncJobs.Summary(text), Changes: SyncJobs.Changes(text), LogFile: logFile}
    }

    ; The files and folders a run (or a preview, /L) lists in its log: [{Kind: "Copy" | "Delete", Folder, Path, Label, Size}...], at
    ; most limit (Total: how many there are; NewDirs: the folders it creates, 0 with /MT, which lists no folders). How the lines are
    ; read: SyncProgress
    static Changes(logText, limit := 2000) {
        progress := SyncProgress("", 0, limit)
        progress.Feed(logText, true)
        list := progress.Changes
        list.Total := progress.Total, list.NewDirs := progress.NewDirs
        return list
    }

    ; A reason not to run (a problem code), "" when it may run. A preview changes nothing: only the settings' errors stop it
    static _Problem(task, preview) {
        if SyncJobs.IsRunning(task["Name"])
            return "Running"
        for problem in RoboSync.Check(task)
            if (problem.Level = "Error")
                return problem.Code
        if (!preview && task["Mode"] = "Mirror")
            return RoboSync.CheckSource(task)
        return ""
    }

    static _Tell(problem, name) => MsgBox(I18n.T("Sync.Problem." problem, name), App.Name, "Icon!")

    static _Failed(name, code, logFile) {
        if (MsgBox(I18n.T("Sync.Failed", name, code) "`n`n" I18n.T("Sync.OpenLog"), App.Name, "YesNo Icon!") = "Yes")
            Path.OpenText(logFile)
    }

    ; The task as robocopy runs it: without a log file of its own, the default one (appended to, a file a day); a preview logs to a
    ; file of its own (overwritten) that the summary and the list are read from. The edit window shows its command
    static AsRun(task, preview := false) {
        copy := task.Clone()
        if preview {
            copy["LogFile"] := SyncJobs.LogDir "\" SyncJobs.FileName(task["Name"]) "_preview.log", copy["AppendLog"] := 0
            copy["Extra"] := Trim(RegExReplace(RoboSync._Get(task, "Extra"), "i)(^|\s)/(NJS|NFL|NDL|V|X|TEE)(?=\s|$)"))   ; the summary and the list are read from the log
        } else if (RoboSync._Get(task, "LogFile") = "") {
            copy["LogFile"] := SyncJobs.LogDir "\" SyncJobs.FileName(task["Name"]) "_{date}.log", copy["AppendLog"] := 1
        }
        return copy
    }

    ;---------------------------------------------------------------------------
    ; Processes
    ;---------------------------------------------------------------------------
    ; Starts robocopy for the task, without asking anything. onDone(job, exitCode) when it ends. Returns "" or a problem code
    ; (Running, StartFailed). It runs the task as AsRun gives it (the default log, the preview's log)
    static Start(task, preview := false, onDone := "") {
        key := SyncJobs.Key(task["Name"])
        if SyncJobs.Jobs.Has(key)
            return "Running"
        copy := SyncJobs.AsRun(task, preview)
        if !preview
            SyncJobs.CleanLogs()
        command := RoboSync.Build(copy, preview)
        logFile := RoboSync.Expand(copy["LogFile"])
        SplitPath(logFile, , &logDir)
        try DirCreate(logDir)                                               ; robocopy doesn't create the log's folder
        ; /UNILOG+ starts a new file in the code page, not Unicode (Chinese names become ??): a file that already starts with the UTF-16
        ; mark is appended to in UTF-16. A file an older version wrote (no mark) is left as it is
        if (RoboSync._Get(copy, "AppendLog") && RoboSync.IsAbsolute(logFile) && !FileExist(logFile))
            try FileOpen(logFile, "w", "UTF-16").Close()
        logStart := 0                                                       ; where this run's part of the log starts (a run appends to the day's log)
        if (RoboSync._Get(copy, "AppendLog") && FileExist(logFile))
            try logStart := FileGetSize(logFile)
        pid := 0
        try Run('"' command.Program '" ' command.Args, A_Temp, "Hide", &pid)   ; the folders are full paths (RoboSync.Check); a relative log file goes to TEMP
        catch as e {
            Logger.Error("SyncJobs: cannot start robocopy for " task["Name"] " - " e.Message)
            return "StartFailed"
        }
        handle := DllCall("OpenProcess", "UInt", 0x101000, "Int", 0, "UInt", pid, "Ptr")   ; SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION
        SyncJobs.Jobs[key] := {Name: task["Name"], Pid: pid, Handle: handle, Preview: preview, Monitor: !preview && RoboSync._Get(task, "Monitor")
            , Started: A_Now, StartTick: A_TickCount, LogFile: logFile, Progress: SyncProgress(logFile, logStart), OnDone: onDone}
        Logger.Debug("SyncJobs: started " task["Name"] (preview ? " (preview)" : "") ", pid " pid)
        WinTimer.Set(SyncJobs._pollTimer, SyncJobs.PollMs)
        SyncJobs._Notify()
        return ""
    }

    static _Notify() {
        if IsObject(SyncJobs.Changed)
            SyncJobs.Changed.Call()
    }

    ; Deletes the logs in LogDir older than KeepLogDays (0 = keeps them all). Only ALTRun's own folder: a task's own log file is the user's
    static CleanLogs(days := "") {
        if (days = "") {
            sync := AppSettings.Feature("Sync")
            days := sync.Has("KeepLogDays") ? sync["KeepLogDays"] : 30
        }
        if (!IsInteger(days) || days <= 0 || !DirExist(SyncJobs.LogDir))
            return 0
        cutoff := DateAdd(A_Now, -days, "Days"), deleted := 0
        Loop Files SyncJobs.LogDir "\*.log" {
            if (A_LoopFileTimeModified < cutoff) {
                try FileDelete(A_LoopFileFullPath), deleted++
            }
        }
        if deleted
            Logger.Debug("SyncJobs: deleted " deleted " logs older than " days " days")
        return deleted
    }

    static Stop(name) {
        key := SyncJobs.Key(name)
        if !SyncJobs.Jobs.Has(key)
            return false
        job := SyncJobs.Jobs[key]
        job.OnDone := "", job.Stopped := true                               ; stopped on purpose: no "failed" message
        try ProcessClose(job.Pid)                                           ; TerminateProcess only starts ending it: wait (2 s at most) so it is gone when this returns
        if job.Handle
            DllCall("WaitForSingleObject", "Ptr", job.Handle, "UInt", SyncJobs.StopWaitMs, "UInt")
        else if job.Pid
            ProcessWaitClose(job.Pid, SyncJobs.StopWaitMs / 1000)
        SyncJobs._Poll()                                                    ; still running after that: the poll timer ends the job later
        return true
    }

    static StopAll() {
        for key, job in SyncJobs.Jobs.Clone()
            SyncJobs.Stop(job.Name)
    }

    ; Once a second while something runs: the jobs that ended (their exit code), then onDone or the notification
    static _Poll() {
        for key, job in SyncJobs.Jobs.Clone() {
            code := 259                                                     ; STILL_ACTIVE
            if job.Handle
                DllCall("GetExitCodeProcess", "Ptr", job.Handle, "UInt*", &code)
            else if !ProcessExist(job.Pid)                                  ; taken over without a handle: no exit code
                code := -1
            if (code = 259 && (!job.Handle || DllCall("WaitForSingleObject", "Ptr", job.Handle, "UInt", 0, "UInt") != 0))
                continue
            if job.Handle
                DllCall("CloseHandle", "Ptr", job.Handle)
            SyncJobs.Jobs.Delete(key)
            SyncJobs.Last[key] := {Code: code, Ended: A_Now, Preview: job.Preview}
            Logger.Debug("SyncJobs: " job.Name (job.Preview ? " (preview)" : "") " ended, exit code " code)
            SyncJobs._Notify()
            if job.HasOwnProp("Stopped")
                SyncPreview.Finish(job, "Stopped")
            else if IsObject(job.OnDone)
                job.OnDone.Call(job, code)
            else if (!job.Preview && code != -1)
                SyncJobs._Report(job, code)
        }
        if !SyncJobs.Jobs.Count
            WinTimer.Set(SyncJobs._pollTimer, 0)
    }

    ; A run that ended on its own: how it went (copied files from the log's last summary)
    ; The sync window shows it when it is open; a notification as well when the window isn't in front, or instead of it when closed
    static _Report(job, code) {
        shown := SyncPreview.Finish(job, code)
        active := shown && SyncPreview.IsActive(job.Name)
        if !RoboSync.Result(code).Ok {
            Logger.Warn("SyncJobs: " job.Name " failed, exit code " code ", log " job.LogFile)
            if !shown
                return SyncJobs._Failed(job.Name, code, job.LogFile)
            return active ? "" : App.Toast(I18n.T("Sync.Failed", job.Name, code))
        }
        summary := ""
        try summary := SyncJobs.Summary(SyncJobs.ReadLog(job.LogFile))
        if IsObject(summary)
            Logger.Debug("SyncJobs: " job.Name " done, copied " summary.Files.Copied " files, deleted " summary.Files.Extras " files and " summary.Dirs.Extras " folders")
        if !active
            App.Toast(IsObject(summary) ? I18n.T("Sync.Done", job.Name, summary.Files.Copied, summary.Files.Extras) : I18n.T("Sync.DoneShort", job.Name))
    }

    ;---------------------------------------------------------------------------
    ; Quitting and restarting
    ;---------------------------------------------------------------------------
    ; Quitting stops the tasks; a restart leaves them running and writes them down for the new ALTRun
    static Shutdown(exitBy) {
        if !SyncJobs.Jobs.Count
            return
        if (SubStr(exitBy, 1, 7) != "restart")
            return SyncJobs.StopAll()
        list := []
        for key, job in SyncJobs.Jobs
            if !job.Preview
                list.Push(Map("Name", job.Name, "Pid", job.Pid, "Started", job.Started, "LogFile", job.LogFile, "Monitor", job.Monitor ? 1 : 0))
        try JSON.WriteFile(SyncJobs.StateFile, list)
    }

    ; At start-up: the tasks a restart left running (still robocopy with the same process id)
    static Adopt() {
        if !FileExist(SyncJobs.StateFile)
            return
        list := []
        try list := JSON.Parse(FileRead(SyncJobs.StateFile, "UTF-8"))
        try FileDelete(SyncJobs.StateFile)
        for entry in (list is Array ? list : []) {
            pid := entry["Pid"]
            if !(ProcessExist(pid) && StrLower(ProcessGetName(pid)) = "robocopy.exe")
                continue
            handle := DllCall("OpenProcess", "UInt", 0x101000, "Int", 0, "UInt", pid, "Ptr")
            SyncJobs.Jobs[SyncJobs.Key(entry["Name"])] := {Name: entry["Name"], Pid: pid, Handle: handle, Preview: false, Monitor: entry["Monitor"] ? true : false
                , Started: entry["Started"], StartTick: A_TickCount, LogFile: entry["LogFile"], OnDone: ""
                , Progress: SyncProgress(entry["LogFile"], FileExist(entry["LogFile"]) ? FileGetSize(entry["LogFile"]) : 0)}
        }
        if SyncJobs.Jobs.Count
            WinTimer.Set(SyncJobs._pollTimer, SyncJobs.PollMs)
    }

    ;---------------------------------------------------------------------------
    ; The summary robocopy writes at the end of a run
    ;---------------------------------------------------------------------------
    ; The table at the end of a run: rows Dirs, Files, Bytes, Times; columns Total, Copied, Skipped, Mismatch, FAILED, Extras. The
    ; labels are in Windows' language, so rows are found by their six numbers: the last two such rows in a row are Dirs and Files
    ; (Bytes may have units such as "1.5 m"; in a log several runs append, the last table counts)
    static Summary(logText) {
        rows := [], lines := StrSplit(StrReplace(logText, "`r"), "`n")
        for index, line in lines
            if RegExMatch(line, "^\s*[^\d\s][^:：]*[:：]\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*$", &m)
                rows.Push({Index: index, Values: [m[1], m[2], m[3], m[4], m[5], m[6]]})
        found := 0
        Loop rows.Length - 1 {
            i := rows.Length - A_Index                                      ; from the end
            if (rows[i + 1].Index = rows[i].Index + 1 && (i = 1 || rows[i - 1].Index != rows[i].Index - 1)) {
                found := i
                break
            }
        }
        if !found
            return ""
        row(values) => {Total: Integer(values[1]), Copied: Integer(values[2]), Skipped: Integer(values[3]), Mismatch: Integer(values[4])
            , Failed: Integer(values[5]), Extras: Integer(values[6])}
        bytes := ""                                                         ; the row after Files: sizes, bytes or with a unit ("400.00 m")
        next := rows[found + 1].Index + 1
        if (next <= lines.Length && RegExMatch(lines[next], "^\s*[^\d\s][^:：]*[:：](.*)$", &m)) {
            values := [], at := 1
            while (at := RegExMatch(m[1], "i)(\d+(?:\.\d+)?)(?:\s([kmgt]))?(?=\s|$)", &v, at)) {
                values.Push(SyncProgress.ParseSize(v[0])), at += StrLen(v[0])
            }
            if (values.Length = 6)
                bytes := row(values)
        }
        return {Dirs: row(rows[found].Values), Files: row(rows[found + 1].Values), Bytes: bytes}
    }

    ; For the search results: "Running since 10:05", "Monitoring since 10:05", "Last run 09:30: OK", ...
    ; The state in a word or two, for the Preferences list: "Running", "Monitoring", "OK 09:30", "Failed 09:30"; "" when not run yet
    static ShortState(name) {
        key := SyncJobs.Key(name)
        if SyncJobs.Jobs.Has(key) {
            job := SyncJobs.Jobs[key]
            return I18n.T(job.Preview ? "Sync.Short.Checking" : job.Monitor ? "Sync.Short.Monitoring" : "Sync.Short.Running")
        }
        if SyncJobs.Last.Has(key) {
            last := SyncJobs.Last[key]
            return I18n.T(RoboSync.Result(last.Code).Ok ? "Sync.Short.Ok" : "Sync.Short.Failed", FormatTime(last.Ended, "HH:mm"))
        }
        return ""
    }

    static State(name) {
        key := SyncJobs.Key(name)
        if SyncJobs.Jobs.Has(key) {
            job := SyncJobs.Jobs[key]
            return I18n.T(job.Preview ? "Sync.State.Checking" : job.Monitor ? "Sync.State.Monitoring" : "Sync.State.Running", FormatTime(job.Started, "HH:mm"))
        }
        if SyncJobs.Last.Has(key) {
            last := SyncJobs.Last[key]
            return I18n.T(RoboSync.Result(last.Code).Ok ? "Sync.State.LastOk" : "Sync.State.LastFailed", FormatTime(last.Ended, "HH:mm"), last.Code)
        }
        return ""
    }
}
