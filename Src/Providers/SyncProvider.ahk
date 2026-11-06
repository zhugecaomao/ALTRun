;===============================================================================
; SyncProvider.ahk - the sync tasks (RoboSync) in the search results (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Each task is a result, found by its name ("projects") or with the keyword (Features.Sync.Keyword, "sync": "sync " lists
; them all, "sync proj" only the tasks). Enter opens the task's sync window, or starts a preview when opened through the
; keyword. The actions panel: Sync Now (SyncJobs.Run: checks, a mirror is previewed and asks before deleting) and Preview /
; Show progress and Stop, Open log, open the source or destination.
; The subtitle says what it does and how it is doing ("Running since 10:05", "Last run 09:30: OK").
; Without sync tasks nothing shows.
;===============================================================================

class SyncProvider {
    static Id := "Sync"
    static Icon => IconCache.Own("Sync", A_ScriptDir "\Resources\ALTRun.ico")   ; Resources\Icons\Sync.ico (imageres.dll's icon numbers differ between Windows builds)

    static Init() {
        SyncWindowState.Start()
    }

    static Search(query) {
        tasks := RoboSync.Items, results := []
        if !tasks.Length
            return results
        keyword := AppSettings.Feature("Sync")["Keyword"]
        only := (keyword != "" && query.HasRest && query.MatchKeyword([keyword], &term))
        needle := only ? term : query.Text
        if (keyword != "" && !query.HasRest && query.MatchKeyword([keyword], &term))      ; "sync": a hint; Enter completes "sync "
            results.Push(ResultItem.KeywordHint(I18n.T("Sync.Keyword"), keyword, SyncProvider.Icon))
        for task in tasks {
            name := RoboSync._Get(task, "Name")
            if (name = "")
                continue
            score := (only && Trim(needle) = "") ? 50 : FuzzyMatcher.Score(needle, name)
            if (score <= 0)
                continue
            item := SyncProvider.ItemFor(task, score + (only ? 40 : 5), only)
            item.Exclusive := only
            results.Push(item)
        }
        return results
    }

    static ItemFor(task, score, autoPreview := false) {
        name := task["Name"], running := SyncJobs.IsRunning(name)
        state := SyncJobs.State(name)
        subtitle := I18n.T(task["Mode"] = "Mirror" ? "Sync.Mirror" : task["Mode"] = "Update" ? "Sync.Update" : "Sync.Copy", RoboSync.Expand(task["Source"]), RoboSync.Expand(task["Destination"]))
        if (state != "")
            subtitle := state " · " subtitle
        item := ResultItem(I18n.T("Sync.RunTitle", name), subtitle, {Icon: SyncProvider.Icon, Score: score
            , Uid: "sync:" SyncJobs.Key(name), Source: task, OnRun: (*) => SyncPreview.OpenTask(task, autoPreview)})
        add(titleKey, icon, fn) => item.Actions.Push(ResultItem(I18n.T(titleKey), "", {Icon: icon, OnRun: fn}))
        if running {
            add("Sync.Action.Show", SyncProvider.Icon, (*) => SyncPreview.ShowJob(name))
            add("Sync.Action.Stop", "res:imageres.dll,-98", (*) => SyncJobs.Stop(name))
        }
        else {
            add("Sync.Action.Run", SyncProvider.Icon, (*) => SyncJobs.Run(task))
            add("Sync.Action.Preview", "res:imageres.dll,-177", (*) => SyncJobs.Preview(task))
        }
        add("Sync.Action.OpenLog", "res:imageres.dll,-102", (*) => SyncProvider.OpenLog(task))
        add("Sync.Action.OpenSource", "folder:", (*) => ActionCatalog.OpenFolder(RoboSync.Expand(task["Source"])))
        add("Sync.Action.OpenDestination", "folder:", (*) => ActionCatalog.OpenFolder(RoboSync.Expand(task["Destination"])))
        return item
    }

    ; Today's log of the task (its own log file, or the one in %TEMP%\ALTRun\SyncLogs); its folder when there is none yet
    static OpenLog(task) {
        logFile := (RoboSync._Get(task, "LogFile") != "") ? RoboSync.Expand(task["LogFile"])
            : SyncJobs.LogDir "\" SyncJobs.FileName(task["Name"]) "_" FormatTime(, "yyyyMMdd") ".log"
        if FileExist(logFile)
            return Path.OpenText(logFile)
        SplitPath(logFile, , &dir)
        if DirExist(dir)
            return ActionCatalog.OpenFolder(dir)
        App.Notify(I18n.T("Sync.NoLog"))
    }
}

; Persists the Folder Sync windows' normal geometry and maximized state. Hooking window messages keeps the state handling
; independent of the preview phases: previews, runs and task windows all share the same GUI.
class SyncWindowState {
    static File => AppSettings.DataDir "\SyncPreviewWindows.json"
    static Started := false, Restored := Map(), Restoring := Map(), Timers := Map()

    static Start() {
        if SyncWindowState.Started
            return
        SyncWindowState.Started := true
        OnMessage(0x18, (wParam, lParam, msg, hwnd) => SyncWindowState._Show(wParam, hwnd))   ; WM_SHOWWINDOW
        OnMessage(0x5, (wParam, lParam, msg, hwnd) => SyncWindowState._Size(wParam, hwnd))     ; WM_SIZE
        OnMessage(0x232, (wParam, lParam, msg, hwnd) => SyncWindowState._Moved(hwnd))          ; WM_EXITSIZEMOVE
        OnMessage(0x82, (wParam, lParam, msg, hwnd) => SyncWindowState._Destroyed(hwnd))       ; WM_NCDESTROY
    }

    static _Window(hwnd) {
        for key, w in SyncPreview.Windows
            if (w.Gui.Hwnd = hwnd)
                return w
        return ""
    }

    static _Show(show, hwnd) {
        if !show || !IsObject(w := SyncWindowState._Window(hwnd)) || SyncWindowState.Restored.Has(hwnd)
            return
        SyncWindowState.Restored[hwnd] := true
        SyncWindowState.Restoring[hwnd] := true
        state := SyncWindowState._Load(w.Key)
        SetTimer(() => SyncWindowState._Restore(w, state), -20)
    }

    static _Restore(w, state) {
        hwnd := w.Gui.Hwnd
        try {
            if (state.Has("X") && state.Has("Y") && state.Has("W") && state.Has("H")
                    && IsNumber(state["X"]) && IsNumber(state["Y"]) && IsNumber(state["W"]) && IsNumber(state["H"])) {
                x := state["X"], y := state["Y"], width := Max(640, state["W"]), height := Max(380, state["H"])
                area := Win.WorkAreaAt(x + width // 2, y + height // 2)
                width := Min(width, area.Right - area.Left), height := Min(height, area.Bottom - area.Top)
                x := Max(area.Left, Min(x, area.Right - width)), y := Max(area.Top, Min(y, area.Bottom - height))
                WinRestore("ahk_id " hwnd)
                WinMove(x, y, width, height, "ahk_id " hwnd)
                if (state.Has("Maximized") && state["Maximized"])
                    WinMaximize("ahk_id " hwnd)
            }
        } catch as e {
            Logger.Debug("SyncWindowState: restore failed - " e.Message)
        }
        SetTimer(() => SyncWindowState.Restoring.Delete(hwnd), -300)
    }

    static _Size(minMax, hwnd) {
        if !IsObject(w := SyncWindowState._Window(hwnd)) || SyncWindowState.Restoring.Has(hwnd)
            return
        if (minMax = 1)                                                     ; do not replace the saved normal geometry when minimized
            return
        state := SyncWindowState._Load(w.Key)
        if (minMax = 2) {
            state["Maximized"] := 1
        } else {
            state["Maximized"] := 0
            try {
                WinGetPos(&x, &y, &width, &height, "ahk_id " hwnd)
                if (width > 0 && height > 0)
                    state["X"] := x, state["Y"] := y, state["W"] := width, state["H"] := height
            }
        }
        SyncWindowState._ScheduleSave(w.Key, state, hwnd)
    }

    static _Moved(hwnd) {
        if !IsObject(w := SyncWindowState._Window(hwnd)) || SyncWindowState.Restoring.Has(hwnd)
            return
        if (WinGetMinMax("ahk_id " hwnd) != 0)
            return
        state := SyncWindowState._Load(w.Key)
        try {
            WinGetPos(&x, &y, &width, &height, "ahk_id " hwnd)
            if (width > 0 && height > 0)
                state["X"] := x, state["Y"] := y, state["W"] := width, state["H"] := height
        }
        SyncWindowState._ScheduleSave(w.Key, state, hwnd)
    }

    static _ScheduleSave(key, state, hwnd) {
        SyncWindowState.Timers[hwnd] := {Key: key, State: state}
        if !SyncWindowState.Timers.Has("Timer" hwnd)
            SyncWindowState.Timers["Timer" hwnd] := () => SyncWindowState._SavePending(hwnd)
        SetTimer(SyncWindowState.Timers["Timer" hwnd], -250)
    }

    static _SavePending(hwnd) {
        if !SyncWindowState.Timers.Has(hwnd)
            return
        pending := SyncWindowState.Timers[hwnd]
        SyncWindowState._Write(pending.Key, pending.State)
        SyncWindowState.Timers.Delete(hwnd)
    }

    static _Destroyed(hwnd) {
        if SyncWindowState.Timers.Has(hwnd) {
            pending := SyncWindowState.Timers[hwnd]
            SyncWindowState._Write(pending.Key, pending.State)
            SyncWindowState.Timers.Delete(hwnd)
        }
        for map in [SyncWindowState.Restored, SyncWindowState.Restoring]
            if map.Has(hwnd)
                map.Delete(hwnd)
        if SyncWindowState.Timers.Has("Timer" hwnd) {
            try SetTimer(SyncWindowState.Timers["Timer" hwnd], 0)
            SyncWindowState.Timers.Delete("Timer" hwnd)
        }
    }

    static _Load(key) {
        try {
            if FileExist(SyncWindowState.File) {
                data := JSON.Parse(FileRead(SyncWindowState.File, "UTF-8"))
                if (data is Map && data.Has(key) && data[key] is Map)
                    return data[key]
            }
        } catch as e {
            Logger.Debug("SyncWindowState: cannot read state - " e.Message)
        }
        return Map("Maximized", 0, "X", 120, "Y", 120, "W", 1000, "H", 600)
    }

    static _Write(key, state) {
        try {
            data := Map()
            if FileExist(SyncWindowState.File) {
                try data := JSON.Parse(FileRead(SyncWindowState.File, "UTF-8"))
                if !(data is Map)
                    data := Map()
            }
            data[key] := state
            DirCreate(AppSettings.DataDir)
            JSON.WriteFile(SyncWindowState.File, data)
        } catch as e {
            Logger.Error("SyncWindowState: cannot save state - " e.Message)
        }
    }
}
