;===============================================================================
; SyncProvider.ahk - the sync tasks (RoboSync) in the search results (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Each task is a result, found by its name ("projects") or with the keyword (Features.Sync.Keyword, "sync": "sync " lists
; them all, "sync proj" only the tasks). Enter runs the task (SyncJobs.Run: checks, a mirror is previewed and asks before
; deleting), or stops it when it runs. The actions panel: Preview, Run / Stop, Open log, open the source or destination.
; The subtitle says what it does and how it is doing ("Running since 10:05", "Last run 09:30: OK").
; Without sync tasks nothing shows.
;===============================================================================

class SyncProvider {
    static Id := "Sync"
    static Icon := "res:imageres.dll,-5375"                                 ; sync arrows

    static Init() {
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
            item := SyncProvider.ItemFor(task, score + (only ? 40 : 5))
            item.Exclusive := only
            results.Push(item)
        }
        return results
    }

    static ItemFor(task, score) {
        name := task["Name"], running := SyncJobs.IsRunning(name)
        state := SyncJobs.State(name)
        subtitle := I18n.T(task["Mode"] = "Mirror" ? "Sync.Mirror" : "Sync.Copy", RoboSync.Expand(task["Source"]), RoboSync.Expand(task["Destination"]))
        if (state != "")
            subtitle := state " · " subtitle
        item := ResultItem(I18n.T(running ? "Sync.StopTitle" : "Sync.RunTitle", name), subtitle, {Icon: SyncProvider.Icon, Score: score
            , Uid: "sync:" SyncJobs.Key(name), Source: task, OnRun: running ? (*) => SyncJobs.Stop(name) : (*) => SyncJobs.Run(task)})
        add(titleKey, icon, fn) => item.Actions.Push(ResultItem(I18n.T(titleKey), "", {Icon: icon, OnRun: fn}))
        if running
            add("Sync.Action.Stop", "res:imageres.dll,-98", (*) => SyncJobs.Stop(name))
        else {
            add("Sync.Action.Run", SyncProvider.Icon, (*) => SyncJobs.Run(task))
            add("Sync.Action.Preview", "res:imageres.dll,-177", (*) => SyncJobs.Preview(task))
        }
        add("Sync.Action.OpenLog", "res:imageres.dll,-102", (*) => SyncProvider.OpenLog(task))
        add("Sync.Action.OpenSource", "folder:", (*) => ActionCatalog.OpenFolder(RoboSync.Expand(task["Source"])))
        add("Sync.Action.OpenDestination", "folder:", (*) => ActionCatalog.OpenFolder(RoboSync.Expand(task["Destination"])))
        return item
    }

    ; Today's log of the task (its own log file, or the one in Data\SyncLogs); its folder when there is none yet
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
