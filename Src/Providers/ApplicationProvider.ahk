;===============================================================================
; ApplicationProvider.ahk - application search (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Indexes the shortcuts and programs in the Start menu and on the desktop (Features.Applications.Folders), and
; Store apps (PowerShell Get-StartApps, run in the background without blocking the UI). Chinese names also
; match by pinyin initials ("wx" -> 微信).
;
; The index is cached in Data\AppIndex.json and read directly at start-up. It is rescanned in the background in batches (at most about 15 ms per timer tick, so hotkeys aren't blocked),
; without notifications; when the search window is open (or opening) during a scan it pauses (if an old index is available) and goes on after it closes, so it never competes with showing or searching:
;   - watching Folders (with subfolders, Windows folder change notifications): files added, deleted or renamed rebuild after WatchQuietMs of quiet
;     (installers create several shortcuts in a row), at least WatchMinGapMs apart. Newly installed software can be found a few seconds later
;   - as a fallback (what watching misses: software installed while ALTRun wasn't running, folders that didn't exist or couldn't be watched at start-up): it also rebuilds when the cache is older than
;     FallbackMinutes, checked at start-up and every minute after (formerly the RefreshMinutes setting, no longer read)
;   - the cache write after a background rebuild waits until keyboard and mouse have been idle a while (writing the JSON is done in one go), and is done at once on exit if still pending
; Store apps (PowerShell) are scanned in the background at most once an hour.
; The tray menu's "Rebuild index" or the system command "Rebuild ALTRun Index" rebuilds immediately (in one go, showing the result).
;
; Unwanted applications: Ctrl+Del (or right-click "Delete") in the results adds them to the Hidden list so they no longer show,
; not even after rebuilding (nothing is uninstalled and AppIndex.json isn't changed). Deleting the line from Hidden in
; Preferences -> Applications brings it back.
;
; Settings (ALTRun.json -> Features.Applications):
;   Folders / FileTypes / Depth / Exclude / Hidden / StoreApps / MatchPinyin
;===============================================================================

class ApplicationProvider {
    static Id        := "Applications"
    static IndexFile := AppSettings.DataDir "\AppIndex.json"
    static Apps      := []            ; [Map("Title", "Target", "Detail", "Search")...]
    static _storePid := 0, _storeFile := "", _storeTimer := "", _storeScanned := 0
    static _job := "", _scanTimer := "", _refreshTimer := ""         ; background rebuild: see _NewJob
    static WatchQuietMs := 5000, WatchMinGapMs := 60000, SaveIdleMs := 2000, FallbackMinutes := 60
    static _watcher := "", _watchTimer := "", _watchTypes := ""
    static _changedAt := 0, _lastBackground := 0, _savePending := false, _saveTimer := ""
    static _entries := [], _keys := [], _keysFor := "", _lastNeedle := "", _lastMatches := []

    static Init() {
        if !ApplicationProvider._LoadCache() || ApplicationProvider._CacheExpired()
            SetTimer(() => ApplicationProvider.RefreshInBackground(), -3000)   ; show the window first, then scan in the background
        SetTimer(() => ApplicationProvider._SearchKeys(), -500)             ; compute the search keys in advance, so the first search doesn't wait
        ApplicationProvider._refreshTimer := Logger.Watch("app refresh check", () => ApplicationProvider._CheckRefresh())
        SetTimer(ApplicationProvider._refreshTimer, 60000)
        ApplicationProvider.StartWatching()
    }

    ;---------------------------------------------------------------------------
    ; Watch the indexed folders (FolderWatcher, a quick look every 2 s without waiting): rebuild in the background when application-type files (FileTypes) or folders
    ; are added, deleted, renamed or modified. The file names are known, so other files changing on the desktop don't count; neither do ALTRun's own folders
    ; (on the desktop, cache and log writes would trigger it all the time)
    ;---------------------------------------------------------------------------
    static StartWatching() {
        ApplicationProvider.StopWatching()
        options := AppSettings.Feature("Applications")
        folders := []
        for folder in options["Folders"]
            folders.Push(Path.Resolve(folder))
        ApplicationProvider._watcher := FolderWatcher(folders)
        ApplicationProvider._watchTypes := ApplicationProvider._TypesPattern(options)
        if (ApplicationProvider._watchTimer = "")
            ApplicationProvider._watchTimer := Logger.Watch("app folder watch", () => ApplicationProvider._PollWatches())
        if ApplicationProvider._watcher.Count
            SetTimer(ApplicationProvider._watchTimer, 2000)
        if (ApplicationProvider._watcher.Count < folders.Length)
            Logger.Debug("ApplicationProvider: watching " ApplicationProvider._watcher.Count " of " folders.Length " folders")
    }

    static StopWatching() {
        if (ApplicationProvider._watchTimer != "")
            SetTimer(ApplicationProvider._watchTimer, 0)
        if IsObject(ApplicationProvider._watcher)
            ApplicationProvider._watcher.Stop()
        ApplicationProvider._watcher := "", ApplicationProvider._changedAt := 0
    }

    ; Every 2 s: read the changes that have arrived (no waiting) and note the time of application-related ones; rebuild after WatchQuietMs of quiet, at least WatchMinGapMs after the last background rebuild
    static _PollWatches() {
        if !IsObject(ApplicationProvider._watcher)
            return
        for fullPath in ApplicationProvider._watcher.Poll()
            if (fullPath = "*" || ApplicationProvider.IsWatchedChange(fullPath, ApplicationProvider._watchTypes)) {
                ApplicationProvider._changedAt := A_TickCount
                break
            }
        if (!ApplicationProvider._changedAt || A_TickCount - ApplicationProvider._changedAt < ApplicationProvider.WatchQuietMs)
            return
        if (ApplicationProvider._lastBackground && A_TickCount - ApplicationProvider._lastBackground < ApplicationProvider.WatchMinGapMs)
            return
        if IsObject(ApplicationProvider._job)
            return
        ApplicationProvider._changedAt := 0
        Logger.Debug("ApplicationProvider: folders changed, rebuilding")
        ApplicationProvider.RefreshInBackground()
    }

    ; Pure function: whether a change at this path needs a rebuild -- an application-type file (types: _TypesPattern), or one without an extension (probably a folder,
    ; e.g. a program's folder in the Start menu deleted / renamed); not inside ALTRun's own folders
    static IsWatchedChange(fullPath, types, ownDir := A_ScriptDir) {
        if (ownDir != "" && InStr(fullPath, RTrim(ownDir, "\") "\") = 1)
            return false
        if DirExist(fullPath)
            return true
        SplitPath(fullPath, &name, , &ext)
        return (ext = "" || RegExMatch(name, types)) ? true : false
    }

    ; FileTypes ("*.lnk"...) -> one regex ("i)^(?:.*\.lnk|...)$")
    static _TypesPattern(options) {
        types := ""
        for pattern in options["FileTypes"]
            types .= (types = "" ? "" : "|") StrReplace(StrReplace(RegExReplace(pattern, "[\\.^$+()\[\]{}|]", "\$0"), "*", ".*"), "?", ".")
        return "i)^(?:" types ")$"
    }

    static Search(query) {
        needle := StrLower(query.Text)
        if (needle = "")
            return []
        keys := ApplicationProvider._SearchKeys()
        apps := ApplicationProvider._entries

        ; When typing on (the new input starts with the previous one), only the previous matches need searching.
        ; Below 3 letters the "letters in order" rule doesn't apply, so those results can't be used to narrow down.
        candidates := ""
        if (StrLen(ApplicationProvider._lastNeedle) >= 3 && InStr(needle, ApplicationProvider._lastNeedle) = 1)
            candidates := ApplicationProvider._lastMatches
        scores := Map()
        if IsObject(candidates) {
            for index in candidates
                if (score := FuzzyMatcher.BestKey(needle, keys[index]))
                    scores[index] := score
        } else {
            for index, entryKeys in keys
                if (score := FuzzyMatcher.BestKey(needle, entryKeys))
                    scores[index] := score
        }
        matches := [], ranks := Map()
        for index, score in scores {
            matches.Push(index)
            ranks[index] := score + Knowledge.Boost(query.Text, "app:" StrLower(apps[index]["Target"]))   ; often chosen applications aren't pushed out of the top results
        }
        ApplicationProvider._lastNeedle := needle, ApplicationProvider._lastMatches := matches

        results := []
        for index in FuzzyMatcher.TopIndexes(ranks, ProviderRegistry.MaxResults) {
            entry := apps[index]
            isStore := InStr(entry["Target"], "shell:AppsFolder\") = 1
            subtitle := isStore ? I18n.T("App.Subtitle.Store") : (entry["Detail"] != "") ? entry["Detail"] : entry["Target"]
            results.Push(ResultItem(entry["Title"], subtitle, {
                Kind: "file", Arg: entry["Target"], Source: entry,
                Icon: entry["Target"], Uid: "app:" StrLower(entry["Target"]), Score: scores[index] + 10
            }))
        }
        return results
    }

    ; compute the search keys while idle after start-up, so the first input doesn't wait
    static Warm() {
        ApplicationProvider._SearchKeys()
    }

    ; For the icon prefetch after start-up (see ProviderRegistry.IconSpecs): the applications that aren't hidden
    static IconItems() {
        ApplicationProvider._SearchKeys()
        items := []
        for entry in ApplicationProvider._entries
            items.Push({Icon: entry["Target"], Uid: "app:" StrLower(entry["Target"])})
        return items
    }

    ; The applications that aren't hidden (_entries) and their search keys (name + pinyin initials);
    ; recomputed when Apps becomes a new array or the hidden list changes
    static _SearchKeys() {
        version := ObjPtr(ApplicationProvider.Apps) "|" ApplicationProvider._HiddenSignature()
        if (ApplicationProvider._keysFor = version)
            return ApplicationProvider._keys
        hidden := ApplicationProvider._HiddenTargets()
        entries := [], keys := []
        for entry in ApplicationProvider.Apps {
            if hidden.Has(entry["Target"])
                continue
            pinyinText := entry.Has("Search") ? entry["Search"] : ""
            entries.Push(entry)
            keys.Push([FuzzyMatcher.Key(entry["Title"]), (pinyinText != "") ? FuzzyMatcher.Key(pinyinText) : ""])
        }
        ApplicationProvider._entries := entries, ApplicationProvider._keys := keys
        ApplicationProvider._keysFor := version
        ApplicationProvider._lastNeedle := "", ApplicationProvider._lastMatches := []
        return keys
    }

    ;---------------------------------------------------------------------------
    ; Hides an unwanted application (Ctrl+Del / right-click "Delete" in the results)
    ;---------------------------------------------------------------------------
    static DeleteItem(item) {
        return ApplicationProvider.Hide(item.Arg)
    }

    static DeletePrompt(item) {
        return I18n.T("App.ConfirmHide", item.Title)
    }

    static Hide(target) {
        hidden := AppSettings.Feature("Applications")["Hidden"]
        if ApplicationProvider._HiddenTargets().Has(target)
            return true
        hidden.Push(target)
        return AppSettings.Save()
    }

    ; The hidden list -> Map (case-insensitive), rebuilt when the list changes (a new array or a different count)
    static _HiddenTargets() {
        static cache := "", cacheFor := ""
        signature := ApplicationProvider._HiddenSignature()
        if (cacheFor != signature) {
            cache := Map()
            cache.CaseSense := "Off"
            for target in AppSettings.Feature("Applications")["Hidden"]
                cache[target] := true
            cacheFor := signature
        }
        return cache
    }

    static _HiddenSignature() {
        hidden := AppSettings.Feature("Applications")["Hidden"]
        return ObjPtr(hidden) ":" hidden.Length
    }

    ; Rescans now (in one go) and returns the number of applications found. Stops a background rebuild first
    static Rebuild() {
        ApplicationProvider._StopJob()
        job := ApplicationProvider._NewJob()
        ApplicationProvider._Step(job, 0)
        return ApplicationProvider._Finish(job, true)
    }

    ; Rescans in the background in batches and replaces Apps when done (searches meanwhile use the old one); no second scan while one runs
    static RefreshInBackground() {
        if IsObject(ApplicationProvider._job)
            return
        ApplicationProvider._lastBackground := A_TickCount
        ApplicationProvider._job := ApplicationProvider._NewJob()
        if (ApplicationProvider._scanTimer = "")
            ApplicationProvider._scanTimer := Logger.Watch("app scan", () => ApplicationProvider._ScanSome())
        SetTimer(ApplicationProvider._scanTimer, 10)
    }

    ; Every minute: rebuild in the background when the cache is out of date (a fallback; normally the folder watch does it). Waits for the next time while the search window is open. No need to wait for keyboard and mouse idle:
    ; scanning is split up and stops as soon as the window opens, and the final cache write also waits for idle (see _Finish)
    static _CheckRefresh() {
        if (IsObject(ApplicationProvider._job) || ApplicationProvider._savePending || !BackgroundWork.CanRun())
            return                                                          ; _savePending: just rebuilt, the cache file isn't written yet (its time is old)
        if ApplicationProvider._CacheExpired()
            ApplicationProvider.RefreshInBackground()
    }

    static _ScanSome() {
        job := ApplicationProvider._job
        if !IsObject(job)
            return SetTimer(ApplicationProvider._scanTimer, 0)
        if (ApplicationProvider.Apps.Length && !BackgroundWork.CanRun())    ; pause while the search window is open or opening, go on
            return                                                          ; when it closes (the cache write too); with no index yet (first run), scan anyway
        if !ApplicationProvider._Step(job, 15)
            return
        SetTimer(ApplicationProvider._scanTimer, 0)
        ApplicationProvider._job := ""
        ApplicationProvider._Finish(job, false)
        SetTimer(() => ApplicationProvider._SearchKeys(), -500)             ; compute the new search keys in advance too
    }

    static _StopJob() {
        if (ApplicationProvider._scanTimer != "")
            SetTimer(ApplicationProvider._scanTimer, 0)
        ApplicationProvider._job := ""
    }

    ; One scan: the Folders one by one in order, each with a stack of folders to scan instead of recursion, so it can stop any time and go on later.
    ; Each folder is listed once and the file names compared with FileTypes (wildcards) here (listing once per type is several times slower)
    static _NewJob() {
        options := AppSettings.Feature("Applications")
        roots := []
        for folder in options["Folders"] {
            folder := RTrim(Path.Resolve(folder), "\")
            if (folder != "" && DirExist(folder))
                roots.Push(folder)
        }
        found := Map()
        found.CaseSense := "Off"
        return {Options: options, Roots: roots, Next: 1, Stack: [], Types: ApplicationProvider._TypesPattern(options), Found: found}
    }

    ; Scans a part: stops after budget ms (0 = all at once); true when done
    static _Step(job, budget) {
        options := job.Options, exclude := options["Exclude"], maxDepth := options["Depth"]
        start := A_TickCount
        loop {
            if !job.Stack.Length {
                if (job.Next > job.Roots.Length)
                    return true
                job.Stack := [[job.Roots[job.Next++], 0]]
            }
            entry := job.Stack.Pop()
            folder := entry[1], depth := entry[2]
            subfolders := []
            Loop Files, folder "\*", "FD" {
                if InStr(A_LoopFileAttrib, "D") {
                    if (depth < maxDepth)
                        subfolders.Push(A_LoopFileFullPath)
                    continue
                }
                if !RegExMatch(A_LoopFileName, job.Types)
                    continue
                SplitPath(A_LoopFileName, , , &ext, &title)
                if (exclude != "" && RegExMatch(title, exclude))
                    continue
                detail := ""
                if (ext = "lnk") {                                          ; a program with shortcuts in several places is kept once
                    try FileGetShortcut(A_LoopFileFullPath, &linkTarget, , &linkArgs)
                    catch                                                   ; a broken shortcut: on an error both variables are left unassigned
                        linkTarget := "", linkArgs := ""
                    detail := Trim(linkTarget " " linkArgs)
                    key := StrLower(title "|" detail)
                } else {
                    key := StrLower(title "|" A_LoopFileFullPath)
                }
                if !job.Found.Has(key)
                    job.Found[key] := ApplicationProvider._Entry(title, A_LoopFileFullPath, detail, options)
            }
            Loop subfolders.Length                                          ; pushed in reverse, so the earlier ones are scanned first
                job.Stack.Push([subfolders[subfolders.Length - A_Index + 1], depth + 1])
            if (budget && A_TickCount - start >= budget)
                return false
        }
    }

    ; Done: swap in the new Apps and write the cache; Store apps are always scanned for a manual rebuild, at most once an hour in the background
    static _Finish(job, manual) {
        apps := []
        for key, entry in job.Found
            apps.Push(entry)
        ApplicationProvider._MergeStoreApps(apps)                          ; keep the previous Store apps and replace them when the background update finishes
        ApplicationProvider.Apps := apps
        if manual
            ApplicationProvider._SaveCache()
        else
            ApplicationProvider._SaveWhenIdle()
        if (job.Options["StoreApps"] && (manual || !ApplicationProvider._storeScanned
                || A_TickCount - ApplicationProvider._storeScanned >= 3600000))
            ApplicationProvider._StartStoreScan()
        Logger.Debug("ApplicationProvider: indexed " apps.Length " apps" (manual ? "" : " (background)"))
        return apps.Length
    }

    static _Entry(title, target, detail, options) {
        search := options["MatchPinyin"] ? Pinyin.Initials(title) : ""
        return Map("Title", title, "Target", target, "Detail", detail, "Search", (search != title) ? search : "")
    }

    ;---------------------------------------------------------------------------
    ; Microsoft Store apps (PowerShell in the background)
    ;---------------------------------------------------------------------------
    static _StartStoreScan() {
        if (ApplicationProvider._storePid && ProcessExist(ApplicationProvider._storePid))
            return
        outFile := A_Temp "\ALTRun_StoreApps.csv"
        try FileDelete(outFile)
        command := "powershell.exe -NoProfile -Command `"Get-StartApps | Select-Object Name, AppID | ConvertTo-Csv -NoTypeInformation | Out-File -Encoding UTF8 -FilePath '" outFile "'`""
        try {
            Run(command, , "Hide", &pid)
        } catch as e {
            Logger.Error("ApplicationProvider: cannot start Get-StartApps - " e.Message)
            return
        }
        ApplicationProvider._storePid := pid, ApplicationProvider._storeScanned := A_TickCount
        ApplicationProvider._storeFile := outFile
        if (ApplicationProvider._storeTimer = "")
            ApplicationProvider._storeTimer := Logger.Watch("store apps scan", () => ApplicationProvider._CheckStoreScan())
        SetTimer(ApplicationProvider._storeTimer, 500)
    }

    static _CheckStoreScan() {
        if ProcessExist(ApplicationProvider._storePid)
            return
        SetTimer(ApplicationProvider._storeTimer, 0)
        ApplicationProvider._storePid := 0
        csvFile := ApplicationProvider._storeFile
        if !FileExist(csvFile)
            return
        options := AppSettings.Feature("Applications")
        storeApps := []
        for line in StrSplit(FileRead(csvFile, "UTF-8"), "`n", "`r") {
            if (A_Index = 1 || !RegExMatch(line, '^"(.*)","(.*)"$', &m))    ; the first line is the header
                continue
            if (m[1] = "" || m[2] = "" || RegExMatch(m[1], options["Exclude"]))
                continue
            if !InStr(m[2], "!")                                            ; only real Store apps of the form "package!app"; normal programs are already indexed from the Start menu
                continue
            storeApps.Push(ApplicationProvider._Entry(m[1], "shell:AppsFolder\" m[2], "", options))
        }
        try FileDelete(csvFile)

        apps := []
        for entry in ApplicationProvider.Apps
            if (InStr(entry["Target"], "shell:AppsFolder\") != 1)
                apps.Push(entry)
        for entry in storeApps
            apps.Push(entry)
        ApplicationProvider.Apps := apps
        ApplicationProvider._SaveCache()
        Logger.Debug("ApplicationProvider: indexed " storeApps.Length " store apps")
    }

    static _MergeStoreApps(apps) {
        for entry in ApplicationProvider.Apps
            if (InStr(entry["Target"], "shell:AppsFolder\") = 1)
                apps.Push(entry)
    }

    ;---------------------------------------------------------------------------
    ; Cache
    ;---------------------------------------------------------------------------
    static _LoadCache() {
        if !FileExist(ApplicationProvider.IndexFile)
            return false
        try {
            data := JSON.Parse(FileRead(ApplicationProvider.IndexFile, "UTF-8"))
            ApplicationProvider.Apps := data["Apps"]
            return true
        } catch as e {
            Logger.Error("ApplicationProvider: cannot read cache - " e.Message)
            return false
        }
    }

    ; Whether the cache is older than FallbackMinutes; no cache file counts as out of date
    ; Makes the cache count as out of date (the scan scope changed in Preferences): the next start-up still reads it first, then rescans in the background.
    ; A pending cache write is dropped too (scanned with the old scope; writing it on exit would make its time current again)
    static ExpireCache() {
        ApplicationProvider._savePending := false
        try FileSetTime("20000101000000", ApplicationProvider.IndexFile, "M")
    }

    static _CacheExpired() {
        minutes := ApplicationProvider.FallbackMinutes
        try return DateDiff(A_Now, FileGetTime(ApplicationProvider.IndexFile, "M"), "Minutes") >= minutes
        return true
    }

    ; The cache write after a background rebuild: waits until keyboard and mouse have been idle for SaveIdleMs and the search window is closed, so it never stalls while the user works
    static _SaveWhenIdle() {
        ApplicationProvider._savePending := true
        if (ApplicationProvider._saveTimer = "")
            ApplicationProvider._saveTimer := Logger.Watch("app index save", () => ApplicationProvider._SaveIfIdle())
        SetTimer(ApplicationProvider._saveTimer, 1000)
    }

    static _SaveIfIdle() {
        if (A_TimeIdlePhysical < ApplicationProvider.SaveIdleMs || !BackgroundWork.CanRun())
            return
        SetTimer(ApplicationProvider._saveTimer, 0)
        ApplicationProvider.SavePending()
    }

    ; write a pending cache now (on exit)
    static SavePending() {
        if !ApplicationProvider._savePending
            return
        ApplicationProvider._savePending := false
        ApplicationProvider._SaveCache()
    }

    static _SaveCache() {
        ApplicationProvider._savePending := false
        try {
            JSON.WriteFile(ApplicationProvider.IndexFile, Map("Apps", ApplicationProvider.Apps))
        } catch as e {
            Logger.Error("ApplicationProvider: cannot write cache - " e.Message)
        }
    }
}
