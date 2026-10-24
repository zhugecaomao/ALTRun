;===============================================================================
; FileIndex.ahk - the built-in file / folder index (used when Everything isn't running) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Indexes the folders in Features.FileSearch.ScopeFolders (default: Desktop, Documents, Downloads),
; ScopeDepth levels deep, folders and files, at most MaxEntries entries.
;
; Scanning runs in the background in batches (at most about 15 ms per timer tick, with a stack of folders to scan instead of recursion),
; so the UI keeps responding; the result is cached in Data\FileIndex.json and read at the next start-up. When the search window is open (or opening)
; during a scan it pauses (if an old index is available) and goes on after it closes; the cache write after a scan (about 0.8 s for 30,000 entries, in one go) waits
; until keyboard and mouse have been idle for SaveIdleMs, and is done at once on exit if still pending. When it rescans:
;   - watching ScopeFolders (FolderWatcher): files / folders added, deleted or renamed (not file modifications, temporary files, anything deeper than ScopeDepth,
;     matching ScopeExclude or in ALTRun's own folders); rescans after WatchQuietMs of quiet, at least WatchMinGapMs apart
;   - as a fallback: the cache is older than RefreshMinutes (at least 5), checked at start-up and every minute after, waiting for RefreshIdleMs of keyboard and mouse idle
;   - never while Everything is used (UseEverything and running)
; If Everything is running at start-up (and UseEverything), neither the cache is read nor a scan done (up to 30,000 entries that wouldn't be used);
; it is read when the built-in index is really needed (Everything was closed).
;
; Usage:
;   FileIndex.Start()                      at start-up (reads the cache, starts a background scan if needed)
;   FileIndex.Search("report", 20)         -> [{Path, IsFolder, Score}] sorted by score
;   FileIndex.Search("cs", 20)             Chinese file names also match by pinyin initials ("测试.bat"), 5 points below a direct match of the same tier
;===============================================================================

class FileIndex {
    static File     := AppSettings.DataDir "\FileIndex.json"
    static Paths    := [], Names := [], Folders := []     ; three parallel arrays: full path / lower-case file name / is a folder
    ; Lower-case pinyin initials of Chinese file names: Map(index, "cs.bat") (only names with Chinese; search only needs these). _initialsFor = the Names they belong to
    ; (after the index changed, or when tests set Names directly, they don't match and pinyin isn't used). Computed in the background in batches after reading the cache (about 0.25 s for 10,000 Chinese names), so start-up isn't slowed
    static Initials := Map(), _initialsFor := "", _initialsTimer := "", _initialsJob := ""
    static Scanning := false
    static _stack := [], _new := "", _timer := "", _lastNeedle := "", _lastMatches := ""
    static Deferred := false                                               ; Everything ran at start-up; the cache isn't read yet
    static _refreshTimer := "", _onDone := ""
    static RefreshIdleMs := 30000                                          ; a scheduled rebuild waits for keyboard and mouse to be idle this long
    static WatchQuietMs := 5000, WatchMinGapMs := 300000, SaveIdleMs := 2000
    static _watcher := "", _watchTimer := "", _changedAt := 0, _lastScan := 0, _savePending := false, _saveTimer := ""
    ; temporary files that don't count as changes: downloads in progress, Office lock files for open documents, editors' temporary files
    static TempNames := "i)(\.(crdownload|part|partial|download|tmp|temp)$|^~\$|^\.~lock\.|~$)"

    static Start() {
        FileIndex._refreshTimer := Logger.Watch("file index refresh check", () => FileIndex._CheckRefresh())
        SetTimer(FileIndex._refreshTimer, 60000)
        options := AppSettings.Feature("FileSearch")
        if (options["UseEverything"] && Everything.IsRunning()) {
            FileIndex.Deferred := true
            Logger.Debug("FileIndex: Everything is running, built-in index not loaded")
            return
        }
        FileIndex._Load(5000)
    }

    ; Reads the cache; without one, or if it is out of date, rescans in the background after delay ms. Starts watching the folders (the built-in index is in use)
    static _Load(delay) {
        FileIndex.Deferred := false
        FileIndex.StartWatching()
        if FileIndex._LoadCache() && !FileIndex._CacheExpired()
            return
        SetTimer(() => FileIndex.Rebuild(), -delay)
    }

    ;---------------------------------------------------------------------------
    ; Watch the indexed folders: rescan in the background on changes (see the top of the file)
    ;---------------------------------------------------------------------------
    static StartWatching() {
        FileIndex.StopWatching()
        folders := []
        for folder in AppSettings.Feature("FileSearch")["ScopeFolders"]
            folders.Push(Path.Resolve(folder))
        FileIndex._watcher := FolderWatcher(folders, FolderWatcher.NameChanges)  ; the index only has paths: modified files need no rescan
        if (FileIndex._watchTimer = "")
            FileIndex._watchTimer := Logger.Watch("file index watch", () => FileIndex._PollWatches())
        if FileIndex._watcher.Count
            SetTimer(FileIndex._watchTimer, 2000)
    }

    static StopWatching() {
        if (FileIndex._watchTimer != "")
            SetTimer(FileIndex._watchTimer, 0)
        if IsObject(FileIndex._watcher)
            FileIndex._watcher.Stop()
        FileIndex._watcher := "", FileIndex._changedAt := 0
    }

    ; Every 2 s: note the time of relevant changes; rescan after WatchQuietMs of quiet, at least WatchMinGapMs after the last scan, when Everything isn't used
    static _PollWatches() {
        if !IsObject(FileIndex._watcher)
            return
        options := AppSettings.Feature("FileSearch")
        if (changes := FileIndex._watcher.Poll()).Length {
            roots := []
            for watch in FileIndex._watcher.Watches
                roots.Push(watch.Folder)
            for fullPath in changes
                if (fullPath = "*" || FileIndex.IsWatchedChange(fullPath, roots, options["ScopeDepth"], options["ScopeExclude"])) {
                    FileIndex._changedAt := A_TickCount
                    break
                }
        }
        if (!FileIndex._changedAt || A_TickCount - FileIndex._changedAt < FileIndex.WatchQuietMs || FileIndex.Scanning)
            return
        if (FileIndex._lastScan && A_TickCount - FileIndex._lastScan < FileIndex.WatchMinGapMs)
            return
        if (FileIndex.Deferred || (options["UseEverything"] && Everything.IsRunning()))
            return
        FileIndex._changedAt := 0
        Logger.Debug("FileIndex: folders changed, rescanning")
        FileIndex.Rebuild()
    }

    ; Pure function: whether a change at this path needs a rescan -- inside an indexed folder, no deeper than depth levels (as in scanning: the first level is 0),
    ; not a temporary file, not matching exclude (a regex on the full path), not in ALTRun's own folders
    static IsWatchedChange(fullPath, roots, depth, exclude := "", ownDir := A_ScriptDir) {
        if (ownDir != "" && InStr(fullPath, RTrim(ownDir, "\") "\") = 1)
            return false
        SplitPath(fullPath, &name)
        if RegExMatch(name, FileIndex.TempNames)
            return false
        if (exclude != "" && RegExMatch(fullPath, exclude))
            return false
        for root in roots {
            if (InStr(fullPath, root "\") != 1)
                continue
            relative := SubStr(fullPath, StrLen(root) + 2)
            return (StrLen(relative) - StrLen(StrReplace(relative, "\")) < depth) ? true : false
        }
        return false
    }

    ; The built-in index is needed now (Everything isn't running): read it now if it was skipped at start-up
    static EnsureLoaded() {
        if FileIndex.Deferred
            FileIndex._Load(10)
    }

    ; Every minute: when the built-in index is in use (read, Everything not used) and the cache is out of date, rescan in the background.
    ; While the user is working (idle less than RefreshIdleMs) or the search window is open, wait for the next time
    static _CheckRefresh() {
        if (FileIndex.Scanning || FileIndex.Deferred || FileIndex._savePending || !BackgroundWork.CanRun() || A_TimeIdlePhysical < FileIndex.RefreshIdleMs)
            return                                                          ; _savePending: just scanned, the cache file isn't written yet (its time is old)
        options := AppSettings.Feature("FileSearch")
        if (options["UseEverything"] && Everything.IsRunning())
            return
        if FileIndex._CacheExpired()
            FileIndex.Rebuild()
    }

    ; onDone: optional, called as onDone(count) after the scan (in a new thread, so it may show a dialog); during a scan it waits for that one to finish
    static Rebuild(onDone := "") {
        if IsObject(onDone)
            FileIndex._onDone := onDone
        if FileIndex.Scanning
            return
        FileIndex.Deferred := false, FileIndex._lastScan := A_TickCount
        options := AppSettings.Feature("FileSearch")
        FileIndex._stack := []
        for folder in options["ScopeFolders"] {
            resolved := RTrim(Path.Resolve(folder), "\")
            if (resolved != "" && DirExist(resolved))
                FileIndex._stack.Push([resolved, 0])
        }
        FileIndex._new := {Paths: [], Names: [], Folders: [], Initials: Map()}
        FileIndex.Scanning := true
        if (FileIndex._timer = "")
            FileIndex._timer := Logger.Watch("file index scan", () => FileIndex._ScanSome())
        SetTimer(FileIndex._timer, 10)
    }

    static _ScanSome() {
        if (FileIndex.Paths.Length && !BackgroundWork.CanRun())             ; pause while the search window is open or opening, go on
            return                                                          ; when it closes (the cache write too); with no index yet (first run), scan anyway
        options := AppSettings.Feature("FileSearch")
        exclude := options["ScopeExclude"]
        start := A_TickCount
        found := FileIndex._new
        while (FileIndex._stack.Length && A_TickCount - start < 15) {
            entry := FileIndex._stack.Pop()
            folder := entry[1], depth := entry[2]
            try {
                Loop Files, folder "\*", "FD" {
                    if (A_LoopFileAttrib ~= "[HS]")                             ; skip hidden / system files
                        continue
                    if (exclude != "" && RegExMatch(A_LoopFileFullPath, exclude))
                        continue
                    isFolder := InStr(A_LoopFileAttrib, "D") ? true : false
                    found.Paths.Push(A_LoopFileFullPath)
                    found.Names.Push(StrLower(A_LoopFileName))
                    if ((initials := FileIndex.InitialsOf(A_LoopFileName)) != "")
                        found.Initials[found.Paths.Length] := initials
                    found.Folders.Push(isFolder)
                    if (isFolder && depth + 1 < options["ScopeDepth"])
                        FileIndex._stack.Push([A_LoopFileFullPath, depth + 1])
                    if (found.Paths.Length >= options["MaxEntries"]) {
                        FileIndex._stack := []
                        break
                    }
                }
            }
        }
        if FileIndex._stack.Length
            return
        SetTimer(FileIndex._timer, 0)
        FileIndex.Paths := found.Paths, FileIndex.Names := found.Names, FileIndex.Folders := found.Folders
        FileIndex.Initials := found.Initials, FileIndex._initialsFor := found.Names
        FileIndex._new := "", FileIndex.Scanning := false
        FileIndex._lastNeedle := "", FileIndex._lastMatches := ""
        FileIndex._SaveWhenIdle()
        Logger.Debug("FileIndex: " FileIndex.Paths.Length " entries")
        if IsObject(done := FileIndex._onDone) {
            FileIndex._onDone := "", count := FileIndex.Paths.Length
            SetTimer(() => done.Call(count), -1)                            ; no dialog inside the scan timer (it would count towards its time)
        }
    }

    ; Matches file names: exact 100 / start 90 / start of a word 80 / contains 60.
    ; While typing on, only the previous matches are searched (only "contains"-type rules, so the set can only shrink).
    ; foldersOnly: folders only ("folder budget")
    ; extensions: with Map("pdf", true, ...) only files with these extensions (file type filter, e.g. "doc report")
    static Search(needle, limit, foldersOnly := false, extensions := "") {
        FileIndex.EnsureLoaded()
        needle := StrLower(Trim(needle))
        if (needle = "")
            return []
        candidates := (FileIndex._lastNeedle != "" && InStr(needle, FileIndex._lastNeedle) = 1) ? FileIndex._lastMatches : ""
        scores := Map(), matches := []
        names := FileIndex.Names, initials := FileIndex.Initials
        usePinyin := IsObject(FileIndex._initialsFor) && FileIndex._initialsFor == names && FileIndex.UsePinyin(needle)   ; pinyin: the initials are ready and the input is one English word
        if IsObject(candidates) {
            for index in candidates
                if (score := FileIndex.ScoreName(needle, names[index]))
                        || (usePinyin && initials.Has(index) && (score := FileIndex.PinyinScore(needle, initials[index])))
                    scores[index] := score, matches.Push(index)
        } else {
            for index, name in names
                if (score := FileIndex.ScoreName(needle, name))
                    scores[index] := score, matches.Push(index)
            if usePinyin                                                    ; only Chinese file names, filtered roughly with InStr first (most don't match)
                for index, text in initials
                    if (InStr(text, needle) && !scores.Has(index) && (score := FileIndex.PinyinScore(needle, text)))
                        scores[index] := score, matches.Push(index)
        }
        FileIndex._lastNeedle := needle, FileIndex._lastMatches := matches
        if foldersOnly {
            folders := FileIndex.Folders
            for index in matches
                if !folders[index]
                    scores.Delete(index)
        }
        if IsObject(extensions) {
            folders := FileIndex.Folders, paths := FileIndex.Paths
            for index in matches {
                if folders[index] || !extensions.Has(StrLower(FileIndex._Ext(paths[index])))
                    scores.Delete(index)
            }
        }

        results := []
        for index in FuzzyMatcher.TopIndexes(scores, limit)
            results.Push({Path: FileIndex.Paths[index], IsFolder: FileIndex.Folders[index], Score: scores[index]})
        return results
    }

    static _Ext(filePath) {
        SplitPath(filePath, , , &ext)
        return ext
    }

    ; Pinyin initials only when the input is one English word (several words wouldn't match ScoreName anyway; input with Chinese matches the text itself)
    static UsePinyin(needle) => !InStr(needle, " ") && !RegExMatch(needle, "[^\x{00}-\x{7f}]")

    ; Lower-case pinyin initials of a file name ("测试.bat" -> "cs.bat"); "" without Chinese (no separate match needed)
    static InitialsOf(name) => RegExMatch(name, "[^\x{00}-\x{ff}]") ? StrLower(Pinyin.Initials(name)) : ""

    ; Score by pinyin initials: the same rules as ScoreName, 5 points lower (direct name matches of the same tier come first), at least 50; 0 if no match
    static PinyinScore(needle, initials) => (score := FileIndex.ScoreName(needle, initials)) ? Max(score - 5, 50) : 0

    static ScoreName(needle, name) {
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

    static _LoadCache() {
        if !FileExist(FileIndex.File)
            return false
        try {
            data := JSON.Parse(FileRead(FileIndex.File, "UTF-8"))
            FileIndex.Paths := data["Paths"], FileIndex.Folders := data["Folders"]
            names := []
            for fullPath in FileIndex.Paths {
                SplitPath(fullPath, &fileName)
                names.Push(StrLower(fileName))
            }
            FileIndex.Names := names
            FileIndex.BuildInitials()
            return true
        } catch as e {
            Logger.Error("FileIndex: cannot read cache - " e.Message)
            return false
        }
    }

    ; Computes the pinyin initials of Names in the background in batches (at most about 15 ms per timer tick) and swaps them in when done; dropped if the index changed meanwhile (a rescan finished)
    static BuildInitials() {
        FileIndex._initialsJob := {Names: FileIndex.Names, Pos: 0, Initials: Map()}
        if (FileIndex._initialsTimer = "")
            FileIndex._initialsTimer := Logger.Watch("file index pinyin", () => FileIndex._InitialsSome())
        SetTimer(FileIndex._initialsTimer, 10)
    }

    static _InitialsSome() {
        job := FileIndex._initialsJob
        if (!IsObject(job) || job.Names !== FileIndex.Names) {
            FileIndex._initialsJob := ""
            return SetTimer(FileIndex._initialsTimer, 0)
        }
        names := job.Names, count := names.Length, i := job.Pos, start := A_TickCount
        while (i < count && A_TickCount - start < 15) {
            i++
            if ((initials := FileIndex.InitialsOf(names[i])) != "")
                job.Initials[i] := initials
        }
        job.Pos := i
        if (i < count)
            return
        SetTimer(FileIndex._initialsTimer, 0)
        FileIndex.Initials := job.Initials, FileIndex._initialsFor := names, FileIndex._initialsJob := ""
        FileIndex._lastNeedle := "", FileIndex._lastMatches := ""            ; results without pinyin from narrowing down earlier
    }

    ; Whether the cache is older than RefreshMinutes (at least 5); no cache file counts as out of date
    ; Makes the cache count as out of date (the scan scope changed in Preferences): the next start-up still reads it first, then rescans in the background.
    ; A pending cache write is dropped too (scanned with the old scope; writing it on exit would make its time current again)
    static ExpireCache() {
        FileIndex._savePending := false
        try FileSetTime("20000101000000", FileIndex.File, "M")
    }

    static _CacheExpired() {
        minutes := Max(5, AppSettings.Feature("FileSearch")["RefreshMinutes"])
        try return DateDiff(A_Now, FileGetTime(FileIndex.File, "M"), "Minutes") >= minutes
        return true
    }

    ; The cache write after a scan: waits until keyboard and mouse have been idle for SaveIdleMs and the search window is closed, so it never stalls while the user works
    static _SaveWhenIdle() {
        FileIndex._savePending := true
        if (FileIndex._saveTimer = "")
            FileIndex._saveTimer := Logger.Watch("file index save", () => FileIndex._SaveIfIdle())
        SetTimer(FileIndex._saveTimer, 1000)
    }

    static _SaveIfIdle() {
        if (A_TimeIdlePhysical < FileIndex.SaveIdleMs || !BackgroundWork.CanRun())
            return
        SetTimer(FileIndex._saveTimer, 0)
        FileIndex.SavePending()
    }

    ; write a pending cache now (on exit)
    static SavePending() {
        if FileIndex._savePending
            FileIndex._SaveCache()
    }

    static _SaveCache() {
        FileIndex._savePending := false
        try {
            JSON.WriteFile(FileIndex.File, Map("Paths", FileIndex.Paths, "Folders", FileIndex.Folders))
        } catch as e {
            Logger.Error("FileIndex: cannot write cache - " e.Message)
        }
    }
}
