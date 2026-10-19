;===============================================================================
; FileIndex.ahk - 内置的文件 / 文件夹索引 (Everything 没有运行时使用) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 索引 Features.FileSearch.ScopeFolders 里的文件夹 (默认: 桌面、文档、下载),
; 深度 ScopeDepth 层, 包括文件夹和文件, 最多 MaxEntries 项。
;
; 扫描在后台分批进行 (每次定时器最多约 15 ms, 用一个待扫描文件夹的栈代替递归),
; 扫描期间界面照常响应; 结果缓存在 Data\FileIndex.json, 下次启动直接读。扫描中搜索窗口打开 (或正在呼出)
; 就暂停 (已经有旧索引可用时), 关了再继续; 扫完后的写缓存 (3 万条约 0.8 秒, 不分段) 等键盘鼠标空闲
; SaveIdleMs 再写, 退出时还没写就马上写。什么时候重新扫描:
;   - 监视 ScopeFolders (FolderWatcher): 有文件 / 文件夹新增、删除、改名 (修改文件不算; 也不算临时文件、比 ScopeDepth 深的、
;     ScopeExclude 的、ALTRun 自己的文件夹), 安静 WatchQuietMs 后重新扫描, 两次至少隔 WatchMinGapMs
;   - 兜底: 缓存超过 RefreshMinutes 分钟 (最少 5 分钟), 启动时和之后每分钟检查, 要等键盘鼠标空闲 RefreshIdleMs
;   - 用 Everything 时 (UseEverything 并且在运行) 都不扫描
; 启动时 Everything 正在运行 (并且 UseEverything) 就不读缓存也不扫描 (最多 3 万条, 用不上),
; 等真的要用内置索引时 (Everything 关掉了) 再读。
;
; 用法:
;   FileIndex.Start()                      启动时 (读缓存, 需要时开始后台扫描)
;   FileIndex.Search("report", 20)         -> [{Path, IsFolder, Score}] 按分数排序
;   FileIndex.Search("cs", 20)             中文文件名也按拼音首字母匹配 ("测试.bat"), 比同一档的直接匹配低 5 分
;===============================================================================

class FileIndex {
    static File     := AppSettings.DataDir "\FileIndex.json"
    static Paths    := [], Names := [], Folders := []     ; 三个平行数组: 完整路径 / 小写文件名 / 是否文件夹
    ; 中文文件名的小写拼音首字母: Map(下标, "cs.bat") (只有含中文的; 搜索时只需要看这些)。_initialsFor = 对应的 Names
    ; (换了索引、测试直接设 Names 时对不上, 不按拼音找)。读缓存后在后台分批算 (1 万个中文名约 0.25 秒), 不拖慢启动
    static Initials := Map(), _initialsFor := "", _initialsTimer := "", _initialsJob := ""
    static Scanning := false
    static _stack := [], _new := "", _timer := "", _lastNeedle := "", _lastMatches := ""
    static Deferred := false                                               ; 启动时有 Everything, 还没读缓存
    static _refreshTimer := ""
    static RefreshIdleMs := 30000                                          ; 定时重建要等键盘鼠标空闲这么久
    static WatchQuietMs := 5000, WatchMinGapMs := 300000, SaveIdleMs := 2000
    static _watcher := "", _watchTimer := "", _changedAt := 0, _lastScan := 0, _savePending := false, _saveTimer := ""
    ; 不当作变化的临时文件: 浏览器下载中、Office 打开文档时的锁文件、编辑器的临时文件
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

    ; 读缓存; 没有缓存或已经过期时 delay 毫秒后在后台重新扫描。开始监视文件夹 (用内置索引了)
    static _Load(delay) {
        FileIndex.Deferred := false
        FileIndex.StartWatching()
        if FileIndex._LoadCache() && !FileIndex._CacheExpired()
            return
        SetTimer(() => FileIndex.Rebuild(), -delay)
    }

    ;---------------------------------------------------------------------------
    ; 监视索引的文件夹: 有变化时在后台重新扫描 (见文件开头)
    ;---------------------------------------------------------------------------
    static StartWatching() {
        FileIndex.StopWatching()
        folders := []
        for folder in AppSettings.Feature("FileSearch")["ScopeFolders"]
            folders.Push(Path.Resolve(folder))
        FileIndex._watcher := FolderWatcher(folders, FolderWatcher.NameChanges)  ; 索引里只有路径: 修改文件不用重新扫描
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

    ; 每 2 秒: 有要算的变化就记下时间; 安静 WatchQuietMs、离上一次扫描至少 WatchMinGapMs、没在用 Everything 时重新扫描
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

    ; 纯函数: 这个路径的变化要不要重新扫描 -- 在某个索引的文件夹里、不比 depth 层深 (和扫描一样: 第一层是 0)、
    ; 不是临时文件、不匹配 exclude (正则, 对完整路径)、不在 ALTRun 自己的文件夹里
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

    ; 要用内置索引了 (Everything 没有运行): 启动时跳过的话现在读
    static EnsureLoaded() {
        if FileIndex.Deferred
            FileIndex._Load(10)
    }

    ; 每分钟: 在用内置索引 (已经读过、Everything 没在用) 并且缓存过期了, 就在后台重新扫描。
    ; 用户正在用电脑 (空闲不到 RefreshIdleMs) 或搜索窗口开着时等下一次
    static _CheckRefresh() {
        if (FileIndex.Scanning || FileIndex.Deferred || FileIndex._savePending || !SearchWindow._CanPrefetch() || A_TimeIdlePhysical < FileIndex.RefreshIdleMs)
            return                                                          ; _savePending: 刚扫完, 缓存文件还没写 (修改时间是旧的)
        options := AppSettings.Feature("FileSearch")
        if (options["UseEverything"] && Everything.IsRunning())
            return
        if FileIndex._CacheExpired()
            FileIndex.Rebuild()
    }

    static Rebuild() {
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
        if (FileIndex.Paths.Length && !SearchWindow._CanPrefetch())         ; 搜索窗口开着 / 正在呼出时先停, 关了再继续 (包括最后写缓存);
            return                                                          ; 还没有索引 (第一次运行) 时照常扫
        options := AppSettings.Feature("FileSearch")
        exclude := options["ScopeExclude"]
        start := A_TickCount
        found := FileIndex._new
        while (FileIndex._stack.Length && A_TickCount - start < 15) {
            entry := FileIndex._stack.Pop()
            folder := entry[1], depth := entry[2]
            try {
                Loop Files, folder "\*", "FD" {
                    if (A_LoopFileAttrib ~= "[HS]")                             ; 跳过隐藏 / 系统文件
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
    }

    ; 按文件名匹配: 完全相同 100 / 开头 90 / 单词开头 80 / 包含 60。
    ; 继续输入时只在上一次匹配到的里面找 (只用 "包含" 类规则, 范围只会缩小)。
    ; foldersOnly: 只要文件夹 ("folder bk")
    ; extensions: Map("pdf", true, ...) 时只要这些扩展名的文件 (文件类型筛选, 例如 "doc 报告")
    static Search(needle, limit, foldersOnly := false, extensions := "") {
        FileIndex.EnsureLoaded()
        needle := StrLower(Trim(needle))
        if (needle = "")
            return []
        candidates := (FileIndex._lastNeedle != "" && InStr(needle, FileIndex._lastNeedle) = 1) ? FileIndex._lastMatches : ""
        scores := Map(), matches := []
        names := FileIndex.Names, initials := FileIndex.Initials
        usePinyin := IsObject(FileIndex._initialsFor) && FileIndex._initialsFor == names && FileIndex.UsePinyin(needle)   ; 拼音: 首字母算好了, 输入是一个英文词
        if IsObject(candidates) {
            for index in candidates
                if (score := FileIndex.ScoreName(needle, names[index]))
                        || (usePinyin && initials.Has(index) && (score := FileIndex.PinyinScore(needle, initials[index])))
                    scores[index] := score, matches.Push(index)
        } else {
            for index, name in names
                if (score := FileIndex.ScoreName(needle, name))
                    scores[index] := score, matches.Push(index)
            if usePinyin                                                    ; 只看中文文件名, 先用 InStr 粗筛 (绝大多数对不上)
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

    ; 拼音首字母匹配: 只在输入是一个英文词时 (多个词 ScoreName 本来就对不上; 输入有中文时直接按原文匹配)
    static UsePinyin(needle) => !InStr(needle, " ") && !RegExMatch(needle, "[^\x{00}-\x{7f}]")

    ; 文件名的小写拼音首字母 ("测试.bat" -> "cs.bat"); 没有中文时 "" (不用另外匹配)
    static InitialsOf(name) => RegExMatch(name, "[^\x{00}-\x{ff}]") ? StrLower(Pinyin.Initials(name)) : ""

    ; 按拼音首字母打分: 和 ScoreName 同样的规则, 低 5 分 (同一档里名称直接匹配的在前), 最低 50; 对不上时 0
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

    ; 在后台分批算 Names 的拼音首字母 (每次定时器最多约 15 ms), 算完再换上; 期间换了索引 (重新扫描完) 就不要了
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
        FileIndex._lastNeedle := "", FileIndex._lastMatches := ""            ; 之前缩小范围时还没有拼音的结果
    }

    ; 缓存超过 RefreshMinutes 分钟 (最少 5 分钟) 了吗; 没有缓存文件也算过期
    ; 让缓存算作过期 (偏好设置里改了扫描范围): 下次启动照常先读它, 再在后台重新扫描。
    ; 还没写的缓存也不写了 (按旧的范围扫的; 退出时写的话修改时间又成了现在)
    static ExpireCache() {
        FileIndex._savePending := false
        try FileSetTime("20000101000000", FileIndex.File, "M")
    }

    static _CacheExpired() {
        minutes := Max(5, AppSettings.Feature("FileSearch")["RefreshMinutes"])
        try return DateDiff(A_Now, FileGetTime(FileIndex.File, "M"), "Minutes") >= minutes
        return true
    }

    ; 扫完后的写缓存: 等键盘鼠标空闲 SaveIdleMs 并且搜索窗口没开时再写, 不在用户正操作时卡那一下
    static _SaveWhenIdle() {
        FileIndex._savePending := true
        if (FileIndex._saveTimer = "")
            FileIndex._saveTimer := Logger.Watch("file index save", () => FileIndex._SaveIfIdle())
        SetTimer(FileIndex._saveTimer, 1000)
    }

    static _SaveIfIdle() {
        if (A_TimeIdlePhysical < FileIndex.SaveIdleMs || !SearchWindow._CanPrefetch())
            return
        SetTimer(FileIndex._saveTimer, 0)
        FileIndex.SavePending()
    }

    ; 还没写的缓存马上写 (退出时)
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
