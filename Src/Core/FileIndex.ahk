;===============================================================================
; FileIndex.ahk - 内置的文件 / 文件夹索引 (Everything 没有运行时使用) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 索引 Features.FileSearch.ScopeFolders 里的文件夹 (默认: 桌面、文档、下载),
; 深度 ScopeDepth 层, 包括文件夹和文件, 最多 MaxEntries 项。
;
; 扫描在后台分批进行 (每次定时器最多约 15 ms, 用一个待扫描文件夹的栈代替递归),
; 扫描期间界面照常响应; 结果缓存在 Data\FileIndex.json, 下次启动直接读,
; 超过 RefreshMinutes 分钟再重新扫描。
; 启动时 Everything 正在运行 (并且 UseEverything) 就不读缓存也不扫描 (最多 3 万条, 用不上),
; 等真的要用内置索引时 (Everything 关掉了) 再读。
;
; 用法:
;   FileIndex.Start()                      启动时 (读缓存, 需要时开始后台扫描)
;   FileIndex.Search("report", 20)         -> [{Path, IsFolder, Score}] 按分数排序
;===============================================================================

class FileIndex {
    static File     := A_ScriptDir "\Data\FileIndex.json"
    static Paths    := [], Names := [], Folders := []     ; 三个平行数组: 完整路径 / 小写文件名 / 是否文件夹
    static Scanning := false
    static _stack := [], _new := "", _timer := "", _lastNeedle := "", _lastMatches := ""
    static Deferred := false                                               ; 启动时有 Everything, 还没读缓存

    static Start() {
        options := AppSettings.Feature("FileSearch")
        if (options["UseEverything"] && Everything.IsRunning()) {
            FileIndex.Deferred := true
            Logger.Debug("FileIndex: Everything is running, built-in index not loaded")
            return
        }
        FileIndex._Load(5000)
    }

    ; 读缓存; 没有缓存或已经过期时 delay 毫秒后在后台重新扫描
    static _Load(delay) {
        FileIndex.Deferred := false
        if FileIndex._LoadCache() && !FileIndex._CacheExpired()
            return
        SetTimer(() => FileIndex.Rebuild(), -delay)
    }

    ; 要用内置索引了 (Everything 没有运行): 启动时跳过的话现在读
    static EnsureLoaded() {
        if FileIndex.Deferred
            FileIndex._Load(10)
    }

    static Rebuild() {
        if FileIndex.Scanning
            return
        FileIndex.Deferred := false
        options := AppSettings.Feature("FileSearch")
        FileIndex._stack := []
        for folder in options["ScopeFolders"] {
            resolved := RTrim(Path.Resolve(folder), "\")
            if (resolved != "" && DirExist(resolved))
                FileIndex._stack.Push([resolved, 0])
        }
        FileIndex._new := {Paths: [], Names: [], Folders: []}
        FileIndex.Scanning := true
        if (FileIndex._timer = "")
            FileIndex._timer := () => FileIndex._ScanSome()
        SetTimer(FileIndex._timer, 10)
    }

    static _ScanSome() {
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
        FileIndex._new := "", FileIndex.Scanning := false
        FileIndex._lastNeedle := "", FileIndex._lastMatches := ""
        FileIndex._SaveCache()
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
        names := FileIndex.Names
        if IsObject(candidates) {
            for index in candidates
                if (score := FileIndex.ScoreName(needle, names[index]))
                    scores[index] := score, matches.Push(index)
        } else {
            for index, name in names
                if (score := FileIndex.ScoreName(needle, name))
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
            return true
        } catch as e {
            Logger.Error("FileIndex: cannot read cache - " e.Message)
            return false
        }
    }

    static _CacheExpired() {
        minutes := AppSettings.Feature("FileSearch")["RefreshMinutes"]
        return DateDiff(A_Now, FileGetTime(FileIndex.File, "M"), "Minutes") >= minutes
    }

    static _SaveCache() {
        try {
            DirCreate(AppSettings.DataDir)
            FileOpen(FileIndex.File, "w", "UTF-8").Write(JSON.Stringify(Map("Paths", FileIndex.Paths, "Folders", FileIndex.Folders)))
        } catch as e {
            Logger.Error("FileIndex: cannot write cache - " e.Message)
        }
    }
}
