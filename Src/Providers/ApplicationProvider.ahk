;===============================================================================
; ApplicationProvider.ahk - 应用程序搜索 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 索引开始菜单、桌面 (Features.Applications.Folders) 里的快捷方式和程序, 以及
; 应用商店应用 (PowerShell Get-StartApps, 在后台运行不卡界面)。中文名称同时
; 按拼音首字母匹配 ("wx" -> 微信)。
;
; 索引缓存在 Data\AppIndex.json, 启动时直接读缓存。在后台分批重新扫描 (每次定时器最多约 15 ms, 不挡热键),
; 不弹通知, 扫描中搜索窗口打开 (或正在呼出) 就暂停 (已经有旧索引可用时), 关了再继续, 不和呼出、搜索抢时间:
;   - 监视 Folders (包括子文件夹, Windows 的文件夹变化通知): 有文件新增、删除、改名, 安静 WatchQuietMs 后重建
;     (装软件时会连续建好几个快捷方式), 两次之间至少隔 WatchMinGapMs。新装的软件几秒后就能搜到
;   - 兜底 (监视不到的: ALTRun 没运行时装的软件、启动时还不存在或监视不了的文件夹): 缓存超过
;     FallbackMinutes 分钟也重建, 启动时和之后每分钟检查一次 (以前是设置 RefreshMinutes, 现在不再读取)
;   - 后台重建后的写缓存等键盘鼠标空闲一会儿再写 (写 JSON 的那一下不分段), 退出时还没写就马上写
; 商店应用 (PowerShell) 后台最多每小时扫一次。
; 托盘菜单 "重建索引" 或系统命令 "Rebuild ALTRun Index" 立即重建 (一次做完, 显示结果)。
;
; 不需要的应用: 在搜索结果里按 Ctrl+Del (或右键 "删除") 把它加入 Hidden 列表, 不再显示,
; 重建索引后也不会回来 (不会卸载程序, 也不改 AppIndex.json)。偏好设置 -> 应用搜索里
; 删掉 Hidden 的一行即可恢复。
;
; 设置 (ALTRun.json -> Features.Applications):
;   Folders / FileTypes / Depth / Exclude / Hidden / StoreApps / MatchPinyin
;===============================================================================

class ApplicationProvider {
    static Id        := "Applications"
    static IndexFile := AppSettings.DataDir "\AppIndex.json"
    static Apps      := []            ; [Map("Title", "Target", "Detail", "Search")...]
    static _storePid := 0, _storeFile := "", _storeTimer := "", _storeScanned := 0
    static _job := "", _scanTimer := "", _refreshTimer := ""         ; 后台重建: 见 _NewJob
    static WatchQuietMs := 5000, WatchMinGapMs := 60000, SaveIdleMs := 2000, FallbackMinutes := 60, WatchRetryMs := 30000
    static _watches := [], _watchTimer := "", _watchTypes := ""
    static _changedAt := 0, _lastBackground := 0, _savePending := false, _saveTimer := ""
    static _entries := [], _keys := [], _keysFor := "", _lastNeedle := "", _lastMatches := []

    static Init() {
        if !ApplicationProvider._LoadCache() || ApplicationProvider._CacheExpired()
            SetTimer(() => ApplicationProvider.RefreshInBackground(), -3000)   ; 先让窗口出来, 再在后台扫描
        SetTimer(() => ApplicationProvider._SearchKeys(), -500)             ; 预先算好搜索 Key, 第一次搜索不用等
        ApplicationProvider._refreshTimer := Logger.Watch("app refresh check", () => ApplicationProvider._CheckRefresh())
        SetTimer(ApplicationProvider._refreshTimer, 60000)
        ApplicationProvider.StartWatching()
    }

    ;---------------------------------------------------------------------------
    ; 监视索引的文件夹 (ReadDirectoryChangesW, 异步, 每 2 秒看一眼不等待): 应用类型的文件 (FileTypes) 或文件夹
    ; 新增、删除、改名、修改时在后台重建。能拿到文件名, 所以桌面上别的文件变了不算; ALTRun 自己的文件夹也不算
    ; (放在桌面上时, 写缓存、日志会一直触发)
    ;---------------------------------------------------------------------------
    static StartWatching() {
        ApplicationProvider.StopWatching()
        options := AppSettings.Feature("Applications")
        for folder in options["Folders"] {
            folder := RTrim(Path.Resolve(folder), "\\")
            if (folder = "" || !DirExist(folder))
                continue
            watch := ApplicationProvider._CreateWatch(folder)
            if IsObject(watch)
                ApplicationProvider._watches.Push(watch)
            else
                Logger.Debug("ApplicationProvider: cannot watch " folder)
        }
        ApplicationProvider._watchTypes := ApplicationProvider._TypesPattern(options)
        if (ApplicationProvider._watchTimer = "")
            ApplicationProvider._watchTimer := Logger.Watch("app folder watch", () => ApplicationProvider._PollWatches())
        if ApplicationProvider._watches.Length
            SetTimer(ApplicationProvider._watchTimer, 2000)
    }

    static StopWatching() {
        if (ApplicationProvider._watchTimer != "")
            SetTimer(ApplicationProvider._watchTimer, 0)
        for watch in ApplicationProvider._watches
            ApplicationProvider._CloseWatch(watch)
        ApplicationProvider._watches := [], ApplicationProvider._changedAt := 0
    }

    static _CloseWatch(watch) {
        DllCall("CancelIoEx", "Ptr", watch.Handle, "Ptr", 0)                ; 取消还在等的读取, 之后才能放掉缓冲区
        DllCall("CloseHandle", "Ptr", watch.Handle)
        if watch.Event
            DllCall("CloseHandle", "Ptr", watch.Event)
    }

    ; Open a fresh watcher for recovery after a completed read cannot be re-armed.
    static _CreateWatch(folder) {
        static FILE_LIST_DIRECTORY := 0x1, SHARE_ALL := 0x7, OPEN_EXISTING := 3, BACKUP_SEMANTICS := 0x02000000, OVERLAPPED := 0x40000000
        if !DirExist(folder)
            return ""
        handle := DllCall("CreateFileW", "Str", folder, "UInt", FILE_LIST_DIRECTORY, "UInt", SHARE_ALL, "Ptr", 0, "UInt", OPEN_EXISTING
            , "UInt", BACKUP_SEMANTICS | OVERLAPPED, "Ptr", 0, "Ptr")
        if (handle = -1 || !handle)
            return ""
        watch := {Folder: folder, Handle: handle, Event: DllCall("CreateEventW", "Ptr", 0, "Int", true, "Int", false, "Ptr", 0, "Ptr")
            , Overlapped: Buffer(A_PtrSize * 2 + 8 + A_PtrSize, 0), Buffer: Buffer(32768)}
        if !watch.Event {
            DllCall("CloseHandle", "Ptr", handle)
            return ""
        }
        NumPut("Ptr", watch.Event, watch.Overlapped, A_PtrSize * 2 + 8)
        if !ApplicationProvider._ReadChanges(watch) {
            ApplicationProvider._CloseWatch(watch)
            return ""
        }
        return watch
    }

    ; 开始 (再次) 异步等这个文件夹 (包括子文件夹) 的下一批变化: 文件名、文件夹名、修改时间
    static _ReadChanges(watch) {
        DllCall("ResetEvent", "Ptr", watch.Event)
        return DllCall("ReadDirectoryChangesW", "Ptr", watch.Handle, "Ptr", watch.Buffer, "UInt", watch.Buffer.Size, "Int", true
            , "UInt", 0x1 | 0x2 | 0x10, "Ptr", 0, "Ptr", watch.Overlapped, "Ptr", 0)
    }

    ; 每 2 秒: 读出已经到了的变化 (不等待), 有应用相关的就记下时间; 安静 WatchQuietMs、离上一次后台重建至少 WatchMinGapMs 时重建
    static _PollWatches() {
        for index, watch in ApplicationProvider._watches {
            if (watch.HasOwnProp("Disabled") && watch.Disabled) {
                if (A_TickCount < watch.RetryAt)
                    continue
                replacement := ApplicationProvider._CreateWatch(watch.Folder)
                if IsObject(replacement) {
                    ApplicationProvider._CloseWatch(watch)
                    ApplicationProvider._watches[index] := replacement
                    Logger.Debug("ApplicationProvider: watch recovered for " replacement.Folder)
                } else {
                    watch.RetryAt := A_TickCount + ApplicationProvider.WatchRetryMs
                }
                continue
            }
            if (DllCall("WaitForSingleObject", "Ptr", watch.Event, "UInt", 0, "UInt") != 0)
                continue
            bytes := 0
            if !DllCall("GetOverlappedResult", "Ptr", watch.Handle, "Ptr", watch.Overlapped, "UInt*", &bytes, "Int", false) {
                error := A_LastError
                Logger.Debug("ApplicationProvider: watch read failed for " watch.Folder " (error " error ")")
                relevant := true                                             ; Errors may have lost changes; refresh conservatively
            } else {
                names := ApplicationProvider._ChangedNames(watch.Buffer, bytes)
                relevant := (bytes = 0)                                     ; 0 = changes overflowed the buffer
                for name in names
                    if ApplicationProvider.IsWatchedChange(watch.Folder "\\" name, ApplicationProvider._watchTypes) {
                        relevant := true
                        break
                    }
            }
            if relevant
                ApplicationProvider._changedAt := A_TickCount
            if !ApplicationProvider._ReadChanges(watch) {
                watch.Disabled := true
                watch.RetryAt := A_TickCount + ApplicationProvider.WatchRetryMs
                Logger.Debug("ApplicationProvider: watch re-arm failed for " watch.Folder "; retry scheduled")
            }
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

    ; FILE_NOTIFY_INFORMATION 链表 -> 相对路径的数组
    static _ChangedNames(buffer, bytes) {
        names := [], offset := 0
        while (bytes && offset < bytes) {
            next := NumGet(buffer, offset, "UInt"), length := NumGet(buffer, offset + 8, "UInt")
            names.Push(StrGet(buffer.Ptr + offset + 12, length // 2, "UTF-16"))
            if !next
                break
            offset += next
        }
        return names
    }

    ; 纯函数: 这个路径的变化要不要重建 -- 应用类型的文件 (types: _TypesPattern), 或没有扩展名的 (多半是文件夹,
    ; 例如开始菜单里一个程序的文件夹被删掉 / 改名); ALTRun 自己文件夹里的不算
    static IsWatchedChange(fullPath, types, ownDir := A_ScriptDir) {
        if (ownDir != "" && InStr(fullPath, RTrim(ownDir, "\\") "\\") = 1)
            return false
        if DirExist(fullPath)
            return true
        SplitPath(fullPath, &name, , &ext)
        return (ext = "" || RegExMatch(name, types)) ? true : false
    }

    ; FileTypes ("*.lnk"...) -> 一个正则 ("i)^(?:.*\.lnk|...)$")
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

        ; 继续输入 (新输入以上一次的输入开头) 时只需要在上一次匹配到的里面找。
        ; 3 个字母以下时 "按顺序出现的字母" 规则不生效, 那时的结果不能用来缩小范围。
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
            ranks[index] := score + Knowledge.Boost(query.Text, "app:" StrLower(apps[index]["Target"]))   ; 常选的应用不会被挤出前几名
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

    ; 启动后空闲时先算好搜索 Key, 第一次输入不用等
    static Warm() {
        ApplicationProvider._SearchKeys()
    }

    ; 启动后在后台预读图标用 (见 ProviderRegistry.PrefetchIcons): 没有隐藏的应用
    static IconItems() {
        ApplicationProvider._SearchKeys()
        items := []
        for entry in ApplicationProvider._entries
            items.Push({Icon: entry["Target"], Uid: "app:" StrLower(entry["Target"])})
        return items
    }

    ; 没有隐藏的应用 (_entries) 和它们的搜索 Key (名称 + 拼音首字母);
    ; Apps 换成新数组或隐藏列表变化时重新计算
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
    ; 隐藏不需要的应用 (搜索结果里 Ctrl+Del / 右键 "删除")
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

    ; 隐藏列表 -> Map (不区分大小写), 列表变化 (换了数组或条数变了) 时重新生成
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

    ; 立即重新扫描 (一次做完), 返回找到的应用数量。正在后台重建的话先停掉
    static Rebuild() {
        ApplicationProvider._StopJob()
        job := ApplicationProvider._NewJob()
        ApplicationProvider._Step(job, 0)
        return ApplicationProvider._Finish(job, true)
    }

    ; 在后台分批重新扫描, 扫完再替换 Apps (中途搜索用的还是旧的); 已经在扫的话不重复
    static RefreshInBackground() {
        if IsObject(ApplicationProvider._job)
            return
        ApplicationProvider._lastBackground := A_TickCount
        ApplicationProvider._job := ApplicationProvider._NewJob()
        if (ApplicationProvider._scanTimer = "")
            ApplicationProvider._scanTimer := Logger.Watch("app scan", () => ApplicationProvider._ScanSome())
        SetTimer(ApplicationProvider._scanTimer, 10)
    }

    ; 每分钟: 缓存过期了就在后台重建 (兜底, 平时靠监视文件夹)。搜索窗口开着时等下一次。不用等键盘鼠标空闲:
    ; 扫描是分段的, 窗口一打开就停, 最后写缓存也等空闲再写 (见 _Finish)
    static _CheckRefresh() {
        if (IsObject(ApplicationProvider._job) || ApplicationProvider._savePending || !SearchWindow._CanPrefetch())
            return                                                          ; _savePending: 刚重建完, 缓存文件还没写 (修改时间是旧的)
        if ApplicationProvider._CacheExpired()
            ApplicationProvider.RefreshInBackground()
    }

    static _ScanSome() {
        job := ApplicationProvider._job
        if !IsObject(job)
            return SetTimer(ApplicationProvider._scanTimer, 0)
        if (ApplicationProvider.Apps.Length && !SearchWindow._CanPrefetch())   ; 搜索窗口开着 / 正在呼出时先停, 关了再继续 (包括最后写缓存);
            return                                                          ; 还没有索引 (第一次运行) 时照常扫
        if !ApplicationProvider._Step(job, 15)
            return
        SetTimer(ApplicationProvider._scanTimer, 0)
        ApplicationProvider._job := ""
        ApplicationProvider._Finish(job, false)
        SetTimer(() => ApplicationProvider._SearchKeys(), -500)             ; 新的搜索 Key 也先算好
    }

    static _StopJob() {
        if (ApplicationProvider._scanTimer != "")
            SetTimer(ApplicationProvider._scanTimer, 0)
        ApplicationProvider._job := ""
    }

    ; 一次扫描: 按 Folders 的顺序一个一个扫, 每个用一个待扫描文件夹的栈代替递归, 可以随时停下, 下次接着扫。
    ; 每个文件夹只列一次, 文件名和 FileTypes (通配符) 自己比 (分别按每种类型列目录要慢好几倍)
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

    ; 扫一部分: 用时超过 budget 毫秒就停 (0 = 一次扫完); 扫完返回 true
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
                if (ext = "lnk") {                                          ; 同一个程序在多个位置有快捷方式时只保留一个
                    linkTarget := "", linkArgs := ""
                    try FileGetShortcut(A_LoopFileFullPath, &linkTarget, , &linkArgs)
                    detail := Trim(linkTarget " " linkArgs)
                    key := StrLower(title "|" detail)
                } else {
                    key := StrLower(title "|" A_LoopFileFullPath)
                }
                if !job.Found.Has(key)
                    job.Found[key] := ApplicationProvider._Entry(title, A_LoopFileFullPath, detail, options)
            }
            Loop subfolders.Length                                          ; 倒着压栈, 先扫前面的
                job.Stack.Push([subfolders[subfolders.Length - A_Index + 1], depth + 1])
            if (budget && A_TickCount - start >= budget)
                return false
        }
    }

    ; 扫完: 换上新的 Apps, 写缓存; 商店应用手动重建时总是扫, 后台最多每小时一次
    static _Finish(job, manual) {
        apps := []
        for key, entry in job.Found
            apps.Push(entry)
        ApplicationProvider._MergeStoreApps(apps)                          ; 保留上一次的商店应用, 后台更新完成后再替换
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
    ; Microsoft Store apps (后台 PowerShell)
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
            if (A_Index = 1 || !RegExMatch(line, '^"(.*)","(.*)"$', &m))    ; 第一行是表头
                continue
            if (m[1] = "" || m[2] = "" || RegExMatch(m[1], options["Exclude"]))
                continue
            if !InStr(m[2], "!")                                            ; 只要 "包名!应用" 形式的真正商店应用, 普通程序已在开始菜单里索引过
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

    ; 缓存超过 FallbackMinutes 分钟了吗; 没有缓存文件也算过期
    static _CacheExpired() {
        minutes := ApplicationProvider.FallbackMinutes
        try return DateDiff(A_Now, FileGetTime(ApplicationProvider.IndexFile, "M"), "Minutes") >= minutes
        return true
    }

    ; 后台重建后的写缓存: 等键盘鼠标空闲 SaveIdleMs 并且搜索窗口没开时再写, 不在用户正操作时卡那一下
    static _SaveWhenIdle() {
        ApplicationProvider._savePending := true
        if (ApplicationProvider._saveTimer = "")
            ApplicationProvider._saveTimer := Logger.Watch("app index save", () => ApplicationProvider._SaveIfIdle())
        SetTimer(ApplicationProvider._saveTimer, 1000)
    }

    static _SaveIfIdle() {
        if (A_TimeIdlePhysical < ApplicationProvider.SaveIdleMs || !SearchWindow._CanPrefetch())
            return
        SetTimer(ApplicationProvider._saveTimer, 0)
        ApplicationProvider.SavePending()
    }

    ; 还没写的缓存马上写 (退出时)
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
