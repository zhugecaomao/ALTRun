;===============================================================================
; ApplicationProvider.ahk - 应用程序搜索 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 索引开始菜单、桌面 (Features.Applications.Folders) 里的快捷方式和程序, 以及
; 应用商店应用 (PowerShell Get-StartApps, 在后台运行不卡界面)。中文名称同时
; 按拼音首字母匹配 ("wx" -> 微信)。
;
; 索引缓存在 Data\AppIndex.json, 启动时直接读缓存, 超过 RefreshMinutes 分钟
; 才在后台重新扫描; 托盘菜单 "重建索引" 或系统命令 "Rebuild ALTRun Index"
; 会立即重建。
;
; 设置 (ALTRun.json -> Features.Applications):
;   Folders / FileTypes / Depth / Exclude / StoreApps / MatchPinyin / RefreshMinutes
;===============================================================================

class ApplicationProvider {
    static Id        := "Applications"
    static IndexFile := A_ScriptDir "\Data\AppIndex.json"
    static Apps      := []            ; [Map("Title", "Target", "Detail", "Search")...]
    static _storePid := 0, _storeFile := "", _storeTimer := ""

    static Init() {
        if !ApplicationProvider._LoadCache() || ApplicationProvider._CacheExpired()
            SetTimer(() => ApplicationProvider.Rebuild(), -3000)            ; 先让窗口出来, 再在后台扫描
    }

    static Search(query) {
        results := []
        if (StrLen(query.Text) < 1)
            return results
        for entry in ApplicationProvider.Apps {
            score := FuzzyMatcher.Best(query.Text, [entry["Title"], entry["Search"]])
            if (score <= 0)
                continue
            isStore := InStr(entry["Target"], "shell:AppsFolder\") = 1
            subtitle := isStore ? I18n.T("App.Subtitle.Store") : (entry["Detail"] != "") ? entry["Detail"] : entry["Target"]
            results.Push(ResultItem(entry["Title"], subtitle, {
                Kind: "file", Arg: entry["Target"],
                Icon: entry["Target"], Uid: "app:" StrLower(entry["Target"]), Score: score + 10
            }))
        }
        return results
    }

    ; 重新扫描, 返回找到的应用数量
    static Rebuild() {
        options := AppSettings.Feature("Applications")
        found := Map()
        found.CaseSense := "Off"
        for folder in options["Folders"]
            ApplicationProvider._ScanFolder(Path.Resolve(folder), options, found)

        apps := []
        for key, entry in found
            apps.Push(entry)
        ApplicationProvider._MergeStoreApps(apps)                          ; 保留上一次的商店应用, 后台更新完成后再替换
        ApplicationProvider.Apps := apps
        ApplicationProvider._SaveCache()
        if options["StoreApps"]
            ApplicationProvider._StartStoreScan()
        Logger.Debug("ApplicationProvider: indexed " apps.Length " apps")
        return apps.Length
    }

    static _ScanFolder(folder, options, found) {
        folder := RTrim(folder, "\")
        if (folder = "" || !DirExist(folder))
            return
        exclude := options["Exclude"]
        maxDepth := options["Depth"]
        for pattern in options["FileTypes"] {
            Loop Files, folder "\" pattern, "R" {
                relative := SubStr(A_LoopFileFullPath, StrLen(folder) + 2)
                depth := StrLen(relative) - StrLen(StrReplace(relative, "\"))
                if (depth > maxDepth)
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
                if !found.Has(key)
                    found[key] := ApplicationProvider._Entry(title, A_LoopFileFullPath, detail, options)
            }
        }
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
        ApplicationProvider._storePid := pid
        ApplicationProvider._storeFile := outFile
        if (ApplicationProvider._storeTimer = "")
            ApplicationProvider._storeTimer := () => ApplicationProvider._CheckStoreScan()
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

    static _CacheExpired() {
        minutes := AppSettings.Feature("Applications")["RefreshMinutes"]
        return DateDiff(A_Now, FileGetTime(ApplicationProvider.IndexFile, "M"), "Minutes") >= minutes
    }

    static _SaveCache() {
        try {
            DirCreate(AppSettings.DataDir)
            FileOpen(ApplicationProvider.IndexFile, "w", "UTF-8").Write(JSON.Stringify(Map("Apps", ApplicationProvider.Apps)))
        } catch as e {
            Logger.Error("ApplicationProvider: cannot write cache - " e.Message)
        }
    }
}
