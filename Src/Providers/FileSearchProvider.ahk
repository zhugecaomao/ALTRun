;===============================================================================
; FileSearchProvider.ahk - 文件 / 文件夹搜索 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 两种用法 (和 Alfred 一样):
;   1. 默认结果 (InDefaultResults, 默认关闭): 直接输入名称, 匹配的文件和文件夹显示在
;      应用和命令下面 (最多 DefaultResultsLimit 条, 至少输入 MinQueryLength 个字)
;   2. 专门搜索文件:  空的搜索框里先按空格 (SpacePrefix, 见 SearchWindow) 再输入 report,
;                     或  'report  /  open report  /  find report
;   3. 只搜文件夹:    folder bk (FolderKeywords), 文件搜索模式里也可以写 "folder bk"
;   4. 按类型搜索:    doc 报告 / pic logo / cad 平面图 ... (TypeFilters: "关键字 = 扩展名 扩展名 ..."),
;                     和 Listary 的文件类型筛选一样; 文件搜索模式里也可以写 "doc 报告"
;   5. 浏览文件夹:    输入路径 C:\  D:\Projects\rep  \\server\share\  ~\  %OneDrive%\ ,
;                     列出这个文件夹里的内容 (最后一段过滤), 文件夹在前; Tab 进入选中的文件夹, Backspace 回到上一级
;                     (和 PowerToys Run / Flow Launcher 的文件夹浏览一样)。隐藏和系统文件不列出
; 结果按名称匹配程度排序 (完全相同 > 名称开头 > 单词开头 > 包含), 同分时文件夹在前,
; 再按修改时间。Everything 的语法 (folder:、ext:、path:...) 原样传给 Everything。
;
; 数据来源 (自动选择):
;   - Everything 在运行: 通过 IPC 直接查询 (Lib\Everything.ahk), 全盘, 不需要额外文件
;   - 否则: 内置索引 (Src\Core\FileIndex.ahk), 只包括 ScopeFolders 里的文件夹
;
; 设置 (ALTRun.json -> Features.FileSearch):
;   Keywords / FolderKeywords / TypeFilters / SpacePrefix / QuotePrefix / MaxResults / InDefaultResults / DefaultResultsLimit /
;   MinQueryLength / UseEverything / EverythingFilter / EverythingPath /
;   ScopeFolders / ScopeDepth / ScopeExclude / MaxEntries / RefreshMinutes
;===============================================================================

class FileSearchProvider {
    static Id := "FileSearch"

    static Init() {
        FileIndex.Start()
    }

    static Search(query) {
        options := AppSettings.Feature("FileSearch")
        term := ""
        if IsObject(browse := FileSearchProvider.BrowsePath(query.Raw))
            return FileSearchProvider.BrowseResults(browse)
        if (options["QuotePrefix"] && query.MatchPrefix("'", &term))
            return FileSearchProvider._KeywordResults(term, options, false, true)
        ; open / find / folder: 输入空格之后才只显示文件结果; 只输入 "folder" 时名字带 folder 的命令照常显示
        if query.MatchKeyword(options["Keywords"], &term)
            return FileSearchProvider._KeywordResults(term, options, false, query.HasRest)
        if query.MatchKeyword(options["FolderKeywords"], &term)
            return FileSearchProvider._KeywordResults(term, options, true, query.HasRest)
        if IsObject(filter := FileSearchProvider.MatchType(query, &term))
            return FileSearchProvider._KeywordResults(term, options, filter, query.HasRest)
        if (options["InDefaultResults"] && StrLen(query.Text) >= options["MinQueryLength"] && !Calc.Looks(query.Text))
            return FileSearchProvider._DefaultResults(query.Text, options)
        return []
    }

    ; 搜索窗口的文件搜索模式 (空格开头) 调用: 只搜文件, 空文字不返回结果; "folder bk" 只搜文件夹, "doc 报告" 只搜文档
    static SearchFiles(term) {
        term := Trim(term)
        if (term = "")
            return []
        if IsObject(browse := FileSearchProvider.BrowsePath(term))
            return FileSearchProvider.BrowseResults(browse)
        options := AppSettings.Feature("FileSearch")
        query := SearchQuery(term), rest := ""
        if (query.HasRest && query.MatchKeyword(options["FolderKeywords"], &rest) && rest != "")
            return FileSearchProvider._KeywordResults(rest, options, true, true)
        if (query.HasRest && IsObject(filter := FileSearchProvider.MatchType(query, &rest)) && rest != "")
            return FileSearchProvider._KeywordResults(rest, options, filter, true)
        return FileSearchProvider._KeywordResults(term, options, false, true)
    }

    ; 输入的是文件夹路径时返回 {Dir: "C:\Projects\", Filter: "rep"}, 否则 ""; 文件夹不存在时 {Dir, Filter, Missing: true}
    ;   C:\ / C:\Projects\rep / \\server\share\ / ~\ (用户文件夹) / %OneDrive%\ (环境变量)
    ;   中文输入法打出的 "D：、" (全角冒号、顿号) 也当作 "D:\"
    static BrowsePath(text) {
        text := LTrim(text)
        if RegExMatch(text, "^[A-Za-z][:：][\\、]")
            text := SubStr(text, 1, 1) ":\" StrReplace(SubStr(text, 4), "、", "\")
        if (text = "~")
            text := EnvGet("UserProfile") "\"
        else if RegExMatch(text, "^~\\")
            text := EnvGet("UserProfile") SubStr(text, 2)
        else if RegExMatch(text, "^%[^%\\]+%\\") {
            size := DllCall("ExpandEnvironmentStringsW", "WStr", text, "Ptr", 0, "UInt", 0, "UInt")
            buf := Buffer(size * 2)
            DllCall("ExpandEnvironmentStringsW", "WStr", text, "Ptr", buf, "UInt", size, "UInt")
            text := StrGet(buf)
        }
        if !RegExMatch(text, "^([A-Za-z]:\\|\\\\[^\\]+\\[^\\]+\\)")
            return ""
        pos := InStr(text, "\", , -1)
        dir := SubStr(text, 1, pos), filter := SubStr(text, pos + 1)
        if !DirExist(dir)
            return {Dir: dir, Filter: filter, Missing: true}
        return {Dir: dir, Filter: filter}
    }

    ; 文件夹里的内容 (文件夹在前), 按最后一段过滤; 文件夹的 Tab 补全成 "路径\", 进入下一级
    static BrowseResults(browse) {
        if browse.HasOwnProp("Missing")                                     ; 例如没有 D 盘: 提示, 而不是什么都不显示
            return [ResultItem(I18n.T("Files.FolderNotFound"), browse.Dir, {Icon: "folder:", Valid: false, Score: 150, Exclusive: true})]
        entries := FileSearchProvider.ListFolder(browse.Dir)
        needle := StrLower(browse.Filter), scores := Map()
        for index, entry in entries {
            score := (needle = "") ? 100 : FuzzyMatcher.Score(needle, entry.Name)
            if (score > 0)
                scores[index] := score + (entry.IsFolder ? 0.5 : 0) - index * 0.00001
        }
        results := []
        for index in FuzzyMatcher.TopIndexes(scores, ProviderRegistry.MaxResults) {
            entry := entries[index]
            item := ResultItem(entry.Name, browse.Dir, {
                Kind: entry.IsFolder ? "folder" : "file", Arg: entry.Path, Icon: entry.IsFolder ? "folder:" : entry.Path,
                Uid: "file:" StrLower(entry.Path), Score: 150 + scores[index] / 1000, Exclusive: true,
                AutoComplete: entry.IsFolder ? entry.Path "\" : ""
            })
            results.Push(item)
        }
        if !results.Length
            results.Push(ResultItem(I18n.T("Files.EmptyFolder"), browse.Dir, {Icon: "folder:", Valid: false, Score: 150, Exclusive: true}))
        return results
    }

    ; 文件夹的内容 [{Name, Path, IsFolder}], 文件夹在前; 3 秒内再次浏览同一个文件夹时直接用 (边打字边过滤)
    static ListFolder(dir) {
        static cache := Map()
        key := StrLower(dir)
        if (cache.Has(key) && A_TickCount - cache[key].Time < 3000)
            return cache[key].Entries
        folders := [], files := []
        try {
            Loop Files, dir "*", "FD" {
                if InStr(A_LoopFileAttrib, "H") || InStr(A_LoopFileAttrib, "S")
                    continue
                entry := {Name: A_LoopFileName, Path: A_LoopFileFullPath, IsFolder: InStr(A_LoopFileAttrib, "D") ? true : false}
                (entry.IsFolder ? folders : files).Push(entry)
                if (folders.Length + files.Length >= 5000)                  ; 特别大的文件夹只列前 5000 项
                    break
            }
        }
        for entry in files
            folders.Push(entry)
        if (cache.Count > 20)
            cache := Map()
        cache[key] := {Time: A_TickCount, Entries: folders}
        return folders
    }

    ; 文件类型筛选: 第一个词是 TypeFilters 里的关键字时返回 {Keyword, Extensions: Map, Everything: "ext:a;b"}
    static MatchType(query, &term) {
        term := ""
        for filter in FileSearchProvider.TypeFilters()
            if query.MatchKeyword([filter.Keyword], &term)
                return filter
        return ""
    }

    ; TypeFilters ("doc = doc docx pdf ...", 每行一个) 解析后的列表; 设置没变时不重新解析
    static TypeFilters() {
        static source := "", parsed := []
        options := AppSettings.Feature("FileSearch")
        lines := options.Has("TypeFilters") ? options["TypeFilters"] : []
        key := ""
        for line in lines
            key .= line "`n"
        if (key == source)
            return parsed
        source := key, parsed := []
        for line in lines
            if IsObject(filter := FileSearchProvider.ParseTypeFilter(line))
                parsed.Push(filter)
        return parsed
    }

    ; "doc = doc docx .pdf, xls" -> {Keyword: "doc", Extensions: Map, Everything: "ext:doc;docx;pdf;xls"}; 写错时返回 ""
    static ParseTypeFilter(line) {
        if !RegExMatch(line, "^\s*([^=\s]+)\s*=\s*(.+)$", &m)
            return ""
        extensions := Map(), list := ""
        for ext in StrSplit(RegExReplace(m[2], "[\s,;]+", " "), " ") {
            ext := StrLower(LTrim(Trim(ext), "*."))
            if (ext = "" || extensions.Has(ext))
                continue
            extensions[ext] := true, list .= (list = "" ? "" : ";") ext
        }
        return extensions.Count ? {Keyword: StrLower(m[1]), Extensions: extensions, Everything: "ext:" list} : ""
    }

    ; 'xxx / open xxx / folder xxx / doc xxx: 只显示文件 (或文件夹、某类文件) 搜索结果
    ; scope: false = 文件和文件夹, true = 只要文件夹, 或者 TypeFilters 里的一项
    ; exclusive: 已经输入了关键字后面的空格, 只显示这些结果
    static _KeywordResults(term, options, scope, exclusive) {
        if (term = "") {
            if IsObject(scope)
                hint := ResultItem(I18n.T("Files.TypeKeyword", StrReplace(SubStr(scope.Everything, 5), ";", " ")), scope.Keyword " ...", {Icon: "folder:", Valid: false})
            else {
                keywords := options[scope ? "FolderKeywords" : "Keywords"]
                keyword := keywords.Length ? keywords[1] : ""
                hint := scope ? ResultItem(I18n.T("Folders.Keyword"), keyword " ...", {Icon: "folder:", Valid: false})
                              : ResultItem(I18n.T("Files.Keyword"), "'..." (keyword != "" ? " / " keyword " ..." : ""), {Icon: "folder:", Valid: false})
            }
            hint.Score := exclusive ? 150 : 5                               ; 还没输入空格时排在后面, 不挡住其它结果
            hint.Exclusive := exclusive
            return [hint]
        }
        results := []
        for found in FileSearchProvider.Query(term, options["MaxResults"], false, scope) {
            item := FileSearchProvider._ToItem(found, 150 - A_Index * 0.01)
            item.Exclusive := exclusive
            results.Push(item)
        }
        fallback := FileSearchProvider.FallbackItem(term, scope)
        fallback.Score := results.Length ? 0 : 150
        fallback.Exclusive := exclusive
        results.Push(fallback)
        return results
    }

    ; 默认结果: 分数压低 (最高约 40), 排在应用 / 命令后面
    static _DefaultResults(text, options) {
        results := []
        for found in FileSearchProvider.Query(text, options["DefaultResultsLimit"], true)
            results.Push(FileSearchProvider._ToItem(found, 5 + (found.Score + (found.IsFolder ? 1 : 0)) * 0.35))
        return results
    }

    ; 返回 [{Path, IsFolder, Score}], 按匹配程度排序
    ; preferPrefix: 默认结果只要名称开头 / 单词开头匹配的, 避免一大堆只是 "包含" 的文件
    ; scope:        true = 只要文件夹 (folder bk); TypeFilters 里的一项 = 只要这些扩展名的文件 (doc 报告)
    ; Everything 按修改时间返回; 只取前 limit 条的话, 名称最匹配但不是最近修改的文件夹 (例如
    ; "bk" 匹配到大量文件时) 就排不进来, 所以多取一些 (EverythingFetch 条), 再按名称匹配程度挑
    static Query(term, limit, preferPrefix, scope := false) {
        options := AppSettings.Feature("FileSearch")
        needle := StrLower(Trim(term))
        filter := IsObject(scope) ? scope : "", foldersOnly := !IsObject(scope) && scope
        if (options["UseEverything"] && Everything.IsRunning()) {
            found := []
            search := (preferPrefix ? "startwith:" : "") FileSearchProvider._ScopePrefix(scope) FileSearchProvider._EverythingTerm(term) " " options["EverythingFilter"]
            fetch := preferPrefix ? limit * 4 : Max(limit, FileSearchProvider.EverythingFetch)
            scoreNeedle := FileSearchProvider.ScoreNeedle(needle)
            seen := Map()
            for item in Everything.Query(Trim(search), fetch, Everything.SORT_DATE_MODIFIED_DESC) {
                SplitPath(item.Path, &name)
                score := FileIndex.ScoreName(scoreNeedle, StrLower(name))
                found.Push({Path: item.Path, IsFolder: item.IsFolder, Score: score ? score : 50})
                seen[StrLower(item.Path)] := true
            }
            ; 文件和文件夹一起搜时, 同名的文件很多 (例如 ppie 文件夹里有几百个 ppie_xxx.dwg) 的话, 按修改时间取的前 fetch 条
            ; 可能全是文件, 名称完全相同的文件夹反而排不进来: 另外单独取一次文件夹
            if (!IsObject(scope) && !scope && !RegExMatch(term, "i)(^|\s)(folder|file|files|ext):")) {   ; 已经写了 folder: / ext: 等就不另外取
                folderSearch := (preferPrefix ? "startwith:" : "") "folder:" FileSearchProvider._EverythingTerm(term) " " options["EverythingFilter"]
                for item in Everything.Query(Trim(folderSearch), 50, Everything.SORT_DATE_MODIFIED_DESC) {
                    if seen.Has(StrLower(item.Path))
                        continue
                    SplitPath(item.Path, &name)
                    score := FileIndex.ScoreName(scoreNeedle, StrLower(name))
                    found.Push({Path: item.Path, IsFolder: true, Score: score ? score : 50})
                }
            }
            return FileSearchProvider._Best(found, limit)
        }
        found := FileIndex.Search(needle, preferPrefix ? limit * 4 : limit, foldersOnly, IsObject(filter) ? filter.Extensions : "")
        if preferPrefix {
            kept := []
            for item in found
                if (item.Score >= 70)                                  ; 名称开头 (90) 或单词开头 (80, 长名称略低于 80)
                    kept.Push(item)
            found := kept
        }
        return FileSearchProvider._Best(found, limit)
    }

    static EverythingFetch := 300

    ; Everything 搜索的前缀: "folder:" 或 "ext:doc;docx " (后面还要接搜索的文字)
    static _ScopePrefix(scope) {
        return IsObject(scope) ? scope.Everything " " : scope ? "folder:" : ""
    }

    ; 给结果打分用的文字: 去掉 Everything 的语法前缀 ("folder:bk" -> "bk", "ext:pdf report" -> "pdf report")
    static ScoreNeedle(needle) {
        return Trim(RegExReplace(needle, "(^|\s)[a-z]+:", "$1"))
    }

    ; 多个词的输入交给 Everything 时保持原样 (Everything 本身就按空格分词); 引号去掉避免语法错误
    static _EverythingTerm(term) {
        return StrReplace(Trim(term), '"')
    }

    static _Best(found, limit) {
        scores := Map()
        for index, item in found
            scores[index] := item.Score + (item.IsFolder ? 1 : 0)            ; 同分时文件夹优先
        best := []
        for index in FuzzyMatcher.TopIndexes(scores, limit)
            best.Push(found[index])
        return best
    }

    static _ToItem(found, score) {
        SplitPath(found.Path, &name, &parentDir)
        return ResultItem(name, parentDir, {
            Kind: found.IsFolder ? "folder" : "file", Arg: found.Path, Icon: found.IsFolder ? "folder:" : found.Path,
            Uid: "file:" StrLower(found.Path), Score: score
        })
    }

    ; 没有结果时: 在 Everything 里搜索 (装了 Everything) 或用 Windows 搜索
    static FallbackItem(term, scope := false) {
        exe := FileSearchProvider._EverythingExe()
        if (exe != "") {
            search := FileSearchProvider._ScopePrefix(scope) StrReplace(term, '"')
            return ResultItem(I18n.T("Files.OpenEverything", search), exe, {
                Icon: exe, OnRun: (*) => Run('"' exe '" -s "' search '"')
            })
        }
        search := (IsObject(scope) ? "" : scope ? "kind:folder " : "") term
        return ResultItem(I18n.T("Files.WindowsSearch", search), "search-ms:", {
            Icon: "res:imageres.dll,-8", OnRun: (*) => Run("search-ms:query=" Url.Encode(search))
        })
    }

    static _EverythingExe() {
        static cached := ""
        if (cached != "")
            return (cached = "-") ? "" : cached
        configured := Path.Resolve(AppSettings.Feature("FileSearch")["EverythingPath"])
        candidates := []
        if (configured != "")
            candidates.Push(DirExist(configured) ? configured "\Everything.exe" : configured)
        for folder in [A_ScriptDir, A_ProgramFiles "\Everything", EnvGet("ProgramFiles(x86)") "\Everything", EnvGet("LocalAppData") "\Everything"]
            candidates.Push(folder "\Everything.exe")
        for candidate in candidates {
            if FileExist(candidate) {
                cached := candidate
                return candidate
            }
        }
        cached := "-"
        return ""
    }
}
