;===============================================================================
; FileSearchProvider.ahk - File / folder search (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Ways to use it (as in Alfred):
;   1. Default results (InDefaultResults, off by default): type a name, and matching files and folders show
;      below apps and commands (at most DefaultResultsLimit, from MinQueryLength characters)
;   2. Files only:      press Space in the empty search box first (SpacePrefix, see SearchWindow), then type report
;                       (with report already typed, move the cursor to the start and press Space),
;                       or  'report  /  open report  /  find report  /  file report
;   3. Folders only:    folder budget (FolderKeywords); "folder budget" also works in file search mode
;   4. One file type:   doc report / pic logo / cad plan ... (TypeFilters: "keyword = ext ext ..."),
;                       like Listary's file type filters; "doc report" also works in file search mode
;   4b. Recent files:   recent report (RecentKeywords): files and folders opened recently in any program (Windows "Recent
;                       items"), newest first, without Everything or the index; "recent " alone lists the newest
;   5. Browse a folder: type a path such as C:\  D:\Projects\rep  \\server\share\  ~\  %OneDrive%\ to list
;                       that folder (the last part filters), folders first; Tab opens the selected folder, Backspace goes up
;                       (like folder browsing in PowerToys Run / Flow Launcher). Hidden and system files are not listed
; Results are sorted by how well the name matches (exact > start of the name > start of a word > contains), folders
; first on ties, then by modification time. Everything syntax (folder:, ext:, path:...) is passed to Everything as is.
;
; Sources (chosen automatically):
;   - Everything running: queried directly over IPC (Lib\Everything.ahk), the whole disk, no extra files needed
;   - Otherwise: the built-in index (Src\Core\FileIndex.ahk), covering only the ScopeFolders
;
; Settings (ALTRun.json -> Features.FileSearch):
;   Keywords / FolderKeywords / RecentKeywords / TypeFilters / SpacePrefix / QuotePrefix / MaxResults / InDefaultResults / DefaultResultsLimit /
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
        ; open / find / file / folder: only after a space do just file results show; typing only "folder" still shows commands with folder in their name
        if query.MatchKeyword(options["Keywords"], &term)
            return FileSearchProvider._KeywordResults(term, options, false, query.HasRest)
        if query.MatchKeyword(options["FolderKeywords"], &term)
            return FileSearchProvider._KeywordResults(term, options, true, query.HasRest)
        if query.MatchKeyword(options["RecentKeywords"], &term)
            return FileSearchProvider.RecentResults(term, options, query.HasRest)
        if IsObject(filter := FileSearchProvider.MatchType(query, &term))
            return FileSearchProvider._KeywordResults(term, options, filter, query.HasRest)
        if (options["InDefaultResults"] && StrLen(query.Text) >= options["MinQueryLength"] && !Calc.Looks(query.Text))
            return FileSearchProvider._DefaultResults(query.Text, options)
        return []
    }

    ; ProviderRegistry.SearchOnly: the search window's file search mode (starting with Space): files only, nothing for empty text; "folder budget" searches folders only, "doc report" documents only
    static SearchOnly(term) {
        term := Trim(term)
        if (term = "")
            return []
        if IsObject(browse := FileSearchProvider.BrowsePath(term))
            return FileSearchProvider.BrowseResults(browse)
        options := AppSettings.Feature("FileSearch")
        query := SearchQuery(term), rest := ""
        if (query.HasRest && query.MatchKeyword(options["FolderKeywords"], &rest) && rest != "")
            return FileSearchProvider._KeywordResults(rest, options, true, true)
        if (query.HasRest && query.MatchKeyword(options["RecentKeywords"], &rest))
            return FileSearchProvider.RecentResults(rest, options, true)
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

    ;---------------------------------------------------------------------------
    ; Recent files: Windows "Recent items", read by Lib\WindowsRecent.ahk
    ;---------------------------------------------------------------------------
    ; "recent report": the recent files and folders whose names match, newest first on ties; "recent " alone lists the
    ; newest; "recent" without the space is only a hint below the other results. Only the newest RecentScan shortcuts
    static RecentScan := 300
    static RecentResults(term, options, exclusive) {
        keywords := options["RecentKeywords"], keyword := keywords.Length ? keywords[1] : ""
        if !exclusive
            return [ResultItem(I18n.T("Files.RecentKeyword"), keyword " ...", {Icon: "folder:", Valid: false, Score: 5})]
        items := WindowsRecent.Items(FileSearchProvider.RecentScan)
        needle := StrLower(Trim(term)), scores := Map()
        for index, item in items {
            score := (needle = "") ? 100 : FuzzyMatcher.Score(needle, item.Name)
            if (score > 0)
                scores[index] := score - index * 0.001                      ; the newer one first on ties
        }
        results := []
        for index in FuzzyMatcher.TopIndexes(scores, options["MaxResults"]) {
            item := items[index]
            SplitPath(item.Path, , &dir)
            results.Push(ResultItem(item.Name, dir, {
                Kind: item.IsFolder ? "folder" : "file", Arg: item.Path, Icon: item.IsFolder ? "folder:" : item.Path,
                Uid: "file:" StrLower(item.Path), Score: 150 + scores[index] / 1000, Exclusive: true,
                AutoComplete: item.IsFolder ? RTrim(item.Path, "\") "\" : ""
            }))
        }
        if !results.Length
            results.Push(ResultItem(I18n.T("Files.RecentNone"), keyword " " term, {Icon: "folder:", Valid: false, Score: 150, Exclusive: true}))
        return results
    }

    ; 默认结果: 分数压低 (最高约 40), 排在应用 / 命令后面
    static _DefaultResults(text, options) {
        results := []
        for found in FileSearchProvider.Query(text, options["DefaultResultsLimit"], true)
            results.Push(FileSearchProvider._ToItem(found, 5 + (found.Score + (found.IsFolder ? 1 : 0)) * 0.35))
        return results
    }

    ; Returns [{Path, IsFolder, Score}], sorted by how well they match
    ; preferPrefix: default results keep only name-start / word-start matches, avoiding piles of files that merely "contain" the text
    ; scope:        true = folders only (folder budget); an entry of TypeFilters = only files with those extensions (doc report)
    ; Everything returns results by modification time; taking only the first limit, the folder with the best name match but not
    ; modified recently (e.g. when a short word matches lots of files) would not get in, so fetch more (EverythingFetch) and pick by name match
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
                found.Push({Path: item.Path, IsFolder: item.IsFolder, Score: FileSearchProvider.NameScore(scoreNeedle, name)})
                seen[StrLower(item.Path)] := true
            }
            if (Everything.LastProblem != "")
                Logger.Warn("FileSearch: Everything " Everything.LastProblem)
            ; When files and folders are searched together and many files share the name (e.g. a "plans" folder holding hundreds of
            ; plans_xxx.dwg), the first fetch results by modification time may all be files and the folder of that exact name is left out: fetch folders separately
            if (!IsObject(scope) && !scope && !RegExMatch(term, "i)(^|\s)(folder|file|files|ext):")) {   ; not when folder: / ext: etc. is already given
                folderSearch := (preferPrefix ? "startwith:" : "") "folder:" FileSearchProvider._EverythingTerm(term) " " options["EverythingFilter"]
                for item in Everything.Query(Trim(folderSearch), 50, Everything.SORT_DATE_MODIFIED_DESC) {
                    if seen.Has(StrLower(item.Path))
                        continue
                    SplitPath(item.Path, &name)
                    found.Push({Path: item.Path, IsFolder: true, Score: FileSearchProvider.NameScore(scoreNeedle, name)})
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

    ; Everything 返回的结果按文件名重新打分 (排序用)。Everything 打开了拼音 (1.5 的 pinyin 选项、1.4 的 IbEverythingExt) 时,
    ; "cs" 找到的 "测试.bat" 文件名里没有 "cs": 再按拼音首字母打一次分, 比同一档的直接匹配低 5 分。
    ; 名称都对不上 (路径匹配、全拼、多个词等) 时 50 分。多个词 ("pt 2024") 时 ScoreName 本来就对不上, 不算拼音 (白算)
    static NameScore(needle, name) {
        if (needle = "")                                                    ; 只输入了 Everything 的语法 (例如 "d:"、"folder:"): 按 Everything 的顺序
            return 50
        if (score := FileIndex.ScoreName(needle, StrLower(name)))
            return score
        if (FileIndex.UsePinyin(needle) && (initials := FileIndex.InitialsOf(name)) != "" && (score := FileIndex.PinyinScore(needle, initials)))
            return score
        return 50
    }

    ; Everything 搜索的前缀: "folder:" 或 "ext:doc;docx " (后面还要接搜索的文字)
    static _ScopePrefix(scope) {
        return IsObject(scope) ? scope.Everything " " : scope ? "folder:" : ""
    }

    ; Text used to score results: Everything syntax prefixes removed ("folder:budget" -> "budget", "ext:pdf report" -> "pdf report")
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
