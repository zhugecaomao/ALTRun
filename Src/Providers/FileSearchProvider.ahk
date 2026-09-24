;===============================================================================
; FileSearchProvider.ahk - 文件 / 文件夹搜索 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 两种用法 (和 Alfred 一样):
;   1. 默认结果 (InDefaultResults, 默认关闭): 直接输入名称, 匹配的文件和文件夹显示在
;      应用和命令下面 (最多 DefaultResultsLimit 条, 至少输入 MinQueryLength 个字)
;   2. 专门搜索文件:  空的搜索框里先按空格 (SpacePrefix, 见 SearchWindow) 再输入 report,
;                     或  'report  /  open report  /  find report
;   3. 只搜文件夹:    folder bk (FolderKeywords), 文件搜索模式里也可以写 "folder bk"
; 结果按名称匹配程度排序 (完全相同 > 名称开头 > 单词开头 > 包含), 同分时文件夹在前,
; 再按修改时间。Everything 的语法 (folder:、ext:、path:...) 原样传给 Everything。
;
; 数据来源 (自动选择):
;   - Everything 在运行: 通过 IPC 直接查询 (Lib\Everything.ahk), 全盘, 不需要额外文件
;   - 否则: 内置索引 (Src\Core\FileIndex.ahk), 只包括 ScopeFolders 里的文件夹
;
; 设置 (ALTRun.json -> Features.FileSearch):
;   Keywords / FolderKeywords / SpacePrefix / QuotePrefix / MaxResults / InDefaultResults / DefaultResultsLimit /
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
        if (options["QuotePrefix"] && query.MatchPrefix("'", &term))
            return FileSearchProvider._KeywordResults(term, options, false, true)
        ; open / find / folder: 输入空格之后才只显示文件结果; 只输入 "folder" 时名字带 folder 的命令照常显示
        if query.MatchKeyword(options["Keywords"], &term)
            return FileSearchProvider._KeywordResults(term, options, false, query.HasRest)
        if query.MatchKeyword(options["FolderKeywords"], &term)
            return FileSearchProvider._KeywordResults(term, options, true, query.HasRest)
        if (options["InDefaultResults"] && StrLen(query.Text) >= options["MinQueryLength"] && !Calc.Looks(query.Text))
            return FileSearchProvider._DefaultResults(query.Text, options)
        return []
    }

    ; 搜索窗口的文件搜索模式 (空格开头) 调用: 只搜文件, 空文字不返回结果; "folder bk" 只搜文件夹
    static SearchFiles(term) {
        term := Trim(term)
        if (term = "")
            return []
        options := AppSettings.Feature("FileSearch")
        query := SearchQuery(term), rest := ""
        if (query.HasRest && query.MatchKeyword(options["FolderKeywords"], &rest) && rest != "")
            return FileSearchProvider._KeywordResults(rest, options, true, true)
        return FileSearchProvider._KeywordResults(term, options, false, true)
    }

    ; 'xxx / open xxx / folder xxx: 只显示文件 (或文件夹) 搜索结果
    ; exclusive: 已经输入了关键字后面的空格, 只显示这些结果
    static _KeywordResults(term, options, foldersOnly, exclusive) {
        if (term = "") {
            keywords := options[foldersOnly ? "FolderKeywords" : "Keywords"]
            keyword := keywords.Length ? keywords[1] : ""
            hint := foldersOnly ? ResultItem(I18n.T("Folders.Keyword"), keyword " ...", {Icon: "folder:", Valid: false})
                                : ResultItem(I18n.T("Files.Keyword"), "'..." (keyword != "" ? " / " keyword " ..." : ""), {Icon: "folder:", Valid: false})
            hint.Score := exclusive ? 150 : 5                               ; 还没输入空格时排在后面, 不挡住其它结果
            hint.Exclusive := exclusive
            return [hint]
        }
        results := []
        for found in FileSearchProvider.Query(term, options["MaxResults"], false, foldersOnly) {
            item := FileSearchProvider._ToItem(found, 150 - A_Index * 0.01)
            item.Exclusive := exclusive
            results.Push(item)
        }
        fallback := FileSearchProvider.FallbackItem(term, foldersOnly)
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
    ; foldersOnly:  只要文件夹 (folder bk)
    ; Everything 按修改时间返回; 只取前 limit 条的话, 名称最匹配但不是最近修改的文件夹 (例如
    ; "bk" 匹配到大量文件时) 就排不进来, 所以多取一些 (EverythingFetch 条), 再按名称匹配程度挑
    static Query(term, limit, preferPrefix, foldersOnly := false) {
        options := AppSettings.Feature("FileSearch")
        needle := StrLower(Trim(term))
        if (options["UseEverything"] && Everything.IsRunning()) {
            found := []
            search := (preferPrefix ? "startwith:" : "") (foldersOnly ? "folder:" : "") FileSearchProvider._EverythingTerm(term) " " options["EverythingFilter"]
            fetch := preferPrefix ? limit * 4 : Max(limit, FileSearchProvider.EverythingFetch)
            scoreNeedle := FileSearchProvider.ScoreNeedle(needle)
            for item in Everything.Query(Trim(search), fetch, Everything.SORT_DATE_MODIFIED_DESC) {
                SplitPath(item.Path, &name)
                score := FileIndex.ScoreName(scoreNeedle, StrLower(name))
                found.Push({Path: item.Path, IsFolder: item.IsFolder, Score: score ? score : 50})
            }
            return FileSearchProvider._Best(found, limit)
        }
        found := FileIndex.Search(needle, preferPrefix ? limit * 4 : limit, foldersOnly)
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
    static FallbackItem(term, foldersOnly := false) {
        exe := FileSearchProvider._EverythingExe()
        if (exe != "") {
            search := (foldersOnly ? "folder:" : "") StrReplace(term, '"')
            return ResultItem(I18n.T("Files.OpenEverything", search), exe, {
                Icon: exe, OnRun: (*) => Run('"' exe '" -s "' search '"')
            })
        }
        search := (foldersOnly ? "kind:folder " : "") term
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
