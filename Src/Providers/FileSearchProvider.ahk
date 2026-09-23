;===============================================================================
; FileSearchProvider.ahk - 文件 / 文件夹搜索 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 两种用法 (和 Alfred 一样):
;   1. 默认结果 (InDefaultResults, 默认关闭): 直接输入名称, 匹配的文件和文件夹显示在
;      应用和命令下面 (最多 DefaultResultsLimit 条, 至少输入 MinQueryLength 个字)
;   2. 专门搜索文件:  空的搜索框里先按空格 (SpacePrefix, 见 SearchWindow) 再输入 report,
;                     或  'report  /  open report  /  find report
;
; 数据来源 (自动选择):
;   - Everything 在运行: 通过 IPC 直接查询 (Lib\Everything.ahk), 全盘, 不需要额外文件
;   - 否则: 内置索引 (Src\Core\FileIndex.ahk), 只包括 ScopeFolders 里的文件夹
;
; 设置 (ALTRun.json -> Features.FileSearch):
;   Keywords / SpacePrefix / QuotePrefix / MaxResults / InDefaultResults / DefaultResultsLimit /
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
        keywordMode := (options["QuotePrefix"] && query.MatchPrefix("'", &term))
        if !keywordMode
            keywordMode := query.MatchKeyword(options["Keywords"], &term)
        if keywordMode
            return FileSearchProvider._KeywordResults(term, options)
        if (options["InDefaultResults"] && StrLen(query.Text) >= options["MinQueryLength"] && !Calc.Looks(query.Text))
            return FileSearchProvider._DefaultResults(query.Text, options)
        return []
    }

    ; 搜索窗口的文件搜索模式 (空格开头) 调用: 只搜文件, 空文字不返回结果
    static SearchFiles(term) {
        if (Trim(term) = "")
            return []
        return FileSearchProvider._KeywordResults(Trim(term), AppSettings.Feature("FileSearch"))
    }

    ; 'xxx / open xxx: 只显示文件搜索结果
    static _KeywordResults(term, options) {
        if (term = "")
            return [ResultItem(I18n.T("Files.Keyword"), "'... / " options["Keywords"][1] " ...", {Icon: "folder:", Valid: false, Score: 150, Exclusive: true})]
        results := []
        for found in FileSearchProvider.Query(term, options["MaxResults"], false) {
            item := FileSearchProvider._ToItem(found, 150 - A_Index * 0.01)
            item.Exclusive := true
            results.Push(item)
        }
        fallback := FileSearchProvider.FallbackItem(term)
        fallback.Score := results.Length ? 0 : 150
        fallback.Exclusive := true
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
    static Query(term, limit, preferPrefix) {
        options := AppSettings.Feature("FileSearch")
        needle := StrLower(Trim(term))
        if (options["UseEverything"] && Everything.IsRunning()) {
            found := []
            search := (preferPrefix ? "startwith:" : "") FileSearchProvider._EverythingTerm(term) " " options["EverythingFilter"]
            for item in Everything.Query(Trim(search), preferPrefix ? limit * 4 : limit, Everything.SORT_DATE_MODIFIED_DESC) {
                SplitPath(item.Path, &name)
                score := FileIndex.ScoreName(needle, StrLower(name))
                found.Push({Path: item.Path, IsFolder: item.IsFolder, Score: score ? score : 50})
            }
            return FileSearchProvider._Best(found, limit)
        }
        found := FileIndex.Search(needle, preferPrefix ? limit * 4 : limit)
        if preferPrefix {
            kept := []
            for item in found
                if (item.Score >= 70)                                  ; 名称开头 (90) 或单词开头 (80, 长名称略低于 80)
                    kept.Push(item)
            found := kept
        }
        return FileSearchProvider._Best(found, limit)
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
    static FallbackItem(term) {
        exe := FileSearchProvider._EverythingExe()
        if (exe != "") {
            return ResultItem(I18n.T("Files.OpenEverything", term), exe, {
                Icon: exe, OnRun: (*) => Run('"' exe '" -s "' StrReplace(term, '"') '"')
            })
        }
        return ResultItem(I18n.T("Files.WindowsSearch", term), "search-ms:", {
            Icon: "res:imageres.dll,-8", OnRun: (*) => Run("search-ms:query=" Url.Encode(term))
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
