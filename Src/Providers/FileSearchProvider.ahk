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
        term := "", typed := ""
        if IsObject(browse := FileSearchProvider.BrowsePath(query.Raw))
            return FileSearchProvider.BrowseResults(browse)
        if (options["QuotePrefix"] && query.MatchPrefix("'", &term))
            return FileSearchProvider._KeywordResults(term, options, false, true)
        ; open / find / file / folder: only after a space do just file results show; typing only "folder" still shows commands with folder in their name
        if query.MatchKeyword(options["Keywords"], &term, &typed)
            return FileSearchProvider._KeywordResults(term, options, false, query.HasRest, typed)
        if query.MatchKeyword(options["FolderKeywords"], &term, &typed)
            return FileSearchProvider._KeywordResults(term, options, true, query.HasRest, typed)
        if query.MatchKeyword(options["RecentKeywords"], &term, &typed)
            return FileSearchProvider.RecentResults(term, options, query.HasRest, typed)
        if IsObject(filter := FileSearchProvider.MatchType(query, &term))
            return FileSearchProvider._KeywordResults(term, options, filter, query.HasRest, filter.Keyword)
        results := FileSearchProvider._PrefixHints(query, options)
        if (options["InDefaultResults"] && StrLen(query.Text) >= options["MinQueryLength"] && !Calc.Looks(query.Text))
            results.Push(FileSearchProvider._DefaultResults(query.Text, options)*)
        return results
    }

    ; "fol" / "rec" / "vid": a keyword not typed in full yet gets its hint at the end of the list (Enter / Tab complete "folder " ...)
    static _PrefixHints(query, options) {
        hints := [], keyword := ""
        if query.StartsKeyword(options["Keywords"], &keyword)
            hints.Push(FileSearchProvider._KeywordResults("", options, false, false, keyword)*)
        if query.StartsKeyword(options["FolderKeywords"], &keyword)
            hints.Push(FileSearchProvider._KeywordResults("", options, true, false, keyword)*)
        if query.StartsKeyword(options["RecentKeywords"], &keyword)
            hints.Push(FileSearchProvider.RecentResults("", options, false, keyword)*)
        for filter in FileSearchProvider.TypeFilters()
            if query.StartsKeyword([filter.Keyword], &keyword) {
                hints.Push(FileSearchProvider._KeywordResults("", options, filter, false, keyword)*)
                break
            }
        for hint in hints
            hint.Score := ResultItem.KeywordPrefixHintScore
        return hints
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

    ; When the input is a folder path returns {Dir: "C:\Projects\", Filter: "rep"}, otherwise ""; {Dir, Filter, Missing: true} when the folder doesn't exist
    ;   C:\ / C:\Projects\rep / \\server\share\ / ~\ (user folder) / %OneDrive%\ (environment variable)
    ;   "D：、" typed with a Chinese IME (full-width colon, ideographic comma) counts as "D:\" too
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

    ; The folder's contents (folders first), filtered by the last part; Tab on a folder completes to "path\" to go one level down
    static BrowseResults(browse) {
        if browse.HasOwnProp("Missing")                                     ; e.g. no drive D: show a hint rather than nothing
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

    ; The folder's contents [{Name, Path, IsFolder}], folders first; reused when the same folder is browsed again within 3 s (filtering while typing)
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
                if (folders.Length + files.Length >= 5000)                  ; very large folders list only the first 5000 entries
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

    ; File type filter: when the first word is a keyword in TypeFilters returns {Keyword, Extensions: Map, Everything: "ext:a;b"}
    static MatchType(query, &term) {
        term := ""
        for filter in FileSearchProvider.TypeFilters()
            if query.MatchKeyword([filter.Keyword], &term)
                return filter
        return ""
    }

    ; TypeFilters ("doc = doc docx pdf ...", one per line) parsed into a list; not parsed again while the setting is unchanged
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

    ; "doc = doc docx .pdf, xls" -> {Keyword: "doc", Extensions: Map, Everything: "ext:doc;docx;pdf;xls"}; "" if it is malformed
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

    ; 'xxx / open xxx / folder xxx / doc xxx: show only file (or folder, or one kind of file) results
    ; scope: false = files and folders, true = folders only, or an entry of TypeFilters
    ; exclusive: the space after the keyword is typed: show only these results
    ; typed: the keyword typed; the hint for the keyword alone completes to it and a space (Enter / Tab / click)
    static _KeywordResults(term, options, scope, exclusive, typed := "") {
        if (term = "") {
            if IsObject(scope)
                hint := ResultItem(I18n.T("Files.TypeKeyword", StrReplace(SubStr(scope.Everything, 5), ";", " ")), scope.Keyword " ...", {Icon: "folder:", Valid: false})
            else {
                keywords := options[scope ? "FolderKeywords" : "Keywords"]
                keyword := keywords.Length ? keywords[1] : ""
                hint := scope ? ResultItem(I18n.T("Folders.Keyword"), keyword " ...", {Icon: "folder:", Valid: false})
                              : ResultItem(I18n.T("Files.Keyword"), "'..." (keyword != "" ? " / " keyword " ..." : ""), {Icon: "folder:", Valid: false})
            }
            hint.Score := exclusive ? 150 : ResultItem.KeywordHintScore     ; before the space is typed, names starting with the word still come first
            if (!exclusive && typed != "")
                hint.AutoComplete := typed " "
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
    ; newest; "recent" without the space is only a hint (ResultItem.KeywordHintScore). Only the newest RecentScan shortcuts
    static RecentScan := 300
    static RecentResults(term, options, exclusive, typed := "") {
        keywords := options["RecentKeywords"], keyword := keywords.Length ? keywords[1] : ""
        if !exclusive
            return [ResultItem(I18n.T("Files.RecentKeyword"), keyword " ...", {Icon: "folder:", Valid: false, Score: ResultItem.KeywordHintScore, AutoComplete: (typed != "") ? typed " " : ""})]
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

    ; default results: scores lowered (at most about 40), after applications / commands
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
                if (item.Score >= 70)                                  ; start of the name (90) or of a word (80, a little below 80 for long names)
                    kept.Push(item)
            found := kept
        }
        return FileSearchProvider._Best(found, limit)
    }

    static EverythingFetch := 300

    ; Results from Everything are scored again by file name (for sorting). With pinyin in Everything (the pinyin option of 1.5, IbEverythingExt for 1.4),
    ; "测试.bat" found by "cs" has no "cs" in its name: score it again by pinyin initials, 5 points below a direct match of the same tier.
    ; 50 when the name doesn't match at all (path matches, full pinyin, several words etc.). With several words ("pt 2024") ScoreName can't match anyway, so pinyin isn't tried (wasted work)
    static NameScore(needle, name) {
        if (needle = "")                                                    ; only Everything syntax typed (e.g. "d:", "folder:"): keep Everything's order
            return 50
        if (score := FileIndex.ScoreName(needle, StrLower(name)))
            return score
        if (FileIndex.UsePinyin(needle) && (initials := FileIndex.InitialsOf(name)) != "" && (score := FileIndex.PinyinScore(needle, initials)))
            return score
        return 50
    }

    ; Prefix for an Everything search: "folder:" or "ext:doc;docx " (the search text follows)
    static _ScopePrefix(scope) {
        return IsObject(scope) ? scope.Everything " " : scope ? "folder:" : ""
    }

    ; Text used to score results: Everything syntax prefixes removed ("folder:budget" -> "budget", "ext:pdf report" -> "pdf report")
    static ScoreNeedle(needle) {
        return Trim(RegExReplace(needle, "(^|\s)[a-z]+:", "$1"))
    }

    ; input with several words goes to Everything as it is (Everything splits at spaces itself); quotes are removed to avoid syntax errors
    static _EverythingTerm(term) {
        return StrReplace(Trim(term), '"')
    }

    static _Best(found, limit) {
        scores := Map()
        for index, item in found
            scores[index] := item.Score + (item.IsFolder ? 1 : 0)            ; folders first on ties
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

    ; No results: search in Everything (if installed) or with Windows Search
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
