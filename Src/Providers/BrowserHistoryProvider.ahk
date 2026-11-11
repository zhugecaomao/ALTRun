;===============================================================================
; BrowserHistoryProvider.ahk - browser history (AutoHotkey v2)
;-------------------------------------------------------------------------------
; "h word" searches the history of Chromium browsers (Chrome, Edge, Brave, Vivaldi; every profile): pages whose
; title or URL contains every word (Chinese titles also by their pinyin initials), most recently visited first.
; "h " alone lists the most recent pages. Only with the keyword: history is private, so it never mixes into normal results.
; Enter opens the page in the default browser, Alt+Enter copies the URL.
;
; Each profile's History is an SQLite database the browser keeps open, read through Windows' own winsqlite3.dll
; (Lib\Sqlite.ahk) without locking it (see ReadFile). Read only when the keyword is typed, never at start-up or in the
; background; read again when searched more than ReloadMs after the last read. Nothing is written or sent anywhere.
;
; Settings (ALTRun.json -> Features.BrowserHistory): Enabled / Keyword
;===============================================================================

class BrowserHistoryProvider {
    static Id := "BrowserHistory"
    static Items := []                                                      ; [{Title, Url, Time, Text, Cjk, Initials}], most recent first
    static MaxItems := 10000                                                ; the most recent pages read from each History file
    static ReloadMs := 60000
    static _loaded := 0                                                     ; A_TickCount of the last read (0 = never)
    static _active := false                                                 ; the last search was a history search

    static Init() {                                                         ; nothing at start-up: read only when searched
    }

    static Search(query) {
        options := AppSettings.Feature("BrowserHistory")
        if !query.MatchKeyword([options["Keyword"]], &rest, &keyword) {
            BrowserHistoryProvider._active := false
            return []
        }
        if !query.HasRest                                                   ; "h" alone: the keyword's hint, Enter completes "h "
            return [ResultItem.KeywordHint(I18n.T("BrowserHistory.Keyword"), keyword, "url:")]
        if !Sqlite.Available()
            return [ResultItem(I18n.T("BrowserHistory.Unavailable"), "", {Icon: "url:", Valid: false, Exclusive: true})]
        ; read when history search starts (not on every key press while typing in it) and the last read is over ReloadMs old
        if (!BrowserHistoryProvider._loaded || !BrowserHistoryProvider._active && A_TickCount - BrowserHistoryProvider._loaded > BrowserHistoryProvider.ReloadMs)
            BrowserHistoryProvider.Load()
        BrowserHistoryProvider._active := true
        results := []
        for page in BrowserHistoryProvider.Matches(rest, ProviderRegistry.MaxResults)
            results.Push(ResultItem(page.Title, BrowserHistoryProvider.TimeText(page.Time) "  ·  " page.Url, {
                Kind: "url", Arg: page.Url, Icon: "url:", Uid: "history:" page.Url,
                Score: 50, Exclusive: true                                  ; all the same score: they stay most recent first
            }))
        if !results.Length                                                  ; say so, instead of the normal results or "Search Google for h ..."
            results.Push(ResultItem(I18n.T(BrowserHistoryProvider.Items.Length ? "BrowserHistory.NoMatch" : "BrowserHistory.NoBrowser"), "", {
                Icon: "url:", Valid: false, Exclusive: true}))
        return results
    }

    ; The pages whose title or URL contains every word of term (most recent first, at most limit); term "" = the most recent
    static Matches(term, limit) {
        words := []
        for word in StrSplit(StrLower(Trim(term)), " ")
            if (word != "")
                words.Push(word)
        found := []
        for page in BrowserHistoryProvider.Items {
            matched := true
            for word in words {
                if InStr(page.Text, word)
                    continue
                if (page.Cjk && InStr(BrowserHistoryProvider._Initials(page), word))   ; "zwbt" finds "中文标题"
                    continue
                matched := false
                break
            }
            if matched {
                found.Push(page)
                if (found.Length >= limit)
                    break
            }
        }
        return found
    }

    ; The History files of every profile of every browser (the same browsers as the bookmarks)
    static Files() {
        files := []
        for browser in BookmarkProvider.Browsers {
            root := EnvGet("LOCALAPPDATA") "\" browser[2]
            if !DirExist(root)
                continue
            Loop Files, root "\*", "D"
                if FileExist(A_LoopFileFullPath "\History")
                    files.Push(A_LoopFileFullPath "\History")
        }
        return files
    }

    static Load() {
        start := Logger.Ms()
        lines := "", pages := [], seen := Map()
        for historyPath in BrowserHistoryProvider.Files() {
            try {
                for row in BrowserHistoryProvider.ReadFile(historyPath) {
                    pageUrl := row[1]
                    if seen.Has(pageUrl) {                                      ; a page visited in several browsers / profiles: its last visit counts
                        page := pages[seen[pageUrl]]
                        if (row[3] > page.Time)
                            page.Time := row[3]
                        continue
                    }
                    title := (row[2] != "") ? row[2] : pageUrl
                    pages.Push({Title: title, Url: pageUrl, Time: row[3], Text: StrLower(title " " pageUrl), Cjk: RegExMatch(title, "[\x{2E80}-\x{9FFF}]") > 0, Initials: ""})
                    seen[pageUrl] := pages.Length
                }
            } catch as e {
                Logger.Error("BrowserHistoryProvider: " historyPath " - " e.Message)
            }
        }
        for index, page in pages                                            ; most recent first across every file
            lines .= Format("{:020}", page.Time) "`t" index "`n"
        items := []
        if (lines != "")
            for line in StrSplit(Sort(RTrim(lines, "`n"), "R"), "`n")
                items.Push(pages[Integer(StrSplit(line, "`t")[2])])
        BrowserHistoryProvider.Items := items, BrowserHistoryProvider._loaded := A_TickCount
        Logger.Time("BrowserHistoryProvider: " items.Length " pages", start)
    }

    ; One History file -> [[url, title, last visit time], ...], most recent first, http / https only.
    ; The browser keeps the file open and writes it at any time. It is read in place without any locking (immutable: the
    ; browser is never blocked, nothing is copied); if that fails (the browser was in the middle of a write), from a copy in
    ; %TEMP% with its journal, opened read-write so SQLite rolls back the unfinished write; the copy is deleted right away
    static Sql := "SELECT url, title, last_visit_time FROM urls"
        . " WHERE hidden = 0 AND last_visit_time > 0 AND (url LIKE 'http://%' OR url LIKE 'https://%')"
        . " ORDER BY last_visit_time DESC LIMIT "
    static ReadFile(historyPath) {
        sql := BrowserHistoryProvider.Sql BrowserHistoryProvider.MaxItems
        try return Sqlite.Query(historyPath, sql, 0, "immutable")
        catch as e
            Logger.Debug("BrowserHistoryProvider: in place failed, reading a copy - " e.Message)
        temp := A_Temp "\ALTRun-History-" DllCall("GetCurrentProcessId")
        DirCreate(temp)
        try {
            copy := temp "\History"
            FileCopy(historyPath, copy, true)
            if FileExist(historyPath "-journal")
                FileCopy(historyPath "-journal", copy "-journal", true)
            return Sqlite.Query(copy, sql, 0, "write")
        } finally {
            try DirDelete(temp, true)                                       ; no copy of the history is left behind
        }
    }

    ; Chromium's time (microseconds since 1601-01-01 UTC) -> "14:05" today, "11-09 14:05" this year, "2025-12-31" before
    static TimeText(chromeTime, now := "") {
        if (chromeTime <= 0)
            return ""
        static offset := DateDiff(A_Now, A_NowUTC, "Seconds")
        stamp := DateAdd("16010101", chromeTime // 1000000 + offset, "Seconds")
        now := (now != "") ? now : A_Now
        if (SubStr(stamp, 1, 8) = SubStr(now, 1, 8))
            return FormatTime(stamp, "HH:mm")
        if (SubStr(stamp, 1, 4) = SubStr(now, 1, 4))
            return FormatTime(stamp, "MM-dd HH:mm")
        return FormatTime(stamp, "yyyy-MM-dd")
    }

    static _Initials(page) {
        if (page.Initials = "")
            page.Initials := StrLower(StrReplace(Pinyin.Initials(page.Title), " "))   ; "中文 标题" -> "zwbt"
        return page.Initials
    }
}
