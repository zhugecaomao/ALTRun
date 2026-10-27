;===============================================================================
; BookmarkProvider.ahk - browser bookmarks (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like Alfred's Web Bookmarks: typing a bookmark's name finds it (InDefaultResults, after applications and commands),
; "bm word" (Keyword) searches bookmarks only. Enter opens in the default browser, Alt+Enter copies the URL.
; Reads the Bookmarks file (JSON) of every profile (Default, Profile 1...) of Chromium browsers (Chrome, Edge, Brave, Vivaldi).
; Firefox keeps its bookmarks in an SQLite database, which can't be read.
; Read once at start-up; then every 30 s a background check whether the files changed (modified time) reads them again (not while typing; one read takes 20 ~ 30 ms).
;
; Settings (ALTRun.json -> Features.Bookmarks): Enabled / Keyword / InDefaultResults
;===============================================================================

class BookmarkProvider {
    static Id := "Bookmarks"
    static Items := []                                                      ; [{Title, Url, Key, Browser}]
    static _stamp := "", _checked := 0
    static Browsers := [["Chrome", "Google\Chrome\User Data"], ["Edge", "Microsoft\Edge\User Data"]
                      , ["Brave", "BraveSoftware\Brave-Browser\User Data"], ["Vivaldi", "Vivaldi\User Data"]]

    static Init() {
        BookmarkProvider.Load()
        ; WinTimer, not SetTimer: it stays set while ALTRun is idle, and a SetTimer timer would wake AutoHotkey every 10 ms (see Lib\WinTimer.ahk)
        WinTimer.Set(Logger.Watch("bookmarks refresh", () => BookmarkProvider._RefreshIfChanged()), 30000)        ; check the bookmark files in the background, not during searches while typing
    }

    static Search(query) {
        options := AppSettings.Feature("Bookmarks")
        exclusive := false, term := query.Text, keyword := ""
        if query.MatchKeyword([options["Keyword"]], &rest, &keyword) {
            if !query.HasRest                                               ; "bm" alone: the keyword's hint, Enter completes "bm "
                return [ResultItem.KeywordHint(I18n.T("Bookmarks.Keyword"), keyword, "url:")]
            term := rest, exclusive := true
        } else if !options["InDefaultResults"] || StrLen(query.Text) < 2 {
            return []
        }
        if !BookmarkProvider._checked                                       ; bookmarks were off at start-up and switched on later
            BookmarkProvider.Init()
        needle := StrLower(Trim(term))
        results := []
        if (needle = "")
            return results
        scores := Map()
        for index, bookmark in BookmarkProvider.Items
            if ((score := FuzzyMatcher.BestKey(needle, bookmark.Keys)) > 0)
                scores[index] := score
        for index in FuzzyMatcher.TopIndexes(scores, exclusive ? ProviderRegistry.MaxResults : 8) {
            bookmark := BookmarkProvider.Items[index]
            results.Push(ResultItem(bookmark.Title, bookmark.Url, {
                Kind: "url", Arg: bookmark.Url, Icon: "url:", Uid: "bookmark:" bookmark.Url,
                Score: scores[index] - (exclusive ? 0 : 12), Exclusive: exclusive     ; in the default results after applications and commands
            }))
        }
        return results
    }

    ; The Bookmarks files of every profile of every browser
    static Files() {
        files := []
        for browser in BookmarkProvider.Browsers {
            root := EnvGet("LOCALAPPDATA") "\" browser[2]
            if !DirExist(root)
                continue
            Loop Files, root "\*", "D"
                if FileExist(A_LoopFileFullPath "\Bookmarks")
                    files.Push([browser[1], A_LoopFileFullPath "\Bookmarks"])
        }
        return files
    }

    static Load() {
        items := [], seen := Map(), stamp := ""
        for bookmarkFile in BookmarkProvider.Files() {
            try {
                fileStamp := bookmarkFile[2] FileGetTime(bookmarkFile[2], "M") "|"
                for bookmark in BookmarkProvider.ParseFile(FileRead(bookmarkFile[2], "UTF-8"))
                    if !seen.Has(bookmark.Url) {                            ; a bookmark several browsers have is kept once
                        seen[bookmark.Url] := true
                        bookmark.Browser := bookmarkFile[1]
                        items.Push(bookmark)
                    }
                stamp .= fileStamp                                          ; noted only after a successful read: reading fails while the browser is writing the file, so read again next time
            } catch as e {
                Logger.Error("BookmarkProvider: " bookmarkFile[2] " - " e.Message)
            }
        }
        BookmarkProvider.Items := items, BookmarkProvider._stamp := stamp, BookmarkProvider._checked := A_TickCount
        Logger.Debug("BookmarkProvider: " items.Length " bookmarks")
    }

    ; Bookmarks file (JSON) -> [{Title, Url, Keys}], only http / https URLs
    static ParseFile(text) {
        data := JSON.Parse(text)
        list := []
        if (data is Map && data.Has("roots") && data["roots"] is Map)
            for name, root in data["roots"]
                BookmarkProvider._Walk(root, list)
        return list
    }

    static _Walk(node, list) {
        if !(node is Map)
            return
        if (node.Has("type") && node["type"] = "url" && node.Has("url") && RegExMatch(node["url"], "i)^https?://")) {
            title := (node.Has("name") && node["name"] != "") ? node["name"] : node["url"]
            host := RegExReplace(node["url"], "i)^https?://(www\.)?([^/:?#]+).*$", "$2")
            pinyinText := Pinyin.Initials(title)
            list.Push({Title: title, Url: node["url"], Keys: [FuzzyMatcher.Key(title), FuzzyMatcher.Key(host), (pinyinText != title) ? FuzzyMatcher.Key(pinyinText) : ""]})
            return
        }
        if (node.Has("children") && node["children"] is Array)
            for child in node["children"]
                BookmarkProvider._Walk(child, list)
    }

    static _RefreshIfChanged() {
        if (A_TickCount - BookmarkProvider._checked < 25000 || !ProviderRegistry.IsEnabled(BookmarkProvider))
            return
        BookmarkProvider._checked := A_TickCount
        stamp := ""
        for bookmarkFile in BookmarkProvider.Files()
            try stamp .= bookmarkFile[2] FileGetTime(bookmarkFile[2], "M") "|"
        if (stamp != BookmarkProvider._stamp)
            BookmarkProvider.Load()
    }
}
