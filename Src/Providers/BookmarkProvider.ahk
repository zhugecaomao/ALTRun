;===============================================================================
; BookmarkProvider.ahk - browser bookmarks (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like Alfred's Web Bookmarks: typing a bookmark's name finds it (InDefaultResults, after applications and commands),
; "bm word" (Keyword) searches bookmarks only. Enter opens in the browser (ActionCatalog.OpenUrl), Alt+Enter copies the URL.
; Reads the Bookmarks file (JSON) of every profile (Default, Profile 1...) of Chromium browsers (Chrome, Edge, Brave, Vivaldi),
; and of the folders in BrowserFolders (portable or other Chromium browsers: their User Data folder, or one profile folder).
; Firefox keeps its bookmarks in an SQLite database, which can't be read.
; Read once at start-up; then every 30 s a background check whether the files changed (modified time) reads them again (not while typing; one read takes 20 ~ 30 ms).
;
; Settings (ALTRun.json -> Features.Bookmarks): Enabled / Keyword / InDefaultResults / BrowserFolders (also used by browser history)
;===============================================================================

class BookmarkProvider {
    static Id := "Bookmarks"
    static Items := []                                                      ; [{Title, Url, Key, Browser}]
    static _stamp := "", _checked := 0
    static Browsers := [["Chrome", "Google\Chrome\User Data"], ["Edge", "Microsoft\Edge\User Data"]
                      , ["Brave", "BraveSoftware\Brave-Browser\User Data"], ["Vivaldi", "Vivaldi\User Data"]]

    static Init() {
        BookmarkProvider.Load()
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

    ; The Bookmarks files of every profile of every browser: [[browser, path]]
    static Files() => BookmarkProvider.ProfileFiles("Bookmarks")

    ; The file named fileName ("Bookmarks", "History") in every profile of every browser: [[browser, path]], each file once.
    ; The four browsers' User Data folders under %LocalAppData%, then the folders in BrowserFolders (browser "Other")
    static ProfileFiles(fileName) {
        roots := []
        for browser in BookmarkProvider.Browsers
            roots.Push([browser[1], EnvGet("LOCALAPPDATA") "\" browser[2]])
        for folder in BookmarkProvider.ExtraFolders()
            roots.Push(["Other", folder])
        files := [], seen := Map()
        for root in roots
            for profileFile in BookmarkProvider.FilesIn(root[2], fileName)
                if !seen.Has(key := StrLower(profileFile)) {                ; a folder in BrowserFolders may be one of the four too
                    seen[key] := true
                    files.Push([root[1], profileFile])
                }
        return files
    }

    ; The full paths of the folders in Features.Bookmarks.BrowserFolders (variables expanded; relative to the program folder)
    static ExtraFolders() {
        folders := []
        options := AppSettings.Feature("Bookmarks")
        if (options.Has("BrowserFolders") && options["BrowserFolders"] is Array)
            for folder in options["BrowserFolders"]
                if (!IsObject(folder) && Trim(folder) != "")
                    folders.Push(Path.Full(folder))
        return folders
    }

    ; The number of browser profiles in folder (a profile folder has a Preferences file): for Preferences to show whether a folder is right
    static ProfileCount(folder) => BookmarkProvider.FilesIn(folder, "Preferences").Length

    ; fileName in folder: a User Data folder (one subfolder per profile: Default, Profile 1...) or a single profile folder
    static FilesIn(folder, fileName) {
        files := []
        if !DirExist(folder)
            return files
        if FileExist(folder "\" fileName)                                   ; a profile folder itself
            files.Push(folder "\" fileName)
        Loop Files, folder "\*", "D"
            if FileExist(A_LoopFileFullPath "\" fileName)
                files.Push(A_LoopFileFullPath "\" fileName)
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
