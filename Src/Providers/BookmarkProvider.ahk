;===============================================================================
; BookmarkProvider.ahk - 浏览器书签 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 Alfred 的 Web Bookmarks 一样: 直接输入书签名称就能找到 (InDefaultResults, 排在应用和命令后面),
; "bm 关键词" (Keyword) 只搜书签。Enter 用默认浏览器打开, Alt+Enter 复制网址。
; 读 Chromium 系浏览器 (Chrome、Edge、Brave、Vivaldi) 每个用户配置 (Default、Profile 1...) 的
; Bookmarks 文件 (JSON)。Firefox 的书签存在 SQLite 数据库里, 读不了。
; 启动时读一次; 搜索时最多每 30 秒看一下文件有没有变 (修改时间), 变了再重新读。
;
; 设置 (ALTRun.json -> Features.Bookmarks): Enabled / Keyword / InDefaultResults
;===============================================================================

class BookmarkProvider {
    static Id := "Bookmarks"
    static Items := []                                                      ; [{Title, Url, Key, Browser}]
    static _stamp := "", _checked := 0
    static Browsers := [["Chrome", "Google\Chrome\User Data"], ["Edge", "Microsoft\Edge\User Data"]
                      , ["Brave", "BraveSoftware\Brave-Browser\User Data"], ["Vivaldi", "Vivaldi\User Data"]]

    static Init() {
        BookmarkProvider.Load()
    }

    static Search(query) {
        options := AppSettings.Feature("Bookmarks")
        exclusive := false, term := query.Text
        if query.MatchKeyword([options["Keyword"]], &rest) {
            if !query.HasRest
                return []
            term := rest, exclusive := true
        } else if !options["InDefaultResults"] || StrLen(query.Text) < 2 {
            return []
        }
        BookmarkProvider._RefreshIfChanged()
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
                Score: scores[index] - (exclusive ? 0 : 12), Exclusive: exclusive     ; 默认结果里排在应用和命令后面
            }))
        }
        return results
    }

    ; 所有浏览器的所有配置里的 Bookmarks 文件
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
                    if !seen.Has(bookmark.Url) {                            ; 几个浏览器都有的书签只留一条
                        seen[bookmark.Url] := true
                        bookmark.Browser := bookmarkFile[1]
                        items.Push(bookmark)
                    }
                stamp .= fileStamp                                          ; 读成功了才记下: 浏览器正在写文件时读失败, 下次再读
            } catch as e {
                Logger.Error("BookmarkProvider: " bookmarkFile[2] " - " e.Message)
            }
        }
        BookmarkProvider.Items := items, BookmarkProvider._stamp := stamp, BookmarkProvider._checked := A_TickCount
        Logger.Debug("BookmarkProvider: " items.Length " bookmarks")
    }

    ; Bookmarks 文件 (JSON) -> [{Title, Url, Keys}], 只要 http / https 网址
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
        if (A_TickCount - BookmarkProvider._checked < 30000)
            return
        BookmarkProvider._checked := A_TickCount
        stamp := ""
        for bookmarkFile in BookmarkProvider.Files()
            try stamp .= bookmarkFile[2] FileGetTime(bookmarkFile[2], "M") "|"
        if (stamp != BookmarkProvider._stamp)
            BookmarkProvider.Load()
    }
}
