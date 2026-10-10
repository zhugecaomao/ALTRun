;===============================================================================
; WebSearchProvider.ahk - 网页搜索 + 兜底结果 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; ALTRun.json -> Features.WebSearch.Engines, 每个引擎:
;   { "Id": "google", "Keyword": "g", "Title": "Google", "Url": "https://...?q={query}", "Icon": "" }
; 输入 "g 关键词" 用对应引擎搜索; 只输入关键字 "g" 时提示补全。
; Url 里写 {query|默认值} 时, 只输入关键字也能直接打开: 用默认值, 例如 "http://localhost:{query|8000}"。
; Icon: 结果的图标, 可以是 .ico / .png / .exe 等文件 (相对路径从 Data 文件夹算起), 或 IconCache 的写法
; (res:、ext: 等); 不写时用默认浏览器的图标。
;
; Fallbacks: 其它功能都没有结果时显示的兜底项, 例如 "Search Google for 'xxx'"。
; 列表里写引擎的 Id, 或 "files" (交给 FileSearchProvider 搜索文件)。
;===============================================================================

class WebSearchProvider {
    static Id := "WebSearch"

    static Init() {
    }

    static Engines() {
        return AppSettings.Feature("WebSearch")["Engines"]
    }

    static Search(query) {
        results := []
        for engine in WebSearchProvider.Engines() {
            if !(engine is Map) || !engine.Has("Keyword") || engine["Keyword"] = ""
                continue
            if query.MatchKeyword([engine["Keyword"]], &term) {
                item := WebSearchProvider.ItemFor(engine, term)
                item.Score := 150
                item.Exclusive := query.HasRest
                results.Push(item)
            } else if (!query.HasRest && StrLen(query.Text) >= 2 && FuzzyMatcher.Score(query.Text, engine["Title"]) >= 80) {
                item := WebSearchProvider.ItemFor(engine, "")                ; 输入引擎名称 -> 提示 "g " 补全
                item.Score := 20
                results.Push(item)
            }
        }
        return results
    }

    static ItemFor(engine, term) {
        icon := WebSearchProvider.IconFor(engine)
        autoComplete := engine["Keyword"] " " term
        if (term = "")
            term := WebSearchProvider.DefaultTerm(engine["Url"])
        if (term = "") {
            return ResultItem(I18n.T("Web.SearchEmpty", engine["Title"]), I18n.T("Search.TypeAfter", engine["Keyword"]), {
                Kind: "url", Icon: icon, Valid: false, AutoComplete: autoComplete, Source: engine
            })
        }
        target := WebSearchProvider.Fill(engine["Url"], term)
        return ResultItem(I18n.T("Web.SearchFor", engine["Title"], term), target, {
            Kind: "url", Arg: target, Icon: icon, AutoComplete: autoComplete, Source: engine
        })
    }

    ; 网址里的 {query} / {query|默认值} 换成搜索词 (编码); 搜索词是空的时用默认值
    static Fill(template, term) {
        if (term = "")
            term := WebSearchProvider.DefaultTerm(template)
        return RegExReplace(template, "\{query(?:\|[^}]*)?\}", Url.Encode(term))   ; 编码后只有字母、数字和 %-_.~, 没有 $
    }

    ; {query|默认值} 里的默认值, 没有时 ""
    static DefaultTerm(template) {
        return RegExMatch(template, "\{query\|([^}]*)\}", &m) ? m[1] : ""
    }

    ; 引擎的 Icon: 图片文件用缩略图, .ico / .exe 等用文件自己的图标; 没写或文件不存在时用默认浏览器的图标
    static IconFor(engine) {
        spec := engine.Has("Icon") ? Trim(engine["Icon"], " `t`"") : ""
        if (spec = "")
            return "url:"
        if RegExMatch(spec, "i)^(res|ext|url|folder|thumb|shell):")
            return spec
        if !RegExMatch(spec, "i)^([a-z]:\\|\\\\|%|A_)")                    ; 相对路径: 从 Data 文件夹算起
            spec := AppSettings.DataDir "\" spec
        iconFile := Path.Resolve(spec)
        if !FileExist(iconFile)
            return "url:"
        return RegExMatch(iconFile, "i)\.(png|jpe?g|bmp|gif)$") ? "thumb:" iconFile : iconFile
    }

    static EditorFields() {
        return [ItemEditor.Field("Title", "Prefs.Col.Title", "text", true)
              , ItemEditor.Field("Keyword", "Prefs.Col.Keyword", "text", true)
              , ItemEditor.Field("Url", "Prefs.Col.Url", "text", true, "", I18n.T("Web.Field.Url"))
              , ItemEditor.Field("Id", "Prefs.Col.Id", "text", true)
              , ItemEditor.Field("Icon", "Prefs.Col.Icon", "file", false, "", I18n.T("Web.Field.Icon"))]
    }

    static NewEngine() {
        return Map("Id", "", "Keyword", "", "Title", "", "Url", "https://", "Icon", "")
    }

    ; 搜索结果里 F3 / 右键 "编辑": 修改对应的搜索引擎
    static EditItem(item) {
        engine := item.Source
        if !IsObject(engine)
            return false
        edited := ItemEditor.Edit(ItemEditor.Owner(), I18n.T("Prefs.Page.WebSearch"), WebSearchProvider.EditorFields(), engine)
        if !IsObject(edited)
            return false
        for key, value in edited
            engine[key] := value
        return AppSettings.Save()
    }

    static DeleteItem(item) {
        engines := WebSearchProvider.Engines()
        for index, engine in engines {
            if (ObjPtr(engine) = ObjPtr(item.Source)) {
                engines.RemoveAt(index)
                return AppSettings.Save()
            }
        }
        return false
    }

    static Fallbacks(query) {
        results := []
        for id in AppSettings.Feature("WebSearch")["Fallbacks"] {
            if (id = "files") {
                if ProviderRegistry.IsEnabled(FileSearchProvider)
                    results.Push(FileSearchProvider.FallbackItem(query.Text))
                continue
            }
            for engine in WebSearchProvider.Engines() {
                if (engine is Map && engine.Has("Id") && engine["Id"] = id) {
                    results.Push(WebSearchProvider.ItemFor(engine, query.Text))
                    break
                }
            }
        }
        return results
    }
}
