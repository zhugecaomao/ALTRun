;===============================================================================
; WebSearchProvider.ahk - 网页搜索 + 兜底结果 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; ALTRun.json -> Features.WebSearch.Engines, 每个引擎:
;   { "Id": "google", "Keyword": "g", "Title": "Google", "Url": "https://...?q={query}" }
; 输入 "g 关键词" 用对应引擎搜索; 只输入关键字 "g" 时提示补全。
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
        if (term = "") {
            return ResultItem(I18n.T("Web.SearchEmpty", engine["Title"]), engine["Keyword"] " ...", {
                Kind: "url", Icon: "url:", Valid: false, AutoComplete: engine["Keyword"] " "
            })
        }
        target := StrReplace(engine["Url"], "{query}", Url.Encode(term))
        return ResultItem(I18n.T("Web.SearchFor", engine["Title"], term), target, {
            Kind: "url", Arg: target, Icon: "url:", AutoComplete: engine["Keyword"] " " term
        })
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
