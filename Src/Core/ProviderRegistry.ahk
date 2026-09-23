;===============================================================================
; ProviderRegistry.ahk - 管理所有搜索功能 (Provider), 汇总排序搜索结果 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 每个功能 (Src\Providers\*.ahk) 是一个只有静态方法的类, 约定:
;   static Id            功能名, 同时也是 ALTRun.json 里 Features 下的设置键名
;   static Init()        启动时调用一次 (建立索引等)
;   static Search(query) 参数是 SearchQuery, 返回 [ResultItem...]
; 没有启用 (Features.<Id>.Enabled = 0) 的功能不会被初始化, 也不参与搜索。
;
; Search() 把所有功能的结果合在一起, 加上 Knowledge 的学习加分后按分数排序,
; 同一个 Uid 只保留分数最高的一条; 一条都没有时显示兜底项 (WebSearch 的 Fallbacks)。
;
; 用法:
;   ProviderRegistry.Register(ApplicationProvider)    启动时按顺序注册
;   ProviderRegistry.InitAll()
;   ProviderRegistry.Search("note")                   -> [ResultItem...]
;===============================================================================

class ProviderRegistry {
    static Providers  := []
    static MaxResults := 50

    static Register(provider) {
        ProviderRegistry.Providers.Push(provider)
    }

    static IsEnabled(provider) {
        try return AppSettings.Feature(provider.Id)["Enabled"] ? true : false
        return false
    }

    static InitAll() {
        for provider in ProviderRegistry.Providers {
            if !ProviderRegistry.IsEnabled(provider)
                continue
            try {
                provider.Init()
            } catch as e {
                Logger.Error("ProviderRegistry: " provider.Id ".Init failed - " e.Message)
            }
        }
    }

    static Search(rawText) {
        query := SearchQuery(rawText)
        if (query.Text = "")
            return []

        results := []
        for provider in ProviderRegistry.Providers {
            if !ProviderRegistry.IsEnabled(provider)
                continue
            try {
                items := provider.Search(query)
            } catch as e {
                Logger.Error("ProviderRegistry: " provider.Id ".Search failed - " e.Message)
                continue
            }
            for item in items {
                item.Provider := provider.Id
                if (item.Uid != "")
                    item.Score += Knowledge.Boost(query.Text, item.Uid)
                results.Push(item)
            }
        }

        results := ProviderRegistry._Dedupe(ProviderRegistry.SortByScore(results))
        if (results.Length > ProviderRegistry.MaxResults)
            results.Length := ProviderRegistry.MaxResults
        if (!results.Length && ProviderRegistry.IsEnabled(WebSearchProvider))
            results := WebSearchProvider.Fallbacks(query)
        return results
    }

    ; 分数从高到低; 分数相同时保持原来的顺序 (功能注册的顺序)
    static SortByScore(items) {
        if (items.Length < 2)
            return items
        lines := ""
        for index, item in items
            lines .= Format("{:.6f}", item.Score - index * 0.00001) "`t" index "`n"
        sorted := []
        for line in StrSplit(RTrim(Sort(lines, "N R"), "`n"), "`n")
            sorted.Push(items[Integer(StrSplit(line, "`t")[2])])
        return sorted
    }

    static _Dedupe(items) {
        seen := Map(), kept := []
        for item in items {
            if (item.Uid != "") {
                if seen.Has(item.Uid)
                    continue
                seen[item.Uid] := true
            }
            kept.Push(item)
        }
        return kept
    }
}
