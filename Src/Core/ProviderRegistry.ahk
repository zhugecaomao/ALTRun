;===============================================================================
; ProviderRegistry.ahk - 管理所有搜索功能 (Provider), 汇总排序搜索结果 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 每个功能 (Src\Providers\*.ahk) 是一个只有静态方法的类, 约定:
;   static Id            功能名, 同时也是 ALTRun.json 里 Features 下的设置键名
;   static Init()        启动时调用一次 (建立索引等)
;   static Search(query) 参数是 SearchQuery, 返回 [ResultItem...]
;   static EditItem(item) / DeleteItem(item)   可选: 搜索结果里 F3 编辑 / Ctrl+Del 删除
;   static DeletePrompt(item)                  可选: 删除前的确认文字 (默认 "确定删除 ... 吗?")
;                        (结果的 Source 指向设置里的那一条, 见 ActionCatalog.CanEdit)
; 没有启用 (Features.<Id>.Enabled = 0) 的功能不会被初始化, 也不参与搜索。
;
; Search() 把所有功能的结果合在一起, 加上 Knowledge 的学习加分后按分数排序,
; 同一个 Uid 只保留分数最高的一条; 一条都没有时显示兜底项 (WebSearch 的 Fallbacks)。
; 某个功能进入关键字模式时 (结果带 Exclusive, 例如 "clip "), 只显示这些结果。
;
; 用法:
;   ProviderRegistry.Register(ApplicationProvider)    启动时按顺序注册
;   ProviderRegistry.InitAll()
;   ProviderRegistry.WarmUp()                         启动后空闲时: 预先算好搜索用的数据 (可选的 Warm())
;   ProviderRegistry.Search("note")                   -> [ResultItem...]
;===============================================================================

class ProviderRegistry {
    static Providers  := []
    static MaxResults := 50

    static Register(provider) {
        ProviderRegistry.Providers.Push(provider)
    }

    static ById(id) {
        for provider in ProviderRegistry.Providers
            if (provider.Id = id)
                return provider
        return ""
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

    static WarmUp() {
        for provider in ProviderRegistry.Providers {
            if !ProviderRegistry.IsEnabled(provider) || !HasMethod(provider, "Warm")
                continue
            try {
                provider.Warm()
            } catch as e {
                Logger.Error("ProviderRegistry: " provider.Id ".Warm failed - " e.Message)
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

        results := ProviderRegistry._KeepExclusive(results)
        results := ProviderRegistry._Dedupe(ProviderRegistry.SortByScore(results))
        if (results.Length > ProviderRegistry.MaxResults)
            results.Length := ProviderRegistry.MaxResults
        if (!results.Length && ProviderRegistry.IsEnabled(WebSearchProvider))
            results := WebSearchProvider.Fallbacks(query)
        return results
    }

    ; 搜索窗口的文件搜索模式: 只要文件搜索的结果 (已按匹配程度排好)
    static SearchFiles(text) {
        if !ProviderRegistry.IsEnabled(FileSearchProvider)
            return []
        results := FileSearchProvider.SearchFiles(text)
        for item in results
            item.Provider := FileSearchProvider.Id
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

    static _KeepExclusive(items) {
        exclusive := []
        for item in items
            if item.Exclusive
                exclusive.Push(item)
        return exclusive.Length ? exclusive : items
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
