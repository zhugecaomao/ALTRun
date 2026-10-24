;===============================================================================
; ProviderRegistry.ahk - the search features (providers): registers them, searches them all and ranks the results
;   (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Each feature (Src\Providers\*.ahk) is a class with static methods only:
;   static Id            the feature's name, also its settings key under Features in ALTRun.json
;   static Init()        called once at start-up (builds indexes etc.)
;   static Search(query) takes a SearchQuery, returns [ResultItem...]
;   Optional:
;   static EditItem(item) / DeleteItem(item)   F3 edit / Ctrl+Del delete in the results (the result's Source points
;                                              to its entry in the settings, see ResultActions.CanEdit)
;   static DeletePrompt(item)                  the confirmation text before deleting (default "Delete ...?")
;   static EmptyResults()                      results shown in the empty search box (e.g. the update notice)
;   static Fallbacks(query)                    results when nothing else matched (web search engines)
;   static SearchOnly(text)                    the provider's own search mode (the file search mode)
;   static Warm() / IconItems()                see WarmUp / IconSpecs
; Disabled features (Features.<Id>.Enabled = 0) are neither initialized nor searched.
;
; Search() merges every feature's results, adds the Knowledge boost and sorts by score, keeping the best result per
; Uid; with no results it shows the fallbacks. When a feature enters its keyword mode (its results are Exclusive,
; e.g. "clip "), only those results are shown.
;
; Usage:
;   ProviderRegistry.Register(ApplicationProvider)    at start-up, in order
;   ProviderRegistry.InitAll()
;   ProviderRegistry.WarmUp()                         after start-up, when idle: prepare the search data (Warm())
;   IconCache.Prefetch(ProviderRegistry.IconSpecs())  after start-up, when idle: read the common items' icons
;   ProviderRegistry.Search("note")                   -> [ResultItem...]
;===============================================================================

class ProviderRegistry {
    static Providers  := []
    static MaxResults := 50
    static SlowSearchMs := 30                                               ; 超过这么多毫秒的搜索写进调试日志

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
                start := Logger.Ms()
                provider.Init()
                Logger.Time("startup: " provider.Id ".Init", start)
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
                start := Logger.Ms()
                provider.Warm()
                Logger.Time("warm-up: " provider.Id, start)
            } catch as e {
                Logger.Error("ProviderRegistry: " provider.Id ".Warm failed - " e.Message)
            }
        }
    }

    ; The icons worth reading in the background after start-up, most useful first: pinned / recent items, then the
    ; providers' items (IconItems: [{Icon, Uid}...], e.g. applications and custom commands), picked ones by pick count
    static IconSpecs() {
        specs := [], others := []
        try {
            for item in ProviderRegistry.EmptyResults()
                specs.Push(item.Icon)
        }
        for provider in ProviderRegistry.Providers {
            if !ProviderRegistry.IsEnabled(provider) || !HasMethod(provider, "IconItems")
                continue
            try {
                for item in provider.IconItems()
                    others.Push(item)
            } catch as e {
                Logger.Error("ProviderRegistry: " provider.Id ".IconItems failed - " e.Message)
            }
        }
        for item in ProviderRegistry.ByPicks(others)
            specs.Push(item.Icon)
        return specs
    }

    ; 按 Knowledge 里选中的次数从多到少排 (没选过的保持原来的顺序, 排在后面)
    static ByPicks(items) {
        picked := [], rest := []
        for item in items {
            count := Knowledge.Picks.Has(item.Uid) && Knowledge.Picks[item.Uid] is Map && Knowledge.Picks[item.Uid].Has("Count")
                ? Knowledge.Picks[item.Uid]["Count"] : 0
            if (IsNumber(count) && count > 0)
                picked.Push({Item: item, Count: count, Order: picked.Length})
            else
                rest.Push(item)
        }
        sorted := []
        for entry in ProviderRegistry._SortPicks(picked)
            sorted.Push(entry.Item)
        for item in rest
            sorted.Push(item)
        return sorted
    }

    static _SortPicks(entries) {                                            ; 次数多的在前, 一样多时保持原来的顺序 (插入排序, 选过的通常只有几十个)
        Loop entries.Length - 1 {
            i := A_Index + 1, entry := entries[i], j := i - 1
            while (j >= 1 && entries[j].Count < entry.Count)
                entries[j + 1] := entries[j], j -= 1
            entries[j + 1] := entry
        }
        return entries
    }

    ; skip: 可选, 这一次不搜的功能 (例如恢复上次的搜索时先不搜文件, 窗口显示后再完整搜一次)
    static Search(rawText, skip := "") {
        query := SearchQuery(rawText)
        if (query.Text = "")
            return ProviderRegistry.EmptyResults()

        results := [], started := Logger.Ms(), slow := ""
        for provider in ProviderRegistry.Providers {
            if (!ProviderRegistry.IsEnabled(provider) || provider = skip)
                continue
            try {
                start := Logger.Ms()
                items := provider.Search(query)
                if ((elapsed := Logger.Ms() - start) >= 10)
                    slow .= " " provider.Id "=" Round(elapsed)
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
        if !results.Length
            results := ProviderRegistry._Fallbacks(query)
        if (Logger.Enabled && (total := Logger.Ms() - started) >= ProviderRegistry.SlowSearchMs)   ; 慢的搜索记进日志 (不记输入的文字)
            Logger.Debug("Perf: search (" StrLen(rawText) " chars) " Round(total) " ms," slow)
        return results
    }

    ; 空搜索框: 平时没有结果, 只有 Provider 的 EmptyResults() (例如后台发现了新版本)
    static EmptyResults() {
        results := []
        for provider in ProviderRegistry.Providers {
            if !ProviderRegistry.IsEnabled(provider) || !HasMethod(provider, "EmptyResults")
                continue
            try {
                start := Logger.Ms()
                Logger.Trace("empty results: " provider.Id)
                for item in provider.EmptyResults() {
                    item.Provider := provider.Id
                    results.Push(item)
                }
                if (Logger.Enabled && Logger.Ms() - start >= 30)              ; 慢的记进日志 (置顶 / 最近使用的图标、路径)
                    Logger.Time("empty results: " provider.Id, start)
            } catch as e {
                Logger.Error("ProviderRegistry: " provider.Id ".EmptyResults failed - " e.Message)
            }
        }
        return results
    }

    ; Nothing matched: the providers' Fallbacks(query) (e.g. "Search Google for ...")
    static _Fallbacks(query) {
        results := []
        for provider in ProviderRegistry.Providers {
            if !ProviderRegistry.IsEnabled(provider) || !HasMethod(provider, "Fallbacks")
                continue
            try {
                for item in provider.Fallbacks(query)
                    results.Push(item)
            } catch as e {
                Logger.Error("ProviderRegistry: " provider.Id ".Fallbacks failed - " e.Message)
            }
        }
        return results
    }

    ; One provider's own search mode (the search window's file search mode): only its results, already sorted
    static SearchOnly(provider, text) {
        if !ProviderRegistry.IsEnabled(provider)
            return []
        results := provider.SearchOnly(text)
        for item in results
            item.Provider := provider.Id
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
