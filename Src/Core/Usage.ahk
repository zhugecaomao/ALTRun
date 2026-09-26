;===============================================================================
; Usage.ahk - 使用统计, 和 Alfred 的 Usage 一样 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 按天记录每个功能用了几次, 存在 Data\Usage.json:
;   {"Since": "2026-09-26", "Days": {"2026-09-26": {"Show": 12, "Applications": 5, ...}, ...}}
; 只记次数, 不记输入了什么、打开了什么。计数只改内存里的 Map, 几秒后合并写一次盘,
; 对搜索速度没有影响。保留最近 KeepDays 天。
;
; 用法:
;   Usage.Load() / Usage.Save()
;   Usage.Count("SnippetExpand")        用了某个功能 (键见 Features; "Show" = 呼出搜索窗口)
;   Usage.CountItem(item)               执行了一个搜索结果 (按它的 Provider 计数)
;   Usage.Summary(days)                 最近 days 天 (0 = 全部) 每个功能的次数: Map(功能 -> 次数)
;   Usage.Total(summary)                使用次数 = 各功能相加 (不含 "Show")
;   Usage.Daily(days)                   最近 days 天每天的使用次数 [[日期, 次数], ...], 旧的在前
;   Usage.Clear()
;===============================================================================

class Usage {
    static File     := A_ScriptDir "\Data\Usage.json"
    static KeepDays := 400
    ; 统计页上的顺序 (次数相同时); "Show" 单独显示, 不算在使用次数里
    static Features := ["Applications", "CustomCommands", "FileSearch", "WebSearch", "Calculator", "Clipboard",
                        "Snippets", "System", "Terminal", "SnippetExpand", "QuickSwitch", "AutoDate"]
    static Days  := Map()        ; "yyyy-MM-dd" -> Map(功能 -> 次数)
    static Since := ""           ; 开始统计的日期
    static _saveTimer := ""

    static Load() {
        Usage.Days := Map(), Usage.Since := ""
        if !FileExist(Usage.File)
            return
        try {
            data := JSON.Parse(FileRead(Usage.File, "UTF-8"))
            if (data.Has("Days") && data["Days"] is Map)
                Usage.Days := data["Days"]
            if data.Has("Since")
                Usage.Since := data["Since"]
        } catch as e {
            Logger.Error("Usage.Load: " e.Message)
        }
    }

    static Save() {
        try {
            Usage._Prune()
            DirCreate(AppSettings.DataDir)
            tmpFile := Usage.File ".tmp"
            try FileDelete(tmpFile)
            FileAppend(JSON.Stringify(Map("Since", Usage.Since, "Days", Usage.Days)), tmpFile, "UTF-8")
            FileMove(tmpFile, Usage.File, true)
        } catch as e {
            Logger.Error("Usage.Save: " e.Message)
        }
    }

    static Count(feature, today := "") {
        if (today = "")
            today := FormatTime(, "yyyy-MM-dd")
        if !Usage.Days.Has(today)
            Usage.Days[today] := Map()
        day := Usage.Days[today]
        day[feature] := (day.Has(feature) ? day[feature] : 0) + 1
        if (Usage.Since = "")
            Usage.Since := today
        Usage._SaveLater()
    }

    static CountItem(item) {
        if !IsObject(item)
            return
        feature := item.Provider
        if (feature = "")                                                   ; 没有结果时的兜底项: 网页搜索或文件搜索
            feature := (item.Kind = "url") ? "WebSearch" : "FileSearch"
        if (feature != "Help")                                              ; 速查表只是打开 Wiki, 不算
            Usage.Count(feature)
    }

    static Summary(days := 0, today := "") {
        cutoff := (days > 0) ? Usage._DaysAgo(days - 1, today) : ""
        result := Map()
        for date, day in Usage.Days {
            if ((cutoff != "" && StrCompare(date, cutoff) < 0) || !(day is Map))    ; 日期是 yyyy-MM-dd, 按文字比较
                continue
            for feature, n in day
                result[feature] := (result.Has(feature) ? result[feature] : 0) + n
        }
        return result
    }

    static Total(summary) {
        total := 0
        for feature, n in summary
            if (feature != "Show")
                total += n
        return total
    }

    static Daily(days, today := "") {
        result := []
        Loop days {
            date := Usage._DaysAgo(days - A_Index, today)
            day := Usage.Days.Has(date) ? Usage.Days[date] : Map()
            result.Push([date, (day is Map) ? Usage.Total(day) : 0])
        }
        return result
    }

    static Clear() {
        Usage.Days := Map(), Usage.Since := ""
        Usage.Save()
    }

    ; n 天前的日期 "yyyy-MM-dd" (today 留空 = 今天)
    static _DaysAgo(n, today := "") {
        base := (today = "") ? A_Now : StrReplace(today, "-") "000000"
        return FormatTime(DateAdd(base, -n, "Days"), "yyyy-MM-dd")
    }

    static _Prune() {
        cutoff := Usage._DaysAgo(Usage.KeepDays - 1)
        old := []
        for date in Usage.Days
            if (StrCompare(date, cutoff) < 0)
                old.Push(date)
        for date in old
            Usage.Days.Delete(date)
    }

    ; 连续计数时只写一次盘
    static _SaveLater() {
        if (Usage._saveTimer = "")
            Usage._saveTimer := () => Usage.Save()
        SetTimer(Usage._saveTimer, -5000)
    }
}
