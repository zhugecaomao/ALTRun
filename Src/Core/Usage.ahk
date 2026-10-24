;===============================================================================
; Usage.ahk - usage statistics, like Alfred's Usage (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Counts how often each feature is used per day, stored in Data\Usage.json:
;   {"Since": "2026-09-26", "Days": {"2026-09-26": {"Show": 12, "Applications": 5, ...}, ...}}
; Only counts, never what was typed or opened. Counting only changes the Map in memory; it is written once a few seconds later,
; so search speed isn't affected. The last KeepDays days are kept.
;
; Usage:
;   Usage.Load() / Usage.Save()
;   Usage.Count("SnippetExpand")        a feature was used (keys as in Features; "Show" = the search window was shown)
;   Usage.CountItem(item)               a search result was run (counted for its provider)
;   Usage.Summary(days)                 counts per feature for the last days days (0 = all): Map(feature -> count)
;   Usage.Total(summary)                uses = the features added up (without "Show")
;   Usage.Daily(days)                   uses per day for the last days days [[date, count], ...], oldest first
;   Usage.Clear()
;===============================================================================

class Usage {
    static File     := AppSettings.DataDir "\Usage.json"
    static KeepDays := 400
    ; order on the statistics page (for equal counts); "Show" is shown separately and isn't counted as a use
    static Features := ["Applications", "CustomCommands", "FileSearch", "WebSearch", "Bookmarks", "Calculator", "Clipboard",
                        "Snippets", "System", "Windows", "Scripts", "Recent", "Terminal", "SnippetExpand", "Selection", "QuickSwitch", "AutoDate"]
    static Days  := Map()        ; "yyyy-MM-dd" -> Map(feature -> count)
    static Since := ""           ; the date counting started
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
            JSON.WriteFile(Usage.File, Map("Since", Usage.Since, "Days", Usage.Days))
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
        if (feature = "")                                                   ; fallbacks when there are no results: web search or file search
            feature := (item.Kind = "url") ? "WebSearch" : "FileSearch"
        if (feature != "Help")                                              ; the cheat sheet only opens the Wiki: not counted
            Usage.Count(feature)
    }

    static Summary(days := 0, today := "") {
        cutoff := (days > 0) ? Usage._DaysAgo(days - 1, today) : ""
        result := Map()
        for date, day in Usage.Days {
            if ((cutoff != "" && StrCompare(date, cutoff) < 0) || !(day is Map))    ; dates are yyyy-MM-dd and compare as text
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

    ; the date n days ago as "yyyy-MM-dd" (today empty = today)
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

    ; Writes once while counting in a row; not while the search window is open (typing), so typing never stalls
    static _SaveLater() {
        if (Usage._saveTimer = "")
            Usage._saveTimer := Logger.Watch("save usage", () => Usage._SaveWhenIdle())
        SetTimer(Usage._saveTimer, -5000)
    }

    static _SaveWhenIdle() {
        if !BackgroundWork.CanRun()
            return SetTimer(Usage._saveTimer, -3000)
        Usage.Save()
    }
}
