;===============================================================================
; Knowledge.ahk - learned ranking + search history (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like Alfred's Knowledge, it remembers "what was typed -> which result was chosen in the end":
;   - how often each result was chosen in total and when last (more used = higher)
;   - which result an input (and its prefixes) chose (after choosing Notepad for "no", it comes first next time
;     for "no" or "not")
;   - recent search text, recalled with ↑ in the empty search box
; The data is in Data\Knowledge.json; deleting it only makes the ranking start learning again.
;
; Usage:
;   Knowledge.Load()                    at start-up
;   Knowledge.Record(queryText, uid)    after running a result
;   Knowledge.Boost(queryText, uid)     the bonus when ranking search results
;   Knowledge.History                   recent search text (newest first)
;   Knowledge.Prune()                   before each save: forgets the results chosen longest ago past MaxPicks / MaxQueries
;===============================================================================

class Knowledge {
    static File        := AppSettings.DataDir "\Knowledge.json"
    static Picks       := Map()      ; uid -> Map("Count", n, "Last", "yyyyMMddHHmmss")
    static QueryPicks  := Map()      ; lower-case input -> Map(uid -> count)
    static History     := []
    static Recent      := []         ; recently opened items (newest first), see RecentProvider
    static MaxQueryLen := 20
    static MaxPicks    := 3000       ; past these the oldest are forgotten when saving (Prune), so the file and each save stay small
    static MaxQueries  := 5000
    static _saveTimer  := ""

    static Load() {
        Knowledge.Picks := Map(), Knowledge.QueryPicks := Map(), Knowledge.History := [], Knowledge.Recent := []
        if !FileExist(Knowledge.File)
            return
        try {
            data := JSON.Parse(FileRead(Knowledge.File, "UTF-8"))
            Knowledge.Picks := Knowledge.ValidPicks(data.Has("Picks") ? data["Picks"] : "")
            Knowledge.QueryPicks := Knowledge.ValidQueryPicks(data.Has("QueryPicks") ? data["QueryPicks"] : "")
            if (data.Has("History") && data["History"] is Array)
                for text in data["History"]
                    if (!IsObject(text) && text != "")
                        Knowledge.History.Push(text)
            if (data.Has("Recent") && data["Recent"] is Array)
                for entry in data["Recent"]
                    if (entry is Map && entry.Has("Uid"))
                        Knowledge.Recent.Push(entry)
        } catch as e {
            Logger.Error("Knowledge.Load: " e.Message)
        }
    }

    ; Only entries the ranking can use, from a file that may have been edited by hand or cut short: uid -> Map with a number Count
    ; (Boost would fail on every search otherwise); input -> Map(uid -> number)
    static ValidPicks(picks) {
        valid := Map()
        if (picks is Map)
            for uid, entry in picks
                if (entry is Map && entry.Has("Count") && IsNumber(entry["Count"]))
                    valid[uid] := entry
        return valid
    }
    static ValidQueryPicks(queryPicks) {
        valid := Map()
        if (queryPicks is Map)
            for queryText, picks in queryPicks {
                if !(picks is Map)
                    continue
                kept := Map()
                for uid, count in picks
                    if IsNumber(count)
                        kept[uid] := count
                if kept.Count
                    valid[queryText] := kept
            }
        return valid
    }

    static Save() {
        Knowledge.Prune()
        try {
            JSON.WriteFile(Knowledge.File, Map("Picks", Knowledge.Picks, "QueryPicks", Knowledge.QueryPicks, "History", Knowledge.History, "Recent", Knowledge.Recent))
        } catch as e {
            Logger.Error("Knowledge.Save: " e.Message)
        }
    }

    ; Past MaxPicks results, the ones chosen longest ago are forgotten, and what they were chosen for with them; past MaxQueries inputs,
    ; the ones whose results were chosen longest ago. Ties at the cut are kept, so a few more may stay
    static Prune() {
        if (Knowledge.Picks.Count > Knowledge.MaxPicks) {
            cutoff := Knowledge._Cutoff(Knowledge.Picks, Knowledge.MaxPicks, (entry) => Knowledge._Last(entry))
            gone := []
            for uid, entry in Knowledge.Picks
                if (Knowledge._Last(entry) < cutoff)
                    gone.Push(uid)
            for uid in gone
                Knowledge.Picks.Delete(uid)
            empty := []
            for queryText, picks in Knowledge.QueryPicks {
                if !(picks is Map) {
                    empty.Push(queryText)
                    continue
                }
                for uid in gone
                    if picks.Has(uid)
                        picks.Delete(uid)
                if !picks.Count
                    empty.Push(queryText)
            }
            for queryText in empty
                Knowledge.QueryPicks.Delete(queryText)
        }
        if (Knowledge.QueryPicks.Count > Knowledge.MaxQueries) {
            cutoff := Knowledge._Cutoff(Knowledge.QueryPicks, Knowledge.MaxQueries, (picks) => Knowledge._NewestPick(picks))
            gone := []
            for queryText, picks in Knowledge.QueryPicks
                if (Knowledge._NewestPick(picks) < cutoff)
                    gone.Push(queryText)
            for queryText in gone
                Knowledge.QueryPicks.Delete(queryText)
        }
    }

    ; The time (yyyyMMddHHmmss as a number, 0 = unknown) of the keep-th newest value in items; those older than it are dropped
    static _Cutoff(items, keep, timeOf) {
        times := ""
        for key, value in items
            times .= timeOf(value) "`n"
        return Number(StrSplit(Sort(RTrim(times, "`n"), "N R"), "`n")[keep])
    }
    static _Last(entry) => (entry is Map && entry.Has("Last") && IsNumber(entry["Last"])) ? Number(entry["Last"]) : 0
    static _NewestPick(picks) {
        newest := 0
        if (picks is Map)
            for uid in picks
                if (Knowledge.Picks.Has(uid) && (last := Knowledge._Last(Knowledge.Picks[uid])) > newest)
                    newest := last
        return newest
    }

    static Record(queryText, uid) {
        queryText := StrLower(Trim(queryText))
        Knowledge.AddHistory(queryText)
        if (uid = "")
            return Knowledge._SaveLater()

        entry := Knowledge.Picks.Has(uid) ? Knowledge.Picks[uid] : Map("Count", 0, "Last", "")
        entry["Count"] := entry["Count"] + 1
        entry["Last"]  := A_Now
        Knowledge.Picks[uid] := entry

        if (queryText != "" && StrLen(queryText) <= Knowledge.MaxQueryLen) {
            picks := Knowledge.QueryPicks.Has(queryText) ? Knowledge.QueryPicks[queryText] : Map()
            picks[uid] := (picks.Has(uid) ? picks[uid] : 0) + 1
            Knowledge.QueryPicks[queryText] := picks
        }
        Knowledge._SaveLater()
    }

    static AddHistory(queryText) {
        if (queryText = "")
            return
        for index, previous in Knowledge.History {
            if (previous = queryText) {
                Knowledge.History.RemoveAt(index)
                break
            }
        }
        Knowledge.History.InsertAt(1, queryText)
        limit := AppSettings.General["HistorySize"]
        while (Knowledge.History.Length > limit)
            Knowledge.History.Pop()
    }

    ; Ranking bonus: how often it is used (up to about +25) + used recently (+5) + chosen for this input and its prefixes (up to about +60)
    static Boost(queryText, uid) {
        if (uid = "" || !Knowledge.Picks.Has(uid))
            return 0
        entry := Knowledge.Picks[uid]
        boost := 8 * Ln(1 + entry["Count"])
        if (entry["Last"] != "" && DateDiff(A_Now, entry["Last"], "Days") <= 7)
            boost += 5

        queryText := StrLower(Trim(queryText))
        len := Min(StrLen(queryText), Knowledge.MaxQueryLen)
        Loop len {
            prefix := SubStr(queryText, 1, len - A_Index + 1)
            if (Knowledge.QueryPicks.Has(prefix) && Knowledge.QueryPicks[prefix].Has(uid)) {
                exact := (A_Index = 1) ? 1.5 : 1                            ; an identical input weighs more
                boost += Min(25 * exact * Ln(1 + Knowledge.QueryPicks[prefix][uid]), 60)
                break
            }
        }
        return boost
    }

    ; Writes once after several commands in a row; not while the search window is open (typing), so typing never stalls
    static _SaveLater() {
        if (Knowledge._saveTimer = "")
            Knowledge._saveTimer := Logger.Watch("save knowledge", () => Knowledge._SaveWhenIdle())
        SetTimer(Knowledge._saveTimer, -3000)
    }

    static _SaveWhenIdle() {
        if !BackgroundWork.CanRun()
            return SetTimer(Knowledge._saveTimer, -2000)
        Knowledge.Save()
    }
}
