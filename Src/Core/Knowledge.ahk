;===============================================================================
; Knowledge.ahk - 学习排序 + 搜索历史 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 Alfred 的 Knowledge 一样, 记住 "输入了什么 -> 最后选了哪一项":
;   - 每项结果被选中的总次数和最近一次时间 (越常用越靠前)
;   - 某个输入 (以及它的前缀) 选中了哪一项 (输入 "no" 选过记事本, 下次输入
;     "no" 或 "not" 时记事本排第一)
;   - 最近的搜索文字, 搜索框为空时按 ↑ 调出
; 数据存在 Data\Knowledge.json, 删掉只会让排序重新开始学习。
;
; 用法:
;   Knowledge.Load()                    启动时
;   Knowledge.Record(queryText, uid)    执行某一项后
;   Knowledge.Boost(queryText, uid)     搜索排序时的加分
;   Knowledge.History                   最近的搜索文字 (新的在前)
;===============================================================================

class Knowledge {
    static File        := A_ScriptDir "\Data\Knowledge.json"
    static Picks       := Map()      ; uid -> Map("Count", n, "Last", "yyyyMMddHHmmss")
    static QueryPicks  := Map()      ; 小写输入 -> Map(uid -> 次数)
    static History     := []
    static MaxQueryLen := 20
    static _saveTimer  := ""

    static Load() {
        Knowledge.Picks := Map(), Knowledge.QueryPicks := Map(), Knowledge.History := []
        if !FileExist(Knowledge.File)
            return
        try {
            data := JSON.Parse(FileRead(Knowledge.File, "UTF-8"))
            if (data.Has("Picks") && data["Picks"] is Map)
                Knowledge.Picks := data["Picks"]
            if (data.Has("QueryPicks") && data["QueryPicks"] is Map)
                Knowledge.QueryPicks := data["QueryPicks"]
            if (data.Has("History") && data["History"] is Array)
                Knowledge.History := data["History"]
        } catch as e {
            Logger.Error("Knowledge.Load: " e.Message)
        }
    }

    static Save() {
        try {
            DirCreate(AppSettings.DataDir)
            data := Map("Picks", Knowledge.Picks, "QueryPicks", Knowledge.QueryPicks, "History", Knowledge.History)
            tmpFile := Knowledge.File ".tmp"
            try FileDelete(tmpFile)
            FileAppend(JSON.Stringify(data), tmpFile, "UTF-8")
            FileMove(tmpFile, Knowledge.File, true)
        } catch as e {
            Logger.Error("Knowledge.Save: " e.Message)
        }
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

    ; 排序加分: 常用程度 (最多约 +25) + 最近用过 (+5) + 这个输入及其前缀选过它 (最多约 +60)
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
                exact := (A_Index = 1) ? 1.5 : 1                            ; 完全相同的输入权重更高
                boost += Min(25 * exact * Ln(1 + Knowledge.QueryPicks[prefix][uid]), 60)
                break
            }
        }
        return boost
    }

    ; 连续执行多个命令时只写一次盘
    static _SaveLater() {
        if (Knowledge._saveTimer = "")
            Knowledge._saveTimer := () => Knowledge.Save()
        SetTimer(Knowledge._saveTimer, -3000)
    }
}
