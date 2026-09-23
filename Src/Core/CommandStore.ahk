;===============================================================================
; CommandStore.ahk - 命令内存缓存 + 排序/用量/历史 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 Lib/AppData.ahk(负责 ALTRun.json 的读写)是姊妹关系:
;   AppData       - 磁盘上的 ALTRun.json 长什么样
;   CommandStore  - 运行时从 g_CMDDATA 派生出来的、给搜索/排序用的内存结构
;                   (g_COMMANDS/g_CMDINDEX/g_FALLBACK), 以及排序权重(Rank)、
;                   使用统计(Usage/RunCount)、执行历史(History)的更新逻辑
;
; 用法 (ALTRun.ahk 里):
;   CommandStore.LoadCommands()             ; 启动时 / 增删改命令后重建内存缓存
;   CommandStore.LoadHistory()              ; 启动时按配置整理历史记录
;   CommandStore.UpdateRank(cmd, showRank)  ; 命中命令后调整其权重(SmartRank)
;   CommandStore.UpdateUsage()              ; 记一次"今日激活次数"
;   CommandStore.UpdateRunCount()           ; 记一次"命令执行总次数"
;   CommandStore.UpdateHistory(cmd)         ; 追加一条执行历史
;
; RankUp()/RankDown() 两个全局函数必须留在 ALTRun.ahk 里, 不能挪进这个类:
; Options 窗口的 FuncList 把它们的函数名当字符串存进 g_HOTKEY[Trigger*], 运行时
; 靠 CommandRunner.Execute() 里的 %cmdPath%() 按名字动态调用, 只认裸的全局函数名, 不认
; Class.Method - 这两个函数体本身只是薄薄一层, 直接调用 CommandStore.UpdateRank()。
;===============================================================================

Class CommandStore {

    static UpdateRank(originCmd, showRank := false, inc := 1) {
        if (g_CONFIG["SmartRank"] = false || originCmd = "")
            return

        AppData.LoadAppData()

        for _, section in ["DefaultCommand", "UserCommand", "Index"] {
            if !g_CMDDATA[section].Has(originCmd)
                continue

            rankValue := g_CMDDATA[section][originCmd]
            rankValue := IsInteger(rankValue) ? rankValue + inc : inc
            rankValue := (rankValue < 0) ? -1 : rankValue

            g_CMDDATA[section][originCmd] := rankValue
            AppData.SaveAppData()
            if (showRank)
                MainWindow.SetStatus("UpdateRank: Rank for current command : " rankValue)

            g_LOG.Debug("UpdateRank: Rank updated for command..." originCmd "=" rankValue)
            break
        }

        ; Reload in-memory cache so the updated rank takes effect immediately.
        CommandStore.LoadCommands()
    }

    ; UpdateUsage/UpdateRunCount/UpdateHistory only mutate in-memory state; the
    ; caller is responsible for calling AppData.SaveAppData() once all of them are done,
    ; so one command execution costs a single ALTRun.json write, not three.

    static UpdateUsage() {
        currDate := A_YYYY . A_MM . A_DD
        g_USAGE[currDate] := g_USAGE.Has(currDate) ? g_USAGE[currDate] + 1 : 1
        g_RUNTIME["Max"] := Max(g_RUNTIME["Max"], g_USAGE[currDate])
    }

    static UpdateRunCount() {
        g_CONFIG["RunCount"]++
        g_LOG.Debug("UpdateRunCount: RunCount update to..." g_CONFIG["RunCount"])
    }

    static UpdateHistory(originCmd) {
        if (g_CONFIG["SaveHistory"] = false || originCmd = "")
            return

        g_HISTORY.InsertAt(1, originCmd " Arg=" g_RUNTIME["Arg"])

        if (g_HISTORY.Length > g_CONFIG["HistoryLen"])
            g_HISTORY.Pop()
    }

    static LoadCommands() {
        ; Rebuild runtime command caches from the JSON command store.
        Global g_COMMANDS, g_CMDINDEX, g_FALLBACK
        g_COMMANDS := Array()
        g_CMDINDEX := Array()
        g_FALLBACK := Array()
        Local rankRows := ""

        AppData.LoadAppData()                                                   ; Loads (and migrates/creates) ALTRun.json once per session

        for _, sectionName in ["DefaultCommand", "UserCommand", "Index"] {
        for commandText, rankValue in g_CMDDATA[sectionName] {
            if (commandText = "" || !IsInteger(rankValue) || rankValue <= 0)
                continue

            parts := StrSplit(commandText, " | ")
            cmdPath := parts.Has(2) ? parts[2] : ""
            cmdDesc := parts.Has(3) ? parts[3] : ""

            cmdType := parts.Has(1) ? parts[1] : ""
            if (cmdType = "Clip") {
                ; A Clip's field 2 is the snippet body, only its short name (desc) is searchable.
                searchable := cmdDesc
            } else if (g_CONFIG["MatchPath"]) {
                searchable := cmdPath " " cmdDesc
            } else {
                SplitPath(cmdPath, &fileName)
                searchable := fileName " " cmdDesc
            }
            if (g_CONFIG["MatchPinyin"])
                searchable := Pinyin.Initials(searchable)

            rankRows .= rankValue "`t" commandText "`t" searchable "`n"
        }
        }

        ; Sort by rank descending, then rebuild arrays.
        rankRows := Sort(rankRows, "R N")
        for _, line in StrSplit(rankRows, "`n", "`r") {
            if !Trim(line)
                continue

            rowParts := StrSplit(line, "`t") ; rank, command, searchable
            if (rowParts.Length < 3)
                continue
            g_COMMANDS.Push(rowParts[2])
            g_CMDINDEX.Push(rowParts[3])
        }

        ; Fallback commands.
        for _, line in g_CMDDATA["FallbackCommand"] {
            line := Trim(line)
            if (line != "" && SubStr(line, 1, 1) != ";")
                g_FALLBACK.Push(line)
        }

        g_LOG.Debug("LoadCommands: Loaded COMMANDS=" g_COMMANDS.Length ", FALLBACK=" g_FALLBACK.Length)
        return
    }

    static LoadHistory() {                                                     ; g_HISTORY is already populated by AppData.LoadAppData(); just apply policy
        if (!g_CONFIG["SaveHistory"]) {
            if (g_HISTORY.Length) {
                g_HISTORY.Length := 0
                AppData.SaveAppData()
            }
            g_LOG.Debug("LoadHistory: History disabled, cleared.")
            return
        }
        if (g_HISTORY.Length > g_CONFIG["HistoryLen"])
            g_HISTORY.Length := g_CONFIG["HistoryLen"]
        g_LOG.Debug("LoadHistory: Loaded history..." g_HISTORY.Length)
    }
}
