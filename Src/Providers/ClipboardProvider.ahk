;===============================================================================
; ClipboardProvider.ahk - 剪贴板历史 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 Alfred 的 Clipboard History 一样: 记住复制过的文字, 输入 "clip" (Keyword)
; 或按热键 (默认 Ctrl+Alt+C) 列出, "clip 关键词" 过滤, Enter 粘贴到前台窗口。
;
; 隐私:
;   - 密码管理器等程序复制的内容 (带 ExcludeClipboardContentFromMonitorProcessing
;     或 CanIncludeInClipboardHistory = 0 标记) 不会被记录
;   - IgnoreApps 里的程序 (进程名) 复制的内容不会被记录
;   - Persist = 0 时只保存在内存里, 退出即清空
;
; 设置 (ALTRun.json -> Features.Clipboard):
;   Enabled / Keyword / Hotkey / MaxItems / MaxItemLength / Persist / IgnoreApps
;
; 用法:
;   ClipboardProvider.PauseRecording(ms)   ALTRun 自己临时借用剪贴板时调用, 这段时间不记录
;   ClipboardProvider.Add(text, source)    手动加入一条 (source = 来源程序的进程名)
;===============================================================================

class ClipboardProvider {
    static Id      := "Clipboard"
    static File    := A_ScriptDir "\Data\ClipboardHistory.json"
    static Entries := []              ; [Map("Text", "Time", "App")], 最新的在前
    static _pausedUntil := 0, _saveTimer := ""

    static Init() {
        if AppSettings.Feature("Clipboard")["Persist"]
            ClipboardProvider._Load()
        OnClipboardChange((dataType) => ClipboardProvider._OnChange(dataType))
    }

    static PauseRecording(milliseconds) {
        ClipboardProvider._pausedUntil := A_TickCount + milliseconds
    }

    static Search(query) {
        options := AppSettings.Feature("Clipboard")
        if !query.MatchKeyword([options["Keyword"]], &term)
            return []
        icon := "res:imageres.dll,-5314"
        exclusive := query.HasRest                                          ; "clip " 之后只显示剪贴板历史
        if !ClipboardProvider.Entries.Length
            return [ResultItem(I18n.T("Clipboard.Empty"), I18n.T("Clipboard.EmptyHint", Win.HotkeyLabel(options["Hotkey"])), {Icon: icon, Valid: false, Score: 300, Exclusive: exclusive})]

        results := []
        tokens := StrSplit(Trim(term), " ")
        for index, entry in ClipboardProvider.Entries {
            if !ClipboardProvider._Matches(entry["Text"], tokens)
                continue
            item := ClipboardProvider._ToItem(entry, icon, 300 - results.Length * 0.001)
            item.Exclusive := exclusive
            results.Push(item)
            if (results.Length >= ProviderRegistry.MaxResults - 1)
                break
        }
        if (term = "")
            results.Push(ResultItem(I18n.T("Clipboard.Clear"), I18n.T("Clipboard.ClearHint", ClipboardProvider.Entries.Length), {
                Icon: "res:shell32.dll,-32", Score: 0, Exclusive: exclusive, OnRun: (*) => ClipboardProvider.Clear()
            }))
        return results
    }

    static _Matches(text, tokens) {
        for token in tokens
            if (token != "" && !InStr(text, token))
                return false
        return true
    }

    static _ToItem(entry, icon, score) {
        text := entry["Text"]
        subtitle := I18n.T("Clipboard.Subtitle", ClipboardProvider._FormatTime(entry["Time"]), entry["App"] != "" ? entry["App"] : "?", StrLen(text))
        return ResultItem(ClipboardProvider._Preview(text), subtitle, {
            Kind: "text", Arg: text, Icon: icon, Score: score, LargeText: text,
            OnRun: (item) => ClipboardProvider.Paste(item.Arg),
            Actions: [ResultItem(I18n.T("Clipboard.Delete"), "", {Icon: "res:shell32.dll,-32", OnRun: (item) => ClipboardProvider.Remove(item.Arg)})]
        })
    }

    ; 粘贴一条历史到前台窗口, 并把它移到最前面
    static Paste(text) {
        ClipboardProvider.Add(text, ClipboardProvider._EntryApp(text))
        ActionCatalog.PasteText(text)
    }

    ;---------------------------------------------------------------------------
    ; History
    ;---------------------------------------------------------------------------
    static Add(text, source := "") {
        options := AppSettings.Feature("Clipboard")
        if (Trim(text, " `t`r`n") = "" || StrLen(text) > options["MaxItemLength"])
            return false
        ClipboardProvider._RemoveText(text)
        ClipboardProvider.Entries.InsertAt(1, Map("Text", text, "Time", A_Now, "App", source))
        while (ClipboardProvider.Entries.Length > options["MaxItems"])
            ClipboardProvider.Entries.Pop()
        ClipboardProvider._SaveLater()
        return true
    }

    static Remove(text) {
        ClipboardProvider._RemoveText(text)
        ClipboardProvider._SaveLater()
    }

    static Clear() {
        ClipboardProvider.Entries := []
        ClipboardProvider._SaveLater()
        App.Notify(I18n.T("Clipboard.Cleared"))
    }

    static _RemoveText(text) {
        for index, entry in ClipboardProvider.Entries {
            if (entry["Text"] == text) {
                ClipboardProvider.Entries.RemoveAt(index)
                return
            }
        }
    }

    static _EntryApp(text) {
        for entry in ClipboardProvider.Entries
            if (entry["Text"] == text)
                return entry["App"]
        return ""
    }

    ;---------------------------------------------------------------------------
    ; Recording
    ;---------------------------------------------------------------------------
    static _OnChange(dataType) {
        if (dataType != 1 || A_TickCount < ClipboardProvider._pausedUntil)  ; 1 = 文字 (包括复制的文件路径)
            return
        if ClipboardProvider._IsPrivate()
            return
        processName := ""
        try processName := WinGetProcessName("A")
        for ignored in AppSettings.Feature("Clipboard")["IgnoreApps"]
            if (processName = ignored)
                return
        text := ""
        try text := A_Clipboard
        ClipboardProvider.Add(text, processName)
    }

    ; 密码管理器等会给剪贴板内容加上 "不要记录" 的标记 (Windows 剪贴板历史也遵守)
    static _IsPrivate() {
        static excludeFormat := DllCall("RegisterClipboardFormat", "Str", "ExcludeClipboardContentFromMonitorProcessing", "UInt")
        static historyFormat := DllCall("RegisterClipboardFormat", "Str", "CanIncludeInClipboardHistory", "UInt")
        if DllCall("IsClipboardFormatAvailable", "UInt", excludeFormat)
            return true
        if !DllCall("IsClipboardFormatAvailable", "UInt", historyFormat)
            return false
        allowed := 1
        if DllCall("OpenClipboard", "Ptr", A_ScriptHwnd) {
            if (handle := DllCall("GetClipboardData", "UInt", historyFormat, "Ptr")) {
                if (pointer := DllCall("GlobalLock", "Ptr", handle, "Ptr")) {
                    allowed := NumGet(pointer, "UInt")
                    DllCall("GlobalUnlock", "Ptr", handle)
                }
            }
            DllCall("CloseClipboard")
        }
        return !allowed
    }

    ;---------------------------------------------------------------------------
    ; Formatting
    ;---------------------------------------------------------------------------
    static _Preview(text) {
        text := Trim(RegExReplace(text, "\s+", " "))
        return (StrLen(text) > 120) ? SubStr(text, 1, 120) "..." : text
    }

    static _FormatTime(timestamp) {
        if (SubStr(timestamp, 1, 8) = FormatTime(, "yyyyMMdd"))
            return FormatTime(timestamp, "HH:mm")
        return FormatTime(timestamp, "yyyy-MM-dd HH:mm")
    }

    ;---------------------------------------------------------------------------
    ; Persistence
    ;---------------------------------------------------------------------------
    static _Load() {
        if !FileExist(ClipboardProvider.File)
            return
        try {
            data := JSON.Parse(FileRead(ClipboardProvider.File, "UTF-8"))
            if (data is Map && data.Has("Entries") && data["Entries"] is Array)
                ClipboardProvider.Entries := data["Entries"]
        } catch as e {
            Logger.Error("ClipboardProvider: cannot read history - " e.Message)
        }
    }

    static Save() {
        if !AppSettings.Feature("Clipboard")["Persist"]
            return
        try {
            DirCreate(AppSettings.DataDir)
            tmpFile := ClipboardProvider.File ".tmp"
            try FileDelete(tmpFile)
            FileAppend(JSON.Stringify(Map("Entries", ClipboardProvider.Entries)), tmpFile, "UTF-8")
            FileMove(tmpFile, ClipboardProvider.File, true)
        } catch as e {
            Logger.Error("ClipboardProvider: cannot write history - " e.Message)
        }
    }

    static _SaveLater() {
        if (ClipboardProvider._saveTimer = "")
            ClipboardProvider._saveTimer := () => ClipboardProvider.Save()
        SetTimer(ClipboardProvider._saveTimer, -2000)
    }
}
