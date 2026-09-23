;===============================================================================
; IniMigration.ahk - 一次性把旧版 ALTRun.ini 迁移到 ALTRun.json (AutoHotkey v2)
;-------------------------------------------------------------------------------
; ALTRun 早期版本(v1)把设置和命令都存在 ALTRun.ini 里, 现在统一改存
; ALTRun.json (原因见 Lib/AppData.ahk 顶部注释)。这个文件只做一件事: 如果
; 检测到用户机器上还有旧的 ALTRun.ini, 且 ALTRun.json 里还没有对应的数据,
; 就把 ini 里的内容读出来、原样搬进 ALTRun.json, 然后把 ini 里已经搬完的
; 小节删掉(整份 ini 文件本身会先备份成 ALTRun.ini.bak, 不会直接删除)。
; 已经在用 JSON 的用户不受任何影响 - 一旦 ALTRun.json 有了对应小节, 这里
; 所有函数都直接跳过, 不会覆盖已有数据。
;
; 用法 (只有 Src\Core\AppData.ahk 的 LoadAppData() 调用, 不需要在别处调用):
;   movedSections := IniMigration.MigrateFromIni(data)   ; data 是从 ALTRun.json
;                                                          ; 解析出来的 Map, 迁移
;                                                          ; 结果直接写回 data 里
;   IniMigration.FinishIniMigration(movedSections)        ; LoadAppData() 把 data
;                                                          ; 存回 ALTRun.json 成功
;                                                          ; 之后再调用这个, 清理 ini
;
; ini 小节名和 ALTRun.json 里的顶层 key 是完全一样的字符串(DefaultCommand/
; UserCommand/Index/FallbackCommand/Config/Hotkey/Gui/Usage/History/Benchmark),
; 所以这里不需要额外的名字映射表。
;===============================================================================

Class IniMigration {

    static IniPath() {
        return A_ScriptDir . "\ALTRun.ini"
    }

    ; One-off: pull whatever ALTRun.ini still has out into `data` (the in-memory
    ; Map that LoadAppData() is about to fold into g_CMDDATA / g_CONFIG / etc).
    ; Nothing is written to disk here - LoadAppData() saves ALTRun.json itself and
    ; only then, via FinishIniMigration(), strips the sections that made it across.
    ; Commands and settings migrate independently, so this is safe to call even
    ; when only one half was ever pulled out of the ini before (as is the case for
    ; anyone who already has an ALTRun.json containing just commands).
    static MigrateFromIni(data) {
        moved := Array()
        if (!FileExist(IniMigration.IniPath()))
            return moved

        iniText := ""
        try iniText := FileRead(IniMigration.IniPath(), "UTF-8")              ; Read the raw file, NOT IniRead - that is what hits the 64 KB cap
        if (InStr(iniText, Chr(0)))                                         ; Old ini saved as UTF-16 by Notepad
            try iniText := FileRead(IniMigration.IniPath(), "UTF-16")
        if (iniText = "")
            return moved

        ; --- Commands: DefaultCommand / UserCommand / Index / FallbackCommand ---
        if (!data.Has("DefaultCommand")) {
            for _, name in ["DefaultCommand", "UserCommand", "Index"] {
                body := IniMigration.ReadIniSectionRaw(iniText, name)
                if (body = "")
                    continue
                data[name] := AppData.ParseCommandBlock(body, true)
                moved.Push(name)
            }

            body := IniMigration.ReadIniSectionRaw(iniText, "FallbackCommand")
            if (body != "") {
                list := Array()
                for _, line in StrSplit(body, "`n", "`r") {
                    line := Trim(IniMigration.UnescapeCommandKey(line))
                    if (line != "" && SubStr(line, 1, 1) != ";")
                        list.Push(line)
                }
                data["FallbackCommand"] := list
                moved.Push("FallbackCommand")
            }
        }

        ; --- Settings: Config / Hotkey / Gui / Usage / History / Benchmark ---
        if (!data.Has("Config")) {
            cfg := IniMigration.ReadIniMapLikeDefaults(g_CONFIG, "Config")
            if (cfg.Count) {
                data["Config"] := cfg
                moved.Push("Config")
            }

            hk := IniMigration.ReadIniMapLikeDefaults(g_HOTKEY, "Hotkey")
            if (hk.Count) {
                data["Hotkey"] := hk
                moved.Push("Hotkey")
            }

            gui := IniMigration.ReadIniMapLikeDefaults(g_GUI, "Gui")
            if (gui.Count) {
                data["Gui"] := gui
                moved.Push("Gui")
            }

            bench := IniMigration.ReadIniMapLikeDefaults(g_BENCH, "Benchmark")
            if (bench.Count) {
                data["Benchmark"] := bench
                moved.Push("Benchmark")
            }

            usageText := ""
            Try usageText := IniRead(IniMigration.IniPath(), "Usage")         ; Whole-section read throws if the section is missing
            if (usageText != "") {
                usage := Map()
                for line in StrSplit(usageText, "`n") {
                    parts := StrSplit(Trim(line, "`r"), "=")
                    if (parts.Length >= 2 && parts[1] != "")
                        usage[parts[1]] := parts[2] + 0
                }
                if (usage.Count) {
                    data["Usage"] := usage
                    moved.Push("Usage")
                }
            }

            history := Array()
            Loop 50 {                                                       ; Generous: the ini's HistoryLen may differ from today's default
                item := IniRead(IniMigration.IniPath(), "History", A_Index, "")
                if (item = "")
                    continue
                history.Push(item)
            }
            if (history.Length) {
                data["History"] := history
                moved.Push("History")
            }
        }

        return moved
    }

    ; Reads an ini section into a Map, coercing each value to the same type
    ; (Integer/Float vs String) as the matching key in `defaultsMap` - e.g. so
    ; "255" migrates as the number 255 but "0xFFFFFF" stays the color string it is.
    static ReadIniMapLikeDefaults(defaultsMap, sectionName) {
        result := Map()
        for key, defVal in defaultsMap {
            val := IniRead(IniMigration.IniPath(), sectionName, key, "")
            if (val = "")
                continue
            result[key] := (Type(defVal) = "Integer" || Type(defVal) = "Float") ? val + 0 : val
        }
        return result
    }

    ; Runs once, right after LoadAppData() has successfully written the migrated
    ; data into ALTRun.json: backs up the old ini, then removes just the sections
    ; that were moved. The ini file itself is kept (as ALTRun.ini.bak plus the
    ; original, now-emptied-of-migrated-sections file) rather than deleted.
    static FinishIniMigration(movedSections) {
        try FileCopy(IniMigration.IniPath(), IniMigration.IniPath() ".bak", true)
        for _, name in movedSections
            Try IniDelete(IniMigration.IniPath(), name)

        g_LOG.Debug("FinishIniMigration: Moved sections [" . IniMigration.JoinArray(movedSections, ", ") . "] from ini to json")
        MsgBox("The following ALTRun.ini section(s) have been moved into ALTRun.json:`n`n"
             . IniMigration.JoinArray(movedSections, ", ") . "`n`n"
             . "The old file was backed up as ALTRun.ini.bak.`n"
             . "You can delete ALTRun.ini once you've confirmed everything still works.", g_TITLE, 64)
    }

    static JoinArray(arr, sep) {
        out := ""
        for _, v in arr
            out .= (out = "" ? "" : sep) . v
        return out
    }

    static ReadIniSectionRaw(iniText, sectionName) {                       ; Section body straight from the file text, no size limit
        if !RegExMatch(iniText, "im)^\[" sectionName "\][ \t]*$", &m)
            return ""
        rest := SubStr(iniText, m.Pos + m.Len)
        if RegExMatch(rest, "m)^\[[^\]\r\n]+\][ \t]*$", &nextSec)
            rest := SubStr(rest, 1, nextSec.Pos - 1)
        return Trim(rest, " `t`r`n")
    }

    ; UnescapeCommandKey(): reverses the old "_Equal_"/"_Semicolon_" encoding that
    ; used to be needed so a command line could survive being an ini key (ini keys
    ; can't contain '=' or ';'). Commands live in JSON now so nothing encodes them
    ; this way anymore, but MigrateFromIni()/AppData.ParseCommandBlock() still call
    ; this when reading a pre-migration ALTRun.ini, since old command lines saved
    ; there were encoded with it.
    static UnescapeCommandKey(cmdLine) {
        if (cmdLine = "")
            return ""

        cmdLine := StrReplace(cmdLine, "_Equal_", "=")      ; equals sign
        cmdLine := StrReplace(cmdLine, "_Semicolon_", ";")  ; semicolon
        return cmdLine
    }
}
