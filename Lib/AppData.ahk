;===================================================
; AppData.ahk - ALTRun.json 读写 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Why not ini: IniRead/IniWrite go through the Windows
; profile API, which truncates a whole section at 64 KB.
; Past that limit commands silently disappear from the
; list. JSON also removes the need to escape "=" and ";"
; in command lines.
;
; Parsing/serializing is done by the shared JSON class in
; Lib\JSON.ahk - AutoHotkey v2 auto-includes it the first
; time JSON.parse()/JSON.stringify() is referenced, because
; the class name matches the file name in the Lib folder
; (same convention already used for Logger, MD5, etc.).
; No #Include line is needed.
;
; ALTRun.json is the single source of truth for everything, including PT
; Tools' own settings (see "PTTools" below, owned by Lib\PTTools.ahk /
; PTToolsWindow). Anyone still on a pre-JSON ALTRun.ini gets migrated
; automatically the next time they start ALTRun - see Lib\IniMigration.ahk
; for that one-off detection/conversion, called from LoadAppData() below.
;
; File layout:
; {
;   "Config"         : { <setting name>: <value>, ... },
;   "Gui"            : { <setting name>: <value>, ... },
;   "Hotkey"         : { <setting name>: <value>, ... },
;   "Usage"          : { "<yyyymmdd>": <run count>, ... },
;   "History"        : [ "<command line> Arg=<arg>", ... ],
;   "Benchmark"      : { <metric name>: <value>, ... },
;   "PTTools"        : { <setting name>: <value>, ... },
;   "DefaultCommand" : { "<command line>": <rank>, ... },
;   "UserCommand"    : { "<command line>": <rank>, ... },
;   "Index"          : { "<command line>": <rank>, ... },
;   "FallbackCommand": [ "<command line>", ... ]
; }
;
; 用法: AppData.LoadAppData() / AppData.SaveAppData() - 都是 ALTRun.ahk 里
; 原来 LoadAppData()/SaveAppData() 的直接替身, 调用方式和返回值完全没变。
;===================================================

Class AppData {

    static LoadAppData(forceReload := false) {
        Global g_CMDDATA, g_HISTORYS

        static loaded := false
        if (!forceReload && loaded)
            return g_CMDDATA
        loaded := true

        g_CMDDATA := Map("DefaultCommand", Map(), "UserCommand", Map(), "Index", Map(), "FallbackCommand", Array())

        data := Map()
        if FileExist(g_JSON) {
            try {
                parsed := JSON.parse(FileRead(g_JSON, "UTF-8"))              ; keepbooltype=false, as_map=true (both defaults)
                if (parsed is Map)
                    data := parsed
            } catch as e {
                g_LOG.Debug("LoadAppData: Invalid JSON - " e.Message)
                try FileMove(g_JSON, g_JSON ".bad", true)
                MsgBox("ALTRun.json could not be parsed:`n`n" e.Message "`n`nIt was renamed to ALTRun.json.bad and the defaults will be rebuilt.", g_TITLE, 48)
                data := Map()
            }
        }

        movedSections := IniMigration.MigrateFromIni(data)                  ; Fills in whatever a pre-JSON ALTRun.ini still has that data[] is missing

        ; --- Commands (DefaultCommand / UserCommand / Index / FallbackCommand) ---
        for _, name in ["DefaultCommand", "UserCommand", "Index"] {
            if !(data.Has(name) && data[name] is Map)
                continue
            for cmdLine, rank in data[name] {
                cmdLine := Trim(cmdLine)
                if (cmdLine = "")
                    continue
                g_CMDDATA[name][cmdLine] := IsInteger(rank) ? rank + 0 : 1
            }
        }
        if (data.Has("FallbackCommand") && data["FallbackCommand"] is Array) {
            for _, cmdLine in data["FallbackCommand"] {
                if (Trim(cmdLine) != "")
                    g_CMDDATA["FallbackCommand"].Push(Trim(cmdLine))
            }
        }

        dirty := movedSections.Length > 0
        if (!g_CMDDATA["DefaultCommand"].Count) {
            g_CMDDATA["DefaultCommand"] := AppData.ParseCommandBlock(AppData.DefaultCommandText())
            dirty := true
        }
        if (!g_CMDDATA["UserCommand"].Count) {
            g_CMDDATA["UserCommand"] := AppData.ParseCommandBlock(AppData.UserCommandText())
            dirty := true
        }
        if (!g_CMDDATA["FallbackCommand"].Length) {
            for _, line in StrSplit(AppData.FallbackCommandText(), "`n", "`r") {
                line := Trim(line)
                if (line != "" && SubStr(line, 1, 1) != ";")
                    g_CMDDATA["FallbackCommand"].Push(line)
            }
            dirty := true
        }

        ; --- Settings (Config / Gui / Hotkey) - overlay JSON values onto the hardcoded defaults ---
        AppData.MergeIntoDefaults(g_CONFIG, data.Get("Config", ""))
        AppData.MergeIntoDefaults(g_HOTKEY, data.Get("Hotkey", ""))
        AppData.MergeIntoDefaults(g_GUI,    data.Get("Gui", ""))
        AppData.MergeIntoDefaults(g_BENCH,  data.Get("Benchmark", ""))
        PTToolsWindow.Load(data.Get("PTTools", ""))
        g_RUNTIME["RegEx"] := g_CONFIG["MatchBeginning"] ? "imS)^" : "imS)"

        ; --- Usage: keep only the last 30 days, same trimming rule as before ---
        if (data.Has("Usage") && data["Usage"] is Map) {
            for dateKey, dayCount in data["Usage"]
                g_USAGE[dateKey] := dayCount
        }
        offsetDate := DateAdd(A_Now, -30, "Days")
        for dateKey in g_USAGE.Clone()                                      ; Clone: we mutate g_USAGE while iterating it
            if (dateKey <= SubStr(offsetDate, 1, 8))
                g_USAGE.Delete(dateKey)
        Loop 30 {
            offsetDate := DateAdd(offsetDate, 1, "Days")
            dateKey := SubStr(offsetDate, 1, 8)
            g_USAGE[dateKey] := g_USAGE.Has(dateKey) ? g_USAGE[dateKey] : 0
            g_RUNTIME["Max"] := Max(g_RUNTIME["Max"], g_USAGE[dateKey])
        }

        ; --- History ---
        g_HISTORYS.Length := 0
        if (data.Has("History") && data["History"] is Array) {
            for _, entry in data["History"]
                if (Trim(entry) != "")
                    g_HISTORYS.Push(entry)
        }

        if (dirty && AppData.SaveAppData() && movedSections.Length)
            IniMigration.FinishIniMigration(movedSections)                  ; Backs up ALTRun.ini, then strips the now-migrated sections

        g_LOG.Debug("LoadAppData: Default=" g_CMDDATA["DefaultCommand"].Count
            . ", User=" g_CMDDATA["UserCommand"].Count
            . ", Index=" g_CMDDATA["Index"].Count
            . ", Fallback=" g_CMDDATA["FallbackCommand"].Length
            . ", History=" g_HISTORYS.Length)

        if (!g_CMDDATA["Index"].Count) {
            if (MsgBox(g_LNG[804], g_TITLE, 4161) = "OK")
                Reindex()
        }
        return g_CMDDATA
    }

    static MergeIntoDefaults(defaultsMap, sourceMap) {                     ; Overlay JSON values onto a defaults Map, key by key
        if !(sourceMap is Map)                                              ; Keeps default (and Map order) for any key the file doesn't have yet
            return                                                          ; - e.g. a setting added in a newer version of ALTRun.
        for key, _ in defaultsMap
            if sourceMap.Has(key)
                defaultsMap[key] := sourceMap[key]
    }

    static SaveAppData() {
        ordered := Map(
            "Config",  g_CONFIG,
            "Gui",     g_GUI,
            "Hotkey",  g_HOTKEY,
            "Usage",   g_USAGE,
            "History", g_HISTORYS,
            "Benchmark", g_BENCH,
            "PTTools", PTToolsWindow.Settings,
            "DefaultCommand", Map(),
            "UserCommand",    Map(),
            "Index",          Map(),
            "FallbackCommand", Array()
        )
        ; Rebuild ranks as plain integers so JSON.stringify emits numbers, not strings.
        for _, name in ["DefaultCommand", "UserCommand", "Index"] {
            for cmdLine, rank in g_CMDDATA[name]
                ordered[name][cmdLine] := IsInteger(rank) ? rank + 0 : 1
        }
        for _, cmdLine in g_CMDDATA["FallbackCommand"]
            ordered["FallbackCommand"].Push(cmdLine)

        out := JSON.stringify(ordered)                                     ; new JSON.ahk always indents 2 spaces, no separate "space" param

        tmpFile := g_JSON ".tmp"
        try {
            if FileExist(tmpFile)
                FileDelete(tmpFile)
            FileAppend(out, tmpFile, "UTF-8")                               ; Write a temp file first, so a crash can never truncate the real one
            FileMove(tmpFile, g_JSON, true)
        } catch as e {
            g_LOG.Debug("SaveAppData: Write failed - " e.Message)
            MsgBox("Could not save ALTRun.json:`n`n" e.Message, g_TITLE, 48)
            return false
        }
        return true
    }

    ; An OnExit callback that RETURNS a nonzero/true value cancels the exit (this
    ; is documented AutoHotkey v2 behavior, not a bug) - so this must NOT simply
    ; forward a true/false result the way an inline `(*) => SaveAppData()` would.
    ; No explicit `return` here means this always yields "" (falsy), so
    ; Reload()/ExitApp() are never blocked, even if the save itself fails.
    static OnAppExit(*) {
        AppData.SaveAppData()
        Logger.Flush()                                                     ; Buffered log lines are lost otherwise - see Lib/Logger.ahk
    }

    static ParseCommandBlock(blockText, legacyUnescape := false) {         ; "command line=rank" lines -> Map
        result := Map()
        for _, line in StrSplit(blockText, "`n", "`r") {
            line := Trim(line)
            if (!line || SubStr(line, 1, 1) = ";" || SubStr(line, 1, 1) = "[")
                continue
            if (legacyUnescape)
                line := IniMigration.UnescapeCommandKey(line)
            if !RegExMatch(line, "^(.*)=(\d+)\s*$", &m)                     ; Split on the LAST '=', so the command itself may contain '='
                continue
            cmdLine := Trim(m.1)
            rank    := m.2 + 0
            if (cmdLine != "" && rank > 0)
                result[cmdLine] := rank
        }
        return result
    }

    static DefaultCommandText() {
        return "
        (
            ; This section is Built-In commands with high priority
            ; App will auto generate this section while it is empty
            ; Please make sure App is not running before modifying.
            ;
            Func | About | Help & About (F1)=99
            Func | Options | Setting Options (F2)=99
            Func | Reload | Reload ALTRun=99
            Func | EditCommand | Edit current command (F3)=99
            Func | UserCommand | Edit command database ALTRun.json (F4)=99
            Func | NewCommand | New Command=99
            Func | NewClip | New Clip (text snippet)=99
            Func | OpenContainer | Locate cmd's dir with File Manager=99
            Func | Usage | ALTRun Usage Status=99
            Func | Reindex | Reindex search database=99
            Func | Everything | Search by Everything=99
            Func | PTTools | PT Tools (Rebar/BRC calculator + SPF2M)=99
            Func | AhkRun | Run Command use AutoHotkey Run=99
            Func | Google | Search Clipboard or Input by Google=99
            Func | Bing | Search Clipboard or Input by Bing=99
            Func | EmptyRecycle | Empty Recycle Bin=99
            Func | TurnMonitorOff | Turn off Monitor, Close Monitor=99
            Func | MuteVolume | Mute Volume=99
            File | %Temp%\ALTRun.log | ALTRun Log File=99
            Dir | A_ScriptDir | ALTRun Program Dir=99
            Dir | A_Startup | Current User Startup Dir=99
            Dir | A_StartupCommon | All User Startup Dir=99
            Dir | A_ProgramsCommon | Windows Search.Index.Cortana Dir=99
            CMD | explorer.exe | Windows File Explorer=99
            CMD | cmd.exe | Windows Command Processor=99
            CMD | Shell:AppsFolder | AppsFolder Applications=66
            CMD | ::{645FF040-5081-101B-9F08-00AA002F954E} | Recycle Bin=66
            CMD | Notepad.exe | Notepad=66
            CMD | WF.msc | Windows Defender Firewall with Advanced Security=66
            CMD | TaskSchd.msc | Task Scheduler=66
            CMD | DevMgmt.msc | Device Manager=66
            CMD | EventVwr.msc | Event Viewer=66
            CMD | CompMgmt.msc | Computer Manager=66
            CMD | TaskMgr.exe | Task Manager=66
            CMD | Calc.exe | Calculator=66
            CMD | MsPaint.exe | Paint=66
            CMD | Regedit.exe | Registry Editor=66
            CMD | CleanMgr.exe | Disk Space Clean-up Manager=66
            CMD | GpEdit.msc | Group Policy=66
            CMD | DiskMgmt.msc | Disk Management=66
            CMD | DxDiag.exe | Directx Diagnostic Tool=66
            CMD | LusrMgr.msc | Local Users and Groups=66
            CMD | MsConfig.exe | System Configuration=66
            CMD | PerfMon.exe /Res | Resources Monitor=66
            CMD | PerfMon.exe | Performance Monitor=66
            CMD | WinVer.exe | About Windows=66
            CMD | Services.msc | Services=66
            CMD | NetPlWiz | User Accounts=66
            CMD | Control | Control Panel=66
            CMD | Control Intl.cpl | Region and Language Options=66
            CMD | Control Firewall.cpl | Windows Defender Firewall=66
            CMD | Control AppWiz.cpl | Programs and Features=66
            CMD | Control Sysdm.cpl | System Properties=66
            CMD | Control AdminTools | Windows Tools=66
            CMD | Control Inetcpl.cpl,,4 | Internet Properties=66
            CMD | Control UserPasswords | User Accounts=66
        )"
    }

    static UserCommandText() {
        return "
        (
            ; This section is User-Defined commands, modify as desired
            ; Format: Command Type | Command | Description=Rank
            ; Command type: File, Dir, CMD, URL, some sample below
            ; Please make sure App is not running before modifying
            ;
            File | C:\Windows\Notepad.exe=9
            Dir | %AppData%\Microsoft\Windows\SendTo | Windows SendTo Dir=9
            Dir | %OneDrive% | OneDrive=9
            Dir | A_Desktop | Desktop=99
            CMD | cmd.exe /k ipconfig | Check IP Address=9
            CMD | explorer /Select,C:\Program Files | Open and select C:\Program Files=9
            CMD | Control Printers | Devices and Printers=66
            CMD | ::{20D04FE0-3AEA-1069-A2D8-08002B30309D} | This PC=9
            URL | www.google.com | Google=9
            Clip | Dear Sir,\n\nThank you for your email.\n\nBest regards,\nLiming | sig=9
            Clip | {date} | today=9
        )"
    }

    static FallbackCommandText() {
        return "
        (
            ; Fallback Commands show when search result is empty
            ; Commands in order, modify as desired
            ; Format: Command Type | Command | Description
            ; Command Type: File, Dir, CMD, URL
            ;
            Func | NewCommand | New Command
            Func | Everything | Search by Everything
            Func | Google | Search Clipboard or Input by Google
            Func | AhkRun | Run Command use AutoHotkey Run
            Func | Bing | Search Clipboard or Input by Bing
            CMD | Calc.exe | Calculator
        )"
    }
}
