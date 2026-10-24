;===============================================================================
; SchemaMigration.ahk - ALTRun.json 版本升级 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 每个版本的升级是一个独立的 _FromN(data) 方法, 负责把版本 N 的数据转换成
; 版本 N+1。Upgrade() 从文件当前的版本开始逐级调用, 直到 CurrentVersion:
;   ALTRun.ini           已发布的 2.x (到 v2026.08.12) 的设置文件, ReadLegacyIni() 读成
;                        版本 2 的结构, 再走下面的逐级升级 (见 AppSettings.Load)
;   2 -> 3   _From2()   ALTRun 2.x (Config/Gui/Hotkey/UserCommand...) -> 3.0
;   3 -> 4   _From3()   文件搜索默认不再混进普通结果 (InDefaultResults 1 -> 0)
;   4 -> 5   _From4()   片段的粘贴方式 Clipboard -> Auto; 补上后来加进默认值的窗口 / 关键字; 小清理 (见 issue #176)
;   5 -> 6   _From5()   window lists become arrays; StandardDialogs split out of DialogWindows; two hotkeys renamed to
;                       Hotkey; unused keys removed; old ScopeExclude default and Compact themes replaced (#176)
;
; 以后改设置结构时: AppSettings.CurrentVersion + 1, 再在这里加一个 _FromN()。
; 升级前 Backup() 会把原文件复制成 ALTRun.v<N>.backup.json。
;
; 2.x -> 3.0 保留的内容:
;   - 呼出热键、开机启动、托盘、失焦隐藏、文件管理器等通用设置
;   - 用户命令 (UserCommand): File/Dir/CMD/URL -> CustomCommands, Clip -> Snippets
;   - 索引目录/类型/深度、结构计算开关、片段粘贴方式
;   - Listary 路径跳转、Ctrl+D 加日期、PT Tools 的设置
;   - 条件热键 (例如 RAPT 里中键打开 PT Tools) -> Hotkeys
; 不再保留: 内置命令列表 (DefaultCommand, 已由系统命令功能取代)、旧索引
; (Index, 会重新建立)、执行历史、使用统计、列表外观相关的旧选项。
;===============================================================================

class SchemaMigration {

    static DetectVersion(data) {
        if (data.Has("SchemaVersion") && IsInteger(data["SchemaVersion"]))
            return data["SchemaVersion"] + 0
        for key in ["Config", "UserCommand", "DefaultCommand", "Hotkey"]
            if data.Has(key)
                return 2
        return AppSettings.CurrentVersion
    }

    static Upgrade(data, fromVersion) {
        version := fromVersion
        while (version < AppSettings.CurrentVersion) {
            step := "_From" version
            if !HasMethod(SchemaMigration, step)
                throw Error("No settings migration from version " version)
            data := SchemaMigration.%step%(data)
            version++
            data["SchemaVersion"] := version
            Logger.Debug("SchemaMigration: upgraded settings to version " version)
        }
        return data
    }

    ;---------------------------------------------------------------------------
    ; ALTRun.ini (2.x)
    ;---------------------------------------------------------------------------
    static LegacyIniFile(jsonFile) {
        SplitPath(jsonFile, , &dir)
        return dir "\ALTRun.ini"
    }

    ; 读 2.x 的 ALTRun.ini -> Map("Config", Map(...), "Gui", ..., "Hotkey", ..., "UserCommand", ...)
    ; 和 _From2() 期望的版本 2 结构一样。只读需要转换的节; Index / History / Usage /
    ; DefaultCommand / FallbackCommand 不再使用。
    ;   普通节: "键=值", 按第一个 = 分开, 数字转成整数
    ;   UserCommand: "类型 | 目标 | 说明=排名", 按最后一个 = 分开; 键里的 = 和 ; 在 2.x 里
    ;                写成 _Equal_ / _Semicolon_, 这里还原
    static ReadLegacyIni(iniFile) {
        data := Map()
        for section in ["Config", "Gui", "Hotkey", "UserCommand"] {
            text := ""
            try text := IniRead(iniFile, section)
            entries := Map()
            isCommands := (section = "UserCommand")
            for line in StrSplit(text, "`n", "`r") {
                line := Trim(line)
                if (line = "" || SubStr(line, 1, 1) = ";")
                    continue
                pos := isCommands ? InStr(line, "=", , -1) : InStr(line, "=")
                if (pos <= 1)
                    continue
                key := Trim(SubStr(line, 1, pos - 1)), value := Trim(SubStr(line, pos + 1))
                if isCommands
                    key := StrReplace(StrReplace(key, "_Equal_", "="), "_Semicolon_", ";")
                entries[key] := IsInteger(value) ? Integer(value) : value
            }
            data[section] := entries
        }
        return data
    }

    static BackupFile(file, version) {
        SplitPath(file, , &dir)
        return dir "\ALTRun.v" version ".backup.json"
    }

    static Backup(file, version) {
        if !FileExist(file)
            return ""
        backup := SchemaMigration.BackupFile(file, version)
        try FileCopy(file, backup, true)
        catch as e {                                                        ; the upgrade still goes on; the log says the backup is missing
            Logger.Error("SchemaMigration: cannot back up the settings as " backup " - " e.Message)
            return ""
        }
        return backup
    }

    ;---------------------------------------------------------------------------
    ; 5 -> 6
    ;---------------------------------------------------------------------------
    ; One settings format from here on (#176):
    ;   - window lists (QuickSwitch DialogWindows / ExcludeWindows / AutoSwitchExclude, AutoDate RenameWindows /
    ;     AppendWindows, Snippets ExpandExclude) were comma-separated strings: now arrays, like every other list (a window
    ;     title with a comma in it can be written now). "ahk_class #32770" in DialogWindows becomes StandardDialogs = 1
    ;   - QuickSwitch.TotalCmdHotkey and AutoDate.RenameHotkey become Hotkey (one hotkey each since 2026.10.20);
    ;     ExplorerHotkey, AppendHotkey and Applications.RefreshMinutes were no longer read: removed
    ;   - FileSearch.ScopeExclude still at the old default: the new default (skips more caches); the index is rescanned
    ;   - the removed DarkCompact / LightCompact / MidnightCompact themes: Dark / Light / Midnight at Normal density
    ;     (unless the user has a theme file of that name). ThemeManager.LegacyCompact still reads them (a theme's Base)
    static OldScopeExclude := "i)\\(node_modules|\.git|__pycache__|\$RECYCLE\.BIN)(\\|$)"
    static _From5(data) {
        features := SchemaMigration._Section(data, "Features")
        extensions := SchemaMigration._Section(data, "Extensions")

        quick := SchemaMigration._Section(extensions, "QuickSwitch")
        SchemaMigration._RenameKey(quick, "TotalCmdHotkey", "Hotkey")
        SchemaMigration._DeleteKey(quick, "ExplorerHotkey")
        if (quick.Has("DialogWindows") && quick["DialogWindows"] is String) {
            others := [], standard := 0
            for window in SchemaMigration._SplitList(quick["DialogWindows"])
                if (window = "ahk_class #32770")
                    standard := 1
                else
                    others.Push(window)
            quick["StandardDialogs"] := standard, quick["DialogWindows"] := others
        }
        SchemaMigration._ListFromText(quick, "ExcludeWindows")
        SchemaMigration._ListFromText(quick, "AutoSwitchExclude")

        dateStamp := SchemaMigration._Section(extensions, "AutoDate")
        SchemaMigration._RenameKey(dateStamp, "RenameHotkey", "Hotkey")
        SchemaMigration._DeleteKey(dateStamp, "AppendHotkey")
        SchemaMigration._ListFromText(dateStamp, "RenameWindows")
        SchemaMigration._ListFromText(dateStamp, "AppendWindows")

        SchemaMigration._ListFromText(SchemaMigration._Section(features, "Snippets"), "ExpandExclude")
        SchemaMigration._DeleteKey(SchemaMigration._Section(features, "Applications"), "RefreshMinutes")

        fileSearch := SchemaMigration._Section(features, "FileSearch")
        if (fileSearch.Has("ScopeExclude") && fileSearch["ScopeExclude"] == SchemaMigration.OldScopeExclude) {
            fileSearch["ScopeExclude"] := AppSettings.Defaults()["Features"]["FileSearch"]["ScopeExclude"]
            FileIndex.ExpireCache()                                         ; rescan with the new exclusions
        }

        appearance := SchemaMigration._Section(data, "Appearance")
        if (appearance.Has("Theme") && RegExMatch(appearance["Theme"], "i)^(Dark|Light|Midnight)Compact$", &m)
            && !FileExist(AppSettings.ThemesDir "\" appearance["Theme"] ".json")) {
            appearance["Theme"] := m[1]
            if (!appearance.Has("Density") || appearance["Density"] = "Comfortable")
                appearance["Density"] := "Normal"
        }
        return data
    }

    ; section[key]: a comma-separated string -> an array of its items (already an array, or missing: left alone)
    static _ListFromText(section, key) {
        if (section.Has(key) && section[key] is String)
            section[key] := SchemaMigration._SplitList(section[key])
    }

    static _RenameKey(section, oldKey, newKey) {
        if !section.Has(oldKey)
            return
        if !section.Has(newKey)
            section[newKey] := section[oldKey]
        section.Delete(oldKey)
    }

    static _DeleteKey(section, key) {
        if section.Has(key)
            section.Delete(key)
    }

    ;---------------------------------------------------------------------------
    ; 4 -> 5
    ;---------------------------------------------------------------------------
    ; 只改老用户不会觉得意外的地方 (看得见的改动, 例如主题跟随系统、Alt+R, 只对新安装生效; 清单见 GitHub issue #176):
    ;   - 片段的粘贴方式: Clipboard (旧的默认值, 几乎都是当时自动写入的) -> Auto (更快, 不动剪贴板); Type 不变
    ;   - 后来补进默认值的窗口 / 关键字, 老用户的列表里也补上 (缺什么补什么, 用户自己加的保留; 清空了的列表不动)
    ;   - 残留的 "F1 -> 关于" 热键 (F1 已经内置) 删掉; 钢筋直径存成数字
    static _From4(data) {
        features := SchemaMigration._Section(data, "Features")
        snippets := SchemaMigration._Section(features, "Snippets")
        if (snippets.Has("PasteMode") && snippets["PasteMode"] = "Clipboard")
            snippets["PasteMode"] := "Auto"

        fileSearch := SchemaMigration._Section(features, "FileSearch")
        if (fileSearch.Has("Keywords") && fileSearch["Keywords"] is Array && fileSearch["Keywords"].Length) {
            found := false
            for keyword in fileSearch["Keywords"]
                found := found || (keyword = "file")
            if !found
                fileSearch["Keywords"].Push("file")
        }
        calculator := SchemaMigration._Section(features, "Calculator")
        if (calculator.Has("BarSizes") && calculator["BarSizes"] is Array)
            for i, size in calculator["BarSizes"]
                if (size is String && IsNumber(Trim(size)))
                    calculator["BarSizes"][i] := Trim(size) + 0

        extensions := SchemaMigration._Section(data, "Extensions")
        dateStamp := SchemaMigration._Section(extensions, "AutoDate")
        SchemaMigration._AddWindows(dateStamp, "RenameWindows", ["ahk_class TSTDTREEDLG", "ahk_class TCOMBOINPUT"])   ; Total Commander 的对话框 / 改名输入框
        if (dateStamp.Has("AppendWindows") && dateStamp["AppendWindows"] = "ahk_class TCmtEditForm,ahk_exe Notepad4.exe")
            dateStamp["AppendWindows"] := "ahk_class TCmtEditForm, ahk_exe Notepad4.exe"   ; 旧的默认值, 只是统一写法
        switcher := SchemaMigration._Section(extensions, "QuickSwitch")
        SchemaMigration._AddWindows(switcher, "ExcludeWindows", ["ahk_exe 7zG.exe", "ahk_exe TOTALCMD64.exe"])   ; 7-Zip 解压、TC 的对话框

        if (data.Has("Hotkeys") && data["Hotkeys"] is Array) {
            i := data["Hotkeys"].Length
            while (i >= 1) {
                entry := data["Hotkeys"][i]
                if (entry is Map && entry.Has("Key") && entry["Key"] = "F1" && entry.Has("Action") && entry["Action"] = "About"
                    && entry.Has("WinTitle") && entry["WinTitle"] = "ALTRun")
                    data["Hotkeys"].RemoveAt(i)
                i--
            }
        }
        return data
    }

    ; 逗号分隔的窗口列表 section[key] 里补上 windows 里还没有的 (不区分大小写); 没有这一项、或者用户清空了的不动
    static _AddWindows(section, key, windows) {
        if (!section.Has(key) || !(section[key] is String) || Trim(section[key]) = "")
            return
        have := Map(), have.CaseSense := "Off"
        for item in SchemaMigration._SplitList(section[key])
            have[item] := true
        text := RTrim(Trim(section[key]), ",")
        for window in windows
            if !have.Has(window)
                text .= ", " window
        section[key] := text
    }

    ;---------------------------------------------------------------------------
    ; 3 -> 4
    ;---------------------------------------------------------------------------
    ; 3.0 早期默认在普通结果里显示全盘匹配的文件, 会混进 NIRMALA.TTF 这类无关文件;
    ; 现在默认关闭, 需要时用 空格 / ' / open / find 专门搜索文件。旧文件里的 1 只是当时
    ; 自动写入的默认值, 这里改成新的默认值 (想要的话可以在偏好设置里重新打开)。
    static _From3(data) {
        features := SchemaMigration._Section(data, "Features")
        fileSearch := SchemaMigration._Section(features, "FileSearch")
        if fileSearch.Has("InDefaultResults")
            fileSearch["InDefaultResults"] := 0
        return data
    }

    ;---------------------------------------------------------------------------
    ; 2.x -> 3
    ;---------------------------------------------------------------------------
    static _From2(old) {
        config := SchemaMigration._Section(old, "Config")
        oldHotkeys := SchemaMigration._Section(old, "Hotkey")
        pick(section, key, fallback) => section.Has(key) ? section[key] : fallback

        data := AppSettings.Defaults()

        general := data["General"]
        general["Hotkey"]               := SchemaMigration._CleanHotkey(pick(oldHotkeys, "GlobalHotkey1", "!Space"))
        general["SecondaryHotkey"]      := SchemaMigration._CleanHotkey(pick(oldHotkeys, "GlobalHotkey2", ""))
        general["Language"]             := pick(config, "Chinese", 0) ? "zh-CN" : "auto"
        general["LaunchAtLogin"]        := pick(config, "AutoStartup", 1)
        general["ShowTrayIcon"]         := pick(config, "ShowTrayIcon", 1)
        general["HideOnDeactivate"]     := pick(config, "HideOnLostFocus", 1)
        general["SwitchToEnglishInput"] := pick(config, "AutoEngIME", 0)
        general["SpaceToRun"]           := pick(config, "SpaceToRun", 0)
        general["KeepLastQuery"]        := pick(config, "KeepInput", 1)     ; 2.x 默认保留输入
        general["FileManager"]          := pick(config, "FileMgr", "explorer.exe")
        general["SendToMenu"]           := pick(config, "EnableSendTo", 1)
        general["StartMenuShortcut"]    := pick(config, "InStartMenu", 1)
        general["CheckForUpdates"]      := pick(config, "AutoUpdateCheck", 1)
        general["SaveLog"]              := pick(config, "SaveLog", 0)

        apps := data["Features"]["Applications"]
        if (pick(config, "IndexDir", "") != "")
            apps["Folders"] := SchemaMigration._SplitList(config["IndexDir"])
        if (pick(config, "IndexType", "") != "")
            apps["FileTypes"] := SchemaMigration._SplitList(config["IndexType"])
        apps["Depth"]       := pick(config, "IndexDepth", apps["Depth"])
        apps["StoreApps"]   := pick(config, "IndexStoreApp", 1)
        apps["MatchPinyin"] := pick(config, "MatchPinyin", 1)

        data["Features"]["Calculator"]["StructuralCalc"] := pick(config, "StruCalc", 0)
        data["Features"]["Snippets"]["PasteMode"]  := (pick(config, "ClipSendMode", 1) = 2) ? "Type" : "Clipboard"
        data["Features"]["Snippets"]["PasteDelay"] := pick(config, "ClipPasteDelay", 300)

        everythingPath := pick(config, "Everything", "")
        if (everythingPath != "" && everythingPath != "C:\Apps\Everything.exe")
            data["Features"]["FileSearch"]["EverythingPath"] := everythingPath

        ; The window lists are written as the comma-separated text of 2.x (the later steps expect that and turn them into
        ; arrays in _From5); the hotkeys go straight to their current names. 2.x's Explorer / append hotkeys are not used
        quick := data["Extensions"]["QuickSwitch"]
        quick["Hotkey"]         := pick(oldHotkeys, "TotalCMDDir", quick["Hotkey"])
        quick["AutoSwitch"]     := pick(config, "AutoSwitchDir", 0)
        quick["DialogWindows"]  := pick(config, "DialogWin", "ahk_class #32770")
        quick["ExcludeWindows"] := pick(config, "ExcludeWin", TextTools.JoinList(quick["ExcludeWindows"]))

        dateOptions := data["Extensions"]["AutoDate"]
        dateOptions["Hotkey"]        := pick(oldHotkeys, "AutoDateBEHKey", dateOptions["Hotkey"])
        dateOptions["RenameWindows"] := pick(oldHotkeys, "AutoDateBefExt", TextTools.JoinList(dateOptions["RenameWindows"]))
        dateOptions["AppendWindows"] := pick(oldHotkeys, "AutoDateAtEnd", TextTools.JoinList(dateOptions["AppendWindows"]))

        if (old.Has("PTTools") && old["PTTools"] is Map)
            data["Extensions"]["PTTools"] := old["PTTools"]

        data["Hotkeys"] := SchemaMigration._ConditionalHotkeys(oldHotkeys)

        commands := [], snippets := []
        userCommands := SchemaMigration._Section(old, "UserCommand")
        for cmdLine, rank in userCommands
            SchemaMigration._ConvertCommand(cmdLine, commands, snippets)
        if (userCommands.Count) {                                           ; 用户有自己的命令时, 不再塞默认示例
            data["CustomCommands"] := commands
            data["Snippets"] := snippets.Length ? snippets : data["Snippets"]
        }
        return data
    }

    ; "Type | Target | Description" -> CustomCommands / Snippets 条目
    static _ConvertCommand(cmdLine, commands, snippets) {
        parts := StrSplit(cmdLine, " | ")
        if (parts.Length < 2)
            return
        cmdType := Trim(parts[1])
        target  := Trim(parts[2])
        title   := parts.Length >= 3 ? Trim(parts[3]) : ""
        if (target = "")
            return

        if (cmdType = "Clip") {
            name := (title != "") ? title : SubStr(target, 1, 30)
            snippets.Push(Map("Name", name, "Keyword", StrLower(StrReplace(name, " ")), "Text", SchemaMigration._UnescapeClip(target)))
            return
        }

        typeMap := Map()
        typeMap.CaseSense := "Off"                                          ; 只能在 Map 为空时设置
        typeMap.Set("File", "File", "Dir", "Folder", "CMD", "Command", "URL", "Url", "App", "File")
        if !typeMap.Has(cmdType)
            return                                                          ; Func 等旧内置命令不再迁移

        newType := typeMap[cmdType]
        arguments := ""
        if (newType = "Command") {                                          ; "cmd.exe /k ipconfig" -> 程序 + 参数
            if RegExMatch(target, "^(`"[^`"]+`"|\S+)\s+(.+)$", &m) {
                target := Trim(m[1], "`"")
                arguments := m[2]
            }
        }
        if (title = "") {
            SplitPath(target, &fileName, , , &nameNoExt)
            title := (newType = "File" && nameNoExt != "") ? nameNoExt : (fileName != "") ? fileName : target
        }
        commands.Push(Map("Title", title, "Type", newType, "Target", target, "Arguments", arguments, "Keyword", ""))
    }

    ; 旧的 "条件热键" (CondTitle/CondHotkey/CondAction) -> Hotkeys 列表
    static _ConditionalHotkeys(oldHotkeys) {
        result := []
        key    := oldHotkeys.Has("CondHotkey") ? oldHotkeys["CondHotkey"] : ""
        action := oldHotkeys.Has("CondAction") ? oldHotkeys["CondAction"] : ""
        if (key = "" || key = "None" || action = "" || action = "Unset")
            return result
        winTitle := oldHotkeys.Has("CondTitle") ? oldHotkeys["CondTitle"] : ""
        result.Push(Map("Key", key, "Action", action, "WinTitle", winTitle))
        return result
    }

    static _Section(data, name) {
        return (data.Has(name) && data[name] is Map) ? data[name] : Map()
    }

    ; A comma-separated value of an old version -> its items (2.x wrote C:\Path\IndexLocation as a placeholder)
    static _SplitList(text) {
        list := []
        for item in TextTools.SplitList(text)
            if (item != "C:\Path\IndexLocation")
                list.Push(item)
        return list
    }

    static _CleanHotkey(key) {
        key := StrReplace(key, "~")
        return (key = "None") ? "" : key
    }

    ; 2.x 的 Clip 正文存成一行: \n = 换行, \t = Tab, \\ = 反斜杠
    static _UnescapeClip(text) {
        text := StrReplace(text, "\\", Chr(1))
        text := StrReplace(text, "\n", "`r`n")
        text := StrReplace(text, "\t", "`t")
        return StrReplace(text, Chr(1), "\")
    }
}
