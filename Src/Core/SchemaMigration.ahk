;===============================================================================
; SchemaMigration.ahk - ALTRun.json 版本升级 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 每个版本的升级是一个独立的 _FromN(data) 方法, 负责把版本 N 的数据转换成
; 版本 N+1。Upgrade() 从文件当前的版本开始逐级调用, 直到 CurrentVersion:
;   2 -> 3   _From2()   ALTRun 2.x (Config/Gui/Hotkey/UserCommand...) -> 3.0
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

    static BackupFile(file, version) {
        SplitPath(file, , &dir)
        return dir "\ALTRun.v" version ".backup.json"
    }

    static Backup(file, version) {
        if !FileExist(file)
            return ""
        backup := SchemaMigration.BackupFile(file, version)
        try FileCopy(file, backup, true)
        return backup
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
        general["Language"]             := pick(config, "Chinese", 0) ? "zh" : "auto"
        general["LaunchAtLogin"]        := pick(config, "AutoStartup", 1)
        general["ShowTrayIcon"]         := pick(config, "ShowTrayIcon", 1)
        general["HideOnDeactivate"]     := pick(config, "HideOnLostFocus", 1)
        general["SwitchToEnglishInput"] := pick(config, "AutoEngIME", 0)
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

        everything := pick(config, "Everything", "")
        if (everything != "" && everything != "C:\Apps\Everything.exe")
            data["Features"]["FileSearch"]["EverythingPath"] := everything

        quick := data["Extensions"]["QuickSwitch"]
        quick["ExplorerHotkey"] := pick(oldHotkeys, "ExplorerDir", quick["ExplorerHotkey"])
        quick["TotalCmdHotkey"] := pick(oldHotkeys, "TotalCMDDir", quick["TotalCmdHotkey"])
        quick["AutoSwitch"]     := pick(config, "AutoSwitchDir", 0)
        quick["DialogWindows"]  := pick(config, "DialogWin", quick["DialogWindows"])
        quick["ExcludeWindows"] := pick(config, "ExcludeWin", quick["ExcludeWindows"])

        dateOptions := data["Extensions"]["AutoDate"]
        dateOptions["RenameHotkey"]  := pick(oldHotkeys, "AutoDateBEHKey", dateOptions["RenameHotkey"])
        dateOptions["RenameWindows"] := pick(oldHotkeys, "AutoDateBefExt", dateOptions["RenameWindows"])
        dateOptions["AppendHotkey"]  := pick(oldHotkeys, "AutoDateAEHKey", dateOptions["AppendHotkey"])
        dateOptions["AppendWindows"] := pick(oldHotkeys, "AutoDateAtEnd", dateOptions["AppendWindows"])

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

    static _SplitList(text) {
        list := []
        for item in StrSplit(text, ",")
            if (Trim(item) != "" && Trim(item) != "C:\Path\IndexLocation")
                list.Push(Trim(item))
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
