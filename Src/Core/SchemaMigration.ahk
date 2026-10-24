;===============================================================================
; SchemaMigration.ahk - upgrades ALTRun.json between versions (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Each version's upgrade is a separate _FromN(data) method that turns version N data into
; version N+1. Upgrade() calls them one after another from the file's version up to CurrentVersion:
;   ALTRun.ini           the settings file of the released 2.x (up to v2026.08.12); ReadLegacyIni() reads it into
;                        the version 2 structure, which then goes through the upgrades below (see AppSettings.Load)
;   2 -> 3   _From2()   ALTRun 2.x (Config/Gui/Hotkey/UserCommand...) -> 3.0
;   3 -> 4   _From3()   file search no longer mixes into the normal results by default (InDefaultResults 1 -> 0)
;   4 -> 5   _From4()   snippet paste mode Clipboard -> Auto; windows / keywords added to the defaults later are added; small cleanups (see issue #176)
;   5 -> 6   _From5()   window lists become arrays; StandardDialogs split out of DialogWindows; two hotkeys renamed to
;                       Hotkey; unused keys removed; old ScopeExclude default and Compact themes replaced (#176)
;
; When the settings structure changes later: AppSettings.CurrentVersion + 1, and add a _FromN() here.
; Before upgrading, Backup() copies the original file to ALTRun.v<N>.backup.json.
;
; Kept from 2.x -> 3.0:
;   - general settings: show hotkey, start with Windows, tray, hide on focus loss, file manager
;   - user commands (UserCommand): File/Dir/CMD/URL -> CustomCommands, Clip -> Snippets
;   - index folders / types / depth, the structural calculation switch, snippet paste mode
;   - Listary-style folder jumps, Ctrl+D date stamp, PT Tools settings
;   - conditional hotkeys (e.g. middle button opens PT Tools in RAPT) -> Hotkeys
; Not kept: the built-in command list (DefaultCommand, replaced by the system commands feature), the old index
; (Index, rebuilt), run history, usage statistics and the old list-appearance options.
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

    ; Reads 2.x's ALTRun.ini -> Map("Config", Map(...), "Gui", ..., "Hotkey", ..., "UserCommand", ...),
    ; the version 2 structure _From2() expects. Only the sections that are converted are read; Index / History / Usage /
    ; DefaultCommand / FallbackCommand are no longer used.
    ;   normal sections: "key=value", split at the first =, numbers become integers
    ;   UserCommand: "type | target | description=rank", split at the last =; = and ; in keys were written
    ;                as _Equal_ / _Semicolon_ by 2.x and are restored here
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
    ; Only changes old users won't be surprised by (visible changes, e.g. the theme following the system or Alt+R, apply to new installs only; list in GitHub issue #176):
    ;   - snippet paste mode: Clipboard (the old default, almost always written automatically back then) -> Auto (faster, leaves the clipboard alone); Type stays
    ;   - windows / keywords added to the defaults later are added to old users' lists too (only what is missing; the user's own entries stay; emptied lists are left alone)
    ;   - a leftover "F1 -> About" hotkey (F1 is built in now) is removed; bar diameters are stored as numbers
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
        SchemaMigration._AddWindows(dateStamp, "RenameWindows", ["ahk_class TSTDTREEDLG", "ahk_class TCOMBOINPUT"])   ; Total Commander's dialogs / rename box
        if (dateStamp.Has("AppendWindows") && dateStamp["AppendWindows"] = "ahk_class TCmtEditForm,ahk_exe Notepad4.exe")
            dateStamp["AppendWindows"] := "ahk_class TCmtEditForm, ahk_exe Notepad4.exe"   ; the old default, only written the same way
        switcher := SchemaMigration._Section(extensions, "QuickSwitch")
        SchemaMigration._AddWindows(switcher, "ExcludeWindows", ["ahk_exe 7zG.exe", "ahk_exe TOTALCMD64.exe"])   ; 7-Zip extract, TC's dialogs

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

    ; Adds the windows not yet in the comma-separated window list section[key] (case-insensitive); a missing or emptied list is left alone
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
    ; Early 3.0 showed files matching on the whole disk in the normal results by default, mixing in unrelated files like NIRMALA.TTF;
    ; now it is off by default, and Space / ' / open / find search files on purpose. The 1 in old files was only the default
    ; written automatically back then, so it becomes the new default here (it can be switched on again in Preferences).
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
        general["KeepLastQuery"]        := pick(config, "KeepInput", 1)     ; 2.x kept the input by default
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
        if (userCommands.Count) {                                           ; users with their own commands don't get the default examples
            data["CustomCommands"] := commands
            data["Snippets"] := snippets.Length ? snippets : data["Snippets"]
        }
        return data
    }

    ; "Type | Target | Description" -> a CustomCommands / Snippets entry
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
        typeMap.CaseSense := "Off"                                          ; can only be set while the Map is empty
        typeMap.Set("File", "File", "Dir", "Folder", "CMD", "Command", "URL", "Url", "App", "File")
        if !typeMap.Has(cmdType)
            return                                                          ; Func and other old built-in commands aren't migrated

        newType := typeMap[cmdType]
        arguments := ""
        if (newType = "Command") {                                          ; "cmd.exe /k ipconfig" -> program + arguments
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

    ; the old "conditional hotkey" (CondTitle / CondHotkey / CondAction) -> the Hotkeys list
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

    ; 2.x stored a Clip's text on one line: \n = line break, \t = Tab, \\ = backslash
    static _UnescapeClip(text) {
        text := StrReplace(text, "\\", Chr(1))
        text := StrReplace(text, "\n", "`r`n")
        text := StrReplace(text, "\t", "`t")
        return StrReplace(text, Chr(1), "\")
    }
}
