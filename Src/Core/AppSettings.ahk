;===============================================================================
; AppSettings.ahk - reads and writes the user settings in Data\ALTRun.json, with their defaults (AutoHotkey v2)
;-------------------------------------------------------------------------------
; ALTRun.json only holds the user's settings and data (custom commands / snippets / web search engines /
; custom hotkeys). It and the data made at run time (application index, learned ranking, clipboard history) are all in
; the Data\ folder, each in a file of its own: deleting an index only makes it rebuild, the settings aren't affected.
; 3.x to v2026.09.26 kept ALTRun.json in the program folder; at start-up it is moved to Data\ (MoveLegacyFile).
;
; File structure (SchemaVersion 6):
; {
;   "SchemaVersion" : 5,
;   "General"       : { show hotkey / start with Windows / tray icon / hide on focus loss / language ... },
;   "Appearance"    : { theme / window width / visible rows },
;   "Features"      : { "<feature>": { "Enabled": 1, ...the feature's own settings } },
;   "Extensions"    : { "QuickSwitch": {...}, "AutoDate": {...}, "PTTools": {...} },
;   "Hotkeys"       : [ { "Key", "Action", "WinTitle" } ],   custom hotkeys -> system commands
;   "CustomCommands": [ { "Title", "Type", "Target", "Arguments", "Keyword" } ],
;   "Snippets"      : [ { "Name", "Keyword", "Text", "AutoExpand" } ],
;   "DataLocation"  : "D:\\OneDrive\\ALTRun"   optional, only in the file at the default location: use this data folder instead (see ResolveDataDir)
; }
;
; Old versions (the 2.x format without SchemaVersion) are upgraded by SchemaMigration; the original file is
; backed up first, see SchemaMigration.ahk.
;
; Usage:
;   AppSettings.Load()                       once at start-up
;   AppSettings.General["Hotkey"]            one value of a section
;   AppSettings.Feature("WebSearch")         a feature's settings (Map)
;   AppSettings.Save()                       writes Data\ALTRun.json
;===============================================================================

class AppSettings {
    static CurrentVersion := 6
    static Portable := Path.IsWritable(A_ScriptDir)                          ; the program folder is writable: data goes in the program folder (portable)
    static UserDir := A_AppData "\ALTRun"                                   ; used when the program folder isn't writable (e.g. installed in Program Files)
    static DataDir := AppSettings.ResolveDataDir()
    static File := AppSettings.DataDir "\ALTRun.json"
    static ThemesDir := (AppSettings.Portable ? A_ScriptDir : AppSettings.UserDir) "\Themes"   ; the user's own theme files (see ThemeManager)
    static LegacyFile := A_ScriptDir "\ALTRun.json"      ; the location of old versions; 2.x's ALTRun.ini is in this folder too
    static Data := Map()
    static MigratedFrom := 0          ; the version upgraded from at this start-up (0 = no upgrade)
    static ImportedFrom := ""         ; the old ALTRun.ini imported at this start-up ("" = none)
    static MovedFrom := ""            ; moved from the program folder to Data\ at this start-up ("" = no)

    static General        => AppSettings.Data["General"]
    static Appearance     => AppSettings.Data["Appearance"]
    static Hotkeys        => AppSettings.Data["Hotkeys"]
    static CustomCommands => AppSettings.Data["CustomCommands"]
    static Snippets       => AppSettings.Data["Snippets"]

    ; The data folder (settings, indexes, learned ranking, clipboard history...):
    ;   1. ALTRun.json at the default location has "DataLocation" (Preferences → Advanced → Change location): use that folder,
    ;      e.g. one in OneDrive so several PCs share the settings (%OneDrive%\ALTRun works; relative paths start from the program folder).
    ;      Only this entry at the default location counts; an ALTRun.json in that folder doesn't redirect again
    ;   2. the program folder is writable: <program folder>\Data (portable, the default)
    ;   3. otherwise: %APPDATA%\ALTRun\Data
    static ResolveDataDir() {
        if ((location := AppSettings.DataLocationIn(AppSettings.DefaultDataDir() "\ALTRun.json")) != "")
            return location
        return AppSettings.DefaultDataDir()
    }

    static DefaultDataDir() => AppSettings.Portable ? A_ScriptDir "\Data" : AppSettings.UserDir "\Data"

    ; The folder given by "DataLocation" in the settings file ("" when missing, empty or the default location). Called at start-up:
    ; it only looks for this entry without parsing the whole file (without a redirect the same file is fully read right after)
    static DataLocationIn(settingsFile) {
        if !FileExist(settingsFile)
            return ""
        try {
            if !RegExMatch(FileRead(settingsFile, "UTF-8"), '"DataLocation"\s*:\s*("(?:[^"\\]|\\.)*")', &m)
                return ""
            location := Trim(JSON.Parse("[" m[1] "]")[1])
            if (location != "" && (location := Path.Full(location)) != AppSettings.DefaultDataDir())
                return location
        }
        return ""
    }

    ; Changes the data folder to folder ("" or the default location = back to the default): written to ALTRun.json at the default location ("DataLocation"); takes effect after reloading
    static SetDataLocation(folder) {
        settingsFile := AppSettings.DefaultDataDir() "\ALTRun.json"
        reset := (folder = "" || Path.Full(folder) = AppSettings.DefaultDataDir())
        data := Map()
        if FileExist(settingsFile)
            data := JSON.Parse(FileRead(settingsFile, "UTF-8"))
        else if reset
            return
        if !reset
            data["DataLocation"] := folder
        else if data.Has("DataLocation")
            data.Delete("DataLocation")
        DirCreate(AppSettings.DefaultDataDir())
        JSON.WriteFile(settingsFile, data)
        if (AppSettings.DataDir = AppSettings.DefaultDataDir()) {           ; the settings at the default location are in use: keep this entry when saving later too
            if !reset
                AppSettings.Data["DataLocation"] := folder
            else if AppSettings.Data.Has("DataLocation")
                AppSettings.Data.Delete("DataLocation")
        }
    }

    ; Kept when resetting to the defaults (Preferences → Advanced): content the user added, not options
    static KeptOnReset := [["CustomCommands"], ["Snippets"], ["Hotkeys"], ["Features", "WebSearch", "Engines"]
        , ["Features", "Recent", "Pinned"], ["Features", "System", "Hidden"], ["Extensions", "PTTools"], ["DataLocation"]]

    ; current after a reset to the defaults: the defaults + the KeptOnReset parts of current
    static ResetData(current) {
        data := AppSettings.Defaults()
        for keys in AppSettings.KeptOnReset {
            source := current, target := data
            for index, key in keys {
                if !(source is Map && source.Has(key))
                    break
                if (index = keys.Length) {
                    target[key] := source[key]
                    break
                }
                source := source[key]
                if !(target.Has(key) && target[key] is Map)
                    target[key] := Map()
                target := target[key]
            }
        }
        return data
    }

    static Feature(name) {
        return AppSettings.Data["Features"][name]
    }

    static Extension(name) {
        return AppSettings.Data["Extensions"][name]
    }

    static Load() {
        AppSettings.MoveLegacyFile()
        data := AppSettings._ReadFile()
        changed := false

        ; The released 2.x (up to v2026.08.12) kept its settings in ALTRun.ini: imported from it while there is no ALTRun.json yet.
        ; ALTRun.ini stays unchanged, so going back to the old version still works
        if (!data.Count && !FileExist(AppSettings.File)) {
            iniFile := SchemaMigration.LegacyIniFile(AppSettings.LegacyFile)
            if FileExist(iniFile) {
                data := SchemaMigration.ReadLegacyIni(iniFile)
                AppSettings.ImportedFrom := iniFile
            }
        }

        if (data.Count) {
            fromVersion := SchemaMigration.DetectVersion(data)
            if (fromVersion < AppSettings.CurrentVersion) {
                SchemaMigration.Backup(AppSettings.File, fromVersion)
                data := SchemaMigration.Upgrade(data, fromVersion)
                AppSettings.MigratedFrom := fromVersion
                changed := true
            }
        } else {
            changed := true                                                 ; first run: write a file with the defaults
        }

        if (data.Count && AppSettings._KeepPrevious(data))
            changed := true
        if (data.Count && AppSettings._UpgradeEngines(data))
            changed := true
        if AppSettings._MergeDefaults(data, AppSettings.Defaults())
            changed := true
        if (AppSettings.DataDir != AppSettings.DefaultDataDir() && data.Has("DataLocation"))
            data.Delete("DataLocation")                                     ; this entry has no effect in the folder redirected to: remove it to avoid confusion
        data["SchemaVersion"] := AppSettings.CurrentVersion
        AppSettings.Data := data

        if (changed)
            AppSettings.Save()
    }

    static Save() {
        try {
            JSON.WriteFile(AppSettings.File, AppSettings.Data)
            return true
        } catch as e {
            Logger.Error("AppSettings.Save: " e.Message)
            MsgBox(I18n.T("Settings.SaveError", e.Message), App.Name, "Icon! 0x40000")   ; always on top: there is no window yet at start-up, and hidden behind another window it would look hung
            return false
        }
    }

    ; An old ALTRun.json in the program folder: while there is no Data\ALTRun.json, move it together with the upgrade backups (ALTRun.v*.backup.json)
    ; and the ALTRun.json.bad left by a failed parse. If both exist the one in Data\ is used and the old file is left alone.
    ; If it can't be moved (e.g. the file is in use), keep using the old location this time and try again at the next start-up
    static MoveLegacyFile() {
        legacy := AppSettings.LegacyFile, target := AppSettings.File
        if (legacy = target || !FileExist(legacy) || FileExist(target))
            return
        SplitPath(legacy, , &fromDir)
        SplitPath(target, , &toDir)
        try {
            DirCreate(toDir)
            FileMove(legacy, target)
        } catch as e {
            Logger.Error("AppSettings: cannot move " legacy " to " toDir " - " e.Message)
            AppSettings.File := legacy
            return
        }
        AppSettings.MovedFrom := legacy
        Loop Files, fromDir "\ALTRun.*", "F" {
            if (A_LoopFileName = "ALTRun.json.bad" || RegExMatch(A_LoopFileName, "i)^ALTRun\.v\d+\.backup\.json$"))
                if !FileExist(toDir "\" A_LoopFileName)
                    try FileMove(A_LoopFileFullPath, toDir "\" A_LoopFileName)
        }
        Logger.Debug("AppSettings: moved " legacy " to " target)
    }

    static _ReadFile() {
        if !FileExist(AppSettings.File)
            return Map()
        try {
            parsed := JSON.Parse(FileRead(AppSettings.File, "UTF-8"))
            return (parsed is Map) ? parsed : Map()
        } catch as e {
            badFile := AppSettings.File ".bad"
            try FileMove(AppSettings.File, badFile, true)
            catch as moveError                                              ; the defaults written next would replace the user's file
                Logger.Error("AppSettings: cannot keep the invalid file as " badFile " - " moveError.Message)
            Logger.Error("AppSettings: invalid JSON - " e.Message)
            MsgBox(I18n.T("Settings.ParseError", e.Message, badFile), App.Name, "Icon! 0x40000")
            return Map()
        }
    }

    ; Adds the keys that defaults has and target lacks (recursively; only Maps are filled in, arrays count as a whole).
    ; Returns whether anything changed. A wrong type (e.g. a Map edited into a string by hand) is replaced by the default too.
    ; New settings that an existing settings file doesn't have yet: keep the old behavior instead of the new install's default (see the two below). Returns whether one was added
    static _KeepPrevious(data) {
        keptDensity := AppSettings._KeepDensity(data)
        keptLargeType := AppSettings._KeepLargeType(data)                  ; not called largeType: that would shadow the LargeType class
        return keptDensity || keptLargeType
    }

    ; An existing settings file without Features.LargeType (before 2026.10.19.15): Large Type was always on, keep it on
    static _KeepLargeType(data) {
        if !(data.Has("Features") && data["Features"] is Map) || data["Features"].Has("LargeType")
            return false
        data["Features"]["LargeType"] := Map("Enabled", 1)
        return true
    }

    ; An existing settings file without Appearance.Density (before 2026.10.19.12): keep the old size instead of the new install's default "Normal".
    ; The old size = Comfortable; with one of the former compact themes (DarkCompact etc.) = Normal (the same size as them). Returns whether it was added
    static _KeepDensity(data) {
        if !(data.Has("Appearance") && data["Appearance"] is Map) || data["Appearance"].Has("Density")
            return false
        appearance := data["Appearance"]
        theme := appearance.Has("Theme") ? appearance["Theme"] : ""
        appearance["Density"] := RegExMatch(theme, "i)^(Dark|Light|Midnight)Compact$") ? "Normal" : "Comfortable"
        return true
    }

    ; Since 2026.10.19.18 the built-in search engines use {query|clipboard}: built-in engines (by Id) whose URL in an existing settings file is still
    ; the old default ({query}) get the new form; changed URLs are left alone. Returns whether anything changed
    static _UpgradeEngines(data) {
        if !(data.Has("Features") && data["Features"] is Map && data["Features"].Has("WebSearch") && data["Features"]["WebSearch"] is Map)
            return false
        webSearch := data["Features"]["WebSearch"]
        if !(webSearch.Has("Engines") && webSearch["Engines"] is Array)
            return false
        upgrades := Map()
        for engine in AppSettings._DefaultEngines()
            upgrades[engine["Id"]] := [StrReplace(engine["Url"], "{query|clipboard}", "{query}"), engine["Url"]]
        changed := false
        for engine in webSearch["Engines"] {
            if !(engine is Map && engine.Has("Id") && engine.Has("Url") && upgrades.Has(engine["Id"]))
                continue
            if (engine["Url"] == upgrades[engine["Id"]][1])
                engine["Url"] := upgrades[engine["Id"]][2], changed := true
        }
        return changed
    }

    static _MergeDefaults(target, defaults) {
        changed := false
        for key, defaultValue in defaults {
            if !target.Has(key) {
                target[key] := defaultValue
                changed := true
            } else if (defaultValue is Map) {
                if (target[key] is Map) {
                    if AppSettings._MergeDefaults(target[key], defaultValue)
                        changed := true
                } else {
                    target[key] := defaultValue
                    changed := true
                }
            } else if (defaultValue is Array && !(target[key] is Array)) {
                target[key] := defaultValue
                changed := true
            }
        }
        return changed
    }

    ; a new copy on each call: callers may change the result freely
    static Defaults() {
        return Map(
            "SchemaVersion", AppSettings.CurrentVersion,
            "General", Map(
                "Hotkey"              , "!Space",
                "SecondaryHotkey"     , "!r",                               ; Alt+R
                "DoubleTap"           , "",                                ; double-tap Ctrl / Shift shows the search window (as in Listary), "" = off
                "CapsLock"            , "",                                 ; CapsLock: Layout = switch input language, Mode = toggle Chinese / English mode, "" = as usual (see CapsLockSwitch)
                "CapsLockNotify"      , 1,                                  ; show a "Caps Lock: on / off" notice when Caps Lock is toggled
                "CapsLockElevated"    , 0,                                  ; also in administrator windows: a small helper that only handles CapsLock runs as administrator from a scheduled task (see CapsLockHelper)
                "SelectionHotkey"     , "^!\",                              ; after selecting text / files: open their actions directly (Ctrl+Alt+\)
                "HotkeySides"         , 0,                                  ; hotkeys recorded in Preferences keep the side of Ctrl / Alt / Shift / Win pressed (<!Space = left Alt + Space)
                "Language"            , "auto",                             ; auto or a language code in Resources\Lang (zh-CN / zh-TW / ja), en
                "LaunchAtLogin"       , 1,
                "ShowTrayIcon"        , 1,
                "HideOnDeactivate"    , 1,
                "SwitchToEnglishInput", 0,                                 ; switch to English when shown: 1 = English (US) keyboard, 2 = English mode of the current input method
                "RestoreInput"        , 1,                                 ; with SwitchToEnglishInput: switch back to the previous input language after the window hides
                "SpaceToRun"          , 0,                                  ; Space runs the selected result (Shift+Space types a space)
                "TabKey"              , "Complete",                         ; Tab: "Complete" = complete the name / browse into a folder; "Next" = next result (Shift+Tab previous, Ctrl+Tab completes)
                "KeepLastQuery"       , 0,                                  ; restore the last search when shown (text selected)
                "KeepLastQueryMinutes", 5,                                  ; ... only within this many minutes after hiding (0 = no limit)
                "ShowTips"            , 1,                                  ; rotate usage tips in the empty search box
                "FileManager"         , "explorer.exe",
                "SendToMenu"          , 1,
                "StartMenuShortcut"   , 1,
                "CheckForUpdates"     , 1,
                "SaveLog"             , 0,
                "HistorySize"         , 30
            ),
            "Appearance", Map(
                "Theme"      , "System",                                    ; a built-in theme (see ThemeManager; System = follow Windows dark / light) or Themes\<name>.json
                "Density"    , "Normal",                                    ; density (any theme): Comfortable (the theme's own size) / Normal / Compact, see ThemeManager.DensityScale
                "Width"      , 700,
                "VisibleRows", 8,
                "Placeholder", "",                                          ; gray text in the empty search box, empty = the default ("ALTRun Search")
                "StatusBar"  , 1,                                           ; status bar at the bottom of the search window (☰ menu, usage tips, the selected result's Enter action and →)
                "Fade"       , 0,                                           ; fade in / out when showing / hiding (see WindowFade)
                "ShowOn"     , "Mouse",                                     ; which screen: Mouse = under the mouse / Primary = primary screen / Active = the active window's
                "RememberPosition", 0,                                      ; remember the position after the window is dragged
                "Position"   , Map("X", 500, "Y", 200)                      ; the remembered position: per mille of the screen (see SearchWindow.Place)
            ),
            "Features", Map(
                "Applications", Map(
                    "Enabled"       , 1,
                    "Folders"       , ["A_Programs", "A_ProgramsCommon", "A_Desktop", "A_DesktopCommon"],
                    "FileTypes"     , ["*.lnk", "*.exe", "*.url", "*.appref-ms"],
                    "Depth"         , 3,
                    "Exclude"       , "i)(uninstall|卸载|readme|help|documentation)",
                    "Hidden"        , [],                                   ; paths of applications deleted from the results (Ctrl+Del)
                    "StoreApps"     , 1,
                    "MatchPinyin"   , 1
                ),
                "CustomCommands", Map("Enabled", 1),
                "Snippets", Map(
                    "Enabled"     , 1,
                    "Keyword"     , "snip",
                    "SearchText"  , 1,                                      ; also search the text (at least 3 characters)
                    "PasteMode"   , "Auto",                                 ; Auto = short single-line text is typed, the rest goes through the clipboard; Clipboard = clipboard + Ctrl+V; Type = typed character by character
                    "PasteDelay"  , 300,
                    "AutoExpand"  , 1,                                      ; typing prefix + keyword in any program expands it
                    "ExpandPrefix", ";",
                    "ExpandExclude", ["ahk_exe mstsc.exe", "ahk_exe KeePass.exe", "ahk_exe KeePassXC.exe"]   ; no expansion in these windows
                ),
                "Clipboard", Map(
                    "Enabled"      , 1,
                    "Keyword"      , "clip",
                    "Hotkey"       , "^!c",                                 ; opens the clipboard history directly
                    "MaxItems"     , 200,
                    "MaxItemLength", 100000,                                ; content longer than this many characters isn't recorded
                    "Persist"      , 1,                                     ; 0 = in memory only, cleared on exit
                    "Images"       , 1,                                     ; record images too (saved as Clipboard\*.png, needs Persist = 1)
                    "LocalFiles"   , 1,                                     ; where it is stored: 0 = all in Data; 1 = images and very long entries only on this PC (%LOCALAPPDATA%\ALTRun); 2 = the history too (see ClipboardProvider.PlaceStorage)
                    "MaxImages"    , 50,
                    "MergeDoubleCopy", 0,                                   ; press Ctrl+C twice quickly: append to the previous entry
                    "IgnoreApps"   , ["KeePass.exe", "KeePassXC.exe", "1Password.exe", "Bitwarden.exe"]
                ),
                "Calculator", Map(
                    "Enabled"       , 1,
                    "StructuralCalc", 0,                                    ; beam main bars / steel area below the result
                    "RebarCover"    , 40,                                   ; cover (mm, each side)
                    "MaxBarSpacing" , 300,                                  ; maximum spacing of main bars (mm)
                    "BarSizes"      , [13, 16, 20, 25, 32],                 ; bar diameters listed in the steel area row (mm)
                    "BarPrefix"     , "H",                                  ; bar symbol prefix (H / T / Y / Φ...)
                    "Currency"      , 0                                     ; currency conversion (100 usd to sgd): exchange rates downloaded daily from frankfurter.dev, off by default
                ),
                "WebSearch", Map(
                    "Enabled"  , 1,
                    "Engines"  , AppSettings._DefaultEngines(),
                    "Fallbacks", ["google", "files", "bing"]                ; fallbacks shown when there are no results at all (engine Id or "files")
                ),
                "FileSearch", Map(
                    "Enabled"            , 1,
                    "Keywords"           , ["open", "find", "file"],
                    "FolderKeywords"     , ["folder"],                      ; "folder budget" searches folders only
                    "RecentKeywords"     , ["recent"],                      ; "recent budget" searches Windows' recent files and folders
                    "TypeFilters"        , [                                ; "doc report" searches only files with these extensions; one "keyword = ext ..." per line
                        "doc = doc docx xls xlsx xlsm csv ppt pptx pdf txt md rtf odt ods",
                        "pic = jpg jpeg png gif bmp tif tiff webp svg ico heic",
                        "video = mp4 mkv avi mov wmv flv webm m4v",
                        "audio = mp3 wav flac aac m4a ogg wma",
                        "zip = zip rar 7z tar gz tgz bz2 xz iso",
                        "exe = exe msi bat cmd ps1 lnk",
                        "cad = dwg dxf dgn dwf rvt ifc skp"
                    ],
                    "SpacePrefix"        , 1,                               ; Space first in the empty search box = search files only, as in Alfred
                    "QuotePrefix"        , 1,                               ; starting with ' searches files directly, as in Alfred
                    "MaxResults"         , 30,
                    "InDefaultResults"   , 0,                               ; also show matching files / folders when typing a name (off by default, so unrelated files from the whole disk don't mix in)
                    "DefaultResultsLimit", 6,
                    "DefaultResultsMatch", "Start",                         ; files in normal results: Start = the name or a word in it starts with the text, Contains = anywhere in the name
                    "MinQueryLength"     , 2,
                    "UseEverything"      , 1,                               ; while Everything runs, query the whole disk over IPC
                    "EverythingFilter"   , "!C:\Windows\ !\AppData\ !\$Recycle.Bin\",
                    "EverythingPath"     , "",                              ; where Everything.exe is; empty = find it (only used for "Search in Everything")
                    "ScopeFolders"       , ["A_Desktop", "A_MyDocuments", "%UserProfile%\Downloads"],   ; what the built-in index covers when there is no Everything
                    "ScopeDepth"         , 4,
                    "ScopeExclude"       , "i)\\(node_modules|\.git|\.svn|\.hg|__pycache__|\.venv|venv|\.tox|\.cache|\.idea|\.vs|\$RECYCLE\.BIN)(\\|$)",   ; code repositories, virtual environments, IDE caches: often tens of thousands of files that would fill MaxEntries
                    "MaxEntries"         , 30000,
                    "RefreshMinutes"     , 60
                ),
                "Terminal", Map(
                    "Enabled", 1,
                    "Prefix" , ">",
                    "Shell"  , "cmd"                                        ; cmd / powershell / pwsh / wt
                ),
                "Help", Map("Enabled", 1),                                  ; typing ? shows every input syntax and shortcut
                "LargeType", Map("Enabled", 0),                             ; "Large Type" with Ctrl+L / in the actions panel (off for new installs; older settings files keep it on)
                "ContextMenu", Map("Enabled", 0),                           ; "Windows Context Menu" in the actions panel of files and folders
                "Bookmarks", Map(
                    "Enabled"         , 1,
                    "Keyword"         , "bm",                               ; "bm word" searches bookmarks only
                    "InDefaultResults", 1                                   ; also show matching bookmarks when typing a name (after applications and commands)
                ),
                "Scripts", Map(
                    "Enabled"    , 1,                                       ; scripts in the Scripts\ folder become commands
                    "Keyword"    , "script"                                 ; "script backup" searches scripts only, "script " lists them all
                ),
                "Recent", Map(
                    "Enabled"    , 1,
                    "RecentCount", 0,                                       ; how many recently opened items the empty search box shows (0 = none; by default only the pinned ones)
                    "Pinned"     , []                                       ; items pinned to the empty search box ("Pin" in the actions panel)
                ),
                "Windows", Map(
                    "Enabled"         , 1,
                    "Keyword"         , "w",                                ; "w title" searches open windows only, "w " lists them all
                    "InDefaultResults", 1                                   ; also show windows whose title matches well when typing (at most 3)
                ),
                "System", Map(
                    "Enabled"       , 1,
                    "ConfirmActions", 1,                                    ; confirm before shutdown / restart / sign out / emptying the Recycle Bin
                    "SettingsPages" , 1,                                    ; search the pages of Windows Settings (display, Bluetooth, WLAN...)
                    "Hidden"        , []                                    ; Ids of system commands deleted from the results (Ctrl+Del)
                )
            ),
            "Extensions", Map(
                "QuickSwitch", Map(
                    "Enabled"          , 1,
                    "Hotkey"           , "^g",                              ; the jump hotkey: folder of the TC / Opus / Explorer window used last
                    "MenuHotkey"       , "^+g",                             ; folder menu: every TC panel, Explorer window and recent folder
                    "RecentFolders"    , 10,                                ; how many recent folders the menu lists, 0 = none
                    "PanelSearch"      , "folders",                         ; panel search box: folders = folders only, all = folders and files
                    "ShowPanel"        , 1,                                 ; show the folder panel below a dialog as soon as it appears (as in Listary); one click jumps there
                    "AutoSwitch"       , 0,
                    "StandardDialogs"  , 1,                                 ; the standard Open / Save dialogs (ahk_class #32770)
                    "DialogWindows"    , [],                                ; dialogs programs draw themselves, e.g. "ahk_class Qt5QWindowIcon" (WPS)
                    "BrowseDialogs"    , 1,                                 ; also the old "Browse for Folder" dialog (folder tree only, SHBrowseForFolder)
                    "ExcludeWindows"   , ["ahk_class SysListView32", "ahk_exe Explorer.exe", "ahk_exe 7zG.exe", "ahk_exe TOTALCMD64.exe"],   ; 7-Zip's extract and Total Commander's dialogs are not Open / Save dialogs
                    "AutoSwitchExclude", []                                 ; no automatic jump in these dialogs (the hotkey still works)
                ),
                "AutoDate", Map(
                    "Enabled"      , 1,
                    "DateFormat"   , "dd.MM.yyyy",
                    "Hotkey"       , "^d",                                  ; one hotkey: the window decides between renaming and appending
                    "RenameWindows", ["ahk_class CabinetWClass", "ahk_class Progman", "ahk_class WorkerW", "ahk_class #32770", "ahk_class TTOTAL_CMD", "ahk_class TSTDTREEDLG", "ahk_class TCOMBOINPUT"],   ; TC: main window, dialogs, rename box
                    "AppendWindows", ["ahk_class TCmtEditForm", "ahk_exe Notepad4.exe"]
                ),
                "PTTools", Map()                                            ; PTToolsWindow fills in its own defaults
            ),
            "Hotkeys", [                                                    ; examples: global, only in one program, only in ALTRun
                Map("Key", "^!#v", "Action", "PastePlain", "WinTitle", ""),                         ; same key as PowerToys "Paste as plain text"
                Map("Key", "^!t", "Action", "TerminalHere", "WinTitle", "ahk_class CabinetWClass"), ; File Explorer
                Map("Key", "^r", "Action", "RebuildIndex", "WinTitle", "ALTRun")
            ],
            "CustomCommands", AppSettings._DefaultCommands(),
            "Snippets", AppSettings._DefaultSnippets()
        )
    }

    ; The default custom commands: one of each type, so new users see what they can do (existing users aren't affected)
    static _DefaultCommands() {
        commands := []
        for row in [
            ; file / program: runs the program or opens the file; Arguments are passed along
            ["Notepad"             , "File"   , "notepad.exe"            , ""                                       , "np"],
            ["Edit hosts file"     , "File"   , "notepad.exe"            , "C:\Windows\System32\drivers\etc\hosts"  , "hosts"],
            ["Paint"               , "File"   , "mspaint.exe"            , ""                                       , ""],
            ; folder: opens in the file manager; built-in variables such as A_Desktop or environment variables such as %UserProfile% work
            ["Desktop"             , "Folder" , "A_Desktop"              , ""                                       , ""],
            ["Downloads"           , "Folder" , "%UserProfile%\Downloads", ""                                       , "dl"],
            ["ALTRun Folder"       , "Folder" , "A_ScriptDir"            , ""                                       , ""],
            ["Documents subfolder" , "Folder" , "%UserProfile%\Documents\{query}", ""                               , "docs"],   ; {query}: "docs <folder name>"
            ; command line: program + arguments
            ["IP Configuration"    , "Command", "cmd.exe"                , "/k ipconfig /all"                       , "ipconfig"],
            ["Ping"                , "Command", "cmd.exe"                , "/k ping {query}"                        , "ping"],     ; {query}: "ping 10.0.0.1"
            ; URL: a web page, or a system link such as ms-settings:
            ["ALTRun on GitHub"    , "Url"    , "https://github.com/zhugecaomao/ALTRun", ""                     , "altrun"],
            ["Google Maps"         , "Url"    , "https://www.google.com/maps/search/{query}", ""                 , "maps"],     ; {query}: "maps Changi Airport"
            ["Windows Update"      , "Url"    , "ms-settings:windowsupdate", ""                                  , ""]
        ]
            commands.Push(Map("Title", row[1], "Type", row[2], "Target", row[3], "Arguments", row[4], "Keyword", row[5]))
        return commands
    }

    ; The default snippets: dates, the clipboard and a few common email lines (type ;keyword to expand, or search the name in the search window)
    static _DefaultSnippets() {
        snippets := []
        for row in [
            ["Today's date"             , "today", "{date}"],
            ["Date and time"            , "now"  , "{datetime}"],
            ["Paste as plain text"      , "plain", "{clipboard}"],
            ["Email: thank you"         , "thx"  , "Hi {cursor},`r`n`r`nThank you for your email. I will look into it and get back to you shortly.`r`n`r`nBest regards,"],
            ["Email: please find attached", "pfa", "Hi,`r`n`r`nPlease find attached {cursor} for your reference.`r`n`r`nBest regards,"],
            ["Email: meeting request"   , "meet" , "Hi,`r`n`r`nWould you be available for a short meeting on {cursor}? Please let me know a time that suits you.`r`n`r`nThanks,"],
            ["Email signature"          , "sig"  , "Best regards,`r`n`r`nYour Name`r`nJob Title | Company`r`nPhone: +1 234 567 890"],
            ["My address"               , "addr" , "123 Example Street`r`nCity, Country 12345"]
        ]
            snippets.Push(Map("Name", row[1], "Keyword", row[2], "Text", row[3], "AutoExpand", 1))
        return snippets
    }

    static _DefaultEngines() {
        engines := []
        for row in [
            ["google"   , "g"   , "Google"   , "https://www.google.com/search?q={query|clipboard}"],
            ["bing"     , "bing", "Bing"     , "https://www.bing.com/search?q={query|clipboard}"],
            ["baidu"    , "bd"  , "Baidu"    , "https://www.baidu.com/s?wd={query|clipboard}"],
            ["github"   , "gh"  , "GitHub"   , "https://github.com/search?q={query|clipboard}"],
            ["wikipedia", "wiki", "Wikipedia", "https://en.wikipedia.org/wiki/Special:Search?search={query|clipboard}"],
            ["youtube"  , "yt"  , "YouTube"  , "https://www.youtube.com/results?search_query={query|clipboard}"],
            ["taobao"   , "tb"  , "Taobao"   , "https://s.taobao.com/search?q={query|clipboard}"],
            ["jd"       , "jd"  , "JD.com"   , "https://search.jd.com/Search?keyword={query|clipboard}&enc=utf-8"],
            ["translate", "tr"  , "Google Translate", "https://translate.google.com/?sl=auto&tl=zh-CN&text={query|clipboard}"]
        ]
            engines.Push(Map("Id", row[1], "Keyword", row[2], "Title", row[3], "Url", row[4]))
        return engines
    }
}
