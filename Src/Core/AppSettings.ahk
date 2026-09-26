;===============================================================================
; AppSettings.ahk - 用户设置 Data\ALTRun.json 的读写和默认值 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; ALTRun.json 只保存用户设置和用户数据 (自定义命令 / 片段 / 网页搜索引擎 /
; 自定义热键)。它和运行时生成的数据 (应用索引、学习记录、剪贴板历史) 都放在
; Data\ 目录里, 各是一个单独的文件: 删掉索引等文件只会让它们重新生成, 不影响设置。
; 3.x 到 v2026.09.26 把 ALTRun.json 放在程序目录, 启动时自动移到 Data\ (MoveLegacyFile)。
;
; 文件结构 (SchemaVersion 3):
; {
;   "SchemaVersion" : 3,
;   "General"       : { 呼出热键 / 开机启动 / 托盘图标 / 失焦隐藏 / 语言 ... },
;   "Appearance"    : { 主题 / 窗口宽度 / 可见行数 },
;   "Features"      : { "<功能名>": { "Enabled": 1, ...该功能自己的设置 } },
;   "Extensions"    : { "QuickSwitch": {...}, "AutoDate": {...}, "PTTools": {...} },
;   "Hotkeys"       : [ { "Key", "Action", "WinTitle" } ],   自定义热键 -> 系统命令
;   "CustomCommands": [ { "Title", "Type", "Target", "Arguments", "Keyword" } ],
;   "Snippets"      : [ { "Name", "Keyword", "Text", "AutoExpand" } ]
; }
;
; 旧版本 (没有 SchemaVersion 的 2.x 格式) 由 SchemaMigration 自动升级, 升级前
; 会先把原文件备份一份, 见 SchemaMigration.ahk。
;
; 用法:
;   AppSettings.Load()                       启动时调用一次
;   AppSettings.General["Hotkey"]            读某一节的某个值
;   AppSettings.Feature("WebSearch")         某个功能的设置 (Map)
;   AppSettings.Save()                       写回 Data\ALTRun.json
;===============================================================================

class AppSettings {
    static CurrentVersion := 4
    static DataDir := A_ScriptDir "\Data"
    static File := A_ScriptDir "\Data\ALTRun.json"
    static LegacyFile := A_ScriptDir "\ALTRun.json"      ; 旧版本的位置, 2.x 的 ALTRun.ini 也在这个目录
    static Data := Map()
    static MigratedFrom := 0          ; 本次启动时从哪个版本升级过来的 (0 = 没有升级)
    static ImportedFrom := ""         ; 本次启动时从旧版本的 ALTRun.ini 导入 ("" = 没有)
    static MovedFrom := ""            ; 本次启动时从程序目录移到 Data\ ("" = 没有)

    static General        => AppSettings.Data["General"]
    static Appearance     => AppSettings.Data["Appearance"]
    static Hotkeys        => AppSettings.Data["Hotkeys"]
    static CustomCommands => AppSettings.Data["CustomCommands"]
    static Snippets       => AppSettings.Data["Snippets"]

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

        ; 已发布的 2.x (到 v2026.08.12) 把设置存在 ALTRun.ini: 还没有 ALTRun.json 时从它导入。
        ; ALTRun.ini 保持不变, 退回旧版本时仍然可以用
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
            changed := true                                                 ; 第一次运行: 写出一份带默认值的文件
        }

        if AppSettings._MergeDefaults(data, AppSettings.Defaults())
            changed := true
        data["SchemaVersion"] := AppSettings.CurrentVersion
        AppSettings.Data := data

        if (changed)
            AppSettings.Save()
    }

    static Save() {
        tmpFile := AppSettings.File ".tmp"
        try {
            SplitPath(AppSettings.File, , &dir)
            DirCreate(dir)
            if FileExist(tmpFile)
                FileDelete(tmpFile)
            FileAppend(JSON.Stringify(AppSettings.Data), tmpFile, "UTF-8")  ; 先写临时文件, 写一半崩溃也不会弄坏正式文件
            FileMove(tmpFile, AppSettings.File, true)
            return true
        } catch as e {
            Logger.Error("AppSettings.Save: " e.Message)
            MsgBox(I18n.T("Settings.SaveError", e.Message), App.Name, 48)
            return false
        }
    }

    ; 旧版本的 ALTRun.json 在程序目录: 还没有 Data\ALTRun.json 时连同升级备份 (ALTRun.v*.backup.json)
    ; 和解析失败留下的 ALTRun.json.bad 一起移过去。两边都有时用 Data\ 里的, 旧文件不动。
    ; 移不动 (例如文件被占用) 时这次继续用旧位置, 下次启动再试
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
            Logger.Error("AppSettings: invalid JSON - " e.Message)
            MsgBox(I18n.T("Settings.ParseError", e.Message, badFile), App.Name, 48)
            return Map()
        }
    }

    ; 把 defaults 里有、target 里没有的键补上 (逐层递归, 只补 Map, 数组整体看待)。
    ; 返回是否有改动。类型不一致 (比如该是 Map 的地方被手改成了字符串) 也用默认值覆盖。
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

    ; 每次调用都新建一份, 调用方可以随意修改返回值
    static Defaults() {
        return Map(
            "SchemaVersion", AppSettings.CurrentVersion,
            "General", Map(
                "Hotkey"              , "!Space",
                "SecondaryHotkey"     , "",
                "Language"            , "auto",                             ; auto / en / zh
                "LaunchAtLogin"       , 1,
                "ShowTrayIcon"        , 1,
                "HideOnDeactivate"    , 1,
                "SwitchToEnglishInput", 0,
                "SpaceToRun"          , 0,                                  ; 空格执行选中项 (Shift+空格输入空格)
                "KeepLastQuery"       , 0,                                  ; 呼出窗口时恢复上一次的搜索 (文字全选)
                "ShowTips"            , 1,                                  ; 空搜索框里轮换显示使用提示
                "FileManager"         , "explorer.exe",
                "SendToMenu"          , 1,
                "StartMenuShortcut"   , 1,
                "CheckForUpdates"     , 1,
                "SaveLog"             , 0,
                "HistorySize"         , 30
            ),
            "Appearance", Map(
                "Theme"      , "Light",                                     ; 内置主题名 (见 ThemeManager) 或 Themes\<名称>.json
                "Width"      , 700,
                "VisibleRows", 8,
                "ShowOn"     , "Mouse",                                     ; 显示在哪块屏幕: Mouse 鼠标所在 / Primary 主屏幕 / Active 当前窗口所在
                "RememberPosition", 0,                                      ; 拖动窗口后记住位置
                "Position"   , Map("X", 500, "Y", 200)                      ; 记住的位置: 屏幕里的千分比 (见 SearchWindow.Place)
            ),
            "Features", Map(
                "Applications", Map(
                    "Enabled"       , 1,
                    "Folders"       , ["A_Programs", "A_ProgramsCommon", "A_Desktop", "A_DesktopCommon"],
                    "FileTypes"     , ["*.lnk", "*.exe", "*.url", "*.appref-ms"],
                    "Depth"         , 3,
                    "Exclude"       , "i)(uninstall|卸载|readme|help|documentation)",
                    "Hidden"        , [],                                   ; 在搜索结果里删除 (Ctrl+Del) 的应用的路径
                    "StoreApps"     , 1,
                    "MatchPinyin"   , 1,
                    "RefreshMinutes", 60
                ),
                "CustomCommands", Map("Enabled", 1),
                "Snippets", Map(
                    "Enabled"     , 1,
                    "Keyword"     , "snip",
                    "PasteMode"   , "Clipboard",                            ; Clipboard = 剪贴板 + Ctrl+V; Type = 逐字输入
                    "PasteDelay"  , 300,
                    "AutoExpand"  , 1,                                      ; 在任何程序里输入 前缀+关键字 自动展开
                    "ExpandPrefix", ";"
                ),
                "Clipboard", Map(
                    "Enabled"      , 1,
                    "Keyword"      , "clip",
                    "Hotkey"       , "^!c",                                 ; 直接打开剪贴板历史
                    "MaxItems"     , 200,
                    "MaxItemLength", 100000,                                ; 超过这么多字的内容不记录
                    "Persist"      , 1,                                     ; 0 = 只在内存里, 退出即清空
                    "IgnoreApps"   , ["KeePass.exe", "KeePassXC.exe", "1Password.exe", "Bitwarden.exe"]
                ),
                "Calculator", Map(
                    "Enabled"       , 1,
                    "StructuralCalc", 0                                     ; 结果下方附带梁主筋 / 配筋面积计算
                ),
                "WebSearch", Map(
                    "Enabled"  , 1,
                    "Engines"  , AppSettings._DefaultEngines(),
                    "Fallbacks", ["google", "files", "bing"]                ; 没有任何结果时显示的兜底项 (引擎 Id 或 "files")
                ),
                "FileSearch", Map(
                    "Enabled"            , 1,
                    "Keywords"           , ["open", "find"],
                    "FolderKeywords"     , ["folder"],                      ; "folder bk" 只搜文件夹
                    "SpacePrefix"        , 1,                               ; 空的搜索框里先按空格 = 只搜文件, 和 Alfred 一样
                    "QuotePrefix"        , 1,                               ; 以 ' 开头直接搜索文件, 和 Alfred 一样
                    "MaxResults"         , 30,
                    "InDefaultResults"   , 0,                               ; 直接输入名称时也显示匹配的文件 / 文件夹 (默认关闭, 避免全盘的无关文件混进来)
                    "DefaultResultsLimit", 6,
                    "MinQueryLength"     , 2,
                    "UseEverything"      , 1,                               ; Everything 在运行时通过 IPC 查询全盘
                    "EverythingFilter"   , "!C:\Windows\ !\AppData\ !\$Recycle.Bin\",
                    "EverythingPath"     , "",                              ; Everything.exe 的位置, 留空自动查找 (只用于 "在 Everything 中搜索")
                    "ScopeFolders"       , ["A_Desktop", "A_MyDocuments", "%UserProfile%\Downloads"],   ; 没有 Everything 时的内置索引范围
                    "ScopeDepth"         , 4,
                    "ScopeExclude"       , "i)\\(node_modules|\.git|__pycache__|\$RECYCLE\.BIN)(\\|$)",
                    "MaxEntries"         , 30000,
                    "RefreshMinutes"     , 30
                ),
                "Terminal", Map(
                    "Enabled", 1,
                    "Prefix" , ">",
                    "Shell"  , "cmd"                                        ; cmd / powershell / pwsh / wt
                ),
                "Help", Map("Enabled", 1),                                  ; 输入 ? 显示所有输入语法和快捷键
                "System", Map(
                    "Enabled"       , 1,
                    "ConfirmActions", 1                                     ; 关机/重启/注销/清空回收站前确认
                )
            ),
            "Extensions", Map(
                "QuickSwitch", Map(
                    "Enabled"          , 1,
                    "ExplorerHotkey"   , "^e",
                    "TotalCmdHotkey"   , "^g",
                    "AutoSwitch"       , 0,
                    "DialogWindows"    , "ahk_class #32770",
                    "ExcludeWindows"   , "ahk_class SysListView32, ahk_exe Explorer.exe"
                ),
                "AutoDate", Map(
                    "Enabled"      , 1,
                    "DateFormat"   , "dd.MM.yyyy",
                    "RenameHotkey" , "^d",
                    "RenameWindows", "ahk_class CabinetWClass,ahk_class Progman,ahk_class WorkerW,ahk_class #32770,ahk_class TTOTAL_CMD",
                    "AppendHotkey" , "^d",
                    "AppendWindows", "ahk_class TCmtEditForm,ahk_exe Notepad4.exe"
                ),
                "PTTools", Map()                                            ; 由 PTToolsWindow 自己补默认值
            ),
            "Hotkeys", [
                Map("Key", "~MButton", "Action", "PTTools", "WinTitle", "ahk_exe RAPTW.exe")
            ],
            "CustomCommands", [
                Map("Title", "Desktop", "Type", "Folder", "Target", "A_Desktop", "Arguments", "", "Keyword", ""),
                Map("Title", "ALTRun Folder", "Type", "Folder", "Target", "A_ScriptDir", "Arguments", "", "Keyword", ""),
                Map("Title", "IP Configuration", "Type", "Command", "Target", "cmd.exe", "Arguments", "/k ipconfig /all", "Keyword", "ipconfig")
            ],
            "Snippets", [
                Map("Name", "Today's date", "Keyword", "today", "Text", "{date}")
            ]
        )
    }

    static _DefaultEngines() {
        engines := []
        for row in [
            ["google"   , "g"   , "Google"   , "https://www.google.com/search?q={query}"],
            ["bing"     , "bing", "Bing"     , "https://www.bing.com/search?q={query}"],
            ["baidu"    , "bd"  , "Baidu"    , "https://www.baidu.com/s?wd={query}"],
            ["github"   , "gh"  , "GitHub"   , "https://github.com/search?q={query}"],
            ["wikipedia", "wiki", "Wikipedia", "https://en.wikipedia.org/wiki/Special:Search?search={query}"],
            ["youtube"  , "yt"  , "YouTube"  , "https://www.youtube.com/results?search_query={query}"],
            ["taobao"   , "tb"  , "Taobao"   , "https://s.taobao.com/search?q={query}"],
            ["jd"       , "jd"  , "JD.com"   , "https://search.jd.com/Search?keyword={query}&enc=utf-8"],
            ["translate", "tr"  , "Google Translate", "https://translate.google.com/?sl=auto&tl=zh-CN&text={query}"]
        ]
            engines.Push(Map("Id", row[1], "Keyword", row[2], "Title", row[3], "Url", row[4]))
        return engines
    }
}
