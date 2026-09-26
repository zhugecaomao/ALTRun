;===============================================================================
; ThemeManager.ahk - 搜索窗口的配色和尺寸 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 主题从哪里来:
;   Light                     写在代码里 (Defaults), 所有主题都在它的基础上修改;
;                             主题文件全部丢失时也能正常显示
;   Resources\Themes\*.json   内置主题 (参照 Alfred 自带的几套), 随程序发布, 升级时整体替换:
;                             Dark / DarkCompact / Classic / Midnight / Frost / Graphite / Ocean / Paper
;   Themes\*.json             用户自己的主题, 升级不动; 和内置主题同名时用户的优先
;   System                    跟随 Windows 的浅色 / 深色设置 (Light 或 Dark), 系统切换时自动更新
;
; 主题文件只写要改的键; "Base" 指定从哪个主题开始 (默认 Light), 例如 Themes\MyBlue.json:
;   { "Base": "Dark", "SelectedBackground": "1D4ED8", "Opacity": 240 }
; 用户主题和内置主题同名并且 "Base" 写自己 (例如 Themes\Dark.json 里 "Base": "Dark"),
; 表示在内置的那一套上修改。偏好设置 -> 外观 -> "复制为自定义主题" 会生成一份完整的文件。
;
; 可用的键 (颜色一律 "RRGGBB"; 字号单位 pt; 尺寸是 96 DPI 下的像素, 会按 DPI 缩放):
;   FontName  InputFontSize  TitleFontSize  SubtitleFontSize  ShortcutFontSize
;   Padding  RowHeight  IconSize
;   Background  Border  InputText  Separator  Title  Subtitle  Shortcut
;   SelectedBackground  SelectedTitle  SelectedSubtitle  SelectedShortcut
;   SelectedRadius   选中行的圆角半径, 0 = 整行直角选中条
;   Opacity          窗口不透明度 1~255, 255 = 不透明
;
; 用法:
;   ThemeManager.Load("Dark")
;   ThemeManager.Get("Background")       -> "1E1F22"
;   ThemeManager.FontName()              -> 实际使用的字体
;   ThemeManager.Names()                 -> System, Light, 内置主题, 用户主题
;   ThemeManager.Resolve("Ocean")        -> 完整的键值 Map (找不到返回 "")
;===============================================================================

class ThemeManager {
    static Current  := Map()
    static Name     := "Light"                      ; 设置里选的主题
    static Resolved := "Light"                      ; 实际使用的主题 (System -> Light / Dark)
    static BuiltinDir := A_ScriptDir "\Resources\Themes"
    static UserDir    := A_ScriptDir "\Themes"
    static BuiltinOrder := ["Dark", "DarkCompact", "Classic", "Midnight", "Frost", "Graphite", "Ocean", "Paper"]   ; 列表里的顺序
    static _listening := false

    static Load(themeName) {
        ThemeManager.Name := themeName
        if (themeName = "System") {
            themeName := ThemeManager.SystemUsesDark() ? "Dark" : "Light"
            ThemeManager._WatchSystemTheme()
        }
        theme := ThemeManager.Resolve(themeName)
        if !IsObject(theme) {
            Logger.Error("ThemeManager: theme '" themeName "' not found, using Light")
            themeName := "Light"
            theme := ThemeManager.Defaults()
        }
        ThemeManager.Current := theme
        ThemeManager.Resolved := themeName
    }

    static Get(key) {
        return ThemeManager.Current.Has(key) ? ThemeManager.Current[key] : ""
    }

    static FontName() {
        name := ThemeManager.Get("FontName")
        if (name != "" && name != "auto")
            return name
        return (I18n.Lang = "zh") ? "Microsoft YaHei UI" : (I18n.Lang = "ja") ? "Yu Gothic UI" : "Segoe UI"
    }

    ; 主题名 -> 完整的键值 Map: 从 Light 开始, 沿 "Base" 逐层叠加。builtinOnly = 只找内置主题
    static Resolve(themeName, builtinOnly := false, depth := 0) {
        if (themeName = "Light" && !FileExist(ThemeManager.UserDir "\Light.json") || depth > 8)
            return ThemeManager.Defaults()
        themeFile := ThemeManager.FileFor(themeName, builtinOnly)
        if (themeFile = "")
            return (themeName = "Light") ? ThemeManager.Defaults() : ""
        try {
            data := JSON.Parse(FileRead(themeFile, "UTF-8"))
        } catch as e {
            Logger.Error("ThemeManager: cannot read " themeFile " - " e.Message)
            return ""
        }
        if !(data is Map)
            return ""
        baseName := data.Has("Base") ? data["Base"] : "Light"
        if (baseName = "" || baseName = "System")
            baseName := "Light"
        ; "Base" 是自己: 用户同名主题在内置那一套上修改
        theme := (baseName = themeName) ? ThemeManager.Resolve(baseName, true, depth + 1) : ThemeManager.Resolve(baseName, false, depth + 1)
        if !IsObject(theme)
            theme := ThemeManager.Defaults()
        for key, value in data
            if (key != "Base")
                theme[key] := value
        return theme
    }

    ; 用户主题优先, 其次内置主题; 都没有返回 ""
    static FileFor(themeName, builtinOnly := false) {
        folders := builtinOnly ? [ThemeManager.BuiltinDir] : [ThemeManager.UserDir, ThemeManager.BuiltinDir]
        for folder in folders
            if FileExist(folder "\" themeName ".json")
                return folder "\" themeName ".json"
        return ""
    }

    static IsBuiltin(themeName) {
        return (themeName = "System" || themeName = "Light" || FileExist(ThemeManager.BuiltinDir "\" themeName ".json") != "")
    }

    ; System, Light, 内置主题 (按 BuiltinOrder, 其余按文件名), 用户主题
    static Names() {
        names := ["System", "Light"]
        for themeName in ThemeManager.BuiltinOrder
            if FileExist(ThemeManager.BuiltinDir "\" themeName ".json")
                names.Push(themeName)
        for folder in [ThemeManager.BuiltinDir, ThemeManager.UserDir] {
            Loop Files, folder "\*.json" {
                SplitPath(A_LoopFileName, , , , &themeName)
                if !ThemeManager._Contains(names, themeName)
                    names.Push(themeName)
            }
        }
        return names
    }

    ; Windows 设置 -> 个性化 -> 颜色 -> "选择默认应用模式"
    static SystemUsesDark() {
        try return !RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme")
        return false
    }

    ; 系统切换浅色 / 深色时 Windows 广播 WM_SETTINGCHANGE ("ImmersiveColorSet"), 重新载入换主题
    static _WatchSystemTheme() {
        if ThemeManager._listening
            return
        ThemeManager._listening := true
        OnMessage(0x1A, (wParam, lParam, *) => ThemeManager._OnSettingChange(lParam))
    }

    static _OnSettingChange(lParam) {
        if (ThemeManager.Name != "System" || !lParam || StrGet(lParam) != "ImmersiveColorSet")
            return
        wanted := ThemeManager.SystemUsesDark() ? "Dark" : "Light"
        if (wanted != ThemeManager.Resolved)
            SetTimer(() => App.Reload(), -500)                              ; 等系统切换完成, 也避免在消息回调里退出
    }

    static _Contains(list, value) {
        for item in list
            if (item = value)
                return true
        return false
    }

    ; Light: 完整的一套键, 其它主题只写和它不同的键
    static Defaults() {
        return Map(
            "FontName"          , "auto",
            "InputFontSize"     , 20,
            "TitleFontSize"     , 13,
            "SubtitleFontSize"  , 9.5,
            "ShortcutFontSize"  , 9,
            "Padding"           , 14,
            "RowHeight"         , 54,
            "IconSize"          , 32,
            "SelectedRadius"    , 0,
            "Opacity"           , 255,
            "Background"        , "FAFAFA",
            "Border"            , "C8C8C8",
            "InputText"         , "1F1F1F",
            "Separator"         , "E4E4E4",
            "Title"             , "1F1F1F",
            "Subtitle"          , "808080",
            "Shortcut"          , "A0A0A0",
            "SelectedBackground", "DDE7F6",
            "SelectedTitle"     , "000000",
            "SelectedSubtitle"  , "4A5568",
            "SelectedShortcut"  , "4A5568"
        )
    }
}
