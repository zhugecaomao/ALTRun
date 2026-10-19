;===============================================================================
; ThemeManager.ahk - 搜索窗口的配色和尺寸 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 主题从哪里来:
;   Light                     写在代码里 (Defaults), 所有主题都在它的基础上修改;
;                             主题文件全部丢失时也能正常显示
;   Resources\Themes\*.json   内置主题 (参照 Alfred 自带的几套), 随程序发布, 升级时整体替换:
;                             Dark / Classic / Midnight / Frost / Graphite / Ocean / Paper / DarkModern / LightModern /
;                             Monokai / OneDark / TokyoNight / Dracula / CatppuccinMocha / GruvboxDark / SolarizedLight (常见编辑器配色)
;   Themes\*.json             用户自己的主题, 升级不动; 和内置主题同名时用户的优先
;   System                    跟随 Windows 的浅色 / 深色设置 (Light 或 Dark), 系统切换时自动更新
; 弹出菜单 (托盘、☰、结果的右键菜单、Quick Switch 的文件夹菜单) 跟随主题的深浅, 见 ApplyMenuMode
; 紧凑 (Appearance.Compact): 任何主题的字号、边距、行高、图标按 CompactScale 缩小, 一屏显示更多结果, 见 Compacted。
; 以前的内置主题 DarkCompact / LightCompact / MidnightCompact 已经去掉: 设置里还是这几个名字 (或者用户主题的 "Base"
; 写的是它们) 时, 当作 Dark / Light / Midnight + 紧凑, 见 LegacyCompact
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
;   Highlight  SelectedHighlight   搜索结果标题里和输入匹配的字的颜色 (普通行 / 选中行)
;   SelectedRadius   选中行的圆角半径, 0 = 整行直角选中条
;   Opacity          窗口不透明度 1~255, 255 = 不透明
;
; 用法:
;   ThemeManager.Load("Dark")
;   ThemeManager.Load("Dark", true)      紧凑
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
    static UserDir    := AppSettings.Portable ? A_ScriptDir "\Themes" : AppSettings.UserDir "\Themes"
    static BuiltinOrder := ["Dark", "Classic", "Midnight", "Frost", "Graphite", "Ocean", "Paper"
                          , "DarkModern", "LightModern", "Monokai", "OneDark", "TokyoNight", "Dracula", "CatppuccinMocha", "GruvboxDark", "SolarizedLight"]   ; 列表里的顺序
    static Compact := false                         ; 实际是不是紧凑 (设置里的, 或者以前的 *Compact 主题)
    ; 紧凑时每个尺寸乘的系数: 浅色的尺寸 (20 / 13 / 9.5 / 9 / 14 / 54 / 32) 正好变成以前的 Dark Compact (17 / 11.5 / 9 / 8.5 / 12 / 48 / 28)
    static CompactScale := Map("InputFontSize", 0.85, "TitleFontSize", 0.885, "SubtitleFontSize", 0.95, "ShortcutFontSize", 0.95
                             , "Padding", 0.86, "RowHeight", 0.89, "IconSize", 0.875)
    static LegacyNames := Map("DarkCompact", "Dark", "LightCompact", "Light", "MidnightCompact", "Midnight")
    static _listening := false

    static Load(themeName, compact := false) {
        if ((base := ThemeManager.LegacyCompact(themeName)) != "")
            themeName := base, compact := true
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
        ThemeManager.Current := compact ? ThemeManager.Compacted(theme) : theme
        ThemeManager.Compact := compact ? true : false
        ThemeManager.Resolved := themeName
        ThemeManager.ApplyMenuMode(ThemeManager.IsDark())
    }

    ; 当前主题是不是深色 (按搜索窗口的背景色)
    static IsDark() => ThemeManager.IsDarkColor(ThemeManager.Get("Background"))

    ; "1E1E1E" 这种颜色是不是深色 (亮度低于一半)
    static IsDarkColor(hex) {
        if !RegExMatch(hex, "i)^[0-9a-f]{6}$")
            return false
        value := Integer("0x" hex)
        return (((value >> 16) & 0xFF) * 299 + ((value >> 8) & 0xFF) * 587 + (value & 0xFF) * 114) / 1000 < 128
    }

    ; 整个程序的弹出菜单用深色还是浅色: uxtheme 的 SetPreferredAppMode (序号 135; 2 = 深色, 3 = 浅色) 再 FlushMenuThemes (136)
    ; 刷新。这两个函数没有公开文档, 资源管理器、Notepad++ 等都在用; Windows 10 1903 (18362) 起才有, 1809 的 135 号是
    ; 另一个函数, 所以按版本号判断, 更早的系统菜单保持浅色。只能用系统自带的深色菜单, 不能用主题里的颜色。返回是否设置了
    static MenuMode := 0                                                    ; 已经设置的: 0 = 没设置过, 2 = 深色, 3 = 浅色
    static ApplyMenuMode(dark) {
        mode := dark ? 2 : 3
        if (mode = ThemeManager.MenuMode || !ThemeManager.SupportsMenuMode())
            return false
        try {
            module := DllCall("GetModuleHandle", "Str", "uxtheme", "Ptr") || DllCall("LoadLibrary", "Str", "uxtheme", "Ptr")
            setMode := DllCall("GetProcAddress", "Ptr", module, "Ptr", 135, "Ptr")
            flush := DllCall("GetProcAddress", "Ptr", module, "Ptr", 136, "Ptr")
            if (!setMode || !flush)
                return false
            DllCall(setMode, "Int", mode)
            DllCall(flush)
            ThemeManager.MenuMode := mode
            return true
        } catch as e {
            Logger.Error("ThemeManager: cannot set menu mode - " e.Message)
        }
        return false
    }

    static SupportsMenuMode(version := A_OSVersion) {
        parts := StrSplit(version, ".")
        return parts.Length >= 3 && parts[1] >= 10 && IsInteger(parts[3]) && parts[3] >= 18362
    }

    static Get(key) {
        return ThemeManager.Current.Has(key) ? ThemeManager.Current[key] : ""
    }

    static FontName() {
        return ThemeManager.FontFrom(ThemeManager.Get("FontName"))
    }

    ; 主题的 FontName -> 实际字体: "auto" (或空) 按界面语言选
    static FontFrom(name) {
        if (name != "" && name != "auto")
            return name
        return (I18n.Font != "") ? I18n.Font : "Segoe UI"                    ; 语言文件的 "_font"
    }

    ; 主题的尺寸按 CompactScale 缩小 (返回新的 Map): 字号取到 0.5 pt, 边距、行高取整, 图标取偶数
    static Compacted(theme) {
        compacted := theme.Clone()
        for key, factor in ThemeManager.CompactScale {
            if !(compacted.Has(key) && IsNumber(compacted[key]))
                continue
            value := compacted[key] * factor
            value := InStr(key, "FontSize") ? Round(value * 2) / 2 : (key = "IconSize") ? Round(value / 2) * 2 : Round(value)
            compacted[key] := (value = Round(value)) ? Round(value) : value    ; 17.0 -> 17 (和主题文件里的写法一样)
        }
        return compacted
    }

    ; 以前的内置紧凑主题 (DarkCompact 等) -> 对应的主题名 ("Dark"); 不是, 或者用户自己有同名的主题文件时 ""
    static LegacyCompact(themeName) {
        if !ThemeManager.LegacyNames.Has(themeName) || ThemeManager.FileFor(themeName) != ""
            return ""
        return ThemeManager.LegacyNames[themeName]
    }

    ; 主题名 -> 完整的键值 Map: 从 Light 开始, 沿 "Base" 逐层叠加。builtinOnly = 只找内置主题
    static Resolve(themeName, builtinOnly := false, depth := 0) {
        if (themeName = "Light" && !FileExist(ThemeManager.UserDir "\Light.json") || depth > 8)
            return ThemeManager.Defaults()
        themeFile := ThemeManager.FileFor(themeName, builtinOnly)
        if (themeFile = "" && ThemeManager.LegacyNames.Has(themeName))      ; 以前的紧凑主题 (用户主题的 "Base" 还写着它)
            return ThemeManager.Compacted(ThemeManager.Resolve(ThemeManager.LegacyNames[themeName], false, depth + 1))
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
            "SelectedShortcut"  , "4A5568",
            "Highlight"         , "005FE0",
            "SelectedHighlight" , "0047B8"
        )
    }
}
