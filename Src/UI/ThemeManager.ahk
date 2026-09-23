;===============================================================================
; ThemeManager.ahk - 搜索窗口的配色和尺寸 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 内置主题 (参照 Alfred 自带的几套主题):
;   System    跟随 Windows 的浅色 / 深色设置 (Light 或 Dark), 系统切换时自动更新
;   Light     浅色, 默认                 Dark      深色
;   Classic   浅灰底 + 蓝色整行选中条     Midnight  接近纯黑, 圆角选中
;   Frost     半透明的冷白色              Graphite  macOS 深色 + 蓝色强调色
;   Ocean     蓝灰色 (Nord 配色)          Paper     米黄色纸张
;
; 自定义主题: Themes\<名称>.json, 只写要改的键; "Base" 指定从哪个内置主题开始
; (默认 Light)。例如 Themes\MyBlue.json:
;   { "Base": "Dark", "SelectedBackground": "1D4ED8", "Opacity": 240 }
; 然后在 ALTRun.json 里设置 "Appearance": { "Theme": "MyBlue" } (或在偏好设置里选)。
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
;   ThemeManager.Names()                 -> 内置主题 + Themes\*.json 的名称
;===============================================================================

class ThemeManager {
    static Current  := Map()
    static Name     := "Light"                      ; 设置里选的主题
    static Resolved := "Light"                      ; 实际使用的主题 (System -> Light / Dark)
    static BuiltinNames := ["System", "Light", "Dark", "Classic", "Midnight", "Frost", "Graphite", "Ocean", "Paper"]
    static _listening := false

    static Load(themeName) {
        ThemeManager.Name := themeName
        if (themeName = "System") {
            themeName := ThemeManager.SystemUsesDark() ? "Dark" : "Light"
            ThemeManager._WatchSystemTheme()
        }
        theme := ThemeManager.Builtin("Light")
        if ThemeManager.IsBuiltin(themeName) {
            ThemeManager._Merge(theme, ThemeManager.Builtin(themeName))
        } else {
            themeFile := A_ScriptDir "\Themes\" themeName ".json"
            try {
                custom := JSON.Parse(FileRead(themeFile, "UTF-8"))
                if (custom.Has("Base") && ThemeManager.IsBuiltin(custom["Base"]) && custom["Base"] != "System")
                    ThemeManager._Merge(theme, ThemeManager.Builtin(custom["Base"]))
                ThemeManager._Merge(theme, custom)
            } catch as e {
                Logger.Error("ThemeManager: cannot load " themeFile " - " e.Message)
                themeName := "Light"
            }
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
        return (I18n.Lang = "zh") ? "Microsoft YaHei UI" : "Segoe UI"
    }

    static IsBuiltin(themeName) {
        for name in ThemeManager.BuiltinNames
            if (name = themeName)
                return true
        return false
    }

    ; 内置主题 + Themes 文件夹里的自定义主题
    static Names() {
        names := ThemeManager.BuiltinNames.Clone()
        Loop Files, A_ScriptDir "\Themes\*.json" {
            SplitPath(A_LoopFileName, , , , &themeName)
            if !ThemeManager.IsBuiltin(themeName)
                names.Push(themeName)
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

    static _Merge(target, source) {
        for key, value in source
            if (key != "Base")
                target[key] := value
        return target
    }

    ; Light 是完整的一套, 其它主题只写和 Light 不同的键
    static Builtin(themeName) {
        switch themeName, false {
            case "Dark":
                return Map(
                    "Background", "1E1F22", "Border", "3C3F44", "InputText", "F2F2F2", "Separator", "34363B",
                    "Title", "E8E8E8", "Subtitle", "8E9196", "Shortcut", "6E7176",
                    "SelectedBackground", "34445E", "SelectedTitle", "FFFFFF", "SelectedSubtitle", "C3CCDA", "SelectedShortcut", "C3CCDA")
            case "Classic":                                                 ; 经典: 浅灰底, 醒目的蓝色整行选中条
                return Map(
                    "RowHeight", 50, "SelectedRadius", 0,
                    "Background", "ECECEC", "Border", "B4B4B4", "InputText", "2B2B2B", "Separator", "D2D2D2",
                    "Title", "1A1A1A", "Subtitle", "767676", "Shortcut", "9A9A9A",
                    "SelectedBackground", "3874D8", "SelectedTitle", "FFFFFF", "SelectedSubtitle", "DCE8FB", "SelectedShortcut", "DCE8FB")
            case "Midnight":                                                ; 午夜: 接近纯黑, 低调的圆角选中
                return Map(
                    "SelectedRadius", 8,
                    "Background", "121212", "Border", "2C2C2C", "InputText", "FFFFFF", "Separator", "262626",
                    "Title", "EDEDED", "Subtitle", "7C7C7C", "Shortcut", "5A5A5A",
                    "SelectedBackground", "2E2E2E", "SelectedTitle", "FFFFFF", "SelectedSubtitle", "ABABAB", "SelectedShortcut", "ABABAB")
            case "Frost":                                                   ; 霜白: 半透明冷白色
                return Map(
                    "SelectedRadius", 8, "Opacity", 238,
                    "Background", "F4F7FB", "Border", "CFD9E6", "InputText", "1C2B3A", "Separator", "DFE6EF",
                    "Title", "1C2B3A", "Subtitle", "6E7F95", "Shortcut", "98A7BA",
                    "SelectedBackground", "D3E3F8", "SelectedTitle", "0B2540", "SelectedSubtitle", "3E5A7A", "SelectedShortcut", "3E5A7A")
            case "Graphite":                                                ; 石墨: macOS 深色 + 系统蓝
                return Map(
                    "SelectedRadius", 6,
                    "Background", "2B2B2D", "Border", "48484B", "InputText", "F5F5F7", "Separator", "3A3A3C",
                    "Title", "F5F5F7", "Subtitle", "98989D", "Shortcut", "6E6E73",
                    "SelectedBackground", "0A64D6", "SelectedTitle", "FFFFFF", "SelectedSubtitle", "D6E4F7", "SelectedShortcut", "D6E4F7")
            case "Ocean":                                                   ; 海洋: Nord 蓝灰配色
                return Map(
                    "SelectedRadius", 6,
                    "Background", "2E3440", "Border", "434C5E", "InputText", "ECEFF4", "Separator", "3B4252",
                    "Title", "ECEFF4", "Subtitle", "8F9BB3", "Shortcut", "6B7489",
                    "SelectedBackground", "434C5E", "SelectedTitle", "88C0D0", "SelectedSubtitle", "D8DEE9", "SelectedShortcut", "88C0D0")
            case "Paper":                                                   ; 纸张: 米黄色, 适合长时间看
                return Map(
                    "SelectedRadius", 6,
                    "Background", "FBF7EE", "Border", "D9CFBA", "InputText", "3B3226", "Separator", "ECE4D3",
                    "Title", "3B3226", "Subtitle", "93866F", "Shortcut", "B0A58F",
                    "SelectedBackground", "F0E3C6", "SelectedTitle", "2A2218", "SelectedSubtitle", "6F6352", "SelectedShortcut", "6F6352")
            case "Light":
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
        return Map()
    }
}
