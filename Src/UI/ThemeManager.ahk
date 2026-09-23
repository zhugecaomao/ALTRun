;===============================================================================
; ThemeManager.ahk - 搜索窗口的配色和尺寸 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 内置两套主题 Light / Dark。也可以在 Themes\<名称>.json 里写自己的主题, 只需要
; 写出和 Light 不同的键, 其余自动沿用 Light 的值; 然后在 ALTRun.json 里设置
; "Appearance": { "Theme": "<名称>" }。
;
; 颜色一律写 "RRGGBB"; 字号单位是 pt; 尺寸单位是 96 DPI 下的像素 (会按 DPI 缩放)。
;
; 用法:
;   ThemeManager.Load("Dark")
;   ThemeManager.Get("Background")       -> "1E1E1E"
;   ThemeManager.FontName()              -> 实际使用的字体
;===============================================================================

class ThemeManager {
    static Current := Map()
    static Name    := "Light"

    static Load(themeName) {
        theme := ThemeManager.Builtin("Light")
        if (themeName = "Dark") {
            for key, value in ThemeManager.Builtin("Dark")
                theme[key] := value
        } else if (themeName != "Light") {
            themeFile := A_ScriptDir "\Themes\" themeName ".json"
            try {
                custom := JSON.Parse(FileRead(themeFile, "UTF-8"))
                for key, value in custom
                    theme[key] := value
            } catch as e {
                Logger.Error("ThemeManager: cannot load " themeFile " - " e.Message)
                themeName := "Light"
            }
        }
        ThemeManager.Current := theme
        ThemeManager.Name := themeName
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

    static Builtin(themeName) {
        if (themeName = "Dark") {
            return Map(
                "Background"        , "1E1F22",
                "Border"            , "3C3F44",
                "InputText"         , "F2F2F2",
                "Separator"         , "34363B",
                "Title"             , "E8E8E8",
                "Subtitle"          , "8E9196",
                "Shortcut"          , "6E7176",
                "SelectedBackground", "34445E",
                "SelectedTitle"     , "FFFFFF",
                "SelectedSubtitle"  , "C3CCDA",
                "SelectedShortcut"  , "C3CCDA"
            )
        }
        return Map(
            "FontName"          , "auto",
            "InputFontSize"     , 20,
            "TitleFontSize"     , 13,
            "SubtitleFontSize"  , 9.5,
            "ShortcutFontSize"  , 9,
            "Padding"           , 14,
            "RowHeight"         , 54,
            "IconSize"          , 32,
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
