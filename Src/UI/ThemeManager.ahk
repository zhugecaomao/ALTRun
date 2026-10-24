;===============================================================================
; ThemeManager.ahk - colors and sizes of the search window (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Where themes come from:
;   Light                     written in the code (Defaults); every theme changes it;
;                             shows fine even if all theme files are lost
;   Resources\Themes\*.json   built-in themes (modeled on the ones that come with Alfred), shipped with the program and replaced on upgrades:
;                             Dark / Classic / Midnight / Frost / Graphite / Ocean / Paper / DarkModern / LightModern /
;                             Monokai / OneDark / TokyoNight / Dracula / CatppuccinMocha / GruvboxDark / SolarizedLight (common editor color schemes)
;   Themes\*.json             the user's own themes, untouched by upgrades; the user's wins over a built-in theme of the same name
;   System                    follows the Windows light / dark setting (Light or Dark), updated when the system switches
; Popup menus (tray, ☰, the results' context menu, Quick Switch's folder menu) follow the theme's darkness, see ApplyMenuMode
; Density (Appearance.Density): font sizes, padding, row height and icons of any theme scale by DensityScale (see Scaled):
;   Comfortable = the theme's own size; Normal (the default for new installs) = the former Dark Compact; Compact = one step smaller
; The former built-in themes DarkCompact / LightCompact / MidnightCompact are gone: when the settings still name them (or a user theme's "Base"
; does), they mean Dark / Light / Midnight + Normal, see LegacyCompact
;
; A theme file only lists the keys it changes; "Base" says which theme it starts from (default Light), e.g. Themes\MyBlue.json:
;   { "Base": "Dark", "SelectedBackground": "1D4ED8", "Opacity": 240 }
; A user theme with the name of a built-in one whose "Base" is itself (e.g. "Base": "Dark" in Themes\Dark.json)
; changes the built-in one. Preferences -> Appearance -> "Copy as custom theme" writes a complete file.
;
; Available keys (colors always "RRGGBB"; font sizes in pt; sizes in pixels at 96 DPI, scaled with the DPI):
;   FontName  InputFontSize  TitleFontSize  SubtitleFontSize  ShortcutFontSize
;   Padding  RowHeight  IconSize
;   Background  Border  InputText  Separator  Title  Subtitle  Shortcut
;   SelectedBackground  SelectedTitle  SelectedSubtitle  SelectedShortcut
;   Highlight  SelectedHighlight   color of the characters in result titles that match the input (normal row / selected row)
;   SelectedRadius   corner radius of the selected row, 0 = a square bar across the row
;   Opacity          window opacity 1~255, 255 = opaque
;
; Usage:
;   ThemeManager.Load("Dark")
;   ThemeManager.Load("Dark", "Compact") compact
;   ThemeManager.Get("Background")       -> "1E1F22"
;   ThemeManager.FontName()              -> the font actually used
;   ThemeManager.Names()                 -> System, Light, built-in themes, user themes
;   ThemeManager.Resolve("Ocean")        -> the complete Map of keys ("" if not found)
;===============================================================================

class ThemeManager {
    static Current  := Map()
    static Name     := "Light"                      ; the theme chosen in the settings
    static Resolved := "Light"                      ; the theme actually used (System -> Light / Dark)
    static BuiltinDir := A_ScriptDir "\Resources\Themes"
    static UserDir    := AppSettings.ThemesDir
    static BuiltinOrder := ["Dark", "Classic", "Midnight", "Frost", "Graphite", "Ocean", "Paper", "Lavender", "Rose", "Forest", "Sunset", "HighContrast"
                          , "DarkModern", "LightModern", "Monokai", "OneDark", "TokyoNight", "Dracula", "CatppuccinMocha", "GruvboxDark", "SolarizedLight"]   ; order in the list
    static Densities := ["Comfortable", "Normal", "Compact"]
    static Density := "Comfortable"                 ; the density actually used
    ; The factor sizes are multiplied by per density (Comfortable = 1). Light's sizes (20 / 13 / 9.5 / 9 / 14 / 54 / 32) at Normal are exactly the former Dark Compact
    ; (17 / 11.5 / 9 / 8.5 / 12 / 48 / 28); at Compact they are 15 / 10.5 / 8.5 / 8 / 10 / 42 / 24
    static DensityScale := Map(
        "Normal" , Map("InputFontSize", 0.85, "TitleFontSize", 0.885, "SubtitleFontSize", 0.95, "ShortcutFontSize", 0.95
                     , "Padding", 0.86, "RowHeight", 0.89, "IconSize", 0.875),
        "Compact", Map("InputFontSize", 0.75, "TitleFontSize", 0.808, "SubtitleFontSize", 0.895, "ShortcutFontSize", 0.889
                     , "Padding", 0.714, "RowHeight", 0.778, "IconSize", 0.75))
    static LegacyNames := Map("DarkCompact", "Dark", "LightCompact", "Light", "MidnightCompact", "Midnight")
    static _listening := false

    static Load(themeName, density := "Comfortable") {
        if !ThemeManager.DensityScale.Has(density)
            density := "Comfortable"                                        ; Comfortable, or an unknown value
        if ((base := ThemeManager.LegacyCompact(themeName)) != "")         ; the former compact themes: at least Normal (Compact stays Compact)
            themeName := base, density := (density = "Comfortable") ? "Normal" : density
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
        ThemeManager.Current := ThemeManager.Scaled(theme, density)
        ThemeManager.Density := ThemeManager.DensityScale.Has(density) ? density : "Comfortable"
        ThemeManager.Resolved := themeName
        ThemeManager.ApplyMenuMode(ThemeManager.IsDark())
    }

    ; Whether the current theme is dark (by the search window's background color)
    static IsDark() => ThemeManager.IsDarkColor(ThemeManager.Get("Background"))

    ; Whether a color like "1E1E1E" is dark (luminance below half)
    static IsDarkColor(hex) {
        if !RegExMatch(hex, "i)^[0-9a-f]{6}$")
            return false
        value := Integer("0x" hex)
        return (((value >> 16) & 0xFF) * 299 + ((value >> 8) & 0xFF) * 587 + (value & 0xFF) * 114) / 1000 < 128
    }

    ; Dark or light popup menus for the whole program: uxtheme's SetPreferredAppMode (ordinal 135; 2 = dark, 3 = light), then FlushMenuThemes (136)
    ; to refresh. Both are undocumented but used by Explorer, Notepad++ and others; they exist since Windows 10 1903 (18362), and ordinal 135 in 1809 is
    ; another function, so the version decides and older systems keep light menus. Only the system's own dark menus are possible, not the theme's colors. Returns whether it was set
    static MenuMode := 0                                                    ; what was set: 0 = nothing yet, 2 = dark, 3 = light
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

    ; The theme's FontName -> the actual font: "auto" (or empty) chooses by UI language
    static FontFrom(name) {
        if (name != "" && name != "auto")
            return name
        return (I18n.Font != "") ? I18n.Font : "Segoe UI"                    ; the language file's "_font"
    }

    ; The theme's sizes scaled by DensityScale of density (a new Map; Comfortable returns it as it is): font sizes to 0.5 pt, padding and row height rounded, icons to even numbers
    static Scaled(theme, density) {
        if !ThemeManager.DensityScale.Has(density)
            return theme
        scaled := theme.Clone()
        for key, factor in ThemeManager.DensityScale[density] {
            if !(scaled.Has(key) && IsNumber(scaled[key]))
                continue
            value := scaled[key] * factor
            value := InStr(key, "FontSize") ? Round(value * 2) / 2 : (key = "IconSize") ? Round(value / 2) * 2 : Round(value)
            scaled[key] := (value = Round(value)) ? Round(value) : value    ; 17.0 -> 17 (as written in theme files)
        }
        return scaled
    }

    ; A former built-in compact theme (DarkCompact etc.) -> the matching theme name ("Dark"); "" if it isn't one, or the user has a theme file of that name
    static LegacyCompact(themeName) {
        if !ThemeManager.LegacyNames.Has(themeName) || ThemeManager.FileFor(themeName) != ""
            return ""
        return ThemeManager.LegacyNames[themeName]
    }

    ; Theme name -> the complete Map of keys: start from Light and layer each "Base". builtinOnly = built-in themes only
    static Resolve(themeName, builtinOnly := false, depth := 0) {
        if (themeName = "Light" && !FileExist(ThemeManager.UserDir "\Light.json") || depth > 8)
            return ThemeManager.Defaults()
        themeFile := ThemeManager.FileFor(themeName, builtinOnly)
        if (themeFile = "" && ThemeManager.LegacyNames.Has(themeName))      ; a former compact theme (a user theme's "Base" still names it)
            return ThemeManager.Scaled(ThemeManager.Resolve(ThemeManager.LegacyNames[themeName], false, depth + 1), "Normal")
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
        ; "Base" is itself: a user theme of the same name changes the built-in one
        theme := (baseName = themeName) ? ThemeManager.Resolve(baseName, true, depth + 1) : ThemeManager.Resolve(baseName, false, depth + 1)
        if !IsObject(theme)
            theme := ThemeManager.Defaults()
        for key, value in data
            if (key != "Base")
                theme[key] := value
        return theme
    }

    ; user themes first, then built-in themes; "" if neither
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

    ; System, Light, built-in themes (in BuiltinOrder, the rest by file name), user themes
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

    ; Windows Settings -> Personalization -> Colors -> "Choose your default app mode"
    static SystemUsesDark() {
        try return !RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme")
        return false
    }

    ; When the system switches between light and dark, Windows broadcasts WM_SETTINGCHANGE ("ImmersiveColorSet"): reload to change the theme
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
            SetTimer(() => App.Reload(), -500)                              ; wait until the system has switched, and don't exit inside a message callback
    }

    static _Contains(list, value) {
        for item in list
            if (item = value)
                return true
        return false
    }

    ; Light: the complete set of keys; other themes only list the keys that differ from it
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
