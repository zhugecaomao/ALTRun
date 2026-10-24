;===============================================================================
; NavIcons.ahk - icons of the page list on the left of Preferences (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Drawn with Windows' own icon fonts (Windows 11: Segoe Fluent Icons, Windows 10: Segoe MDL2 Assets,
; the same monochrome line icons as Windows Settings and PowerToys): no icon files to ship, sharp at high DPI.
; Without either font (e.g. Wine) it returns 0 and the list shows text only.
;
; Usage:
;   il := NavIcons.ImageList(PreferencesWindow.PageKeys)   ; 0 = no icon font
;===============================================================================

class NavIcons {
    static Fonts := ["Segoe Fluent Icons", "Segoe MDL2 Assets"]
    ; page -> glyph (the same code points in both fonts)
    static Glyphs := Map(
        "Prefs.Page.General",      0xE713,    ; Settings
        "Prefs.Page.Window",       0xE721,    ; Search
        "Prefs.Page.Keyboard",     0xE765,    ; KeyboardClassic
        "Prefs.Page.Appearance",   0xE790,    ; Color
        "Prefs.Page.Features",     0xEA86,    ; Puzzle
        "Prefs.Page.Applications", 0xE71D,    ; AllApps
        "Prefs.Page.FileSearch",   0xEC50,    ; FileExplorer
        "Prefs.Page.FileIndex",    0xE8F1,    ; Library
        "Prefs.Page.Commands",     0xE945,    ; LightningBolt
        "Prefs.Page.Snippets",     0xE70F,    ; Edit
        "Prefs.Page.Clipboard",    0xE77F,    ; Paste
        "Prefs.Page.WebSearch",    0xE774,    ; Globe
        "Prefs.Page.Calculator",   0xE8EF,    ; Calculator
        "Prefs.Page.Scripts",      0xE943,    ; Code
        "Prefs.Page.Hotkeys",      0xEDA7,    ; KeyboardShortcut
        "Prefs.Page.QuickSwitch",  0xE8AB,    ; Switch
        "Prefs.Page.QSPanel",      0xE8A0,    ; OpenPane
        "Prefs.Page.DateStamp",    0xE787,    ; Calendar
        "Prefs.Page.Usage",        0xE9D2,    ; AreaChart
        "Prefs.Page.Advanced",     0xE90F)    ; Repair
    static Color := 0x505050                                                ; RGB, a dark gray close to the body text

    ; An image list with one icon per page (in the order of keys), cells 28 wide and 26 high (scaled with DPI), 16 px glyphs centered
    static ImageList(keys) {
        font := NavIcons.FontName()
        if (font = "")
            return 0
        scale := A_ScreenDPI / 96
        w := Round(28 * scale), h := Round(26 * scale)
        il := DllCall("comctl32\ImageList_Create", "Int", w, "Int", h, "UInt", 0x20, "Int", keys.Length, "Int", 0, "Ptr")   ; ILC_COLOR32
        hdc := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
        hFont := DllCall("CreateFontW", "Int", -Round(16 * scale), "Int", 0, "Int", 0, "Int", 0, "Int", 400, "UInt", 0, "UInt", 0, "UInt", 0
                       , "UInt", 1, "UInt", 0, "UInt", 0, "UInt", 4, "UInt", 0, "Str", font, "Ptr")   ; 4 = ANTIALIASED_QUALITY (grayscale, without ClearType's color fringes)
        oldFont := DllCall("SelectObject", "Ptr", hdc, "Ptr", hFont, "Ptr")
        DllCall("SetTextColor", "Ptr", hdc, "UInt", 0xFFFFFF)
        DllCall("SetBkMode", "Ptr", hdc, "Int", 1)
        r := (NavIcons.Color >> 16) & 0xFF, g := (NavIcons.Color >> 8) & 0xFF, b := NavIcons.Color & 0xFF
        for key in keys {
            info := Buffer(40, 0)                                           ; BITMAPINFOHEADER: 32-bit, top-down
            NumPut("UInt", 40, "Int", w, "Int", -h, "UShort", 1, "UShort", 32, info)
            hbm := DllCall("CreateDIBSection", "Ptr", hdc, "Ptr", info, "UInt", 0, "Ptr*", &bits := 0, "Ptr", 0, "UInt", 0, "Ptr")
            old := DllCall("SelectObject", "Ptr", hdc, "Ptr", hbm, "Ptr")
            if NavIcons.Glyphs.Has(key) {
                rect := Buffer(16, 0)
                NumPut("Int", 0, "Int", 0, "Int", w, "Int", h, rect)
                DllCall("DrawTextW", "Ptr", hdc, "Str", Chr(NavIcons.Glyphs[key]), "Int", -1, "Ptr", rect, "UInt", 0x825)   ; CENTER | VCENTER | SINGLELINE | NOPREFIX
                DllCall("GdiFlush")
                ; white text on black: the brightness is the opacity; turn it into premultiplied-alpha dark gray
                Loop w * h {
                    offset := (A_Index - 1) * 4
                    if (a := NumGet(bits, offset, "UChar"))
                        NumPut("UInt", (a << 24) | ((r * a // 255) << 16) | ((g * a // 255) << 8) | (b * a // 255), bits, offset)
                }
            }
            DllCall("SelectObject", "Ptr", hdc, "Ptr", old)
            DllCall("comctl32\ImageList_Add", "Ptr", il, "Ptr", hbm, "Ptr", 0)
            DllCall("DeleteObject", "Ptr", hbm)
        }
        DllCall("SelectObject", "Ptr", hdc, "Ptr", oldFont)
        DllCall("DeleteObject", "Ptr", hFont)
        DllCall("DeleteDC", "Ptr", hdc)
        return il
    }

    ; pages without an icon (for tests), separated by spaces
    static MissingPages(keys) {
        missing := ""
        for key in keys
            if !NavIcons.Glyphs.Has(key)
                missing .= key " "
        return missing
    }

    ; glyphs missing from the icon font (for tests; Win10's Segoe MDL2 Assets has fewer glyphs than Win11's font)
    static MissingGlyphs() {
        missing := ""
        hdc := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
        hFont := DllCall("CreateFontW", "Int", -16, "Int", 0, "Int", 0, "Int", 0, "Int", 400, "UInt", 0, "UInt", 0, "UInt", 0
                       , "UInt", 1, "UInt", 0, "UInt", 0, "UInt", 0, "UInt", 0, "Str", NavIcons.FontName(), "Ptr")
        old := DllCall("SelectObject", "Ptr", hdc, "Ptr", hFont, "Ptr")
        for key, code in NavIcons.Glyphs {
            DllCall("GetGlyphIndicesW", "Ptr", hdc, "Str", Chr(code), "Int", 1, "UShort*", &index := 0, "UInt", 1)   ; GGI_MARK_NONEXISTING_GLYPHS
            if (index = 0xFFFF)
                missing .= key " "
        }
        DllCall("SelectObject", "Ptr", hdc, "Ptr", old)
        DllCall("DeleteObject", "Ptr", hFont)
        DllCall("DeleteDC", "Ptr", hdc)
        return missing
    }

    ; the first icon font installed, "" if none
    static FontName() {
        static found := ""
        if (found = "") {
            found := "-"
            hdc := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
            for name in NavIcons.Fonts {
                hFont := DllCall("CreateFontW", "Int", -16, "Int", 0, "Int", 0, "Int", 0, "Int", 400, "UInt", 0, "UInt", 0, "UInt", 0
                               , "UInt", 1, "UInt", 0, "UInt", 0, "UInt", 0, "UInt", 0, "Str", name, "Ptr")
                old := DllCall("SelectObject", "Ptr", hdc, "Ptr", hFont, "Ptr")
                face := Buffer(64 * 2, 0)
                DllCall("GetTextFaceW", "Ptr", hdc, "Int", 64, "Ptr", face)
                DllCall("SelectObject", "Ptr", hdc, "Ptr", old)
                DllCall("DeleteObject", "Ptr", hFont)
                if (StrGet(face, "UTF-16") = name) {                        ; without the font Windows substitutes another
                    found := name
                    break
                }
            }
            DllCall("DeleteDC", "Ptr", hdc)
        }
        return (found = "-") ? "" : found
    }
}
