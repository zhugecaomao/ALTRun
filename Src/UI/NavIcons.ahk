;===============================================================================
; NavIcons.ahk - 偏好设置左边页面列表的图标 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 用 Windows 自带的图标字体画 (Windows 11: Segoe Fluent Icons, Windows 10: Segoe MDL2 Assets,
; 和 Windows 设置、PowerToys 同一套单色线条图标): 不用另外带图标文件, 高 DPI 下也清晰。
; 两种字体都没有时 (例如 Wine) 返回 0, 列表只显示文字。
;
; 用法:
;   il := NavIcons.ImageList(PreferencesWindow.PageKeys)   ; 0 = 没有图标字体
;===============================================================================

class NavIcons {
    static Fonts := ["Segoe Fluent Icons", "Segoe MDL2 Assets"]
    ; 页面 -> 字形 (两种字体的编码相同)
    static Glyphs := Map(
        "Prefs.Page.General",      0xE713,    ; Settings
        "Prefs.Page.Window",       0xE721,    ; Search
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
        "Prefs.Page.Hotkeys",      0xE765,    ; KeyboardClassic
        "Prefs.Page.QuickSwitch",  0xE8AB,    ; Switch
        "Prefs.Page.QSPanel",      0xE8A0,    ; OpenPane
        "Prefs.Page.DateStamp",    0xE787,    ; Calendar
        "Prefs.Page.Usage",        0xE9D2,    ; AreaChart
        "Prefs.Page.Advanced",     0xE90F)    ; Repair
    static Color := 0x505050                                                ; RGB, 和正文接近的深灰

    ; 每页一个图标的图像列表 (keys 的顺序), 格子宽 28、高 26 (按 DPI 放大), 字形 16 px 居中
    static ImageList(keys) {
        font := NavIcons.FontName()
        if (font = "")
            return 0
        scale := A_ScreenDPI / 96
        w := Round(28 * scale), h := Round(26 * scale)
        il := DllCall("comctl32\ImageList_Create", "Int", w, "Int", h, "UInt", 0x20, "Int", keys.Length, "Int", 0, "Ptr")   ; ILC_COLOR32
        hdc := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
        hFont := DllCall("CreateFontW", "Int", -Round(16 * scale), "Int", 0, "Int", 0, "Int", 0, "Int", 400, "UInt", 0, "UInt", 0, "UInt", 0
                       , "UInt", 1, "UInt", 0, "UInt", 0, "UInt", 4, "UInt", 0, "Str", font, "Ptr")   ; 4 = ANTIALIASED_QUALITY (灰度, 不用 ClearType 的彩边)
        oldFont := DllCall("SelectObject", "Ptr", hdc, "Ptr", hFont, "Ptr")
        DllCall("SetTextColor", "Ptr", hdc, "UInt", 0xFFFFFF)
        DllCall("SetBkMode", "Ptr", hdc, "Int", 1)
        r := (NavIcons.Color >> 16) & 0xFF, g := (NavIcons.Color >> 8) & 0xFF, b := NavIcons.Color & 0xFF
        for key in keys {
            info := Buffer(40, 0)                                           ; BITMAPINFOHEADER: 32 位, 从上到下
            NumPut("UInt", 40, "Int", w, "Int", -h, "UShort", 1, "UShort", 32, info)
            hbm := DllCall("CreateDIBSection", "Ptr", hdc, "Ptr", info, "UInt", 0, "Ptr*", &bits := 0, "Ptr", 0, "UInt", 0, "Ptr")
            old := DllCall("SelectObject", "Ptr", hdc, "Ptr", hbm, "Ptr")
            if NavIcons.Glyphs.Has(key) {
                rect := Buffer(16, 0)
                NumPut("Int", 0, "Int", 0, "Int", w, "Int", h, rect)
                DllCall("DrawTextW", "Ptr", hdc, "Str", Chr(NavIcons.Glyphs[key]), "Int", -1, "Ptr", rect, "UInt", 0x825)   ; CENTER | VCENTER | SINGLELINE | NOPREFIX
                DllCall("GdiFlush")
                ; 白字画在黑底上, 亮度就是不透明度: 换成预乘 Alpha 的深灰
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

    ; 没有图标的页面 (测试用), 空格分开
    static MissingPages(keys) {
        missing := ""
        for key in keys
            if !NavIcons.Glyphs.Has(key)
                missing .= key " "
        return missing
    }

    ; 图标字体里没有的字形 (测试用; Win10 的 Segoe MDL2 Assets 比 Win11 的字体少一些字形)
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

    ; 装了的第一种图标字体, 都没有时 ""
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
                if (StrGet(face, "UTF-16") = name) {                        ; 没有这种字体时 Windows 换成别的字体
                    found := name
                    break
                }
            }
            DllCall("DeleteDC", "Ptr", hdc)
        }
        return (found = "-") ? "" : found
    }
}
