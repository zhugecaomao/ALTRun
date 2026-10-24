;===============================================================================
; ResultRows.ahk - draws the rows of the search window's result list (AutoHotkey v2)
;-------------------------------------------------------------------------------
; The result list is a ListView whose rows are drawn here (NM_CUSTOMDRAW): icon, title with the matching characters
; highlighted, subtitle, the Ctrl+N hint, the rounded selection, marks and pins. Also owns the fonts and brushes
; (Gdi) the status bar shares, and the text measuring helpers. Part of SearchWindow, whose state it reads.
;
; Usage:
;   ResultRows.CreateGdi()                  after the theme is loaded (SearchWindow.Create)
;   ResultRows.OnCustomDraw(lParam)         the list's NM_CUSTOMDRAW handler
;   ResultRows.HighlightRanges(item)        the highlighted ranges of a result's title
;===============================================================================

class ResultRows {
    static Gdi := Map()

    static CreateGdi() {
        font := ThemeManager.FontName()
        gdi := ResultRows.Gdi
        gdi["TitleFont"]    := ResultRows._CreateFont(font, ThemeManager.Get("TitleFontSize"), 400)
        gdi["SubtitleFont"] := ResultRows._CreateFont(font, ThemeManager.Get("SubtitleFontSize"), 400)
        gdi["ShortcutFont"] := ResultRows._CreateFont(font, ThemeManager.Get("ShortcutFontSize"), 400)
        gdi["StatusFont"]   := gdi["SubtitleFont"]                          ; the status bar text is like the description text
        gdi["KeyFont"]      := ResultRows._CreateFont(font, Max(7, ThemeManager.Get("SubtitleFontSize") - 1), 400)     ; text in the status bar keycaps (Ctrl, K)
        gdi["Background"]   := DllCall("CreateSolidBrush", "UInt", Win.ColorToBgr(ThemeManager.Get("Background")), "Ptr")
        gdi["SelectedBgr"]  := Win.ColorToBgr(ThemeManager.Get("SelectedBackground"))
        gdi["Selected"]     := DllCall("CreateSolidBrush", "UInt", gdi["SelectedBgr"], "Ptr")
        for key in ["Title", "Subtitle", "Shortcut", "SelectedTitle", "SelectedSubtitle", "SelectedShortcut", "Highlight", "SelectedHighlight"]
            gdi[key "Color"] := Win.ColorToBgr(ThemeManager.Get(key))
        ResultRows._ScheduleFontWarmUp()
    }

    ; Warms up the fallback fonts: Segoe UI in the English UI has no CJK characters, so the first time they are drawn Windows loads a fallback font (Microsoft YaHei etc.).
    ; Right after boot that takes hundreds of milliseconds, and inside the result list's custom-draw callback (WM_PAINT) it stalls for a long time: the result rows stay blank
    ; and timers and hotkeys don't run either (found on a freshly started Windows on GitHub, 2026-10). So 1 s after the fonts are created, CJK text is measured once
    ; in each result font on a memory DC to load the fallback fonts first; drawing later doesn't wait
    static FontWarmUpText := "记事本 メモ帳 한글", FontWarmUpDelay := 1000, _fontWarmUp := ""

    static _ScheduleFontWarmUp() {
        if (ResultRows._fontWarmUp = "")
            ResultRows._fontWarmUp := () => ResultRows.WarmUpFonts()
        SetTimer(ResultRows._fontWarmUp, -ResultRows.FontWarmUpDelay)
    }

    static WarmUpFonts() {
        gdi := ResultRows.Gdi
        if !(gdi.Has("TitleFont") && gdi.Has("SubtitleFont"))
            return
        start := Logger.Ms()
        hdc := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
        oldFont := DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["TitleFont"], "Ptr")
        ResultRows.Extents(hdc, ResultRows.FontWarmUpText)
        DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["SubtitleFont"])
        ResultRows.Extents(hdc, ResultRows.FontWarmUpText)
        DllCall("SelectObject", "Ptr", hdc, "Ptr", oldFont)
        DllCall("DeleteDC", "Ptr", hdc)
        Logger.Time("font warm-up", start)
    }

    static _CreateFont(name, pointSize, weight) {
        height := -Round(pointSize * A_ScreenDPI / 72)
        return DllCall("CreateFontW", "Int", height, "Int", 0, "Int", 0, "Int", 0, "Int", weight
            , "UInt", 0, "UInt", 0, "UInt", 0, "UInt", 1, "UInt", 0, "UInt", 0, "UInt", 5, "UInt", 0   ; 1 = DEFAULT_CHARSET, 5 = CLEARTYPE_QUALITY
            , "WStr", name, "Ptr")
    }

    ;---------------------------------------------------------------------------
    ; Painting (NM_CUSTOMDRAW)
    ;---------------------------------------------------------------------------
    ; Not interruptible (Critical): when painting a row takes over about 15 ms (fonts loading the first time, a busy machine...), AutoHotkey lets timers run in the middle of the callback.
    ; An icon timer slipping in then asks for a list repaint (_Repaint) while the list is painting; that doesn't re-enter, so the repaint request stays pending and Windows keeps sending WM_PAINT:
    ; AutoHotkey keeps handling WM_PAINT and never returns to the thread painting the row, and the lower-priority WM_TIMER never gets a turn either -- the result rows stay blank, and timers, hotkeys and
    ; the tray don't respond, yet the window doesn't count as "not responding" (confirmed by a controlled experiment on Windows on GitHub, 2026-10). When not interruptible, timers wait until painting is done
    static OnCustomDraw(lParam) {
        static CDDS_PREPAINT := 0x1, CDDS_ITEMPREPAINT := 0x10001
        static CDRF_DODEFAULT := 0x0, CDRF_SKIPDEFAULT := 0x4, CDRF_NOTIFYITEMDRAW := 0x20
        Critical("On")
        stage := NumGet(lParam, A_PtrSize * 3, "UInt")
        if (stage = CDDS_PREPAINT) {
            SearchWindow._paints++                                          ; the custom-draw callback ran (see SearchWindow._CheckPaint)
            return CDRF_NOTIFYITEMDRAW
        }
        if (stage != CDDS_ITEMPREPAINT)
            return CDRF_DODEFAULT

        hdc := NumGet(lParam, A_PtrSize = 8 ? 32 : 16, "Ptr")
        row := NumGet(lParam, A_PtrSize = 8 ? 56 : 36, "UPtr")             ; dwItemSpec, 0-based
        ; Take this row's result before painting: another thread may slip in while painting (e.g. hiding the window after running a command, or a new search from clearing the box)
        ; and swap in fewer results with SetResults; reading SearchWindow.Results by index afterwards would be out of range ("Invalid index")
        results := SearchWindow.Results, index := SearchWindow.Offset + row + 1
        if (index > results.Length) {
            SearchWindow._LogBlankRow(row + 1, results.Length)
            return CDRF_SKIPDEFAULT
        }
        item := results[index]

        rect := Buffer(16, 0)                                               ; left = LVIR_BOUNDS (0)
        SendMessage(0x100E, row, rect.Ptr, SearchWindow.List.Hwnd)          ; LVM_GETITEMRECT
        ResultRows._PaintRow(hdc, rect, item, index = SearchWindow.Selected, row + 1)
        return CDRF_SKIPDEFAULT
    }

    ; the pin mark Resources\Icons\Pinned.ico, read once per size (0 = no such file)
    static _PinBadge(size) {
        static icons := Map(), file := A_ScriptDir "\Resources\Icons\Pinned.ico"
        if !icons.Has(size)
            icons[size] := FileExist(file) ? DllCall("LoadImageW", "Ptr", 0, "WStr", file, "UInt", 1, "Int", size, "Int", size, "UInt", 0x10, "Ptr") : 0   ; IMAGE_ICON, LR_LOADFROMFILE
        return icons[size]
    }

    static _PaintRow(hdc, rect, item, selected, visibleRow) {
        static DT_RIGHT := 0x2, DT_VCENTER := 0x4, DT_BOTTOM := 0x8, DT_SINGLELINE := 0x20
        static DT_NOPREFIX := 0x800, DT_PATH_ELLIPSIS := 0x4000, DT_END_ELLIPSIS := 0x8000
        gdi := ResultRows.Gdi
        left := NumGet(rect, 0, "Int"), top := NumGet(rect, 4, "Int")
        right := NumGet(rect, 8, "Int"), bottom := NumGet(rect, 12, "Int")
        pad := SearchWindow.Padding, iconSize := SearchWindow.IconSize, rowH := bottom - top

        if (selected && SearchWindow.SelectedRadius > 0) {                 ; rounded selection: a little margin left and right, an antialiased rounded rectangle
            DllCall("FillRect", "Ptr", hdc, "Ptr", rect, "Ptr", gdi["Background"])
            inset := Win.Scale(6), gap := Win.Scale(2)
            Win.FillRoundRect(hdc, left + inset, top + gap, right - inset + 1, bottom - gap + 1, SearchWindow.SelectedRadius, gdi["SelectedBgr"])
        } else {
            DllCall("FillRect", "Ptr", hdc, "Ptr", rect, "Ptr", selected ? gdi["Selected"] : gdi["Background"])
        }
        DllCall("SetBkMode", "Ptr", hdc, "Int", 1)                          ; TRANSPARENT
        if (SearchWindow.Mode = "results" && SearchWindow.IsMarked(item)) {  ; marked: a vertical bar in the highlight color on the left
            mark := Buffer(16)
            NumPut("Int", left + Win.Scale(2), "Int", top + Win.Scale(4), "Int", left + Win.Scale(6), "Int", bottom - Win.Scale(4), mark)
            brush := DllCall("CreateSolidBrush", "UInt", gdi["HighlightColor"], "Ptr")
            DllCall("FillRect", "Ptr", hdc, "Ptr", mark, "Ptr", brush)
            DllCall("DeleteObject", "Ptr", brush)
        }

        iconY := top + (rowH - iconSize) // 2
        if (hIcon := IconCache.Get(item.Icon))
            DllCall("DrawIconEx", "Ptr", hdc, "Int", left + pad, "Int", iconY, "Ptr", hIcon, "Int", iconSize, "Int", iconSize, "UInt", 0, "Ptr", 0, "UInt", 3)
        if (item.Pinned && (hBadge := ResultRows._PinBadge(badge := Max(Win.Scale(13), Round(iconSize * 0.55))))) {     ; pinned: a pin at the bottom right
            offset := iconSize - badge + Win.Scale(3)
            DllCall("DrawIconEx", "Ptr", hdc, "Int", left + pad + offset, "Int", iconY + offset, "Ptr", hBadge, "Int", badge, "Int", badge, "UInt", 0, "Ptr", 0, "UInt", 3)
        }

        textLeft := left + pad + iconSize + Win.Scale(12)
        shortcutW := Win.Scale(56)
        oldFont := DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["ShortcutFont"], "Ptr")
        textRight := right - pad - shortcutW - Win.Scale(8)

        if (visibleRow <= 9 && SearchWindow.Mode = "results") {             ; Ctrl+N hint on the right
            DllCall("SetTextColor", "Ptr", hdc, "UInt", selected ? gdi["SelectedShortcutColor"] : gdi["ShortcutColor"])
            ResultRows.DrawText(hdc, "Ctrl+" visibleRow, right - pad - shortcutW, top, right - pad, bottom, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX)
        } else if (SearchWindow.Mode = "actions") {                         ; actions panel: shortcut or description on the right, with room for its actual width
            hintW := 0, gap := Win.Scale(12)
            if (item.Subtitle != "") {
                hintW := ResultRows.TextWidth(hdc, item.Subtitle)
                DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["TitleFont"])
                titleW := ResultRows.TextWidth(hdc, item.Title)
                DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["ShortcutFont"])
                hintW := ResultRows.HintWidth(right - pad - textLeft, titleW, hintW, gap)
                DllCall("SetTextColor", "Ptr", hdc, "UInt", selected ? gdi["SelectedShortcutColor"] : gdi["ShortcutColor"])
                ResultRows.DrawText(hdc, item.Subtitle, right - pad - hintW, top, right - pad, bottom, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS)
            }
            textRight := right - pad - (hintW ? hintW + gap : 0)
        }

        DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["TitleFont"])
        DllCall("SetTextColor", "Ptr", hdc, "UInt", selected ? gdi["SelectedTitleColor"] : gdi["TitleColor"])
        hasSubtitle := (item.Subtitle != "" && SearchWindow.Mode = "results")
        middle := top + Round(rowH * 0.54)
        ranges := (SearchWindow.Mode = "results") ? ResultRows.HighlightRanges(item) : []
        if hasSubtitle
            ResultRows.DrawTitle(hdc, item.Title, ranges, textLeft, top, textRight, middle, DT_BOTTOM | DT_SINGLELINE | DT_NOPREFIX, selected)
        else
            ResultRows.DrawTitle(hdc, item.Title, ranges, textLeft, top, textRight, bottom, DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX, selected)

        if hasSubtitle {
            DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["SubtitleFont"])
            DllCall("SetTextColor", "Ptr", hdc, "UInt", selected ? gdi["SelectedSubtitleColor"] : gdi["SubtitleColor"])
            ellipsis := (item.Kind = "file" || item.Kind = "folder") ? DT_PATH_ELLIPSIS : DT_END_ELLIPSIS
            ResultRows.DrawText(hdc, item.Subtitle, textLeft, middle + Win.Scale(2), textRight, bottom, DT_SINGLELINE | DT_NOPREFIX | ellipsis)
        }
        DllCall("SelectObject", "Ptr", hdc, "Ptr", oldFont)
    }

    ; Width of the description on the right of an actions panel row: the name first (at most half), the description gets the rest; both in full when they fit
    static HintWidth(available, titleW, hintW, gap) {
        keep := Min(titleW, available // 2)
        return Max(0, Min(hintW, available - keep - gap))
    }

    ; Positions to highlight in the title, cached on the result per search text (not recomputed for repaints or selection moves)
    static HighlightRanges(item) {
        text := SearchWindow.HighlightText
        if (Trim(text) = "" || item.Title = "")
            return []
        if (!item.HasOwnProp("HighlightFor") || item.HighlightFor !== text)
            item.HighlightFor := text, item.HighlightRanges := FuzzyMatcher.MatchRanges(text, item.Title)
        return item.HighlightRanges
    }

    ; Draws the title: matching characters in the Highlight color; cut off with "…" when it doesn't fit.
    ; Each character's position is measured once for the whole title (GetTextExtentExPoint) and the parts are placed at those positions with ExtTextOut;
    ; titles without highlights are drawn the same way. Measuring each part separately and joining them, or using two ways of drawing with and without highlights (DrawText adjusts spacing),
    ; is off by about 1 pixel, so the spacing would jump as the highlight edges move while typing
    static DrawTitle(hdc, title, ranges, left, top, right, bottom, flags, selected) {
        static DT_VCENTER := 0x4, DT_BOTTOM := 0x8, ETO_CLIPPED := 0x4
        if (title = "")
            return
        gdi := ResultRows.Gdi
        available := right - left
        extents := ResultRows.Extents(hdc, title)
        if (extents.Length && extents[extents.Length] > available) {
            ellipsis := ResultRows.TextWidth(hdc, "…"), fit := 0
            while (fit < extents.Length && extents[fit + 1] <= available - ellipsis)
                fit++
            title := SubStr(title, 1, fit) "…"
            extents := ResultRows.Extents(hdc, title)
        }
        metrics := Buffer(60, 0)                                            ; TEXTMETRICW, tmHeight comes first
        DllCall("GetTextMetricsW", "Ptr", hdc, "Ptr", metrics)
        lineHeight := NumGet(metrics, 0, "Int")
        y := (flags & DT_BOTTOM) ? bottom - lineHeight : (flags & DT_VCENTER) ? top + (bottom - top - lineHeight) // 2 : top
        clip := Buffer(16)
        NumPut("Int", left, "Int", top, "Int", right, "Int", bottom, clip)
        normal := selected ? gdi["SelectedTitleColor"] : gdi["TitleColor"]
        highlight := selected ? gdi["SelectedHighlightColor"] : gdi["HighlightColor"]
        pos := 1
        for segment in FuzzyMatcher.Segments(title, ranges) {
            count := StrLen(segment[1])
            advances := Buffer(count * 4)                                   ; width of each character (difference of two adjacent cumulative widths)
            Loop count
                NumPut("Int", extents[pos + A_Index - 1] - (pos + A_Index - 2 >= 1 ? extents[pos + A_Index - 2] : 0), advances, (A_Index - 1) * 4)
            x := left + (pos > 1 ? extents[pos - 1] : 0)
            DllCall("SetTextColor", "Ptr", hdc, "UInt", segment[2] ? highlight : normal)
            DllCall("ExtTextOutW", "Ptr", hdc, "Int", x, "Int", y, "UInt", ETO_CLIPPED, "Ptr", clip, "WStr", segment[1], "UInt", count, "Ptr", advances)
            pos += count
        }
        DllCall("SetTextColor", "Ptr", hdc, "UInt", normal)
    }

    ; Width from the start to the end of each character (pixels, measured once for the whole text): extents[i] = width of the first i characters
    static Extents(hdc, text) {
        count := StrLen(text), extents := []
        if !count
            return extents
        fit := 0, partial := Buffer(count * 4, 0), size := Buffer(8, 0)
        DllCall("GetTextExtentExPointW", "Ptr", hdc, "WStr", text, "Int", count, "Int", 0x7FFFFFFF, "Int*", &fit, "Ptr", partial, "Ptr", size)
        Loop count
            extents.Push(NumGet(partial, (A_Index - 1) * 4, "Int"))
        return extents
    }

    static TextWidth(hdc, text) {
        size := Buffer(8, 0)
        DllCall("GetTextExtentPoint32W", "Ptr", hdc, "WStr", text, "Int", StrLen(text), "Ptr", size)
        return NumGet(size, 0, "Int")
    }

    static DrawText(hdc, text, left, top, right, bottom, flags) {
        rect := Buffer(16)
        NumPut("Int", left, "Int", top, "Int", right, "Int", bottom, rect)
        DllCall("DrawTextW", "Ptr", hdc, "WStr", text, "Int", -1, "Ptr", rect, "UInt", flags)
    }
}
