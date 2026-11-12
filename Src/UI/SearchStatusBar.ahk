;===============================================================================
; SearchStatusBar.ahk - the status bar at the bottom of the search window (AutoHotkey v2)
;-------------------------------------------------------------------------------
; A picture control below the result rows, drawn into a bitmap: the mode on the left, and what Enter and → do with
; the selected result on the right (round key caps). Part of SearchWindow, whose window, fonts (ResultRows.Gdi) and
; layout it uses. The menu is the icon right of the search box (SearchWindow.ShowMenu).
;
; Usage:
;   SearchStatusBar.Create(gui, width, padding)   while the window is built
;   SearchStatusBar.Place(visibleRows)            after the layout changed
;   SearchStatusBar.Update()                      after the results / selection changed (repaints only on changes)
;===============================================================================

class SearchStatusBar {
    ; Status bar at the bottom (Window.StatusBar, like Listary / Raycast), shown when there are result rows. On the left
    ; the file search / actions panel state; on the right what Enter and → do with the selected result (round key caps),
    ; as Raycast does: the other actions and their keys (Ctrl+Enter, F3...) are in the actions panel.
    ; The usage tip is in the empty search box (its cue banner), not here
    static Height := 0, Control := "", _key := ""

    ; Status bar: a picture control whose contents (background, separators, icons, tip, keycaps) are drawn into a bitmap, see _Render
    static Create(g, w, pad) {
        SearchStatusBar.Height := 0, SearchStatusBar.Control := "", SearchStatusBar._key := ""
        if !SearchStatusBar.Enabled()
            return
        SearchStatusBar.Height := Win.Scale(38)
        SearchStatusBar.Control := g.AddPicture("x0 y0 w" w " h" SearchStatusBar.Height " Hidden +0xE")     ; 0xE = SS_BITMAP
        SearchStatusBar.Place(0)
    }

    static Enabled() {
        window := AppSettings.Window
        return !window.Has("StatusBar") || window["StatusBar"]
    }

    ; status bar height: only shown with result rows
    static Shown(visibleCount) => (SearchStatusBar.Height && visibleCount > 0)

    ; the status bar goes below the results; hidden without result rows
    static Place(visibleCount) {
        if !SearchStatusBar.Height
            return
        shown := SearchStatusBar.Shown(visibleCount)
        if shown
            SearchStatusBar.Control.Move(0, SearchWindow._ContentHeight(visibleCount))
        if (SearchStatusBar.Control.Visible != shown)
            SearchStatusBar.Control.Visible := shown
    }

    ; only repaint the status bar when its contents changed
    static Update() {
        if !SearchStatusBar.Height
            return
        texts := SearchStatusBar.Texts()
        key := texts.Left "`n" texts.Enter "`n" texts.Actions
        if (key == SearchStatusBar._key)
            return
        SearchStatusBar._key := key
        SearchStatusBar._Render(texts)
    }

    ; {Left, Enter, Actions}: the text on the left (file search mode / Esc in the actions panel; the usage tip is in the
    ; empty box instead), what Enter does with the selected result ("" = nothing to show) and whether "Actions →" is shown
    static Texts() {
        actions := SearchWindow.Session.Mode = "actions"
        left := actions ? I18n.T("Status.Back") : SearchWindow.Session.FileMode ? I18n.T("Status.FileMode") : ""
        item := SearchWindow.SelectedItem()
        enter := !IsObject(item) ? "" : actions ? I18n.T("Status.Run") : SearchStatusBar.EnterTitle(item)
        return {Left: left, Enter: enter, Actions: (!actions && IsObject(item) && item.Valid) ? 1 : 0}
    }

    ; What Enter does for the selected result (a short name): the result's own (RunTitle) > by type (open / copy / paste / run); hint rows that can't run: complete
    static EnterTitle(item) {
        if !item.Valid
            return (item.AutoComplete != "") ? I18n.T("Status.Complete") : ""
        if (item.RunTitle != "")
            return item.RunTitle
        switch item.Kind {
            case "file", "folder", "url": return I18n.T("Status.Open")
            case "text": return IsObject(item.OnRun) ? I18n.T("Status.Paste") : I18n.T("Status.Copy")
        }
        return I18n.T("Status.Run")
    }

    ; The parts on the right: [{Label, Keys: [key caps...]}], a line between two parts: Open ↵ | Actions →
    static Parts(texts) {
        parts := []
        if (texts.Enter != "")
            parts.Push({Label: texts.Enter, Keys: ["↵"]})
        if texts.Actions
            parts.Push({Label: I18n.T("Status.Actions"), Keys: ["→"]})
        return parts
    }

    ; Sizes on the status bar's right side: drawing it and checking that it fits use the same ones
    static _Metrics() => {KeyH: Win.Scale(20), KeyPad: Win.Scale(6), Gap: Win.Scale(6), KeyGap: Win.Scale(4), Divider: Win.Scale(8)}

    ; Arrows are drawn in a key cap 1.3 times as wide as high; other keys fit their text
    static _KeyWidth(hdc, keyText, m) => (keyText = "↵" || keyText = "→") ? Round(m.KeyH * 1.3) : Max(m.KeyH, ResultRows.TextWidth(hdc, keyText) + 2 * m.KeyPad)

    ; Width of the parts on the right as _Render draws them (the fonts must exist, see ResultRows.CreateGdi)
    static PartsWidth(parts) {
        m := SearchStatusBar._Metrics(), gdi := ResultRows.Gdi
        hdc := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr"), width := 0
        oldFont := DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["KeyFont"], "Ptr")
        for index, part in parts {
            DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["KeyFont"])
            for keyText in part.Keys
                width += SearchStatusBar._KeyWidth(hdc, keyText, m) + m.KeyGap
            width += m.Gap - m.KeyGap
            DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["StatusFont"])
            width += ResultRows.TextWidth(hdc, part.Label) + (index > 1 ? 2 * m.Divider : 0)
        }
        DllCall("SelectObject", "Ptr", hdc, "Ptr", oldFont)
        DllCall("DeleteDC", "Ptr", hdc)
        return width
    }

    ; Mixes two colors ("RRGGBB") by t (0 ~ 1); returns BGR for GDI
    static MixColor(base, other, t) {
        a := Win.ColorToBgr(base), b := Win.ColorToBgr(other), mixed := 0
        for shift in [0, 8, 16] {
            ca := (a >> shift) & 0xFF, cb := (b >> shift) & 0xFF
            mixed |= Round(ca + (cb - ca) * t) << shift
        }
        return mixed
    }

    ; Draws the status bar bitmap and replaces the one in the picture control. The background is the background color moved slightly towards the text color (a bit darker in light themes, lighter in dark ones),
    ; keycaps a bit more; the ↵ in a keycap uses Segoe UI Symbol (normal fonts may lack it).
    ; A 24-bit DIB (no alpha, like ThemePreview): in a 32-bit bitmap the background and text GDI draws have alpha 0 and only the keycaps GDI+ draws have 255,
    ; and the picture control shows a 32-bit bitmap with its alpha, so the text and background would be invisible
    static _Render(texts) {
        w := SearchWindow.Width, h := SearchStatusBar.Height, pad := SearchWindow.Padding
        background := ThemeManager.Get("Background"), title := ThemeManager.Get("Title")
        gdi := ResultRows.Gdi
        screen := DllCall("GetDC", "Ptr", 0, "Ptr")
        hdc := DllCall("CreateCompatibleDC", "Ptr", screen, "Ptr")
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", screen)
        info := Buffer(40, 0)                                               ; BITMAPINFOHEADER: top-down, 24-bit
        NumPut("UInt", 40, "Int", w, "Int", -h, "UShort", 1, "UShort", 24, info)
        bits := 0
        bitmap := DllCall("CreateDIBSection", "Ptr", hdc, "Ptr", info, "UInt", 0, "Ptr*", &bits, "Ptr", 0, "UInt", 0, "Ptr")
        oldBitmap := DllCall("SelectObject", "Ptr", hdc, "Ptr", bitmap, "Ptr")
        SearchStatusBar.FillRect(hdc, 0, 0, w, h, SearchStatusBar.MixColor(background, title, 0.045))
        SearchStatusBar.FillRect(hdc, 0, 0, w, 1, Win.ColorToBgr(ThemeManager.Get("Separator")))
        DllCall("SetBkMode", "Ptr", hdc, "Int", 1)                         ; TRANSPARENT
        oldFont := DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["StatusFont"], "Ptr")

        ; the right side goes right to left: label + keycap, with a vertical line between the parts
        m := SearchStatusBar._Metrics()
        keyH := m.KeyH, gap := m.Gap, keyGap := m.KeyGap, divider := m.Divider
        top := (h - keyH) // 2 + 1, x := w - pad
        parts := SearchStatusBar.Parts(texts)
        labelBgr := SearchStatusBar.MixColor(title, background, 0.15), keyBgr := SearchStatusBar.MixColor(background, title, 0.11)
        keyTextBgr := SearchStatusBar.MixColor(title, background, 0.3), lineBgr := SearchStatusBar.MixColor(background, title, 0.18)
        index := parts.Length
        while (index >= 1) {
            part := parts[index]
            keyIndex := part.Keys.Length
            while (keyIndex >= 1) {
                keyText := part.Keys[keyIndex]
                DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["KeyFont"])
                arrowKey := (keyText = "↵" || keyText = "→")                    ; the arrows are drawn (the font's arrows are too small and sit too low)
                keyW := SearchStatusBar._KeyWidth(hdc, keyText, m)
                Win.FillRoundRect(hdc, x - keyW, top, x, top + keyH, Win.Scale(5), keyBgr)
                if arrowKey {
                    SearchStatusBar.DrawKeyArrow(hdc, keyText, x - keyW / 2, top + keyH / 2, keyH * 0.5, keyTextBgr)
                } else {
                    DllCall("SetTextColor", "Ptr", hdc, "UInt", keyTextBgr)
                    ResultRows.DrawText(hdc, keyText, x - keyW, top, x, top + keyH, 0x825)      ; DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX
                }
                x -= keyW + keyGap
                keyIndex--
            }
            x -= gap - keyGap
            DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["StatusFont"])
            labelW := ResultRows.TextWidth(hdc, part.Label)
            DllCall("SetTextColor", "Ptr", hdc, "UInt", labelBgr)
            ResultRows.DrawText(hdc, part.Label, x - labelW, 1, x, h, 0x824)             ; DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX
            x -= labelW
            if (index > 1) {                                                ; the vertical line to the previous part
                x -= divider
                SearchStatusBar.FillRect(hdc, x, top + Win.Scale(3), x + Max(1, Win.Scale(1)), top + keyH - Win.Scale(3), lineBgr)
                x -= divider
            }
            index--
        }

        ; Left: the mode (with an ellipsis if it doesn't fit)
        DllCall("SetTextColor", "Ptr", hdc, "UInt", gdi["SubtitleColor"])
        ResultRows.DrawText(hdc, texts.Left, pad, 1, x - Win.Scale(16), h, 0x8824)      ; DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS

        DllCall("SelectObject", "Ptr", hdc, "Ptr", oldFont)
        DllCall("SelectObject", "Ptr", hdc, "Ptr", oldBitmap)
        DllCall("DeleteDC", "Ptr", hdc)
        ctrl := SearchStatusBar.Control.Hwnd
        previous := SendMessage(0x172, 0, bitmap, ctrl)                     ; STM_SETIMAGE: returns the previous bitmap
        if previous
            DllCall("DeleteObject", "Ptr", previous)
        if (SendMessage(0x173, 0, 0, ctrl) != bitmap)                       ; STM_GETIMAGE: when the control made its own copy (a 32-bit bitmap), ours is no longer needed
            DllCall("DeleteObject", "Ptr", bitmap)
    }

    ; An arrow in a keycap (antialiased lines); (cx, cy) is the center, size the width and height:
    ;   ↵ Enter: down from the top right, then left, with the arrowhead at the left end;  → right arrow key: a horizontal line with the arrowhead at the right end
    static DrawKeyArrow(hdc, key, cx, cy, size, bgr) {
        graphics := 0, pen := 0
        if (!ClipboardData.StartGdiplus() || DllCall("gdiplus\GdipCreateFromHDC", "Ptr", hdc, "Ptr*", &graphics) || !graphics)
            return
        DllCall("gdiplus\GdipSetSmoothingMode", "Ptr", graphics, "Int", 4)    ; AntiAlias
        argb := 0xFF000000 | ((bgr & 0xFF) << 16) | (bgr & 0xFF00) | ((bgr >> 16) & 0xFF)
        DllCall("gdiplus\GdipCreatePen1", "UInt", argb, "Float", Max(1.2, size * 0.13), "Int", 2, "Ptr*", &pen)   ; 2 = UnitPixel
        DllCall("gdiplus\GdipSetPenLineJoin", "Ptr", pen, "Int", 2)             ; LineJoinRound
        DllCall("gdiplus\GdipSetPenStartCap", "Ptr", pen, "Int", 2), DllCall("gdiplus\GdipSetPenEndCap", "Ptr", pen, "Int", 2)   ; LineCapRound
        half := size / 2, head := size * 0.3
        if (key = "→") {
            points := [[cx - half, cy], [cx + half, cy]]
            tip := [[cx + half - head, cy - head], [cx + half, cy], [cx + half - head, cy + head]]
        } else {
            baseY := cy + size * 0.18
            points := [[cx + half, cy - half], [cx + half, baseY], [cx - half, baseY]]      ; vertical line + horizontal line
            tip := [[cx - half + head, baseY - head], [cx - half, baseY], [cx - half + head, baseY + head]]
        }
        for lines in [points, tip] {                                        ; line + arrowhead (two slanted lines meeting at one end of the line)
            buf := Buffer(8 * lines.Length)
            for index, point in lines
                NumPut("Float", point[1], "Float", point[2], buf, (index - 1) * 8)
            DllCall("gdiplus\GdipDrawLines", "Ptr", graphics, "Ptr", pen, "Ptr", buf, "Int", lines.Length)
        }
        DllCall("gdiplus\GdipDeletePen", "Ptr", pen)
        DllCall("gdiplus\GdipDeleteGraphics", "Ptr", graphics)
    }

    static FillRect(hdc, left, top, right, bottom, bgr) {
        rect := Buffer(16)
        NumPut("Int", left, "Int", top, "Int", right, "Int", bottom, rect)
        brush := DllCall("CreateSolidBrush", "UInt", bgr, "Ptr")
        DllCall("FillRect", "Ptr", hdc, "Ptr", rect, "Ptr", brush)
        DllCall("DeleteObject", "Ptr", brush)
    }
}
