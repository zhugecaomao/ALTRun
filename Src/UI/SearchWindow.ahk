;===============================================================================
; SearchWindow.ahk - the Alfred-style search window (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Layout: a borderless window with rounded corners, a large input box at the top and the result list below (each row: icon +
; title + gray subtitle + a Ctrl+N hint on the right); the window height follows the number of results.
;
; The result list is a ListView that draws every row itself (NM_CUSTOMDRAW); it only holds the rows currently visible
; (at most VisibleRows), and Offset does the scrolling, so there is never a scroll bar.
;
; Keyboard (all handled in _OnKeyDown, catching WM_KEYDOWN / WM_SYSKEYDOWN with OnMessage):
;   Enter / Ctrl+Enter / Alt+Enter   run / reveal or paste / copy   (see ActionCatalog)
;   ↑ ↓  Ctrl+P Ctrl+N               up / down one row (moves the selection only, as in Alfred; not mixed with history)
;   PgUp PgDn                        up / down one page (VisibleRows rows)
;   Ctrl+↑ / Ctrl+↓                  previous / next search in the history; past the newest the typed text comes back
;   Ctrl+1 ~ Ctrl+9                  run visible row N directly
;   Tab                              complete the name; on a folder, browse into it (the box becomes "path\")
;                                    with TabKey = "Next": Tab / Shift+Tab = next / previous result (wraps), Ctrl+Tab completes
;   Insert                           mark / unmark files and folders (several can be marked); → acts on all marked ones
;   →  (caret at the end)            open the actions panel; ← / Esc go back
;   Space (empty search box)         file search mode (files only, hint "Search files..."); Backspace goes back
;   Space (text typed, SpaceToRun)   run the selected result; Shift+Space types a space
;   Ctrl+C (no text selected in the box)  copy the current result
;   Ctrl+L                           Large Type
;   Ctrl+E                           edit the selected file (when its type has an "Edit" verb, see ActionCatalog.EditorFor)
;   F3                               edit the current result (no results: a new custom command from the typed text)
;   Ctrl+Del (caret at the end)      delete the current result (after confirming)
;   F1                               about ALTRun (the Advanced page of Preferences)
;   F2 / Ctrl+,                      Preferences
;   F4                               edit Data\ALTRun.json in Notepad
;   Esc                              close the actions panel / hide the window
; Mouse: click selects, double-click runs, right-click shows the result's action menu (the same as the actions panel);
;        dragging the blank area around the input box moves the window
;
; Position: Appearance.ShowOn decides the screen (Mouse = under the mouse / Primary = primary screen /
; Active = the active window's). By default centered horizontally, 20% from the top; with Appearance.RememberPosition,
; a dragged position is saved relative to the screen (per mille, Appearance.Position), so it lands in the matching place on another screen.
;
; Usage:
;   SearchWindow.Create()             once at start-up
;   SearchWindow.Show([text])  Hide()  Toggle()
;===============================================================================

class SearchWindow {
    static Gui := "", Input := "", List := "", Separator := ""
    static Results := [], Selected := 0, Offset := 0
    static Mode := "results"                        ; results = search results; actions = the actions panel
    static HighlightText := ""                      ; characters in result titles matching it are highlighted (Highlight color)
    static FileMode := false                        ; file search mode: Space in the empty search box enters it; files only
    static ActionSource := "", SavedQuery := "", AllActions := []
    static _actionsOnly := false                                            ; an actions panel opened directly (actions for the selection): Esc / ← close the window
    static HistoryIndex := 0                         ; which history entry is shown (Knowledge.History, 1 = most recent), 0 = not browsing
    static Marked := Map()                           ; files / folders marked with Insert: lower-case path -> path (cleared each time the window is shown)
    static _historyDraft := ""                       ; the text in the box before browsing the history; restored by Ctrl+↓
    static Width := 0, Padding := 0, InputHeight := 0, RowHeight := 0, IconSize := 0, VisibleRows := 8
    static SelectedRadius := 0
    static _gdi := Map()
    static _searchTimer := "", _hideTimer := ""
    static _shownRows := -1                                                 ; how many result rows the window is currently sized for
    static _tip := ""                                                       ; the usage tip for this showing (HelpProvider.NextTip)
    static LastArea := "", HiddenAt := 0                                    ; work area of the screen at the last hide, and its time (for Hud)
    static _last := ""                                                      ; the search at the last hide {Text, FileMode, Selected} (KeepLastQuery)
    static _layoutBefore := 0                                               ; input language of the foreground window before showing (switched back after hiding when SwitchToEnglishInput = 1)
    static _imeBefore := ""                                                 ; {Open, Conv} of the IME when shown (restored after hiding when SwitchToEnglishInput = 2)
    static _keepOpen := false                        ; don't hide on focus loss during a context menu / delete confirmation
    ; Status bar at the bottom (Appearance.StatusBar, like Listary / Raycast), shown when there are result rows. On the left
    ; the ☰ menu button (Preferences, cheat sheet, rebuild index, check for updates, reload, quit) and the file search /
    ; actions panel state; on the right what Enter, Ctrl+Enter, F3 and → do with the selected result (round key caps).
    ; The usage tip is in the empty search box (its cue banner), not here
    static StatusHeight := 0, StatusBar := "", _statusKey := ""

    ;---------------------------------------------------------------------------
    ; Create
    ;---------------------------------------------------------------------------
    static IconPrefetchDelay := 2000                                        ; how long after start-up the common icons start being prefetched (see IconCache.Prefetch)
    static Create() {
        SearchWindow.Width       := Win.Scale(AppSettings.Appearance["Width"])
        SearchWindow.VisibleRows := Max(1, Min(AppSettings.Appearance["VisibleRows"], 9))
        SearchWindow.Padding     := Win.Scale(ThemeManager.Get("Padding"))
        SearchWindow.RowHeight   := Win.Scale(ThemeManager.Get("RowHeight"))
        SearchWindow.IconSize    := Win.Scale(ThemeManager.Get("IconSize"))
        SearchWindow.SelectedRadius := Win.Scale(ThemeManager.Get("SelectedRadius"))
        if (IconCache.Size != SearchWindow.IconSize)                        ; the theme / DPI changed: cached icons were scaled for the old size
            IconCache.Clear()
        IconCache.Size := SearchWindow.IconSize
        IconCache.OnLoaded := () => SearchWindow._Repaint()
        BackgroundWork.IsBusy := () => !SearchWindow._AllowsBackgroundWork()
        SearchWindow._CreateGdiObjects()

        w := SearchWindow.Width, pad := SearchWindow.Padding
        background := ThemeManager.Get("Background")

        g := Gui("-Caption +AlwaysOnTop +ToolWindow -DPIScale" (SearchWindow._IsWindows11() ? "" : " +Border"), App.Name)
        g.BackColor := background
        g.MarginX := 0, g.MarginY := 0
        g.SetFont("s" ThemeManager.Get("InputFontSize") " c" ThemeManager.Get("InputText"), ThemeManager.FontName())

        searchBox := g.AddEdit("x" pad " y" pad " w" (w - 2 * pad) " r1 -Multi -E0x200 -VScroll -WantReturn Background" background)
        searchBox.GetPos(, , , &editH)
        SearchWindow.InputHeight := Max(editH, Win.Scale(40))
        metrics := Win.CapMetrics(searchBox.Hwnd)
        searchBox.Move(pad, pad + SearchWindow.InputTop(SearchWindow.InputHeight, editH, pad, metrics.Ascent, metrics.Cap))
        searchBox.OnEvent("Change", (*) => SearchWindow._OnInputChange())
        Win.SetCueBanner(searchBox.Hwnd, SearchWindow.Placeholder())

        listTop := pad + SearchWindow.InputHeight + pad // 2
        separator := g.AddText("x0 y" listTop " w" w " h1 Hidden Background" ThemeManager.Get("Separator"))

        ; 0x2000 = LVS_NOSCROLL; LV0x10000 = double buffering, against flicker
        list := g.AddListView("x0 y" (listTop + 1) " w" w " h0 Hidden -Hdr -Multi -E0x200 -TabStop +0x2000 +LV0x10000 Background" background, ["Result"])
        list.ModifyCol(1, w)
        rowImageList := DllCall("comctl32\ImageList_Create", "Int", 1, "Int", SearchWindow.RowHeight, "UInt", 0x21, "Int", 1, "Int", 1, "Ptr")
        list.SetImageList(rowImageList, 1)                                  ; an empty small image list only to set the row height
        ; The real row height: the ListView is 1 pixel taller than the image list (2026-10: a theme row of 48 pixels was really 49); when the list and window counted 48,
        ; each row was off by 1 pixel and with 8 rows the bottom 8 pixels of the last row were cut off (the lower corners of the selected row disappeared). Lay out with the measured height
        SearchWindow.RowHeight := Win.ListRowHeight(list, SearchWindow.RowHeight)
        list.OnNotify(-12, (ctrl, lParam) => SearchWindow._OnCustomDraw(lParam))    ; NM_CUSTOMDRAW
        list.OnEvent("Click", (ctrl, row) => SearchWindow._OnRowClick(row))
        list.OnEvent("DoubleClick", (ctrl, row) => SearchWindow._OnRowDoubleClick(row))
        list.OnEvent("ContextMenu", (ctrl, row, isRightClick, x, y) => SearchWindow._OnRowContextMenu(row, x, y))

        g.OnEvent("Close", (*) => SearchWindow.Hide())

        SearchWindow.Gui := g, SearchWindow.Input := searchBox, SearchWindow.List := list, SearchWindow.Separator := separator
        SearchWindow._CreateStatusBar(g, w, pad)

        OnMessage(0x100, (p*) => SearchWindow._OnKeyMessage(p*))            ; WM_KEYDOWN
        OnMessage(0x104, (p*) => SearchWindow._OnKeyMessage(p*))            ; WM_SYSKEYDOWN (Alt+...)
        OnMessage(0x6,   (p*) => SearchWindow._OnActivate(p*))              ; WM_ACTIVATE
        OnMessage(0x20A, (p*) => SearchWindow._OnMouseWheel(p*))            ; WM_MOUSEWHEEL
        OnMessage(0x201, (p*) => SearchWindow._OnLButtonDown(p*))           ; WM_LBUTTONDOWN: drag the window
        OnMessage(0x232, (p*) => SearchWindow._OnMoved(p*))                 ; WM_EXITSIZEMOVE: dragging ended
        OnMessage(0x7E,  (*) => Logger.Debug("Display changed: " MonitorGetCount() " monitor(s)"))   ; WM_DISPLAYCHANGE: to compare with the log when it hangs

        g.Show("Hide w" w " h" SearchWindow._WindowHeight(0))
        try DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", g.Hwnd, "UInt", 3, "Int*", 1, "UInt", 4)   ; DWMWA_TRANSITIONS_FORCEDISABLED: no fade animation, appear as soon as the hotkey is pressed
        Win.SetCorner(g.Hwnd)
        Win.SetBorderColor(g.Hwnd, ThemeManager.Get("Border"))
        if (SearchWindow.Opacity() < 255)
            WindowFade.SetAlpha(g.Hwnd, SearchWindow.Opacity())
    }

    ; the theme's window opacity (Opacity, 1-254); 255 = opaque when not set
    static Opacity() {
        opacity := ThemeManager.Get("Opacity")
        return (IsNumber(opacity) && opacity >= 1 && opacity < 255) ? Round(opacity) : 255
    }

    ; Status bar: a picture control whose contents (background, separators, icons, tip, keycaps) are drawn into a bitmap, see _RenderStatusBar
    static _CreateStatusBar(g, w, pad) {
        SearchWindow.StatusHeight := 0, SearchWindow.StatusBar := "", SearchWindow._statusKey := ""
        if !SearchWindow.StatusBarEnabled()
            return
        SearchWindow.StatusHeight := Win.Scale(38)
        SearchWindow.StatusBar := g.AddPicture("x0 y0 w" w " h" SearchWindow.StatusHeight " Hidden +0xE")   ; 0xE = SS_BITMAP
        SearchWindow.StatusBar.OnEvent("Click", (*) => SearchWindow._OnStatusClick())   ; click on ☰: the menu (clicks elsewhere on the status bar do nothing)
        SearchWindow._PlaceStatusBar(0)
    }

    static StatusBarEnabled() {
        appearance := AppSettings.Appearance
        return !appearance.Has("StatusBar") || appearance["StatusBar"]
    }

    ; status bar height: only shown with result rows
    static _StatusShown(visibleCount) => (SearchWindow.StatusHeight && visibleCount > 0)

    ; the status bar goes below the results; hidden without result rows
    static _PlaceStatusBar(visibleCount) {
        if !SearchWindow.StatusHeight
            return
        shown := SearchWindow._StatusShown(visibleCount)
        if shown
            SearchWindow.StatusBar.Move(0, SearchWindow._ContentHeight(visibleCount))
        if (SearchWindow.StatusBar.Visible != shown)
            SearchWindow.StatusBar.Visible := shown
    }

    ; only repaint the status bar when its contents changed
    static _UpdateStatusBar() {
        if !SearchWindow.StatusHeight
            return
        texts := SearchWindow.StatusTexts()
        key := texts.Left "`n" texts.Enter "`n" texts.Ctrl "`n" texts.Edit "`n" texts.Actions
        if (key == SearchWindow._statusKey)
            return
        SearchWindow._statusKey := key
        SearchWindow._RenderStatusBar(texts)
    }

    ; {Left, Enter, Ctrl, Edit, Actions}: the text on the left (file search mode / Esc in the actions panel; the usage
    ; tip is in the empty box instead), what Enter / Ctrl+Enter do with the selected result ("" = nothing to show),
    ; whether F3 edits it and whether "Actions →" is shown
    static StatusTexts() {
        actions := SearchWindow.Mode = "actions"
        left := actions ? I18n.T("Status.Back") : SearchWindow.FileMode ? I18n.T("Status.FileMode") : ""
        item := SearchWindow.SelectedItem()
        enter := !IsObject(item) ? "" : actions ? I18n.T("Status.Run") : SearchWindow.EnterTitle(item)
        usable := !actions && IsObject(item) && item.Valid
        ctrl := usable ? SearchWindow.CtrlEnterTitle(item, enter) : ""
        return {Left: left, Enter: enter, Ctrl: ctrl, Edit: (usable && ResultActions.CanEdit(item)) ? 1 : 0, Actions: usable ? 1 : 0}
    }

    ; What Ctrl+Enter does (ActionCatalog.RunModifier); "" when it is the same as Enter
    static CtrlEnterTitle(item, enter) {
        switch item.Kind {
            case "file", "folder": title := I18n.T("Status.Reveal")
            case "text"          : title := I18n.T("Status.Paste")
            case "url"           : title := I18n.T("Status.Copy")
            default              : return ""
        }
        return (title = enter) ? "" : title
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

    ; The parts on the right: [{Label, Keys: [key caps...]}], a line between two parts:
    ; Open ↵ | Reveal Ctrl ↵ | Edit F3 | Actions →
    static StatusParts(texts) {
        parts := []
        if (texts.Enter != "")
            parts.Push({Label: texts.Enter, Keys: ["↵"]})
        if (texts.Ctrl != "")
            parts.Push({Label: texts.Ctrl, Keys: ["Ctrl", "↵"]})
        if texts.Edit
            parts.Push({Label: I18n.T("Status.Edit"), Keys: ["F3"]})
        if texts.Actions
            parts.Push({Label: I18n.T("Status.Actions"), Keys: ["→"]})
        return parts
    }

    ; Sizes on the status bar's right side: drawing it and checking that it fits use the same ones
    static _StatusMetrics() => {KeyH: Win.Scale(20), KeyPad: Win.Scale(6), Gap: Win.Scale(6), KeyGap: Win.Scale(4), Divider: Win.Scale(8)}

    ; Arrows are drawn in a key cap 1.3 times as wide as high; other keys fit their text
    static _KeyWidth(hdc, keyText, m) => (keyText = "↵" || keyText = "→") ? Round(m.KeyH * 1.3) : Max(m.KeyH, SearchWindow._TextWidth(hdc, keyText) + 2 * m.KeyPad)

    ; Width of the parts on the right as _RenderStatusBar draws them (the fonts must exist, see _CreateGdiObjects)
    static StatusPartsWidth(parts) {
        m := SearchWindow._StatusMetrics(), gdi := SearchWindow._gdi
        hdc := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr"), width := 0
        oldFont := DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["KeyFont"], "Ptr")
        for index, part in parts {
            DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["KeyFont"])
            for keyText in part.Keys
                width += SearchWindow._KeyWidth(hdc, keyText, m) + m.KeyGap
            width += m.Gap - m.KeyGap
            DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["StatusFont"])
            width += SearchWindow._TextWidth(hdc, part.Label) + (index > 1 ? 2 * m.Divider : 0)
        }
        DllCall("SelectObject", "Ptr", hdc, "Ptr", oldFont)
        DllCall("DeleteDC", "Ptr", hdc)
        return width
    }

    ; width taken by the ☰ button (from the left of the status bar)
    static StatusMenuWidth() => SearchWindow.StatusMenuWidthFor(SearchWindow.Padding)
    static StatusMenuWidthFor(pad) => pad + Win.Scale(28)

    ; The status bar was clicked: show the menu when the click is on ☰
    static _OnStatusClick() {
        point := Buffer(8)
        DllCall("GetCursorPos", "Ptr", point)
        DllCall("ScreenToClient", "Ptr", SearchWindow.StatusBar.Hwnd, "Ptr", point)
        if (NumGet(point, 0, "Int") <= SearchWindow.StatusMenuWidth())
            SearchWindow.ShowMainMenu()
    }

    ; The ☰ menu's items: [[name, function], ...], "" = separator
    static MainMenuItems() {
        later := (fn) => (*) => (SearchWindow.Hide(), SetTimer(fn, -1))     ; hide the search window first, act after the menu closes
        return [[I18n.T("Tray.Preferences") "`tCtrl+,", later(() => App.OpenPreferences())]
              , [I18n.T("Menu.CheatSheet") "`t?", (*) => SearchWindow._SetInput("?")]
              , ""
              , [I18n.T("Tray.RebuildIndex"), later(() => App.RebuildIndex())]
              , [I18n.T("Tray.CheckUpdate"), later(() => UpdateChecker.Check())]
              , ""
              , [I18n.T("Tray.Reload"), later(() => App.Reload())]
              , [I18n.T("Tray.Exit"), later(() => App.Quit())]]
    }

    ; Shows the menu above ☰ (the status bar is at the bottom of the window; the system puts the menu where it fits)
    static ShowMainMenu() {
        mainMenu := Menu()
        for entry in SearchWindow.MainMenuItems()
            IsObject(entry) ? mainMenu.Add(entry[1], entry[2]) : mainMenu.Add()
        SearchWindow.StatusBar.GetPos(&x, &y)
        SearchWindow._keepOpen := true
        CoordMode("Menu", "Client")
        mainMenu.Show(SearchWindow.Padding, y)
        SearchWindow._keepOpen := false
        if (SearchWindow.IsVisible() && !WinActive("ahk_id " SearchWindow.Gui.Hwnd))
            try WinActivate("ahk_id " SearchWindow.Gui.Hwnd)
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
    static _RenderStatusBar(texts) {
        w := SearchWindow.Width, h := SearchWindow.StatusHeight, pad := SearchWindow.Padding
        background := ThemeManager.Get("Background"), title := ThemeManager.Get("Title")
        gdi := SearchWindow._gdi
        screen := DllCall("GetDC", "Ptr", 0, "Ptr")
        hdc := DllCall("CreateCompatibleDC", "Ptr", screen, "Ptr")
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", screen)
        info := Buffer(40, 0)                                               ; BITMAPINFOHEADER: top-down, 24-bit
        NumPut("UInt", 40, "Int", w, "Int", -h, "UShort", 1, "UShort", 24, info)
        bits := 0
        bitmap := DllCall("CreateDIBSection", "Ptr", hdc, "Ptr", info, "UInt", 0, "Ptr*", &bits, "Ptr", 0, "UInt", 0, "Ptr")
        oldBitmap := DllCall("SelectObject", "Ptr", hdc, "Ptr", bitmap, "Ptr")
        SearchWindow._FillRect(hdc, 0, 0, w, h, SearchWindow.MixColor(background, title, 0.045))
        SearchWindow._FillRect(hdc, 0, 0, w, 1, Win.ColorToBgr(ThemeManager.Get("Separator")))
        DllCall("SetBkMode", "Ptr", hdc, "Int", 1)                         ; TRANSPARENT
        oldFont := DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["StatusFont"], "Ptr")

        ; the right side goes right to left: label + keycap, with a vertical line between the parts
        m := SearchWindow._StatusMetrics()
        keyH := m.KeyH, gap := m.Gap, keyGap := m.KeyGap, divider := m.Divider
        top := (h - keyH) // 2 + 1, x := w - pad
        parts := SearchWindow.StatusParts(texts)
        labelBgr := SearchWindow.MixColor(title, background, 0.15), keyBgr := SearchWindow.MixColor(background, title, 0.11)
        keyTextBgr := SearchWindow.MixColor(title, background, 0.3), lineBgr := SearchWindow.MixColor(background, title, 0.18)
        index := parts.Length
        while (index >= 1) {
            part := parts[index]
            keyIndex := part.Keys.Length
            while (keyIndex >= 1) {
                keyText := part.Keys[keyIndex]
                DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["KeyFont"])
                arrowKey := (keyText = "↵" || keyText = "→")                    ; the arrows are drawn (the font's arrows are too small and sit too low)
                keyW := SearchWindow._KeyWidth(hdc, keyText, m)
                Win.FillRoundRect(hdc, x - keyW, top, x, top + keyH, Win.Scale(5), keyBgr)
                if arrowKey {
                    SearchWindow.DrawKeyArrow(hdc, keyText, x - keyW / 2, top + keyH / 2, keyH * 0.5, keyTextBgr)
                } else {
                    DllCall("SetTextColor", "Ptr", hdc, "UInt", keyTextBgr)
                    SearchWindow._DrawText(hdc, keyText, x - keyW, top, x, top + keyH, 0x825)   ; DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX
                }
                x -= keyW + keyGap
                keyIndex--
            }
            x -= gap - keyGap
            DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["StatusFont"])
            labelW := SearchWindow._TextWidth(hdc, part.Label)
            DllCall("SetTextColor", "Ptr", hdc, "UInt", labelBgr)
            SearchWindow._DrawText(hdc, part.Label, x - labelW, 1, x, h, 0x824)          ; DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX
            x -= labelW
            if (index > 1) {                                                ; the vertical line to the previous part
                x -= divider
                SearchWindow._FillRect(hdc, x, top + Win.Scale(3), x + Max(1, Win.Scale(1)), top + keyH - Win.Scale(3), lineBgr)
                x -= divider
            }
            index--
        }

        ; Left: the ☰ menu button (three lines) + the tip (with an ellipsis if it doesn't fit)
        lineW := Win.Scale(14), lineH := Max(1, Round(Win.Scale(1) * 1.5)), lineGap := Win.Scale(4)
        lineTop := (h - 3 * lineH - 2 * lineGap) // 2 + 1
        Loop 3
            SearchWindow._FillRect(hdc, pad, lineTop + (A_Index - 1) * (lineH + lineGap), pad + lineW, lineTop + (A_Index - 1) * (lineH + lineGap) + lineH, labelBgr)
        left := SearchWindow.StatusMenuWidth()
        DllCall("SetTextColor", "Ptr", hdc, "UInt", gdi["SubtitleColor"])
        SearchWindow._DrawText(hdc, texts.Left, left, 1, x - Win.Scale(16), h, 0x8824)  ; DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS

        DllCall("SelectObject", "Ptr", hdc, "Ptr", oldFont)
        DllCall("SelectObject", "Ptr", hdc, "Ptr", oldBitmap)
        DllCall("DeleteDC", "Ptr", hdc)
        ctrl := SearchWindow.StatusBar.Hwnd
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

    static _FillRect(hdc, left, top, right, bottom, bgr) {
        rect := Buffer(16)
        NumPut("Int", left, "Int", top, "Int", right, "Int", bottom, rect)
        brush := DllCall("CreateSolidBrush", "UInt", bgr, "Ptr")
        DllCall("FillRect", "Ptr", hdc, "Ptr", rect, "Ptr", brush)
        DllCall("DeleteObject", "Ptr", brush)
    }

    static _CreateGdiObjects() {
        font := ThemeManager.FontName()
        gdi := SearchWindow._gdi
        gdi["TitleFont"]    := SearchWindow._CreateFont(font, ThemeManager.Get("TitleFontSize"), 400)
        gdi["SubtitleFont"] := SearchWindow._CreateFont(font, ThemeManager.Get("SubtitleFontSize"), 400)
        gdi["ShortcutFont"] := SearchWindow._CreateFont(font, ThemeManager.Get("ShortcutFontSize"), 400)
        gdi["StatusFont"]   := gdi["SubtitleFont"]                          ; the status bar text is like the description text
        gdi["KeyFont"]      := SearchWindow._CreateFont(font, Max(7, ThemeManager.Get("SubtitleFontSize") - 1), 400)   ; text in the status bar keycaps (Ctrl, K)
        gdi["Background"]   := DllCall("CreateSolidBrush", "UInt", Win.ColorToBgr(ThemeManager.Get("Background")), "Ptr")
        gdi["SelectedBgr"]  := Win.ColorToBgr(ThemeManager.Get("SelectedBackground"))
        gdi["Selected"]     := DllCall("CreateSolidBrush", "UInt", gdi["SelectedBgr"], "Ptr")
        for key in ["Title", "Subtitle", "Shortcut", "SelectedTitle", "SelectedSubtitle", "SelectedShortcut", "Highlight", "SelectedHighlight"]
            gdi[key "Color"] := Win.ColorToBgr(ThemeManager.Get(key))
        SearchWindow._ScheduleFontWarmUp()
    }

    ; Warms up the fallback fonts: Segoe UI in the English UI has no CJK characters, so the first time they are drawn Windows loads a fallback font (Microsoft YaHei etc.).
    ; Right after boot that takes hundreds of milliseconds, and inside the result list's custom-draw callback (WM_PAINT) it stalls for a long time: the result rows stay blank
    ; and timers and hotkeys don't run either (found on a freshly started Windows on GitHub, 2026-10). So 1 s after the fonts are created, CJK text is measured once
    ; in each result font on a memory DC to load the fallback fonts first; drawing later doesn't wait
    static FontWarmUpText := "记事本 メモ帳 한글", FontWarmUpDelay := 1000, _fontWarmUp := ""
    static _ScheduleFontWarmUp() {
        if (SearchWindow._fontWarmUp = "")
            SearchWindow._fontWarmUp := () => SearchWindow.WarmUpFonts()
        SetTimer(SearchWindow._fontWarmUp, -SearchWindow.FontWarmUpDelay)
    }

    static WarmUpFonts() {
        gdi := SearchWindow._gdi
        if !(gdi.Has("TitleFont") && gdi.Has("SubtitleFont"))
            return
        start := Logger.Ms()
        hdc := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
        oldFont := DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["TitleFont"], "Ptr")
        SearchWindow._Extents(hdc, SearchWindow.FontWarmUpText)
        DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["SubtitleFont"])
        SearchWindow._Extents(hdc, SearchWindow.FontWarmUpText)
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

    static _IsWindows11() {
        return VerCompare(A_OSVersion, "10.0.22000") >= 0
    }

    ;---------------------------------------------------------------------------
    ; Show / Hide
    ;---------------------------------------------------------------------------
    static IsVisible() {
        return IsObject(SearchWindow.Gui) && DllCall("IsWindowVisible", "Ptr", SearchWindow.Gui.Hwnd)
    }

    static IsActive() {
        return IsObject(SearchWindow.Gui) && WinActive("ahk_id " SearchWindow.Gui.Hwnd) ? true : false
    }

    ; text: the text to search; empty if not given; with "keep the last search" (KeepLastQuery) the last
    ; input, file search mode and selected row come back with the text selected: Enter runs it again, typing starts a new search
    static _showing := false, _showStarted := 0, _showStep := ""
    static StuckShowMs := 3000                                              ; when the last showing hasn't finished after this long, it counts as stuck and no longer blocks a new one
    ; typeAhead: characters typed while showing are kept and added to the search box afterwards (see _StartTypeAhead)
    static Show(text := "", typeAhead := true) {
        if !IsObject(SearchWindow.Gui)
            return
        if SearchWindow._ShouldSkipShow()
            return
        if (SearchWindow._prepareTimer != "")
            SetTimer(SearchWindow._prepareTimer, 0)
        if typeAhead
            SearchWindow._StartTypeAhead()
        SearchWindow._showing := true, SearchWindow._showStarted := A_TickCount, SearchWindow._showStep := "start"
        try SearchWindow._Show(text)
        finally {
            SearchWindow._showing := false
            typed := SearchWindow._TakeTypeAhead()
        }
        SearchWindow._ApplyTypeAhead(typed)
    }

    ; From the hotkey press until the search box has focus takes tens of milliseconds (first the empty box is searched, then the window shown; longer right after boot), and characters typed
    ; in that time used to go to the foreground program, lost or typed into another window. Now InputHook catches and keeps the keys that type characters during that time
    ; (Ctrl / Alt combinations, arrow keys etc. as usual) and adds them to the search box once it has focus; an Enter pressed meanwhile runs the first result after the search, Esc closes the window.
    ; At most TypeAheadMs: if showing hangs, typing isn't held forever
    static TypeAheadMs := 3000
    static _typeAhead := ""
    static _StartTypeAhead() {
        if (IsObject(SearchWindow._typeAhead) || SearchWindow.IsActive())
            return
        try {
            ih := InputHook("T" SearchWindow.TypeAheadMs / 1000)
            ih.KeyOpt("{Enter}{NumpadEnter}{Esc}", "ES")
            ih.Start()
            SearchWindow._typeAhead := ih
        }
    }

    ; Stops recording; returns {Text: the typed characters, Key: "Enter" / "NumpadEnter" / "Escape" / "", Modifier: Ctrl / Alt held with Enter}
    static _TakeTypeAhead() {
        ih := SearchWindow._typeAhead, SearchWindow._typeAhead := ""
        if !IsObject(ih)
            return {Text: "", Key: "", Modifier: ""}
        try ih.Stop()
        key := (ih.EndReason = "EndKey") ? ih.EndKey : ""
        modifier := (key = "") ? "" : InStr(ih.EndMods, "^") ? "ctrl" : InStr(ih.EndMods, "!") ? "alt" : ""
        return {Text: RegExReplace(ih.Input, "[\x00-\x1F\x7F]"), Key: key, Modifier: modifier}   ; no Tab or other control characters
    }

    ; Treats the characters typed while showing as typed in the search box: they replace selected text (the restored last search); a leading space enters file search mode
    static _ApplyTypeAhead(typed) {
        if (typed.Text = "" && typed.Key = "")
            return
        if Logger.Enabled
            Logger.Debug("SearchWindow: typed while showing " SearchWindow.TypeAheadSummary(typed))
        if (typed.Key = "Escape")
            return SearchWindow.Hide("Esc")
        text := typed.Text
        if (text != "") {
            if (SubStr(text, 1, 1) = " " && SearchWindow.Mode = "results" && !SearchWindow.FileMode) {
                if SearchWindow._AllTextSelected()
                    SearchWindow.Input.Value := ""
                if SearchWindow._CanEnterFileMode() {
                    SearchWindow.FileMode := true
                    SearchWindow._UpdateCueBanner()
                    text := SubStr(text, 2)
                }
            }
            SendMessage(0xC2, 1, StrPtr(text), SearchWindow.Input.Hwnd)   ; EM_REPLACESEL: replace the selection / insert at the caret
            if (SearchWindow._searchTimer != "")
                SetTimer(SearchWindow._searchTimer, 0)
            SearchWindow._RunSearch()                                       ; search now: a following Enter runs the results for these characters
        }
        if (typed.Key != "")
            SearchWindow._Execute(typed.Modifier)
    }

    ; For the log: shown again while still visible (see _BringToFront), and which program had the focus
    static ShowNote(wasVisible, active, foreground) {
        return !wasVisible ? "" : active ? " (already visible, active)" : " (already visible, not active; foreground: " foreground ")"
    }

    ; For the log: the program and window class in front, e.g. "notepad.exe Notepad" (no window titles)
    static ForegroundText() {
        try return WinGetProcessName("A") " " WinGetClass("A")
        return "none"
    }

    ; For the log: how much was typed, not what (the log never holds the search text), e.g. "5 chars (leading space) + Enter"
    static TypeAheadSummary(typed) {
        return StrLen(typed.Text) " chars" (SubStr(typed.Text, 1, 1) = " " ? " (leading space)" : "") (typed.Key != "" ? " + " typed.Key : "")
    }

    ; The last showing hasn't finished (e.g. the hotkey was pressed again meanwhile): no re-entry. But when it has been stuck too long (2026-10: stuck for 2.5 min after turning off a monitor,
    ; with the hotkey and tray not working), stop waiting and show again; the log notes how long it was stuck and at which step
    static _ShouldSkipShow() {
        if !SearchWindow._showing
            return false
        stuck := A_TickCount - SearchWindow._showStarted
        Logger.Debug("SearchWindow: show skipped, still showing (" stuck " ms, at " SearchWindow._showStep ")")
        if (stuck < SearchWindow.StuckShowMs)
            return true
        Logger.Warn("SearchWindow: previous show stuck for " stuck " ms at " SearchWindow._showStep ", showing again")
        return false
    }

    static _Step(name) {
        SearchWindow._showStep := name
        Logger.Trace("show: " name)
    }

    ; Background work (icon prefetch, index scans, cache writes) runs only while the window is hidden, so it doesn't
    ; compete with typing. Not while showing either: halfway through, the window isn't up yet (it counts as hidden);
    ; once showing takes more than about 15 ms a timer can slip in, and one COM call that reads an icon can take
    ; seconds, so showing waits for it (2.3 s on an old PC, 2026-10)
    static _AllowsBackgroundWork() => !SearchWindow._showing && !SearchWindow.IsVisible()

    ; Logs a step of showing that is particularly slow (Trace only writes right after start-up; this always does)
    static SlowStepMs := 100
    static _LogSlowStep(name, start) {
        elapsed := Logger.Ms() - start
        if (Logger.Enabled && elapsed >= SearchWindow.SlowStepMs)
            Logger.Debug("Perf: slow show step " name " " Round(elapsed) " ms")
    }

    static _Show(text) {
        started := Logger.Ms()
        SearchWindow._blankLogged := false, SearchWindow._skipLogged := false
        wasVisible := SearchWindow.IsVisible()
        if Logger.Enabled                                                   ; pairs with "Perf: show" below: a stuck show is stuck in here
            Logger.Debug("SearchWindow: show" SearchWindow.ShowNote(wasVisible, wasVisible && SearchWindow.IsActive()
                , wasVisible ? SearchWindow.ForegroundText() : ""))
        if (text = "" && wasVisible)                                        ; already open: bring it to the front as it is
            return SearchWindow._BringToFront(started)
        SearchWindow._Step("previous window")
        App.RememberActiveWindow()
        SearchWindow._actionsOnly := false
        SearchWindow.Marked := Map()
        Usage.Count("Show")
        last := SearchWindow._last
        restore := (text = "" && AppSettings.General["KeepLastQuery"] && IsObject(last) && last.Text != "")
        ; Restoring the last search with the box and results still as they were when hidden: show them as they are without searching again (only new input searches).
        ; A search (e.g. file search) takes tens of milliseconds; searching before showing meant the first key pressed meanwhile was still in the old window, lost or delayed
        reuse := restore && SearchWindow.Mode = "results" && SearchWindow.Input.Value == last.Text && SearchWindow.FileMode = last.FileMode
        ; the empty search box was already prepared after hiding (PrepareEmpty): no search first either, refresh after showing (_RefreshShown)
        prepared := !restore && text = "" && !wasVisible && SearchWindow.IsPrepared()
        SearchWindow._prepared := false                                     ; use it once; prepared again after the next hide
        SearchWindow.Mode := "results"
        SearchWindow.FileMode := restore ? last.FileMode : false
        SearchWindow._tip := AppSettings.General["ShowTips"] ? HelpProvider.NextTip() : ""
        SearchWindow._UpdateCueBanner()
        SearchWindow._UpdateStatusBar()
        SearchWindow.HistoryIndex := 0
        SearchWindow._Step("results")
        if (!reuse && !prepared) {
            if restore                                                      ; restoring the last search: show the results that don't need waiting first, then search fully after showing (_FinishRestore)
                SearchWindow._ShowQuickResults(last.Text)
            else
                SearchWindow._SetInput(text)
        }

        SearchWindow._Step("position")
        appearance := AppSettings.Appearance
        pos := SearchWindow.Place(SearchWindow._ScreenArea(appearance["ShowOn"]), SearchWindow.Width
            , SearchWindow._WindowHeight(SearchWindow.VisibleRows), appearance["RememberPosition"] ? appearance["Position"] : "")
        ; Repaint requests while the window is hidden are dropped; repainting only after showing makes the result list flash white first.
        ; So DWM hides (cloaks) the window first, and it is revealed after showing and painting the whole window at once
        hwnd := SearchWindow.Gui.Hwnd
        SearchWindow._Step("window")
        SearchWindow._HookKeysIfModal()
        hidden := !SearchWindow.IsVisible()
        WindowFade.Stop()                                                   ; the shadow of the last hide is still fading out: remove it now
        cloaked := hidden && SearchWindow._Cloak(hwnd, true)
        try {
            if (hidden && AppSettings.Appearance["Fade"])                   ; fade in (see WindowFade)
                WindowFade.In(hwnd, SearchWindow.Opacity())
            SearchWindow.Gui.Show("x" pos.X " y" pos.Y " w" SearchWindow.Width " h" SearchWindow._WindowHeight(SearchWindow._VisibleCount()))
            SearchWindow._shownRows := SearchWindow._VisibleCount()
            SearchWindow._Step("redraw")
            DllCall("RedrawWindow", "Ptr", hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x185)   ; RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW
        } finally {
            if cloaked
                SearchWindow._Cloak(hwnd, false)
        }
        SearchWindow._Step("activate")
        try WinActivate("ahk_id " SearchWindow.Gui.Hwnd)                    ; take focus first, so all keys pressed from now go into the search box
        SearchWindow._Step("focus")
        SearchWindow.Input.Focus()
        len := StrLen(SearchWindow.Input.Value)
        if restore
            SendMessage(0xB1, 0, -1, SearchWindow.Input.Hwnd)               ; restored text is selected: typing replaces it
        else
            SendMessage(0xB1, len, len, SearchWindow.Input.Hwnd)            ; the Edit selects all on focus: put the caret back at the end
        switch AppSettings.General["SwitchToEnglishInput"] {
            case 1:                                                         ; switch to the English (US) keyboard
                SearchWindow._Step("input language: read")                  ; logged in two steps: right after boot this once took 17 s; see which step
                step := Logger.Ms()
                if (!wasVisible || !SearchWindow._layoutBefore)                 ; switch back to the previous input language on hide
                    SearchWindow._layoutBefore := App.PreviousWindow ? Win.KeyboardLayout(App.PreviousWindow) : 0
                SearchWindow._LogSlowStep("input language: read", step)
                SearchWindow._Step("input language: switch")
                step := Logger.Ms()
                Win.SwitchToEnglishIME()
                SearchWindow._LogSlowStep("input language: switch", step)
            case 2:                                                         ; keep the input method, switch it to English mode
                SearchWindow._Step("input mode")
                step := Logger.Ms()
                SearchWindow._SwitchImeToEnglish(wasVisible)
                SearchWindow._LogSlowStep("input mode", step)
        }
        if prepared {
            gen := SearchWindow._searchGen
            SetTimer(() => SearchWindow._RefreshShown(gen), -1)
        }
        if (restore && !reuse) {
            gen := SearchWindow._searchGen
            SetTimer(() => SearchWindow._FinishRestore(gen, last), -1)
        }
        Logger.Time(prepared ? "show (prepared)" : "show", started)
    }

    ; Shown again while still open (the hotkey while it is not active, the tray's Show, or ALTRun.exe started again, which
    ; hands over to this instance): keep the search text, results, mode and caret, only activate it. It used to start
    ; over with an empty box, so what was being typed was lost (2026-10, appinn: another program started ALTRun.exe).
    ; The window "before ALTRun" stays the one from the real show, as in _Reopen
    static _BringToFront(started) {
        Usage.Count("Show")
        hwnd := SearchWindow.Input.Hwnd
        selection := SendMessage(0xB0, 0, 0, hwnd)                          ; EM_GETSEL
        SearchWindow._Step("activate")
        try WinActivate("ahk_id " SearchWindow.Gui.Hwnd)
        SearchWindow._Step("focus")
        SearchWindow.Input.Focus()
        SendMessage(0xB1, selection & 0xFFFF, (selection >> 16) & 0xFFFF, hwnd)   ; EM_SETSEL: focusing the box selects all
        Logger.Time("show (already open)", started)
    }

    ; Switches the search box's IME to English mode (setting the open status and conversion mode like the CapsLock Chinese / English toggle), keeping the old state to restore on hide.
    ; No extra English (US) keyboard needed: the input method stays the same, only in English (suggested in #154, 2026-10)
    static _SwitchImeToEnglish(wasVisible) {
        if !(imeWnd := DllCall("imm32\ImmGetDefaultIMEWnd", "Ptr", SearchWindow.Input.Hwnd, "Ptr"))
            return
        before := Win.ImeState(imeWnd)
        if (!wasVisible || !IsObject(SearchWindow._imeBefore))
            SearchWindow._imeBefore := before
        target := Win.ImeEnglish(before.Conv)
        Win.SetImeState(imeWnd, target.Open, target.Conv)
        if Logger.Enabled {
            after := Win.ImeState(imeWnd)
            Logger.Debug(Format("IME to English: open {} conv 0x{:X} -> now open {} conv 0x{:X}", before.Open, before.Conv, after.Open, after.Conv))
        }
    }

    ; Showing switched to an English input language: switch back on hide. Only needed when all programs share one input language (the Windows default);
    ; with "use a different input method for each app window" only ALTRun itself switched and other programs are unaffected
    static _RestoreInputLanguage() {
        SearchWindow._RestoreImeMode()
        layout := SearchWindow._layoutBefore, SearchWindow._layoutBefore := 0
        if (!layout || AppSettings.General["SwitchToEnglishInput"] != 1 || !AppSettings.General["RestoreInput"]
            || Win.PerWindowInputMethod() || layout = Win.KeyboardLayout())
            return
        if SearchWindow.IsActive() {                                        ; still in the foreground: switch back ourselves, which applies to all programs
            try DllCall("ActivateKeyboardLayout", "Ptr", layout, "UInt", 0)
        } else if (hwnd := WinExist("A")) {                                 ; focus already lost: ask the foreground window to switch
            try PostMessage(0x50, 0, layout, , "ahk_id " hwnd)              ; WM_INPUTLANGCHANGEREQUEST
        }
    }

    ; Showing switched the IME to English mode: restore the old open status and conversion mode on hide. Most Chinese IMEs share one Chinese / English state across programs,
    ; so without restoring, other programs would be in English too. If focus is already lost, set it on the foreground window (like switching the input language back above)
    static _RestoreImeMode() {
        before := SearchWindow._imeBefore, SearchWindow._imeBefore := ""
        if (!IsObject(before) || AppSettings.General["SwitchToEnglishInput"] != 2 || !AppSettings.General["RestoreInput"])
            return
        if !(hwnd := WinExist("A")) || !(imeWnd := Win.ImeWindow(hwnd))
            return
        Win.SetImeState(imeWnd, before.Open, before.Conv)
        if Logger.Enabled {
            after := Win.ImeState(imeWnd)
            Logger.Debug(Format("IME restore: open {} conv 0x{:X}{} -> now open {} conv 0x{:X}", before.Open, before.Conv
                , SearchWindow.IsActive() ? "" : " (other window)", after.Open, after.Conv))
        }
    }

    ; DWMWA_CLOAK: the window shows and paints as usual, DWM just doesn't put it on screen. true = success (Windows 8 and later)
    static _Cloak(hwnd, on) {
        try return DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "UInt", 13, "Int*", on ? 1 : 0, "UInt", 4) = 0
        return false
    }

    ; The window's position {X, Y} in area. position: {X, Y} per mille (0 = left / top edge, 1000 = right / bottom edge;
    ; Y is where the window's top edge sits within the work area height), "" = the default (centered horizontally, 20% from the top).
    ; maxHeight: the window height with the most results; however low the position, the whole window stays on screen
    static Place(area, width, maxHeight, position := "") {
        fx := 500, fy := 200
        if (position is Map && position.Has("X") && position.Has("Y") && IsNumber(position["X"]) && IsNumber(position["Y"]))
            fx := position["X"], fy := position["Y"]
        areaW := area.Right - area.Left, areaH := area.Bottom - area.Top
        x := area.Left + Round((areaW - width) * Max(0, Min(fx, 1000)) / 1000)
        y := area.Top + Round(areaH * Max(0, Min(fy, 1000)) / 1000)
        x := Max(area.Left, Min(x, area.Right - width))
        y := Max(area.Top, Min(y, area.Bottom - maxHeight))
        return {X: x, Y: y}
    }

    ; The reverse of Place: the per mille position of the window's top-left corner (x, y) in area
    static RelativePosition(area, width, x, y) {
        areaW := area.Right - area.Left, areaH := area.Bottom - area.Top
        fx := (areaW > width) ? Round((x - area.Left) * 1000 / (areaW - width)) : 500
        fy := (areaH > 0) ? Round((y - area.Top) * 1000 / areaH) : 200
        return Map("X", Max(0, Min(fx, 1000)), "Y", Max(0, Min(fy, 1000)))
    }

    ; showOn: Mouse = the screen under the mouse / Primary = primary screen / Active = the screen of the window active before showing (the mouse's if none)
    static _ScreenArea(showOn) {
        switch showOn, false {
            case "Primary": return Win.WorkArea(MonitorGetPrimary())
            case "Active":
                area := Win.WorkAreaOfWindow(App.PreviousWindow)
                if IsObject(area)
                    return area
        }
        return Win.WorkAreaAtMouse()
    }

    ; Dragging the window itself (the blank area around the input box, not the box or the result list) works like dragging a title bar
    static _OnLButtonDown(wParam, lParam, msg, hwnd) {
        if (!IsObject(SearchWindow.Gui) || hwnd != SearchWindow.Gui.Hwnd)
            return
        PostMessage(0xA1, 2, , , "ahk_id " hwnd)                            ; WM_NCLBUTTONDOWN, HTCAPTION
        return 0
    }

    ; Dragging ended: save the position when "remember position" is on (otherwise it only lasts while the window is shown this time)
    static _OnMoved(wParam, lParam, msg, hwnd) {
        if (!IsObject(SearchWindow.Gui) || hwnd != SearchWindow.Gui.Hwnd || !AppSettings.Appearance["RememberPosition"])
            return
        WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
        area := Win.WorkAreaAt(x + w // 2, y + SearchWindow.InputHeight // 2)
        AppSettings.Appearance["Position"] := SearchWindow.RelativePosition(area, w, x, y)
        AppSettings.Save()
    }

    ; reason: for the log ("run", "Esc", "hotkey", "deactivated by ..."); when the search box seems to empty itself
    ; (2026-10, appinn), the log shows whether the window was hidden and shown again in between
    static Hide(reason := "") {
        if SearchWindow.IsVisible() {
            if Logger.Enabled
                Logger.Debug("SearchWindow: hide" (reason != "" ? " (" reason ")" : ""))
            try {                                                           ; a notice right after hiding (e.g. "Pinned") shows on the same screen, see Hud
                WinGetPos(&x, &y, &w, &h, "ahk_id " SearchWindow.Gui.Hwnd)
                SearchWindow.LastArea := Win.WorkAreaAt(x + w // 2, y + SearchWindow.InputHeight // 2), SearchWindow.HiddenAt := A_TickCount
            }
            SearchWindow._RememberQuery()
            SearchWindow._RestoreInputLanguage()
            if AppSettings.Appearance["Fade"]
                WindowFade.Out(SearchWindow.Gui.Hwnd)                       ; the window hides at once as usual, leaving a shadow that fades out
            SearchWindow.Gui.Hide()
            SearchWindow._UnhookKeys()
            SearchWindow._PrepareLater()
        }
        LargeType.Close()
    }

    ; After hiding, prepare the next showing in the background: the empty search box and its results (pinned, recently opened items, the new version notice).
    ; Showing then needs no search first: it shows and takes focus at once, so typing can start right after the hotkey; it searches again after showing (_RefreshShown) and swaps in only if something changed.
    ; Prepared after a short wait: actions on the current results right after hiding (menus, confirmations...) finish first. With "keep the last search"
    ; the box and results stay as they are (see reuse in _Show), nothing to prepare
    static PrepareDelay := 200, _prepareTimer := "", _prepared := false
    static _PrepareLater() {
        if AppSettings.General["KeepLastQuery"]
            return
        if (SearchWindow._prepareTimer = "")
            SearchWindow._prepareTimer := () => SearchWindow.PrepareEmpty()
        SetTimer(SearchWindow._prepareTimer, -SearchWindow.PrepareDelay)
    }

    static PrepareEmpty() {
        if (!IsObject(SearchWindow.Gui) || SearchWindow.IsVisible())
            return
        started := Logger.Ms()
        SearchWindow.Mode := "results", SearchWindow.FileMode := false
        SearchWindow._SetInput("")
        SearchWindow._prepared := true
        Logger.Time("prepare empty box", started)
    }

    ; Restoring the last search (kept on showing, but the box and results are no longer as they were when hidden): file search (Everything) usually takes
    ; tens of milliseconds, and when Everything is busy it waits for the timeout (1 s) with the window not appearing (2026-10 log: show 1089 ms). So first only the results of
    ; features that don't make you wait (applications, commands, snippets...), show the window at once, then search fully after showing (_FinishRestore)
    static _quietText := ""
    static _ShowQuickResults(text) {
        SearchWindow._quietText := text
        SearchWindow._inputLen := StrLen(text), SearchWindow.Input.Value := text
        ++SearchWindow._searchGen
        browse := FileSearchProvider.BrowsePath(text)
        SearchWindow.HighlightText := IsObject(browse) ? browse.Filter : text
        SearchWindow.SetResults(SearchWindow.FileMode ? [] : ProviderRegistry.Search(text, FileSearchProvider))
    }

    ; Full search after showing, selecting the row selected last time; not needed once typing started / the actions panel opened
    static _FinishRestore(gen, last) {
        SearchWindow._quietText := ""
        if (gen != SearchWindow._searchGen || !SearchWindow.IsVisible() || SearchWindow.Mode != "results"
            || SearchWindow.Input.Value != last.Text)
            return
        if (SearchWindow._searchTimer != "")
            SetTimer(SearchWindow._searchTimer, 0)
        SearchWindow._RunSearch()
        if (SearchWindow.Input.Value == last.Text)
            SearchWindow.MoveSelection(last.Selected - SearchWindow.Selected)
    }

    ; the prepared empty search box hasn't been used or changed yet
    static IsPrepared() {
        return SearchWindow._prepared && SearchWindow.Mode = "results" && !SearchWindow.FileMode && SearchWindow.Input.Value == ""
    }

    ; Searches again after showing the prepared results: swaps in only if something changed while hidden (e.g. a new version was just found); not needed once typing started
    static _RefreshShown(gen) {
        if (gen != SearchWindow._searchGen || !SearchWindow.IsVisible() || SearchWindow.Mode != "results" || SearchWindow.FileMode
            || SearchWindow.Input.Value != "")
            return
        results := ProviderRegistry.Search("")
        if (gen = SearchWindow._searchGen && !SearchWindow.SameResults(results, SearchWindow.Results))
            SearchWindow.SetResults(results)
    }

    static SameResults(a, b) {
        if (a.Length != b.Length)
            return false
        for index, item in a
            if (item.Title !== b[index].Title || item.Subtitle !== b[index].Subtitle)
                return false
        return true
    }

    ; Remembers the search on hide (for KeepLastQuery); in the actions panel the search before the panel opened
    static _RememberQuery() {
        actions := (SearchWindow.Mode = "actions")
        SearchWindow._last := {Text: actions ? SearchWindow.SavedQuery : SearchWindow.Input.Value
                             , FileMode: SearchWindow.FileMode, Selected: actions ? 1 : Max(1, SearchWindow.Selected)}
    }

    static Toggle() {
        if (SearchWindow.IsVisible() && WinActive("ahk_id " SearchWindow.Gui.Hwnd))
            SearchWindow.Hide("hotkey")
        else
            SearchWindow.Show()
    }

    ;---------------------------------------------------------------------------
    ; Search / results
    ;---------------------------------------------------------------------------
    static _SetInput(text) {
        if (text = "" && Logger.Enabled && SearchWindow.IsVisible())
            SearchWindow._LogEmptied(StrLen(SearchWindow.Input.Value), "by ALTRun, mode " SearchWindow.Mode)
        SearchWindow._inputLen := StrLen(text), SearchWindow.Input.Value := text
        len := StrLen(text)
        SendMessage(0xB1, len, len, SearchWindow.Input.Hwnd)                ; EM_SETSEL: caret to the end
        SearchWindow._RunSearch()
    }

    ; The search box went from 2 or more characters to empty in one step: by ALTRun (_SetInput) or by one edit (select all
    ; and delete, an input method, another program). Only the count is logged, never the text
    static _inputLen := 0
    static _LogEmptied(before, how) {
        if (before >= 2)
            Logger.Debug("SearchWindow: input emptied (" before " chars, " how ")")
    }

    static _OnInputChange() {
        if (SearchWindow._quietText != "" && SearchWindow.Input.Value == SearchWindow._quietText)
            return                                                          ; text filled in by restoring the last search: search after showing (_FinishRestore)
        text := SearchWindow.Input.Value
        if (text = "" && Logger.Enabled)
            SearchWindow._LogEmptied(SearchWindow._inputLen, "by an edit")
        SearchWindow._inputLen := StrLen(text)
        if RegExMatch(text, "[`r`n`t]") {                                   ; multi-line text was pasted: join it into one line, the search box always has one line
            SearchWindow.Input.Value := RegExReplace(text, "\s*[`r`n`t]+\s*", " ")
            len := StrLen(SearchWindow.Input.Value)
            SendMessage(0xB1, len, len, SearchWindow.Input.Hwnd)
        }
        SearchWindow.HistoryIndex := 0
        if (SearchWindow._searchTimer = "")
            SearchWindow._searchTimer := () => SearchWindow._RunSearch()
        SetTimer(SearchWindow._searchTimer, -1)                             ; while typing continuously only the last input is searched
    }

    ; Searches on every character typed (the box's Change event, each its own AHK thread). When a search is slow, the search for the next character
    ; slips in and finishes first; the earlier results are out of date then and aren't shown (they would cover the new ones, and the window height and list wouldn't match)
    static _searchGen := 0
    static _RunSearch() {
        gen := ++SearchWindow._searchGen
        text := SearchWindow.Input.Value
        if (SearchWindow.Mode = "actions") {
            filtered := []
            for action in SearchWindow.AllActions
                if (Trim(text) = "" || FuzzyMatcher.Score(text, action.Title) > 0)
                    filtered.Push(action)
            return SearchWindow.SetResults(filtered)
        }
        browse := FileSearchProvider.BrowsePath(text)                       ; browsing a folder: only the last part (the filter text) is highlighted, not the words in the path
        SearchWindow.HighlightText := IsObject(browse) ? browse.Filter : text
        results := SearchWindow.FileMode ? ProviderRegistry.SearchOnly(FileSearchProvider, text) : ProviderRegistry.Search(text)
        if (gen != SearchWindow._searchGen)                                 ; more characters were typed during the search: use the newer search's results
            return
        SearchWindow.SetResults(results)
    }

    ; Swaps in new results and adjusts the list and window height. Not interruptible (Critical): if another search slipped in halfway,
    ; the row counts of both would mix (a tall window with an empty list). Only the controls change here; painting happens afterwards as usual
    static SetResults(results) {
        Critical("On")
        SearchWindow._NoteResults(results.Length)
        SearchWindow.Results  := results
        SearchWindow.Selected := results.Length ? 1 : 0
        SearchWindow.Offset   := 0
        SearchWindow._Layout()
        Critical("Off")
        SearchWindow._Repaint()                                             ; repaint once what was painted during Critical (the custom-draw callback can't run then)
        if (SearchWindow._repaintLater = "")                                ; paint again after this input is handled: in case no result rows were painted at that moment (blank rows), they appear right after
            SearchWindow._repaintLater := () => SearchWindow._AfterResults()
        SetTimer(SearchWindow._repaintLater, -30)
    }
    static _repaintLater := ""

    ; 30 ms after the results changed: paint the list again (without erasing: each row fills itself, so an unchanged repaint is invisible and doesn't flicker),
    ; and check that the list's row count and window height match the results; if not, log and fix it (a tall window with an empty list)
    static _AfterResults() {
        DllCall("InvalidateRect", "Ptr", SearchWindow.List.Hwnd, "Ptr", 0, "Int", 0)
        SearchWindow._WatchPaint()
        if !SearchWindow.IsVisible()
            return
        rows := SearchWindow._VisibleCount(), client := Buffer(16)
        DllCall("GetClientRect", "Ptr", SearchWindow.Gui.Hwnd, "Ptr", client)
        height := NumGet(client, 12, "Int")
        if (SearchWindow.List.GetCount() = rows && height = SearchWindow._WindowHeight(rows))
            return
        if Logger.Enabled {
            Logger.Warn("SearchWindow: layout out of date (results " SearchWindow.Results.Length ", visible rows " rows ", list rows " SearchWindow.List.GetCount()
                . ", window height " height " instead of " SearchWindow._WindowHeight(rows) "), fixing it")
            for entry in SearchWindow._recent
                Logger.Debug("  recent results: " entry)
        }
        SearchWindow._shownRows := -1                                       ; make _Layout set the window height again
        SearchWindow._Layout()
    }

    ; To diagnose "the window is several rows high but the rows are blank": the last 10 result changes are kept in memory
    ; (time, input length, results, list rows) and written to the debug log when a row has no result to draw (at most once
    ; per show). Only the length of the input, never the text
    static _recent := [], _blankLogged := false
    static _NoteResults(count) {
        SearchWindow._recent.Push(FormatTime(, "HH:mm:ss") "." Format("{:03}", A_MSec) " input " StrLen(SearchWindow.Input.Value) " chars, mode " SearchWindow.Mode
            . (SearchWindow.FileMode ? " (files)" : "") " results " count " gen " SearchWindow._searchGen " list rows before " SearchWindow.List.GetCount())
        if (SearchWindow._recent.Length > 10)
            SearchWindow._recent.RemoveAt(1)
    }
    static _LogBlankRow(row, count) {
        if (SearchWindow._blankLogged || !Logger.Enabled)
            return
        SearchWindow._blankLogged := true
        Logger.Warn("SearchWindow: row " row " has no result (results " count ", offset " SearchWindow.Offset ", list rows " SearchWindow.List.GetCount()
            . ", window rows " SearchWindow._shownRows ", mode " SearchWindow.Mode ", input '" SubStr(SearchWindow.Input.Value, 1, 30) "')")
        for entry in SearchWindow._recent
            Logger.Debug("  recent results: " entry)
    }

    static _VisibleCount() {
        return Min(SearchWindow.Results.Length, SearchWindow.VisibleRows)
    }

    static _WindowHeight(visibleCount) {
        return SearchWindow._ContentHeight(visibleCount) + (SearchWindow._StatusShown(visibleCount) ? SearchWindow.StatusHeight : 0)
    }

    ; Height of the input box and result rows (without the status bar)
    static _ContentHeight(visibleCount) {
        pad := SearchWindow.Padding
        if !visibleCount
            return pad + SearchWindow.InputHeight + pad
        ; input box + gap + separator + result rows + bottom margin (as laid out in Create())
        return pad + SearchWindow.InputHeight + pad // 2 + 1 + visibleCount * SearchWindow.RowHeight + pad // 2
    }

    ; Called on every key press: with the same row count only the list is repainted, the window size and controls stay (resizing is much slower than repainting)
    static _Layout() {
        visibleCount := SearchWindow._VisibleCount()
        list := SearchWindow.List
        if (list.GetCount() != visibleCount) {
            list.Opt("-Redraw")
            list.Delete()
            Loop visibleCount
                list.Add("", "")
            list.Move(, , , visibleCount * SearchWindow.RowHeight)
            list.Opt("+Redraw")
        }
        if (list.Visible != (visibleCount > 0)) {
            list.Visible := visibleCount > 0
            SearchWindow.Separator.Visible := visibleCount > 0
        }
        SearchWindow._PlaceStatusBar(visibleCount)
        SearchWindow._Repaint()
        if (SearchWindow.IsVisible() && SearchWindow._shownRows != visibleCount) {
            SearchWindow.Gui.Show("NA w" SearchWindow.Width " h" SearchWindow._WindowHeight(visibleCount))
            SearchWindow._shownRows := visibleCount
        }
    }

    static _Repaint() {
        DllCall("InvalidateRect", "Ptr", SearchWindow.List.Hwnd, "Ptr", 0, "Int", 1)
        SearchWindow._WatchPaint()
        SearchWindow._UpdateStatusBar()                                     ; the results / selected row changed
    }

    ; Diagnoses "the window is several rows tall but the result rows are blank" (when the row count matches the results): about 100 ms after each list repaint request it checks
    ; whether the list has finished painting (no area waiting) without the custom-draw callback ever running (_paints unchanged), which means ListView painted it the default way:
    ; the rows' text is empty, so they come out blank. Then it logs (at most once per showing) and paints again, at most 3 extra times.
    ; A possible cause: during the short time an AutoHotkey thread can't be interrupted (e.g. WM_PAINT handled while reading an icon or waiting for Everything),
    ; the callback is skipped
    static _paints := 0, _paintsBefore := 0, _paintCheck := "", _paintTries := 0, _paintRetries := 0, _skipLogged := false
    static _WatchPaint() {
        SearchWindow._paintsBefore := SearchWindow._paints, SearchWindow._paintTries := 0, SearchWindow._paintRetries := 0
        if (SearchWindow._paintCheck = "")
            SearchWindow._paintCheck := () => SearchWindow._CheckPaint()
        SetTimer(SearchWindow._paintCheck, -100)
    }
    static _CheckPaint() {
        list := SearchWindow.List
        if (!SearchWindow.IsVisible() || !list.Visible || SearchWindow._paints != SearchWindow._paintsBefore)
            return
        if DllCall("GetUpdateRect", "Ptr", list.Hwnd, "Ptr", 0, "Int", 0) {   ; not painted yet (e.g. covered by another window): look again later, for at most 2 s
            if (++SearchWindow._paintTries < 20)
                SetTimer(SearchWindow._paintCheck, -100)
            return
        }
        if (Logger.Enabled && !SearchWindow._skipLogged) {
            SearchWindow._skipLogged := true
            Logger.Warn("SearchWindow: custom draw skipped (results " SearchWindow.Results.Length ", list rows " list.GetCount() ", mode " SearchWindow.Mode
                . (SearchWindow.FileMode ? " (files)" : "") ", input '" SubStr(SearchWindow.Input.Value, 1, 30) "'), repainting")
            for entry in SearchWindow._recent
                Logger.Debug("  recent results: " entry)
        }
        if (++SearchWindow._paintRetries > 3)
            return
        DllCall("InvalidateRect", "Ptr", list.Hwnd, "Ptr", 0, "Int", 0)
        SearchWindow._paintTries := 0
        SetTimer(SearchWindow._paintCheck, -100)
    }

    static SelectedItem() {
        if (SearchWindow.Selected < 1 || SearchWindow.Selected > SearchWindow.Results.Length)
            return ""
        return SearchWindow.Results[SearchWindow.Selected]
    }

    ; wrap: past the last result go back to the first (and the other way round), used by Tab cycling
    static MoveSelection(step, wrap := false) {
        count := SearchWindow.Results.Length
        if !count
            return
        if (wrap && (SearchWindow.Selected + step > count || SearchWindow.Selected + step < 1))
            SearchWindow.Selected := (step > 0) ? 1 : count
        else
            SearchWindow.Selected := Max(1, Min(count, SearchWindow.Selected + step))
        visible := SearchWindow.VisibleRows
        if (SearchWindow.Selected <= SearchWindow.Offset)
            SearchWindow.Offset := SearchWindow.Selected - 1
        else if (SearchWindow.Selected > SearchWindow.Offset + visible)
            SearchWindow.Offset := SearchWindow.Selected - visible
        SearchWindow._Repaint()
    }

    ;---------------------------------------------------------------------------
    ; Execute
    ;---------------------------------------------------------------------------
    static _Execute(modifier := "") {
        item := SearchWindow.SelectedItem()
        if !IsObject(item)
            return

        if (SearchWindow.Mode = "actions") {
            source := SearchWindow.ActionSource
            Knowledge.Record(SearchWindow.SavedQuery, source.Uid)
            RecentProvider.Remember(source)
            Usage.CountItem(source)
            SearchWindow.Marked := Map()
            SearchWindow.Hide("run")
            SearchWindow._SafeRun(() => item.OnRun.Call(source))
            return
        }

        if !item.Valid {
            if (item.AutoComplete != "")
                SearchWindow._SetInput(item.AutoComplete)
            return
        }
        Knowledge.Record(SearchWindow.Input.Value, item.Uid)
        RecentProvider.Remember(item)
        Usage.CountItem(item)
        SearchWindow.Hide("run")
        if (modifier = "")
            SearchWindow._SafeRun(() => ActionCatalog.RunDefault(item))
        else
            SearchWindow._SafeRun(() => ActionCatalog.RunModifier(item, modifier))
    }

    static _SafeRun(fn) {
        try {
            fn()
        } catch as e {
            Logger.Error("SearchWindow: action failed - " e.Message)
            MsgBox(e.Message, App.Name, 48)
        }
    }

    static _RunVisibleRow(row) {
        if (row > SearchWindow._VisibleCount())
            return
        SearchWindow.Selected := SearchWindow.Offset + row
        SearchWindow._Execute()
    }

    ;---------------------------------------------------------------------------
    ; Action panel
    ;---------------------------------------------------------------------------
    ; Opens the actions panel of an item directly (not in the search results, e.g. actions for the selection): Esc / ← close the window
    static ShowActions(item) {
        SearchWindow.Show("", false)
        if SearchWindow._OpenActionsFor(item)
            SearchWindow._actionsOnly := true
    }

    static _OpenActions() {
        return SearchWindow._OpenActionsFor(SearchWindow.SelectedItem())
    }

    static _OpenActionsFor(item) {
        if (!IsObject(item) || !item.Valid || SearchWindow.Mode = "actions")
            return false
        if SearchWindow.Marked.Count {                                      ; files are marked: act on all marked ones together
            paths := []
            for key, markedPath in SearchWindow.Marked
                paths.Push(markedPath)
            item := ResultItem(I18n.T("Search.MarkedItems", paths.Length), "", {Icon: "folder:"})
            actions := ResultActions.ListForMany(paths)
        } else
            actions := ResultActions.ListFor(item)
        if !actions.Length
            return false
        SearchWindow.SavedQuery := SearchWindow.Input.Value
        SearchWindow.ActionSource := item
        SearchWindow.AllActions := actions
        SearchWindow.Mode := "actions"
        Win.SetCueBanner(SearchWindow.Input.Hwnd, I18n.T("Search.ActionsFor", item.Title))
        SearchWindow._SetInput("")
        return true
    }

    static _CloseActions() {
        SearchWindow.Mode := "results"
        SearchWindow._UpdateCueBanner()
        SearchWindow._SetInput(SearchWindow.SavedQuery)
    }

    ;---------------------------------------------------------------------------
    ; File search mode (Space in the empty search box, as in Alfred)
    ; With text typed, Space with the caret moved to the start also enters it (the text stays, now searching files); in file search mode Backspace with the caret at the start
    ; goes back to normal search (the text stays too), as if the leading space had been deleted
    ;---------------------------------------------------------------------------
    static _CanEnterFileMode() {
        return SearchWindow.Mode = "results" && !SearchWindow.FileMode && (SearchWindow.Input.Value = "" || SearchWindow._CaretAtStart())
            && AppSettings.Feature("FileSearch")["SpacePrefix"] && ProviderRegistry.IsEnabled(FileSearchProvider)
    }

    static _SetFileMode(enabled) {
        SearchWindow.FileMode := enabled
        SearchWindow._UpdateCueBanner()
        SearchWindow._RunSearch()
    }

    ; Gray cue text: normal search / file search mode (the actions panel's cue is set in _OpenActions)
    ; Gray text in the empty box (the cue banner, so it shows only while the box is empty): "Search files..." in file
    ; search mode, otherwise this time's usage tip, or the placeholder when tips are off
    static _UpdateCueBanner() {
        if SearchWindow.FileMode
            text := I18n.T("Search.FilesPlaceholder")
        else
            text := (SearchWindow._tip != "") ? SearchWindow.FitCue(SearchWindow._tip) : SearchWindow.Placeholder()
        Win.SetCueBanner(SearchWindow.Input.Hwnd, text)
    }

    ; The cue banner uses the box's big font and just cuts off what doesn't fit: shorten it with "…" instead
    static FitCue(text) {
        hwnd := SearchWindow.Input.Hwnd, rect := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", hwnd, "Ptr", rect)
        margins := SendMessage(0xD4, 0, 0, hwnd)                            ; EM_GETMARGINS: left / right margin
        room := NumGet(rect, 8, "Int") - (margins & 0xFFFF) - (margins >> 16) - 2
        return SearchWindow.FitWidth(text, room, (part) => Win.TextExtent(hwnd, part))
    }

    ; The longest start of text (+ "…") that measure() says fits in room; the text itself when it fits
    static FitWidth(text, room, measure) {
        if (room <= 0 || measure(text) <= room)
            return text
        low := 0, high := StrLen(text)
        while (low < high) {
            middle := (low + high + 1) // 2
            if (measure(RTrim(SubStr(text, 1, middle)) "…") <= room)
                low := middle
            else
                high := middle - 1
        }
        return RTrim(SubStr(text, 1, low)) "…"
    }

    ; Top margin of the input box within the input band (height bandH): centered on the capital-letter height (ascent: baseline to font top, cap: capital height).
    ; Centering on the whole font height also counts room for descenders such as g and y below, so the text looked too high (#158, 2026-10; at 150%, 7 pixels free above,
    ; 17 below). Moved down by at most pad / 2 (there is margin below, and without results the window also keeps pad below the input band); centered as usual if the glyph can't be measured
    static InputTop(bandH, editH, pad, ascent, cap) {
        centered := (bandH - editH) // 2
        if (ascent <= 0 || cap <= 0 || cap > ascent)
            return centered
        return Max(centered, Min(Round(bandH / 2 - (ascent - cap / 2)), bandH - editH + pad // 2))
    }

    ; Gray text in the empty search box: the setting (Appearance.Placeholder) if given, "ALTRun Search" if empty;
    ; only spaces = show nothing
    static Placeholder() {
        text := AppSettings.Appearance.Get("Placeholder", "")
        return (text = "") ? I18n.T("Search.Placeholder") : Trim(text)
    }

    ;---------------------------------------------------------------------------
    ; Edit / delete (F3, Ctrl+Del, context menu, actions panel)
    ;---------------------------------------------------------------------------
    ; The item acted on: in the actions panel the item the panel was opened for, otherwise the selected result
    static _TargetItem() {
        return (SearchWindow.Mode = "actions") ? SearchWindow.ActionSource : SearchWindow.SelectedItem()
    }

    static _CurrentQuery() {
        return (SearchWindow.Mode = "actions") ? SearchWindow.SavedQuery : SearchWindow.Input.Value
    }

    ; show the dialog after the key message is handled, not waiting inside the OnMessage callback
    static _Later(fn) {
        SetTimer(fn, -1)
    }

    static _EditSelected() {
        item := SearchWindow._TargetItem()
        isFallback := IsObject(item) && item.Provider = ""                  ; the fallbacks when nothing matches (web search etc.)
        if (IsObject(item) && !isFallback && ResultActions.CanEdit(item))
            return SearchWindow._Later(() => SearchWindow.EditItem(item))
        query := Trim(SearchWindow._CurrentQuery())
        if (query != "" && (!IsObject(item) || isFallback))                 ; no results: a new command from the typed text
            return SearchWindow._Later(() => SearchWindow.NewCommand(query))
        App.Notify(I18n.T("Search.NotEditable"))
    }

    ; after the edit dialog closes, back to the same search, so the change shows right away
    static EditItem(item) {
        query := SearchWindow._CurrentQuery()
        SearchWindow.Hide()
        SearchWindow._SafeRun(() => ResultActions.EditItem(item))
        SearchWindow._Reopen(query)
    }

    static NewCommand(title) {
        SearchWindow.Hide()
        SearchWindow._SafeRun(() => CustomCommandProvider.Edit("", Map("Title", title)))
        SearchWindow._Reopen(title)
    }

    static DeleteItem(item) {
        query := SearchWindow._CurrentQuery()
        wasVisible := SearchWindow.IsVisible()
        SearchWindow._keepOpen := true
        answer := MsgBox(ResultActions.DeletePrompt(item), App.Name, "YesNo Icon! Default2" (wasVisible ? " Owner" SearchWindow.Gui.Hwnd : ""))
        SearchWindow._keepOpen := false
        if (answer = "Yes")
            SearchWindow._SafeRun(() => ResultActions.DeleteItem(item))
        selected := SearchWindow.Selected
        SearchWindow._Reopen(query)
        if (answer = "Yes" && selected > 1)
            SearchWindow.MoveSelection(selected - 2)                        ; select the row above the deleted one (MoveSelection is relative to row 1)
    }

    ; show the search window again without changing the "window before showing" (pasting etc. still goes back to that window)
    static _Reopen(query) {
        previous := App.PreviousWindow
        SearchWindow.Show(query)
        App.PreviousWindow := previous
    }

    static _OnRowContextMenu(row, x, y) {
        if (row < 1 || SearchWindow.Offset + row > SearchWindow.Results.Length)
            return
        if (SearchWindow.Mode = "results") {
            SearchWindow.Selected := SearchWindow.Offset + row
            SearchWindow._Repaint()
        }
        item := SearchWindow._TargetItem()
        if (!IsObject(item) || !item.Valid && !ResultActions.CanEdit(item))
            return
        if (SearchWindow.Mode = "actions")                                  ; right-click in the actions panel = run that item
            return SearchWindow._Execute()
        contextMenu := Menu()
        for action in ResultActions.ListFor(item) {
            label := StrReplace(action.Title, "&", "&&") (action.Subtitle != "" ? "`t" action.Subtitle : "")
            contextMenu.Add(label, SearchWindow._MenuHandler(action, item))
            if (A_Index = 1)
                contextMenu.Default := label
        }
        ; While the menu shows, AHK can't handle NM_CUSTOMDRAW and the list would paint in the default style (blue selection bar):
        ; remove ListView's own selection state first and paint in our style before showing the menu
        Critical("Off")
        SearchWindow.List.Modify(0, "-Select -Focus")
        SearchWindow._Repaint()
        DllCall("UpdateWindow", "Ptr", SearchWindow.List.Hwnd)
        SearchWindow._keepOpen := true
        contextMenu.Show(x, y)
        SearchWindow._keepOpen := false
        if (SearchWindow.IsVisible() && !WinActive("ahk_id " SearchWindow.Gui.Hwnd))
            try WinActivate("ahk_id " SearchWindow.Gui.Hwnd)
    }

    ; a separate method makes the closure, so each menu item keeps its own action
    static _MenuHandler(action, item) {
        return (*) => SearchWindow._RunMenuAction(action, item)
    }

    static _RunMenuAction(action, item) {
        SearchWindow._keepOpen := false
        Knowledge.Record(SearchWindow.Input.Value, item.Uid)
        RecentProvider.Remember(item)
        Usage.CountItem(item)
        SearchWindow.SavedQuery := SearchWindow.Input.Value
        SearchWindow.Hide()
        SearchWindow._SafeRun(() => action.OnRun.Call(item))
    }

    ;---------------------------------------------------------------------------
    ; Keyboard / mouse
    ;---------------------------------------------------------------------------
    static _IsOwnWindow(hwnd) {
        return IsObject(SearchWindow.Gui) && (hwnd = SearchWindow.Gui.Hwnd || hwnd = SearchWindow.Input.Hwnd || hwnd = SearchWindow.List.Hwnd)
    }

    static _OnKeyMessage(wParam, lParam, msg, hwnd) {
        result := SearchWindow._OnKeyDown(wParam, lParam, msg, hwnd)
        SearchWindow._keySeen := (result = "" && SearchWindow._IsOwnWindow(hwnd))   ; dispatched next: the subclass (if any) lets it through
        return result
    }

    ; While ALTRun shows a system dialog (MsgBox, FileSelect, DirSelect...), that dialog's modal message loop
    ; dispatches posted messages without calling OnMessage, so the keys never reach _OnKeyDown: typing still
    ; searches, but Enter, Esc, arrows... do nothing. Shown at such a time, the window subclasses the search box
    ; and the list and handles the keys in their window procedure instead (until it hides)
    static _keyCallback := 0, _keyHooked := false, _keySeen := false, _swallowChar := false

    ; A visible dialog (#32770) of this process: MsgBox, FileSelect, DirSelect... (about 0.04 ms: only dialogs are looked at)
    static InSystemDialog() {
        pid := DllCall("GetCurrentProcessId"), hwnd := 0
        while (hwnd := DllCall("FindWindowEx", "Ptr", 0, "Ptr", hwnd, "Str", "#32770", "Ptr", 0, "Ptr")) {
            DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "UInt*", &owner := 0)
            if (owner = pid && DllCall("IsWindowVisible", "Ptr", hwnd))
                return hwnd
        }
        return 0
    }

    static _HookKeysIfModal() {
        if SearchWindow.InSystemDialog()
            SearchWindow._HookKeys()
    }

    static _HookKeys() {
        if SearchWindow._keyHooked
            return
        if !SearchWindow._keyCallback                                       ; created once and kept: it may be running when the window hides
            SearchWindow._keyCallback := CallbackCreate((hwnd, msg, wParam, lParam, id, data) => SearchWindow._KeyProc(hwnd, msg, wParam, lParam), , 6)
        SearchWindow._keySeen := false, SearchWindow._swallowChar := false
        for hwnd in [SearchWindow.Input.Hwnd, SearchWindow.List.Hwnd]
            DllCall("comctl32\SetWindowSubclass", "Ptr", hwnd, "Ptr", SearchWindow._keyCallback, "UPtr", 1, "UPtr", 0)
        SearchWindow._keyHooked := true
        Logger.Debug("SearchWindow: keys handled in the window procedure (a system dialog is open)")
    }

    static _UnhookKeys() {
        if !SearchWindow._keyHooked
            return
        for hwnd in [SearchWindow.Input.Hwnd, SearchWindow.List.Hwnd]
            DllCall("comctl32\RemoveWindowSubclass", "Ptr", hwnd, "Ptr", SearchWindow._keyCallback, "UPtr", 1)
        SearchWindow._keyHooked := false
    }

    static _KeyProc(hwnd, msg, wParam, lParam) {
        if (msg = 0x100 || msg = 0x104) {                                   ; WM_KEYDOWN / WM_SYSKEYDOWN
            SearchWindow._swallowChar := false
            if SearchWindow._keySeen                                        ; OnMessage already saw it (no modal loop)
                SearchWindow._keySeen := false
            else if (SearchWindow._OnKeyDown(wParam, lParam, msg, hwnd) != "") {
                SearchWindow._swallowChar := true                           ; the WM_CHAR made from it (Enter, Esc, Space...) must not reach the box either
                return 0
            }
        } else if ((msg = 0x102 || msg = 0x106) && SearchWindow._swallowChar) {   ; WM_CHAR / WM_SYSCHAR
            SearchWindow._swallowChar := false
            return 0
        }
        return DllCall("comctl32\DefSubclassProc", "Ptr", hwnd, "UInt", msg, "UPtr", wParam, "Ptr", lParam, "Ptr")
    }

    ; 0 = the key is handled and not passed to the input box; empty = handled the normal way
    static _OnKeyDown(wParam, lParam, msg, hwnd) {
        if !SearchWindow._IsOwnWindow(hwnd)
            return
        vk := wParam
        ctrl  := GetKeyState("Ctrl")
        shift := GetKeyState("Shift")
        alt   := (msg = 0x104) || GetKeyState("Alt")
        actions := SearchWindow.Mode = "actions"

        switch vk {
            case 0x0D:                                                      ; Enter
                SearchWindow._Execute(ctrl ? "ctrl" : alt ? "alt" : "")
                return 0
            case 0x1B:                                                      ; Esc
                if (actions && SearchWindow._actionsOnly)
                    SearchWindow.Hide("Esc")
                else if actions
                    SearchWindow._CloseActions()
                else
                    SearchWindow.Hide("Esc")
                return 0
            case 0x20:                                                      ; Space: in the empty search box (or with the caret at the start), enter file search mode
                if (ctrl || alt)
                    return
                ; with all text selected (the restored last search) a space would replace it, so treat it as an empty search box
                if (!shift && !actions && !SearchWindow.FileMode && SearchWindow._AllTextSelected()) {
                    SearchWindow._SetInput("")
                    if SearchWindow._CanEnterFileMode() {
                        SearchWindow._SetFileMode(true)
                        return 0
                    }
                    return
                }
                if (!shift && SearchWindow._CanEnterFileMode()) {
                    SearchWindow._SetFileMode(true)
                    return 0
                }
                ; SpaceToRun (the 2.x option of the same name): with text typed, Space runs the selected result; Shift+Space types a space as usual
                if (!shift && AppSettings.General["SpaceToRun"] && SearchWindow.Input.Value != "" && IsObject(SearchWindow.SelectedItem())) {
                    SearchWindow._Execute()
                    return 0
                }
                return
            case 0x26:                                                      ; ↑ / Ctrl+↑ previous search in the history
                if ctrl
                    SearchWindow.RecallHistory(1)
                else
                    SearchWindow.MoveSelection(-1)
                return 0
            case 0x28:                                                      ; ↓ / Ctrl+↓ next search in the history
                if ctrl
                    SearchWindow.RecallHistory(-1)
                else
                    SearchWindow.MoveSelection(1)
                return 0
            case 0x21:                                                      ; PgUp
                SearchWindow.MoveSelection(-SearchWindow.VisibleRows)
                return 0
            case 0x22:                                                      ; PgDn
                SearchWindow.MoveSelection(SearchWindow.VisibleRows)
                return 0
            case 0x2D:                                                      ; Insert: mark / unmark (as in Total Commander) and move to the next row
                if (!actions && !ctrl && !shift && !alt) {
                    SearchWindow.ToggleMark()
                    return 0
                }
                return
            case 0x09:                                                      ; Tab: complete, or (TabKey = "Next") next / previous result; Ctrl+Tab always completes
                if (!ctrl && AppSettings.General["TabKey"] = "Next")
                    SearchWindow.MoveSelection(shift ? -1 : 1, true)
                else
                    SearchWindow._AutoComplete()
                return 0
            case 0x27:                                                      ; →
                ; Only a fresh →: holding → to move the caret doesn't open the actions panel by auto-repeat when the caret reaches the end (lParam bit 30 = already down)
                ; With text selected (e.g. the restored last search selected all), like a normal input box it only deselects and moves the caret to the end, so editing an expression doesn't open the panel
                if (!(lParam & 0x40000000) && SearchWindow._CaretAtEnd() && !SearchWindow._HasTextSelection() && SearchWindow._OpenActions())
                    return 0
                return
            case 0x25:                                                      ; ←
                if (actions && SearchWindow.Input.Value = "") {
                    if SearchWindow._actionsOnly
                        SearchWindow.Hide()
                    else
                        SearchWindow._CloseActions()
                    return 0
                }
                return
            case 0x70:                                                      ; F1
                SearchWindow.Hide()
                App.About()
                return 0
            case 0x72:                                                      ; F3
                SearchWindow._EditSelected()
                return 0
            case 0x71:                                                      ; F2
                SearchWindow.Hide()
                App.OpenPreferences()
                return 0
            case 0x73:                                                      ; F4
                SearchWindow.Hide()
                App.EditSettingsFile()
                return 0
            case 0x2E:                                                      ; Ctrl+Del
                if (ctrl && SearchWindow._CaretAtEnd() && IsObject(item := SearchWindow._TargetItem()) && ResultActions.CanDelete(item)) {
                    SearchWindow._Later(() => SearchWindow.DeleteItem(item))
                    return 0
                }
                return
            case 0x08:                                                      ; Backspace
                if (actions && SearchWindow.Input.Value = "") {
                    if SearchWindow._actionsOnly
                        SearchWindow.Hide()
                    else
                        SearchWindow._CloseActions()
                    return 0
                }
                ; pressed again after emptying the box in file search mode, or with the caret at the start: back to normal search
                if (SearchWindow.FileMode && (SearchWindow.Input.Value = "" || !actions && SearchWindow._CaretAtStart())) {
                    SearchWindow._SetFileMode(false)
                    return 0
                }
                if ctrl {                                                   ; Ctrl+Backspace: delete the previous word (a normal Edit control doesn't support it)
                    SearchWindow._DeletePreviousWord()
                    return 0
                }
                return
        }

        if !ctrl
            return
        switch vk {
            case 0x4E:                                                      ; Ctrl+N
                SearchWindow.MoveSelection(1)
                return 0
            case 0x50:                                                      ; Ctrl+P
                SearchWindow.MoveSelection(-1)
                return 0
            case 0x43:                                                      ; Ctrl+C
                if (shift || SearchWindow._HasTextSelection())
                    return
                SearchWindow._CopySelected()
                return 0
            case 0x4C:                                                      ; Ctrl+L
                if LargeType.Enabled()
                    SearchWindow._ShowLargeType()
                return 0
            case 0x45:                                                      ; Ctrl+E
                SearchWindow._EditFileSelected()
                return 0
            case 0xBC:                                                      ; Ctrl+,
                SearchWindow.Hide()
                App.OpenPreferences()
                return 0
        }
        if (vk >= 0x31 && vk <= 0x39) {                                     ; Ctrl+1..9
            SearchWindow._RunVisibleRow(vk - 0x30)
            return 0
        }
    }

    ; step: 1 = an older entry (Ctrl+↑), -1 = a newer one (Ctrl+↓). As soon as the text in the box changes it starts from the newest again
    ; (_OnInputChange resets HistoryIndex). Entries equal to the current text are skipped, so each press changes something
    static RecallHistory(step) {
        if (SearchWindow.Mode = "actions" || SearchWindow.FileMode)
            return
        history := Knowledge.History
        current := SearchWindow.Input.Value
        index := SearchWindow.HistoryIndex
        if (index = 0)
            SearchWindow._historyDraft := current
        Loop {
            index += step
            if (index < 0 || index > history.Length)
                return                                                      ; already the oldest / newest entry
            text := index ? history[index] : SearchWindow._historyDraft
            if (text != current || index = 0)
                break
        }
        SearchWindow._SetInput(text)
        SearchWindow.HistoryIndex := index                                  ; _SetInput doesn't go through the Change event: keep the history position here
    }

    ; Marks / unmarks the selected file or folder, then moves to the next row
    static ToggleMark() {
        item := SearchWindow.SelectedItem()
        if !(IsObject(item) && (item.Kind = "file" || item.Kind = "folder") && item.Arg != "")
            return false
        key := StrLower(item.Arg)
        if SearchWindow.Marked.Has(key)
            SearchWindow.Marked.Delete(key)
        else
            SearchWindow.Marked[key] := item.Arg
        if (SearchWindow.Selected < SearchWindow.Results.Length)
            SearchWindow.MoveSelection(1)
        else
            SearchWindow._Repaint()
        return true
    }

    static IsMarked(item) {
        return SearchWindow.Marked.Count && IsObject(item) && (item.Kind = "file" || item.Kind = "folder") && SearchWindow.Marked.Has(StrLower(item.Arg))
    }

    static _AutoComplete() {
        item := SearchWindow.SelectedItem()
        if (!IsObject(item) || SearchWindow.Mode = "actions")
            return
        text := (item.AutoComplete != "") ? item.AutoComplete : item.Title
        if (item.AutoComplete = "" && item.Kind = "folder" && RegExMatch(item.Arg, "^([A-Za-z]:\\|\\\\)"))
            text := RTrim(item.Arg, "\") "\"                                ; folders: Tab enters them to browse their contents
        SearchWindow._SetInput(text)
    }

    static _CopySelected() {
        item := SearchWindow.SelectedItem()
        if !IsObject(item)
            return
        SearchWindow.Hide()
        ActionCatalog.CopyText(item.CopyText())
    }

    ; When the selected file's type has an "Edit" verb, open it with that editor ("Edit" in the actions panel); other results do nothing
    static _EditFileSelected() {
        item := SearchWindow.SelectedItem()
        if !(IsObject(item) && item.Kind = "file" && IsObject(ActionCatalog.EditorFor(item.Arg)))
            return
        SearchWindow.Hide()
        SearchWindow._SafeRun(() => ActionCatalog.EditFile(item.Arg))
    }

    static _ShowLargeType() {
        item := SearchWindow.SelectedItem()
        text := IsObject(item) ? item.DisplayText() : SearchWindow.Input.Value
        SearchWindow.Hide()
        LargeType.Show(text)
    }

    ; Caret at the start, no text selected
    static _CaretAtStart() {
        return SendMessage(0xB0, 0, 0, SearchWindow.Input.Hwnd) = 0             ; EM_GETSEL: start and end both 0
    }

    static _CaretAtEnd() {
        selection := SendMessage(0xB0, 0, 0, SearchWindow.Input.Hwnd)        ; EM_GETSEL
        return ((selection >> 16) & 0xFFFF) >= StrLen(SearchWindow.Input.Value)
    }

    ; All text in the box is selected (after restoring the last search): a key press now replaces it
    static _AllTextSelected() {
        len := StrLen(SearchWindow.Input.Value)
        selection := SendMessage(0xB0, 0, 0, SearchWindow.Input.Hwnd)        ; EM_GETSEL
        return len && (selection & 0xFFFF) = 0 && ((selection >> 16) & 0xFFFF) >= len
    }

    static _HasTextSelection() {
        selection := SendMessage(0xB0, 0, 0, SearchWindow.Input.Hwnd)
        return (selection & 0xFFFF) != ((selection >> 16) & 0xFFFF)
    }

    static _DeletePreviousWord() {
        text := SearchWindow.Input.Value
        caret := (SendMessage(0xB0, 0, 0, SearchWindow.Input.Hwnd) >> 16) & 0xFFFF
        before := RegExReplace(SubStr(text, 1, caret), "\S*\s*$")
        SearchWindow.Input.Value := before SubStr(text, caret + 1)
        len := StrLen(before)
        SendMessage(0xB1, len, len, SearchWindow.Input.Hwnd)
        SearchWindow._OnInputChange()
    }

    static _OnRowClick(row) {
        if (row >= 1 && SearchWindow.Offset + row <= SearchWindow.Results.Length) {
            SearchWindow.Selected := SearchWindow.Offset + row
            SearchWindow._Repaint()
        }
        SearchWindow.Input.Focus()
    }

    static _OnRowDoubleClick(row) {
        if (row >= 1 && SearchWindow.Offset + row <= SearchWindow.Results.Length) {
            SearchWindow.Selected := SearchWindow.Offset + row
            SearchWindow._Execute()
        }
    }

    static _OnMouseWheel(wParam, lParam, msg, hwnd) {
        if !SearchWindow._IsOwnWindow(hwnd)
            return
        delta := (wParam >> 16) & 0xFFFF
        SearchWindow.MoveSelection(delta >= 0x8000 ? 1 : -1)                ; the high word is the signed scroll amount
        return 0
    }

    static _OnActivate(wParam, lParam, msg, hwnd) {
        if (!IsObject(SearchWindow.Gui) || hwnd != SearchWindow.Gui.Hwnd)
            return
        if ((wParam & 0xFFFF) = 0 && AppSettings.General["HideOnDeactivate"]) {
            if (SearchWindow._hideTimer = "")
                SearchWindow._hideTimer := () => SearchWindow._HideIfInactive()
            SetTimer(SearchWindow._hideTimer, -50)
        }
    }

    static _HideIfInactive() {
        if SearchWindow._keepOpen
            return
        if (SearchWindow.IsVisible() && !WinActive("ahk_id " SearchWindow.Gui.Hwnd))
            SearchWindow.Hide(Logger.Enabled ? "deactivated by " SearchWindow.ForegroundText() : "")
    }

    ;---------------------------------------------------------------------------
    ; Painting (NM_CUSTOMDRAW)
    ;---------------------------------------------------------------------------
    ; Not interruptible (Critical): when painting a row takes over about 15 ms (fonts loading the first time, a busy machine...), AutoHotkey lets timers run in the middle of the callback.
    ; An icon timer slipping in then asks for a list repaint (_Repaint) while the list is painting; that doesn't re-enter, so the repaint request stays pending and Windows keeps sending WM_PAINT:
    ; AutoHotkey keeps handling WM_PAINT and never returns to the thread painting the row, and the lower-priority WM_TIMER never gets a turn either -- the result rows stay blank, and timers, hotkeys and
    ; the tray don't respond, yet the window doesn't count as "not responding" (confirmed by a controlled experiment on Windows on GitHub, 2026-10). When not interruptible, timers wait until painting is done
    static _OnCustomDraw(lParam) {
        static CDDS_PREPAINT := 0x1, CDDS_ITEMPREPAINT := 0x10001
        static CDRF_DODEFAULT := 0x0, CDRF_SKIPDEFAULT := 0x4, CDRF_NOTIFYITEMDRAW := 0x20
        Critical("On")
        stage := NumGet(lParam, A_PtrSize * 3, "UInt")
        if (stage = CDDS_PREPAINT) {
            SearchWindow._paints++                                          ; the custom-draw callback ran (see _CheckPaint)
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
        SearchWindow._PaintRow(hdc, rect, item, index = SearchWindow.Selected, row + 1)
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
        gdi := SearchWindow._gdi
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
        if (item.Pinned && (hBadge := SearchWindow._PinBadge(badge := Max(Win.Scale(13), Round(iconSize * 0.55))))) {   ; pinned: a pin at the bottom right
            offset := iconSize - badge + Win.Scale(3)
            DllCall("DrawIconEx", "Ptr", hdc, "Int", left + pad + offset, "Int", iconY + offset, "Ptr", hBadge, "Int", badge, "Int", badge, "UInt", 0, "Ptr", 0, "UInt", 3)
        }

        textLeft := left + pad + iconSize + Win.Scale(12)
        shortcutW := Win.Scale(56)
        oldFont := DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["ShortcutFont"], "Ptr")
        textRight := right - pad - shortcutW - Win.Scale(8)

        if (visibleRow <= 9 && SearchWindow.Mode = "results") {             ; Ctrl+N hint on the right
            DllCall("SetTextColor", "Ptr", hdc, "UInt", selected ? gdi["SelectedShortcutColor"] : gdi["ShortcutColor"])
            SearchWindow._DrawText(hdc, "Ctrl+" visibleRow, right - pad - shortcutW, top, right - pad, bottom, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX)
        } else if (SearchWindow.Mode = "actions") {                         ; actions panel: shortcut or description on the right, with room for its actual width
            hintW := 0, gap := Win.Scale(12)
            if (item.Subtitle != "") {
                hintW := SearchWindow._TextWidth(hdc, item.Subtitle)
                DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["TitleFont"])
                titleW := SearchWindow._TextWidth(hdc, item.Title)
                DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["ShortcutFont"])
                hintW := SearchWindow.HintWidth(right - pad - textLeft, titleW, hintW, gap)
                DllCall("SetTextColor", "Ptr", hdc, "UInt", selected ? gdi["SelectedShortcutColor"] : gdi["ShortcutColor"])
                SearchWindow._DrawText(hdc, item.Subtitle, right - pad - hintW, top, right - pad, bottom, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS)
            }
            textRight := right - pad - (hintW ? hintW + gap : 0)
        }

        DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["TitleFont"])
        DllCall("SetTextColor", "Ptr", hdc, "UInt", selected ? gdi["SelectedTitleColor"] : gdi["TitleColor"])
        hasSubtitle := (item.Subtitle != "" && SearchWindow.Mode = "results")
        middle := top + Round(rowH * 0.54)
        ranges := (SearchWindow.Mode = "results") ? SearchWindow.HighlightRanges(item) : []
        if hasSubtitle
            SearchWindow._DrawTitle(hdc, item.Title, ranges, textLeft, top, textRight, middle, DT_BOTTOM | DT_SINGLELINE | DT_NOPREFIX, selected)
        else
            SearchWindow._DrawTitle(hdc, item.Title, ranges, textLeft, top, textRight, bottom, DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX, selected)

        if hasSubtitle {
            DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["SubtitleFont"])
            DllCall("SetTextColor", "Ptr", hdc, "UInt", selected ? gdi["SelectedSubtitleColor"] : gdi["SubtitleColor"])
            ellipsis := (item.Kind = "file" || item.Kind = "folder") ? DT_PATH_ELLIPSIS : DT_END_ELLIPSIS
            SearchWindow._DrawText(hdc, item.Subtitle, textLeft, middle + Win.Scale(2), textRight, bottom, DT_SINGLELINE | DT_NOPREFIX | ellipsis)
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
    static _DrawTitle(hdc, title, ranges, left, top, right, bottom, flags, selected) {
        static DT_VCENTER := 0x4, DT_BOTTOM := 0x8, ETO_CLIPPED := 0x4
        if (title = "")
            return
        gdi := SearchWindow._gdi
        available := right - left
        extents := SearchWindow._Extents(hdc, title)
        if (extents.Length && extents[extents.Length] > available) {
            ellipsis := SearchWindow._TextWidth(hdc, "…"), fit := 0
            while (fit < extents.Length && extents[fit + 1] <= available - ellipsis)
                fit++
            title := SubStr(title, 1, fit) "…"
            extents := SearchWindow._Extents(hdc, title)
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
    static _Extents(hdc, text) {
        count := StrLen(text), extents := []
        if !count
            return extents
        fit := 0, partial := Buffer(count * 4, 0), size := Buffer(8, 0)
        DllCall("GetTextExtentExPointW", "Ptr", hdc, "WStr", text, "Int", count, "Int", 0x7FFFFFFF, "Int*", &fit, "Ptr", partial, "Ptr", size)
        Loop count
            extents.Push(NumGet(partial, (A_Index - 1) * 4, "Int"))
        return extents
    }

    static _TextWidth(hdc, text) {
        size := Buffer(8, 0)
        DllCall("GetTextExtentPoint32W", "Ptr", hdc, "WStr", text, "Int", StrLen(text), "Ptr", size)
        return NumGet(size, 0, "Int")
    }

    static _DrawText(hdc, text, left, top, right, bottom, flags) {
        rect := Buffer(16)
        NumPut("Int", left, "Int", top, "Int", right, "Int", bottom, rect)
        DllCall("DrawTextW", "Ptr", hdc, "WStr", text, "Int", -1, "Ptr", rect, "UInt", flags)
    }
}
