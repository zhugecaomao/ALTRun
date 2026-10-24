;===============================================================================
; SearchWindow.ahk - the Alfred-style search window (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Layout: a borderless window with rounded corners, a large input box at the top and the result list below (each row: icon +
; title + gray subtitle + a Ctrl+N hint on the right); the window height follows the number of results.
;
; The result list is a ListView that draws every row itself (NM_CUSTOMDRAW); it only holds the rows currently visible
; (at most VisibleRows), and Offset does the scrolling, so there is never a scroll bar.
;
; Keyboard (all handled in SearchKeys.OnKeyDown, catching WM_KEYDOWN / WM_SYSKEYDOWN with OnMessage):
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
; Helpers in their own files: SearchKeys (keyboard), ResultRows (draws the rows), SearchStatusBar, TypeAhead (keys typed while showing),
; SearchInputLanguage (English input while shown), ResultActions (actions panel, edit, delete).
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
    static _searchTimer := "", _hideTimer := ""
    static _shownRows := -1                                                 ; how many result rows the window is currently sized for
    static _tip := ""                                                       ; the usage tip for this showing (HelpProvider.NextTip)
    static LastArea := "", HiddenAt := 0                                    ; work area of the screen at the last hide, and its time (for Hud)
    static _last := ""                                                      ; the search at the last hide {Text, FileMode, Selected} (KeepLastQuery)
    static _keepOpen := false                                               ; don't hide on focus loss during a context menu / delete confirmation

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
        ResultRows.CreateGdi()

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
        list.OnNotify(-12, (ctrl, lParam) => ResultRows.OnCustomDraw(lParam))       ; NM_CUSTOMDRAW
        list.OnEvent("Click", (ctrl, row) => SearchWindow._OnRowClick(row))
        list.OnEvent("DoubleClick", (ctrl, row) => SearchWindow._OnRowDoubleClick(row))
        list.OnEvent("ContextMenu", (ctrl, row, isRightClick, x, y) => SearchWindow._OnRowContextMenu(row, x, y))

        g.OnEvent("Close", (*) => SearchWindow.Hide())

        SearchWindow.Gui := g, SearchWindow.Input := searchBox, SearchWindow.List := list, SearchWindow.Separator := separator
        SearchStatusBar.Create(g, w, pad)

        OnMessage(0x100, (p*) => SearchKeys.OnKeyMessage(p*))               ; WM_KEYDOWN
        OnMessage(0x104, (p*) => SearchKeys.OnKeyMessage(p*))               ; WM_SYSKEYDOWN (Alt+...)
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
    ; catchTyping: characters typed while showing are kept and added to the search box afterwards (see TypeAhead)
    static Show(text := "", catchTyping := true) {
        if !IsObject(SearchWindow.Gui)
            return
        if SearchWindow._ShouldSkipShow()
            return
        if (SearchWindow._prepareTimer != "")
            SetTimer(SearchWindow._prepareTimer, 0)
        if catchTyping
            TypeAhead.Start()
        SearchWindow._showing := true, SearchWindow._showStarted := A_TickCount, SearchWindow._showStep := "start"
        try SearchWindow._Show(text)
        finally {
            SearchWindow._showing := false
            typed := TypeAhead.Take()
        }
        SearchWindow._ApplyTypeAhead(typed)
    }

    ; Treats the characters typed while showing as typed in the search box: they replace selected text (the restored last search); a leading space enters file search mode
    static _ApplyTypeAhead(typed) {
        if (typed.Text = "" && typed.Key = "")
            return
        if Logger.Enabled
            Logger.Debug("SearchWindow: typed while showing " TypeAhead.Summary(typed))
        if (typed.Key = "Escape")
            return SearchWindow.Hide("Esc")
        text := typed.Text
        if (text != "") {
            if (SubStr(text, 1, 1) = " " && SearchWindow.Mode = "results" && !SearchWindow.FileMode) {
                if SearchKeys.AllTextSelected()
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
        SearchStatusBar.Update()
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
        SearchKeys.HookIfModal()
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
        SearchInputLanguage.SwitchOnShow(wasVisible)
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
            SearchInputLanguage.Restore()
            if AppSettings.Appearance["Fade"]
                WindowFade.Out(SearchWindow.Gui.Hwnd)                       ; the window hides at once as usual, leaving a shadow that fades out
            SearchWindow.Gui.Hide()
            SearchKeys.Unhook()
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
        return SearchWindow._ContentHeight(visibleCount) + (SearchStatusBar.Shown(visibleCount) ? SearchStatusBar.Height : 0)
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
        SearchStatusBar.Place(visibleCount)
        SearchWindow._Repaint()
        if (SearchWindow.IsVisible() && SearchWindow._shownRows != visibleCount) {
            SearchWindow.Gui.Show("NA w" SearchWindow.Width " h" SearchWindow._WindowHeight(visibleCount))
            SearchWindow._shownRows := visibleCount
        }
    }

    static _Repaint() {
        DllCall("InvalidateRect", "Ptr", SearchWindow.List.Hwnd, "Ptr", 0, "Int", 1)
        SearchWindow._WatchPaint()
        SearchStatusBar.Update()                                            ; the results / selected row changed
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
        return SearchWindow.Mode = "results" && !SearchWindow.FileMode && (SearchWindow.Input.Value = "" || SearchKeys.CaretAtStart())
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
    ; Keyboard actions (keys: see SearchKeys) / mouse
    ;---------------------------------------------------------------------------
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
        if !SearchKeys.IsOwnWindow(hwnd)
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

}
