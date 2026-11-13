;===============================================================================
; PreferencesWindow.ahk - the Preferences window (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like Alfred's Preferences: categories on the left, the category's settings on the right.
; All changes go to a copy of the settings (Working) first; the buttons work like in any Windows settings window:
;   OK       writes ALTRun.json, closes the window, and ALTRun reloads in the background (closes directly without changes)
;   Cancel   discards the changes (confirms first if there are any); Esc and the close button do the same
;   Apply    writes and reloads, and Preferences reopens on the same page at the same position; grayed out without changes
;   Help     (F1) opens the Wiki page for the current page
; Settings take effect after reloading (hotkeys, theme, indexes...), so the window flashes on "Apply".
;
; The pages are built in PrefsPages (one BuildXxx per page, the first time it is shown) with the form helpers in
; PrefsForm, which also holds the layout constants. Each setting takes one line there, e.g.
;   PrefsForm.Check("General.LaunchAtLogin", "Prefs.LaunchAtLogin", , , "Prefs.Group.Startup")
; Paths are keys in ALTRun.json joined with "."; like Alfred, a setting has a gray description below it when I18n
; has "<label key>.Desc".
;
; Usage:
;   PreferencesWindow.Show([page, x, y])
;   PreferencesWindow.AddCommands(paths)   "Send to" while Preferences is open: added to the Commands page
;===============================================================================

class PreferencesWindow {
    static Gui := "", Working := "", Pages := [], Binds := [], PageList := ""
    static _gapLines := []                                                  ; the y of the line in each gap of the page list (client pixels)
    static _lists := Map()                                                  ; setting path of a list page (e.g. "CustomCommands") -> the method that adds / edits items
    ; the order of the pages; search entries like "ALTRun Preferences: Appearance" come from it without opening the window.
    ; Commands and tests open a page by its key (or by its name, the key without "Prefs.Page.", on the command line), never by its number
    static PageKeys := ["Prefs.Page.General", "Prefs.Page.Window", "Prefs.Page.Keyboard", "Prefs.Page.Appearance", "Prefs.Page.Features"
        , "Prefs.Page.Applications", "Prefs.Page.FileSearch", "Prefs.Page.FileIndex", "Prefs.Page.WebSearch", "Prefs.Page.Calculator", "Prefs.Page.Clipboard", "Prefs.Page.Snippets"
        , "Prefs.Page.Commands", "Prefs.Page.Scripts", "Prefs.Page.Selection", "Prefs.Page.Hotkeys"
        , "Prefs.Page.QuickSwitch", "Prefs.Page.QSPanel", "Prefs.Page.DateStamp"
        , "Prefs.Page.DelayedStart", "Prefs.Page.Sync"
        , "Prefs.Page.Usage", "Prefs.Page.Advanced"]
    ; the first page of each group after the first (basics / search / commands and triggers / window tools / automation / usage and advanced):
    ; the page list has a gap with a line above each of them
    static GroupStarts := ["Prefs.Page.Applications", "Prefs.Page.Commands", "Prefs.Page.QuickSwitch", "Prefs.Page.DelayedStart", "Prefs.Page.Usage"]
    static _page := 0
    static _dirty := false, _ready := false, ApplyButton := ""
    static MinWindowWidth := 450                                            ; the search window: room for the status bar's four parts
    static _positionReset := false
    ; pageIndex: the page's key (e.g. "Prefs.Page.Advanced"), or its number
    static Show(pageIndex := 1, x := "", y := "") {
        if IsObject(PreferencesWindow.Gui) {
            index := IsInteger(pageIndex) ? pageIndex : PreferencesWindow._PageIndex(pageIndex)
            if (index >= 1 && index <= PreferencesWindow.Pages.Length)
                PreferencesWindow._SelectRow(index)
            WinActivate("ahk_id " PreferencesWindow.Gui.Hwnd)
            return
        }
        SearchWindow.Hide()
        PreferencesWindow.Working := PreferencesWindow.DeepCopy(AppSettings.Data)
        PreferencesWindow.ShowActualStartup(PreferencesWindow.Working)
        HotkeyBox.Sides := PreferencesWindow.GetPath(PreferencesWindow.Working, "General.HotkeySides") ? true : false
        PreferencesWindow.Binds := [], PrefsForm.Wider := Map(), PreferencesWindow._lists := Map(), PreferencesWindow._helpButton := ""
        PreferencesWindow._dirty := false, PreferencesWindow._ready := false, PreferencesWindow._positionReset := false, PreferencesWindow._page := 0
        ; Pages are built when needed (_BuildPage): opening builds only the page shown, the others when first clicked. Opening used to build all 19 pages,
        ; over 300 ms each time (a second or two on old PCs), and most pages are never looked at
        PreferencesWindow.Pages := []
        for index, key in PreferencesWindow.PageKeys
            PreferencesWindow.Pages.Push({Key: key, Name: I18n.T(key), Wiki: PreferencesWindow.WikiPage(key), Controls: [], Built: false
                , Builder: PreferencesWindow.PageBuilders[key]})

        g := Gui("-MinimizeBox", I18n.T("Prefs.Title"))
        g.SetFont("s9", ThemeManager.FontName())
        g.OnEvent("Close", (*) => PreferencesWindow.Cancel())                ; true = don't close (chose to keep editing)
        g.OnEvent("Escape", (*) => PreferencesWindow.Cancel())
        PreferencesWindow.Gui := g

        ; The page list on the left is built first: the Tab order is list -> the page's controls -> bottom buttons (pages built later have their controls moved before the bottom buttons, see _BuildPage)
        ; A ListView rather than a ListBox: an icon before each page (drawn with the system icon font, see NavIcons); 0x20 full-row select, 0x10000 double buffering
        pageList := g.AddListView("x12 y12 w160 h" (PrefsForm.ButtonY - 24) " -Hdr -Multi LV0x10020", [""])
        if !IsInteger(pageIndex)
            pageIndex := PreferencesWindow._PageIndex(pageIndex)
        pageIndex := Max(1, Min(pageIndex, PreferencesWindow.Pages.Length))
        PreferencesWindow._BuildPage(pageIndex)
        if !IsObject(PreferencesWindow.Gui)                                 ; the window was closed while a page was being built (see _BuildPage)
            return

        names := []
        for page in PreferencesWindow.Pages
            names.Push(page.Name)
        ; The pages are in groups with a gap and a line between them (GroupStarts). The rows of a report ListView can't have gaps, so the
        ; list shows tiles one row high and as wide as the list, each put at its place: hover, selection and ↑ ↓ Home End work as in a list.
        ; Where tile view isn't available (Wine) it stays a plain list without gaps
        if IsObject(tiles := PreferencesWindow._TileView(pageList)) {
            for index, name in names
                pageList.Add("Icon" index, name)
            ; the list's content (LVM_GETVIEWRECT) ends 4 px below the last tile (Windows 11), and when it is taller than the list both scroll bars appear:
            ; lay out in less height until it fits (the gaps take the rounding, so it can take a second try)
            height := tiles.Height, view := Buffer(16, 0)
            Loop 5 {
                layout := PreferencesWindow._PlaceTiles(pageList, tiles, height)
                if (!SendMessage(0x1022, 0, view, pageList) || (over := NumGet(view, 12, "Int") - tiles.Height) <= 0)
                    break
                height -= over
            }
            if (NavIcons.FontName() != "")
                pageList.SetImageList(NavIcons.ImageList(PreferencesWindow.PageKeys, Round(layout.RowH * 96 / A_ScreenDPI) - 2), 0)   ; tiles use the normal image list
            PreferencesWindow._gapLines := layout.Lines
        } else {
            ; The image list's height is the row height (a roomier sidebar). The rows are measured with an empty image list first, then the
            ; icons are drawn once at the height that makes the rows fill the list, leaving only a few pixels below the last one
            cellH := PreferencesWindow.PageRowHeight()
            sizer := (NavIcons.FontName() != "") ? NavIcons.EmptyList(cellH) : 0
            if sizer
                pageList.SetImageList(sizer, 1)
            for index, name in names
                pageList.Add("Icon" index, name)
            if sizer {
                old := pageList.SetImageList(NavIcons.ImageList(PreferencesWindow.PageKeys, cellH + PreferencesWindow._SpareRowHeight(pageList)), 1)
                if old
                    DllCall("comctl32\ImageList_Destroy", "Ptr", old)
            }
            pageList.ModifyCol(1, "AutoHdr")                                ; the only column fills the list width (a column narrower than the list would add a column separator on the right)
            PreferencesWindow._gapLines := []
        }
        DllCall("uxtheme\SetWindowTheme", "Ptr", pageList.Hwnd, "Str", "Explorer", "Ptr", 0)   ; the same hover / selection look as Explorer
        pageList.OnNotify(-101, (ctrl, lParam) => PreferencesWindow._OnPageRow(lParam))   ; LVN_ITEMCHANGED (the ItemSelect event sometimes doesn't arrive)
        pageList.OnNotify(-12, (ctrl, lParam) => PreferencesWindow._OnPageListDraw(lParam))   ; NM_CUSTOMDRAW: no dotted focus rectangle; the lines between the groups
        for code in [-2, -3, -5, -6]                                        ; NM_CLICK / NM_DBLCLK / NM_RCLICK / NM_RDBLCLK (a fast second click counts as a double-click)
            pageList.OnNotify(code, (ctrl, lParam) => PreferencesWindow._OnPageClick(lParam))
        PreferencesWindow.PageList := pageList

        buttonY := " y" PrefsForm.ButtonY
        help := g.AddButton("x12" buttonY " w85", I18n.T("Prefs.Help"))
        help.OnEvent("Click", (*) => PreferencesWindow.Help())
        PreferencesWindow._helpButton := help
        g.AddButton("x475" buttonY " w85 Default", I18n.T("Prefs.OK")).OnEvent("Click", (*) => PreferencesWindow.OK())
        g.AddButton("x570" buttonY " w85", I18n.T("Prefs.Cancel")).OnEvent("Click", (*) => PreferencesWindow.Cancel())
        apply := g.AddButton("x665" buttonY " w85 Disabled", I18n.T("Prefs.Apply"))
        apply.OnEvent("Click", (*) => PreferencesWindow.Apply())
        PreferencesWindow.ApplyButton := apply
        HotIfWinActive("ahk_id " g.Hwnd)
        Hotkey("F1", (*) => PreferencesWindow.Help())
        HotIfWinActive()

        PreferencesWindow._SelectRow(pageIndex)
        g.Show((IsInteger(x) && IsInteger(y) ? "x" x " y" y " " : "") "w765 h" (PrefsForm.ButtonY + 41))
        pageList.Focus()                                                    ; focus in the page list: ↑ ↓ switch pages, so the first page's settings can't be changed by accident
        ; values the program fills in when opening don't count as changes; start recording once the controls' notifications are handled
        SetTimer(() => (PreferencesWindow._ready := true), -300)
    }

    static _SelectRow(index) {
        PreferencesWindow.SelectPage(index)
        if IsObject(PreferencesWindow.Gui)
            PreferencesWindow.PageList.Modify(index, "Select Focus Vis")
    }

    ; A row was selected: show that page
    static _OnPageRow(lParam) {
        row := NumGet(lParam, 3 * A_PtrSize, "Int") + 1                     ; NMLISTVIEW: iItem (0-based), uNewState, uOldState
        newState := NumGet(lParam, 3 * A_PtrSize + 8, "UInt"), oldState := NumGet(lParam, 3 * A_PtrSize + 12, "UInt")
        if !((newState ^ oldState) & 0x2)                                   ; LVIS_SELECTED didn't change (e.g. only the focus changed)
            return
        if ((newState & 0x2) && row != PreferencesWindow._page)
            PreferencesWindow.SelectPage(row)
    }

    ; Switches the (still empty) page list to tiles -> {Width, Height} (its client size, measured while empty: no scroll bar); "" when tile view isn't available
    static _TileView(pageList) {
        client := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", pageList.Hwnd, "Ptr", client)
        if (SendMessage(0x108E, 4, 0, pageList) != 1)                       ; LVM_SETVIEW LV_VIEW_TILE (-1: not available)
            return ""
        return {Width: NumGet(client, 8, "Int"), Height: NumGet(client, 12, "Int")}
    }

    ; Lays out the tiles in height pixels: one row high and as wide as the list, each at its place (gaps before the group starts) -> PageLayout(...)
    static _PlaceTiles(pageList, tiles, height) {
        starts := []
        for key in PreferencesWindow.GroupStarts
            starts.Push(PreferencesWindow._PageIndex(key))
        layout := PreferencesWindow.PageLayout(height, PreferencesWindow.PageKeys.Length, starts, A_ScreenDPI / 96)
        info := Buffer(40, 0)                                               ; LVTILEVIEWINFO: LVTVIM_TILESIZE | LVTVIM_COLUMNS, LVTVIF_FIXEDSIZE, no second line
        NumPut("UInt", 40, "UInt", 0x3, "UInt", 0x3, "Int", tiles.Width, "Int", layout.RowH, "Int", 0, info)
        SendMessage(0x10A2, 0, info, pageList)                              ; LVM_SETTILEVIEWINFO
        for top in layout.Tops {
            point := Buffer(8, 0)
            NumPut("Int", 0, "Int", top, point)
            SendMessage(0x1031, A_Index - 1, point, pageList)               ; LVM_SETITEMPOSITION32
        }
        return layout
    }

    ; Where the rows of the page list go (pixels): rows as high as fit (at most 28 at 96 DPI) with a gap of 9 to 12 before each group start
    ; (row numbers in starts), the rest of the height below the last row. -> {RowH, Gap, Tops (each row's top), Lines (the y of each gap's line)}
    static PageLayout(clientH, rows, starts, scale := 1) {
        gap := Round(9 * scale)
        rowH := Min(Round(28 * scale), (clientH - starts.Length * gap) // rows)
        if starts.Length
            gap := Min(Round(12 * scale), Max(gap, (clientH - rows * rowH) // starts.Length))
        tops := [], lines := [], y := 0
        Loop rows {
            row := A_Index                                                  ; (A_Index inside the for loop below is the for loop's)
            for start in starts
                if (start = row)
                    lines.Push(y + gap // 2), y += gap
            tops.Push(y), y += rowH
        }
        return {RowH: rowH, Gap: gap, Tops: tops, Lines: lines}
    }

    ; The selected row is shown by its highlight only, without a dotted focus rectangle (like the navigation in Windows Settings). Windows decides whether to show focus rectangles
    ; by "whether the keyboard was used recently", so the first time Preferences opens the dots sometimes show and sometimes don't; CDIS_FOCUS is removed before each row is drawn
    ; The icon cell height of the plain page list (without tile view) to start with: 26 (a roomy sidebar), less when all pages wouldn't fit without scrolling
    ; (a row is about the cell plus 2; the list is ButtonY - 24 high with a 2 px border). _SpareRowHeight then measures the real rows (with NavIcons.EmptyList)
    static PageRowHeight() => Min(26, (PrefsForm.ButtonY - 24 - 4) // PreferencesWindow.PageKeys.Length - 2)

    ; How much taller (at 96 DPI) each row of the page list can be so that all rows still fit: the list's client height divided by the rows,
    ; minus the real row height (LVM_GETITEMRECT; Windows adds 1 or 2 pixels to the icon cell depending on the version and theme)
    static _SpareRowHeight(pageList) {
        rect := Buffer(16, 0)
        NumPut("Int", 0, rect, 0)                                           ; LVIR_BOUNDS
        if !SendMessage(0x100E, 0, rect, pageList)                          ; LVM_GETITEMRECT, row 0
            return 0
        rowH := NumGet(rect, 12, "Int") - NumGet(rect, 4, "Int")
        client := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", pageList.Hwnd, "Ptr", client)
        target := NumGet(client, 12, "Int") // pageList.GetCount()
        return (rowH > 0 && target > rowH) ? (target - rowH) * 96 // A_ScreenDPI : 0
    }

    ; No focus rectangle on the rows (see above); after painting, a line in each gap between the groups: a light gray between the
    ; window color and gray text (so it follows high contrast themes), inset like the rows' highlight
    static _OnPageListDraw(lParam) {
        stage := NumGet(lParam, 3 * A_PtrSize, "UInt")                      ; NMCUSTOMDRAW.dwDrawStage
        if (stage = 0x1)                                                    ; CDDS_PREPAINT
            return 0x30                                                     ; CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYPOSTPAINT
        if (stage = 0x10001) {                                              ; CDDS_ITEMPREPAINT
            offset := 6 * A_PtrSize + 16                                    ; NMCUSTOMDRAW.uItemState
            NumPut("UInt", NumGet(lParam, offset, "UInt") & ~0x10, lParam, offset)   ; CDIS_FOCUS
        } else if (stage = 0x2 && PreferencesWindow._gapLines.Length) {     ; CDDS_POSTPAINT
            hdc := NumGet(lParam, 4 * A_PtrSize, "Ptr")
            client := Buffer(16, 0)
            DllCall("GetClientRect", "Ptr", NumGet(lParam, 0, "Ptr"), "Ptr", client)
            window := DllCall("GetSysColor", "Int", 5, "UInt"), gray := DllCall("GetSysColor", "Int", 17, "UInt")   ; COLOR_WINDOW, COLOR_GRAYTEXT
            color := 0
            Loop 3
                shift := (A_Index - 1) * 8, color |= ((((window >> shift) & 0xFF) * 3 + ((gray >> shift) & 0xFF)) // 4) << shift
            pen := DllCall("CreatePen", "Int", 0, "Int", 1, "UInt", color, "Ptr")
            old := DllCall("SelectObject", "Ptr", hdc, "Ptr", pen, "Ptr")
            inset := Round(6 * A_ScreenDPI / 96)
            for y in PreferencesWindow._gapLines {
                DllCall("MoveToEx", "Ptr", hdc, "Int", inset, "Int", y, "Ptr", 0)
                DllCall("LineTo", "Ptr", hdc, "Int", NumGet(client, 8, "Int") - inset, "Int", y)
            }
            DllCall("SelectObject", "Ptr", hdc, "Ptr", old)
            DllCall("DeleteObject", "Ptr", pen)
        }
        return 0                                                            ; CDRF_DODEFAULT
    }

    ; A click on the blank area below the list (not on a row) makes ListView deselect: select the current page again after the mouse is released.
    ; Not right at deselecting: clicking another row also deselects the previous row first, and that row would flash
    static _OnPageClick(lParam) {
        if (NumGet(lParam, 3 * A_PtrSize, "Int") < 0 && !PreferencesWindow.PageList.GetNext())   ; NMITEMACTIVATE.iItem = -1
            PreferencesWindow.PageList.Modify(PreferencesWindow._page, "Select Focus")
    }

    ; a page key -> its number in the list (1 when there is no such page)
    static _PageIndex(nameKey) {
        for index, key in PreferencesWindow.PageKeys
            if (key = nameKey)
                return index
        return 1
    }

    ; A control was changed by the user: "Apply" becomes available and Cancel asks first; force: counts even right after opening (commands added by "Send to")
    static MarkDirty(force := false) {
        if ((!PreferencesWindow._ready && !force) || PreferencesWindow._dirty)
            return
        PreferencesWindow._dirty := true
        try PreferencesWindow.ApplyButton.Enabled := true
    }

    ; Switching pages only hides the old page's controls, shows the new one's and repaints just the page area on the right:
    ; the page list and bottom buttons aren't repainted, so switching doesn't flicker (no WM_SETREDRAW on the whole window: a top-level window that stops redrawing leaves its child controls unpainted)
    static SelectPage(pageIndex) {
        PreferencesWindow._BuildPage(pageIndex)
        if !IsObject(PreferencesWindow.Gui)
            return
        hwnd := PreferencesWindow.Gui.Hwnd
        previous := PreferencesWindow._page
        PreferencesWindow._page := pageIndex
        shown := PreferencesWindow.Pages[pageIndex]
        if (shown.HasOwnProp("OnShow") && IsObject(shown.OnShow))           ; a page that shows what other pages changed (Selection: the commands of each type)
            shown.OnShow.Call()
        for index, page in PreferencesWindow.Pages {
            if (index = pageIndex || index = previous || previous = 0)      ; right after building, every page has to be hidden once
                for ctrl in page.Controls
                    ctrl.Visible := (index = pageIndex)
        }
        ; check boxes and descriptions have a transparent background: repainting without erasing would draw the text twice, making it bold
        DllCall("RedrawWindow", "Ptr", hwnd, "Ptr", PreferencesWindow._PageArea(), "Ptr", 0, "UInt", 0x185)   ; RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW
    }

    ; The RECT of the page area (in the window's client area, right of the page list, above the bottom buttons)
    static _PageArea() {
        list := Buffer(16), client := Buffer(16), area := Buffer(16)
        DllCall("GetWindowRect", "Ptr", PreferencesWindow.PageList.Hwnd, "Ptr", list)
        DllCall("MapWindowPoints", "Ptr", 0, "Ptr", PreferencesWindow.Gui.Hwnd, "Ptr", list, "UInt", 2)   ; screen coordinates -> client coordinates
        DllCall("GetClientRect", "Ptr", PreferencesWindow.Gui.Hwnd, "Ptr", client)
        NumPut("Int", NumGet(list, 8, "Int"), "Int", 0, "Int", NumGet(client, 8, "Int"), "Int", NumGet(list, 12, "Int"), area)
        return area
    }

    ; The window's handle while Preferences is open, else 0. For results that arrive later (a file index rebuild runs in the background):
    ; Gui is still set for a moment after the window is destroyed (Close), and a timer can run in between; a dialog owned by the
    ; destroyed window failed with "Invalid window handle" (#290)
    static OpenHwnd() {
        try {
            if (IsObject(PreferencesWindow.Gui) && DllCall("IsWindow", "Ptr", hwnd := PreferencesWindow.Gui.Hwnd))
                return hwnd
        }
        return 0
    }

    static Close() {
        HotkeyBox.CancelActive()                                            ; closing the window while a hotkey is being recorded: end the recording first and restore ALTRun's hotkeys
        if IsObject(PreferencesWindow.Gui)
            PreferencesWindow.Gui.Destroy()
        PreferencesWindow.Gui := "", PreferencesWindow._lists := Map()
    }

    ; "Send to" while Preferences is open: switch to the Commands page and add the files / folders to the list; they are saved by OK / Apply
    ; (writing to the settings file directly would be overwritten by the next save). As without Preferences open: one opens the edit dialog
    ; (editing the existing command if there is one), several are added directly, skipping existing ones
    static AddCommands(paths) {
        PreferencesWindow.Show("Prefs.Page.Commands")
        list := PreferencesWindow._lists["CustomCommands"]
        if (paths.Length = 1) {
            existing := CustomCommandProvider.FindByTarget(paths[1], list.Items)
            if IsObject(existing) {
                App.Notify(I18n.T("Custom.Exists", existing["Title"]), 2500)
                list.Edit.Call(existing)
            } else {
                list.Insert.Call(CustomCommandProvider.FromPath(paths[1]))
            }
            return
        }
        added := 0, skipped := 0
        for target in paths {
            if IsObject(CustomCommandProvider.FindByTarget(target, list.Items)) {
                skipped += 1
                continue
            }
            list.Append.Call(CustomCommandProvider.FromPath(target))
            added += 1
        }
        App.Toast(I18n.T("Custom.AddedMany", added) (skipped ? " " I18n.T("Custom.SkippedExisting", skipped) : ""), 3000)
    }

    static OK() {
        if PreferencesWindow._dirty
            PreferencesWindow.Save(false)
        else
            PreferencesWindow.Close()
    }

    static Apply() {
        if PreferencesWindow._dirty
            PreferencesWindow.Save(true)
    }

    ; true = not closed (there were changes and the user chose to keep editing)
    static Cancel() {
        if (PreferencesWindow._dirty && MsgBox(I18n.T("Prefs.DiscardChanges"), I18n.T("Prefs.Title"), "YesNo Icon? Default2 Owner" PreferencesWindow.Gui.Hwnd) != "Yes")
            return true
        PreferencesWindow.Close()
        return false
    }

    static Help() {
        page := PreferencesWindow.Pages[Max(1, PreferencesWindow._page)]
        ActionCatalog.OpenUrl(HelpProvider.WikiPage(page.Wiki))
    }

    ; the Wiki page for each page (Help button / F1)
    static WikiPage(nameKey) {
        static pages := Map("Prefs.Page.General", "Configuration", "Prefs.Page.Window", "Usage", "Prefs.Page.Keyboard", "Hotkeys", "Prefs.Page.Appearance", "Themes"
                          , "Prefs.Page.Features", "Usage", "Prefs.Page.Applications", "Configuration", "Prefs.Page.FileSearch", "File-Search"
                          , "Prefs.Page.FileIndex", "File-Search", "Prefs.Page.Commands", "Custom-Commands", "Prefs.Page.Selection", "Usage", "Prefs.Page.Snippets", "Snippets"
                          , "Prefs.Page.Clipboard", "Clipboard", "Prefs.Page.WebSearch", "Web-Search", "Prefs.Page.Calculator", "Calculator"
                          , "Prefs.Page.Scripts", "Scripts", "Prefs.Page.Hotkeys", "Hotkeys", "Prefs.Page.DelayedStart", "Delayed-Start", "Prefs.Page.Sync", "Sync-Tasks", "Prefs.Page.QuickSwitch", "Quick-Switch"
                          , "Prefs.Page.QSPanel", "Quick-Switch-Panel", "Prefs.Page.DateStamp", "Date-Stamp", "Prefs.Page.Usage", "Usage")
        return pages.Has(nameKey) ? pages[nameKey] : "Configuration"
    }

    ; UAC shield right after the checkbox text: changing the option asks for UAC (imageres.dll icon 78, as for "Run as
    ; administrator" in the actions panel; IDI_SHIELD when missing). A checkbox can't hold an icon (BCM_SETSHIELD is for
    ; buttons only), so it is a picture beside it, not on top: one over the checkbox was painted over when it was enabled / disabled
    static _ShieldAfter(check) {
        size := Win.Scale(16), hIcon := 0
        DllCall("user32\PrivateExtractIconsW", "WStr", "imageres.dll", "Int", -78, "Int", size, "Int", size, "Ptr*", &hIcon, "Ptr", 0, "UInt", 1, "UInt", 0)
        if !hIcon
            hIcon := DllCall("LoadImageW", "Ptr", 0, "Ptr", 32518, "UInt", 1, "Int", size, "Int", size, "UInt", 0x8000, "Ptr")   ; IDI_SHIELD, IMAGE_ICON, LR_SHARED
        if !hIcon
            return
        check.GetPos(&x, &y, , &h)
        textW := Ceil(Win.TextExtent(check.Hwnd, check.Text) * 96 / A_ScreenDPI)
        check.Move(, , 18 + textW + 2)                                      ; box and gap + text: the checkbox ends where its text does
        saved := PrefsForm._y
        PrefsForm._y := y + (h - 16) // 2
        try PrefsForm.Add("Picture", "x" (x + 18 + textW + 6) " w16 h16", "HICON:" hIcon)
        PrefsForm._y := saved
    }

    ; Turning on "also in administrator windows" creates the scheduled task, turning it off deletes it (one UAC prompt each). If creating fails (UAC canceled), explain and turn the option off again;
    ; when ALTRun's folder can be changed without administrator rights, say what that means and turn it on only if the user still wants it.
    ; false = don't save (the user chose to go back and change it)
    static _ApplyCapsLockHelper(keyboard) {
        was := AppSettings.Keyboard.Has("CapsLockElevated") && AppSettings.Keyboard["CapsLockElevated"]
        now := keyboard.Has("CapsLockElevated") && keyboard["CapsLockElevated"]
        if (now && !CapsLockSwitch.IsMode(keyboard["CapsLock"]))
            keyboard["CapsLockElevated"] := now := 0                         ; the option has no effect while CapsLock works as usual
        if (now && !was && (folder := CapsLockHelper.WritableFolder()) != ""
            && MsgBox(I18n.T("Prefs.CapsLockElevatedWritable", folder), I18n.T("Prefs.Title"), "YesNo Icon! Default2 Owner" PreferencesWindow.Gui.Hwnd) != "Yes")
            keyboard["CapsLockElevated"] := now := 0                        ; the other settings are saved; the option stays off
        if (now && !was && !CapsLockHelper.Install()) {
            MsgBox(I18n.T("Prefs.CapsLockElevatedFailed"), I18n.T("Prefs.Title"), "Icon! Owner" PreferencesWindow.Gui.Hwnd)
            keyboard["CapsLockElevated"] := 0
        } else if (was && !now) {
            CapsLockHelper.Uninstall()
        }
        return true
    }

    ; "Launch ALTRun at login" as it really is: off when it was turned off in Task Manager → Startup apps (saving then keeps it off)
    static ShowActualStartup(working, offInTaskManager := "") {
        general := working["General"]
        if (offInTaskManager = "")
            offInTaskManager := general["LaunchAtLogin"] && App.LaunchOffInTaskManager()
        if offInTaskManager
            general["LaunchAtLogin"] := 0
    }

    ; Command line "-Preferences page [x y]" (reopening after Apply) -> {Page (a page key, or 1), X, Y}. page is the page's name, the key
    ; without "Prefs.Page." (e.g. "Hotkeys", "QSPanel"; any case); a number still works (in the current order)
    static ParseArgs(args) {
        value(i) => (args.Length >= i && IsInteger(args[i])) ? Integer(args[i]) : ""
        page := 1
        if (args.Length >= 2) {
            if IsInteger(args[2])
                page := (args[2] >= 1 && args[2] <= PreferencesWindow.PageKeys.Length) ? PreferencesWindow.PageKeys[Integer(args[2])] : 1
            else
                for key in PreferencesWindow.PageKeys
                    if (PreferencesWindow.PageName(key) = args[2])
                        page := key
        }
        return {Page: page, X: value(3), Y: value(4)}
    }

    ; A page's name on the command line: its key without "Prefs.Page." ("Prefs.Page.Hotkeys" -> "Hotkeys")
    static PageName(key) => RegExReplace(key, "^Prefs\.Page\.")

    ; The regular expression boxes (Exclude on Applications and Search Sources): the first one that isn't valid -> {Page (index in PageKeys),
    ; Label, Error}; "" when they all are
    static BadPattern(working) {
        for field in [["Applications", "Exclude", "Prefs.AppExclude", "Prefs.Page.Applications"], ["FileSearch", "ScopeExclude", "Prefs.ScopeExclude", "Prefs.Page.FileIndex"]]
            if ((problem := TextTools.RegexError(working["Features"][field[1]][field[2]])) != "")
                for index, key in PreferencesWindow.PageKeys
                    if (key = field[4])
                        return {Page: index, Label: I18n.T(field[3]), Error: problem}
        return ""
    }

    ; reopen: Apply = reload and reopen on the same page at the same position; OK = only reload in the background
    static Save(reopen := true) {
        for bind in PreferencesWindow.Binds
            PreferencesWindow.SetPath(PreferencesWindow.Working, bind.Path, bind.Read.Call())
        general := PreferencesWindow.Working["General"]
        appearance := PreferencesWindow.Working["Appearance"]
        AppSettings.ClampNumbers(PreferencesWindow.Working)                 ; e.g. an empty number box (0) or a width narrower than the status bar
        if (general["Hotkey"] = "")
            general["Hotkey"] := "!Space"
        if IsObject(bad := PreferencesWindow.BadPattern(PreferencesWindow.Working)) {     ; nothing is saved: a bad pattern would only be ignored later
            PreferencesWindow._SelectRow(bad.Page)
            MsgBox(I18n.T("Prefs.BadPattern", bad.Label, I18n.T(PreferencesWindow.PageKeys[bad.Page]), bad.Error), I18n.T("Prefs.Title"), "Icon! Owner" PreferencesWindow.Gui.Hwnd)
            return
        }
        clash := PreferencesWindow.DuplicateHotkey(PreferencesWindow._GlobalHotkeys(PreferencesWindow.Working))
        if (IsObject(clash) && MsgBox(I18n.T("Prefs.HotkeyClash", HotkeyBox.Label(clash.Key), clash.First, clash.Second)
                , I18n.T("Prefs.Title"), "YesNo Icon! Default2 Owner" PreferencesWindow.Gui.Hwnd) != "Yes")
            return
        if !PreferencesWindow._positionReset                                ; the search window was dragged after Preferences opened: use the new position
            appearance["Position"] := PreferencesWindow.DeepCopy(AppSettings.Appearance["Position"])

        if !PreferencesWindow._ApplyCapsLockHelper(PreferencesWindow.Working["Keyboard"])
            return
        stale := PreferencesWindow.StaleIndexes(AppSettings.Data, PreferencesWindow.Working)
        delayedBefore := AppSettings.Data["DelayedStart"]
        persistBefore := AppSettings.Feature("Clipboard")["Persist"]
        AppSettings.Data := PreferencesWindow.Working
        if !AppSettings.Save()
            return
        ; saving the clipboard history was turned off: the history saved so far stays on disk unless the user wants it gone (pinned entries too)
        if (persistBefore && !AppSettings.Feature("Clipboard")["Persist"] && (FileExist(ClipboardProvider.File) || DirExist(ClipboardProvider.Folder))
            && MsgBox(I18n.T("Prefs.ClipPersistDelete"), I18n.T("Prefs.Title"), "YesNo Icon? Default2 Owner" PreferencesWindow.Gui.Hwnd) = "Yes")
            ClipboardProvider.DeleteSaved()
        DelayedStart.AfterSave(delayedBefore, AppSettings.Data["DelayedStart"], PreferencesWindow.Gui.Hwnd)   ; the Task Scheduler tasks follow the list
        if general["LaunchAtLogin"]                                         ; ticked: on in Task Manager → Startup apps too
            App.TurnOnInTaskManager()
        if stale.Apps                                                       ; the index scope changed: rescan in the background after restarting (the old index is used until then)
            ApplicationProvider.ExpireCache()
        if stale.Files
            FileIndex.ExpireCache()
        page := PreferencesWindow.PageName(PreferencesWindow.Pages[Max(1, PreferencesWindow._page)].Key)
        WinGetPos(&x, &y, , , PreferencesWindow.Gui.Hwnd)
        PreferencesWindow.Close()
        App.Restart(reopen ? "-Preferences " page " " x " " y : "-Reloaded")
    }

    ;---------------------------------------------------------------------------
    ; Pages
    ;---------------------------------------------------------------------------
    ; how each page is built (PrefsPages.BuildXxx), by page key
    static PageBuilders := Map("Prefs.Page.General", "BuildGeneral", "Prefs.Page.Window", "BuildWindow", "Prefs.Page.Keyboard", "BuildKeyboard"
        , "Prefs.Page.Appearance", "BuildAppearance", "Prefs.Page.Features", "BuildFeatures", "Prefs.Page.Applications", "BuildApplications"
        , "Prefs.Page.FileSearch", "BuildFileSearch", "Prefs.Page.FileIndex", "BuildFileIndex", "Prefs.Page.WebSearch", "BuildWebSearch"
        , "Prefs.Page.Calculator", "BuildCalculator", "Prefs.Page.Clipboard", "BuildClipboard", "Prefs.Page.Snippets", "BuildSnippets"
        , "Prefs.Page.Commands", "BuildCommands", "Prefs.Page.Scripts", "BuildScripts", "Prefs.Page.Selection", "BuildSelection", "Prefs.Page.Hotkeys", "BuildHotkeys"
        , "Prefs.Page.QuickSwitch", "BuildQuickSwitch", "Prefs.Page.QSPanel", "BuildQuickSwitchPanel", "Prefs.Page.DateStamp", "BuildDateStamp"
        , "Prefs.Page.DelayedStart", "BuildDelayedStart", "Prefs.Page.Sync", "BuildSync", "Prefs.Page.Usage", "BuildUsage", "Prefs.Page.Advanced", "BuildAdvanced")
    static _helpButton := ""
    ; Builds every page (for the tests that check each page's layout)
    static BuildAll() {
        Loop PreferencesWindow.Pages.Length
            PreferencesWindow._BuildPage(A_Index)
    }

    ; Builds a page (does nothing if already built). Each page has a fixed left column width (set for the Chinese / English labels); when another language's labels are longer and don't fit,
    ; the left column is widened to the actual text width and the page built again (the controls from the first build are hidden and unused; the width is remembered so later openings build once).
    ; Building can be interrupted by another thread (Sleep, tray menu, Esc / close button); after the window closes Gui is "": add no more controls and return false
    static _BuildPage(index) {
        if !IsObject(PreferencesWindow.Gui)
            return false
        page := PreferencesWindow.Pages[index]
        if page.Built
            return false
        wasReady := PreferencesWindow._ready
        PreferencesWindow._ready := false                                   ; values the program fills in don't count as changes (as right after opening)
        Loop 2 {
            PrefsForm.Wider := Map(), binds := PreferencesWindow.Binds.Length
            try PrefsPages.%page.Builder%()
            catch as e {
                if IsObject(PreferencesWindow.Gui)
                    throw e
                Logger.Debug("PreferencesWindow: closed while building " page.Key)
                return false
            }
            if (!PrefsForm.Wider.Has(page.Key) || A_Index = 2)
                break
            PrefsForm.LabelWidths[I18n.Lang " " page.Key] := PrefsForm.Wider[page.Key]
            for ctrl in page.Controls
                ctrl.Visible := false, ctrl.Enabled := false
            if (PreferencesWindow.Binds.Length > binds)
                PreferencesWindow.Binds.RemoveAt(binds + 1, PreferencesWindow.Binds.Length - binds)
        }
        if !IsObject(PreferencesWindow.Gui)
            return false
        if IsObject(PreferencesWindow._helpButton)                          ; the window is already built: new controls move before the bottom buttons (Tab order)
            PreferencesWindow._BeforeButtons(page.Controls)
        if wasReady {                                                       ; handle the notifications the controls sent while building (Change events) first, then start recording changes
            Sleep(-1)
            PreferencesWindow._ready := true
        }
        if !IsObject(PreferencesWindow.Gui)                                 ; the window was closed during Sleep
            return false
        return true
    }

    ; Puts the controls in order before the "Help" button (Tab follows the controls' Z order: list -> page -> bottom buttons)
    static _BeforeButtons(controls) {
        after := DllCall("GetWindow", "Ptr", PreferencesWindow._helpButton.Hwnd, "UInt", 3, "Ptr")   ; GW_HWNDPREV
        for ctrl in controls {
            DllCall("SetWindowPos", "Ptr", ctrl.Hwnd, "Ptr", after, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x13)   ; NOSIZE | NOMOVE | NOACTIVATE
            after := ctrl.Hwnd
        }
    }

    ;---------------------------------------------------------------------------
    ; Data helpers (pure functions, with unit tests)
    ;---------------------------------------------------------------------------
    ; Every hotkey -> [{Key, Label, WinTitle, Entry}...]: the show hotkey, the second show hotkey, the selection hotkey, clipboard history
    ; (when on), and each custom hotkey (Entry = its Map, named as in the Hotkeys list; WinTitle = the windows it works in, "" = all).
    ; Quick switch and date stamp only work in particular windows (two Ctrl+D are fine), so they aren't included.
    ; value(path): reads a setting; Preferences passes one reading its boxes (BoundValue), which may not be applied yet
    static AllHotkeys(data, value := "") {
        if !HasMethod(value)
            value := (path) => PreferencesWindow.GetPath(data, path)
        list := []
        add(key, label, winTitle := "", entry := "") => (Trim(key) != "") ? list.Push({Key: key, Label: label, WinTitle: winTitle, Entry: entry}) : ""
        add(value("General.Hotkey"), I18n.T("Prefs.Hotkey"))
        add(value("General.SecondaryHotkey"), I18n.T("Prefs.SecondaryHotkey"))
        add(value("General.SelectionHotkey"), I18n.T("Prefs.SelectionHotkey"))
        if value("Features.Clipboard.Enabled")
            add(value("Features.Clipboard.Hotkey"), I18n.T("Prefs.Page.Clipboard"))
        commands := (data.Has("CustomCommands") && data["CustomCommands"] is Array) ? data["CustomCommands"] : []
        if (data.Has("Hotkeys") && data["Hotkeys"] is Array)
            for entry in data["Hotkeys"]
                if (entry is Map && entry.Has("Key"))
                    add(entry["Key"], I18n.T("Prefs.Page.Hotkeys") " (" HotkeyEditor.ListText(entry, commands) ")", entry.Has("WinTitle") ? Trim(entry["WinTitle"]) : "", entry)
        return list
    }

    ; Global hotkeys (not tied to a window) -> [{Key, Label, WinTitle, Entry}...] (see AllHotkeys)
    static _GlobalHotkeys(data) {
        list := []
        for each in PreferencesWindow.AllHotkeys(data)
            if (each.WinTitle = "")
                list.Push(each)
        return list
    }

    ; A setting as Preferences shows it now: the box bound to it (changed but not applied yet), otherwise Working
    static BoundValue(path) {
        for bind in PreferencesWindow.Binds
            if (bind.Path = path)
                return bind.Read.Call()
        return PreferencesWindow.GetPath(PreferencesWindow.Working, path)
    }

    ; The first duplicate hotkey -> {Key, First, Second} (the names of both); "" without duplicates. ~ * $ don't change whether it is the same key.
    ; Two hotkeys clash when one key press triggers both: Alt+Space clashes with LAlt+Space (either side includes the left one), LAlt+Space
    ; and RAlt+Space don't
    static DuplicateHotkey(list) {
        seen := []
        for entry in list {
            parts := PreferencesWindow.HotkeyParts(entry.Key)
            for before in seen
                if PreferencesWindow.SameKeyPress(before.Parts, parts)
                    return {Key: entry.Key, First: before.Label, Second: entry.Label}
            seen.Push({Parts: parts, Label: entry.Label})
        }
        return ""
    }

    ; "~<!^C" -> {Key: "c", Mods: Map("^", "", "!", "<")}: the key in lower case and each modifier with its side ("<" left, ">" right, "" either)
    static HotkeyParts(hk) {
        hk := StrLower(RegExReplace(Trim(hk), "^[~*$]+")), mods := Map(), side := "", i := 1
        while (i < StrLen(hk) && InStr("<>^!+#", c := SubStr(hk, i, 1))) {
            if InStr("<>", c)
                side := c
            else
                mods[c] := side, side := ""
            i++
        }
        return {Key: SubStr(hk, i), Mods: mods}
    }

    ; Whether two HotkeyParts are one key press (a modifier without a side matches either side)
    static SameKeyPress(a, b) {
        if (a.Key != b.Key || a.Mods.Count != b.Mods.Count)
            return false
        for m, side in a.Mods
            if !(b.Mods.Has(m) && (side = "" || b.Mods[m] = "" || side = b.Mods[m]))
                return false
        return true
    }

    ; Pure function: whether the scan scope of the application index / file index changed between before and after saving -> {Apps, Files}
    static StaleIndexes(before, after) {
        changed(group, keys) {
            for key in keys
                if (JSON.Stringify(PreferencesWindow.GetPath(before, group "." key)) != JSON.Stringify(PreferencesWindow.GetPath(after, group "." key)))
                    return true
            return false
        }
        return {Apps: changed("Features.Applications", ["Folders", "FileTypes", "Exclude", "Depth"])
              , Files: changed("Features.FileSearch", ["ScopeFolders", "ScopeDepth", "ScopeExclude", "MaxEntries"])}
    }

    static DeepCopy(value) {
        return JSON.Parse(JSON.Stringify(value))
    }

    static GetPath(data, path) {
        node := data
        for key in StrSplit(path, ".") {
            if !(node is Map) || !node.Has(key)
                return ""
            node := node[key]
        }
        return node
    }

    static SetPath(data, path, value) {
        keys := StrSplit(path, ".")
        node := data
        Loop keys.Length - 1 {
            key := keys[A_Index]
            if !(node.Has(key) && node[key] is Map)
                node[key] := Map()
            node := node[key]
        }
        node[keys[keys.Length]] := value
    }
}
