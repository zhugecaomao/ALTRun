;===============================================================================
; QuickSwitchPanel.ahk - Quick Switch's folder panel below open / save dialogs (AutoHotkey v2)
;-------------------------------------------------------------------------------
; The search box + folder list attached below a file dialog: building, drawing (NM_CUSTOMDRAW, ALTRun's theme colors),
; placing it next to the dialog, searching (the list first, then Everything / the built-in index) and its keys.
; Part of QuickSwitch (Src\Extensions\QuickSwitch.ahk), which shows and hides it as dialogs come and go.
;
; Usage:
;   QuickSwitchPanel._UpdatePanel(dialog) / HidePanel()     from QuickSwitch's window watching
;   QuickSwitchPanel.JumpTo(folder)                          jump the dialog the panel belongs to
;===============================================================================

class QuickSwitchPanel {
    static _panel := "", _panelFor := 0, _panelFolders := [], _panelPos := ""

    ;---------------------------------------------------------------------------
    ; The folder panel (like Listary's Quick Switch window): attached below the dialog, as wide as it;
    ; the search box on top (filters the list first, then finds folders and files with Everything / the built-in index), the list below.
    ; Clicking an item or Enter in the search box jumps there (files: see PickFile); Esc goes back to the dialog
    ;---------------------------------------------------------------------------
    static _UpdatePanel(dialog) {
        if !dialog
            return QuickSwitchPanel.HidePanel()
        if (dialog != QuickSwitchPanel._panelFor) {                              ; a new dialog, or back from another window: list the folders again
            if !QuickSwitch._PanelEnabled()
                return QuickSwitchPanel.HidePanel()
            QuickSwitchPanel._BuildPanel(dialog)
        }
        QuickSwitchPanel._PlacePanel(dialog)
    }

    static HidePanel() {
        if IsObject(QuickSwitchPanel._panel)
            try QuickSwitchPanel._panel.Destroy()
        QuickSwitchPanel._panel := "", QuickSwitchPanel._panelFor := 0, QuickSwitchPanel._panelPos := "", QuickSwitchPanel._panelRows := -1
        QuickSwitchPanel._panelSearch := "", QuickSwitchPanel._panelList := "", QuickSwitchPanel._panelLine := "", QuickSwitchPanel._panelStar := ""   ; the controls are destroyed: no more references
        QuickSwitchPanel._searchHwnd := 0, QuickSwitchPanel._listHwnd := 0
    }

    static PanelRows := 8

    static _BuildPanel(dialog) {
        QuickSwitchPanel.HidePanel()
        QuickSwitchPanel._panelFor := dialog
        QuickSwitchPanel._panelBase := QuickSwitchFavorites.MenuFolders(false)
        rect := QuickSwitchPanel._FrameRect(dialog)
        width := IsObject(rect) ? QuickSwitchPanel._PanelWidthAt(rect) : QuickSwitchPanel.PanelWidthFor(600)
        colors := QuickSwitchPanel.PanelColors()                                 ; follows ALTRun's theme (light / dark...)
        ; -DPIScale: the panel is as wide as the dialog, in the dialog's real pixels; the row height is measured in pixels too. Letting AHK scale by DPI again
        ; would make the panel half again as wide as the dialog at 150%, with an empty area below. Fixed margins are converted with Win.Scale
        panel := Gui("-Caption +ToolWindow +AlwaysOnTop +Border -DPIScale", "ALTRun Quick Switch")
        panel.MarginX := 0, panel.MarginY := 0, panel.BackColor := colors.Background
        panel.SetFont("s10 c" colors.Text, QuickSwitchPanel._FontName())
        search := panel.AddEdit("x" Win.Scale(10) " y" Win.Scale(8) " w" QuickSwitchPanel.SearchWidth(width) " r1 -Multi -E0x200 Background" colors.Background)
        search.GetPos(, &searchY, , &searchH)
        star := panel.AddPicture("x" QuickSwitchPanel.StarX(width) " y" (searchY + (searchH - QuickSwitchPanel.StarSize()) // 2) " w" QuickSwitchPanel.StarSize() " h" QuickSwitchPanel.StarSize() " Icon44", "shell32.dll")
        star.OnEvent("Click", (*) => QuickSwitchFavorites.ShowFavoritesMenu())       ; the favorite folders (like the favorites button of Listary's Quick Switch window)
        Win.AddTooltip(star.Hwnd, I18n.T("QuickSwitch.FavoritesButton") "  (Alt+↓, F4)")
        hint := (QuickSwitch.Options.Has("MenuHotkey") && QuickSwitch.Options["MenuHotkey"] != "") ? "  (" Win.HotkeyLabel(QuickSwitch.Options["MenuHotkey"]) ")" : ""
        DllCall("SendMessage", "Ptr", search.Hwnd, "UInt", 0x1501, "Ptr", 1, "WStr", " " I18n.T(QuickSwitch._SearchesFiles() ? "QuickSwitch.SearchCue" : "QuickSwitch.SearchCueFolders") hint)   ; EM_SETCUEBANNER: gray cue text
        search.OnEvent("Change", (*) => SetTimer(QuickSwitchPanel._searchTimer, -150))
        QuickSwitchPanel._barHeight := searchY + searchH + searchY               ; the panel shown as only its search box (see PanelLayout), client height
        line := panel.AddText("x0 y+" Win.Scale(6) " w" width " h1 Background" colors.Separator)   ; separator between the search box and the list
        panel.SetFont("s9 c" colors.Text)
        list := panel.AddListView("x0 y+" Win.Scale(2) " w" width " h100 -Hdr -Multi -E0x200 +LV0x10400 Background" colors.Background, ["Name", "Folder", "From"])   ; 0x400 = the full path on hover when it is too long
        icons := IL_Create(3)
        IL_Add(icons, "shell32.dll", 4)                                     ; 1 = folder
        IL_Add(icons, "shell32.dll", 1)                                     ; 2 = file
        IL_Add(icons, "shell32.dll", 44)                                    ; 3 = favorite folder (the star)
        list.SetImageList(icons)
        list.OnEvent("Click", (ctrl, row) => QuickSwitchPanel._PanelJump(row))
        list.OnEvent("ContextMenu", (ctrl, row, *) => QuickSwitchFavorites._ShowRowMenu(row))
        list.OnNotify(-12, (ctrl, lParam) => QuickSwitchPanel._OnCustomDraw(lParam))   ; NM_CUSTOMDRAW: the selected row and gray text in the theme's colors
        list.Add("Icon1", "X")                                              ; the row height can only be measured with one row in (LVM_APPROXIMATEVIEWRECT)
        QuickSwitchPanel._rowTop := SendMessage(0x1040, 1, 0, list) >> 16
        QuickSwitchPanel._rowHeight := (SendMessage(0x1040, 2, 0, list) >> 16) - QuickSwitchPanel._rowTop
        list.Delete()
        list.GetPos(, &listY)
        panel.Show("NA Hide w" width " h100")                               ; the border around the client area: the panel's height for a number of rows
        panel.GetPos(, , , &outerH), panel.GetClientPos(, , , &clientH)
        QuickSwitchPanel._listY := listY, QuickSwitchPanel._frameH := outerH - clientH, QuickSwitchPanel._panelRows := -1
        QuickSwitchPanel._panel := panel, QuickSwitchPanel._panelList := list, QuickSwitchPanel._panelSearch := search, QuickSwitchPanel._panelWidth := width
        QuickSwitchPanel._panelLine := line, QuickSwitchPanel._panelDialogW := IsObject(rect) ? rect.W : 0, QuickSwitchPanel._panelColors := colors, QuickSwitchPanel._panelStar := star
        QuickSwitchPanel._searchHwnd := search.Hwnd, QuickSwitchPanel._listHwnd := list.Hwnd
        if colors.Dark
            try DllCall("uxtheme\SetWindowTheme", "Ptr", list.Hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)   ; dark theme: dark scroll bars and selected row too
        try DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", panel.Hwnd, "UInt", 33, "Int*", 3, "UInt", 4)   ; Win11: small rounded corners
        QuickSwitchPanel._ShowFolders(QuickSwitchPanel._panelBase)
        if (QuickSwitch._placeTimer != "")
            SetTimer(QuickSwitch._placeTimer, -500)                         ; once more: some dialogs move or resize themselves right after they are shown
        if WinExist("ahk_class dopus.lister")                               ; reading Opus takes a moment: show the panel first
            SetTimer(QuickSwitchPanel._AddDOpusFolders.Bind(QuickSwitch, dialog), -1)
    }

    static _AddDOpusFolders(dialog) {
        if (QuickSwitchPanel._panelFor != dialog || !IsObject(QuickSwitchPanel._panelList))
            return
        QuickSwitchPanel._panelBase := QuickSwitchFavorites.MenuFolders()
        if (QuickSwitchPanel._panelFor = dialog && IsObject(QuickSwitchPanel._panelSearch) && QuickSwitchPanel._panelSearch.Value = "")   ; not while searching
            QuickSwitchPanel._ShowFolders(QuickSwitchPanel._panelBase)
    }

    static _searchHwnd := 0, _listHwnd := 0, _panelColors := "", _panelBase := [], _panelList := "", _panelSearch := "", _panelLine := "", _panelStar := "", _panelWidth := 600, _panelDialogW := 0, _rowTop := 20, _rowHeight := 18
    static _listY := 40, _frameH := 2, _barHeight := 40, _panelRows := -1

    ; The search box and the ☆ button right of it (real pixels): the box ends a gap before the button, the button keeps the box's right margin
    static StarSize() => Win.Scale(20)
    static StarX(width) => width - Win.Scale(10) - QuickSwitchPanel.StarSize()
    static SearchWidth(width) => QuickSwitchPanel.StarX(width) - Win.Scale(8) - Win.Scale(10)

    ; Panel width = dialog width, at least 420 (converted by DPI) and at most maxWidth (the work area of the dialog's monitor); all real pixels
    static PanelWidthFor(dialogWidth, maxWidth := 0) => Max(Win.Scale(420), maxWidth > 0 ? Min(dialogWidth, maxWidth) : dialogWidth)

    static _PanelWidthAt(rect) {
        area := Win.WorkAreaAt(rect.X + rect.W // 2, rect.Y + rect.H // 2)
        return QuickSwitchPanel.PanelWidthFor(rect.W, area.Right - area.Left)
    }

    ; The panel's colors: background / text / separator colors of ALTRun's current theme; light while no theme is loaded yet
    static PanelColors() {
        get(key, fallback) => ((value := ThemeManager.Get(key)) != "") ? value : fallback
        background := get("Background", "FFFFFF")
        return {Background: background, Text: get("Title", "1F1F1F"), Subtitle: get("Subtitle", "808080"), Separator: get("Separator", "E4E4E4")
              , Selected: get("SelectedBackground", "DDE7F6"), SelectedText: get("SelectedTitle", "000000"), SelectedSubtitle: get("SelectedSubtitle", "4A5568")
              , Dark: QuickSwitchPanel.IsDarkColor(background)}
    }

    ; "RRGGBB" -> a Windows COLORREF (0x00BBGGRR)
    static ColorRef(hex) {
        value := Integer("0x" hex)
        return ((value & 0xFF) << 16) | (value & 0xFF00) | ((value >> 16) & 0xFF)
    }

    ; Each cell sets its own colors: the selected row in the theme's selection color, paths and sources in gray (NMLVCUSTOMDRAW)
    static _OnCustomDraw(lParam) {
        static CDDS_PREPAINT := 0x1, CDDS_ITEMPREPAINT := 0x10001, CDDS_SUBITEMPREPAINT := 0x30001
        static CDRF_NEWFONT := 0x2, CDRF_NOTIFYITEMDRAW := 0x20, CDRF_NOTIFYSUBITEMDRAW := 0x20
        x64 := (A_PtrSize = 8)
        stage := NumGet(lParam, A_PtrSize * 3, "UInt")
        if (stage = CDDS_PREPAINT)
            return CDRF_NOTIFYITEMDRAW
        if (stage = CDDS_ITEMPREPAINT)
            return CDRF_NOTIFYSUBITEMDRAW
        if (stage != CDDS_SUBITEMPREPAINT)
            return 0
        colors := QuickSwitchPanel._panelColors
        row := NumGet(lParam, x64 ? 56 : 36, "UPtr")
        column := NumGet(lParam, x64 ? 88 : 56, "Int")
        selected := SendMessage(0x102C, row, 0x2, QuickSwitchPanel._panelList) & 0x2   ; LVM_GETITEMSTATE / LVIS_SELECTED
        stateOffset := x64 ? 64 : 40
        NumPut("UInt", NumGet(lParam, stateOffset, "UInt") & ~0x11, lParam, stateOffset)   ; remove CDIS_SELECTED / CDIS_FOCUS so the system's blue highlight isn't drawn
        text := selected ? (column ? colors.SelectedSubtitle : colors.SelectedText) : (column ? colors.Subtitle : colors.Text)
        NumPut("UInt", QuickSwitchPanel.ColorRef(text), lParam, x64 ? 80 : 48)
        NumPut("UInt", QuickSwitchPanel.ColorRef(selected ? colors.Selected : colors.Background), lParam, x64 ? 84 : 52)
        return CDRF_NEWFONT
    }

    static IsDarkColor(hex) => ThemeManager.IsDarkColor(hex)

    static _FontName() {
        try return ThemeManager.FontName()
        return "Segoe UI"
    }

    ; the dialog's width changed: the panel follows
    static _ResizePanel(rect) {
        width := QuickSwitchPanel._PanelWidthAt(rect)
        QuickSwitchPanel._panelDialogW := rect.W
        if (width = QuickSwitchPanel._panelWidth)
            return
        QuickSwitchPanel._panelWidth := width
        QuickSwitchPanel._panelSearch.Move(, , QuickSwitchPanel.SearchWidth(width))
        QuickSwitchPanel._panelStar.Move(QuickSwitchPanel.StarX(width))
        QuickSwitchPanel._panelLine.Move(, , width)
        QuickSwitchPanel._panelList.Move(, , width)
        QuickSwitchPanel._ShowFolders(QuickSwitchPanel._panelFolders)                ; Recomputes the column widths and adjusts the panel's size and position
    }
    static _searchTimer := () => QuickSwitchPanel._SearchPanel()

    static ColumnsWidth(list) {
        total := 0
        loop 3
            total += SendMessage(0x101D, A_Index - 1, 0, list)            ; LVM_GETCOLUMNWIDTH
        return total
    }

    static ClientWidth(list) {
        rect := Buffer(16)
        DllCall("GetClientRect", "Ptr", list.Hwnd, "Ptr", rect)
        return NumGet(rect, 8, "Int")
    }

    ; List: folder name | full path | source (TC / Explorer / recent / search)
    static _ShowFolders(folders) {
        list := QuickSwitchPanel._panelList
        if !IsObject(list)
            return
        QuickSwitchPanel._panelFolders := folders
        list.Opt("-Redraw")
        list.Delete()
        for entry in folders
            if entry.HasOwnProp("Hint")
                list.Add("Icon1", "", entry.Text, "")                      ; the hint row: its text in the wide column
            else
                list.Add("Icon" QuickSwitchPanel.RowIcon(entry), QuickSwitchPanel.FolderName(entry.Path), entry.Path, entry.Tag)
        if folders.Length
            list.Modify(1, "Select Focus")
        list.Opt("+Redraw")
        layout := QuickSwitchPanel._Layout(QuickSwitchPanel._FrameRect(QuickSwitchPanel._panelFor))
        QuickSwitchPanel._SizeList(layout ? layout.Rows : QuickSwitchPanel.WantedRows(folders.Length))
        if DllCall("IsWindowVisible", "Ptr", QuickSwitchPanel._panel.Hwnd) {     ; The search results changed: adjust the size and position again
            QuickSwitchPanel._panelPos := ""
            QuickSwitchPanel._PlacePanel(QuickSwitchPanel._panelFor)
        }
    }

    ; The list's height for this many rows (0 = the panel shows only its search box), and the column widths that go with it
    static _SizeList(rows) {
        list := QuickSwitchPanel._panelList, folders := QuickSwitchPanel._panelFolders
        QuickSwitchPanel._panelRows := rows, QuickSwitchPanel._panelPos := ""
        list.Visible := QuickSwitchPanel._panelLine.Visible := rows > 0
        if !rows
            return
        list.ModifyCol(3, "Auto"), list.ModifyCol(1, "Auto")
        nameW := Max(Min(SendMessage(0x101D, 0, 0, list), QuickSwitchPanel._panelWidth // 3), Win.Scale(120))   ; LVM_GETCOLUMNWIDTH
        list.ModifyCol(1, nameW)
        list.ModifyCol(2, QuickSwitchPanel._panelWidth - nameW - SendMessage(0x101D, 2, 0, list) - Win.Scale(folders.Length > rows ? 22 : 4))
        list.Move(, , , QuickSwitchPanel.ListHeight(rows))
        list.ModifyCol(3, "AutoHdr")                                        ; the last column fills to the right so the selected row's color spans the row; computed after the height is set (with a scroll bar where needed), so no horizontal scroll bar appears
        ; The last column is never narrower than its (hidden) header text: with no rows, or only short tags, it can stick out
        ; past the right edge and bring up a horizontal scrollbar. Take the difference off the path column
        if (over := QuickSwitchPanel.ColumnsWidth(list) - QuickSwitchPanel.ClientWidth(list)) > 0
            list.ModifyCol(2, Max(0, SendMessage(0x101D, 1, 0, list) - over))
        if (row := list.GetNext(0, "F"))
            list.Modify(row, "Vis")                                         ; fewer rows: the selected one stays in view
    }

    static ListHeight(rows) => QuickSwitchPanel._rowTop + (rows - 1) * QuickSwitchPanel._rowHeight + Win.Scale(4)

    ; The height follows the content: 3 ~ PanelRows rows
    static WantedRows(count) => Max(QuickSwitchPanel.MinRows, Min(count, QuickSwitchPanel.PanelRows))
    static MinRows := 3

    ; The row's icon in the panel's image list: 1 folder, 2 file, 3 favorite folder (a star, so favorites stand out)
    static RowIcon(entry) => (entry.HasOwnProp("IsFile") && entry.IsFile) ? 2 : (entry.HasOwnProp("Group") && entry.Group = "favorite") ? 3 : 1

    ; "D:\Projects\Tower" -> "Tower"; drive roots show as "D:\"
    static FolderName(folder) {
        trimmed := RTrim(folder, "\")
        if (trimmed ~= "^[A-Za-z]:$")
            return trimmed "\"
        SplitPath(trimmed, &name)
        return (name != "") ? name : folder
    }

    ; The text in the search box: filter the list and the favorite folders first (paths containing every word), then add the folders and files
    ; found (Everything or the built-in index, at most 30). The favorites aren't in the list itself (they are in the ☆ menu), only in searches
    static _SearchPanel() {
        if !IsObject(QuickSwitchPanel._panelSearch)
            return
        term := Trim(QuickSwitchPanel._panelSearch.Value)
        if (term = "")
            return QuickSwitchPanel._ShowFolders(QuickSwitchPanel._panelBase)
        found := QuickSwitchPanel.FilterFolders(QuickSwitchPanel.SearchBase(QuickSwitchPanel._panelBase), term, QuickSwitchPanel._FindItems(term))
        if (!found.Length && !QuickSwitchPanel._EverythingInUse())
            found.Push(QuickSwitchPanel.IndexHint(AppSettings.Feature("FileSearch")["ScopeFolders"]))
        QuickSwitchPanel._ShowFolders(found)
    }

    static _EverythingInUse() => AppSettings.Feature("FileSearch")["UseEverything"] && Everything.IsRunning()

    ; The row shown when a search finds nothing and Everything isn't used: the built-in index covers only the ScopeFolders, so the
    ; folder looked for may simply be outside it. Choosing the row (click or Enter) opens Preferences -> File Search to add folders
    static IndexHint(scopeFolders) {
        names := "", count := 0
        for folder in scopeFolders {
            if ((resolved := RTrim(Path.Resolve(folder), "\")) = "")
                continue
            if (++count <= 4)
                names .= (names = "" ? "" : ", ") QuickSwitchPanel.FolderName(resolved)
        }
        if (count > 4)
            names .= ", ..."
        return {Path: "", Hint: true, Group: "hint", Tag: "", Text: I18n.T("QuickSwitch.NoMatchIndex", names = "" ? "-" : names)}
    }

    ; What a search filters: the favorite folders first, then the panel's folders that aren't favorites
    static SearchBase(base) {
        result := QuickSwitchFavorites.FavoriteEntries(), seen := Map()
        for entry in result
            seen[StrLower(RTrim(entry.Path, "\"))] := true
        for entry in base
            if !seen.Has(StrLower(RTrim(entry.Path, "\")))
                result.Push(entry)
        return result
    }

    ; [{Path, IsFolder}], folders first. Folders are searched separately (searched together with files, a folder with exactly the name might not make it in for not being recently modified);
    ; PanelSearch = "folders" searches folders only
    static _FindItems(term) {
        found := [], seen := Map()
        try {
            for item in FileSearchProvider.Query(term, 15, false, true)
                found.Push({Path: item.Path, IsFolder: true}), seen[StrLower(item.Path)] := true
            if !QuickSwitch._SearchesFiles()
                return found
            for item in FileSearchProvider.Query(term, 30, false)
                if (!item.IsFolder && !seen.Has(StrLower(item.Path)))
                    found.Push({Path: item.Path, IsFolder: false})
        }
        return found
    }

    ; Filters base (case-insensitive, every word must appear in the path), then appends the items of extra that aren't duplicates:
    ; each item of extra is a path (folder) or {Path, IsFolder}; folders get the source "search", files "file"
    static FilterFolders(base, term, extra) {
        result := [], seen := Map(), words := StrSplit(Trim(term), " ")
        for entry in base {
            matched := true
            for word in words
                if (word != "" && !InStr(entry.Path, word))
                    matched := false
            if matched {
                result.Push(entry)
                seen[StrLower(RTrim(entry.Path, "\"))] := true
            }
        }
        for item in extra {
            itemPath := IsObject(item) ? item.Path : item
            isFile := IsObject(item) && !item.IsFolder
            if seen.Has(StrLower(RTrim(itemPath, "\")))
                continue
            seen[StrLower(RTrim(itemPath, "\"))] := true
            result.Push({Path: itemPath, Group: "search", IsFile: isFile, Tag: I18n.T(isFile ? "QuickSwitch.TagFile" : "QuickSwitch.TagSearch")})
        }
        return result
    }

    static _PanelJump(row) {
        if (row < 1 || row > QuickSwitchPanel._panelFolders.Length || !QuickSwitchPanel._panelFor)
            return
        entry := QuickSwitchPanel._panelFolders[row]
        if entry.HasOwnProp("Hint")
            return App.OpenPreferences("Prefs.Page.FileSearch")
        if (entry.HasOwnProp("IsFile") && entry.IsFile) {
            dialog := QuickSwitchPanel._panelFor
            QuickSwitchPanel._BackToDialog()
            FileDialogs.PickFile(entry.Path, dialog)
            QuickSwitchPanel._ClearSearch()
        } else
            QuickSwitchPanel.JumpTo(entry.Path)
    }

    ; The dialog jumps to folder (a row of the panel, an item of the ☆ menu: * takes the menu's item name, position and menu)
    static JumpTo(folder, *) {
        if !(dialog := QuickSwitchPanel._panelFor)
            return
        QuickSwitchPanel._BackToDialog()
        FileDialogs.SetDialogPath(RTrim(folder, "\") "\", dialog)
        QuickSwitchPanel._ClearSearch()
    }

    static _ClearSearch() {
        if (IsObject(QuickSwitchPanel._panelSearch) && QuickSwitchPanel._panelSearch.Value != "") {   ; after jumping, clear the search so the list is back as it was
            QuickSwitchPanel._panelSearch.Value := ""
            QuickSwitchPanel._ShowFolders(QuickSwitchPanel._panelBase)
        }
    }

    static _BackToDialog() {
        dialog := QuickSwitchPanel._panelFor
        if (!dialog || WinActive("ahk_id " dialog))
            return
        try {
            WinActivate("ahk_id " dialog)
            WinWaitActive("ahk_id " dialog, , 1)
        }
    }

    ; Keys in the search box: ↑ ↓ choose, Enter jumps, Esc goes back to the dialog
    static PanelKey(key) {
        list := QuickSwitchPanel._panelList
        if !IsObject(list)
            return
        count := list.GetCount(), row := list.GetNext()
        switch key {
            case "Up", "Down":
                if !count
                    return
                row := (key = "Down") ? Min(row + 1, count) : Max(row - 1, 1)
                list.Modify(0, "-Select"), list.Modify(row, "Select Focus Vis")
            case "Enter":
                QuickSwitchPanel._PanelJump(row ? row : 1)
            case "Escape":
                QuickSwitchPanel._BackToDialog()
        }
    }

    static _PlacePanel(dialog) {
        if !IsObject(QuickSwitchPanel._panel)
            return
        rect := QuickSwitchPanel._FrameRect(dialog)
        if !IsObject(rect)
            return QuickSwitchPanel.HidePanel()
        if (QuickSwitchPanel._panelDialogW && rect.W != QuickSwitchPanel._panelDialogW)
            return QuickSwitchPanel._ResizePanel(rect)
        layout := QuickSwitchPanel._Layout(rect)
        if (layout.Rows != QuickSwitchPanel._panelRows)                          ; more or less room (the dialog moved), or the caret came into / left the panel
            QuickSwitchPanel._SizeList(layout.Rows)
        key := layout.X "," layout.Y "," layout.Rows
        if (key != QuickSwitchPanel._panelPos) {                                 ; reposition only when the dialog moved
            height := layout.Rows ? QuickSwitchPanel._listY + QuickSwitchPanel.ListHeight(layout.Rows) : QuickSwitchPanel._barHeight
            QuickSwitchPanel._panel.Show("NA x" layout.X " y" layout.Y " w" QuickSwitchPanel._panelWidth " h" height)   ; client size
            QuickSwitchPanel._panelPos := key
        }
    }

    ; Where the panel goes for the dialog at rect, with how many rows ("" without a dialog); the caret in the panel opens its list
    static _Layout(rect) {
        if !(IsObject(rect) && IsObject(QuickSwitchPanel._panel))
            return ""
        area := Win.WorkAreaAt(rect.X + rect.W // 2, rect.Y + rect.H // 2)
        rowsH := QuickSwitchPanel._frameH + QuickSwitchPanel._listY + QuickSwitchPanel.ListHeight(0)   ; the panel's height with 0 rows; each row adds _rowHeight
        open := DllCall("GetForegroundWindow", "Ptr") = QuickSwitchPanel._panel.Hwnd
        return QuickSwitchPanel.PanelLayout(rect.X, rect.Y, rect.W, rect.H, QuickSwitchPanel._panelWidth, QuickSwitchPanel.WantedRows(QuickSwitchPanel._panelFolders.Length)
            , rowsH, QuickSwitchPanel._rowHeight, QuickSwitchPanel._frameH + QuickSwitchPanel._barHeight, area, open)
    }

    ; the window's visible frame (Win10/11 windows have an invisible resize border around them, which WinGetPos includes)
    static _FrameRect(hwnd) {
        rect := Buffer(16, 0)
        if !DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", hwnd, "UInt", 9, "Ptr", rect, "UInt", 16) {   ; DWMWA_EXTENDED_FRAME_BOUNDS
            left := NumGet(rect, 0, "Int"), top := NumGet(rect, 4, "Int")
            if (NumGet(rect, 8, "Int") > left)
                return {X: left, Y: top, W: NumGet(rect, 8, "Int") - left, H: NumGet(rect, 12, "Int") - top}
        }
        try {
            WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
            return {X: x, Y: y, W: w, H: h}
        }
        return ""
    }

    ; Panel position and rows {X, Y, Rows} (as in Listary), never over the dialog's buttons and file list (#440): right below the dialog,
    ; left-aligned, else above it, with all wanted rows; else on the side with more room, with the rows that fit there (at least MinRows,
    ; the list scrolls); else (a tall or maximized dialog) only the search box (Rows 0): below or above the dialog, or inside it at its top
    ; (over the title bar). open = the caret is in the panel: then it opens its list with all rows from there, over the dialog where needed.
    ; The panel's height with n rows is rowsH + n * rowH, with none barH; all in real pixels
    static PanelLayout(x, y, w, h, panelW, wanted, rowsH, rowH, barH, area, open := false) {
        left := Max(area.Left, Min(x, area.Right - panelW))
        below := area.Bottom - (y + h), above := y - area.Top
        fullH := rowsH + wanted * rowH
        if (fullH <= below)
            return {X: left, Y: y + h, Rows: wanted}
        if (fullH <= above)
            return {X: left, Y: y - fullH, Rows: wanted}
        fitBelow := (below - rowsH) // rowH, fitAbove := (above - rowsH) // rowH
        if (fitBelow >= QuickSwitchPanel.MinRows && fitBelow >= fitAbove)
            return {X: left, Y: y + h, Rows: fitBelow}
        if (fitAbove >= QuickSwitchPanel.MinRows)
            return {X: left, Y: y - (rowsH + fitAbove * rowH), Rows: fitAbove}
        barY := (barH <= below) ? y + h : (barH <= above) ? y - barH : Max(area.Top, y)
        if !open
            return {X: left, Y: barY, Rows: 0}
        return {X: left, Y: Max(area.Top, Min(barY, area.Bottom - fullH)), Rows: wanted}
    }
}
