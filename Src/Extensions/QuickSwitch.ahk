;===============================================================================
; QuickSwitch.ahk - quick folder jumps in open / save dialogs (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like Listary's "Quick Switch": in an Open / Save dialog, the hotkey (Ctrl+G by default) jumps the dialog to the
; folder open in the file manager used last: Total Commander, Directory Opus or File Explorer (see FileManagerFolder).
; AutoSwitch = 1: switching from one of them to a dialog jumps automatically.
; The folder panel (ShowPanel, on by default, like Listary's Quick Switch window): as soon as a dialog appears, a search box +
; folder list (both panels of every TC window, Explorer windows, recently used folders) sits below it; one click jumps there, no hotkey to remember;
; the search box also searches folders (with PanelSearch = all, files too: choosing a file jumps to its folder). Hidden when the dialog isn't in front,
; refreshed when back in the dialog (a folder change in TC shows right away). The colors follow ALTRun's theme.
; Hotkey (MenuHotkey, default Ctrl+Shift+G): the caret jumps to the panel's search box to choose with the keyboard; shows the panel for the moment when automatic showing is off.
;
; Settings (ALTRun.json -> Extensions.QuickSwitch):
;   Enabled / Hotkey (the jump hotkey, for all three) / MenuHotkey / RecentFolders / ShowPanel / PanelSearch / AutoSwitch /
;   StandardDialogs (ahk_class #32770) / BrowseDialogs / DialogWindows / ExcludeWindows / AutoSwitchExclude
;   The window lists are arrays of window conditions (ahk_class / ahk_exe / part of the title)
;
; Usage:
;   QuickSwitch.Init(AppSettings.Extension("QuickSwitch"))    at start-up
;   QuickSwitch.FolderOfWindow(hwnd)     current folder of a TC / Opus / Explorer window (read only, for "terminal here")
;   QuickSwitch.MenuFolders()            the folders in the panel [{Path, Group, Tag}]
;   QuickSwitch.RecentFolders(n)         the n most recently used folders (opened files count by their folder)
;===============================================================================

class QuickSwitch {
    static Options := Map(), _hasExclude := false
    static _lastWasManager := false
    static _panel := "", _panelFor := 0, _panelFolders := [], _panelPos := ""

    static Init(options) {
        QuickSwitch.Options := options
        if !options["Enabled"]
            return
        if options["StandardDialogs"]
            GroupAdd("ALTRunDialogs", "ahk_class #32770")
        for window in options["DialogWindows"]
            GroupAdd("ALTRunDialogs", window)
        QuickSwitch._hasExclude := false
        for window in options["ExcludeWindows"]
            GroupAdd("ALTRunDialogExclude", window), QuickSwitch._hasExclude := true
        for window in options["AutoSwitchExclude"]
            GroupAdd("ALTRunAutoSwitchExclude", window)

        HotIf((*) => QuickSwitch.IsFileDialog())
        try {
            if (options["Hotkey"] != "")
                Hotkey(options["Hotkey"], (*) => QuickSwitch.SyncFileManagerPath())
            if (options["MenuHotkey"] != "")
                Hotkey(options["MenuHotkey"], (*) => QuickSwitch.FocusPanel())
        } catch as e {
            Logger.Error("QuickSwitch: cannot register hotkeys - " e.Message)
        }
        OnMessage(0x100, (p*) => QuickSwitch._OnKeyDown(p*))                ; WM_KEYDOWN: ↑ ↓ Enter Esc in the panel's search box

        SetTimer(Logger.Watch("quick switch watch", () => QuickSwitch._Watch()), 250)
    }

    ; WM_KEYDOWN rather than a hotkey: while a Chinese IME is composing (not yet committed) keys are VK_PROCESSKEY and aren't taken for Enter / Esc
    ; compare only the window handles noted: after the panel closes its control objects are destroyed and their Hwnd can't be read
    static _OnKeyDown(wParam, lParam, msg, hwnd) {
        if !(hwnd && (hwnd = QuickSwitch._searchHwnd || hwnd = QuickSwitch._listHwnd))
            return
        switch wParam {
            case 0x26: QuickSwitch.PanelKey("Up")
            case 0x28: QuickSwitch.PanelKey("Down")
            case 0x0D: QuickSwitch.PanelKey("Enter")
            case 0x1B: QuickSwitch.PanelKey("Escape")
            default:   return
        }
        return 0
    }

    ; Every 250 ms: with AutoSwitch, jump when switching from TC to the dialog; show / hide / position the folder panel
    static _Watch() {
        if QuickSwitch.PanelActive() {                                      ; typing in the panel's search box: the dialog and panel stay as they are
            if !WinExist("ahk_id " QuickSwitch._panelFor)                   ; the dialog was closed (by its program): close the panel too
                QuickSwitch.HidePanel()
            return
        }
        isDialog := QuickSwitch.IsFileDialog()
        if (isDialog && QuickSwitch.Options["AutoSwitch"] && QuickSwitch._lastWasManager && !WinActive("ahk_group ALTRunAutoSwitchExclude"))
            QuickSwitch.SyncFileManagerPath(true)
        QuickSwitch._lastWasManager := QuickSwitch.IsFileManager(WinExist("A"))
        if (QuickSwitch._PanelEnabled() || IsObject(QuickSwitch._panel))    ; with automatic showing off, a panel opened by the hotkey still follows the dialog
            QuickSwitch._UpdatePanel(isDialog)
    }

    static _SearchesFiles() => QuickSwitch.Options.Has("PanelSearch") && QuickSwitch.Options["PanelSearch"] = "all"

    static _PanelEnabled() => QuickSwitch.Options.Has("ShowPanel") && QuickSwitch.Options["ShowPanel"]

    ; Hotkey (MenuHotkey, default Ctrl+Shift+G): the caret jumps to the panel's search box to choose with the keyboard; shows the panel for the moment when automatic showing is off
    static FocusPanel() {
        dialog := WinExist("A")
        if (!IsObject(QuickSwitch._panel) || QuickSwitch._panelFor != dialog) {
            QuickSwitch._BuildPanel(dialog)
            QuickSwitch._PlacePanel(dialog)
        }
        if !IsObject(QuickSwitch._panel)
            return
        try {
            WinActivate("ahk_id " QuickSwitch._panel.Hwnd)
            ControlFocus(QuickSwitch._panelSearch)
        }
    }

    ;---------------------------------------------------------------------------
    ; The folder panel (like Listary's Quick Switch window): attached below the dialog, as wide as it;
    ; the search box on top (filters the list first, then finds folders and files with Everything / the built-in index), the list below.
    ; Clicking an item or Enter in the search box jumps there (files: see PickFile); Esc goes back to the dialog
    ;---------------------------------------------------------------------------
    static _UpdatePanel(isDialog) {
        dialog := 0
        try dialog := isDialog ? WinGetID("A") : 0
        if !dialog
            return QuickSwitch.HidePanel()
        if (dialog != QuickSwitch._panelFor) {                              ; a new dialog, or back from another window: list the folders again
            if !QuickSwitch._PanelEnabled()
                return QuickSwitch.HidePanel()
            QuickSwitch._BuildPanel(dialog)
        }
        QuickSwitch._PlacePanel(dialog)
    }

    ; typing in the panel's search box (the panel is the foreground window)
    static PanelActive(*) => IsObject(QuickSwitch._panel) && WinActive("ahk_id " QuickSwitch._panel.Hwnd)

    static HidePanel() {
        if IsObject(QuickSwitch._panel)
            try QuickSwitch._panel.Destroy()
        QuickSwitch._panel := "", QuickSwitch._panelFor := 0, QuickSwitch._panelPos := ""
        QuickSwitch._panelSearch := "", QuickSwitch._panelList := "", QuickSwitch._panelLine := ""   ; the controls are destroyed: no more references
        QuickSwitch._searchHwnd := 0, QuickSwitch._listHwnd := 0
    }

    static PanelRows := 8

    static _BuildPanel(dialog) {
        QuickSwitch.HidePanel()
        QuickSwitch._panelFor := dialog
        QuickSwitch._panelBase := QuickSwitch.MenuFolders(false)
        rect := QuickSwitch._FrameRect(dialog)
        width := QuickSwitch.PanelWidthFor(IsObject(rect) ? rect.W : 600)
        colors := QuickSwitch.PanelColors()                                 ; follows ALTRun's theme (light / dark...)
        ; -DPIScale: the panel is as wide as the dialog, in the dialog's real pixels; the row height is measured in pixels too. Letting AHK scale by DPI again
        ; would make the panel half again as wide as the dialog at 150%, with an empty area below. Fixed margins are converted with Win.Scale
        panel := Gui("-Caption +ToolWindow +AlwaysOnTop +Border -DPIScale", "ALTRun Quick Switch")
        panel.MarginX := 0, panel.MarginY := 0, panel.BackColor := colors.Background
        panel.SetFont("s10 c" colors.Text, QuickSwitch._FontName())
        search := panel.AddEdit("x" Win.Scale(10) " y" Win.Scale(8) " w" (width - Win.Scale(20)) " r1 -Multi -E0x200 Background" colors.Background)
        hint := (QuickSwitch.Options.Has("MenuHotkey") && QuickSwitch.Options["MenuHotkey"] != "") ? "  (" Win.HotkeyLabel(QuickSwitch.Options["MenuHotkey"]) ")" : ""
        DllCall("SendMessage", "Ptr", search.Hwnd, "UInt", 0x1501, "Ptr", 1, "WStr", " " I18n.T(QuickSwitch._SearchesFiles() ? "QuickSwitch.SearchCue" : "QuickSwitch.SearchCueFolders") hint)   ; EM_SETCUEBANNER: gray cue text
        search.OnEvent("Change", (*) => SetTimer(QuickSwitch._searchTimer, -150))
        line := panel.AddText("x0 y+" Win.Scale(6) " w" width " h1 Background" colors.Separator)   ; separator between the search box and the list
        panel.SetFont("s9 c" colors.Text)
        list := panel.AddListView("x0 y+" Win.Scale(2) " w" width " h100 -Hdr -Multi -E0x200 +LV0x10400 Background" colors.Background, ["Name", "Folder", "From"])   ; 0x400 = the full path on hover when it is too long
        icons := IL_Create(2)
        IL_Add(icons, "shell32.dll", 4)                                     ; 1 = folder
        IL_Add(icons, "shell32.dll", 1)                                     ; 2 = file
        list.SetImageList(icons)
        list.OnEvent("Click", (ctrl, row) => QuickSwitch._PanelJump(row))
        list.OnNotify(-12, (ctrl, lParam) => QuickSwitch._OnCustomDraw(lParam))   ; NM_CUSTOMDRAW: the selected row and gray text in the theme's colors
        list.Add("Icon1", "X")                                              ; the row height can only be measured with one row in (LVM_APPROXIMATEVIEWRECT)
        QuickSwitch._rowTop := SendMessage(0x1040, 1, 0, list) >> 16
        QuickSwitch._rowHeight := (SendMessage(0x1040, 2, 0, list) >> 16) - QuickSwitch._rowTop
        list.Delete()
        QuickSwitch._panel := panel, QuickSwitch._panelList := list, QuickSwitch._panelSearch := search, QuickSwitch._panelWidth := width
        QuickSwitch._panelLine := line, QuickSwitch._panelDialogW := IsObject(rect) ? rect.W : 0, QuickSwitch._panelColors := colors
        QuickSwitch._searchHwnd := search.Hwnd, QuickSwitch._listHwnd := list.Hwnd
        if colors.Dark
            try DllCall("uxtheme\SetWindowTheme", "Ptr", list.Hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)   ; dark theme: dark scroll bars and selected row too
        try DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", panel.Hwnd, "UInt", 33, "Int*", 3, "UInt", 4)   ; Win11: small rounded corners
        QuickSwitch._ShowFolders(QuickSwitch._panelBase)
        if WinExist("ahk_class dopus.lister")                               ; reading Opus takes a moment: show the panel first
            SetTimer(QuickSwitch._AddDOpusFolders.Bind(QuickSwitch, dialog), -1)
    }

    static _AddDOpusFolders(dialog) {
        if (QuickSwitch._panelFor != dialog || !IsObject(QuickSwitch._panelList))
            return
        QuickSwitch._panelBase := QuickSwitch.MenuFolders()
        if (QuickSwitch._panelFor = dialog && IsObject(QuickSwitch._panelSearch) && QuickSwitch._panelSearch.Value = "")   ; not while searching
            QuickSwitch._ShowFolders(QuickSwitch._panelBase)
    }

    static _searchHwnd := 0, _listHwnd := 0, _panelColors := "", _panelBase := [], _panelList := "", _panelSearch := "", _panelLine := "", _panelWidth := 600, _panelDialogW := 0, _rowTop := 20, _rowHeight := 18

    ; Panel width = dialog width (420 ~ 1000, converted by DPI; all real pixels)
    static PanelWidthFor(dialogWidth) => Max(Win.Scale(420), Min(dialogWidth, Win.Scale(1000)))

    ; The panel's colors: background / text / separator colors of ALTRun's current theme; light while no theme is loaded yet
    static PanelColors() {
        get(key, fallback) => ((value := ThemeManager.Get(key)) != "") ? value : fallback
        background := get("Background", "FFFFFF")
        return {Background: background, Text: get("Title", "1F1F1F"), Subtitle: get("Subtitle", "808080"), Separator: get("Separator", "E4E4E4")
              , Selected: get("SelectedBackground", "DDE7F6"), SelectedText: get("SelectedTitle", "000000"), SelectedSubtitle: get("SelectedSubtitle", "4A5568")
              , Dark: QuickSwitch.IsDarkColor(background)}
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
        colors := QuickSwitch._panelColors
        row := NumGet(lParam, x64 ? 56 : 36, "UPtr")
        column := NumGet(lParam, x64 ? 88 : 56, "Int")
        selected := SendMessage(0x102C, row, 0x2, QuickSwitch._panelList) & 0x2   ; LVM_GETITEMSTATE / LVIS_SELECTED
        stateOffset := x64 ? 64 : 40
        NumPut("UInt", NumGet(lParam, stateOffset, "UInt") & ~0x11, lParam, stateOffset)   ; remove CDIS_SELECTED / CDIS_FOCUS so the system's blue highlight isn't drawn
        text := selected ? (column ? colors.SelectedSubtitle : colors.SelectedText) : (column ? colors.Subtitle : colors.Text)
        NumPut("UInt", QuickSwitch.ColorRef(text), lParam, x64 ? 80 : 48)
        NumPut("UInt", QuickSwitch.ColorRef(selected ? colors.Selected : colors.Background), lParam, x64 ? 84 : 52)
        return CDRF_NEWFONT
    }

    static IsDarkColor(hex) => ThemeManager.IsDarkColor(hex)

    static _FontName() {
        try return ThemeManager.FontName()
        return "Segoe UI"
    }

    ; the dialog's width changed: the panel follows
    static _ResizePanel(dialogWidth) {
        width := QuickSwitch.PanelWidthFor(dialogWidth)
        QuickSwitch._panelDialogW := dialogWidth
        if (width = QuickSwitch._panelWidth)
            return
        QuickSwitch._panelWidth := width
        QuickSwitch._panelSearch.Move(, , width - Win.Scale(20))
        QuickSwitch._panelLine.Move(, , width)
        QuickSwitch._panelList.Move(, , width)
        QuickSwitch._ShowFolders(QuickSwitch._panelFolders)                ; Recomputes the column widths and adjusts the panel's size and position
    }
    static _searchTimer := () => QuickSwitch._SearchPanel()

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
        list := QuickSwitch._panelList
        if !IsObject(list)
            return
        QuickSwitch._panelFolders := folders
        list.Opt("-Redraw")
        list.Delete()
        for entry in folders
            list.Add(entry.HasOwnProp("IsFile") && entry.IsFile ? "Icon2" : "Icon1", QuickSwitch.FolderName(entry.Path), entry.Path, entry.Tag)
        list.ModifyCol(3, "Auto"), list.ModifyCol(1, "Auto")
        nameW := Max(Min(SendMessage(0x101D, 0, 0, list), QuickSwitch._panelWidth // 3), Win.Scale(120))   ; LVM_GETCOLUMNWIDTH
        list.ModifyCol(1, nameW)
        list.ModifyCol(2, QuickSwitch._panelWidth - nameW - SendMessage(0x101D, 2, 0, list) - Win.Scale(folders.Length > QuickSwitch.PanelRows ? 22 : 4))
        if folders.Length
            list.Modify(1, "Select Focus")
        list.Opt("+Redraw")
        rows := Max(3, Min(folders.Length, QuickSwitch.PanelRows))           ; the height follows the content (3 ~ PanelRows rows)
        list.Move(, , , QuickSwitch._rowTop + (rows - 1) * QuickSwitch._rowHeight + Win.Scale(4))
        list.ModifyCol(3, "AutoHdr")                                        ; the last column fills to the right so the selected row's color spans the row; computed after the height is set (with a scroll bar where needed), so no horizontal scroll bar appears
        ; The last column is never narrower than its (hidden) header text: with no rows, or only short tags, it can stick out
        ; past the right edge and bring up a horizontal scrollbar. Take the difference off the path column
        if (over := QuickSwitch.ColumnsWidth(list) - QuickSwitch.ClientWidth(list)) > 0
            list.ModifyCol(2, Max(0, SendMessage(0x101D, 1, 0, list) - over))
        if DllCall("IsWindowVisible", "Ptr", QuickSwitch._panel.Hwnd) {     ; The search results changed: adjust the size and position again
            QuickSwitch._panelPos := ""
            QuickSwitch._PlacePanel(QuickSwitch._panelFor)
        }
    }

    ; "D:\Projects\Tower" -> "Tower"; drive roots show as "D:\"
    static FolderName(folder) {
        trimmed := RTrim(folder, "\")
        if (trimmed ~= "^[A-Za-z]:$")
            return trimmed "\"
        SplitPath(trimmed, &name)
        return (name != "") ? name : folder
    }

    ; The text in the search box: filter the list first (paths containing every word), then add the folders and files found (Everything or the built-in index, at most 30)
    static _SearchPanel() {
        if !IsObject(QuickSwitch._panelSearch)
            return
        term := Trim(QuickSwitch._panelSearch.Value)
        QuickSwitch._ShowFolders(QuickSwitch.FilterFolders(QuickSwitch._panelBase, term, term = "" ? [] : QuickSwitch._FindItems(term)))
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
        if (row < 1 || row > QuickSwitch._panelFolders.Length || !QuickSwitch._panelFor)
            return
        entry := QuickSwitch._panelFolders[row], dialog := QuickSwitch._panelFor
        QuickSwitch._BackToDialog()
        if (entry.HasOwnProp("IsFile") && entry.IsFile)
            QuickSwitch.PickFile(entry.Path, dialog)
        else
            QuickSwitch.SetDialogPath(RTrim(entry.Path, "\") "\", dialog)
        if (IsObject(QuickSwitch._panelSearch) && QuickSwitch._panelSearch.Value != "") {   ; after jumping, clear the search so the list is back as it was
            QuickSwitch._panelSearch.Value := ""
            QuickSwitch._ShowFolders(QuickSwitch._panelBase)
        }
    }

    static _BackToDialog() {
        dialog := QuickSwitch._panelFor
        if (!dialog || WinActive("ahk_id " dialog))
            return
        try {
            WinActivate("ahk_id " dialog)
            WinWaitActive("ahk_id " dialog, , 1)
        }
    }

    ; Keys in the search box: ↑ ↓ choose, Enter jumps, Esc goes back to the dialog
    static PanelKey(key) {
        list := QuickSwitch._panelList
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
                QuickSwitch._PanelJump(row ? row : 1)
            case "Escape":
                QuickSwitch._BackToDialog()
        }
    }

    static _PlacePanel(dialog) {
        if !IsObject(QuickSwitch._panel)
            return
        rect := QuickSwitch._FrameRect(dialog)
        if !IsObject(rect)
            return QuickSwitch.HidePanel()
        if (QuickSwitch._panelDialogW && rect.W != QuickSwitch._panelDialogW)
            return QuickSwitch._ResizePanel(rect.W)
        if !QuickSwitch._panelPos                                           ; first showing, or the size changed: compute the size from the content first
            QuickSwitch._panel.Show(DllCall("IsWindowVisible", "Ptr", QuickSwitch._panel.Hwnd) ? "NA AutoSize" : "NA Hide AutoSize")
        QuickSwitch._panel.GetPos(, , &panelW, &panelH)
        area := Win.WorkAreaAt(rect.X + rect.W // 2, rect.Y + rect.H // 2)
        pos := QuickSwitch.PanelPosition(rect.X, rect.Y, rect.W, rect.H, panelW, panelH, area)
        key := pos.X "," pos.Y
        if (key != QuickSwitch._panelPos) {                                 ; reposition only when the dialog moved
            QuickSwitch._panel.Show("NA x" pos.X " y" pos.Y)
            QuickSwitch._panelPos := key
        }
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

    ; Panel position (as in Listary): right below the dialog, left-aligned; above the dialog when there is no room below;
    ; inside the dialog at its bottom when there is room neither above nor below
    static PanelPosition(x, y, w, h, panelW, panelH, area) {
        left := Max(area.Left, Min(x, area.Right - panelW))
        if (y + h + panelH <= area.Bottom)
            return {X: left, Y: y + h}
        if (y - panelH >= area.Top)
            return {X: left, Y: y - panelH}
        return {X: left, Y: Max(area.Top, area.Bottom - panelH)}
    }

    ; Only real open / save file dialogs count (a file name box + a file list + buttons or an address bar)
    static IsFileDialog(*) {
        browse := QuickSwitch.Options.Has("BrowseDialogs") && QuickSwitch.Options["BrowseDialogs"]
        listed := WinActive("ahk_group ALTRunDialogs")
        if !(listed || browse && WinActive("ahk_class #32770"))
            return false
        if (QuickSwitch._hasExclude && WinActive("ahk_group ALTRunDialogExclude"))
            return false
        controls := [], title := "", winClass := ""                         ; read once for both kinds (this runs every 250 ms)
        try controls := WinGetControls("A")
        try title := WinGetTitle("A")
        try winClass := WinGetClass("A")
        if (browse && QuickSwitch.LooksLikeBrowseDialog(winClass, controls, QuickSwitch._TreeId(controls, "A")))
            return true
        return listed && QuickSwitch.LooksLikeFileDialog(winClass, controls, title)
    }

    ; Decided by the window's controls (independent of the system language). #32770 is a system dialog: it needs Explorer's file view to count;
    ; one with only a plain list (e.g. the main window of "System Information" msinfo32, with a find box, a list and buttons) doesn't.
    ; Other programs ("other dialogs" added in Preferences) count with a plain list too; those whose controls can't be read (Qt etc.) are decided by title
    static LooksLikeFileDialog(winClass, controls, title := "") {
        if !controls.Length
            return RegExMatch(title, "i)(open|save|select|import|export|browse|打开|另存|选择|导入|导出)") > 0
        hasNameEdit := QuickSwitch._HasControl(controls, "^Edit\d+$")
        hasFileView := (winClass = "#32770")
            ? QuickSwitch._HasControl(controls, "^DirectUIHWND\d+$", "^SHELLDLL_DefView\d*$")
            : QuickSwitch._HasControl(controls, "^DirectUIHWND\d+$", "^SHELLDLL_DefView\d*$", "^SysListView32\d*$")
        hasButtons  := QuickSwitch._HasControl(controls, "^Button\d+$", "^ToolbarWindow32\d+$", "^ComboBoxEx32\d+$", "^ComboBox\d+$")
        return hasNameEdit && hasFileView && hasButtons
    }

    ; The old "Browse for Folder" dialog (SHBrowseForFolder): only a folder tree, no file list, so it is not one
    ; of the dialogs above. Told apart by fixed names, whatever the system language: the new style puts the tree
    ; in a "SHBrowseForFolder ShellNameSpace Control", the classic style has the tree itself with control ID 0x3741.
    ; Turned on by Works in → "Browse for Folder dialogs" (BrowseDialogs); ExcludeWindows applies too
    static IsBrowseWindow(hwnd) {
        controls := [], winClass := ""
        try controls := WinGetControls("ahk_id " hwnd), winClass := WinGetClass("ahk_id " hwnd)
        return QuickSwitch.LooksLikeBrowseDialog(winClass, controls, QuickSwitch._TreeId(controls, "ahk_id " hwnd))
    }

    static LooksLikeBrowseDialog(winClass, controls, treeId := 0) {
        return winClass = "#32770" && (treeId = 0x3741 || QuickSwitch._HasControl(controls, "^SHBrowseForFolder ShellNameSpace Control\d+$"))
    }

    ; Control ID of the dialog's first tree view (0 = none)
    static _TreeId(controls, win) {
        for name in controls
            if (name = "SysTreeView321")
                try return DllCall("GetDlgCtrlID", "Ptr", ControlGetHwnd(name, win), "Int")
        return 0
    }

    ; Select a folder in a "Browse for Folder" dialog's tree, like Listary: BFFM_SETSELECTIONW expands the tree to it.
    ; The message carries a pointer, so the path is first written into the dialog's own process. Doesn't press OK
    static SelectInBrowseDialog(folder, dialog) {
        folder := RTrim(folder, "\")
        if RegExMatch(folder, "^[A-Za-z]:$")
            folder .= "\"                                                   ; a drive needs its backslash
        pid := 0
        try pid := WinGetPID("ahk_id " dialog)
        if !pid || !(process := DllCall("OpenProcess", "UInt", 0x38, "Int", false, "UInt", pid, "Ptr"))   ; PROCESS_VM_OPERATION | VM_READ | VM_WRITE
            return false
        size := (StrLen(folder) + 1) * 2, ok := false
        remote := DllCall("VirtualAllocEx", "Ptr", process, "Ptr", 0, "UPtr", size, "UInt", 0x3000, "UInt", 0x04, "Ptr")   ; MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE
        if remote {
            DllCall("IsWow64Process", "Ptr", process, "Int*", &wow64 := 0)
            text := Buffer(size), StrPut(folder, text, "UTF-16")
            if (!(wow64 && remote > 0xFFFFFFFF)                             ; a 32-bit program only sees the low 4 GB
                && DllCall("WriteProcessMemory", "Ptr", process, "Ptr", remote, "Ptr", text, "UPtr", size, "Ptr", 0)) {
                try {
                    SendMessage(0x467, 1, remote, , "ahk_id " dialog, , , , 2000)   ; BFFM_SETSELECTIONW (WM_USER + 103); wParam TRUE = lParam is a path
                    ok := true
                }
            }
            DllCall("VirtualFreeEx", "Ptr", process, "Ptr", remote, "UPtr", 0, "UInt", 0x8000)   ; MEM_RELEASE
        }
        DllCall("CloseHandle", "Ptr", process)
        if !ok
            Logger.Error("QuickSwitch: cannot select " folder " in the Browse for Folder dialog")
        return ok
    }

    static _HasControl(controls, patterns*) {
        for name in controls
            for pattern in patterns
                if RegExMatch(name, "i)" pattern)
                    return true
        return false
    }

    ; The jump hotkey: the folder of the file manager used last (TC, Opus or Explorer, see FileManagerFolder)
    static SyncFileManagerPath(silent := false) {
        manager := QuickSwitch.FileManagerFolder()
        if !(IsObject(manager) && manager.Path != "")
            return silent ? "" : MsgBox(I18n.T("QuickSwitch.NoFileManager"), App.Name, 48)
        QuickSwitch.SetDialogPath(RTrim(manager.Path, "\") "\")             ; with a trailing backslash: AutoCAD and others need it
    }

    ; The folder of a Total Commander panel: asked directly through TC's WM_COPYDATA interface (since TC 8.0), without touching the clipboard.
    ; hwnd = 0 takes the frontmost TC window; otherPanel = true takes the other panel
    static TotalCmdFolder(hwnd := 0, otherPanel := false) {
        if !(hwnd := hwnd ? hwnd : WinExist("ahk_class TTOTAL_CMD"))
            return ""
        folder := RTrim(QuickSwitch.TotalCmdAsk(hwnd, otherPanel ? "TP" : "SP"), "\")
        return (folder ~= "^([A-Za-z]:|\\\\)") ? folder : ""               ; only real paths (no FTP, inside archives etc.)
    }

    ; Sends WM_COPYDATA (dwData = "GW", an ANSI command) to TC, which replies at once with WM_COPYDATA (dwData = "RW", UTF-16)
    ; to the window in wParam. Commands: A = which side is active (L / R); two letters: first L left / R right / S active / T the other,
    ; second P path / C file count / I caret position / N file name under the caret. "" if TC doesn't reply
    static _tcWaiting := false, _tcReply := "", _tcListening := false
    static TotalCmdAsk(hwnd, command, timeoutMs := 500) {
        if !QuickSwitch._tcListening {
            OnMessage(0x4A, (p*) => QuickSwitch._OnTotalCmdReply(p*))         ; WM_COPYDATA
            QuickSwitch._tcListening := true
        }
        request := Buffer(StrPut(command, "CP0"))
        StrPut(command, request, "CP0")
        copyData := Buffer(A_PtrSize * 3, 0)                                    ; COPYDATASTRUCT
        NumPut("UPtr", Ord("G") + 256 * Ord("W"), copyData, 0)
        NumPut("UInt", request.Size, copyData, A_PtrSize)
        NumPut("Ptr", request.Ptr, copyData, A_PtrSize * 2)
        QuickSwitch._tcReply := "", QuickSwitch._tcWaiting := true, accepted := 0
        DllCall("SendMessageTimeoutW", "Ptr", hwnd, "UInt", 0x4A, "Ptr", A_ScriptHwnd, "Ptr", copyData.Ptr
              , "UInt", 2, "UInt", timeoutMs, "UPtr*", &accepted)                 ; 2 = SMTO_ABORTIFHUNG: don't wait when TC hangs
        QuickSwitch._tcWaiting := false                                         ; TC replies while handling this message, so no extra wait
        return QuickSwitch._tcReply
    }

    static _OnTotalCmdReply(wParam, lParam, msg, hwnd) {
        if (!QuickSwitch._tcWaiting || NumGet(lParam, 0, "UPtr") != Ord("R") + 256 * Ord("W"))
            return                                                              ; not the reply waited for: leave it to other handlers (Everything)
        size := NumGet(lParam, A_PtrSize, "UInt"), data := NumGet(lParam, A_PtrSize * 2, "Ptr")
        reply := (data && size >= 2) ? StrGet(data, size // 2, "UTF-16") : ""  ; up to the terminating 0
        QuickSwitch._tcReply := reply, QuickSwitch._tcWaiting := false
        return true
    }

    ;---------------------------------------------------------------------------
    ; The folders in the panel
    ;---------------------------------------------------------------------------
    ; [{Path, Group: "open" / "recent", Tag}]: the active and other panel of each TC window, each Explorer window,
    ; recent folders (RecentFolders of them), without duplicates
    ; opus = false: without Directory Opus (the panel adds those a moment later, see _AddDOpusFolders)
    static MenuFolders(opus := true) {
        list := [], seen := Map()
        add(folder, group, tag) {
            if (folder = "" || seen.Has(StrLower(RTrim(folder, "\"))))
                return
            seen[StrLower(RTrim(folder, "\"))] := true
            list.Push({Path: folder, Group: group, Tag: tag})
        }
        for hwnd in WinGetList("ahk_class TTOTAL_CMD") {
            add(QuickSwitch.TotalCmdFolder(hwnd), "open", I18n.T("QuickSwitch.TagTC"))
            add(QuickSwitch.TotalCmdFolder(hwnd, true), "open", I18n.T("QuickSwitch.TagTCOther"))
        }
        if opus                                                             ; the source tab of each lister, then the other side
            for state in [1, 2]
                for tab in QuickSwitch.DOpusTabs()
                    if (tab.State = state)
                        add(tab.Path, "open", I18n.T(state = 1 ? "QuickSwitch.TagOpus" : "QuickSwitch.TagOpusOther"))
        try {
            for window in ComObject("Shell.Application").Windows
                try add(window.Document.Folder.Self.Path, "open", I18n.T("QuickSwitch.TagExplorer"))
        }
        count := QuickSwitch.Options.Has("RecentFolders") ? QuickSwitch.Options["RecentFolders"] : 10
        for folder in QuickSwitch.RecentFolders(count)
            add(folder, "recent", I18n.T("QuickSwitch.TagRecent"))
        return list
    }

    ; The newest limit folders of Windows "Recent items" (WindowsRecent); a file gives its folder.
    ; Only the newest limit * 6 shortcuts are looked at, not all of the hundreds there may be
    static RecentFolders(limit, recentDir := "") {
        folders := [], seen := Map()
        if (limit <= 0)
            return folders
        for item in WindowsRecent.Items(limit * 6, recentDir) {
            folder := RTrim(item.IsFolder ? item.Path : QuickSwitch._Parent(item.Path), "\")
            if (folder ~= "^[A-Za-z]:$")
                folder .= "\"
            if (folder = "" || seen.Has(StrLower(folder)))
                continue
            seen[StrLower(folder)] := true
            folders.Push(folder)
            if (folders.Length >= limit)
                break
        }
        return folders
    }

    static _Parent(target) {
        SplitPath(RTrim(target, "\"), , &dir)
        return dir
    }

    ; The current folder of an Explorer window; hwnd = 0 takes the frontmost Explorer window
    static ExplorerFolder(hwnd := 0) {
        if !hwnd
            hwnd := WinExist("ahk_class CabinetWClass")
        if !hwnd
            return ""
        try {
            for window in ComObject("Shell.Application").Windows
                if (window.HWND = hwnd)
                    return window.Document.Folder.Self.Path
        }
        return ""
    }

    ; Current folder of the TC / Opus / Explorer window nearest the top: {Path, Tag, Hwnd}, "" when none.
    ; Windows are listed front to back, so with a dialog in front this is the file manager used last
    ; (the jump hotkey, auto jump, "Copy to ..." in the actions)
    static FileManagerFolder() {
        for hwnd in WinGetList() {
            className := ""
            try className := WinGetClass("ahk_id " hwnd)
            switch className {
                case "TTOTAL_CMD"   : return {Path: QuickSwitch.TotalCmdFolder(hwnd), Tag: "Total Commander", Hwnd: hwnd}
                case "dopus.lister" : return {Path: QuickSwitch.DOpusFolder(hwnd), Tag: "Directory Opus", Hwnd: hwnd}
                case "CabinetWClass": return {Path: QuickSwitch.ExplorerFolder(hwnd), Tag: I18n.T("QuickSwitch.TagExplorer"), Hwnd: hwnd}
            }
        }
        return ""
    }

    static IsFileManager(hwnd) {
        className := ""
        try className := WinGetClass("ahk_id " hwnd)
        return className = "TTOTAL_CMD" || className = "dopus.lister" || className = "CabinetWClass"
    }

    static FolderOfWindow(hwnd) {
        if (!hwnd || !WinExist("ahk_id " hwnd))
            return ""
        switch WinGetClass("ahk_id " hwnd) {
            case "TTOTAL_CMD"   : return QuickSwitch.TotalCmdFolder(hwnd)
            case "dopus.lister" : return QuickSwitch.DOpusFolder(hwnd)
            case "CabinetWClass": return QuickSwitch.ExplorerFolder(hwnd)
        }
        return ""
    }

    ;---------------------------------------------------------------------------
    ; Directory Opus: its own "dopusrt.exe /info <file>,paths" writes every lister's tabs to a small XML file
    ; (read only, no clipboard). It runs only when an Opus window is open and a folder is actually needed,
    ; and a result is reused for DOpusCacheMs
    ;---------------------------------------------------------------------------
    static DOpusCacheMs := 1500, _dopusTabs := [], _dopusAt := 0

    ; Source (focused) tab of a lister; hwnd 0 = the active lister
    static DOpusFolder(hwnd := 0) => QuickSwitch.PickDOpusFolder(QuickSwitch.DOpusTabs(), hwnd)

    static PickDOpusFolder(tabs, hwnd := 0) {
        for tab in tabs
            if (tab.State = 1 && (hwnd ? tab.Lister = hwnd : tab.ActiveLister))
                return tab.Path
        for tab in tabs                                                     ; lister not found (e.g. Opus just opened it): the active one
            if (tab.State = 1 && tab.ActiveLister)
                return tab.Path
        return ""
    }

    static DOpusTabs() {
        if !(lister := WinExist("ahk_class dopus.lister"))
            return []
        if (QuickSwitch._dopusAt && A_TickCount - QuickSwitch._dopusAt < QuickSwitch.DOpusCacheMs)
            return QuickSwitch._dopusTabs
        started := Logger.Ms(), tabs := []
        try {
            SplitPath(WinGetProcessPath("ahk_id " lister), , &dir)
            xmlFile := A_Temp "\ALTRun.dopus-paths.xml"
            try FileDelete(xmlFile)
            Run('"' dir '\dopusrt.exe" /info "' xmlFile '",paths', , "Hide", &pid)
            if ProcessWaitClose(pid, 3)                                     ; still running after 3 s (returns its PID): never hang the dialog on it
                try ProcessClose(pid)
            tabs := QuickSwitch.ParseDOpusPaths(FileRead(xmlFile, "UTF-8"))
            try FileDelete(xmlFile)
        } catch as e {
            Logger.Error("QuickSwitch: cannot read the Directory Opus folders - " e.Message)
        }
        QuickSwitch._dopusTabs := tabs, QuickSwitch._dopusAt := A_TickCount
        if Logger.Enabled
            Logger.Debug(Format("Perf: Directory Opus paths {} ms, {} tabs", Round(Logger.Ms() - started), tabs.Length))
        return tabs
    }

    ; <path active_lister="1" active_tab="1" tab_state="1" side="1" lister="0x70e2a" tab="0x50e1a" ...>C:\Folder</path>
    ; tab_state: 1 = the source (focused) tab of its lister, 2 = the tab shown on the other side of a dual lister
    static ParseDOpusPaths(xml) {
        tabs := [], pos := 1, attrs := ""
        value(name) => RegExMatch(attrs, "\b" name '="([^"]*)"', &a) ? a[1] : ""
        while (pos := RegExMatch(xml, "s)<path\b([^>]*)>(.*?)</path>", &m, pos)) {
            pos += m.Len
            attrs := m[1], folder := m[2]
            for entity in [["&lt;", "<"], ["&gt;", ">"], ["&quot;", '"'], ["&apos;", "'"], ["&amp;", "&"]]   ; &amp; last
                folder := StrReplace(folder, entity[1], entity[2])
            lister := value("lister"), state := value("tab_state")
            tabs.Push({Path: Trim(folder), Lister: IsInteger(lister) ? Integer(lister) : 0, State: IsInteger(state) ? Integer(state) : 0
                , ActiveLister: value("active_lister") = "1"})
        }
        return tabs
    }

    ; Choosing a file in a dialog: for open and save dialogs alike, jump to its folder (without opening or saving for the user)
    static PickFile(filePath, dialog := 0) {
        SplitPath(filePath, , &dir)
        if (dir != "")
            QuickSwitch.SetDialogPath(RTrim(dir, "\") "\", dialog)
    }

    ; fill the old name back only when the file name box held a normal file name (not when empty, a path or a filter like *.txt)
    static ShouldRestoreName(saved) {
        saved := Trim(saved)
        return saved != "" && !RegExMatch(saved, "[\\/:*?]")
    }

    ; dialog: the dialog to jump in (the current window by default). Before simulating keyboard input, make sure it is the foreground window,
    ; otherwise the text would go to another window (e.g. switching back to the dialog failed and the panel is still in front)
    static SetDialogPath(folder, dialog := 0) {
        if (folder = "" || !FileExist(folder))
            return
        if !dialog
            dialog := WinExist("A")
        if !(dialog && WinExist("ahk_id " dialog))
            return
        Usage.Count("QuickSwitch")
        if QuickSwitch.IsBrowseWindow(dialog)                               ; folder tree only: select it there, pressing Enter would mean OK
            return QuickSwitch.SelectInBrowseDialog(folder, dialog)
        className := "", hasEdit := false
        try className := WinGetClass("ahk_id " dialog)
        try hasEdit := ControlGetHwnd("Edit1", "ahk_id " dialog) != 0
        method := QuickSwitch.JumpMethod(className, hasEdit)
        if (method = "edit") {
            try {
                saved := ""
                try saved := ControlGetText("Edit1", "ahk_id " dialog)         ; what the file name box held (e.g. a name Save As filled in)
                ControlFocus("Edit1", "ahk_id " dialog)
                ControlSetText(folder, "Edit1", "ahk_id " dialog)
                ControlSend("{Enter}", "Edit1", "ahk_id " dialog)
                if QuickSwitch.ShouldRestoreName(saved) {                   ; put it back after the jump, so the suggested name isn't lost
                    Sleep(150)
                    try ControlSetText(saved, "Edit1", "ahk_id " dialog)
                }
                return
            }
            method := (className = "#32770") ? "address" : "type"           ; the box refused: as if there were none
        }
        if !WinActive("ahk_id " dialog)                                     ; typing goes to the foreground window: only into the dialog
            return
        if (method = "address") {                                           ; a system dialog without the box: its address bar
            SendInput("!d")
            Sleep(60)
        }
        SendText(folder)                                                    ; dialogs a program draws itself (WPS / Qt, GTK, Java...): the
        SendInput("{Enter}")                                                ; file name box has the focus; a folder path + Enter goes there
    }

    ; How to put the folder into a dialog:
    ;   edit    = into its file name box (Edit1) + Enter: system dialogs and programs with real edit controls
    ;   address = a system dialog (#32770) without that box: Alt+D (address bar), type, Enter
    ;   type    = a dialog a program draws itself (WPS / Qt, GTK, Java...): type into the focused file name box, Enter.
    ;             No Alt+D there: it isn't the address bar and may press some button of the dialog
    static JumpMethod(className, hasEdit) {
        if hasEdit
            return "edit"
        return (className = "#32770") ? "address" : "type"
    }
}
