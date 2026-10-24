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
; Pages are built with a few small helpers, one line per setting:
;   _Check("General.LaunchAtLogin", "Prefs.LaunchAtLogin", , , "Prefs.Group.Startup")  check box (the last one is the group label in the left column)
;   _Field("Features.Snippets.Keyword", "Prefs.SnippetKeyword", "M")   text box, width "S" / "M" / "L"
;   _Field("General.Hotkey", "Prefs.Hotkey", "K", "hotkey")     a box that records a hotkey (HotkeyBox): shows Alt+Space, saves !Space
;   _Choice("General.Language", "Prefs.Language", values, display texts)
;   _LinesButton("Features.Applications.Folders", "Prefs.AppFolders")  one item per line: summary + "Edit..." window
;   _Csv("Features.FileSearch.Keywords", "Prefs.FileKeywords", "M")   array <-> comma-separated
;   _Button("Prefs.ClipClear", fn) / _Section("Prefs.QuickSwitch") / _Info(label, text)
;   _List(...)                                                  list + Add / Edit / Delete (ItemEditor)
; Paths are keys in ALTRun.json joined with ".".
; The layout is a two-column form: right-aligned labels on the left (LabelW), every control starts at _InputX(), so every page lines up.
; Like Alfred, each setting has a gray description below it: when I18n has "<label key>.Desc" (e.g.
; Prefs.LaunchAtLogin.Desc) it is shown automatically, nothing extra on the page.
;
; Usage:
;   PreferencesWindow.Show([page, x, y])
;   PreferencesWindow.AddCommands(paths)   "Send to" while Preferences is open: added to the Commands page
;===============================================================================

class PreferencesWindow {
    static Gui := "", Working := "", Pages := [], Binds := [], PageList := ""
    static _lists := Map()                                                  ; setting path of a list page (e.g. "CustomCommands") -> the method that adds / edits items
    ; the order of the pages (matching _BuildPages, checked by a test); search entries like "ALTRun Preferences: Appearance" come from it without opening the window
    static PageKeys := ["Prefs.Page.General", "Prefs.Page.Window", "Prefs.Page.Keyboard", "Prefs.Page.Appearance", "Prefs.Page.Features", "Prefs.Page.Applications"
        , "Prefs.Page.FileSearch", "Prefs.Page.FileIndex", "Prefs.Page.Commands", "Prefs.Page.Snippets", "Prefs.Page.Clipboard", "Prefs.Page.WebSearch"
        , "Prefs.Page.Calculator", "Prefs.Page.Scripts", "Prefs.Page.Hotkeys", "Prefs.Page.QuickSwitch", "Prefs.Page.QSPanel", "Prefs.Page.DateStamp"
        , "Prefs.Page.Usage", "Prefs.Page.Advanced"]
    static _page := 0, _y := 0
    static _dirty := false, _ready := false, ApplyButton := ""
    ; Two-column form (as in Alfred / Listary): right-aligned labels on the left, and a page's controls start at one vertical line (_InputX);
    ; the left column's width is set per page by its labels (_BeginPage); input boxes have only three widths, check boxes get group labels in the left column,
    ; buttons share one size, and gray descriptions align with the controls
    static ContentX := 190, ContentW := 560, LabelW := 190
    static MinWindowWidth := 450                                            ; the search window: room for the status bar's four parts
    static WidthS := 70, WidthM := 220, ButtonW := 200, ButtonH := 26       ; WidthL = the whole control column (_InputW)
    static WidthK := 150                                                    ; hotkey box ("K")
    static _pairX := 0, _pairFirstW := 0, _pairSecondW := 0, _pairSizes := ["", ""]
    static ButtonY := 619                                                   ; where the bottom buttons are; page content must stay above (y < ButtonY - 10)
    static LabelMax := 250                                                  ; widest the left column gets; longer labels wrap
    static _positionReset := false
    static _pageKey := "", _wider := Map()                                  ; the page being built; pages whose labels don't fit -> the left column width they need
    static _labelWidths := Map()                                            ; "language page" -> the widened left column (remembered once opened)

    ; pageIndex: page number, or the page's key (e.g. "Prefs.Page.Advanced")
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
        PreferencesWindow.Binds := [], PreferencesWindow._wider := Map(), PreferencesWindow._lists := Map(), PreferencesWindow._helpButton := ""
        PreferencesWindow._dirty := false, PreferencesWindow._ready := false, PreferencesWindow._positionReset := false, PreferencesWindow._page := 0
        ; Pages are built when needed (_BuildPage): opening builds only the page shown, the others when first clicked. Opening used to build all 19 pages,
        ; over 300 ms each time (a second or two on old PCs), and most pages are never looked at
        PreferencesWindow.Pages := []
        for index, key in PreferencesWindow.PageKeys
            PreferencesWindow.Pages.Push({Key: key, Name: I18n.T(key), Wiki: PreferencesWindow.WikiPage(key), Controls: [], Built: false
                , Builder: PreferencesWindow.PageBuilders[index]})

        g := Gui("-MinimizeBox", I18n.T("Prefs.Title"))
        g.SetFont("s9", ThemeManager.FontName())
        g.OnEvent("Close", (*) => PreferencesWindow.Cancel())                ; true = don't close (chose to keep editing)
        g.OnEvent("Escape", (*) => PreferencesWindow.Cancel())
        PreferencesWindow.Gui := g

        ; The page list on the left is built first: the Tab order is list -> the page's controls -> bottom buttons (pages built later have their controls moved before the bottom buttons, see _BuildPage)
        ; A ListView rather than a ListBox: an icon before each page (drawn with the system icon font, see NavIcons); 0x20 full-row select, 0x10000 double buffering
        pageList := g.AddListView("x12 y12 w160 h" (PreferencesWindow.ButtonY - 24) " -Hdr -Multi LV0x10020", [""])
        if !IsInteger(pageIndex)
            pageIndex := PreferencesWindow._PageIndex(pageIndex)
        pageIndex := Max(1, Min(pageIndex, PreferencesWindow.Pages.Length))
        PreferencesWindow._BuildPage(pageIndex)
        if !IsObject(PreferencesWindow.Gui)                                 ; the window was closed while a page was being built (see _BuildPage)
            return

        names := []
        for page in PreferencesWindow.Pages
            names.Push(page.Name)
        icons := NavIcons.ImageList(PreferencesWindow.PageKeys)             ; the image list's height is the row height (a roomier sidebar)
        if icons
            pageList.SetImageList(icons, 1)
        for index, name in names
            pageList.Add("Icon" index, name)
        pageList.ModifyCol(1, "AutoHdr")                                    ; the only column fills the list width (a column narrower than the list would add a column separator on the right)
        DllCall("uxtheme\SetWindowTheme", "Ptr", pageList.Hwnd, "Str", "Explorer", "Ptr", 0)   ; the same hover / selection look as Explorer
        pageList.OnNotify(-101, (ctrl, lParam) => PreferencesWindow._OnPageRow(lParam))   ; LVN_ITEMCHANGED (the ItemSelect event sometimes doesn't arrive)
        pageList.OnNotify(-12, (ctrl, lParam) => PreferencesWindow._OnPageListDraw(lParam))   ; NM_CUSTOMDRAW: no dotted focus rectangle
        for code in [-2, -3, -5, -6]                                        ; NM_CLICK / NM_DBLCLK / NM_RCLICK / NM_RDBLCLK (a fast second click counts as a double-click)
            pageList.OnNotify(code, (ctrl, lParam) => PreferencesWindow._OnPageClick(lParam))
        PreferencesWindow.PageList := pageList

        buttonY := " y" PreferencesWindow.ButtonY
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
        g.Show((IsInteger(x) && IsInteger(y) ? "x" x " y" y " " : "") "w765 h" (PreferencesWindow.ButtonY + 41))
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

    ; The selected row is shown by its highlight only, without a dotted focus rectangle (like the navigation in Windows Settings). Windows decides whether to show focus rectangles
    ; by "whether the keyboard was used recently", so the first time Preferences opens the dots sometimes show and sometimes don't; CDIS_FOCUS is removed before each row is drawn
    static _OnPageListDraw(lParam) {
        stage := NumGet(lParam, 3 * A_PtrSize, "UInt")                      ; NMCUSTOMDRAW.dwDrawStage
        if (stage = 0x1)                                                    ; CDDS_PREPAINT
            return 0x20                                                     ; CDRF_NOTIFYITEMDRAW
        if (stage = 0x10001) {                                              ; CDDS_ITEMPREPAINT
            offset := 6 * A_PtrSize + 16                                    ; NMCUSTOMDRAW.uItemState
            NumPut("UInt", NumGet(lParam, offset, "UInt") & ~0x10, lParam, offset)   ; CDIS_FOCUS
        }
        return 0                                                            ; CDRF_DODEFAULT
    }

    ; A click on the blank area below the list (not on a row) makes ListView deselect: select the current page again after the mouse is released.
    ; Not right at deselecting: clicking another row also deselects the previous row first, and that row would flash
    static _OnPageClick(lParam) {
        if (NumGet(lParam, 3 * A_PtrSize, "Int") < 0 && !PreferencesWindow.PageList.GetNext())   ; NMITEMACTIVATE.iItem = -1
            PreferencesWindow.PageList.Modify(PreferencesWindow._page, "Select Focus")
    }

    static _PageIndex(nameKey) {
        for index, page in PreferencesWindow.Pages
            if (page.Key = nameKey)
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
                          , "Prefs.Page.FileIndex", "File-Search", "Prefs.Page.Commands", "Custom-Commands", "Prefs.Page.Snippets", "Snippets"
                          , "Prefs.Page.Clipboard", "Clipboard", "Prefs.Page.WebSearch", "Web-Search", "Prefs.Page.Calculator", "Calculator"
                          , "Prefs.Page.Scripts", "Scripts", "Prefs.Page.Hotkeys", "Hotkeys", "Prefs.Page.QuickSwitch", "Quick-Switch"
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
        saved := PreferencesWindow._y
        PreferencesWindow._y := y + (h - 16) // 2
        try PreferencesWindow._Add("Picture", "x" (x + 18 + textW + 6) " w16 h16", "HICON:" hIcon)
        PreferencesWindow._y := saved
    }

    ; Turning on "also in administrator windows" creates the scheduled task, turning it off deletes it (one UAC prompt each). If creating fails (UAC canceled), explain and turn the option off again;
    ; false = don't save (the user chose to go back and change it)
    static _ApplyCapsLockHelper(general) {
        was := AppSettings.General.Has("CapsLockElevated") && AppSettings.General["CapsLockElevated"]
        now := general.Has("CapsLockElevated") && general["CapsLockElevated"]
        if (now && !CapsLockSwitch.IsMode(general["CapsLock"]))
            general["CapsLockElevated"] := now := 0                         ; the option has no effect while CapsLock works as usual
        if (now && !was && !CapsLockHelper.Install()) {
            MsgBox(I18n.T("Prefs.CapsLockElevatedFailed"), I18n.T("Prefs.Title"), "Icon! Owner" PreferencesWindow.Gui.Hwnd)
            general["CapsLockElevated"] := 0
        } else if (was && !now) {
            CapsLockHelper.Uninstall()
        }
        return true
    }

    ; Command line "-Preferences page [x y]" (reopening after Apply) -> {Page, X, Y}
    static ParseArgs(args) {
        value(i) => (args.Length >= i && IsInteger(args[i])) ? Integer(args[i]) : ""
        page := value(2)
        return {Page: (page = "") ? 1 : page, X: value(3), Y: value(4)}
    }

    ; reopen: Apply = reload and reopen on the same page at the same position; OK = only reload in the background
    static Save(reopen := true) {
        for bind in PreferencesWindow.Binds
            PreferencesWindow.SetPath(PreferencesWindow.Working, bind.Path, bind.Read.Call())
        general := PreferencesWindow.Working["General"]
        appearance := PreferencesWindow.Working["Appearance"]
        appearance["Width"] := Max(PreferencesWindow.MinWindowWidth, appearance["Width"])
        if (general["Hotkey"] = "")
            general["Hotkey"] := "!Space"
        clash := PreferencesWindow.DuplicateHotkey(PreferencesWindow._GlobalHotkeys(PreferencesWindow.Working))
        if (IsObject(clash) && MsgBox(I18n.T("Prefs.HotkeyClash", HotkeyBox.Label(clash.Key), clash.First, clash.Second)
                , I18n.T("Prefs.Title"), "YesNo Icon! Default2 Owner" PreferencesWindow.Gui.Hwnd) != "Yes")
            return
        if !PreferencesWindow._positionReset                                ; the search window was dragged after Preferences opened: use the new position
            appearance["Position"] := PreferencesWindow.DeepCopy(AppSettings.Appearance["Position"])

        if !PreferencesWindow._ApplyCapsLockHelper(general)
            return
        stale := PreferencesWindow.StaleIndexes(AppSettings.Data, PreferencesWindow.Working)
        AppSettings.Data := PreferencesWindow.Working
        if !AppSettings.Save()
            return
        if stale.Apps                                                       ; the index scope changed: rescan in the background after restarting (the old index is used until then)
            ApplicationProvider.ExpireCache()
        if stale.Files
            FileIndex.ExpireCache()
        page := PreferencesWindow._page
        WinGetPos(&x, &y, , , PreferencesWindow.Gui.Hwnd)
        PreferencesWindow.Close()
        App.Restart(reopen ? "-Preferences " page " " x " " y : "-Reloaded")
    }

    ;---------------------------------------------------------------------------
    ; Pages
    ;---------------------------------------------------------------------------
    ; how each page is built, one per PageKeys entry
    static PageBuilders := ["_BuildGeneral", "_BuildWindow", "_BuildKeyboard", "_BuildAppearance", "_BuildFeatures", "_BuildApplications"
        , "_BuildFileSearch", "_BuildFileIndex", "_BuildCommands", "_BuildSnippets", "_BuildClipboard", "_BuildWebSearch"
        , "_BuildCalculator", "_BuildScripts", "_BuildHotkeys", "_BuildQuickSwitch", "_BuildQuickSwitchPanel", "_BuildDateStamp"
        , "_BuildUsage", "_BuildAdvanced"]
    static _current := "", _helpButton := ""

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
            PreferencesWindow._wider := Map(), binds := PreferencesWindow.Binds.Length
            try PreferencesWindow.%page.Builder%()
            catch as e {
                if IsObject(PreferencesWindow.Gui)
                    throw e
                Logger.Debug("PreferencesWindow: closed while building " page.Key)
                return false
            }
            if (!PreferencesWindow._wider.Has(page.Key) || A_Index = 2)
                break
            PreferencesWindow._labelWidths[I18n.Lang " " page.Key] := PreferencesWindow._wider[page.Key]
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

    static _BuildGeneral() {
        PreferencesWindow._BeginPage("Prefs.Page.General", 120)
        PreferencesWindow._Pair(["General.Hotkey", "Prefs.Hotkey", "K", "hotkey"], ["General.SecondaryHotkey", "Prefs.SecondaryHotkey", "K", "hotkey"])
        PreferencesWindow._Field("General.SelectionHotkey", "Prefs.SelectionHotkey", "K", "hotkey")
        PreferencesWindow._Pair(["General.DoubleTap", "Prefs.DoubleTap", "K", "choice", ["", "Ctrl", "Shift"], [I18n.T("Prefs.DoubleTap.None"), I18n.T("Prefs.DoubleTap.Ctrl"), I18n.T("Prefs.DoubleTap.Shift")]]
            , ["General.Language", "Prefs.Language", "K", "choice", PreferencesWindow._LanguageValues(), PreferencesWindow._LanguageLabels()])
        PreferencesWindow._Gap()
        for row in [["LaunchAtLogin", "Prefs.LaunchAtLogin", "Prefs.Group.Startup"], ["ShowTrayIcon", "Prefs.ShowTrayIcon", ""]
                   , ["SendToMenu", "Prefs.SendToMenu", "Prefs.Group.Integration"], ["StartMenuShortcut", "Prefs.StartMenu", ""]
                   , ["CheckForUpdates", "Prefs.CheckUpdates", "Prefs.Group.Updates"]]   ; writing the debug log is under Advanced → Diagnostics
            PreferencesWindow._Check("General." row[1], row[2], , , row[3])
        PreferencesWindow._Gap()
        PreferencesWindow._Field("General.FileManager", "Prefs.FileManager", "L", "file")
    }

    ; Search window behavior (the settings are still in General, just on a page of their own)
    static _BuildWindow() {
        PreferencesWindow._BeginPage("Prefs.Page.Window", 16)                ; arranged in sections, each section's content indented a little
        PreferencesWindow._Section("Prefs.Group.Window")
        PreferencesWindow._Check("General.HideOnDeactivate", "Prefs.HideOnDeactivate")
        PreferencesWindow._Check("General.KeepLastQuery", "Prefs.KeepLastQuery")
        PreferencesWindow._Check("Appearance.StatusBar", "Prefs.StatusBar")       ; the usage tips are in the status bar (in the empty search box when the status bar is off): kept together
        PreferencesWindow._Check("General.ShowTips", "Prefs.ShowTips")
        PreferencesWindow._Section("Prefs.Section.EmptyBox")
        placeholder := PreferencesWindow._InlineField("Appearance.Placeholder", "Prefs.Placeholder", "M")
        Win.SetCueBanner(placeholder.Hwnd, I18n.T("Search.Placeholder"))  ; empty shows the default text in gray
        recent := PreferencesWindow._InlineField("Features.Recent.RecentCount", "Prefs.RecentCount", "S", "number")
        PreferencesWindow._AlignLeft(placeholder, recent)                   ; the input boxes on this page share one left edge
    }

    ; Keyboard: input method switching (Caps Lock, English when shown) and the keys in the search box.
    ; A two-column form like the General page: labels right-aligned on the left, every control starting on one line
    static _BuildKeyboard() {
        PreferencesWindow._BeginPage("Prefs.Page.Keyboard", 150)
        PreferencesWindow._Section("Prefs.Group.InputMethod")
        capsLabels := [I18n.T("Prefs.CapsLock.Off"), I18n.T("Prefs.CapsLock.Layout"), I18n.T("Prefs.CapsLock.Mode")]
        capsLock := PreferencesWindow._Choice("General.CapsLock", "Prefs.CapsLock", ["", "Layout", "Mode"], capsLabels)   ; Caps Lock switches the input method (CapsLockSwitch)
        notice := PreferencesWindow._Check("General.CapsLockNotify", "Prefs.CapsLockNotify")   ; hint when Caps Lock turns on / off
        elevated := PreferencesWindow._Check("General.CapsLockElevated", "Prefs.CapsLockElevated")   ; also in administrator windows (CapsLockHelper)
        controls := PreferencesWindow._current.Controls
        elevatedDesc := controls[controls.Length]                           ; the description just added (sized for two lines, so the shorter text fits too)
        PreferencesWindow._ShieldAfter(elevated)                            ; turning it on / off asks for UAC
        update := (*) => (notice.Enabled := elevated.Enabled := capsLock.Value > 1
            , elevatedDesc.Text := I18n.T(elevated.Enabled ? "Prefs.CapsLockElevated.Desc" : "Prefs.CapsLockElevated.NeedsMode"))   ; when grayed out, say why
        update()
        capsLock.OnEvent("Change", update)
        englishLabels := [I18n.T("Prefs.EnglishInput.Off"), I18n.T("Prefs.EnglishInput.Keyboard"), I18n.T("Prefs.EnglishInput.Mode")]
        english := PreferencesWindow._Choice("General.SwitchToEnglishInput", "Prefs.EnglishInput", [0, 1, 2], englishLabels)   ; switch to English when shown
        restore := PreferencesWindow._Check("General.RestoreInput", "Prefs.RestoreInput")
        restore.Enabled := english.Value > 1
        english.OnEvent("Change", (*) => restore.Enabled := english.Value > 1)
        PreferencesWindow._Section("Prefs.Group.SearchKeys")
        PreferencesWindow._Check("General.SpaceToRun", "Prefs.SpaceToRun", , , "Prefs.Group.SpaceKey")
        tabLabels := [I18n.T("Prefs.TabKey.Complete"), I18n.T("Prefs.TabKey.Next")]
        tab := PreferencesWindow._Choice("General.TabKey", "Prefs.TabKey", ["Complete", "Next"], tabLabels)
        PreferencesWindow._Field("General.HistorySize", "Prefs.HistorySize", "S", "number")   ; Ctrl+↑ / ↓ recall earlier searches
        PreferencesWindow._SameWidth(capsLock, english, tab)
    }

    ; Drop-downs in one column get one width: wide enough for the longest option of any of them, at least "M"
    ; (a button in the column just takes the same width)
    static _SameWidth(ctrls*) {
        width := PreferencesWindow.WidthM
        for ctrl in ctrls
            if (ctrl.Type = "DDL")
                width := Max(width, PreferencesWindow._ChoiceWidth(ctrl))
        width := Min(width, PreferencesWindow._InputW())
        for ctrl in ctrls
            ctrl.Move(, , width)
    }

    static _BuildAppearance() {
        PreferencesWindow._BeginPage("Prefs.Page.Appearance", 150)
        appearance := PreferencesWindow.Working["Appearance"]
        if ((base := ThemeManager.LegacyCompact(appearance["Theme"])) != "") {  ; a former compact theme: select the matching theme, density at least Normal (takes effect after saving)
            appearance["Theme"] := base
            if (appearance["Density"] = "Comfortable")
                appearance["Density"] := "Normal"
        }
        PreferencesWindow._Section("Prefs.Section.Theme")
        themes := ThemeManager.Names()
        gallery := PreferencesWindow._ThemeGallery("Appearance.Theme", themes, 164)   ; one and a half rows of thumbnails (shows it can scroll), leaving room for "Compact" below
        PreferencesWindow._Buttons("Prefs.Group.CustomThemes"
            , ["Prefs.CopyTheme", (*) => PreferencesWindow._CopyTheme(gallery, themes)]
            , ["Prefs.OpenThemes", (*) => PreferencesWindow._OpenFolder(ThemeManager.UserDir)])
        PreferencesWindow._Section("Prefs.Section.Size")
        PreferencesWindow._Pair(["Appearance.Width", "Prefs.Width", "S", "number"]
            , ["Appearance.VisibleRows", "Prefs.VisibleRows", "S", "choice", [1, 2, 3, 4, 5, 6, 7, 8, 9], ["1", "2", "3", "4", "5", "6", "7", "8", "9"]])   ; a list: nothing out of range to type
        for ctrl in PreferencesWindow._current.Controls                      ; the drop-down _Pair just added
            if (ctrl.Type = "DDL")
                visible := ctrl
        density := PreferencesWindow._Choice("Appearance.Density", "Prefs.Density", ThemeManager.Densities   ; any theme can be scaled (ThemeManager.Scaled)
            , [I18n.T("Prefs.Density.Comfortable"), I18n.T("Prefs.Density.Normal"), I18n.T("Prefs.Density.Compact")])
        PreferencesWindow._Check("Appearance.Fade", "Prefs.Fade", , , "Prefs.Group.Animation")
        PreferencesWindow._Section("Prefs.WindowPosition")
        showOn := PreferencesWindow._Choice("Appearance.ShowOn", "Prefs.ShowOn", ["Mouse", "Primary", "Active"]
            , [I18n.T("Prefs.ShowOn.Mouse"), I18n.T("Prefs.ShowOn.Primary"), I18n.T("Prefs.ShowOn.Active")])
        PreferencesWindow._Check("Appearance.RememberPosition", "Prefs.RememberPosition", , , "Prefs.Group.Dragging")
        reset := PreferencesWindow._Button("Prefs.ResetPosition", (button, *) => PreferencesWindow._ResetPosition(button))
        PreferencesWindow._SameWidth(density, showOn, reset)                ; one column: the two drop-downs and the button below them
        ; "Visible results" a little wider than its longest option, and that column ends on the same line as it
        visible.GetPos(&vx, , &vw), density.GetPos(&dx, , &dw)
        right := Min(Max(vx + vw + 16, dx + dw), PreferencesWindow.ContentX + PreferencesWindow.ContentW)
        visible.Move(, , right - vx)
        for ctrl in [density, showOn, reset]
            ctrl.Move(, , right - dx)
    }

    ; the remembered position goes back to the default (takes effect when saved)
    static _ResetPosition(button) {
        PreferencesWindow.Working["Appearance"]["Position"] := AppSettings.Defaults()["Appearance"]["Position"]
        PreferencesWindow._positionReset := true
        PreferencesWindow.MarkDirty()
        button.Enabled := false
    }

    ; Options for the UI language: Auto + the languages in Resources\Lang (each shown in its own language)
    static _LanguageValues() {
        values := ["auto"]
        for language in I18n.Languages()
            values.Push(language[1])
        return values
    }

    static _LanguageLabels() {
        labels := [I18n.T("Prefs.Language.auto")]
        for language in I18n.Languages()
            labels.Push(language[2])
        return labels
    }

    ; built-in themes show their translated names, user themes their file names
    static _ThemeLabel(themeName) {
        label := I18n.T("Theme." themeName)
        return (label = "Theme." themeName) ? themeName : label
    }

    ; Theme list: a thumbnail per theme (ThemePreview, a mini search window in the theme's colors), picked by clicking like in Alfred / Listary.
    ; A ListView in large-icon view: its own scroll bar and arrow-key selection, room for any number of themes. Returns the ListView (row = index in themes)
    static _ThemeGallery(path, themes, height) {
        thumbW := Win.Scale(112), thumbH := Win.Scale(76)
        list := PreferencesWindow._Add("ListView", "x" PreferencesWindow.ContentX " w" PreferencesWindow.ContentW " h" height
            . " Icon -Multi -Hdr", ["Theme"])
        DllCall("uxtheme\SetWindowTheme", "Ptr", list.Hwnd, "WStr", "Explorer", "Ptr", 0)   ; the selected item gets a light frame instead of tinting the thumbnail blue
        list.SetImageList(ThemePreview.ImageList(themes, thumbW, thumbH), 0)   ; 0 = large icons
        SendMessage(0x1035, 0, PreferencesWindow.GalleryColumnWidth(list, thumbW) | ((thumbH + Win.Scale(30)) << 16), list.Hwnd)   ; LVM_SETICONSPACING
        current := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        selected := 1
        for index, themeName in themes {
            list.Add("Icon" index, PreferencesWindow._ThemeLabel(themeName))
            if (themeName = current)
                selected := index
        }
        list.Modify(selected, "Select Focus Vis")
        list.OnNotify(-101, (ctrl, lParam) => PreferencesWindow._OnGalleryChange(lParam))   ; LVN_ITEMCHANGED (selection by the program counts too)
        PreferencesWindow._Bind(path, () => themes[list.GetNext() || selected])
        PreferencesWindow._Below(8, list)
        return list
    }

    ; Width of each cell of the theme list: exactly GalleryColumns thumbnails per row, spread evenly (no big gap on the right), with 2 pixels to spare so the last doesn't wrap.
    ; more than two rows of themes have a vertical scroll bar: subtract it first
    static GalleryColumns := 4
    static GalleryColumnWidth(list, thumbW) {
        rect := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", list.Hwnd, "Ptr", rect)
        width := NumGet(rect, 8, "Int") - (DllCall("GetWindowLong", "Ptr", list.Hwnd, "Int", -16, "Int") & 0x200000 ? 0 : SysGet(2))   ; subtract SM_CXVSCROLL while WS_VSCROLL hasn't appeared yet
        return Max(thumbW + Win.Scale(8), (width - 2) // PreferencesWindow.GalleryColumns)
    }

    ; The table's last column fills the remaining width (like AutoHdr) but first leaves room for the vertical scroll bar: Windows doesn't draw it for a list not yet shown,
    ; so AutoHdr counts without it and a horizontal scroll bar appears once shown. Call again when the list contents change (add, delete, filter)
    static FitColumns(listView) {
        static LVM_GETCOLUMNWIDTH := 0x101D, LVM_SETCOLUMNWIDTH := 0x101E, LVM_GETHEADER := 0x101F, HDM_GETITEMCOUNT := 0x1200
        columns := SendMessage(HDM_GETITEMCOUNT, 0, 0, SendMessage(LVM_GETHEADER, 0, 0, listView.Hwnd))
        width := PreferencesWindow.ListWidth(listView)
        Loop columns - 1
            width -= SendMessage(LVM_GETCOLUMNWIDTH, A_Index - 1, 0, listView.Hwnd)
        SendMessage(LVM_SETCOLUMNWIDTH, columns - 1, Max(width, 50), listView.Hwnd)
    }

    ; The width the table's columns should add up to: the client width, minus the scroll bar's width when the rows don't fit on one screen (the scroll bar is there or will be)
    static ListWidth(listView) {
        static LVM_GETCOUNTPERPAGE := 0x1028
        rect := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", listView.Hwnd, "Ptr", rect)
        hasVScroll := DllCall("GetWindowLong", "Ptr", listView.Hwnd, "Int", -16, "Int") & 0x200000   ; WS_VSCROLL
        needVScroll := hasVScroll || listView.GetCount() > SendMessage(LVM_GETCOUNTPERPAGE, 0, 0, listView.Hwnd)
        return NumGet(rect, 8, "Int") + (hasVScroll ? SysGet(2) : 0) - (needVScroll ? SysGet(2) : 0)   ; SM_CXVSCROLL
    }

    ; another theme was selected in the theme list: "Apply" becomes available
    static _OnGalleryChange(lParam) {
        static LVIS_SELECTED := 0x2
        offset := A_PtrSize * 3 + 4                                         ; NMLISTVIEW: hdr, iItem, iSubItem, uNewState
        newState := NumGet(lParam, offset + 4, "UInt"), oldState := NumGet(lParam, offset + 8, "UInt")
        if ((newState & LVIS_SELECTED) && !(oldState & LVIS_SELECTED))
            PreferencesWindow.MarkDirty()
    }

    ; Copies the selected theme (expanded to all keys) to Themes\<new name>.json, opens it in Notepad and selects it in the list
    static _CopyTheme(gallery, themes) {
        source := themes[gallery.GetNext() || 1]
        if (source = "System")
            source := ThemeManager.SystemUsesDark() ? "Dark" : "Light"
        theme := ThemeManager.Resolve(source)
        if !IsObject(theme)
            return
        PreferencesWindow.Gui.Opt("+OwnDialogs")
        answer := InputBox(I18n.T("Prefs.CopyThemePrompt"), I18n.T("Prefs.CopyTheme"), "w340 h130", source " Custom")
        themeName := Trim(RegExReplace(answer.Value, '[\\/:*?"<>|]'))
        if (answer.Result != "OK" || themeName = "")
            return
        themeFile := ThemeManager.UserDir "\" themeName ".json"
        if (FileExist(themeFile) && MsgBox(I18n.T("Prefs.ThemeExists", themeName), App.Name, "YesNo Icon! Default2") != "Yes")
            return
        try {
            JSON.WriteFile(themeFile, theme)
        } catch as e {
            return MsgBox(e.Message, App.Name, 48)
        }
        index := 0
        for existing, value in themes
            if (value = themeName)
                index := existing
        if !index {
            themes.Push(themeName)
            imageList := SendMessage(0x1002, 0, 0, gallery.Hwnd)           ; LVM_GETIMAGELIST (LVSIL_NORMAL)
            DllCall("comctl32\ImageList_GetIconSize", "Ptr", imageList, "Int*", &thumbW := 0, "Int*", &thumbH := 0)
            gallery.Add("Icon" (ThemePreview.AddTo(imageList, themeName, thumbW, thumbH) + 1), themeName)
            index := themes.Length
        }
        gallery.Modify(0, "-Select")
        gallery.Modify(index, "Select Focus Vis")
        Run('notepad.exe "' themeFile '"')
    }

    static _BuildFeatures() {
        PreferencesWindow._BeginPage("Prefs.Page.Features", 140)
        PreferencesWindow._Section("Prefs.EnabledFeatures")
        features := ["Applications", "CustomCommands", "Snippets", "Clipboard", "Calculator", "WebSearch", "Bookmarks", "Windows", "Recent", "FileSearch", "Scripts", "Terminal", "System", "Help"]
        startY := PreferencesWindow._y, columnW := PreferencesWindow._InputW() // 2      ; two columns: the English names are long, three would wrap
        rowH := 20                                                          ; row spacing: at least 4 more than a check box (Chinese and Japanese check boxes are taller)
        for index, feature in features {
            column := Mod(index - 1, 2), row := (index - 1) // 2
            PreferencesWindow._y := startY + row * rowH
            check := PreferencesWindow._Check("Features." feature ".Enabled", "Prefs.Feature." feature
                , PreferencesWindow._InputX() + column * columnW, , (index = 1) ? "Prefs.Group.SearchFeatures" : "", columnW)
            if (index = 1) {
                check.GetPos(, , , &checkH)
                rowH := Max(rowH, checkH + 4)
            }
        }
        PreferencesWindow._y := startY + (features.Length // 2) * rowH        ; last row: Large Type (not a search feature), Windows Settings pages (part of the system commands)
        PreferencesWindow._Check("Features.LargeType.Enabled", "Prefs.Feature.LargeType", PreferencesWindow._InputX(), , "", columnW)
        PreferencesWindow._y := startY + (features.Length // 2) * rowH
        PreferencesWindow._Check("Features.System.SettingsPages", "Prefs.SettingsPages", PreferencesWindow._InputX() + columnW, , "", columnW)
        PreferencesWindow._y := startY + Ceil((features.Length + 1) / 2) * rowH
        PreferencesWindow._Section("Prefs.Section.FeatureOptions")
        PreferencesWindow._Check("Features.System.ConfirmActions", "Prefs.ConfirmActions", , , "Prefs.Feature.System")
        PreferencesWindow._LinesButton("Features.System.Hidden", "Prefs.SysHidden")
        PreferencesWindow._y += 4                                           ; a little less than _Gap: this page is full
        PreferencesWindow._Pair(["Features.Terminal.Prefix", "Prefs.TerminalPrefix", "S"]
            , ["Features.Terminal.Shell", "Prefs.TerminalShell", "M", "choice", ["cmd", "powershell", "pwsh", "wt"], ["Command Prompt (cmd)", "Windows PowerShell", "PowerShell 7 (pwsh)", "Windows Terminal (wt)"]])
        for row in [["Bookmarks", "Prefs.Feature.Bookmarks"], ["Windows", "Prefs.Feature.Windows"]] {   ; bookmarks and windows: also in normal search + a keyword that searches only them
            rowY := PreferencesWindow._y
            PreferencesWindow._Check("Features." row[1] ".InDefaultResults", "Prefs.ShowInNormal", , , row[2], columnW)
            field := PreferencesWindow._SideField(rowY, PreferencesWindow._InputX() + columnW, "Features." row[1] ".Keyword", "Prefs.FeatureKeyword")   ; aligned with the right column of the features above
            field.GetPos(, &fieldY, , &fieldH)
            PreferencesWindow._y := Max(rowY + 25, fieldY + fieldH + 8)     ; keyword boxes on two rows at least 5 apart (Chinese / Japanese input boxes are taller)
        }
        PreferencesWindow._y -= 3
        PreferencesWindow._Desc("Prefs.FeatureKeyword", PreferencesWindow._InputX())
    }

    static _BuildApplications() {
        PreferencesWindow._BeginPage("Prefs.Page.Applications", 130)
        PreferencesWindow._Section("Prefs.Section.AppIndex")
        PreferencesWindow._LinesButton("Features.Applications.Folders", "Prefs.AppFolders")
        PreferencesWindow._Csv("Features.Applications.FileTypes", "Prefs.AppFileTypes", "L")
        PreferencesWindow._Field("Features.Applications.Exclude", "Prefs.AppExclude", "L")
        PreferencesWindow._FieldButton("Features.Applications.Depth", "Prefs.AppDepth", "Prefs.RebuildIndex", (button, *) => PreferencesWindow._RebuildApps(button))   ; no refresh interval: the folders are watched and updated on changes
        PreferencesWindow._Section("Prefs.Section.AppResults")
        PreferencesWindow._LinesButton("Features.Applications.Hidden", "Prefs.AppHidden", {Summary: (lines) => PreferencesWindow.NameSummary(lines)})
        PreferencesWindow._Check("Features.Applications.StoreApps", "Prefs.StoreApps", , , "Prefs.Group.Options")
        PreferencesWindow._Check("Features.Applications.MatchPinyin", "Prefs.MatchPinyin")
    }

    static _BuildFileSearch() {
        PreferencesWindow._BeginPage("Prefs.Page.FileSearch", 130)
        PreferencesWindow._Section("Prefs.Section.FileStart")
        PreferencesWindow._Check("Features.FileSearch.SpacePrefix", "Prefs.SpacePrefix", , , "Prefs.Group.StartFileSearch")
        PreferencesWindow._Check("Features.FileSearch.QuotePrefix", "Prefs.QuotePrefix")
        PreferencesWindow._Pair(["Features.FileSearch.Keywords", "Prefs.FileKeywords", "M", "csv"], ["Features.FileSearch.FolderKeywords", "Prefs.FolderKeywords", "S", "csv"])
        PreferencesWindow._Csv("Features.FileSearch.RecentKeywords", "Prefs.RecentKeywords", "M")
        PreferencesWindow._LinesButton("Features.FileSearch.TypeFilters", "Prefs.TypeFilters", {Summary: (lines) => PreferencesWindow.TypeFilterSummary(lines), AlignPairs: true})
        PreferencesWindow._Section("Prefs.Section.FileResults")
        PreferencesWindow._Check("Features.FileSearch.InDefaultResults", "Prefs.FileInDefault", , , "Prefs.Group.NormalSearch")
        PreferencesWindow._Pair(["Features.FileSearch.MaxResults", "Prefs.FileMaxResults", "S", "number"], ["Features.FileSearch.DefaultResultsLimit", "Prefs.FileDefaultLimit", "S", "number"])
        PreferencesWindow._Field("Features.FileSearch.MinQueryLength", "Prefs.FileMinLength", "S", "number")
        PreferencesWindow._Section("Prefs.Section.Browse")                  ; browsing folders is always available: just a description
    }

    ; Search sources: Everything (searches the whole disk while it runs) and the built-in file index when Everything isn't running
    static _BuildFileIndex() {
        PreferencesWindow._BeginPage("Prefs.Page.FileIndex", 130)
        PreferencesWindow._Section("Prefs.Section.Everything")
        pinyinHelp := HelpProvider.WikiPage("File-Search") (SubStr(I18n.Lang, 1, 2) = "zh" ? "#常见问题" : "#troubleshooting")
        PreferencesWindow._LinkDesc(I18n.T("Prefs.EverythingPinyin", pinyinHelp), PreferencesWindow.ContentX)   ; the section's description: how to make Everything support pinyin (to compare with the pinyin note in the built-in index section)
        PreferencesWindow._Check("Features.FileSearch.UseEverything", "Prefs.UseEverything", , I18n.T(Everything.IsRunning() ? "Prefs.Running" : "Prefs.NotRunning"))   ; whether it is running now
        PreferencesWindow._Field("Features.FileSearch.EverythingFilter", "Prefs.EverythingFilter", "L")
        PreferencesWindow._Field("Features.FileSearch.EverythingPath", "Prefs.EverythingPath", "L", "folder")
        PreferencesWindow._Section("Prefs.Section.BuiltinIndex")
        PreferencesWindow._LinesButton("Features.FileSearch.ScopeFolders", "Prefs.ScopeFolders")
        PreferencesWindow._FieldButton("Features.FileSearch.ScopeDepth", "Prefs.ScopeDepth", "Prefs.RebuildFileIndex", (button, *) => PreferencesWindow._RebuildFiles(button))
        PreferencesWindow._Field("Features.FileSearch.ScopeExclude", "Prefs.ScopeExclude", "L")
        PreferencesWindow._Pair(["Features.FileSearch.MaxEntries", "Prefs.MaxEntries", "S", "number"], ["Features.FileSearch.RefreshMinutes", "Prefs.RefreshMinutes", "S", "number"])
    }

    static _BuildCommands() {
        PreferencesWindow._BeginPage("Prefs.Page.Commands")
        PreferencesWindow._List("CustomCommands", 540
            , [["Prefs.Col.Title", "Title", 140], ["Prefs.Col.Type", "Type", 70], ["Prefs.Col.Target", "Target", 190], ["Prefs.Col.Keyword", "Keyword", 70], ["Prefs.Col.Status", "Status", 70]]
            , CustomCommandProvider.EditorFields(), () => CustomCommandProvider.NewCommand()
            , (item, roots) => CustomCommandProvider.CheckTarget(item, roots))
        PreferencesWindow._Note("Prefs.CommandsNote")
    }

    static _BuildSnippets() {
        PreferencesWindow._BeginPage("Prefs.Page.Snippets", 150)
        PreferencesWindow._List("Snippets", 220
            , [["Prefs.Col.Name", "Name", 150], ["Prefs.Col.Keyword", "Keyword", 90], ["Prefs.Col.Text", "Text", 300]]
            , SnippetProvider.EditorFields(), () => SnippetProvider.NewSnippet(), ""
            , [{Label: I18n.T("Snippet.ImportExport"), Width: 100, Click: (list, button) => PreferencesWindow._SnippetMenu(list, button)}])
        PreferencesWindow._Section("Prefs.Section.Options")
        PreferencesWindow._Check("Features.Snippets.SearchText", "Prefs.SnippetSearchText", , , "Prefs.Group.Search")
        PreferencesWindow._Check("Features.Snippets.AutoExpand", "Prefs.SnippetAutoExpand", , , "Prefs.Group.AutoExpand")
        ; Two rows of the same short boxes: Search keyword / Paste by on the left, Keyword prefix / Delay on the right
        keywordRow := PreferencesWindow._Pair(["Features.Snippets.Keyword", "Prefs.SnippetKeyword", "K"], ["Features.Snippets.ExpandPrefix", "Prefs.ExpandPrefix", "K"])
        pasteRow := PreferencesWindow._Pair(["Features.Snippets.PasteMode", "Prefs.PasteMode", "K", "choice", ["Auto", "Clipboard", "Type"], [I18n.T("Prefs.PasteMode.Auto"), I18n.T("Prefs.PasteMode.Clipboard"), I18n.T("Prefs.PasteMode.Type")]]
            , ["Features.Snippets.PasteDelay", "Prefs.PasteDelay", "K", "number"])
        PreferencesWindow._AlignPairRows(keywordRow, pasteRow)
        PreferencesWindow._LinesButton("Features.Snippets.ExpandExclude", "Prefs.ExpandExclude")
    }

    ; The "Import / Export" button on the Snippets page: shows a menu
    static _SnippetMenu(list, button) {
        popup := Menu()
        popup.Add(I18n.T("Snippet.Import"), (*) => PreferencesWindow._ImportSnippets(list))
        popup.Add(I18n.T("Snippet.Export"), (*) => PreferencesWindow._ExportSnippets(list))
        button.GetPos(&x, &y, , &h)
        popup.Show(x, y + h)
    }

    ; Import from a table: paste what was copied from Excel, or read a CSV / TSV file; merged into the list (closing Preferences with Cancel doesn't save it)
    static _ImportSnippets(list) {
        owner := PreferencesWindow.Gui, title := I18n.T("Snippet.Import.Title"), state := {Message: ""}
        g := Gui("+Owner" owner.Hwnd " -MinimizeBox", title)
        g.SetFont("s9", ThemeManager.FontName())
        g.MarginX := 14, g.MarginY := 12
        g.AddText("xm w560", I18n.T("Snippet.Import.Help"))
        g.SetFont("s8")
        g.AddText("xm y+4 w560 cGray", I18n.T("Snippet.Import.Columns"))
        g.SetFont("s9")
        box := g.AddEdit("xm y+8 w560 r12 +Multi +WantTab")
        g.AddButton("xm y+12 w100", I18n.T("Snippet.Import.File")).OnEvent("Click", ChooseFile)
        g.AddButton("x" (14 + 560 - 170) " yp w80 Default", I18n.T("Snippet.Import.Ok")).OnEvent("Click", DoImport)
        g.AddButton("x+10 yp w80", I18n.T("Prefs.Cancel")).OnEvent("Click", (*) => g.Destroy())
        g.OnEvent("Escape", (*) => g.Destroy())
        g.OnEvent("Close", (*) => g.Destroy())
        owner.Opt("+Disabled")
        g.Show()
        WinWaitClose(g.Hwnd)
        owner.Opt("-Disabled")
        try WinActivate("ahk_id " owner.Hwnd)
        if (state.Message != "")
            MsgBox(state.Message, title, "Iconi Owner" owner.Hwnd)
        return

        ChooseFile(*) {
            g.Opt("+OwnDialogs")
            selected := FileSelect(1, , title, I18n.T("Snippet.Import.Filter"))
            if (selected = "")
                return
            try box.Value := DelimitedText.ReadFile(selected)
            catch as e
                MsgBox(e.Message, title, "Icon! Owner" g.Hwnd)
        }
        DoImport(*) {
            prefix := PreferencesWindow.GetPath(PreferencesWindow.Working, "Features.Snippets.ExpandPrefix")
            result := SnippetProvider.ImportRows(DelimitedText.Parse(box.Value), list.Items, prefix)
            if !(result.Added || result.Updated) {
                MsgBox(I18n.T("Snippet.Import.None"), title, "Icon! Owner" g.Hwnd)
                return
            }
            PreferencesWindow.MarkDirty(true)
            list.Reload.Call()
            state.Message := I18n.T("Snippet.Import.Done", result.Added, result.Updated)
                . (result.Skipped ? "`n" I18n.T("Snippet.Import.Skipped", result.Skipped) : "")
            g.Destroy()
        }
    }

    ; Export as CSV (UTF-8 with BOM, so Excel opens it without garbled text): the list in Preferences now, including unsaved changes
    static _ExportSnippets(list) {
        owner := PreferencesWindow.Gui, title := I18n.T("Snippet.Export")
        owner.Opt("+OwnDialogs")
        target := FileSelect("S16", "Snippets.csv", title, I18n.T("Snippet.Export.Filter"))
        if (target = "")
            return
        if !RegExMatch(target, "i)\.(csv|txt)$")
            target .= ".csv"
        rows := SnippetProvider.ExportRows(list.Items)
        try {
            if FileExist(target)
                FileDelete(target)
            FileAppend(DelimitedText.ToCsv(rows), target, "UTF-8")
        } catch as e {
            MsgBox(e.Message, title, "Icon! Owner" owner.Hwnd)
            return
        }
        MsgBox(I18n.T("Snippet.Export.Done", rows.Length - 1, target), title, "Iconi Owner" owner.Hwnd)
    }

    static _BuildClipboard() {
        PreferencesWindow._BeginPage("Prefs.Page.Clipboard", 120)
        PreferencesWindow._Pair(["Features.Clipboard.Hotkey", "Prefs.ClipHotkey", "K", "hotkey"], ["Features.Clipboard.Keyword", "Prefs.ClipKeyword", "S"])
        PreferencesWindow._Pair(["Features.Clipboard.MaxItems", "Prefs.ClipMaxItems", "S", "number"], ["Features.Clipboard.MaxItemLength", "Prefs.ClipMaxLength", "S", "number"])
        PreferencesWindow._Check("Features.Clipboard.Persist", "Prefs.ClipPersist", , , "Prefs.Group.History")
        PreferencesWindow._Check("Features.Clipboard.Images", "Prefs.ClipImages")
        PreferencesWindow._Field("Features.Clipboard.MaxImages", "Prefs.ClipMaxImages", "S", "number")
        PreferencesWindow._Choice("Features.Clipboard.LocalFiles", "Prefs.ClipStorage", [0, 1, 2]   ; where it is stored (ClipboardProvider.PlaceStorage)
            , [I18n.T("Prefs.ClipStorage.Data"), I18n.T("Prefs.ClipStorage.Files"), I18n.T("Prefs.ClipStorage.All")], "L")   ; long options: full width
        PreferencesWindow._Check("Features.Clipboard.MergeDoubleCopy", "Prefs.ClipMerge", , , "Prefs.Group.Copy")
        PreferencesWindow._Gap()
        PreferencesWindow._LinesButton("Features.Clipboard.IgnoreApps", "Prefs.ClipIgnoreApps")
        PreferencesWindow._Button("Prefs.ClipClear", (*) => ClipboardProvider.Clear())
    }

    static _BuildWebSearch() {
        PreferencesWindow._BeginPage("Prefs.Page.WebSearch", 130)
        PreferencesWindow._List("Features.WebSearch.Engines", 320
            , [["Prefs.Col.Keyword", "Keyword", 70], ["Prefs.Col.Title", "Title", 130], ["Prefs.Col.Url", "Url", 340]]
            , WebSearchProvider.EditorFields(), () => WebSearchProvider.NewEngine())
        PreferencesWindow._Csv("Features.WebSearch.Fallbacks", "Prefs.Fallbacks", "L")
    }

    ; Calculator: currency conversion, structural calculation (beam main bars / steel area) and its parameters
    static _BuildCalculator() {
        base := "Features.Calculator."
        PreferencesWindow._BeginPage("Prefs.Page.Calculator", 150)
        PreferencesWindow._Section("Prefs.Section.Calculator")
        PreferencesWindow._Check(base "Currency", "Prefs.Currency")
        PreferencesWindow._Section("Prefs.Section.Structural")
        PreferencesWindow._Check(base "StructuralCalc", "Prefs.StructuralCalc")
        PreferencesWindow._Field(base "RebarCover", "Prefs.RebarCover", "S", "number")      ; one row per item: labels and input boxes each aligned
        PreferencesWindow._Field(base "MaxBarSpacing", "Prefs.MaxBarSpacing", "S", "number")
        PreferencesWindow._Csv(base "BarSizes", "Prefs.BarSizes", "M", true)
        PreferencesWindow._Field(base "BarPrefix", "Prefs.BarPrefix", "S")
    }

    ; Scripts extension: description, the Scripts folder, the scripts found (double-click edits in Notepad); the switch is on the Features page
    static _BuildScripts() {
        PreferencesWindow._BeginPage("Prefs.Page.Scripts", 110)
        PreferencesWindow._Section("Prefs.Section.Scripts")
        folder := PreferencesWindow._Info("Prefs.Group.ScriptsFolder", ScriptProvider.Dir, " r1 cGray 0x8100")   ; like the data folder on the Advanced page: an ellipsis in the middle when too long
        Win.AddTooltip(folder.Hwnd, ScriptProvider.Dir)
        PreferencesWindow._y += 4
        PreferencesWindow._Button("Prefs.ScriptsFolder", (*) => (ScriptProvider.OpenFolder(), PreferencesWindow._FillScripts()))
        PreferencesWindow._Field("Features.Scripts.Keyword", "Prefs.ScriptKeyword", "S")
        PreferencesWindow._Section("Prefs.Section.ScriptList")
        listView := PreferencesWindow._Add("ListView", "w" PreferencesWindow.ContentW " h170 -Multi NoSort Grid ReadOnly"
            , [I18n.T("Prefs.Col.Name"), I18n.T("Prefs.Col.Keyword"), I18n.T("Prefs.Col.Argument"), I18n.T("Prefs.Col.Mode"), I18n.T("Prefs.Col.File")])
        for index, width in [145, 70, 95, 135, 100]
            listView.ModifyCol(index, width)
        listView.OnEvent("DoubleClick", (ctrl, row) => (row && row <= ScriptProvider.Scripts.Length) ? Run('notepad.exe "' ScriptProvider.Scripts[row].Path '"') : "")
        PreferencesWindow._scriptList := listView
        PreferencesWindow._FillScripts()
        PreferencesWindow._Below(6, listView)
        PreferencesWindow._Note("Prefs.ScriptsNote")
    }

    static _scriptList := ""
    static _FillScripts() {
        static modes := Map("window", "Script.Mode.Window", "silent", "Script.Mode.Silent", "output", "Script.Mode.Output")
        listView := PreferencesWindow._scriptList
        if !IsObject(listView)
            return
        ScriptProvider.Load()
        listView.Delete()
        for script in ScriptProvider.Scripts
            listView.Add(, script.Title, script.Keyword, script.Argument, I18n.T(modes[script.Mode]), script.Name)
        if !ScriptProvider.Scripts.Length
            listView.Add(, I18n.T("Prefs.NoScripts"))
        PreferencesWindow.FitColumns(listView)
    }

    static _BuildHotkeys() {
        PreferencesWindow._BeginPage("Prefs.Page.Hotkeys", 100)                ; narrow label column: the list starts at the left edge (Caps Lock is on the Keyboard page)
        actions := [["ToggleWindow", "Show / hide ALTRun (ToggleWindow)"]]
        for command in SystemProvider.Commands()
            actions.Push([command["Id"], command["Title"] " (" command["Id"] ")"])
        PreferencesWindow._List("Hotkeys", 500
            , [["Prefs.Col.Key", "Key", 160], ["Prefs.Col.Action", "Action", 150], ["Prefs.Col.WinTitle", "WinTitle", 230]]
            , [ItemEditor.Field("Key", "Prefs.Col.Key", "hotkey", true, "", I18n.T("Prefs.HotkeyHint"))
             , ItemEditor.Field("Action", "Prefs.Col.Action", "choice", false, actions)
             , ItemEditor.Field("WinTitle", "Prefs.Col.WinTitle", "text", false, "", I18n.T("Prefs.WinTitleHint"))]
            , () => Map("Key", "", "Action", "ToggleWindow", "WinTitle", ""))
        PreferencesWindow._Note("Prefs.HotkeysNote")
    }

    ; Quick switch for dialogs (Listary's Quick Switch)
    static _BuildQuickSwitch() {
        base := "Extensions.QuickSwitch."
        PreferencesWindow._BeginPage("Prefs.Page.QuickSwitch", 150)
        PreferencesWindow._Section("Prefs.QuickSwitch")
        PreferencesWindow._Check(base "Enabled", "Prefs.EnableExtension", , , "Prefs.Group.Status")
        PreferencesWindow._Field(base "Hotkey", "Prefs.QSHotkey", "K", "hotkey")   ; one hotkey for TC / Opus / Explorer: the one used last

        PreferencesWindow._Section("Prefs.Section.QSDialogs")
        PreferencesWindow._Check(base "StandardDialogs", "Prefs.QSStandard", , , "Prefs.Group.DialogWindows")   ; ahk_class #32770
        PreferencesWindow._Check(base "BrowseDialogs", "Prefs.QSBrowse")   ; "Browse for Folder": folder tree only (QuickSwitch.LooksLikeBrowseDialog)
        PreferencesWindow._LinesButton(base "DialogWindows", "Prefs.QSOtherDialogs")
        PreferencesWindow._LinesButton(base "ExcludeWindows", "Prefs.QSExclude")

        PreferencesWindow._Section("Prefs.Section.QSAuto")
        PreferencesWindow._Check(base "AutoSwitch", "Prefs.QSAuto", , , "Prefs.Group.Options")
        PreferencesWindow._LinesButton(base "AutoSwitchExclude", "Prefs.QSAutoExclude")
    }

    ; The dialog panel (part of QuickSwitch): the folder list and search box below open / save dialogs, and the folder menu for the keyboard
    static _BuildQuickSwitchPanel() {
        base := "Extensions.QuickSwitch."
        PreferencesWindow._BeginPage("Prefs.Page.QSPanel", 150)
        PreferencesWindow._Section("Prefs.QSPanelSection")
        PreferencesWindow._Check(base "ShowPanel", "Prefs.QSPanel", , , "Prefs.Group.Status")
        PreferencesWindow._Choice(base "PanelSearch", "Prefs.QSPanelSearch", ["folders", "all"], [I18n.T("Prefs.QSPanelSearch.Folders"), I18n.T("Prefs.QSPanelSearch.All")], "K")
        PreferencesWindow._Field(base "RecentFolders", "Prefs.QSRecent", "S", "number")
        PreferencesWindow._Section("Prefs.QSMenuSection")
        PreferencesWindow._Field(base "MenuHotkey", "Prefs.QSMenu", "K", "hotkey")
    }

    ; Date stamp (AutoDate): hotkeys and target windows for renaming files and for comment boxes
    static _BuildDateStamp() {
        base := "Extensions.AutoDate."
        PreferencesWindow._BeginPage("Prefs.Page.DateStamp", 150)
        PreferencesWindow._Section("Prefs.AutoDate")
        PreferencesWindow._Check(base "Enabled", "Prefs.EnableExtension", , , "Prefs.Group.Status")
        PreferencesWindow._Field(base "Hotkey", "Prefs.RenameHotkey", "K", "hotkey")   ; one hotkey: the window decides between the two below
        PreferencesWindow._Field(base "DateFormat", "Prefs.DateFormat", "K")   ; as wide as the hotkey box above
        PreferencesWindow._Section("Prefs.Section.DateRename")
        PreferencesWindow._LinesButton(base "RenameWindows", "Prefs.RenameWindows")
        PreferencesWindow._Section("Prefs.Section.DateAppend")
        PreferencesWindow._LinesButton(base "AppendWindows", "Prefs.AppendWindows")
    }

    ; "About" at the top (icon, name, version, home page, check for updates), the settings and data file locations and actions below
    static _BuildAdvanced() {
        PreferencesWindow._BeginPage("Prefs.Page.Advanced", 110)
        textX := PreferencesWindow._InputX(), textW := PreferencesWindow._InputW(), top := PreferencesWindow._y + 6
        PreferencesWindow._y := top                                         ; the icon in the left column; text and buttons aligned with the control column below
        if FileExist(App.IconFile)                                         ; without the icon file (e.g. only the exe was copied) use the icon in the exe
            PreferencesWindow._Add("Picture", "x" (textX - 64 - 18) " w64 h64", App.IconFile)
        else if A_IsCompiled
            PreferencesWindow._Add("Picture", "x" (textX - 64 - 18) " w64 h64 Icon1", A_ScriptFullPath)
        PreferencesWindow.Gui.SetFont("s15 bold")
        PreferencesWindow._Add("Text", "x" textX " w" textW, App.Name)
        PreferencesWindow.Gui.SetFont("s9 norm")
        ; tagline, version (what a problem report asks for, on its own line), copyright and license: the same font size, the last two gray, a little more line spacing
        PreferencesWindow._y := top + 33
        PreferencesWindow._Add("Text", "x" textX " w" textW, I18n.T("App.Tagline"))
        PreferencesWindow._y := top + 56
        PreferencesWindow._Add("Text", "x" textX " w" textW " cGray", I18n.T("Prefs.Version", App.Version) App.AdminSuffix())
        PreferencesWindow._y := top + 78
        PreferencesWindow._Add("Text", "x" textX " w" textW " cGray", "© 2013–" SubStr(App.Version, 1, 4) " zhugecaomao   ·   GPL-3.0")   ; the version starts with the release year
        PreferencesWindow._y := top + 105
        links := ""
        for link in [[App.Website, I18n.T("Prefs.Homepage")], [App.RepoUrl, "GitHub"], [App.RepoUrl "/releases", I18n.T("Update.ReleaseNotes")]
                   , [App.RepoUrl "/issues/new/choose", I18n.T("Prefs.ReportIssue")]]
            links .= (links = "" ? "" : "   ·   ") '<a href="' link[1] '">' link[2] '</a>'
        PreferencesWindow._Add("Link", "x" textX " w" textW, links)
        PreferencesWindow._y := top + 134
        PreferencesWindow._Add("Button", "x" textX " w" PreferencesWindow.ButtonW " h" PreferencesWindow.ButtonH, I18n.T("Tray.CheckUpdate"))
            .OnEvent("Click", (*) => UpdateChecker.Check())
        PreferencesWindow._y := top + 134 + PreferencesWindow.ButtonH + 22

        PreferencesWindow._Section("Prefs.Section.Data")
        ; The settings file is ALTRun.json in the data folder, so no extra line; an ellipsis in the middle when the path is too long (SS_PATHELLIPSIS), the full path on hover
        folder := PreferencesWindow._Info("Prefs.Group.DataFolder", AppSettings.DataDir, " r1 cGray 0x8100")
        Win.AddTooltip(folder.Hwnd, AppSettings.DataDir)
        PreferencesWindow._y += 4
        PreferencesWindow._Buttons(""
            , ["Prefs.EditJson", (*) => (PreferencesWindow.Cancel() || App.EditSettingsFile())]
            , ["Prefs.OpenDataFolder", (*) => PreferencesWindow._OpenFolder(AppSettings.DataDir)]
            , ["Prefs.ChangeDataFolder", (*) => PreferencesWindow._ChangeDataFolder()])

        ; What a problem report needs: the debug log (moved here from the General page, next to opening the log) and system info
        PreferencesWindow._Section("Prefs.Group.Diagnostics")
        PreferencesWindow._Check("General.SaveLog", "Prefs.SaveLog")
        PreferencesWindow._y += 4
        PreferencesWindow._Buttons(""
            , ["Prefs.OpenLog", (*) => PreferencesWindow._OpenLog()]
            , ["Prefs.CopySystemInfo", (*) => (A_Clipboard := App.SystemInfo(), App.Notify(I18n.T("Prefs.SystemInfoCopied")))])

        PreferencesWindow._Section("Prefs.Section.Reset")
        PreferencesWindow._Buttons(""
            , ["Prefs.ResetLearning", (*) => PreferencesWindow._ResetLearning()]
            , ["Prefs.RestoreDefaults", (*) => PreferencesWindow._RestoreDefaults()])
    }

    ; while there is no log yet ("Write debug log" off, or just switched on and not applied), say how to get one
    static _OpenLog() {
        Logger.Flush()
        if !FileExist(Logger.File)
            return MsgBox(I18n.T("Prefs.NoLog"), App.Name, "Iconi Owner" PreferencesWindow.Gui.Hwnd)
        App.OpenLog()
    }

    ; Reset to defaults: after confirming, back up ALTRun.json, keep what the user added (see AppSettings.KeptOnReset), replace the rest with the defaults and reload.
    ; changes not yet applied in Preferences are dropped
    static _RestoreDefaults() {
        if (MsgBox(I18n.T("Prefs.RestoreDefaultsConfirm"), App.Name, "YesNo Icon! Default2 Owner" PreferencesWindow.Gui.Hwnd) != "Yes")
            return
        backup := AppSettings.File ".bak"
        try FileCopy(AppSettings.File, backup, true)
        catch as e
            return MsgBox(I18n.T("Settings.SaveError", e.Message), App.Name, "Icon! Owner" PreferencesWindow.Gui.Hwnd)
        AppSettings.Data := AppSettings.ResetData(AppSettings.Data)
        if !AppSettings.Save()
            return
        Logger.Debug("Preferences: settings restored to defaults, old file kept as " backup)
        App.Reload()
    }

    ; Usage statistics (like Alfred's Usage): totals, a bar chart per day for the last 30 days, the count per feature
    static _usage := ""
    static _BuildUsage() {
        PreferencesWindow._BeginPage("Prefs.Page.Usage")
        PreferencesWindow._Section("Usage.Title")
        w := PreferencesWindow.ContentW, x := PreferencesWindow.ContentX
        totals := PreferencesWindow._Add("Text", "w" w)
        PreferencesWindow._Below(4, totals)
        since := PreferencesWindow._Add("Text", "w" w " cGray")
        PreferencesWindow._Below(12, since)
        top := PreferencesWindow._y, barW := 14, step := 18, chartH := 100
        bars := []
        Loop 30 {
            PreferencesWindow._y := top
            bars.Push(PreferencesWindow._Add("Progress", "x" (x + (A_Index - 1) * step) " w" barW " h" chartH " Vertical c3B82F6 BackgroundE6EBF2"))
        }
        PreferencesWindow._y := top + chartH + 4
        PreferencesWindow.Gui.SetFont("s8")
        first := PreferencesWindow._Add("Text", "x" x " w100 cGray")
        chart := PreferencesWindow._Add("Text", "x" (x + 100) " w" (30 * step - barW - 200 + 10) " Center cGray")
        PreferencesWindow._Add("Text", "x" (x + 30 * step - barW - 90) " w" (90 + barW - 4) " Right cGray", I18n.T("Usage.Today"))
        PreferencesWindow.Gui.SetFont("s9")
        PreferencesWindow._y += 22
        headers := []
        for key in ["Feature", "Today", "Week", "Month", "All", "Share"]
            headers.Push(I18n.T("Usage.Col." key))
        list := PreferencesWindow._Add("ListView", "w" w " h272 -Multi NoSort Grid ReadOnly", headers)
        for index, width in [178, 66, 66, 66, 80, 80]
            list.ModifyCol(index, width (index > 1 ? " Right" : ""))
        PreferencesWindow._Below(8, list)
        PreferencesWindow._Button("Usage.Clear", (*) => PreferencesWindow._ClearUsage())
        PreferencesWindow._usage := {Totals: totals, Since: since, Bars: bars, First: first, Chart: chart, List: list}
        PreferencesWindow._RefreshUsage()
    }

    static _RefreshUsage() {
        ui := PreferencesWindow._usage
        num := (n) => Calc.Thousands(n)
        periods := [Usage.Summary(1), Usage.Summary(7), Usage.Summary(30), Usage.Summary(0)]
        ui.Totals.Value := I18n.T("Usage.Totals", num(Usage.Total(periods[1])), num(Usage.Total(periods[2])), num(Usage.Total(periods[3])), num(Usage.Total(periods[4])))
        all := periods[4]
        ui.Since.Value := I18n.T("Usage.Since", (Usage.Since != "") ? Usage.Since : FormatTime(, "yyyy-MM-dd"), num(all.Has("Show") ? all["Show"] : 0))
        daily := Usage.Daily(30)
        most := 0
        for day in daily
            most := Max(most, day[2])
        for index, bar in ui.Bars {
            bar.Opt("Range0-" Max(most, 1))
            bar.Value := daily[index][2]
        }
        ui.First.Value := SubStr(daily[1][1], 6)
        ui.Chart.Value := I18n.T("Usage.Chart", num(most))

        ; most used first (ties in the order of Usage.Features), "show search window" last
        rows := []
        for order, feature in Usage.Features
            rows.Push({Id: feature, Order: order, All: all.Has(feature) ? all[feature] : 0})
        Loop rows.Length - 1 {
            Loop rows.Length - A_Index {
                a := rows[A_Index], b := rows[A_Index + 1]
                if (b.All > a.All || (b.All = a.All && b.Order < a.Order))
                    rows[A_Index] := b, rows[A_Index + 1] := a
            }
        }
        rows.Push({Id: "Show", Order: 0, All: all.Has("Show") ? all["Show"] : 0})
        total := Usage.Total(all)
        list := ui.List
        list.Delete()
        for row in rows {
            counts := []
            for period in periods
                counts.Push(num(period.Has(row.Id) ? period[row.Id] : 0))
            share := (row.Id = "Show") ? "-" : (total ? Format("{:.1f}%", row.All * 100 / total) : "0%")
            list.Add("", I18n.T("Usage.F." row.Id), counts[1], counts[2], counts[3], counts[4], share)
        }
        PreferencesWindow.FitColumns(list)
    }

    static _ClearUsage() {
        if (MsgBox(I18n.T("Usage.ClearConfirm"), App.Name, "YesNo Icon? Default2 Owner" PreferencesWindow.Gui.Hwnd) != "Yes")
            return
        Usage.Clear()
        PreferencesWindow._RefreshUsage()
    }

    static _ResetLearning() {
        Knowledge.Picks := Map(), Knowledge.QueryPicks := Map(), Knowledge.History := [], Knowledge.Recent := []
        Knowledge.Save()
        MsgBox(I18n.T("Prefs.ResetDone"), App.Name, 64)
    }

    ; Switch to another data folder (e.g. one in OneDrive shared by several PCs): copy the current settings and data there (if ALTRun settings are already there,
    ; they can be used directly); the location goes into ALTRun.json at the default location ("DataLocation") and ALTRun reloads. The old folder is left alone, like a backup
    static _ChangeDataFolder() {
        current := AppSettings.DataDir
        folder := DirSelect("*" current, 3, I18n.T("Prefs.DataFolderPrompt"))
        if (folder = "")
            return
        folder := RTrim(folder, "\")
        owner := " Owner" PreferencesWindow.Gui.Hwnd
        if (folder = current)
            return
        if (InStr(folder "\", current "\") = 1 || InStr(current "\", folder "\") = 1)
            return MsgBox(I18n.T("Prefs.DataFolderNested"), App.Name, "Icon!" owner)
        useExisting := false
        if FileExist(folder "\ALTRun.json") {
            answer := MsgBox(I18n.T("Prefs.DataFolderExisting", folder), App.Name, "YesNoCancel Icon?" owner)
            if (answer = "Cancel")
                return
            useExisting := (answer = "Yes")
        } else if (MsgBox(I18n.T("Prefs.DataFolderCopy", folder), App.Name, "OKCancel Icon?" owner) != "OK") {
            return
        }
        if PreferencesWindow.Cancel()                                       ; with unsaved changes, ask whether to discard them first
            return
        try {
            if !useExisting {
                AppSettings.Save(), Knowledge.Save(), Usage.Save(), ClipboardProvider.Save()
                DirCopy(current, folder, true)
            }
            AppSettings.SetDataLocation(folder)
        } catch as e {
            return MsgBox(I18n.T("Prefs.DataFolderFailed", folder, e.Message), App.Name, "Icon!")
        }
        App.Reload()
    }

    static _OpenFolder(folder) {
        DirCreate(folder)
        ActionCatalog.OpenFolder(folder)                                    ; the file manager in the settings (e.g. Total Commander)
    }

    ;---------------------------------------------------------------------------
    ; Builders
    ;---------------------------------------------------------------------------
    ; labelW: width of this page's left-column labels (set by its longest label, so the content isn't pushed right as a whole)
    static _BeginPage(nameKey, labelW := 190) {
        page := PreferencesWindow.Pages[PreferencesWindow._PageIndex(nameKey)]
        page.Controls := [], page.Built := true
        PreferencesWindow._current := page
        PreferencesWindow._y := 14
        PreferencesWindow._pageKey := nameKey
        PreferencesWindow._pairX := 0                                       ; right-column position of this page's paired settings, see _Pair
        PreferencesWindow.LabelW := Max(labelW, PreferencesWindow._labelWidths.Get(I18n.Lang " " nameKey, 0))
    }

    ; when a left-column label doesn't fit (it would wrap), note the width this page needs, see Show
    static _FitLabel(ctrl, text) {
        need := Ceil(Win.TextExtent(ctrl.Hwnd, text) * 96 / A_ScreenDPI) + 12 + 2
        if (need > PreferencesWindow.LabelW && PreferencesWindow.LabelW < PreferencesWindow.LabelMax) {
            key := PreferencesWindow._pageKey
            PreferencesWindow._wider[key] := Min(PreferencesWindow.LabelMax, Max(need, PreferencesWindow._wider.Get(key, 0)))
        }
    }

    static _Add(type, options, text := "") {
        pos := InStr(options, " x") || InStr(options, "x") = 1 ? "" : "x" PreferencesWindow.ContentX " "
        ctrl := PreferencesWindow.Gui.Add(type, pos "y" PreferencesWindow._y " " options " Hidden", text)
        PreferencesWindow._current.Controls.Push(ctrl)
        switch type, false {                                                ; a user change -> "Apply" becomes available (type names are case-insensitive)
            case "Edit", "DropDownList", "ComboBox": ctrl.OnEvent("Change", (*) => PreferencesWindow.MarkDirty())
            case "CheckBox", "Radio":                ctrl.OnEvent("Click", (*) => PreferencesWindow.MarkDirty())
        }
        return ctrl
    }

    static _Bind(path, readFn) {
        PreferencesWindow.Binds.Push({Path: path, Read: readFn})
    }

    ; Moves _y below the lowest of these controls (their real height, so wrapped text never overlaps)
    static _Below(gap, controls*) {
        bottom := PreferencesWindow._y
        for ctrl in controls {
            ctrl.GetPos(, &top, , &height)
            bottom := Max(bottom, top + height)
        }
        PreferencesWindow._y := bottom + gap
    }

    ; Section title: bold + a thin line below; a little space above when it isn't the page's first row
    static _Section(labelKey) {
        if (PreferencesWindow._y > 20)
            PreferencesWindow._y += 8
        PreferencesWindow.Gui.SetFont("bold")
        ctrl := PreferencesWindow._Add("Text", "w" PreferencesWindow.ContentW, I18n.T(labelKey))
        PreferencesWindow.Gui.SetFont("norm")
        PreferencesWindow._Below(3, ctrl)
        line := PreferencesWindow._Add("Text", "w" PreferencesWindow.ContentW " h1 0x10")   ; SS_ETCHEDHORZ
        PreferencesWindow._Below(PreferencesWindow._HasDesc(labelKey) ? 4 : 10, line)
        PreferencesWindow._Desc(labelKey, PreferencesWindow.ContentX)
    }

    static _HasDesc(labelKey) {
        return I18n.Strings.Has(labelKey ".Desc")
    }

    ; Settings whose description has a Wiki link: {1} in the description becomes the URL of [page, anchor on the Chinese page, anchor on the English page]
    static DescLinks := Map("Prefs.SelectionHotkey", ["Usage", "#选中内容的操作", "#actions-on-selected-content"])

    ; The small gray description below a setting ("<labelKey>.Desc" in I18n); nothing if there is none. x: aligned with the left of the control above
    static _Desc(labelKey, x, text := "") {
        controlType := "Text"
        if (text = "") {
            if !PreferencesWindow._HasDesc(labelKey)
                return
            text := I18n.T(labelKey ".Desc", PreferencesWindow._DescLink(labelKey))
            if PreferencesWindow.DescLinks.Has(labelKey)
                controlType := "Link"                                              ; <a href="...">...</a>, opens the web page when clicked
        }
        PreferencesWindow.Gui.SetFont("s8")
        ctrl := PreferencesWindow._Add(controlType, "x" x " w" (PreferencesWindow.ContentX + PreferencesWindow.ContentW - x) " cGray", text)
        PreferencesWindow.Gui.SetFont("s9")
        PreferencesWindow._Below(10, ctrl)
    }

    static _DescLink(labelKey) {
        if !PreferencesWindow.DescLinks.Has(labelKey)
            return ""
        link := PreferencesWindow.DescLinks[labelKey]
        return HelpProvider.WikiPage(link[1]) (SubStr(I18n.Lang, 1, 2) = "zh" ? link[2] : link[3])
    }

    ; A gray description with links (<a href="...">...</a>, opens the page when clicked), below a section title that has no description of its own
    static _LinkDesc(text, x) {
        PreferencesWindow._y -= 6                                           ; like a section's own description, 4 pixels below the title's line
        PreferencesWindow.Gui.SetFont("s8")
        ctrl := PreferencesWindow._Add("Link", "x" x " w" (PreferencesWindow.ContentX + PreferencesWindow.ContentW - x) " cGray", text)
        PreferencesWindow.Gui.SetFont("s9")
        PreferencesWindow._Below(10, ctrl)
    }

    ; A full-width gray description (below lists)
    static _Note(textKey) {
        PreferencesWindow._Desc("", PreferencesWindow.ContentX, I18n.T(textKey))
    }

    static _Gap() {
        PreferencesWindow._y += 10
    }

    ; A label and an input box on the same row: the label moves down 3 pixels to align with the box's text
    static _Label(labelKey) {
        PreferencesWindow._y += 3
        ctrl := PreferencesWindow._Add("Text", "w" (PreferencesWindow.LabelW - 12) " Right", I18n.T(labelKey))
        PreferencesWindow._y -= 3
        PreferencesWindow._FitLabel(ctrl, I18n.T(labelKey))
        return ctrl
    }

    static _InputX() {
        return PreferencesWindow.ContentX + PreferencesWindow.LabelW
    }

    ; width of the control column (the width of a long "L" input box)
    static _InputW() {
        return PreferencesWindow.ContentW - PreferencesWindow.LabelW
    }

    ; Input box width: "S" numbers / "M" hotkeys, keywords / "L" paths, filters (the whole control column; leaves room for a browse button if there is one)
    static _Width(size, withBrowse := false) {
        switch size {
            case "S": return PreferencesWindow.WidthS
            case "M": return PreferencesWindow.WidthM
            case "K": return PreferencesWindow.WidthK
        }
        return PreferencesWindow._InputW() - (withBrowse ? 34 : 0)
    }

    ; Check box in the control column; group: the group label in the left column (only on the first check box of a group).
    ; desc: a description shown instead of "<labelKey>.Desc" (e.g. text that changes with a state); width: to the right edge of the page by default
    static _Check(path, labelKey, x := 0, desc := "", group := "", width := 0) {
        x := x ? x : PreferencesWindow._InputX()
        if (group != "")
            PreferencesWindow._GroupLabel(group)
        ctrl := PreferencesWindow._Add("Checkbox", "x" x " w" (width ? width : PreferencesWindow.ContentX + PreferencesWindow.ContentW - x), I18n.T(labelKey))
        ctrl.Value := PreferencesWindow.GetPath(PreferencesWindow.Working, path) ? 1 : 0
        PreferencesWindow._Bind(path, () => ctrl.Value)
        hasDesc := (desc != "" || PreferencesWindow._HasDesc(labelKey))
        PreferencesWindow._Below(hasDesc ? 0 : 6, ctrl)
        PreferencesWindow._Desc(labelKey, x + 18, desc)                     ; aligned with the check box's text
        return ctrl
    }

    ; A group label in the left column, on the same row as the first control on the right
    static _GroupLabel(labelKey, dy := 1) {
        PreferencesWindow._y += dy
        ctrl := PreferencesWindow._Add("Text", "x" PreferencesWindow.ContentX " w" (PreferencesWindow.LabelW - 12) " Right", I18n.T(labelKey))
        PreferencesWindow._y -= dy
        PreferencesWindow._FitLabel(ctrl, I18n.T(labelKey))
        return ctrl
    }

    ; Left-column label + some text (or a link) on the right, e.g. "Version: 2026.09.26"
    static _Info(labelKey, text, options := "", type := "Text") {
        label := PreferencesWindow._GroupLabel(labelKey)
        ctrl := PreferencesWindow._Add(type, "x" PreferencesWindow._InputX() " w" PreferencesWindow._InputW() options, text)
        PreferencesWindow._Below(6, label, ctrl)
        return ctrl
    }

    ; A label directly followed by an input box (for pages arranged in sections, without the left column): "Search history entries [30]"
    static _InlineField(path, labelKey, size, kind := "text") {
        x := PreferencesWindow._InputX()
        PreferencesWindow._y += 3
        label := PreferencesWindow._Add("Text", "x" x, I18n.T(labelKey))
        PreferencesWindow._y -= 3
        label.GetPos(, , &labelW)
        value := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        ctrl := PreferencesWindow._Add("Edit", "x" (x + labelW + 8) " w" PreferencesWindow._Width(size) " r1 -Multi" (kind = "number" ? " Number" : ""), value)
        if (kind = "number")
            PreferencesWindow._Bind(path, () => IsInteger(ctrl.Value) ? Integer(ctrl.Value) : 0)
        else
            PreferencesWindow._Bind(path, () => ctrl.Value)
        PreferencesWindow._Below(PreferencesWindow._HasDesc(labelKey) ? 3 : 8, label, ctrl)
        PreferencesWindow._Desc(labelKey, x)
        return ctrl
    }

    ; Aligns the left edges of several controls (the input boxes of _InlineField rows) to the rightmost one
    static _AlignLeft(controls*) {
        left := 0
        for ctrl in controls {
            ctrl.GetPos(&x)
            left := Max(left, x)
        }
        for ctrl in controls {
            ctrl.GetPos(&x)
            if (x < left)
                ctrl.Move(left)
        }
    }

    ; kind: text / number / file / folder; size: "S" / "M" / "L" (see _Width)
    static _Field(path, labelKey, size, kind := "text") {
        label := PreferencesWindow._Label(labelKey)
        if (kind = "hotkey")
            return PreferencesWindow._BelowInput(labelKey, label, PreferencesWindow._HotkeyBox(PreferencesWindow._InputX(), path))
        value := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        browseKind := (kind = "file" || kind = "folder")
        ctrl := PreferencesWindow._Add("Edit", "x" PreferencesWindow._InputX() " w" PreferencesWindow._Width(size, browseKind) " r1 -Multi" (kind = "number" ? " Number" : ""), value)
        if browseKind {
            browse := PreferencesWindow._Add("Button", "x+4 yp-1 w30", I18n.T("Prefs.Browse"))
            browseFn := ItemEditor._Browser(ctrl, kind, PreferencesWindow.Gui)
            browse.OnEvent("Click", (p*) => (browseFn(p*), PreferencesWindow.MarkDirty()))
        }
        if (kind = "number")
            PreferencesWindow._Bind(path, () => IsInteger(ctrl.Value) ? Integer(ctrl.Value) : 0)
        else
            PreferencesWindow._Bind(path, () => ctrl.Value)
        PreferencesWindow._BelowInput(labelKey, label, ctrl)
    }

    ; After a label + input box row: a description goes right below the box
    static _BelowInput(labelKey, controls*) {
        hasDesc := PreferencesWindow._HasDesc(labelKey)
        PreferencesWindow._Below(hasDesc ? 3 : 8, controls*)
        PreferencesWindow._Desc(labelKey, PreferencesWindow._InputX())
    }

    static _Choice(path, labelKey, values, labels, size := "M") {
        label := PreferencesWindow._Label(labelKey)
        ctrl := PreferencesWindow._DropDown(PreferencesWindow._InputX(), path, values, labels, size)
        PreferencesWindow._BelowInput(labelKey, label, ctrl)
        return ctrl
    }

    ; A drop-down bound to path. The option matching the default value is marked "(default)",
    ; and the box widens (up to the right edge of the page) when an option would be cut off
    static _DropDown(x, path, values, labels, size) {
        static defaults := AppSettings.Defaults()
        default := PreferencesWindow.GetPath(defaults, path)
        marked := []
        for index, label in labels
            marked.Push(values[index] = default ? I18n.T("Prefs.DefaultChoice", label) : label)
        ctrl := PreferencesWindow._Add("DropDownList", "x" x " w" PreferencesWindow._Width(size), marked)
        width := Min(PreferencesWindow._ChoiceWidth(ctrl), PreferencesWindow.ContentX + PreferencesWindow.ContentW - x)
        if (width > PreferencesWindow._Width(size))
            ctrl.Move(, , width)
        current := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        ctrl.Value := 1
        for index, value in values
            if (value = current)
                ctrl.Value := index
        PreferencesWindow._Bind(path, () => values[ctrl.Value])
        return ctrl
    }

    ; Width a drop-down needs to show its longest option in full
    static _ChoiceWidth(ctrl) {
        width := 0
        for text in ControlGetItems(ctrl)
            width := Max(width, Ceil(Win.TextExtent(ctrl.Hwnd, text) * 96 / A_ScreenDPI) + 32)   ; plus the arrow button and border
        return width
    }

    ; A list with one item per line, shown as a read-only summary + "Edit...", which opens it in a roomy window of its own
    ; (ItemEditor.EditLines, with the longer "<labelKey>.Help" text): the page stays short and long lines are easy to read.
    ; The setting at path is an array. options (all optional):
    ;   Summary:    (lines) => the text shown on the page (default: the items joined with commas; the box cuts off what doesn't fit)
    ;   AlignPairs: the button goes under the right column of the paired settings above (_Pair), not at the right edge
    static _LinesButton(path, labelKey, options := "") {
        label := PreferencesWindow._Label(labelKey)
        value := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        state := {Lines: (value is Array) ? value.Clone() : []}
        summarize := (IsObject(options) && options.HasOwnProp("Summary")) ? options.Summary : (lines) => TextTools.JoinList(lines)
        alignPairs := IsObject(options) && options.HasOwnProp("AlignPairs") && options.AlignPairs
        x := PreferencesWindow._InputX(), paired := alignPairs && PreferencesWindow._pairX
        buttonW := paired ? PreferencesWindow._pairSecondW : PreferencesWindow.WidthS
        buttonX := (paired ? PreferencesWindow._pairX + PreferencesWindow._pairSecondW : x + PreferencesWindow._InputW()) - buttonW
        box := PreferencesWindow._Add("Edit", "x" x " w" (buttonX - 8 - x) " r1 -Multi ReadOnly", summarize(state.Lines))
        box.GetPos(, , , &boxH)
        button := PreferencesWindow._Add("Button", "x" buttonX " yp-1 w" buttonW " h" (boxH + 2), I18n.T("Prefs.EditList"))
        button.OnEvent("Click", (*) => PreferencesWindow._EditLines(labelKey, state, box, summarize))
        PreferencesWindow._Bind(path, () => state.Lines)
        PreferencesWindow._BelowInput(labelKey, label, box, button)
    }

    static _EditLines(labelKey, state, box, summarize) {
        help := I18n.Strings.Has(labelKey ".Help") ? I18n.T(labelKey ".Help") : I18n.T(labelKey ".Desc")
        edited := ItemEditor.EditLines(PreferencesWindow.Gui, I18n.T(labelKey), help, state.Lines)
        if (!(edited is Array) || TextTools.Join(edited) == TextTools.Join(state.Lines))
            return
        state.Lines := edited, box.Value := summarize(edited)
        PreferencesWindow.MarkDirty()
    }

    ; File types: just the keywords ("doc, pic, video, ...")
    static TypeFilterSummary(lines) {
        keywords := []
        for line in lines
            if IsObject(filter := FileSearchProvider.ParseTypeFilter(line))
                keywords.Push(filter.Keyword)
        return TextTools.JoinList(keywords)
    }

    ; Hidden apps: the app names (Disk Cleanup, NVIDIACorp.NVIDIAControlPanel), not the long paths
    static NameSummary(lines) {
        names := []
        for line in lines {
            SplitPath(line, &name)
            if InStr(name, "!")                                             ; shell:AppsFolder\<family>!<app>
                name := SubStr(name, InStr(name, "!", , -1) + 1)
            names.Push(RegExReplace(name, "i)\.lnk$"))
        }
        return TextTools.JoinList(names)
    }

    ; numeric: a list of numbers (e.g. bar diameters), stored as numbers rather than text
    static _Csv(path, labelKey, size, numeric := false) {
        label := PreferencesWindow._Label(labelKey)
        value := TextTools.JoinList(PreferencesWindow.GetPath(PreferencesWindow.Working, path))
        ctrl := PreferencesWindow._Add("Edit", "x" PreferencesWindow._InputX() " w" PreferencesWindow._Width(size) " r1 -Multi", value)
        PreferencesWindow._Bind(path, () => TextTools.SplitList(ctrl.Value, numeric))
        PreferencesWindow._BelowInput(labelKey, label, ctrl)
    }

    ; Button in the control column, one size for all; group: the group label in the left column
    static _Button(labelKey, fn, group := "") {
        if (group != "")
            PreferencesWindow._GroupLabel(group, 5)                         ; aligned with the button's text
        ctrl := PreferencesWindow._Add("Button", "x" PreferencesWindow._InputX() " w" PreferencesWindow.ButtonW " h" PreferencesWindow.ButtonH, I18n.T(labelKey))
        ctrl.OnEvent("Click", fn)
        PreferencesWindow._Below(PreferencesWindow._HasDesc(labelKey) ? 2 : 6, ctrl)
        PreferencesWindow._Desc(labelKey, PreferencesWindow._InputX())
        return ctrl
    }

    ; Several buttons side by side: specs = [labelKey, fn], ...; the description is the first button's "<labelKey>.Desc"
    static _Buttons(group, specs*) {
        if (group != "")
            PreferencesWindow._GroupLabel(group, 5)
        x := PreferencesWindow._InputX(), controls := []
        buttonW := Min(PreferencesWindow.ButtonW, (PreferencesWindow._InputW() - 10 * (specs.Length - 1)) // specs.Length)   ; narrower buttons when they don't fit after the left column was widened
        for spec in specs {
            ctrl := PreferencesWindow._Add("Button", "x" x " w" buttonW " h" PreferencesWindow.ButtonH, I18n.T(spec[1]))
            ctrl.OnEvent("Click", spec[2])
            controls.Push(ctrl)
            x += buttonW + 10
        }
        PreferencesWindow._Below(PreferencesWindow._HasDesc(specs[1][1]) ? 2 : 6, controls*)
        PreferencesWindow._Desc(specs[1][1], PreferencesWindow._InputX())
    }

    ; Two settings on one row: the left one's label in the left column, the right one's label right after the left input box.
    ; left / right = [path, labelKey, size, kind := "text", values, labels] (kind: text / number / csv / choice)
    ; the description is the left one's "<labelKey>.Desc"
    static _Pair(left, right) {
        label := PreferencesWindow._Label(left[2])
        first := PreferencesWindow._Input(PreferencesWindow._InputX(), left*)
        first.GetPos(&x, , &w)
        sameAsFirstRow := PreferencesWindow._pairX && left[3] = PreferencesWindow._pairSizes[1] && right[3] = PreferencesWindow._pairSizes[2]
        if sameAsFirstRow {                                                 ; the same kind of input box as the first row of the page: the same width, right side aligned to the same column
            w := PreferencesWindow._pairFirstW
            first.Move(, , w)
        }
        PreferencesWindow._y += 3
        label2 := PreferencesWindow._Add("Text", "x" (x + w + 18), I18n.T(right[2]))
        PreferencesWindow._y -= 3
        label2.GetPos(&x2, , &w2)
        ; With a widened left column the control column is narrower: if it doesn't fit, shorten the left input box (down to "S"); with the same kind of box on both sides shorten both, keeping them equal
        secondW := sameAsFirstRow ? PreferencesWindow._pairSecondW : 0
        overflow := x2 + w2 + 8 + (secondW ? secondW : PreferencesWindow._Width(right[3])) - (PreferencesWindow.ContentX + PreferencesWindow.ContentW)
        if (!sameAsFirstRow && overflow > 0 && left[3] = right[3] && w > PreferencesWindow.WidthS) {
            w := Max(PreferencesWindow.WidthS, (PreferencesWindow.ContentX + PreferencesWindow.ContentW - x - 18 - w2 - 8) // 2)
            first.Move(, , w)
            label2.Move(x + w + 18)
            label2.GetPos(&x2)
            secondW := w, overflow := x2 + w2 + 8 + w - (PreferencesWindow.ContentX + PreferencesWindow.ContentW)
        }
        if (overflow > 0 && w > PreferencesWindow.WidthS) {
            shrink := Min(overflow, w - PreferencesWindow.WidthS)
            w -= shrink, overflow -= shrink
            first.Move(, , w)
            label2.Move(x + w + 18)
            label2.GetPos(&x2)
        }
        if (overflow > 0) {                                                 ; still doesn't fit: the right one moves to the next row
            first.GetPos(, &firstY, , &firstH)
            PreferencesWindow._y := firstY + firstH + 6
            label2.Move(PreferencesWindow._InputX(), PreferencesWindow._y + 3)
            label2.GetPos(&x2)
        }
        secondX := x2 + w2 + 8
        ; Paired settings on several rows of a page, with the left boxes as wide or narrower: the right boxes align with the first row's right box (labels right-aligned against the boxes)
        if (overflow <= 0) {
            target := PreferencesWindow._pairX
            if !target
                PreferencesWindow._pairX := secondX, PreferencesWindow._pairFirstW := w, PreferencesWindow._pairSizes := [left[3], right[3]]
                , PreferencesWindow._pairSecondW := secondW ? secondW : PreferencesWindow._Width(right[3])
            else if (w <= PreferencesWindow._pairFirstW && target >= secondX && target + (secondW ? secondW : PreferencesWindow._Width(right[3])) <= PreferencesWindow.ContentX + PreferencesWindow.ContentW) {
                label2.Move(target - 8 - w2)
                secondX := target
            }
        }
        second := PreferencesWindow._Input(secondX, right*)
        if secondW
            second.Move(, , secondW)
        PreferencesWindow._BelowInput(left[2], label, first, label2, second)
        return {Label2: label2, Second: second}
    }

    ; Rows from _Pair whose right boxes should start at the same x even when a later row's right label is wider:
    ; move the right label and box of the other rows to the rightmost one
    static _AlignPairRows(rows*) {
        target := 0
        for row in rows
            row.Second.GetPos(&x), target := Max(target, x)
        for row in rows {
            row.Second.GetPos(&x), row.Label2.GetPos(&labelX)
            row.Second.Move(target), row.Label2.Move(labelX + target - x)
        }
    }

    ; A single input control (without a label) bound to path; kind as in _Pair
    static _Input(x, path, labelKey, size, kind := "text", values := "", labels := "") {
        current := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        if (kind = "hotkey")
            return PreferencesWindow._HotkeyBox(x, path)
        if (kind = "choice")
            return PreferencesWindow._DropDown(x, path, values, labels, size)
        text := (kind = "csv") ? TextTools.JoinList(current) : current
        ctrl := PreferencesWindow._Add("Edit", "x" x " w" PreferencesWindow._Width(size) " r1 -Multi" (kind = "number" ? " Number" : ""), text)
        switch kind {
            case "number": PreferencesWindow._Bind(path, () => IsInteger(ctrl.Value) ? Integer(ctrl.Value) : 0)
            case "csv":    PreferencesWindow._Bind(path, () => TextTools.SplitList(ctrl.Value))
            default:       PreferencesWindow._Bind(path, () => ctrl.Value)
        }
        return ctrl
    }

    ; The hotkey recording box: not through _Add (the program changing its text isn't a change); marked as changed only when a new hotkey is recorded
    static _HotkeyBox(x, path) {
        ctrl := HotkeyBox.Add(PreferencesWindow.Gui, "x" x " y" PreferencesWindow._y " w" PreferencesWindow.WidthK " Hidden"
            , PreferencesWindow.GetPath(PreferencesWindow.Working, path), false, () => PreferencesWindow.MarkDirty())
        PreferencesWindow._current.Controls.Push(ctrl)
        PreferencesWindow._Bind(path, () => HotkeyBox.Value(ctrl))
        return ctrl
    }

    ; A button on the same row as the short input box ("S") above (e.g. "Subfolder depth [4]  [Rebuild file index]")
    ; "label [input box]" right of a check box on the same row: the label starts at x, the box follows right after, centered vertically with the check box
    static _SideField(rowY, x, path, labelKey, size := "S") {
        label := PreferencesWindow._Add("Text", "x" x " y" (rowY + 1), I18n.T(labelKey))
        label.GetPos(&labelX, , &labelW)
        ctrl := PreferencesWindow._Input(labelX + labelW + 8, path, labelKey, size)
        ctrl.Move(, rowY - 3)
        return ctrl
    }

    ; "Rebuild index": rebuild now (in one go), show the result, OK closes it
    static _RebuildApps(button) {
        button.Enabled := false
        count := ApplicationProvider.Rebuild()
        button.Enabled := true
        PreferencesWindow._RebuildResult(I18n.T("Prefs.RebuildIndex.Done", count), I18n.T("Prefs.RebuildIndex"))
    }

    ; "Rebuild file index": scans in the background in batches, the button shows "Rebuilding...", the result is shown when done
    static _RebuildFiles(button) {
        label := button.Text
        button.Text := I18n.T("Prefs.Rebuilding"), button.Enabled := false
        FileIndex.Rebuild(Done)

        Done(count) {
            try button.Text := label, button.Enabled := true                ; Preferences may have closed already
            options := AppSettings.Feature("FileSearch")
            message := I18n.T("Prefs.RebuildFileIndex.Done", count)
            if (count >= options["MaxEntries"])
                message .= "`n`n" I18n.T("Prefs.RebuildFileIndex.Limit", options["MaxEntries"])
            if (options["UseEverything"] && Everything.IsRunning())
                message .= "`n`n" I18n.T("Prefs.RebuildFileIndex.Everything")
            PreferencesWindow._RebuildResult(message, label)
        }
    }

    ; The rebuild result: a dialog while Preferences is open (OK closes it); a notification if it has closed
    static _RebuildResult(message, title) {
        if !IsObject(PreferencesWindow.Gui)
            return App.Toast(message)
        if PreferencesWindow._dirty                                         ; the rebuild uses the saved settings
            message .= "`n`n" I18n.T("Prefs.RebuildUnsaved")
        MsgBox(message, title, "Iconi Owner" PreferencesWindow.Gui.Hwnd)
    }

    ; A number box (width "S") followed by a button (e.g. "Subfolder depth" + "Rebuild index"), with the gray description (of labelKey) below both
    static _FieldButton(path, labelKey, buttonKey, fn) {
        label := PreferencesWindow._Label(labelKey)
        rowY := PreferencesWindow._y
        ctrl := PreferencesWindow._Add("Edit", "x" PreferencesWindow._InputX() " w" PreferencesWindow.WidthS " r1 -Multi Number", PreferencesWindow.GetPath(PreferencesWindow.Working, path))
        PreferencesWindow._Bind(path, () => IsInteger(ctrl.Value) ? Integer(ctrl.Value) : 0)
        button := PreferencesWindow._Add("Button", "x" (PreferencesWindow._InputX() + PreferencesWindow.WidthS + 12) " y" (rowY - 2) " w" PreferencesWindow.ButtonW " h" PreferencesWindow.ButtonH, I18n.T(buttonKey))
        button.OnEvent("Click", fn)
        PreferencesWindow._BelowInput(labelKey, label, ctrl, button)
        return ctrl
    }

    ; List page: a ListView shows the array at path; Add / Edit / Delete use ItemEditor
    ; check: optional, (item, roots) => "OK" / "Missing" / "Unavailable" / "Skipped".
    ; When given there is a "Check paths" button, with the results in a "Status" column (the items themselves have no such key)
    ; The filter box right of the buttons: only items containing all the typed words (searched in every column's text, case-insensitive)
    static _List(path, height, columns, fields, newItem, check := "", buttons := "") {
        items := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        headers := []
        for column in columns
            headers.Push(I18n.T(column[1]))
        listView := PreferencesWindow._Add("ListView", "w" PreferencesWindow.ContentW " h" (height - 40) " -Multi NoSort Grid", headers)
        for index, column in columns
            listView.ModifyCol(index, column[3])
        pageName := PreferencesWindow._current.Name
        statuses := Map()                                                   ; ObjPtr(item) -> check result, only after checking
        shown := []                                                         ; which item in items each row shows (filtering shows only some)
        filterBox := ""

        Refresh(selectRow := 0) {
            listView.Opt("-Redraw")
            listView.Delete()
            shown.Length := 0
            filter := IsObject(filterBox) ? filterBox.Value : ""
            for index, item in items {
                cells := []
                for column in columns
                    cells.Push((column[2] = "Status") ? StatusText(item) : PreferencesWindow._Cell(item, column[2]))
                if !PreferencesWindow.FilterMatch(filter, cells)
                    continue
                listView.Add("", cells*)
                shown.Push(index)
            }
            if (selectRow && shown.Length)
                listView.Modify(Min(selectRow, shown.Length), "Select Focus Vis")
            listView.Opt("+Redraw")
            PreferencesWindow.FitColumns(listView)
        }
        RowOf(index) {                                                      ; The row showing item index of items, 0 if not shown
            for row, shownIndex in shown
                if (shownIndex = index)
                    return row
            return 0
        }
        StatusText(item) {
            key := ObjPtr(item)
            return statuses.Has(key) ? I18n.T("Prefs.Status." statuses[key]) : ""
        }
        Recheck(item) {                                                     ; after a check, added / changed items are checked right away too
            if statuses.Count
                statuses[ObjPtr(item)] := check(item, Map())
        }
        AddItem(*) => Insert(newItem())
        Insert(item) {                                                      ; Opens the edit dialog (item = the prefilled content) and adds the result to the list on OK
            edited := ItemEditor.Edit(PreferencesWindow.Gui, pageName, fields, item)
            if IsObject(edited)
                Append(edited)
        }
        Append(item) {
            items.Push(item)
            PreferencesWindow.MarkDirty(true)
            Recheck(item)
            if IsObject(filterBox)
                filterBox.Value := ""                                       ; clear the filter so the new item is surely visible
            Refresh(items.Length)
        }
        EditExisting(item) {                                                ; Selects this item in the list and opens the edit dialog
            for index, existing in items {
                if (existing = item) {
                    if IsObject(filterBox)
                        filterBox.Value := ""
                    Refresh(index)
                    return EditItem()
                }
            }
        }
        EditItem(*) {
            row := listView.GetNext()
            if !row
                return
            index := shown[row]
            edited := ItemEditor.Edit(PreferencesWindow.Gui, pageName, fields, ItemEditor.WithDefaults(items[index], newItem()))
            if IsObject(edited) {
                if statuses.Has(ObjPtr(items[index]))
                    statuses.Delete(ObjPtr(items[index]))
                items[index] := edited
                PreferencesWindow.MarkDirty()
                Recheck(edited)
                Refresh(row)
            }
        }
        CheckItems(button, *) {
            button.Enabled := false
            statuses.Clear()
            roots := Map()
            counts := Map("OK", 0, "Missing", 0, "Unavailable", 0, "Skipped", 0)
            firstProblem := 0
            for index, item in items {
                result := check(item, roots)
                statuses[ObjPtr(item)] := result
                counts[result] += 1
                if (!firstProblem && (result = "Missing" || result = "Unavailable"))
                    firstProblem := index
            }
            Refresh(RowOf(firstProblem))
            button.Enabled := true
            if !firstProblem
                message := I18n.T("Prefs.CheckAllOk", counts["OK"])
            else
                message := I18n.T("Prefs.CheckResult", counts["Missing"], counts["Unavailable"])
            if counts["Skipped"]
                message .= "`n`n" I18n.T("Prefs.CheckSkipped", counts["Skipped"])
            MsgBox(message, pageName, (firstProblem ? "Icon!" : "Iconi") " Owner" PreferencesWindow.Gui.Hwnd)
        }
        DeleteItem(*) {
            row := listView.GetNext()
            if !row
                return
            if (MsgBox(I18n.T("Prefs.ConfirmDelete", listView.GetText(row, 1)), pageName, "YesNo Icon? Owner" PreferencesWindow.Gui.Hwnd) != "Yes")
                return
            items.RemoveAt(shown[row])
            PreferencesWindow.MarkDirty()
            Refresh(row)
        }

        listView.OnEvent("DoubleClick", EditItem)
        Refresh()
        ReloadList() {                                                      ; after items were changed from outside (import)
            if IsObject(filterBox)
                filterBox.Value := ""
            Refresh()
        }
        list := {Items: items, Insert: Insert, Append: Append, Edit: EditExisting, Reload: ReloadList}
        PreferencesWindow._lists[path] := list                              ; items added from outside (AddCommands)
        PreferencesWindow._y += height - 34
        PreferencesWindow._Add("Button", "w80", I18n.T("Prefs.Add")).OnEvent("Click", AddItem)
        PreferencesWindow._Add("Button", "x+8 yp w80", I18n.T("Prefs.Edit")).OnEvent("Click", EditItem)
        PreferencesWindow._Add("Button", "x+8 yp w80", I18n.T("Prefs.Delete")).OnEvent("Click", DeleteItem)
        if IsObject(check)
            PreferencesWindow._Add("Button", "x+24 yp w100", I18n.T("Prefs.CheckTargets")).OnEvent("Click", CheckItems)
        if IsObject(buttons)                                                ; the page's own buttons: [{Label, Width, Click: (list, button) => ...}]
            for extra in buttons
                PreferencesWindow._Add("Button", "x+24 yp w" extra.Width, extra.Label).OnEvent("Click", ((extra, button, *) => extra.Click.Call(list, button)).Bind(extra))
        listView.GetPos(&listX, , &listW)
        filterBox := PreferencesWindow.Gui.Add("Edit", "x" (listX + listW - 160) " yp+1 w160 r1 -Multi Hidden")   ; not _Add: filtering doesn't change the settings
        PreferencesWindow._current.Controls.Push(filterBox)
        Win.SetCueBanner(filterBox.Hwnd, I18n.T("Prefs.Filter"))
        filterBox.OnEvent("Change", (*) => Refresh())
        PreferencesWindow._y += 38
    }

    ; Filter box: shown only when every word of filter (separated by spaces) appears in some column, case-insensitive; all shown when empty
    static FilterMatch(filter, cells) {
        filter := Trim(filter)
        if (filter = "")
            return true
        text := ""
        for cell in cells
            text .= cell "`n"
        for word in StrSplit(filter, " ")
            if (word != "" && !InStr(text, word))
                return false
        return true
    }

    static _Cell(item, key) {
        if !(item is Map) || !item.Has(key)
            return ""
        if (key = "Target" && item.Has("Arguments") && Trim(item["Arguments"]) != "")   ; custom commands: the arguments follow the target
            return RegExReplace(item["Target"] " " item["Arguments"], "\s+", " ")
        if (key = "Key")                                                    ; custom hotkeys: shown as Ctrl+Alt+P, Middle button
            return HotkeyBox.Label(item[key])
        if (key = "Type") {                                                 ; the type of a custom command shows its translated name
            label := I18n.T("Prefs.TypeShort." item[key])
            if (label != "Prefs.TypeShort." item[key])
                return label
        }
        return RegExReplace(item[key] "", "\s+", " ")
    }

    ;---------------------------------------------------------------------------
    ; Data helpers (pure functions, with unit tests)
    ;---------------------------------------------------------------------------
    ; Global hotkeys (not tied to a window) -> [{Key, Label}...]: the show hotkey, the second show hotkey, clipboard history, custom hotkeys without a window.
    ; Quick switch and date stamp only work in particular windows (two Ctrl+D are fine), so they aren't included
    static _GlobalHotkeys(data) {
        list := []
        add(key, label) => (Trim(key) != "") ? list.Push({Key: key, Label: label}) : ""
        general := data["General"]
        add(general["Hotkey"], I18n.T("Prefs.Hotkey"))
        add(general["SecondaryHotkey"], I18n.T("Prefs.SecondaryHotkey"))
        add(general.Has("SelectionHotkey") ? general["SelectionHotkey"] : "", I18n.T("Prefs.SelectionHotkey"))
        clipboard := PreferencesWindow.GetPath(data, "Features.Clipboard")
        if (clipboard is Map && clipboard.Has("Enabled") && clipboard["Enabled"] && clipboard.Has("Hotkey"))
            add(clipboard["Hotkey"], I18n.T("Prefs.Page.Clipboard"))
        if (data.Has("Hotkeys") && data["Hotkeys"] is Array)
            for entry in data["Hotkeys"]
                if (entry is Map && entry.Has("Key") && (!entry.Has("WinTitle") || Trim(entry["WinTitle"]) = ""))
                    add(entry["Key"], I18n.T("Prefs.Page.Hotkeys") " (" (entry.Has("Action") ? entry["Action"] : "") ")")
        return list
    }

    ; The first duplicate hotkey -> {Key, First, Second} (the names of both); "" without duplicates. ~ * $ don't change whether it is the same key
    static DuplicateHotkey(list) {
        seen := Map()
        for entry in list {
            key := StrLower(RegExReplace(Trim(entry.Key), "^[~*$]+"))
            if RegExMatch(key, "^([\^!+#]+)(.+)$", &m)                     ; a different modifier order is the same key (!^c = ^!c)
                key := StrReplace(Sort(RegExReplace(m[1], "(.)", "$1`n")), "`n") m[2]
            if seen.Has(key)
                return {Key: entry.Key, First: seen[key], Second: entry.Label}
            seen[key] := entry.Label
        }
        return ""
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
