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
    static _lists := Map()                                                  ; setting path of a list page (e.g. "CustomCommands") -> the method that adds / edits items
    ; the order of the pages (matching PageBuilders, checked by a test); search entries like "ALTRun Preferences: Appearance" come from it without opening the window
    static PageKeys := ["Prefs.Page.General", "Prefs.Page.Window", "Prefs.Page.Keyboard", "Prefs.Page.Appearance", "Prefs.Page.Features", "Prefs.Page.Applications"
        , "Prefs.Page.FileSearch", "Prefs.Page.FileIndex", "Prefs.Page.Commands", "Prefs.Page.Snippets", "Prefs.Page.Clipboard", "Prefs.Page.WebSearch"
        , "Prefs.Page.Calculator", "Prefs.Page.Scripts", "Prefs.Page.Hotkeys", "Prefs.Page.QuickSwitch", "Prefs.Page.QSPanel", "Prefs.Page.DateStamp"
        , "Prefs.Page.Usage", "Prefs.Page.Advanced"]
    static _page := 0
    static _dirty := false, _ready := false, ApplyButton := ""
    static MinWindowWidth := 450                                            ; the search window: room for the status bar's four parts
    static _positionReset := false
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
        PreferencesWindow.Binds := [], PrefsForm.Wider := Map(), PreferencesWindow._lists := Map(), PreferencesWindow._helpButton := ""
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
        saved := PrefsForm._y
        PrefsForm._y := y + (h - 16) // 2
        try PrefsForm.Add("Picture", "x" (x + 18 + textW + 6) " w16 h16", "HICON:" hIcon)
        PrefsForm._y := saved
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
    static PageBuilders := ["BuildGeneral", "BuildWindow", "BuildKeyboard", "BuildAppearance", "BuildFeatures", "BuildApplications"
        , "BuildFileSearch", "BuildFileIndex", "BuildCommands", "BuildSnippets", "BuildClipboard", "BuildWebSearch"
        , "BuildCalculator", "BuildScripts", "BuildHotkeys", "BuildQuickSwitch", "BuildQuickSwitchPanel", "BuildDateStamp"
        , "BuildUsage", "BuildAdvanced"]
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
