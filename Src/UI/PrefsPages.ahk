;===============================================================================
; PrefsPages.ahk - the pages of the Preferences window (AutoHotkey v2)
;-------------------------------------------------------------------------------
; One BuildXxx method per page, in the order of PreferencesWindow.PageKeys, built with the helpers in PrefsForm,
; plus the actions of their buttons (import / export snippets, change the data folder, usage statistics...).
; PreferencesWindow builds a page the first time it is shown (see PreferencesWindow._BuildPage).
;===============================================================================

class PrefsPages {
    static BuildGeneral() {
        PrefsForm.BeginPage("Prefs.Page.General", 120)
        PrefsForm.Pair(["General.Hotkey", "Prefs.Hotkey", "K", "hotkey"], ["General.SecondaryHotkey", "Prefs.SecondaryHotkey", "K", "hotkey"])
        PrefsForm.Field("General.SelectionHotkey", "Prefs.SelectionHotkey", "K", "hotkey")
        PrefsForm.Pair(["General.DoubleTap", "Prefs.DoubleTap", "K", "choice", ["", "Ctrl", "Shift"], [I18n.T("Prefs.DoubleTap.None"), I18n.T("Prefs.DoubleTap.Ctrl"), I18n.T("Prefs.DoubleTap.Shift")]]
            , ["General.Language", "Prefs.Language", "K", "choice", PrefsPages._LanguageValues(), PrefsPages._LanguageLabels()])
        PrefsForm.Gap()
        for row in [["LaunchAtLogin", "Prefs.LaunchAtLogin", "Prefs.Group.Startup"], ["ShowTrayIcon", "Prefs.ShowTrayIcon", ""]
                   , ["SendToMenu", "Prefs.SendToMenu", "Prefs.Group.Integration"], ["StartMenuShortcut", "Prefs.StartMenu", ""]
                   , ["CheckForUpdates", "Prefs.CheckUpdates", "Prefs.Group.Updates"]]   ; writing the debug log is under Advanced → Diagnostics
            PrefsForm.Check("General." row[1], row[2], , , row[3])
        PrefsForm.Gap()
        PrefsForm.Field("General.FileManager", "Prefs.FileManager", "L", "file")
    }

    ; Search window behavior (the settings are still in General, just on a page of their own)
    static BuildWindow() {
        PrefsForm.BeginPage("Prefs.Page.Window", 16)                         ; arranged in sections, each section's content indented a little
        PrefsForm.Section("Prefs.Group.Window")
        PrefsForm.Check("General.HideOnDeactivate", "Prefs.HideOnDeactivate")
        PrefsForm.Check("General.KeepLastQuery", "Prefs.KeepLastQuery")
        PrefsForm.Check("Appearance.StatusBar", "Prefs.StatusBar")                ; the usage tips are in the status bar (in the empty search box when the status bar is off): kept together
        PrefsForm.Check("General.ShowTips", "Prefs.ShowTips")
        PrefsForm.Section("Prefs.Section.EmptyBox")
        placeholder := PrefsForm.InlineField("Appearance.Placeholder", "Prefs.Placeholder", "M")
        Win.SetCueBanner(placeholder.Hwnd, I18n.T("Search.Placeholder"))  ; empty shows the default text in gray
        recent := PrefsForm.InlineField("Features.Recent.RecentCount", "Prefs.RecentCount", "S", "number")
        PrefsForm.AlignLeft(placeholder, recent)                            ; the input boxes on this page share one left edge
    }

    ; Keyboard: input method switching (Caps Lock, English when shown) and the keys in the search box.
    ; A two-column form like the General page: labels right-aligned on the left, every control starting on one line
    static BuildKeyboard() {
        PrefsForm.BeginPage("Prefs.Page.Keyboard", 150)
        PrefsForm.Section("Prefs.Group.InputMethod")
        capsLabels := [I18n.T("Prefs.CapsLock.Off"), I18n.T("Prefs.CapsLock.Layout"), I18n.T("Prefs.CapsLock.Mode")]
        capsLock := PrefsForm.Choice("General.CapsLock", "Prefs.CapsLock", ["", "Layout", "Mode"], capsLabels)            ; Caps Lock switches the input method (CapsLockSwitch)
        notice := PrefsForm.Check("General.CapsLockNotify", "Prefs.CapsLockNotify")            ; hint when Caps Lock turns on / off
        elevated := PrefsForm.Check("General.CapsLockElevated", "Prefs.CapsLockElevated")            ; also in administrator windows (CapsLockHelper)
        controls := PrefsForm._current.Controls
        elevatedDesc := controls[controls.Length]                           ; the description just added (sized for two lines, so the shorter text fits too)
        PreferencesWindow._ShieldAfter(elevated)                            ; turning it on / off asks for UAC
        update := (*) => (notice.Enabled := elevated.Enabled := capsLock.Value > 1
            , elevatedDesc.Text := I18n.T(elevated.Enabled ? "Prefs.CapsLockElevated.Desc" : "Prefs.CapsLockElevated.NeedsMode"))   ; when grayed out, say why
        update()
        capsLock.OnEvent("Change", update)
        englishLabels := [I18n.T("Prefs.EnglishInput.Off"), I18n.T("Prefs.EnglishInput.Keyboard"), I18n.T("Prefs.EnglishInput.Mode")]
        english := PrefsForm.Choice("General.SwitchToEnglishInput", "Prefs.EnglishInput", [0, 1, 2], englishLabels)            ; switch to English when shown
        restore := PrefsForm.Check("General.RestoreInput", "Prefs.RestoreInput")
        restore.Enabled := english.Value > 1
        english.OnEvent("Change", (*) => restore.Enabled := english.Value > 1)
        PrefsForm.Section("Prefs.Group.SearchKeys")
        PrefsForm.Check("General.SpaceToRun", "Prefs.SpaceToRun", , , "Prefs.Group.SpaceKey")
        tabLabels := [I18n.T("Prefs.TabKey.Complete"), I18n.T("Prefs.TabKey.Next")]
        tab := PrefsForm.Choice("General.TabKey", "Prefs.TabKey", ["Complete", "Next"], tabLabels)
        PrefsForm.Field("General.HistorySize", "Prefs.HistorySize", "S", "number")            ; Ctrl+↑ / ↓ recall earlier searches
        PrefsForm.SameWidth(capsLock, english, tab)
    }

    static BuildAppearance() {
        PrefsForm.BeginPage("Prefs.Page.Appearance", 150)
        appearance := PreferencesWindow.Working["Appearance"]
        if ((base := ThemeManager.LegacyCompact(appearance["Theme"])) != "") {  ; a former compact theme: select the matching theme, density at least Normal (takes effect after saving)
            appearance["Theme"] := base
            if (appearance["Density"] = "Comfortable")
                appearance["Density"] := "Normal"
        }
        PrefsForm.Section("Prefs.Section.Theme")
        themes := ThemeManager.Names()
        gallery := PrefsPages._ThemeGallery("Appearance.Theme", themes, 164)          ; one and a half rows of thumbnails (shows it can scroll), leaving room for "Compact" below
        PrefsForm.Buttons("Prefs.Group.CustomThemes"
            , ["Prefs.CopyTheme", (*) => PrefsPages._CopyTheme(gallery, themes)]
            , ["Prefs.OpenThemes", (*) => PrefsPages._OpenFolder(ThemeManager.UserDir)])
        PrefsForm.Section("Prefs.Section.Size")
        PrefsForm.Pair(["Appearance.Width", "Prefs.Width", "S", "number"]
            , ["Appearance.VisibleRows", "Prefs.VisibleRows", "S", "choice", [1, 2, 3, 4, 5, 6, 7, 8, 9], ["1", "2", "3", "4", "5", "6", "7", "8", "9"]])   ; a list: nothing out of range to type
        for ctrl in PrefsForm._current.Controls                              ; the drop-down Pair just added
            if (ctrl.Type = "DDL")
                visible := ctrl
        density := PrefsForm.Choice("Appearance.Density", "Prefs.Density", ThemeManager.Densities            ; any theme can be scaled (ThemeManager.Scaled)
            , [I18n.T("Prefs.Density.Comfortable"), I18n.T("Prefs.Density.Normal"), I18n.T("Prefs.Density.Compact")])
        PrefsForm.Check("Appearance.Fade", "Prefs.Fade", , , "Prefs.Group.Animation")
        PrefsForm.Section("Prefs.WindowPosition")
        showOn := PrefsForm.Choice("Appearance.ShowOn", "Prefs.ShowOn", ["Mouse", "Primary", "Active"]
            , [I18n.T("Prefs.ShowOn.Mouse"), I18n.T("Prefs.ShowOn.Primary"), I18n.T("Prefs.ShowOn.Active")])
        PrefsForm.Check("Appearance.RememberPosition", "Prefs.RememberPosition", , , "Prefs.Group.Dragging")
        reset := PrefsForm.Button("Prefs.ResetPosition", (button, *) => PrefsPages._ResetPosition(button))
        PrefsForm.SameWidth(density, showOn, reset)                         ; one column: the two drop-downs and the button below them
        ; "Visible results" a little wider than its longest option, and that column ends on the same line as it
        visible.GetPos(&vx, , &vw), density.GetPos(&dx, , &dw)
        right := Min(Max(vx + vw + 16, dx + dw), PrefsForm.ContentX + PrefsForm.ContentW)
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
        list := PrefsForm.Add("ListView", "x" PrefsForm.ContentX " w" PrefsForm.ContentW " h" height
            . " Icon -Multi -Hdr", ["Theme"])
        DllCall("uxtheme\SetWindowTheme", "Ptr", list.Hwnd, "WStr", "Explorer", "Ptr", 0)   ; the selected item gets a light frame instead of tinting the thumbnail blue
        list.SetImageList(ThemePreview.ImageList(themes, thumbW, thumbH), 0)   ; 0 = large icons
        SendMessage(0x1035, 0, PrefsPages.GalleryColumnWidth(list, thumbW) | ((thumbH + Win.Scale(30)) << 16), list.Hwnd)          ; LVM_SETICONSPACING
        current := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        selected := 1
        for index, themeName in themes {
            list.Add("Icon" index, PrefsPages._ThemeLabel(themeName))
            if (themeName = current)
                selected := index
        }
        list.Modify(selected, "Select Focus Vis")
        list.OnNotify(-101, (ctrl, lParam) => PrefsPages._OnGalleryChange(lParam))          ; LVN_ITEMCHANGED (selection by the program counts too)
        PrefsForm.Bind(path, () => themes[list.GetNext() || selected])
        PrefsForm.Below(8, list)
        return list
    }

    ; Width of each cell of the theme list: exactly GalleryColumns thumbnails per row, spread evenly (no big gap on the right), with 2 pixels to spare so the last doesn't wrap.
    ; more than two rows of themes have a vertical scroll bar: subtract it first
    static GalleryColumns := 4

    static GalleryColumnWidth(list, thumbW) {
        rect := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", list.Hwnd, "Ptr", rect)
        width := NumGet(rect, 8, "Int") - (DllCall("GetWindowLong", "Ptr", list.Hwnd, "Int", -16, "Int") & 0x200000 ? 0 : SysGet(2))   ; subtract SM_CXVSCROLL while WS_VSCROLL hasn't appeared yet
        return Max(thumbW + Win.Scale(8), (width - 2) // PrefsPages.GalleryColumns)
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

    static BuildFeatures() {
        PrefsForm.BeginPage("Prefs.Page.Features", 140)
        PrefsForm.Section("Prefs.EnabledFeatures")
        features := ["Applications", "CustomCommands", "Snippets", "Clipboard", "Calculator", "WebSearch", "Bookmarks", "Windows", "Recent", "FileSearch", "Scripts", "Terminal", "System", "Help"]
        startY := PrefsForm._y, columnW := PrefsForm.InputW() // 2                       ; two columns: the English names are long, three would wrap
        rowH := 20                                                          ; row spacing: at least 4 more than a check box (Chinese and Japanese check boxes are taller)
        for index, feature in features {
            column := Mod(index - 1, 2), row := (index - 1) // 2
            PrefsForm._y := startY + row * rowH
            check := PrefsForm.Check("Features." feature ".Enabled", "Prefs.Feature." feature
                , PrefsForm.InputX() + column * columnW, , (index = 1) ? "Prefs.Group.SearchFeatures" : "", columnW)
            if (index = 1) {
                check.GetPos(, , , &checkH)
                rowH := Max(rowH, checkH + 4)
            }
        }
        PrefsForm._y := startY + (features.Length // 2) * rowH                ; last row: Large Type (not a search feature), Windows Settings pages (part of the system commands)
        PrefsForm.Check("Features.LargeType.Enabled", "Prefs.Feature.LargeType", PrefsForm.InputX(), , "", columnW)
        PrefsForm._y := startY + (features.Length // 2) * rowH
        PrefsForm.Check("Features.System.SettingsPages", "Prefs.SettingsPages", PrefsForm.InputX() + columnW, , "", columnW)
        PrefsForm._y := startY + Ceil((features.Length + 1) / 2) * rowH
        PrefsForm.Section("Prefs.Section.FeatureOptions")
        PrefsForm.Check("Features.System.ConfirmActions", "Prefs.ConfirmActions", , , "Prefs.Feature.System")
        PrefsForm.LinesButton("Features.System.Hidden", "Prefs.SysHidden")
        PrefsForm._y += 4                                                   ; a little less than Gap: this page is full
        PrefsForm.Pair(["Features.Terminal.Prefix", "Prefs.TerminalPrefix", "S"]
            , ["Features.Terminal.Shell", "Prefs.TerminalShell", "M", "choice", ["cmd", "powershell", "pwsh", "wt"], ["Command Prompt (cmd)", "Windows PowerShell", "PowerShell 7 (pwsh)", "Windows Terminal (wt)"]])
        for row in [["Bookmarks", "Prefs.Feature.Bookmarks"], ["Windows", "Prefs.Feature.Windows"]] {   ; bookmarks and windows: also in normal search + a keyword that searches only them
            rowY := PrefsForm._y
            PrefsForm.Check("Features." row[1] ".InDefaultResults", "Prefs.ShowInNormal", , , row[2], columnW)
            field := PrefsForm.SideField(rowY, PrefsForm.InputX() + columnW, "Features." row[1] ".Keyword", "Prefs.FeatureKeyword")                     ; aligned with the right column of the features above
            field.GetPos(, &fieldY, , &fieldH)
            PrefsForm._y := Max(rowY + 25, fieldY + fieldH + 8)             ; keyword boxes on two rows at least 5 apart (Chinese / Japanese input boxes are taller)
        }
        PrefsForm._y -= 3
        PrefsForm.Desc("Prefs.FeatureKeyword", PrefsForm.InputX())
    }

    static BuildApplications() {
        PrefsForm.BeginPage("Prefs.Page.Applications", 130)
        PrefsForm.Section("Prefs.Section.AppIndex")
        PrefsForm.LinesButton("Features.Applications.Folders", "Prefs.AppFolders")
        PrefsForm.Csv("Features.Applications.FileTypes", "Prefs.AppFileTypes", "L")
        PrefsForm.Field("Features.Applications.Exclude", "Prefs.AppExclude", "L")
        PrefsForm.FieldButton("Features.Applications.Depth", "Prefs.AppDepth", "Prefs.RebuildIndex", (button, *) => PrefsPages._RebuildApps(button))                   ; no refresh interval: the folders are watched and updated on changes
        PrefsForm.Section("Prefs.Section.AppResults")
        PrefsForm.LinesButton("Features.Applications.Hidden", "Prefs.AppHidden", {Summary: (lines) => PrefsPages.NameSummary(lines)})
        PrefsForm.Check("Features.Applications.StoreApps", "Prefs.StoreApps", , , "Prefs.Group.Options")
        PrefsForm.Check("Features.Applications.MatchPinyin", "Prefs.MatchPinyin")
    }

    static BuildFileSearch() {
        PrefsForm.BeginPage("Prefs.Page.FileSearch", 130)
        PrefsForm.Section("Prefs.Section.FileStart")
        PrefsForm.Check("Features.FileSearch.SpacePrefix", "Prefs.SpacePrefix", , , "Prefs.Group.StartFileSearch")
        PrefsForm.Check("Features.FileSearch.QuotePrefix", "Prefs.QuotePrefix")
        PrefsForm.Pair(["Features.FileSearch.Keywords", "Prefs.FileKeywords", "M", "csv"], ["Features.FileSearch.FolderKeywords", "Prefs.FolderKeywords", "S", "csv"])
        PrefsForm.Csv("Features.FileSearch.RecentKeywords", "Prefs.RecentKeywords", "M")
        PrefsForm.LinesButton("Features.FileSearch.TypeFilters", "Prefs.TypeFilters", {Summary: (lines) => PrefsPages.TypeFilterSummary(lines), AlignPairs: true})
        PrefsForm.Section("Prefs.Section.FileResults")
        PrefsForm.Check("Features.FileSearch.InDefaultResults", "Prefs.FileInDefault", , , "Prefs.Group.NormalSearch")
        PrefsForm.Pair(["Features.FileSearch.MaxResults", "Prefs.FileMaxResults", "S", "number"], ["Features.FileSearch.DefaultResultsLimit", "Prefs.FileDefaultLimit", "S", "number"])
        PrefsForm.Field("Features.FileSearch.MinQueryLength", "Prefs.FileMinLength", "S", "number")
        PrefsForm.Section("Prefs.Section.Browse")                           ; browsing folders is always available: just a description
    }

    ; Search sources: Everything (searches the whole disk while it runs) and the built-in file index when Everything isn't running
    static BuildFileIndex() {
        PrefsForm.BeginPage("Prefs.Page.FileIndex", 130)
        PrefsForm.Section("Prefs.Section.Everything")
        pinyinHelp := HelpProvider.WikiPage("File-Search") (SubStr(I18n.Lang, 1, 2) = "zh" ? "#常见问题" : "#troubleshooting")
        PrefsForm.LinkDesc(I18n.T("Prefs.EverythingPinyin", pinyinHelp), PrefsForm.ContentX)                    ; the section's description: how to make Everything support pinyin (to compare with the pinyin note in the built-in index section)
        PrefsForm.Check("Features.FileSearch.UseEverything", "Prefs.UseEverything", , I18n.T(Everything.IsRunning() ? "Prefs.Running" : "Prefs.NotRunning"))            ; whether it is running now
        PrefsForm.Field("Features.FileSearch.EverythingFilter", "Prefs.EverythingFilter", "L")
        PrefsForm.Field("Features.FileSearch.EverythingPath", "Prefs.EverythingPath", "L", "folder")
        PrefsForm.Section("Prefs.Section.BuiltinIndex")
        PrefsForm.LinesButton("Features.FileSearch.ScopeFolders", "Prefs.ScopeFolders")
        PrefsForm.FieldButton("Features.FileSearch.ScopeDepth", "Prefs.ScopeDepth", "Prefs.RebuildFileIndex", (button, *) => PrefsPages._RebuildFiles(button))
        PrefsForm.Field("Features.FileSearch.ScopeExclude", "Prefs.ScopeExclude", "L")
        PrefsForm.Pair(["Features.FileSearch.MaxEntries", "Prefs.MaxEntries", "S", "number"], ["Features.FileSearch.RefreshMinutes", "Prefs.RefreshMinutes", "S", "number"])
    }

    static BuildCommands() {
        PrefsForm.BeginPage("Prefs.Page.Commands")
        PrefsForm.List("CustomCommands", 540
            , [["Prefs.Col.Title", "Title", 140], ["Prefs.Col.Type", "Type", 70], ["Prefs.Col.Target", "Target", 190], ["Prefs.Col.Keyword", "Keyword", 70], ["Prefs.Col.Status", "Status", 70]]
            , CustomCommandProvider.EditorFields(), () => CustomCommandProvider.NewCommand()
            , (item, roots) => CustomCommandProvider.CheckTarget(item, roots))
        PrefsForm.Note("Prefs.CommandsNote")
    }

    static BuildSnippets() {
        PrefsForm.BeginPage("Prefs.Page.Snippets", 150)
        PrefsForm.List("Snippets", 220
            , [["Prefs.Col.Name", "Name", 150], ["Prefs.Col.Keyword", "Keyword", 90], ["Prefs.Col.Text", "Text", 300]]
            , SnippetProvider.EditorFields(), () => SnippetProvider.NewSnippet(), ""
            , [{Label: I18n.T("Snippet.ImportExport"), Width: 100, Click: (list, button) => PrefsPages._SnippetMenu(list, button)}])
        PrefsForm.Section("Prefs.Section.Options")
        PrefsForm.Check("Features.Snippets.SearchText", "Prefs.SnippetSearchText", , , "Prefs.Group.Search")
        PrefsForm.Check("Features.Snippets.AutoExpand", "Prefs.SnippetAutoExpand", , , "Prefs.Group.AutoExpand")
        ; Two rows of the same short boxes: Search keyword / Paste by on the left, Keyword prefix / Delay on the right
        keywordRow := PrefsForm.Pair(["Features.Snippets.Keyword", "Prefs.SnippetKeyword", "K"], ["Features.Snippets.ExpandPrefix", "Prefs.ExpandPrefix", "K"])
        pasteRow := PrefsForm.Pair(["Features.Snippets.PasteMode", "Prefs.PasteMode", "K", "choice", ["Auto", "Clipboard", "Type"], [I18n.T("Prefs.PasteMode.Auto"), I18n.T("Prefs.PasteMode.Clipboard"), I18n.T("Prefs.PasteMode.Type")]]
            , ["Features.Snippets.PasteDelay", "Prefs.PasteDelay", "K", "number"])
        PrefsForm.AlignPairRows(keywordRow, pasteRow)
        PrefsForm.LinesButton("Features.Snippets.ExpandExclude", "Prefs.ExpandExclude")
    }

    ; The "Import / Export" button on the Snippets page: shows a menu
    static _SnippetMenu(list, button) {
        popup := Menu()
        popup.Add(I18n.T("Snippet.Import"), (*) => PrefsPages._ImportSnippets(list))
        popup.Add(I18n.T("Snippet.Export"), (*) => PrefsPages._ExportSnippets(list))
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

    static BuildClipboard() {
        PrefsForm.BeginPage("Prefs.Page.Clipboard", 120)
        PrefsForm.Pair(["Features.Clipboard.Hotkey", "Prefs.ClipHotkey", "K", "hotkey"], ["Features.Clipboard.Keyword", "Prefs.ClipKeyword", "S"])
        PrefsForm.Pair(["Features.Clipboard.MaxItems", "Prefs.ClipMaxItems", "S", "number"], ["Features.Clipboard.MaxItemLength", "Prefs.ClipMaxLength", "S", "number"])
        PrefsForm.Check("Features.Clipboard.Persist", "Prefs.ClipPersist", , , "Prefs.Group.History")
        PrefsForm.Check("Features.Clipboard.Images", "Prefs.ClipImages")
        PrefsForm.Field("Features.Clipboard.MaxImages", "Prefs.ClipMaxImages", "S", "number")
        PrefsForm.Choice("Features.Clipboard.LocalFiles", "Prefs.ClipStorage", [0, 1, 2]            ; where it is stored (ClipboardProvider.PlaceStorage)
            , [I18n.T("Prefs.ClipStorage.Data"), I18n.T("Prefs.ClipStorage.Files"), I18n.T("Prefs.ClipStorage.All")], "L")   ; long options: full width
        PrefsForm.Check("Features.Clipboard.MergeDoubleCopy", "Prefs.ClipMerge", , , "Prefs.Group.Copy")
        PrefsForm.Gap()
        PrefsForm.LinesButton("Features.Clipboard.IgnoreApps", "Prefs.ClipIgnoreApps")
        PrefsForm.Button("Prefs.ClipClear", (*) => ClipboardProvider.Clear())
    }

    static BuildWebSearch() {
        PrefsForm.BeginPage("Prefs.Page.WebSearch", 130)
        PrefsForm.List("Features.WebSearch.Engines", 320
            , [["Prefs.Col.Keyword", "Keyword", 70], ["Prefs.Col.Title", "Title", 130], ["Prefs.Col.Url", "Url", 340]]
            , WebSearchProvider.EditorFields(), () => WebSearchProvider.NewEngine())
        PrefsForm.Csv("Features.WebSearch.Fallbacks", "Prefs.Fallbacks", "L")
    }

    ; Calculator: currency conversion, structural calculation (beam main bars / steel area) and its parameters
    static BuildCalculator() {
        base := "Features.Calculator."
        PrefsForm.BeginPage("Prefs.Page.Calculator", 150)
        PrefsForm.Section("Prefs.Section.Calculator")
        PrefsForm.Check(base "Currency", "Prefs.Currency")
        PrefsForm.Section("Prefs.Section.Structural")
        PrefsForm.Check(base "StructuralCalc", "Prefs.StructuralCalc")
        PrefsForm.Field(base "RebarCover", "Prefs.RebarCover", "S", "number")               ; one row per item: labels and input boxes each aligned
        PrefsForm.Field(base "MaxBarSpacing", "Prefs.MaxBarSpacing", "S", "number")
        PrefsForm.Csv(base "BarSizes", "Prefs.BarSizes", "M", true)
        PrefsForm.Field(base "BarPrefix", "Prefs.BarPrefix", "S")
    }

    ; Scripts extension: description, the Scripts folder, the scripts found (double-click edits in Notepad); the switch is on the Features page
    static BuildScripts() {
        PrefsForm.BeginPage("Prefs.Page.Scripts", 110)
        PrefsForm.Section("Prefs.Section.Scripts")
        folder := PrefsForm.Info("Prefs.Group.ScriptsFolder", ScriptProvider.Dir, " r1 cGray 0x8100")            ; like the data folder on the Advanced page: an ellipsis in the middle when too long
        Win.AddTooltip(folder.Hwnd, ScriptProvider.Dir)
        PrefsForm._y += 4
        PrefsForm.Button("Prefs.ScriptsFolder", (*) => (ScriptProvider.OpenFolder(), PrefsPages._FillScripts()))
        PrefsForm.Field("Features.Scripts.Keyword", "Prefs.ScriptKeyword", "S")
        PrefsForm.Section("Prefs.Section.ScriptList")
        listView := PrefsForm.Add("ListView", "w" PrefsForm.ContentW " h170 -Multi NoSort Grid ReadOnly"
            , [I18n.T("Prefs.Col.Name"), I18n.T("Prefs.Col.Keyword"), I18n.T("Prefs.Col.Argument"), I18n.T("Prefs.Col.Mode"), I18n.T("Prefs.Col.File")])
        for index, width in [145, 70, 95, 135, 100]
            listView.ModifyCol(index, width)
        listView.OnEvent("DoubleClick", (ctrl, row) => (row && row <= ScriptProvider.Scripts.Length) ? Run('notepad.exe "' ScriptProvider.Scripts[row].Path '"') : "")
        PrefsPages._scriptList := listView
        PrefsPages._FillScripts()
        PrefsForm.Below(6, listView)
        PrefsForm.Note("Prefs.ScriptsNote")
    }

    static _scriptList := ""

    static _FillScripts() {
        static modes := Map("window", "Script.Mode.Window", "silent", "Script.Mode.Silent", "output", "Script.Mode.Output")
        listView := PrefsPages._scriptList
        if !IsObject(listView)
            return
        ScriptProvider.Load()
        listView.Delete()
        for script in ScriptProvider.Scripts
            listView.Add(, script.Title, script.Keyword, script.Argument, I18n.T(modes[script.Mode]), script.Name)
        if !ScriptProvider.Scripts.Length
            listView.Add(, I18n.T("Prefs.NoScripts"))
        PrefsForm.FitColumns(listView)
    }

    static BuildHotkeys() {
        PrefsForm.BeginPage("Prefs.Page.Hotkeys", 100)                         ; narrow label column: the list starts at the left edge (Caps Lock is on the Keyboard page)
        actions := [["ToggleWindow", "Show / hide ALTRun (ToggleWindow)"]]
        for command in SystemProvider.Commands()
            actions.Push([command["Id"], command["Title"] " (" command["Id"] ")"])
        PrefsForm.List("Hotkeys", 500
            , [["Prefs.Col.Key", "Key", 160], ["Prefs.Col.Action", "Action", 150], ["Prefs.Col.WinTitle", "WinTitle", 230]]
            , [ItemEditor.Field("Key", "Prefs.Col.Key", "hotkey", true, "", I18n.T("Prefs.HotkeyHint"))
             , ItemEditor.Field("Action", "Prefs.Col.Action", "choice", false, actions)
             , ItemEditor.Field("WinTitle", "Prefs.Col.WinTitle", "text", false, "", I18n.T("Prefs.WinTitleHint"))]
            , () => Map("Key", "", "Action", "ToggleWindow", "WinTitle", ""))
        PrefsForm.Note("Prefs.HotkeysNote")
    }

    ; Quick switch for dialogs (Listary's Quick Switch)
    static BuildQuickSwitch() {
        base := "Extensions.QuickSwitch."
        PrefsForm.BeginPage("Prefs.Page.QuickSwitch", 150)
        PrefsForm.Section("Prefs.QuickSwitch")
        PrefsForm.Check(base "Enabled", "Prefs.EnableExtension", , , "Prefs.Group.Status")
        PrefsForm.Field(base "Hotkey", "Prefs.QSHotkey", "K", "hotkey")            ; one hotkey for TC / Opus / Explorer: the one used last

        PrefsForm.Section("Prefs.Section.QSDialogs")
        PrefsForm.Check(base "StandardDialogs", "Prefs.QSStandard", , , "Prefs.Group.DialogWindows")            ; ahk_class #32770
        PrefsForm.Check(base "BrowseDialogs", "Prefs.QSBrowse")            ; "Browse for Folder": folder tree only (QuickSwitch.LooksLikeBrowseDialog)
        PrefsForm.LinesButton(base "DialogWindows", "Prefs.QSOtherDialogs")
        PrefsForm.LinesButton(base "ExcludeWindows", "Prefs.QSExclude")

        PrefsForm.Section("Prefs.Section.QSAuto")
        PrefsForm.Check(base "AutoSwitch", "Prefs.QSAuto", , , "Prefs.Group.Options")
        PrefsForm.LinesButton(base "AutoSwitchExclude", "Prefs.QSAutoExclude")
    }

    ; The dialog panel (part of QuickSwitch): the folder list and search box below open / save dialogs, and the folder menu for the keyboard
    static BuildQuickSwitchPanel() {
        base := "Extensions.QuickSwitch."
        PrefsForm.BeginPage("Prefs.Page.QSPanel", 150)
        PrefsForm.Section("Prefs.QSPanelSection")
        PrefsForm.Check(base "ShowPanel", "Prefs.QSPanel", , , "Prefs.Group.Status")
        PrefsForm.Choice(base "PanelSearch", "Prefs.QSPanelSearch", ["folders", "all"], [I18n.T("Prefs.QSPanelSearch.Folders"), I18n.T("Prefs.QSPanelSearch.All")], "K")
        PrefsForm.Field(base "RecentFolders", "Prefs.QSRecent", "S", "number")
        PrefsForm.Section("Prefs.QSMenuSection")
        PrefsForm.Field(base "MenuHotkey", "Prefs.QSMenu", "K", "hotkey")
    }

    ; Date stamp (AutoDate): hotkeys and target windows for renaming files and for comment boxes
    static BuildDateStamp() {
        base := "Extensions.AutoDate."
        PrefsForm.BeginPage("Prefs.Page.DateStamp", 150)
        PrefsForm.Section("Prefs.AutoDate")
        PrefsForm.Check(base "Enabled", "Prefs.EnableExtension", , , "Prefs.Group.Status")
        PrefsForm.Field(base "Hotkey", "Prefs.RenameHotkey", "K", "hotkey")            ; one hotkey: the window decides between the two below
        PrefsForm.Field(base "DateFormat", "Prefs.DateFormat", "K")            ; as wide as the hotkey box above
        PrefsForm.Section("Prefs.Section.DateRename")
        PrefsForm.LinesButton(base "RenameWindows", "Prefs.RenameWindows")
        PrefsForm.Section("Prefs.Section.DateAppend")
        PrefsForm.LinesButton(base "AppendWindows", "Prefs.AppendWindows")
    }

    ; "About" at the top (icon, name, version, home page, check for updates), the settings and data file locations and actions below
    static BuildAdvanced() {
        PrefsForm.BeginPage("Prefs.Page.Advanced", 110)
        textX := PrefsForm.InputX(), textW := PrefsForm.InputW(), top := PrefsForm._y + 6
        PrefsForm._y := top                                                 ; the icon in the left column; text and buttons aligned with the control column below
        if FileExist(App.IconFile)                                         ; without the icon file (e.g. only the exe was copied) use the icon in the exe
            PrefsForm.Add("Picture", "x" (textX - 64 - 18) " w64 h64", App.IconFile)
        else if A_IsCompiled
            PrefsForm.Add("Picture", "x" (textX - 64 - 18) " w64 h64 Icon1", A_ScriptFullPath)
        PreferencesWindow.Gui.SetFont("s15 bold")
        PrefsForm.Add("Text", "x" textX " w" textW, App.Name)
        PreferencesWindow.Gui.SetFont("s9 norm")
        ; tagline, version (what a problem report asks for, on its own line), copyright and license: the same font size, the last two gray, a little more line spacing
        PrefsForm._y := top + 33
        PrefsForm.Add("Text", "x" textX " w" textW, I18n.T("App.Tagline"))
        PrefsForm._y := top + 56
        PrefsForm.Add("Text", "x" textX " w" textW " cGray", I18n.T("Prefs.Version", App.Version) App.AdminSuffix())
        PrefsForm._y := top + 78
        PrefsForm.Add("Text", "x" textX " w" textW " cGray", "© 2013–" SubStr(App.Version, 1, 4) " zhugecaomao   ·   GPL-3.0")            ; the version starts with the release year
        PrefsForm._y := top + 105
        links := ""
        for link in [[App.Website, I18n.T("Prefs.Homepage")], [App.RepoUrl, "GitHub"], [App.RepoUrl "/releases", I18n.T("Update.ReleaseNotes")]
                   , [App.RepoUrl "/issues/new/choose", I18n.T("Prefs.ReportIssue")]]
            links .= (links = "" ? "" : "   ·   ") '<a href="' link[1] '">' link[2] '</a>'
        PrefsForm.Add("Link", "x" textX " w" textW, links)
        PrefsForm._y := top + 134
        PrefsForm.Add("Button", "x" textX " w" PrefsForm.ButtonW " h" PrefsForm.ButtonH, I18n.T("Tray.CheckUpdate"))
            .OnEvent("Click", (*) => UpdateChecker.Check())
        PrefsForm._y := top + 134 + PrefsForm.ButtonH + 22

        PrefsForm.Section("Prefs.Section.Data")
        ; The settings file is ALTRun.json in the data folder, so no extra line; an ellipsis in the middle when the path is too long (SS_PATHELLIPSIS), the full path on hover
        folder := PrefsForm.Info("Prefs.Group.DataFolder", AppSettings.DataDir, " r1 cGray 0x8100")
        Win.AddTooltip(folder.Hwnd, AppSettings.DataDir)
        PrefsForm._y += 4
        PrefsForm.Buttons(""
            , ["Prefs.EditJson", (*) => (PreferencesWindow.Cancel() || App.EditSettingsFile())]
            , ["Prefs.OpenDataFolder", (*) => PrefsPages._OpenFolder(AppSettings.DataDir)]
            , ["Prefs.ChangeDataFolder", (*) => PrefsPages._ChangeDataFolder()])

        ; What a problem report needs: the debug log (moved here from the General page, next to opening the log) and system info
        PrefsForm.Section("Prefs.Group.Diagnostics")
        PrefsForm.Check("General.SaveLog", "Prefs.SaveLog")
        PrefsForm._y += 4
        PrefsForm.Buttons(""
            , ["Prefs.OpenLog", (*) => PrefsPages._OpenLog()]
            , ["Prefs.CopySystemInfo", (*) => (A_Clipboard := App.SystemInfo(), App.Notify(I18n.T("Prefs.SystemInfoCopied")))])

        PrefsForm.Section("Prefs.Section.Reset")
        PrefsForm.Buttons(""
            , ["Prefs.ResetLearning", (*) => PrefsPages._ResetLearning()]
            , ["Prefs.RestoreDefaults", (*) => PrefsPages._RestoreDefaults()])
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

    static BuildUsage() {
        PrefsForm.BeginPage("Prefs.Page.Usage")
        PrefsForm.Section("Usage.Title")
        w := PrefsForm.ContentW, x := PrefsForm.ContentX
        totals := PrefsForm.Add("Text", "w" w)
        PrefsForm.Below(4, totals)
        since := PrefsForm.Add("Text", "w" w " cGray")
        PrefsForm.Below(12, since)
        top := PrefsForm._y, barW := 14, step := 18, chartH := 100
        bars := []
        Loop 30 {
            PrefsForm._y := top
            bars.Push(PrefsForm.Add("Progress", "x" (x + (A_Index - 1) * step) " w" barW " h" chartH " Vertical c3B82F6 BackgroundE6EBF2"))
        }
        PrefsForm._y := top + chartH + 4
        PreferencesWindow.Gui.SetFont("s8")
        first := PrefsForm.Add("Text", "x" x " w100 cGray")
        chart := PrefsForm.Add("Text", "x" (x + 100) " w" (30 * step - barW - 200 + 10) " Center cGray")
        PrefsForm.Add("Text", "x" (x + 30 * step - barW - 90) " w" (90 + barW - 4) " Right cGray", I18n.T("Usage.Today"))
        PreferencesWindow.Gui.SetFont("s9")
        PrefsForm._y += 22
        headers := []
        for key in ["Feature", "Today", "Week", "Month", "All", "Share"]
            headers.Push(I18n.T("Usage.Col." key))
        list := PrefsForm.Add("ListView", "w" w " h272 -Multi NoSort Grid ReadOnly", headers)
        for index, width in [178, 66, 66, 66, 80, 80]
            list.ModifyCol(index, width (index > 1 ? " Right" : ""))
        PrefsForm.Below(8, list)
        PrefsForm.Button("Usage.Clear", (*) => PrefsPages._ClearUsage())
        PrefsPages._usage := {Totals: totals, Since: since, Bars: bars, First: first, Chart: chart, List: list}
        PrefsPages._RefreshUsage()
    }

    static _RefreshUsage() {
        ui := PrefsPages._usage
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
        PrefsForm.FitColumns(list)
    }

    static _ClearUsage() {
        if (MsgBox(I18n.T("Usage.ClearConfirm"), App.Name, "YesNo Icon? Default2 Owner" PreferencesWindow.Gui.Hwnd) != "Yes")
            return
        Usage.Clear()
        PrefsPages._RefreshUsage()
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

    ; "Rebuild index": rebuild now (in one go), show the result, OK closes it
    static _RebuildApps(button) {
        button.Enabled := false
        count := ApplicationProvider.Rebuild()
        button.Enabled := true
        PrefsPages._RebuildResult(I18n.T("Prefs.RebuildIndex.Done", count), I18n.T("Prefs.RebuildIndex"))
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
            PrefsPages._RebuildResult(message, label)
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
}
