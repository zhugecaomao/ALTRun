;===============================================================================
; HotkeyEditor.ahk - the edit window of a custom hotkey (Preferences → Hotkeys → Add / Edit) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Top to bottom in the order people decide things, each group with a gray line saying what it is for:
;   Hotkey      the recording box (HotkeyBox, mouse buttons too) and "Keep the key's original function" (a leading ~)
;   Action      Search in ALTRun / Show or hide ALTRun first, then the system commands by name
;   Search text only for Search: what the search box starts with; {selection} = the text selected in the program in front.
;               "Insert" (a drop-down button) adds {selection} or a search scope (folder, doc, recent...), so nothing has to be remembered
;   Works in    all programs, or only in windows matching a WinTitle (ALTRun = only in the search window)
; A line at the bottom says in words what the hotkey does, updated while editing (Summary), so a wrong setting shows before saving.
; "Show" (an empty search box) of older versions is shown as Search.
;
; Usage:
;   edited := HotkeyEditor.Edit(ownerGui, item, original)   -> the edited Map (keys not shown are kept), "" when cancelled
;       original: the hotkey being edited ("" = a new one: the window title says "Add")
;   HotkeyEditor.Normalize(item)     -> a copy with Search / SearchText for the old forms and every key the window shows
;   HotkeyEditor.Summary(item)       -> "Ctrl+Alt+F opens ALTRun and searches "folder <selected text>" in all programs."
;===============================================================================

class HotkeyEditor {
    static Width := 460                                                     ; the content width: narrower than the general editor (the groups go down), wide enough for the hotkey hint on one line
    static Scopes := ["folder", "file", "doc", "pic", "recent", "g"]       ; the search scopes "Insert" offers (prefixes of the search box)

    static Edit(owner, item, original := "") {
        state := {Result: ""}
        ui := HotkeyEditor.Build(owner, item, original, OnOk)
        g := ui.Gui
        owner.Opt("+Disabled")
        g.Show()
        WinWaitClose(g.Hwnd)
        HotkeyBox.CancelActive()                                            ; the window was closed while a hotkey was being recorded
        owner.Opt("-Disabled")
        try WinActivate("ahk_id " owner.Hwnd)
        return state.Result

        OnOk(*) {
            g.Opt("+OwnDialogs")
            edited := HotkeyEditor.Read(ui)
            if (LTrim(edited["Key"], "~") = "") {
                MsgBox(I18n.T("Prefs.Required", I18n.T("Prefs.Col.Key")), g.Title, 48)
                return
            }
            if (edited["WinTitle"] = "" && ui.Only.Value) {
                MsgBox(I18n.T("Prefs.Required", I18n.T("HotkeyEditor.Scope.Only")), g.Title, 48)
                ui.WinTitle.Focus()
                return
            }
            state.Result := edited
            g.Destroy()
        }
    }

    ; The window, not shown yet (the tests check its layout): {Gui, Key, Pass, Action, Actions, SearchText, Insert, All, Only, WinTitle, SummaryText, Item}
    static Build(owner, item, original := "", onOk := "") {
        item := HotkeyEditor.Normalize(item)
        w := HotkeyEditor.Width
        g := Gui("+Owner" owner.Hwnd " -MinimizeBox", I18n.T(IsObject(original) ? "HotkeyEditor.TitleEdit" : "HotkeyEditor.TitleAdd"))
        g.SetFont("s9", ThemeManager.FontName())
        g.MarginX := 14, g.MarginY := 12
        ui := {Gui: g, Item: item}

        section(key, first := false) {
            g.SetFont("bold")
            g.AddText("xm " (first ? "" : "y+16 ") "w" w, I18n.T(key))
            g.SetFont("norm")
            g.AddText("xm y+3 w" w " h1 0x10")                              ; SS_ETCHEDHORZ
        }
        hint(text, indent := 0) {
            g.SetFont("s8")
            ctrl := g.AddText("xm+" indent " y+4 w" (w - indent) " cGray", text)
            g.SetFont("s9")
            return ctrl
        }

        section("HotkeyEditor.Section.Key", true)
        ui.Key := HotkeyBox.Add(g, "xm y+8 w220", LTrim(item["Key"], "~"), true, () => Update())
        ui.Pass := g.AddCheckbox("x+14 yp+3", I18n.T("Hotkey.PassThrough"))
        ui.Pass.Value := (SubStr(item["Key"], 1, 1) = "~") ? 1 : 0
        ui.Key.GetPos(, &keyY, , &keyH)
        g.SetFont("s8")
        g.AddText("xm y" (keyY + keyH + 4) " w" w " cGray", I18n.T("Prefs.HotkeyHint"))   ; below the box (the check box beside it is shorter)
        g.SetFont("s9")

        section("HotkeyEditor.Section.Action")
        ui.Actions := HotkeyEditor.Actions(item["Action"])
        labels := []
        for choice in ui.Actions
            labels.Push(choice[2])
        ui.Action := g.AddDropDownList("xm y+8 w" w, labels)
        ui.Action.Value := HotkeyEditor._Index(ui.Actions, item["Action"])
        ui.Action.OnEvent("Change", (*) => Update())

        section("HotkeyEditor.Section.SearchText")
        ui.SearchText := g.AddEdit("xm y+8 w" (w - 88) " r1 -Multi", item["SearchText"])
        ui.SearchText.OnEvent("Change", (*) => Update())
        ui.Insert := g.AddButton("x+8 yp-1 w80", I18n.T("HotkeyEditor.Insert"))
        insertItems := [{Label: I18n.T("HotkeyEditor.Insert.Selection"), Click: (*) => InsertSelection()}, "-"]
        for scope in HotkeyEditor.Scopes
            insertItems.Push({Label: scope "`t" HotkeyEditor.ScopeText(scope), Click: ((scope, *) => SetScope(scope)).Bind(scope)})
        PrefsForm.MenuButton(ui.Insert, insertItems, "")
        hint(I18n.T("HotkeyEditor.SearchHint"))
        hint(I18n.T("HotkeyEditor.SearchExamples"))

        section("HotkeyEditor.Section.Scope")
        ui.All := g.AddRadio("xm y+8 Group", I18n.T("HotkeyEditor.Scope.All"))
        ui.Only := g.AddRadio("xm y+6", I18n.T("HotkeyEditor.Scope.Only"))
        ui.WinTitle := g.AddEdit("xm+18 y+4 w" (w - 18) " r1 -Multi", item["WinTitle"])
        (item["WinTitle"] = "") ? (ui.All.Value := 1) : (ui.Only.Value := 1)
        for ctrl in [ui.All, ui.Only]
            ctrl.OnEvent("Click", (*) => Update())
        ui.WinTitle.OnEvent("Change", (*) => Update())
        hint(I18n.T("HotkeyEditor.Scope.Hint"), 18)

        g.AddText("xm y+16 w" w " h1 0x10")
        ui.SummaryText := g.AddText("xm y+8 w" w " r2", "")                 ; two lines tall: the longest summary fits

        ok := g.AddButton("xm+" (w - 170) " y+14 w80 Default", "OK")
        if IsObject(onOk)
            ok.OnEvent("Click", onOk)
        g.AddButton("x+10 yp w80", I18n.T("Prefs.Cancel")).OnEvent("Click", (*) => g.Destroy())
        g.OnEvent("Escape", (*) => g.Destroy())
        g.OnEvent("Close", (*) => g.Destroy())
        Update()
        return ui

        Update(*) {
            search := ui.Actions[ui.Action.Value][1] = "Search"
            ui.SearchText.Enabled := search, ui.Insert.Enabled := search    ; kept in place, grayed out: the window doesn't jump
            ui.WinTitle.Enabled := ui.Only.Value
            ui.SummaryText.Value := HotkeyEditor.Summary(HotkeyEditor.Read(ui))
        }
        InsertSelection() {                                                 ; {selection} at the cursor (replacing a selected part)
            ui.SearchText.Focus()
            SendMessage(0xC2, true, StrPtr("{selection}"), ui.SearchText)   ; EM_REPLACESEL, undoable
            Update()
        }
        SetScope(scope) {                                                   ; the scope goes first, replacing one already there
            text := ui.SearchText.Value
            for known in HotkeyEditor.Scopes
                if RegExMatch(text, "i)^\Q" known "\E(\s|$)")
                    text := LTrim(SubStr(text, StrLen(known) + 1))
            ui.SearchText.Value := scope " " text
            ui.SearchText.Focus()
            SendMessage(0xB1, -1, -1, ui.SearchText)                        ; EM_SETSEL: the cursor at the end
            Update()
        }
    }

    ; The item as the window shows it now (the keys it doesn't show are kept)
    static Read(ui) {
        edited := ui.Item.Clone()
        key := HotkeyBox.Value(ui.Key)
        key := LTrim(key, "~")
        edited["Key"] := (key != "" && ui.Pass.Value) ? "~" key : key
        action := ui.Actions[ui.Action.Value][1]
        edited["Action"] := action
        edited["SearchText"] := (action = "Search") ? ui.SearchText.Value : ""
        edited["WinTitle"] := ui.Only.Value ? Trim(ui.WinTitle.Value) : ""
        return edited
    }

    static Normalize(item) {
        result := (item is Map) ? item.Clone() : Map()
        for key in ["Key", "Action", "SearchText", "WinTitle"]
            if !result.Has(key)
                result[key] := ""
        if (result["Action"] = "")
            result["Action"] := "Search"
        searchText := App.HotkeySearchText(result)
        action := App.HotkeyActionId(result)
        if SystemProvider.Aliases.Has(action)                               ; names of old versions (Activate -> Show)
            action := SystemProvider.Aliases[action]
        if (action = "Show")                                                ; Show = Search with an empty search box
            action := "Search"
        result["Action"] := action, result["SearchText"] := searchText
        return result
    }

    ; [[Id, label], ...]: ALTRun's own two first, then the system commands by name; an unknown Id (a newer version's) is kept as it is
    static Actions(current := "") {
        choices := [["Search", I18n.T("HotkeyEditor.Action.Search")], ["ToggleWindow", I18n.T("HotkeyEditor.Action.Toggle")]]
        for command in SystemProvider.Commands()
            choices.Push([command["Id"], command["Title"]])
        if (current != "" && !HotkeyEditor._Index(choices, current, 0))
            choices.Push([current, current])
        return choices
    }

    static ScopeText(scope) {
        switch scope {
            case "folder": return I18n.T("HotkeyEditor.Scope.folder")
            case "file":   return I18n.T("HotkeyEditor.Scope.file")
            case "doc":    return I18n.T("HotkeyEditor.Scope.doc")
            case "pic":    return I18n.T("HotkeyEditor.Scope.pic")
            case "recent": return I18n.T("HotkeyEditor.Scope.recent")
            case "g":      return I18n.T("HotkeyEditor.Scope.g")
        }
        return ""
    }

    ; What the hotkey does, in words: "Ctrl+Alt+F opens ALTRun and searches "folder <selected text>" in all programs."
    static Summary(item) {
        key := LTrim(item.Has("Key") ? item["Key"] : "", "~")
        if (key = "")
            return I18n.T("HotkeyEditor.Summary.NoKey")
        label := HotkeyBox.Label(key)
        action := item.Has("Action") ? item["Action"] : ""
        switch action {
            case "Search":
                text := Trim(item.Has("SearchText") ? item["SearchText"] : "")
                text := RegExReplace(text, "i)\{selection\}", I18n.T("HotkeyEditor.Summary.Selection"))
                does := (text = "") ? I18n.T("HotkeyEditor.Summary.SearchEmpty", label) : I18n.T("HotkeyEditor.Summary.Search", label, text)
            case "ToggleWindow":
                does := I18n.T("HotkeyEditor.Summary.Toggle", label)
            default:
                name := action
                for command in SystemProvider.Commands()
                    if (command["Id"] = action)
                        name := command["Title"]
                does := I18n.T("HotkeyEditor.Summary.Command", label, name)
        }
        winTitle := Trim(item.Has("WinTitle") ? item["WinTitle"] : "")
        where := (winTitle = "") ? I18n.T("HotkeyEditor.Summary.All") : (winTitle = "ALTRun") ? I18n.T("HotkeyEditor.Summary.InALTRun")
            : I18n.T("HotkeyEditor.Summary.Only", winTitle)
        return does " " where
    }

    static _Index(choices, id, fallback := 1) {
        for index, choice in choices
            if (choice[1] = id)
                return index
        return fallback
    }
}
