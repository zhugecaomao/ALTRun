;===============================================================================
; HotkeyEditor.ahk - the edit window of a custom hotkey (Preferences → Hotkeys → Add / Edit) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Top to bottom in the order people decide things, each group with a gray line saying what it is for:
;   Hotkey      the recording box (HotkeyBox, mouse buttons too) and "Keep the key's original function" (a leading ~)
;   Action      the command's name in a read-only box and "Choose..." (CommandPicker: a searchable list of ALTRun's own two, the
;               custom commands and the system commands); below it in gray what it is and how it gets the selection (in red when the
;               custom command was deleted)
;   Search text only for Search: what the search box starts with; {selection} = the text selected in the program in front.
;               "Insert" (a drop-down button) adds {selection} or a search scope (folder, doc, recent...), so nothing has to be remembered
;   Works in    all programs, or only in windows matching a WinTitle (ALTRun = only in the search window)
; A line at the bottom says in words what the hotkey does, updated while editing (Summary), so a wrong setting shows before saving.
; When another hotkey already uses the same keys in the same windows, the gray line under the hotkey box says which one, in red (Clash;
; at the box that causes it, in place of how to record a hotkey; OK still saves, as Apply only warns).
; "Show" (an empty search box) of older versions is shown as Search.
;
; Usage:
;   edited := HotkeyEditor.Edit(ownerGui, item, original, commands, inUse)   -> the edited Map (keys not shown are kept), "" when cancelled
;       original: the hotkey being edited ("" = a new one: the window title says "Add")
;       commands: the custom commands to choose from and name ("" = AppSettings.CustomCommands; Preferences passes the ones being edited)
;       inUse: the hotkeys to check for a clash, PreferencesWindow.AllHotkeys ("" = those of AppSettings.Data); original is left out
;   HotkeyEditor.Normalize(item)     -> a copy with Search / SearchText for the old forms and every key the window shows
;   HotkeyEditor.Summary(item, commands)   -> "Ctrl+Alt+F opens ALTRun and searches "folder <selected text>" in all programs."
;   HotkeyEditor.Clash(item, inUse)        -> the first of inUse the same key press triggers in the same windows ({Key, Label...}), "" when none
;   HotkeyEditor.ClashText(item, inUse)    -> the red line under the hotkey box: "Already used by Hotkeys (Lock computer)", "" when none
;   HotkeyEditor.ActionName(id, commands)  -> the name shown for an Action ("Notepad", "(deleted command) custom:k3x9a2mq")
;   HotkeyEditor.ActionDetail(id, commands) -> {Text, Missing}: the gray line below it
;   HotkeyEditor.ListText(item, commands)  -> the Action column of Preferences → Hotkeys ("Search: folder {selection}", "Run: Notepad", "Lock computer")
;   HotkeyEditor.IsMissing(item, commands) -> true when the hotkey runs a custom command that was deleted (its row is red)
;===============================================================================

class HotkeyEditor {
    static Width := 460                                                     ; the content width: narrower than the general editor (the groups go down), wide enough for the hotkey hint on one line
    static MissingColor := "cC42B1C"                                        ; the line below a deleted custom command (the red of Windows error text)
    static Scopes := ["folder", "file", "doc", "pic", "recent", "g"]       ; the search scopes "Insert" offers (prefixes of the search box)

    static Edit(owner, item, original := "", commands := "", inUse := "") {
        state := {Result: ""}
        ui := HotkeyEditor.Build(owner, item, original, OnOk, commands, inUse)
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

    ; The window, not shown yet (the tests check its layout):
    ; {Gui, Key, Pass, Action, ActionId, Choose, Detail, Commands, InUse, SearchText, Insert, All, Only, WinTitle, SummaryText, KeyHint, Item, SetAction}
    static Build(owner, item, original := "", onOk := "", commands := "", inUse := "") {
        item := HotkeyEditor.Normalize(item)
        w := HotkeyEditor.Width
        g := Gui("+Owner" owner.Hwnd " -MinimizeBox", I18n.T(IsObject(original) ? "HotkeyEditor.TitleEdit" : "HotkeyEditor.TitleAdd"))
        g.SetFont("s9", ThemeManager.FontName())
        g.MarginX := 14, g.MarginY := 12
        ui := {Gui: g, Item: item, ActionId: item["Action"], Commands: HotkeyEditor._Commands(commands), InUse: HotkeyEditor._Others(inUse, original)}

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
        ui.KeyHint := g.AddText("xm y" (keyY + keyH + 4) " w" w " cGray +0x400C", I18n.T("Prefs.HotkeyHint"))   ; below the box (the check box beside it is shorter); one line, a long name ends in "..." (SS_LEFTNOWORDWRAP | SS_ENDELLIPSIS)
        g.SetFont("s9")

        section("HotkeyEditor.Section.Action")
        ui.Action := g.AddEdit("xm y+8 w" (w - 88) " r1 -Multi ReadOnly -TabStop")   ; laid out like the search text and "Insert" below
        ui.Choose := g.AddButton("x+8 yp-1 w80", I18n.T("HotkeyEditor.Choose"))
        ui.Choose.OnEvent("Click", (*) => ChooseCommand())
        g.SetFont("s8")
        ui.Detail := g.AddText("xm y+4 w" w " r2 cGray")                    ; two lines tall: a long target wraps instead of being cut
        g.SetFont("s9")
        ui.SetAction := SetAction

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
        SetAction(ui.ActionId)
        return ui

        SetAction(id) {                                                     ; the command chosen: its name, the line below, the summary
            ui.ActionId := id
            ui.Action.Value := HotkeyEditor.ActionName(id, ui.Commands)
            detail := HotkeyEditor.ActionDetail(id, ui.Commands)
            ui.Detail.SetFont(detail.Missing ? HotkeyEditor.MissingColor : "cGray")
            ui.Detail.Value := detail.Text
            Update()
        }
        ChooseCommand() {
            HotkeyBox.CancelActive()                                        ; a hotkey being recorded stops first
            id := CommandPicker.Choose(g, ui.ActionId, ui.Commands)
            if (id != "")
                SetAction(id)
            ui.Choose.Focus()
        }
        Update(*) {
            search := ui.ActionId = "Search"
            ui.SearchText.Enabled := search, ui.Insert.Enabled := search    ; kept in place, grayed out: the window doesn't jump
            ui.WinTitle.Enabled := ui.Only.Value
            edited := HotkeyEditor.Read(ui)
            ui.SummaryText.Value := HotkeyEditor.Summary(edited, ui.Commands)
            clash := HotkeyEditor.ClashText(edited, ui.InUse)
            text := (clash = "") ? I18n.T("Prefs.HotkeyHint") : clash        ; the clash in red in place of how to record a hotkey
            if (ui.KeyHint.Value != text) {
                ui.KeyHint.SetFont((clash = "") ? "cGray" : HotkeyEditor.MissingColor)
                ui.KeyHint.Value := text
            }
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
        action := ui.ActionId
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

    ; The name shown for an Action: ALTRun's own two, a custom command's name, a system command's name; a deleted custom command
    ; "(deleted command) custom:<Id>", an unknown Id (a newer version's) as it is
    static ActionName(id, commands := "") {
        switch id {
            case "Search":       return I18n.T("HotkeyEditor.Action.Search")
            case "ToggleWindow": return I18n.T("HotkeyEditor.Action.Toggle")
        }
        if (SubStr(id, 1, 7) = "custom:") {
            command := AppSettings.ItemById(HotkeyEditor._Commands(commands), SubStr(id, 8))
            return IsObject(command) ? command["Title"] : I18n.T("HotkeyEditor.Deleted", id)
        }
        command := HotkeyEditor._System(id)
        return IsObject(command) ? command["Title"] : id
    }

    ; The line below the name: what kind of command, and for a custom command its target and how it gets the selection
    ; (the rules of CustomCommandProvider.RunFromHotkey). Missing: the custom command was deleted (shown in red)
    static ActionDetail(id, commands := "") {
        switch id {
            case "Search":       return {Text: I18n.T("HotkeyEditor.Detail.Search"), Missing: false}
            case "ToggleWindow": return {Text: I18n.T("HotkeyEditor.Detail.Toggle"), Missing: false}
        }
        if (SubStr(id, 1, 7) = "custom:") {
            command := AppSettings.ItemById(HotkeyEditor._Commands(commands), SubStr(id, 8))
            if !IsObject(command)
                return {Text: I18n.T("HotkeyEditor.Detail.Missing"), Missing: true}
            target := Trim(RegExReplace((command.Has("Target") ? command["Target"] : "") " " (command.Has("Arguments") ? command["Arguments"] : ""), "\s+", " "))
            if (StrLen(target) > 80)
                target := SubStr(target, 1, 79) "…"
            rule := CustomCommandProvider.TakesQuery(command) ? "HotkeyEditor.Detail.Query"
                : SelectionActions.CommandTypes(command).Length ? "HotkeyEditor.Detail.Types" : "HotkeyEditor.Detail.Plain"
            return {Text: I18n.T("CommandPicker.Kind.Custom") " · " target " · " I18n.T(rule), Missing: false}
        }
        command := HotkeyEditor._System(id)
        return {Text: IsObject(command) ? command["Subtitle"] : I18n.T("HotkeyEditor.Detail.Unknown"), Missing: false}
    }

    ; The Action column of the hotkey list: Search with its text, a custom command as "Run: <name>", everything else by name
    ; (old forms as the edit window shows them, see Normalize)
    static ListText(item, commands := "") {
        normalized := HotkeyEditor.Normalize(item)
        id := normalized["Action"]
        if (id = "Search") {
            text := Trim(normalized["SearchText"])
            return (text = "") ? I18n.T("HotkeyEditor.Action.Search") : I18n.T("Prefs.Hotkey.Search", text)
        }
        if (SubStr(id, 1, 7) = "custom:" && !HotkeyEditor.IsMissing(normalized, commands))
            return I18n.T("Prefs.Hotkey.Run", HotkeyEditor.ActionName(id, commands))
        return HotkeyEditor.ActionName(id, commands)
    }

    static IsMissing(item, commands := "") {
        id := (item is Map && item.Has("Action")) ? item["Action"] "" : ""
        return SubStr(id, 1, 7) = "custom:" && !IsObject(AppSettings.ItemById(HotkeyEditor._Commands(commands), SubStr(id, 8)))
    }

    static _System(id) {
        for command in SystemProvider.Commands()
            if (command["Id"] = id)
                return command
        return ""
    }

    static _Commands(commands) => (commands is Array) ? commands : AppSettings.CustomCommands

    ; The hotkeys in use but the one being edited (inUse "" = those of the saved settings)
    static _Others(inUse, original) {
        others := []
        for each in (inUse is Array) ? inUse : PreferencesWindow.AllHotkeys(AppSettings.Data)
            if !(IsObject(original) && IsObject(each.Entry) && ObjPtr(each.Entry) = ObjPtr(original))
                others.Push(each)
        return others
    }

    ; The first hotkey of inUse that the same key press triggers in the same windows (all programs, or the same WinTitle), "" when none.
    ; A hotkey limited to some windows over one that works everywhere is fine: there the limited one wins
    static Clash(item, inUse) {
        key := item.Has("Key") ? Trim(item["Key"]) : ""
        if (LTrim(key, "~") = "")
            return ""
        parts := PreferencesWindow.HotkeyParts(key)
        winTitle := item.Has("WinTitle") ? Trim(item["WinTitle"]) : ""
        for each in inUse
            if (Trim(each.WinTitle) = winTitle && PreferencesWindow.SameKeyPress(PreferencesWindow.HotkeyParts(each.Key), parts))
                return each
        return ""
    }

    static ClashText(item, inUse) {
        clash := HotkeyEditor.Clash(item, inUse)
        return IsObject(clash) ? I18n.T("HotkeyEditor.Clash", clash.Label) : ""
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
    static Summary(item, commands := "") {
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
            default:                                                        ; a custom command ("(deleted command) custom:<Id>" once deleted) or a system command
                does := I18n.T("HotkeyEditor.Summary.Command", label, HotkeyEditor.ActionName(action, commands))
        }
        winTitle := Trim(item.Has("WinTitle") ? item["WinTitle"] : "")
        where := (winTitle = "") ? I18n.T("HotkeyEditor.Summary.All") : (winTitle = "ALTRun") ? I18n.T("HotkeyEditor.Summary.InALTRun")
            : I18n.T("HotkeyEditor.Summary.Only", winTitle)
        return does " " where
    }
}
