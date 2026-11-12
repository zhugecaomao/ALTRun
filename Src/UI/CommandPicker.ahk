;===============================================================================
; CommandPicker.ahk - "Choose Command": the searchable list a hotkey's command is picked from (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Opened by "Choose..." in the hotkey edit window (HotkeyEditor). One list of everything a hotkey can run, in three kinds:
;   ALTRun           Search in ALTRun, Show / hide ALTRun (no prefix in the hotkey's Action)
;   Custom command   the user's custom commands ("custom:<Id>")
;   System command   SystemProvider's commands (no prefix)
; Typing filters it with FuzzyMatcher (names, keywords, English names of system commands, pinyin initials), keeping the kinds in
; that order and the best matches first within a kind. ↑ ↓ Page Up / Down in the search box move in the list; Enter or a double-click chooses.
; A drop-down can't do this: there are 70+ system commands and the user's commands can be hundreds.
;
; Usage:
;   id := CommandPicker.Choose(ownerGui, currentId, commands)   -> the chosen Action ("Search", "Lock", "custom:<Id>"), "" when cancelled
;       commands: the custom commands to list (in Preferences the ones being edited, so a command added there can be chosen);
;       those without an Id get one first (AppSettings.EnsureIds), so the hotkey can refer to them
;   CommandPicker.Choices(commands)           -> [{Id, Title, Kind, Keyword, Keys}] in the order shown
;   CommandPicker.Filter(choices, text)       -> the choices matching text, in the order shown
;===============================================================================

class CommandPicker {
    static Width := 460, ListRows := 14

    static Choose(owner, current, commands) {
        if (commands is Array)
            AppSettings.EnsureIds(Map("CustomCommands", commands))
        state := {Result: ""}
        ui := CommandPicker.Build(owner, current, commands, OnOk)
        g := ui.Gui
        owner.Opt("+Disabled")
        g.Show()
        ui.Search.Focus()
        WinWaitClose(g.Hwnd)
        ui.Release.Call()
        owner.Opt("-Disabled")
        try WinActivate("ahk_id " owner.Hwnd)
        return state.Result

        OnOk(*) {
            if (id := CommandPicker.Selected(ui)) != "" {
                state.Result := id
                g.Destroy()
            }
        }
    }

    ; The window, not shown yet (the tests check it): {Gui, Search, List, Shown, Choices, Release}; Release.Call() after it is closed
    static Build(owner, current, commands, onOk := "") {
        w := CommandPicker.Width
        g := Gui("+Owner" owner.Hwnd " -MinimizeBox", I18n.T("CommandPicker.Title"))
        g.SetFont("s9", ThemeManager.FontName())
        g.MarginX := 14, g.MarginY := 12
        ui := {Gui: g, Choices: CommandPicker.Choices(commands), Shown: []}
        ui.Search := g.AddEdit("xm w" w " r1 -Multi")
        Win.SetCueBanner(ui.Search.Hwnd, I18n.T("CommandPicker.Filter"))
        ui.List := g.AddListView("xm y+8 w" w " r" CommandPicker.ListRows " -Multi NoSortHdr LV0x10000"   ; LVS_EX_DOUBLEBUFFER: no flicker while typing
            , [I18n.T("Prefs.Col.Name"), I18n.T("CommandPicker.Col.Kind"), I18n.T("Prefs.Col.Keyword")])
        ui.List.ModifyCol(1, 250), ui.List.ModifyCol(2, 120), ui.List.ModifyCol(3, w - 250 - 120 - 4)
        ui.Search.OnEvent("Change", (*) => Fill())
        ok := g.AddButton("xm+" (w - 170) " y+14 w80 Default", "OK")
        if IsObject(onOk) {
            ok.OnEvent("Click", onOk)
            ui.List.OnEvent("DoubleClick", (ctrl, row) => row ? onOk() : "")
        }
        g.AddButton("x+10 yp w80", I18n.T("Prefs.Cancel")).OnEvent("Click", (*) => g.Destroy())
        g.OnEvent("Escape", (*) => g.Destroy())
        g.OnEvent("Close", (*) => g.Destroy())
        keyHandler := CommandPicker._OnKeyDown.Bind(CommandPicker, ui)        ; ↑ ↓ in the search box move in the list
        OnMessage(0x100, keyHandler)                                        ; WM_KEYDOWN
        ui.Release := () => OnMessage(0x100, keyHandler, 0)                 ; called once the window is closed
        Fill(current)
        return ui

        Fill(select := "") {                                                ; the list for the text typed; select: the Id to select (else the first row)
            ui.Shown := CommandPicker.Filter(ui.Choices, ui.Search.Value)
            ui.List.Opt("-Redraw")
            ui.List.Delete()
            row := 1
            for index, choice in ui.Shown {
                ui.List.Add("", choice.Title, choice.Kind, choice.Keyword)
                if (select != "" && choice.Id = select)
                    row := index
            }
            if ui.Shown.Length
                ui.List.Modify(row, "Select Focus Vis")
            ui.List.Opt("+Redraw")
        }
    }

    ; The Id of the selected row, "" when none
    static Selected(ui) {
        row := ui.List.GetNext()
        return (row && row <= ui.Shown.Length) ? ui.Shown[row].Id : ""
    }

    static Choices(commands) {
        own := I18n.T("CommandPicker.Kind.ALTRun"), custom := I18n.T("CommandPicker.Kind.Custom"), system := I18n.T("CommandPicker.Kind.System")
        choices := []
        add(id, title, kind, keyword := "", english := "") {
            keys := [FuzzyMatcher.Key(title)]
            for text in [keyword, english, Pinyin.Initials(title)]
                if (text != "" && text != title)
                    keys.Push(FuzzyMatcher.Key(text))
            choices.Push({Id: id, Title: title, Kind: kind, Keyword: keyword, Keys: keys})
        }
        add("Search", I18n.T("HotkeyEditor.Action.Search"), own, "", I18n.Strings["HotkeyEditor.Action.Search"])
        add("ToggleWindow", I18n.T("HotkeyEditor.Action.Toggle"), own, "", I18n.Strings["HotkeyEditor.Action.Toggle"])
        if (commands is Array)
            for command in commands
                if (command is Map && command.Has("Id") && command["Id"] != "" && command.Has("Title"))
                    add("custom:" command["Id"], command["Title"], custom, command.Has("Keyword") ? command["Keyword"] : "")
        for command in SystemProvider.Commands()
            add(command["Id"], command["Title"], system, "", command["English"])
        return choices
    }

    ; The choices text matches (fuzzy, on the name, keyword, English name or pinyin initials): the kinds stay in their order,
    ; the better matches first within a kind (the list's order where they are equal)
    static Filter(choices, text) {
        needle := StrLower(Trim(text))
        if (needle = "")
            return choices.Clone()
        result := [], group := [], kind := ""
        flush() {
            sorted := CommandPicker._ByScore(group)
            for entry in sorted
                result.Push(entry.Choice)
            group := []
        }
        for index, choice in choices {
            if (choice.Kind != kind)
                flush(), kind := choice.Kind
            score := FuzzyMatcher.BestKey(needle, choice.Keys)
            if (score > 0)
                group.Push({Choice: choice, Score: score, Index: index})
        }
        flush()
        return result
    }

    static _ByScore(entries) {                                              ; insertion sort, stable: a kind has at most a few hundred
        sorted := []
        for entry in entries {
            position := sorted.Length + 1
            while (position > 1 && sorted[position - 1].Score < entry.Score)
                position -= 1
            sorted.InsertAt(position, entry)
        }
        return sorted
    }

    ; WM_KEYDOWN in the search box: the movement keys go to the list (selecting there), everything else stays in the box
    static _OnKeyDown(ui, wParam, lParam, msg, hwnd) {
        static keys := Map(0x26, "Up", 0x28, "Down", 0x21, "PgUp", 0x22, "PgDn")
        try {
            if (hwnd != ui.Search.Hwnd || !keys.Has(wParam))
                return
        } catch                                                             ; the window is gone
            return
        count := ui.List.GetCount()
        if !count
            return 0
        row := ui.List.GetNext()
        page := CommandPicker.ListRows - 1
        switch keys[wParam] {
            case "Up":   row := Max(row - 1, 1)
            case "Down": row := Min(row + 1, count)
            case "PgUp": row := Max(row - page, 1)
            case "PgDn": row := Min(row + page, count)
        }
        ui.List.Modify(0, "-Select")
        ui.List.Modify(row, "Select Focus Vis")
        return 0
    }
}
