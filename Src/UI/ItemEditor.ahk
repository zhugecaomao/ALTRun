;===============================================================================
; ItemEditor.ahk - the general "edit one record" dialog (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Custom commands / snippets / search engines / custom hotkeys in Preferences share this dialog; each kind of
; record only describes its fields:
;   fields := [
;     {Key: "Title",  Label: "Name", Type: "text", Required: true},
;     {Key: "Type",   Label: "Type", Type: "choice", Choices: [["File", "File"], ["Folder", "Folder"]]},
;     {Key: "Target", Label: "Target", Type: "file"},      ; text box + "..." to choose a file
;       a file field can have Button: {Label, OnClick(controls, g)}: one more button after "..."
;     {Key: "Text",   Label: "Text", Type: "multiline"},
;     {Key: "AutoExpand", Label: "Auto expand", Type: "check"}
;   ]
; Type can be: text / multiline / choice / check / file / folder / number / hotkey
; A file field can have FolderWhen: ["Type", "Folder"] = when another field (Type) is set to Folder, "Browse" chooses a folder instead
; hotkey: a hotkey recording box (HotkeyBox, can record middle / side buttons) + a check box "keep the key's own function" (the ~ in front)
; Hint: gray description below the box; Check: (value, editedMap) => warning text; on OK a warning asks "Save anyway?"
;
; Usage:
;   result := ItemEditor.Edit(ownerGui, "Title", fields, itemMap)
;   returns the edited new Map (keeping keys of itemMap not listed), "" when canceled
;   ItemEditor.WithDefaults(itemMap, defaultsMap)   fills missing keys with the defaults (old settings may lack newer keys)
;   ItemEditor.Field("Title", "Prefs.Col.Title", "text", true)   makes a field description (the label is an I18n key)
;===============================================================================

class ItemEditor {
    static RowGap := 14, HintGap := 4                                       ; space between rows, and between a box and the gray description below it

    ; The lowest bottom of the controls added after first (the box, buttons, check box of one row)
    static _RowBottom(g, first) {
        bottom := 0, after := false
        for ctrl in g {
            if (ctrl.Hwnd = first.Hwnd) {
                after := true
                continue
            }
            if after {
                ctrl.GetPos(, &y, , &h)
                bottom := Max(bottom, y + h)
            }
        }
        return bottom
    }

    static Edit(owner, title, fields, item) {
        state := {Result: ""}                                               ; the nested function OnOk hands the result back through it
        g := Gui("+Owner" owner.Hwnd " -MinimizeBox", title)
        g.SetFont("s9", ThemeManager.FontName())
        g.MarginX := 14, g.MarginY := 12
        controls := Map(), passBoxes := Map()
        labelW := 110, inputW := 540                                        ; wider, so long paths and command-line arguments are visible in full

        ; single-line boxes must have r1: without a row count, an initial text wider than the box (e.g. a long path) makes AHK
        ; turn it multi-line and taller, covering the controls below
        for field in fields {
            g.AddText("xm w" labelW " y+" ItemEditor.RowGap " Section", (field.Type = "check") ? "" : field.Label)   ; a check box's text goes after the box
            first := g.AddText("x+0 yp w0 h0")                              ; note which control the row starts with (the gray description goes below the row's lowest control)
            value := item.Has(field.Key) ? item[field.Key] : ""
            switch field.Type {
                case "multiline":
                    ctrl := g.AddEdit("x+8 ys-3 w" inputW " r8 +Multi +WantTab", value)
                case "check":
                    ctrl := g.AddCheckbox("x+8 ys w" inputW, field.Label)
                    ctrl.Value := value ? 1 : 0
                case "choice":
                    labels := []
                    for choice in field.Choices
                        labels.Push(choice[2])
                    ctrl := g.AddDropDownList("x+8 ys-3 w" inputW, labels)
                    ctrl.Value := Max(1, ItemEditor._ChoiceIndex(field.Choices, value))
                case "file", "folder":
                    extraW := field.HasOwnProp("Button") ? 154 : 0           ; the field's own button (e.g. "Download site icon")
                    ctrl := g.AddEdit("x+8 ys-3 w" (inputW - 34 - extraW) " r1 -Multi", value)
                    browse := g.AddButton("x+4 yp-1 w30", I18n.T("Prefs.Browse"))
                    kind := field.HasOwnProp("FolderWhen") ? ItemEditor._KindGetter(field.FolderWhen, fields, controls) : field.Type
                    browse.OnEvent("Click", ItemEditor._Browser(ctrl, kind, g))
                    if extraW
                        g.AddButton("x+4 yp w" (extraW - 4), field.Button.Label).OnEvent("Click", ItemEditor._FieldButton(field.Button, controls, g))
                case "number":
                    ctrl := g.AddEdit("x+8 ys-3 w100 r1 -Multi Number", value)
                case "hotkey":
                    ctrl := HotkeyBox.Add(g, "x+8 ys-3 w200", LTrim(value, "~"), true)
                    passBoxes[field.Key] := g.AddCheckbox("x+14 yp+3", I18n.T("Hotkey.PassThrough"))
                    passBoxes[field.Key].Value := (SubStr(value, 1, 1) = "~") ? 1 : 0
                default:
                    ctrl := g.AddEdit("x+8 ys-3 w" inputW " r1 -Multi", value)
            }
            if field.HasOwnProp("Hint") {                                   ; small gray description, as in Preferences
                indent := (field.Type = "check") ? 18 : 0                   ; check box: aligned with the text after the box
                g.SetFont("s8")                                             ; below the row's lowest control (a check box after a hotkey box is shorter than the box; placed by it, the description would touch the box)
                g.AddText("xs+" (labelW + 8 + indent) " y" (ItemEditor._RowBottom(g, first) + ItemEditor.HintGap) " w" (inputW - indent) " cGray", field.Hint)
                g.SetFont("s9")
            }
            controls[field.Key] := ctrl
        }

        g.AddButton("xm+" (labelW + inputW - 170) " y+18 w80 Default", "OK").OnEvent("Click", OnOk)
        g.AddButton("x+10 yp w80", I18n.T("Prefs.Cancel")).OnEvent("Click", (*) => g.Destroy())
        g.OnEvent("Escape", (*) => g.Destroy())
        g.OnEvent("Close", (*) => g.Destroy())

        owner.Opt("+Disabled")
        g.Show()
        WinWaitClose(g.Hwnd)
        HotkeyBox.CancelActive()                                            ; the dialog was closed while a hotkey was being recorded
        owner.Opt("-Disabled")
        try WinActivate("ahk_id " owner.Hwnd)
        return state.Result

        OnOk(*) {
            g.Opt("+OwnDialogs")                                            ; the message box covers the edit window, which can't be used until it closes
            edited := item.Clone()
            for spec in fields {
                input := controls[spec.Key]
                switch spec.Type {
                    case "check" : newValue := input.Value
                    case "choice": newValue := spec.Choices[input.Value][1]
                    case "number": newValue := IsInteger(input.Value) ? Integer(input.Value) : 0
                    case "hotkey":
                        newValue := HotkeyBox.Value(input)
                        if (newValue != "" && passBoxes[spec.Key].Value && SubStr(newValue, 1, 1) != "~")
                            newValue := "~" newValue
                        else if !passBoxes[spec.Key].Value
                            newValue := LTrim(newValue, "~")
                    default      : newValue := input.Value
                }
                if (spec.HasOwnProp("Required") && spec.Required && Trim(newValue) = "") {
                    MsgBox(I18n.T("Prefs.Required", spec.Label), title, 48)
                    input.Focus()
                    return
                }
                edited[spec.Key] := newValue
            }
            for spec in fields {                                            ; Check: run after all fields are read, so it can look at other fields (returns warning text = a problem)
                if !spec.HasOwnProp("Check") || (warning := spec.Check.Call(edited[spec.Key], edited)) = ""
                    continue
                if (MsgBox(warning "`n`n" I18n.T("Editor.SaveAnyway"), title, "YesNo Icon! Default2") != "Yes") {
                    controls[spec.Key].Focus()
                    return
                }
            }
            state.Result := edited
            g.Destroy()
        }
    }

    ; A list with one item per line in a roomy window of its own (PrefsForm.LinesButton): help text on top, a big
    ; box that doesn't wrap (long paths stay on one line), OK / Cancel. Returns the lines (blank ones dropped), or "" on Cancel
    static EditLines(owner, title, help, lines) {
        state := {Result: ""}
        g := Gui("+Owner" owner.Hwnd " -MinimizeBox", title)
        g.SetFont("s9", ThemeManager.FontName())
        g.MarginX := 14, g.MarginY := 12
        width := 680
        if (help != "")
            g.AddText("w" width, help)
        box := g.AddEdit("xm y+10 w" width " r16 +Multi -Wrap +HScroll +VScroll", TextTools.Join(lines))
        g.AddButton("xm+" (width - 170) " y+14 w80 Default", "OK").OnEvent("Click", OnOk)
        g.AddButton("x+10 yp w80", I18n.T("Prefs.Cancel")).OnEvent("Click", (*) => g.Destroy())
        g.OnEvent("Escape", (*) => g.Destroy())
        g.OnEvent("Close", (*) => g.Destroy())
        owner.Opt("+Disabled")
        g.Show()
        SendMessage(0xB1, 0, 0, box.Hwnd)                                   ; EM_SETSEL: caret at the start, nothing selected
        WinWaitClose(g.Hwnd)
        owner.Opt("-Disabled")
        try WinActivate("ahk_id " owner.Hwnd)
        return state.Result

        OnOk(*) {
            state.Result := TextTools.SplitLines(box.Value)
            g.Destroy()
        }
    }

    ; The field's own button: OnClick(controls, g), controls = field Key -> control
    static _FieldButton(button, controls, g) {
        return (*) => button.OnClick.Call(controls, g)
    }

    static WithDefaults(item, defaults) {
        filled := item.Clone()
        for key, value in defaults
            if !filled.Has(key)
                filled[key] := value
        return filled
    }

    static Field(key, labelKey, type := "text", required := false, choices := "", hint := "") {
        spec := {Key: key, Label: I18n.T(labelKey), Type: type, Required: required}
        if IsObject(choices)
            spec.Choices := choices
        if (hint != "")
            spec.Hint := hint
        return spec
    }

    ; Owner of the edit dialog: Preferences when it is open, otherwise a hidden window (not on the taskbar)
    static Owner() {
        static hiddenOwner := ""
        if IsObject(PreferencesWindow.Gui)
            return PreferencesWindow.Gui
        if !IsObject(hiddenOwner)
            hiddenOwner := Gui("+ToolWindow -Caption", App.Name)
        return hiddenOwner
    }

    static _ChoiceIndex(choices, value) {
        for index, choice in choices
            if (choice[1] = value)
                return index
        return 0
    }

    static _Browser(ctrl, kind, owner) {
        return (*) => ItemEditor._Browse(ctrl, kind, owner)
    }

    ; Only when "Browse" is clicked is the other field's current choice looked at: when[2] chooses a folder, otherwise a file
    static _KindGetter(when, fields, controls) {
        return () => ItemEditor.KindFor(when, fields, controls)
    }

    static KindFor(when, fields, controls) {
        for spec in fields {
            if (spec.Key = when[1] && controls.Has(spec.Key) && spec.HasOwnProp("Choices")) {
                index := controls[spec.Key].Value
                return (index >= 1 && index <= spec.Choices.Length && spec.Choices[index][1] = when[2]) ? "folder" : "file"
            }
        }
        return "file"
    }

    static _Browse(ctrl, kind, owner) {
        if HasMethod(kind)
            kind := kind()
        owner.Opt("+OwnDialogs")
        current := Path.Resolve(ctrl.Value)
        selected := (kind = "folder") ? DirSelect("*" current, 3) : FileSelect(3, current)
        if (selected != "")
            ctrl.Value := selected
    }
}
