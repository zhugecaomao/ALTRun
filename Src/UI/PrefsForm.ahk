;===============================================================================
; PrefsForm.ahk - the form helpers and layout constants of the Preferences pages (AutoHotkey v2)
;-------------------------------------------------------------------------------
; A page is a two-column form: right-aligned labels on the left (LabelW wide), every control starting at InputX(),
; and a gray description below a setting when I18n has "<label key>.Desc". Each setting takes one line, e.g.
;   PrefsForm.BeginPage("Prefs.Page.General", 120)
;   PrefsForm.Check("General.LaunchAtLogin", "Prefs.LaunchAtLogin", , , "Prefs.Group.Startup")
;   PrefsForm.Field("Features.Snippets.Keyword", "Prefs.SnippetKeyword", "M")      widths "S" / "M" / "K" / "L"
;   PrefsForm.Pair([...], [...]) / Choice(...) / LinesButton(...) / List(...) / Button(...)
; Paths are keys in ALTRun.json joined with "."; values are read from and written back to PreferencesWindow.Working.
; The pages themselves are in PrefsPages; the window, page list and buttons in PreferencesWindow.
;===============================================================================

class PrefsForm {
    ;---------------------------------------------------------------------------
    ; Layout (pixels at 96 DPI; the window scales them)
    ;---------------------------------------------------------------------------
    ; Two-column form (as in Alfred / Listary): right-aligned labels on the left, and a page's controls start at one vertical line (InputX);
    ; the left column's width is set per page by its labels (BeginPage); input boxes have only three widths, check boxes get group labels in the left column,
    ; buttons share one size, and gray descriptions align with the controls
    static ContentX := 190, ContentW := 560, LabelW := 190
    static WidthS := 70, WidthM := 220, ButtonW := 200, ButtonH := 26       ; WidthL = the whole control column (InputW)
    static WidthK := 150                                                    ; hotkey box ("K")
    static ButtonY := 619                                                   ; where the bottom buttons are; page content must stay above (y < ButtonY - 10)
    static LabelMax := 250                                                  ; widest the left column gets; longer labels wrap

    ; State of the page being built
    static _current := "", _y := 0                                          ; the page and the y of its next row
    static _pairX := 0, _pairFirstW := 0, _pairSecondW := 0, _pairSizes := ["", ""]
    static _pageKey := "", Wider := Map()                                   ; the page being built; pages whose labels don't fit -> the left column width they need
    static LabelWidths := Map()                                             ; "language page" -> the widened left column (remembered once opened)

    ;---------------------------------------------------------------------------
    ; Rows
    ;---------------------------------------------------------------------------
    ; labelW: width of this page's left-column labels (set by its longest label, so the content isn't pushed right as a whole)
    static BeginPage(nameKey, labelW := 190) {
        page := PreferencesWindow.Pages[PreferencesWindow._PageIndex(nameKey)]
        page.Controls := [], page.Built := true
        PrefsForm._current := page
        PrefsForm._y := 14
        PrefsForm._pageKey := nameKey
        PrefsForm._pairX := 0                                               ; right-column position of this page's paired settings, see Pair
        PrefsForm.LabelW := Max(labelW, PrefsForm.LabelWidths.Get(I18n.Lang " " nameKey, 0))
    }

    ; when a left-column label doesn't fit (it would wrap), note the width this page needs, see Show
    static FitLabel(ctrl, text) {
        need := Ceil(Win.TextExtent(ctrl.Hwnd, text) * 96 / A_ScreenDPI) + 12 + 2
        if (need > PrefsForm.LabelW && PrefsForm.LabelW < PrefsForm.LabelMax) {
            key := PrefsForm._pageKey
            PrefsForm.Wider[key] := Min(PrefsForm.LabelMax, Max(need, PrefsForm.Wider.Get(key, 0)))
        }
    }

    static Add(type, options, text := "") {
        pos := InStr(options, " x") || InStr(options, "x") = 1 ? "" : "x" PrefsForm.ContentX " "
        ctrl := PreferencesWindow.Gui.Add(type, pos "y" PrefsForm._y " " options " Hidden", text)
        PrefsForm._current.Controls.Push(ctrl)
        switch type, false {                                                ; a user change -> "Apply" becomes available (type names are case-insensitive)
            case "Edit", "DropDownList", "ComboBox": ctrl.OnEvent("Change", (*) => PreferencesWindow.MarkDirty())
            case "CheckBox", "Radio":                ctrl.OnEvent("Click", (*) => PreferencesWindow.MarkDirty())
        }
        return ctrl
    }

    static Bind(path, readFn) {
        PreferencesWindow.Binds.Push({Path: path, Read: readFn})
    }

    ; Moves _y below the lowest of these controls (their real height, so wrapped text never overlaps)
    static Below(gap, controls*) {
        bottom := PrefsForm._y
        for ctrl in controls {
            ctrl.GetPos(, &top, , &height)
            bottom := Max(bottom, top + height)
        }
        PrefsForm._y := bottom + gap
    }

    ; Section title: bold + a thin line below; a little space above when it isn't the page's first row
    static Section(labelKey) {
        if (PrefsForm._y > 20)
            PrefsForm._y += 8
        PreferencesWindow.Gui.SetFont("bold")
        ctrl := PrefsForm.Add("Text", "w" PrefsForm.ContentW, I18n.T(labelKey))
        PreferencesWindow.Gui.SetFont("norm")
        PrefsForm.Below(3, ctrl)
        line := PrefsForm.Add("Text", "w" PrefsForm.ContentW " h1 0x10")                    ; SS_ETCHEDHORZ
        PrefsForm.Below(PrefsForm.HasDesc(labelKey) ? 4 : 10, line)
        PrefsForm.Desc(labelKey, PrefsForm.ContentX)
    }

    static HasDesc(labelKey) {
        return I18n.Strings.Has(labelKey ".Desc")
    }

    ; Settings whose description has a Wiki link: {1} in the description becomes the URL of [page, anchor on the Chinese page, anchor on the English page]
    static DescLinks := Map("Prefs.SelectionHotkey", ["Usage", "#选中内容的操作", "#actions-on-selected-content"])

    ; The small gray description below a setting ("<labelKey>.Desc" in I18n); nothing if there is none. x: aligned with the left of the control above
    static Desc(labelKey, x, text := "") {
        controlType := "Text"
        if (text = "") {
            if !PrefsForm.HasDesc(labelKey)
                return
            text := I18n.T(labelKey ".Desc", PrefsForm.DescLink(labelKey))
            if PrefsForm.DescLinks.Has(labelKey)
                controlType := "Link"                                              ; <a href="...">...</a>, opens the web page when clicked
        }
        PreferencesWindow.Gui.SetFont("s8")
        ctrl := PrefsForm.Add(controlType, "x" x " w" (PrefsForm.ContentX + PrefsForm.ContentW - x) " cGray", text)
        PreferencesWindow.Gui.SetFont("s9")
        PrefsForm.Below(10, ctrl)
    }

    static DescLink(labelKey) {
        if !PrefsForm.DescLinks.Has(labelKey)
            return ""
        link := PrefsForm.DescLinks[labelKey]
        return HelpProvider.WikiPage(link[1]) (SubStr(I18n.Lang, 1, 2) = "zh" ? link[2] : link[3])
    }

    ; A gray description with links (<a href="...">...</a>, opens the page when clicked), below a section title that has no description of its own
    static LinkDesc(text, x) {
        PrefsForm._y -= 6                                                   ; like a section's own description, 4 pixels below the title's line
        PreferencesWindow.Gui.SetFont("s8")
        ctrl := PrefsForm.Add("Link", "x" x " w" (PrefsForm.ContentX + PrefsForm.ContentW - x) " cGray", text)
        PreferencesWindow.Gui.SetFont("s9")
        PrefsForm.Below(10, ctrl)
    }

    ; A full-width gray description (below lists)
    static Note(textKey) {
        PrefsForm.Desc("", PrefsForm.ContentX, I18n.T(textKey))
    }

    static Gap() {
        PrefsForm._y += 10
    }

    ; A label and an input box on the same row: the label moves down 3 pixels to align with the box's text
    static Label(labelKey) {
        PrefsForm._y += 3
        ctrl := PrefsForm.Add("Text", "w" (PrefsForm.LabelW - 12) " Right", I18n.T(labelKey))
        PrefsForm._y -= 3
        PrefsForm.FitLabel(ctrl, I18n.T(labelKey))
        return ctrl
    }

    static InputX() {
        return PrefsForm.ContentX + PrefsForm.LabelW
    }

    ; width of the control column (the width of a long "L" input box)
    static InputW() {
        return PrefsForm.ContentW - PrefsForm.LabelW
    }

    ; Input box width: "S" numbers / "M" hotkeys, keywords / "L" paths, filters (the whole control column; leaves room for a browse button if there is one)
    static Width(size, withBrowse := false) {
        switch size {
            case "S": return PrefsForm.WidthS
            case "M": return PrefsForm.WidthM
            case "K": return PrefsForm.WidthK
        }
        return PrefsForm.InputW() - (withBrowse ? 34 : 0)
    }

    ; Check box in the control column; group: the group label in the left column (only on the first check box of a group).
    ; desc: a description shown instead of "<labelKey>.Desc" (e.g. text that changes with a state); width: to the right edge of the page by default
    static Check(path, labelKey, x := 0, desc := "", group := "", width := 0) {
        x := x ? x : PrefsForm.InputX()
        if (group != "")
            PrefsForm.GroupLabel(group)
        ctrl := PrefsForm.Add("Checkbox", "x" x " w" (width ? width : PrefsForm.ContentX + PrefsForm.ContentW - x), I18n.T(labelKey))
        ctrl.Value := PreferencesWindow.GetPath(PreferencesWindow.Working, path) ? 1 : 0
        PrefsForm.Bind(path, () => ctrl.Value)
        hasDesc := (desc != "" || PrefsForm.HasDesc(labelKey))
        PrefsForm.Below(hasDesc ? 0 : 6, ctrl)
        PrefsForm.Desc(labelKey, x + 18, desc)                              ; aligned with the check box's text
        return ctrl
    }

    ; A group label in the left column, on the same row as the first control on the right
    static GroupLabel(labelKey, dy := 1) {
        PrefsForm._y += dy
        ctrl := PrefsForm.Add("Text", "x" PrefsForm.ContentX " w" (PrefsForm.LabelW - 12) " Right", I18n.T(labelKey))
        PrefsForm._y -= dy
        PrefsForm.FitLabel(ctrl, I18n.T(labelKey))
        return ctrl
    }

    ; Left-column label + some text (or a link) on the right, e.g. "Version: 2026.09.26"
    static Info(labelKey, text, options := "", type := "Text") {
        label := PrefsForm.GroupLabel(labelKey)
        ctrl := PrefsForm.Add(type, "x" PrefsForm.InputX() " w" PrefsForm.InputW() options, text)
        PrefsForm.Below(6, label, ctrl)
        return ctrl
    }

    ; A label directly followed by an input box (for pages arranged in sections, without the left column): "Search history entries [30]"
    static InlineField(path, labelKey, size, kind := "text") {
        x := PrefsForm.InputX()
        PrefsForm._y += 3
        label := PrefsForm.Add("Text", "x" x, I18n.T(labelKey))
        PrefsForm._y -= 3
        label.GetPos(, , &labelW)
        value := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        ctrl := PrefsForm.Add("Edit", "x" (x + labelW + 8) " w" PrefsForm.Width(size) " r1 -Multi" (kind = "number" ? " Number" : ""), value)
        if (kind = "number")
            PrefsForm.Bind(path, () => IsInteger(ctrl.Value) ? Integer(ctrl.Value) : 0)
        else
            PrefsForm.Bind(path, () => ctrl.Value)
        PrefsForm.Below(PrefsForm.HasDesc(labelKey) ? 3 : 8, label, ctrl)
        PrefsForm.Desc(labelKey, x)
        return ctrl
    }

    ; Aligns the left edges of several controls (the input boxes of InlineField rows) to the rightmost one
    static AlignLeft(controls*) {
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

    ; kind: text / number / file / folder; size: "S" / "M" / "L" (see Width)
    static Field(path, labelKey, size, kind := "text") {
        label := PrefsForm.Label(labelKey)
        if (kind = "hotkey")
            return PrefsForm.BelowInput(labelKey, label, PrefsForm.HotkeyField(PrefsForm.InputX(), path))
        value := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        browseKind := (kind = "file" || kind = "folder")
        ctrl := PrefsForm.Add("Edit", "x" PrefsForm.InputX() " w" PrefsForm.Width(size, browseKind) " r1 -Multi" (kind = "number" ? " Number" : ""), value)
        if browseKind {
            browse := PrefsForm.Add("Button", "x+4 yp-1 w30", I18n.T("Prefs.Browse"))
            browseFn := ItemEditor._Browser(ctrl, kind, PreferencesWindow.Gui)
            browse.OnEvent("Click", (p*) => (browseFn(p*), PreferencesWindow.MarkDirty()))
        }
        if (kind = "number")
            PrefsForm.Bind(path, () => IsInteger(ctrl.Value) ? Integer(ctrl.Value) : 0)
        else
            PrefsForm.Bind(path, () => ctrl.Value)
        PrefsForm.BelowInput(labelKey, label, ctrl)
    }

    ; After a label + input box row: a description goes right below the box
    static BelowInput(labelKey, controls*) {
        hasDesc := PrefsForm.HasDesc(labelKey)
        PrefsForm.Below(hasDesc ? 3 : 8, controls*)
        PrefsForm.Desc(labelKey, PrefsForm.InputX())
    }

    static Choice(path, labelKey, values, labels, size := "M") {
        label := PrefsForm.Label(labelKey)
        ctrl := PrefsForm.DropDown(PrefsForm.InputX(), path, values, labels, size)
        PrefsForm.BelowInput(labelKey, label, ctrl)
        return ctrl
    }

    ; A drop-down bound to path. The option matching the default value is marked "(default)",
    ; and the box widens (up to the right edge of the page) when an option would be cut off
    static DropDown(x, path, values, labels, size) {
        static defaults := AppSettings.Defaults()
        default := PreferencesWindow.GetPath(defaults, path)
        marked := []
        for index, label in labels
            marked.Push(values[index] = default ? I18n.T("Prefs.DefaultChoice", label) : label)
        ctrl := PrefsForm.Add("DropDownList", "x" x " w" PrefsForm.Width(size), marked)
        width := Min(PrefsForm.ChoiceWidth(ctrl), PrefsForm.ContentX + PrefsForm.ContentW - x)
        if (width > PrefsForm.Width(size))
            ctrl.Move(, , width)
        current := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        ctrl.Value := 1
        for index, value in values
            if (value = current)
                ctrl.Value := index
        PrefsForm.Bind(path, () => values[ctrl.Value])
        return ctrl
    }

    ; Width a drop-down needs to show its longest option in full
    static ChoiceWidth(ctrl) {
        width := 0
        for text in ControlGetItems(ctrl)
            width := Max(width, Ceil(Win.TextExtent(ctrl.Hwnd, text) * 96 / A_ScreenDPI) + 32)   ; plus the arrow button and border
        return width
    }

    ; A list with one item per line, shown as a read-only summary + "Edit...", which opens it in a roomy window of its own
    ; (ItemEditor.EditLines, with the longer "<labelKey>.Help" text): the page stays short and long lines are easy to read.
    ; The setting at path is an array. options (all optional):
    ;   Summary:    (lines) => the text shown on the page (default: the items joined with commas; the box cuts off what doesn't fit)
    ;   AlignPairs: the button goes under the right column of the paired settings above (Pair), not at the right edge
    static LinesButton(path, labelKey, options := "") {
        label := PrefsForm.Label(labelKey)
        value := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        state := {Lines: (value is Array) ? value.Clone() : []}
        summarize := (IsObject(options) && options.HasOwnProp("Summary")) ? options.Summary : (lines) => TextTools.JoinList(lines)
        alignPairs := IsObject(options) && options.HasOwnProp("AlignPairs") && options.AlignPairs
        x := PrefsForm.InputX(), paired := alignPairs && PrefsForm._pairX
        buttonW := paired ? PrefsForm._pairSecondW : PrefsForm.WidthS
        buttonX := (paired ? PrefsForm._pairX + PrefsForm._pairSecondW : x + PrefsForm.InputW()) - buttonW
        box := PrefsForm.Add("Edit", "x" x " w" (buttonX - 8 - x) " r1 -Multi ReadOnly", summarize(state.Lines))
        box.GetPos(, , , &boxH)
        button := PrefsForm.Add("Button", "x" buttonX " yp-1 w" buttonW " h" (boxH + 2), I18n.T("Prefs.EditList"))
        button.OnEvent("Click", (*) => PrefsForm.EditLines(labelKey, state, box, summarize))
        PrefsForm.Bind(path, () => state.Lines)
        PrefsForm.BelowInput(labelKey, label, box, button)
    }

    static EditLines(labelKey, state, box, summarize) {
        help := I18n.Strings.Has(labelKey ".Help") ? I18n.T(labelKey ".Help") : I18n.T(labelKey ".Desc")
        edited := ItemEditor.EditLines(PreferencesWindow.Gui, I18n.T(labelKey), help, state.Lines)
        if (!(edited is Array) || TextTools.Join(edited) == TextTools.Join(state.Lines))
            return
        state.Lines := edited, box.Value := summarize(edited)
        PreferencesWindow.MarkDirty()
    }

    ; numeric: a list of numbers (e.g. bar diameters), stored as numbers rather than text
    static Csv(path, labelKey, size, numeric := false) {
        label := PrefsForm.Label(labelKey)
        value := TextTools.JoinList(PreferencesWindow.GetPath(PreferencesWindow.Working, path))
        ctrl := PrefsForm.Add("Edit", "x" PrefsForm.InputX() " w" PrefsForm.Width(size) " r1 -Multi", value)
        PrefsForm.Bind(path, () => TextTools.SplitList(ctrl.Value, numeric))
        PrefsForm.BelowInput(labelKey, label, ctrl)
    }

    ; Button in the control column, one size for all; group: the group label in the left column
    static Button(labelKey, fn, group := "") {
        if (group != "")
            PrefsForm.GroupLabel(group, 5)                                  ; aligned with the button's text
        ctrl := PrefsForm.Add("Button", "x" PrefsForm.InputX() " w" PrefsForm.ButtonW " h" PrefsForm.ButtonH, I18n.T(labelKey))
        ctrl.OnEvent("Click", fn)
        PrefsForm.Below(PrefsForm.HasDesc(labelKey) ? 2 : 6, ctrl)
        PrefsForm.Desc(labelKey, PrefsForm.InputX())
        return ctrl
    }

    ; Several buttons side by side: specs = [labelKey, fn], ...; the description is the first button's "<labelKey>.Desc"
    static Buttons(group, specs*) {
        if (group != "")
            PrefsForm.GroupLabel(group, 5)
        x := PrefsForm.InputX(), controls := []
        buttonW := Min(PrefsForm.ButtonW, (PrefsForm.InputW() - 10 * (specs.Length - 1)) // specs.Length)                    ; narrower buttons when they don't fit after the left column was widened
        for spec in specs {
            ctrl := PrefsForm.Add("Button", "x" x " w" buttonW " h" PrefsForm.ButtonH, I18n.T(spec[1]))
            ctrl.OnEvent("Click", spec[2])
            controls.Push(ctrl)
            x += buttonW + 10
        }
        PrefsForm.Below(PrefsForm.HasDesc(specs[1][1]) ? 2 : 6, controls*)
        PrefsForm.Desc(specs[1][1], PrefsForm.InputX())
    }

    ; Two settings on one row: the left one's label in the left column, the right one's label right after the left input box.
    ; left / right = [path, labelKey, size, kind := "text", values, labels] (kind: text / number / csv / choice)
    ; the description is the left one's "<labelKey>.Desc"
    static Pair(left, right) {
        label := PrefsForm.Label(left[2])
        first := PrefsForm.Input(PrefsForm.InputX(), left*)
        first.GetPos(&x, , &w)
        sameAsFirstRow := PrefsForm._pairX && left[3] = PrefsForm._pairSizes[1] && right[3] = PrefsForm._pairSizes[2]
        if sameAsFirstRow {                                                 ; the same kind of input box as the first row of the page: the same width, right side aligned to the same column
            w := PrefsForm._pairFirstW
            first.Move(, , w)
        }
        PrefsForm._y += 3
        label2 := PrefsForm.Add("Text", "x" (x + w + 18), I18n.T(right[2]))
        PrefsForm._y -= 3
        label2.GetPos(&x2, , &w2)
        ; With a widened left column the control column is narrower: if it doesn't fit, shorten the left input box (down to "S"); with the same kind of box on both sides shorten both, keeping them equal
        secondW := sameAsFirstRow ? PrefsForm._pairSecondW : 0
        overflow := x2 + w2 + 8 + (secondW ? secondW : PrefsForm.Width(right[3])) - (PrefsForm.ContentX + PrefsForm.ContentW)
        if (!sameAsFirstRow && overflow > 0 && left[3] = right[3] && w > PrefsForm.WidthS) {
            w := Max(PrefsForm.WidthS, (PrefsForm.ContentX + PrefsForm.ContentW - x - 18 - w2 - 8) // 2)
            first.Move(, , w)
            label2.Move(x + w + 18)
            label2.GetPos(&x2)
            secondW := w, overflow := x2 + w2 + 8 + w - (PrefsForm.ContentX + PrefsForm.ContentW)
        }
        if (overflow > 0 && w > PrefsForm.WidthS) {
            shrink := Min(overflow, w - PrefsForm.WidthS)
            w -= shrink, overflow -= shrink
            first.Move(, , w)
            label2.Move(x + w + 18)
            label2.GetPos(&x2)
        }
        if (overflow > 0) {                                                 ; still doesn't fit: the right one moves to the next row
            first.GetPos(, &firstY, , &firstH)
            PrefsForm._y := firstY + firstH + 6
            label2.Move(PrefsForm.InputX(), PrefsForm._y + 3)
            label2.GetPos(&x2)
        }
        secondX := x2 + w2 + 8
        ; Paired settings on several rows of a page, with the left boxes as wide or narrower: the right boxes align with the first row's right box (labels right-aligned against the boxes)
        if (overflow <= 0) {
            target := PrefsForm._pairX
            if !target
                PrefsForm._pairX := secondX, PrefsForm._pairFirstW := w, PrefsForm._pairSizes := [left[3], right[3]]
                , PrefsForm._pairSecondW := secondW ? secondW : PrefsForm.Width(right[3])
            else if (w <= PrefsForm._pairFirstW && target >= secondX && target + (secondW ? secondW : PrefsForm.Width(right[3])) <= PrefsForm.ContentX + PrefsForm.ContentW) {
                label2.Move(target - 8 - w2)
                secondX := target
            }
        }
        second := PrefsForm.Input(secondX, right*)
        if secondW
            second.Move(, , secondW)
        PrefsForm.BelowInput(left[2], label, first, label2, second)
        return {Label2: label2, Second: second}
    }

    ; Rows from Pair whose right boxes should start at the same x even when a later row's right label is wider:
    ; move the right label and box of the other rows to the rightmost one
    static AlignPairRows(rows*) {
        target := 0
        for row in rows
            row.Second.GetPos(&x), target := Max(target, x)
        for row in rows {
            row.Second.GetPos(&x), row.Label2.GetPos(&labelX)
            row.Second.Move(target), row.Label2.Move(labelX + target - x)
        }
    }

    ; A single input control (without a label) bound to path; kind as in Pair
    static Input(x, path, labelKey, size, kind := "text", values := "", labels := "") {
        current := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        if (kind = "hotkey")
            return PrefsForm.HotkeyField(x, path)
        if (kind = "choice")
            return PrefsForm.DropDown(x, path, values, labels, size)
        text := (kind = "csv") ? TextTools.JoinList(current) : current
        ctrl := PrefsForm.Add("Edit", "x" x " w" PrefsForm.Width(size) " r1 -Multi" (kind = "number" ? " Number" : ""), text)
        switch kind {
            case "number": PrefsForm.Bind(path, () => IsInteger(ctrl.Value) ? Integer(ctrl.Value) : 0)
            case "csv":    PrefsForm.Bind(path, () => TextTools.SplitList(ctrl.Value))
            default:       PrefsForm.Bind(path, () => ctrl.Value)
        }
        return ctrl
    }

    ; The hotkey recording box: not through Add (the program changing its text isn't a change); marked as changed only when a new hotkey is recorded
    static HotkeyField(x, path) {
        ctrl := HotkeyBox.Add(PreferencesWindow.Gui, "x" x " y" PrefsForm._y " w" PrefsForm.WidthK " Hidden"
            , PreferencesWindow.GetPath(PreferencesWindow.Working, path), false, () => PreferencesWindow.MarkDirty())
        PrefsForm._current.Controls.Push(ctrl)
        PrefsForm.Bind(path, () => HotkeyBox.Value(ctrl))
        return ctrl
    }

    ; A button on the same row as the short input box ("S") above (e.g. "Subfolder depth [4]  [Rebuild file index]")
    ; "label [input box]" right of a check box on the same row: the label starts at x, the box follows right after, centered vertically with the check box
    static SideField(rowY, x, path, labelKey, size := "S") {
        label := PrefsForm.Add("Text", "x" x " y" (rowY + 1), I18n.T(labelKey))
        label.GetPos(&labelX, , &labelW)
        ctrl := PrefsForm.Input(labelX + labelW + 8, path, labelKey, size)
        ctrl.Move(, rowY - 3)
        return ctrl
    }

    ; A number box (width "S") followed by a button (e.g. "Subfolder depth" + "Rebuild index"), with the gray description (of labelKey) below both
    static FieldButton(path, labelKey, buttonKey, fn) {
        label := PrefsForm.Label(labelKey)
        rowY := PrefsForm._y
        ctrl := PrefsForm.Add("Edit", "x" PrefsForm.InputX() " w" PrefsForm.WidthS " r1 -Multi Number", PreferencesWindow.GetPath(PreferencesWindow.Working, path))
        PrefsForm.Bind(path, () => IsInteger(ctrl.Value) ? Integer(ctrl.Value) : 0)
        button := PrefsForm.Add("Button", "x" (PrefsForm.InputX() + PrefsForm.WidthS + 12) " y" (rowY - 2) " w" PrefsForm.ButtonW " h" PrefsForm.ButtonH, I18n.T(buttonKey))
        button.OnEvent("Click", fn)
        PrefsForm.BelowInput(labelKey, label, ctrl, button)
        return ctrl
    }

    ; List page: a ListView shows the array at path; Add / Edit / Delete use ItemEditor
    ; check: optional, (item, roots) => "OK" / "Missing" / "Unavailable" / "Skipped".
    ; When given there is a "Check paths" button, with the results in a "Status" column (the items themselves have no such key)
    ; The filter box right of the buttons: only items containing all the typed words (searched in every column's text, case-insensitive)
    static List(path, height, columns, fields, newItem, check := "", buttons := "") {
        items := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        headers := []
        for column in columns
            headers.Push(I18n.T(column[1]))
        listView := PrefsForm.Add("ListView", "w" PrefsForm.ContentW " h" (height - 40) " -Multi NoSort Grid", headers)
        for index, column in columns
            listView.ModifyCol(index, column[3])
        pageName := PrefsForm._current.Name
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
                    cells.Push((column[2] = "Status") ? StatusText(item) : PrefsForm.Cell(item, column[2]))
                if !PrefsForm.FilterMatch(filter, cells)
                    continue
                listView.Add("", cells*)
                shown.Push(index)
            }
            if (selectRow && shown.Length)
                listView.Modify(Min(selectRow, shown.Length), "Select Focus Vis")
            listView.Opt("+Redraw")
            PrefsForm.FitColumns(listView)
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
        PrefsForm._y += height - 34
        PrefsForm.Add("Button", "w80", I18n.T("Prefs.Add")).OnEvent("Click", AddItem)
        PrefsForm.Add("Button", "x+8 yp w80", I18n.T("Prefs.Edit")).OnEvent("Click", EditItem)
        PrefsForm.Add("Button", "x+8 yp w80", I18n.T("Prefs.Delete")).OnEvent("Click", DeleteItem)
        if IsObject(check)
            PrefsForm.Add("Button", "x+24 yp w100", I18n.T("Prefs.CheckTargets")).OnEvent("Click", CheckItems)
        if IsObject(buttons)                                                ; the page's own buttons: [{Label, Width, Click: (list, button) => ...}]
            for extra in buttons
                PrefsForm.Add("Button", "x+24 yp w" extra.Width, extra.Label).OnEvent("Click", ((extra, button, *) => extra.Click.Call(list, button)).Bind(extra))
        listView.GetPos(&listX, , &listW)
        filterBox := PreferencesWindow.Gui.Add("Edit", "x" (listX + listW - 160) " yp+1 w160 r1 -Multi Hidden")   ; not Add: filtering doesn't change the settings
        PrefsForm._current.Controls.Push(filterBox)
        Win.SetCueBanner(filterBox.Hwnd, I18n.T("Prefs.Filter"))
        filterBox.OnEvent("Change", (*) => Refresh())
        PrefsForm._y += 38
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

    static Cell(item, key) {
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

    ; Drop-downs in one column get one width: wide enough for the longest option of any of them, at least "M"
    ; (a button in the column just takes the same width)
    static SameWidth(ctrls*) {
        width := PrefsForm.WidthM
        for ctrl in ctrls
            if (ctrl.Type = "DDL")
                width := Max(width, PrefsForm.ChoiceWidth(ctrl))
        width := Min(width, PrefsForm.InputW())
        for ctrl in ctrls
            ctrl.Move(, , width)
    }

    ; The table's last column fills the remaining width (like AutoHdr) but first leaves room for the vertical scroll bar: Windows doesn't draw it for a list not yet shown,
    ; so AutoHdr counts without it and a horizontal scroll bar appears once shown. Call again when the list contents change (add, delete, filter)
    static FitColumns(listView) {
        static LVM_GETCOLUMNWIDTH := 0x101D, LVM_SETCOLUMNWIDTH := 0x101E, LVM_GETHEADER := 0x101F, HDM_GETITEMCOUNT := 0x1200
        columns := SendMessage(HDM_GETITEMCOUNT, 0, 0, SendMessage(LVM_GETHEADER, 0, 0, listView.Hwnd))
        width := PrefsForm.ListWidth(listView)
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
}
