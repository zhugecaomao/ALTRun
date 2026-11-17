; @altrun.title          PT Tools (Rebar / BRC Calculator)
; @altrun.keyword        pttools
; @altrun.description    Rebar and mesh (BRC) areas, the number of bars required and a calculator; results copy to Excel
; @altrun.description_zh 钢筋和钢筋网 (BRC) 面积、所需钢筋根数、计算器; 结果可以复制到 Excel
; @altrun.network        none
;
; Type "pttools" in ALTRun. Changing an input updates the results at once; Enter moves to the next box, F1 opens the help
; (https://github.com/zhugecaomao/ALTRun/wiki/PT-Tools). The inputs and the window position are kept in "PT Tools.ini" next to this
; file. Expressions (Area expression, Calculator) take numbers, + - * / ^ and parentheses only: nothing else is run.
#Requires AutoHotkey v2.0

if (A_LineFile = A_ScriptFullPath)                                          ; run as a program; ALTRun's tests #Include this file instead
    PTToolsWindow.Main()

class PTToolsWindow {
    static Title := "PT Tools"
    static HelpSection := "pt-tools"

    ; Defaults - the keys are the controls' Names (the names registered with the "v" prefix in v2),
    ; so Show() / Save() loop over the same list without a separate mapping table.
    static Defaults := Map(
        "WinLeft", 500, "WinTop", 500,

        ; Rebar area
        "RebarSpanWidth", 3000, "RebarBarDiameter", "20", "RebarBarSpacing", 200,

        ; Bars required
        "RequiredAreaExpression", "1000", "RequiredBarDiameter", "20",
        "SafetyFactorEnabled", 0, "SafetyFactor", 1.0,

        ; Mesh area
        "MeshSpanWidth", 3000, "MeshTopMark", "A7", "MeshBotMark", "A9",

        ; Calculator
        "ExprInput", "100+50"
    )

    ; Window inputs - round-trip through Show()/Save() (the results are calculated, not saved)
    static FieldNames := [
        "RebarSpanWidth", "RebarBarDiameter", "RebarBarSpacing",
        "RequiredAreaExpression", "RequiredBarDiameter", "SafetyFactorEnabled", "SafetyFactor",
        "MeshSpanWidth", "MeshTopMark", "MeshBotMark",
        "ExprInput"
    ]
    ; ComboBox fields (input boxes with common values to pick from) -> their list. .Value is the LIST INDEX when the current text
    ; happens to match a preset item (same as DropDownList), NOT the text - unlike Edit, where .Value is always the text.
    ; Anywhere we read, write or persist these, we have to go through .Text instead of .Value.
    static SpanWidths   := ["3000", "5700", "6000", "7200"]
    static BarDiameters := ["10", "13", "16", "20", "25", "32", "40"]
    static BarSpacings  := ["100", "125", "150", "175", "200", "225", "250", "275", "300", "350", "500", "1000"]
    static RequiredAreas := ["3500", "5000", "7500", "9500", "10000", "12000"]
    static MeshMarks    := ["A6", "A7", "A8", "A9", "A10", "A12", "A13", "B8", "B10", "B12", "D8", "D10", "D12", "E8", "E10", "E12"]
    static SafetyFactors := ["1.1", "1.2", "1.25", "1.5"]
    static Expressions  := ["100+50", "1200*1.09", "(6000/200+1)*201.06", "3.1416*16^2/4", "2^10-24"]   ; examples: + - * / ^ and parentheses
    static ComboFields := Map("RebarSpanWidth", PTToolsWindow.SpanWidths, "RebarBarDiameter", PTToolsWindow.BarDiameters
        , "RebarBarSpacing", PTToolsWindow.BarSpacings, "RequiredAreaExpression", PTToolsWindow.RequiredAreas
        , "RequiredBarDiameter", PTToolsWindow.BarDiameters, "MeshSpanWidth", PTToolsWindow.SpanWidths
        , "MeshTopMark", PTToolsWindow.MeshMarks, "MeshBotMark", PTToolsWindow.MeshMarks
        , "SafetyFactor", PTToolsWindow.SafetyFactors, "ExprInput", PTToolsWindow.Expressions)

    static G := ""                                                        ; the Gui object, while open
    static HotkeysReady := false

    ; Settings: the inputs and the window position, in the .ini file next to this script ([Settings]). Only inputs are saved;
    ; results are calculated again when the window opens. Keys not in Defaults are ignored, a missing or empty one takes its default.
    static IniFile := RegExReplace(A_LineFile, "i)\.ahk$", ".ini")
    static Settings := PTToolsWindow.Defaults.Clone()
    static ExitOnClose := false                                           ; true when this file runs as a program (Main)

    ; Run as a program: the window, or the one already open (in another process) brought to the front
    static Main() {
        SetTitleMatchMode(3)                                                ; the window title exactly
        if WinExist(PTToolsWindow.Title " ahk_class AutoHotkeyGUI") {
            WinActivate()
            ExitApp()
        }
        A_IconHidden := true
        Persistent()
        PTToolsWindow.ExitOnClose := true
        PTToolsWindow.Load()
        PTToolsWindow.Show()
    }

    ; Settings from the .ini file (or the given one), each missing key from Defaults
    static Load(file := PTToolsWindow.IniFile) {
        S := Map()
        for key, value in PTToolsWindow.Defaults {
            saved := IniRead(file, "Settings", key, "")
            S[key] := (saved = "" && value != "") ? value : saved
        }
        PTToolsWindow.Settings := S
        return S
    }

    static WriteIni(file := PTToolsWindow.IniFile) {
        for key in PTToolsWindow.Defaults
            if PTToolsWindow.Settings.Has(key)
                try IniWrite(PTToolsWindow.Settings[key], file, "Settings", key)
    }

    ; F1: this tool's section on the Wiki page, the Chinese page on a Chinese system (the headings are the same in both)
    static HelpUrl(lang := A_Language) => "https://github.com/zhugecaomao/ALTRun/wiki/PT-Tools" (PTToolsWindow.IsChinese(lang) ? "-zh" : "") "#" PTToolsWindow.HelpSection
    static IsChinese(lang := A_Language) => RegExMatch(lang, "i)^(0804|0404|0C04|1004|1404)$") > 0
    static OpenHelp() {
        try Run(PTToolsWindow.HelpUrl())
    }

    ; "" + 0 throws TypeError in v2 (unlike v1, which silently coerced) - and every
    ; Change handler below reads a live Edit/ComboBox that's often briefly empty or
    ; partial while the user is mid-keystroke, so every numeric field read goes
    ; through this instead of a bare "+ 0".
    static Num(val, default := 0) {
        return IsNumber(val) ? val + 0 : default
    }

    ; Adds a "Label:" + Edit control pair on one row (label at x, edit at
    ; editX) and returns the Edit control, so callers can still chain
    ; .OnEvent(...) - shared by every numeric/text field.
    static Field(g, x, editX, y, w, label, name, value, opts := "") {
        if InStr(opts, "ReadOnly")                                          ; Result boxes: Tab / Enter skip them; they can still be selected with the mouse to copy
            opts .= " -Tabstop"
        g.AddText("x" x " y" (y + 3), label)
        return g.AddEdit("x" editX " y" y " w" w " r1 " opts " v" name, value)
    }

    ; Saves, closes the window, and ends the program when this file runs as one
    static Close(GuiObj) {
        PTToolsWindow.Save()
        GuiObj.Destroy()
        PTToolsWindow.G := ""
        if PTToolsWindow.ExitOnClose
            ExitApp()
    }

    ; Formats a computed value to exactly 2 decimal places for display. Round()
    ; alone isn't enough - a result that isn't exactly representable in binary
    ; floating point (most 2-decimal values aren't) can print with a dozen-plus
    ; trailing digits, e.g. Round(9550.54, 2) showing as "9550.5400000000009"
    ; once it round-trips through string concatenation. Non-numeric input (""
    ; for an invalid calc, or an error string from PTCalc.Eval) passes through
    ; unchanged.
    static Fmt2(n) {
        return IsNumber(n) ? Format("{:.2f}", n + 0) : n
    }

    ; ===================================================================
    ; PT Tools window - Rebar area / Bars required / Mesh area / Calculator
    ; ===================================================================
    static Show() {
        if PTToolsWindow.G {
            WinActivate("ahk_id " PTToolsWindow.G.Hwnd)
            return
        }

        S := PTToolsWindow.Settings
        g := Gui("+AlwaysOnTop", PTToolsWindow.Title)
        PTToolsWindow.G := g
        g.SetFont("s10", "Segoe UI")                                        ; the Windows UI font: clearer English and digits
        g.OnEvent("Close", (p*) => PTToolsWindow.Close(p*))

        rebarRight    := PTToolsWindow.BuildRebarAreaGroup(g, S, PTToolsWindow.Margin)
        requiredRight := PTToolsWindow.BuildRequiredQtyGroup(g, S, rebarRight + PTToolsWindow.RebarColGap)
        meshRight     := PTToolsWindow.BuildMeshAreaGroup(g, S, requiredRight + PTToolsWindow.RebarColGap)
        PTToolsWindow.BuildExpressionGroup(g, S, requiredRight)

        ; ComboBox initial text is set last, once every sibling control any of
        ; this window's Change handlers touch already exists - defends against a
        ; programmatic .Text write firing Change synchronously mid-construction
        ; (see the note on ComboFields at the top of the class).
        for name in PTToolsWindow.ComboFields {
            g[name].Text := S[name]
            inner := DllCall("FindWindowEx", "Ptr", g[name].Hwnd, "Ptr", 0, "Str", "Edit", "Ptr", 0, "Ptr")
            if (name != "ExprInput")                                        ; an expression reads from the left
                try ControlSetStyle("+0x2", inner)                          ; the edit box inside is right-aligned too (ES_RIGHT), like the other numbers
            style := InStr(name, "Mark") ? "+0x8"                       ; mesh marks in capitals (ES_UPPERCASE), "a7" -> "A7"
                : (name = "RebarSpanWidth") ? "+0x2000" : ""               ; digits only (ES_NUMBER), as before
            if (style != "")
                try ControlSetStyle(style, inner)
        }

        PTToolsWindow.SetupHotkeys()

        ; the same margin left, right and below the groups (AutoSize would add the GUI's own margins, smaller on the right and below)
        g.Show("x" S["WinLeft"] " y" S["WinTop"] " w" (meshRight + PTToolsWindow.Margin) " h" (PTToolsWindow.TopY + PTToolsWindow.TallGroupH + PTToolsWindow.Margin))

        ; results aren't saved: calculated from the inputs every time the window opens
        PTToolsWindow.UpdateRebarArea()
        PTToolsWindow.OnRequiredExpressionChanged()
        PTToolsWindow.OnMeshFieldChanged()
        PTToolsWindow.OnExpressionChanged()
    }

    static SetupHotkeys() {
        if PTToolsWindow.HotkeysReady
            return
        PTToolsWindow.HotkeysReady := true
        HotIfWinActive(PTToolsWindow.Title)                               ; Enter / numpad Enter = move to next field, same as the old tool
        Hotkey("Enter", (*) => SendInput("{Tab}"))
        Hotkey("NumpadEnter", (*) => SendInput("{Tab}"))
        Hotkey("F1", (*) => PTToolsWindow.OpenHelp())
        HotIfWinActive()
    }

    static Save() {
        g := PTToolsWindow.G
        if !(g)
            return
        S := PTToolsWindow.Settings
        for _, name in PTToolsWindow.FieldNames {
            try S[name] := PTToolsWindow.ComboFields.Has(name) ? g[name].Text : g[name].Value
        }
        try {
            WinGetPos(&x, &y, , , "ahk_id " g.Hwnd)
            S["WinLeft"] := x
            S["WinTop"] := y
        }
        PTToolsWindow.WriteIni()                                            ; written right away
    }

    ; ---------------------------------------------------------------------
    ; Layout constants for the three side-by-side GroupBoxes below. Every
    ; label sits labelGap px to the left of its edit control - wide enough
    ; for the longest label in any of the three groups ("Top mesh + bar (mm²)"
    ; et al) so the text can never visually run into the edit box.
    ; Labels are in sentence case (Windows UI guidelines).
    ; ---------------------------------------------------------------------
    static RebarLabelGap := 145
    static RebarFieldW   := 80
    static RebarColGap   := 10
    static Margin        := 15                                              ; window edge to the groups: left, right and below (the title bar is above)
    static TopY          := 3                                               ; top of the groups
    static TallGroupH    := 320                                             ; Mesh area; Rebar area / Bars required (240) + Calculator (70) end at the same line

    ; A "Label" + ComboBox pair: an input box with common values to pick from (see ComboFields; its text is set in Show)
    static Combo(g, x, editX, y, label, name) {
        g.AddText("x" x " y" (y + 3), label)
        return g.AddComboBox("x" editX " y" y " w" PTToolsWindow.RebarFieldW " v" name, PTToolsWindow.ComboFields[name])
    }

    static BuildRebarAreaGroup(g, S, x0) {
        labelX := x0 + 15
        editX  := labelX + PTToolsWindow.RebarLabelGap
        w := (editX + PTToolsWindow.RebarFieldW + 15) - x0
        y := PTToolsWindow.TopY

        g.Add("GroupBox", "x" x0 " y" PTToolsWindow.TopY " w" w " h240", "Rebar area")
        changed := (p*) => PTToolsWindow.OnRebarFieldChanged(p*)
        PTToolsWindow.Combo(g, labelX, editX, y + 27, "Span width (mm)", "RebarSpanWidth").OnEvent("Change", changed)
        PTToolsWindow.Combo(g, labelX, editX, y + 67, "Bar diameter (mm)", "RebarBarDiameter").OnEvent("Change", changed)
        PTToolsWindow.Combo(g, labelX, editX, y + 107, "Bar spacing (mm)", "RebarBarSpacing").OnEvent("Change", changed)
        PTToolsWindow.Field(g, labelX, editX, y + 147, PTToolsWindow.RebarFieldW, "Rebar area (mm²)", "RebarTotalArea", "", "ReadOnly Right")
        PTToolsWindow.Field(g, labelX, editX, y + 187, PTToolsWindow.RebarFieldW, "Number of bars", "RebarBarCount", "", "ReadOnly Right")

        return x0 + w
    }

    static BuildRequiredQtyGroup(g, S, x0) {
        labelX := x0 + 15
        editX  := labelX + PTToolsWindow.RebarLabelGap
        w := (editX + PTToolsWindow.RebarFieldW + 15) - x0
        y := PTToolsWindow.TopY

        g.Add("GroupBox", "x" x0 " y" PTToolsWindow.TopY " w" w " h240", "Bars required")
        PTToolsWindow.Combo(g, labelX, editX, y + 27, "Area expression", "RequiredAreaExpression")
            .OnEvent("Change", (p*) => PTToolsWindow.OnRequiredExpressionChanged(p*))
        PTToolsWindow.Field(g, labelX, editX, y + 67, PTToolsWindow.RebarFieldW, "Total area (mm²)", "RequiredTotalArea", "", "ReadOnly Right")
        PTToolsWindow.Combo(g, labelX, editX, y + 107, "Bar diameter (mm)", "RequiredBarDiameter")
            .OnEvent("Change", (p*) => PTToolsWindow.OnRequiredFieldChanged(p*))

        ; Safety factor: an optional multiplier on the required area before it's
        ; converted to a bar count, e.g. 1.20 for a 20% design margin. Disabled
        ; by default so it never silently changes a result the user didn't ask for.
        g.AddCheckBox("x" labelX " y" (y + 149) " w130 -Tabstop vSafetyFactorEnabled Checked" S["SafetyFactorEnabled"], "Safety factor")
            .OnEvent("Click", (p*) => PTToolsWindow.OnSafetyFactorToggled(p*))
        g.AddComboBox("x" editX " y" (y + 147) " w" PTToolsWindow.RebarFieldW " vSafetyFactor", PTToolsWindow.SafetyFactors)
            .OnEvent("Change", (p*) => PTToolsWindow.OnRequiredFieldChanged(p*))

        PTToolsWindow.Field(g, labelX, editX, y + 187, PTToolsWindow.RebarFieldW, "Number of bars", "RequiredBarCount", "", "ReadOnly Right")

        g["SafetyFactor"].Enabled := S["SafetyFactorEnabled"]
        return x0 + w
    }

    static BuildMeshAreaGroup(g, S, x0) {
        labelX := x0 + 15
        editX  := labelX + PTToolsWindow.RebarLabelGap
        w := (editX + PTToolsWindow.RebarFieldW + 15) - x0
        y := PTToolsWindow.TopY

        g.Add("GroupBox", "x" x0 " y" PTToolsWindow.TopY " w" w " h" PTToolsWindow.TallGroupH, "Mesh area")
        changed := (p*) => PTToolsWindow.OnMeshFieldChanged(p*)
        PTToolsWindow.Combo(g, labelX, editX, y + 27, "Span width (mm)", "MeshSpanWidth").OnEvent("Change", changed)
        PTToolsWindow.Combo(g, labelX, editX, y + 67, "Top mesh (A/B/D/E)", "MeshTopMark").OnEvent("Change", changed)
        PTToolsWindow.Field(g, labelX, editX, y + 107, PTToolsWindow.RebarFieldW, "Top mesh area (mm²)", "MeshTopArea", "", "ReadOnly Right")
        PTToolsWindow.Field(g, labelX, editX, y + 147, PTToolsWindow.RebarFieldW, "Top mesh + bar (mm²)", "MeshTopCombinedArea", "", "ReadOnly Right")
        PTToolsWindow.Combo(g, labelX, editX, y + 187, "Bot mesh (A/B/D/E)", "MeshBotMark").OnEvent("Change", changed)
        PTToolsWindow.Field(g, labelX, editX, y + 227, PTToolsWindow.RebarFieldW, "Bot mesh area (mm²)", "MeshBotArea", "", "ReadOnly Right")
        PTToolsWindow.Field(g, labelX, editX, y + 267, PTToolsWindow.RebarFieldW, "Bot mesh + bar (mm²)", "MeshBotCombinedArea", "", "ReadOnly Right")
        return x0 + w
    }

    static BuildExpressionGroup(g, S, requiredGroupRight) {
        ; Sits under the Rebar area + Bars required columns only
        ; (the taller Mesh area column stands beside it, not above it).
        x0 := PTToolsWindow.Margin, y := PTToolsWindow.TopY + PTToolsWindow.TallGroupH - 70   ; bottom aligned with Mesh area
        w := requiredGroupRight - x0
        g.Add("GroupBox", "x" x0 " y" y " w" w " h70", "Calculator")

        inputW  := (w - 45) // 2                                            ; 45 = margins (15+15) + "=" sign column (15)
        eqX     := x0 + 15 + inputW + 8
        resultX := eqX + 23
        resultW := w - (resultX - x0) - 15

        g.AddComboBox("x" (x0 + 15) " y" (y + 28) " w" inputW " vExprInput", PTToolsWindow.Expressions)
            .OnEvent("Change", (p*) => PTToolsWindow.OnExpressionChanged(p*))
        g.AddText("x" eqX " y" (y + 31) " w15", "=")
        g.AddEdit("x" resultX " y" (y + 28) " w" resultW " r1 ReadOnly -Tabstop Right vExprResult")
    }

    ; An input of Rebar area changed: its results, and the mesh area's span width follows (values changed by the program don't fire Change,
    ; so the mesh area and "+ bar" are recomputed here too)
    static OnRebarFieldChanged(*) {
        g := PTToolsWindow.G
        PTToolsWindow.UpdateRebarArea()
        g["MeshSpanWidth"].Text := PTToolsWindow.Num(g["RebarSpanWidth"].Text)
        PTToolsWindow.OnMeshFieldChanged()
    }

    static UpdateRebarArea() {
        g := PTToolsWindow.G
        spanWidth := PTToolsWindow.Num(g["RebarSpanWidth"].Text)           ; combo boxes: .Value would be the LIST INDEX when the text matches a preset item
        spacing   := PTToolsWindow.Num(g["RebarBarSpacing"].Text)
        diameter  := PTToolsWindow.Num(g["RebarBarDiameter"].Text)
        count := (spacing != 0) ? Round(spanWidth / spacing, 2) : 0
        g["RebarBarCount"].Value := PTToolsWindow.Fmt2(count)
        g["RebarTotalArea"].Value := PTToolsWindow.Fmt2(count * PTToolsWindow.CalculateRebarArea(diameter))
    }

    static OnRequiredExpressionChanged(*) {
        g := PTToolsWindow.G
        result := PTCalc.Eval(g["RequiredAreaExpression"].Text)              ; numbers/+-*/^() only, no arbitrary eval
        g["RequiredTotalArea"].Value := PTToolsWindow.Fmt2(result)
        PTToolsWindow.OnRequiredFieldChanged()
    }

    static OnRequiredFieldChanged(*) {
        g := PTToolsWindow.G
        totalArea := PTToolsWindow.Num(g["RequiredTotalArea"].Value)
        if (g["SafetyFactorEnabled"].Value)
            totalArea *= PTToolsWindow.Num(g["SafetyFactor"].Text, 1.0)
        area := PTToolsWindow.CalculateRebarArea(PTToolsWindow.Num(g["RequiredBarDiameter"].Text)) ; .Value would be the LIST INDEX when the text matches a preset item
        g["RequiredBarCount"].Value := (area != 0) ? Ceil(totalArea / area) : 0
    }

    ; The factor value only matters while it's actually being applied.
    static OnSafetyFactorToggled(*) {
        g := PTToolsWindow.G
        g["SafetyFactor"].Enabled := g["SafetyFactorEnabled"].Value
        PTToolsWindow.OnRequiredFieldChanged()
    }

    static OnMeshFieldChanged(*) {
        g := PTToolsWindow.G
        spanWidth := PTToolsWindow.Num(g["MeshSpanWidth"].Text)

        topMesh := PTToolsWindow.CalculateMeshArea(g["MeshTopMark"].Text)
        g["MeshTopArea"].Value := IsNumber(topMesh) ? PTToolsWindow.Fmt2(spanWidth / 1000 * topMesh) : ""

        botMesh := PTToolsWindow.CalculateMeshArea(g["MeshBotMark"].Text)
        g["MeshBotArea"].Value := IsNumber(botMesh) ? PTToolsWindow.Fmt2(spanWidth / 1000 * botMesh) : ""

        PTToolsWindow.RecalculateCombinedAreas()
    }

    ; Top / bot "+ bar" fields = that mesh area plus the rebar area figure -
    ; shared by both the Rebar area and the Mesh area group box, so both handlers call this.
    static RecalculateCombinedAreas() {
        g := PTToolsWindow.G
        rebarArea := PTToolsWindow.Num(g["RebarTotalArea"].Value)
        topArea   := g["MeshTopArea"].Value
        botArea   := g["MeshBotArea"].Value
        g["MeshTopCombinedArea"].Value := IsNumber(topArea) ? PTToolsWindow.Fmt2(topArea + rebarArea) : ""
        g["MeshBotCombinedArea"].Value := IsNumber(botArea) ? PTToolsWindow.Fmt2(botArea + rebarArea) : ""
    }

    static OnExpressionChanged(*) {
        g := PTToolsWindow.G
        g["ExprResult"].Value := PTToolsWindow.Fmt2(PTCalc.Eval(g["ExprInput"].Text))
    }

    ; Bar cross-section area from the bar diameter (mm2)
    static CalculateRebarArea(diameter) {
        return Round((diameter / 2) ** 2.0 * 3.1415926, 2)
    }

    ; Mesh cross-section area from the mesh type (e.g. "A7" / "B8" / "D10" / "E6") (mm2/m)
    static CalculateMeshArea(meshMark) {
        wireGrade := Format("{:U}", SubStr(meshMark, 1, 1))
        wireSizeStr := SubStr(meshMark, 2)
        if !IsNumber(wireSizeStr)
            return ""
        wireSize := wireSizeStr + 0
        if (wireGrade = "A")
            return Round((wireSize / 2) ** 2.0 * 3.1415926 * (1000 / 200), 2)
        if (wireGrade = "B" || wireGrade = "D")
            return Round((wireSize / 2) ** 2.0 * 3.1415926 * (1000 / 100), 2)
        if (wireGrade = "E")
            return Round((wireSize / 2) ** 2.0 * 3.1415926 * (1000 / 150), 2)
        return ""
    }
}

;===============================================================================
; PTCalc - evaluates arithmetic expressions (only numbers and + - * / ^ ** ( )); a copy of
; ALTRun's Calc (Lib\Util.ahk), so this file runs on its own; ALTRun's tests check both give the same results
; Deliberately no eval, so no arbitrary code runs. Recursive descent, all in floating point (large numbers don't overflow);
; precedence: parentheses > power (right-associative, 2^3^2 = 2^(3^2)) > sign (-2^2 = -4, 2^-2 = 0.25) > * / > + -
;===============================================================================
class PTCalc {
    ; Returns the result (a float); "" when it isn't a complete expression, divides by 0 or is out of range
    static Eval(expr) {
        expr := RegExReplace(expr, "\s")
        if !RegExMatch(expr, "^[\d\+\-\*/\^\(\)\.]+$")
            return ""
        state := {Text: expr, Pos: 1}
        try {
            value := PTCalc._Sum(state)
            if (state.Pos <= StrLen(expr))                                  ; characters left over, e.g. "(1+2))"
                return ""
        } catch
            return ""
        return (Abs(value) <= 1.7976931348623157e308) ? value : ""          ; overflow (inf) / NaN aren't results
    }

    static _Sum(state) {                                                    ; + -
        value := PTCalc._Product(state)
        while (op := PTCalc._Peek(state, "+-")) {
            state.Pos += 1
            operand := PTCalc._Product(state)
            value := (op = "+") ? value + operand : value - operand
        }
        return value
    }

    static _Product(state) {                                                ; * /
        value := PTCalc._Signed(state)
        while (op := PTCalc._Peek(state, "*/")) {
            state.Pos += 1
            operand := PTCalc._Signed(state)
            value := (op = "*") ? value * operand : value / operand         ; division by 0 throws; Eval returns ""
        }
        return value
    }

    static _Signed(state) {                                                 ; sign
        if (op := PTCalc._Peek(state, "+-")) {
            state.Pos += 1
            return (op = "-") ? -PTCalc._Signed(state) : PTCalc._Signed(state)
        }
        return PTCalc._Power(state)
    }

    static _Power(state) {                                                  ; power: ^ or **
        base := PTCalc._Primary(state)
        if (SubStr(state.Text, state.Pos, 1) = "^")
            state.Pos += 1
        else if (SubStr(state.Text, state.Pos, 2) = "**")
            state.Pos += 2
        else
            return base
        return base ** PTCalc._Signed(state)                                  ; the exponent can have a sign and be a power itself
    }

    static _Primary(state) {                                                ; a number or parentheses
        if (SubStr(state.Text, state.Pos, 1) = "(") {
            state.Pos += 1
            value := PTCalc._Sum(state)
            if (SubStr(state.Text, state.Pos, 1) != ")")
                throw ValueError("Missing )")
            state.Pos += 1
            return value
        }
        if !RegExMatch(state.Text, "\G(\d+\.?\d*|\.\d+)", &m, state.Pos)
            throw ValueError("Number expected")
        state.Pos += m.Len
        return Float(m[0])
    }

    static _Peek(state, chars) {
        ch := SubStr(state.Text, state.Pos, 1)
        return (ch != "" && InStr(chars, ch)) ? ch : ""
    }
}
