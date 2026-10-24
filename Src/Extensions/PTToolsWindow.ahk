;===============================================================================
; PTToolsWindow.ahk - PT Tools: rebar / BRC calculator + SPF2M tendon profile calculator (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Ported and modularized from the old PTTools.ahk in the repository root (AutoHotkey v1, ran as a
; separate process). Only the engineering tools were ported: the rebar / BRC calculator, expression
; evaluation and the SPF2M tendon profile calculator; the rest of the old file was left out.
;
; The rebar / BRC calculator and SPF2M are two separate windows (PTToolsWindow.Show() /
; PTToolsWindow.ShowSpf2m()) instead of tabs, so each can be opened and closed on its own and
; remembers its own position. Both share one Settings map (the "PTTools" node in ALTRun.json),
; but each only reads and writes its own fields (RebarFieldNames / Spf2mFieldNames) and has its
; own window position (WinLeft/WinTop for PT Tools, Spf2mWinLeft/Spf2mWinTop for SPF2M).
;
; SPF2M tendon profiles are calculated by TendonProfile.ahk, with the same results as the original
; SPF2M.EXE (a DOS program that used to run in DOSBox and is no longer shipped).
;
; Settings live in the Extensions.PTTools node of ALTRun.json (AppSettings.Extension("PTTools"));
; Settings points straight at that Map, and saving calls AppSettings.Save(). The field names are
; more descriptive than the old ones (SpanWidth1 / RebarSize1 / ... numbered by GroupBox); the old
; keys were renamed by hand in the only existing ALTRun.json, so there is no migration code.
;
; Usage: the system commands "PTTools" / "SPF2M" (see SystemProvider), or the same Action in custom hotkeys
;   PTToolsWindow.Show()          ; rebar / BRC calculator
;   PTToolsWindow.ShowSpf2m()     ; SPF2M tendon profile calculator
;===============================================================================

Class PTToolsWindow {

    ; ---------------------------------------------------------------------
    ; Defaults / current settings - the keys are the controls' Names (the names registered with the "v" prefix in v2),
    ; so Save() / Show() loop over the same list without a separate mapping table.
    ; ---------------------------------------------------------------------
    static Defaults := Map(
        "WinLeft", 500, "WinTop", 500,
        "Spf2mWinLeft", 550, "Spf2mWinTop", 550,

        ; Total Rebar Area
        "RebarSpanWidth", 3000, "RebarBarDiameter", "20", "RebarBarSpacing", 200,
        "RebarTotalArea", 5969.04, "RebarBarCount", 19,

        ; Rebar Qty Required
        "RequiredAreaExpression", "1000", "RequiredTotalArea", "1000", "RequiredBarDiameter", "20",
        "SafetyFactorEnabled", 0, "SafetyFactor", 1.0, "RequiredBarCount", 4,

        ; BRC Area
        "BrcSpanWidth", 3800, "BrcTopMeshMark", "A7", "BrcTopArea", 731.196,
        "BrcBotMeshMark", "A9", "BrcBotArea", 1208.742,

        ; Expression Evaluation
        "ExprInput", "100+50", "ExprResult", "150",

        ; SPF2M profile
        "ProfileType", 1, "TendonType", 3,                                  ; 12S
        "StartLevel", 450, "StartAtCG", 0,
        "EndLevel", 50, "EndAtCG", 0,
        "HorizontalDistance", 12000,
        "MinRadius", "", "Contraflexure", "", "SupportIntervals", "",       ; empty = SPF2M's default

        ; Duct diameters for "At C.G.", per tendon type - empty = TendonProfile.DuctDiameters
        "DuctDiaMono", "", "DuctDia7s", "", "DuctDia12s", "",
        "DuctDia19s", "", "DuctDia22s", "", "DuctDia31s", ""
    )

    ; PT Tools (Rebar/BRC) window fields - round-trip through Show()/Save().
    static RebarFieldNames := [
        "RebarSpanWidth", "RebarBarDiameter", "RebarBarSpacing", "RebarTotalArea", "RebarBarCount",
        "RequiredAreaExpression", "RequiredTotalArea", "RequiredBarDiameter",
        "SafetyFactorEnabled", "SafetyFactor", "RequiredBarCount",
        "BrcSpanWidth", "BrcTopMeshMark", "BrcTopArea", "BrcBotMeshMark", "BrcBotArea",
        "ExprInput", "ExprResult"
    ]

    ; SPF2M window fields - round-trip through ShowSpf2m()/SaveSpf2m().
    static Spf2mFieldNames := [
        "ProfileType", "TendonType", "StartLevel", "StartAtCG",
        "EndLevel", "EndAtCG", "HorizontalDistance",
        "MinRadius", "Contraflexure", "SupportIntervals"
    ]

    ; ComboBox fields: .Value is the LIST INDEX when the current text happens to
    ; match a preset item (same as DropDownList), NOT the text - unlike Edit,
    ; where .Value is always the text. Anywhere we read or persist these two, we
    ; have to go through .Text instead of .Value.
    static ComboFields := Map("RebarBarDiameter", 1, "RequiredBarDiameter", 1)

    static ProfileTypes := ["Double Parabolic Profile", "Parabolic-Straight-Parabolic",
        "Parabolic-Straight Profile", "Straight-Parabolic Profile"]
    static TendonTypes  := ["Slab", "7S", "12S", "19S", "22S", "31S"]     ; index also selects DuctDia* below
    static BarDiameters := ["10", "13", "16", "20", "25", "32", "40"]

    static Settings := PTToolsWindow.Defaults.Clone()
    static G    := ""                                                     ; PT Tools (Rebar/BRC) Gui object, while open
    static SpfG := ""                                                     ; SPF2M Gui object, while open
    static LastResult := ""                                               ; the result of the last TendonProfile.Calc (for copying the table)
    static HotkeysReady := false
    static Spf2mHotkeysReady := false

    ; Called by App.Start(): saved is Extensions.PTTools from AppSettings, with missing fields filled from the defaults;
    ; after that Settings is the same Map, so AppSettings.Save() writes changes to ALTRun.json.
    static Load(saved) {
        for key, value in PTToolsWindow.Defaults
            if !saved.Has(key)
                saved[key] := value
        PTToolsWindow.Settings := saved
        return saved
    }

    ; "" + 0 throws TypeError in v2 (unlike v1, which silently coerced) - and every
    ; Change handler below reads a live Edit/ComboBox that's often briefly empty or
    ; partial while the user is mid-keystroke, so every numeric field read goes
    ; through this instead of a bare "+ 0".
    static Num(val, default := 0) {
        return IsNumber(val) ? val + 0 : default
    }

    ; Formats a computed value to exactly 2 decimal places for display. Round()
    ; alone isn't enough - a result that isn't exactly representable in binary
    ; floating point (most 2-decimal values aren't) can print with a dozen-plus
    ; trailing digits, e.g. Round(9550.54, 2) showing as "9550.5400000000009"
    ; once it round-trips through string concatenation. Non-numeric input (""
    ; for an invalid calc, or an error string from Calc.Eval) passes through
    ; unchanged.
    static Fmt2(n) {
        return IsNumber(n) ? Format("{:.2f}", n + 0) : n
    }

    ; ===================================================================
    ; PT Tools window - Rebar Area / Rebar Qty Required / BRC Area / Expression
    ; ===================================================================
    static Show() {
        if WinExist("PT Tools") {
            WinActivate("PT Tools")
            return
        }

        S := PTToolsWindow.Settings
        g := Gui("+AlwaysOnTop", "PT Tools")
        PTToolsWindow.G := g
        g.SetFont("s10", "Segoe UI")                                        ; the Windows UI font: clearer English and digits
        g.OnEvent("Close", (p*) => PTToolsWindow.OnClose(p*))

        rebarRight    := PTToolsWindow.BuildRebarAreaGroup(g, S, 20)
        requiredRight := PTToolsWindow.BuildRequiredQtyGroup(g, S, rebarRight + PTToolsWindow.RebarColGap)
        PTToolsWindow.BuildBrcAreaGroup(g, S, requiredRight + PTToolsWindow.RebarColGap)
        PTToolsWindow.BuildExpressionGroup(g, S, requiredRight)

        ; ComboBox initial text is set last, once every sibling control any of
        ; this window's Change handlers touch already exists - defends against a
        ; programmatic .Text write firing Change synchronously mid-construction
        ; (see the note on ComboFields at the top of the class).
        g["RebarBarDiameter"].Text := S["RebarBarDiameter"]
        g["RequiredBarDiameter"].Text := S["RequiredBarDiameter"]
        for name in ["RebarBarDiameter", "RequiredBarDiameter"]             ; the edit box inside the combo box is right-aligned too (ES_RIGHT), like the other numbers
            try ControlSetStyle("+0x2", DllCall("FindWindowEx", "Ptr", g[name].Hwnd, "Ptr", 0, "Str", "Edit", "Ptr", 0, "Ptr"))

        PTToolsWindow.SetupHotkeys()
        PTToolsWindow.SetupInputs(g)

        g.Show("x" S["WinLeft"] " y" S["WinTop"] " AutoSize")

        ; Derived display fields (BrcTopCombinedArea/BrcBotCombinedArea) aren't
        ; persisted, they're recomputed from RebarTotalArea + BRC areas every
        ; time the window opens.
        PTToolsWindow.RecalculateCombinedAreas()
    }

    static SetupHotkeys() {
        if PTToolsWindow.HotkeysReady
            return
        PTToolsWindow.HotkeysReady := true
        HotIfWinActive("PT Tools")                                        ; Enter / numpad Enter = move to next field, same as the old tool
        Hotkey("Enter", (*) => SendInput("{Tab}"))
        Hotkey("NumpadEnter", (*) => SendInput("{Tab}"))
        HotIfWinActive()
    }

    ; Input and result boxes: a little margin on both sides, so a right-aligned caret doesn't touch the digits;
    ; clicking into an input box selects all (tabbing / Enter in already selects all in Windows), so a new number can just be typed.
    ; Select all with PostMessage: after the mouse-down handling (placing the caret at the click) is done
    static SetupInputs(g) {
        margin := Round(4 * A_ScreenDPI / 96), margins := (margin << 16) | margin
        for ctrl in g {
            if (ctrl.Type != "Edit" && ctrl.Type != "ComboBox")
                continue
            inner := (ctrl.Type = "Edit") ? ctrl.Hwnd : DllCall("FindWindowEx", "Ptr", ctrl.Hwnd, "Ptr", 0, "Str", "Edit", "Ptr", 0, "Ptr")
            if inner
                SendMessage(0xD3, 3, margins, inner)                         ; EM_SETMARGINS: EC_LEFTMARGIN | EC_RIGHTMARGIN
            if (ctrl.Type = "ComboBox")
                ctrl.OnEvent("Focus", (c, *) => PostMessage(0x142, 0, 0xFFFF0000, c))   ; CB_SETEDITSEL 0, -1
            else if !(ControlGetStyle(ctrl) & 0x800)                        ; ES_READONLY result boxes don't need select-all
                ctrl.OnEvent("Focus", (c, *) => PostMessage(0xB1, 0, -1, c))   ; EM_SETSEL 0, -1
        }
    }

    static OnClose(GuiObj) {
        PTToolsWindow.Save()
        GuiObj.Destroy()
        PTToolsWindow.G := ""
    }

    static Save() {
        g := PTToolsWindow.G
        if !(g)
            return
        S := PTToolsWindow.Settings
        for _, name in PTToolsWindow.RebarFieldNames {
            try S[name] := PTToolsWindow.ComboFields.Has(name) ? g[name].Text : g[name].Value
        }
        try {
            WinGetPos(&x, &y, , , "ahk_id " g.Hwnd)
            S["WinLeft"] := x
            S["WinTop"] := y
        }
        AppSettings.Save()                                                  ; writes ALTRun.json right away
    }

    ; ---------------------------------------------------------------------
    ; Layout constants for the three side-by-side GroupBoxes below. Every
    ; label sits labelGap px to the left of its edit control - wide enough
    ; for the longest label in any of the three groups ("Top BRC Mark
    ; A/B/D/E" et al) so the text can never visually run into the edit box
    ; the way the old single labelGap=135 layout did.
    ; ---------------------------------------------------------------------
    static RebarLabelGap := 175
    static RebarFieldW   := 80
    static RebarColGap   := 10

    ; Adds a "Label:" + Edit control pair on one row (label at x, edit at
    ; editX) and returns the Edit control, so callers can still chain
    ; .OnEvent(...) - shared by every numeric/text field on this tab.
    static Field(g, x, editX, y, w, label, name, value, opts := "") {
        if InStr(opts, "ReadOnly")                                          ; Result boxes: Tab / Enter skip them; they can still be selected with the mouse to copy
            opts .= " -Tabstop"
        g.AddText("x" x " y" (y + 3), label)
        return g.AddEdit("x" editX " y" y " w" w " r1 " opts " v" name, value)
    }

    static BuildRebarAreaGroup(g, S, x0) {
        labelX := x0 + 15
        editX  := labelX + PTToolsWindow.RebarLabelGap
        w := (editX + PTToolsWindow.RebarFieldW + 15) - x0

        g.Add("GroupBox", "x" x0 " y8 w" w " h240", "Rebar Area")
        PTToolsWindow.Field(g, labelX, editX, 35, PTToolsWindow.RebarFieldW, "Span Width (mm)", "RebarSpanWidth", S["RebarSpanWidth"], "+Number Right")
            .OnEvent("Change", (p*) => PTToolsWindow.OnRebarFieldChanged(p*))

        g.AddText("x" labelX " y78", "Rebar Diameter (mm)")
        g.AddComboBox("x" editX " y75 w" PTToolsWindow.RebarFieldW " vRebarBarDiameter", PTToolsWindow.BarDiameters)
            .OnEvent("Change", (p*) => PTToolsWindow.OnRebarFieldChanged(p*))

        PTToolsWindow.Field(g, labelX, editX, 115, PTToolsWindow.RebarFieldW, "Rebar Spacing (mm)", "RebarBarSpacing", S["RebarBarSpacing"], "Right")
            .OnEvent("Change", (p*) => PTToolsWindow.OnRebarFieldChanged(p*))
        PTToolsWindow.Field(g, labelX, editX, 155, PTToolsWindow.RebarFieldW, "Rebar Area (mm²)", "RebarTotalArea", PTToolsWindow.Fmt2(S["RebarTotalArea"]), "ReadOnly Right")
        PTToolsWindow.Field(g, labelX, editX, 195, PTToolsWindow.RebarFieldW, "Number of Bars", "RebarBarCount", PTToolsWindow.Fmt2(S["RebarBarCount"]), "ReadOnly Right")

        return x0 + w
    }

    static BuildRequiredQtyGroup(g, S, x0) {
        labelX := x0 + 15
        editX  := labelX + PTToolsWindow.RebarLabelGap
        w := (editX + PTToolsWindow.RebarFieldW + 15) - x0

        g.Add("GroupBox", "x" x0 " y8 w" w " h240", "Bars Required")
        PTToolsWindow.Field(g, labelX, editX, 35, PTToolsWindow.RebarFieldW, "Area Expression", "RequiredAreaExpression", S["RequiredAreaExpression"], "Right")
            .OnEvent("Change", (p*) => PTToolsWindow.OnRequiredExpressionChanged(p*))
        PTToolsWindow.Field(g, labelX, editX, 75, PTToolsWindow.RebarFieldW, "Total Area (mm²)", "RequiredTotalArea", PTToolsWindow.Fmt2(S["RequiredTotalArea"]), "ReadOnly Right")

        g.AddText("x" labelX " y118", "Rebar Diameter (mm)")
        g.AddComboBox("x" editX " y115 w" PTToolsWindow.RebarFieldW " vRequiredBarDiameter", PTToolsWindow.BarDiameters)
            .OnEvent("Change", (p*) => PTToolsWindow.OnRequiredFieldChanged(p*))

        ; Safety factor: an optional multiplier on the required area before it's
        ; converted to a bar count, e.g. 1.20 for a 20% design margin. Disabled
        ; by default so it never silently changes a result the user didn't ask for.
        g.AddCheckBox("x" labelX " y157 w130 -Tabstop vSafetyFactorEnabled Checked" S["SafetyFactorEnabled"], "Safety Factor")
            .OnEvent("Click", (p*) => PTToolsWindow.OnSafetyFactorToggled(p*))
        g.AddEdit("x" editX " y155 w" PTToolsWindow.RebarFieldW " r1 Right vSafetyFactor", S["SafetyFactor"])
            .OnEvent("Change", (p*) => PTToolsWindow.OnRequiredFieldChanged(p*))

        PTToolsWindow.Field(g, labelX, editX, 195, PTToolsWindow.RebarFieldW, "Number of Bars", "RequiredBarCount", S["RequiredBarCount"], "ReadOnly Right")

        g["SafetyFactor"].Enabled := S["SafetyFactorEnabled"]
        return x0 + w
    }

    static BuildBrcAreaGroup(g, S, x0) {
        labelX := x0 + 15
        editX  := labelX + PTToolsWindow.RebarLabelGap
        w := (editX + PTToolsWindow.RebarFieldW + 15) - x0

        g.Add("GroupBox", "x" x0 " y8 w" w " h320", "BRC Area")
        PTToolsWindow.Field(g, labelX, editX, 35, PTToolsWindow.RebarFieldW, "Span Width (mm)", "BrcSpanWidth", S["BrcSpanWidth"], "Right")
            .OnEvent("Change", (p*) => PTToolsWindow.OnBrcFieldChanged(p*))
        PTToolsWindow.Field(g, labelX, editX, 75, PTToolsWindow.RebarFieldW, "Top Mesh (A/B/D/E)", "BrcTopMeshMark", S["BrcTopMeshMark"], "Uppercase Right")
            .OnEvent("Change", (p*) => PTToolsWindow.OnBrcFieldChanged(p*))
        PTToolsWindow.Field(g, labelX, editX, 115, PTToolsWindow.RebarFieldW, "Top Mesh Area (mm²)", "BrcTopArea", PTToolsWindow.Fmt2(S["BrcTopArea"]), "ReadOnly Right")
        PTToolsWindow.Field(g, labelX, editX, 155, PTToolsWindow.RebarFieldW, "Top Mesh + Rebar (mm²)", "BrcTopCombinedArea", "0.00", "ReadOnly Right")
        PTToolsWindow.Field(g, labelX, editX, 195, PTToolsWindow.RebarFieldW, "Bottom Mesh (A/B/D/E)", "BrcBotMeshMark", S["BrcBotMeshMark"], "Uppercase Right")
            .OnEvent("Change", (p*) => PTToolsWindow.OnBrcFieldChanged(p*))
        PTToolsWindow.Field(g, labelX, editX, 235, PTToolsWindow.RebarFieldW, "Bottom Mesh Area (mm²)", "BrcBotArea", PTToolsWindow.Fmt2(S["BrcBotArea"]), "ReadOnly Right")
        PTToolsWindow.Field(g, labelX, editX, 275, PTToolsWindow.RebarFieldW, "Bottom Mesh + Rebar (mm²)", "BrcBotCombinedArea", "0.00", "ReadOnly Right")
    }

    static BuildExpressionGroup(g, S, requiredGroupRight) {
        ; Sits under the Total Rebar Area + Rebar Qty Required columns only
        ; (same as the taller BRC Area column standing beside it, not under it).
        w := requiredGroupRight - 20
        g.Add("GroupBox", "x20 y258 w" w " h70", "Calculator")   ; bottom aligned with BRC Area (8 + 320)

        inputW  := (w - 65) // 2                                            ; 65 = margins (35+15) + "=" sign column (15)
        eqX     := 35 + inputW + 8
        resultX := eqX + 23
        resultW := w - (resultX - 20) - 15

        g.AddEdit("x35 y286 w" inputW " r1 vExprInput", S["ExprInput"])
            .OnEvent("Change", (p*) => PTToolsWindow.OnExpressionChanged(p*))
        g.AddText("x" eqX " y289 w15", "=")
        g.AddEdit("x" resultX " y286 w" resultW " r1 ReadOnly -Tabstop Right vExprResult", PTToolsWindow.Fmt2(S["ExprResult"]))
    }

    static OnRebarFieldChanged(*) {
        g := PTToolsWindow.G
        spanWidth := PTToolsWindow.Num(g["RebarSpanWidth"].Value)
        spacing   := PTToolsWindow.Num(g["RebarBarSpacing"].Value)
        diameter  := PTToolsWindow.Num(g["RebarBarDiameter"].Text)         ; .Value would be the LIST INDEX when the text matches a preset item

        g["BrcSpanWidth"].Value := spanWidth
        count := (spacing != 0) ? Round(spanWidth / spacing, 2) : 0
        g["RebarBarCount"].Value := PTToolsWindow.Fmt2(count)
        g["RebarTotalArea"].Value := PTToolsWindow.Fmt2(count * PTToolsWindow.CalculateRebarArea(diameter))

        PTToolsWindow.OnBrcFieldChanged()                                   ; BRC's Span Width follows (values changed by the program don't fire Change): the mesh area and "+ Rebar" are recomputed together
    }

    static OnRequiredExpressionChanged(*) {
        g := PTToolsWindow.G
        result := Calc.Eval(g["RequiredAreaExpression"].Value)             ; Lib\Util.ahk - numbers/+-*/^() only, no arbitrary eval
        g["RequiredTotalArea"].Value := PTToolsWindow.Fmt2(result)
        PTToolsWindow.OnRequiredFieldChanged()
    }

    static OnRequiredFieldChanged(*) {
        g := PTToolsWindow.G
        totalArea := PTToolsWindow.Num(g["RequiredTotalArea"].Value)
        if (g["SafetyFactorEnabled"].Value)
            totalArea *= PTToolsWindow.Num(g["SafetyFactor"].Value, 1.0)
        area := PTToolsWindow.CalculateRebarArea(PTToolsWindow.Num(g["RequiredBarDiameter"].Text)) ; .Value would be the LIST INDEX when the text matches a preset item
        g["RequiredBarCount"].Value := (area != 0) ? Ceil(totalArea / area) : 0
    }

    ; The factor value only matters while it's actually being applied.
    static OnSafetyFactorToggled(*) {
        g := PTToolsWindow.G
        g["SafetyFactor"].Enabled := g["SafetyFactorEnabled"].Value
        PTToolsWindow.OnRequiredFieldChanged()
    }

    static OnBrcFieldChanged(*) {
        g := PTToolsWindow.G
        spanWidth := PTToolsWindow.Num(g["BrcSpanWidth"].Value)

        topMesh := PTToolsWindow.CalculateMeshArea(g["BrcTopMeshMark"].Value)
        g["BrcTopArea"].Value := IsNumber(topMesh) ? PTToolsWindow.Fmt2(spanWidth / 1000 * topMesh) : ""

        botMesh := PTToolsWindow.CalculateMeshArea(g["BrcBotMeshMark"].Value)
        g["BrcBotArea"].Value := IsNumber(botMesh) ? PTToolsWindow.Fmt2(spanWidth / 1000 * botMesh) : ""

        PTToolsWindow.RecalculateCombinedAreas()
    }

    ; TOP/BOT "+ Rebar" fields = that BRC area plus the Total Rebar Area figure -
    ; shared by both the Rebar and the BRC group box, so both handlers call this.
    static RecalculateCombinedAreas() {
        g := PTToolsWindow.G
        rebarArea := PTToolsWindow.Num(g["RebarTotalArea"].Value)
        topArea   := g["BrcTopArea"].Value
        botArea   := g["BrcBotArea"].Value
        g["BrcTopCombinedArea"].Value := IsNumber(topArea) ? PTToolsWindow.Fmt2(topArea + rebarArea) : ""
        g["BrcBotCombinedArea"].Value := IsNumber(botArea) ? PTToolsWindow.Fmt2(botArea + rebarArea) : ""
    }

    static OnExpressionChanged(*) {
        g := PTToolsWindow.G
        g["ExprResult"].Value := PTToolsWindow.Fmt2(Calc.Eval(g["ExprInput"].Value))
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

    ; ===================================================================
    ; SPF2M window - tendon profile calculator
    ; ===================================================================
    ; The parameters used to be typed into SPF2M.EXE in DOSBox automatically; now TendonProfile calculates directly (with results
    ; identical to SPF2M, see TendonProfile.ahk), and changing any input updates the table on the right at once.
    static ShowSpf2m() {
        if WinExist(PTToolsWindow.Spf2mTitle " ahk_class AutoHotkeyGUI") {
            WinActivate()
            return
        }

        S := PTToolsWindow.Settings
        g := Gui("+AlwaysOnTop", PTToolsWindow.Spf2mTitle)
        PTToolsWindow.SpfG := g
        g.SetFont("s9", "Microsoft YaHei")
        g.OnEvent("Close", (p*) => PTToolsWindow.OnCloseSpf2m(p*))

        right := PTToolsWindow.BuildProfileGroup(g, S, 20)
        PTToolsWindow.BuildResultGroup(g, right + 10)
        PTToolsWindow.SetupSpf2mHotkeys()
        PTToolsWindow.SetupInputs(g)
        PTToolsWindow.OnTendonChanged()                                    ; the hints for minimum radius / duct diameter follow the strand type
        g.Show("x" S["Spf2mWinLeft"] " y" S["Spf2mWinTop"] " AutoSize")
    }

    static SetupSpf2mHotkeys() {
        if PTToolsWindow.Spf2mHotkeysReady
            return
        PTToolsWindow.Spf2mHotkeysReady := true
        HotIfWinActive(PTToolsWindow.Spf2mTitle " ahk_class AutoHotkeyGUI")                    ; Enter / numpad Enter = move to next field, same as PT Tools
        Hotkey("Enter", (*) => SendInput("{Tab}"))
        Hotkey("NumpadEnter", (*) => SendInput("{Tab}"))
        HotIfWinActive()
    }

    static OnCloseSpf2m(GuiObj) {
        PTToolsWindow.SaveSpf2m()
        GuiObj.Destroy()
        PTToolsWindow.SpfG := ""
    }

    static SaveSpf2m() {
        g := PTToolsWindow.SpfG
        if !(g)
            return
        S := PTToolsWindow.Settings
        for _, name in PTToolsWindow.Spf2mFieldNames {
            try S[name] := g[name].Value
        }
        S[PTToolsWindow.DuctDiaKey(g["TendonType"].Value)] := g["DuctDia"].Value
        try {
            WinGetPos(&x, &y, , , "ahk_id " g.Hwnd)
            S["Spf2mWinLeft"] := x
            S["Spf2mWinTop"] := y
        }
        AppSettings.Save()
    }

    ; each strand type remembers its own duct diameter ("At C.G." subtracts half the diameter from the levels)
    static DuctDiaKey(tendon) {
        static keys := ["DuctDiaMono", "DuctDia7s", "DuctDia12s", "DuctDia19s", "DuctDia22s", "DuctDia31s"]
        return keys[Max(1, Min(tendon, keys.Length))]
    }

    static Spf2mTitle := "SPF2M Post-Tensioning Tendon Profile Calculator"
    static GroupH := 420                                                   ; heights of the two GroupBoxes

    static ProfileLabelGap := 190
    static ProfileFieldW   := 90
    static ProfileCbGap    := 10
    static ProfileCbW      := 70

    static BuildProfileGroup(g, S, x0) {
        labelX := x0 + 15
        editX  := labelX + PTToolsWindow.ProfileLabelGap
        cbX    := editX + PTToolsWindow.ProfileFieldW + PTToolsWindow.ProfileCbGap
        fullW  := PTToolsWindow.ProfileFieldW + PTToolsWindow.ProfileCbGap + PTToolsWindow.ProfileCbW
        w := (cbX + PTToolsWindow.ProfileCbW + 15) - x0
        recalc := (*) => PTToolsWindow.Recalculate()

        g.Add("GroupBox", "x" x0 " y15 w" w " h" PTToolsWindow.GroupH, "Tendon Profile")
        g.AddText("x" labelX " y45", "Profile Type")
        g.AddDropDownList("x" editX " y42 w" fullW " Choose" S["ProfileType"] " vProfileType AltSubmit", PTToolsWindow.ProfileTypes).OnEvent("Change", recalc)
        g.AddText("x" labelX " y80", "Tendon Type")
        g.AddDropDownList("x" editX " y77 w" fullW " Choose" S["TendonType"] " vTendonType AltSubmit", PTToolsWindow.TendonTypes)
            .OnEvent("Change", (*) => PTToolsWindow.OnTendonChanged(true))

        PTToolsWindow.Field(g, labelX, editX, 112, PTToolsWindow.ProfileFieldW, "Start Level (mm)", "StartLevel", S["StartLevel"]).OnEvent("Change", recalc)
        g.AddCheckBox("x" cbX " y114 w" PTToolsWindow.ProfileCbW " vStartAtCG Checked" S["StartAtCG"], "At C.G.").OnEvent("Click", recalc)
        PTToolsWindow.Field(g, labelX, editX, 147, PTToolsWindow.ProfileFieldW, "End Level (mm)", "EndLevel", S["EndLevel"]).OnEvent("Change", recalc)
        g.AddCheckBox("x" cbX " y149 w" PTToolsWindow.ProfileCbW " vEndAtCG Checked" S["EndAtCG"], "At C.G.").OnEvent("Click", recalc)
        PTToolsWindow.Field(g, labelX, editX, 182, PTToolsWindow.ProfileFieldW, "Horizontal Distance (mm)", "HorizontalDistance", S["HorizontalDistance"]).OnEvent("Change", recalc)

        g.AddText("x" labelX " y222 w" (w - 30) " h1 0x10")                ; separator: optional fields below, empty = SPF2M's default
        PTToolsWindow.Field(g, labelX, editX, 235, PTToolsWindow.ProfileFieldW, "Min. Radius of Curvature (mm)", "MinRadius", S["MinRadius"]).OnEvent("Change", recalc)
        PTToolsWindow.Field(g, labelX, editX, 270, PTToolsWindow.ProfileFieldW, "Point of Contraflexure (mm)", "Contraflexure", S["Contraflexure"]).OnEvent("Change", recalc)
        PTToolsWindow.Field(g, labelX, editX, 305, fullW, "Support Intervals (mm)", "SupportIntervals", S["SupportIntervals"]).OnEvent("Change", recalc)
        Win.SetCueBanner(g["SupportIntervals"].Hwnd, "auto: max. 1000")
        PTToolsWindow.Field(g, labelX, editX, 340, PTToolsWindow.ProfileFieldW, "Duct Dia. for At C.G. (mm)", "DuctDia", "").OnEvent("Change", recalc)
        g.SetFont("s8 cGray")
        g.AddText("x" labelX " y378 w" (w - 30) " h" (PTToolsWindow.GroupH - 370), "Empty fields use the SPF2M defaults (in gray). Intervals: e.g. 500, 1500, 800 ... must add up to the distance. Distances are from the high end.")
        g.SetFont("s9 cDefault")
        return x0 + w
    }

    static BuildResultGroup(g, x0) {
        w := 480
        g.Add("GroupBox", "x" x0 " y15 w" w " h" PTToolsWindow.GroupH, "Tendon Profile - Support Heights (mm)")
        ; laid out bottom-up: the Copy Table button at the bottom of the frame, two or three summary lines above, the rest of the height for the table
        buttonY  := PTToolsWindow.GroupH - 25
        summaryY := buttonY - 53
        list := g.AddListView("x" (x0 + 12) " y40 w" (w - 24) " h" (summaryY - 48) " vResultList -Multi NoSort Grid"
            , ["Distance", "Interval", "Actual", "Beam @ 5mm", "Slab @ 10mm"])
        for index, width in [80, 76, 82, 100, 100]
            list.ModifyCol(index, width " Right")
        g.AddText("x" (x0 + 12) " y" summaryY " w" (w - 24) " h48 vResultSummary")
        g.SetFont("cRed")
        g.AddText("x" (x0 + 12) " y" summaryY " w" (w - 24) " h48 vResultError Hidden")
        g.SetFont("cDefault")
        g.AddButton("x" (x0 + 12) " y" buttonY " w140 h30", "Copy Table").OnEvent("Click", (*) => PTToolsWindow.CopyProfileTable())
    }

    ; The strand type changed: the gray hint for minimum radius becomes this type's default, the duct diameter the value remembered for it
    static OnTendonChanged(saveOld := false) {
        g := PTToolsWindow.SpfG
        S := PTToolsWindow.Settings
        static lastTendon := 0
        tendon := g["TendonType"].Value
        if (saveOld && lastTendon)
            S[PTToolsWindow.DuctDiaKey(lastTendon)] := g["DuctDia"].Value
        lastTendon := tendon
        g["DuctDia"].Value := S[PTToolsWindow.DuctDiaKey(tendon)]
        Win.SetCueBanner(g["MinRadius"].Hwnd, TendonProfile.DefaultRadius(tendon))
        Win.SetCueBanner(g["DuctDia"].Hwnd, TendonProfile.DefaultDuctDiameter(tendon))
        PTToolsWindow.Recalculate()
    }

    ; Window inputs -> the parameters of TendonProfile.Calc ("At C.G." subtracts half the duct diameter from the levels)
    static ProfileInput() {
        g := PTToolsWindow.SpfG
        halfDuct := PTToolsWindow.Num(g["DuctDia"].Value, TendonProfile.DefaultDuctDiameter(g["TendonType"].Value)) / 2
        level(name, cgName) {
            value := Trim(g[name].Value)
            if !IsNumber(value)
                return ""
            return g[cgName].Value ? value - halfDuct : value + 0
        }
        input := Map("Profile", g["ProfileType"].Value, "Tendon", g["TendonType"].Value
            , "Start", level("StartLevel", "StartAtCG"), "End", level("EndLevel", "EndAtCG")
            , "Distance", Trim(g["HorizontalDistance"].Value))
        for name, key in Map("MinRadius", "Radius", "Contraflexure", "Contraflexure") {
            value := Trim(g[name].Value)
            input[key] := IsNumber(value) ? value + 0 : ""
        }
        input["Intervals"] := TendonProfile.ParseIntervals(g["SupportIntervals"].Value)
        if IsNumber(input["Distance"])
            input["Distance"] += 0
        return input
    }

    static Recalculate() {
        g := PTToolsWindow.SpfG
        if !(g)
            return
        result := TendonProfile.Calc(PTToolsWindow.ProfileInput())
        PTToolsWindow.LastResult := result
        list := g["ResultList"]
        list.Opt("-Redraw")
        list.Delete()
        for row in result.Rows
            list.Add("", PTToolsWindow.FmtLevel(row[1]), PTToolsWindow.FmtLevel(row[2]), PTToolsWindow.FmtLevel(row[3]), PTToolsWindow.FmtLevel(row[4]), PTToolsWindow.FmtLevel(row[5]))
        list.Opt("+Redraw")
        g["ResultError"].Visible := (result.Error != "")
        g["ResultSummary"].Visible := (result.Error = "")
        g["ResultError"].Value := result.Error
        g["ResultSummary"].Value := PTToolsWindow.ProfileSummary(result)
        Win.SetCueBanner(g["Contraflexure"].Hwnd, (result.Error = "" && !result.RadiusChanged) ? Integer(result.Contraflexure) : "")
    }

    static ProfileSummary(result) {
        if (result.Error != "")
            return result.Error
        text := "Min. Radius of Curvature = " Integer(result.MinRadius) " mm"
        if (result.RadiusChanged || result.Radius != result.MinRadius)
            text .= "`nRadius of Curvature = " Integer(result.Radius) " mm"
        text .= "`nDist. to Point of Contraflexure = " Integer(result.Contraflexure) " mm (from the high end)"
        return text
    }

    ; levels may have decimals ("At C.G." minus half the duct diameter): integers as usual, otherwise one decimal
    static FmtLevel(v) {
        return (v = Integer(v)) ? Integer(v) : Round(v, 1)
    }

    ; The table copied as tab-separated text, which pastes into Excel column by column
    static CopyProfileTable() {
        result := PTToolsWindow.LastResult
        if !IsObject(result) || result.Error != ""
            return
        g := PTToolsWindow.SpfG
        text := PTToolsWindow.ProfileTypes[g["ProfileType"].Value] " / " PTToolsWindow.TendonTypes[g["TendonType"].Value] "`r`n"
            . "Distance`tInterval`tActual`tBeam @ 5mm`tSlab @ 10mm`r`n"
        for row in result.Rows
            text .= PTToolsWindow.FmtLevel(row[1]) "`t" PTToolsWindow.FmtLevel(row[2]) "`t" PTToolsWindow.FmtLevel(row[3]) "`t" PTToolsWindow.FmtLevel(row[4]) "`t" PTToolsWindow.FmtLevel(row[5]) "`r`n"
        text .= StrReplace(PTToolsWindow.ProfileSummary(result), "`n", "`r`n") "`r`n"
        A_Clipboard := text
        App.Notify("Table copied - paste it into Excel")
    }
}
