;===============================================================================
; PTTools.ahk - PT 工具箱: 钢筋/BRC 计算器 + SPF2M 束线型计算器自动化 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 从仓库根目录旧的 PTTools.ahk (AutoHotkey v1, 独立进程运行) 移植并模块化而来。
;
; 没有移植的部分: 旧文件里的 "RAPT KEY" 标签页会安装一个虚拟驱动
; (MultiKey/HASPHL), 在注册表里写入伪造的加密狗 (dongle) 数据, 让 RAPT 这款商业
; 结构工程软件误以为检测到了已授权的硬件加密狗 —— 这是绕过商业软件许可证保护的
; 工具, 不属于这个项目要维护的范围, 因此完全没有移植。这个 class 只保留了
; Rebar/BRC 计算器、表达式求值、以及 SPF2M 束线型计算器自动化这几项正常的工程
; 小工具。
;
; SPF2M 需要的资源文件 (DOSBox.exe / SDL.dll / SDL_net.dll / SPF2M.exe) 已经从旧
; PTTools.ahk 里内嵌的 Base64 数据还原成真正的二进制文件, 放在 Res\ 目录下, 和
; Run.bat 引用的文件名 (DOSBox.exe SPF2M.exe) 保持一致。
;
; 设置改成存在 ALTRun.json 的 "PTTools" 节点里, 读写方式和 g_CONFIG/g_GUI 一样,
; 都走 ALTRun.ahk 的 LoadAppData()/SaveAppData()。字段名比旧版本 (SpanWidth1/
; RebarSize1/... 按 GroupBox 编号) 更语义化, 因为是这台机器唯一一份数据, 改名时
; 直接手动同步改了 ALTRun.json 里对应的 key, 没有另外写一遍迁移代码。
;
; 用法 (ALTRun.ahk 里):
;   PTTools() { PTToolsWindow.Show() }         ; 见 ALTRun.ahk 里的 PTTools()
;===============================================================================

Class PTToolsWindow {

    ; ---------------------------------------------------------------------
    ; 默认值 / 当前设置 - 键名就是控件的 Name (v2 里用 "v" 前缀注册的那个名字),
    ; 这样 Save()/Show() 都能直接按同一份名单循环, 不用另外维护映射表。
    ; ---------------------------------------------------------------------
    static Defaults := Map(
        "WinLeft", 500, "WinTop", 500,

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
        "ProfileType", 1, "TendonType", 1,
        "StartLevel", 450, "StartAtCG", 0,
        "EndLevel", 50, "EndAtCG", 0,
        "HorizontalDistance", 12000,

        ; SPF2M automation options
        "AutoInputEnabled", 1, "CustomTimingEnabled", 0,
        "InitDelayMs", 1500, "KeyIntervalMs", 25,
        "DuctDiaMono", 25, "DuctDia7s", 70, "DuctDia12s", 90,
        "DuctDia19s", 100, "DuctDia22s", 120, "DuctDia31s", 130,
        "MinCurveRadius", "[DFT]", "ChangeSupportInterval", "N"
    )

    ; The subset of Defaults' keys that live on a GUI control and round-trip
    ; through Save()/Show() - i.e. everything except WinLeft/WinTop, which are
    ; the window position and handled separately via WinGetPos.
    static FieldNames := [
        "RebarSpanWidth", "RebarBarDiameter", "RebarBarSpacing", "RebarTotalArea", "RebarBarCount",
        "RequiredAreaExpression", "RequiredTotalArea", "RequiredBarDiameter",
        "SafetyFactorEnabled", "SafetyFactor", "RequiredBarCount",
        "BrcSpanWidth", "BrcTopMeshMark", "BrcTopArea", "BrcBotMeshMark", "BrcBotArea",
        "ExprInput", "ExprResult",
        "ProfileType", "TendonType", "StartLevel", "StartAtCG",
        "EndLevel", "EndAtCG", "HorizontalDistance",
        "AutoInputEnabled", "CustomTimingEnabled", "InitDelayMs", "KeyIntervalMs",
        "DuctDiaMono", "DuctDia7s", "DuctDia12s", "DuctDia19s", "DuctDia22s", "DuctDia31s",
        "MinCurveRadius", "ChangeSupportInterval"
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
    static G := ""                                                        ; the Gui object, while the window is open
    static HotkeysReady := false

    ; Called from ALTRun.ahk's LoadAppData(), same pattern as g_CONFIG/g_GUI.
    static Load(saved) {
        PTToolsWindow.Settings := PTToolsWindow.Defaults.Clone()
        MergeIntoDefaults(PTToolsWindow.Settings, saved)                   ; shared helper, defined in ALTRun.ahk
        return PTToolsWindow.Settings
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
    ; Window
    ; ===================================================================
    static Show() {
        if WinExist("PT Tools") {
            WinActivate("PT Tools")
            return
        }

        S := PTToolsWindow.Settings
        g := Gui("+AlwaysOnTop", "PT Tools")
        PTToolsWindow.G := g
        g.SetFont("s9", "Microsoft YaHei")
        g.OnEvent("Close", (p*) => PTToolsWindow.OnClose(p*))

        tab := g.AddTab3(, ["REBAR", "SPF2M"])

        tab.UseTab(1)
        PTToolsWindow.BuildRebarTab(g, S)

        tab.UseTab(2)
        PTToolsWindow.BuildSpf2mTab(g, S)

        tab.UseTab(0)

        PTToolsWindow.SetupHotkeys()

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
        for _, name in PTToolsWindow.FieldNames {
            try S[name] := PTToolsWindow.ComboFields.Has(name) ? g[name].Text : g[name].Value
        }
        try {
            WinGetPos(&x, &y, , , "ahk_id " g.Hwnd)
            S["WinLeft"] := x
            S["WinTop"] := y
        }
        SaveAppData()                                                      ; ALTRun.ahk global - writes ALTRun.json right away
    }

    ; ===================================================================
    ; Tab 1 - REBAR
    ; ===================================================================
    ; Adds a "Label:" + Edit control pair on one row (label at x, edit at
    ; editX) and returns the Edit control, so callers can still chain
    ; .OnEvent(...) - shared by every numeric/text field on this tab.
    static Field(g, x, editX, y, w, label, name, value, opts := "") {
        g.AddText("x" x " y" (y + 3), label)
        return g.AddEdit("x" editX " y" y " w" w " r1 " opts " v" name, value)
    }

    static BuildRebarTab(g, S) {
        PTToolsWindow.BuildRebarAreaGroup(g, S)
        PTToolsWindow.BuildRequiredQtyGroup(g, S)
        PTToolsWindow.BuildBrcAreaGroup(g, S)
        PTToolsWindow.BuildExpressionGroup(g, S)

        ; ComboBox initial text is set last, once every sibling control any of
        ; this tab's Change handlers touch already exists - defends against a
        ; programmatic .Text write firing Change synchronously mid-construction
        ; (see the note on ComboFields at the top of the class).
        g["RebarBarDiameter"].Text := S["RebarBarDiameter"]
        g["RequiredBarDiameter"].Text := S["RequiredBarDiameter"]
    }

    static BuildRebarAreaGroup(g, S) {
        g.Add("GroupBox", "x20 y15 w240 h240", "Total Rebar Area")
        PTToolsWindow.Field(g, 35, 170, 42, 80, "Span Width (mm)", "RebarSpanWidth", S["RebarSpanWidth"], "+Number")
            .OnEvent("Change", (p*) => PTToolsWindow.OnRebarFieldChanged(p*))

        g.AddText("x35 y85", "Rebar Diameter (mm)")
        g.AddComboBox("x170 y82 w80 vRebarBarDiameter", PTToolsWindow.BarDiameters)
            .OnEvent("Change", (p*) => PTToolsWindow.OnRebarFieldChanged(p*))

        PTToolsWindow.Field(g, 35, 170, 122, 80, "Rebar Spacing (mm)", "RebarBarSpacing", S["RebarBarSpacing"])
            .OnEvent("Change", (p*) => PTToolsWindow.OnRebarFieldChanged(p*))
        PTToolsWindow.Field(g, 35, 170, 162, 80, "Rebar Area (mm2)", "RebarTotalArea", PTToolsWindow.Fmt2(S["RebarTotalArea"]), "ReadOnly")
        PTToolsWindow.Field(g, 35, 170, 202, 80, "Rebar Count (No.)", "RebarBarCount", PTToolsWindow.Fmt2(S["RebarBarCount"]), "ReadOnly")
    }

    static BuildRequiredQtyGroup(g, S) {
        g.Add("GroupBox", "x270 y15 w240 h240", "Rebar Qty Required")
        PTToolsWindow.Field(g, 285, 420, 42, 80, "Area Expression", "RequiredAreaExpression", S["RequiredAreaExpression"])
            .OnEvent("Change", (p*) => PTToolsWindow.OnRequiredExpressionChanged(p*))
        PTToolsWindow.Field(g, 285, 420, 82, 80, "Total Area (mm2)", "RequiredTotalArea", PTToolsWindow.Fmt2(S["RequiredTotalArea"]), "ReadOnly")

        g.AddText("x285 y125", "Rebar Diameter (mm)")
        g.AddComboBox("x420 y122 w80 vRequiredBarDiameter", PTToolsWindow.BarDiameters)
            .OnEvent("Change", (p*) => PTToolsWindow.OnRequiredFieldChanged(p*))

        ; Safety factor: an optional multiplier on the required area before it's
        ; converted to a bar count, e.g. 1.20 for a 20% design margin. Disabled
        ; by default so it never silently changes a result the user didn't ask for.
        g.AddCheckBox("x285 y164 w130 vSafetyFactorEnabled Checked" S["SafetyFactorEnabled"], "Safety Factor")
            .OnEvent("Click", (p*) => PTToolsWindow.OnSafetyFactorToggled(p*))
        g.AddEdit("x420 y162 w80 r1 vSafetyFactor", S["SafetyFactor"])
            .OnEvent("Change", (p*) => PTToolsWindow.OnRequiredFieldChanged(p*))

        PTToolsWindow.Field(g, 285, 420, 202, 80, "Rebar Count (No.)", "RequiredBarCount", S["RequiredBarCount"], "ReadOnly")

        g["SafetyFactor"].Enabled := S["SafetyFactorEnabled"]
    }

    static BuildBrcAreaGroup(g, S) {
        g.Add("GroupBox", "x520 y15 w240 h320", "BRC Area")
        PTToolsWindow.Field(g, 535, 670, 42, 80, "Span Width (mm)", "BrcSpanWidth", S["BrcSpanWidth"])
            .OnEvent("Change", (p*) => PTToolsWindow.OnBrcFieldChanged(p*))
        PTToolsWindow.Field(g, 535, 670, 82, 80, "Top BRC Mark A/B/D/E", "BrcTopMeshMark", S["BrcTopMeshMark"], "Uppercase")
            .OnEvent("Change", (p*) => PTToolsWindow.OnBrcFieldChanged(p*))
        PTToolsWindow.Field(g, 535, 670, 122, 80, "Top Mesh Area (mm2)", "BrcTopArea", PTToolsWindow.Fmt2(S["BrcTopArea"]), "ReadOnly")
        PTToolsWindow.Field(g, 535, 670, 162, 80, "Top BRC + Rebar", "BrcTopCombinedArea", "0.00", "ReadOnly")
        PTToolsWindow.Field(g, 535, 670, 202, 80, "Bot BRC Mark A/B/D/E", "BrcBotMeshMark", S["BrcBotMeshMark"], "Uppercase")
            .OnEvent("Change", (p*) => PTToolsWindow.OnBrcFieldChanged(p*))
        PTToolsWindow.Field(g, 535, 670, 242, 80, "Bot Mesh Area (mm2)", "BrcBotArea", PTToolsWindow.Fmt2(S["BrcBotArea"]), "ReadOnly")
        PTToolsWindow.Field(g, 535, 670, 282, 80, "Bot BRC + Rebar", "BrcBotCombinedArea", "0.00", "ReadOnly")
    }

    static BuildExpressionGroup(g, S) {
        g.Add("GroupBox", "x20 y270 w490 h75", "Expression Evaluation")
        g.AddEdit("x35 y300 w220 r1 vExprInput", S["ExprInput"])
            .OnEvent("Change", (p*) => PTToolsWindow.OnExpressionChanged(p*))
        g.AddText("x262 y303 w15", "=")
        g.AddEdit("x280 y300 w215 r1 ReadOnly vExprResult", PTToolsWindow.Fmt2(S["ExprResult"]))
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

        PTToolsWindow.RecalculateCombinedAreas()
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

    ; 根据钢筋直径计算钢筋截面积 (mm2)
    static CalculateRebarArea(diameter) {
        return Round((diameter / 2) ** 2.0 * 3.1415926, 2)
    }

    ; 根据钢筋网型号 (如 "A7"/"B8"/"D10"/"E6") 计算钢筋网截面积 (mm2/m)
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
    ; Tab 2 - SPF2M (Post-Tension Profile Calculator automation)
    ; ===================================================================
    static BuildSpf2mTab(g, S) {
        PTToolsWindow.BuildProfileGroup(g, S)
        PTToolsWindow.BuildAutomationOptionsGroup(g, S)
        PTToolsWindow.ApplyFieldEnableState(S["AutoInputEnabled"], S["CustomTimingEnabled"])
    }

    static BuildProfileGroup(g, S) {
        g.Add("GroupBox", "x20 y15 w380 h250", "SPF2M (Post-Tension Profile Calculator)")
        g.AddText("x35 y45", "Profile Type")
        g.AddDropDownList("x160 y42 w220 Choose" S["ProfileType"] " vProfileType", PTToolsWindow.ProfileTypes)
        g.AddText("x35 y80", "Tendon Type")
        g.AddDropDownList("x160 y77 w220 Choose" S["TendonType"] " vTendonType", PTToolsWindow.TendonTypes)

        PTToolsWindow.Field(g, 35, 200, 112, 90, "Start Level (mm)", "StartLevel", S["StartLevel"])
        g.AddCheckBox("x300 y114 w80 vStartAtCG Checked" S["StartAtCG"], "At C.G.")

        PTToolsWindow.Field(g, 35, 200, 147, 90, "End Level (mm)", "EndLevel", S["EndLevel"])
        g.AddCheckBox("x300 y149 w80 vEndAtCG Checked" S["EndAtCG"], "At C.G.")

        PTToolsWindow.Field(g, 35, 200, 182, 90, "Horizontal Distance (mm)", "HorizontalDistance", S["HorizontalDistance"])

        g.AddButton("x35 y220 w140 h30", "Run SPF2M").OnEvent("Click", (p*) => PTToolsWindow.OnRunSpf2mClicked(p*))
    }

    static BuildAutomationOptionsGroup(g, S) {
        g.Add("GroupBox", "x410 y15 w360 h280", "Automation Options")
        cbAuto := g.AddCheckBox("x425 y45 w330 vAutoInputEnabled Checked" S["AutoInputEnabled"], "Automatically Input Data into SPF2M")
        cbAuto.OnEvent("Click", (p*) => PTToolsWindow.OnAutoInputToggled(p*))
        cbCustom := g.AddCheckBox("x425 y75 w330 vCustomTimingEnabled Checked" S["CustomTimingEnabled"], "Use Custom Timing && Duct Settings")
        cbCustom.OnEvent("Click", (p*) => PTToolsWindow.OnCustomTimingToggled(p*))

        PTToolsWindow.Field(g, 425, 545, 109, 50, "Startup Delay (ms)", "InitDelayMs", S["InitDelayMs"])
        PTToolsWindow.Field(g, 610, 710, 109, 50, "Interval (ms)", "KeyIntervalMs", S["KeyIntervalMs"])

        PTToolsWindow.Field(g, 425, 545, 139, 50, "Mono Duct Dia.", "DuctDiaMono", S["DuctDiaMono"])
        PTToolsWindow.Field(g, 610, 710, 139, 50, "7s Duct Dia.", "DuctDia7s", S["DuctDia7s"])

        PTToolsWindow.Field(g, 425, 545, 169, 50, "12s Duct Dia.", "DuctDia12s", S["DuctDia12s"])
        PTToolsWindow.Field(g, 610, 710, 169, 50, "19s Duct Dia.", "DuctDia19s", S["DuctDia19s"])

        PTToolsWindow.Field(g, 425, 545, 199, 50, "22s Duct Dia.", "DuctDia22s", S["DuctDia22s"])
        PTToolsWindow.Field(g, 610, 710, 199, 50, "31s Duct Dia.", "DuctDia31s", S["DuctDia31s"])

        PTToolsWindow.Field(g, 425, 545, 229, 50, "Min. Curve Radius", "MinCurveRadius", S["MinCurveRadius"])
        PTToolsWindow.Field(g, 610, 710, 229, 50, "Chg Support (Y/N)", "ChangeSupportInterval", S["ChangeSupportInterval"])
    }

    static OnAutoInputToggled(*) {
        g := PTToolsWindow.G
        autoInput := g["AutoInputEnabled"].Value
        if (!autoInput)
            g["CustomTimingEnabled"].Value := 0
        PTToolsWindow.ApplyFieldEnableState(autoInput, g["CustomTimingEnabled"].Value)
    }

    static OnCustomTimingToggled(*) {
        g := PTToolsWindow.G
        PTToolsWindow.ApplyFieldEnableState(g["AutoInputEnabled"].Value, g["CustomTimingEnabled"].Value)
    }

    ; Profile fields only matter when we're actually going to type them into
    ; SPF2M, and the timing/duct-dia overrides only matter while that typing
    ; is happening - so both groups are gated behind "auto input" being on.
    static ApplyFieldEnableState(autoInputEnabled, customTimingEnabled) {
        g := PTToolsWindow.G
        for _, name in ["ProfileType", "TendonType", "StartLevel", "StartAtCG", "EndLevel", "EndAtCG", "HorizontalDistance"]
            g[name].Enabled := autoInputEnabled
        g["CustomTimingEnabled"].Enabled := autoInputEnabled

        customEnabled := autoInputEnabled && customTimingEnabled
        for _, name in ["InitDelayMs", "KeyIntervalMs", "DuctDiaMono", "DuctDia7s", "DuctDia12s", "DuctDia19s", "DuctDia22s", "DuctDia31s", "MinCurveRadius", "ChangeSupportInterval"]
            g[name].Enabled := customEnabled
    }

    static OnRunSpf2mClicked(*) {
        g := PTToolsWindow.G

        resDir := A_ScriptDir "\Res"
        dosbox := resDir "\DOSBox.exe"
        if !FileExist(dosbox) || !FileExist(resDir "\SPF2M.exe") {
            MsgBox("SPF2M resource files not found under Res (need DOSBox.exe + SPF2M.exe + SDL.dll + SDL_net.dll).", "PT Tools", 48)
            return
        }

        autoInputEnabled := g["AutoInputEnabled"].Value

        DetectHiddenWindows(true)

        if (!autoInputEnabled) {
            if WinExist("ahk_class SDL_app")
                WinActivate()                                              ; no title = the "last found window" WinExist just set
            else
                Run('"' dosbox '" SPF2M.exe -noconsole', resDir, "Hide")
            return
        }

        ; Duct dia. lookup indexed the same way TendonType is (1=Slab/Mono..6=31s) -
        ; the old v1 tool's dropdown list had a stray blank item here that shifted
        ; this lookup by one; that bug is not reproduced.
        ductDiameters := [PTToolsWindow.Num(g["DuctDiaMono"].Value), PTToolsWindow.Num(g["DuctDia7s"].Value), PTToolsWindow.Num(g["DuctDia12s"].Value),
            PTToolsWindow.Num(g["DuctDia19s"].Value), PTToolsWindow.Num(g["DuctDia22s"].Value), PTToolsWindow.Num(g["DuctDia31s"].Value)]
        tendonType := g["TendonType"].Value

        startLevel := PTToolsWindow.Num(g["StartLevel"].Value)
        if (g["StartAtCG"].Value)
            startLevel -= ductDiameters[tendonType] / 2
        endLevel := PTToolsWindow.Num(g["EndLevel"].Value)
        if (g["EndAtCG"].Value)
            endLevel -= ductDiameters[tendonType] / 2

        keyIntervalMs := PTToolsWindow.Num(g["KeyIntervalMs"].Value, 25)
        initDelayMs   := PTToolsWindow.Num(g["InitDelayMs"].Value, 1500)

        if WinExist("ahk_class SDL_app")                                   ; Clean up a leftover previous run first
            WinClose()

        Sleep(keyIntervalMs)
        Run('"' dosbox '" SPF2M.exe -noconsole', resDir, "Hide")

        if !WinWaitActive("ahk_class SDL_app", , 20) {
            MsgBox("WinWait timed out, SPF2M window not found. Please run again.", "PT Tools", 48)
            return
        }

        Sleep(initDelayMs)
        WinActivate("ahk_class SDL_app")

        keystrokes := [g["ProfileType"].Value, "{Enter}"
            , tendonType, "{Enter}"
            , "{Enter}"                                                    ; Min. Radius of curvature - no change
            , startLevel, "{Enter}"
            , endLevel, "{Enter}"
            , g["HorizontalDistance"].Value, "{Enter}"
            , "{Enter}"                                                    ; Dist. to Point of Contraflexure - no change
            , g["ChangeSupportInterval"].Value, "{Enter}"]

        for _, keystroke in keystrokes {
            Sleep(keyIntervalMs)
            SendInput(keystroke)
        }
    }
}
