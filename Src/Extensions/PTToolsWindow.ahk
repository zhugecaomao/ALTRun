;===============================================================================
; PTToolsWindow.ahk - PT 工具箱: 钢筋/BRC 计算器 + SPF2M 束线型计算器 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 从仓库根目录旧的 PTTools.ahk (AutoHotkey v1, 独立进程运行) 移植并模块化而来。
;
; 没有移植的部分: 旧文件里的 "RAPT KEY" 标签页会安装一个虚拟驱动
; (MultiKey/HASPHL), 在注册表里写入伪造的加密狗 (dongle) 数据, 让 RAPT 这款商业
; 结构工程软件误以为检测到了已授权的硬件加密狗 —— 这是绕过商业软件许可证保护的
; 工具, 不属于这个项目要维护的范围, 因此完全没有移植。这个 class 只保留了
; Rebar/BRC 计算器、表达式求值、以及 SPF2M 束线型计算器这几项正常的工程小工具。
;
; Rebar/BRC 计算器和 SPF2M 分成两个独立窗口(PTToolsWindow.Show() /
; PTToolsWindow.ShowSpf2m()), 不再用 Tab 切换 - 各自可以单独开关、单独记住
; 窗口位置, 互不影响。两边共用同一份 Settings(都存在 ALTRun.json 的
; "PTTools" 节点里), 只是各自只读写自己关心的那部分字段(RebarFieldNames /
; Spf2mFieldNames), 也各自有一份窗口位置(WinLeft/WinTop 给 PT Tools 窗口,
; Spf2mWinLeft/Spf2mWinTop 给 SPF2M 窗口)。
;
; SPF2M 束线型直接由 TendonProfile.ahk 计算, 结果和原来的 SPF2M.EXE 一致。Resources\ 里
; 还有 DOSBox.exe / SPF2M.exe 时, 窗口里多一个 "Original SPF2M..." 按钮, 在 DOSBox 里打开
; 原程序并输入同样的参数, 用来对照。
;
; 设置存在 ALTRun.json 的 Extensions.PTTools 节点里 (AppSettings.Extension("PTTools")),
; Settings 直接指向那个 Map, 保存时调用 AppSettings.Save()。字段名比旧版本 (SpanWidth1/
; RebarSize1/... 按 GroupBox 编号) 更语义化, 因为是这台机器唯一一份数据, 改名时
; 直接手动同步改了 ALTRun.json 里对应的 key, 没有另外写一遍迁移代码。
;
; 用法: 系统命令 "PTTools" / "SPF2M" (见 SystemProvider), 或自定义热键里的同名 Action
;   PTToolsWindow.Show()          ; Rebar/BRC 计算器
;   PTToolsWindow.ShowSpf2m()     ; SPF2M 束线型计算器
;===============================================================================

Class PTToolsWindow {

    ; ---------------------------------------------------------------------
    ; 默认值 / 当前设置 - 键名就是控件的 Name (v2 里用 "v" 前缀注册的那个名字),
    ; 这样 Save()/Show() 都能直接按同一份名单循环, 不用另外维护映射表。
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
        "ProfileType", 1, "TendonType", 1,
        "StartLevel", 450, "StartAtCG", 0,
        "EndLevel", 50, "EndAtCG", 0,
        "HorizontalDistance", 12000,
        "MinRadius", "", "Contraflexure", "", "SupportIntervals", "",       ; 空 = SPF2M 的默认值

        ; Duct diameters (At C.G.) and the "Original SPF2M" automation timing
        "InitDelayMs", 1500, "KeyIntervalMs", 25,
        "DuctDiaMono", 25, "DuctDia7s", 70, "DuctDia12s", 90,
        "DuctDia19s", 100, "DuctDia22s", 120, "DuctDia31s", 130
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
    static LastResult := ""                                               ; 最近一次 TendonProfile.Calc 的结果 (复制表格用)
    static HotkeysReady := false
    static Spf2mHotkeysReady := false

    ; App.Start() 调用: saved 是 AppSettings 里的 Extensions.PTTools, 缺的字段用默认值补上,
    ; 之后 Settings 和它是同一个 Map, 修改后 AppSettings.Save() 就会写进 ALTRun.json。
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
        g.SetFont("s9", "Microsoft YaHei")
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
        g.AddText("x" x " y" (y + 3), label)
        return g.AddEdit("x" editX " y" y " w" w " r1 " opts " v" name, value)
    }

    static BuildRebarAreaGroup(g, S, x0) {
        labelX := x0 + 15
        editX  := labelX + PTToolsWindow.RebarLabelGap
        w := (editX + PTToolsWindow.RebarFieldW + 15) - x0

        g.Add("GroupBox", "x" x0 " y15 w" w " h240", "Total Rebar Area")
        PTToolsWindow.Field(g, labelX, editX, 42, PTToolsWindow.RebarFieldW, "Span Width (mm)", "RebarSpanWidth", S["RebarSpanWidth"], "+Number")
            .OnEvent("Change", (p*) => PTToolsWindow.OnRebarFieldChanged(p*))

        g.AddText("x" labelX " y85", "Rebar Diameter (mm)")
        g.AddComboBox("x" editX " y82 w" PTToolsWindow.RebarFieldW " vRebarBarDiameter", PTToolsWindow.BarDiameters)
            .OnEvent("Change", (p*) => PTToolsWindow.OnRebarFieldChanged(p*))

        PTToolsWindow.Field(g, labelX, editX, 122, PTToolsWindow.RebarFieldW, "Rebar Spacing (mm)", "RebarBarSpacing", S["RebarBarSpacing"])
            .OnEvent("Change", (p*) => PTToolsWindow.OnRebarFieldChanged(p*))
        PTToolsWindow.Field(g, labelX, editX, 162, PTToolsWindow.RebarFieldW, "Rebar Area (mm2)", "RebarTotalArea", PTToolsWindow.Fmt2(S["RebarTotalArea"]), "ReadOnly")
        PTToolsWindow.Field(g, labelX, editX, 202, PTToolsWindow.RebarFieldW, "Rebar Count (No.)", "RebarBarCount", PTToolsWindow.Fmt2(S["RebarBarCount"]), "ReadOnly")

        return x0 + w
    }

    static BuildRequiredQtyGroup(g, S, x0) {
        labelX := x0 + 15
        editX  := labelX + PTToolsWindow.RebarLabelGap
        w := (editX + PTToolsWindow.RebarFieldW + 15) - x0

        g.Add("GroupBox", "x" x0 " y15 w" w " h240", "Rebar Qty Required")
        PTToolsWindow.Field(g, labelX, editX, 42, PTToolsWindow.RebarFieldW, "Area Expression", "RequiredAreaExpression", S["RequiredAreaExpression"])
            .OnEvent("Change", (p*) => PTToolsWindow.OnRequiredExpressionChanged(p*))
        PTToolsWindow.Field(g, labelX, editX, 82, PTToolsWindow.RebarFieldW, "Total Area (mm2)", "RequiredTotalArea", PTToolsWindow.Fmt2(S["RequiredTotalArea"]), "ReadOnly")

        g.AddText("x" labelX " y125", "Rebar Diameter (mm)")
        g.AddComboBox("x" editX " y122 w" PTToolsWindow.RebarFieldW " vRequiredBarDiameter", PTToolsWindow.BarDiameters)
            .OnEvent("Change", (p*) => PTToolsWindow.OnRequiredFieldChanged(p*))

        ; Safety factor: an optional multiplier on the required area before it's
        ; converted to a bar count, e.g. 1.20 for a 20% design margin. Disabled
        ; by default so it never silently changes a result the user didn't ask for.
        g.AddCheckBox("x" labelX " y164 w130 vSafetyFactorEnabled Checked" S["SafetyFactorEnabled"], "Safety Factor")
            .OnEvent("Click", (p*) => PTToolsWindow.OnSafetyFactorToggled(p*))
        g.AddEdit("x" editX " y162 w" PTToolsWindow.RebarFieldW " r1 vSafetyFactor", S["SafetyFactor"])
            .OnEvent("Change", (p*) => PTToolsWindow.OnRequiredFieldChanged(p*))

        PTToolsWindow.Field(g, labelX, editX, 202, PTToolsWindow.RebarFieldW, "Rebar Count (No.)", "RequiredBarCount", S["RequiredBarCount"], "ReadOnly")

        g["SafetyFactor"].Enabled := S["SafetyFactorEnabled"]
        return x0 + w
    }

    static BuildBrcAreaGroup(g, S, x0) {
        labelX := x0 + 15
        editX  := labelX + PTToolsWindow.RebarLabelGap
        w := (editX + PTToolsWindow.RebarFieldW + 15) - x0

        g.Add("GroupBox", "x" x0 " y15 w" w " h320", "BRC Area")
        PTToolsWindow.Field(g, labelX, editX, 42, PTToolsWindow.RebarFieldW, "Span Width (mm)", "BrcSpanWidth", S["BrcSpanWidth"])
            .OnEvent("Change", (p*) => PTToolsWindow.OnBrcFieldChanged(p*))
        PTToolsWindow.Field(g, labelX, editX, 82, PTToolsWindow.RebarFieldW, "Top BRC Mark A/B/D/E", "BrcTopMeshMark", S["BrcTopMeshMark"], "Uppercase")
            .OnEvent("Change", (p*) => PTToolsWindow.OnBrcFieldChanged(p*))
        PTToolsWindow.Field(g, labelX, editX, 122, PTToolsWindow.RebarFieldW, "Top Mesh Area (mm2)", "BrcTopArea", PTToolsWindow.Fmt2(S["BrcTopArea"]), "ReadOnly")
        PTToolsWindow.Field(g, labelX, editX, 162, PTToolsWindow.RebarFieldW, "Top BRC + Rebar", "BrcTopCombinedArea", "0.00", "ReadOnly")
        PTToolsWindow.Field(g, labelX, editX, 202, PTToolsWindow.RebarFieldW, "Bot BRC Mark A/B/D/E", "BrcBotMeshMark", S["BrcBotMeshMark"], "Uppercase")
            .OnEvent("Change", (p*) => PTToolsWindow.OnBrcFieldChanged(p*))
        PTToolsWindow.Field(g, labelX, editX, 242, PTToolsWindow.RebarFieldW, "Bot Mesh Area (mm2)", "BrcBotArea", PTToolsWindow.Fmt2(S["BrcBotArea"]), "ReadOnly")
        PTToolsWindow.Field(g, labelX, editX, 282, PTToolsWindow.RebarFieldW, "Bot BRC + Rebar", "BrcBotCombinedArea", "0.00", "ReadOnly")
    }

    static BuildExpressionGroup(g, S, requiredGroupRight) {
        ; Sits under the Total Rebar Area + Rebar Qty Required columns only
        ; (same as the taller BRC Area column standing beside it, not under it).
        w := requiredGroupRight - 20
        g.Add("GroupBox", "x20 y270 w" w " h75", "Expression Evaluation")

        inputW  := (w - 65) // 2                                            ; 65 = margins (35+15) + "=" sign column (15)
        eqX     := 35 + inputW + 8
        resultX := eqX + 23
        resultW := w - (resultX - 20) - 15

        g.AddEdit("x35 y300 w" inputW " r1 vExprInput", S["ExprInput"])
            .OnEvent("Change", (p*) => PTToolsWindow.OnExpressionChanged(p*))
        g.AddText("x" eqX " y303 w15", "=")
        g.AddEdit("x" resultX " y300 w" resultW " r1 ReadOnly vExprResult", PTToolsWindow.Fmt2(S["ExprResult"]))
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
    ; SPF2M window - tendon profile calculator
    ; ===================================================================
    ; 以前是把参数自动敲进 DOSBox 里的 SPF2M.EXE; 现在由 TendonProfile 直接计算 (结果和
    ; SPF2M 完全一致, 见 TendonProfile.ahk), 修改任何输入都立即更新右边的表格。
    ; Resources 里还有 DOSBox.exe / SPF2M.exe 时, 可以用 "Original SPF2M" 按钮打开原程序对照。
    static ShowSpf2m() {
        if WinExist("SPF2M ahk_class AutoHotkeyGUI") {
            WinActivate()
            return
        }

        S := PTToolsWindow.Settings
        g := Gui("+AlwaysOnTop", "SPF2M")
        PTToolsWindow.SpfG := g
        g.SetFont("s9", "Microsoft YaHei")
        g.OnEvent("Close", (p*) => PTToolsWindow.OnCloseSpf2m(p*))

        right := PTToolsWindow.BuildProfileGroup(g, S, 20)
        PTToolsWindow.BuildResultGroup(g, right + 10)
        PTToolsWindow.SetupSpf2mHotkeys()
        PTToolsWindow.OnTendonChanged()                                    ; 最小半径 / 管道直径的提示跟着钢绞线类型
        g.Show("x" S["Spf2mWinLeft"] " y" S["Spf2mWinTop"] " AutoSize")
    }

    static SetupSpf2mHotkeys() {
        if PTToolsWindow.Spf2mHotkeysReady
            return
        PTToolsWindow.Spf2mHotkeysReady := true
        HotIfWinActive("SPF2M ahk_class AutoHotkeyGUI")                    ; Enter / numpad Enter = move to next field, same as PT Tools
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

    ; 每种钢绞线类型各自记住管道直径 ("At C.G." 时标高减去半个直径)
    static DuctDiaKey(tendon) {
        static keys := ["DuctDiaMono", "DuctDia7s", "DuctDia12s", "DuctDia19s", "DuctDia22s", "DuctDia31s"]
        return keys[Max(1, Min(tendon, keys.Length))]
    }

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

        g.Add("GroupBox", "x" x0 " y15 w" w " h400", "Tendon Profile")
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

        g.AddText("x" labelX " y222 w" (w - 30) " h1 0x10")                ; 分隔线: 下面是可选项, 空着 = SPF2M 的默认值
        PTToolsWindow.Field(g, labelX, editX, 235, PTToolsWindow.ProfileFieldW, "Min. Radius of Curvature (mm)", "MinRadius", S["MinRadius"]).OnEvent("Change", recalc)
        PTToolsWindow.Field(g, labelX, editX, 270, PTToolsWindow.ProfileFieldW, "Point of Contraflexure (mm)", "Contraflexure", S["Contraflexure"]).OnEvent("Change", recalc)
        PTToolsWindow.Field(g, labelX, editX, 305, fullW, "Support Intervals (mm)", "SupportIntervals", S["SupportIntervals"]).OnEvent("Change", recalc)
        Win.SetCueBanner(g["SupportIntervals"].Hwnd, "auto: max. 1000")
        PTToolsWindow.Field(g, labelX, editX, 340, PTToolsWindow.ProfileFieldW, "Duct Dia. for At C.G. (mm)", "DuctDia", "").OnEvent("Change", recalc)
        g.SetFont("s8 cGray")
        g.AddText("x" labelX " y375 w" (w - 30), "Empty fields use the SPF2M defaults (in gray). Intervals: e.g. 500, 1500, 800 ... must add up to the distance. Distances are from the high end.")
        g.SetFont("s9 cDefault")
        return x0 + w
    }

    static BuildResultGroup(g, x0) {
        w := 480
        g.Add("GroupBox", "x" x0 " y15 w" w " h400", "Tendon Profile - Support Heights (mm)")
        list := g.AddListView("x" (x0 + 12) " y40 w" (w - 24) " h240 vResultList -Multi NoSort Grid"
            , ["Distance", "Interval", "Actual", "Beam @ 5mm", "Slab @ 10mm"])
        for index, width in [80, 76, 82, 100, 100]
            list.ModifyCol(index, width " Right")
        g.AddText("x" (x0 + 12) " y290 w" (w - 24) " h60 vResultSummary")
        g.SetFont("cRed")
        g.AddText("x" (x0 + 12) " y290 w" (w - 24) " h60 vResultError Hidden")
        g.SetFont("cDefault")
        g.AddButton("x" (x0 + 12) " y365 w140 h30", "Copy Table").OnEvent("Click", (*) => PTToolsWindow.CopyProfileTable())
        if FileExist(A_ScriptDir "\Resources\DOSBox.exe") && FileExist(A_ScriptDir "\Resources\SPF2M.exe")
            g.AddButton("x+10 yp w160 h30", "Original SPF2M...").OnEvent("Click", (*) => PTToolsWindow.RunOriginalSpf2m())
    }

    ; 钢绞线类型变了: 最小半径的灰色提示换成这种类型的默认值, 管道直径换成这种类型记住的值
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
        PTToolsWindow.Recalculate()
    }

    ; 窗口里的输入 -> TendonProfile.Calc 的参数 ("At C.G." 的标高减去半个管道直径)
    static ProfileInput() {
        g := PTToolsWindow.SpfG
        halfDuct := PTToolsWindow.Num(g["DuctDia"].Value) / 2
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

    ; 标高可能带小数 ("At C.G." 减去半个管道直径): 整数照常显示, 否则保留一位小数
    static FmtLevel(v) {
        return (v = Integer(v)) ? Integer(v) : Round(v, 1)
    }

    ; 表格复制成制表符分隔的文字, 粘贴到 Excel 正好是一列一列的
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
        ToolTip("Table copied - paste it into Excel")
        SetTimer(() => ToolTip(), -1500)
    }

    ; 在 DOSBox 里打开原来的 SPF2M.EXE 并输入同样的参数, 用来对照结果 (没有 DOSBox 时不显示这个按钮)
    static RunOriginalSpf2m() {
        resDir := A_ScriptDir "\Resources"
        dosbox := resDir "\DOSBox.exe"
        input := PTToolsWindow.ProfileInput()
        S := PTToolsWindow.Settings
        keyIntervalMs := PTToolsWindow.Num(S["KeyIntervalMs"], 25)
        initDelayMs   := PTToolsWindow.Num(S["InitDelayMs"], 1500)

        DetectHiddenWindows(true)
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

        keystrokes := [input["Profile"], "{Enter}", input["Tendon"], "{Enter}"
            , input["Radius"], "{Enter}"                                   ; 空 = 默认的最小半径
            , input["Start"], "{Enter}", input["End"], "{Enter}", input["Distance"], "{Enter}"
            , input["Contraflexure"], "{Enter}"]
        if (input["Contraflexure"] != "")
            keystrokes.Push("{Enter}")                                     ; 改了反弯点时 SPF2M 显示新的半径后再问一次
        keystrokes.Push("N", "{Enter}")                                    ; 不改支架间距
        for _, keystroke in keystrokes {
            Sleep(keyIntervalMs)
            if (keystroke != "")
                SendInput((SubStr(keystroke, 1, 1) = "{") ? keystroke : "{Text}" keystroke)
        }
    }
}
