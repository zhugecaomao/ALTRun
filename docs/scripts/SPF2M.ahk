; @altrun.title          SPF2M Post-Tensioning Tendon Profile Calculator
; @altrun.keyword        spf2m
; @altrun.description    Post-tensioning tendon profiles: the tendon height at each support, with the same results as the SPF2M program; the table copies to Excel
; @altrun.description_zh 后张预应力束线型: 各支架处的束高, 结果和 SPF2M 程序一致; 表格可以复制到 Excel
; @altrun.network        none
;
; Type "spf2m" in ALTRun. Changing an input updates the table at once; Enter moves to the next box, F1 opens the help
; (https://github.com/zhugecaomao/ALTRun/wiki/PT-Tools#spf2m), Copy Table copies it for Excel. The inputs and the window position
; are kept in "SPF2M.ini" next to this file.
#Requires AutoHotkey v2.0

if (A_LineFile = A_ScriptFullPath)                                          ; run as a program; ALTRun's tests #Include this file instead
    Spf2mWindow.Main()

;===============================================================================
; Spf2mWindow - the tendon profile calculator. The parameters used to be typed into SPF2M.EXE (a DOS program) in DOSBox;
; TendonProfile (below) calculates directly with identical results, and changing any input updates the table on the right at once.
;===============================================================================
class Spf2mWindow {
    static Title := "SPF2M Post-Tensioning Tendon Profile Calculator"
    static HelpSection := "spf2m"

    ; Defaults - the keys are the controls' Names; DuctDia* are per strand type (the window has one DuctDia box)
    static Defaults := Map(
        "WinLeft", 550, "WinTop", 550,
        "ProfileType", 1, "TendonType", 3,                                  ; 12S
        "StartLevel", 450, "StartAtCG", 0,
        "EndLevel", 50, "EndAtCG", 0,
        "HorizontalDistance", 12000,
        "MinRadius", "", "Contraflexure", "", "SupportIntervals", "",       ; empty = SPF2M's default

        ; Duct diameters for "At C.G.", per tendon type - empty = TendonProfile.DuctDiameters
        "DuctDiaMono", "", "DuctDia7s", "", "DuctDia12s", "",
        "DuctDia19s", "", "DuctDia22s", "", "DuctDia31s", ""
    )

    ; Window fields - round-trip through Show()/Save()
    static FieldNames := [
        "ProfileType", "TendonType", "StartLevel", "StartAtCG",
        "EndLevel", "EndAtCG", "HorizontalDistance",
        "MinRadius", "Contraflexure", "SupportIntervals"
    ]

    static ProfileTypes := ["Double Parabolic Profile", "Parabolic-Straight-Parabolic",
        "Parabolic-Straight Profile", "Straight-Parabolic Profile"]
    static TendonTypes  := ["Slab", "7S", "12S", "19S", "22S", "31S"]     ; index also selects DuctDia* above

    static G := ""                                                        ; the Gui object, while open
    static LastResult := ""                                               ; the result of the last TendonProfile.Calc (for copying the table)
    static HotkeysReady := false

    static Margin := 15                                                    ; window edge to the groups: left, right and below (the title bar is above)
    static TopY   := 3                                                     ; top of the groups
    static ColGap := 10                                                    ; between the two groups

    ; Settings: the inputs and the window position, in the .ini file next to this script ([Settings]). Only inputs are saved;
    ; results are calculated again when the window opens. Keys not in Defaults are ignored, a missing or empty one takes its default.
    static IniFile := RegExReplace(A_LineFile, "i)\.ahk$", ".ini")
    static Settings := Spf2mWindow.Defaults.Clone()
    static ExitOnClose := false                                           ; true when this file runs as a program (Main)

    ; Run as a program: the window, or the one already open (in another process) brought to the front
    static Main() {
        SetTitleMatchMode(3)                                                ; the window title exactly
        if WinExist(Spf2mWindow.Title " ahk_class AutoHotkeyGUI") {
            WinActivate()
            ExitApp()
        }
        A_IconHidden := true
        Persistent()
        Spf2mWindow.ExitOnClose := true
        Spf2mWindow.Load()
        Spf2mWindow.Show()
    }

    ; Settings from the .ini file (or the given one), each missing key from Defaults
    static Load(file := Spf2mWindow.IniFile) {
        S := Map()
        for key, value in Spf2mWindow.Defaults {
            saved := IniRead(file, "Settings", key, "")
            S[key] := (saved = "" && value != "") ? value : saved
        }
        Spf2mWindow.Settings := S
        return S
    }

    static WriteIni(file := Spf2mWindow.IniFile) {
        for key in Spf2mWindow.Defaults
            if Spf2mWindow.Settings.Has(key)
                try IniWrite(Spf2mWindow.Settings[key], file, "Settings", key)
    }

    ; F1: this tool's section on the Wiki page, the Chinese page on a Chinese system (the headings are the same in both)
    static HelpUrl(lang := A_Language) => "https://github.com/zhugecaomao/ALTRun/wiki/PT-Tools" (Spf2mWindow.IsChinese(lang) ? "-zh" : "") "#" Spf2mWindow.HelpSection
    static IsChinese(lang := A_Language) => RegExMatch(lang, "i)^(0804|0404|0C04|1004|1404)$") > 0
    static OpenHelp() {
        try Run(Spf2mWindow.HelpUrl())
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
        Spf2mWindow.Save()
        GuiObj.Destroy()
        Spf2mWindow.G := ""
        if Spf2mWindow.ExitOnClose
            ExitApp()
    }

    ; The gray hint text of an Edit box (EM_SETCUEBANNER, shown also while it has the focus)
    static SetCueBanner(hwnd, text) {
        try DllCall("User32\SendMessageW", "Ptr", hwnd, "UInt", 0x1501, "Ptr", 1, "WStr", text)
        DllCall("InvalidateRect", "Ptr", hwnd, "Ptr", 0, "Int", 1)          ; changing the cue of a box already shown needs a repaint
    }

    ; A short notice under the mouse pointer
    static Notify(text) {
        ToolTip(text)
        SetTimer(() => ToolTip(), -1500)
    }

    static Show() {
        if Spf2mWindow.G {
            WinActivate("ahk_id " Spf2mWindow.G.Hwnd)
            return
        }

        S := Spf2mWindow.Settings
        g := Gui("+AlwaysOnTop", Spf2mWindow.Title)
        Spf2mWindow.G := g
        g.SetFont("s9", "Microsoft YaHei")
        g.OnEvent("Close", (p*) => Spf2mWindow.Close(p*))

        right := Spf2mWindow.BuildProfileGroup(g, S, Spf2mWindow.Margin)
        right := Spf2mWindow.BuildResultGroup(g, right + Spf2mWindow.ColGap)
        Spf2mWindow.SetupHotkeys()
        Spf2mWindow.OnTendonChanged()                                    ; the hints for minimum radius / duct diameter follow the strand type
        g.Show("x" S["WinLeft"] " y" S["WinTop"] " w" (right + Spf2mWindow.Margin) " h" (Spf2mWindow.TopY + Spf2mWindow.GroupH + Spf2mWindow.Margin))   ; the same margins as PT Tools
    }

    static SetupHotkeys() {
        if Spf2mWindow.HotkeysReady
            return
        Spf2mWindow.HotkeysReady := true
        HotIfWinActive(Spf2mWindow.Title " ahk_class AutoHotkeyGUI")   ; Enter / numpad Enter = move to next field, same as PT Tools
        Hotkey("Enter", (*) => SendInput("{Tab}"))
        Hotkey("NumpadEnter", (*) => SendInput("{Tab}"))
        Hotkey("F1", (*) => Spf2mWindow.OpenHelp())
        HotIfWinActive()
    }

    static Save() {
        g := Spf2mWindow.G
        if !(g)
            return
        S := Spf2mWindow.Settings
        for _, name in Spf2mWindow.FieldNames {
            try S[name] := g[name].Value
        }
        S[Spf2mWindow.DuctDiaKey(g["TendonType"].Value)] := g["DuctDia"].Value
        try {
            WinGetPos(&x, &y, , , "ahk_id " g.Hwnd)
            S["WinLeft"] := x
            S["WinTop"] := y
        }
        Spf2mWindow.WriteIni()
    }

    ; each strand type remembers its own duct diameter ("At C.G." subtracts half the diameter from the levels)
    static DuctDiaKey(tendon) {
        static keys := ["DuctDiaMono", "DuctDia7s", "DuctDia12s", "DuctDia19s", "DuctDia22s", "DuctDia31s"]
        return keys[Max(1, Min(tendon, keys.Length))]
    }

    static GroupH := 420                                                   ; heights of the two GroupBoxes

    static ProfileLabelGap := 190
    static ProfileFieldW   := 90
    static ProfileCbGap    := 10
    static ProfileCbW      := 70

    static BuildProfileGroup(g, S, x0) {
        labelX := x0 + 15
        editX  := labelX + Spf2mWindow.ProfileLabelGap
        cbX    := editX + Spf2mWindow.ProfileFieldW + Spf2mWindow.ProfileCbGap
        fullW  := Spf2mWindow.ProfileFieldW + Spf2mWindow.ProfileCbGap + Spf2mWindow.ProfileCbW
        w := (cbX + Spf2mWindow.ProfileCbW + 15) - x0
        recalc := (*) => Spf2mWindow.Recalculate()
        y := Spf2mWindow.TopY

        g.Add("GroupBox", "x" x0 " y" y " w" w " h" Spf2mWindow.GroupH, "Tendon profile")
        g.AddText("x" labelX " y" (y + 30), "Profile type")
        g.AddDropDownList("x" editX " y" (y + 27) " w" fullW " Choose" S["ProfileType"] " vProfileType AltSubmit", Spf2mWindow.ProfileTypes).OnEvent("Change", recalc)
        g.AddText("x" labelX " y" (y + 65), "Tendon type")
        g.AddDropDownList("x" editX " y" (y + 62) " w" fullW " Choose" S["TendonType"] " vTendonType AltSubmit", Spf2mWindow.TendonTypes)
            .OnEvent("Change", (*) => Spf2mWindow.OnTendonChanged(true))

        Spf2mWindow.Field(g, labelX, editX, y + 97, Spf2mWindow.ProfileFieldW, "Start level (mm)", "StartLevel", S["StartLevel"]).OnEvent("Change", recalc)
        g.AddCheckBox("x" cbX " y" (y + 99) " w" Spf2mWindow.ProfileCbW " vStartAtCG Checked" S["StartAtCG"], "At C.G.").OnEvent("Click", recalc)
        Spf2mWindow.Field(g, labelX, editX, y + 132, Spf2mWindow.ProfileFieldW, "End level (mm)", "EndLevel", S["EndLevel"]).OnEvent("Change", recalc)
        g.AddCheckBox("x" cbX " y" (y + 134) " w" Spf2mWindow.ProfileCbW " vEndAtCG Checked" S["EndAtCG"], "At C.G.").OnEvent("Click", recalc)
        Spf2mWindow.Field(g, labelX, editX, y + 167, Spf2mWindow.ProfileFieldW, "Horizontal distance (mm)", "HorizontalDistance", S["HorizontalDistance"]).OnEvent("Change", recalc)

        g.AddText("x" labelX " y" (y + 207) " w" (w - 30) " h1 0x10")                ; separator: optional fields below, empty = SPF2M's default
        Spf2mWindow.Field(g, labelX, editX, y + 220, Spf2mWindow.ProfileFieldW, "Min. radius of curvature (mm)", "MinRadius", S["MinRadius"]).OnEvent("Change", recalc)
        Spf2mWindow.Field(g, labelX, editX, y + 255, Spf2mWindow.ProfileFieldW, "Point of contraflexure (mm)", "Contraflexure", S["Contraflexure"]).OnEvent("Change", recalc)
        Spf2mWindow.Field(g, labelX, editX, y + 290, fullW, "Support intervals (mm)", "SupportIntervals", S["SupportIntervals"]).OnEvent("Change", recalc)
        Spf2mWindow.SetCueBanner(g["SupportIntervals"].Hwnd, "auto: max. 1000")
        Spf2mWindow.Field(g, labelX, editX, y + 325, Spf2mWindow.ProfileFieldW, "Duct dia. for At C.G. (mm)", "DuctDia", "").OnEvent("Change", recalc)
        g.SetFont("s8 cGray")
        g.AddText("x" labelX " y" (y + 363) " w" (w - 30) " h" (Spf2mWindow.GroupH - 370), "Empty fields use the SPF2M defaults (in gray). Intervals: e.g. 500, 1500, 800 ... must add up to the distance. Distances are from the high end.")
        g.SetFont("s9 cDefault")
        return x0 + w
    }

    static BuildResultGroup(g, x0) {
        w := 480, y := Spf2mWindow.TopY
        g.Add("GroupBox", "x" x0 " y" y " w" w " h" Spf2mWindow.GroupH, "Tendon profile - support heights (mm)")
        ; laid out bottom-up: the Copy Table button at the bottom of the frame, two or three summary lines above, the rest of the height for the table
        buttonY  := y + Spf2mWindow.GroupH - 40
        summaryY := buttonY - 53
        list := g.AddListView("x" (x0 + 12) " y" (y + 25) " w" (w - 24) " h" (summaryY - y - 33) " vResultList -Multi NoSort Grid"
            , ["Distance", "Interval", "Actual", "Beam @ 5mm", "Slab @ 10mm"])
        Loop 5
            list.ModifyCol(A_Index, "Right")
        Spf2mWindow.FitColumns(list, [80, 76, 82, 100, 100])
        g.AddText("x" (x0 + 12) " y" summaryY " w" (w - 24) " h48 vResultSummary")
        g.SetFont("cRed")
        g.AddText("x" (x0 + 12) " y" summaryY " w" (w - 24) " h48 vResultError Hidden")
        g.SetFont("cDefault")
        g.AddButton("x" (x0 + 12) " y" buttonY " w140 h30", "Copy Table").OnEvent("Click", (*) => Spf2mWindow.CopyProfileTable())
        return x0 + w
    }

    ; Column widths in these proportions, filling the list's width less a vertical scroll bar: the table never needs a horizontal one
    static FitColumns(list, weights) {
        rect := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", list.Hwnd, "Ptr", rect)
        room := NumGet(rect, 8, "Int") - SysGet(2)                          ; SM_CXVSCROLL
        total := 0
        for weight in weights
            total += weight
        used := 0
        for index, weight in weights {
            width := (index = weights.Length) ? room - used : Floor(room * weight / total)
            used += width
            SendMessage(0x101E, index - 1, width, list)                     ; LVM_SETCOLUMNWIDTH, in pixels
        }
    }

    ; The strand type changed: the gray hint for minimum radius becomes this type's default, the duct diameter the value remembered for it
    static OnTendonChanged(saveOld := false) {
        g := Spf2mWindow.G
        S := Spf2mWindow.Settings
        static lastTendon := 0
        tendon := g["TendonType"].Value
        if (saveOld && lastTendon)
            S[Spf2mWindow.DuctDiaKey(lastTendon)] := g["DuctDia"].Value
        lastTendon := tendon
        g["DuctDia"].Value := S[Spf2mWindow.DuctDiaKey(tendon)]
        Spf2mWindow.SetCueBanner(g["MinRadius"].Hwnd, TendonProfile.DefaultRadius(tendon))
        Spf2mWindow.SetCueBanner(g["DuctDia"].Hwnd, TendonProfile.DefaultDuctDiameter(tendon))
        Spf2mWindow.Recalculate()
    }

    ; Window inputs -> the parameters of TendonProfile.Calc ("At C.G." subtracts half the duct diameter from the levels)
    static ProfileInput() {
        g := Spf2mWindow.G
        halfDuct := Spf2mWindow.Num(g["DuctDia"].Value, TendonProfile.DefaultDuctDiameter(g["TendonType"].Value)) / 2
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
        g := Spf2mWindow.G
        if !(g)
            return
        result := TendonProfile.Calc(Spf2mWindow.ProfileInput())
        Spf2mWindow.LastResult := result
        list := g["ResultList"]
        list.Opt("-Redraw")
        list.Delete()
        for row in result.Rows
            list.Add("", Spf2mWindow.FmtLevel(row[1]), Spf2mWindow.FmtLevel(row[2]), Spf2mWindow.FmtLevel(row[3]), Spf2mWindow.FmtLevel(row[4]), Spf2mWindow.FmtLevel(row[5]))
        list.Opt("+Redraw")
        g["ResultError"].Visible := (result.Error != "")
        g["ResultSummary"].Visible := (result.Error = "")
        g["ResultError"].Value := result.Error
        g["ResultSummary"].Value := Spf2mWindow.ProfileSummary(result)
        Spf2mWindow.SetCueBanner(g["Contraflexure"].Hwnd, (result.Error = "" && !result.RadiusChanged) ? Integer(result.Contraflexure) : "")
    }

    static ProfileSummary(result) {
        if (result.Error != "")
            return result.Error
        text := "Min. radius of curvature = " Integer(result.MinRadius) " mm"
        if (result.RadiusChanged || result.Radius != result.MinRadius)
            text .= "`nRadius of curvature = " Integer(result.Radius) " mm"
        text .= "`nDist. to point of contraflexure = " Integer(result.Contraflexure) " mm (from the high end)"
        return text
    }

    ; levels may have decimals ("At C.G." minus half the duct diameter): integers as usual, otherwise one decimal
    static FmtLevel(v) {
        return (v = Integer(v)) ? Integer(v) : Round(v, 1)
    }

    ; The table copied as tab-separated text, which pastes into Excel column by column
    static CopyProfileTable() {
        result := Spf2mWindow.LastResult
        if !IsObject(result) || result.Error != ""
            return
        g := Spf2mWindow.G
        text := Spf2mWindow.ProfileTypes[g["ProfileType"].Value] " / " Spf2mWindow.TendonTypes[g["TendonType"].Value] "`r`n"
            . "Distance`tInterval`tActual`tBeam @ 5mm`tSlab @ 10mm`r`n"
        for row in result.Rows
            text .= Spf2mWindow.FmtLevel(row[1]) "`t" Spf2mWindow.FmtLevel(row[2]) "`t" Spf2mWindow.FmtLevel(row[3]) "`t" Spf2mWindow.FmtLevel(row[4]) "`t" Spf2mWindow.FmtLevel(row[5]) "`r`n"
        text .= StrReplace(Spf2mWindow.ProfileSummary(result), "`n", "`r`n") "`r`n"
        A_Clipboard := text
        Spf2mWindow.Notify(Spf2mWindow.IsChinese() ? "表格已复制, 可以粘贴到 Excel" : "Table copied - paste it into Excel")
    }
}

;===============================================================================
; TendonProfile - post-tensioning tendon profile calculation with results identical to SPF2M (AutoHotkey v2)
;-------------------------------------------------------------------------------
; SPF2M.EXE is the company's in-house 16-bit DOS program written in Turbo Basic in 1993, which used to need DOSBox
; to run (no longer shipped with the program). This calculates directly with the same formulas: they come from PT-Profile-VLX, an AutoCAD tool of the same series
; (POBLIC.lsp), checked one by one against 139 result sets produced by SPF2M itself (Tests\Fixtures\SPF2M-Reference.json in ALTRun's repository,
; over 1,100 values, all identical).
;
; Coordinates: distances are measured from the high end (SPF2M's tables always start at the high point), levels are the tendon height (mm).
;   1 double parabola              a parabola on the low side + a reverse parabola on the high side, tangent at the contraflexure point
;   2 parabola-straight-parabola   parabolas at both ends, straight in the middle
;   3 parabola-straight            a parabola on the high side, straight elsewhere
;   4 straight-parabola            straight on the high side, a parabola on the low side
; The contraflexure distance is measured from the high point; by default it follows from the minimum radius of curvature, and when given the radius is computed from it.
; Support spacing at most 1000 mm; when the span isn't whole meters, the remainder goes on the high side: over 599 it is a segment of its own, otherwise it is
; merged with one 1000 and split into two segments (rounded to 100). The segments can also be given by hand.
; The ends show the input levels; points in between: Actual truncated, Beam rounded to 5 mm, Slab to 10 mm.
;
; Usage:
;   result := TendonProfile.Calc(Map("Profile", 1, "Tendon", 3, "Start", 600, "End", 100, "Distance", 8000))
;   optional: "Radius" minimum radius of curvature (default by strand type), "Contraflexure" contraflexure distance, "Intervals" an array of support spacings
;   result.Rows  [[distance, spacing, Actual, Beam, Slab], ...]    result.Error  "" or an error description
;   result.MinRadius / result.Radius (the radius actually used) / result.Contraflexure
;===============================================================================

class TendonProfile {
    static Profiles := ["Double Parabolic", "Parabolic-Straight-Parabolic", "Parabolic-Straight", "Straight-Parabolic"]
    static Tendons  := ["Slab", "7S", "12S", "19S", "22S", "31S"]
    static Radii    := [5000, 3200, 4200, 5300, 5700, 6700]                 ; minimum radius of curvature per strand type (mm), as in SPF2M
    static DuctDiameters := [25, 70, 90, 100, 120, 130]                     ; duct diameter (mm); "At C.G." subtracts half of it from the levels
    static MaxInterval := 1000

    static DefaultRadius(tendon) {
        return (IsInteger(tendon) && tendon >= 1 && tendon <= TendonProfile.Radii.Length) ? TendonProfile.Radii[tendon] : TendonProfile.Radii[1]
    }

    static DefaultDuctDiameter(tendon) {
        return (IsInteger(tendon) && tendon >= 1 && tendon <= TendonProfile.DuctDiameters.Length) ? TendonProfile.DuctDiameters[tendon] : TendonProfile.DuctDiameters[1]
    }

    static Calc(input) {
        result := {Rows: [], Error: "", MinRadius: 0, Radius: 0, Contraflexure: 0, RadiusChanged: false}
        get := (key) => (input.Has(key) && input[key] != "") ? input[key] : ""
        profile := get("Profile"), tendon := get("Tendon")
        start := get("Start"), finish := get("End"), l := get("Distance")
        if !(IsInteger(profile) && profile >= 1 && profile <= 4)
            return TendonProfile._Fail(result, "Choose a profile type.")
        if !(IsNumber(start) && IsNumber(finish))
            return TendonProfile._Fail(result, "Enter the start and end levels.")
        if !(IsNumber(l) && l > 0)
            return TendonProfile._Fail(result, "Enter the horizontal distance.")
        minRadius := IsNumber(get("Radius")) ? get("Radius") + 0 : TendonProfile.DefaultRadius(tendon)
        if (minRadius <= 0)
            return TendonProfile._Fail(result, "The radius of curvature must be greater than 0.")
        lo := Min(start, finish) + 0, hi := Max(start, finish) + 0, ly := hi - lo, l += 0
        result.MinRadius := minRadius

        ; contraflexure point (from the high point) and the radius actually used
        rad := minRadius
        contra := TendonProfile.ContraflexureFor(profile, rad, ly, l)
        wanted := get("Contraflexure")
        if IsNumber(wanted) && (wanted + 0 != Integer(contra)) {             ; the same as the default = unchanged
            wanted += 0
            if (ly = 0)
                return TendonProfile._Fail(result, "Start and end levels are equal, the profile is straight - leave the point of contraflexure empty.")
            if (wanted <= 0 || !TendonProfile._ContraflexureFits(profile, wanted, l))
                return TendonProfile._Fail(result, "The point of contraflexure must lie within the span"
                    . ((profile <= 2) ? " (at most half of it)." : "."))
            rad := TendonProfile.RadiusFor(profile, wanted, ly, l)
            contra := wanted
            result.RadiusChanged := true
        }
        if (contra = "" || rad < minRadius || (profile = 1 && contra > l / 2))
            return TendonProfile._Fail(result, "MIN. RADIUS OF CURVATURE NOT ACHIEVABLE - increase the distance or reduce the level difference.")
        result.Radius := rad
        result.Contraflexure := contra

        intervals := get("Intervals")
        if (intervals = "")
            intervals := TendonProfile.DefaultIntervals(l)
        else if (msg := TendonProfile._CheckIntervals(intervals, l))
            return TendonProfile._Fail(result, msg)

        rows := [[0, 0, hi, hi, hi]]                                        ; the ends show the input levels
        s := 0
        for index, interval in intervals {
            s += interval
            if (index = intervals.Length) {
                rows.Push([l, interval, lo, lo, lo])
                break
            }
            y := TendonProfile.Height(profile, rad, lo, hi, l, s)
            rows.Push([s, interval, Integer(y), Integer((y + 2.5) / 5) * 5, Integer((y + 5) / 10) * 10])
        }
        result.Rows := rows
        return result
    }

    ; The default contraflexure distance (from the high point); "" when the minimum radius can't be met
    static ContraflexureFor(profile, rad, ly, l) {
        switch profile {
            case 1: return 2 * rad * ly / l
            case 2:
                disc := l * l - 4 * rad * ly
                return (disc < 0) ? "" : (l - Sqrt(disc)) / 2
            default:
                disc := l * l - 2 * rad * ly
                return (disc < 0) ? "" : l - Sqrt(disc)
        }
    }

    ; The radius of curvature for a given contraflexure distance (the inverse of ContraflexureFor)
    static RadiusFor(profile, contra, ly, l) {
        switch profile {
            case 1: return contra * l / (2 * ly)
            case 2: return contra * (l - contra) / ly
            default: return contra * (2 * l - contra) / (2 * ly)
        }
    }

    static _ContraflexureFits(profile, contra, l) {
        return (profile <= 2) ? contra <= l / 2 : contra < l
    }

    ; Tendon height at s from the high point
    static Height(profile, rad, lo, hi, l, s) {
        ly := hi - lo
        switch profile {
            case 1:                                                         ; double parabola: t measured from the low point
                t := l - s
                x := l - 2 * rad * ly / l
                if (t <= x)
                    return lo + (l - x) / (2 * rad * x) * t * t
                return lo + l * t / rad + ly - l * l / (2 * rad) - t * t / (2 * rad)
            case 2:                                                         ; parabola-straight-parabola
                x1 := l / 2 - Sqrt(l * l - 4 * rad * ly) / 2
                if (s < x1)
                    return hi - s * s / (2 * rad)
                if (s > l - x1)
                    return lo + (s - l) ** 2 / (2 * rad)
                return hi + x1 * x1 / (2 * rad) - x1 * s / rad
            case 3:                                                         ; parabola-straight
                x := l - Sqrt(l * l - 2 * rad * ly)
                if (s <= x)
                    return hi - s * s / (2 * rad)
                return lo + l * x / rad - x * s / rad
            default:                                                        ; straight-parabola
                x1 := Sqrt(l * l - 2 * rad * ly)
                if (s < x1)
                    return hi + s * (x1 - l) / rad
                return lo + (s - l) ** 2 / (2 * rad)
        }
    }

    ; Support spacings of at most 1000 mm, with the remainder on the high side (as in SPF2M)
    static DefaultIntervals(l) {
        step := TendonProfile.MaxInterval
        whole := Integer(l / step)
        rest := l - whole * step
        intervals := []
        if (rest = 0) {
            Loop whole
                intervals.Push(step)
        } else if (rest > 599 || whole = 0) {
            intervals.Push(TendonProfile._Clean(rest))
            Loop whole
                intervals.Push(step)
        } else {                                                            ; the remainder is too short: merge it with one 1000 and split into two segments
            total := rest + step
            second := Integer(total / 200) * 100
            first := total - second
            if (Abs(second - first) >= 100) {
                second := (Integer(total / 200) + 1) * 100
                first := total - second
            }
            intervals.Push(TendonProfile._Clean(first), second)
            Loop whole - 1
                intervals.Push(step)
        }
        return intervals
    }

    ; Text like "500, 1500, 800" -> an array; "" when empty (= automatic)
    static ParseIntervals(text) {
        text := Trim(text)
        if (text = "")
            return ""
        list := []
        for part in StrSplit(RegExReplace(text, "[\s,;/]+", ","), ",") {
            if (part = "")
                continue
            if !IsNumber(part)
                return part                                                 ; not numbers: returned as they are, and _CheckIntervals reports the error
            list.Push(part + 0)
        }
        return list
    }

    static _CheckIntervals(intervals, l) {
        if !(intervals is Array)
            return "Support intervals: '" intervals "' is not a number."
        total := 0
        for interval in intervals {
            if !(IsNumber(interval) && interval > 0)
                return "Support intervals must be greater than 0."
            total += interval
        }
        if (Abs(total - l) > 0.5)
            return "Support intervals add up to " TendonProfile._Clean(total) " mm, not the horizontal distance " TendonProfile._Clean(l) " mm."
        return ""
    }

    static _Clean(v) {
        return (v = Integer(v)) ? Integer(v) : Round(v, 1)
    }

    static _Fail(result, message) {
        result.Error := message
        return result
    }
}
