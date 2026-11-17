;===============================================================================
; PTToolsTests.ahk - the script library's PT Tools.ahk and SPF2M.ahk (docs\scripts): settings in their .ini iniFile, help pages,
;   the two windows, the tendon profile calculation against the original SPF2M.EXE, PTCalc against ALTRun's Calc
;-------------------------------------------------------------------------------
; Included by Tests\RunTests.ahk, which also #Includes both scripts: they open their window only when run as a program.
; Every static method whose name doesn't start with "_" is a test; each one must put back the shared state it changes.
;===============================================================================
class PTToolsTests {
    ; Settings: read from and written to the .ini file next to the script (a temporary one here); a missing key takes its default,
    ; keys no longer used are ignored. F1 opens the tool's section on the Wiki, the Chinese page on a Chinese system
    static PTToolsSettings() {
        eq := (n, a, e) => TestRunner.Equal("PTToolsSettings." n, a, e)
        iniFile := A_Temp "\ALTRun-pttools-test.ini"
        saved := [PTToolsWindow.Settings, Spf2mWindow.Settings]
        try {
            eq("ini next to the script", RegExMatch(PTToolsWindow.IniFile, "i)\\docs\\scripts\\PT Tools\.ini$") > 0 && RegExMatch(Spf2mWindow.IniFile, "i)\\docs\\scripts\\SPF2M\.ini$") > 0, true)
            eq("not run as a program when included", PTToolsWindow.ExitOnClose || Spf2mWindow.ExitOnClose, false)
            try FileDelete(iniFile)
            S := PTToolsWindow.Load(iniFile)
            eq("defaults without a file", S["RebarSpanWidth"] "|" S["MeshTopMark"] "|" S["SafetyFactorEnabled"] "|" S["WinLeft"], "3000|A7|0|500")
            IniWrite("5700", iniFile, "Settings", "RebarSpanWidth"), IniWrite("3800", iniFile, "Settings", "BrcSpanWidth")
            S := PTToolsWindow.Load(iniFile)
            eq("saved value read", S["RebarSpanWidth"], "5700")
            eq("key no longer used ignored", S.Has("BrcSpanWidth"), false)
            S["ExprInput"] := "2^10", S["WinLeft"] := 123
            PTToolsWindow.WriteIni(iniFile)
            S := PTToolsWindow.Load(iniFile)
            eq("round trip", S["ExprInput"] "|" S["WinLeft"] "|" S["RebarSpanWidth"], "2^10|123|5700")
            stray := ""
            for name in PTToolsWindow.FieldNames
                stray .= PTToolsWindow.Defaults.Has(name) ? "" : name " "
            for name in Spf2mWindow.FieldNames
                stray .= Spf2mWindow.Defaults.Has(name) ? "" : name " "
            eq("every saved field has a default", stray, "")

            try FileDelete(iniFile)                                            ; SPF2M: optional fields stay empty, a duct diameter per strand type
            S := Spf2mWindow.Load(iniFile)
            eq("SPF2M defaults", S["TendonType"] "|" S["MinRadius"] "|" S["DuctDia12s"] "|" S["WinLeft"], "3|||550")
            S["MinRadius"] := 4500, S["DuctDia12s"] := 80
            Spf2mWindow.WriteIni(iniFile)
            S := Spf2mWindow.Load(iniFile)
            eq("SPF2M round trip", S["MinRadius"] "|" S["DuctDia12s"] "|" S["DuctDia7s"], "4500|80|")

            eq("help: English page", PTToolsWindow.HelpUrl("0409"), HelpProvider.WikiUrl "PT-Tools#pt-tools")
            eq("help: Chinese page", Spf2mWindow.HelpUrl("0804"), HelpProvider.WikiUrl "PT-Tools-zh#spf2m")
            eq("help: Traditional Chinese", Spf2mWindow.IsChinese("0404") && !Spf2mWindow.IsChinese("0411"), true)
            for page in ["PT-Tools", "PT-Tools-zh"] {                        ; the sections F1 opens exist on both pages
                text := FileRead(A_ScriptDir "\..\docs\wiki\" page ".md", "UTF-8")
                eq("help sections on " page, RegExMatch(text, "m)^## PT Tools$") > 0 && RegExMatch(text, "m)^## SPF2M$") > 0, true)
            }
        } finally {
            PTToolsWindow.Settings := saved[1], Spf2mWindow.Settings := saved[2]
            try FileDelete(iniFile)
        }
    }

    ; PTCalc is a copy of Calc (Lib\Util.ahk), so PT Tools.ahk runs on its own: the same code apart from the class name, the same results
    static PTCalcSameAsCalc() {
        body := (text, start, finish) => (i := InStr(text, start)) && (j := InStr(text, finish, , i)) ? SubStr(text, i, j - i) : "(not found)"
        start := "    `; Returns the result (a float)"
        ours := body(FileRead(A_ScriptDir "\..\Lib\Util.ahk", "UTF-8"), start, "    `; Thousands separators")
        theirs := StrReplace(body(FileRead(A_ScriptDir "\..\docs\scripts\PT Tools.ahk", "UTF-8"), start, "`n}"), "PTCalc.", "Calc.") "`n`n"
        TestRunner.Equal("PTCalcSameAsCalc.code", StrReplace(theirs, "`r"), StrReplace(ours, "`r"))
        different := ""
        for expr in ["1+2*3", "(1+2)*3", "2^3^2", "-2^2", "2^-2", "2**10", "10/4", "1/0", "(1+2))", "1+", "", "abc", "3.1416*16^2/4"
                   , "(6000/200+1)*201.06", "1.5.2", "--3", "9^999", ".5+.5", " 1 + 1 "]
            if (PTCalc.Eval(expr) "" != Calc.Eval(expr) "")
                different .= "[" expr "] "
        TestRunner.Equal("PTCalcSameAsCalc.results", different, "")
    }

    ; PT Tools window: result boxes take no Tab stop (from Area expression straight to Bar diameter), units written as mm², even margins,
    ; inputs are combo boxes with their common values, the span width follows into the mesh area
    static PTToolsWindowUi() {
        eq := (n, a, e) => TestRunner.Equal("PTToolsWindowUi." n, a, e)
        saved := [PTToolsWindow.Settings["WinLeft"], PTToolsWindow.Settings["WinTop"]]
        PTToolsWindow.Settings["WinLeft"] := -3000, PTToolsWindow.Settings["WinTop"] := -3000
        PTToolsWindow.Show()
        g := PTToolsWindow.G
        tabbed := "", oldUnit := ""
        for ctrl in g {
            style := ControlGetStyle(ctrl)
            if (ctrl.Type = "Edit" && (style & 0x800) && (style & 0x10000))     ; ES_READONLY still with WS_TABSTOP
                tabbed .= ctrl.Name " "
            if InStr(ctrl.Text, "mm2")
                oldUnit .= ctrl.Text " | "
        }
        eq("even margins", PTToolsTests._EvenMargins(g), true)
        eq("results calculated on opening", g["RebarBarCount"].Value != "" && g["RequiredBarCount"].Value != "" && g["MeshTopCombinedArea"].Value != "" && g["ExprResult"].Value != "", true)
        eq("read-only fields skip tab", tabbed, "")
        eq("units", oldUnit, "")
        next := DllCall("GetNextDlgTabItem", "Ptr", g.Hwnd, "Ptr", g["RequiredAreaExpression"].Hwnd, "Int", 0, "Ptr")
        eq("expression -> diameter", next = g["RequiredBarDiameter"].Hwnd, true)
        editable := ""                                                      ; every input is a combo box (native: the first click selects all); Edit only for results
        for ctrl in g
            if (ctrl.Type = "Edit" && !(ControlGetStyle(ctrl) & 0x800))
                editable .= ctrl.Name " "
        eq("inputs are combo boxes", editable, "")
        types := ""
        for name, list in PTToolsWindow.ComboFields
            types .= (g[name].Type = "ComboBox" && g[name].Text = PTToolsWindow.Settings[name]) ? "" : name " "
        eq("combo boxes with the saved text", types, "")
        spanEdit := DllCall("FindWindowEx", "Ptr", g["RebarSpanWidth"].Hwnd, "Ptr", 0, "Str", "Edit", "Ptr", 0, "Ptr")
        eq("span width digits only", (ControlGetStyle(spanEdit) & 0x2000) != 0, true)
        meshEdit := DllCall("FindWindowEx", "Ptr", g["MeshTopMark"].Hwnd, "Ptr", 0, "Str", "Edit", "Ptr", 0, "Ptr")
        eq("mesh mark in capitals", (ControlGetStyle(meshEdit) & 0x8) != 0, true)
        g["RebarSpanWidth"].Text := "5700"                                  ; a list item: .Value would be its index, the handlers read .Text
        g["RebarBarSpacing"].Text := "150"
        PTToolsWindow.OnRebarFieldChanged()
        eq("bar count from the combo texts", g["RebarBarCount"].Value, "38.00")
        eq("mesh span follows", g["MeshSpanWidth"].Text, "5700")
        eq("top mesh area recalculated", g["MeshTopArea"].Value, PTToolsWindow.Fmt2(5.7 * PTToolsWindow.CalculateMeshArea(g["MeshTopMark"].Text)))
        eq("bottom mesh area recalculated", g["MeshBotArea"].Value, PTToolsWindow.Fmt2(5.7 * PTToolsWindow.CalculateMeshArea(g["MeshBotMark"].Text)))
        g["RequiredAreaExpression"].Text := "5000"
        PTToolsWindow.OnRequiredExpressionChanged()
        eq("area expression from the combo text", g["RequiredTotalArea"].Value, "5000.00")
        g["ExprInput"].Text := "(6000/200+1)*201.06"
        PTToolsWindow.OnExpressionChanged()
        eq("calculator from the combo text", g["ExprResult"].Value, "6232.86")
        examples := ""
        for expr in PTToolsWindow.Expressions
            examples .= IsNumber(PTCalc.Eval(expr)) ? "" : expr " "
        eq("calculator examples all work", examples, "")
        safety := g["SafetyFactorEnabled"]
        eq("safety factor checkbox skips tab", (ControlGetStyle(safety) & 0x10000) != 0, false)
        ; Numbers in the bar diameter combo box are right-aligned: the first character's position (EM_POSFROMCHAR) is in the right half of the box (Wine doesn't support changing alignment after creation, so only checked on Windows)
        diaEdit := DllCall("FindWindowEx", "Ptr", g["RebarBarDiameter"].Hwnd, "Ptr", 0, "Str", "Edit", "Ptr", 0, "Ptr")
        eq("diameter diaEdit has ES_RIGHT", (ControlGetStyle(diaEdit) & 0x2) != 0, true)
        if !DllCall("GetProcAddress", "Ptr", DllCall("GetModuleHandle", "Str", "ntdll", "Ptr"), "AStr", "wine_get_version", "Ptr") {
            WinGetClientPos(, , &ew, , diaEdit)
            firstX := SendMessage(0xD6, 0, 0, diaEdit) & 0xFFFF                     ; EM_POSFROMCHAR: x of character 0
            TestRunner.True("PTToolsWindowUi.diameter text on the right (first char at " firstX " of " ew ")", firstX > ew // 2)
        }
        g.Destroy(), PTToolsWindow.G := ""
        PTToolsWindow.Settings["WinLeft"] := saved[1], PTToolsWindow.Settings["WinTop"] := saved[2]
    }

    ; SPF2M window: the same margins as PT Tools; the table's columns leave room for a vertical scroll bar, no horizontal one
    static Spf2mWindowUi() {
        eq := (n, a, e) => TestRunner.Equal("Spf2mWindowUi." n, a, e)
        saved := [Spf2mWindow.Settings["WinLeft"], Spf2mWindow.Settings["WinTop"]]
        Spf2mWindow.Settings["WinLeft"] := -3000, Spf2mWindow.Settings["WinTop"] := -3000
        Spf2mWindow.Show()
        eq("even margins", PTToolsTests._EvenMargins(Spf2mWindow.G), true)
        eq("table filled on opening", Spf2mWindow.G["ResultList"].GetCount() > 0, true)
        list := Spf2mWindow.G["ResultList"], columns := 0
        Loop 5
            columns += SendMessage(0x101D, A_Index - 1, 0, list)            ; LVM_GETCOLUMNWIDTH
        rect := Buffer(16, 0), DllCall("GetClientRect", "Ptr", list.Hwnd, "Ptr", rect)
        eq("table fits", columns <= NumGet(rect, 8, "Int") && !(ControlGetStyle(list) & 0x100000), true)   ; WS_HSCROLL
        Spf2mWindow.G.Destroy(), Spf2mWindow.G := ""
        Spf2mWindow.Settings["WinLeft"] := saved[1], Spf2mWindow.Settings["WinTop"] := saved[2]
    }

    ; The same margin left, right and below the group boxes (the title bar is above)
    static _EvenMargins(g) {
        left := 99999, right := 0, bottom := 0
        for ctrl in g
            if (ctrl.Type = "GroupBox") {
                ControlGetPos(&cx, &cy, &cw, &ch, ctrl)
                left := Min(left, cx), right := Max(right, cx + cw), bottom := Max(bottom, cy + ch)
            }
        WinGetClientPos(, , &clientW, &clientH, g)
        return Abs(clientW - right - left) <= 1 && Abs(clientH - bottom - left) <= 1
    }

    ; Input checks, support spacing, cases SPF2M didn't handle well (crashes / shows nothing)
    static TendonProfileInputs() {
        eq := (n, a, e) => TestRunner.Equal("TendonProfileInputs." n, a, e)
        profileOf := (profile, start, finish, l, extra := "") => TendonProfile.Calc(TendonProfileInputsMap(profile, start, finish, l, extra))
        TendonProfileInputsMap(profile, start, finish, l, extra) {
            input := Map("Profile", profile, "Tendon", 3, "Start", start, "End", finish, "Distance", l)
            if IsObject(extra)
                for key, value in extra
                    input[key] := value
            return input
        }
        intervals := (l) => TendonProfileJoin(TendonProfile.DefaultIntervals(l))
        TendonProfileJoin(list) {
            text := ""
            for v in list
                text .= (text = "" ? "" : ",") v
            return text
        }
        eq("intervals whole metres", intervals(8000), "1000,1000,1000,1000,1000,1000,1000,1000")
        eq("intervals single split", intervals(9750), "750,1000,1000,1000,1000,1000,1000,1000,1000,1000")
        eq("intervals double split", intervals(7350), "650,700,1000,1000,1000,1000,1000,1000")
        eq("intervals double split 2", intervals(8420), "720,700,1000,1000,1000,1000,1000,1000,1000")
        eq("intervals short span", intervals(450), "450")
        eq("intervals 1300", intervals(1300), "600,700")
        eq("parse intervals", TendonProfileJoin(TendonProfile.ParseIntervals("500, 1500 800;550")), "500,1500,800,550")
        eq("parse empty = auto", TendonProfile.ParseIntervals("  "), "")
        eq("intervals must add up", InStr(profileOf(1, 600, 100, 8000, Map("Intervals", [1000, 1000])).Error, "add up") > 0, true)
        eq("interval not a number", profileOf(1, 600, 100, 8000, Map("Intervals", TendonProfile.ParseIntervals("1000 abc"))).Error != "", true)
        eq("missing distance", profileOf(1, 600, 100, "").Error != "", true)
        eq("missing level", profileOf(1, "", 100, 8000).Error != "", true)
        ; where SPF2M crashes: an error description here
        eq("psp not achievable", InStr(profileOf(2, 900, 100, 3000, Map("Tendon", 6)).Error, "NOT ACHIEVABLE") > 0, true)
        ; where SPF2M shows nothing: the contraflexure point beyond half the span
        eq("contraflexure beyond half span", InStr(profileOf(1, 600, 100, 9000, Map("Contraflexure", 5000)).Error, "half") > 0, true)
        eq("contraflexure with equal levels", profileOf(1, 300, 300, 6000, Map("Contraflexure", 500)).Error != "", true)
        result := profileOf(1, 600, 100, 9000, Map("Contraflexure", 1200))
        eq("contraflexure -> radius", Integer(result.Radius) "|" result.RadiusChanged, "10800|1")
        result := profileOf(1, 600, 100, 8000, Map("Contraflexure", 525))
        eq("contraflexure equal to default keeps radius", result.Radius "|" result.RadiusChanged, "4200|0")
        eq("default radius 22s", TendonProfile.DefaultRadius(5), 5700)
        eq("default duct dia", TendonProfile.DefaultDuctDiameter(1) "|" TendonProfile.DefaultDuctDiameter(3) "|" TendonProfile.DefaultDuctDiameter(6), "25|90|130")
        eq("default tendon 12S", Spf2mWindow.Defaults["TendonType"], 3)
    }

    ; The tendon profile calculation compared value by value with the original SPF2M.EXE: Tests\Fixtures\SPF2M-Reference.json holds the results of running
    ; SPF2M in DOSBox (4 profile types x 6 strand types, rising / falling, spans that aren't whole meters, custom radius / contraflexure /
    ; support spacing, and the cases where SPF2M reports an error or crashes); how it was made: see Tests\Tools\SPF2M
    static TendonProfileVsSpf2m() {
        data := JSON.Parse(FileRead(A_ScriptDir "\Fixtures\SPF2M-Reference.json", "UTF-8"))
        checked := 0, failed := 0
        for ref in data["Cases"] {
            input := Map()
            for key in ["Profile", "Tendon", "Start", "End", "Distance", "Radius", "Contraflexure", "Intervals"]
                if ref.Has(key)
                    input[key] := ref[key]
            result := TendonProfile.Calc(input)
            name := "TendonProfile." ref["Profile"] "/" ref["Tendon"] " " ref["Start"] "->" ref["End"] " L" ref["Distance"]
                . (ref.Has("Radius") ? " R" ref["Radius"] : "") . (ref.Has("Contraflexure") ? " C" ref["Contraflexure"] : "")
            if ref.Has("Expect") {
                TestRunner.True(name " error", result.Error != "")
                continue
            }
            if (result.Error != "") {
                TestRunner.Fail(name, "unexpected error: " result.Error)
                continue
            }
            got := "", want := ""
            for row in result.Rows
                got .= row[1] "," row[2] "," row[3] "," row[4] "," row[5] ";"
            for row in ref["Rows"]
                want .= row[1] "," row[2] "," row[3] "," row[4] "," row[5] ";"
            if (got != want) {
                failed += 1
                TestRunner.Equal(name " rows", got, want)
            }
            checked += ref["Rows"].Length
            TestRunner.Equal(name " contraflexure", Integer(result.Contraflexure), ref["ShownContraflexure"])
            if ref.Has("ShownRadius")
                TestRunner.Equal(name " radius", Integer(result.Radius), ref["ShownRadius"])
        }
        TestRunner.True("TendonProfile.rows checked (" checked ")", checked > 1000 && !failed)
    }
}
