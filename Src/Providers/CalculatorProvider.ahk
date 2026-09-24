;===============================================================================
; CalculatorProvider.ahk - 计算器 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 输入算式直接显示结果 ("12*(3+4)" / "=2^10"), 最多两位小数, Enter 复制结果。
; 打开 Features.Calculator.StructuralCalc 后, 结果下方附带两行结构计算:
;   - 把结果当作梁宽 (mm): 主筋根数和间距 (保护层 40 mm, 最大间距 300 mm)
;   - 把结果当作配筋面积 As (mm²): H13 / H16 / H20 / H25 / H32 需要的根数
;===============================================================================

class CalculatorProvider {
    static Id := "Calculator"

    static Init() {
    }

    static Search(query) {
        results := []
        expression := query.Text
        if (SubStr(expression, 1, 1) = "=")
            expression := SubStr(expression, 2)
        else if !Calc.Looks(expression)
            return results
        expression := StrReplace(StrReplace(StrReplace(expression, ","), "×", "*"), "÷", "/")

        value := Calc.Eval(expression)
        if !IsNumber(value)
            return results
        text := CalculatorProvider.Format(value)

        results.Push(ResultItem(text, I18n.T("Calc.Subtitle") " · " Trim(expression), {
            Kind: "text", Arg: text, Icon: "res:imageres.dll,-182", Score: 200, LargeText: text
        }))
        if AppSettings.Feature("Calculator")["StructuralCalc"]
            CalculatorProvider._AddStructural(results, value)
        return results
    }

    ; 最多两位小数, 去掉多余的 0, 整数部分加千分位: 1234567.504 -> 1,234,567.5, 10/3 -> 3.33
    ; (浮点数直接转文字会带出 774.39999999999998 这样的尾巴)。
    ; 不到 0.005 的数按两位小数会变成 0, 改为保留到第一位有效数字后一位: 0.004, 0.000031
    static Format(value) {
        decimals := 2
        if (value != 0 && Abs(value) < 0.005)
            decimals := Min(Ceil(-Log(Abs(value))) + 1, 10)
        text := Format("{:." decimals "f}", value)
        if InStr(text, ".")
            text := RTrim(RTrim(text, "0"), ".")
        if (text = "-0")
            text := "0"
        return Calc.Thousands(text)
    }

    static _AddStructural(results, value) {
        if (value <= 80)
            return
        barCount := Ceil((value - 80) / 300 + 1)
        spacing  := Max(Round((value - 80) / (barCount - 0.999)), 0)
        beamText := I18n.T("Calc.BeamWidth", CalculatorProvider.Format(value), barCount, spacing)
        results.Push(ResultItem(beamText, "", {Kind: "text", Arg: beamText, Icon: "res:imageres.dll,-182", Score: 199}))

        bars := ""
        for bar in [[13, 132.7], [16, 201.1], [20, 314.2], [25, 490.9], [32, 804.2]]
            bars .= (bars = "" ? "" : "  ") Ceil(value / bar[2]) "H" bar[1]
        areaText := I18n.T("Calc.RebarArea", CalculatorProvider.Format(value), bars)
        results.Push(ResultItem(areaText, "", {Kind: "text", Arg: areaText, Icon: "res:imageres.dll,-182", Score: 198}))
    }
}
