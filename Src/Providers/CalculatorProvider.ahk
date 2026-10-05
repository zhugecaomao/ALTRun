;===============================================================================
; CalculatorProvider.ahk - 计算器 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 输入算式直接显示结果 ("12*(3+4)" / "=2^10"), 最多两位小数, Enter 复制结果。
; 打开 Features.Calculator.StructuralCalc 后, 结果下方附带两行结构计算:
;   - 把结果当作梁宽 (mm): 主筋根数和间距 (梁边到主筋中心 BarEdge, 默认 40 mm; 最大间距 MaxBarSpacing, 默认 300 mm)
;   - 把结果当作配筋面积 As (mm²): BarSizes 里每种直径需要的根数 (默认 H13 / H16 / H20 / H25 / H32, 前缀 BarPrefix)
; 单位换算 (Lib\Units.ahk): "10 km in mi"、"100 f to c"、"20 mpa in psi"...; 货币换算 "100 usd to sgd"
; 要打开 Features.Calculator.Currency (默认关闭, 汇率每天从 frankfurter.dev 下载, 见 CurrencyRates)。
; 进制: "0xFF" / "0b1010" / "0o17" 显示十进制、十六进制、二进制; "255 in hex" / "0xff to dec" / "10 in bin" / "8 in oct"。
; 日期: "today + 30 days" / "2026-12-25 - 2 weeks" / "today + 3 months" (单位 d w m y, 也可以写 天 周 月 年);
;       "2026-12-25 - today" 两个日期相差几天。今天也可以写 now / 今天。
;===============================================================================

class CalculatorProvider {
    static Id := "Calculator"

    static Init() {
        if AppSettings.Feature("Calculator")["Currency"]
            CurrencyRates.Start()
    }

    static Search(query) {
        results := []
        if (bases := CalculatorProvider._Bases(query.Text)).Length
            return bases
        if (dates := CalculatorProvider._Dates(query.Text)).Length
            return dates
        if IsObject(conversion := Units.Parse(query.Text))
            return CalculatorProvider._Convert(conversion)
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

    ; "10 km in mi" -> 一条结果 "6.21 mi" (Enter 复制数字); 货币还没有汇率时给出提示
    static _Convert(conversion) {
        icon := "res:imageres.dll,-182"
        value := Units.Convert(conversion.Value, conversion.From, conversion.To)
        if IsNumber(value) {
            text := CalculatorProvider.Format(value)
            label := Units.Label(conversion.To)
            subtitle := CalculatorProvider.Format(conversion.Value) " " Units.Label(conversion.From) " = " text " " label
            if (Units._Lookup(conversion.To).Kind = "currency")
                subtitle .= " · " I18n.T("Calc.RatesOf", CurrencyRates.Date)
            return [ResultItem(text " " label, subtitle, {Kind: "text", Arg: text, Icon: icon, Score: 200, LargeText: text " " label})]
        }
        ; 两边都像货币代码 (3 个字母) 又不是单位: 提示打开货币换算 / 还没有下载汇率
        if (RegExMatch(conversion.From, "^[a-z]{3}$") && RegExMatch(conversion.To, "^[a-z]{3}$")
                && !IsObject(Units._Lookup(conversion.From)) && !IsObject(Units._Lookup(conversion.To))) {
            hint := AppSettings.Feature("Calculator")["Currency"] ? "Calc.RatesNotYet" : "Calc.CurrencyOff"
            return [ResultItem(I18n.T("Calc.Currency"), I18n.T(hint), {Icon: icon, Score: 150, Valid: false})]
        }
        return []
    }

    static _Item(text, subtitle, score := 200) {
        return ResultItem(text, subtitle, {Kind: "text", Arg: text, Icon: "res:imageres.dll,-182", Score: score, LargeText: text})
    }

    ; 进制换算: 带前缀的数 (0x / 0b / 0o) 单独输入时列出三种写法; "数 in hex|bin|oct|dec" 换成指定的进制
    static _Bases(text) {
        static names := Map("hex", 16, "bin", 2, "oct", 8, "dec", 10)
        if !RegExMatch(Trim(text), "i)^(0x[0-9a-f]+|0b[01]+|0o[0-7]+|\d+)(?:\s+(?:in|to)\s+(hex|bin|oct|dec))?$", &m)
            return []
        if (m[2] = "" && !RegExMatch(m[1], "i)^0[xbo]"))                   ; 普通数字不当作进制换算
            return []
        value := CalculatorProvider.ParseInteger(m[1])
        if (value = "")
            return []
        subtitle := I18n.T("Calc.Subtitle") " · " Trim(text)
        if (m[2] != "")
            return [CalculatorProvider._Item(CalculatorProvider.ToBase(value, names[StrLower(m[2])]), subtitle)]
        return [CalculatorProvider._Item(String(value), subtitle, 200)
              , CalculatorProvider._Item("0x" CalculatorProvider.ToBase(value, 16), subtitle, 199)
              , CalculatorProvider._Item("0b" CalculatorProvider.ToBase(value, 2), subtitle, 198)]
    }

    ; "0xff" / "0b101" / "0o17" / "255" -> 整数 (太大时 "")
    static ParseInteger(text) {
        base := 10, digits := text
        if RegExMatch(text, "i)^0([xbo])(.+)$", &m)
            base := (StrLower(m[1]) = "x") ? 16 : (StrLower(m[1]) = "b") ? 2 : 8, digits := m[2]
        if (StrLen(digits) > ((base = 2) ? 62 : (base = 8) ? 20 : 15))     ; 不超过 64 位整数
            return ""
        value := 0
        Loop Parse, StrLower(digits)
            value := value * base + InStr("0123456789abcdef", A_LoopField) - 1
        return value
    }

    static ToBase(value, base) {
        if (base = 10)
            return String(value)
        if (value = 0)
            return "0"
        digits := ""
        while (value > 0) {
            digits := SubStr("0123456789ABCDEF", Mod(value, base) + 1, 1) digits
            value := value // base
        }
        return digits
    }

    ; 日期加减: "today + 30 days" / "2026-12-25 - 2w" -> 那一天; "2026-12-25 - today" -> 相差几天
    static _Dates(text, today := "") {
        static dateRe := "(today|now|今天|\d{4}-\d{1,2}-\d{1,2})"
        text := Trim(StrLower(text)), today := (today != "") ? today : FormatTime(, "yyyyMMdd")
        if RegExMatch(text, "^" dateRe "\s*([+-])\s*(\d+)\s*(d|days?|天|w|weeks?|周|星期|m|months?|个月|月|y|years?|年)$", &m) {
            start := CalculatorProvider._ParseDate(m[1], today)
            if (start = "")
                return []
            amount := (m[2] = "-") ? -m[3] : m[3] + 0
            unit := m[4]
            if RegExMatch(unit, "^(d|days?|天)$")
                result := DateAdd(start, amount, "Days")
            else if RegExMatch(unit, "^(w|weeks?|周|星期)$")
                result := DateAdd(start, amount * 7, "Days")
            else
                result := CalculatorProvider.AddMonths(start, RegExMatch(unit, "^(y|years?|年)$") ? amount * 12 : amount)
            label := FormatTime(result, "yyyy-MM-dd") " (" FormatTime(result, "ddd") ")"
            return [CalculatorProvider._Item(label, I18n.T("Calc.Subtitle") " · " text)]
        }
        if RegExMatch(text, "^" dateRe "\s*-\s*" dateRe "$", &m) {
            a := CalculatorProvider._ParseDate(m[1], today), b := CalculatorProvider._ParseDate(m[2], today)
            if (a = "" || b = "")
                return []
            days := DateDiff(a, b, "Days")
            return [CalculatorProvider._Item(I18n.T("Calc.Days", days), I18n.T("Calc.Subtitle") " · " text)]
        }
        return []
    }

    static _ParseDate(text, today) {
        if RegExMatch(text, "^(today|now|今天)$")
            return today
        if !RegExMatch(text, "^(\d{4})-(\d{1,2})-(\d{1,2})$", &m)
            return ""
        stamp := Format("{:04}{:02}{:02}", m[1], m[2], m[3])
        try {
            if (FormatTime(stamp, "yyyyMMdd") = stamp)                      ; 2026-02-30 这样不存在的日期不算
                return stamp
        }
        return ""
    }

    ; 加几个月; 那个月没有这一天时用月底 (1 月 31 日 + 1 个月 = 2 月 28 / 29 日)
    static AddMonths(stamp, months) {
        year := SubStr(stamp, 1, 4) + 0, month := SubStr(stamp, 5, 2) + 0, day := SubStr(stamp, 7, 2) + 0
        total := year * 12 + (month - 1) + months
        year := total // 12, month := Mod(total, 12) + 1
        nextMonth := (month = 12) ? Format("{:04}0101", year + 1) : Format("{:04}{:02}01", year, month + 1)
        lastDay := SubStr(DateAdd(nextMonth, -1, "Days"), 7, 2) + 0
        return Format("{:04}{:02}{:02}", year, month, Min(day, lastDay))
    }

    static _AddStructural(results, value) {
        options := AppSettings.Feature("Calculator")
        number(key, fallback, minimum) {                                    ; 设置里的数字, 写错时用默认值
            raw := options.Has(key) ? options[key] : fallback
            return (IsNumber(raw) && raw >= minimum) ? raw + 0 : fallback
        }
        edge := number("BarEdge", 40, 0), maxSpacing := number("MaxBarSpacing", 300, 1)
        inner := value - 2 * edge                                           ; 两边主筋中心之间的距离
        if (inner <= 0)
            return
        barCount := Ceil(inner / maxSpacing + 1)
        spacing  := Max(Round(inner / (barCount - 0.999)), 0)
        beamText := I18n.T("Calc.BeamWidth", CalculatorProvider.Format(value), barCount, spacing)
        results.Push(ResultItem(beamText, "", {Kind: "text", Arg: beamText, Icon: "res:imageres.dll,-182", Score: 199}))

        prefix := options.Has("BarPrefix") ? Trim(options["BarPrefix"]) : "H"
        sizes := (options.Has("BarSizes") && options["BarSizes"] is Array) ? options["BarSizes"] : [13, 16, 20, 25, 32]
        bars := ""
        for size in sizes {
            if !(IsNumber(size) && size > 0)
                continue
            area := 3.141592653589793 * size * size / 4                     ; 一根的面积 (mm²)
            bars .= (bars = "" ? "" : "  ") Ceil(Round(value / area, 6)) prefix CalculatorProvider.Format(size + 0)
        }
        if (bars = "")
            return
        areaText := I18n.T("Calc.RebarArea", CalculatorProvider.Format(value), bars)
        results.Push(ResultItem(areaText, "", {Kind: "text", Arg: areaText, Icon: "res:imageres.dll,-182", Score: 198}))
    }
}
