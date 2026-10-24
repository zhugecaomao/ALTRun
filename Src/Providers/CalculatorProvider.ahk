;===============================================================================
; CalculatorProvider.ahk - calculator (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Typing an expression shows the result directly ("12*(3+4)" / "=2^10"), at most two decimals; Enter copies the result.
; With Features.Calculator.StructuralCalc on, two rows of structural calculation follow the result:
;   - the result as a beam width (mm): number and spacing of main bars (cover RebarCover, default 40 mm; maximum spacing MaxBarSpacing, default 300 mm)
;   - the result as a steel area As (mm²): the number of bars needed for each diameter in BarSizes (default H13 / H16 / H20 / H25 / H32, prefix BarPrefix)
; Unit conversion (Lib\Units.ahk): "10 km in mi", "100 f to c", "20 mpa in psi"...; currency conversion "100 usd to sgd"
; needs Features.Calculator.Currency (off by default; exchange rates downloaded daily from frankfurter.dev, see CurrencyRates).
; Number bases: "0xFF" / "0b1010" / "0o17" show decimal, hex and binary; "255 in hex" / "0xff to dec" / "10 in bin" / "8 in oct".
; Dates: "today + 30 days" / "2026-12-25 - 2 weeks" / "today + 3 months" (units d w m y, or 天 周 月 年);
;       "2026-12-25 - today" gives the days between two dates. today can also be now / 今天.
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
            Kind: "text", Arg: text, Icon: IconCache.Own("Calculator", "res:imageres.dll,-182"), Score: 200, LargeText: text
        }))
        if AppSettings.Feature("Calculator")["StructuralCalc"]
            CalculatorProvider._AddStructural(results, value)
        return results
    }

    ; At most two decimals, trailing 0s dropped, thousands separators in the integer part: 1234567.504 -> 1,234,567.5, 10/3 -> 3.33
    ; (turning a float into text directly brings tails like 774.39999999999998).
    ; Numbers below 0.005 would become 0 with two decimals, so they keep one digit after the first significant one: 0.004, 0.000031
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

    ; "10 km in mi" -> one result "6.21 mi" (Enter copies the number); a hint when there are no exchange rates yet for currencies
    static _Convert(conversion) {
        icon := IconCache.Own("Calculator", "res:imageres.dll,-182")
        value := Units.Convert(conversion.Value, conversion.From, conversion.To)
        if IsNumber(value) {
            text := CalculatorProvider.Format(value)
            label := Units.Label(conversion.To)
            subtitle := CalculatorProvider.Format(conversion.Value) " " Units.Label(conversion.From) " = " text " " label
            if (Units._Lookup(conversion.To).Kind = "currency")
                subtitle .= " · " I18n.T("Calc.RatesOf", CurrencyRates.Date)
            return [ResultItem(text " " label, subtitle, {Kind: "text", Arg: text, Icon: icon, Score: 200, LargeText: text " " label})]
        }
        ; both sides look like currency codes (3 letters) but aren't units: hint to turn on currency conversion / rates not downloaded yet
        if (RegExMatch(conversion.From, "^[a-z]{3}$") && RegExMatch(conversion.To, "^[a-z]{3}$")
                && !IsObject(Units._Lookup(conversion.From)) && !IsObject(Units._Lookup(conversion.To))) {
            if AppSettings.Feature("Calculator")["Currency"]
                return [ResultItem(I18n.T("Calc.Currency"), I18n.T("Calc.RatesNotYet"), {Icon: icon, Score: 150, Valid: false})]
            return [ResultItem(I18n.T("Calc.Currency"), I18n.T("Calc.CurrencyOff"), {Icon: icon, Score: 150   ; Enter: opens the Calculator settings page
                , OnRun: (*) => App.OpenPreferences("Prefs.Page.Calculator")})]
        }
        return []
    }

    static _Item(text, subtitle, score := 200) {
        return ResultItem(text, subtitle, {Kind: "text", Arg: text, Icon: IconCache.Own("Calculator", "res:imageres.dll,-182"), Score: score, LargeText: text})
    }

    ; Number bases: a prefixed number (0x / 0b / 0o) typed alone lists the three forms; "n in hex|bin|oct|dec" converts to that base
    static _Bases(text) {
        static names := Map("hex", 16, "bin", 2, "oct", 8, "dec", 10)
        if !RegExMatch(Trim(text), "i)^(0x[0-9a-f]+|0b[01]+|0o[0-7]+|\d+)(?:\s+(?:in|to)\s+(hex|bin|oct|dec))?$", &m)
            return []
        if (m[2] = "" && !RegExMatch(m[1], "i)^0[xbo]"))                   ; plain numbers aren't base conversions
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

    ; "0xff" / "0b101" / "0o17" / "255" -> an integer ("" when too big)
    static ParseInteger(text) {
        base := 10, digits := text
        if RegExMatch(text, "i)^0([xbo])(.+)$", &m)
            base := (StrLower(m[1]) = "x") ? 16 : (StrLower(m[1]) = "b") ? 2 : 8, digits := m[2]
        if (StrLen(digits) > ((base = 2) ? 62 : (base = 8) ? 20 : 15))     ; at most a 64-bit integer
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

    ; Date arithmetic: "today + 30 days" / "2026-12-25 - 2w" -> that day; "2026-12-25 - today" -> days between
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
            if (FormatTime(stamp, "yyyyMMdd") = stamp)                      ; dates that don't exist, like 2026-02-30, don't count
                return stamp
        }
        return ""
    }

    ; Adds months; when that month lacks the day, its last day is used (31 Jan + 1 month = 28 / 29 Feb)
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
        number(key, fallback, minimum) {                                    ; a number from the settings; the default if it is invalid
            raw := options.Has(key) ? options[key] : fallback
            return (IsNumber(raw) && raw >= minimum) ? raw + 0 : fallback
        }
        cover := number("RebarCover", 40, 0), maxSpacing := number("MaxBarSpacing", 300, 1)
        inner := value - 2 * cover                                          ; minus the cover on both sides
        if (inner <= 0)
            return
        barCount := Ceil(inner / maxSpacing + 1)
        spacing  := Max(Round(inner / (barCount - 0.999)), 0)
        beamText := I18n.T("Calc.BeamWidth", CalculatorProvider.Format(value), barCount, spacing)
        results.Push(ResultItem(beamText, "", {Kind: "text", Arg: beamText, Icon: IconCache.Own("Calculator", "res:imageres.dll,-182"), Score: 199}))

        prefix := options.Has("BarPrefix") ? Trim(options["BarPrefix"]) : "H"
        sizes := (options.Has("BarSizes") && options["BarSizes"] is Array) ? options["BarSizes"] : [13, 16, 20, 25, 32]
        bars := ""
        for size in sizes {
            if !(IsNumber(size) && size > 0)
                continue
            area := 3.141592653589793 * size * size / 4                     ; area of one bar (mm²)
            bars .= (bars = "" ? "" : "  ") Ceil(Round(value / area, 6)) prefix CalculatorProvider.Format(size + 0)
        }
        if (bars = "")
            return
        areaText := I18n.T("Calc.RebarArea", CalculatorProvider.Format(value), bars)
        results.Push(ResultItem(areaText, "", {Kind: "text", Arg: areaText, Icon: IconCache.Own("Calculator", "res:imageres.dll,-182"), Score: 198}))
    }
}
