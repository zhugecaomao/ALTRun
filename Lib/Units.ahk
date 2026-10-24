;===============================================================================
; Units.ahk - unit and currency conversion (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Form: number + unit + (in / to / as / = / -> / 转 / 换成) + target unit, e.g.
;   10 km in mi     5ft to cm     100 f to c     2.5 t in kg     1 GB to MB
;   300 kn to kip   20 mpa in psi   3 亩 in m2     100 usd to sgd (needs exchange rates)
; Units are case-insensitive; ² ³ can be written as 2 3 (m2, cm3). Length, mass, area, volume, speed, time, data,
; pressure / stress, force, energy, power, temperature. GB / MB count in 1024s.
; Currencies: Units.Rates is a Map(code -> units per 1 euro) filled by the caller (CurrencyRates in ALTRun).
;
; Usage:
;   r := Units.Parse("10 km in mi")   -> {Value: 10, From: "km", To: "mi"} ("" when it isn't a conversion)
;   Units.Convert(10, "km", "mi")      -> 6.2137... ("" for unknown units or different kinds)
;   Units.Label("mi")                  -> "mi" (display name of the target unit)
;===============================================================================

class Units {
    static Rates := Map()                                                   ; currency code (upper case) -> units per 1 EUR
    static _unitMap := ""

    static Parse(text) {
        text := Trim(StrLower(text))
        if !RegExMatch(text, "^([-+]?\d[\d,]*(?:\.\d+)?)\s*([^\s\d][^\s]*?)\s*(?:\s(?:in|to|as)\s|=|->|>|\s转|\s换成|转|换成|\s)\s*([^\s\d][^\s]*)$", &m)
            return ""
        value := StrReplace(m[1], ",")
        if !IsNumber(value)
            return ""
        return {Value: value + 0, From: m[2], To: m[3]}
    }

    ; Returns the converted value; "" for unknown units, different kinds or a missing exchange rate
    static Convert(value, from, to) {
        a := Units._Lookup(from), b := Units._Lookup(to)
        if !IsObject(a) || !IsObject(b) || a.Kind != b.Kind
            return ""
        if (a.Kind = "temperature")
            return Units._FromKelvin(Units._ToKelvin(value, a.Name), b.Name)
        return value * a.Factor / b.Factor
    }

    static Label(unit) {
        info := Units._Lookup(unit)
        return IsObject(info) ? info.Label : unit
    }

    ; unit -> {Kind, Factor (to the base unit), Name, Label}
    static _Lookup(unit) {
        unit := StrReplace(StrReplace(StrLower(Trim(unit)), "²", "2"), "³", "3")
        unit := RegExReplace(unit, "^°")                                   ; °c / °f
        table := Units._Table()
        if table.Has(unit)
            return table[unit]
        code := StrUpper(unit)                                              ; currency code
        static names := Map("rmb", "CNY", "人民币", "CNY", "元", "CNY", "美元", "USD", "欧元", "EUR", "日元", "JPY", "英镑", "GBP"
            , "港币", "HKD", "港元", "HKD", "新币", "SGD", "新加坡元", "SGD", "马币", "MYR", "令吉", "MYR", "澳元", "AUD", "加元", "CAD"
            , "台币", "TWD", "韩元", "KRW", "泰铢", "THB", "瑞士法郎", "CHF")
        if names.Has(unit)
            code := names[unit]
        if (StrLen(code) = 3 && Units.Rates.Has(code))
            return {Kind: "currency", Factor: 1 / Units.Rates[code], Name: code, Label: code}
        return ""
    }

    static _Table() {
        if IsObject(Units._unitMap)
            return Units._unitMap
        table := Map()
        add(kind, factor, label, names*) {
            for name in names
                table[name] := {Kind: kind, Factor: factor, Name: names[1], Label: label}
        }
        ; length (m)
        add("length", 0.001, "mm", "mm", "millimeter", "millimeters", "毫米")
        add("length", 0.01, "cm", "cm", "centimeter", "centimeters", "厘米")
        add("length", 1, "m", "m", "meter", "meters", "metre", "metres", "米")
        add("length", 1000, "km", "km", "kilometer", "kilometers", "公里", "千米")
        add("length", 0.0254, "in", "in", "inch", "inches", "英寸")
        add("length", 0.3048, "ft", "ft", "foot", "feet", "英尺")
        add("length", 0.9144, "yd", "yd", "yard", "yards", "码")
        add("length", 1609.344, "mi", "mi", "mile", "miles", "英里")
        add("length", 1852, "nmi", "nmi", "海里")
        ; mass (kg)
        add("mass", 0.000001, "mg", "mg", "毫克")
        add("mass", 0.001, "g", "g", "gram", "grams", "克")
        add("mass", 1, "kg", "kg", "kilogram", "kilograms", "kilo", "公斤", "千克")
        add("mass", 1000, "t", "t", "tonne", "tonnes", "ton", "tons", "吨")
        add("mass", 0.45359237, "lb", "lb", "lbs", "pound", "pounds", "磅")
        add("mass", 0.028349523125, "oz", "oz", "ounce", "ounces", "盎司")
        add("mass", 0.5, "斤", "斤", "jin")
        add("mass", 0.05, "两", "两", "liang")
        ; area (m²)
        add("area", 0.000001, "mm²", "mm2", "sqmm")
        add("area", 0.0001, "cm²", "cm2", "sqcm")
        add("area", 1, "m²", "m2", "sqm", "平方米", "平米")
        add("area", 1000000, "km²", "km2", "sqkm", "平方公里")
        add("area", 10000, "ha", "ha", "hectare", "hectares", "公顷")
        add("area", 4046.8564224, "acre", "acre", "acres", "英亩")
        add("area", 0.09290304, "ft²", "ft2", "sqft")
        add("area", 0.00064516, "in²", "in2", "sqin")
        add("area", 10000 / 15, "亩", "亩", "mu")
        ; volume (L)
        add("volume", 0.001, "mL", "ml", "milliliter", "毫升")
        add("volume", 0.01, "cL", "cl")
        add("volume", 1, "L", "l", "liter", "liters", "litre", "litres", "升")
        add("volume", 1000, "m³", "m3", "立方米", "方")
        add("volume", 0.001, "cm³", "cm3", "cc")
        add("volume", 28.316846592, "ft³", "ft3")
        add("volume", 3.785411784, "gal", "gal", "gallon", "gallons", "加仑")
        add("volume", 0.946352946, "qt", "qt", "quart", "quarts")
        add("volume", 0.473176473, "pt", "pt", "pint", "pints")
        add("volume", 0.2365882365, "cup", "cup", "cups")
        add("volume", 0.0295735295625, "fl oz", "floz")
        ; speed (m/s)
        add("speed", 1, "m/s", "m/s", "mps")
        add("speed", 1 / 3.6, "km/h", "km/h", "kmh", "kph")
        add("speed", 0.44704, "mph", "mph")
        add("speed", 1852 / 3600, "kn", "knot", "knots", "节")
        add("speed", 0.3048, "ft/s", "ft/s", "fps")
        ; time (s)
        add("time", 0.001, "ms", "ms", "millisecond", "milliseconds", "毫秒")
        add("time", 1, "s", "s", "sec", "second", "seconds", "秒")
        add("time", 60, "min", "min", "mins", "minute", "minutes", "分钟")
        add("time", 3600, "h", "h", "hr", "hrs", "hour", "hours", "小时")
        add("time", 86400, "day", "day", "days", "d", "天")
        add("time", 604800, "week", "week", "weeks", "wk", "周", "星期")
        add("time", 31557600, "year", "year", "years", "yr", "年")
        ; data (bytes, in 1024s)
        add("data", 0.125, "bit", "bit", "bits")
        add("data", 1, "B", "b", "byte", "bytes")
        add("data", 1024, "KB", "kb", "kib")
        add("data", 1024 ** 2, "MB", "mb", "mib")
        add("data", 1024 ** 3, "GB", "gb", "gib")
        add("data", 1024 ** 4, "TB", "tb", "tib")
        ; pressure / stress (Pa)
        add("pressure", 1, "Pa", "pa", "n/m2")
        add("pressure", 1000, "kPa", "kpa", "kn/m2")
        add("pressure", 1000000, "MPa", "mpa", "n/mm2")
        add("pressure", 100000, "bar", "bar")
        add("pressure", 6894.757293168, "psi", "psi")
        add("pressure", 6894757.293168, "ksi", "ksi")
        add("pressure", 101325, "atm", "atm")
        add("pressure", 47.880259, "psf", "psf")
        ; force (N)
        add("force", 1, "N", "n", "newton", "newtons", "牛")
        add("force", 1000, "kN", "kn", "kilonewton", "千牛")
        add("force", 4.4482216152605, "lbf", "lbf")
        add("force", 4448.2216152605, "kip", "kip", "kips")
        add("force", 9.80665, "kgf", "kgf")
        add("force", 9806.65, "tf", "tf", "tonf")
        ; energy (J)
        add("energy", 1, "J", "j", "joule", "joules")
        add("energy", 1000, "kJ", "kj")
        add("energy", 4.184, "cal", "cal")
        add("energy", 4184, "kcal", "kcal", "千卡", "大卡")
        add("energy", 3600, "Wh", "wh")
        add("energy", 3600000, "kWh", "kwh", "度")
        ; power (W)
        add("power", 1, "W", "w", "watt", "watts", "瓦")
        add("power", 1000, "kW", "kw", "千瓦")
        add("power", 745.69987158227, "hp", "hp", "马力")
        ; temperature (calculated separately)
        add("temperature", 1, "°C", "c", "celsius", "摄氏度")
        add("temperature", 1, "°F", "f", "fahrenheit", "华氏度")
        add("temperature", 1, "K", "k", "kelvin")
        Units._unitMap := table
        return table
    }

    static _ToKelvin(value, name) {
        switch name {
            case "c": return value + 273.15
            case "f": return (value - 32) * 5 / 9 + 273.15
        }
        return value
    }

    static _FromKelvin(value, name) {
        switch name {
            case "c": return value - 273.15
            case "f": return (value - 273.15) * 9 / 5 + 32
        }
        return value
    }
}
