;===============================================================================
; CurrencyRates.ahk - exchange rates for currency conversion (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Only used with Features.Calculator.Currency on (off by default): the only site besides GitHub that ALTRun ever visits.
; Rates come from Frankfurter (https://frankfurter.dev, reference rates published by the ECB and other central banks, free, no key needed),
; updated once per working day. The download is stored in Data\Currency.json; 1.5 minutes after start-up and every hour after,
; it downloads again when the last download is 12 hours old. Failures only go to the log, and the last rates stay in use.
;
; Usage:
;   CurrencyRates.Start()    at start-up (only called when the setting is on): reads the cache, downloads in the background if needed
;   CurrencyRates.Date       date of the rates (yyyy-MM-dd), shown in conversion results
;   Units.Rates              the rates conversions use (Lib\Units.ahk)
;===============================================================================

class CurrencyRates {
    static Url   := "https://api.frankfurter.dev/v1/latest?base=EUR"
    static File  := AppSettings.DataDir "\Currency.json"
    static Date  := "", Fetched := ""
    static StaleHours := 12
    static _timer := ""

    static Start() {
        CurrencyRates.Load()
        CurrencyRates._timer := () => CurrencyRates._Tick()
        ; WinTimer, not SetTimer: it stays set while ALTRun is idle, and a SetTimer timer would wake AutoHotkey every 10 ms (see Lib\WinTimer.ahk)
        WinTimer.Set(CurrencyRates._timer, -90000)
    }

    static _Tick() {
        if CurrencyRates.IsStale(CurrencyRates.Fetched, A_Now)
            CurrencyRates.Refresh()
        WinTimer.Set(CurrencyRates._timer, -3600000)
    }

    static IsStale(fetched, now) {
        if (fetched = "")
            return true
        try return DateDiff(now, fetched, "Hours") >= CurrencyRates.StaleHours || DateDiff(now, fetched, "Hours") < 0
        return true
    }

    static Load() {
        try {
            data := JSON.Parse(FileRead(CurrencyRates.File, "UTF-8"))
            CurrencyRates.Apply(data)
            CurrencyRates.Fetched := data.Has("Fetched") ? data["Fetched"] : ""
        }
    }

    ; Frankfurter's reply ({"base": "EUR", "date": "...", "rates": {...}}) or the cache ({"Date", "Rates"}) -> Units.Rates
    static Apply(data) {
        rates := data.Has("rates") ? data["rates"] : data.Has("Rates") ? data["Rates"] : ""
        if !(rates is Map) || !rates.Count
            throw Error("No exchange rates in the data.")
        table := Map("EUR", 1)
        for code, rate in rates
            if (IsNumber(rate) && rate > 0)
                table[StrUpper(code)] := rate + 0
        Units.Rates := table
        CurrencyRates.Date := data.Has("date") ? data["date"] : data.Has("Date") ? data["Date"] : ""
    }

    ; Downloads in the background (ALTRun doesn't hang) and swaps in the new rates when done
    static Refresh() {
        HttpGet.Start(CurrencyRates.Url, (text) => CurrencyRates._Downloaded(text), (message) => Logger.Error("CurrencyRates: " message))
    }

    static _Downloaded(text) {
        CurrencyRates.Apply(JSON.Parse(text))
        CurrencyRates.Fetched := A_Now
        rates := Map()
        for code, rate in Units.Rates
            rates[code] := rate
        JSON.WriteFile(CurrencyRates.File, Map("Date", CurrencyRates.Date, "Fetched", CurrencyRates.Fetched, "Rates", rates))
        Logger.Debug("CurrencyRates: " Units.Rates.Count " rates of " CurrencyRates.Date)
    }
}
