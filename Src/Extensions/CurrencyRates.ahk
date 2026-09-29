;===============================================================================
; CurrencyRates.ahk - 货币换算用的汇率 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 只在打开 Features.Calculator.Currency 时使用 (默认关闭): 这是 ALTRun 除 GitHub 之外唯一会访问的网站。
; 汇率来自 Frankfurter (https://frankfurter.dev, 欧洲央行等央行公布的参考汇率, 免费, 不需要 key),
; 工作日每天更新一次。下载的结果存在 Data\Currency.json; 启动 1.5 分钟后、之后每小时看一下,
; 距离上次下载满 12 小时再下载。下载失败只写日志, 继续用上次的汇率。
;
; 用法:
;   CurrencyRates.Start()    启动时 (设置打开了才调用): 读缓存, 需要时在后台下载
;   CurrencyRates.Date       汇率的日期 (yyyy-MM-dd), 显示在换算结果里
;   Units.Rates              换算用的汇率 (Lib\Units.ahk)
;===============================================================================

class CurrencyRates {
    static Url   := "https://api.frankfurter.dev/v1/latest?base=EUR"
    static File  := A_ScriptDir "\Data\Currency.json"
    static Date  := "", Fetched := ""
    static StaleHours := 12
    static _timer := ""

    static Start() {
        CurrencyRates.Load()
        CurrencyRates._timer := () => CurrencyRates._Tick()
        SetTimer(CurrencyRates._timer, -90000)
    }

    static _Tick() {
        if CurrencyRates.IsStale(CurrencyRates.Fetched, A_Now)
            CurrencyRates.Refresh()
        SetTimer(CurrencyRates._timer, -3600000)
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

    ; Frankfurter 的回复 ({"base": "EUR", "date": "...", "rates": {...}}) 或缓存 ({"Date", "Rates"}) -> Units.Rates
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

    static Refresh() {
        tmpFile := A_Temp "\ALTRun_rates.json"
        try {
            Download(CurrencyRates.Url, tmpFile)
            data := JSON.Parse(FileRead(tmpFile, "UTF-8"))
            try FileDelete(tmpFile)
            CurrencyRates.Apply(data)
            CurrencyRates.Fetched := A_Now
            rates := Map()
            for code, rate in Units.Rates
                rates[code] := rate
            DirCreate(AppSettings.DataDir)
            try FileDelete(CurrencyRates.File)
            FileAppend(JSON.Stringify(Map("Date", CurrencyRates.Date, "Fetched", CurrencyRates.Fetched, "Rates", rates)), CurrencyRates.File, "UTF-8")
            Logger.Debug("CurrencyRates: " Units.Rates.Count " rates of " CurrencyRates.Date)
        } catch as e {
            Logger.Error("CurrencyRates: " e.Message)
        }
    }
}
