; @altrun.title          Unix Timestamp
; @altrun.keyword        ts
; @altrun.mode           filter
; @altrun.description    Live results: Unix time now, a timestamp as a date (seconds or milliseconds), or a date as a timestamp; Enter copies
; @altrun.description_zh 实时结果: 现在的 Unix 时间戳, 时间戳转成日期 (秒或毫秒), 或日期转成时间戳; 回车复制
; @altrun.network        none
;
; Type in ALTRun:
;   ts                      now, in seconds and milliseconds
;   ts 1794933000           a timestamp (10 digits: seconds, 13: milliseconds) as local time, UTC and ISO 8601
;   ts 2026-11-17 16:30     a date and time in your time zone (also 2026/11/17, seconds optional) as a timestamp
#Requires AutoHotkey v2.0
#NoTrayIcon

Chinese := RegExMatch(A_Language, "i)^(0804|0404|0C04|1004|1404)$") > 0
text := Trim(A_Args.Length ? A_Args[1] : "")
rows := []
if (text = "" || text = "now") {
    utc := A_NowUTC, ms := Format("{:03}", A_MSec)
    seconds := DateDiff(utc, "19700101000000", "Seconds")
    Row(seconds, Chinese ? "现在的 Unix 时间戳 (秒)" : "Unix time now (seconds)")
    Row(seconds ms, Chinese ? "毫秒" : "Milliseconds")
    Row(FormatTime(utc, "yyyy-MM-dd'T'HH:mm:ss") "Z", "ISO 8601 (UTC)")
} else if RegExMatch(text, "^-?\d{1,14}$") {
    milliseconds := StrLen(LTrim(text, "-")) >= 12
    seconds := milliseconds ? Floor(text / 1000) : Integer(text)
    utc := DateAdd("19700101000000", seconds, "Seconds"), here := UtcToLocal(utc)
    unit := milliseconds ? (Chinese ? "毫秒" : "milliseconds") : (Chinese ? "秒" : "seconds")
    Row(FormatTime(here, "yyyy-MM-dd HH:mm:ss"), (Chinese ? "本地时间 · 按" : "Local time · read as ") unit)
    Row(FormatTime(utc, "yyyy-MM-dd HH:mm:ss") " UTC", "UTC")
    Row(FormatTime(utc, "yyyy-MM-dd'T'HH:mm:ss") "Z", "ISO 8601 (UTC)")
} else if RegExMatch(text, "^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$", &m) {
    stamp := Format("{:04}{:02}{:02}{:02}{:02}{:02}", m[1], m[2], m[3], m[4] = "" ? 0 : m[4], m[5] = "" ? 0 : m[5], m[6] = "" ? 0 : m[6])
    if !IsTime(stamp)
        Hint(Chinese ? "不是有效的日期: " text : "Not a valid date: " text)
    seconds := DateDiff(LocalToUtc(stamp), "19700101000000", "Seconds")
    Row(seconds, Chinese ? "Unix 时间戳 (秒), 按本地时间" : "Unix time (seconds), read as local time")
    Row(seconds "000", Chinese ? "毫秒" : "Milliseconds")
} else
    Hint(Chinese ? "输入时间戳 (10 位或 13 位) 或日期, 例如 2026-11-17 16:30" : "Type a timestamp (10 or 13 digits) or a date such as 2026-11-17 16:30")
Print()
ExitApp

Row(value, subtitle) => rows.Push('{"title": "' Json(value) '", "subtitle": "' Json(subtitle) '", "arg": "' Json(value) '", "type": "text"}')

Hint(message) {                                     ; one row that can't be chosen, then end
    rows.Push('{"title": "' Json(message) '", "valid": false}')
    Print()
    ExitApp
}

Print() {
    out := ""
    for row in rows
        out .= (out = "" ? "" : ",`n") row
    FileAppend('{"items": [' out ']}', "*", "UTF-8-RAW")
}

; YYYYMMDDHH24MISS between UTC and this PC's time zone, with the summer time rules of that date
UtcToLocal(stamp) => FromSystemTime(Convert("SystemTimeToTzSpecificLocalTime", ToSystemTime(stamp)))
LocalToUtc(stamp) => FromSystemTime(Convert("TzSpecificLocalTimeToSystemTime", ToSystemTime(stamp)))

Convert(function, st) {
    result := Buffer(16, 0)
    DllCall(function, "Ptr", 0, "Ptr", st, "Ptr", result)
    return result
}

ToSystemTime(stamp) {
    st := Buffer(16, 0)
    for index, part in [[0, 1, 4], [2, 5, 2], [6, 7, 2], [8, 9, 2], [10, 11, 2], [12, 13, 2]]   ; offset, position, length
        NumPut("UShort", Integer(SubStr(stamp, part[2], part[3])), st, part[1])
    return st
}

FromSystemTime(st) => Format("{:04}{:02}{:02}{:02}{:02}{:02}", NumGet(st, 0, "UShort"), NumGet(st, 2, "UShort"), NumGet(st, 6, "UShort")
    , NumGet(st, 8, "UShort"), NumGet(st, 10, "UShort"), NumGet(st, 12, "UShort"))

Json(s) {                                           ; text -> inside a JSON string
    s := StrReplace(StrReplace(s, "\", "\\"), '"', '\"')
    return StrReplace(StrReplace(StrReplace(s, "`r", "\r"), "`n", "\n"), "`t", "\t")
}
