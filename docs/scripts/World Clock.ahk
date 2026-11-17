; @altrun.title          World Clock
; @altrun.keyword        tz
; @altrun.mode           filter
; @altrun.filter         altrun
; @altrun.description    Live results: the time now in 60 cities, with the date and the difference from yours; Enter copies it
; @altrun.description_zh 实时结果: 60 个城市现在的时间、日期和与本地的时差; 回车复制
; @altrun.network        none
;
; Type "tz" and a space in ALTRun, then a city in English or Chinese ("tz london", "tz 东京", or pinyin initials "tz dj").
; The times come from Windows' own time zone data, summer time included; nothing is downloaded.
; Add a city: a line in Cities below with its English and Chinese name and its Windows time zone (the names "tzutil /l" lists).
#Requires AutoHotkey v2.0
#NoTrayIcon

Cities := [
    ["London", "伦敦", "GMT Standard Time"], ["Dublin", "都柏林", "GMT Standard Time"], ["Lisbon", "里斯本", "GMT Standard Time"],
    ["Paris", "巴黎", "Romance Standard Time"], ["Madrid", "马德里", "Romance Standard Time"], ["Brussels", "布鲁塞尔", "Romance Standard Time"],
    ["Berlin", "柏林", "W. Europe Standard Time"], ["Amsterdam", "阿姆斯特丹", "W. Europe Standard Time"], ["Rome", "罗马", "W. Europe Standard Time"],
    ["Zurich", "苏黎世", "W. Europe Standard Time"], ["Stockholm", "斯德哥尔摩", "W. Europe Standard Time"], ["Warsaw", "华沙", "Central European Standard Time"],
    ["Athens", "雅典", "GTB Standard Time"], ["Helsinki", "赫尔辛基", "FLE Standard Time"], ["Istanbul", "伊斯坦布尔", "Turkey Standard Time"],
    ["Moscow", "莫斯科", "Russian Standard Time"], ["Cairo", "开罗", "Egypt Standard Time"], ["Johannesburg", "约翰内斯堡", "South Africa Standard Time"],
    ["Lagos", "拉各斯", "W. Central Africa Standard Time"], ["Nairobi", "内罗毕", "E. Africa Standard Time"], ["Riyadh", "利雅得", "Arab Standard Time"],
    ["Doha", "多哈", "Arab Standard Time"], ["Dubai", "迪拜", "Arabian Standard Time"], ["Tehran", "德黑兰", "Iran Standard Time"],
    ["Karachi", "卡拉奇", "Pakistan Standard Time"], ["Mumbai", "孟买", "India Standard Time"], ["New Delhi", "新德里", "India Standard Time"],
    ["Kathmandu", "加德满都", "Nepal Standard Time"], ["Dhaka", "达卡", "Bangladesh Standard Time"], ["Yangon", "仰光", "Myanmar Standard Time"],
    ["Bangkok", "曼谷", "SE Asia Standard Time"], ["Jakarta", "雅加达", "SE Asia Standard Time"], ["Ho Chi Minh City", "胡志明市", "SE Asia Standard Time"],
    ["Singapore", "新加坡", "Singapore Standard Time"], ["Kuala Lumpur", "吉隆坡", "Singapore Standard Time"], ["Manila", "马尼拉", "Singapore Standard Time"],
    ["Beijing", "北京", "China Standard Time"], ["Shanghai", "上海", "China Standard Time"], ["Hong Kong", "香港", "China Standard Time"],
    ["Taipei", "台北", "Taipei Standard Time"], ["Seoul", "首尔", "Korea Standard Time"], ["Tokyo", "东京", "Tokyo Standard Time"],
    ["Perth", "珀斯", "W. Australia Standard Time"], ["Brisbane", "布里斯班", "E. Australia Standard Time"], ["Sydney", "悉尼", "AUS Eastern Standard Time"],
    ["Melbourne", "墨尔本", "AUS Eastern Standard Time"], ["Auckland", "奥克兰", "New Zealand Standard Time"], ["Honolulu", "檀香山", "Hawaiian Standard Time"],
    ["Anchorage", "安克雷奇", "Alaskan Standard Time"], ["Los Angeles", "洛杉矶", "Pacific Standard Time"], ["San Francisco", "旧金山", "Pacific Standard Time"],
    ["Vancouver", "温哥华", "Pacific Standard Time"], ["Denver", "丹佛", "Mountain Standard Time"], ["Chicago", "芝加哥", "Central Standard Time"],
    ["Mexico City", "墨西哥城", "Central Standard Time (Mexico)"], ["New York", "纽约", "Eastern Standard Time"], ["Toronto", "多伦多", "Eastern Standard Time"],
    ["São Paulo", "圣保罗", "E. South America Standard Time"], ["Buenos Aires", "布宜诺斯艾利斯", "Argentina Standard Time"], ["UTC", "协调世界时", "UTC"]
]
Chinese := RegExMatch(A_Language, "i)^(0804|0404|0C04|1004|1404)$") > 0
Days := Chinese ? ["周日", "周一", "周二", "周三", "周四", "周五", "周六"] : ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]

zones := TimeZones()
if !zones.Count {                                   ; Windows 7, or no time zone data
    FileAppend('{"items": [{"title": "' (Chinese ? "找不到 Windows 的时区数据" : "Windows' time zone data wasn't found") '", "valid": false}]}', "*", "UTF-8-RAW")
    ExitApp
}
utc := Buffer(16, 0), DllCall("GetSystemTime", "Ptr", utc)
here := SystemTimeStamp(LocalOf(utc, ""))
items := ""
for city in Cities {
    if !zones.Has(city[3])
        continue                                    ; a zone this Windows doesn't have
    there := LocalOf(utc, zones[city[3]]), stamp := SystemTimeStamp(there)
    hours := DateDiff(stamp, here, "Minutes") / 60
    difference := (hours = 0) ? (Chinese ? "同一时间" : "same time")
        : Chinese ? (hours > 0 ? "早 " : "晚 ") HourText(Abs(hours)) " 小时" : HourText(Abs(hours)) (Abs(hours) = 1 ? " hour " : " hours ") (hours > 0 ? "ahead" : "behind")
    name := Chinese ? city[2] : city[1], time := FormatTime(stamp, "HH:mm")
    date := Days[FormatTime(stamp, "WDay")] " " FormatTime(stamp, Chinese ? "M月d日" : "d MMM")
    items .= (items = "" ? "" : ",`n") '{"title": "' time "  " Json(name) '", "subtitle": "' Json(date " · " Offset(utc, there) " · " difference)
        . '", "arg": "' Json(name " " time " (" date ")") '", "type": "text", "match": "' Json(city[1] " " city[2] " " city[3])
        . '", "uid": "' Json(city[1]) '"}'
}
FileAppend('{"items": [' items ']}', "*", "UTF-8-RAW")
ExitApp

; Windows' time zones by key name ("Tokyo Standard Time") -> DYNAMIC_TIME_ZONE_INFORMATION (432 bytes; TimeZoneKeyName at 172)
TimeZones() {
    zones := Map()
    Loop {
        info := Buffer(432, 0)
        if (DllCall("advapi32\EnumDynamicTimeZoneInformation", "UInt", A_Index - 1, "Ptr", info) != 0)
            break
        zones[StrGet(info.Ptr + 172, 128, "UTF-16")] := info
    }
    return zones
}

; SYSTEMTIME (UTC) -> SYSTEMTIME in the zone (zone "": the PC's own), summer time by that date's rules
LocalOf(utc, zone) {
    result := Buffer(16, 0)
    if (zone = "")
        DllCall("SystemTimeToTzSpecificLocalTime", "Ptr", 0, "Ptr", utc, "Ptr", result)
    else
        DllCall("SystemTimeToTzSpecificLocalTimeEx", "Ptr", zone, "Ptr", utc, "Ptr", result)
    return result
}

SystemTimeStamp(st) {                               ; SYSTEMTIME -> YYYYMMDDHH24MISS
    return Format("{:04}{:02}{:02}{:02}{:02}{:02}", NumGet(st, 0, "UShort"), NumGet(st, 2, "UShort"), NumGet(st, 6, "UShort")
        , NumGet(st, 8, "UShort"), NumGet(st, 10, "UShort"), NumGet(st, 12, "UShort"))
}

Offset(utc, there) {                                ; "UTC+8", "UTC-3:30"
    minutes := DateDiff(SystemTimeStamp(there), SystemTimeStamp(utc), "Minutes")
    return "UTC" (minutes < 0 ? "-" : "+") Floor(Abs(minutes) / 60) (Mod(Abs(minutes), 60) ? Format(":{:02}", Mod(Abs(minutes), 60)) : "")
}

HourText(h) => (h = Floor(h)) ? Integer(h) : RTrim(Format("{:.2f}", h), "0")

Json(s) {                                           ; text -> inside a JSON string
    s := StrReplace(StrReplace(s, "\", "\\"), '"', '\"')
    return StrReplace(StrReplace(StrReplace(s, "`r", "\r"), "`n", "\n"), "`t", "\t")
}
