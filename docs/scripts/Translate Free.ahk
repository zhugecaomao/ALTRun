; @altrun.title          Translate (Free)
; @altrun.keyword        fy
; @altrun.mode           filter
; @altrun.author         sfcboy
; @altrun.description    Free live translation, no API key: Chinese <-> English and more. Enter copies, Ctrl+Enter pastes
; @altrun.description_zh 免费实时翻译, 无需 API key: 中英互译及更多语言。回车复制, Ctrl+Enter 粘贴到原窗口
; @altrun.network        api.mymemory.translated.net, fanyi.youdao.com
;
; HOW TO USE / 用法
; Type "fy text" in ALTRun: the translation appears while you type (Chinese -> English, other languages -> Chinese).
; Enter copies the result, Ctrl+Enter pastes it into the window you came from. No API key or signup needed.
; Primary engine: MyMemory (free, no key). Fallback: Youdao's public endpoint. To force a direction, start the text
; with "en>zh", "zh>en", "ja>zh" etc. ("fy en>zh hello" translates English to Chinese).
#Requires AutoHotkey v2.0
#NoTrayIcon

text := A_Args.Length ? Trim(A_Args[1]) : ""
if (text = "")
    Show("输入要翻译的文字", "中文 -> 英文, 其他语言 -> 中文; 也可写 en>zh hello", "输入 fy 后接文字", false)

; Optional direction prefix, e.g. "en>zh hello". Direction detection is automatic when absent.
from := "", to := ""
if RegExMatch(text, "^([a-z]{2,3})>([a-z]{2,3})\s+", &m) {
    from := m[1], to := m[2], text := Trim(SubStr(text, m[0].Length + 1))
    if (text = "")
        Show("方向已设置, 还缺文字", "例如: fy en>zh hello", "Direction set, add text", false)
}
if (from = "")
    from := "auto"
if (to = "")
    to := RegExMatch(text, "[\x{3400}-\x{9FFF}]") ? "en" : "zh-CN"

translated := Translate(text, from, to)
if (translated = "")
    Fail("翻译服务未返回结果 / No translation returned")
Show(translated, "免费翻译 · " from " → " to)

; ---- Translate: MyMemory first, Youdao as fallback. Returns "" on total failure. ----
Translate(text, from, to) {
    result := MyMemory(text, from, to)
    if (result != "")
        return result
    return Youdao(text, from)
}

MyMemory(text, from, to) {
    ; MyMemory expects langpair like "zh-CN|en"; "auto" is not supported, guess source here.
    src := (from = "auto") ? (RegExMatch(text, "[\x{3400}-\x{9FFF}]") ? "zh-CN" : "en") : Normalize(from)
    dst := Normalize(to)
    pair := src "|" dst
    url := "https://api.mymemory.translated.net/get?q=" UrlEncode(text) "&langpair=" UrlEncode(pair)
    try {
        req := ComObject("WinHttp.WinHttpRequest.5.1")
        req.SetTimeouts(1500, 1500, 1500, 3000)     ; stay under ALTRun's 10 s filter budget
        req.Open("GET", url, false)
        req.Send()
        if (req.Status != 200)
            return ""
        body := Utf8(req.ResponseBody)
        if !RegExMatch(body, '"translatedText"\s*:\s*"((?:[^"\\]|\\.)*)"', &m)
            return ""
        value := Unescape(m[1])
        ; MyMemory returns the source unchanged with an error-ish message when out of quota
        if (value = text)
            return ""
        return value
    } catch
        return ""
}

Youdao(text, from) {
    ; Public web endpoint; no key. Works for auto direction.
    src := (from = "auto") ? "AUTO" : StrUpper(Normalize(from))
    url := "https://fanyi.youdao.com/translate?doctype=json&from=" src "&to=AUTO&i=" UrlEncode(text)
    try {
        req := ComObject("WinHttp.WinHttpRequest.5.1")
        req.SetTimeouts(1500, 1500, 1500, 3000)
        req.Open("GET", url, false)
        req.Send()
        if (req.Status != 200)
            return ""
        body := Utf8(req.ResponseBody)
        ; Result shape: {"translateResult":[[{"tgt":"...","src":"..."}]]}
        if !RegExMatch(body, '"tgt"\s*:\s*"((?:[^"\\]|\\.)*)"', &m)
            return ""
        value := Unescape(m[1])
        if (value = text)
            return ""
        return value
    } catch
        return ""
}

Normalize(tag) {
    tag := StrLower(tag)
    if (tag = "zh" || tag = "cn" || tag = "zh-cn")
        return "zh-CN"
    if (tag = "en")
        return "en"
    return tag
}

; ---- Result / error rows as Alfred-style JSON, then exit (same as the official Translate.ahk). ----
Show(title, subtitle, arg := "", valid := true) {
    if (arg = "")
        arg := title
    FileAppend('{"items": [{"title": "' Escape(title) '", "subtitle": "' Escape(subtitle) '", "arg": "' Escape(arg)
        . '", "type": "text", "valid": ' (valid ? "true" : "false") '}]}', "*", "UTF-8-RAW")
    ExitApp
}

Fail(message) {
    FileAppend(message "`n", "**", "UTF-8-RAW")
    ExitApp(1)
}

UrlEncode(s) {
    size := StrPut(s, "UTF-8"), bytes := Buffer(size), StrPut(s, bytes, "UTF-8"), out := ""
    Loop size - 1 {
        b := NumGet(bytes, A_Index - 1, "UChar")
        out .= (b >= 0x30 && b <= 0x39 || b >= 0x41 && b <= 0x5A || b >= 0x61 && b <= 0x7A || b = 0x2D || b = 0x2E || b = 0x5F || b = 0x7E)
            ? Chr(b) : Format("%{:02X}", b)
    }
    return out
}

Utf8(safeArray) {
    return StrGet(NumGet(ComObjValue(safeArray), 8 + A_PtrSize, "Ptr"), safeArray.MaxIndex() + 1, "UTF-8")
}

Escape(s) {
    s := StrReplace(StrReplace(s, "\", "\\"), '"', '\"')
    s := StrReplace(StrReplace(StrReplace(s, "`r", "\r"), "`n", "\n"), "`t", "\t")
    while RegExMatch(s, "[\x01-\x1F]", &c)
        s := StrReplace(s, c[0], Format("\u{:04X}", Ord(c[0])))
    return s
}

Unescape(s) {
    out := "", i := 1
    while (i <= StrLen(s)) {
        c := SubStr(s, i, 1)
        if (c != "\") {
            out .= c, i += 1
            continue
        }
        n := SubStr(s, i + 1, 1)
        switch n {
            case "n": out .= "`n"
            case "r": out .= "`r"
            case "t": out .= "`t"
            case "b": out .= "`b"
            case "f": out .= "`f"
            case "u": out .= Chr(Integer("0x" SubStr(s, i + 2, 4))), i += 4
            default:  out .= n
        }
        i += 2
    }
    return out
}
