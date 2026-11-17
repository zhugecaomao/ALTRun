; @altrun.title          Translate (DeepL)
; @altrun.keyword        fy
; @altrun.mode           filter
; @altrun.description    Live translation with DeepL: Chinese to English, other languages to Chinese
; @altrun.description_zh 用 DeepL 边打字边翻译: 中文译成英文, 其他语言译成中文
; @altrun.requires       Your own DeepL API key (the free plan is enough) in the user environment variable DEEPL_AUTH_KEY
; @altrun.requires_zh    自己的 DeepL API key (免费版就够), 放在用户环境变量 DEEPL_AUTH_KEY
; @altrun.network        api-free.deepl.com, api.deepl.com
;
; Live translation with DeepL: "fy text" shows the translation (Chinese -> English, anything else -> Chinese).
; Enter copies it, Ctrl+Enter pastes it into the window you came from. The text typed after "fy" is sent to DeepL.
; Needs your own DeepL API key (the free plan is enough) in the user environment variable DEEPL_AUTH_KEY.
#Requires AutoHotkey v2.0
#NoTrayIcon

text := A_Args.Length ? Trim(A_Args[1]) : ""
key := EnvGet("DEEPL_AUTH_KEY")
if (key = "")                                       ; set after ALTRun started: read it where Windows keeps it
    try key := RegRead("HKCU\Environment", "DEEPL_AUTH_KEY")
if (key = "")
    Show("Set DEEPL_AUTH_KEY first", "Your DeepL API key, as a user environment variable", false)
if (text = "")
    Show("Type the text to translate", "Chinese -> English, other languages -> Chinese", false)

target := RegExMatch(text, "[\x{3400}-\x{9FFF}]") ? "EN-US" : "ZH-HANS"
host := (SubStr(key, -3) = ":fx") ? "api-free.deepl.com" : "api.deepl.com"   ; free plan keys end with :fx
try {
    request := ComObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts(2000, 2000, 2000, 3500)      ; together under ALTRun's 10 s, so a slow network shows the error below
    request.Open("POST", "https://" host "/v2/translate", false)
    request.SetRequestHeader("Authorization", "DeepL-Auth-Key " key)
    request.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    request.Send("text=" UrlEncode(text) "&target_lang=" target)
} catch as e                                        ; offline, proxy, time-out: an error row instead of an error dialog
    Fail("Can't reach DeepL: " Trim(e.Message, " `r`n"))
switch request.Status {
    case 200:
    case 403: Fail("DeepL didn't accept the key in DEEPL_AUTH_KEY")
    case 456: Fail("DeepL: this month's characters are used up")
    default:  Fail("DeepL answered " request.Status " " request.StatusText)
}
body := Utf8(request.ResponseBody)
if !RegExMatch(body, '"text"\s*:\s*"((?:[^"\\]|\\.)*)"', &m)
    Fail("Unexpected answer from DeepL")
from := RegExMatch(body, '"detected_source_language"\s*:\s*"(\w+)"', &source) ? source[1] " → " : ""
Show(Unescape(m[1]), "DeepL · " from target)

; One result (or a hint that can't be chosen, valid = false) as JSON, then end
Show(title, subtitle, valid := true) {
    FileAppend('{"items": [{"title": "' Escape(title) '", "subtitle": "' Escape(subtitle) '", "arg": "' Escape(title)
        . '", "type": "text", "valid": ' (valid ? "true" : "false") '}]}', "*", "UTF-8-RAW")
    ExitApp
}

Fail(message) {                                     ; ALTRun shows the first line of standard error as an error row
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

Utf8(safeArray) {                                   ; the answer's bytes as UTF-8 text (ResponseText guesses the code page)
    return StrGet(NumGet(ComObjValue(safeArray), 8 + A_PtrSize, "Ptr"), safeArray.MaxIndex() + 1, "UTF-8")
}

Escape(s) {                                         ; text -> inside a JSON string
    s := StrReplace(StrReplace(s, "\", "\\"), '"', '\"')
    s := StrReplace(StrReplace(StrReplace(s, "`r", "\r"), "`n", "\n"), "`t", "\t")
    while RegExMatch(s, "[\x01-\x1F]", &c)            ; other control characters, which JSON doesn't allow as they are
        s := StrReplace(s, c[0], Format("\u{:04X}", Ord(c[0])))
    return s
}

Unescape(s) {                                       ; inside a JSON string -> text
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
