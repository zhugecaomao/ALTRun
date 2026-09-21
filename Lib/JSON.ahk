;===============================================================================
; JSON.ahk - 极简 JSON 读写器 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 对象 -> Map, 数组 -> Array, true/false -> 1/0, null -> ""
; Stringify 输出带缩进的可读格式, 方便你直接用记事本改配置。
;
; 用法:
;   data := JSON.Parse(FileRead("x.json", "UTF-8"))
;   text := JSON.Stringify(data)
;===============================================================================

class JSON {

    ;--- 解析入口 ----------------------------------------------------------------
    static Parse(text) {
        pos := 1
        value := JSON._Value(text, &pos)
        JSON._Space(text, &pos)
        return value
    }

    ;--- 序列化入口 --------------------------------------------------------------
    ; obj    : Map / Array / 字符串 / 数字
    ; indent : 当前缩进层级(递归用, 外部调用不用传)
    static Stringify(obj, indent := 0) {
        pad    := JSON._Pad(indent)
        padIn  := JSON._Pad(indent + 1)

        if (obj is Map) {
            if (obj.Count = 0)
                return "{}"
            parts := []
            for key, val in obj
                parts.Push(padIn '"' JSON.Escape(key) '": ' JSON.Stringify(val, indent + 1))
            return "{`r`n" JSON._Join(parts, ",`r`n") "`r`n" pad "}"
        }

        if (obj is Array) {
            if (obj.Length = 0)
                return "[]"
            parts := []
            for val in obj
                parts.Push(padIn JSON.Stringify(val, indent + 1))
            return "[`r`n" JSON._Join(parts, ",`r`n") "`r`n" pad "]"
        }

        ; 整数/浮点直接输出; 其余一律当字符串
        if (obj is Integer || obj is Float)
            return obj ""

        return '"' JSON.Escape(obj "") '"'
    }

    ;--- 字符串转义 --------------------------------------------------------------
    static Escape(str) {
        str := StrReplace(str, "\", "\\")
        str := StrReplace(str, '"', '\"')
        str := StrReplace(str, "`r", "\r")
        str := StrReplace(str, "`n", "\n")
        str := StrReplace(str, "`t", "\t")
        return str
    }

    ;=== 以下为内部实现 ==========================================================

    static _Pad(level) {
        pad := ""
        Loop level
            pad .= "  "
        return pad
    }

    static _Join(parts, sep) {
        out := ""
        for i, p in parts
            out .= (i = 1 ? "" : sep) p
        return out
    }

    ; 跳过空白。先做单字符快判, 绝大多数 token 前没有空白, 可省掉一次正则。
    static _Space(text, &pos) {
        ch := SubStr(text, pos, 1)
        if (ch != " " && ch != "`t" && ch != "`r" && ch != "`n")
            return
        if RegExMatch(text, "\s*", &m, pos)
            pos += m.Len
    }

    static _Value(text, &pos) {
        JSON._Space(text, &pos)
        ch := SubStr(text, pos, 1)

        switch ch, true {
            case "{" : return JSON._Object(text, &pos)
            case "[" : return JSON._Array(text, &pos)
            case '"' : return JSON._String(text, &pos)
        }
        if (SubStr(text, pos, 4) = "true") {
            pos += 4
            return 1
        }
        if (SubStr(text, pos, 5) = "false") {
            pos += 5
            return 0
        }
        if (SubStr(text, pos, 4) = "null") {
            pos += 4
            return ""
        }
        if (RegExMatch(text, "-?\d++(\.\d++)?([eE][-+]?\d++)?", &m, pos) && m.Pos = pos) {
            pos += m.Len
            return m[0] + 0
        }
        throw Error("JSON: 位置 " pos " 处有无法识别的字符")
    }

    static _Object(text, &pos) {
        obj := Map()
        pos++                                       ; 跳过 {
        JSON._Space(text, &pos)
        if (SubStr(text, pos, 1) = "}") {
            pos++
            return obj
        }
        loop {
            JSON._Space(text, &pos)
            key := JSON._String(text, &pos)
            JSON._Space(text, &pos)
            if (SubStr(text, pos, 1) != ":")
                throw Error("JSON: 位置 " pos " 处缺少 ':'")
            pos++
            obj[key] := JSON._Value(text, &pos)
            JSON._Space(text, &pos)
            ch := SubStr(text, pos, 1)
            pos++
            if (ch = ",")
                continue
            if (ch = "}")
                return obj
            throw Error("JSON: 位置 " (pos - 1) " 处缺少 ',' 或 '}'")
        }
    }

    static _Array(text, &pos) {
        arr := Array()
        pos++                                       ; 跳过 [
        JSON._Space(text, &pos)
        if (SubStr(text, pos, 1) = "]") {
            pos++
            return arr
        }
        loop {
            arr.Push(JSON._Value(text, &pos))
            JSON._Space(text, &pos)
            ch := SubStr(text, pos, 1)
            pos++
            if (ch = ",")
                continue
            if (ch = "]")
                return arr
            throw Error("JSON: 位置 " (pos - 1) " 处缺少 ',' 或 ']'")
        }
    }

    static _String(text, &pos) {
        ; 一次正则吃掉整个字符串, 比逐字符扫描快很多
        static rx := '"((?:[^"\\]|\\.)*+)"'
        if !(RegExMatch(text, rx, &m, pos) && m.Pos = pos)
            throw Error("JSON: 位置 " pos " 处字符串格式错误")
        pos += m.Len
        return JSON._Unescape(m[1])
    }

    static _Unescape(str) {
        if !InStr(str, "\")                         ; 没有转义符就直接返回
            return str
        out := "", i := 1, len := StrLen(str)
        while (i <= len) {
            ch := SubStr(str, i, 1)
            if (ch != "\") {
                out .= ch
                i++
                continue
            }
            esc := SubStr(str, i + 1, 1)
            switch esc, true {
                case '"': out .= '"'
                case "\": out .= "\"
                case "/": out .= "/"
                case "b": out .= Chr(8)
                case "f": out .= Chr(12)
                case "n": out .= "`n"
                case "r": out .= "`r"
                case "t": out .= "`t"
                case "u":
                    out .= Chr("0x" SubStr(str, i + 2, 4))
                    i += 4
                default : out .= esc
            }
            i += 2
        }
        return out
    }
}
