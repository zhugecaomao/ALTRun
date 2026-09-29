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
        value := JSON._Value(&text, &pos)
        JSON._Space(&text, &pos)
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

    ; 整段文本一律按引用 (&text) 传给内部函数: AutoHotkey 按值传字符串时会复制一份, 每读一个值
    ; 就复制一次整个文件, 总耗时随文件大小平方增长 (6000 条路径要 9 秒, 3 万条要几分钟)。
    ; 读取时也只用 InStr / SubStr 按位置前进, 不对整段文本做正则。

    ; 跳过空白 (空格、Tab、回车、换行)
    static _Space(&text, &pos) {
        while ((ch := Ord(SubStr(text, pos, 1))) = 32 || ch = 9 || ch = 13 || ch = 10)
            pos++
    }

    static _Value(&text, &pos) {
        JSON._Space(&text, &pos)
        ch := SubStr(text, pos, 1)

        switch ch, true {
            case "{" : return JSON._Object(&text, &pos)
            case "[" : return JSON._Array(&text, &pos)
            case '"' : return JSON._String(&text, &pos)
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
        stop := pos                                 ; 数字: 先取出连续的数字字符, 再只对这一小段用正则检查
        while ((c := SubStr(text, stop, 1)) != "" && InStr("0123456789+-.eE", c, true))
            stop++
        numText := SubStr(text, pos, stop - pos)
        if (numText != "" && RegExMatch(numText, "^-?\d+(\.\d+)?([eE][-+]?\d+)?$")) {
            pos := stop
            return numText + 0
        }
        throw Error("JSON: 位置 " pos " 处有无法识别的字符")
    }

    static _Object(&text, &pos) {
        obj := Map()
        pos++                                       ; 跳过 {
        JSON._Space(&text, &pos)
        if (SubStr(text, pos, 1) = "}") {
            pos++
            return obj
        }
        loop {
            JSON._Space(&text, &pos)
            if (SubStr(text, pos, 1) != '"')
                throw Error("JSON: 位置 " pos " 处应该是键名")
            key := JSON._String(&text, &pos)
            JSON._Space(&text, &pos)
            if (SubStr(text, pos, 1) != ":")
                throw Error("JSON: 位置 " pos " 处缺少 ':'")
            pos++
            obj[key] := JSON._Value(&text, &pos)
            JSON._Space(&text, &pos)
            ch := SubStr(text, pos, 1)
            pos++
            if (ch = ",")
                continue
            if (ch = "}")
                return obj
            throw Error("JSON: 位置 " (pos - 1) " 处缺少 ',' 或 '}'")
        }
    }

    static _Array(&text, &pos) {
        items := Array()
        pos++                                       ; 跳过 [
        JSON._Space(&text, &pos)
        if (SubStr(text, pos, 1) = "]") {
            pos++
            return items
        }
        loop {
            items.Push(JSON._Value(&text, &pos))
            JSON._Space(&text, &pos)
            ch := SubStr(text, pos, 1)
            pos++
            if (ch = ",")
                continue
            if (ch = "]")
                return items
            throw Error("JSON: 位置 " (pos - 1) " 处缺少 ',' 或 ']'")
        }
    }

    ; pos 指向开头的引号: 用 InStr 找下一个引号, 前面有奇数个反斜杠的是转义的引号, 继续往后找
    static _String(&text, &pos) {
        from := pos + 1
        loop {
            quote := InStr(text, '"', true, from)
            if !quote
                throw Error("JSON: 位置 " pos " 处的字符串没有结束")
            back := quote - 1
            while (SubStr(text, back, 1) = "\")
                back--
            if Mod(quote - 1 - back, 2) = 0         ; 偶数个反斜杠: 这个引号就是结尾
                break
            from := quote + 1
        }
        raw := SubStr(text, pos + 1, quote - pos - 1)
        pos := quote + 1
        return JSON._Unescape(raw)
    }

    ; 按反斜杠分段拼接 (不逐个字符处理)
    static _Unescape(str) {
        if !InStr(str, "\")                        ; 没有转义符就直接返回
            return str
        out := "", i := 1
        while (j := InStr(str, "\", true, i)) {
            out .= SubStr(str, i, j - i)
            esc := SubStr(str, j + 1, 1)
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
                    out .= Chr("0x" SubStr(str, j + 2, 4))
                    j += 4
                default : out .= esc
            }
            i := j + 2
        }
        return out SubStr(str, i)
    }
}
