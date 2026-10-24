;===============================================================================
; JSON.ahk - a minimal JSON reader / writer (AutoHotkey v2)
;-------------------------------------------------------------------------------
; object -> Map, array -> Array, true / false -> 1 / 0, null -> ""
; Stringify writes an indented, readable form, so the settings can be edited in Notepad.
;
; Usage:
;   data := JSON.Parse(FileRead("x.json", "UTF-8"))
;   text := JSON.Stringify(data)
;   JSON.WriteFile("x.json", data)      writes a temporary file first and then replaces the file; UTF-8 without BOM
;===============================================================================

class JSON {

    ;--- Parsing -----------------------------------------------------------------
    static Parse(text) {
        pos := 1
        value := JSON._Value(&text, &pos)
        JSON._Space(&text, &pos)
        return value
    }

    ;--- Serializing -------------------------------------------------------------
    ; obj    : Map / Array / string / number
    ; indent : current indent level (for recursion; callers leave it out)
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

        ; integers / floats as they are; everything else as a string
        if (obj is Integer || obj is Float)
            return obj ""

        return '"' JSON.Escape(obj "") '"'
    }

    ;--- Writing files -----------------------------------------------------------
    ; Writes a temporary file in the same folder, then replaces the original: a crash or power cut halfway can't damage it.
    ; Creates the folder if needed; UTF-8 without BOM (JSON standard RFC 8259).
    ; Sync folders such as OneDrive and antivirus software hold a just-changed file for a moment: on failure retry every 100 ms, up to MoveTries times, then throw
    static MoveTries := 20

    static WriteFile(file, data, indent := 0) {
        SplitPath(file, , &dir)
        if (dir != "")
            DirCreate(dir)
        tmpFile := file ".tmp"
        try FileDelete(tmpFile)
        if FileExist(tmpFile)                                               ; a temporary file left over from last time can't be deleted (in use): use another name
            tmpFile := file "." A_TickCount ".tmp"
        FileAppend(JSON.Stringify(data, indent), tmpFile, "UTF-8-RAW")
        Loop JSON.MoveTries {
            try {
                FileMove(tmpFile, file, true)
                return
            } catch as e {
                if (A_Index = JSON.MoveTries) {
                    try FileDelete(tmpFile)
                    throw e
                }
                Sleep(100)
            }
        }
    }

    ;--- String escaping ---------------------------------------------------------
    static Escape(str) {
        str := StrReplace(str, "\", "\\")
        str := StrReplace(str, '"', '\"')
        str := StrReplace(str, "`r", "\r")
        str := StrReplace(str, "`n", "\n")
        str := StrReplace(str, "`t", "\t")
        if RegExMatch(str, "[\x00-\x1F]") {                                ; other control characters (rare, e.g. ESC copied from a terminal): \u00XX
            out := ""
            Loop Parse, str
                out .= (Ord(A_LoopField) < 32) ? Format("\u{:04X}", Ord(A_LoopField)) : A_LoopField
            str := out
        }
        return str
    }

    ;=== Internals ===============================================================

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

    ; The whole text is always passed by reference (&text) to the internal functions: AutoHotkey copies a string passed by value,
    ; so each value read would copy the whole file, and the total time grew with the square of its size (9 s for 6,000 paths, minutes for 30,000).
    ; Reading also moves along with InStr / SubStr only, never running a regex over the whole text.

    ; skip white space (space, Tab, CR, LF)
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
        stop := pos                                 ; numbers: take the run of number characters first, then check only that short piece with a regex
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
        pos++                                       ; skip {
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
        pos++                                       ; skip [
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

    ; pos is at the opening quote: InStr finds the next quote; one after an odd number of backslashes is escaped, so look further
    static _String(&text, &pos) {
        from := pos + 1
        loop {
            quote := InStr(text, '"', true, from)
            if !quote
                throw Error("JSON: 位置 " pos " 处的字符串没有结束")
            back := quote - 1
            while (SubStr(text, back, 1) = "\")
                back--
            if Mod(quote - 1 - back, 2) = 0         ; an even number of backslashes: this quote is the end
                break
            from := quote + 1
        }
        raw := SubStr(text, pos + 1, quote - pos - 1)
        pos := quote + 1
        return JSON._Unescape(raw)
    }

    ; joined piece by piece between backslashes (not character by character)
    static _Unescape(str) {
        if !InStr(str, "\")                        ; no escapes: return it as it is
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
