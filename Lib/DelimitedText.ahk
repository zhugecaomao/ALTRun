;===============================================================================
; DelimitedText.ahk - table text (CSV / TSV / copied from Excel) <-> two-dimensional array (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Pure functions. The separator is detected (Tab if the first line has the most Tabs, then ; and ,); separators and line breaks inside double quotes are kept,
; "" stands for one ", as Excel copies and saves it.
;
; Usage:
;   rows := DelimitedText.Parse(text)           [["Name", "Keyword"], ["Email", "em"]...], blank lines skipped
;   csv  := DelimitedText.ToCsv(rows)           comma-separated, quoted where needed, `r`n between rows
;   text := DelimitedText.ReadFile(path)        UTF-8 (with or without BOM); otherwise the system code page (e.g. GBK)
;===============================================================================

class DelimitedText {
    static Parse(text) {
        text := StrReplace(StrReplace(text, "`r`n", "`n"), "`r", "`n")
        delimiter := DelimitedText.Delimiter(text)
        rows := [], row := [], field := "", quoted := false, position := 1, length := StrLen(text)
        while (position <= length) {
            char := SubStr(text, position, 1)
            if quoted {
                if (char = '"') {
                    if (SubStr(text, position + 1, 1) = '"')
                        field .= '"', position++
                    else
                        quoted := false
                } else {
                    field .= char
                }
            } else if (char = '"' && field = "") {                          ; only a quote at the start of a field counts
                quoted := true
            } else if (char = delimiter) {
                row.Push(field), field := ""
            } else if (char = "`n") {
                row.Push(field), field := ""
                DelimitedText._AddRow(rows, row), row := []
            } else {
                field .= char
            }
            position++
        }
        row.Push(field)
        DelimitedText._AddRow(rows, row)
        return rows
    }

    ; In the first non-empty line (outside quotes): Tab if it has the most Tabs, otherwise ; if there are more ; than , (European Excel), , by default
    static Delimiter(text) {
        counts := Map("`t", 0, ";", 0, ",", 0), quoted := false, seen := false
        Loop Parse, text {
            if (A_LoopField = '"')
                quoted := !quoted
            else if (!quoted && A_LoopField = "`n") {
                if seen                                                     ; blank lines before it don't count as the first line
                    break
            } else if !quoted
                seen := true
            if (!quoted && counts.Has(A_LoopField))
                counts[A_LoopField] += 1
        }
        if (counts["`t"] && counts["`t"] >= counts[";"] && counts["`t"] >= counts[","])
            return "`t"
        return (counts[";"] > counts[","]) ? ";" : ","
    }

    static ToCsv(rows) {
        lines := []
        for row in rows {
            cells := []
            for cell in row {
                cell := cell ""
                if RegExMatch(cell, '[",\r\n]|^\s|\s$')
                    cell := '"' StrReplace(cell, '"', '""') '"'
                cells.Push(cell)
            }
            line := ""
            for index, cell in cells
                line .= (index > 1 ? "," : "") cell
            lines.Push(line)
        }
        text := ""
        for line in lines
            text .= line "`r`n"
        return text
    }

    static ReadFile(path) {
        text := FileRead(path, "UTF-8")
        return InStr(text, Chr(0xFFFD)) ? FileRead(path, "CP0") : text   ; bytes that don't decode: not UTF-8
    }

    static _AddRow(rows, row) {
        for cell in row
            if (cell != "")
                return rows.Push(row)
    }
}
