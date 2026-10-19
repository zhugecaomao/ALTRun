;===============================================================================
; DelimitedText.ahk - 表格文字 (CSV / TSV / 从 Excel 复制的内容) <-> 二维数组 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 纯函数。分隔符自动判断 (第一行里 Tab 最多就是 Tab, 其次 ; 和 ,); 双引号里的分隔符、换行照原样保留,
; "" 表示一个 ", 和 Excel 复制 / 保存的格式一样。
;
; 用法:
;   rows := DelimitedText.Parse(text)           [["名称", "关键字"], ["邮箱", "em"]...], 空行跳过
;   csv  := DelimitedText.ToCsv(rows)           逗号分隔, 需要时加引号, 行之间 `r`n
;   text := DelimitedText.ReadFile(path)        UTF-8 (有没有 BOM 都行), 不是 UTF-8 时按系统编码 (例如 GBK)
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
            } else if (char = '"' && field = "") {                          ; 只有字段开头的引号才算
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

    ; 第一个不空的行 (引号外) 里 Tab 最多就用 Tab, 否则 ; 比 , 多时用 ; (欧洲的 Excel), 默认 ,
    static Delimiter(text) {
        counts := Map("`t", 0, ";", 0, ",", 0), quoted := false, seen := false
        Loop Parse, text {
            if (A_LoopField = '"')
                quoted := !quoted
            else if (!quoted && A_LoopField = "`n") {
                if seen                                                     ; 前面的空行不算第一行
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
        return InStr(text, Chr(0xFFFD)) ? FileRead(path, "CP0") : text   ; 有解不出来的字节: 不是 UTF-8
    }

    static _AddRow(rows, row) {
        for cell in row
            if (cell != "")
                return rows.Push(row)
    }
}
