;===============================================================================
; TextTools.ahk - 文字转换 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 纯函数, 输入一段文字返回转换后的文字, 不碰剪贴板。统一按 `n / `r`n 拆行,
; 合并时统一用 `r`n。
;
; 用法: TextTools.SortLines(text) / TextTools.DedupeLines(text) / ...
;===============================================================================

class TextTools {
    static Upper(text)     => StrUpper(text)
    static Lower(text)     => StrLower(text)
    static TitleCase(text) => StrTitle(text)

    static Reverse(text) {
        out := ""
        Loop Parse text
            out := A_LoopField . out
        return out
    }

    static Lines(text) => StrSplit(text, "`n", "`r")

    ; One item per line: trimmed, blank lines dropped (the lists edited in Preferences)
    static SplitLines(text) {
        items := []
        for line in StrSplit(text, "`n", "`r")
            if ((line := Trim(line)) != "")
                items.Push(line)
        return items
    }

    static Join(lines) => TextTools.JoinList(lines, "`r`n")

    ; Comma-separated items (a one-line list in Preferences, an old settings value): trimmed, blanks dropped;
    ; numeric: numbers become numbers ("10, 12" -> [10, 12])
    static SplitList(text, numeric := false) {
        items := []
        for item in StrSplit(text, ",")
            if ((item := Trim(item)) != "")
                items.Push((numeric && IsNumber(item)) ? item + 0 : item)
        return items
    }

    ; The items joined with separator (default ", "); a value that isn't an array is returned as it is
    static JoinList(items, separator := ", ") {
        if !(items is Array)
            return items
        out := ""
        for index, item in items
            out .= (index = 1 ? "" : separator) item
        return out
    }

    static SortLines(text, descending := false) {
        return TextTools.Join(TextTools.Lines(Sort(StrReplace(text, "`r`n", "`n"), descending ? "R" : "")))
    }

    ; 逐行去掉首尾空格/Tab, 保留空行本身
    static TrimLines(text) {
        trimmed := []
        for line in TextTools.Lines(text)
            trimmed.Push(Trim(line, " `t"))
        return TextTools.Join(trimmed)
    }

    static RemoveBlankLines(text) {
        kept := []
        for line in TextTools.Lines(text)
            if (Trim(line) != "")
                kept.Push(line)
        return TextTools.Join(kept)
    }

    ; 按原顺序去重 (Sort 的 U 选项会顺带重新排序)
    static DedupeLines(text) {
        seen := Map(), kept := []
        for line in TextTools.Lines(text) {
            if seen.Has(line)
                continue
            seen[line] := true
            kept.Push(line)
        }
        return TextTools.Join(kept)
    }
}
