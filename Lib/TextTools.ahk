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

    static Join(lines) {
        out := ""
        for index, line in lines
            out .= (index = 1 ? "" : "`r`n") . line
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
