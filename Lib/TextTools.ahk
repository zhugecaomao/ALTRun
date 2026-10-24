;===============================================================================
; TextTools.ahk - text conversions (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Pure functions: take text, return the converted text, never touch the clipboard. Lines are split at `n / `r`n
; and joined with `r`n.
;
; Usage: TextTools.SortLines(text) / TextTools.DedupeLines(text) / ...
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

    ; Trims spaces / Tabs at both ends of each line; blank lines stay
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

    ; Removes duplicates keeping the order (Sort's U option would also re-sort)
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
