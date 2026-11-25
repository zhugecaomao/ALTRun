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
        out := "", i := 1
        while (i <= StrLen(text)) {
            char := TextTools.CharAt(text, i)                               ; an emoji stays one character, not two swapped halves
            out := char . out, i += StrLen(char)
        }
        return out
    }

    ; The character at UTF-16 position i: both halves of a surrogate pair (an emoji, a CJK Extension B character) together. Loop Parse and
    ; SubStr count UTF-16 units and split such a pair, and a lone half makes RegExMatch throw (PCRE error -10, #653)
    static CharAt(text, i) {
        high := Ord(SubStr(text, i, 1))
        if (high >= 0xD800 && high <= 0xDBFF) {
            low := Ord(SubStr(text, i + 1, 1))
            if (low >= 0xDC00 && low <= 0xDFFF)
                return SubStr(text, i, 2)
        }
        return SubStr(text, i, 1)
    }

    ; True for half of a surrogate pair standing alone (cut off, or so in the data): not a character RegExMatch accepts
    static IsLoneSurrogate(char) => StrLen(char) = 1 && Ord(char) >= 0xD800 && Ord(char) <= 0xDFFF

    ; The first count UTF-16 units of text, one fewer when the last would be the first half of a surrogate pair (for "..." previews)
    static Head(text, count) {
        head := SubStr(text, 1, count), last := Ord(SubStr(head, -1))
        return (StrLen(head) = count && last >= 0xD800 && last <= 0xDBFF) ? SubStr(head, 1, count - 1) : head
    }

    static Lines(text) => StrSplit(text, "`n", "`r")

    ; "" when pattern is a valid regular expression (or empty), otherwise the error, e.g. "Compile error 14 at offset 4: missing )"
    static RegexError(pattern) {
        try RegExMatch("", pattern)
        catch as e
            return e.Message
        return ""
    }

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
