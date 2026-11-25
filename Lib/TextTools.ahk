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

    ; AutoHotkey's RegExMatch / RegExReplace throw (PCRE error -10) when the text STARTS with half a surrogate pair or a character in
    ; U+F800-U+FFFF (fullwidth ，（）Ａ１, halfwidth katakana, U+FFFD), or ENDS with a first half or a character in U+F800-U+FBFF (CJK
    ; compatibility ideographs like 﨑, the ﬁ ligature), wherever the match starts; the same characters in between are fine (#653).
    ; Match / Replace put U+0001 at such an end and take it out of the result: the same result as RegExMatch / RegExReplace for patterns
    ; that treat U+0001 like that character (".*", "\S+", "[^,]+", ...) or don't look at the very start / end; not for one that needs
    ; that character itself right after ^ or before $ (e.g. "^\p{L}"). Match: m is then a TextTools.MatchInfo (m[n], m.Pos(n),
    ; m.Len(n), m.Count, m.Name(n), m.Value(n), named groups) without U+0001
    static RegExUnsafe(text) => TextTools._UnsafeStart(text) || TextTools._UnsafeEnd(text)
    static _UnsafeStart(text) => ((cp := Ord(text)) >= 0xD800 && cp <= 0xDFFF) || (cp >= 0xF800 && cp <= 0xFFFF)   ; Ord: a whole pair is one code point
    static _UnsafeEnd(text) => ((u := Ord(SubStr(text, -1))) >= 0xD800 && u <= 0xDBFF) || (u >= 0xF800 && u <= 0xFBFF)

    static Match(text, pattern, &m := "", start := 1) {
        if !TextTools.RegExUnsafe(text)
            return RegExMatch(text, pattern, &m, start)
        before := TextTools._UnsafeStart(text) ? 1 : 0
        padded := (before ? Chr(1) : "") text (TextTools._UnsafeEnd(text) ? Chr(1) : "")
        if !RegExMatch(padded, pattern, &found, start + before * (start > 1))
            return 0
        m := TextTools.MatchInfo(found, before, StrLen(text))
        return m.Pos(0)
    }

    ; A RegExMatchInfo of the text with U+0001 added (Match): positions and values of the text itself
    class MatchInfo {
        __New(found, before, length) => (this._found := found, this._before := before, this._length := length)
        Pos(n := 0) => (this._found.Len(n) = 0 && this._found.Pos(n) = 0) ? 0 : Min(Max(this._found.Pos(n) - this._before, 1), this._length + 1)
        Len(n := 0) => StrLen(this.Value(n))
        Value(n := 0) {
            start := this._found.Pos(n) - this._before, last := start + this._found.Len(n) - 1   ; in the text: 0 or length + 1 = U+0001
            return (this._found.Len(n) = 0) ? "" : SubStr(this._found.Value(n), (start < 1) ? 2 : 1, Min(last, this._length) - Max(start, 1) + 1)
        }
        Name(n) => this._found.Name(n)
        Count => this._found.Count
        Mark => this._found.Mark
        __Item[n := 0] => this.Value(n)
        __Get(name, params) => this.Value(name)
    }

    static Replace(text, pattern, replacement := "") {
        if !TextTools.RegExUnsafe(text)
            return RegExReplace(text, pattern, replacement)
        before := TextTools._UnsafeStart(text), after := TextTools._UnsafeEnd(text)
        out := RegExReplace((before ? Chr(1) : "") text (after ? Chr(1) : ""), pattern, replacement)
        if (before && SubStr(out, 1, 1) = Chr(1))                           ; still there: not part of a match
            out := SubStr(out, 2)
        if (after && SubStr(out, -1) = Chr(1))
            out := SubStr(out, 1, -1)
        return out
    }

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
