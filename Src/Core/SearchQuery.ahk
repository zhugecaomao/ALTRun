;===============================================================================
; SearchQuery.ahk - parses the text in the search box (AutoHotkey v2)
;-------------------------------------------------------------------------------
; "g hello world" -> Keyword = "g", Rest = "hello world", HasRest = true
; "clip "         -> Keyword = "clip", Rest = "", HasRest = true   (a space after the keyword)
; "notepad"       -> Keyword = "notepad", Rest = "", HasRest = false
;
; Usage:
;   q := SearchQuery("g hello")
;   if q.MatchKeyword(["g", "google"], &term)   ; term = "hello"
;   SearchQuery("scrip").StartsKeyword(["script"], &keyword)   ; true, keyword = "script" (a keyword not typed in full yet)
;   SearchQuery.Keywords("g, ㅎ, google")        ["g", "ㅎ", "google"] (the keyword field can hold several, separated by commas)
;===============================================================================

class SearchQuery {
    __New(raw) {
        this.Raw  := raw
        this.Text := Trim(raw)
        leading := LTrim(raw)                                               ; keep a trailing space: "clip " means keyword mode has started
        spacePos := InStr(leading, " ")
        this.Keyword := StrLower(spacePos ? SubStr(leading, 1, spacePos - 1) : this.Text)
        this.Rest    := spacePos ? Trim(SubStr(leading, spacePos + 1)) : ""
        this.HasRest := spacePos > 0
    }

    ; true when the first word is one of keywords (case-insensitive); the text after it goes into term, the matched keyword into matched
    MatchKeyword(keywords, &term, &matched := "") {
        term := "", matched := ""
        for keyword in keywords {
            if (keyword != "" && this.Keyword = keyword) {
                term := this.Rest, matched := keyword
                return true
            }
        }
        return false
    }

    ; true when the text (one word, at least PrefixMin letters, no space yet) is the start of a longer keyword, put into keyword
    static PrefixMin := 3
    StartsKeyword(keywords, &keyword) {
        keyword := ""
        length := StrLen(this.Text)
        if (this.HasRest || length < SearchQuery.PrefixMin)
            return false
        for candidate in keywords {
            if (StrLen(candidate) > length && SubStr(candidate, 1, length) = this.Text) {
                keyword := StrLower(candidate)
                return true
            }
        }
        return false
    }

    ; The keyword field of web searches, custom commands and snippets can hold several, separated by commas (half- or full-width), e.g. "g, ㅎ, google"
    ; (works whichever keyboard is active, Chinese / English / Korean). Split, trimmed, empty ones dropped. Cached per text (used on every key press): don't change the result
    static Keywords(text) {
        static cache := Map()
        text := text ""
        if !cache.Has(text) {
            cache[text] := TextTools.SplitList(StrReplace(text, "，", ","))   ; a full-width comma separates too
        }
        return cache[text]
    }

    ; true when it starts with prefix (no space needed after the prefix, e.g. ">dir" or "'report")
    MatchPrefix(prefix, &term) {
        term := ""
        if (prefix = "" || SubStr(this.Text, 1, StrLen(prefix)) != prefix)
            return false
        term := Trim(SubStr(this.Text, StrLen(prefix) + 1))
        return true
    }
}
