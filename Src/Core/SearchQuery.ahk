;===============================================================================
; SearchQuery.ahk - 解析搜索框里的文字 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; "g hello world" -> Keyword = "g", Rest = "hello world", HasRest = true
; "notepad"       -> Keyword = "notepad", Rest = "", HasRest = false
;
; 用法:
;   q := SearchQuery("g hello")
;   if q.MatchKeyword(["g", "google"], &term)   ; term = "hello"
;===============================================================================

class SearchQuery {
    __New(raw) {
        this.Raw  := raw
        this.Text := Trim(raw)
        spacePos := InStr(this.Text, " ")
        this.Keyword := StrLower(spacePos ? SubStr(this.Text, 1, spacePos - 1) : this.Text)
        this.Rest    := spacePos ? Trim(SubStr(this.Text, spacePos + 1)) : ""
        this.HasRest := spacePos > 0
    }

    ; 第一个词是 keywords 之一 (不区分大小写) 时返回 true, 并把后面的文字放进 term
    MatchKeyword(keywords, &term) {
        term := ""
        for keyword in keywords {
            if (keyword != "" && this.Keyword = keyword) {
                term := this.Rest
                return true
            }
        }
        return false
    }

    ; 以 prefix 开头 (前缀后面不需要空格, 例如 ">dir" 或 "'report") 时返回 true
    MatchPrefix(prefix, &term) {
        term := ""
        if (prefix = "" || SubStr(this.Text, 1, StrLen(prefix)) != prefix)
            return false
        term := Trim(SubStr(this.Text, StrLen(prefix) + 1))
        return true
    }
}
