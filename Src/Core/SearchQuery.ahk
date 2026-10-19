;===============================================================================
; SearchQuery.ahk - 解析搜索框里的文字 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; "g hello world" -> Keyword = "g", Rest = "hello world", HasRest = true
; "clip "         -> Keyword = "clip", Rest = "", HasRest = true   (关键字后面有空格)
; "notepad"       -> Keyword = "notepad", Rest = "", HasRest = false
;
; 用法:
;   q := SearchQuery("g hello")
;   if q.MatchKeyword(["g", "google"], &term)   ; term = "hello"
;   SearchQuery.Keywords("g, ㅎ, google")        ["g", "ㅎ", "google"] (关键字一栏可以写几个, 用逗号分隔)
;===============================================================================

class SearchQuery {
    __New(raw) {
        this.Raw  := raw
        this.Text := Trim(raw)
        leading := LTrim(raw)                                               ; 保留结尾空格: "clip " 表示已经进入关键字模式
        spacePos := InStr(leading, " ")
        this.Keyword := StrLower(spacePos ? SubStr(leading, 1, spacePos - 1) : this.Text)
        this.Rest    := spacePos ? Trim(SubStr(leading, spacePos + 1)) : ""
        this.HasRest := spacePos > 0
    }

    ; 第一个词是 keywords 之一 (不区分大小写) 时返回 true, 并把后面的文字放进 term, 对上的关键字放进 matched
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

    ; 网页搜索、自定义命令、片段的关键字一栏可以写几个, 用逗号 (半角或全角) 分隔, 例如 "g, ㅎ, google"
    ; (中英文 / 韩文键盘切换时都能用)。拆开、去掉空白, 空的不要。结果按原文缓存 (每次输入都要用), 不要修改
    static Keywords(text) {
        static cache := Map()
        text := text ""
        if !cache.Has(text) {
            keywords := []
            for keyword in StrSplit(StrReplace(text, "，", ","), ",")
                if ((keyword := Trim(keyword)) != "")
                    keywords.Push(keyword)
            cache[text] := keywords
        }
        return cache[text]
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
