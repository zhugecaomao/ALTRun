;===============================================================================
; FuzzyMatcher.ahk - 搜索匹配打分 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Score(needle, text) 返回 0 (不匹配) ~ 100 (完全相同), 规则按 Alfred 的习惯:
;   100  完全相同                     "notepad"  -> Notepad
;    90  开头相同                     "note"     -> Notepad
;    80  某个单词的开头               "code"     -> Visual Studio Code
;    75  单词首字母                   "vsc"      -> Visual Studio Code
;    60  包含 (越靠前分越高)          "pad"      -> Notepad
;    45  多个关键词都能匹配上         "st co"    -> Visual Studio Code
;    30  按顺序出现的字母 (至少 3 个, 首字母相同)  "ntpd" -> Notepad
; 同一档里, 文字越短分数越高 (更精确)。都不区分大小写。
;
; 用法:
;   FuzzyMatcher.Score("vsc", "Visual Studio Code")
;   FuzzyMatcher.Best("wx", ["微信", "WX"])      取多个候选文字里的最高分
;===============================================================================

class FuzzyMatcher {

    static Best(needle, candidates) {
        best := 0
        for text in candidates {
            if (text = "")
                continue
            score := FuzzyMatcher.Score(needle, text)
            if (score > best)
                best := score
        }
        return best
    }

    static Score(needle, text) {
        needle := Trim(needle)
        if (needle = "")
            return 1
        if (text = "")
            return 0

        n := StrLower(needle)
        h := StrLower(text)
        lengthPenalty := Min(Max(StrLen(h) - StrLen(n), 0), 40) * 0.1      ; 最多扣 4 分

        if (h = n)
            return 100
        if (SubStr(h, 1, StrLen(n)) = n)
            return 90 - lengthPenalty
        if RegExMatch(h, "(?:^|[\s\-_.\\/()\[\]])\Q" FuzzyMatcher._QuoteSafe(n) "\E")
            return 80 - lengthPenalty
        if (!InStr(n, " ") && StrLen(n) >= 2 && InStr(FuzzyMatcher.Initials(text), n) = 1)   ; 用原文, 驼峰需要大小写
            return 75 - lengthPenalty
        if (pos := InStr(h, n))
            return 60 - Min(pos, 20) * 0.25 - lengthPenalty

        if InStr(n, " ") {                                                  ; 多个关键词: 每个都要能匹配上
            total := 0, count := 0
            for token in StrSplit(n, " ") {
                if (token = "")
                    continue
                tokenScore := FuzzyMatcher.Score(token, h)
                if (tokenScore <= 1)
                    return 0
                total += tokenScore, count++
            }
            return count ? Min(total / count, 90) * 0.5 : 0
        }

        ; 按顺序出现的字母: 至少 3 个, 且第一个字母相同, 否则会命中太多无关项
        if (StrLen(n) >= 3 && SubStr(n, 1, 1) = SubStr(h, 1, 1) && FuzzyMatcher.IsSubsequence(n, h))
            return 30 - lengthPenalty
        return 0
    }

    ; "Visual Studio Code" -> "vsc"; "my_file-name" -> "mfn"; "AutoHotkey" -> "ah" (驼峰也算单词)
    static Initials(text) {
        text := RegExReplace(text, "([a-z])([A-Z])", "$1 $2")
        out := ""
        for word in StrSplit(RegExReplace(text, "[^\p{L}\p{N}]+", " "), " ")
            if (word != "")
                out .= SubStr(word, 1, 1)
        return StrLower(out)
    }

    static IsSubsequence(needle, text) {
        pos := 1
        Loop Parse needle {
            pos := InStr(text, A_LoopField, false, pos)
            if !pos
                return false
            pos++
        }
        return true
    }

    ; \Q...\E 里如果出现 \E 会提前结束引用, 需要拆开
    static _QuoteSafe(text) {
        return StrReplace(text, "\E", "\E\\E\Q", true)                   ; 区分大小写, \e 不需要处理
    }
}
