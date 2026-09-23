;===============================================================================
; FuzzyMatcher.ahk - 搜索匹配打分 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 返回 0 (不匹配) ~ 100 (完全相同), 规则按 Alfred 的习惯:
;   100  完全相同                     "notepad"  -> Notepad
;    90  开头相同                     "note"     -> Notepad
;    80  某个单词的开头               "code"     -> Visual Studio Code
;    75  单词首字母                   "vsc"      -> Visual Studio Code
;    60  包含 (越靠前分越高)          "pad"      -> Notepad
;    45  多个关键词都能匹配上         "st co"    -> Visual Studio Code
;    30  按顺序出现的字母 (至少 3 个, 首字母相同)  "ntpd" -> Notepad
; 同一档里, 文字越短分数越高 (更精确)。都不区分大小写。
;
; 性能: 每次按键都要给几千个候选打分, 所以候选文字预先算好 SearchKey (小写、
; 单词边界、首字母), 打分时只剩几次 InStr, 不用正则。不匹配的候选 (绝大多数)
; 一般两次 InStr 就被排除。
;
; 用法:
;   key := FuzzyMatcher.Key("Visual Studio Code")          候选文字建一次, 缓存起来
;   FuzzyMatcher.ScoreKey("vsc", key)                      needle 必须已经是小写
;   FuzzyMatcher.BestKey("wx", [titleKey, pinyinKey])      多个候选取最高分
;   FuzzyMatcher.Score("vsc", "Visual Studio Code")        一次性打分 (内部建 Key)
;   FuzzyMatcher.TopIndexes(scores, limit)                 分数数组 -> 分数最高的 limit 个下标
;===============================================================================

class FuzzyMatcher {

    static Score(needle, text) {
        needle := StrLower(Trim(needle))
        if (needle = "")
            return 1
        return FuzzyMatcher.ScoreKey(needle, FuzzyMatcher.Key(text))
    }

    static Best(needle, candidates) {
        needle := StrLower(Trim(needle))
        best := 0
        for text in candidates {
            if (text = "")
                continue
            score := (needle = "") ? 1 : FuzzyMatcher.ScoreKey(needle, FuzzyMatcher.Key(text))
            if (score > best)
                best := score
        }
        return best
    }

    ; 预先计算的候选文字: Lower = 小写原文, Words = " " + 单词之间只用空格分隔的小写文字,
    ; Initials = 各单词首字母 (驼峰也算单词)
    static Key(text) {
        lower := StrLower(text)
        words := " " RegExReplace(lower, "[\s\-_.\\/()\[\]]+", " ")
        return {Lower: lower, Words: words, Initials: FuzzyMatcher.Initials(text), Length: StrLen(lower)}
    }

    static BestKey(needle, keys) {
        best := 0
        for key in keys {
            if !IsObject(key)
                continue
            score := FuzzyMatcher.ScoreKey(needle, key)
            if (score > best)
                best := score
        }
        return best
    }

    ; needle: 已经 Trim + 小写
    static ScoreKey(needle, key) {
        if (needle = "")
            return 1
        h := key.Lower
        if (h = "")
            return 0
        needleLength := StrLen(needle)
        lengthPenalty := Min(Max(key.Length - needleLength, 0), 40) * 0.1  ; 最多扣 4 分

        if (pos := InStr(h, needle)) {
            if (key.Length = needleLength)
                return 100
            if (pos = 1)
                return 90 - lengthPenalty
            if (InStr(key.Words, " " needle) || InStr(" -_.\/()[]", SubStr(h, pos - 1, 1)))   ; 第二个条件: needle 本身带分隔符时
                return 80 - lengthPenalty
            return 60 - Min(pos, 20) * 0.25 - lengthPenalty
        }
        if (needleLength >= 2 && InStr(key.Initials, needle) = 1)
            return 75 - lengthPenalty

        if InStr(needle, " ") {                                             ; 多个关键词: 每个都要能匹配上
            total := 0, count := 0
            for token in StrSplit(needle, " ") {
                if (token = "")
                    continue
                tokenScore := FuzzyMatcher.ScoreKey(token, key)
                if (tokenScore <= 1)
                    return 0
                total += tokenScore, count++
            }
            return count ? Min(total / count, 90) * 0.5 : 0
        }

        ; 按顺序出现的字母: 至少 3 个, 且第一个字母相同, 否则会命中太多无关项
        if (needleLength >= 3 && SubStr(needle, 1, 1) = SubStr(h, 1, 1) && FuzzyMatcher.IsSubsequence(needle, h))
            return 30 - lengthPenalty
        return 0
    }

    ; "Visual Studio Code" -> "vsc"; "my_file-name" -> "mfn"; "AutoHotkey" -> "ah" (驼峰也算单词)
    static Initials(text) {
        text := RegExReplace(text, "([a-z])([A-Z])", "$1 $2")
        initials := ""
        for word in StrSplit(RegExReplace(text, "[^\p{L}\p{N}]+", " "), " ")
            if (word != "")
                initials .= SubStr(word, 1, 1)
        return StrLower(initials)
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

    ; scores: 下标 -> 分数 (Map), 返回分数最高的 limit 个下标 (分数相同时下标小的在前)
    static TopIndexes(scores, limit) {
        if !scores.Count
            return []
        lines := ""
        for index, score in scores
            lines .= Format("{:09.4f}", score) "`t" Format("{:07}", 9999999 - index) "`n"
        indexes := []
        for line in StrSplit(Sort(RTrim(lines, "`n"), "R"), "`n") {
            indexes.Push(9999999 - Integer(StrSplit(line, "`t")[2]))
            if (indexes.Length >= limit)
                break
        }
        return indexes
    }
}
