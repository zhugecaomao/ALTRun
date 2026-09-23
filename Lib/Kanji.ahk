;===============================================================================
; Kanji.ahk - 简体/繁体中文逐字互转 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 移植自旧版 ALTRun (AutoHotkey v1) 的 Plugins/Kanji.ahk 插件, 对照表原始来源
; 见 https://autohotkey.com/boards/viewtopic.php?t=9133 - 本地文本文件查表,
; 不依赖任何外部 API/网络请求。
;
; Resources/Kanji.txt: 一整行, 空格分隔的"简体+繁体"两字词, 例如 "锕錒 皑皚 嗳噯"
; - 每个词第 1 个字符是简体, 第 2 个字符是对应的繁体(逐字对照, 不是词典,
; 一简体只对应一个繁体, 少数一简对多繁的情况这份表没有收录, 够日常使用)。
;
; 用法: Kanji.ToSimplified(text) / Kanji.ToTraditional(text)
;===============================================================================

Class Kanji {
    static _s2t := ""    ; 简体字符 -> 繁体字符
    static _t2s := ""    ; 繁体字符 -> 简体字符

    static _Load() {
        if (Kanji._s2t != "")
            return
        Kanji._s2t := Map()
        Kanji._t2s := Map()

        dataFile := A_ScriptDir "\Resources\Kanji.txt"
        if !FileExist(dataFile) {
            Logger.Debug("Kanji: Resources\Kanji.txt not found, simplified/traditional conversion disabled")
            return
        }
        for _, token in StrSplit(FileRead(dataFile, "UTF-8"), " ", "`r`n") {
            if (StrLen(token) != 2)
                continue
            Kanji._s2t[SubStr(token, 1, 1)] := SubStr(token, 2, 1)
            Kanji._t2s[SubStr(token, 2, 1)] := SubStr(token, 1, 1)
        }
    }

    static ToTraditional(text) {
        Kanji._Load()
        return Kanji._Convert(text, Kanji._s2t)
    }

    static ToSimplified(text) {
        Kanji._Load()
        return Kanji._Convert(text, Kanji._t2s)
    }

    static _Convert(text, map) {
        out := ""
        Loop Parse text
            out .= map.Has(A_LoopField) ? map[A_LoopField] : A_LoopField
        return out
    }
}
