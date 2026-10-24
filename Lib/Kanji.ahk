;===============================================================================
; Kanji.ahk - converts Chinese between Simplified and Traditional, character by character (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Ported from the Plugins/Kanji.ahk plugin of the old ALTRun (AutoHotkey v1); the table originally comes from
; https://autohotkey.com/boards/viewtopic.php?t=9133 - looked up in a local text file,
; no external API or network requests.
;
; Resources/Kanji.txt: one line of space-separated "Simplified + Traditional" pairs, e.g. "锕錒 皑皚 嗳噯"
; - the first character of each pair is Simplified, the second the matching Traditional one (a character table, not a dictionary:
; one Simplified character maps to one Traditional character; the few one-to-many cases aren't in the table, which is enough for daily use).
;
; Usage: Kanji.ToSimplified(text) / Kanji.ToTraditional(text)
;===============================================================================

Class Kanji {
    static _s2t := ""    ; Simplified character -> Traditional character
    static _t2s := ""    ; Traditional character -> Simplified character

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
