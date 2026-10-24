;===============================================================================
; TypeAhead.ahk - keeps what is typed while the search window is still opening (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Usage (SearchWindow.Show):
;   TypeAhead.Start()            before showing
;   typed := TypeAhead.Take()    once the search box has focus -> {Text, Key, Modifier}
;   TypeAhead.Summary(typed)     for the log: how much was typed, never what
;===============================================================================

class TypeAhead {
    ; From the hotkey press until the search box has focus takes tens of milliseconds (first the empty box is searched, then the window shown; longer right after boot), and characters typed
    ; in that time used to go to the foreground program, lost or typed into another window. Now InputHook catches and keeps the keys that type characters during that time
    ; (Ctrl / Alt combinations, arrow keys etc. as usual) and adds them to the search box once it has focus; an Enter pressed meanwhile runs the first result after the search, Esc closes the window.
    ; At most TimeoutMs: if showing hangs, typing isn't held forever
    static TimeoutMs := 3000
    static _hook := ""

    static Start() {
        if (IsObject(TypeAhead._hook) || SearchWindow.IsActive())
            return
        try {
            ih := InputHook("T" TypeAhead.TimeoutMs / 1000)
            ih.KeyOpt("{Enter}{NumpadEnter}{Esc}", "ES")
            ih.Start()
            TypeAhead._hook := ih
        }
    }

    ; Stops recording; returns {Text: the typed characters, Key: "Enter" / "NumpadEnter" / "Escape" / "", Modifier: Ctrl / Alt held with Enter}
    static Take() {
        ih := TypeAhead._hook, TypeAhead._hook := ""
        if !IsObject(ih)
            return {Text: "", Key: "", Modifier: ""}
        try ih.Stop()
        key := (ih.EndReason = "EndKey") ? ih.EndKey : ""
        modifier := (key = "") ? "" : InStr(ih.EndMods, "^") ? "ctrl" : InStr(ih.EndMods, "!") ? "alt" : ""
        return {Text: RegExReplace(ih.Input, "[\x00-\x1F\x7F]"), Key: key, Modifier: modifier}   ; no Tab or other control characters
    }

    ; For the log: how much was typed, not what (the log never holds the search text), e.g. "5 chars (leading space) + Enter"
    static Summary(typed) {
        return StrLen(typed.Text) " chars" (SubStr(typed.Text, 1, 1) = " " ? " (leading space)" : "") (typed.Key != "" ? " + " typed.Key : "")
    }
}
