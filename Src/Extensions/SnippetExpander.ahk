;===============================================================================
; SnippetExpander.ahk - auto expansion of snippet keywords (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like Alfred's snippet auto expansion: typing "prefix + keyword" in any program (default prefix
; ";", e.g. ";sig") deletes the typed text and replaces it with the snippet's text (placeholders are replaced on expansion).
; The prefix avoids triggering by accident in normal typing. Not expanded in: ALTRun's own search window, the windows in ExpandExclude
; (e.g. remote desktop, password managers), and Windows password boxes.
;
; Settings:
;   Features.Snippets.AutoExpand      1 = on
;   Features.Snippets.ExpandPrefix    the prefix before the keyword, default ";"
;   Features.Snippets.ExpandExclude   windows to leave alone, an array (ahk_exe xxx.exe / ahk_class xxx / part of the title)
;   Snippets[n].AutoExpand            0 turns auto expansion off for one snippet (it can still be searched)
;   Snippets[n].Keyword               can hold several, separated by commas ("sig, signature"), each expands
; Call Refresh() after snippets change, and the new keywords work right away.
;
; Usage:
;   SnippetExpander.Init()            at start-up, after SearchWindow.Create()
;===============================================================================

class SnippetExpander {
    static Count := 0
    static MaxAbbreviation := 40                                            ; AHK hotstring abbreviations (prefix + keyword) are at most 40 characters
    static _texts := Map()                                                  ; registered abbreviation -> text
    static _exclude := []                                                   ; windows where nothing expands
    static _criterion := (*) => SnippetExpander.ShouldExpand()              ; all hotstrings share one condition (turning them off needs the same one)

    static Init() {
        SnippetExpander.Refresh()
    }

    ; Registers again from the current AppSettings.Snippets: removed abbreviations are turned off, new / changed ones registered.
    ; Called after a snippet is edited / deleted in the results, without reloading ALTRun.
    static Refresh() {
        options := AppSettings.Feature("Snippets")
        wanted := Map()
        wanted.CaseSense := "Off"                                           ; case-insensitive like the hotstrings
        if (options["Enabled"] && options["AutoExpand"]) {
            for snippet in AppSettings.Snippets {
                for abbreviation in SnippetExpander.Abbreviations(snippet, options["ExpandPrefix"])
                    wanted[abbreviation] := snippet["Text"]
            }
        }
        SnippetExpander._exclude := options["ExpandExclude"]
        HotIf(SnippetExpander._criterion)
        for abbreviation in SnippetExpander._texts
            if !wanted.Has(abbreviation)
                try Hotstring(":*?:" abbreviation, , "Off")
        registered := Map()
        registered.CaseSense := "Off"
        for abbreviation, text in wanted {
            try {
                ; * = no end character needed, ? = also fires right after other characters; AHK deletes the typed abbreviation itself
                Hotstring(":*?:" abbreviation, SnippetExpander._Expander(abbreviation), "On")
                registered[abbreviation] := text
            } catch as e {
                Logger.Error("SnippetExpander: cannot register " abbreviation " - " e.Message)
            }
        }
        HotIf()
        SnippetExpander._texts := registered
        SnippetExpander.Count := registered.Count
        Logger.Debug("SnippetExpander: " SnippetExpander.Count " snippets")
    }

    ; Checked only when an abbreviation was typed (not on every key): nothing expands in the search window, the windows in ExpandExclude or password boxes (ES_PASSWORD)
    static ShouldExpand() {
        if (IsObject(SearchWindow.Gui) && WinActive("ahk_id " SearchWindow.Gui.Hwnd))
            return false
        for window in SnippetExpander._exclude
            try if WinActive(window)
                return false
        try {
            focused := ControlGetFocus("A")
            if (InStr(ControlGetClassNN(focused), "Edit") = 1 && (ControlGetStyle(focused) & 0x20))
                return false
        }
        return true
    }

    ; Snippet -> the abbreviations that trigger it, one per keyword ([] = no auto expansion)
    static Abbreviations(snippet, prefix) {
        abbreviations := []
        if !(snippet is Map) || !snippet.Has("Keyword") || !snippet.Has("Text")
            return abbreviations
        if (snippet.Has("AutoExpand") && !snippet["AutoExpand"])
            return abbreviations
        for keyword in SearchQuery.Keywords(snippet["Keyword"]) {
            if RegExMatch(keyword, "[\s``]")                                ; white space and backticks can't be used in abbreviations
                continue
            if (StrLen(prefix keyword) > SnippetExpander.MaxAbbreviation)   ; too long to register: search only (the editor already warned)
                continue
            abbreviations.Push(prefix keyword)
        }
        return abbreviations
    }

    ; a separate method makes the closure; the text is looked up when it fires, so nothing needs regenerating after a snippet changes
    static _Expander(abbreviation) {
        return (*) => (Usage.Count("SnippetExpand"), SnippetProvider.Paste(SnippetExpander._texts[abbreviation], false))
    }
}
