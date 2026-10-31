;===============================================================================
; SearchSession.ahk - the state of the search window, apart from the window itself (AutoHotkey v2)
;-------------------------------------------------------------------------------
; What is searched and shown: the mode, the results and the selection, the actions panel, marked files, history browsing and the
; search kept from the last showing. SearchWindow draws it and turns keys and clicks into changes of it; nothing here touches a
; window, so it can be tested without one.
;
; Moved here step by step: the code reads and changes the session (SearchWindow.Session.Mode ...); SearchWindow still has the old
; names (SearchWindow.Mode, .Results ...) as properties forwarding here, only for the tests and the benchmark probe until they
; use the session too. Next steps: the changes between states become methods here, then searches become requests whose results
; come back with their number (for searching in the background).
;
; Usage:
;   session := SearchWindow.Session
;   session.Mode / Results / Selected / Offset ...
;===============================================================================

class SearchSession {
    ; Results
    Mode          := "results"   ; "results" = search results; "actions" = the actions panel
    FileMode      := false       ; file search mode: files only (Space in the empty search box enters it)
    Results       := []          ; the ResultItems shown (all of them; the list shows VisibleRows from Offset)
    Selected      := 0           ; the selected result, 1-based in Results (0 = none)
    Offset        := 0           ; the first result shown in the list is Results[Offset + 1]
    HighlightText := ""          ; characters in result titles matching it are highlighted (when browsing a folder only the filter part)
    SearchGen     := 0           ; counts the searches: a search finding a newer one started meanwhile drops its results

    ; Actions panel
    ActionSource  := ""          ; the result the panel was opened for
    SavedQuery    := ""          ; the search box text before the panel opened (put back when it closes)
    AllActions    := []          ; every action of the panel (typing filters them)
    ActionsOnly   := false       ; the panel was opened directly (actions for the selection): Esc / ← close the window

    ; Marked files / folders (Insert): lower-case path -> path; cleared each time the window is shown
    Marked        := Map()

    ; History browsing (Ctrl+↑ / Ctrl+↓)
    HistoryIndex  := 0           ; which entry of Knowledge.History is shown (1 = most recent), 0 = not browsing
    HistoryDraft  := ""          ; the text in the box before browsing; Ctrl+↓ past the newest entry brings it back

    ; Keeping the last search (KeepLastQuery)
    Last              := ""      ; the search at the last hide: {Text, FileMode, Selected, Time}
    RestoredSelection := false   ; the restored search is still all selected: typing over it starts a new, normal search
}
