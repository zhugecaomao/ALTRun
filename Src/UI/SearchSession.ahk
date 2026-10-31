;===============================================================================
; SearchSession.ahk - the state of the search window, apart from the window itself (AutoHotkey v2)
;-------------------------------------------------------------------------------
; What is searched and shown: the mode, the results and the selection, the actions panel, marked files, history browsing and the
; search kept from the last showing. SearchWindow draws it and turns keys and clicks into changes of it; nothing here touches a
; window, so it can be tested without one.
;
; Searches are requests (SearchRequest): the session numbers them, and results are shown only for the newest one, so a search
; finishing after a newer one started (later: in the background) is dropped.
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
    Request       := ""          ; the newest search (SearchRequest); only its results are shown

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

    ;---------------------------------------------------------------------------
    ; Showing and hiding
    ;---------------------------------------------------------------------------
    ; A new showing: search results, no marks or history browsing; restore = the last search comes back (its file search mode too,
    ; selected until the caret moves)
    StartShowing(restore) {
        this.Mode := "results", this.ActionsOnly := false, this.Marked := Map(), this.HistoryIndex := 0
        this.FileMode := restore ? this.Last.FileMode : false
        this.RestoredSelection := restore
    }

    ; At hiding: the search to bring back next time (KeepLastQuery). text = the search box text; in the actions panel the
    ; search it was opened from (its first row then)
    Remember(text, now := A_TickCount) {
        actions := (this.Mode = "actions")
        this.Last := {Text: actions ? this.SavedQuery : text, FileMode: this.FileMode, Selected: actions ? 1 : Max(1, this.Selected), Time: now}
    }

    ; Whether the last search comes back (KeepLastQuery): only within KeepLastQueryMinutes after hiding (0 = no limit), like Alfred's
    ; "show latest query if within N minutes"; after that the window opens with an empty search box
    KeepsLast(general, now := A_TickCount) {
        last := this.Last
        if !(general["KeepLastQuery"] && IsObject(last) && last.Text != "")
            return false
        minutes := general["KeepLastQueryMinutes"]
        return !IsNumber(minutes) || minutes <= 0 || now - last.Time <= minutes * 60000
    }

    ; The restored last search is still all selected and was typed over: a new, normal search (file search mode ends too), like
    ; Alfred, where the file search prefix is part of the text. Returns true when file search mode ended (the window updates its cue)
    LeaveRestored() {
        if !this.RestoredSelection
            return false
        this.RestoredSelection := false
        if !this.FileMode
            return false
        this.FileMode := false
        return true
    }

    ;---------------------------------------------------------------------------
    ; Searches
    ;---------------------------------------------------------------------------
    ; A new search for text: numbered after the ones before it, which no longer count. In search results the highlight follows
    ; the text (browsing a folder: only the filter part, not the words in the path); the actions panel keeps it.
    ; quick: only the features that don't make you wait (no file search; see SearchRequest)
    StartSearch(text, quick := false) {
        if (this.Mode = "results") {
            browse := FileSearchProvider.BrowsePath(text)
            this.HighlightText := IsObject(browse) ? browse.Filter : text
        }
        return this.Request := SearchRequest(++this.SearchGen, text, this.FileMode, quick)
    }

    ; Whether request is still the newest search (its results are shown)
    IsCurrent(request) => IsObject(request) && request.Gen = this.SearchGen

    ;---------------------------------------------------------------------------
    ; Results and selection
    ;---------------------------------------------------------------------------
    SetResults(results) {
        this.Results := results, this.Selected := results.Length ? 1 : 0, this.Offset := 0
    }

    SelectedItem() {
        return (this.Selected >= 1 && this.Selected <= this.Results.Length) ? this.Results[this.Selected] : ""
    }

    ; Moves the selection by step, scrolling so it stays among the visibleRows shown; wrap: past the last result back to the
    ; first and the other way round (Tab cycling). Returns false when there are no results
    MoveSelection(step, visibleRows, wrap := false) {
        count := this.Results.Length
        if !count
            return false
        if (wrap && (this.Selected + step > count || this.Selected + step < 1))
            this.Selected := (step > 0) ? 1 : count
        else
            this.Selected := Max(1, Min(count, this.Selected + step))
        if (this.Selected <= this.Offset)
            this.Offset := this.Selected - 1
        else if (this.Selected > this.Offset + visibleRows)
            this.Offset := this.Selected - visibleRows
        return true
    }

    ;---------------------------------------------------------------------------
    ; Actions panel
    ;---------------------------------------------------------------------------
    ; Opens the panel for source with these actions; query = the search box text, put back when it closes
    OpenActions(source, actions, query) {
        this.SavedQuery := query, this.ActionSource := source, this.AllActions := actions, this.Mode := "actions"
    }

    ; Back to the results; returns the search text to put back
    CloseActions() {
        this.Mode := "results"
        return this.SavedQuery
    }

    ; The panel's actions matching what is typed (all of them for an empty box)
    FilterActions(text) {
        if (Trim(text) = "")
            return this.AllActions
        filtered := []
        for action in this.AllActions
            if (FuzzyMatcher.Score(text, action.Title) > 0)
                filtered.Push(action)
        return filtered
    }

    ;---------------------------------------------------------------------------
    ; Marked files and folders (Insert)
    ;---------------------------------------------------------------------------
    static CanMark(item) => IsObject(item) && (item.Kind = "file" || item.Kind = "folder") && item.Arg != ""

    ; Marks / unmarks item; false when it can't be marked (not a file or folder)
    ToggleMark(item) {
        if !SearchSession.CanMark(item)
            return false
        key := StrLower(item.Arg)
        if this.Marked.Has(key)
            this.Marked.Delete(key)
        else
            this.Marked[key] := item.Arg
        return true
    }

    IsMarked(item) => this.Marked.Count && SearchSession.CanMark(item) && this.Marked.Has(StrLower(item.Arg))

    MarkedPaths() {
        paths := []
        for key, markedPath in this.Marked
            paths.Push(markedPath)
        return paths
    }

    ;---------------------------------------------------------------------------
    ; History browsing (Ctrl+↑ older / Ctrl+↓ newer)
    ;---------------------------------------------------------------------------
    ; The entry step away from the one shown, skipping entries equal to current (so each press changes something); going past
    ; the newest entry brings back the text typed before browsing. Returns {Text, Index}, or "" at the oldest / newest end.
    ; The caller puts the text in the box, then sets HistoryIndex := Index (editing the box resets it)
    HistoryStep(step, history, current) {
        index := this.HistoryIndex
        if (index = 0)
            this.HistoryDraft := current
        Loop {
            index += step
            if (index < 0 || index > history.Length)
                return ""
            text := index ? history[index] : this.HistoryDraft
            if (text != current || index = 0)
                return {Text: text, Index: index}
        }
    }
}

;===============================================================================
; SearchRequest - one search: its number, text and kind (made by SearchSession.StartSearch)
;===============================================================================
class SearchRequest {
    __New(gen, text, fileMode, quick := false) {
        this.Gen := gen, this.Text := text, this.FileMode := fileMode, this.Quick := quick
    }

    ; The results: in file search mode files only; quick = without file search (Everything may wait up to its timeout), nothing
    ; in file search mode
    Run() {
        if this.Quick
            return this.FileMode ? [] : ProviderRegistry.Search(this.Text, FileSearchProvider)
        return this.FileMode ? ProviderRegistry.SearchOnly(FileSearchProvider, this.Text) : ProviderRegistry.Search(this.Text)
    }
}
