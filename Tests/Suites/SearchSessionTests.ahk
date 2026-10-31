;===============================================================================
; SearchSessionTests.ahk - the search window's state without a window (SearchSession)
;-------------------------------------------------------------------------------
; Included by Tests\RunTests.ahk. Every static method whose name doesn't start with "_" is a test;
; TestRunner runs them in name order, and each one must put back the shared state it changes (these use their own sessions).
;===============================================================================
class SearchSessionTests {
    static _Items(count) {
        items := []
        Loop count
            items.Push(ResultItem("Item " A_Index, "", {Kind: "file", Arg: "C:\Work\item" A_Index ".txt"}))
        return items
    }

    ; A new showing resets the mode, marks and history browsing; a restored search brings back its file search mode, selected
    static Showing() {
        eq := (n, a, e) => TestRunner.Equal("SessionShowing." n, a, e)
        s := SearchSession()
        s.Mode := "actions", s.ActionsOnly := true, s.Marked["c:\x"] := "C:\x", s.HistoryIndex := 3, s.FileMode := true
        s.Last := {Text: "report", FileMode: true, Selected: 2, Time: 0}
        s.StartShowing(false)
        eq("plain", s.Mode "|" s.ActionsOnly "|" s.Marked.Count "|" s.HistoryIndex "|" s.FileMode "|" s.RestoredSelection, "results|0|0|0|0|0")
        s.StartShowing(true)
        eq("restored", s.FileMode "|" s.RestoredSelection, "1|1")
        eq("typed over: file search ends", s.LeaveRestored() "|" s.FileMode "|" s.RestoredSelection, "1|0|0")
        eq("typed over again: nothing to do", s.LeaveRestored(), false)
        s.Last.FileMode := false, s.StartShowing(true)
        eq("restored normal search typed over", s.LeaveRestored() "|" s.RestoredSelection, "0|0")
    }

    ; At hiding: the text, file search mode and selection; from the actions panel the search it was opened from
    static Remember() {
        eq := (n, a, e) => TestRunner.Equal("SessionRemember." n, a, e)
        s := SearchSession()
        s.SetResults(SearchSessionTests._Items(5)), s.MoveSelection(2, 8), s.FileMode := true
        s.Remember("report", 1234)
        eq("results", s.Last.Text "|" s.Last.FileMode "|" s.Last.Selected "|" s.Last.Time, "report|1|3|1234")
        s.OpenActions(s.SelectedItem(), [ResultItem("Open")], "budget")
        s.Remember("", 99)
        eq("actions panel", s.Last.Text "|" s.Last.Selected, "budget|1")
        s.SetResults([]), s.Mode := "results", s.Remember("x")
        eq("no results: row 1", s.Last.Selected, 1)
    }

    ; Selection: clamped, scrolled into the visible rows, wrapping only when asked (Tab cycling)
    static Selection() {
        eq := (n, a, e) => TestRunner.Equal("SessionSelection." n, a, e)
        s := SearchSession()
        eq("no results", s.MoveSelection(1, 3) "|" s.Selected "|" s.SelectedItem(), "0|0|")
        s.SetResults(SearchSessionTests._Items(10))
        eq("first selected", s.Selected "|" s.Offset "|" s.SelectedItem().Title, "1|0|Item 1")
        s.MoveSelection(4, 3)
        eq("down: scrolled", s.Selected "|" s.Offset, "5|2")
        s.MoveSelection(-3, 3)
        eq("up: scrolled back", s.Selected "|" s.Offset, "2|1")
        s.MoveSelection(-5, 3)
        eq("clamped at the top", s.Selected "|" s.Offset, "1|0")
        s.MoveSelection(50, 3)
        eq("clamped at the bottom", s.Selected "|" s.Offset, "10|7")
        s.MoveSelection(1, 3, true)
        eq("wraps to the first", s.Selected "|" s.Offset, "1|0")
        s.MoveSelection(-1, 3, true)
        eq("wraps to the last", s.Selected "|" s.Offset, "10|7")
        s.SetResults(SearchSessionTests._Items(2))
        eq("new results start at the top", s.Selected "|" s.Offset, "1|0")
    }

    ; Actions panel: opened for a result with its actions, filtered by typing, closed back to the search text
    static Actions() {
        eq := (n, a, e) => TestRunner.Equal("SessionActions." n, a, e)
        s := SearchSession()
        source := ResultItem("Budget.xlsx")
        s.OpenActions(source, [ResultItem("Open"), ResultItem("Copy Path"), ResultItem("Properties")], "budget")
        eq("opened", s.Mode "|" s.ActionSource.Title "|" s.SavedQuery, "actions|Budget.xlsx|budget")
        eq("empty filter: all", s.FilterActions(" ").Length, 3)
        titles := ""
        for action in s.FilterActions("prop")
            titles .= action.Title "|"
        eq("filtered", titles, "Properties|")
        eq("closed", s.CloseActions() "|" s.Mode, "budget|results")
    }

    ; Marks: files and folders only, by path in any case; listed by path (a Map keeps its keys sorted)
    static Marks() {
        eq := (n, a, e) => TestRunner.Equal("SessionMarks." n, a, e)
        s := SearchSession()
        items := SearchSessionTests._Items(3)
        folder := ResultItem("Work", "", {Kind: "folder", Arg: "C:\Work"})
        eq("url can't be marked", s.ToggleMark(ResultItem("Site", "", {Kind: "url", Arg: "https://example.com"})) "|" s.Marked.Count, "0|0")
        eq("nothing", s.ToggleMark("") "|" s.IsMarked(items[1]), "0|0")
        s.ToggleMark(items[2]), s.ToggleMark(folder)
        eq("marked", s.Marked.Count "|" s.IsMarked(items[2]) "|" s.IsMarked(items[1]), "2|1|0")
        same := ResultItem("Item 2", "", {Kind: "file", Arg: "c:\WORK\ITEM2.TXT"})
        eq("same path in another case", s.IsMarked(same), true)
        eq("paths", TextTools.JoinList(s.MarkedPaths(), ";"), "C:\Work;C:\Work\item2.txt")
        s.ToggleMark(same)
        eq("unmarked", s.Marked.Count "|" s.IsMarked(items[2]), "1|0")
    }

    ; History browsing: older / newer entries, skipping the one already shown, back to the draft past the newest
    static History() {
        eq := (n, a, e) => TestRunner.Equal("SessionHistory." n, a, e)
        s := SearchSession()
        history := ["report", "budget", "notes"]                            ; most recent first
        step(direction, current) {
            entry := s.HistoryStep(direction, history, current)
            if !IsObject(entry)
                return "end"
            s.HistoryIndex := entry.Index
            return entry.Text "@" entry.Index
        }
        eq("older from a draft", step(1, "dra"), "report@1")
        eq("older", step(1, "report"), "budget@2")
        eq("oldest", step(1, "budget"), "notes@3")
        eq("past the oldest", step(1, "notes"), "end")
        eq("newer", step(-1, "notes"), "budget@2")
        eq("back to the draft", step(-1, "budget") "|" step(-1, "report"), "report@1|dra@0")
        eq("past the draft", step(-1, "dra"), "end")
        s.HistoryIndex := 0
        eq("the same text is skipped", step(1, "report"), "budget@2")
    }
}
