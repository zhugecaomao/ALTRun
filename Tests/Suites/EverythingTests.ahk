;===============================================================================
; EverythingTests.ahk - Lib\Everything.ahk against a fake Everything (Tests\FakeEverything.ahk)
;-------------------------------------------------------------------------------
; Included by Tests\RunTests.ahk. Every static method whose name doesn't start with "_" is a test; TestRunner runs them
; in name order. The real Everything 1.4 and 1.5 are checked on Windows by .github/workflows/everything.yml.
; Skipped when a real Everything runs on this PC (its window would answer instead of the fake).
;===============================================================================
class EverythingTests {
    static _Items := [{Path: "C:\Work\Report 2026.docx", IsFolder: false}, {Path: "D:\项目\报告", IsFolder: true}]

    static _Skip(name) {
        if !Everything.IsRunning()
            return false
        TestRunner.Equal("Everything." name " (skipped: Everything runs on this PC)", true, true)
        return true
    }

    ; The query as Everything gets it, and the reply's paths and folder flags
    static Query() {
        eq := (n, a, e) => TestRunner.Equal("EverythingQuery." n, a, e)
        if EverythingTests._Skip("Query")
            return
        eq("not running", Everything.IsRunning() "|" Everything.Query("report").Length, "0|0")
        fake := FakeEverything(EverythingTests._Items)
        try {
            eq("running", Everything.IsRunning(), true)
            results := Everything.Query("report 2026", 30, Everything.SORT_DATE_MODIFIED_DESC)
            q := fake.Queries[1]
            eq("query", q.Search "|" q.Max "|" q.Sort "|" q.RequestFlags "|" (q.ReplyHwnd = A_ScriptHwnd), "report 2026|30|14|4|1")
            eq("results", results.Length, 2)
            if (results.Length = 2)
                eq("paths and folders", results[1].Path "|" results[1].IsFolder "|" results[2].Path "|" results[2].IsFolder
                    , "C:\Work\Report 2026.docx|0|D:\项目\报告|1")
            eq("no problem", Everything.LastProblem, "")
            fake.Results := []
            eq("nothing found", Everything.Query("nothing").Length "|" Everything.LastProblem, "0|")
        } finally {
            fake.Close()
        }
        eq("closed", Everything.IsRunning(), false)
    }

    ; A reply coming later is waited for; none at all ends at the timeout, with the reason
    static Waiting() {
        eq := (n, a, e) => TestRunner.Equal("EverythingWaiting." n, a, e)
        if EverythingTests._Skip("Waiting")
            return
        fake := FakeEverything(EverythingTests._Items)
        try {
            fake.Delay := 60
            started := A_TickCount
            results := Everything.Query("report", 30, 14, 1000)
            elapsed := A_TickCount - started
            eq("late reply waited for", results.Length, 2)
            eq("returns with the reply, not at the timeout", elapsed < 600, true)
            fake.Delay := -1
            started := A_TickCount
            results := Everything.Query("report", 30, 14, 200)
            elapsed := A_TickCount - started
            eq("no reply", results.Length, 0)
            eq("until the timeout", elapsed >= 180 && elapsed < 1000, true)
            eq("reason", InStr(Everything.LastProblem, "no reply in") = 1, true)
            fake.Delay := 0
            eq("next query fine again", Everything.Query("report").Length "|" Everything.LastProblem, "2|")
        } finally {
            fake.Close()
        }
    }

    ; A reply to an earlier query (a search typed past) isn't taken for the current one
    static StaleReply() {
        eq := (n, a, e) => TestRunner.Equal("EverythingStaleReply." n, a, e)
        if EverythingTests._Skip("StaleReply")
            return
        fake := FakeEverything(EverythingTests._Items)
        try {
            fake.StaleFirst := true
            results := Everything.Query("report")
            eq("the current reply", results.Length = 2 && results[1].Path = "C:\Work\Report 2026.docx", true)
            fake.Delay := -1
            Everything.Query("report", 30, 14, 150)
            eq("only a stale reply: counted in the reason", InStr(Everything.LastProblem, "other WM_COPYDATA while waiting: 1") > 0, true)
        } finally {
            fake.Close()
        }
    }

    ; File search with Everything: the separate folder query only when the first one was cut off at its limit (with fewer
    ; results every match is there, folders included), and not after Everything didn't answer (no second wait)
    static FileSearchQueries() {
        eq := (n, a, e) => TestRunner.Equal("EverythingFileSearch." n, a, e)
        if EverythingTests._Skip("FileSearchQueries")
            return
        options := AppSettings.Feature("FileSearch")
        saved := [options["UseEverything"], FileSearchProvider.EverythingTimeoutMs]
        options["UseEverything"] := 1, FileSearchProvider.EverythingTimeoutMs := 100
        fake := FakeEverything(EverythingTests._Items)
        try {
            found := FileSearchProvider.Query("report", 30, false)
            eq("all found in one query", fake.Queries.Length "|" found.Length, "1|2")
            eq("first query", InStr(fake.Queries[1].Search, "report !C:\Windows\") = 1 && fake.Queries[1].Max = FileSearchProvider.EverythingFetch, true)

            many := []
            Loop 20
                many.Push({Path: "C:\Work\report " A_Index ".txt", IsFolder: false})
            fake.Results := many, fake.Queries := []
            FileSearchProvider.Query("report", 5, true)                     ; default results: 4 x 5 = 20 asked for
            eq("cut off: folders asked separately", fake.Queries.Length, 2)
            if (fake.Queries.Length = 2)
                eq("folder query", InStr(fake.Queries[2].Search, "startwith:folder:report ") = 1, true)
            fake.Queries := []
            FileSearchProvider.Query("folder:report", 5, true)
            eq("folder: typed: no second query", fake.Queries.Length, 1)

            fake.Delay := -1, fake.Queries := []
            started := A_TickCount
            found := FileSearchProvider.Query("report", 5, true)
            elapsed := A_TickCount - started
            eq("no answer: no second query", fake.Queries.Length "|" found.Length, "1|0")
            eq("waits EverythingTimeoutMs, not 1 s", elapsed < 600, true)
        } finally {
            fake.Close()
            options["UseEverything"] := saved[1], FileSearchProvider.EverythingTimeoutMs := saved[2]
        }
    }

    ; The 1.5 alphas run as the named instance "1.5a"
    static NamedInstance() {
        eq := (n, a, e) => TestRunner.Equal("EverythingNamedInstance." n, a, e)
        if EverythingTests._Skip("NamedInstance")
            return
        fake := FakeEverything(EverythingTests._Items, "EVERYTHING_TASKBAR_NOTIFICATION_(1.5a)")
        try {
            eq("found", Everything.IsRunning(), true)
            eq("queried", Everything.Query("report").Length, 2)
        } finally {
            fake.Close()
        }
    }
}
