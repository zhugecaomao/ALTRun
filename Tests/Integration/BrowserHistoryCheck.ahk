;===============================================================================
; BrowserHistoryCheck.ahk - reading the History of a running browser (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Run by .github/workflows/browser-history.yml after it opened a few pages in Edge and Chrome and left them running
; (the file named on the command line lists them, one "url|title" per line in the order they were opened). Checks the part that depends on Windows and the browsers:
;   - winsqlite3.dll is there
;   - BrowserHistoryProvider.ReadFile reads each browser's History while the browser keeps it open (in place, without
;     locking it): every page is found, with its title (a Chinese one too), most recent first
;   - the browser still runs afterwards, and no copy is left in %TEMP%
; The rest (several profiles, filters, matching) is in the unit tests (ProviderTests.BrowserHistory) with test databases.
; Prints PASS / FAIL lines; the exit code is the number of failures.
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Off
#Warn All, StdOut
#Warn VarUnset, Off                                                         ; the provider names ALTRun classes this check doesn't load (not called here)
#Include %A_ScriptDir%\..\..\Lib\Sqlite.ahk
#Include %A_ScriptDir%\..\..\Src\Providers\BrowserHistoryProvider.ahk

class Logger {                                                              ; ReadFile writes a debug line when it falls back to a copy
    static Lines := []
    static Debug(text) => Logger.Lines.Push(text)
}

Main()

Main() {
    failures := 0
    out(text) => FileAppend(text "`n", "*", "UTF-8")
    check(name, ok, detail := "") {
        out((ok ? "PASS " : "FAIL ") name (detail != "" ? " (" detail ")" : ""))
        failures += ok ? 0 : 1
    }
    out("Windows " A_OSVersion)
    check("winsqlite3.dll", Sqlite.Available())
    if !Sqlite.Available()
        ExitApp(failures)
    expected := []                                                          ; [url, title] in the order they were opened (the last is the most recent)
    for line in StrSplit(FileRead(A_Args[1], "UTF-8"), "`n", "`r")
        if (line != "")
            expected.Push(StrSplit(line, "|", , 2))
    check("pages to look for", expected.Length >= 2, expected.Length)
    for browser in [["Edge", "Microsoft\Edge\User Data\Default\History", "msedge.exe"], ["Chrome", "Google\Chrome\User Data\Default\History", "chrome.exe"]] {
        historyPath := EnvGet("LOCALAPPDATA") "\" browser[2]
        if !FileExist(historyPath) {
            check(browser[1] ": History exists", false, historyPath)
            continue
        }
        check(browser[1] ": running while read", ProcessExist(browser[3]) > 0)
        start := A_TickCount
        try {
            rows := BrowserHistoryProvider.ReadFile(historyPath)
        } catch as e {
            check(browser[1] ": read", false, e.Message)
            continue
        }
        elapsed := A_TickCount - start
        check(browser[1] ": read in place", !Logger.Lines.Length, Logger.Lines.Length ? Logger.Lines[1] : elapsed " ms, " rows.Length " pages")
        Logger.Lines := []
        found := Map()
        for index, row in rows
            found[row[1]] := {Index: index, Title: row[2], Time: row[3]}
        previous := 0
        for page in expected {
            ok := found.Has(page[1]) && found[page[1]].Title = page[2]
            check(browser[1] ": " page[1], ok, ok ? "" : (found.Has(page[1]) ? "title " found[page[1]].Title : "not found"))
            if (ok && previous)
                check(browser[1] ": " page[1] " before the page opened earlier", found[page[1]].Index < previous)
            if ok
                previous := found[page[1]].Index
        }
        check(browser[1] ": still running", ProcessExist(browser[3]) > 0)
    }
    check("no copy in TEMP", !DirExist(A_Temp "\ALTRun-History-" DllCall("GetCurrentProcessId")))
    ExitApp(failures)
}
