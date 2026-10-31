;===============================================================================
; EverythingCheck.ahk - ALTRun's Everything IPC (Lib\Everything.ahk) against a real, running Everything (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Run by .github/workflows/everything.yml on Windows with Everything 1.4 and 1.5; also works on your own PC while
; Everything runs. Prints one line per check (PASS / FAIL) and the IPC windows found; the exit code is the number of
; failed checks.
;
;   EverythingCheck.ahk <folder> <name>
;     <folder>\<name>.txt (a file) and <folder>\<name>_dir (a folder) must exist; <name> should be unique on the PC,
;     so that searching it finds those two and nothing else
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut                                                           ; never a warning dialog: on a CI runner nobody closes it
#Include ..\..\Lib\Everything.ahk

EverythingCheck.Run(A_Args[1], A_Args[2])

class EverythingCheck {
    static Failures := 0

    static Run(folder, name) {
        folder := RTrim(folder, "\")
        started := A_TickCount
        while (!Everything.IsRunning() && A_TickCount - started < 60000)
            Sleep(250)
        EverythingCheck.Check("Everything is running", Everything.IsRunning(), "after " (A_TickCount - started) " ms")
        EverythingCheck.Log("IPC windows: " EverythingCheck.IpcWindows())
        if !Everything.IsRunning()
            ExitApp(EverythingCheck.Failures)

        ; Right after starting, Everything is still building its index: ask until both test items are found
        started := A_TickCount, results := []
        while (A_TickCount - started < 120000) {
            results := Everything.Query(name, 10, Everything.SORT_NAME, 2000)
            if (results.Length >= 2)
                break
            Sleep(500)
        }
        EverythingCheck.Check("query finds the test file and folder", results.Length = 2
            , results.Length " results after " (A_TickCount - started) " ms" (Everything.LastProblem != "" ? ", " Everything.LastProblem : ""))
        found := Map()
        for item in results
            found[StrLower(item.Path)] := item.IsFolder
        filePath := folder "\" name ".txt", dirPath := folder "\" name "_dir"
        EverythingCheck.Check("file: full path, not a folder", found.Has(StrLower(filePath)) && !found[StrLower(filePath)], filePath)
        EverythingCheck.Check("folder: full path, a folder", found.Has(StrLower(dirPath)) && found[StrLower(dirPath)], dirPath)

        EverythingCheck.Check("no results for a name nobody has", Everything.Query(name "_nothing_like_it", 10).Length = 0)

        ; How long a query takes once the index is ready (information only: a CI runner is no PC). Timed with the
        ; performance counter: A_TickCount only moves in 15.6 ms steps
        times := []
        Loop 20 {
            t := EverythingCheck.NowMs()
            Everything.Query(name, 10)
            times.Push(Round(EverythingCheck.NowMs() - t, 1))
        }
        EverythingCheck.Log("query time over 20 queries: median " EverythingCheck.Median(times) " ms, max " Max(times*) " ms")
        ExitApp(EverythingCheck.Failures)
    }

    ; Classes of the top-level windows whose class starts with EVERYTHING (the IPC window is hidden)
    static IpcWindows() {
        DetectHiddenWindows(true)
        classes := ""
        for hwnd in WinGetList() {
            try cls := WinGetClass(hwnd)
            catch
                continue
            if (InStr(cls, "EVERYTHING") = 1 && !InStr(classes, cls ";"))
                classes .= cls ";"
        }
        return classes = "" ? "(none)" : RTrim(classes, ";")
    }

    static Check(label, ok, detail := "") {
        EverythingCheck.Log((ok ? "PASS " : "FAIL ") label (detail != "" ? " (" detail ")" : ""))
        if !ok
            EverythingCheck.Failures++
    }

    static Log(text) => FileAppend(text "`n", "*", "UTF-8")

    static NowMs() {
        static freq := 0
        if !freq
            DllCall("QueryPerformanceFrequency", "Int64*", &freq)
        DllCall("QueryPerformanceCounter", "Int64*", &now := 0)
        return now * 1000 / freq
    }

    static Median(values) {
        sorted := []
        for v in values
            sorted.Push(v)
        Loop sorted.Length - 1 {                                            ; insertion sort: 20 values
            i := A_Index + 1, v := sorted[i]
            while (i > 1 && sorted[i - 1] > v)
                sorted[i] := sorted[i - 1], i--
            sorted[i] := v
        }
        return sorted[(sorted.Length + 1) // 2]
    }
}
