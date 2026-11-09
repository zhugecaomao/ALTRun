;===============================================================================
; ErrorHandler.ahk - errors nothing caught, once ALTRun has started (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Without it, an error nobody caught in a hotkey, timer, window event or message handler shows AutoHotkey's raw English dialog
; and isn't logged. Once start-up is done (App.Start installs it), such an error is written to the log with its place and call
; stack, even when the debug log is off, and a short notice says what went wrong and where the details are (at most one notice
; every 10 seconds). The thread that failed stops; ALTRun keeps running. Errors that end ALTRun (mode "ExitApp") and errors during
; start-up still show the dialog: there is nothing to keep running then.
; The same error can come again and again (a timer that fails each second), and ALTRun runs for weeks: from one place (file and
; line) the first FullTimes are written in full, later ones only counted, and once an hour (and on exit) one line says how many
; more there were. The log is only cut at start-up (Logger.Rotate), so it would otherwise grow by megabytes an hour.
;
; Usage:
;   ErrorHandler.Install()                       once, after start-up
;   ErrorHandler.Handle(thrown, mode, notify)    what OnError calls; notify := false in the tests
;   ErrorHandler.WriteCounts()                   the counted errors not written yet (on exit; otherwise once an hour)
;===============================================================================

class ErrorHandler {
    static Count := 0                                                       ; errors handled since start-up
    static NoticeGapMs := 10000
    static FullTimes := 5, CountsEveryMs := 3600000
    static _lastNotice := 0
    static _handlerFn := "", _countsTimer := ""
    static _places := Map()                                                 ; "file:line" -> {Times, Unwritten}

    static Install() {
        if (ErrorHandler._handlerFn = "") {
            ErrorHandler._handlerFn := (thrown, mode) => ErrorHandler.Handle(thrown, mode)
            OnError(ErrorHandler._handlerFn)
        }
    }

    ; 1 = handled (no dialog, the thread stops), 0 = AutoHotkey shows its dialog
    static Handle(thrown, mode, notify := true) {
        ErrorHandler.Count++
        place := ErrorHandler.Place(thrown)
        if !ErrorHandler._places.Has(place)
            ErrorHandler._places[place] := {Times: 0, Unwritten: 0}
        seen := ErrorHandler._places[place], seen.Times++
        if (seen.Times <= ErrorHandler.FullTimes || mode = "ExitApp") {
            try Logger.Record("ERR", "Uncaught error (" mode "): " ErrorHandler.Describe(thrown)
                . (seen.Times = ErrorHandler.FullTimes ? " (the same error here again is only counted; the count is written once an hour)" : ""))
        } else {
            seen.Unwritten++
            if (ErrorHandler._countsTimer = "")
                ErrorHandler._countsTimer := () => ErrorHandler.WriteCounts()
            SetTimer(ErrorHandler._countsTimer, ErrorHandler.CountsEveryMs)
        }
        if (mode = "ExitApp")
            return 0
        if (notify && (!ErrorHandler._lastNotice || A_TickCount - ErrorHandler._lastNotice >= ErrorHandler.NoticeGapMs)) {
            ErrorHandler._lastNotice := A_TickCount
            try App.Toast(I18n.T("App.UncaughtError", ErrorHandler.Summary(thrown), Logger.File), 8000)
        }
        return 1
    }

    ; Where an error comes from, to count the same one: file and line ("" for a thrown value that isn't an Error)
    static Place(thrown) {
        if !(thrown is Error)
            return ""
        SplitPath(thrown.File, &fileName)
        return fileName ":" thrown.Line
    }

    ; One line per place with errors counted but not written yet; the timer stops until the next one
    static WriteCounts() {
        if (ErrorHandler._countsTimer != "")
            SetTimer(ErrorHandler._countsTimer, 0)
        for place, seen in ErrorHandler._places
            if seen.Unwritten {
                try Logger.Record("ERR", "Uncaught error at " (place = "" ? "(not an Error)" : place) ": " seen.Unwritten " more since the last note ("
                    . seen.Times " since start-up)")
                seen.Unwritten := 0
            }
    }

    ; One line for the log: type, message, extra, file and line, then the call stack
    static Describe(thrown) {
        if !(thrown is Error) {
            try return "thrown value: " String(thrown)
            return "thrown value of type " Type(thrown)
        }
        text := Type(thrown) ": " thrown.Message (thrown.Extra != "" ? " (" thrown.Extra ")" : "")
        SplitPath(thrown.File, &fileName)
        text .= " at " fileName ":" thrown.Line
        stack := Trim(thrown.Stack, " `t`r`n")
        return stack = "" ? text : text "`nstack: " stack                  ; the log turns the line breaks into " | "
    }

    ; For the notice: the message alone
    static Summary(thrown) {
        if (thrown is Error)
            return thrown.Message
        try return String(thrown)
        return Type(thrown)
    }
}
