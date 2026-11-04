;===============================================================================
; ErrorHandler.ahk - errors nothing caught, once ALTRun has started (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Without it, an error nobody caught in a hotkey, timer, window event or message handler shows AutoHotkey's raw English dialog
; and isn't logged. Once start-up is done (App.Start installs it), such an error is written to the log with its place and call
; stack, even when the debug log is off, and a short notice says what went wrong and where the details are (at most one notice
; every 10 seconds). The thread that failed stops; ALTRun keeps running. Errors that end ALTRun (mode "ExitApp") and errors during
; start-up still show the dialog: there is nothing to keep running then.
;
; Usage:
;   ErrorHandler.Install()                       once, after start-up
;   ErrorHandler.Handle(thrown, mode, notify)    what OnError calls; notify := false in the tests
;===============================================================================

class ErrorHandler {
    static Count := 0                                                       ; errors handled since start-up
    static NoticeGapMs := 10000
    static _lastNotice := 0
    static _handlerFn := ""

    static Install() {
        if (ErrorHandler._handlerFn = "") {
            ErrorHandler._handlerFn := (thrown, mode) => ErrorHandler.Handle(thrown, mode)
            OnError(ErrorHandler._handlerFn)
        }
    }

    ; 1 = handled (no dialog, the thread stops), 0 = AutoHotkey shows its dialog
    static Handle(thrown, mode, notify := true) {
        ErrorHandler.Count++
        try Logger.Record("ERR", "Uncaught error (" mode "): " ErrorHandler.Describe(thrown))
        if (mode = "ExitApp")
            return 0
        if (notify && (!ErrorHandler._lastNotice || A_TickCount - ErrorHandler._lastNotice >= ErrorHandler.NoticeGapMs)) {
            ErrorHandler._lastNotice := A_TickCount
            try App.Toast(I18n.T("App.UncaughtError", ErrorHandler.Summary(thrown), Logger.File), 8000)
        }
        return 1
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
