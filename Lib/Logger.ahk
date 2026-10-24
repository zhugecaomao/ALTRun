;===============================================================================
; Logger.ahk - a light logger (AutoHotkey v2)
;-------------------------------------------------------------------------------
; The log file is %Temp%\ALTRun.log by default.
; During start-up lines are buffered and written in batches (dozens at once); after start-up (Logger.Immediate) each line is written at once:
; if the program hangs and is killed, the log still shows how far it got. Warnings and errors are always written at once.
; One line per entry (CRLF); line breaks in a message become " | ".
;
; Usage:
;   Logger.Debug("xxx")  /  Logger.Warn("xxx")  /  Logger.Error("xxx")
;   Logger.Flush()   ; writes the buffer now; rarely needed by hand, it is called once on exit
;   Logger.Close()   ; closes the log file (normally kept open, reopened on the next write); called once on exit
;   Logger.Rotate()  ; when the log is too big, keeps it as .old and starts a new one; called once at start-up
;   start := Logger.Ms(), ..., Logger.Time("load index", start)   ; logs the time "Perf: load index 12 ms"
;   SetTimer(Logger.Watch("save usage", fn), ...)   ; logs "Perf: slow save usage 80 ms" when one run of a background task exceeds SlowMs
;
; Logger.Enabled decides whether anything is written; App sets it from General.SaveLog at start-up.
;===============================================================================

class Logger {
    static File    := A_Temp "\ALTRun.log"
    static _buffer := []
    static _file := "", _filePath := ""         ; the log file, kept open (see _Open)
    static _maxBuf := 60                        ; write to disk once this many lines are buffered
    static Enabled := true
    static Immediate := false                   ; true = write each line at once (switched on when App has started)
    static TraceUntil := 0                      ; Trace() only writes while A_TickCount is before this (a short time after start-up)

    static Debug(msg) => Logger._Write("DBG", msg)

    ; High-resolution milliseconds (QueryPerformanceCounter; A_TickCount is only accurate to 10 ~ 16 ms)
    static Ms() {
        static freq := 0
        if !freq
            DllCall("QueryPerformanceFrequency", "Int64*", &freq)
        DllCall("QueryPerformanceCounter", "Int64*", &now := 0)
        return now * 1000 / freq
    }

    ; Logs the time from start (Logger.Ms()) until now; returns the milliseconds
    static Time(label, start) {
        elapsed := Logger.Ms() - start
        if Logger.Enabled
            Logger._Write("DBG", "Perf: " label " " Round(elapsed) " ms")
        return elapsed
    }
    ; Wraps a function with timing (for timers): a run over SlowMs logs "Perf: slow <label> N ms", to find background tasks that stall now and then.
    ; Time spent in an interruption (a hotkey etc.) counts too
    static SlowMs := 50
    static Watch(label, fn) => (params*) => Logger._Watched(label, fn, params)

    static _Watched(label, fn, params) {
        start := Logger.Ms()
        try return fn(params*)
        finally {
            elapsed := Logger.Ms() - start
            if (Logger.Enabled && elapsed >= Logger.SlowMs)
                Logger._Write("DBG", "Perf: slow " label " " Round(elapsed) " ms")
        }
    }

    ; Finer steps (each step of showing the window, icons loaded...), only logged for a short time after start-up so they don't fill the log
    static Trace(msg) {
        if (Logger.Enabled && A_TickCount < Logger.TraceUntil)
            Logger._Write("DBG", msg)
    }
    static Warn(msg)  => Logger._Write("WRN", msg)
    static Error(msg) => Logger._Write("ERR", msg)

    static _Write(level, msg) {
        if !Logger.Enabled
            return
        msg := RegExReplace(msg, "\R+", " | ")                               ; one entry per line, always ending in CRLF
        Logger._buffer.Push(FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") " [" level "] " msg "`r`n")
        if (Logger.Immediate || level != "DBG" || Logger._buffer.Length >= Logger._maxBuf)
            Logger.Flush()
    }

    ; Appends the buffer to the file in one go. Called once before exit (see OnExit in App.ahk),
    ; but if the program is killed or crashes, the lines not yet flushed are lost.
    static Flush() {
        if (Logger._buffer.Length = 0)
            return
        text := ""
        for line in Logger._buffer
            text .= line
        Logger._buffer := []
        try {
            out := Logger._Open()
            out.Write(text)
            handle := out.Handle                                               ; reading Handle hands the write buffer to the system: nothing is lost if the program is killed
            return
        }
        Logger.Close()
        try FileAppend(text, Logger.File, "UTF-8")
    }

    ; The log file stays open: opening and closing it for each line is slow (on old PCs antivirus checks every close, 20 ~ 30 ms per line).
    ; Shared for read / write / delete, so other programs can open it while it is written; reopened when Logger.File changes or the file is deleted.
    static _Open() {
        if (Logger._file != "" && (Logger._filePath != Logger.File || !FileExist(Logger.File)))
            Logger.Close()
        if (Logger._file = "") {
            if !(out := FileOpen(Logger.File, "a", "UTF-8"))
                throw OSError()
            Logger._file := out, Logger._filePath := Logger.File
        }
        return Logger._file
    }

    static Close() {
        if (Logger._file != "")
            try Logger._file.Close()
        Logger._file := "", Logger._filePath := ""
    }

    ; Cuts the log when it gets too big, so it can't grow forever. App calls it once at start-up.
    static Rotate(maxBytes := 1048576) {
        Logger.Close()
        try {
            if (FileExist(Logger.File) && FileGetSize(Logger.File) > maxBytes)
                FileMove(Logger.File, Logger.File ".old", true)
        }
    }
}
