;===============================================================================
; Logger.ahk - 轻量日志 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 日志文件默认写到 %Temp%\ALTRun.log。
; 启动时缓冲后批量落盘 (启动时一下子写几十条), 启动完成后 (Logger.Immediate) 每条马上写入:
; 程序卡住被强制结束时, 日志里也能看到最后做到哪一步。警告和错误总是马上写入。
; 每条日志一行 (CRLF), 消息里的换行换成 " | "。
;
; 用法:
;   Logger.Debug("xxx")  /  Logger.Warn("xxx")  /  Logger.Error("xxx")
;   Logger.Flush()   ; 立即把缓冲区写盘, 一般不用手动调, 退出时会自动调用一次
;   Logger.Rotate()  ; 日志文件过大时截断保留一份 .old, 启动时自动调用一次
;   start := Logger.Ms(), ..., Logger.Time("load index", start)   ; 记录耗时 "Perf: load index 12 ms"
;
; 是否写盘由 Logger.Enabled 决定, App 启动时按设置 General.SaveLog 赋值。
;===============================================================================

class Logger {
    static File    := A_Temp "\ALTRun.log"
    static _buffer := []
    static _maxBuf := 60                        ; 攒够这么多条才写一次盘
    static Enabled := true
    static Immediate := false                   ; true = 每条马上写盘 (App 启动完成后打开)
    static TraceUntil := 0                      ; A_TickCount 在这之前时 Trace() 才写 (启动后的一小段时间)

    static Debug(msg) => Logger._Write("DBG", msg)

    ; 高精度的毫秒数 (QueryPerformanceCounter; A_TickCount 的精度只有 10 ~ 16 ms)
    static Ms() {
        static freq := 0
        if !freq
            DllCall("QueryPerformanceFrequency", "Int64*", &freq)
        DllCall("QueryPerformanceCounter", "Int64*", &now := 0)
        return now * 1000 / freq
    }

    ; 从 start (Logger.Ms()) 到现在的耗时写进日志, 返回毫秒数
    static Time(label, start) {
        elapsed := Logger.Ms() - start
        if Logger.Enabled
            Logger._Write("DBG", "Perf: " label " " Round(elapsed) " ms")
        return elapsed
    }
    ; 更细的步骤 (呼出窗口的每一步、加载的图标...), 只在启动后的一小段时间里记, 平时不占日志
    static Trace(msg) {
        if (Logger.Enabled && A_TickCount < Logger.TraceUntil)
            Logger._Write("DBG", msg)
    }
    static Warn(msg)  => Logger._Write("WRN", msg)
    static Error(msg) => Logger._Write("ERR", msg)

    static _Write(level, msg) {
        if !Logger.Enabled
            return
        msg := RegExReplace(msg, "\R+", " | ")                               ; 一条一行, 行尾统一 CRLF
        Logger._buffer.Push(FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") " [" level "] " msg "`r`n")
        if (Logger.Immediate || level != "DBG" || Logger._buffer.Length >= Logger._maxBuf)
            Logger.Flush()
    }

    ; 把缓冲区一次性追加到文件。退出前会自动调用一次(见 App.ahk 的 OnExit),
    ; 但如果程序是被强制结束/崩溃的, 还没来得及 Flush 的那几条会丢失。
    static Flush() {
        if (Logger._buffer.Length = 0)
            return
        text := ""
        for line in Logger._buffer
            text .= line
        Logger._buffer := []
        try FileAppend(text, Logger.File, "UTF-8")
    }

    ; 日志过大时截断, 防止无限增长。App 启动时调用一次。
    static Rotate(maxBytes := 1048576) {
        try {
            if (FileExist(Logger.File) && FileGetSize(Logger.File) > maxBytes)
                FileMove(Logger.File, Logger.File ".old", true)
        }
    }
}
