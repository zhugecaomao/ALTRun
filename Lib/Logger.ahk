;===============================================================================
; Logger.ahk - 轻量日志 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 日志文件默认写到 %Temp%\ALTRun.log。
; 缓冲后批量落盘, 避免每条日志都开关一次文件 (启动时会写几百条)。
;
; 用法:
;   Logger.Debug("xxx")  /  Logger.Warn("xxx")  /  Logger.Error("xxx")
;   Logger.Flush()   ; 立即把缓冲区写盘, 一般不用手动调, 退出时会自动调用一次
;   Logger.Rotate()  ; 日志文件过大时截断保留一份 .old, 启动时自动调用一次
;
; 是否写盘由 Logger.Enabled 决定, App 启动时按设置 General.SaveLog 赋值。
;===============================================================================

class Logger {
    static File    := A_Temp "\ALTRun.log"
    static _buffer := []
    static _maxBuf := 60                        ; 攒够这么多条才写一次盘
    static Enabled := true

    static Debug(msg) => Logger._Write("DBG", msg)
    static Warn(msg)  => Logger._Write("WRN", msg)
    static Error(msg) => Logger._Write("ERR", msg)

    static _Write(level, msg) {
        if !Logger.Enabled
            return
        Logger._buffer.Push(FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") " [" level "] " msg "`r`n")
        if (Logger._buffer.Length >= Logger._maxBuf)
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
