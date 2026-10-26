;===============================================================================
; WinTimer.ahk - timers that wait without waking AutoHotkey every 10 ms (AutoHotkey v2)
;-------------------------------------------------------------------------------
; While any SetTimer timer is set, AutoHotkey checks its timers every 10 ms, even when the only one is due in an hour: about 330 million
; CPU cycles a minute while idle (measured on Windows Server 2025, idle minute: no timer 0.8, one 2 s SetTimer 330, one one-hour one-shot
; SetTimer 328, one Windows timer 4.5 Mcycles). WinTimer uses Windows timers instead (SetTimer on a message-only window, WM_TIMER), so
; nothing runs until a timer is due. Meant for timers that stay set while ALTRun is idle (watching folders every 2 s, hourly checks);
; short-lived ones (batches, delays while typing) keep SetTimer.
; Like SetTimer, the function runs in a new thread; WM_TIMER is a low-priority message, so a busy moment can delay it a little.
;
; Usage:
;   WinTimer.Set(fn, 2000)       every 2 s; setting the same function again restarts it with the new time
;   WinTimer.Set(fn, -60000)     once, in a minute
;   WinTimer.Set(fn, 0)          stop
;   WinTimer.Count               how many are set
;===============================================================================

class WinTimer {
    static _byFn := Map(), _byId := Map(), _nextId := 0, _gui := ""

    static Count => WinTimer._byId.Count

    static Set(fn, ms) {
        if WinTimer._byFn.Has(fn)
            id := WinTimer._byFn[fn]
        else if !ms
            return
        else
            id := ++WinTimer._nextId
        if !ms
            return WinTimer._Delete(id)
        hwnd := WinTimer._Window()
        WinTimer._byFn[fn] := id, WinTimer._byId[id] := {Fn: fn, Once: ms < 0}
        if !DllCall("SetTimer", "Ptr", hwnd, "UPtr", id, "UInt", Abs(ms), "Ptr", 0, "UPtr")   ; the same id again: restarts it
            throw OSError(A_LastError, "WinTimer.Set")
    }

    static _Delete(id) {
        if !WinTimer._byId.Has(id)
            return
        DllCall("KillTimer", "Ptr", WinTimer._gui.Hwnd, "UPtr", id)
        WinTimer._byFn.Delete(WinTimer._byId[id].Fn), WinTimer._byId.Delete(id)
    }

    static _Window() {
        if !IsObject(WinTimer._gui) {
            WinTimer._gui := Gui(, "ALTRun timers")
            DllCall("SetParent", "Ptr", WinTimer._gui.Hwnd, "Ptr", -3)     ; HWND_MESSAGE: only receives messages, never listed among the windows
            OnMessage(0x113, (wParam, lParam, msg, hwnd) => WinTimer._OnTimer(wParam, hwnd), 10)   ; WM_TIMER; up to 10 at once (one timer's function may still be running)
        }
        return WinTimer._gui.Hwnd
    }

    static _OnTimer(id, hwnd) {
        if (hwnd != WinTimer._gui.Hwnd || !WinTimer._byId.Has(id))
            return                                                          ; someone else's timer (AutoHotkey uses WM_TIMER too)
        timer := WinTimer._byId[id]
        if timer.Once
            WinTimer._Delete(id)
        timer.Fn.Call()
        return 0
    }
}
