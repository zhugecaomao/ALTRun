;===============================================================================
; WinTimer.ahk - timers that wait without waking AutoHotkey every 10 ms (AutoHotkey v2)
;-------------------------------------------------------------------------------
; While any SetTimer timer is set, AutoHotkey keeps a 10 ms Windows timer running to check its timers, even when the only one is due
; in an hour. AutoHotkey v2.0 source:
;   globaldata.h  SLEEP_INTERVAL 10; SET_MAIN_TIMER = SetTimer(g_hWnd, TIMER_ID_MAIN, SLEEP_INTERVAL, NULL)
;                 https://github.com/AutoHotkey/AutoHotkey/blob/v2.0/source/globaldata.h#L218-L239
;   script.cpp    enabling a script timer: "SET_MAIN_TIMER // Ensure the API timer is always running when there is at least one enabled
;                 timed subroutine"; KILL_MAIN_TIMER only when none is left
;                 https://github.com/AutoHotkey/AutoHotkey/blob/v2.0/source/script.cpp#L3941-L3991
; Measured on Windows Server 2025 (GitHub Actions), CPU cycles of a small script in an idle minute: no timer 0.8, one 2 s SetTimer 330,
; one one-hour one-shot SetTimer 328, one 2 s Windows timer 4.5 million. ALTRun always has such timers (folder watches, hourly checks),
; so it used 240 to 420 million cycles a minute while idle (Tests\Bench), the constant 0.1% in Task Manager (#290).
; WinTimer uses Windows timers instead (SetTimer on a message-only window, WM_TIMER), so nothing runs until a timer is due.
; Meant for timers that stay set while ALTRun is idle; short-lived ones (batches, delays while typing) keep SetTimer.
; A single SetTimer timer left set while idle keeps the 10 ms timer running for all (the bench showed no gain while Quick Switch still
; polled with SetTimer), so every one of them must be a WinTimer: ProjectTests.LongTimers flags long SetTimer periods.
; Result (Tests\Bench, median of 6 idle minutes): 301 million cycles a minute before, 18 million with these timers on WinTimer and
; Quick Switch on window events (#291, #295).
; Like SetTimer, the function runs in a new thread, and not again while it is still running (a later tick is skipped, not queued);
; WM_TIMER is a low-priority message, so a busy moment can delay it a little.
; The timers are coalescable (SetCoalescableTimer, a tolerance of a tenth of the time, at most 1 s): Windows may line them up with other
; timers to wake the CPU less often, as recommended for background work. None of these needs to run to the millisecond.
;
; Usage:
;   WinTimer.Set(fn, 2000)       every 2 s; setting the same function again restarts it with the new time
;   WinTimer.Set(fn, -60000)     once, in a minute
;   WinTimer.Set(fn, 0)          stop
;   WinTimer.Count               how many are set
;===============================================================================

class WinTimer {
    static _byFn := Map(), _byId := Map(), _nextId := 0, _gui := ""

    static MaxToleranceMs := 1000

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
        running := WinTimer._byId.Has(id) && WinTimer._byId[id].Running     ; set again from its own function: it is still running
        WinTimer._byFn[fn] := id, WinTimer._byId[id] := {Fn: fn, Once: ms < 0, Running: running}
        if !WinTimer._SetTimer(hwnd, id, Abs(ms))                           ; the same id again: restarts it
            throw OSError(A_LastError, "WinTimer.Set")
    }

    ; Windows 8+: SetCoalescableTimer; SetTimer where it is missing
    static _SetTimer(hwnd, id, ms) {
        try return DllCall("SetCoalescableTimer", "Ptr", hwnd, "UPtr", id, "UInt", ms, "Ptr", 0, "UInt", WinTimer.Tolerance(ms), "UPtr")
        return DllCall("SetTimer", "Ptr", hwnd, "UPtr", id, "UInt", ms, "Ptr", 0, "UPtr")
    }

    static Tolerance(ms) => Min(Abs(ms) // 10, WinTimer.MaxToleranceMs)

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
        if timer.Running                                                    ; its function hasn't finished yet: skip this tick, like SetTimer
            return 0
        if timer.Once
            WinTimer._Delete(id)
        timer.Running := true
        try timer.Fn.Call()
        finally {
            if (WinTimer._byId.Has(id) && WinTimer._byId[id].Fn = timer.Fn)   ; may have been stopped or set again meanwhile
                WinTimer._byId[id].Running := false
        }
        return 0
    }
}
