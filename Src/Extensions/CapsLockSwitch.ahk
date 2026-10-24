;===============================================================================
; CapsLockSwitch.ahk - as on macOS: tap CapsLock to switch the input language, hold it for Caps Lock (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Settings (ALTRun.json -> General.CapsLock):
;   ""       off (default), CapsLock as usual
;   "Layout" switch to the next input language (like Win+Space, e.g. English (US) <-> Microsoft Pinyin)
;   "Mode"   toggle the Chinese / English mode of the current Chinese IME (like Shift in Microsoft Pinyin)
; Caps Lock toggles only after holding for HoldMs (without waiting for the release), with a "Caps Lock: on / off" notice;
; Shift+CapsLock toggles Caps Lock directly as before (with the notice too).
;
; Press and release are two hotkeys; the input language switches only on release (when not held for HoldMs), nothing switches while held.
; No KeyWait for the release: switching input languages (or some IMEs) refreshes the key state, KeyWait thinks it was released,
; and every auto-repeat while held would switch once. Auto-repeat while held (consecutive presses) doesn't count as a new press.
; The hotkeys use the keyboard hook ($): a lone CapsLock would be registered with the system's RegisterHotkey, which can't stop CapsLock's own toggle.
;
; Remote control (Chrome Remote Desktop etc.): pressing CapsLock locally changes the local Caps Lock too, and before the next key the remote software sends
; an extra CapsLock (press and release together, under 1 ms) to align the remote PC's Caps Lock. Press-to-release under MinTapMs
; is taken as such an extra one and caught without doing anything (a human press takes at least tens of milliseconds); otherwise every character typed would switch the input language.
; Timed with the high-resolution QueryPerformanceCounter: A_TickCount is only accurate to 10 ~ 16 ms, too coarse to tell.
; The extra press often comes with other keys held (Alt+Space sends Alt+CapsLock, also Ctrl+C, Shift+letters...), so the hotkeys have *:
; caught whatever modifiers are held, otherwise the one that slips through really toggles Caps Lock (Caps Lock on after showing the window).
; Holding an arrow key etc. sends one extra press per auto-repeat, quickly passing AutoHotkey's limit of "at most 70 hotkeys in 2 s" and showing
; the A_MaxHotkeysPerInterval warning (keys aren't caught during the warning, so Caps Lock toggles again); registering turns that check off
; (A_HotkeyInterval := 0): the check guards against Send triggering its own hotkey in an endless loop, and hotkeys with $ aren't triggered by their own Send.
;
; Usage:
;   when App registers hotkeys: CapsLockSwitch.Mode := "Layout", CapsLockSwitch.Register(Hotkey)
;===============================================================================

class CapsLockSwitch {
    static Mode := "", HoldMs := 300
    static Key := "*$CapsLock"                                              ; * = counts with modifiers held, $ = keyboard hook, see above
    static RepeatGapMs := 1000                                              ; when the last press (or auto-repeat) is this old without a release, the release event was lost: a new press
    static MinTapMs := 25                                                   ; a shorter tap: sent by remote software to align Caps Lock, doesn't count
    static Clock := () => CapsLockSwitch.Now()                              ; replaced by a fake clock in tests
    static _down := false, _held := false, _shift := false, _downAt := 0, _lastDown := 0, _holdTimer := ""

    static IsMode(value) => (value = "Layout" || value = "Mode")

    ; High-resolution milliseconds (fractional)
    static Now() {
        static frequency := 0
        if !frequency
            DllCall("QueryPerformanceFrequency", "Int64*", &frequency)
        counter := 0
        DllCall("QueryPerformanceCounter", "Int64*", &counter)
        return counter * 1000 / frequency
    }

    ; register: the function that registers hotkeys (key, callback), usually Hotkey or App._TryHotkey
    static Register(register) {
        A_HotkeyInterval := 0                                               ; turn off the "too many hotkeys" warning, see above
        register(CapsLockSwitch.Key, (*) => CapsLockSwitch.Down())
        register(CapsLockSwitch.Key " up", (*) => CapsLockSwitch.Up())
    }

    static Down() {
        now := CapsLockSwitch.Clock.Call()
        if (CapsLockSwitch._down && now - CapsLockSwitch._lastDown < CapsLockSwitch.RepeatGapMs) {   ; auto-repeat while held
            CapsLockSwitch._lastDown := now
            return
        }
        CapsLockSwitch._down := true, CapsLockSwitch._held := false, CapsLockSwitch._shift := CapsLockSwitch.ShiftDown()
        CapsLockSwitch._downAt := now, CapsLockSwitch._lastDown := now
        if (CapsLockSwitch._holdTimer = "")
            CapsLockSwitch._holdTimer := () => CapsLockSwitch.Hold()
        SetTimer(CapsLockSwitch._holdTimer, -CapsLockSwitch.HoldMs)
    }

    ; held for HoldMs: toggle Caps Lock (once)
    static Hold() {
        if (!CapsLockSwitch._down || CapsLockSwitch._held)
            return
        CapsLockSwitch._held := true
        CapsLockSwitch.ToggleCapsLock()
    }

    static Up() {
        if !CapsLockSwitch._down
            return
        CapsLockSwitch._down := false
        if (CapsLockSwitch._holdTimer != "")
            SetTimer(CapsLockSwitch._holdTimer, 0)
        if CapsLockSwitch._held
            return
        pressed := CapsLockSwitch.Clock.Call() - CapsLockSwitch._downAt
        if (pressed < CapsLockSwitch.MinTapMs)                              ; sent by remote software, see above
            return
        if (CapsLockSwitch._shift || pressed >= CapsLockSwitch.HoldMs) {    ; Shift+CapsLock: as before; the timer hasn't had a chance to run yet (the program is busy): counts as held too
            CapsLockSwitch.ToggleCapsLock()
        } else {
            if Logger.Enabled
                Logger.Debug("CapsLock: pressed " Round(pressed) " ms, switching input (" CapsLockSwitch.Mode ")")
            CapsLockSwitch.Switch()
        }
    }

    ; Toggles Caps Lock and shows with the HUD whether it is on or off (there is no other feedback while holding, and the local light isn't visible under remote control)
    static ToggleCapsLock() {
        on := !GetKeyState("CapsLock", "T")
        SetCapsLockState(on ? "On" : "Off")
        CapsLockSwitch.Notify.Call(on)
    }

    static NotifyMs := 1000
    static Notify := (on) => App.CapsLockNotice(on)                        ; in the helper for administrator windows ALTRun shows it instead (CapsLockHelper)
    static ShiftDown() => GetKeyState("Shift")
    static StateText(on) => I18n.T(on ? "Caps.On" : "Caps.Off")

    static Switch() {
        if (CapsLockSwitch.Mode = "Mode")
            Win.ToggleImeMode()
        else
            Win.NextInputLanguage()
    }
}
