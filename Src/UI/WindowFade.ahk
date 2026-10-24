;===============================================================================
; WindowFade.ahk - fades the search window in and out when showing / hiding (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Used when Appearance.Fade is on (Preferences → Window, #158):
;   - fade in: the window's opacity is set to 0 before showing and raised on a timer afterwards to the theme's opacity (Opacity, opaque if none).
;     The window still gets focus at once, and characters typed during the fade go into the search box
;   - fade out: the real window still hides at once (focus goes back to the previous program; running programs / pasting next is unaffected and no slower).
;     Before hiding, its look on screen is copied into a layered window (the shadow) that doesn't take focus and lets clicks through, and the shadow fades out
;   - one fade in or out at a time: starting a new one, or showing the window again, ends the previous one at once
;===============================================================================

class WindowFade {
    static InMs := 120, OutMs := 100, StepMs := 10
    static _fade := "", _timer := "", _ghost := ""

    ; Opacity elapsed ms after the start, fast then slow (ease-out), exactly to at ms
    static Alpha(elapsed, ms, from, to) {
        t := (ms > 0) ? Max(0, Min(elapsed / ms, 1)) : 1
        return Round(from + (to - from) * (1 - (1 - t) ** 3))
    }

    ; Call before the window shows: opacity 0, raised to target (1-255) after showing
    static In(hwnd, target) {
        WindowFade.Stop()
        WindowFade.SetAlpha(hwnd, 0)
        WindowFade._Start(WindowFade.InMs, 0, target, (a) => WindowFade.SetAlpha(hwnd, a))
    }

    ; Call before the window hides: shows its current look at the same place (the shadow), which fades out after hiding. When it can't be copied (the window isn't on screen,
    ; the screen is locked...) there is no shadow and the window just hides as usual
    static Out(hwnd) {
        ghost := WindowFade._ShowGhost(hwnd)
        WindowFade.Stop()                                                   ; fading in: the window gets back the theme's opacity (covered by the shadow, not visible)
        if ghost
            WindowFade._Start(WindowFade.OutMs, 255, 0, (a) => WindowFade._GhostAlpha(a), () => WindowFade._HideGhost())
    }

    ; Ends a fade in progress: the window goes straight to its final opacity, the shadow disappears at once
    static Stop() {
        if (WindowFade._timer != "")
            SetTimer(WindowFade._timer, 0)
        fade := WindowFade._fade, WindowFade._fade := ""
        if IsObject(fade)
            WindowFade._Finish(fade)
    }

    static Running => IsObject(WindowFade._fade)

    ; Opacity of the whole window (a layered window, WS_EX_LAYERED); can be set while it is hidden too
    static SetAlpha(hwnd, alpha) {
        saved := A_DetectHiddenWindows
        DetectHiddenWindows(true)
        try WinSetTransparent(Round(alpha), "ahk_id " hwnd)
        DetectHiddenWindows(saved)
    }

    static _Start(ms, from, to, apply, done := "") {
        WindowFade._fade := {Start: A_TickCount, Ms: ms, From: from, To: to, Apply: apply, Done: done}
        if (WindowFade._timer = "")
            WindowFade._timer := () => WindowFade._Tick()
        SetTimer(WindowFade._timer, WindowFade.StepMs)
    }

    static _Tick() {
        fade := WindowFade._fade
        if !IsObject(fade)
            return SetTimer(WindowFade._timer, 0)
        elapsed := A_TickCount - fade.Start
        if (elapsed >= fade.Ms)
            return WindowFade.Stop()
        try fade.Apply.Call(WindowFade.Alpha(elapsed, fade.Ms, fade.From, fade.To))
    }

    static _Finish(fade) {
        try fade.Apply.Call(fade.To)
        if IsObject(fade.Done)
            try fade.Done.Call()
    }

    ;---------------------------------------------------------------------------
    ; The shadow: a layered window whose content is given in one go with UpdateLayeredWindow (it appears complete, without a white flash)
    ;---------------------------------------------------------------------------
    static _GhostGui() {
        if !IsObject(WindowFade._ghost)                                     ; layered + click-through + doesn't take focus
            WindowFade._ghost := Gui("-Caption +ToolWindow +AlwaysOnTop -DPIScale +E0x8080020")
        return WindowFade._ghost
    }

    static _ShowGhost(hwnd) {
        if !DllCall("IsWindowVisible", "Ptr", hwnd)
            return false
        try WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
        catch
            return false
        if (w <= 0 || h <= 0)
            return false
        screen := DllCall("GetDC", "Ptr", 0, "Ptr")
        mem := DllCall("CreateCompatibleDC", "Ptr", screen, "Ptr")
        bitmap := DllCall("CreateCompatibleBitmap", "Ptr", screen, "Int", w, "Int", h, "Ptr")
        old := DllCall("SelectObject", "Ptr", mem, "Ptr", bitmap, "Ptr")
        ok := DllCall("BitBlt", "Ptr", mem, "Int", 0, "Int", 0, "Int", w, "Int", h, "Ptr", screen, "Int", x, "Int", y, "UInt", 0x00CC0020)   ; SRCCOPY: the look on screen
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", screen)
        if ok {
            g := WindowFade._GhostGui()
            pos := Buffer(8), NumPut("Int", x, "Int", y, pos)
            size := Buffer(8), NumPut("Int", w, "Int", h, size)
            ok := DllCall("UpdateLayeredWindow", "Ptr", g.Hwnd, "Ptr", 0, "Ptr", pos, "Ptr", size, "Ptr", mem, "Int64*", 0
                , "UInt", 0, "UInt*", 255 << 16, "UInt", 2)                 ; BLENDFUNCTION opaque, ULW_ALPHA
            if ok                                                           ; shown on top, not activated
                DllCall("SetWindowPos", "Ptr", g.Hwnd, "Ptr", -1, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x53)
        }
        DllCall("SelectObject", "Ptr", mem, "Ptr", old)
        DllCall("DeleteObject", "Ptr", bitmap)
        DllCall("DeleteDC", "Ptr", mem)
        return ok ? true : false
    }

    static _GhostAlpha(alpha) {                                             ; same content, only the opacity changes
        DllCall("UpdateLayeredWindow", "Ptr", WindowFade._ghost.Hwnd, "Ptr", 0, "Ptr", 0, "Ptr", 0, "Ptr", 0, "Ptr", 0
            , "UInt", 0, "UInt*", Round(alpha) << 16, "UInt", 2)
    }

    static _HideGhost() {
        if IsObject(WindowFade._ghost)
            DllCall("ShowWindow", "Ptr", WindowFade._ghost.Hwnd, "Int", 0)
    }

    static GhostVisible => IsObject(WindowFade._ghost) && DllCall("IsWindowVisible", "Ptr", WindowFade._ghost.Hwnd)
}
