;===============================================================================
; WindowFade.ahk - 搜索窗口显示 / 隐藏时淡入淡出 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 打开 Appearance.Fade 时用 (偏好设置 → 窗口, #158):
;   - 淡入: 显示前把窗口的透明度设成 0, 显示后用计时器逐步升到主题的透明度 (Opacity, 没有时不透明)。
;     窗口照常马上拿到焦点, 淡入期间打的字照样进搜索框
;   - 淡出: 真正的窗口照常马上隐藏 (焦点回到原来的程序、接着运行程序 / 粘贴都不受影响, 也不变慢)。
;     隐藏前把它在屏幕上的样子复制到一个不抢焦点、鼠标点得穿的分层窗口 (影子) 里, 让影子淡出
;   - 同一时间只有一个淡入或淡出: 开始新的、或者再次显示窗口时, 上一个马上结束
;===============================================================================

class WindowFade {
    static InMs := 120, OutMs := 100, StepMs := 10
    static _fade := "", _timer := "", _ghost := ""

    ; 开始后第 elapsed 毫秒时的透明度, 先快后慢 (ease-out), 到 ms 时正好是 to
    static Alpha(elapsed, ms, from, to) {
        t := (ms > 0) ? Max(0, Min(elapsed / ms, 1)) : 1
        return Round(from + (to - from) * (1 - (1 - t) ** 3))
    }

    ; 窗口显示之前调用: 透明度设成 0, 显示后升到 target (1-255)
    static In(hwnd, target) {
        WindowFade.Stop()
        WindowFade.SetAlpha(hwnd, 0)
        WindowFade._Start(WindowFade.InMs, 0, target, (a) => WindowFade.SetAlpha(hwnd, a))
    }

    ; 窗口隐藏之前调用: 在同一个位置显示它现在的样子 (影子), 隐藏后影子淡出。复制不了时 (窗口不在屏幕上、
    ; 锁屏时...) 不显示影子, 窗口照常直接隐藏
    static Out(hwnd) {
        ghost := WindowFade._ShowGhost(hwnd)
        WindowFade.Stop()                                                   ; 正在淡入: 窗口恢复主题的透明度 (被影子盖着, 看不到)
        if ghost
            WindowFade._Start(WindowFade.OutMs, 255, 0, (a) => WindowFade._GhostAlpha(a), () => WindowFade._HideGhost())
    }

    ; 结束正在进行的淡入 / 淡出: 窗口直接到最终的透明度, 影子直接消失
    static Stop() {
        if (WindowFade._timer != "")
            SetTimer(WindowFade._timer, 0)
        fade := WindowFade._fade, WindowFade._fade := ""
        if IsObject(fade)
            WindowFade._Finish(fade)
    }

    static Running => IsObject(WindowFade._fade)

    ; 窗口整体的透明度 (分层窗口, WS_EX_LAYERED); 窗口隐藏时也能设
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
    ; 影子: 一个分层窗口, 内容用 UpdateLayeredWindow 一次给好 (显示出来就是完整的样子, 不会先白一下)
    ;---------------------------------------------------------------------------
    static _GhostGui() {
        if !IsObject(WindowFade._ghost)                                     ; 分层 + 鼠标点得穿 + 不抢焦点
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
        ok := DllCall("BitBlt", "Ptr", mem, "Int", 0, "Int", 0, "Int", w, "Int", h, "Ptr", screen, "Int", x, "Int", y, "UInt", 0x00CC0020)   ; SRCCOPY: 屏幕上看到的样子
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", screen)
        if ok {
            g := WindowFade._GhostGui()
            pos := Buffer(8), NumPut("Int", x, "Int", y, pos)
            size := Buffer(8), NumPut("Int", w, "Int", h, size)
            ok := DllCall("UpdateLayeredWindow", "Ptr", g.Hwnd, "Ptr", 0, "Ptr", pos, "Ptr", size, "Ptr", mem, "Int64*", 0
                , "UInt", 0, "UInt*", 255 << 16, "UInt", 2)                 ; BLENDFUNCTION 不透明, ULW_ALPHA
            if ok                                                           ; 放到最上面显示, 不激活
                DllCall("SetWindowPos", "Ptr", g.Hwnd, "Ptr", -1, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x53)
        }
        DllCall("SelectObject", "Ptr", mem, "Ptr", old)
        DllCall("DeleteObject", "Ptr", bitmap)
        DllCall("DeleteDC", "Ptr", mem)
        return ok ? true : false
    }

    static _GhostAlpha(alpha) {                                             ; 内容不变, 只改透明度
        DllCall("UpdateLayeredWindow", "Ptr", WindowFade._ghost.Hwnd, "Ptr", 0, "Ptr", 0, "Ptr", 0, "Ptr", 0, "Ptr", 0
            , "UInt", 0, "UInt*", Round(alpha) << 16, "UInt", 2)
    }

    static _HideGhost() {
        if IsObject(WindowFade._ghost)
            DllCall("ShowWindow", "Ptr", WindowFade._ghost.Hwnd, "Int", 0)
    }

    static GhostVisible => IsObject(WindowFade._ghost) && DllCall("IsWindowVisible", "Ptr", WindowFade._ghost.Hwnd)
}
