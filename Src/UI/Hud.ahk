;===============================================================================
; Hud.ahk - 操作后的简短提示 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 Alfred 的 "Copied to clipboard" 一样: 一块跟随当前主题颜色和字体的圆角小窗,
; 淡入, 停留 duration 毫秒后淡出。不抢焦点, 鼠标可以点穿。
; 位置固定: 搜索窗口开着时在它下方居中 (下面放不下就放上面), 否则在鼠标所在
; 屏幕的中间偏下。
;
; 用法:
;   Hud.Show("已复制")                    一般通过 App.Notify(text, duration) 调用
;   Hud.Hide()
;===============================================================================

class Hud {
    static Gui := ""
    static MaxWidth := 520                  ; 超过这个宽度 (96 DPI 下的像素) 自动换行
    static Alpha := 240
    static _alpha := 0
    static _timer := ""
    static _fadingOut := false

    static Show(text, duration := 1500) {
        Hud.Hide()
        if (Trim(text) = "")
            return

        g := Gui("-Caption +AlwaysOnTop +ToolWindow -DPIScale +E0x08000020", "ALTRun HUD")   ; WS_EX_NOACTIVATE | WS_EX_TRANSPARENT
        g.BackColor := Hud._Color("Background", "FAFAFA")
        g.MarginX := Win.Scale(20), g.MarginY := Win.Scale(11)
        g.SetFont("s11 c" Hud._Color("Title", "1F1F1F"), ThemeManager.FontName())
        maxW := Win.Scale(Hud.MaxWidth)
        label := g.AddText("Center", text)
        label.GetPos(, , &textW)
        if (textW > maxW) {
            label.Destroy()
            label := g.AddText("Center w" maxW, text)
        }
        g.Show("Hide AutoSize")
        g.GetPos(, , &w, &h)
        pos := Hud.Position(w, h, Hud._Anchor())
        WinSetTransparent(0, g.Hwnd)
        g.Show("NA x" pos.X " y" pos.Y)
        Win.SetCorner(g.Hwnd)
        Win.SetBorderColor(g.Hwnd, Hud._Color("Border", "C8C8C8"))
        Hud.Gui := g

        Hud._alpha := 0, Hud._fadingOut := false
        Hud._timer := () => Hud._Step()
        SetTimer(Hud._timer, 15)
        SetTimer(() => Hud._StartFadeOut(g), -Max(duration, 300))
    }

    static Hide() {
        if IsObject(Hud._timer)
            SetTimer(Hud._timer, 0)
        Hud._timer := ""
        if IsObject(Hud.Gui) {
            try Hud.Gui.Destroy()
            Hud.Gui := ""
        }
    }

    ; w x h 的提示放在哪里。anchor: {Window: 搜索窗口 {X, Y, W, H} 或 "", Area: 工作区 {Left, Top, Right, Bottom}}
    static Position(w, h, anchor) {
        area := anchor.Area, gap := Win.Scale(12)
        if IsObject(anchor.Window) {
            box := anchor.Window
            x := box.X + (box.W - w) // 2
            y := box.Y + box.H + gap
            if (y + h > area.Bottom)
                y := box.Y - gap - h
        } else {
            x := area.Left + (area.Right - area.Left - w) // 2
            y := area.Top + Round((area.Bottom - area.Top) * 2 / 3) - h // 2
        }
        x := Max(area.Left, Min(x, area.Right - w))
        y := Max(area.Top, Min(y, area.Bottom - h))
        return {X: x, Y: y}
    }

    static _Anchor() {
        if SearchWindow.IsVisible() {
            try {
                WinGetPos(&x, &y, &w, &h, "ahk_id " SearchWindow.Gui.Hwnd)
                return {Window: {X: x, Y: y, W: w, H: h}, Area: Win.WorkAreaAt(x + w // 2, y + h // 2)}
            }
        }
        return {Window: "", Area: Win.WorkAreaAtMouse()}
    }

    static _Color(key, fallback) {
        color := ThemeManager.Get(key)
        return (color != "") ? color : fallback
    }

    static _StartFadeOut(g) {
        if (Hud.Gui != g)                   ; 已经换成了新的提示
            return
        Hud._fadingOut := true
        if !IsObject(Hud._timer) {
            Hud._timer := () => Hud._Step()
            SetTimer(Hud._timer, 15)
        }
    }

    ; 每 15 ms 调整一次透明度: 淡入约 100 ms, 淡出约 150 ms
    static _Step() {
        if !IsObject(Hud.Gui)
            return Hud.Hide()
        if Hud._fadingOut {
            Hud._alpha -= 25
            if (Hud._alpha <= 0)
                return Hud.Hide()
        } else {
            Hud._alpha := Min(Hud._alpha + 40, Hud.Alpha)
            if (Hud._alpha >= Hud.Alpha) {
                SetTimer(Hud._timer, 0)
                Hud._timer := ""
            }
        }
        try WinSetTransparent(Hud._alpha, Hud.Gui.Hwnd)
    }
}
