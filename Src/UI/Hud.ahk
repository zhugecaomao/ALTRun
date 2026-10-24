;===============================================================================
; Hud.ahk - short notices after an action (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like Alfred's "Copied to clipboard": a small rounded window in the current theme's colors and font,
; fading in, staying duration ms, fading out. Doesn't take focus; clicks go through.
; Fixed position: below the search window, centered, while it is open (above if there is no room below); otherwise low in the middle of a screen:
;   - the search window just hid (e.g. after "Pin", within RecentWindowMs): the screen it was on
;   - other times (snippet expansion, background notices): chosen like the search window by Appearance.ShowOn (primary screen / under the mouse / active window's)
;
; Usage:
;   Hud.Show("Copied")                    usually called through App.Notify(text, duration)
;   Hud.Hide()
;===============================================================================

class Hud {
    static Gui := ""
    static MaxWidth := 520                  ; wider than this (pixels at 96 DPI) wraps
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
        if (textW > maxW) {                                                 ; too wide: wrap at the maximum width (controls can't be deleted one by one, so it is hidden first)
            label.Visible := false
            label := g.AddText("xm ym Center w" maxW, text)
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

    ; Where a w x h notice goes. anchor: {Window: the search window {X, Y, W, H} or "", Area: work area {Left, Top, Right, Bottom}}
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

    static RecentWindowMs := 3000

    static _Anchor() {
        if SearchWindow.IsVisible() {
            try {
                WinGetPos(&x, &y, &w, &h, "ahk_id " SearchWindow.Gui.Hwnd)
                return {Window: {X: x, Y: y, W: w, H: h}, Area: Win.WorkAreaAt(x + w // 2, y + h // 2)}
            }
        }
        if (IsObject(SearchWindow.LastArea) && A_TickCount - SearchWindow.HiddenAt < Hud.RecentWindowMs)
            return {Window: "", Area: SearchWindow.LastArea}
        return {Window: "", Area: Hud.ScreenArea(AppSettings.Appearance["ShowOn"])}
    }

    ; screen chosen like the search window's Appearance.ShowOn; Active uses the current active window
    static ScreenArea(showOn) {
        switch showOn, false {
            case "Primary": return Win.WorkArea(MonitorGetPrimary())
            case "Active":
                area := ""
                try area := Win.WorkAreaOfWindow(WinExist("A"))
                if IsObject(area)
                    return area
        }
        return Win.WorkAreaAtMouse()
    }

    static _Color(key, fallback) {
        color := ThemeManager.Get(key)
        return (color != "") ? color : fallback
    }

    static _StartFadeOut(g) {
        if (Hud.Gui != g)                   ; already replaced by a new notice
            return
        Hud._fadingOut := true
        if !IsObject(Hud._timer) {
            Hud._timer := () => Hud._Step()
            SetTimer(Hud._timer, 15)
        }
    }

    ; adjust the opacity every 15 ms: fade in about 100 ms, fade out about 150 ms
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
                if IsObject(Hud._timer)                                     ; Hide / a new Show may have run in between and cleared it
                    SetTimer(Hud._timer, 0)
                Hud._timer := ""
            }
        }
        try WinSetTransparent(Hud._alpha, Hud.Gui.Hwnd)
    }
}
