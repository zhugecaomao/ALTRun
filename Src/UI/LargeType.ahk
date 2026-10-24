;===============================================================================
; LargeType.ahk - Large Type (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like Alfred's Large Type: shows text in a huge font in the middle of the screen (phone numbers,
; calculation results, passwords...); any key, a click or switching windows closes it. The font size follows the text length.
; Can be turned off in Preferences → Features (Features.LargeType.Enabled, off for new installs): then Ctrl+L does nothing and the actions panel doesn't list it.
;
; Usage:
;   LargeType.Show("138 0013 8000")
;===============================================================================

class LargeType {
    static Gui := ""
    static _handlersReady := false

    static Enabled() {
        try return AppSettings.Feature("LargeType")["Enabled"] ? true : false
        return false
    }

    static Show(text) {
        LargeType.Close()
        if (Trim(text) = "")
            return
        LargeType._EnsureHandlers()

        area := Win.WorkAreaAtMouse()
        areaW := area.Right - area.Left, areaH := area.Bottom - area.Top
        len := StrLen(text)
        fontSize := (len <= 12) ? 96 : (len <= 30) ? 64 : (len <= 80) ? 40 : (len <= 300) ? 28 : 18
        pad := Win.Scale(40)

        g := Gui("-Caption +AlwaysOnTop +ToolWindow -DPIScale", "ALTRun Large Type")
        g.BackColor := "141414"
        g.MarginX := pad, g.MarginY := pad
        g.SetFont("s" fontSize " cFFFFFF", ThemeManager.FontName())
        maxW := Round(areaW * 0.85) - 2 * pad
        label := g.AddText("Center", text)
        label.GetPos(, , &textW)
        if (textW > maxW) {                                                 ; too wide: wrap at the maximum width (controls can't be deleted one by one, so it is hidden first)
            label.Visible := false
            label := g.AddText("xm ym Center w" maxW, text)
        }
        label.OnEvent("Click", (*) => LargeType.Close())
        g.OnEvent("Escape", (*) => LargeType.Close())
        g.Show("Hide AutoSize")
        g.GetPos(, , &winW, &winH)
        winH := Min(winH, areaH)
        g.Show("x" (area.Left + (areaW - winW) // 2) " y" (area.Top + (areaH - winH) // 2) " h" winH)
        WinSetTransparent(235, g.Hwnd)
        Win.SetCorner(g.Hwnd)
        LargeType.Gui := g
    }

    static Close() {
        if IsObject(LargeType.Gui) {
            try LargeType.Gui.Destroy()
            LargeType.Gui := ""
        }
    }

    static _EnsureHandlers() {
        if LargeType._handlersReady
            return
        LargeType._handlersReady := true
        OnMessage(0x100, (p*) => LargeType._OnKey(p*))                      ; WM_KEYDOWN
        OnMessage(0x104, (p*) => LargeType._OnKey(p*))                      ; WM_SYSKEYDOWN
        OnMessage(0x6,   (p*) => LargeType._OnActivate(p*))                 ; WM_ACTIVATE
    }

    static _IsOwn(hwnd) {
        return IsObject(LargeType.Gui) && (hwnd = LargeType.Gui.Hwnd || DllCall("GetAncestor", "Ptr", hwnd, "UInt", 2, "Ptr") = LargeType.Gui.Hwnd)
    }

    static _OnKey(wParam, lParam, msg, hwnd) {
        if !LargeType._IsOwn(hwnd)
            return
        LargeType.Close()
        return 0
    }

    static _OnActivate(wParam, lParam, msg, hwnd) {
        if (LargeType._IsOwn(hwnd) && (wParam & 0xFFFF) = 0)
            SetTimer(() => LargeType.Close(), -1)
    }
}
