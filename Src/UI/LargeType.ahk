;===============================================================================
; LargeType.ahk - 大字显示 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 Alfred 的 Large Type 一样: 屏幕中央用超大字号显示一段文字 (电话号码、
; 计算结果、密码...), 按任意键、点击或切换窗口后关闭。字号按文字长度自动选。
;
; 用法:
;   LargeType.Show("138 0013 8000")
;===============================================================================

class LargeType {
    static Gui := ""
    static _handlersReady := false

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
        if (textW > maxW) {                                                 ; 太宽就按最大宽度自动换行
            label.Destroy()
            label := g.AddText("Center w" maxW, text)
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
