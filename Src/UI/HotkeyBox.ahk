;===============================================================================
; HotkeyBox.ahk - 录制热键的输入框 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 Alfred / PowerToys 一样: 框里显示 "Ctrl+Alt+C" 这样的写法, 点一下开始录制,
; 按下想要的组合键就记下来; 设置里保存的仍是 AutoHotkey 的写法 (^!c), 设置文件格式不变。
;   Esc / Tab 取消 (保留原来的热键), Backspace / Delete 清除 (不用这个热键)
;   单独的字母、数字、空格等 (会影响平时打字) 要加 Ctrl / Alt / Win; F1 ~ F24 等可以单独用
;   mouse = true 时, 在框里点鼠标中键 / 侧键录成鼠标热键 (MButton / XButton1 / XButton2)
; 录制期间用 InputHook 拦下所有按键 (Alt+Space、Win+E 不会触发系统或其它程序的功能), 并暂停
; ALTRun 自己的热键 (Suspend), 否则按呼出热键会弹出搜索窗口, 而不是被记录。
; 录完或取消后焦点移到下一个控件: 框里不能打字, 想重新录制就再点一下。
; 旧设置里的特殊写法 (~、*、CapsLock & J...) 在重新录制之前原样保留。
;
; 用法:
;   box := HotkeyBox.Add(gui, "x10 y10 w150", "!Space", mouse := false, onChange := fn)
;   HotkeyBox.Value(box)                -> "!Space" (AutoHotkey 写法; 录到新热键时调用 onChange)
;   HotkeyBox.Label("~MButton")         -> "鼠标中键" (列表等处显示用)
;   HotkeyBox.Compose("^!", "c")        -> "^!c"
;   HotkeyBox.IsAllowed("a")            -> false (要加 Ctrl / Alt / Win)
;===============================================================================

class HotkeyBox {
    static Boxes := Map()                                                   ; 输入框 Hwnd -> {Ctrl, Gui, Value, Mouse, OnChange}
    static _active := "", _hook := "", _suspended := false, _mouseHooked := false
    static TipId := 20

    static Add(g, options, value, mouse := false, onChange := "") {
        ctrl := g.AddEdit(options " r1 -Multi -TabStop", "")                ; 不接受 Tab 焦点: 打开窗口、按 Tab 经过时不会开始录制
        state := {Ctrl: ctrl, Gui: g, Value: value, Mouse: mouse, OnChange: onChange}
        HotkeyBox.Boxes[ctrl.Hwnd] := state
        HotkeyBox._Show(state)
        ctrl.OnEvent("Focus", (*) => HotkeyBox.Start(state))
        ctrl.OnEvent("LoseFocus", (*) => HotkeyBox.Stop(state))
        if (mouse && !HotkeyBox._mouseHooked) {
            HotkeyBox._mouseHooked := true
            OnMessage(0x207, (wParam, lParam, msg, hwnd) => HotkeyBox._OnMouse(wParam, msg, hwnd))   ; WM_MBUTTONDOWN
            OnMessage(0x20B, (wParam, lParam, msg, hwnd) => HotkeyBox._OnMouse(wParam, msg, hwnd))   ; WM_XBUTTONDOWN
        }
        return ctrl
    }

    static Value(ctrl) {
        return HotkeyBox.Boxes.Has(ctrl.Hwnd) ? HotkeyBox.Boxes[ctrl.Hwnd].Value : ctrl.Value
    }

    static Label(hk) {
        return Win.HotkeyLabel(hk, HotkeyBox.Names())
    }

    ; 鼠标按键的显示名称 (键盘按键用 Windows 的英文名称: Space、F1、Home...)
    static Names() {
        return Map("MButton", I18n.T("Hotkey.MButton"), "XButton1", I18n.T("Hotkey.XButton1"), "XButton2", I18n.T("Hotkey.XButton2")
                 , "LButton", I18n.T("Hotkey.LButton"), "RButton", I18n.T("Hotkey.RButton"))
    }

    ; mods: "^" "!" "+" "#" 的任意组合; key: GetKeyName 给出的名称。修饰符按 Ctrl Alt Shift Win 排好, 字母小写
    static Compose(mods, key) {
        ordered := ""
        for m in ["^", "!", "+", "#"]
            if InStr(mods, m)
                ordered .= m
        return ordered ((StrLen(key) = 1) ? StrLower(key) : key)
    }

    ; 会影响平时打字的键 (字母、数字、符号、空格、回车、小键盘...) 要和 Ctrl / Alt / Win 一起用; Shift 不算
    static IsAllowed(hk) {
        if !RegExMatch(hk, "^([\^!+#]*)(.+)$", &m)
            return false
        if RegExMatch(m[1], "[\^!#]")
            return true
        key := m[2]
        return !(StrLen(key) = 1 || RegExMatch(key, "i)^(Space|Enter|Tab|Backspace|Delete|Escape|Numpad.*)$"))
    }

    ;---------------------------------------------------------------------------
    ; 录制
    ;---------------------------------------------------------------------------
    static Start(state) {
        if (HotkeyBox._active == state)
            return
        HotkeyBox._End()
        HotkeyBox._active := state
        state.Ctrl.Value := ""
        Win.SetCueBanner(state.Ctrl.Hwnd, I18n.T("Hotkey.Press"))
        HotkeyBox._ShowTip(state, I18n.T("Hotkey.PressTip"))
        if !A_IsSuspended {
            try TraySetIcon(, , true)                                       ; 暂停热键时托盘图标不要变成 "S"
            Suspend(true)
            HotkeyBox._suspended := true
        }
        ih := InputHook("L0")
        ih.KeyOpt("{All}", "NS")                                            ; 所有按键: 通知 + 拦下
        ih.OnKeyDown := (ih, vk, sc) => HotkeyBox._OnKeyDown(vk, sc)
        ih.OnKeyUp := (ih, vk, sc) => HotkeyBox._OnModifiers()
        ih.Start()
        HotkeyBox._hook := ih
    }

    ; 失去焦点 (点了别处、切换了窗口): 取消录制, 保留原来的热键
    static Stop(state) {
        if (HotkeyBox._active == state) {
            HotkeyBox._End()
            HotkeyBox._Show(state)
        }
    }

    static _OnKeyDown(vk, sc) {
        state := HotkeyBox._active
        if !IsObject(state)
            return
        name := GetKeyName(Format("vk{:x}sc{:x}", vk, sc))
        if RegExMatch(name, "i)^[LR]?(Control|Ctrl|Shift|Alt|Win)$")
            return HotkeyBox._OnModifiers()
        mods := HotkeyBox._Mods()
        if (name = "Tab" && (mods = "" || mods = "+"))
            return HotkeyBox._Finish(state, state.Value)
        if (mods = "") {
            switch name {
                case "Escape": return HotkeyBox._Finish(state, state.Value)
                case "Backspace", "Delete": return HotkeyBox._Finish(state, "")
            }
        }
        hk := HotkeyBox.Compose(mods, name)
        if !HotkeyBox.IsAllowed(hk) {
            state.Ctrl.Value := ""
            HotkeyBox._ShowTip(state, I18n.T("Hotkey.NeedModifier", HotkeyBox.Label(hk)))
            return
        }
        HotkeyBox._Finish(state, hk)
    }

    ; 按住修饰键时先显示 "Ctrl+Alt+"
    static _OnModifiers() {
        state := HotkeyBox._active
        if !IsObject(state)
            return
        mods := HotkeyBox._Mods()
        state.Ctrl.Value := (mods = "") ? "" : HotkeyBox.Label(mods "x")
        if (mods != "")
            state.Ctrl.Value := SubStr(state.Ctrl.Value, 1, -1)             ; "Ctrl+Alt+X" -> "Ctrl+Alt+"
    }

    static _Mods() {
        mods := ""
        for pair in [["Ctrl", "^"], ["Alt", "!"], ["Shift", "+"], ["LWin", "#"], ["RWin", "#"]]
            if (GetKeyState(pair[1], "P") && !InStr(mods, pair[2]))
                mods .= pair[2]
        return mods
    }

    ; 在框里点鼠标中键 / 侧键 (只在 mouse = true 的框里)
    static _OnMouse(wParam, msg, hwnd) {
        if !HotkeyBox.Boxes.Has(hwnd) || !HotkeyBox.Boxes[hwnd].Mouse
            return
        key := (msg = 0x207) ? "MButton" : (((wParam >> 16) & 0xFFFF) = 1) ? "XButton1" : "XButton2"
        HotkeyBox._Finish(HotkeyBox.Boxes[hwnd], HotkeyBox.Compose(HotkeyBox._Mods(), key))
        return 0
    }

    static _Finish(state, value) {
        HotkeyBox._End()
        changed := (value != state.Value)
        state.Value := value
        HotkeyBox._Show(state)
        if (changed && IsObject(state.OnChange))
            state.OnChange.Call()
        try PostMessage(0x28, 0, 0, , "ahk_id " state.Gui.Hwnd)          ; WM_NEXTDLGCTL: 焦点移到下一个控件
    }

    static _End() {
        if IsObject(HotkeyBox._hook) {
            try HotkeyBox._hook.Stop()
            HotkeyBox._hook := ""
        }
        HotkeyBox._active := ""
        ToolTip(, , , HotkeyBox.TipId)
        if HotkeyBox._suspended {
            Suspend(false)
            try TraySetIcon(, , false)
            HotkeyBox._suspended := false
        }
    }

    static _Show(state) {
        Win.SetCueBanner(state.Ctrl.Hwnd, I18n.T("Hotkey.None"))
        state.Ctrl.Value := HotkeyBox.Label(state.Value)
    }

    static _ShowTip(state, text) {
        try {
            WinGetPos(&x, &y, , &h, "ahk_id " state.Ctrl.Hwnd)
            CoordMode("ToolTip", "Screen")
            ToolTip(text, x, y + h + 2, HotkeyBox.TipId)
        }
    }
}
