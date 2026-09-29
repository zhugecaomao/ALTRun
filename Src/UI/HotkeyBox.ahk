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
; 录制期间每 200 ms 检查一次: 框不再有焦点、或窗口已经关掉时结束录制 (关窗口时不一定收到 LoseFocus,
; 不结束的话 ALTRun 的热键会一直暂停、按键一直被拦下)。
;
; 用法:
;   box := HotkeyBox.Add(gui, "x10 y10 w150", "!Space", mouse := false, onChange := fn)
;   HotkeyBox.Value(box)                -> "!Space" (AutoHotkey 写法; 录到新热键时调用 onChange)
;   HotkeyBox.CancelActive()            关闭窗口前调用: 正在录制时结束录制, 恢复 ALTRun 的热键
;   HotkeyBox.Label("~MButton")         -> "鼠标中键" (列表等处显示用)
;   HotkeyBox.Compose("^!", "c")        -> "^!c"
;   HotkeyBox.IsAllowed("a")            -> false (要加 Ctrl / Alt / Win)
;===============================================================================

class HotkeyBox {
    static Boxes := Map()                                                   ; 输入框 Hwnd -> {Ctrl, Gui, Value, Mouse, OnChange}
    static _active := "", _hook := "", _suspended := false, _mouseHooked := false, _watchTimer := ""
    static _held := Map()                                                   ; 录制时按住的修饰键 (LControl, RAlt...) -> true
    static TipId := 20

    static Add(g, options, value, mouse := false, onChange := "") {
        HotkeyBox.Prune()
        ctrl := g.AddEdit(options " r1 -Multi -TabStop", "")                ; 不接受 Tab 焦点: 打开窗口、按 Tab 经过时不会开始录制
        state := {Ctrl: ctrl, Hwnd: ctrl.Hwnd, GuiHwnd: g.Hwnd, Value: value, Mouse: mouse, OnChange: onChange}
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

    ; 去掉已经关掉的窗口里的框: 窗口句柄会被系统重新使用, 留着的话在别的控件上点中键会被当成录制
    static Prune() {
        for hwnd in [HotkeyBox.Boxes*]
            if !HotkeyBox.IsAlive(HotkeyBox.Boxes[hwnd])
                HotkeyBox.Boxes.Delete(hwnd)
    }

    static IsAlive(state) {
        try return (state.Ctrl.Hwnd = state.Hwnd)
        return false
    }

    static CancelActive() {
        state := HotkeyBox._active
        HotkeyBox._End()                                                    ; 不管是否还在录制都执行一次: 被打断的 _End 可能没做完
        if IsObject(state)
            HotkeyBox._Show(state)
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
        Win.SetCueBanner(state.Hwnd, I18n.T("Hotkey.Press"))
        HotkeyBox._ShowTip(state, I18n.T("Hotkey.PressTip"))
        if !A_IsSuspended {
            try TraySetIcon(, , true)                                       ; 暂停热键时托盘图标不要变成 "S"
            Suspend(true)
            HotkeyBox._suspended := true
        }
        HotkeyBox._held := Map()
        for key in ["LControl", "RControl", "LAlt", "RAlt", "LShift", "RShift", "LWin", "RWin"]   ; 开始录制前已经按住的
            if GetKeyState(key, "P")
                HotkeyBox._held[key] := true
        ih := InputHook("L0")
        ih.KeyOpt("{All}", "NS")                                            ; 所有按键: 通知 + 拦下
        ih.OnKeyDown := (ih, vk, sc) => HotkeyBox._OnKeyDown(vk, sc)
        ih.OnKeyUp := (ih, vk, sc) => HotkeyBox._OnKeyUp(vk, sc)
        ih.Start()
        HotkeyBox._hook := ih
        if !IsObject(HotkeyBox._watchTimer)
            HotkeyBox._watchTimer := () => HotkeyBox._Watch()
        SetTimer(HotkeyBox._watchTimer, 200)
    }

    ; 框没有焦点了 (点了别处、切换了窗口) 或窗口已经关掉: 结束录制
    static _Watch() {
        state := HotkeyBox._active
        if !IsObject(state)
            return SetTimer(HotkeyBox._watchTimer, 0)
        if (!DllCall("IsWindow", "Ptr", state.Hwnd) || DllCall("GetFocus", "Ptr") != state.Hwnd)
            HotkeyBox.CancelActive()
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
        if !DllCall("IsWindow", "Ptr", state.Hwnd)                          ; 窗口已经关掉
            return HotkeyBox.CancelActive()
        name := GetKeyName(Format("vk{:x}sc{:x}", vk, sc))
        if HotkeyBox.IsModifier(name) {
            HotkeyBox._held[name] := true
            return HotkeyBox._OnModifiers()
        }
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

    static _OnKeyUp(vk, sc) {
        name := GetKeyName(Format("vk{:x}sc{:x}", vk, sc))
        if HotkeyBox.IsModifier(name) {
            if HotkeyBox._held.Has(name)
                HotkeyBox._held.Delete(name)
            HotkeyBox._OnModifiers()
        }
    }

    static IsModifier(name) {
        return RegExMatch(name, "i)^[LR]?(Control|Ctrl|Shift|Alt|Win)$") ? true : false
    }

    ; 按住修饰键时先显示 "Ctrl+Alt+"
    static _OnModifiers() {
        state := HotkeyBox._active
        if !IsObject(state)
            return
        if !DllCall("IsWindow", "Ptr", state.Hwnd)                          ; 窗口已经关掉
            return HotkeyBox.CancelActive()
        mods := HotkeyBox._Mods()
        state.Ctrl.Value := (mods = "") ? "" : HotkeyBox.Label(mods "x")
        if (mods != "")
            state.Ctrl.Value := SubStr(state.Ctrl.Value, 1, -1)             ; "Ctrl+Alt+X" -> "Ctrl+Alt+"
    }

    ; 按住的修饰键 -> "^!+#" 的组合。录制时以 InputHook 收到的按下 / 松开为准 (_held): 远程桌面
    ; (Chrome Remote Desktop、RDP...) 和其它程序模拟的按键不算 "物理按下", 被拦下后逻辑状态也不变,
    ; 只看 GetKeyState 的话 Ctrl+Alt+K 会被当成单独的 K。不在录制时 (在框里点鼠标中键) 看按键状态
    static _Mods() {
        mods := ""
        for pair in [["Control", "^"], ["Alt", "!"], ["Shift", "+"], ["Win", "#"]] {
            down := false
            for side in ["L", "R", ""]                                     ; "" = 不分左右的 Control / Shift / Alt (有的程序这样模拟按键)
                if (HotkeyBox._held.Has(side pair[1]) || (side != "" && !IsObject(HotkeyBox._active) && GetKeyState(side pair[1])))
                    down := true
            if (pair[1] = "Control" && HotkeyBox._held.Has("Ctrl"))
                down := true
            if down
                mods .= pair[2]
        }
        return mods
    }

    ; 在框里点鼠标中键 / 侧键 (只在 mouse = true 的框里)
    static _OnMouse(wParam, msg, hwnd) {
        if !HotkeyBox.Boxes.Has(hwnd)
            return
        if !HotkeyBox.IsAlive(state := HotkeyBox.Boxes[hwnd])
            return HotkeyBox.Prune()
        if !state.Mouse
            return
        key := (msg = 0x207) ? "MButton" : (((wParam >> 16) & 0xFFFF) = 1) ? "XButton1" : "XButton2"
        HotkeyBox._Finish(state, HotkeyBox.Compose(HotkeyBox._Mods(), key))
        return 0
    }

    static _Finish(state, value) {
        HotkeyBox._End()
        changed := (value != state.Value)
        state.Value := value
        HotkeyBox._Show(state)
        if (changed && IsObject(state.OnChange))
            state.OnChange.Call()
        try PostMessage(0x28, 0, 0, , "ahk_id " state.GuiHwnd)           ; WM_NEXTDLGCTL: 焦点移到下一个控件
    }

    ; 结束录制, 可以重复调用。先恢复 ALTRun 的热键、停掉拦截按键的 InputHook: 后面的 ToolTip 等会处理
    ; 窗口消息, 这时正在关的窗口可能打断这个线程, 放在后面的话热键可能一直暂停
    static _End() {
        if HotkeyBox._suspended {
            HotkeyBox._suspended := false
            Suspend(false)
        }
        if IsObject(ih := HotkeyBox._hook) {
            HotkeyBox._hook := ""
            try ih.Stop()
        }
        HotkeyBox._active := "", HotkeyBox._held := Map()
        if IsObject(HotkeyBox._watchTimer)
            SetTimer(HotkeyBox._watchTimer, 0)
        try TraySetIcon(, , false)
        ToolTip(, , , HotkeyBox.TipId)
    }

    static _Show(state) {
        if !DllCall("IsWindow", "Ptr", state.Hwnd)                          ; 窗口已经关掉
            return
        Win.SetCueBanner(state.Hwnd, I18n.T("Hotkey.None"))
        try state.Ctrl.Value := HotkeyBox.Label(state.Value)
    }

    static _ShowTip(state, text) {
        try {
            WinGetPos(&x, &y, , &h, "ahk_id " state.Hwnd)
            CoordMode("ToolTip", "Screen")
            ToolTip(text, x, y + h + 2, HotkeyBox.TipId)
        }
    }
}
