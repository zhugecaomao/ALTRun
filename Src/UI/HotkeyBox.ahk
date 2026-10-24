;===============================================================================
; HotkeyBox.ahk - an input box that records a hotkey (AutoHotkey v2)
;-------------------------------------------------------------------------------
; As in Alfred / PowerToys: the box shows forms like "Ctrl+Alt+C"; click it to start recording
; and press the combination to record it. The settings still store the AutoHotkey form (^!c), so the settings file format is unchanged.
;   Esc / Tab cancel (the old hotkey stays), Backspace / Delete clear it (no hotkey)
;   single letters, digits, Space etc. (which would get in the way of typing) need Ctrl / Alt / Win; F1 ~ F24 etc. work alone
;   with mouse = true, a middle / side button click in the box records a mouse hotkey (MButton / XButton1 / XButton2)
; While recording, InputHook catches every key (Alt+Space, Win+E don't trigger system or other programs' functions), and
; ALTRun's own hotkeys are suspended (Suspend), otherwise pressing the show hotkey would show the search window instead of being recorded.
; After recording or canceling, focus moves to the next control (without selecting its text): the box can't be typed in; click it again to record again.
; Special forms in old settings (~, *, CapsLock & J...) stay as they are until recorded again.
; While recording, every 200 ms: recording ends when the box has lost focus or the window has closed (closing a window doesn't always send LoseFocus,
; and without ending ALTRun's hotkeys would stay suspended and keys caught); when the window moved, the tip below the box follows.
;
; Usage:
;   box := HotkeyBox.Add(gui, "x10 y10 w150", "!Space", mouse := false, onChange := fn)
;   HotkeyBox.Value(box)                -> "!Space" (AutoHotkey form; onChange is called when a new hotkey is recorded)
;   HotkeyBox.CancelActive()            before closing a window: ends a recording in progress and restores ALTRun's hotkeys
;   HotkeyBox.Label("~MButton")         -> "Middle button" (for display in lists etc.)
;   HotkeyBox.Compose("^!", "c")        -> "^!c"
;   HotkeyBox.IsAllowed("a")            -> false (needs Ctrl / Alt / Win)
;===============================================================================

class HotkeyBox {
    static Boxes := Map()                                                   ; input box Hwnd -> {Ctrl, Gui, Value, Mouse, OnChange}
    static _active := "", _hook := "", _suspended := false, _mouseHooked := false, _watchTimer := ""
    static _held := Map()                                                   ; modifier keys held while recording (LControl, RAlt...) -> true
    static TipId := 20

    static Add(g, options, value, mouse := false, onChange := "") {
        HotkeyBox.Prune()
        ctrl := g.AddEdit(options " r1 -Multi -TabStop", "")                ; no Tab focus: opening the window or tabbing through doesn't start recording
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

    ; drop boxes of windows that have closed: window handles get reused by the system, and a middle click on another control would count as recording
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
        HotkeyBox._End()                                                    ; run once whether or not it is still recording: an interrupted _End may not have finished
        if IsObject(state)
            HotkeyBox._Show(state)
    }

    static Value(ctrl) {
        return HotkeyBox.Boxes.Has(ctrl.Hwnd) ? HotkeyBox.Boxes[ctrl.Hwnd].Value : ctrl.Value
    }

    static Label(hk) {
        return Win.HotkeyLabel(hk, HotkeyBox.Names())
    }

    ; Display names of mouse buttons (keyboard keys use Windows' English names: Space, F1, Home...)
    static Names() {
        return Map("MButton", I18n.T("Hotkey.MButton"), "XButton1", I18n.T("Hotkey.XButton1"), "XButton2", I18n.T("Hotkey.XButton2")
                 , "LButton", I18n.T("Hotkey.LButton"), "RButton", I18n.T("Hotkey.RButton"))
    }

    ; mods: any combination of "^" "!" "+" "#"; key: the name GetKeyName gives. Modifiers sorted Ctrl Alt Shift Win, letters lower case
    static Compose(mods, key) {
        ordered := ""
        for m in ["^", "!", "+", "#"]
            if InStr(mods, m)
                ordered .= m
        return ordered ((StrLen(key) = 1) ? StrLower(key) : key)
    }

    ; Keys that would get in the way of typing (letters, digits, symbols, Space, Enter, numpad...) need Ctrl / Alt / Win; Shift doesn't count
    static IsAllowed(hk) {
        if !RegExMatch(hk, "^([\^!+#]*)(.+)$", &m)
            return false
        if RegExMatch(m[1], "[\^!#]")
            return true
        key := m[2]
        return !(StrLen(key) = 1 || RegExMatch(key, "i)^(Space|Enter|Tab|Backspace|Delete|Escape|Numpad.*)$"))
    }

    ;---------------------------------------------------------------------------
    ; Recording
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
            try TraySetIcon(, , true)                                       ; the tray icon shouldn't turn into "S" while hotkeys are suspended
            Suspend(true)
            HotkeyBox._suspended := true
        }
        HotkeyBox._held := Map()
        for key in ["LControl", "RControl", "LAlt", "RAlt", "LShift", "RShift", "LWin", "RWin"]   ; already held before recording started
            if GetKeyState(key, "P")
                HotkeyBox._held[key] := true
        ih := InputHook("L0")
        ih.KeyOpt("{All}", "NS")                                            ; all keys: notify + catch
        ih.OnKeyDown := (ih, vk, sc) => HotkeyBox._OnKeyDown(vk, sc)
        ih.OnKeyUp := (ih, vk, sc) => HotkeyBox._OnKeyUp(vk, sc)
        ih.Start()
        HotkeyBox._hook := ih
        if !IsObject(HotkeyBox._watchTimer)
            HotkeyBox._watchTimer := () => HotkeyBox._Watch()
        SetTimer(HotkeyBox._watchTimer, 200)
    }

    ; the box lost focus (clicked elsewhere, switched windows) or the window closed: end recording
    static _Watch() {
        state := HotkeyBox._active
        if !IsObject(state)
            return SetTimer(HotkeyBox._watchTimer, 0)
        if (!DllCall("IsWindow", "Ptr", state.Hwnd) || DllCall("GetFocus", "Ptr") != state.Hwnd)
            return HotkeyBox.CancelActive()
        if (state.HasOwnProp("Tip") && state.Tip != "")
            HotkeyBox._ShowTip(state, state.Tip, true)
    }

    ; Lost focus (clicked elsewhere, switched windows): cancel recording, keep the old hotkey
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
        if !DllCall("IsWindow", "Ptr", state.Hwnd)                          ; the window has closed
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

    ; while modifiers are held, show "Ctrl+Alt+" first
    static _OnModifiers() {
        state := HotkeyBox._active
        if !IsObject(state)
            return
        if !DllCall("IsWindow", "Ptr", state.Hwnd)                          ; the window has closed
            return HotkeyBox.CancelActive()
        mods := HotkeyBox._Mods()
        state.Ctrl.Value := (mods = "") ? "" : HotkeyBox.Label(mods "x")
        if (mods != "")
            state.Ctrl.Value := SubStr(state.Ctrl.Value, 1, -1)             ; "Ctrl+Alt+X" -> "Ctrl+Alt+"
    }

    ; Held modifiers -> a combination of "^!+#". While recording, the key downs / ups InputHook received count (_held): keys from remote desktops
    ; (Chrome Remote Desktop, RDP...) and other programs' simulated keys aren't "physically down", and caught keys don't change the logical state,
    ; so by GetKeyState alone Ctrl+Alt+K would be taken as K. When not recording (a middle click in the box) the key states are used
    static _Mods() {
        mods := ""
        for pair in [["Control", "^"], ["Alt", "!"], ["Shift", "+"], ["Win", "#"]] {
            down := false
            for side in ["L", "R", ""]                                     ; "" = Control / Shift / Alt without left or right (some programs simulate keys that way)
                if (HotkeyBox._held.Has(side pair[1]) || (side != "" && !IsObject(HotkeyBox._active) && GetKeyState(side pair[1])))
                    down := true
            if (pair[1] = "Control" && HotkeyBox._held.Has("Ctrl"))
                down := true
            if down
                mods .= pair[2]
        }
        return mods
    }

    ; a middle / side button click in the box (only in boxes with mouse = true)
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
        HotkeyBox._FocusNext(state)
    }

    ; Moves focus to the next control Tab can reach; for an edit box, its text isn't selected (caret at the start). Not WM_NEXTDLGCTL: it
    ; selects all the text (as if tabbed into), and an edit box that got focus when the window opened keeps everything selected. With a list like "target windows" below,
    ; the next character typed would replace the whole list
    static _FocusNext(state) {
        try {
            if !(next := DllCall("GetNextDlgTabItem", "Ptr", state.GuiHwnd, "Ptr", state.Hwnd, "Int", false, "Ptr"))
                return
            DllCall("SetFocus", "Ptr", next)
            if (WinGetClass("ahk_id " next) = "Edit")
                SendMessage(0xB1, 0, 0, next)                               ; EM_SETSEL: nothing selected
        }
    }

    ; Ends recording; may be called repeatedly. ALTRun's hotkeys are restored and the key-catching InputHook stopped first: ToolTip etc. below handle
    ; window messages, and a window that is closing may interrupt this thread then; done later, the hotkeys might stay suspended
    static _End() {
        if HotkeyBox._suspended {
            HotkeyBox._suspended := false
            Suspend(false)
        }
        if IsObject(ih := HotkeyBox._hook) {
            HotkeyBox._hook := ""
            try ih.Stop()
        }
        if IsObject(state := HotkeyBox._active)
            state.Tip := ""
        HotkeyBox._active := "", HotkeyBox._held := Map()
        if IsObject(HotkeyBox._watchTimer)
            SetTimer(HotkeyBox._watchTimer, 0)
        try TraySetIcon(, , false)
        ToolTip(, , , HotkeyBox.TipId)
    }

    static _Show(state) {
        if !DllCall("IsWindow", "Ptr", state.Hwnd)                          ; the window has closed
            return
        Win.SetCueBanner(state.Hwnd, I18n.T("Hotkey.None"))
        try state.Ctrl.Value := HotkeyBox.Label(state.Value)
    }

    ; The tip below the box (ToolTip uses screen coordinates and doesn't follow the window). onlyIfMoved: not shown again if the box didn't move (no flicker)
    static _ShowTip(state, text, onlyIfMoved := false) {
        try {
            WinGetPos(&x, &y, , &h, "ahk_id " state.Hwnd)
            at := x "," (y + h + 2)
            if (onlyIfMoved && state.HasOwnProp("TipAt") && state.TipAt = at)
                return
            CoordMode("ToolTip", "Screen")
            ToolTip(text, x, y + h + 2, HotkeyBox.TipId)
            state.Tip := text, state.TipAt := at
        }
    }
}
