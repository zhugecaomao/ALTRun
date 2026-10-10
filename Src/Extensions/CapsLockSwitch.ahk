;===============================================================================
; CapsLockSwitch.ahk - 和 macOS 一样: 按一下 CapsLock 切换输入法, 按住才是大写锁定 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 设置 (ALTRun.json -> General.CapsLock):
;   ""       不用 (默认), CapsLock 照常
;   "Layout" 切换到下一个输入法 (和 Win+Space 一样, 例如 英语(美国) <-> 微软拼音)
;   "Mode"   切换当前中文输入法的 中 / 英 模式 (和微软拼音里按 Shift 一样)
; 按住超过 HoldMs 才开 / 关大写锁定 (不用等松开); Shift+CapsLock 等组合键照常。
;
; 按下和松开分成两个热键, 只在松开时 (而且没有按住满 HoldMs) 切换输入法, 按住期间什么都不切换。
; 不用 KeyWait 等松开: 切换输入法时 (或者有的输入法) 会刷新按键状态, KeyWait 以为已经松开,
; 按住时自动重复的每一下都切换一次输入法。按住时的自动重复 (连续的按下) 不算新的一次。
; 热键用键盘钩子 ($): 单独的 CapsLock 会用系统的 RegisterHotkey 注册, 拦不住 CapsLock 本身的大写切换。
;
; 用法:
;   App 注册热键时: CapsLockSwitch.Mode := "Layout", CapsLockSwitch.Register(Hotkey)
;===============================================================================

class CapsLockSwitch {
    static Mode := "", HoldMs := 300
    static Key := "$CapsLock"                                               ; $ = 用键盘钩子, 见上面
    static RepeatGapMs := 1000                                              ; 上一次按下 (或自动重复) 过了这么久还没松开: 当作松开事件丢了, 是新的一次
    static Clock := () => A_TickCount                                       ; 测试时换成假的时钟
    static _down := false, _held := false, _downAt := 0, _lastDown := 0, _holdTimer := ""

    static IsMode(value) => (value = "Layout" || value = "Mode")

    ; register: 注册热键的函数 (key, callback), 一般是 Hotkey 或 App._TryHotkey
    static Register(register) {
        register(CapsLockSwitch.Key, (*) => CapsLockSwitch.Down())
        register(CapsLockSwitch.Key " up", (*) => CapsLockSwitch.Up())
    }

    static Down() {
        now := CapsLockSwitch.Clock.Call()
        if (CapsLockSwitch._down && now - CapsLockSwitch._lastDown < CapsLockSwitch.RepeatGapMs) {   ; 按住时的自动重复
            CapsLockSwitch._lastDown := now
            return
        }
        CapsLockSwitch._down := true, CapsLockSwitch._held := false
        CapsLockSwitch._downAt := now, CapsLockSwitch._lastDown := now
        if (CapsLockSwitch._holdTimer = "")
            CapsLockSwitch._holdTimer := () => CapsLockSwitch.Hold()
        SetTimer(CapsLockSwitch._holdTimer, -CapsLockSwitch.HoldMs)
    }

    ; 按住满 HoldMs: 开 / 关大写锁定 (只一次)
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
        if (CapsLockSwitch.Clock.Call() - CapsLockSwitch._downAt >= CapsLockSwitch.HoldMs)   ; 定时器还没来得及运行 (程序忙): 也算按住
            CapsLockSwitch.ToggleCapsLock()
        else
            CapsLockSwitch.Switch()
    }

    static ToggleCapsLock() {
        SetCapsLockState(GetKeyState("CapsLock", "T") ? "Off" : "On")
    }

    static Switch() {
        if (CapsLockSwitch.Mode = "Mode")
            Win.ToggleImeMode()
        else
            Win.NextInputLanguage()
    }
}
