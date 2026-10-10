;===============================================================================
; CapsLockSwitch.ahk - 和 macOS 一样: 按一下 CapsLock 切换输入法, 按住才是大写锁定 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 设置 (ALTRun.json -> General.CapsLock):
;   ""       不用 (默认), CapsLock 照常
;   "Layout" 切换到下一个输入法 (和 Win+Space 一样, 例如 英语(美国) <-> 微软拼音)
;   "Mode"   切换当前中文输入法的 中 / 英 模式 (和微软拼音里按 Shift 一样)
; 按住超过 HoldMs 才开 / 关大写锁定 (不用等松开), 并显示 "大写锁定: 开 / 关" 的提示;
; Shift+CapsLock 和原来一样直接开 / 关大写锁定 (也显示提示)。
;
; 按下和松开分成两个热键, 只在松开时 (而且没有按住满 HoldMs) 切换输入法, 按住期间什么都不切换。
; 不用 KeyWait 等松开: 切换输入法时 (或者有的输入法) 会刷新按键状态, KeyWait 以为已经松开,
; 按住时自动重复的每一下都切换一次输入法。按住时的自动重复 (连续的按下) 不算新的一次。
; 热键用键盘钩子 ($): 单独的 CapsLock 会用系统的 RegisterHotkey 注册, 拦不住 CapsLock 本身的大写切换。
;
; 远程控制 (Chrome 远程桌面等): 本机按 CapsLock 时本机的大写锁定也会变, 远程软件发下一个按键前会补发
; 一次 CapsLock (按下和松开连在一起, 不到 1 ms) 把远程电脑的大写锁定对齐。按下到松开不到 MinTapMs 的
; 当作这种补发, 拦下但什么都不做 (人按一下至少几十毫秒); 否则每打一个字都会切换一次输入法。
; 计时用高精度的 QueryPerformanceCounter: A_TickCount 的精度只有 10 ~ 16 ms, 分不清。
; 补发时常常按着别的键 (按 Alt+Space 时补发的是 Alt+CapsLock, 还有 Ctrl+C、Shift+字母...), 所以热键带 *:
; 不管按着什么修饰键都拦下, 否则漏过去的那一次会真的切换大写锁定 (呼出窗口后大写被锁定)。
; 按住方向键等时每次自动重复都会补发一次, 很快超过 AutoHotkey "2 秒内最多 70 次热键" 的限制, 弹出
; A_MaxHotkeysPerInterval 的警告 (警告期间按键没被拦下, 大写锁定又被切换); 注册时关掉这个检查
; (A_HotkeyInterval := 0): 这个检查是防止 Send 触发自己的热键无限循环, 带 $ 的热键不会被自己 Send 触发。
;
; 用法:
;   App 注册热键时: CapsLockSwitch.Mode := "Layout", CapsLockSwitch.Register(Hotkey)
;===============================================================================

class CapsLockSwitch {
    static Mode := "", HoldMs := 300
    static Key := "*$CapsLock"                                              ; * = 按着修饰键也算, $ = 用键盘钩子, 见上面
    static RepeatGapMs := 1000                                              ; 上一次按下 (或自动重复) 过了这么久还没松开: 当作松开事件丢了, 是新的一次
    static MinTapMs := 25                                                   ; 更短的按一下: 远程软件补发来对齐大写锁定的, 不算
    static Clock := () => CapsLockSwitch.Now()                              ; 测试时换成假的时钟
    static _down := false, _held := false, _shift := false, _downAt := 0, _lastDown := 0, _holdTimer := ""

    static IsMode(value) => (value = "Layout" || value = "Mode")

    ; 高精度的毫秒数 (小数)
    static Now() {
        static frequency := 0
        if !frequency
            DllCall("QueryPerformanceFrequency", "Int64*", &frequency)
        counter := 0
        DllCall("QueryPerformanceCounter", "Int64*", &counter)
        return counter * 1000 / frequency
    }

    ; register: 注册热键的函数 (key, callback), 一般是 Hotkey 或 App._TryHotkey
    static Register(register) {
        A_HotkeyInterval := 0                                               ; 关掉 "热键太频繁" 的警告, 见上面
        register(CapsLockSwitch.Key, (*) => CapsLockSwitch.Down())
        register(CapsLockSwitch.Key " up", (*) => CapsLockSwitch.Up())
    }

    static Down() {
        now := CapsLockSwitch.Clock.Call()
        if (CapsLockSwitch._down && now - CapsLockSwitch._lastDown < CapsLockSwitch.RepeatGapMs) {   ; 按住时的自动重复
            CapsLockSwitch._lastDown := now
            return
        }
        CapsLockSwitch._down := true, CapsLockSwitch._held := false, CapsLockSwitch._shift := CapsLockSwitch.ShiftDown()
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
        pressed := CapsLockSwitch.Clock.Call() - CapsLockSwitch._downAt
        if (pressed < CapsLockSwitch.MinTapMs)                              ; 远程软件补发的, 见上面
            return
        if (CapsLockSwitch._shift || pressed >= CapsLockSwitch.HoldMs) {    ; Shift+CapsLock: 和原来一样; 定时器还没来得及运行 (程序忙): 也算按住
            CapsLockSwitch.ToggleCapsLock()
        } else {
            if Logger.Enabled
                Logger.Debug("CapsLock: pressed " Round(pressed) " ms, switching input (" CapsLockSwitch.Mode ")")
            CapsLockSwitch.Switch()
        }
    }

    ; 开 / 关大写锁定, 并用 HUD 提示现在是开还是关 (按住时看不到别的反馈, 远程控制时也看不到本机的指示灯)
    static ToggleCapsLock() {
        on := !GetKeyState("CapsLock", "T")
        SetCapsLockState(on ? "On" : "Off")
        App.Notify(CapsLockSwitch.StateText(on), CapsLockSwitch.NotifyMs)
    }

    static NotifyMs := 1000
    static ShiftDown() => GetKeyState("Shift")
    static StateText(on) => I18n.T(on ? "Caps.On" : "Caps.Off")

    static Switch() {
        if (CapsLockSwitch.Mode = "Mode")
            Win.ToggleImeMode()
        else
            Win.NextInputLanguage()
    }
}
