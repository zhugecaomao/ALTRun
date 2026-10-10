;===============================================================================
; CapsLockSwitch.ahk - 和 macOS 一样: 按一下 CapsLock 切换输入法, 按住才是大写锁定 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 设置 (ALTRun.json -> General.CapsLock):
;   ""       不用 (默认), CapsLock 照常
;   "Layout" 切换到下一个输入法 (和 Win+Space 一样, 例如 英语(美国) <-> 微软拼音)
;   "Mode"   切换当前中文输入法的 中 / 英 模式 (和微软拼音里按 Shift 一样)
; 按住超过 HoldMs 才开 / 关大写锁定 (不用等松开); Shift+CapsLock 等组合键照常。
;
; 用法:
;   App 注册热键时: CapsLockSwitch.Mode := "Layout", Hotkey("CapsLock", (*) => CapsLockSwitch.Press())
;===============================================================================

class CapsLockSwitch {
    static Mode := "", HoldMs := 300

    static IsMode(value) => (value = "Layout" || value = "Mode")

    static Press() {
        if KeyWait("CapsLock", "T" CapsLockSwitch.HoldMs / 1000) {         ; 很快松开: 切换输入法
            CapsLockSwitch.Switch()
            return
        }
        SetCapsLockState(GetKeyState("CapsLock", "T") ? "Off" : "On")       ; 按住: 原来的大写锁定
        KeyWait("CapsLock")                                                 ; 等松开, 按住时的自动重复不再触发
    }

    static Switch() {
        if (CapsLockSwitch.Mode = "Mode")
            Win.ToggleImeMode()
        else
            Win.NextInputLanguage()
    }
}
