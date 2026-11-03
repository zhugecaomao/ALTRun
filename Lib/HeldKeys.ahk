;===============================================================================
; HeldKeys.ahk - modifier keys a window saw go down but not up, and keys still held (AutoHotkey v2)
;-------------------------------------------------------------------------------
; A hotkey such as Alt+Space shows another window while Alt is down: the window in front got Alt's key-down, but Alt's key-up goes
; to the new window. Windows doesn't send it to the first window later, so a program that goes by the key messages it got still
; thinks Alt is down when it is back in front, and takes the next key as Alt+key (Total Commander's Synchronize directories window
; took a stray Enter key-up as Alt+Enter and opened a properties window). Release sends that window the missing key-ups; before
; them an unassigned key (vkE8, as AutoHotkey's #MenuMaskKey), so a lone Alt key-up doesn't open its menu bar.
;
; Usage:
;   held := HeldKeys.Held()                    ; ["LAlt"]: the modifiers physically down now
;   HeldKeys.Release(held)                     ; to the window in front: key-ups of those not held any more -> how many were sent
;   HeldKeys.WaitUp(["Enter", "NumpadEnter"], 500)   ; until none of them is held, at most 500 ms -> ms waited
;===============================================================================

class HeldKeys {
    static Modifiers := Map("LAlt", 0xA4, "RAlt", 0xA5, "LCtrl", 0xA2, "RCtrl", 0xA3, "LShift", 0xA0, "RShift", 0xA1)
    static Extended := Map(0xA5, true, 0xA3, true)                          ; right Alt / Ctrl: extended keys
    static MaskKey := 0xE8

    static Held() {
        held := []
        for name in HeldKeys.Modifiers
            if GetKeyState(name, "P")
                held.Push(name)
        return held
    }

    ; Key-ups to the window in front for the keys in names that aren't physically down now; the mask key first when one of them is Alt
    static Release(names, isDown := "") {
        isDown := IsObject(isDown) ? isDown : (name) => GetKeyState(name, "P")
        ups := []
        for name in names
            if (HeldKeys.Modifiers.Has(name) && !isDown(name))
                ups.Push(HeldKeys.Modifiers[name])
        if !ups.Length
            return 0
        events := []
        for vk in ups
            if (vk = 0xA4 || vk = 0xA5) {
                events.Push([HeldKeys.MaskKey, false], [HeldKeys.MaskKey, true])
                break
            }
        for vk in ups
            events.Push([vk, true])
        HeldKeys._Send(events)
        return ups.Length
    }

    ; Waits until none of names is physically held, at most maxMs. isDown / wait are replaced in the tests
    static WaitUp(names, maxMs, isDown := "", wait := "") {
        isDown := IsObject(isDown) ? isDown : (name) => GetKeyState(name, "P")
        wait := IsObject(wait) ? wait : Sleep
        started := A_TickCount, waited := 0
        loop {
            held := false
            for name in names
                held := held || isDown(name)
            if (!held || waited >= maxMs)
                return waited
            wait(10)
            waited := A_TickCount - started
        }
    }

    ; SendInput: [[vk, up]...] as one batch, so nothing typed meanwhile comes in between
    static _Send(events) {
        size := 8 + 3 * A_PtrSize + 8                                       ; INPUT: type, then KEYBDINPUT padded to the union's size (40 bytes on 64-bit)
        buf := Buffer(size * events.Length, 0)
        for index, event in events {
            offset := (index - 1) * size, vk := event[1]
            NumPut("UInt", 1, buf, offset)                                  ; INPUT_KEYBOARD
            NumPut("UShort", vk, buf, offset + A_PtrSize)
            NumPut("UShort", DllCall("MapVirtualKey", "UInt", vk, "UInt", 0, "UInt"), buf, offset + A_PtrSize + 2)
            NumPut("UInt", (event[2] ? 0x2 : 0) | (HeldKeys.Extended.Has(vk) ? 0x1 : 0), buf, offset + A_PtrSize + 4)   ; KEYEVENTF_KEYUP, KEYEVENTF_EXTENDEDKEY
        }
        return DllCall("SendInput", "UInt", events.Length, "Ptr", buf, "Int", size, "UInt")
    }
}
