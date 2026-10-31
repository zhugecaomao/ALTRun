;===============================================================================
; KeyWatch.ahk - a diagnostic log of a few keys and the window in front (AutoHotkey v2)
;-------------------------------------------------------------------------------
; For finding out which keys reach which window, e.g. when Total Commander opens a properties dialog right after a result
; was run with Enter (2026-10). Runs on its own, so it keeps logging while ALTRun exits and restarts. Not shipped with ALTRun.
;
; Logged to %TEMP%\KeyWatch.log (a new file each start; the previous one is kept as KeyWatch.old.log):
;   - key down / up of Enter, Numpad Enter, Alt, Space, Esc, Ctrl, Shift, Win, the Menu key and F10 only (no letters or
;     digits, so nothing typed is recorded), with the window in front, and whether the key came from the keyboard or was
;     sent by a program ("sent"; "sent by AHK" when an AutoHotkey script such as ALTRun sent it)
;   - every change of the window in front: class, title, program, owner window
;
; Usage: run it with AutoHotkey v2 (double-click), keep it running while you work; when the problem happens, send the
; lines around that time. Exit from its tray icon.
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

KeyWatch.Start()

class KeyWatch {
    static File := A_Temp "\KeyWatch.log"
    static Keys := Map(0x0D, "Enter", 0x10, "Shift", 0x11, "Ctrl", 0x12, "Alt", 0x1B, "Esc", 0x20, "Space", 0x5B, "LWin", 0x5C, "RWin"
                     , 0x5D, "Menu", 0x79, "F10", 0xA0, "LShift", 0xA1, "RShift", 0xA2, "LCtrl", 0xA3, "RCtrl", 0xA4, "LAlt", 0xA5, "RAlt")
    static _lines := "", _hook := 0, _winHook := 0, _keyProc := 0, _winProc := 0, _lastDown := Map()

    static Start() {
        if FileExist(KeyWatch.File)
            FileMove(KeyWatch.File, A_Temp "\KeyWatch.old.log", true)
        KeyWatch._Add("KeyWatch started; " A_OSVersion ", AutoHotkey " A_AhkVersion ", ALTRun " (ProcessExist("ALTRun.exe") ? "running" : "not running (or run as a script)"))
        KeyWatch._keyProc := CallbackCreate(KeyWatch._OnKey.Bind(KeyWatch), "F", 3)
        KeyWatch._hook := DllCall("SetWindowsHookEx", "Int", 13, "Ptr", KeyWatch._keyProc, "Ptr", DllCall("GetModuleHandle", "Ptr", 0, "Ptr"), "UInt", 0, "Ptr")   ; WH_KEYBOARD_LL
        KeyWatch._winProc := CallbackCreate(KeyWatch._OnForeground.Bind(KeyWatch), "F", 7)
        KeyWatch._winHook := DllCall("SetWinEventHook", "UInt", 0x0003, "UInt", 0x0003, "Ptr", 0, "Ptr", KeyWatch._winProc, "UInt", 0, "UInt", 0, "UInt", 0, "Ptr")   ; EVENT_SYSTEM_FOREGROUND, out of context
        if !KeyWatch._hook || !KeyWatch._winHook
            KeyWatch._Add("could not install the hooks (keyboard " KeyWatch._hook ", window " KeyWatch._winHook ")")
        A_IconTip := "KeyWatch: logging to " KeyWatch.File
        A_TrayMenu.Insert("1&", "Open the log", (*) => Run('notepad.exe "' KeyWatch.File '"'))
        A_TrayMenu.Insert("2&")
        OnExit((*) => KeyWatch._Stop())
        SetTimer(() => KeyWatch._Flush(), 500)
    }

    ; KBDLLHOOKSTRUCT: vkCode, scanCode, flags (0x01 extended, 0x10 injected, 0x80 up), time, dwExtraInfo.
    ; Kept short: the hook only notes the line, the timer writes it
    static _OnKey(nCode, wParam, lParam) {
        if (nCode = 0) {
            vk := NumGet(lParam, 0, "UInt"), flags := NumGet(lParam, 8, "UInt")
            if KeyWatch.Keys.Has(vk) {
                up := flags & 0x80
                name := (vk = 0x0D && (flags & 0x01)) ? "NumpadEnter" : KeyWatch.Keys[vk]
                repeat := !up && KeyWatch._lastDown.Has(vk) && KeyWatch._lastDown[vk]
                KeyWatch._lastDown[vk] := !up
                extra := NumGet(lParam, 16, "UPtr") & 0xFFFFFFFF                ; AutoHotkey marks what it sends with 0xFFC3D4xx (the send level in the low bits)
                source := !(flags & 0x10) ? "keyboard" : ((extra & 0xFFFFFF00) = 0xFFC3D400) ? "sent by AHK" : "sent"
                KeyWatch._Add(Format("{:-11} {:-6} {:-11} in {}", name, up ? "up" : repeat ? "repeat" : "down", source, KeyWatch._Window(DllCall("GetForegroundWindow", "Ptr"))))
            }
        }
        return DllCall("CallNextHookEx", "Ptr", 0, "Int", nCode, "UPtr", wParam, "Ptr", lParam, "Ptr")
    }

    static _OnForeground(hook, event, hwnd, idObject, idChild, thread, time) {
        if (idObject != 0 || !hwnd)
            return
        owner := DllCall("GetWindow", "Ptr", hwnd, "UInt", 4, "Ptr")         ; GW_OWNER
        KeyWatch._Add("FOREGROUND  " KeyWatch._Window(hwnd) (owner ? ", owner " KeyWatch._Window(owner, false) : ""))
    }

    ; "TTOTAL_CMD of TOTALCMD64.exe 'Total Commander'" (0x... hwnd first)
    static _Window(hwnd, withTitle := true) {
        if !hwnd
            return "(none)"
        cls := "", title := "", exe := ""
        try cls := WinGetClass("ahk_id " hwnd)
        try exe := WinGetProcessName("ahk_id " hwnd)
        if withTitle
            try title := SubStr(WinGetTitle("ahk_id " hwnd), 1, 60)
        return Format("0x{:X} {} of {}", hwnd, cls, exe) (withTitle && title != "" ? " '" title "'" : "")
    }

    static _Add(text) {
        KeyWatch._lines .= FormatTime(, "HH:mm:ss") "." Format("{:03}", A_MSec) "  " text "`n"
    }

    static _Flush() {
        if (KeyWatch._lines = "")
            return
        text := KeyWatch._lines, KeyWatch._lines := ""
        try FileAppend(text, KeyWatch.File, "UTF-8")
    }

    static _Stop() {
        if KeyWatch._hook
            DllCall("UnhookWindowsHookEx", "Ptr", KeyWatch._hook)
        if KeyWatch._winHook
            DllCall("UnhookWinEvent", "Ptr", KeyWatch._winHook)
        KeyWatch._Add("KeyWatch stopped")
        KeyWatch._Flush()
    }
}
