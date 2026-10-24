;===============================================================================
; FakeTotalCmd.ahk - for tests: imitates Total Commander's WM_COPYDATA interface (AutoHotkey v2)
;-------------------------------------------------------------------------------
; On a WM_COPYDATA with dwData = "GW" (an ANSI command) it replies at once, like TC, with dwData = "RW" (UTF-16)
; to the window in wParam: SP = path of the active panel, TP = path of the other panel, A = which side is active.
; Started by RunTests.ahk; exits by itself after 20 s.
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon

replies := Map("SP", "C:\Work\Current\", "TP", "\\server\share\Other\", "A", "L", "LP", "ftp://example.com/pub/")
OnMessage(0x4A, OnCopyData)
g := Gui("+ToolWindow", "ALTRun Fake Total Commander")
g.Show("NA x0 y0 w160 h40")
SetTimer(() => ExitApp(), -20000)

OnCopyData(wParam, lParam, msg, hwnd) {
    if (NumGet(lParam, 0, "UPtr") != Ord("G") + 256 * Ord("W"))
        return
    command := StrGet(NumGet(lParam, A_PtrSize * 2, "Ptr"), NumGet(lParam, A_PtrSize, "UInt"), "CP0")    ; up to the terminating 0
    reply := replies.Has(command) ? replies[command] : ""
    data := Buffer((StrLen(reply) + 1) * 2, 0)
    StrPut(reply, data, "UTF-16")
    copyData := Buffer(A_PtrSize * 3, 0)
    NumPut("UPtr", Ord("R") + 256 * Ord("W"), copyData, 0)
    NumPut("UInt", data.Size, copyData, A_PtrSize)
    NumPut("Ptr", data.Ptr, copyData, A_PtrSize * 2)
    DllCall("SendMessageW", "Ptr", wParam, "UInt", 0x4A, "Ptr", hwnd, "Ptr", copyData.Ptr)
    return true
}
