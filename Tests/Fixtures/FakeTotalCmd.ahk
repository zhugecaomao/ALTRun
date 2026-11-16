;===============================================================================
; FakeTotalCmd.ahk - for tests: imitates Total Commander's WM_COPYDATA interface (AutoHotkey v2)
;-------------------------------------------------------------------------------
; On a WM_COPYDATA with dwData = "GW" (an ANSI command) it replies at once, like TC, with dwData = "RW" (UTF-16)
; to the window in wParam: SP = path of the active panel, TP = path of the other panel, A = which side is active.
; Commands sent to it (as to TC, by TC Commands.ahk) are added to the file in its first argument, if given, one per line:
; "cm <number>" for WM_USER+51, "em <name>" for WM_COPYDATA with dwData = "EM" (an ANSI command name).
; Started by the tests; exits by itself after 20 s.
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon

replies := Map("SP", "C:\Work\Current\", "TP", "\\server\share\Other\", "A", "L", "LP", "ftp://example.com/pub/")
OnMessage(0x4A, OnCopyData)
OnMessage(0x433, (wParam, *) => Record("cm " wParam))                    ; WM_USER+51
g := Gui("+ToolWindow", "ALTRun Fake Total Commander")
g.Show("NA x0 y0 w160 h40")
SetTimer(() => ExitApp(), -20000)

OnCopyData(wParam, lParam, msg, hwnd) {
    if (NumGet(lParam, 0, "UPtr") = Ord("E") + 256 * Ord("M"))
        return (Record("em " StrGet(NumGet(lParam, A_PtrSize * 2, "Ptr"), "CP0")), true)
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

Record(line) {
    if A_Args.Length
        FileAppend(line "`n", A_Args[1], "UTF-8")
    return 1
}
