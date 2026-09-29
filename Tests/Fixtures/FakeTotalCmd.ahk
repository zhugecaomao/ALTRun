;===============================================================================
; FakeTotalCmd.ahk - 测试用: 模拟 Total Commander 的 WM_COPYDATA 接口 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 收到 dwData = "GW" 的 WM_COPYDATA (内容是 ANSI 命令) 时, 和 TC 一样马上用 dwData = "RW" (UTF-16)
; 回复给 wParam 里的窗口: SP = 当前面板路径, TP = 另一个面板路径, A = 当前是哪一边。
; 由 RunTests.ahk 启动, 20 秒后自己退出。
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
    command := StrGet(NumGet(lParam, A_PtrSize * 2, "Ptr"), NumGet(lParam, A_PtrSize, "UInt"), "CP0")    ; 到结尾的 0 为止
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
