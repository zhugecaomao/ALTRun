;===============================================================================
; ProbeFont.ahk - 临时诊断: 刚启动的 Windows 上第一次量中文文字要多久 (GetTextExtentExPointW / DrawText)
;   AutoHotkey64.exe ProbeFont.ahk 输出文件
;===============================================================================
#Requires AutoHotkey v2.0
#NoTrayIcon

out := A_Args.Length ? A_Args[1] : A_Temp "\ProbeFont.txt"
Note(text) => FileAppend(Format("{:7} ", A_TickCount - started) text "`n", out, "UTF-8")
started := A_TickCount
try FileDelete(out)

hdc := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
for face in ["Segoe UI", "Microsoft YaHei UI"] {
    font := DllCall("CreateFontW", "Int", -30, "Int", 0, "Int", 0, "Int", 0, "Int", 400, "UInt", 0, "UInt", 0, "UInt", 0, "UInt", 1
        , "UInt", 0, "UInt", 0, "UInt", 5, "UInt", 0, "WStr", face, "Ptr")
    old := DllCall("SelectObject", "Ptr", hdc, "Ptr", font, "Ptr")
    for text in ["Notepad", "记事本", "记事本 2", "メモ帳"] {
        Note(face ": measuring '" text "'")
        t := A_TickCount, fit := 0, partial := Buffer(StrLen(text) * 4, 0), size := Buffer(8, 0)
        DllCall("GetTextExtentExPointW", "Ptr", hdc, "WStr", text, "Int", StrLen(text), "Int", 0x7FFFFFFF, "Int*", &fit, "Ptr", partial, "Ptr", size)
        Note(face ": '" text "' width " NumGet(size, 0, "Int") " in " (A_TickCount - t) " ms")
    }
    DllCall("SelectObject", "Ptr", hdc, "Ptr", old)
    DllCall("DeleteObject", "Ptr", font)
}
DllCall("DeleteDC", "Ptr", hdc)

g := Gui("+ToolWindow", "ProbeFont")
g.SetFont("s14", "Segoe UI")
g.AddText("w300", "看一下: 记事本 ノート")
t := A_TickCount
g.Show("NA x0 y0")
DllCall("UpdateWindow", "Ptr", g.Hwnd)
Note("GUI with CJK text shown in " (A_TickCount - t) " ms")
Note("done")
ExitApp
