; 临时诊断: 空搜索框的提示文字 (cue banner) 和输入同样文字时的上下位置 (按大写字母高度居中后)
#Requires AutoHotkey v2.0
#SingleInstance Force
#Include ..\Lib\JSON.ahk

out := A_Args[1], tag := A_Args[2]
DirCreate(out)
OnError((e, *) => (Log("ERROR " e.Message " (line " e.Line ") " e.Extra), ExitApp(1)))
app := "C:\ALTRunDiag", repo := A_ScriptDir "\.."
try DirDelete(app, true)
DirCreate(app)
FileCopy(repo "\ALTRun.ahk", app "\ALTRun.ahk")
for folder in ["Lib", "Src", "Resources"]
    DirCopy(repo "\" folder, app "\" folder)
Log("DPI " A_ScreenDPI)
for theme in ["Light", "Dark"]
for lang in ["en", "zh-CN"] {
    for statusBar in [1, 0] {
        name := tag "-" theme "-" lang "-sb" statusBar
        DirCreate(app "\Data")
        try FileDelete(app "\Data\ALTRun.json")
        settings := Map("SchemaVersion", 4
            , "General", Map("Language", lang, "LaunchAtLogin", 0, "HideOnDeactivate", 0, "SendToMenu", 0, "StartMenuShortcut", 0, "CheckForUpdates", 0, "ShowTips", 0)
            , "Appearance", Map("Theme", theme, "Width", 700, "StatusBar", statusBar))
        FileAppend(JSON.Stringify(settings, 4), app "\Data\ALTRun.json", "UTF-8")
        Run('"' A_AhkPath '" "' app '\ALTRun.ahk"', app, , &pid)
        Sleep(4500)
        hwnd := WinWait("ALTRun ahk_class AutoHotkeyGUI ahk_pid " pid, , 10)
        if !hwnd {
            Log(name ": no search window; windows of the process:")
            for w in WinGetList("ahk_pid " pid)
                Log("  " WinGetClass(w) " '" WinGetTitle(w) "' visible " ((WinGetStyle(w) & 0x10000000) ? 1 : 0))
            ProcessClose(pid)
            continue
        }
        Sleep(2500)
        WinGetPos(&wx, &wy, &ww, &wh, hwnd)
        ControlGetPos(&ex, &ey, &ew, &eh, "Edit1", hwnd)
        cueText := (lang = "en") ? "ALTRun Search" : "ALTRun 搜索"
        Log(name ": window " ww "x" wh ", edit x" ex " y" ey " w" ew " h" eh ", cue '" cueText "'")
        Measure(name "-cue", hwnd, ex, ey, ew, eh)
        ControlSetText(cueText, "Edit1", hwnd)
        Sleep(1500)
        ControlSend("{End}", "Edit1", hwnd)
        Sleep(1500)
        Measure(name "-typed", hwnd, ex, ey, ew, eh)
        ControlSetText("微信 Wx 记事本 Notepad", "Edit1", hwnd)
        Sleep(1500)
        Measure(name "-mixed", hwnd, ex, ey, ew, eh)
        ProcessClose(pid)
        ProcessWaitClose(pid, 5)
        Sleep(500)
    }
}
ExitApp()

; 输入框区域里和背景不同的像素: 最上 / 最下一行 (窗口坐标), 并存一张窗口上部的截图
Measure(name, hwnd, ex, ey, ew, eh) {
    global out
    WinGetPos(&wx, &wy, &ww, &wh, hwnd)
    h := Min(wh, ey + eh + ey)
    hdcScreen := DllCall("GetDC", "Ptr", 0, "Ptr")
    hdc := DllCall("CreateCompatibleDC", "Ptr", hdcScreen, "Ptr")
    hbm := DllCall("CreateCompatibleBitmap", "Ptr", hdcScreen, "Int", ww, "Int", h, "Ptr")
    old := DllCall("SelectObject", "Ptr", hdc, "Ptr", hbm, "Ptr")
    DllCall("BitBlt", "Ptr", hdc, "Int", 0, "Int", 0, "Int", ww, "Int", h, "Ptr", hdcScreen, "Int", wx, "Int", wy, "UInt", 0x40CC0020)
    bg := DllCall("GetPixel", "Ptr", hdc, "Int", ww - 3, "Int", ey + eh // 2, "UInt")
    top := -1, bottom := -1
    y := 0
    while (y < h) {
        x := ex + 2
        while (x < ex + ew - 2) {
            c := DllCall("GetPixel", "Ptr", hdc, "Int", x, "Int", y, "UInt")
            if (Abs((c & 0xFF) - (bg & 0xFF)) + Abs(((c >> 8) & 0xFF) - ((bg >> 8) & 0xFF)) + Abs(((c >> 16) & 0xFF) - ((bg >> 16) & 0xFF)) > 60) {
                if (top < 0)
                    top := y
                bottom := y
                break
            }
            x += 1
        }
        y += 1
    }
    DllCall("SelectObject", "Ptr", hdc, "Ptr", old)
    SavePng(hbm, out "\" name ".png")
    DllCall("DeleteObject", "Ptr", hbm), DllCall("DeleteDC", "Ptr", hdc), DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdcScreen)
    Log(Format("  {}: ink rows {}..{} (center {}), edit rows {}..{} (center {}), input band 0..{} (center {})"
        , name, top, bottom, (top + bottom) / 2, ey, ey + eh - 1, ey + (eh - 1) / 2, 2 * ey + eh - 1, (2 * ey + eh - 1) / 2))
}

SavePng(hbm, file) {
    static token := 0
    if !token {
        DllCall("LoadLibrary", "Str", "gdiplus")
        input := Buffer(24, 0)
        NumPut("UInt", 1, input)
        DllCall("gdiplus\GdiplusStartup", "Ptr*", &token, "Ptr", input, "Ptr", 0)
    }
    bitmap := 0
    DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "Ptr", hbm, "Ptr", 0, "Ptr*", &bitmap)
    clsid := Buffer(16)
    DllCall("ole32\CLSIDFromString", "Str", "{557CF406-1A04-11D3-9A73-0000F81EF32E}", "Ptr", clsid)
    try FileDelete(file)
    DllCall("gdiplus\GdipSaveImageToFile", "Ptr", bitmap, "WStr", file, "Ptr", clsid, "Ptr", 0)
    DllCall("gdiplus\GdipDisposeImage", "Ptr", bitmap)
}

Log(text) {
    global out
    FileAppend(text "`n", out "\diag.log", "UTF-8")
}
