; @altrun.title          Save Screenshot
; @altrun.keyword        shot
; @altrun.mode           silent
; @altrun.author         sfcboy
; @altrun.description    Save the current clipboard image as a PNG file; Save As opens in the current Explorer folder
; @altrun.description_zh 把剪贴板里当前的图片另存为 PNG 文件; "另存为"对话框默认打开当前资源管理器所在文件夹
; @altrun.network        none
;
;
; Trigger options (pick one or both):
;   1) Keyword: type "shot" in ALTRun and press Enter.
;   2) Hotkey:  Preferences -> Hotkeys -> add one, Action = "script:Save Screenshot.ahk"
;              (e.g. Ctrl+Alt+S). Press it right after taking a screenshot.
;
; Optional settings via environment variables (user level):
;   ALTRUN_SHOT_DIR   default folder the Save As dialog opens in (e.g. D:\Pictures\Screenshots)
;   ALTRUN_SHOT_PREFIX default filename prefix (default: screenshot)
; If ALTRUN_SHOT_DIR is set, the file name is auto-generated as <prefix>_yyyyMMdd_HHmmss.png
; and the dialog opens there; otherwise it opens in the current Explorer folder (or the Desktop).
#Requires AutoHotkey v2.0
#NoTrayIcon

Chinese := RegExMatch(A_Language, "i)^(0804|0404|0C04|1004|1404)$") > 0
; In silent mode ALTRun shows the last line printed as a notification (read as UTF-8). try: run any other way
; (double-clicked, or the mode changed to window) there is no standard output, and the save has already worked.
Say(msg) {
    try FileAppend(msg "`n", "*", "UTF-8-RAW")
}
Fail(msg) {
    Say(msg)
    ExitApp(1)
}

; The folder that was in front when the user triggered ALTRun (ALTRUN_HWND) - i.e. the Explorer
; window whose folder the user is looking at. Falls back to the frontmost Explorer window.
ActiveFolder() {
    hwndStr := EnvGet("ALTRUN_HWND")
    hwnd := (hwndStr != "" && IsInteger(hwndStr)) ? Integer(hwndStr) : 0
    shell := ComObject("Shell.Application")
    ; Map each Explorer window handle to its current folder path
    paths := Map()
    for window in shell.Windows() {
        try {
            p := window.Document.Folder.Self.Path
            h := window.HWND
            ; Windows 11 tabs: each tab is an item with the window's HWND; the tab shown wins
            if (p != "" && (!paths.Has(h) || TabOf(window) = CurrentTab(h)))
                paths[h] := p
        }
    }
    ; Only VISIBLE (not minimized) Explorer windows count. A minimized folder window is not the
    ; "current folder" - when all folders are minimized, the user is effectively on the desktop.
    ; 1) Exact window that was in front when ALTRun was triggered, if still visible
    if (hwnd && paths.Has(hwnd)) {
        mm := 0
        try mm := WinGetMinMax("ahk_id " hwnd)
        if (mm != -1)
            return paths[hwnd]
    }
    ; 2) Frontmost VISIBLE Explorer window by Z-order (WinGetList is front-to-back)
    for h in WinGetList() {
        cls := ""
        try cls := WinGetClass("ahk_id " h)
        if (cls = "CabinetWClass" || cls = "ExploreWClass") {
            mm := 0
            try mm := WinGetMinMax("ahk_id " h)
            if (mm = -1)                                ; skip minimized windows
                continue
            if (paths.Has(h))
                return paths[h]
        }
    }
    return ""                                           ; no visible folder -> caller falls back to Desktop
}

; The tab view shown in an Explorer window: its frontmost ShellTabWindowClass child
CurrentTab(hwnd) => DllCall("FindWindowExW", "Ptr", hwnd, "Ptr", 0, "Str", "ShellTabWindowClass", "Ptr", 0, "Ptr")

; The tab view of a Shell.Application window item: IShellBrowser::GetWindow (IOleWindow, the fourth method)
TabOf(window) {
    static IID_IShellBrowser := "{000214E2-0000-0000-C000-000000000046}"
    tab := 0
    try ComCall(3, ComObjQuery(window, IID_IShellBrowser, IID_IShellBrowser), "Ptr*", &tab)
    return tab
}

; --- The image is read from the current clipboard, not from ALTRun's data folder ---
if !ClipboardHasImage()
    Fail(Chinese ? "剪贴板里没有图片" : "No image in the clipboard")

; --- Decide starting folder and default file name ---
startDir := EnvGet("ALTRUN_SHOT_DIR")
prefix := EnvGet("ALTRUN_SHOT_PREFIX")
if (prefix = "")
    prefix := "screenshot"
if (startDir = "")
    startDir := ActiveFolder()          ; folder the user was in when ALTRun was invoked
if (startDir = "" || !DirExist(startDir)) {
    ; No visible folder window -> the user is on the desktop, so default there.
    if DirExist(A_Desktop)
        startDir := A_Desktop
    else
        startDir := A_MyDocuments
}
defaultName := prefix "_" FormatTime(A_Now, "yyyyMMdd_HHmmss") ".png"

; --- Ask where to save. Overwrite is checked below, after the extension is normalized:
;     FileSelect only sees the name as typed (e.g. "a") and cannot know that "a.png" exists. ---
title := Chinese ? "保存截图" : "Save Screenshot"
filter := Chinese ? "PNG 图片 (*.png)" : "PNG Images (*.png)"
dest := FileSelect("S", startDir "\" defaultName, title, filter)
if (dest = "")                                      ; user cancelled
    ExitApp(0)

; Add .png when the user typed a name without an extension
if !RegExMatch(dest, "i)\.png$")
    dest .= ".png"

; The real path is known only now, so this is the one place to confirm an overwrite.
if FileExist(dest) {
    if Chinese
        question := "文件已存在: " dest "`n要覆盖吗?"
    else
        question := "File already exists: " dest "`nOverwrite?"
    answer := MsgBox(question, title, 4)             ; 4 = Yes/No
    if (answer = "No")
        ExitApp(0)
}

; --- Save the clipboard image to the chosen path (overwrite already confirmed above) ---
ok := false
try ok := SaveClipboardImage(dest)                  ; throws when the clipboard is busy or the file can't be written
if !ok
    Fail((Chinese ? "保存失败: " : "Failed to save: ") dest)

SplitPath(dest, &name)
Say((Chinese ? "已保存: " : "Saved: ") name)

;-------------------------------------------------------------------------------
; Does the current clipboard contain an image? CF_BITMAP is also reported when a program put only
; CF_DIB / CF_DIBV5 there (Windows converts between them).
;-------------------------------------------------------------------------------
ClipboardHasImage() {
    static CF_BITMAP := 2
    pngFmt := DllCall("RegisterClipboardFormat", "Str", "PNG", "UInt")
    return DllCall("IsClipboardFormatAvailable", "UInt", CF_BITMAP)
        || (pngFmt && DllCall("IsClipboardFormatAvailable", "UInt", pngFmt))
}

;-------------------------------------------------------------------------------
; Save the current clipboard image as a PNG, with no temporary file. A "PNG" format (browsers, Office)
; is written as it is, which keeps transparency; otherwise CF_BITMAP is encoded with GDI+.
;-------------------------------------------------------------------------------
SaveClipboardImage(file) {
    loop 10 {                                       ; another program may hold the clipboard for a moment
        if DllCall("OpenClipboard", "Ptr", 0, "Int")
            break
        if (A_Index = 10)
            throw Error("Cannot open the clipboard")
        Sleep(20)
    }
    try {
        if SavePngFormat(file)
            return true
        if (hBitmap := DllCall("GetClipboardData", "UInt", 2, "Ptr"))      ; CF_BITMAP
            return SaveBitmap(hBitmap, file)
        return false
    } finally {
        DllCall("CloseClipboard")
    }
}

; Writes the clipboard's "PNG" format to the file; false when there is none or it isn't PNG data
SavePngFormat(file) {
    pngFmt := DllCall("RegisterClipboardFormat", "Str", "PNG", "UInt")
    if !(pngFmt && (hPng := DllCall("GetClipboardData", "UInt", pngFmt, "Ptr")))
        return false
    if !(pPng := DllCall("GlobalLock", "Ptr", hPng, "Ptr"))
        return false
    try {
        size := DllCall("GlobalSize", "Ptr", hPng, "UPtr")
        if (size < 8 || NumGet(pPng, 0, "UInt") != 0x474E5089)            ; signature 89 50 4E 47 ("\x89PNG")
            return false
        f := FileOpen(file, "w")                    ; overwrite was confirmed by the caller
        f.RawWrite(pPng, size)
        f.Close()
        return true
    } finally {
        DllCall("GlobalUnlock", "Ptr", hPng)
    }
}

; Encodes an HBITMAP as PNG with GDI+. The HBITMAP belongs to the clipboard, so it isn't deleted here.
SaveBitmap(hBitmap, file) {
    clsid := Buffer(16, 0)                          ; PNG encoder, built into GDI+
    DllCall("ole32\CLSIDFromString", "Str", "{557CF406-1A04-11D3-9A73-0000F81EF32E}", "Ptr", clsid)
    DllCall("LoadLibrary", "Str", "gdiplus", "Ptr")
    input := Buffer(24, 0)                          ; GdiplusStartupInput: version 1
    NumPut("UInt", 1, input)
    token := 0, bitmap := 0
    if DllCall("gdiplus\GdiplusStartup", "Ptr*", &token, "Ptr", input, "Ptr", 0)
        return false
    try {
        if DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "Ptr", hBitmap, "Ptr", 0, "Ptr*", &bitmap)
            return false
        try FileDelete(file)                        ; overwrite was confirmed by the caller
        return DllCall("gdiplus\GdipSaveImageToFile", "Ptr", bitmap, "WStr", file, "Ptr", clsid, "Ptr", 0) = 0
    } finally {
        if bitmap
            DllCall("gdiplus\GdipDisposeImage", "Ptr", bitmap)
        DllCall("gdiplus\GdiplusShutdown", "Ptr", token)
    }
}
