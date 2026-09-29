;===============================================================================
; ClipboardData.ahk - 剪贴板里的图片和文件 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; AutoHotkey 的 A_Clipboard 只处理文字; 这里补上图片 (存成 / 读回 PNG, 用 GDI+) 和文件列表 (CF_HDROP)。
;
; 用法:
;   ClipboardData.HasImage() / HasFiles()
;   size := ClipboardData.SaveImage("C:\x.png")      -> {Width, Height}, 没有图片或失败时 ""
;   ClipboardData.SetImage("C:\x.png")                 把 PNG 放到剪贴板 (CF_DIB), 成功返回 true
;   ClipboardData.Files()                               -> [路径...] (剪贴板里复制的文件)
;   ClipboardData.SetFiles(["C:\a.txt", "C:\b"])       把文件放到剪贴板 (在资源管理器里 Ctrl+V 就是粘贴文件)
;   ClipboardData.SetFiles(paths, true)                剪切: 粘贴时移动文件
;   ClipboardData.StartGdiplus()                       启动 GDI+ (只启动一次)
;===============================================================================

class ClipboardData {
    static CF_BITMAP := 2, CF_DIB := 8, CF_HDROP := 15
    static _token := 0

    static HasImage() => (DllCall("IsClipboardFormatAvailable", "UInt", ClipboardData.CF_DIB) || DllCall("IsClipboardFormatAvailable", "UInt", ClipboardData.CF_BITMAP)) ? true : false
    static HasFiles() => DllCall("IsClipboardFormatAvailable", "UInt", ClipboardData.CF_HDROP) ? true : false

    static Files() {
        files := []
        if !ClipboardData.HasFiles()
            return files
        for line in StrSplit(A_Clipboard, "`n", "`r")                     ; 复制文件时 A_Clipboard 是每行一个路径
            if (line != "")
                files.Push(line)
        return files
    }

    ; 先读 CF_DIB (设备无关位图, 别的程序放的图片也能读), 读不到再读 CF_BITMAP
    static SaveImage(file) {
        if !ClipboardData.HasImage() || !ClipboardData.StartGdiplus() || !ClipboardData._Open()
            return ""
        bitmap := 0
        try {
            bitmap := ClipboardData._FromDib()
            if (!bitmap && (hbm := DllCall("GetClipboardData", "UInt", ClipboardData.CF_BITMAP, "Ptr")))   ; 位图归剪贴板所有, 不要删除
                DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "Ptr", hbm, "Ptr", 0, "Ptr*", &bitmap)
        } finally {
            DllCall("CloseClipboard")
        }
        if !bitmap
            return ""
        width := 0, height := 0
        DllCall("gdiplus\GdipGetImageWidth", "Ptr", bitmap, "UInt*", &width)
        DllCall("gdiplus\GdipGetImageHeight", "Ptr", bitmap, "UInt*", &height)
        clsid := Buffer(16)
        DllCall("ole32\CLSIDFromString", "Str", "{557CF406-1A04-11D3-9A73-0000F81EF32E}", "Ptr", clsid)   ; PNG 编码器
        try FileDelete(file)
        status := DllCall("gdiplus\GdipSaveImageToFile", "Ptr", bitmap, "WStr", file, "Ptr", clsid, "Ptr", 0)
        DllCall("gdiplus\GdipDisposeImage", "Ptr", bitmap)
        return status ? "" : {Width: width, Height: height}
    }

    ; 剪贴板已打开时调用; 返回复制出来的 GDI+ 位图 (解锁之后剪贴板的内存就不能再用了), 失败时 0
    static _FromDib() {
        if !(hMem := DllCall("GetClipboardData", "UInt", ClipboardData.CF_DIB, "Ptr"))
            return 0
        if !(info := DllCall("GlobalLock", "Ptr", hMem, "Ptr"))
            return 0
        copy := 0
        try {
            headerSize := NumGet(info, 0, "UInt"), bitCount := NumGet(info, 14, "UShort")
            compression := NumGet(info, 16, "UInt"), colorsUsed := NumGet(info, 32, "UInt")
            colors := colorsUsed ? colorsUsed : (bitCount <= 8) ? (1 << bitCount) : 0
            masks := (compression = 3 && headerSize = 40) ? 12 : 0          ; BI_BITFIELDS: 头后面还有 3 个颜色掩码
            source := 0
            if (!DllCall("gdiplus\GdipCreateBitmapFromGdiDib", "Ptr", info, "Ptr", info + headerSize + masks + colors * 4, "Ptr*", &source) && source) {
                width := 0, height := 0
                DllCall("gdiplus\GdipGetImageWidth", "Ptr", source, "UInt*", &width)
                DllCall("gdiplus\GdipGetImageHeight", "Ptr", source, "UInt*", &height)
                DllCall("gdiplus\GdipCloneBitmapAreaI", "Int", 0, "Int", 0, "Int", width, "Int", height, "Int", 0x22009, "Ptr", source, "Ptr*", &copy)   ; 32bppRGB
                DllCall("gdiplus\GdipDisposeImage", "Ptr", source)
            }
        } finally {
            DllCall("GlobalUnlock", "Ptr", hMem)
        }
        return copy
    }

    ; 放进去的是 CF_DIB (所有程序都认), Windows 会自动提供 CF_BITMAP 等其他格式
    static SetImage(file) {
        if !FileExist(file) || !ClipboardData.StartGdiplus()
            return false
        bitmap := 0, hbm := 0
        if DllCall("gdiplus\GdipCreateBitmapFromFile", "WStr", file, "Ptr*", &bitmap) || !bitmap
            return false
        DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "Ptr", bitmap, "Ptr*", &hbm, "UInt", 0xFFFFFFFF)   ; 透明的地方填白色
        width := 0, height := 0
        DllCall("gdiplus\GdipGetImageWidth", "Ptr", bitmap, "UInt*", &width)
        DllCall("gdiplus\GdipGetImageHeight", "Ptr", bitmap, "UInt*", &height)
        DllCall("gdiplus\GdipDisposeImage", "Ptr", bitmap)
        if !hbm
            return false
        hMem := DllCall("GlobalAlloc", "UInt", 0x42, "UPtr", 40 + width * height * 4, "Ptr")
        if !hMem {
            DllCall("DeleteObject", "Ptr", hbm)
            return false
        }
        info := DllCall("GlobalLock", "Ptr", hMem, "Ptr")
        NumPut("UInt", 40, "Int", width, "Int", height, "UShort", 1, "UShort", 32, info)   ; BITMAPINFOHEADER, 从下往上, 32 位 BI_RGB
        hdc := DllCall("GetDC", "Ptr", 0, "Ptr")
        lines := DllCall("GetDIBits", "Ptr", hdc, "Ptr", hbm, "UInt", 0, "UInt", height, "Ptr", info + 40, "Ptr", info, "UInt", 0)
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdc)
        DllCall("GlobalUnlock", "Ptr", hMem)
        DllCall("DeleteObject", "Ptr", hbm)
        if (!lines || !ClipboardData._Open()) {
            DllCall("GlobalFree", "Ptr", hMem)
            return false
        }
        DllCall("EmptyClipboard")
        ok := DllCall("SetClipboardData", "UInt", ClipboardData.CF_DIB, "Ptr", hMem, "Ptr")   ; 成功后内存归剪贴板所有
        DllCall("CloseClipboard")
        if !ok
            DllCall("GlobalFree", "Ptr", hMem)
        return ok ? true : false
    }

    ; DROPFILES 结构 (20 字节) + 以两个空字符结尾的 UTF-16 路径列表;
    ; cut = true 时再放一个 "Preferred DropEffect" = 移动, 资源管理器 / TC 粘贴时就是移动文件
    static SetFiles(paths, cut := false) {
        size := 20 + 2
        for filePath in paths
            size += (StrLen(filePath) + 1) * 2
        hMem := DllCall("GlobalAlloc", "UInt", 0x42, "UPtr", size, "Ptr")  ; GMEM_MOVEABLE | GMEM_ZEROINIT
        if !hMem
            return false
        pointer := DllCall("GlobalLock", "Ptr", hMem, "Ptr")
        NumPut("UInt", 20, pointer)                                         ; pFiles: 路径列表的位置
        NumPut("Int", 1, pointer + 16)                                      ; fWide: UTF-16
        offset := 20
        for filePath in paths {
            StrPut(filePath, pointer + offset, "UTF-16")
            offset += (StrLen(filePath) + 1) * 2
        }
        DllCall("GlobalUnlock", "Ptr", hMem)
        if !ClipboardData._Open() {
            DllCall("GlobalFree", "Ptr", hMem)
            return false
        }
        DllCall("EmptyClipboard")
        ok := DllCall("SetClipboardData", "UInt", ClipboardData.CF_HDROP, "Ptr", hMem, "Ptr")
        if ok && (hEffect := DllCall("GlobalAlloc", "UInt", 0x42, "UPtr", 4, "Ptr")) {
            NumPut("UInt", cut ? 2 : 5, DllCall("GlobalLock", "Ptr", hEffect, "Ptr"))   ; DROPEFFECT_MOVE / COPY | LINK
            DllCall("GlobalUnlock", "Ptr", hEffect)
            if !DllCall("SetClipboardData", "UInt", ClipboardData.DropEffectFormat(), "Ptr", hEffect, "Ptr")
                DllCall("GlobalFree", "Ptr", hEffect)
        }
        DllCall("CloseClipboard")
        if !ok
            DllCall("GlobalFree", "Ptr", hMem)
        return ok ? true : false
    }

    static DropEffectFormat() {
        static format := DllCall("RegisterClipboardFormat", "Str", "Preferred DropEffect", "UInt")
        return format
    }

    ; 剪贴板里的文件是剪切的 (粘贴时移动) 时返回 true
    static IsCut() {
        if !DllCall("IsClipboardFormatAvailable", "UInt", ClipboardData.DropEffectFormat()) || !ClipboardData._Open()
            return false
        effect := 0
        if (hMem := DllCall("GetClipboardData", "UInt", ClipboardData.DropEffectFormat(), "Ptr")) && (pointer := DllCall("GlobalLock", "Ptr", hMem, "Ptr")) {
            effect := NumGet(pointer, "UInt")
            DllCall("GlobalUnlock", "Ptr", hMem)
        }
        DllCall("CloseClipboard")
        return (effect & 2) && !(effect & 1)
    }

    ; 别的程序正占用剪贴板时稍等再试
    static _Open() {
        Loop 10 {
            if DllCall("OpenClipboard", "Ptr", A_ScriptHwnd)
                return true
            Sleep(20)
        }
        return false
    }

    ; GDI+ 只需启动一次 (IconCache 画图片缩略图也用)
    static StartGdiplus() {
        if ClipboardData._token
            return true
        DllCall("LoadLibrary", "Str", "gdiplus")
        input := Buffer(24, 0)
        NumPut("UInt", 1, input)
        token := 0
        DllCall("gdiplus\GdiplusStartup", "Ptr*", &token, "Ptr", input, "Ptr", 0)
        ClipboardData._token := token
        return token ? true : false
    }
}
