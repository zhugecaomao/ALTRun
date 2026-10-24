;===============================================================================
; ClipboardData.ahk - images and files on the clipboard (AutoHotkey v2)
;-------------------------------------------------------------------------------
; AutoHotkey's A_Clipboard only handles text; this adds images (saved as / read back from PNG with GDI+) and file lists (CF_HDROP).
;
; Usage:
;   ClipboardData.HasImage() / HasFiles()
;   size := ClipboardData.SaveImage("C:\x.png")      -> {Width, Height}, "" when there is no image or it fails
;   ClipboardData.SetImage("C:\x.png")                 puts the PNG on the clipboard (CF_DIB), true on success
;   ClipboardData.Files()                               -> [paths...] (files copied to the clipboard)
;   ClipboardData.SetFiles(["C:\a.txt", "C:\b"])       puts files on the clipboard (Ctrl+V in Explorer pastes the files)
;   ClipboardData.SetFiles(paths, true)                cut: pasting moves the files
;   ClipboardData.StartGdiplus()                       starts GDI+ (only once)
;   ClipboardData.Borrow(ms)                           ALTRun puts temporary content on the clipboard for ms
;   ClipboardData.Borrowed()                           -> true until then (the clipboard history doesn't record it)
;===============================================================================

class ClipboardData {
    static CF_BITMAP := 2, CF_DIB := 8, CF_HDROP := 15
    static _token := 0
    static BorrowedUntil := 0                                               ; A_TickCount until which the clipboard is borrowed

    static Borrow(milliseconds) => ClipboardData.BorrowedUntil := A_TickCount + milliseconds
    static Borrowed() => A_TickCount < ClipboardData.BorrowedUntil

    static HasImage() => (DllCall("IsClipboardFormatAvailable", "UInt", ClipboardData.CF_DIB) || DllCall("IsClipboardFormatAvailable", "UInt", ClipboardData.CF_BITMAP)) ? true : false
    static HasFiles() => DllCall("IsClipboardFormatAvailable", "UInt", ClipboardData.CF_HDROP) ? true : false

    static Files() {
        files := []
        if !ClipboardData.HasFiles()
            return files
        for line in StrSplit(A_Clipboard, "`n", "`r")                     ; for copied files A_Clipboard has one path per line
            if (line != "")
                files.Push(line)
        return files
    }

    ; Reads CF_DIB first (device-independent bitmap, so images put there by other programs work too), then CF_BITMAP.
    ; The clipboard is open only while the raw data is copied out; converting and saving the PNG happen after it is closed, so other programs don't wait
    static SaveImage(file) {
        if !ClipboardData.HasImage() || !ClipboardData.StartGdiplus() || !ClipboardData._Open()
            return ""
        bitmap := 0, dib := ""
        try {
            dib := ClipboardData._CopyDib()
            if (!dib && (hbm := DllCall("GetClipboardData", "UInt", ClipboardData.CF_BITMAP, "Ptr")))   ; the bitmap belongs to the clipboard: don't delete it
                DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "Ptr", hbm, "Ptr", 0, "Ptr*", &bitmap)
        } finally {
            DllCall("CloseClipboard")
        }
        if IsObject(dib)
            bitmap := ClipboardData._FromDib(dib)
        if !bitmap
            return ""
        width := 0, height := 0
        DllCall("gdiplus\GdipGetImageWidth", "Ptr", bitmap, "UInt*", &width)
        DllCall("gdiplus\GdipGetImageHeight", "Ptr", bitmap, "UInt*", &height)
        clsid := Buffer(16)
        DllCall("ole32\CLSIDFromString", "Str", "{557CF406-1A04-11D3-9A73-0000F81EF32E}", "Ptr", clsid)   ; PNG encoder
        try FileDelete(file)
        status := DllCall("gdiplus\GdipSaveImageToFile", "Ptr", bitmap, "WStr", file, "Ptr", clsid, "Ptr", 0)
        DllCall("gdiplus\GdipDisposeImage", "Ptr", bitmap)
        return status ? "" : {Width: width, Height: height}
    }

    ; Called with the clipboard open: copies the CF_DIB contents into a Buffer (that memory can't be used after the clipboard closes), "" if there is none
    static _CopyDib() {
        if !(hMem := DllCall("GetClipboardData", "UInt", ClipboardData.CF_DIB, "Ptr"))
            return ""
        size := DllCall("GlobalSize", "Ptr", hMem, "UPtr")
        if (size < 40 || !(source := DllCall("GlobalLock", "Ptr", hMem, "Ptr")))
            return ""
        dib := Buffer(size)
        DllCall("RtlMoveMemory", "Ptr", dib, "Ptr", source, "UPtr", size)
        DllCall("GlobalUnlock", "Ptr", hMem)
        return dib
    }

    ; CF_DIB contents (BITMAPINFO + pixels) -> a GDI+ bitmap, 0 on failure
    static _FromDib(dib) {
        info := dib.Ptr
        copy := 0
        try {
            headerSize := NumGet(info, 0, "UInt"), bitCount := NumGet(info, 14, "UShort")
            compression := NumGet(info, 16, "UInt"), colorsUsed := NumGet(info, 32, "UInt")
            colors := colorsUsed ? colorsUsed : (bitCount <= 8) ? (1 << bitCount) : 0
            masks := (compression = 3 && headerSize = 40) ? 12 : 0          ; BI_BITFIELDS: three color masks follow the header
            source := 0
            if (!DllCall("gdiplus\GdipCreateBitmapFromGdiDib", "Ptr", info, "Ptr", info + headerSize + masks + colors * 4, "Ptr*", &source) && source) {
                width := 0, height := 0
                DllCall("gdiplus\GdipGetImageWidth", "Ptr", source, "UInt*", &width)
                DllCall("gdiplus\GdipGetImageHeight", "Ptr", source, "UInt*", &height)
                DllCall("gdiplus\GdipCloneBitmapAreaI", "Int", 0, "Int", 0, "Int", width, "Int", height, "Int", 0x22009, "Ptr", source, "Ptr*", &copy)   ; 32bppRGB
                DllCall("gdiplus\GdipDisposeImage", "Ptr", source)
            }
        }
        return copy
    }

    ; CF_DIB goes in (every program understands it); Windows provides CF_BITMAP and the other formats itself
    static SetImage(file) {
        if !FileExist(file) || !ClipboardData.StartGdiplus()
            return false
        bitmap := 0, hbm := 0
        if DllCall("gdiplus\GdipCreateBitmapFromFile", "WStr", file, "Ptr*", &bitmap) || !bitmap
            return false
        DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "Ptr", bitmap, "Ptr*", &hbm, "UInt", 0xFFFFFFFF)   ; transparent areas become white
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
        NumPut("UInt", 40, "Int", width, "Int", height, "UShort", 1, "UShort", 32, info)   ; BITMAPINFOHEADER, bottom-up, 32-bit BI_RGB
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
        ok := DllCall("SetClipboardData", "UInt", ClipboardData.CF_DIB, "Ptr", hMem, "Ptr")   ; on success the memory belongs to the clipboard
        DllCall("CloseClipboard")
        if !ok
            DllCall("GlobalFree", "Ptr", hMem)
        return ok ? true : false
    }

    ; DROPFILES structure (20 bytes) + a UTF-16 path list ending in two null characters;
    ; cut = true also sets "Preferred DropEffect" = move, so pasting in Explorer / TC moves the files
    static SetFiles(paths, cut := false) {
        size := 20 + 2
        for filePath in paths
            size += (StrLen(filePath) + 1) * 2
        hMem := DllCall("GlobalAlloc", "UInt", 0x42, "UPtr", size, "Ptr")  ; GMEM_MOVEABLE | GMEM_ZEROINIT
        if !hMem
            return false
        pointer := DllCall("GlobalLock", "Ptr", hMem, "Ptr")
        NumPut("UInt", 20, pointer)                                         ; pFiles: where the path list starts
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

    ; true when the files on the clipboard were cut (pasting moves them)
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

    ; another program holds the clipboard: wait a moment and retry
    static _Open() {
        Loop 10 {
            if DllCall("OpenClipboard", "Ptr", A_ScriptHwnd)
                return true
            Sleep(20)
        }
        return false
    }

    ; GDI+ needs starting only once (IconCache uses it for image thumbnails too)
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
