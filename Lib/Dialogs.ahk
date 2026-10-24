;===============================================================================
; Dialogs.ahk - the system font / color dialogs (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Calls the Win32 ChooseFont / ChooseColor common dialogs directly; doesn't depend on the settings or the UI.
;
; Usage:
;   fontObj := FontDialog.Choose(fontObj, hwnd)      ; fontObj: see the fields below
;   color   := ColorDialog.Choose(color, hwnd, &customColors, "full")
;
; Naming: AutoHotkey names are case-insensitive, so the parameters of these two classes deliberately avoid
; names like "font" / "color" / "FontDialog" / "ColorDialog" that would collide with class names or other identifiers.
; Path.Resolve() in Lib/Util.ahk once had a real bug because its parameter was called path (the same as the class Path);
; this is the lesson from it.
;===============================================================================

;===============================================================================
; ColorDialog - color dialog + RGB / BGR conversion
;-------------------------------------------------------------------------------
; Win32 color structures are in BGR byte order while AHK / HTML use RGB, so every color value exchanged with
; the Win32 API goes through RgbBgr() to swap the channels.
;===============================================================================
class ColorDialog {
    ; RGB <-> BGR (swaps the red and blue channels); Win32 color structures are all in BGR byte order
    static RgbBgr(rgbOrBgr) {
        rgbOrBgr := rgbOrBgr + 0
        return ((rgbOrBgr & 0xFF) << 16) | (rgbOrBgr & 0xFF00) | ((rgbOrBgr >> 16) & 0xFF)
    }

    ; output is always "0xRRGGBB"
    static Hex(rgbValue) {
        return Format("0x{:06X}", rgbValue & 0xFFFFFF)
    }

    ; Shows the system color dialog (ChooseColor).
    ; initColor    : initial color (RGB).
    ; ownerHwnd    : optional, owner window handle.
    ; customColors : optional array / Map (1..16) of custom colors, used for input and output.
    ; full         : true = full panel (with custom colors), false = basic panel.
    ; Returns -1 when the user cancels; otherwise the chosen color "0xRRGGBB", and customColors is updated.
    static Choose(initColor := 0, ownerHwnd := 0, &customColors := "", full := 1) {
        initColor := (initColor = "") ? 0 : initColor              ; empty means black
        panel := full ? 0x3 : 0x1                                  ; 0x3 = full panel, 0x1 = basic panel
        bgrColor := ColorDialog.RgbBgr(initColor)                  ; convert the input color, RGB -> BGR

        CUSTOM := Buffer(16 * 4, 0)                                 ; custom color buffer (16 colors, 4 bytes each)

        CHOOSECOLOR := Buffer(9 * A_PtrSize, 0)                     ; CHOOSECOLOR structure buffer
        size := CHOOSECOLOR.size

        if (IsObject(customColors)) {
            Loop 16 {
                if (customColors.Has(A_Index)) {
                    col := customColors[A_Index] = "" ? 0 : customColors[A_Index]  ; empty is treated as black
                    custBgr := ColorDialog.RgbBgr(col)
                    NumPut "UInt", custBgr, CUSTOM, ((A_Index - 1) * 4)
                }
            }
        }

        NumPut "UInt", size, CHOOSECOLOR, 0
        NumPut "UPtr", ownerHwnd, CHOOSECOLOR, A_PtrSize
        NumPut "UInt", bgrColor, CHOOSECOLOR, 3 * A_PtrSize
        NumPut "UInt", panel, CHOOSECOLOR, 5 * A_PtrSize
        NumPut "UPtr", CUSTOM.ptr, CHOOSECOLOR, 4 * A_PtrSize

        ret := DllCall("comdlg32\ChooseColor", "UPtr", CHOOSECOLOR.ptr, "UInt")

        if !ret
            return -1

        customColors := Array()
        Loop 16 {
            newCustBgr := NumGet(CUSTOM, (A_Index - 1) * 4, "UInt")
            customColors.InsertAt(A_Index, ColorDialog.Hex(ColorDialog.RgbBgr(newCustBgr)))
        }

        pickedBgr := NumGet(CHOOSECOLOR, 3 * A_PtrSize, "UInt")
        return ColorDialog.Hex(ColorDialog.RgbBgr(pickedBgr))       ; convert the output color, BGR -> RGB
    }
}

;===============================================================================
; FontDialog - font dialog
;-------------------------------------------------------------------------------
; Shows the system font dialog (ChooseFont).
; spec (optional) : initial font Map with name / size / color / bold / italic / underline / strike.
;                   Only the fields to preselect are needed, e.g. Map("name", "Terminal", "size", 14).
; ownerHwnd   : optional owner window handle; with it the dialog is modal.
; withEffects : true = show the underline / strikeout effect options, false = don't.
; Returns false when the user cancels; otherwise spec with these fields added:
;   ["str"] : an options string that can be passed to SetFont directly.
;===============================================================================
class FontDialog {
    static Choose(spec := "", ownerHwnd := 0, withEffects := 1) {
        spec := (spec = "") ? Map() : spec
        logfont := Buffer((A_PtrSize = 4) ? 60 : 92, 0)
        dc := DllCall("GetDC", "Ptr", 0, "Ptr")
        dpi := DllCall("GetDeviceCaps", "Ptr", dc, "UInt", 90, "Int")
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", dc)
        effectFlags := 0x041 + (withEffects ? 0x100 : 0)

        initName := spec.Has("name") ? spec["name"] : ""
        initBold := spec.Has("bold") ? spec["bold"] : 0
        initBold := initBold ? 700 : 400
        initItalic := spec.Has("italic") ? spec["italic"] : 0
        initUnderline := spec.Has("underline") ? spec["underline"] : 0
        initStrikeout := spec.Has("strike") ? spec["strike"] : 0
        initSize := spec.Has("size") ? spec["size"] : 10
        initSize := initSize ? Floor(initSize * dpi / 72) : 16
        initColorVal := spec.Has("color") ? spec["color"] : 0
        initBgr := ColorDialog.RgbBgr(initColorVal)                ; convert the input color, RGB -> BGR

        NumPut "UInt", initSize, logfont
        NumPut "UInt", initBold, "UChar", initItalic, "UChar", initUnderline, "UChar", initStrikeout, logfont, 16

        choosefont := Buffer(A_PtrSize = 8 ? 104 : 60, 0), cap := choosefont.size
        NumPut "UInt", cap, choosefont, 0
        NumPut "UPtr", ownerHwnd, choosefont, A_PtrSize
        offset1 := (A_PtrSize = 8) ? 24 : 12
        offset2 := (A_PtrSize = 8) ? 36 : 20
        offset3 := (A_PtrSize = 4) ? 6 * A_PtrSize : 5 * A_PtrSize
        NumPut "UPtr", logfont.ptr, choosefont, offset1
        NumPut "UInt", effectFlags, choosefont, offset2
        NumPut "UInt", initBgr, choosefont, offset3

        StrPut(initName, logfont.ptr + 28, "UTF-16")
        ok := DllCall("comdlg32\ChooseFont", "UPtr", choosefont.ptr)
        pickedName := StrGet(logfont.ptr + 28, "UTF-16")

        if !ok
            return false

        spec["bold"] := NumGet(logfont, 16, "UChar")
        spec["italic"] := NumGet(logfont, 20, "UChar")
        spec["underline"] := NumGet(logfont, 21, "UChar")
        spec["strike"] := NumGet(logfont, 22, "UChar")

        spec["bold"] := (spec["bold"] < 188) ? 0 : 1

        pickedBgr := NumGet(choosefont, (A_PtrSize = 4) ? 6 * A_PtrSize : 5 * A_PtrSize, "UInt")
        spec["color"] := ColorDialog.Hex(ColorDialog.RgbBgr(pickedBgr))  ; convert the output color, BGR -> RGB

        pickedSize := NumGet(choosefont, A_PtrSize = 8 ? 32 : 16, "UInt") / 10  ; iPointSize is in tenths of a point
        spec["size"] := pickedSize
        spec["name"] := pickedName

        str := "norm"
        if (spec["bold"])
            str .= " bold"
        if (spec["italic"])
            str .= " italic"
        if (spec["strike"])
            str .= " strike"
        if (spec["color"])
            str .= " c" spec["color"]
        if (spec["size"])
            str .= " s" spec["size"]
        if (spec["underline"])
            str .= " underline"
        spec["str"] := str
        return spec
    }
}
