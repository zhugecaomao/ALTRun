;===============================================================================
; Dialogs.ahk - 系统字体/颜色选择对话框 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 直接调用 Win32 ChooseFont / ChooseColor 通用对话框, 不依赖 Config/UI。
;
; 用法:
;   fontObj := FontDialog.Choose(fontObj, hwnd)      ; fontObj 见下方字段说明
;   color   := ColorDialog.Choose(color, hwnd, &customColors, "full")
;
; 命名注意: AutoHotkey 变量名不区分大小写, 所以这两个类的方法参数都刻意避免
; 叫 "font"/"color"/"FontDialog"/"ColorDialog" 这类会和类名/其他标识符撞车的
; 名字 —— Lib/Util.ahk 的 Path.Resolve() 就因为参数叫 path(和类名 Path 撞了)
; 出过一次真实的 bug, 这里引以为戒。
;===============================================================================

;===============================================================================
; ColorDialog - 颜色选择对话框 + RGB/BGR 颜色转换
;-------------------------------------------------------------------------------
; Win32 的颜色结构是 BGR 字节序, AHK/HTML 习惯用的是 RGB, 所以每次跟 Win32 API
; 交换颜色值都要过一遍 RgbBgr() 转一次通道顺序。
;===============================================================================
class ColorDialog {
    ; RGB <-> BGR 互转(交换红蓝通道), Win32 颜色结构都是 BGR 字节序
    static RgbBgr(rgbOrBgr) {
        rgbOrBgr := rgbOrBgr + 0
        return ((rgbOrBgr & 0xFF) << 16) | (rgbOrBgr & 0xFF00) | ((rgbOrBgr >> 16) & 0xFF)
    }

    ; 输出格式统一为 "0xRRGGBB"
    static Hex(rgbValue) {
        return Format("0x{:06X}", rgbValue & 0xFFFFFF)
    }

    ; 调用系统颜色选择对话框(ChooseColor)。
    ; initColor    : 初始颜色(RGB)。
    ; ownerHwnd    : 可选, 父窗口句柄。
    ; customColors : 可选, 自定义颜色数组/Map(1..16), 输入与输出共用。
    ; full         : 真值=完整面板(含自定义颜色), 假值=基础面板。
    ; 返回: -1 表示用户取消; 否则是用户选的颜色 "0xRRGGBB", 并更新 customColors。
    static Choose(initColor := 0, ownerHwnd := 0, &customColors := "", full := 1) {
        initColor := (initColor = "") ? 0 : initColor              ; 空值兜底为黑色
        panel := full ? 0x3 : 0x1                                  ; 0x3=完整面板, 0x1=基础面板
        bgrColor := ColorDialog.RgbBgr(initColor)                  ; 输入颜色转换, RGB -> BGR

        CUSTOM := Buffer(16 * 4, 0)                                 ; 自定义颜色缓冲区(16 色, 每色 4 字节)

        CHOOSECOLOR := Buffer(9 * A_PtrSize, 0)                     ; CHOOSECOLOR 结构缓冲区
        size := CHOOSECOLOR.size

        if (IsObject(customColors)) {
            Loop 16 {
                if (customColors.Has(A_Index)) {
                    col := customColors[A_Index] = "" ? 0 : customColors[A_Index]  ; 空值按黑色处理
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
        return ColorDialog.Hex(ColorDialog.RgbBgr(pickedBgr))       ; 输出颜色转换, BGR -> RGB
    }
}

;===============================================================================
; FontDialog - 字体选择对话框
;-------------------------------------------------------------------------------
; 调用系统字体选择对话框(ChooseFont)。
; spec (可选) : 字体初始值 Map, 支持 name/size/color/bold/italic/underline/strike。
;               只需要设置想预选的字段, 例如: Map("name","Terminal","size",14)。
; ownerHwnd   : 可选, 父窗口句柄, 传入后对话框为模态。
; withEffects : 真值=显示下划线/删除线等效果选项, 假值=不显示。
; 返回: false 表示用户取消; 否则返回 spec, 并补上以下字段:
;   ["str"] : 可直接传给 SetFont 的样式字符串。
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
        initBgr := ColorDialog.RgbBgr(initColorVal)                ; 输入颜色转换, RGB -> BGR

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
        spec["color"] := ColorDialog.Hex(ColorDialog.RgbBgr(pickedBgr))  ; 输出颜色转换, BGR -> RGB

        pickedSize := NumGet(choosefont, A_PtrSize = 8 ? 32 : 16, "UInt") / 10  ; iPointSize 的缩放值
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
