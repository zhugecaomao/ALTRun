;===============================================================================
; ThemePreview.ahk - 主题缩略图 (偏好设置 -> 外观 的主题列表) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 按主题自己的颜色和圆角画一个简化的迷你搜索窗口 (不写字, 只用色块): 输入框里一段输入文字色,
; 分隔线, 三行结果 (图标 + 一条标题); 第一行是选中行, 标题前一段是高亮色。
; 不用图片文件, 自定义主题也有缩略图;
; System (跟随系统) 左半边是 Light, 右半边是 Dark。
;
; 用法:
;   hbm := ThemePreview.Bitmap(ThemeManager.Resolve("Dark"), 128, 88)   ; HBITMAP, 用完 DeleteObject
;   il  := ThemePreview.ImageList(["System", "Light", "Dark"], 128, 88)   ; 顺序和名称一样, 找不到的主题用 Light
;===============================================================================

class ThemePreview {
    static TitleWidths := [0.52, 0.40, 0.46]                                    ; 三行标题的长度 (占宽度的比例)

    ; 一个主题 (完整的键值 Map) 的缩略图, w x h 像素
    static Bitmap(theme, w, h) {
        screen := DllCall("GetDC", "Ptr", 0, "Ptr")
        dc := DllCall("CreateCompatibleDC", "Ptr", screen, "Ptr")
        hbm := DllCall("CreateCompatibleBitmap", "Ptr", screen, "Int", w, "Int", h, "Ptr")
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", screen)
        old := DllCall("SelectObject", "Ptr", dc, "Ptr", hbm, "Ptr")
        ThemePreview._Paint(dc, theme, w, h)
        DllCall("SelectObject", "Ptr", dc, "Ptr", old)
        DllCall("DeleteDC", "Ptr", dc)
        return hbm
    }

    ; 主题名 -> 缩略图; "System" 拼成左浅右深
    static ForName(themeName, w, h) {
        if (themeName != "System") {
            theme := ThemeManager.Resolve(themeName)
            return ThemePreview.Bitmap(IsObject(theme) ? theme : ThemeManager.Defaults(), w, h)
        }
        light := ThemePreview.Bitmap(ThemeManager.Defaults(), w, h)
        dark := ThemeManager.Resolve("Dark")
        if !IsObject(dark)
            return light
        darkBitmap := ThemePreview.Bitmap(dark, w, h)
        screen := DllCall("GetDC", "Ptr", 0, "Ptr")
        target := DllCall("CreateCompatibleDC", "Ptr", screen, "Ptr"), source := DllCall("CreateCompatibleDC", "Ptr", screen, "Ptr")
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", screen)
        oldTarget := DllCall("SelectObject", "Ptr", target, "Ptr", light, "Ptr")
        oldSource := DllCall("SelectObject", "Ptr", source, "Ptr", darkBitmap, "Ptr")
        half := w // 2
        DllCall("BitBlt", "Ptr", target, "Int", half, "Int", 0, "Int", w - half, "Int", h, "Ptr", source, "Int", half, "Int", 0, "UInt", 0x00CC0020)   ; SRCCOPY
        DllCall("SelectObject", "Ptr", target, "Ptr", oldTarget)
        DllCall("SelectObject", "Ptr", source, "Ptr", oldSource)
        DllCall("DeleteDC", "Ptr", target), DllCall("DeleteDC", "Ptr", source)
        DllCall("DeleteObject", "Ptr", darkBitmap)
        return light
    }

    ; 每个主题一张缩略图的图像列表 (ListView 的大图标); ListView 销毁时一起释放
    static ImageList(names, w, h) {
        il := DllCall("comctl32\ImageList_Create", "Int", w, "Int", h, "UInt", 0x20, "Int", names.Length, "Int", 4, "Ptr")   ; ILC_COLOR32
        for themeName in names
            ThemePreview.AddTo(il, themeName, w, h)
        return il
    }

    static AddTo(il, themeName, w, h) {
        hbm := ThemePreview.ForName(themeName, w, h)
        index := DllCall("comctl32\ImageList_Add", "Ptr", il, "Ptr", hbm, "Ptr", 0, "Int")
        DllCall("DeleteObject", "Ptr", hbm)
        return index
    }

    static _Paint(dc, theme, w, h) {
        color := (key) => Win.ColorToBgr(theme.Has(key) ? theme[key] : ThemeManager.Defaults()[key])
        fill(left, top, right, bottom, bgr) {
            rect := Buffer(16)
            NumPut("Int", left, "Int", top, "Int", right, "Int", bottom, rect)
            brush := DllCall("CreateSolidBrush", "UInt", bgr, "Ptr")
            DllCall("FillRect", "Ptr", dc, "Ptr", rect, "Ptr", brush)
            DllCall("DeleteObject", "Ptr", brush)
        }
        roundRect(left, top, right, bottom, radius, bgr) {
            if (radius <= 0)
                return fill(left, top, right, bottom, bgr)
            brush := DllCall("CreateSolidBrush", "UInt", bgr, "Ptr")
            oldBrush := DllCall("SelectObject", "Ptr", dc, "Ptr", brush, "Ptr")
            oldPen := DllCall("SelectObject", "Ptr", dc, "Ptr", DllCall("GetStockObject", "Int", 8, "Ptr"), "Ptr")   ; NULL_PEN
            DllCall("RoundRect", "Ptr", dc, "Int", left, "Int", top, "Int", right + 1, "Int", bottom + 1, "Int", radius * 2, "Int", radius * 2)
            DllCall("SelectObject", "Ptr", dc, "Ptr", oldPen)
            DllCall("SelectObject", "Ptr", dc, "Ptr", oldBrush)
            DllCall("DeleteObject", "Ptr", brush)
        }
        bar(left, top, width, height, bgr) => roundRect(left, top, left + width, top + height, height // 2, bgr)   ; 代表一段文字

        fill(0, 0, w, h, color("Background"))
        fill(0, 0, w, 1, color("Border")), fill(0, h - 1, w, h, color("Border"))
        fill(0, 0, 1, h, color("Border")), fill(w - 1, 0, w, h, color("Border"))

        pad := Max(5, Round(w * 0.08))
        inputH := Round(h * 0.26)
        inputBar := Max(3, Round(inputH * 0.22))
        bar(pad, (inputH - inputBar) // 2, Round(w * 0.24), inputBar, color("InputText"))
        separatorY := inputH
        fill(0, separatorY, w, separatorY + 1, color("Separator"))

        rowTop := separatorY + 3, rowH := (h - rowTop - 3) / 3
        radius := Min(Round((theme.Has("SelectedRadius") ? theme["SelectedRadius"] : 0) * w / 260), Round(rowH / 3))
        titleH := Max(3, Round(rowH * 0.22))
        for index, titleW in ThemePreview.TitleWidths {
            top := Round(rowTop + (index - 1) * rowH), bottom := Round(rowTop + index * rowH)
            selected := (index = 1)
            if selected {
                inset := radius ? Max(2, Round(w * 0.02)) : 0
                roundRect(inset, top + (radius ? 1 : 0), w - 1 - inset, bottom - (radius ? 1 : 0), radius, color("SelectedBackground"))
            }
            iconSize := Round(rowH * 0.42), iconTop := top + Round((rowH - iconSize) / 2)
            roundRect(pad, iconTop, pad + iconSize, iconTop + iconSize, Max(1, iconSize // 4), color(selected ? "SelectedSubtitle" : "Subtitle"))
            x := pad + iconSize + Max(3, Round(w * 0.05))
            titleY := top + Round((rowH - titleH) / 2), titleLen := Round(w * titleW)
            if selected {                                                   ; 选中行: 前一段是和输入匹配的高亮色
                matchW := Round(w * 0.12), gap := Max(1, Round(w * 0.012))
                bar(x, titleY, matchW, titleH, color("SelectedHighlight"))
                bar(x + matchW + gap, titleY, titleLen - matchW - gap, titleH, color("SelectedTitle"))
            } else
                bar(x, titleY, titleLen, titleH, color("Title"))
        }
    }
}
