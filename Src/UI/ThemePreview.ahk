;===============================================================================
; ThemePreview.ahk - theme thumbnails (the theme list in Preferences -> Appearance) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Draws a simplified mini search window in the theme's own colors and corner radius (no text, only color blocks): a stretch of input text color in the input box,
; a separator, three result rows (a colored icon + a title bar); the first row is selected, with the start of its title in the highlight color.
; No image files, so custom themes get thumbnails too;
; System (follow the system) is Light on the left half and Dark on the right.
;
; Usage:
;   hbm := ThemePreview.Bitmap(ThemeManager.Resolve("Dark"), 128, 88)   ; HBITMAP; DeleteObject when done
;   il  := ThemePreview.ImageList(["System", "Light", "Dark"], 128, 88)   ; in the order of the names; unknown themes use Light
;===============================================================================

class ThemePreview {
    static TitleWidths := [0.52, 0.40, 0.46]                                    ; lengths of the three titles (fraction of the width)
    static IconColors := ["3B82F6", "F59E0B", "10B981"]                        ; the "icons" on the left of the result rows (blue / orange / green, the same in every theme)

    ; Thumbnail of one theme (the complete Map of keys), w x h pixels. A 24-bit DIB (no alpha): the rounded corners are drawn by GDI+,
    ; which writes alpha = 255 on a 32-bit bitmap while the background GDI draws has alpha 0, and an image list seeing alpha treats the background as transparent
    static Bitmap(theme, w, h) {
        screen := DllCall("GetDC", "Ptr", 0, "Ptr")
        dc := DllCall("CreateCompatibleDC", "Ptr", screen, "Ptr")
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", screen)
        info := Buffer(40, 0)                                               ; BITMAPINFOHEADER
        NumPut("UInt", 40, "Int", w, "Int", -h, "UShort", 1, "UShort", 24, info)   ; top-down, 24-bit
        bits := 0
        hbm := DllCall("CreateDIBSection", "Ptr", dc, "Ptr", info, "UInt", 0, "Ptr*", &bits, "Ptr", 0, "UInt", 0, "Ptr")
        old := DllCall("SelectObject", "Ptr", dc, "Ptr", hbm, "Ptr")
        ThemePreview._Paint(dc, theme, w, h)
        DllCall("SelectObject", "Ptr", dc, "Ptr", old)
        DllCall("DeleteDC", "Ptr", dc)
        return hbm
    }

    ; theme name -> thumbnail; "System" is light on the left and dark on the right
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

    ; An image list with one thumbnail per theme (the ListView's large icons); freed with the ListView
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
            Win.FillRoundRect(dc, left, top, right + 1, bottom + 1, radius, bgr)   ; antialiased (GDI+)
        }
        bar(left, top, width, height, bgr) => roundRect(left, top, left + width, top + height, height // 2, bgr)   ; stands for a piece of text
        soft(key, backKey, t) {                                             ; the color moved towards the background by t (0~1): thin lines aren't as harsh
            a := color(key), b := color(backKey), mixed := 0
            Loop 3 {
                shift := (A_Index - 1) * 8
                mixed |= Round(((a >> shift) & 0xFF) * (1 - t) + ((b >> shift) & 0xFF) * t) << shift
            }
            return mixed
        }

        fill(0, 0, w, h, color("Background"))
        fill(0, 0, w, 1, color("Border")), fill(0, h - 1, w, h, color("Border"))
        fill(0, 0, 1, h, color("Border")), fill(w - 1, 0, w, h, color("Border"))

        pad := Max(5, Round(w * 0.08))
        inputH := Round(h * 0.26)
        inputBar := Max(2, Round(inputH * 0.14))
        bar(pad, (inputH - inputBar) // 2, Round(w * 0.22), inputBar, soft("InputText", "Background", 0.35))
        separatorY := inputH
        fill(0, separatorY, w, separatorY + 1, color("Separator"))

        rowTop := separatorY + 3, rowH := (h - rowTop - 3) / 3
        radius := Min(Round((theme.Has("SelectedRadius") ? theme["SelectedRadius"] : 0) * w / 260), Round(rowH / 3))
        titleH := Max(2, Round(rowH * 0.14))
        for index, titleW in ThemePreview.TitleWidths {
            top := Round(rowTop + (index - 1) * rowH), bottom := Round(rowTop + index * rowH)
            selected := (index = 1)
            if selected {
                inset := radius ? Max(2, Round(w * 0.02)) : 0
                roundRect(inset, top + (radius ? 1 : 0), w - 1 - inset, bottom - (radius ? 1 : 0), radius, color("SelectedBackground"))
            }
            back := selected ? "SelectedBackground" : "Background"
            iconSize := Round(rowH * 0.4), iconTop := top + Round((rowH - iconSize) / 2)
            roundRect(pad, iconTop, pad + iconSize, iconTop + iconSize, Max(1, iconSize // 4), Win.ColorToBgr(ThemePreview.IconColors[index]))
            x := pad + iconSize + Max(3, Round(w * 0.05))
            titleY := top + Round((rowH - titleH) / 2), titleLen := Round(w * titleW)
            if selected {                                                   ; selected row: the start is in the highlight color that matches the input
                matchW := Round(w * 0.12), gap := Max(1, Round(w * 0.012))
                bar(x, titleY, matchW, titleH, color("SelectedHighlight"))
                bar(x + matchW + gap, titleY, titleLen - matchW - gap, titleH, soft("SelectedTitle", back, 0.3))
            } else
                bar(x, titleY, titleLen, titleH, soft("Title", back, 0.45))
        }
    }
}
