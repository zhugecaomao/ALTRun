;===============================================================================
; Util.ahk - general helper functions
;-------------------------------------------------------------------------------
; Grouped into small classes by purpose; all pure functions that don't depend on the settings or the UI, so they are easy to test on their own.
;===============================================================================


;===============================================================================
; SafeFile - writing a file so that a crash or power cut halfway can't damage it
;-------------------------------------------------------------------------------
; Writes a temporary file in the same folder, then replaces the original. Creates the folder if needed; UTF-8 without BOM.
; Sync folders such as OneDrive and antivirus software hold a just-changed file for a moment: on failure retry every 100 ms,
; up to MoveTries times, then throw (the original is kept, the temporary file deleted)
;===============================================================================
class SafeFile {
    static MoveTries := 20

    static WriteText(file, text) {
        SplitPath(file, , &dir)
        if (dir != "")
            DirCreate(dir)
        tmpFile := file ".tmp"
        try FileDelete(tmpFile)
        if FileExist(tmpFile)                                               ; a temporary file left over from last time can't be deleted (in use): use another name
            tmpFile := file "." A_TickCount ".tmp"
        FileAppend(text, tmpFile, "UTF-8-RAW")
        Loop SafeFile.MoveTries {
            try {
                FileMove(tmpFile, file, true)
                return
            } catch as e {
                if (A_Index = SafeFile.MoveTries) {
                    try FileDelete(tmpFile)
                    throw e
                }
                Sleep(100)
            }
        }
    }
}


;===============================================================================
; Path - path resolution
;-------------------------------------------------------------------------------
; Three forms are supported:
;   A_ScriptDir\x.txt   starts with an AHK built-in variable
;   %Temp%\x.txt        environment-variable placeholder
;   notepad.exe         bare program name (looked up in PATH)
;===============================================================================
class Path {
    static _varCache := ""

    ; A UNC path, or a drive letter Windows calls a network drive (GetDriveType doesn't touch the network; each letter
    ; is asked once)
    static IsRemote(target) {
        static drives := Map()
        if (SubStr(target, 1, 2) = "\\")
            return true
        if !RegExMatch(target, "^([A-Za-z]):", &match)
            return false
        letter := StrUpper(match[1])
        if !drives.Has(letter)
            drives[letter] := DllCall("GetDriveTypeW", "WStr", letter ":\", "UInt") = 4      ; DRIVE_REMOTE
        return drives[letter]
    }

    ; Only what looks like an extension counts (1-6 letters / digits, at least one letter: pdf, docx, dwg, sldprt):
    ; folders with a dot in the name such as "26. Main Street (A)", "10.Project-2310", "Design.2019" have none
    static Extension(path) {
        SplitPath(RTrim(path, "\/"), , , &ext)
        return (RegExMatch(ext, "^(?=.*[A-Za-z])[A-Za-z0-9]{1,6}$") || ext = "appref-ms") ? ext : ""
    }

    ; The built-in variable table is filled on first use (reading it at start-up slows a cold start)
    static _Vars() {
        if (Path._varCache != "")
            return Path._varCache
        Path._varCache := Map(
            "A_ScriptDir"      , A_ScriptDir,
            "A_Temp"           , A_Temp,
            "A_Startup"        , A_Startup,
            "A_StartMenu"      , A_StartMenu,
            "A_Programs"       , A_Programs,
            "A_AppData"        , A_AppData,
            "A_Desktop"        , A_Desktop,
            "A_DesktopCommon"  , A_DesktopCommon,
            "A_MyDocuments"    , A_MyDocuments,
            "A_ProgramFiles"   , A_ProgramFiles,
            "A_ProgramsCommon" , A_ProgramsCommon,
            "A_StartupCommon"  , A_StartupCommon,
            "A_StartMenuCommon", A_StartMenuCommon,
            "A_WinDir"         , A_WinDir
        )
        return Path._varCache
    }

    ; A folder path the user wrote -> full path: expands %environment variables% and built-in variables starting with A_; relative paths start from base
    static Full(raw, base := A_ScriptDir) {
        raw := Trim(raw)
        if (InStr(raw, "A_") = 1)
            raw := Path.Resolve(raw)
        size := DllCall("ExpandEnvironmentStringsW", "WStr", raw, "Ptr", 0, "UInt", 0, "UInt")
        buf := Buffer(size * 2)
        DllCall("ExpandEnvironmentStringsW", "WStr", raw, "Ptr", buf, "UInt", size, "UInt")
        raw := StrGet(buf)
        if !RegExMatch(raw, "^([A-Za-z]:|\\\\)")
            raw := RTrim(base, "\") "\" raw
        buf := Buffer(32767 * 2)
        if DllCall("GetFullPathNameW", "WStr", raw, "UInt", 32767, "Ptr", buf, "Ptr", 0, "UInt")
            raw := StrGet(buf)
        return RTrim(raw, "\")
    }

    ; Whether files can be written in this folder (tries to create a temporary file and delete it)
    static IsWritable(dir) {
        probe := dir "\ALTRun.write-test.tmp"
        try {
            FileAppend("", probe)
            FileDelete(probe)
            return true
        }
        return false
    }

    ; To an absolute path. keepRunAs = true keeps the "*RunAs " prefix (used when actually running it)
    ; Note: the parameter must not be called path. AHK names are case-insensitive, so "path" and the class name "Path"
    ; are the same identifier, and every Path.xxx self-reference in the method body would read the parameter instead.
    static Resolve(raw, keepRunAs := false) {
        raw := Trim(raw)
        if (raw = "")
            return ""

        if (!keepRunAs)
            raw := StrReplace(raw, "*RunAs ", "")

        vars := Path._Vars()

        ; Like "A_ScriptDir\sub\x.exe": it must start with A_, otherwise something like "Plot A_IGLS" would be mistaken for one
        if (InStr(raw, "A_") = 1) {
            if (InStr(raw, "\")) {
                parts := StrSplit(raw, "\",, 2)
                if (vars.Has(parts[1]))
                    raw := vars[parts[1]] "\" (parts.Has(2) ? parts[2] : "")
            } else if (vars.Has(raw)) {
                raw := vars[raw]
            }
        }

        ; environment-variable placeholder
        raw := StrReplace(raw, "%Temp%"    , A_Temp)
        raw := StrReplace(raw, "%OneDrive%", EnvGet("OneDrive"))
        raw := StrReplace(raw, "%AppData%" , A_AppData)
        raw := StrReplace(raw, "%UserProfile%", EnvGet("UserProfile"))

        ; Bare file name: search PATH; without an extension look for .exe ("notepad" -> notepad.exe)
        ; Check for "\" first: a full path doesn't touch the disk (FileExist hangs when a network drive is disconnected)
        if (!InStr(raw, "\") && !FileExist(raw)) {
            buf := Buffer(260 * 2)
            if DllCall("kernel32\SearchPathW", "Ptr", 0, "WStr", raw, "WStr", ".exe",
                       "UInt", buf.Size // 2, "Ptr", buf, "Ptr", 0)
                raw := StrGet(buf, "UTF-16")
        }
        return raw
    }

    ; The reverse: common folders back to placeholders, so the settings file works on other machines
    static Shorten(path) {
        path := StrReplace(path, A_Temp, "%Temp%")
        if (od := EnvGet("OneDrive"))
            path := StrReplace(path, od, "%OneDrive%")
        return path
    }

    ; Just the file name, for lists
    ; Opens a text file with the default program for its type (e.g. .log set to open with Notepad4), or Notepad if that fails.
    ; Only for types whose default action is viewing (.log .txt): scripts (.ahk .bat) run by default, .json often has no association.
    ; runner is passed only by tests (instead of Run)
    static OpenText(file, runner := "") {
        runner := IsObject(runner) ? runner : Run
        try
            runner('"' file '"')
        catch
            runner('notepad.exe "' file '"')
    }

    static Leaf(path) {
        SplitPath(path, &name)
        return name != "" ? name : path
    }
}


;===============================================================================
; Fonts - font description string <-> AHK SetFont options
; The settings store a human-readable form such as "Microsoft YaHei, norm s10"
;===============================================================================
class Fonts {
    static Spec(text, defName := "Microsoft YaHei", defOpt := "norm s10") {
        parts := StrSplit(text, ",")
        name  := parts.Length >= 1 ? Trim(parts[1]) : ""
        opt   := parts.Length >= 2 ? Trim(parts[2]) : ""
        return { name: name != "" ? name : defName, opt: opt != "" ? opt : defOpt }
    }

    ; Shows the system font dialog; returns {name, opt} or false
    static Choose(initName := "", ownerHwnd := 0) {
        lf := Buffer((A_PtrSize = 4) ? 60 : 92, 0)
        dc := DllCall("GetDC", "Ptr", 0, "Ptr")
        dpi := DllCall("GetDeviceCaps", "Ptr", dc, "UInt", 90, "Int")
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", dc)

        NumPut("UInt", Floor(10 * dpi / 72), lf)            ; 10 pt by default
        NumPut("UInt", 400, lf, 16)
        StrPut(initName, lf.Ptr + 28, "UTF-16")

        cf  := Buffer(A_PtrSize = 8 ? 104 : 60, 0)
        NumPut("UInt", cf.Size, cf, 0)
        NumPut("UPtr", ownerHwnd, cf, A_PtrSize)
        NumPut("UPtr", lf.Ptr, cf, (A_PtrSize = 8) ? 24 : 12)
        NumPut("UInt", 0x141, cf, (A_PtrSize = 8) ? 36 : 20)  ; SCREENFONTS|EFFECTS|INITTOLOGFONT

        if !DllCall("comdlg32\ChooseFont", "UPtr", cf.Ptr)
            return false

        name := StrGet(lf.Ptr + 28, "UTF-16")
        size := NumGet(cf, A_PtrSize = 8 ? 32 : 16, "UInt") / 10
        bold := NumGet(lf, 16, "UInt") >= 600
        ital := NumGet(lf, 20, "UChar")

        opt := "norm" (bold ? " bold" : "") (ital ? " italic" : "") " s" size
        return { name: name, opt: opt }
    }
}


;===============================================================================
; Win - small window helpers
;===============================================================================
class Win {
    ; Windows 11 rounded corners
    static SetCorner(hwnd, rounded := true) {
        static ATTR := 33, ROUND := 2, SQUARE := 1
        try DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "UInt", ATTR,
                    "Int*", rounded ? ROUND : SQUARE, "UInt", 4)
    }

    ; Width of one line of text in a control (pixels, in the control's own font)
    ; The real height of one ListView row (pixels): measured by adding a row for a moment (LVM_GETITEMRECT); fallback if that fails
    static ListRowHeight(list, fallback) {
        added := !list.GetCount() && list.Add()
        rect := Buffer(16, 0)                                               ; left = LVIR_BOUNDS (0)
        height := SendMessage(0x100E, 0, rect.Ptr, list.Hwnd) ? NumGet(rect, 12, "Int") - NumGet(rect, 4, "Int") : 0   ; LVM_GETITEMRECT
        if added
            list.Delete()
        return (height > 0) ? height : fallback
    }

    ; The control font's {Ascent: baseline to top, Cap: height of a capital H} (pixels); both 0 if they can't be measured
    static CapMetrics(hwnd) {
        result := {Ascent: 0, Cap: 0}
        hdc := DllCall("GetDC", "Ptr", hwnd, "Ptr")
        font := SendMessage(0x31, 0, 0, hwnd)                               ; WM_GETFONT
        old := DllCall("SelectObject", "Ptr", hdc, "Ptr", font, "Ptr")
        tm := Buffer(60, 0)                                                 ; TEXTMETRICW
        if DllCall("GetTextMetricsW", "Ptr", hdc, "Ptr", tm)
            result.Ascent := NumGet(tm, 4, "Int")
        gm := Buffer(20, 0), mat := Buffer(16, 0)                           ; GLYPHMETRICS, MAT2 (identity matrix)
        NumPut("Short", 1, mat, 2), NumPut("Short", 1, mat, 14)
        if (DllCall("GetGlyphOutlineW", "Ptr", hdc, "UInt", Ord("H"), "UInt", 0, "Ptr", gm, "UInt", 0, "Ptr", 0, "Ptr", mat, "UInt") != 0xFFFFFFFF)
            result.Cap := NumGet(gm, 12, "Int")                             ; gmptGlyphOrigin.y: how far the top of the glyph is above the baseline
        DllCall("SelectObject", "Ptr", hdc, "Ptr", old)
        DllCall("ReleaseDC", "Ptr", hwnd, "Ptr", hdc)
        return result
    }

    static TextExtent(hwnd, text) {
        hdc := DllCall("GetDC", "Ptr", hwnd, "Ptr")
        font := SendMessage(0x31, 0, 0, hwnd)                               ; WM_GETFONT
        old := DllCall("SelectObject", "Ptr", hdc, "Ptr", font, "Ptr")
        size := Buffer(8, 0)
        DllCall("GetTextExtentPoint32W", "Ptr", hdc, "WStr", text, "Int", StrLen(text), "Ptr", size)
        DllCall("SelectObject", "Ptr", hdc, "Ptr", old)
        DllCall("ReleaseDC", "Ptr", hwnd, "Ptr", hdc)
        return NumGet(size, 0, "Int")
    }

    ; Windows 11 window border color, color as "RRGGBB"
    static SetBorderColor(hwnd, color) {
        static ATTR := 34
        bgr := Win.ColorToBgr(color)
        try DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "UInt", ATTR,
                    "UInt*", bgr, "UInt", 4)
    }

    ; "RRGGBB" / "#RRGGBB" -> 0x00BBGGRR as GDI uses it
    static ColorToBgr(color) {
        color := LTrim(color, "#")
        rgb := Integer("0x" color)
        return ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF)
    }

    ; An antialiased filled rounded rectangle (selected row, theme thumbnails): GDI's RoundRect draws arcs on whole pixels, so the
    ; corners are jagged; GDI+ fills an antialiased path that blends into the background. right / bottom are exclusive (as in FillRect);
    ; PixelOffsetModeHalf puts integer coordinates on pixel edges, so straight edges stay sharp. Falls back to RoundRect without GDI+
    static FillRoundRect(hdc, left, top, right, bottom, radius, bgr) {
        w := right - left, h := bottom - top
        if (w <= 0 || h <= 0)
            return
        d := Min(radius * 2, w, h), graphics := 0, shape := 0, brush := 0
        if (d > 0 && ClipboardData.StartGdiplus() && !DllCall("gdiplus\GdipCreateFromHDC", "Ptr", hdc, "Ptr*", &graphics) && graphics) {
            DllCall("gdiplus\GdipSetSmoothingMode", "Ptr", graphics, "Int", 4)        ; SmoothingModeAntiAlias
            DllCall("gdiplus\GdipSetPixelOffsetMode", "Ptr", graphics, "Int", 4)      ; PixelOffsetModeHalf
            DllCall("gdiplus\GdipCreatePath", "Int", 0, "Ptr*", &shape)
            for arc in [[left, top, 180], [right - d, top, 270], [right - d, bottom - d, 0], [left, bottom - d, 90]]
                DllCall("gdiplus\GdipAddPathArc", "Ptr", shape, "Float", arc[1], "Float", arc[2], "Float", d, "Float", d, "Float", arc[3], "Float", 90)
            DllCall("gdiplus\GdipClosePathFigure", "Ptr", shape)
            argb := 0xFF000000 | ((bgr & 0xFF) << 16) | (bgr & 0xFF00) | ((bgr >> 16) & 0xFF)
            DllCall("gdiplus\GdipCreateSolidFill", "UInt", argb, "Ptr*", &brush)
            DllCall("gdiplus\GdipFillPath", "Ptr", graphics, "Ptr", brush, "Ptr", shape)
            DllCall("gdiplus\GdipDeleteBrush", "Ptr", brush)
            DllCall("gdiplus\GdipDeletePath", "Ptr", shape)
            DllCall("gdiplus\GdipDeleteGraphics", "Ptr", graphics)
            return
        }
        brush := DllCall("CreateSolidBrush", "UInt", bgr, "Ptr")
        oldBrush := DllCall("SelectObject", "Ptr", hdc, "Ptr", brush, "Ptr")
        oldPen := DllCall("SelectObject", "Ptr", hdc, "Ptr", DllCall("GetStockObject", "Int", 8, "Ptr"), "Ptr")   ; NULL_PEN
        DllCall("RoundRect", "Ptr", hdc, "Int", left, "Int", top, "Int", right + 1, "Int", bottom + 1, "Int", d, "Int", d)
        DllCall("SelectObject", "Ptr", hdc, "Ptr", oldPen)
        DllCall("SelectObject", "Ptr", hdc, "Ptr", oldBrush)
        DllCall("DeleteObject", "Ptr", brush)
    }

    ; Scales a pixel value by the system DPI (windows are created with -DPIScale and scale themselves)
    static Scale(px) {
        return Round(px * A_ScreenDPI / 96)
    }

    ; Work area of the monitor under the mouse, as {Left, Top, Right, Bottom}
    static WorkAreaAtMouse() {
        CoordMode("Mouse", "Screen")
        MouseGetPos(&mouseX, &mouseY)
        return Win.WorkAreaAt(mouseX, mouseY)
    }

    ; Work area of the monitor that contains screen point (x, y); the primary monitor if it is on none
    static WorkAreaAt(x, y) {
        Loop MonitorGetCount() {
            MonitorGet(A_Index, &left, &top, &right, &bottom)
            if (x >= left && x < right && y >= top && y < bottom)
                return Win.WorkArea(A_Index)
        }
        return Win.WorkArea(MonitorGetPrimary())
    }

    ; Work area of the monitor under the window's center; "" if the window doesn't exist
    static WorkAreaOfWindow(hwnd) {
        if !hwnd
            return ""
        try WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
        catch
            return ""
        return Win.WorkAreaAt(x + w // 2, y + h // 2)
    }

    static WorkArea(monitor) {
        MonitorGetWorkArea(monitor, &left, &top, &right, &bottom)
        return {Left: left, Top: top, Right: right, Bottom: bottom}
    }

    ; Dark title bar on Windows 10/11
    static SetDarkTitleBar(hwnd, dark := true) {
        static ATTR := 20
        try DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "UInt", ATTR,
                    "Int*", dark ? 1 : 0, "UInt", 4)
    }

    ; Switch to an English input language (US keyboard layout)
    static SwitchToEnglishIME() {
        try DllCall("ActivateKeyboardLayout", "UInt", 0x04090409, "UInt", 0)
    }

    ; Switches the foreground window to the next installed input language (like Win+Space); returns the new one (HKL), 0 with only one
    static NextInputLanguage() {
        if !(hwnd := WinExist("A"))
            return 0
        count := DllCall("GetKeyboardLayoutList", "Int", 0, "Ptr", 0, "Int")
        if (count < 2)
            return 0
        buf := Buffer(count * A_PtrSize)
        DllCall("GetKeyboardLayoutList", "Int", count, "Ptr", buf, "Int")
        layouts := []
        Loop count
            layouts.Push(NumGet(buf, (A_Index - 1) * A_PtrSize, "Ptr"))
        current := Win.KeyboardLayout(hwnd)
        next := Win.NextLayout(layouts, current)
        ; When the foreground is ALTRun's own window (e.g. the search window), switch directly: 2026-10 report that CapsLock couldn't switch
        ; to Chinese in the search window (with "switch to English when shown" on, the window had switched to English with
        ; ActivateKeyboardLayout), and the message sent to our own window had no effect. Like restoring on hide (SearchInputLanguage.Restore), use ActivateKeyboardLayout
        own := DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "Ptr", 0, "UInt") = DllCall("GetCurrentThreadId", "UInt")
        if own {
            try DllCall("ActivateKeyboardLayout", "Ptr", next, "UInt", 0)
        } else {
            try PostMessage(0x50, 0, next, , "ahk_id " hwnd)               ; WM_INPUTLANGCHANGEREQUEST: ask the foreground window to switch
        }
        if Logger.Enabled
            Logger.Debug(Format("Input language: {:08X} -> {:08X} ({}), now {:08X}", current & 0xFFFFFFFF, next & 0xFFFFFFFF
                , own ? "ALTRun window" : "posted", Win.KeyboardLayout(hwnd) & 0xFFFFFFFF))   ; While sending, now is still the old one (the other window switches a little later)
        return next
    }

    ; The one after current in layouts (after the last, the first again); the first if current isn't in it
    static NextLayout(layouts, current) {
        for index, layout in layouts
            if (layout = current)
                return layouts[Mod(index, layouts.Length) + 1]
        return layouts.Length ? layouts[1] : 0
    }

    ; Toggles the Chinese / English mode of the input method in the foreground window (like Shift in Microsoft Pinyin). true = a toggle was sent (false when the window has no input method).
    ; Chinese = open the IME + add the native bit to the conversion mode; English = remove the native bit + close the IME. Both are set because input methods differ:
    ; WeChat IME only reacts to open / close; changing only the conversion mode leaves its Chinese / English indicator as it is (log of #154, 2026-10).
    ; After each toggle the reading and the side it went to are kept; if the next reading is unchanged, that is used (for IMEs whose reading doesn't follow), see ImeTarget
    static _imeLast := Map()                                                ; window -> {Open, Conv: reading after the toggle, Chinese: switched to Chinese}
    static ToggleImeMode() {
        if !(hwnd := WinExist("A")) || !(imeWnd := Win.ImeWindow(hwnd))
            return false
        before := Win.ImeState(imeWnd)
        open := before.Open, conv := before.Conv
        last := Win._imeLast.Has(hwnd) ? Win._imeLast[hwnd] : ""
        plan := Win.ImeTarget(open, conv, last)
        Win.SetImeState(imeWnd, plan.Open, plan.Conv)
        after := Win.ImeState(imeWnd)
        after.Chinese := plan.Chinese
        if (Win._imeLast.Count > 50)                                        ; only the windows used recently are kept
            Win._imeLast := Map()
        Win._imeLast[hwnd] := after
        if Logger.Enabled
            Logger.Debug(Format("IME mode: open {} conv 0x{:X}{} -> {} (set open {}, conv 0x{:X}), now open {} conv 0x{:X}"
                , open, conv, plan.Remembered ? " (readings unchanged, using the last switch)" : "", plan.Chinese ? "Chinese" : "English"
                , plan.Open, plan.Conv, after.Open, after.Conv))
        return true
    }

    ; The IME window of the window's focused control (or the window itself), for WM_IME_CONTROL; 0 if not found
    static ImeWindow(hwnd) {
        info := Buffer(24 + 6 * A_PtrSize, 0)                                ; GUITHREADINFO: find the focused control
        NumPut("UInt", info.Size, info)
        threadId := DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "Ptr", 0, "UInt")
        focus := DllCall("GetGUIThreadInfo", "UInt", threadId, "Ptr", info) ? NumGet(info, 8 + A_PtrSize, "Ptr") : 0
        return DllCall("imm32\ImmGetDefaultIMEWnd", "Ptr", focus ? focus : hwnd, "Ptr")
    }

    ; The IME's current {Open: open status, Conv: conversion mode}
    static ImeState(imeWnd) {
        return {Open: Win._ImeControl(imeWnd, 5), Conv: Win._ImeControl(imeWnd, 1)}    ; IMC_GETOPENSTATUS, IMC_GETCONVERSIONMODE
    }

    ; Sets the IME's open status and conversion mode, in the order given by ImeSteps
    static SetImeState(imeWnd, open, conv) {
        for step in Win.ImeSteps(open, conv)
            Win._ImeControl(imeWnd, step[1], step[2])
    }

    ; The [WM_IME_CONTROL command, value] pairs to send: when opening, open first and then set the mode (some IMEs reset the mode on open); when closing, set the mode first and then close
    static ImeSteps(open, conv) {
        static IMC_SETCONVERSIONMODE := 2, IMC_SETOPENSTATUS := 6
        if open
            return [[IMC_SETOPENSTATUS, 1], [IMC_SETCONVERSIONMODE, conv]]
        return [[IMC_SETCONVERSIONMODE, conv], [IMC_SETOPENSTATUS, 0]]
    }

    ; The {Open, Conv} for English mode: remove the native bit (IME_CMODE_NATIVE = 1) and close the IME
    static ImeEnglish(conv) {
        return {Open: 0, Conv: conv & ~1}
    }

    static _ImeControl(imeWnd, command, value := 0) {                        ; 2 = SMTO_ABORTIFHUNG: don't hang along with a hung foreground program
        result := 0
        DllCall("SendMessageTimeoutW", "Ptr", imeWnd, "UInt", 0x283, "Ptr", command, "Ptr", value, "UInt", 2, "UInt", 300, "Ptr*", &result)   ; WM_IME_CONTROL
        return result
    }

    ; What a Chinese / English toggle does: {Chinese: switch to Chinese, Open: open status to set, Conv: mode to set, Remembered: decided from the kept state}
    ;   open / conv: the open status and conversion mode read now (IME_CMODE_NATIVE = 1 is Chinese)
    ;   last:        {Open, Conv, Chinese} after the last toggle, "" if none. A reading equal to that one = the reading didn't follow (or nothing changed),
    ;                so it is taken as still being on the side switched to last time
    static ImeTarget(open, conv, last := "") {
        remembered := IsObject(last) && last.Open = open && last.Conv = conv
        chinese := remembered ? last.Chinese : (open && (conv & 1))
        if chinese
            return {Chinese: false, Open: 0, Conv: conv & ~1, Remembered: remembered}
        return {Chinese: true, Open: 1, Conv: conv | 1, Remembered: remembered}
    }

    ; The input language (HKL) the window's thread uses; hwnd = 0: the current thread
    static KeyboardLayout(hwnd := 0) {
        threadId := hwnd ? DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "Ptr", 0, "UInt") : 0
        return DllCall("GetKeyboardLayout", "UInt", threadId, "Ptr")
    }

    ; Whether the process runs as administrator (elevated token); false if it can't be opened (e.g. a normal process asking about an elevated one)
    static IsElevated(pid) {
        elevated := 0
        if !(process := DllCall("OpenProcess", "UInt", 0x1000, "Int", 0, "UInt", pid, "Ptr"))   ; PROCESS_QUERY_LIMITED_INFORMATION
            return false
        token := 0
        if DllCall("advapi32\OpenProcessToken", "Ptr", process, "UInt", 0x8, "Ptr*", &token) {   ; TOKEN_QUERY
            DllCall("advapi32\GetTokenInformation", "Ptr", token, "Int", 20, "UInt*", &elevated, "UInt", 4, "UInt*", &size := 0)   ; TokenElevation
            DllCall("CloseHandle", "Ptr", token)
        }
        DllCall("CloseHandle", "Ptr", process)
        return elevated != 0
    }

    ; Windows setting "Let me use a different input method for each app window" is on (SPI_GETTHREADLOCALINPUTSETTINGS)
    static PerWindowInputMethod() {
        on := 0
        try DllCall("SystemParametersInfo", "UInt", 0x104E, "UInt", 0, "Int*", &on, "UInt", 0)
        return on != 0
    }

    ; Waits until all modifier keys are up, so Ctrl / Alt don't leak into the next Send
    static WaitModifiersUp(timeout := 0.5) {
        for key in ["Ctrl", "Alt", "Shift", "LWin", "RWin"]
            KeyWait(key, "T" timeout)
    }

    ; Gray system cue text while the box is empty (it stays while the box has focus and is never taken for real input);
    ; it gives way as soon as the user types and comes back when the box is emptied.
    ; A tooltip when the mouse rests on a control (e.g. the full path of one cut off with an ellipsis); a static text needs SS_NOTIFY (0x100) to get mouse messages.
    ; The tooltip belongs to the control's window and is destroyed with it. Returns the tooltip window
    static AddTooltip(ctrlHwnd, text) {
        static TTS_ALWAYSTIP := 0x1, TTS_NOPREFIX := 0x2, TTF_IDISHWND := 0x1, TTF_SUBCLASS := 0x10
        static TTM_ADDTOOLW := 0x432, TTM_SETMAXTIPWIDTH := 0x418
        owner := DllCall("GetAncestor", "Ptr", ctrlHwnd, "UInt", 2, "Ptr")  ; GA_ROOT
        tip := DllCall("CreateWindowExW", "UInt", 0x8, "Str", "tooltips_class32", "Ptr", 0, "UInt", 0x80000000 | TTS_ALWAYSTIP | TTS_NOPREFIX
                     , "Int", 0x80000000, "Int", 0x80000000, "Int", 0x80000000, "Int", 0x80000000, "Ptr", owner, "Ptr", 0, "Ptr", 0, "Ptr", 0, "Ptr")
        info := Buffer(24 + 6 * A_PtrSize, 0)                               ; TOOLINFOW
        NumPut("UInt", info.Size, "UInt", TTF_IDISHWND | TTF_SUBCLASS, info)
        NumPut("Ptr", owner, "Ptr", ctrlHwnd, info, 8)
        NumPut("Ptr", StrPtr(text), info, 24 + 3 * A_PtrSize)               ; the tooltip keeps its own copy of the text
        DllCall("SendMessageW", "Ptr", tip, "UInt", TTM_SETMAXTIPWIDTH, "Ptr", 0, "Ptr", Round(600 * A_ScreenDPI / 96))
        DllCall("SendMessageW", "Ptr", tip, "UInt", TTM_ADDTOOLW, "Ptr", 0, "Ptr", info)
        return tip
    }

    static SetCueBanner(hwnd, text) {
        static EM_SETCUEBANNER := 0x1501
        try DllCall("User32\SendMessageW", "Ptr", hwnd, "UInt", EM_SETCUEBANNER, "Ptr", 1, "WStr", text)
        DllCall("InvalidateRect", "Ptr", hwnd, "Ptr", 0, "Int", 1)          ; changing the cue of a box already shown needs a repaint
    }

    ; Hotkey string -> human-readable label: "^!c" -> "Ctrl+Alt+C", "!Space" -> "Alt+Space", "~MButton" -> "MButton",
    ; "CapsLock & j" -> "CapsLock+J". Only leading modifiers count (so "^+" is Ctrl + the plus key); ~ * $ aren't shown.
    ; names: optional Map that replaces key names with other text, e.g. Map("MButton", "Middle button")
    static HotkeyLabel(hk, names := "") {
        hk := Trim(hk)
        if (hk = "" || hk = "None")
            return ""
        if InStr(hk, " & ") {                                               ; a combination of two keys (CapsLock & j)
            parts := StrSplit(hk, "&", " `t")
            return Win.HotkeyLabel(parts[1], names) "+" Win.HotkeyLabel(parts[2], names)
        }
        mods := "", i := 1
        while (i < StrLen(hk) && InStr("~*$<>^!+#", SubStr(hk, i, 1)))
            mods .= SubStr(hk, i++, 1)
        key := SubStr(hk, i), up := ""
        if RegExMatch(key, "i)^(.+?)\s+up$", &m)
            key := m[1], up := " Up"
        label := (InStr(mods, "^") ? "Ctrl+" : "") (InStr(mods, "!") ? "Alt+" : "") (InStr(mods, "+") ? "Shift+" : "") (InStr(mods, "#") ? "Win+" : "")
        return label Win.KeyLabel(key, names) up
    }

    ; The display name of a key: "space" -> "Space", "c" -> "C", "esc" -> "Escape"
    static KeyLabel(key, names := "") {
        if (StrLen(key) = 1)                                                ; a single character is shown as it is (GetKeyName("+") would give "=", on the same key)
            return StrUpper(key)
        name := GetKeyName(key)
        if (name = "")
            name := key
        if (IsObject(names) && names.Has(name))
            return names[name]
        return (StrLen(name) = 1) ? StrUpper(name) : name
    }
}


;===============================================================================
; Arr - small array helpers
;===============================================================================
class Arr {
    ; Position of the first value in arr (1-based), 0 if not found.
    static IndexOf(value, arr) {
        for index, element in arr
            if (element = value)
                return index
        return 0
    }
}


;===============================================================================
; Pinyin - initials of Chinese characters (Resources\Pinyin.txt, GB2312 ranges as fallback)
;===============================================================================
class Pinyin {
    static _ranges := [
        [-20319,-20284,"A"], [-20283,-19776,"B"], [-19775,-19219,"C"], [-19218,-18711,"D"],
        [-18710,-18527,"E"], [-18526,-18240,"F"], [-18239,-17923,"G"], [-17922,-17418,"H"],
        [-17417,-16475,"J"], [-16474,-16213,"K"], [-16212,-15641,"L"], [-15640,-15166,"M"],
        [-15165,-14923,"N"], [-14922,-14915,"O"], [-14914,-14631,"P"], [-14630,-14150,"Q"],
        [-14149,-14091,"R"], [-14090,-13319,"S"], [-13318,-12839,"T"], [-12838,-12557,"W"],
        [-12556,-11848,"X"], [-11847,-11056,"Y"], [-11055,-10247,"Z"] ]

    ; "记事本" (Notepad) -> "JSB"; other characters stay as they are
    ; Each character is looked up once (_cache): the same characters come up again and again in file and app names
    static _cache := Map(), _buf := Buffer(4)
    static Initials(str) {
        if !RegExMatch(str, "[^\x{00}-\x{ff}]")     ; No Chinese: return it as it is
            return str
        out := "", cache := Pinyin._cache
        for ch in StrSplit(str) {
            code := Ord(ch)
            if (code >= 0x2E80 && code <= 0x9FFF)
                out .= cache.Has(ch) ? cache[ch] : (cache[ch] := Pinyin._Initial(ch))
            else
                out .= ch
        }
        return out
    }

    ; The initial of one character: CJK U+3400 - U+9FFF from Resources\Pinyin.txt (one ASCII letter per code point,
    ; made by Tests\Tools\MakePinyin.py; "-" = no reading), so GB2312 level-2, GBK and traditional characters work too.
    ; Other characters, or without the file: the GB2312 ranges below (only the 3755 level-1 characters, sorted by pinyin)
    static DataFile := A_ScriptDir "\Resources\Pinyin.txt", TableStart := 0x3400, _letters := Buffer(0), _tried := false
    static _Initial(ch) {
        if !Pinyin._tried
            Pinyin._Load()
        index := Ord(ch) - Pinyin.TableStart
        if (index >= 0 && index < Pinyin._letters.Size) {
            letter := NumGet(Pinyin._letters, index, "UChar")
            return (letter = 45) ? "" : Chr(letter)                         ; 45 = "-"
        }
        buf := Pinyin._buf
        StrPut(ch, buf, "CP936")
        gbk := (NumGet(buf, 0, "UChar") << 8) + NumGet(buf, 1, "UChar") - 65536
        for r in Pinyin._ranges
            if (gbk >= r[1] && gbk <= r[2])
                return r[3]
        return ""
    }

    ; Read once, the first time a character is looked up (a Buffer: no string copies; empty when the file is missing,
    ; then the ranges still work)
    static _Load() {
        Pinyin._tried := true
        try Pinyin._letters := FileRead(Pinyin.DataFile, "RAW")
    }
}


;===============================================================================
; Calc - evaluates arithmetic expressions (only numbers and + - * / ^ ** ( ))
; Deliberately no eval, so no arbitrary code runs. Recursive descent, all in floating point (large numbers don't overflow);
; precedence: parentheses > power (right-associative, 2^3^2 = 2^(3^2)) > sign (-2^2 = -4, 2^-2 = 0.25) > * / > + -
;===============================================================================
class Calc {
    ; Whether it looks like an expression
    static Looks(expr) {
        return RegExMatch(expr, "[\+\-\*/\^]") && RegExMatch(expr, "^[\d\+\-\*/\^\(\)\.\s]+$")
    }

    ; Returns the result (a float); "" when it isn't a complete expression, divides by 0 or is out of range
    static Eval(expr) {
        expr := RegExReplace(expr, "\s")
        if !RegExMatch(expr, "^[\d\+\-\*/\^\(\)\.]+$")
            return ""
        state := {Text: expr, Pos: 1}
        try {
            value := Calc._Sum(state)
            if (state.Pos <= StrLen(expr))                                  ; characters left over, e.g. "(1+2))"
                return ""
        } catch
            return ""
        return (Abs(value) <= 1.7976931348623157e308) ? value : ""          ; overflow (inf) / NaN aren't results
    }

    static _Sum(state) {                                                    ; + -
        value := Calc._Product(state)
        while (op := Calc._Peek(state, "+-")) {
            state.Pos += 1
            operand := Calc._Product(state)
            value := (op = "+") ? value + operand : value - operand
        }
        return value
    }

    static _Product(state) {                                                ; * /
        value := Calc._Signed(state)
        while (op := Calc._Peek(state, "*/")) {
            state.Pos += 1
            operand := Calc._Signed(state)
            value := (op = "*") ? value * operand : value / operand         ; division by 0 throws; Eval returns ""
        }
        return value
    }

    static _Signed(state) {                                                 ; sign
        if (op := Calc._Peek(state, "+-")) {
            state.Pos += 1
            return (op = "-") ? -Calc._Signed(state) : Calc._Signed(state)
        }
        return Calc._Power(state)
    }

    static _Power(state) {                                                  ; power: ^ or **
        base := Calc._Primary(state)
        if (SubStr(state.Text, state.Pos, 1) = "^")
            state.Pos += 1
        else if (SubStr(state.Text, state.Pos, 2) = "**")
            state.Pos += 2
        else
            return base
        return base ** Calc._Signed(state)                                  ; the exponent can have a sign and be a power itself
    }

    static _Primary(state) {                                                ; a number or parentheses
        if (SubStr(state.Text, state.Pos, 1) = "(") {
            state.Pos += 1
            value := Calc._Sum(state)
            if (SubStr(state.Text, state.Pos, 1) != ")")
                throw ValueError("Missing )")
            state.Pos += 1
            return value
        }
        if !RegExMatch(state.Text, "\G(\d+\.?\d*|\.\d+)", &m, state.Pos)
            throw ValueError("Number expected")
        state.Pos += m.Len
        return Float(m[0])
    }

    static _Peek(state, chars) {
        ch := SubStr(state.Text, state.Pos, 1)
        return (ch != "" && InStr(chars, ch)) ? ch : ""
    }

    ; Thousands separators (negative numbers too): -1234567.5 -> -1,234,567.5
    static Thousands(num) {
        return RegExReplace(num "", "\G-?\d+?(?=(\d{3})+(?:\D|$))", "$0,")
    }
}


;===============================================================================
; Url - URL helpers
;===============================================================================
class Url {
    ; Percent-encodes as UTF-8, to put search words into a {query} URL
    static Encode(text) {
        out := ""
        for ch in StrSplit(text) {
            if RegExMatch(ch, "^[0-9A-Za-z\-_.~]$") {
                out .= ch
                continue
            }
            buf := Buffer(8, 0)
            len := StrPut(ch, buf, "UTF-8") - 1
            Loop len
                out .= Format("%{:02X}", NumGet(buf, A_Index - 1, "UChar"))
        }
        return out
    }
}
