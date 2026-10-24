;===============================================================================
; Util.ahk - 通用工具函数
;-------------------------------------------------------------------------------
; 按用途分成几个小类, 都是纯函数, 不依赖 Config/UI, 方便单独测试。
;===============================================================================


;===============================================================================
; Path - 路径解析
;-------------------------------------------------------------------------------
; 支持三种写法:
;   A_ScriptDir\x.txt   AHK 内置变量开头
;   %Temp%\x.txt        环境变量风格占位符
;   notepad.exe         裸可执行名 (自动在 PATH 里找)
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

    ; 延迟初始化内置变量表 (启动时就取会拖慢冷启动)
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

    ; 用户写的文件夹路径 -> 完整路径: 展开 %环境变量% 和 A_ 开头的内置变量, 相对路径从 base 算起
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

    ; 能不能在这个文件夹里写文件 (试着建一个临时文件再删掉)
    static IsWritable(dir) {
        probe := dir "\ALTRun.write-test.tmp"
        try {
            FileAppend("", probe)
            FileDelete(probe)
            return true
        }
        return false
    }

    ; 转成绝对路径。keepRunAs=true 时保留 "*RunAs " 前缀(用于实际执行)
    ; 注意: 参数不能叫 path —— AHK 变量名不区分大小写, "path" 和类名 "Path" 会
    ; 被当成同一个标识符, 方法体内所有 Path.xxx 的自引用都会失效(改成读参数)。
    static Resolve(raw, keepRunAs := false) {
        raw := Trim(raw)
        if (raw = "")
            return ""

        if (!keepRunAs)
            raw := StrReplace(raw, "*RunAs ", "")

        vars := Path._Vars()

        ; 形如 "A_ScriptDir\sub\x.exe" —— 必须以 A_ 开头才算, 否则像 "Plot A_IGLS" 会被误判
        if (InStr(raw, "A_") = 1) {
            if (InStr(raw, "\")) {
                parts := StrSplit(raw, "\",, 2)
                if (vars.Has(parts[1]))
                    raw := vars[parts[1]] "\" (parts.Has(2) ? parts[2] : "")
            } else if (vars.Has(raw)) {
                raw := vars[raw]
            }
        }

        ; 环境变量占位符
        raw := StrReplace(raw, "%Temp%"    , A_Temp)
        raw := StrReplace(raw, "%OneDrive%", EnvGet("OneDrive"))
        raw := StrReplace(raw, "%AppData%" , A_AppData)
        raw := StrReplace(raw, "%UserProfile%", EnvGet("UserProfile"))

        ; 裸文件名: 去 PATH 里搜, 没写扩展名时按 .exe 找 ("notepad" -> notepad.exe)
        ; 先判断有没有 "\": 完整路径不访问磁盘 (网络驱动器断开时 FileExist 会卡住)
        if (!InStr(raw, "\") && !FileExist(raw)) {
            buf := Buffer(260 * 2)
            if DllCall("kernel32\SearchPathW", "Ptr", 0, "WStr", raw, "WStr", ".exe",
                       "UInt", buf.Size // 2, "Ptr", buf, "Ptr", 0)
                raw := StrGet(buf, "UTF-16")
        }
        return raw
    }

    ; 反向: 把常见目录换回占位符, 让配置文件跨机器可移植
    static Shorten(path) {
        path := StrReplace(path, A_Temp, "%Temp%")
        if (od := EnvGet("OneDrive"))
            path := StrReplace(path, od, "%OneDrive%")
        return path
    }

    ; 只取文件名, 用于列表显示
    ; 用这种文件的默认程序打开文本文件 (例如把 .log 设成用 Notepad4 打开), 打不开时用记事本。
    ; 只用于默认动作就是查看的类型 (.log .txt): 脚本 (.ahk .bat) 的默认动作是运行, .json 常常没有关联。
    ; runner 只在测试时传入 (代替 Run)
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
; Fonts - 字体描述串 <-> AHK SetFont 参数
; 配置里存 "Microsoft YaHei, norm s10" 这种人类可读格式
;===============================================================================
class Fonts {
    static Spec(text, defName := "Microsoft YaHei", defOpt := "norm s10") {
        parts := StrSplit(text, ",")
        name  := parts.Length >= 1 ? Trim(parts[1]) : ""
        opt   := parts.Length >= 2 ? Trim(parts[2]) : ""
        return { name: name != "" ? name : defName, opt: opt != "" ? opt : defOpt }
    }

    ; 调用系统字体对话框, 返回 {name, opt} 或 false
    static Choose(initName := "", ownerHwnd := 0) {
        lf := Buffer((A_PtrSize = 4) ? 60 : 92, 0)
        dc := DllCall("GetDC", "Ptr", 0, "Ptr")
        dpi := DllCall("GetDeviceCaps", "Ptr", dc, "UInt", 90, "Int")
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", dc)

        NumPut("UInt", Floor(10 * dpi / 72), lf)            ; 默认 10pt
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
; Win - 窗口相关的小工具
;===============================================================================
class Win {
    ; Win11 圆角
    static SetCorner(hwnd, rounded := true) {
        static ATTR := 33, ROUND := 2, SQUARE := 1
        try DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "UInt", ATTR,
                    "Int*", rounded ? ROUND : SQUARE, "UInt", 4)
    }

    ; 控件里一行文字的宽度 (像素, 用控件自己的字体)
    ; ListView 一行实际的高度 (像素): 临时加一行量出来 (LVM_GETITEMRECT); 量不到时 fallback
    static ListRowHeight(list, fallback) {
        added := !list.GetCount() && list.Add()
        rect := Buffer(16, 0)                                               ; left = LVIR_BOUNDS (0)
        height := SendMessage(0x100E, 0, rect.Ptr, list.Hwnd) ? NumGet(rect, 12, "Int") - NumGet(rect, 4, "Int") : 0   ; LVM_GETITEMRECT
        if added
            list.Delete()
        return (height > 0) ? height : fallback
    }

    ; 控件字体的 {Ascent: 基线到顶部, Cap: 大写字母 H 的高度} (像素); 量不到时都是 0
    static CapMetrics(hwnd) {
        result := {Ascent: 0, Cap: 0}
        hdc := DllCall("GetDC", "Ptr", hwnd, "Ptr")
        font := SendMessage(0x31, 0, 0, hwnd)                               ; WM_GETFONT
        old := DllCall("SelectObject", "Ptr", hdc, "Ptr", font, "Ptr")
        tm := Buffer(60, 0)                                                 ; TEXTMETRICW
        if DllCall("GetTextMetricsW", "Ptr", hdc, "Ptr", tm)
            result.Ascent := NumGet(tm, 4, "Int")
        gm := Buffer(20, 0), mat := Buffer(16, 0)                           ; GLYPHMETRICS, MAT2 (单位矩阵)
        NumPut("Short", 1, mat, 2), NumPut("Short", 1, mat, 14)
        if (DllCall("GetGlyphOutlineW", "Ptr", hdc, "UInt", Ord("H"), "UInt", 0, "Ptr", gm, "UInt", 0, "Ptr", 0, "Ptr", mat, "UInt") != 0xFFFFFFFF)
            result.Cap := NumGet(gm, 12, "Int")                             ; gmptGlyphOrigin.y: 字形顶部在基线上面多高
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

    ; Win11 窗口边框颜色, color 为 "RRGGBB"
    static SetBorderColor(hwnd, color) {
        static ATTR := 34
        bgr := Win.ColorToBgr(color)
        try DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "UInt", ATTR,
                    "UInt*", bgr, "UInt", 4)
    }

    ; "RRGGBB" / "#RRGGBB" -> GDI 用的 0x00BBGGRR
    static ColorToBgr(color) {
        color := LTrim(color, "#")
        rgb := Integer("0x" color)
        return ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF)
    }

    ; 抗锯齿的实心圆角矩形 (选中行、主题缩略图): GDI 的 RoundRect 按整像素画弧, 圆角有台阶,
    ; 这里用 GDI+ 的 AntiAlias 填充路径, 边缘和背景平滑过渡。right / bottom 不含 (和 FillRect 一样);
    ; PixelOffsetModeHalf 让整数坐标落在像素边界上, 直边不会发虚。GDI+ 用不了时退回 RoundRect
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

    ; 按系统 DPI 缩放像素值 (窗口用 -DPIScale 创建, 自己控制缩放)
    static Scale(px) {
        return Round(px * A_ScreenDPI / 96)
    }

    ; 鼠标所在显示器的工作区, 返回 {Left, Top, Right, Bottom}
    static WorkAreaAtMouse() {
        CoordMode("Mouse", "Screen")
        MouseGetPos(&mouseX, &mouseY)
        return Win.WorkAreaAt(mouseX, mouseY)
    }

    ; 包含屏幕坐标 (x, y) 的显示器的工作区; 不在任何显示器上时用主显示器
    static WorkAreaAt(x, y) {
        Loop MonitorGetCount() {
            MonitorGet(A_Index, &left, &top, &right, &bottom)
            if (x >= left && x < right && y >= top && y < bottom)
                return Win.WorkArea(A_Index)
        }
        return Win.WorkArea(MonitorGetPrimary())
    }

    ; 窗口中心所在显示器的工作区; 窗口不存在时返回 ""
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

    ; Win10/11 深色标题栏
    static SetDarkTitleBar(hwnd, dark := true) {
        static ATTR := 20
        try DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "UInt", ATTR,
                    "Int*", dark ? 1 : 0, "UInt", 4)
    }

    ; 切到英文输入法 (US 键盘布局)
    static SwitchToEnglishIME() {
        try DllCall("ActivateKeyboardLayout", "UInt", 0x04090409, "UInt", 0)
    }

    ; 前台窗口切到下一个已安装的输入法 (和 Win+Space 一样); 返回切到的输入法 (HKL), 只有一个输入法时 0
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
        ; 前台是 ALTRun 自己的窗口 (例如搜索窗口) 时直接切换: 2026-10 反馈搜索窗口里按 CapsLock 切不到中文 (打开了 "呼出时切换到
        ; 英文输入法", 搜索窗口已经用 ActivateKeyboardLayout 切到了英文), 发给自己窗口的消息没有生效。
        ; 和隐藏时切回原来的输入法一样 (SearchWindow._RestoreInputLanguage) 用 ActivateKeyboardLayout
        own := DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "Ptr", 0, "UInt") = DllCall("GetCurrentThreadId", "UInt")
        if own {
            try DllCall("ActivateKeyboardLayout", "Ptr", next, "UInt", 0)
        } else {
            try PostMessage(0x50, 0, next, , "ahk_id " hwnd)               ; WM_INPUTLANGCHANGEREQUEST: 请前台窗口切换
        }
        if Logger.Enabled
            Logger.Debug(Format("Input language: {:08X} -> {:08X} ({}), now {:08X}", current & 0xFFFFFFFF, next & 0xFFFFFFFF
                , own ? "ALTRun window" : "posted", Win.KeyboardLayout(hwnd) & 0xFFFFFFFF))   ; 发消息时 now 还是原来的 (对方稍后才切换)
        return next
    }

    ; layouts 里 current 的下一个 (最后一个之后回到第一个); current 不在里面时第一个
    static NextLayout(layouts, current) {
        for index, layout in layouts
            if (layout = current)
                return layouts[Mod(index, layouts.Length) + 1]
        return layouts.Length ? layouts[1] : 0
    }

    ; 切换前台窗口里输入法的 中 / 英 模式 (和微软拼音里按 Shift 一样)。返回 true = 发出了切换 (前台窗口没有输入法时 false)。
    ; 中文 = 打开输入法 + 转换模式加上中文位, 英文 = 去掉中文位 + 关闭输入法: 两个都设, 因为输入法认的不一样 ——
    ; 微信输入法只认打开 / 关闭, 只改转换模式时界面上的中 / 英不变 (2026-10 #154 的日志)。
    ; 另外每次切换后记下读数和切到了哪边, 下次读数还和那时一样就按记下的来 (读数不跟着变的输入法), 见 ImeTarget
    static _imeLast := Map()                                                ; 窗口 -> {Open, Conv: 切换后的读数, Chinese: 切到了中文}
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
        if (Win._imeLast.Count > 50)                                        ; 只记最近用到的窗口
            Win._imeLast := Map()
        Win._imeLast[hwnd] := after
        if Logger.Enabled
            Logger.Debug(Format("IME mode: open {} conv 0x{:X}{} -> {} (set open {}, conv 0x{:X}), now open {} conv 0x{:X}"
                , open, conv, plan.Remembered ? " (readings unchanged, using the last switch)" : "", plan.Chinese ? "Chinese" : "English"
                , plan.Open, plan.Conv, after.Open, after.Conv))
        return true
    }

    ; 窗口里有焦点的控件 (没有时窗口本身) 的输入法窗口, 用来发 WM_IME_CONTROL; 找不到时 0
    static ImeWindow(hwnd) {
        info := Buffer(24 + 6 * A_PtrSize, 0)                                ; GUITHREADINFO: 找有焦点的控件
        NumPut("UInt", info.Size, info)
        threadId := DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "Ptr", 0, "UInt")
        focus := DllCall("GetGUIThreadInfo", "UInt", threadId, "Ptr", info) ? NumGet(info, 8 + A_PtrSize, "Ptr") : 0
        return DllCall("imm32\ImmGetDefaultIMEWnd", "Ptr", focus ? focus : hwnd, "Ptr")
    }

    ; 输入法现在的 {Open: 打开状态, Conv: 转换模式}
    static ImeState(imeWnd) {
        return {Open: Win._ImeControl(imeWnd, 5), Conv: Win._ImeControl(imeWnd, 1)}    ; IMC_GETOPENSTATUS, IMC_GETCONVERSIONMODE
    }

    ; 设置输入法的打开状态和转换模式, 顺序见 ImeSteps
    static SetImeState(imeWnd, open, conv) {
        for step in Win.ImeSteps(open, conv)
            Win._ImeControl(imeWnd, step[1], step[2])
    }

    ; 要发的 [WM_IME_CONTROL 命令, 值]: 打开时先打开再设转换模式 (打开时有的输入法会重设转换模式), 关闭时先设转换模式再关闭
    static ImeSteps(open, conv) {
        static IMC_SETCONVERSIONMODE := 2, IMC_SETOPENSTATUS := 6
        if open
            return [[IMC_SETOPENSTATUS, 1], [IMC_SETCONVERSIONMODE, conv]]
        return [[IMC_SETCONVERSIONMODE, conv], [IMC_SETOPENSTATUS, 0]]
    }

    ; 切到英文模式要设的 {Open, Conv}: 去掉中文位 (IME_CMODE_NATIVE = 1) 并关闭输入法
    static ImeEnglish(conv) {
        return {Open: 0, Conv: conv & ~1}
    }

    static _ImeControl(imeWnd, command, value := 0) {                        ; 2 = SMTO_ABORTIFHUNG: 前台程序卡住时不跟着卡
        result := 0
        DllCall("SendMessageTimeoutW", "Ptr", imeWnd, "UInt", 0x283, "Ptr", command, "Ptr", value, "UInt", 2, "UInt", 300, "Ptr*", &result)   ; WM_IME_CONTROL
        return result
    }

    ; 切换中 / 英要做什么: {Chinese: 切到中文, Open: 要设的打开状态, Conv: 要设的转换模式, Remembered: 按记下的状态判断}
    ;   open / conv: 现在读到的打开状态和转换模式 (IME_CMODE_NATIVE = 1 是中文)
    ;   last:        上次切换后的 {Open, Conv, Chinese}, 没有时 ""。现在的读数和那时一样 = 读数没跟着变 (或者没有变过),
    ;                就当作还是上次切到的那边
    static ImeTarget(open, conv, last := "") {
        remembered := IsObject(last) && last.Open = open && last.Conv = conv
        chinese := remembered ? last.Chinese : (open && (conv & 1))
        if chinese
            return {Chinese: false, Open: 0, Conv: conv & ~1, Remembered: remembered}
        return {Chinese: true, Open: 1, Conv: conv | 1, Remembered: remembered}
    }

    ; 窗口 (所在线程) 正在用的输入法 (HKL); hwnd = 0: 当前线程
    static KeyboardLayout(hwnd := 0) {
        threadId := hwnd ? DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "Ptr", 0, "UInt") : 0
        return DllCall("GetKeyboardLayout", "UInt", threadId, "Ptr")
    }

    ; 进程是否以管理员身份运行 (提升了权限的令牌); 打不开 (例如普通权限查管理员的进程) 时 false
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

    ; Windows 设置里打开了 "允许我为每个应用窗口使用不同的输入法" (SPI_GETTHREADLOCALINPUTSETTINGS)
    static PerWindowInputMethod() {
        on := 0
        try DllCall("SystemParametersInfo", "UInt", 0x104E, "UInt", 0, "Int*", &on, "UInt", 0)
        return on != 0
    }

    ; 等所有修饰键松开, 防止 Ctrl/Alt 漏进后续 Send
    static WaitModifiersUp(timeout := 0.5) {
        for key in ["Ctrl", "Alt", "Shift", "LWin", "RWin"]
            KeyWait(key, "T" timeout)
    }

    ; 输入框留空时显示系统原生的灰色提示文字(焦点在框里时也不消失, 不会被
    ; 误当成一次真实输入), 一旦用户开始打字就自动让位, 清空后又自动回来。
    ; 鼠标停在控件上时显示的提示 (例如被省略号截断的路径显示完整路径); 静态文本要有 SS_NOTIFY (0x100) 才收得到鼠标。
    ; 提示窗口属于控件所在的窗口, 跟着一起销毁。返回提示窗口
    static AddTooltip(ctrlHwnd, text) {
        static TTS_ALWAYSTIP := 0x1, TTS_NOPREFIX := 0x2, TTF_IDISHWND := 0x1, TTF_SUBCLASS := 0x10
        static TTM_ADDTOOLW := 0x432, TTM_SETMAXTIPWIDTH := 0x418
        owner := DllCall("GetAncestor", "Ptr", ctrlHwnd, "UInt", 2, "Ptr")  ; GA_ROOT
        tip := DllCall("CreateWindowExW", "UInt", 0x8, "Str", "tooltips_class32", "Ptr", 0, "UInt", 0x80000000 | TTS_ALWAYSTIP | TTS_NOPREFIX
                     , "Int", 0x80000000, "Int", 0x80000000, "Int", 0x80000000, "Int", 0x80000000, "Ptr", owner, "Ptr", 0, "Ptr", 0, "Ptr", 0, "Ptr")
        info := Buffer(24 + 6 * A_PtrSize, 0)                               ; TOOLINFOW
        NumPut("UInt", info.Size, "UInt", TTF_IDISHWND | TTF_SUBCLASS, info)
        NumPut("Ptr", owner, "Ptr", ctrlHwnd, info, 8)
        NumPut("Ptr", StrPtr(text), info, 24 + 3 * A_PtrSize)               ; 提示窗口会复制一份文字
        DllCall("SendMessageW", "Ptr", tip, "UInt", TTM_SETMAXTIPWIDTH, "Ptr", 0, "Ptr", Round(600 * A_ScreenDPI / 96))
        DllCall("SendMessageW", "Ptr", tip, "UInt", TTM_ADDTOOLW, "Ptr", 0, "Ptr", info)
        return tip
    }

    static SetCueBanner(hwnd, text) {
        static EM_SETCUEBANNER := 0x1501
        try DllCall("User32\SendMessageW", "Ptr", hwnd, "UInt", EM_SETCUEBANNER, "Ptr", 1, "WStr", text)
        DllCall("InvalidateRect", "Ptr", hwnd, "Ptr", 0, "Int", 1)          ; 输入框已经显示时换提示文字, 要重画才看得到
    }

    ; 热键字符串 -> 人类可读标签: "^!c" -> "Ctrl+Alt+C", "!Space" -> "Alt+Space", "~MButton" -> "MButton",
    ; "CapsLock & j" -> "CapsLock+J"。只认开头的修饰符 (所以 "^+" 是 Ctrl + 加号键); ~ * $ 不显示。
    ; names: 可选的 Map, 把按键名换成别的文字, 例如 Map("MButton", "鼠标中键")
    static HotkeyLabel(hk, names := "") {
        hk := Trim(hk)
        if (hk = "" || hk = "None")
            return ""
        if InStr(hk, " & ") {                                               ; 两个键的组合 (CapsLock & j)
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

    ; 一个按键的显示名称: "space" -> "Space", "c" -> "C", "esc" -> "Escape"
    static KeyLabel(key, names := "") {
        if (StrLen(key) = 1)                                                ; 单个字符原样显示 (GetKeyName("+") 会给出同一个键上的 "=")
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
; Arr - 数组小工具
;===============================================================================
class Arr {
    ; value 在 arr 里第一次出现的位置(下标从 1 开始), 找不到就返回 0。
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

    ; "记事本" -> "JSB"; other characters stay as they are
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
; Calc - 四则运算表达式求值 (只认数字和 + - * / ^ ** ( ))
; 故意不走 eval, 避免任意代码执行。递归下降解析, 全程用浮点数计算 (大数不会溢出);
; 优先级: 括号 > 乘方 (右结合, 2^3^2 = 2^(3^2)) > 正负号 (-2^2 = -4, 2^-2 = 0.25) > 乘除 > 加减
;===============================================================================
class Calc {
    ; 看起来像不像一个算式
    static Looks(expr) {
        return RegExMatch(expr, "[\+\-\*/\^]") && RegExMatch(expr, "^[\d\+\-\*/\^\(\)\.\s]+$")
    }

    ; 返回计算结果 (浮点数); 不是完整的算式、除以 0、结果超出范围时返回 ""
    static Eval(expr) {
        expr := RegExReplace(expr, "\s")
        if !RegExMatch(expr, "^[\d\+\-\*/\^\(\)\.]+$")
            return ""
        state := {Text: expr, Pos: 1}
        try {
            value := Calc._Sum(state)
            if (state.Pos <= StrLen(expr))                                  ; 后面还有没用上的字符, 例如 "(1+2))"
                return ""
        } catch
            return ""
        return (Abs(value) <= 1.7976931348623157e308) ? value : ""          ; 溢出 (inf) / NaN 都不算结果
    }

    static _Sum(state) {                                                    ; 加减
        value := Calc._Product(state)
        while (op := Calc._Peek(state, "+-")) {
            state.Pos += 1
            operand := Calc._Product(state)
            value := (op = "+") ? value + operand : value - operand
        }
        return value
    }

    static _Product(state) {                                                ; 乘除
        value := Calc._Signed(state)
        while (op := Calc._Peek(state, "*/")) {
            state.Pos += 1
            operand := Calc._Signed(state)
            value := (op = "*") ? value * operand : value / operand         ; 除以 0 抛异常, Eval 返回 ""
        }
        return value
    }

    static _Signed(state) {                                                 ; 正负号
        if (op := Calc._Peek(state, "+-")) {
            state.Pos += 1
            return (op = "-") ? -Calc._Signed(state) : Calc._Signed(state)
        }
        return Calc._Power(state)
    }

    static _Power(state) {                                                  ; 乘方 ^ 或 **
        base := Calc._Primary(state)
        if (SubStr(state.Text, state.Pos, 1) = "^")
            state.Pos += 1
        else if (SubStr(state.Text, state.Pos, 2) = "**")
            state.Pos += 2
        else
            return base
        return base ** Calc._Signed(state)                                  ; 指数可以带正负号, 也可以再乘方
    }

    static _Primary(state) {                                                ; 数字或括号
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

    ; 千分位 (负数也加): -1234567.5 -> -1,234,567.5
    static Thousands(num) {
        return RegExReplace(num "", "\G-?\d+?(?=(\d{3})+(?:\D|$))", "$0,")
    }
}


;===============================================================================
; Url - 网址工具
;===============================================================================
class Url {
    ; 按 UTF-8 做百分号编码, 用于把搜索词拼进 {query} 网址
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
