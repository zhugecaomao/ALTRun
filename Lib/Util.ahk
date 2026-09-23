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
            "A_MyDocuments"    , A_MyDocuments,
            "A_ProgramFiles"   , A_ProgramFiles,
            "A_ProgramsCommon" , A_ProgramsCommon,
            "A_StartupCommon"  , A_StartupCommon,
            "A_StartMenuCommon", A_StartMenuCommon,
            "A_WinDir"         , A_WinDir
        )
        return Path._varCache
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
        if (!FileExist(raw) && !InStr(raw, "\")) {
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

    ; 按系统 DPI 缩放像素值 (窗口用 -DPIScale 创建, 自己控制缩放)
    static Scale(px) {
        return Round(px * A_ScreenDPI / 96)
    }

    ; 鼠标所在显示器的工作区, 返回 {Left, Top, Right, Bottom}
    static WorkAreaAtMouse() {
        CoordMode("Mouse", "Screen")
        MouseGetPos(&mouseX, &mouseY)
        Loop MonitorGetCount() {
            MonitorGet(A_Index, &left, &top, &right, &bottom)
            if (mouseX >= left && mouseX < right && mouseY >= top && mouseY < bottom) {
                MonitorGetWorkArea(A_Index, &left, &top, &right, &bottom)
                return {Left: left, Top: top, Right: right, Bottom: bottom}
            }
        }
        MonitorGetWorkArea(MonitorGetPrimary(), &left, &top, &right, &bottom)
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

    ; 等所有修饰键松开, 防止 Ctrl/Alt 漏进后续 Send
    static WaitModifiersUp(timeout := 0.5) {
        for key in ["Ctrl", "Alt", "Shift", "LWin", "RWin"]
            KeyWait(key, "T" timeout)
    }

    ; 输入框留空时显示系统原生的灰色提示文字(焦点在框里时也不消失, 不会被
    ; 误当成一次真实输入), 一旦用户开始打字就自动让位, 清空后又自动回来。
    static SetCueBanner(hwnd, text) {
        static EM_SETCUEBANNER := 0x1501
        try DllCall("User32\SendMessageW", "Ptr", hwnd, "UInt", EM_SETCUEBANNER, "Ptr", 1, "WStr", text)
    }

    ; 热键字符串 -> 人类可读标签, 例如 "^g" -> "Ctrl+G"
    static HotkeyLabel(hk) {
        if (!hk || hk = "None")
            return ""
        label := ""
        if InStr(hk, "^")
            label .= "Ctrl+"
        if InStr(hk, "!")
            label .= "Alt+"
        if InStr(hk, "+")
            label .= "Shift+"
        if InStr(hk, "#")
            label .= "Win+"
        base := RegExReplace(hk, "[\^\!\+\#\<\>\*\~\$\s]")
        return label (StrLen(base) = 1 ? StrUpper(base) : base)
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
; Pinyin - 中文拼音首字母
; 用 GBK 编码区间映射, 不需要词库, 够用且零依赖
;===============================================================================
class Pinyin {
    static _ranges := [
        [-20319,-20284,"A"], [-20283,-19776,"B"], [-19775,-19219,"C"], [-19218,-18711,"D"],
        [-18710,-18527,"E"], [-18526,-18240,"F"], [-18239,-17923,"G"], [-17922,-17418,"H"],
        [-17417,-16475,"J"], [-16474,-16213,"K"], [-16212,-15641,"L"], [-15640,-15166,"M"],
        [-15165,-14923,"N"], [-14922,-14915,"O"], [-14914,-14631,"P"], [-14630,-14150,"Q"],
        [-14149,-14091,"R"], [-14090,-13319,"S"], [-13318,-12839,"T"], [-12838,-12557,"W"],
        [-12556,-11848,"X"], [-11847,-11056,"Y"], [-11055,-10247,"Z"] ]

    ; "记事本" -> "JSB"; 非中文字符原样保留
    static Initials(str) {
        if !RegExMatch(str, "[^\x{00}-\x{ff}]")     ; 没有中文就直接返回
            return str
        out := ""
        for ch in StrSplit(str) {
            code := Ord(ch)
            if (code >= 0x2E80 && code <= 0x9FFF) {
                buf := Buffer(4)
                StrPut(ch, buf, "CP936")
                gbk := (NumGet(buf, 0, "UChar") << 8) + NumGet(buf, 1, "UChar") - 65536
                for r in Pinyin._ranges {
                    if (gbk >= r[1] && gbk <= r[2]) {
                        out .= r[3]
                        break
                    }
                }
            } else {
                out .= ch
            }
        }
        return out
    }
}


;===============================================================================
; Calc - 四则运算表达式求值 (只认数字和 + - * / ^ ( ))
; 故意不走 eval, 避免任意代码执行
;===============================================================================
class Calc {
    ; 看起来像不像一个算式
    static Looks(expr) {
        return RegExMatch(expr, "[\+\-\*/\^]") && RegExMatch(expr, "^[\d\+\-\*/\^\(\)\.\s]+$")
    }

    static Eval(expr, depth := 0) {
        if (depth > 12)
            return ""
        expr := StrReplace(expr, " ")
        if (!RegExMatch(expr, "^[\d\+\-\*/\^\(\)\.]*$"))
            return ""

        ; 先递归消掉括号
        while RegExMatch(expr, "\(([^()]*)\)", &m) {
            inner := Calc.Eval(m[1], depth + 1)
            if (inner = "")
                return ""
            expr := StrReplace(expr, m[0], inner)
        }
        return Calc._Flat(expr)
    }

    ; 无括号表达式: 按 幂 -> 乘除 -> 加减 的优先级逐步归约
    static _Flat(expr) {
        ; 乘方从右往左结合(2^3^2 = 2^(3^2), 不是 (2^3)^2), 且底数不吃掉前面的负号
        ; (-2^2 = -(2^2) = -4, 不是 (-2)^2 = 4) - 指数本身仍允许带负号(2^-2 合法)。
        ; 用贪婪 ".*" 顶到字符串最右边再回溯, 天然找到"最靠右"的一组底数^指数。
        while RegExMatch(expr, "(.*)(\d+(?:\.\d+)?)(\^|\*\*)(-?\d+(?:\.\d+)?)", &m)
            expr := m[1] . (m[2] ** m[4]) . SubStr(expr, m.Pos + m.Len)

        while RegExMatch(expr, "(-?\d+(?:\.\d+)?)([*/])(-?\d+(?:\.\d+)?)", &m) {
            a := m[1] + 0, b := m[3] + 0
            if (m[2] = "*")
                r := a * b
            else
                r := (Abs(b) < 1e-12) ? 0 : a / b   ; 除零直接给 0, 不抛异常
            expr := StrReplace(expr, m[0], r)
        }

        while RegExMatch(expr, "(-?\d+(?:\.\d+)?)([+\-])(-?\d+(?:\.\d+)?)", &m) {
            a := m[1] + 0, b := m[3] + 0
            expr := StrReplace(expr, m[0], m[2] = "+" ? a + b : a - b)
        }
        return expr
    }

    ; 千分位
    static Thousands(num) {
        return RegExReplace(num "", "\G\d+?(?=(\d{3})+(?:\D|$))", "$0,")
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
