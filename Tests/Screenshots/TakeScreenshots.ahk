;===============================================================================
; TakeScreenshots.ahk - 自动生成 README / Wiki 用的界面截图 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 在一个临时文件夹里准备好演示用的 ALTRun (示例设置、自定义命令、应用快捷方式、示例文件),
; 逐个场景启动 ALTRun, 输入搜索内容, 把窗口截成 PNG。不改动你自己的 ALTRun.json。
; 截图用英文界面, 演示数据都是虚构的通用内容 (不要放个人或工作相关的信息)。
;
;   AutoHotkey64.exe Tests\Screenshots\TakeScreenshots.ahk [输出文件夹] [场景名...]
;
; 输出文件夹默认 docs\images\screenshots; 不写场景名 = 全部场景。
; GitHub Actions 的 "Screenshots" 工作流在 Windows 上运行它, 并把截图提交回分支。
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn All, StdOut
#Include ..\..\Lib\JSON.ahk

Shots.Run(A_Args)
ExitApp(Shots.Failures)

class Shots {
    static RepoDir  := Shots._FullPath(A_ScriptDir "\..\..")
    static AppDir   := A_Temp "\ALTRunScreenshots"
    static OutDir   := ""
    static DemoDir  := ""
    static Pid      := 0
    static Failures := 0

    ; 场景名 -> [主题, 启动参数, 函数]
    static Scenes() {
        return [
            ["search",      "Light", "", () => Shots.Search("re")],
            ["pinyin",      "Light", "", () => Shots.Search("jsb")],
            ["actions",     "Light", "", () => Shots.Actions("website")],
            ["files",       "Light", "", () => Shots.FileMode("report")],
            ["calculator",  "Light", "", () => Shots.Search("(1200+350)*2.5")],
            ["clipboard",   "Light", "", () => Shots.Clipboard()],
            ["websearch",   "Light", "", () => Shots.Search("g autohotkey v2 hotkeys")],
            ["system",      "Light", "", () => Shots.Search("lock")],
            ["hud",         "Dark",  "", () => Shots.Hud("12*3")],
            ["prefs-general",    "Light", "-Preferences 1", () => Shots.Preferences()],
            ["prefs-appearance", "Light", "-Preferences 3", () => Shots.Preferences()],
            ["prefs-commands",   "Light", "-Preferences 7", () => Shots.Preferences()],
            ["theme-dark",      "Dark",     "", () => Shots.Search("re")],
            ["theme-darkcompact", "DarkCompact", "", () => Shots.Search("re")],
            ["theme-classic",   "Classic",  "", () => Shots.Search("re")],
            ["theme-midnight",  "Midnight", "", () => Shots.Search("re")],
            ["theme-frost",     "Frost",    "", () => Shots.Search("re")],
            ["theme-graphite",  "Graphite", "", () => Shots.Search("re")],
            ["theme-ocean",     "Ocean",    "", () => Shots.Search("re")],
            ["theme-paper",     "Paper",    "", () => Shots.Search("re")]
        ]
    }

    static Run(args) {
        Shots.OutDir := (args.Length >= 1) ? Shots._FullPath(args[1]) : Shots.RepoDir "\docs\images\screenshots"
        wanted := Map()
        Loop args.Length - 1
            wanted[args[A_Index + 1]] := true
        DirCreate(Shots.OutDir)
        Shots.PrepareApp()
        Shots.PrepareDemoFiles()
        Shots.ShowBackdrop()
        for scene in Shots.Scenes() {
            if (wanted.Count && !wanted.Has(scene[1]))
                continue
            Shots.Log("== " scene[1])
            ; 在 GitHub 刚启动的 Windows 上, 最开始的几次启动 ALTRun 有时画不出结果行
            ; (列表里有项目, 窗口没有卡住, 重画也没用; 原因未查明), 这时重新启动 ALTRun 再试
            Loop 4 {
                try {
                    Shots.Launch(scene[2], scene[3])
                    hwnd := scene[4]()
                    if (!Shots.WaitPainted(hwnd, 20) && A_Index < 4) {
                        Shots.Diagnose(hwnd)
                        Shots.Log("relaunching")
                        Shots.Close()
                        continue
                    }
                    Shots.Capture(hwnd, Shots.OutDir "\" scene[1] ".png")
                } catch as e {
                    Shots.Failures += 1
                    Shots.Log("FAIL " scene[1] ": " e.Message " (line " e.Line ")")
                }
                Shots.Close()
                break
            }
        }
        Shots.Log("done, " Shots.Failures " failure(s)")
    }

    ;---------------------------------------------------------------------------
    ; 准备
    ;---------------------------------------------------------------------------
    static PrepareApp() {
        try DirDelete(Shots.AppDir, true)
        DirCreate(Shots.AppDir)
        FileCopy(Shots.RepoDir "\ALTRun.ahk", Shots.AppDir "\ALTRun.ahk")
        for folder in ["Lib", "Src", "Resources"]
            DirCopy(Shots.RepoDir "\" folder, Shots.AppDir "\" folder)
    }

    ; 通用的示例文件夹
    static PrepareDemoFiles() {
        drive := "C:"
        Shots.DemoDir := drive "\Demo Projects"
        try DirDelete(Shots.DemoDir, true)
        files := [
            "Website Redesign\Design\Homepage Mockup.png",
            "Website Redesign\Design\Style Guide.pdf",
            "Website Redesign\Notes\Kickoff Meeting Notes.docx",
            "Website Redesign\Reports\Usability Test Report.pdf",
            "Annual Report 2026\Annual Report 2026 Draft.docx",
            "Annual Report 2026\Budget 2026.xlsx",
            "Annual Report 2026\Report Charts.pptx",
            "Annual Report 2026\Quarterly Report Q3.pdf",
            "Travel\Tokyo Trip Itinerary.pdf"
        ]
        for relative in files {
            SplitPath(Shots.DemoDir "\" relative, , &dir)
            DirCreate(dir)
            FileAppend("", Shots.DemoDir "\" relative)
        }
        ; 应用: 只索引这里的快捷方式, 结果和机器上装了什么无关
        ; (记事本、命令提示符等 Windows 工具是 ALTRun 内置的系统命令, 不用再建)
        appDir := drive "\ALTRun Demo Apps"
        try DirDelete(appDir, true)
        DirCreate(appDir)
        ; 中文名称的记事本: 演示拼音首字母搜索 (jsb)
        for shortcut in [["Remote Desktop Connection", A_WinDir "\System32\mstsc.exe"],
                         ["记事本", A_WinDir "\System32\notepad.exe"],
                         ["Microsoft Edge", A_ProgramFiles " (x86)\Microsoft\Edge\Application\msedge.exe"]]
            if FileExist(shortcut[2])
                FileCreateShortcut(shortcut[2], appDir "\" shortcut[1] ".lnk")
        Shots.AppsDir := appDir
    }
    static AppsDir := ""

    static WriteSettings(theme) {
        demo := Shots.DemoDir
        settings := Map(
            "SchemaVersion", 4,
            "General", Map("Language", "en", "LaunchAtLogin", 0, "HideOnDeactivate", 0, "SendToMenu", 0,
                           "StartMenuShortcut", 0, "CheckForUpdates", 0, "Hotkey", "!Space"),
            "Appearance", Map("Theme", theme, "Width", 700, "VisibleRows", 8),
            "Features", Map(
                "Applications", Map("Folders", [Shots.AppsDir], "StoreApps", 0),
                "Calculator", Map("StructuralCalc", 0),
                "FileSearch", Map("UseEverything", 0, "ScopeFolders", [demo], "InDefaultResults", 0)
            ),
            "CustomCommands", [
                Map("Title", "Website Redesign", "Type", "Folder", "Target", demo "\Website Redesign", "Arguments", "", "Keyword", ""),
                Map("Title", "Annual Report 2026", "Type", "Folder", "Target", demo "\Annual Report 2026", "Arguments", "", "Keyword", ""),
                Map("Title", "Demo Projects", "Type", "Folder", "Target", demo, "Arguments", "", "Keyword", "proj"),
                Map("Title", "Budget 2026", "Type", "File", "Target", demo "\Annual Report 2026\Budget 2026.xlsx", "Arguments", "", "Keyword", "budget"),
                Map("Title", "IP Configuration", "Type", "Command", "Target", "cmd.exe", "Arguments", "/k ipconfig /all", "Keyword", "ip"),
                Map("Title", "ALTRun on GitHub", "Type", "Url", "Target", "https://github.com/zhugecaomao/ALTRun", "Arguments", "", "Keyword", "altrun")
            ]
        )
        try DirDelete(Shots.AppDir "\Data", true)                           ; 每个场景从空的索引和学习记录开始
        try FileDelete(Shots.AppDir "\ALTRun.json")                         ; 旧版本的位置
        DirCreate(Shots.AppDir "\Data")
        FileAppend(JSON.Stringify(settings, 4), Shots.AppDir "\Data\ALTRun.json", "UTF-8")
    }

    ; 纯色背景铺满屏幕, 挡住桌面上的其它窗口 (半透明主题会透出后面的内容)。
    ; 置顶才能盖住控制台窗口; 搜索窗口也是置顶的, 后显示的在上面
    static ShowBackdrop() {
        backdrop := Gui("-Caption +ToolWindow +AlwaysOnTop -DPIScale")
        backdrop.BackColor := "8A9BB0"
        backdrop.Show("NA x0 y0 w" A_ScreenWidth " h" A_ScreenHeight)
        Shots.Backdrop := backdrop
    }
    static Backdrop := ""

    ;---------------------------------------------------------------------------
    ; 启动 / 关闭
    ;---------------------------------------------------------------------------
    static Launch(theme, args := "") {
        Shots.WriteSettings(theme)
        Run('"' A_AhkPath '" "' Shots.AppDir '\ALTRun.ahk" ' args, Shots.AppDir, , &pid)
        Shots.Pid := pid
        ; 启动时屏幕上方的 "ALTRun 已在运行" 提示 3 秒后消失
        Sleep(4500)
    }

    static Close() {
        if Shots.Pid
            try ProcessClose(Shots.Pid)
        ProcessWaitClose(Shots.Pid, 5)
        Shots.Pid := 0
        Sleep(500)
    }

    static SearchWindow() {
        hwnd := WinWait("ALTRun ahk_class AutoHotkeyGUI ahk_pid " Shots.Pid, , 10)
        if !hwnd
            throw Error("search window not found")
        return hwnd
    }

    static SetQuery(hwnd, text) {
        ControlSetText(text, "Edit1", hwnd)
        ControlSend("{End}", "Edit1", hwnd)
        Sleep(2000)                                                         ; 搜索 + 图标异步载入
    }

    ;---------------------------------------------------------------------------
    ; 场景
    ;---------------------------------------------------------------------------
    static Search(text) {
        hwnd := Shots.SearchWindow()
        Shots.SetQuery(hwnd, text)
        return hwnd
    }

    static Actions(text) {
        hwnd := Shots.Search(text)
        ControlSend("{Right}", "Edit1", hwnd)
        Sleep(1500)
        return hwnd
    }

    static FileMode(text) {
        hwnd := Shots.SearchWindow()
        Sleep(4000)                                                         ; 内置文件索引
        ControlSetText("", "Edit1", hwnd)
        ControlSend("{Space}", "Edit1", hwnd)
        Sleep(300)
        ControlSend("{Text}" text, "Edit1", hwnd)
        Sleep(2000)
        return hwnd
    }

    ; 剪贴板历史记录复制时的前台程序: 从记事本复制, 而不是截图脚本自己
    static Clipboard() {
        hwnd := Shots.SearchWindow()
        Run("notepad.exe", , , &notepad)
        WinWait("ahk_pid " notepad, , 10)
        WinActivate("ahk_pid " notepad)
        Sleep(500)
        for text in ["https://github.com/zhugecaomao/ALTRun",
                     "The meeting is moved to Thursday at 3 pm.",
                     "C:\Demo Projects\Annual Report 2026",
                     "Thanks, the new homepage looks great!",
                     "SELECT name, total FROM orders WHERE total > 100"] {
            A_Clipboard := text
            Sleep(1000)
        }
        ProcessClose(notepad)
        WinActivate(hwnd)
        Shots.SetQuery(hwnd, "clip ")
        return hwnd
    }

    ; 操作后的提示 (HUD): 回车复制计算结果, 搜索窗口关闭, 提示显示在屏幕中间偏下
    static Hud(text) {
        hwnd := Shots.Search(text)
        ControlSend("{Enter}", "Edit1", hwnd)
        hud := WinWait("ALTRun HUD ahk_class AutoHotkeyGUI ahk_pid " Shots.Pid, , 3)
        if !hud
            throw Error("HUD not shown")
        Sleep(300)                                                          ; 淡入
        return hud
    }

    static Preferences() {
        hwnd := WinWait("ahk_class AutoHotkeyGUI ahk_pid " Shots.Pid, , 10)
        if !hwnd
            throw Error("preferences window not found")
        WinSetAlwaysOnTop(1, hwnd)                                          ; 在背景之上
        WinActivate(hwnd)
        Sleep(1500)
        return hwnd
    }

    ;---------------------------------------------------------------------------
    ; 截图: 从屏幕复制窗口区域 (包括半透明效果), 用 GDI+ 存成 PNG
    ;---------------------------------------------------------------------------
    ; 等到列表区域不再是一片空白 (窗口忙, 还没画出结果), 最多等 seconds 秒
    static WaitPainted(hwnd, seconds) {
        start := A_TickCount
        Loop {
            DllCall("RedrawWindow", "Ptr", hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x185)    ; INVALIDATE | ERASE | ALLCHILDREN | UPDATENOW
            Sleep(500)
            hbm := Shots.CaptureBitmap(hwnd, &w, &h, &blank)
            DllCall("DeleteObject", "Ptr", hbm)
            elapsed := (A_TickCount - start) // 1000
            if !blank {
                if (elapsed > 1)
                    Shots.Log("painted after " elapsed " s")
                return true
            }
            if (elapsed >= seconds) {
                Shots.Log("still blank after " elapsed " s")
                return false
            }
            if (Mod(A_Index, 5) = 1)
                Shots.Log("waiting for the list to paint (" elapsed " s, hung: " DllCall("IsHungAppWindow", "Ptr", hwnd) ")")
            Sleep(1500)
        }
    }

    static Capture(hwnd, file) {
        hbm := Shots.CaptureBitmap(hwnd, &w, &h, &blank)
        try {
            Shots.SavePng(hbm, file)
        } finally {
            DllCall("DeleteObject", "Ptr", hbm)
        }
        Shots.Log("saved " file " (" w "x" h ")")
    }

    ; 列表画不出来时, 把 ALTRun 的所有窗口 (包括错误对话框的内容) 和列表状态写到日志
    static Diagnose(hwnd) {
        try Shots.Log("list: items=" SendMessage(0x1004, 0, 0, "SysListView321", hwnd) " visible=" ControlGetVisible("SysListView321", hwnd))
        DetectHiddenWindows(false)
        for window in WinGetList("ahk_pid " Shots.Pid) {
            Shots.Log("window: [" WinGetClass(window) "] " WinGetTitle(window))
            if (WinGetClass(window) = "#32770")
                Shots.Log(WinGetText(window))
        }
    }

    static CaptureBitmap(hwnd, &w, &h, &blank) {
        rect := Buffer(16, 0)
        if (DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", hwnd, "UInt", 9, "Ptr", rect, "UInt", 16) = 0 && NumGet(rect, 8, "Int") > NumGet(rect, 0, "Int")) {
            x := NumGet(rect, 0, "Int"), y := NumGet(rect, 4, "Int")         ; DWMWA_EXTENDED_FRAME_BOUNDS: 不含看不见的边框
            w := NumGet(rect, 8, "Int") - x, h := NumGet(rect, 12, "Int") - y
        } else {
            WinGetPos(&x, &y, &w, &h, hwnd)
        }
        hdcScreen := DllCall("GetDC", "Ptr", 0, "Ptr")
        hdcMem := DllCall("CreateCompatibleDC", "Ptr", hdcScreen, "Ptr")
        hbm := DllCall("CreateCompatibleBitmap", "Ptr", hdcScreen, "Int", w, "Int", h, "Ptr")
        old := DllCall("SelectObject", "Ptr", hdcMem, "Ptr", hbm, "Ptr")
        DllCall("BitBlt", "Ptr", hdcMem, "Int", 0, "Int", 0, "Int", w, "Int", h, "Ptr", hdcScreen, "Int", x, "Int", y, "UInt", 0x40CC0020)
        blank := Shots._ListIsBlank(hdcMem, w, h)
        DllCall("SelectObject", "Ptr", hdcMem, "Ptr", old)
        DllCall("DeleteDC", "Ptr", hdcMem)
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdcScreen)
        return hbm
    }

    ; 搜索窗口: 输入框以下 (约 80 像素起) 每隔几个像素取一个点, 全都一样 = 结果还没画出来
    static _ListIsBlank(hdc, w, h) {
        if (h < 120)
            return false
        first := DllCall("GetPixel", "Ptr", hdc, "Int", 60, "Int", 80, "UInt")
        y := 80
        while (y < h - 4) {
            x := 20
            while (x < w - 20) {
                if (DllCall("GetPixel", "Ptr", hdc, "Int", x, "Int", y, "UInt") != first)
                    return false
                x += 7
            }
            y += 3
        }
        return true
    }

    static SavePng(hbm, file) {
        static token := 0
        if !token {
            DllCall("LoadLibrary", "Str", "gdiplus")
            input := Buffer(24, 0)
            NumPut("UInt", 1, input)
            DllCall("gdiplus\GdiplusStartup", "Ptr*", &token, "Ptr", input, "Ptr", 0)
        }
        bitmap := 0
        DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "Ptr", hbm, "Ptr", 0, "Ptr*", &bitmap)
        clsid := Buffer(16)
        DllCall("ole32\CLSIDFromString", "Str", "{557CF406-1A04-11D3-9A73-0000F81EF32E}", "Ptr", clsid)   ; PNG encoder
        try FileDelete(file)
        status := DllCall("gdiplus\GdipSaveImageToFile", "Ptr", bitmap, "WStr", file, "Ptr", clsid, "Ptr", 0)
        DllCall("gdiplus\GdipDisposeImage", "Ptr", bitmap)
        if status
            throw Error("GdipSaveImageToFile failed: " status)
    }

    static Log(text) {
        FileAppend(text "`n", "*", "UTF-8")
    }

    static _FullPath(path) {
        size := DllCall("GetFullPathName", "Str", path, "UInt", 0, "Ptr", 0, "Ptr", 0)
        buf := Buffer(size * 2)
        DllCall("GetFullPathName", "Str", path, "UInt", size, "Ptr", buf, "Ptr", 0)
        return StrGet(buf)
    }
}
