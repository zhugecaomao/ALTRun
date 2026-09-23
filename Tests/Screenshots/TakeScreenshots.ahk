;===============================================================================
; TakeScreenshots.ahk - 自动生成 README / Wiki 用的界面截图 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 在一个临时文件夹里准备好演示用的 ALTRun (示例设置、自定义命令、应用快捷方式、项目文件),
; 逐个场景启动 ALTRun, 输入搜索内容, 把窗口截成 PNG。不改动你自己的 ALTRun.json。
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
            ["search",      "Light", "", () => Shots.Search("pt")],
            ["pinyin",      "Light", "", () => Shots.Search("jsb")],
            ["actions",     "Light", "", () => Shots.Actions("riverside")],
            ["files",       "Light", "", () => Shots.FileMode("report")],
            ["calculator",  "Light", "", () => Shots.Search("(1200+350)*2.5")],
            ["clipboard",   "Light", "", () => Shots.Clipboard()],
            ["websearch",   "Light", "", () => Shots.Search("g 后张预应力 楼板")],
            ["system",      "Light", "", () => Shots.Search("锁")],
            ["prefs-general",    "Light", "-Preferences 1", () => Shots.Preferences()],
            ["prefs-appearance", "Light", "-Preferences 2", () => Shots.Preferences()],
            ["prefs-commands",   "Light", "-Preferences 6", () => Shots.Preferences()],
            ["theme-dark",      "Dark",     "", () => Shots.Search("pt")],
            ["theme-classic",   "Classic",  "", () => Shots.Search("pt")],
            ["theme-midnight",  "Midnight", "", () => Shots.Search("pt")],
            ["theme-frost",     "Frost",    "", () => Shots.Search("pt")],
            ["theme-graphite",  "Graphite", "", () => Shots.Search("pt")],
            ["theme-ocean",     "Ocean",    "", () => Shots.Search("pt")],
            ["theme-paper",     "Paper",    "", () => Shots.Search("pt")]
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
        ; 预热: 刚建好的快捷方式第一次取图标很慢 (Windows 在扫描新文件), 这期间窗口画不出结果。
        ; 等它第一次画出来, 这一次不截图
        Shots.Log("== warm-up")
        Shots.Launch("Light")
        Shots.WaitPainted(Shots.Search("pt"), 180)
        Shots.Close()
        for scene in Shots.Scenes() {
            if (wanted.Count && !wanted.Has(scene[1]))
                continue
            Shots.Log("== " scene[1])
            try {
                Shots.Launch(scene[2], scene[3])
                hwnd := scene[4]()
                Shots.Capture(hwnd, Shots.OutDir "\" scene[1] ".png")
            } catch as e {
                Shots.Failures += 1
                Shots.Log("FAIL " scene[1] ": " e.Message " (line " e.Line ")")
            }
            Shots.Close()
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

    ; 结构设计项目的示例文件夹
    static PrepareDemoFiles() {
        drive := "C:"
        Shots.DemoDir := drive "\Design Projects"
        try DirDelete(Shots.DemoDir, true)
        files := [
            "PT2415 - Riverside Tower\Drawings\PT2415-S-101 Level 3 PT Layout.dwg",
            "PT2415 - Riverside Tower\Drawings\PT2415-S-102 Level 4 PT Layout.dwg",
            "PT2415 - Riverside Tower\Calculations\PT2415 Transfer Beam Design.xlsx",
            "PT2415 - Riverside Tower\Reports\PT2415 Design Report.docx",
            "PT2415 - Riverside Tower\Reports\PT2415 Slab Deflection Report.pdf",
            "PT2421 - Harbour Carpark\Reports\PT2421 Tender Report.pdf",
            "PT2421 - Harbour Carpark\Reports\PT2421 Site Inspection Report.docx",
            "PT2421 - Harbour Carpark\Drawings\PT2421-S-201 Ramp PT Layout.dwg",
            "Standards\Eurocode 2 - Design of Concrete Structures.pdf"
        ]
        for relative in files {
            SplitPath(Shots.DemoDir "\" relative, , &dir)
            DirCreate(dir)
            FileAppend("", Shots.DemoDir "\" relative)
        }
        ; 应用: 只索引这里的快捷方式, 结果和机器上装了什么无关
        appDir := drive "\ALTRun Demo Apps"
        try DirDelete(appDir, true)
        DirCreate(appDir)
        system := A_WinDir "\System32\"
        for shortcut in [["记事本", "notepad.exe"], ["计算器", "calc.exe"], ["画图", "mspaint.exe"], ["命令提示符", "cmd.exe"],
                         ["任务管理器", "taskmgr.exe"], ["控制面板", "control.exe"], ["注册表编辑器", "regedit.exe"],
                         ["Windows PowerShell", "WindowsPowerShell\v1.0\powershell.exe"], ["远程桌面连接", "mstsc.exe"]] {
            target := FileExist(system shortcut[2]) ? system shortcut[2] : A_WinDir "\" shortcut[2]
            if !FileExist(target)
                target := A_WinDir "\explorer.exe"
            FileCreateShortcut(target, appDir "\" shortcut[1] ".lnk")
        }
        Shots.AppsDir := appDir
    }
    static AppsDir := ""

    static WriteSettings(theme) {
        demo := Shots.DemoDir
        settings := Map(
            "SchemaVersion", 4,
            "General", Map("Language", "zh", "LaunchAtLogin", 0, "HideOnDeactivate", 0, "SendToMenu", 0,
                           "StartMenuShortcut", 0, "CheckForUpdates", 0, "Hotkey", "!Space"),
            "Appearance", Map("Theme", theme, "Width", 700, "VisibleRows", 8),
            "Features", Map(
                "Applications", Map("Folders", [Shots.AppsDir], "StoreApps", 0),
                "Calculator", Map("StructuralCalc", 1),
                "FileSearch", Map("UseEverything", 0, "ScopeFolders", [demo], "InDefaultResults", 0)
            ),
            "CustomCommands", [
                Map("Title", "PT2415 - Riverside Tower", "Type", "Folder", "Target", demo "\PT2415 - Riverside Tower", "Arguments", "", "Keyword", ""),
                Map("Title", "PT2421 - Harbour Carpark", "Type", "Folder", "Target", demo "\PT2421 - Harbour Carpark", "Arguments", "", "Keyword", ""),
                Map("Title", "项目资料", "Type", "Folder", "Target", demo, "Arguments", "", "Keyword", "proj"),
                Map("Title", "Eurocode 2 设计规范", "Type", "File", "Target", demo "\Standards\Eurocode 2 - Design of Concrete Structures.pdf", "Arguments", "", "Keyword", "ec2"),
                Map("Title", "IP 配置", "Type", "Command", "Target", "cmd.exe", "Arguments", "/k ipconfig /all", "Keyword", "ip"),
                Map("Title", "ALTRun 项目主页", "Type", "Url", "Target", "https://github.com/zhugecaomao/ALTRun", "Arguments", "", "Keyword", "altrun")
            ]
        )
        path := Shots.AppDir "\ALTRun.json"
        try FileDelete(path)
        FileAppend(JSON.Stringify(settings, 4), path, "UTF-8")
        try DirDelete(Shots.AppDir "\Data", true)
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
                     "PT2415 转换梁 TB-3: 1200 x 2000, 预应力 4 x 19 束",
                     "C:\Design Projects\PT2415 - Riverside Tower\Calculations",
                     "Please find attached the revised PT layout for Level 3.",
                     "fck = 40 MPa, fpk = 1860 MPa"] {
            A_Clipboard := text
            Sleep(1000)
        }
        ProcessClose(notepad)
        WinActivate(hwnd)
        Shots.SetQuery(hwnd, "clip ")
        return hwnd
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
        Shots.WaitPainted(hwnd, 60)
        hbm := Shots.CaptureBitmap(hwnd, &w, &h, &blank)
        try {
            Shots.SavePng(hbm, file)
        } finally {
            DllCall("DeleteObject", "Ptr", hbm)
        }
        Shots.Log("saved " file " (" w "x" h ")")
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
