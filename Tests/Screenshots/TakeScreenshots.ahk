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
; "demo" 场景不截单张图, 而是把一段操作的每一帧存到 %Temp%\ALTRunDemoFrames (frames.txt 每行: 文件名 Tab 停留的毫秒数
; Tab 这一步按的键),
; 由 MakeDemo.py 合成 README 首页的动图 demo.png。截图的圆角由 RoundCorners.py 加上; 截图前 SetDisplay.ps1 把屏幕调成
; 1920 x 1080、150% 缩放 (见 screenshots.yml), 下面的坐标都按 Shots.Scale 放大。
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
    static AppDir   := "C:\ALTRun"                                   ; 截图里显示的路径 (偏好设置 → 脚本) 简单好看
    static OutDir   := ""
    static DemoDir  := ""
    static Pid      := 0
    static Failures := 0
    static Scale    := A_ScreenDPI / 96                             ; 屏幕缩放 (150% = 1.5)

    ; 场景名 -> [主题, 启动参数, 函数, 启动前准备数据的函数 (可选)]
    static Scenes() {
        return [
            ["search",      "Light", "", () => Shots.Search("re")],
            ["pinyin",      "Light", "", () => Shots.Search("jsb")],
            ["actions",     "Light", "", () => Shots.Actions("website")],
            ["files",       "Light", "", () => Shots.FileMode("report")],
            ["calculator",  "Light", "", () => Shots.Search("(200+100)*2")],     ; 600: 附带结构计算的两行
            ["clipboard",   "Light", "", () => Shots.Search("clip "), () => Shots.WriteClipboard()],
            ["empty",       "Light", "", () => Shots.EmptyBox(), () => Shots.WritePinnedAndRecent()],
            ["browse",      "Light", "", () => Shots.Search(Shots.DemoDir "\Annual Report 2026\")],
            ["websearch",   "Light", "", () => Shots.Search("g autohotkey v2 hotkeys")],
            ["system",      "Light", "", () => Shots.Search("lock")],
            ["hud",         "Dark",  "", () => Shots.Hud("12*3")],
            ["quickswitch", "Light", "", () => Shots.QuickSwitch(), () => Shots.WriteRecentFolders()],
            ["demo",        "Light", "", () => Shots.Demo(), () => Shots.PrepareDemo()],
            ["prefs-general",    "Light", "-Preferences 1", () => Shots.Preferences()],
            ["prefs-appearance", "Light", "-Preferences 3", () => Shots.Preferences()],
            ["prefs-commands",   "Light", "-Preferences 8", () => Shots.Preferences()],
            ["prefs-calculator", "Light", "-Preferences 12", () => Shots.Preferences()],
            ["prefs-scripts",    "Light", "-Preferences 13", () => Shots.Preferences()],
            ["theme-dark",      "Dark",     "", () => Shots.Search("re")],
            ["theme-darkcompact", "DarkCompact", "", () => Shots.Search("re")],
            ["theme-classic",   "Classic",  "", () => Shots.Search("re")],
            ["theme-midnight",  "Midnight", "", () => Shots.Search("re")],
            ["theme-midnightcompact", "MidnightCompact", "", () => Shots.Search("re")],
            ["theme-frost",     "Frost",    "", () => Shots.Search("re")],
            ["theme-graphite",  "Graphite", "", () => Shots.Search("re")],
            ["theme-ocean",     "Ocean",    "", () => Shots.Search("re")],
            ["theme-paper",     "Paper",    "", () => Shots.Search("re")],
            ["theme-darkmodern", "DarkModern", "", () => Shots.Search("re")],
            ["theme-lightmodern", "LightModern", "", () => Shots.Search("re")],
            ["theme-monokai", "Monokai", "", () => Shots.Search("re")],
            ["theme-onedark", "OneDark", "", () => Shots.Search("re")],
            ["theme-tokyonight", "TokyoNight", "", () => Shots.Search("re")],
            ["theme-dracula", "Dracula", "", () => Shots.Search("re")],
            ["theme-catppuccinmocha", "CatppuccinMocha", "", () => Shots.Search("re")],
            ["theme-gruvboxdark", "GruvboxDark", "", () => Shots.Search("re")],
            ["theme-solarizedlight", "SolarizedLight", "", () => Shots.Search("re")],
            ["theme-lightcompact", "LightCompact", "", () => Shots.Search("re")]
        ]
    }

    static Run(args) {
        Shots.OutDir := (args.Length >= 1) ? Shots._FullPath(args[1]) : Shots.RepoDir "\docs\images\screenshots"
        wanted := Map()
        Loop args.Length - 1
            wanted[args[A_Index + 1]] := true
        DirCreate(Shots.OutDir)
        Shots.Log("screen " A_ScreenWidth " x " A_ScreenHeight ", DPI " A_ScreenDPI)
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
                    Shots.Launch(scene[2], scene[3], (scene.Length >= 5) ? scene[5] : "")
                    hwnd := scene[4]()
                    outFile := Shots.OutDir "\" scene[1] ".png"
                    if (hwnd is Array)                                      ; 几个窗口 (对话框 + 文件夹面板) 截成一张
                        Shots.CaptureWindows(hwnd, outFile)
                    else if (hwnd != "") {                                  ; "" = demo: 每一帧已经存好
                        if (!Shots.WaitPainted(hwnd, 20) && A_Index < 4) {
                            Shots.Diagnose(hwnd)
                            Shots.Log("relaunching")
                            Shots.Close()
                            continue
                        }
                        Shots.Capture(hwnd, outFile)
                    }
                } catch as e {
                    Shots.Failures += 1
                    Shots.Log("FAIL " scene[1] ": " e.Message " (line " e.Line ")")
                }
                Shots.Close()
                Shots.CloseHelpers()
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
        ; 脚本扩展的示例 (偏好设置 → 脚本 的列表)
        DirCreate(Shots.AppDir "\Scripts")
        FileAppend("; @altrun.title    Ping host`n; @altrun.keyword  ping`n; @altrun.argument host name`n; @altrun.mode     output`n"
            . "RunWait('ping ' A_Args[1])`n", Shots.AppDir "\Scripts\Ping.ahk", "UTF-8")
        FileAppend("# @altrun.title    Today's date`n# @altrun.keyword  today`n# @altrun.mode     silent`nGet-Date -Format 'dddd, d MMMM yyyy'`n"
            , Shots.AppDir "\Scripts\Today.ps1", "UTF-8")
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
            "Annual Report 2026\Archive\Annual Report 2025.pdf",
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
                           "StartMenuShortcut", 0, "CheckForUpdates", 0, "Hotkey", "!Space", "ShowTips", 0),
            "Appearance", Map("Theme", theme, "Width", 700, "VisibleRows", 8),
            "Features", Map(
                "Applications", Map("Folders", [Shots.AppsDir], "StoreApps", 0),
                "Calculator", Map("StructuralCalc", 1),
                "Recent", Map("Pinned", Shots.Pinned, "RecentCount", 5),
                "FileSearch", Map("UseEverything", 0, "ScopeFolders", [demo], "InDefaultResults", 0)
            ),
            "Extensions", Map("QuickSwitch", Map("RecentFolders", 3)),         ; 文件夹面板只列 WriteRecentFolders 建的最近项目
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

    ; 空搜索框里置顶的项目 (只在 "empty" 场景里有, 见 WritePinnedAndRecent)
    static Pinned := []

    static Entry(title, kind, target) {
        SplitPath(target, , &dir)
        return Map("Title", title, "Subtitle", (kind = "folder") ? target : dir, "Kind", kind, "Arg", target, "Arguments", ""
                 , "Icon", "", "Uid", kind ":" StrLower(target))
    }

    ; 置顶两项 + 最近打开的三项 (Knowledge.json), 空搜索框里显示
    static WritePinnedAndRecent() {
        demo := Shots.DemoDir
        Shots.Pinned := [Shots.Entry("Annual Report 2026", "folder", demo "\Annual Report 2026")
                       , Shots.Entry("Budget 2026.xlsx", "file", demo "\Annual Report 2026\Budget 2026.xlsx")]
        recent := [Shots.Entry("Kickoff Meeting Notes.docx", "file", demo "\Website Redesign\Notes\Kickoff Meeting Notes.docx")
                 , Shots.Entry("Website Redesign", "folder", demo "\Website Redesign")
                 , Shots.Entry("Tokyo Trip Itinerary.pdf", "file", demo "\Travel\Tokyo Trip Itinerary.pdf")]
        Shots.WriteSettings("Light")
        Shots.Pinned := []
        FileAppend(JSON.Stringify(Map("Picks", Map(), "QueryPicks", Map(), "History", [], "Recent", recent)), Shots.AppDir "\Data\Knowledge.json", "UTF-8")
    }

    ; Quick Switch 的 "最近用过的文件夹": 在 Windows 的最近使用的项目里放三个示例文件的快捷方式 (比别的都新)
    static WriteRecentFolders() {
        recent := A_AppData "\Microsoft\Windows\Recent"
        DirCreate(recent)
        stamp := DateAdd(A_Now, 10, "Minutes")
        for relative in ["Travel\Tokyo Trip Itinerary.pdf", "Website Redesign\Design\Style Guide.pdf", "Annual Report 2026\Archive\Annual Report 2025.pdf"] {
            SplitPath(relative, &name)
            FileCreateShortcut(Shots.DemoDir "\" relative, recent "\" name ".lnk")
            FileSetTime(stamp := DateAdd(stamp, -1, "Minutes"), recent "\" name ".lnk")
        }
    }

    ; 动图: 剪贴板历史里有几条记录; 关掉结构计算 (计算器只显示结果, 不附带梁主筋 / 配筋面积, 一般用户看不懂)
    static PrepareDemo() {
        Shots.WriteClipboard()
        settingsFile := Shots.AppDir "\Data\ALTRun.json"
        settings := JSON.Parse(FileRead(settingsFile, "UTF-8"))
        settings["Features"]["Calculator"]["StructuralCalc"] := 0
        FileDelete(settingsFile)
        FileAppend(JSON.Stringify(settings, 4), settingsFile, "UTF-8")
    }

    ; 剪贴板历史: 几条常见的文字, 其中一条置顶
    static WriteClipboard() {
        entries := [], stamp := A_Now
        for item in [["SELECT name, total FROM orders WHERE total > 100", 0],
                     ["Thanks, the new homepage looks great!", 0],
                     ["C:\Demo Projects\Annual Report 2026", 0],
                     ["The meeting is moved to Thursday at 3 pm.", 0],
                     ["https://github.com/zhugecaomao/ALTRun", 1]] {
            entry := Map("Text", item[1], "Time", stamp := DateAdd(stamp, -3, "Minutes"), "App", "notepad.exe")
            if item[2]
                entry["Pinned"] := 1
            entries.Push(entry)
        }
        FileAppend(JSON.Stringify(Map("Entries", entries)), Shots.AppDir "\Data\ClipboardHistory.json", "UTF-8")
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
    static Launch(theme, args := "", prepare := "") {
        Shots.WriteSettings(theme)
        if IsObject(prepare)
            prepare()
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
    ; 空搜索框: 启动时打开的窗口, 等图标载入
    static EmptyBox() {
        hwnd := Shots.SearchWindow()
        Sleep(2500)
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

    ; Quick Switch: 打开两个资源管理器窗口, 另一个程序 (OpenDialog.ahk) 弹出 "打开" 对话框, 下面贴着文件夹面板。
    ; 屏幕可能只有 1024 x 768: 对话框放在上面, 面板放得下
    static QuickSwitch() {
        try ControlSend("{Esc}", "Edit1", Shots.SearchWindow())            ; 启动时打开的搜索窗口先关掉
        for folder in ["Website Redesign", "Annual Report 2026"]
            Run('explorer.exe "' Shots.DemoDir "\" folder '"')
        Sleep(3000)
        helper := Shots.AppDir "\OpenDialog.ahk"
        try FileDelete(helper)
        FileAppend('#NoTrayIcon`nFileSelect(, "' Shots.DemoDir '\Travel", "Open")`n', helper, "UTF-8")
        Run('"' A_AhkPath '" "' helper '"', , , &pid)
        Shots.HelperPid := pid
        dialog := WinWait("Open ahk_class #32770 ahk_pid " pid, , 10)
        if !dialog
            throw Error("Open dialog not shown")
        WinSetAlwaysOnTop(1, dialog)                                        ; 在背景之上
        k := Shots.Scale
        WinMove(Round(160 * k), Round(20 * k), Round(720 * k), Round(440 * k), dialog)
        WinActivate(dialog)
        panel := WinWait("ALTRun Quick Switch ahk_pid " Shots.Pid, , 5)
        if !panel
            throw Error("Quick Switch panel not shown")
        Sleep(1500)                                                         ; 面板的图标和位置
        return [dialog, panel]
    }
    static HelperPid := 0

    ; 关掉 QuickSwitch 场景打开的对话框和资源管理器窗口
    static CloseHelpers() {
        for pid in [Shots.HelperPid, Shots.NotepadPid]
            if pid
                try ProcessClose(pid)
        Shots.HelperPid := 0, Shots.NotepadPid := 0
        for window in Shots.HiddenWindows                                   ; Demo 藏起来的终端窗口
            try WinShow(window)
        Shots.HiddenWindows := []
        for window in WinGetList("ahk_class CabinetWClass")
            try WinClose(window)
    }

    ; README 首页的动图: 7 个场景 (应用、操作面板、文件、剪贴板历史、窗口切换、计算器、系统命令), 每一步把同一块屏幕区域
    ; 存成一帧, 记下这一步按的键。MakeDemo.py 再取出窗口, 放到模糊的背景上, 右下角画出按键。返回 "" (不另外截图)
    static DemoFrames := A_Temp "\ALTRunDemoFrames"
    static NotepadPid := 0, HiddenWindows := []
    static Demo() {
        hwnd := Shots.SearchWindow()
        for folder in ["Website Redesign", "Annual Report 2026"]           ; 窗口切换要有窗口可切 (都在背景后面)
            Run('explorer.exe "' Shots.DemoDir "\" folder '"')
        Run("notepad.exe", , , &pid)
        Shots.NotepadPid := pid
        ; GitHub 虚拟机上运行截图的终端窗口也在任务栏上, 窗口切换会列出它: 先藏起来, 截完再显示
        for title in ["ahk_exe WindowsTerminal.exe", "ahk_class ConsoleWindowClass", "ahk_class CASCADIA_HOSTING_WINDOW_CLASS"]
            for window in WinGetList(title)
                try WinHide(window), Shots.HiddenWindows.Push(window)
        Sleep(3000)
        WinActivate(hwnd)
        Shots.WaitPainted(hwnd, 10)
        dir := Shots.DemoFrames
        try DirDelete(dir, true)
        DirCreate(dir)
        r := Shots.FrameRect(hwnd)
        pad := Round(24 * Shots.Scale)
        region := {X: Max(0, r.X - pad), Y: Max(0, r.Y - pad), W: r.W + 2 * pad, H: Min(A_ScreenHeight - Max(0, r.Y - pad), Round(640 * Shots.Scale))}
        state := {Frames: "", Count: 0}                                     ; 内部函数改外面的变量: 通过对象传
        frame(ms, keys := "") {
            name := Format("frame-{:02}.png", ++state.Count)
            hbm := Shots.CaptureRect(region.X, region.Y, region.W, region.H)
            try Shots.SavePng(hbm, dir "\" name)
            finally DllCall("DeleteObject", "Ptr", hbm)
            state.Frames .= name "`t" ms "`t" keys "`n"
        }
        type(text, ms) {
            ControlSend("{Text}" text, "Edit1", hwnd)
            Sleep(900)
            frame(ms)
        }
        clear() {
            ControlSetText("", "Edit1", hwnd)
            Sleep(400)
        }
        frame(900, "Alt+Space")                                             ; 刚呼出的空搜索框
        type("r", 350), type("e", 350), type("p", 1500)
        ControlSend("{Right}", "Edit1", hwnd), Sleep(1200), frame(1800, "→")
        ControlSend("{Esc}", "Edit1", hwnd), Sleep(300), clear()
        ControlSend("{Space}", "Edit1", hwnd), Sleep(300)
        ControlSend("{Text}report", "Edit1", hwnd), Sleep(1200), frame(2000, "Space")
        clear(), ControlSend("{Backspace}", "Edit1", hwnd), Sleep(300)      ; 回到普通搜索
        WinActivate(hwnd), Send("^!c"), Sleep(1500), frame(2000, "Ctrl+Alt+C")   ; 剪贴板历史的快捷键
        clear(), type("w ", 2000)
        clear(), type("1200*1.09", 1800)
        clear(), type("lock", 1800)
        FileAppend(state.Frames, dir "\frames.txt", "UTF-8-RAW")
        Shots.Log("saved " state.Count " demo frames to " dir)
        return ""
    }

    static Preferences() {
        hwnd := WinWait("ahk_class AutoHotkeyGUI ahk_pid " Shots.Pid, , 10)
        if !hwnd
            throw Error("preferences window not found")
        WinSetAlwaysOnTop(1, hwnd)                                          ; 在背景之上
        r := Shots.FrameRect(hwnd)
        if (r.Y < 0)                                                        ; 150% 时窗口很高, 居中后标题栏会跑到屏幕上面
            WinMove(, 0, , , hwnd)
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

    ; 窗口在屏幕上看得见的范围 {X, Y, W, H}
    static FrameRect(hwnd) {
        rect := Buffer(16, 0)
        if (DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", hwnd, "UInt", 9, "Ptr", rect, "UInt", 16) = 0 && NumGet(rect, 8, "Int") > NumGet(rect, 0, "Int")) {
            x := NumGet(rect, 0, "Int"), y := NumGet(rect, 4, "Int")         ; DWMWA_EXTENDED_FRAME_BOUNDS: 不含看不见的边框
            return {X: x, Y: y, W: NumGet(rect, 8, "Int") - x, H: NumGet(rect, 12, "Int") - y}
        }
        WinGetPos(&x, &y, &w, &h, hwnd)
        return {X: x, Y: y, W: w, H: h}
    }

    ; 几个窗口 (例如对话框和贴着它的面板) 合在一起截成一张: 取包住它们的矩形
    static CaptureWindows(hwnds, file) {
        left := 99999, top := 99999, right := -99999, bottom := -99999
        for hwnd in hwnds {
            r := Shots.FrameRect(hwnd)
            left := Min(left, r.X), top := Min(top, r.Y), right := Max(right, r.X + r.W), bottom := Max(bottom, r.Y + r.H)
        }
        hbm := Shots.CaptureRect(left, top, right - left, bottom - top)
        try Shots.SavePng(hbm, file)
        finally DllCall("DeleteObject", "Ptr", hbm)
        Shots.Log("saved " file " (" (right - left) "x" (bottom - top) ")")
    }

    static CaptureBitmap(hwnd, &w, &h, &blank) {
        r := Shots.FrameRect(hwnd), w := r.W, h := r.H
        return Shots.CaptureRect(r.X, r.Y, w, h, &blank)
    }

    ; 从屏幕复制一块区域, 返回 HBITMAP (调用的人 DeleteObject); blank: 搜索窗口的列表区域还是一片空白
    static CaptureRect(x, y, w, h, &blank := false) {
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
        if (h < Round(120 * Shots.Scale))
            return false
        top := Round(80 * Shots.Scale)
        first := DllCall("GetPixel", "Ptr", hdc, "Int", Round(60 * Shots.Scale), "Int", top, "UInt")
        y := top
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
