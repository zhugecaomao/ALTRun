;===============================================================================
; TakeScreenshots.ahk - makes the UI screenshots for the README / Wiki (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Prepares a demo ALTRun in a temporary folder (sample settings, custom commands, application shortcuts, sample files),
; starts ALTRun for each scene, types the search, and saves the window as a PNG. Your own ALTRun.json isn't touched.
; Screenshots use the English UI, and the demo data is made-up generic content (nothing personal or work-related).
;
;   AutoHotkey64.exe Tests\Screenshots\TakeScreenshots.ahk [output folder] [scene... | -scene...]
;
; The output folder defaults to docs\images\screenshots; no scene names = all scenes; "-scene" = all but these;
; "normal-scale" / "-normal-scale" = the scenes in Shots.NormalScale / all but them.
; The "demo" scene isn't a single shot: every frame of a sequence goes to %Temp%\ALTRunDemoFrames (frames.txt, each line: file name Tab ms shown
; Tab the key pressed for this step),
; and MakeDemo.py turns them into the animation demo.png on the README front page. RoundCorners.py adds the rounded corners; before the shots SetDisplay.ps1 sets the screen to
; 1920 x 1080 at 150% scaling (see screenshots.yml), and the coordinates below scale by Shots.Scale. The Preferences and Quick Switch shots
; (Shots.NormalScale) are too big at 150%, so screenshots.yml sets the scaling back to 100% and takes them separately.
; GitHub Actions' "Screenshots" workflow runs it on Windows and commits the screenshots back to the branch.
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn All, StdOut
#Include ..\..\Lib\JSON.ahk

Shots.Run(A_Args)
ExitApp(Shots.Failures)

class Shots {
    static RepoDir  := Shots._FullPath(A_ScriptDir "\..\..")
    static AppDir   := "C:\ALTRun"                                   ; a short, tidy path shown in the screenshots (Preferences → Scripts)
    static OutDir   := ""
    static DemoDir  := ""
    static Pid      := 0
    static Failures := 0
    static Scale    := A_ScreenDPI / 96                             ; screen scaling (150% = 1.5)

    static NormalScale := ["quickswitch", "sendto", "prefs-general", "prefs-appearance", "prefs-searchsources", "prefs-commands", "prefs-snippets", "prefs-calculator", "prefs-scripts"
                         , "prefs-keyboard", "prefs-hotkeys", "prefs-quickswitch", "prefs-dialogpanel", "prefs-datestamp"]   ; scenes shot at 100% (see screenshots.yml)

    ; scene name -> [theme, start arguments, function, function preparing data before start-up (optional)]
    static Scenes() {
        return [
            ["search",      "Light", "", () => Shots.Search("re")],
            ["pinyin",      "Light", "", () => Shots.Search("jsb")],
            ["actions",     "Light", "", () => Shots.Actions("website")],
            ["files",       "Light", "", () => Shots.FileMode("report")],
            ["filetype",    "Light", "", () => Shots.TypeSearch("doc report")],   ; search by type: documents only (doc in TypeFilters)
            ["calculator",  "Light", "", () => Shots.Search("(200+100)*2")],     ; 600: with the two rows of structural calculation
            ["clipboard",   "Light", "", () => Shots.Search("clip "), () => Shots.WriteClipboard()],
            ["empty",       "Light", "", () => Shots.EmptyBox(), () => Shots.WritePinnedAndRecent()],
            ["browse",      "Light", "", () => Shots.Search(Shots.DemoDir "\Annual Report 2026\")],
            ["websearch",   "Light", "", () => Shots.Search("g autohotkey v2 hotkeys")],
            ["system",      "Light", "", () => Shots.Search("lock")],
            ["hud",         "Dark",  "", () => Shots.Hud("12*3")],
            ["quickswitch", "Light", "", () => Shots.QuickSwitch(), () => Shots.WriteRecentFolders()],
            ["windows",     "Light", "", () => Shots.Windows()],                 ; window switching (w + Space)
            ["sendto",      "Light", "", () => Shots.SendTo(), () => Shots.EnableSendTo()],
            ["demo",        "Light", "", () => Shots.Demo(), () => Shots.PrepareDemo()],
            ["prefs-general",    "Light", "-Preferences 1", () => Shots.Preferences()],
            ["prefs-appearance", "Light", "-Preferences 4", () => Shots.Preferences()],
            ["prefs-commands",   "Light", "-Preferences 9", () => Shots.Preferences()],
            ["prefs-snippets",   "Light", "-Preferences 10", () => Shots.Preferences()],
            ["prefs-calculator", "Light", "-Preferences 13", () => Shots.Preferences()],
            ["prefs-scripts",    "Light", "-Preferences 14", () => Shots.Preferences()],
            ["prefs-searchsources", "Light", "-Preferences 8", () => Shots.Preferences()],   ; Wiki: file search
            ["prefs-keyboard",   "Light", "-Preferences 3", () => Shots.Preferences()],      ; Wiki: Custom Hotkeys (Caps Lock)
            ["prefs-hotkeys",    "Light", "-Preferences 15", () => Shots.Preferences()],     ; Wiki: custom hotkeys
            ["prefs-quickswitch", "Light", "-Preferences 16", () => Shots.Preferences()],    ; Wiki: quick switch for dialogs
            ["prefs-dialogpanel", "Light", "-Preferences 17", () => Shots.Preferences()],    ; Wiki: the dialogs' folder panel
            ["prefs-datestamp",  "Light", "-Preferences 18", () => Shots.Preferences()],     ; Wiki: date stamp
            ["theme-dark",      "Dark",     "", () => Shots.Search("re")],
            ["theme-compact", "Dark|Compact", "", () => Shots.Search("re")],   ; dark + density "Compact" (all other shots use the new install default "Normal")
            ["theme-classic",   "Classic",  "", () => Shots.Search("re")],
            ["theme-midnight",  "Midnight", "", () => Shots.Search("re")],
            ["theme-frost",     "Frost",    "", () => Shots.Search("re")],
            ["theme-graphite",  "Graphite", "", () => Shots.Search("re")],
            ["theme-ocean",     "Ocean",    "", () => Shots.Search("re")],
            ["theme-paper",     "Paper",    "", () => Shots.Search("re")],
            ["theme-lavender",  "Lavender", "", () => Shots.Search("re")],
            ["theme-rose",      "Rose",     "", () => Shots.Search("re")],
            ["theme-forest",    "Forest",   "", () => Shots.Search("re")],
            ["theme-sunset",    "Sunset",   "", () => Shots.Search("re")],
            ["theme-highcontrast", "HighContrast", "", () => Shots.Search("re")],
            ["theme-darkmodern", "DarkModern", "", () => Shots.Search("re")],
            ["theme-lightmodern", "LightModern", "", () => Shots.Search("re")],
            ["theme-monokai", "Monokai", "", () => Shots.Search("re")],
            ["theme-onedark", "OneDark", "", () => Shots.Search("re")],
            ["theme-tokyonight", "TokyoNight", "", () => Shots.Search("re")],
            ["theme-dracula", "Dracula", "", () => Shots.Search("re")],
            ["theme-catppuccinmocha", "CatppuccinMocha", "", () => Shots.Search("re")],
            ["theme-gruvboxdark", "GruvboxDark", "", () => Shots.Search("re")],
            ["theme-solarizedlight", "SolarizedLight", "", () => Shots.Search("re")]
        ]
    }

    static Run(args) {
        Shots.OutDir := (args.Length >= 1) ? Shots._FullPath(args[1]) : Shots.RepoDir "\docs\images\screenshots"
        wanted := Map(), skipped := Map()
        Loop args.Length - 1 {
            name := args[A_Index + 1]
            target := (SubStr(name, 1, 1) = "-") ? skipped : wanted
            name := LTrim(name, "-")
            for scene in (name = "normal-scale") ? Shots.NormalScale : [name]
                target[scene] := true
        }
        DirCreate(Shots.OutDir)
        Shots.Log("screen " A_ScreenWidth " x " A_ScreenHeight ", DPI " A_ScreenDPI)
        Shots.PrepareApp()
        Shots.PrepareDemoFiles()
        Shots.ShowBackdrop()
        for scene in Shots.Scenes() {
            if ((wanted.Count && !wanted.Has(scene[1])) || skipped.Has(scene[1]))
                continue
            Shots.Log("== " scene[1])
            ; When the result rows stay blank, restart ALTRun and try again (a safeguard). The former cause: a timer slipping in while a row was painted asked for a repaint, and the list kept waiting for it
            ; (see ResultRows.OnCustomDraw; the custom-draw callback is no longer interruptible)
            Loop 4 {
                try {
                    Shots.Launch(scene[2], scene[3], (scene.Length >= 5) ? scene[5] : "")
                    hwnd := scene[4]()
                    outFile := Shots.OutDir "\" scene[1] ".png"
                    if (hwnd is Array)                                      ; several windows (dialog + folder panel) in one shot
                        Shots.CaptureWindows(hwnd, outFile)
                    else if (hwnd != "") {                                  ; "" = already saved (demo: every frame, sendto: Explorer and its menus)
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
    ; Preparation
    ;---------------------------------------------------------------------------
    static PrepareApp() {
        try DirDelete(Shots.AppDir, true)
        DirCreate(Shots.AppDir)
        FileCopy(Shots.RepoDir "\ALTRun.ahk", Shots.AppDir "\ALTRun.ahk")
        for folder in ["Lib", "Src", "Resources"]
            DirCopy(Shots.RepoDir "\" folder, Shots.AppDir "\" folder)
        ; examples for the Scripts extension (the list in Preferences → Scripts)
        DirCreate(Shots.AppDir "\Scripts")
        FileAppend("; @altrun.title    Ping host`n; @altrun.keyword  ping`n; @altrun.argument host name`n; @altrun.mode     output`n"
            . "RunWait('ping ' A_Args[1])`n", Shots.AppDir "\Scripts\Ping.ahk", "UTF-8")
        FileAppend("# @altrun.title    Today's date`n# @altrun.keyword  today`n# @altrun.mode     silent`nGet-Date -Format 'dddd, d MMMM yyyy'`n"
            , Shots.AppDir "\Scripts\Today.ps1", "UTF-8")
    }

    ; generic sample folders
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
        ; Applications: only the shortcuts here are indexed, so the results don't depend on what the machine has installed
        ; (Notepad, Command Prompt and other Windows tools are ALTRun's built-in system commands, no need to create them)
        appDir := drive "\ALTRun Demo Apps"
        try DirDelete(appDir, true)
        DirCreate(appDir)
        ; Notepad with a Chinese name: demonstrates pinyin initial search (jsb)
        for shortcut in [["Remote Desktop Connection", A_WinDir "\System32\mstsc.exe"],
                         ["记事本", A_WinDir "\System32\notepad.exe"],
                         ["Microsoft Edge", A_ProgramFiles " (x86)\Microsoft\Edge\Application\msedge.exe"]]
            if FileExist(shortcut[2])
                FileCreateShortcut(shortcut[2], appDir "\" shortcut[1] ".lnk")
        Shots.AppsDir := appDir
    }
    static AppsDir := ""

    ; The current settings version (AppSettings.CurrentVersion). With an older version ALTRun upgrades the settings at start-up and doesn't show the search window while the upgrade notice is up,
    ; so no scene would find the window (that is what happened after the settings version reached 5)
    static SchemaVersion() {
        return RegExMatch(FileRead(Shots.RepoDir "\Src\Core\AppSettings.ahk", "UTF-8"), "static CurrentVersion\s*:=\s*(\d+)", &m) ? Integer(m[1]) : 5
    }

    ; theme: theme name, optionally with "|density" (e.g. "Dark|Compact"); the new install default "Normal" if not given
    static WriteSettings(theme) {
        demo := Shots.DemoDir
        parts := StrSplit(theme, "|"), theme := parts[1], density := (parts.Length >= 2) ? parts[2] : "Normal"
        settings := Map(
            "SchemaVersion", Shots.SchemaVersion(),
            "General", Map("Language", "en", "LaunchAtLogin", 0, "HideOnDeactivate", 0, "SendToMenu", 0,
                           "StartMenuShortcut", 0, "CheckForUpdates", 0, "Hotkey", "!Space", "ShowTips", 0),
            "Appearance", Map("Theme", theme, "Density", density, "Width", 700, "VisibleRows", 8),
            "Features", Map(
                "Applications", Map("Folders", [Shots.AppsDir], "StoreApps", 0),
                "Calculator", Map("StructuralCalc", 1),
                "Recent", Map("Pinned", Shots.Pinned, "RecentCount", 5),
                "FileSearch", Map("UseEverything", 0, "ScopeFolders", [demo], "InDefaultResults", 0)
            ),
            "Extensions", Map("QuickSwitch", Map("RecentFolders", 3)),         ; the folder panel only lists the recent items WriteRecentFolders created
            "CustomCommands", [
                Map("Title", "Website Redesign", "Type", "Folder", "Target", demo "\Website Redesign", "Arguments", "", "Keyword", ""),
                Map("Title", "Annual Report 2026", "Type", "Folder", "Target", demo "\Annual Report 2026", "Arguments", "", "Keyword", ""),
                Map("Title", "Demo Projects", "Type", "Folder", "Target", demo, "Arguments", "", "Keyword", "proj"),
                Map("Title", "Budget 2026", "Type", "File", "Target", demo "\Annual Report 2026\Budget 2026.xlsx", "Arguments", "", "Keyword", "budget"),
                Map("Title", "IP Configuration", "Type", "Command", "Target", "cmd.exe", "Arguments", "/k ipconfig /all", "Keyword", "ip"),
                Map("Title", "ALTRun on GitHub", "Type", "Url", "Target", "https://github.com/zhugecaomao/ALTRun", "Arguments", "", "Keyword", "altrun")
            ]
        )
        try DirDelete(Shots.AppDir "\Data", true)                           ; each scene starts with an empty index and learned ranking
        try FileDelete(Shots.AppDir "\ALTRun.json")                         ; the location of old versions
        DirCreate(Shots.AppDir "\Data")
        FileAppend(JSON.Stringify(settings, 4), Shots.AppDir "\Data\ALTRun.json", "UTF-8")
    }

    ; items pinned to the empty search box (only in the "empty" scene, see WritePinnedAndRecent)
    static Pinned := []

    static Entry(title, kind, target) {
        SplitPath(target, , &dir)
        return Map("Title", title, "Subtitle", (kind = "folder") ? target : dir, "Kind", kind, "Arg", target, "Arguments", ""
                 , "Icon", "", "Uid", kind ":" StrLower(target))
    }

    ; Two pinned + three recently opened items (Knowledge.json), shown in the empty search box
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

    ; Quick Switch's "recently used folders": shortcuts to three sample files in Windows' Recent items (newer than all others)
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

    ; Animation: a few entries in the clipboard history; structural calculation off (the calculator shows only the result, without beam bars / steel area, which most users wouldn't understand)
    static PrepareDemo() {
        Shots.WriteClipboard()
        settingsFile := Shots.AppDir "\Data\ALTRun.json"
        settings := JSON.Parse(FileRead(settingsFile, "UTF-8"))
        settings["Features"]["Calculator"]["StructuralCalc"] := 0
        FileDelete(settingsFile)
        FileAppend(JSON.Stringify(settings, 4), settingsFile, "UTF-8")
    }

    ; Clipboard history: a few common texts, one of them pinned
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

    ; A plain background covering the screen hides other windows on the desktop (translucent themes would show what is behind).
    ; It must be topmost to cover the console window; the search window is topmost too, and the one shown later is on top
    static ShowBackdrop() {
        backdrop := Gui("-Caption +ToolWindow +AlwaysOnTop -DPIScale")
        backdrop.BackColor := "8A9BB0"
        backdrop.Show("NA x0 y0 w" A_ScreenWidth " h" A_ScreenHeight)
        Shots.Backdrop := backdrop
    }
    static Backdrop := ""

    ; Before the shot, put the backdrop right below this window: another topmost window on the runner (e.g. the run log) may slip in between,
    ; and a translucent theme (Frost) would show its text in the shot
    static BackdropBehind(hwnd) {
        if !IsObject(Shots.Backdrop)
            return
        DllCall("SetWindowPos", "Ptr", Shots.Backdrop.Hwnd, "Ptr", hwnd, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x13)   ; SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE
        Sleep(100)                                                          ; wait for DWM to compose again
    }

    ;---------------------------------------------------------------------------
    ; Start / close
    ;---------------------------------------------------------------------------
    static Launch(theme, args := "", prepare := "") {
        Shots.WriteSettings(theme)
        if IsObject(prepare)
            prepare()
        Run('"' A_AhkPath '" "' Shots.AppDir '\ALTRun.ahk" ' args, Shots.AppDir, , &pid)
        Shots.Pid := pid
        ; the "ALTRun is already running" notice at the top of the screen at start-up disappears after 3 s
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
        Sleep(2000)                                                         ; search + icons loaded asynchronously
    }

    ;---------------------------------------------------------------------------
    ; Scenes
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
        Sleep(4000)                                                         ; the built-in file index
        ControlSetText("", "Edit1", hwnd)
        ControlSend("{Space}", "Edit1", hwnd)
        Sleep(300)
        ControlSend("{Text}" text, "Edit1", hwnd)
        Sleep(2000)
        return hwnd
    }

    ; Search by type (doc report): type after the built-in file index is ready
    static TypeSearch(text) {
        hwnd := Shots.SearchWindow()
        Sleep(4000)
        Shots.SetQuery(hwnd, text)
        return hwnd
    }

    ; the clipboard history records the foreground program at copy time: copy from Notepad, not from the screenshot script itself
    ; Empty search box: the window opened at start-up, after the icons load
    static EmptyBox() {
        hwnd := Shots.SearchWindow()
        Sleep(2500)
        return hwnd
    }

    ; A notice after an action (HUD): Enter copies the calculation result, the search window closes, and the notice shows low in the middle of the screen
    static Hud(text) {
        hwnd := Shots.Search(text)
        ControlSend("{Enter}", "Edit1", hwnd)
        hud := WinWait("ALTRun HUD ahk_class AutoHotkeyGUI ahk_pid " Shots.Pid, , 3)
        if !hud
            throw Error("HUD not shown")
        Sleep(300)                                                          ; fade in
        return hud
    }

    ; Quick Switch: two Explorer windows open, another program (OpenDialog.ahk) shows an "Open" dialog with the folder panel attached below.
    ; the screen may be only 1024 x 768: the dialog goes at the top so the panel fits
    static QuickSwitch() {
        try ControlSend("{Esc}", "Edit1", Shots.SearchWindow())            ; close the search window opened at start-up first
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
        WinSetAlwaysOnTop(1, dialog)                                        ; above the backdrop
        k := Shots.Scale
        WinMove(Round(160 * k), Round(20 * k), Round(720 * k), Round(440 * k), dialog)
        WinActivate(dialog)
        panel := WinWait("ALTRun Quick Switch ahk_pid " Shots.Pid, , 5)
        if !panel
            throw Error("Quick Switch panel not shown")
        Sleep(1500)                                                         ; the panel's icons and position
        return [dialog, panel]
    }
    static HelperPid := 0

    ; Closes the windows the scenes opened (Quick Switch dialog, Explorer, Notepad, context menu), shows the hidden terminal, puts the Send to shortcut back
    static CloseHelpers() {
        for pid in [Shots.HelperPid, Shots.SamplePids*]
            if pid
                try ProcessClose(pid)
        Shots.HelperPid := 0, Shots.SamplePids := []
        for window in Shots.HiddenWindows                                   ; the terminal window the demo hid
            try WinShow(window)
        Shots.HiddenWindows := []
        for window in WinGetList("ahk_class CabinetWClass")
            try WinClose(window)
        if WinExist("ahk_class #32768")                                     ; a context menu left open
            try Send("{Esc 3}")
        if Shots.SendToActive {                                             ; the Send to scene: the demo shortcut out, the one put aside back
            try FileDelete(Shots.SendToLink)
            if Shots.SendToBackup
                try FileMove(Shots.SendToBackup, Shots.SendToLink)
            Shots.SendToActive := false, Shots.SendToBackup := ""
        }
    }

    ; The animation on the README front page: 7 scenes (applications, actions panel, files, clipboard history, window switching, calculator, system commands); each step saves the same screen area
    ; as a frame, noting the key pressed for it. MakeDemo.py then cuts out the window, puts it on a blurred background and draws the keys at the bottom right. Returns "" (no separate shot)
    static DemoFrames := A_Temp "\ALTRunDemoFrames"
    static SamplePids := [], HiddenWindows := []
    static Demo() {
        hwnd := Shots.SearchWindow()
        Shots.OpenSampleWindows()
        WinActivate(hwnd)
        Shots.WaitPainted(hwnd, 10)
        dir := Shots.DemoFrames
        try DirDelete(dir, true)
        DirCreate(dir)
        r := Shots.FrameRect(hwnd)
        pad := Round(24 * Shots.Scale)
        region := {X: Max(0, r.X - pad), Y: Max(0, r.Y - pad), W: r.W + 2 * pad, H: Min(A_ScreenHeight - Max(0, r.Y - pad), Round(640 * Shots.Scale))}
        state := {Frames: "", Count: 0}                                     ; inner functions change outer variables: passed through an object
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
        frame(900, "Alt+Space")                                             ; the empty search box just shown
        type("r", 350), type("e", 350), type("p", 1500)
        ControlSend("{Right}", "Edit1", hwnd), Sleep(1200), frame(1800, "→")
        ControlSend("{Esc}", "Edit1", hwnd), Sleep(300), clear()
        ControlSend("{Space}", "Edit1", hwnd), Sleep(300)
        ControlSend("{Text}report", "Edit1", hwnd), Sleep(1200), frame(2000, "Space")
        clear(), ControlSend("{Backspace}", "Edit1", hwnd), Sleep(300)      ; back to normal search
        WinActivate(hwnd), Send("^!c"), Sleep(1500), frame(2000, "Ctrl+Alt+C")   ; the clipboard history shortcut
        clear(), type("w ", 2000)
        clear(), type("1200*1.09", 1800)
        clear(), type("lock", 1800)
        FileAppend(state.Frames, dir "\frames.txt", "UTF-8-RAW")
        Shots.Log("saved " state.Count " demo frames to " dir)
        return ""
    }

    ; Windows to switch to (all behind the backdrop): Explorer, Notepad with a text file, Character Map, Remote Desktop Connection; returns how many didn't show
    ; (Windows 11 opens a second Explorer folder as a tab of the same window, so other programs instead).
    ; On the GitHub VM the terminal window running the shots is on the taskbar too and window switching would list it: hide it first, CloseHelpers shows it again
    static OpenSampleWindows() {
        Run('explorer.exe "' Shots.DemoDir '\Annual Report 2026"')
        notes := Shots.AppDir "\Packing List.txt"
        if !FileExist(notes)
            FileAppend("Passport`r`nCharger`r`nUmbrella`r`n", notes, "UTF-8")
        for command in ['notepad.exe "' notes '"', "charmap.exe", "mstsc.exe"]
            try Run(command, , , &pid), Shots.SamplePids.Push(pid)
        for title in ["ahk_exe WindowsTerminal.exe", "ahk_class ConsoleWindowClass", "ahk_class CASCADIA_HOSTING_WINDOW_CLASS"]
            for window in WinGetList(title)
                try WinHide(window), Shots.HiddenWindows.Push(window)
        ; wait until every window is on screen: some take several seconds (a fixed wait once left only two in the list)
        missing := []
        for title in ["Annual Report 2026 ahk_class CabinetWClass", "Packing List.txt", "Character Map", "Remote Desktop Connection"]
            if !WinWait(title, , 15)
                missing.Push(title)
        for title in missing
            Shots.Log("sample window not shown: " title)
        Sleep(1000)
        return missing.Length
    }

    ; Window switching: "w " lists the open windows
    static Windows() {
        hwnd := Shots.SearchWindow()
        if Shots.OpenSampleWindows()
            throw Error("not all sample windows shown")
        WinActivate(hwnd)
        Shots.SetQuery(hwnd, "w ")
        return hwnd
    }

    ; Explorer's "Send to": ALTRun creates its Send to shortcut at start-up when SendToMenu is on (off in all other scenes).
    ; A Send to shortcut already there (your own ALTRun's, when run locally) is put aside and restored by CloseHelpers
    static SendToLink := RegExReplace(A_StartMenu, "\\Start Menu$", "\SendTo") "\ALTRun.lnk"
    static SendToActive := false, SendToBackup := ""
    static EnableSendTo() {
        if (!Shots.SendToActive && FileExist(Shots.SendToLink)) {
            Shots.SendToBackup := A_Temp "\ALTRun-SendTo-backup.lnk"
            FileMove(Shots.SendToLink, Shots.SendToBackup, true)
        }
        Shots.SendToActive := true
        settingsFile := Shots.AppDir "\Data\ALTRun.json"
        settings := JSON.Parse(FileRead(settingsFile, "UTF-8"))
        settings["General"]["SendToMenu"] := 1
        FileDelete(settingsFile)
        FileAppend(JSON.Stringify(settings, 4), settingsFile, "UTF-8")
    }

    ; A sample file selected in Explorer, its context menu open with "Send to" expanded (ALTRun is the first item).
    ; The full menu (the short Windows 11 menu has no "Send to")
    static SendTo() {
        try ControlSend("{Esc}", "Edit1", Shots.SearchWindow())            ; close the search window opened at start-up first
        Loop 20 {
            if FileExist(Shots.SendToLink)
                break
            Sleep(250)
        }
        if !FileExist(Shots.SendToLink)
            throw Error("Send to shortcut not created")
        Run('explorer.exe /select,"' Shots.DemoDir '\Annual Report 2026\Annual Report 2026 Draft.docx"')
        explorer := WinWait("Annual Report 2026 ahk_class CabinetWClass", , 10)
        if !explorer
            throw Error("Explorer window not shown")
        k := Shots.Scale
        WinMove(Round(60 * k), Round(40 * k), Round(760 * k), Round(580 * k), explorer)   ; tall enough for the menu
        WinSetAlwaysOnTop(1, explorer)                                      ; above the backdrop
        WinActivate(explorer)
        Sleep(2000)                                                         ; the selection and the file icons
        Shots.BackdropBehind(explorer)
        ; WM_CONTEXTMENU at the selection (lParam -1, as from the keyboard) straight to the file list: Shift+F10 goes to the focus,
        ; which in Windows 11 Explorer stays in the address / command bar (a XAML island) and opens that bar's menu.
        ; Windows 11 shows its short menu without "Send to" (screenshots.yml sets the VM to the full menu); its "Show more options" (Shift+F10)
        ; is tried too. Not Shift + WM_CONTEXTMENU: that gives the extended menu, whose Send to also lists every folder in the user profile
        view := 0
        try view := DllCall("FindWindowEx", "Ptr", ControlGetHwnd("SHELLDLL_DefView1", explorer), "Ptr", 0, "Str", "DirectUIHWND", "Ptr", 0, "Ptr")
        if !view
            throw Error("Explorer file list not found")
        PostMessage(0x7B, view, -1, , view)                                 ; WM_CONTEXTMENU
        if !WinWait("ahk_class #32768", , 3) {
            Sleep(500)
            Send("+{F10}")
            if !WinWait("ahk_class #32768", , 5) {
                try Shots.Log("active window: " WinGetClass("A"))
                throw Error("context menu not shown")
            }
        }
        Sleep(1000)
        Send("n")                                                           ; Se&nd to: opens the submenu with its first item selected
        Sleep(1500)                                                         ; the Send to items are read when it opens
        ; Explorer is topmost and the menus above it: captured as they are, without changing the order (making Explorer topmost again would cover the menu)
        left := 99999, top := 99999, right := -99999, bottom := -99999, count := 0
        for hwnd in [explorer, WinGetList("ahk_class #32768")*] {
            r := Shots.FrameRect(hwnd)
            Shots.Log(WinGetClass(hwnd) " at " r.X "," r.Y " " r.W "x" r.H)
            left := Min(left, r.X), top := Min(top, r.Y), right := Max(right, r.X + r.W), bottom := Max(bottom, r.Y + r.H)
            count += 1
        }
        if (count < 3)
            throw Error("Send to submenu not shown")
        outFile := Shots.OutDir "\sendto.png"
        hbm := Shots.CaptureRect(left, top, right - left, bottom - top)
        try Shots.SavePng(hbm, outFile)
        finally DllCall("DeleteObject", "Ptr", hbm)
        Shots.Log("saved " outFile " (" (right - left) "x" (bottom - top) ")")
        return ""                                                           ; already saved
    }

    static Preferences() {
        hwnd := WinWait("ahk_class AutoHotkeyGUI ahk_pid " Shots.Pid, , 10)
        if !hwnd
            throw Error("preferences window not found")
        WinSetAlwaysOnTop(1, hwnd)                                          ; above the backdrop
        r := Shots.FrameRect(hwnd)
        if (r.Y < 0)                                                        ; at 150% the window is tall, and centering would push its title bar above the screen
            WinMove(, 0, , , hwnd)
        WinActivate(hwnd)
        Sleep(1500)
        return hwnd
    }

    ;---------------------------------------------------------------------------
    ; Screenshot: copies the window area from the screen (with translucency) and saves it as a PNG with GDI+
    ;---------------------------------------------------------------------------
    ; Waits until the list area is no longer blank (the window is busy and hasn't drawn the results yet), at most seconds s
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

    ; when the list can't be drawn, log all ALTRun windows (including the contents of error dialogs) and the list state
    static Diagnose(hwnd) {
        try Shots.Log("list: items=" SendMessage(0x1004, 0, 0, "SysListView321", hwnd) " visible=" ControlGetVisible("SysListView321", hwnd))
        DetectHiddenWindows(false)
        for window in WinGetList("ahk_pid " Shots.Pid) {
            Shots.Log("window: [" WinGetClass(window) "] " WinGetTitle(window))
            if (WinGetClass(window) = "#32770")
                Shots.Log(WinGetText(window))
        }
    }

    ; The window's visible area on screen {X, Y, W, H}
    static FrameRect(hwnd) {
        rect := Buffer(16, 0)
        if (DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", hwnd, "UInt", 9, "Ptr", rect, "UInt", 16) = 0 && NumGet(rect, 8, "Int") > NumGet(rect, 0, "Int")) {
            x := NumGet(rect, 0, "Int"), y := NumGet(rect, 4, "Int")         ; DWMWA_EXTENDED_FRAME_BOUNDS: without the invisible border
            return {X: x, Y: y, W: NumGet(rect, 8, "Int") - x, H: NumGet(rect, 12, "Int") - y}
        }
        WinGetPos(&x, &y, &w, &h, hwnd)
        return {X: x, Y: y, W: w, H: h}
    }

    ; Several windows in one shot (dialog + folder panel): bring them to the front in the given order (later ones on top), then put the backdrop right below the first.
    ; Without this ordering: after the previous scene put the backdrop below the search window, the dialog was covered by it and only the panel came out
    static CaptureWindows(hwnds, file) {
        for hwnd in hwnds
            DllCall("SetWindowPos", "Ptr", hwnd, "Ptr", -1, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x13)   ; HWND_TOPMOST, SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE
        Shots.BackdropBehind(hwnds[1])
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
        Shots.BackdropBehind(hwnd)
        r := Shots.FrameRect(hwnd), w := r.W, h := r.H
        return Shots.CaptureRect(r.X, r.Y, w, h, &blank)
    }

    ; Copies an area of the screen and returns an HBITMAP (the caller calls DeleteObject); blank: the search window's list area is still blank
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

    ; search window: points every few pixels below the input box (from about 80 pixels down) all the same = results not drawn yet
    static _ListIsBlank(hdc, w, h) {
        if (h < Round(120 * Shots.Scale))
            return false
        top := Round(80 * Shots.Scale)
        bottom := h - 4 - Round(40 * Shots.Scale)                           ; the status bar at the bottom isn't looked at (its different color used to make it always count as drawn)
        first := DllCall("GetPixel", "Ptr", hdc, "Int", Round(60 * Shots.Scale), "Int", top, "UInt")
        y := top
        while (y < bottom) {
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
