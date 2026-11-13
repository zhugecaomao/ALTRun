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
; "normal-scale" / "-normal-scale" = the scenes in Shots.NormalScale / all but them; "themes" / "-themes" = the theme-... scenes / all but them.
; The "demo" scenes aren't single shots: every frame of a sequence goes to %Temp%\ALTRunDemoFrames\<scene> (frames.txt, each line: file name Tab ms shown),
; and MakeDemo.py (run by MakeDemos.ps1) turns each folder into the animation <scene>.png: demo.png on the README front page, demo-quickswitch.png etc. for the features. RoundCorners.py adds the rounded corners; before the shots SetDisplay.ps1 sets the screen to
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
    static SettleMs := 1200                                         ; after start-up, before a scene starts (see Launch)

    static NormalScale := ["quickswitch", "quickswitch-favorites", "demo-quickswitch", "pttools", "spf2m", "sendto", "prefs-general", "prefs-appearance", "prefs-filesearch", "prefs-searchsources", "prefs-commands", "command-editor", "prefs-selection", "prefs-snippets", "prefs-calculator", "prefs-scripts"
                         , "prefs-keyboard", "prefs-hotkeys", "prefs-quickswitch", "prefs-dialogpanel", "prefs-datestamp", "prefs-delayedstart", "prefs-startupapps", "prefs-sync", "sync-editor", "sync-preview", "hotkey-editor", "hotkey-editor-command", "command-picker", "prefs-hotkeys-commands", "command-editor-hotkeys", "hotkey-editor-clash"
                         , "prefs-window", "prefs-features", "prefs-applications", "prefs-clipboard", "prefs-websearch", "prefs-usage", "prefs-advanced", "data-location", "data-location-list"]   ; scenes shot at 100% (see screenshots.yml)

    ; the theme scenes (theme-dark ...): "themes" in the scene list (screenshots.yml shoots them on a runner of their own)
    static Themes() {
        names := []
        for scene in Shots.Scenes()
            if (SubStr(scene[1], 1, 6) = "theme-")
                names.Push(scene[1])
        return names
    }

    ; scene name -> [theme, start arguments, function, function preparing data before start-up (optional)]
    static Scenes() {
        return [
            ["search",      "Light", "", () => Shots.Search("re")],
            ["pinyin",      "Light", "", () => Shots.Search("jsb")],
            ["actions",     "Light", "", () => Shots.Actions("website")],
            ["actions-command", "Light", "", () => Shots.ActionsFiltered("en", "hotkey"), () => Shots.CommandHotkey()],   ; a custom command's actions panel filtered to "Change Hotkey Ctrl+Alt+N..." (below the first 8 rows; a command without {query}: that one opens the panel only with text typed)
            ["shellmenu",   "Light", "", () => Shots.ShellMenu(), () => Shots.EnableContextMenu()],   ; Wiki: the Windows context menu from the actions panel
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
            ["quickswitch-favorites", "Light", "", () => Shots.QuickSwitchFavorites(), () => Shots.WriteRecentFolders()],   ; Wiki: the panel's star menu
            ["windows",     "Light", "", () => Shots.Windows()],                 ; window switching (w + Space)
            ["sendto",      "Light", "", () => Shots.SendTo(), () => Shots.EnableSendTo()],
            ["pttools",     "Light", "", () => Shots.ToolWindow("PTTools", "PT Tools")],          ; Wiki: extensions
            ["spf2m",       "Light", "", () => Shots.ToolWindow("SPF2M", "SPF2M")],
            ["demo",        "Light", "", () => Shots.Demo(), () => Shots.PrepareDemo()],
            ["demo-quickswitch", "Light", "", () => Shots.DemoQuickSwitch(), () => Shots.WriteRecentFolders()],   ; feature animations (README, site, Wiki)
            ["demo-clipboard",   "Light", "", () => Shots.DemoClipboard(), () => Shots.WriteClipboard()],
            ["demo-snippets",    "Light", "", () => Shots.DemoSnippets()],
            ["demo-browse",      "Light", "", () => Shots.DemoBrowse()],
            ["demo-calculator",  "Light", "", () => Shots.DemoCalculator(), () => Shots.PrepareDemo()],
            ["prefs-general",    "Light", "-Preferences General", () => Shots.Preferences()],
            ["prefs-appearance", "Light", "-Preferences Appearance", () => Shots.Preferences()],
            ["prefs-commands",   "Light", "-Preferences Commands", () => Shots.Preferences()],
            ["command-editor",   "Light", "-Preferences Commands", () => Shots.SyncWindow("Edit...", "Custom Commands", 7), () => Shots.PingForIp()],   ; Wiki: custom commands (edit window of the default Ping: {query}, the type IP and a hotkey)
            ["command-editor-hotkeys", "Light", "-Preferences Commands", () => Shots.SetHotkeyMenu(), () => Shots.PingForIp()],   ; "Set Hotkey..." of a command with a hotkey: the menu
            ["prefs-selection",  "Light", "-Preferences Selection", () => Shots.Preferences(), () => Shots.SelectionTypes()],   ; Wiki: actions on selected content
            ["prefs-snippets",   "Light", "-Preferences Snippets", () => Shots.Preferences()],
            ["prefs-calculator", "Light", "-Preferences Calculator", () => Shots.Preferences()],
            ["prefs-scripts",    "Light", "-Preferences Scripts", () => Shots.Preferences()],
            ["prefs-filesearch", "Light", "-Preferences FileSearch", () => Shots.Preferences(), () => Shots.FilesInNormalResults()],   ; Wiki: file search (files in normal results)
            ["prefs-searchsources", "Light", "-Preferences FileIndex", () => Shots.Preferences()],   ; Wiki: file search
            ["prefs-keyboard",   "Light", "-Preferences Keyboard", () => Shots.Preferences()],      ; Wiki: Custom Hotkeys (Caps Lock)
            ["prefs-hotkeys",    "Light", "-Preferences Hotkeys", () => Shots.Preferences()],     ; Wiki: custom hotkeys
            ["hotkey-editor",    "Light", "-Preferences Hotkeys", () => Shots.SyncWindow("Edit...", "Edit Hotkey", 2)],   ; Wiki: custom hotkeys (edit window, Ctrl+Alt+F)
            ["hotkey-editor-command", "Light", "-Preferences Hotkeys", () => Shots.SyncWindow("Edit...", "Edit Hotkey", 1), () => Shots.CommandHotkey()],   ; Wiki: custom hotkeys (running a custom command)
            ["hotkey-editor-clash", "Light", "-Preferences Hotkeys", () => Shots.SyncWindow("Edit...", "Edit Hotkey", 5), () => Shots.CommandHotkey(true, true)],   ; a hotkey with the keys of another one: the red line below the summary
            ["command-picker",   "Light", "-Preferences Hotkeys", () => Shots.CommandPicker(), () => Shots.CommandHotkey()],   ; Wiki: custom hotkeys (Choose Command)
            ["prefs-hotkeys-commands", "Light", "-Preferences Hotkeys", () => Shots.Preferences(), () => Shots.CommandHotkey(true)],   ; the list with a custom command's hotkey and one whose command was deleted (red)
            ["prefs-quickswitch", "Light", "-Preferences QuickSwitch", () => Shots.Preferences()],    ; Wiki: quick switch for dialogs
            ["prefs-dialogpanel", "Light", "-Preferences QSPanel", () => Shots.Preferences()],    ; Wiki: the dialogs' folder panel
            ["prefs-datestamp",  "Light", "-Preferences DateStamp", () => Shots.Preferences()],     ; Wiki: date stamp
            ["prefs-window",     "Light", "-Preferences Window", () => Shots.Preferences()],      ; Wiki: search and shortcuts
            ["prefs-features",   "Light", "-Preferences Features", () => Shots.Preferences()],      ; Wiki: settings reference (features)
            ["prefs-applications", "Light", "-Preferences Applications", () => Shots.Preferences()],    ; Wiki: automatic app index updates
            ["prefs-clipboard",  "Light", "-Preferences Clipboard", () => Shots.Preferences(), () => Shots.ClipboardDefault()],   ; Wiki: clipboard history
            ["prefs-websearch",  "Light", "-Preferences WebSearch", () => Shots.Preferences()],     ; Wiki: web search
            ["prefs-delayedstart", "Light", "-Preferences DelayedStart", () => Shots.Preferences(), () => Shots.DelayedStartList()],   ; Wiki: delayed start
            ["prefs-startupapps", "Light", "-Preferences DelayedStart", () => Shots.StartupApps(), () => Shots.DelayedStartList()],   ; Wiki: delayed start (startup apps)
            ["prefs-sync",       "Light", "-Preferences Sync", () => Shots.Preferences(), () => Shots.SyncTasks()],    ; Wiki: sync tasks
            ["sync-editor",      "Light", "-Preferences Sync", () => Shots.SyncWindow("Edit...", "Folder Sync Task"), () => Shots.SyncTasks()],   ; Wiki: sync tasks (edit window)
            ["sync-preview",     "Light", "-Preferences Sync", () => Shots.SyncPreview(), () => Shots.SyncTasks(true)],   ; Wiki: sync tasks (preview)
            ["prefs-usage",      "Light", "-Preferences Usage", () => Shots.Preferences(), () => Shots.WriteUsage()],   ; Wiki: usage statistics
            ["prefs-advanced",   "Light", "-Preferences Advanced", () => Shots.Preferences()],     ; Wiki: data files
            ["data-location",    "Light", "-Preferences Advanced", () => Shots.DataLocation(), () => Shots.DemoOneDrive()],   ; Wiki: FAQ, several PCs (Change Location...)
            ["data-location-list", "Light", "-Preferences Advanced", () => Shots.DataLocation(true), () => Shots.DemoOneDrive()],   ; the same with its list of places open
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
        wanted := Map(), skipped := Map(), within := Map()                  ; "name": shoot it, "-name": skip it, "=name": only scenes also listed so
        Loop args.Length - 1 {
            name := args[A_Index + 1]
            if (Trim(name, " -=") = "")                                      ; an empty name would select no scene at all
                continue
            target := (SubStr(name, 1, 1) = "-") ? skipped : (SubStr(name, 1, 1) = "=") ? within : wanted
            name := LTrim(name, "-=")
            for scene in (name = "normal-scale") ? Shots.NormalScale : (name = "themes") ? Shots.Themes() : [name]
                target[scene] := true
        }
        DirCreate(Shots.OutDir)
        Shots.Log("screen " A_ScreenWidth " x " A_ScreenHeight ", DPI " A_ScreenDPI)
        count := 0
        for scene in Shots.Scenes()
            count += !((wanted.Count && !wanted.Has(scene[1])) || skipped.Has(scene[1]) || (within.Count && !within.Has(scene[1])))
        Shots.Log(count " scene(s) to shoot")
        Shots.PrepareApp()
        Shots.PrepareDemoFiles()
        Shots.PrepareAppIndex()
        Shots.ShowBackdrop()
        for scene in Shots.Scenes() {
            if ((wanted.Count && !wanted.Has(scene[1])) || skipped.Has(scene[1]) || (within.Count && !within.Has(scene[1])))
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

    ; Builds the app index once (Data\AppIndex.json, the demo apps folder) with ALTRun started in the background, and keeps a copy
    ; that WriteSettings puts back into each scene's new Data folder. Without it every scene scanned the apps 3 s after start-up,
    ; and with the short wait in Launch the first results missed them (2026-10: "re" without Remote Desktop, "jsb" without Notepad)
    static AppIndexCopy => Shots.AppDir "\AppIndex.ready.json"             ; kept outside Data, which WriteSettings clears for each scene
    static PrepareAppIndex() {
        indexFile := Shots.AppDir "\Data\AppIndex.json"
        try FileDelete(Shots.AppIndexCopy)
        Shots.WriteSettings("Light")
        started := A_TickCount
        Run('"' A_AhkPath '" "' Shots.AppDir '\ALTRun.ahk" -Startup', Shots.AppDir, , &pid)
        while (!FileExist(indexFile) && A_TickCount - started < 30000)
            Sleep(250)
        Sleep(500)                                                          ; written through a temporary file and renamed; a moment more anyway
        try ProcessClose(pid)
        ProcessWaitClose(pid, 5)
        if !FileExist(indexFile)
            return Shots.Log("app index not written: scenes may miss apps")
        FileCopy(indexFile, Shots.AppIndexCopy, true)
        text := FileRead(indexFile, "UTF-8")
        Shots.Log("app index ready after " Round((A_TickCount - started) / 1000, 1) " s (" StrLen(text) " characters, Notepad " (InStr(text, "notepad") ? "in it" : "missing") ")")
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
                "FileSearch", Map("UseEverything", 0, "ScopeFolders", [demo], "InDefaultResults", 0),
                "Clipboard", Map("LocalFiles", 0)                           ; the made-up history stays in this Data folder, away from this PC's %LOCALAPPDATA%
            ),
            "Extensions", Map("QuickSwitch", Map("RecentFolders", 3, "Favorites", [demo, demo "\Annual Report 2026"])),   ; the folder panel: two favorites, and only the recent items WriteRecentFolders created
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
        if FileExist(Shots.AppIndexCopy)                                    ; but with the app index built once (PrepareAppIndex), so no scene waits for a scan
            FileCopy(Shots.AppIndexCopy, Shots.AppDir "\Data\AppIndex.json", true)
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
        ; Started without arguments, ALTRun ends its start-up by showing the search window and the "ALTRun is running" notice (3 s):
        ; hide the notice as soon as it shows instead of waiting for it to go. With arguments (-Preferences ...) there is no notice:
        ; wait for its (hidden) main window. Then a moment for what runs 0.5 s after start-up (icon cache, warm-up)
        if (args = "") {
            if (hud := WinWait("ALTRun HUD ahk_pid " pid, , 10))
                try WinHide(hud)
        } else {
            saved := A_DetectHiddenWindows
            DetectHiddenWindows(true)
            WinWait("ahk_class AutoHotkey ahk_pid " pid, , 10)
            DetectHiddenWindows(saved)
        }
        Sleep(Shots.SettleMs)
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

    ; The actions panel of the result found by text, filtered by typing filter
    static ActionsFiltered(text, filter) {
        hwnd := Shots.Actions(text)
        ControlSend("{Text}" filter, "Edit1", hwnd)
        Sleep(800)
        return hwnd
    }

    ; A file's actions panel filtered to "Windows Context Menu", then Enter: the search window and the Explorer menu over it
    static ShellMenu() {
        hwnd := Shots.Search(Shots.DemoDir "\Annual Report 2026\Budget")
        ControlSend("{Right}", "Edit1", hwnd)
        Sleep(800)
        ControlSend("{Text}context", "Edit1", hwnd)
        Sleep(800)
        ControlSend("{Enter}", "Edit1", hwnd)
        popup := WinWait("ahk_class #32768 ahk_pid " Shots.Pid, , 5)
        if !popup
            throw Error("Context menu not shown")
        Sleep(1500)                                                         ; the menu's icons (owner-drawn items)
        return [hwnd, popup]
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
        dialog := WinWait("Open ahk_class #32770 ahk_pid " pid, , 30)      ; the first file dialog on a fresh runner can take over 10 s (the shell starts up)
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

    ; The panel's star menu: the favorites, the first one's subfolders open (Down selects it, Right opens its submenu)
    static QuickSwitchFavorites() {
        windows := Shots.QuickSwitch()
        star := 0
        for ctrl in WinGetControlsHwnd(windows[2])                          ; the star: the panel's only icon (Static with SS_ICON)
            if (WinGetClass(ctrl) = "Static" && (ControlGetStyle(ctrl) & 0x1F) = 0x3)
                star := ctrl
        if !star
            throw Error("Star button not found")
        ControlClick(star)
        if !WinWait("ahk_class #32768 ahk_pid " Shots.Pid, , 5)
            throw Error("Favorites menu not shown")
        Sleep(500)
        Send("{Down}{Right}")
        Sleep(1500)                                                         ; the submenu and the menus' icons
        popups := WinGetList("ahk_class #32768 ahk_pid " Shots.Pid)
        if (popups.Length < 2)
            throw Error("Favorite's submenu not shown")
        windows.Push(popups*)
        return windows
    }

    ; PT Tools / SPF2M: the system command opens the window (its default inputs), at the top left so it fits the screen
    static ToolWindow(command, title) {
        hwnd := Shots.SearchWindow()
        Shots.SetQuery(hwnd, command)
        ControlSend("{Enter}", "Edit1", hwnd)
        if !(window := WinWait(title " ahk_class AutoHotkeyGUI ahk_pid " Shots.Pid, , 5))
            throw Error(title " not shown")
        WinMove(Round(40 * Shots.Scale), Round(40 * Shots.Scale), , , window)
        Sleep(800)
        return window
    }

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

    ;---------------------------------------------------------------------------
    ; Animations: each step saves the same screen area as a frame (FramesStart / Frame / FramesEnd); MakeDemo.py then cuts out the window
    ; (or, with several windows, the area they cover) and puts it on a blurred background. The scenes return "" (no separate shot)
    ;---------------------------------------------------------------------------
    static DemoFrames := A_Temp "\ALTRunDemoFrames"
    static SamplePids := [], HiddenWindows := []
    static _frames := ""

    ; windows: the frames show several windows (Notepad and the search window, a dialog and its panel): a marker file tells MakeDemo.py to keep the area they cover
    static FramesStart(scene, region, windows := false) {
        dir := Shots.DemoFrames "\" scene
        try DirDelete(dir, true)
        DirCreate(dir)
        if windows
            FileAppend("", dir "\windows")
        Shots._frames := {Dir: dir, Region: region, List: "", Count: 0}
    }

    static Frame(ms) {
        f := Shots._frames, r := f.Region
        name := Format("frame-{:02}.png", ++f.Count)
        hbm := Shots.CaptureRect(r.X, r.Y, r.W, r.H)
        try Shots.SavePng(hbm, f.Dir "\" name)
        finally DllCall("DeleteObject", "Ptr", hbm)
        f.List .= name "`t" ms "`n"
    }

    static FramesEnd() {
        f := Shots._frames
        FileAppend(f.List, f.Dir "\frames.txt", "UTF-8-RAW")
        Shots.Log("saved " f.Count " frames to " f.Dir)
        return ""
    }

    ; The area around the search window, tall enough for the result list it shows later
    static SearchRegion(hwnd) {
        r := Shots.FrameRect(hwnd)
        return {X: r.X, Y: r.Y, W: r.W, H: Max(r.H, Round(600 * Shots.Scale))}
    }

    ; The area around several rectangles {X, Y, W, H} with a margin, within the screen
    static Around(rects) {
        pad := Round(24 * Shots.Scale)
        left := 99999, top := 99999, right := -99999, bottom := -99999
        for r in rects
            left := Min(left, r.X), top := Min(top, r.Y), right := Max(right, r.X + r.W), bottom := Max(bottom, r.Y + r.H)
        left := Max(0, left - pad), top := Max(0, top - pad)
        return {X: left, Y: top, W: Min(A_ScreenWidth, right + pad) - left, H: Min(A_ScreenHeight, bottom + pad) - top}
    }

    ; The animation on the README front page: 7 scenes (applications, actions panel, files, clipboard history, window switching, calculator, system commands)
    static Demo() {
        hwnd := Shots.SearchWindow()
        Shots.OpenSampleWindows()
        WinActivate(hwnd)
        Shots.WaitPainted(hwnd, 10)
        Shots.FramesStart("demo", Shots.Around([Shots.SearchRegion(hwnd)]))
        frame(ms, keys := "") => Shots.Frame(ms)                           ; keys: the key pressed for this step (the status bar shows it)
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
        return Shots.FramesEnd()
    }

    ; Notepad above the backdrop, below the search window's top (the search window, topmost and shown later, covers its top part)
    static OpenNotepad(search) {
        Run("notepad.exe")
        if !(hwnd := WinWait("ahk_class Notepad", , 10))
            throw Error("Notepad not shown")
        Shots.SamplePids.Push(WinGetPID(hwnd))
        k := Shots.Scale
        WinSetAlwaysOnTop(1, hwnd)
        WinMove(search.X - Round(40 * k), search.Y + Round(70 * k), search.W + Round(80 * k), Round(380 * k), hwnd)
        WinActivate(hwnd)
        WinWaitActive(hwnd, , 3)
        Sleep(800)
        return hwnd
    }

    ; Quick Switch: the Open dialog in Travel; Ctrl+G jumps to the folder of the Explorer window in front (Annual Report 2026),
    ; Ctrl+Shift+G puts the caret in the panel's search box, "web" + Enter jumps to Website Redesign
    static DemoQuickSwitch() {
        windows := Shots.QuickSwitch()
        dialog := windows[1], panel := windows[2]
        for window in windows
            DllCall("SetWindowPos", "Ptr", window, "Ptr", -1, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x13)   ; HWND_TOPMOST, no size / move / activate
        Shots.BackdropBehind(dialog)
        p := Shots.FrameRect(panel)
        Shots.FramesStart("demo-quickswitch", Shots.Around([Shots.FrameRect(dialog), {X: p.X, Y: p.Y, W: p.W, H: p.H + Round(40 * Shots.Scale)}]), true)
        SendLevel(1)                                                        ; the hotkeys of another script (ALTRun) see these keys
        Shots.Frame(1800)
        WinActivate(dialog), Sleep(300)
        Send("^g"), Sleep(1800), Shots.Frame(2200)
        Send("^+g"), Sleep(800)
        Send("{Text}web"), Sleep(1200), Shots.Frame(1800)
        Send("{Enter}"), Sleep(1800), Shots.Frame(3000)
        SendLevel(0)
        return Shots.FramesEnd()
    }

    ; Clipboard history: Ctrl+Alt+C in Notepad, "meet" filters, Enter pastes the item into Notepad
    static DemoClipboard() {
        hwnd := Shots.SearchWindow()
        search := Shots.SearchRegion(hwnd)
        ControlSend("{Esc}", "Edit1", hwnd), Sleep(500)
        note := Shots.OpenNotepad(search)
        Send("{Text}Reply to Sam: "), Sleep(400)
        Shots.FramesStart("demo-clipboard", Shots.Around([search, Shots.FrameRect(note)]), true)
        Shots.Frame(1200)
        Send("^!c"), Sleep(1500), Shots.Frame(1800)
        ControlSend("{Text}meet", "Edit1", hwnd), Sleep(1200), Shots.Frame(1800)
        ControlSend("{Enter}", "Edit1", hwnd), Sleep(1500), Shots.Frame(3000)
        return Shots.FramesEnd()
    }

    ; Snippets: typing ;thx in Notepad expands the default "Email: thank you" snippet, with the caret where {cursor} was
    static DemoSnippets() {
        hwnd := Shots.SearchWindow()
        search := Shots.SearchRegion(hwnd)
        ControlSend("{Esc}", "Edit1", hwnd), Sleep(500)
        note := Shots.OpenNotepad(search)
        Shots.FramesStart("demo-snippets", Shots.Around([Shots.FrameRect(note)]), true)
        SendLevel(1)                                                        ; ALTRun's hotstrings see these keys
        Shots.Frame(1000)
        for char in StrSplit(";thx")
            Send("{Text}" char), Sleep(300), Shots.Frame(400)
        Sleep(900), Shots.Frame(1800)
        Send("{Text}Sam"), Sleep(600), Shots.Frame(3000)
        SendLevel(0)
        return Shots.FramesEnd()
    }

    ; Browsing folders: type a path, Tab goes into the selected folder, typing filters, Backspace goes up a level
    static DemoBrowse() {
        hwnd := Shots.SearchWindow()
        WinActivate(hwnd)
        Shots.WaitPainted(hwnd, 10)
        Shots.FramesStart("demo-browse", Shots.Around([Shots.SearchRegion(hwnd)]))
        Shots.Frame(800)
        ControlSend("{Text}" Shots.DemoDir "\", "Edit1", hwnd), Sleep(1500), Shots.Frame(2000)
        ControlSend("{Tab}", "Edit1", hwnd), Sleep(1500), Shots.Frame(2000)
        ControlSend("{Text}bud", "Edit1", hwnd), Sleep(1200), Shots.Frame(2000)
        ControlSend("{Backspace 3}", "Edit1", hwnd), Sleep(600)
        ControlSend("{Backspace}", "Edit1", hwnd), Sleep(1500), Shots.Frame(2500)
        return Shots.FramesEnd()
    }

    ; Calculator: arithmetic and unit conversions, the result shown as you type
    static DemoCalculator() {
        hwnd := Shots.SearchWindow()
        WinActivate(hwnd)
        Shots.WaitPainted(hwnd, 10)
        Shots.FramesStart("demo-calculator", Shots.Around([Shots.SearchRegion(hwnd)]))
        for query in ["(200+100)*2", "10 km in mi", "100 f to c", "1 GB to MB"]
            Shots.SetQuery(hwnd, query), Shots.Frame(1800)
        return Shots.FramesEnd()
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
    ; The Clipboard page shows the default "Save in" (everything on this PC only) instead of the scenes' Data folder (WriteSettings)
    static ClipboardDefault() {
        settingsFile := Shots.AppDir "\Data\ALTRun.json"
        settings := JSON.Parse(FileRead(settingsFile, "UTF-8"))
        settings["Features"]["Clipboard"]["LocalFiles"] := 2
        FileDelete(settingsFile)
        FileAppend(JSON.Stringify(settings, 4), settingsFile, "UTF-8")
    }

    ; The File Search page with "Also show matching files and folders in normal results" ticked, so its options aren't grayed out
    static FilesInNormalResults() {
        settingsFile := Shots.AppDir "\Data\ALTRun.json"
        settings := JSON.Parse(FileRead(settingsFile, "UTF-8"))
        settings["Features"]["FileSearch"]["InDefaultResults"] := 1
        FileDelete(settingsFile)
        FileAppend(JSON.Stringify(settings, 4), settingsFile, "UTF-8")
    }

    ; The startup apps picker of the Delayed Start page, with two made-up startup apps of this user (Run key). They are deleted again once the
    ; picker has listed them, so a run on your own PC leaves nothing behind
    static StartupApps() {
        runKey := "HKCU\Software\Microsoft\Windows\CurrentVersion\Run"
        demo := [["ALTRun Demo Sync Tool", '"C:\Tools\SyncTool\SyncTool.exe" --tray'], ["ALTRun Demo Chat", '"C:\Tools\Chat\Chat.exe" /minimized']]
        for value in demo
            RegWrite(value[2], "REG_SZ", runKey, value[1])
        try {
            prefs := Shots.Preferences()
            button := ControlGetHwnd("Startup Apps...", prefs)               ; the screenshots use the English interface
            Shots.Log("button: visible " ControlGetVisible(button) ", enabled " ControlGetEnabled(button))
            PostMessage(0xF5, 0, 0, button)                                 ; BM_CLICK, as in SyncWindow: a ControlClick was sometimes lost
            picker := WinWait("Startup Apps ahk_class AutoHotkeyGUI ahk_pid " Shots.Pid, , 10)
            if !picker {
                Shots.Diagnose(prefs)
                throw Error("startup apps window not shown")
            }
        } finally {
            for value in demo
                try RegDelete(runKey, value[1])
        }
        WinSetAlwaysOnTop(1, picker)
        Sleep(800)
        return picker
    }

    ; A delayed start list (made up)
    static DelayedStartList() {
        settingsFile := Shots.AppDir "\Data\ALTRun.json"
        settings := JSON.Parse(FileRead(settingsFile, "UTF-8"))
        settings["DelayedStart"] := [Map("Name", "OneDrive", "Target", "%LocalAppData%\Microsoft\OneDrive\OneDrive.exe", "Arguments", "/background", "Delay", 300, "Enabled", 1)
            , Map("Name", "Sync Tool", "Target", "C:\Tools\SyncTool\SyncTool.exe", "Arguments", "--tray", "Delay", 90, "Enabled", 1)
            , Map("Name", "Chat", "Target", "C:\Tools\Chat\Chat.exe", "Arguments", "", "Delay", 600, "Enabled", 0)]
        FileDelete(settingsFile)
        FileAppend(JSON.Stringify(settings, 4), settingsFile, "UTF-8")
    }

    ; The default type IP and the default Ping command set for it, added after the six demo commands (row 7), for the command editor
    static PingForIp() {
        settingsFile := Shots.AppDir "\Data\ALTRun.json"
        settings := JSON.Parse(FileRead(settingsFile, "UTF-8"))
        settings["SelectionTypes"] := [Map("Name", "IP", "Kind", "Text", "Pattern", "^(\d{1,3}\.){3}\d{1,3}$")]   ; as in AppSettings._DefaultSelectionTypes
        settings["CustomCommands"].Push(Map("Id", "pg7k2m4x", "Title", "Ping", "Type", "Command", "Target", "cmd.exe", "Arguments", "/k ping {query}", "Keyword", "ping", "Selection", "IP"))
        settings["Hotkeys"] := [Map("Key", "^!p", "Action", "custom:pg7k2m4x", "WinTitle", "")]   ; "Hotkey Ctrl+Alt+P" on the Triggers line
        FileDelete(settingsFile)
        FileAppend(JSON.Stringify(settings, 4), settingsFile, "UTF-8")
    }

    ; A custom command and a hotkey that runs it, for the hotkey edit window and Choose Command.
    ; withList: also a search, a system command and a hotkey whose command was deleted, for the Hotkeys page
    ; withClash: and a fifth one with the keys of the first (Ctrl+Alt+N)
    static CommandHotkey(withList := false, withClash := false) {
        settingsFile := Shots.AppDir "\Data\ALTRun.json"
        settings := JSON.Parse(FileRead(settingsFile, "UTF-8"))
        settings["CustomCommands"].Push(Map("Id", "np4k2x7q", "Title", "Edit in Notepad", "Type", "File", "Target", "notepad.exe", "Arguments", "", "Keyword", "en", "Selection", "Text files"))
        settings["Hotkeys"] := [Map("Key", "^!n", "Action", "custom:np4k2x7q", "WinTitle", "")]   ; the first one (the demo settings have no Hotkeys: the defaults)
        if withList
            settings["Hotkeys"].Push(Map("Key", "^!f", "Action", "Search", "SearchText", "folder {selection}", "WinTitle", "")
                , Map("Key", "^!t", "Action", "TerminalHere", "WinTitle", "ahk_class CabinetWClass")
                , Map("Key", "^!q", "Action", "custom:q7w2e9rt", "WinTitle", ""))   ; no such command: deleted
        if withClash
            settings["Hotkeys"].Push(Map("Key", "^!n", "Action", "Lock", "WinTitle", ""))
        FileDelete(settingsFile)
        FileAppend(JSON.Stringify(settings, 4), settingsFile, "UTF-8")
    }

    ; "Set Hotkey..." in the edit window of Ping (PingForIp): the menu of its hotkey below the button
    static SetHotkeyMenu() {
        editor := Shots.SyncWindow("Edit...", "Custom Commands", 7)
        button := 0
        for ctrl in WinGetControlsHwnd(editor)
            if (!button && WinGetClass(ctrl) = "Button" && ControlGetText(ctrl) = "Set Hotkey...")
                button := ctrl
        if !button
            throw Error("command editor: Set Hotkey... button not found")
        PostMessage(0xF5, 0, 0, button)                                     ; BM_CLICK, posted: the menu is modal
        popup := WinWait("ahk_class #32768 ahk_pid " Shots.Pid, , 5)
        if !popup
            throw Error("Set Hotkey...: menu not shown")
        Sleep(800)
        return [editor, popup]
    }

    ; "Choose..." in the edit window of the hotkey above: the list with that command selected
    static CommandPicker() {
        editor := Shots.SyncWindow("Edit...", "Edit Hotkey", 1)
        button := 0
        for ctrl in WinGetControlsHwnd(editor)
            if (!button && WinGetClass(ctrl) = "Button" && ControlGetText(ctrl) = "Choose...")
                button := ctrl
        if !button
            throw Error("hotkey editor: Choose... button not found")
        PostMessage(0xF5, 0, 0, button)                                     ; BM_CLICK, posted: the window is modal
        window := WinWait("Choose Command ahk_class AutoHotkeyGUI ahk_pid " Shots.Pid, , 20)
        if !window
            throw Error("Choose...: window not shown")
        WinSetAlwaysOnTop(1, window)
        Sleep(800)
        return window
    }

    ; Content types of a selection and commands that use them (made up), for Preferences → Selection
    static SelectionTypes() {
        settingsFile := Shots.AppDir "\Data\ALTRun.json"
        settings := JSON.Parse(FileRead(settingsFile, "UTF-8"))
        settings["SelectionTypes"] := [Map("Name", "PO", "Kind", "Text", "Pattern", "^PO\d{8}$")
            , Map("Name", "Ticket", "Kind", "Text", "Pattern", "i)^[A-Z]{2,5}-\d+$")
            , Map("Name", "PDF", "Kind", "File", "Pattern", "pdf")
            , Map("Name", "Drawing", "Kind", "File", "Pattern", "dwg, dxf")]
        for row in [["ERP Order", "https://erp.example.com/order?id={query}", "PO"], ["Warehouse Shipment", "https://wms.example.com/ship/{query}", "PO"]
                  , ["Issue Tracker", "https://issues.example.com/browse/{query}", "Ticket"]]
            settings["CustomCommands"].Push(Map("Title", row[1], "Type", "Url", "Target", row[2], "Arguments", "", "Keyword", "", "Selection", row[3]))
        settings["CustomCommands"].Push(Map("Title", "PDF Tool", "Type", "File", "Target", "C:\Tools\PdfTool\PdfTool.exe", "Arguments", "", "Keyword", "", "Selection", "PDF"))
        FileDelete(settingsFile)
        FileAppend(JSON.Stringify(settings, 4), settingsFile, "UTF-8")
    }

    ; Sync tasks (made up). sampleFolders: the first task syncs two sample folders, so Preview has something real to show
    ; (a new file, a changed one, and one only in the destination, which a mirror deletes)
    static SyncTasks(sampleFolders := false) {
        settingsFile := Shots.AppDir "\Data\ALTRun.json"
        settings := JSON.Parse(FileRead(settingsFile, "UTF-8"))
        task(name, source, destination, mode, monitor := 0) => Map("Name", name, "Source", source, "Destination", destination, "Files", "*.*", "Mode", mode
            , "Copy", "DAT", "DirCopy", "DA", "Retries", 1, "Wait", 1, "Threads", 16, "Monitor", monitor, "MonitorChanges", 1, "MonitorMinutes", 5
            , "ExcludeFiles", ["thumbs.db", "desktop.ini", "~$*.*"], "ExcludeDirs", ["System Volume Information", "$Recycle.Bin"], "LogFile", "", "AppendLog", 1, "Extra", "/NP")
        tasks := [task("Projects", "\\fileserver\Projects", "%OneDrive%\Work\Projects", "Mirror", 1)
            , task("Templates", "\\fileserver\Templates", "%OneDrive%\Work\Templates", "Update")
            , task("Photos Backup", "%UserProfile%\Pictures", "E:\Backup\Pictures", "Copy")]
        if sampleFolders {
            root := "C:\Demo Sync"
            try DirDelete(root, true)
            for sample in ["Server\Projects\Plan.docx", "Server\Projects\Budget.xlsx", "Server\Projects\Drawings\Site Plan.pdf", "Server\Projects\Drawings\Elevations.pdf"
                       , "OneDrive\Projects\Plan.docx", "OneDrive\Projects\Old Notes.txt"] {
                SplitPath(root "\" sample, , &dir)
                DirCreate(dir)
                FileAppend("sample", root "\" sample)
            }
            FileSetTime("20260101000000", root "\OneDrive\Projects\Plan.docx")      ; older than the source's: copied again
            tasks[1]["Source"] := root "\Server\Projects", tasks[1]["Destination"] := root "\OneDrive\Projects", tasks[1]["Monitor"] := 0
            settings["General"]["SaveLog"] := 1                             ; errors in the background (the preview ends in a timer) go to ALTRun.log, which Diagnose shows
        }
        settings["SyncTasks"] := tasks
        FileDelete(settingsFile)
        FileAppend(JSON.Stringify(settings, 4), settingsFile, "UTF-8")
    }

    ; Preferences → Sync: selects the first task, clicks a button (Edit / Preview) and returns the window that opens (its title starts with title)
    ; A list page's button that opens a window (also the hotkey editor): row = the list row selected first
    static SyncWindow(buttonText, title, row := 1) {
        prefs := Shots.Preferences()
        list := 0, button := 0
        for ctrl in WinGetControlsHwnd(prefs) {                             ; the controls of the page on show (other pages' are hidden)
            if !ControlGetVisible(ctrl)
                continue
            if (!list && WinGetClass(ctrl) = "SysListView32" && ControlGetClassNN(ctrl) != "SysListView321")   ; SysListView321 = the page list on the left
                list := ctrl
            if (!button && WinGetClass(ctrl) = "Button" && ControlGetText(ctrl) = buttonText)
                button := ctrl
        }
        if !(list && button)
            throw Error("sync page: list or " buttonText " button not found")
        ControlFocus(list), ControlSend("{Home}" (row > 1 ? "{Down " (row - 1) "}" : ""), list)
        Sleep(300)
        PostMessage(0xF5, 0, 0, button)                                     ; BM_CLICK: a posted mouse click (ControlClick) was sometimes lost; posted, so Edit's modal window doesn't block
        window := WinWait(title " ahk_class AutoHotkeyGUI ahk_pid " Shots.Pid, , 20)
        if !window {
            Shots.Diagnose(prefs)
            throw Error(buttonText ": window not shown")
        }
        WinSetAlwaysOnTop(1, window)
        Sleep(1000)
        return window
    }

    ; Preferences → Sync → Sync... opens the task's window (Ready); its Preview button previews the sample folders in it
    static SyncPreview() {
        window := Shots.SyncWindow("Sync...", "Projects - Folder Sync")
        preview := 0
        for ctrl in WinGetControlsHwnd(window)
            if (WinGetClass(ctrl) = "Button" && ControlGetText(ctrl) = "Preview")
                preview := ctrl
        if !preview
            throw Error("sync window: Preview button not found")
        PostMessage(0xF5, 0, 0, preview)                                    ; BM_CLICK
        list := 0                                                           ; robocopy /L of the sample folders, then the list: wait for its rows
        for ctrl in WinGetControlsHwnd(window)
            if (WinGetClass(ctrl) = "SysListView32")
                list := ctrl
        Loop 60 {
            Sleep(250)
            if (list && SendMessage(0x1004, 0, 0, list) > 0)                 ; LVM_GETITEMCOUNT
                break
        }
        if !(list && SendMessage(0x1004, 0, 0, list) > 0) {                 ; no rows: don't save an empty preview (a message box may say why)
            Shots.Diagnose(window)
            throw Error("sync preview: no changes listed after 15 s")
        }
        Sleep(500)
        return window
    }

    ; The action panel's "Windows Context Menu" is off by default
    static EnableContextMenu() {
        settingsFile := Shots.AppDir "\Data\ALTRun.json"
        settings := JSON.Parse(FileRead(settingsFile, "UTF-8"))
        settings["Features"]["ContextMenu"] := Map("Enabled", 1)
        FileDelete(settingsFile)
        FileAppend(JSON.Stringify(settings, 4), settingsFile, "UTF-8")
    }

    ; Made-up usage counts for the last 30 days, so the Usage page has a chart and a table to show
    static WriteUsage() {
        features := [["Show", 18], ["Applications", 9], ["CustomCommands", 4], ["FileSearch", 5], ["WebSearch", 3], ["Calculator", 2]
                   , ["Clipboard", 3], ["Snippets", 1], ["SnippetExpand", 4], ["QuickSwitch", 2], ["AutoDate", 1], ["System", 1]]
        days := Map()
        Loop 30 {
            n := A_Index
            if (Mod(n, 7) = 0 || Mod(n, 7) = 6)                             ; weekends: little use
                continue
            day := Map()
            for pair in features
                day[pair[1]] := Max(0, pair[2] + Mod(n * 7 + StrLen(pair[1]) * 3, 5) - 2)
            days[FormatTime(DateAdd(A_Now, -(n - 1), "Days"), "yyyy-MM-dd")] := day
        }
        usageFile := Shots.AppDir "\Data\Usage.json"
        try FileDelete(usageFile)
        FileAppend(JSON.Stringify(Map("Since", FormatTime(DateAdd(A_Now, -29, "Days"), "yyyy-MM-dd"), "Days", days)), usageFile, "UTF-8")
    }

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

    ; Preferences → Advanced → Change Location...: the dialog with the per-PC OneDrive entry chosen; openList: its list open too (the dialog and the list in one shot)
    static DataLocation(openList := false) {
        prefs := Shots.Preferences()
        button := 0
        for ctrl in WinGetControlsHwnd(prefs)
            if (!button && ControlGetVisible(ctrl) && WinGetClass(ctrl) = "Button" && ControlGetText(ctrl) = "Change Location...")
                button := ctrl
        if !button
            throw Error("advanced page: Change Location... button not found")
        PostMessage(0xF5, 0, 0, button)                                     ; BM_CLICK, posted: the dialog is modal
        window := WinWait("Data Folder Location ahk_class AutoHotkeyGUI ahk_pid " Shots.Pid, , 20)
        if !window
            throw Error("Change Location...: dialog not shown")
        WinSetAlwaysOnTop(1, window)
        Sleep(800)
        combo := ""
        for ctrl in WinGetControlsHwnd(window)
            if (WinGetClass(ctrl) = "ComboBox")
                combo := ctrl
        ControlChooseString("%OneDrive%\ALTRun\Data-%COMPUTERNAME%", combo)                  ; the per-PC OneDrive entry: the dialog's lines follow it
        Sleep(500)
        if !openList
            return window
        SendMessage(0x14F, 1, 0, combo)                                     ; CB_SHOWDROPDOWN
        Sleep(800)
        list := WinExist("ahk_class ComboLBox ahk_pid " Shots.Pid)
        return list ? [window, list] : window
    }

    ; A demo OneDrive folder for the data location dialog (the runner has no OneDrive): ALTRun started after this inherits %OneDrive%
    static DemoOneDrive() {
        DirCreate(Shots.DemoDir "\OneDrive")
        EnvSet("OneDrive", Shots.DemoDir "\OneDrive")
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
        logText := ""                                                       ; the end of ALTRun's own log (errors from timers and background work go there)
        try logText := FileRead(A_Temp "\ALTRun.log")
        lines := StrSplit(Trim(logText, "`r`n"), "`n")
        Loop Min(15, lines.Length)
            Shots.Log("ALTRun.log: " RTrim(lines[lines.Length - Min(15, lines.Length) + A_Index], "`r"))
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
