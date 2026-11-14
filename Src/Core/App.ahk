;===============================================================================
; App.ahk - start-up / tray menu / global hotkeys / exit (AutoHotkey v2)
;-------------------------------------------------------------------------------
; ALTRun.ahk only does the #Includes and calls App.Start(); the start-up order is here:
;   settings -> language -> learned ranking -> theme -> search features -> search window -> tray -> hotkeys
;   -> extensions (snippet expansion / QuickSwitch / AutoDate) -> start-up and other shortcuts -> command-line arguments
;
; Command-line arguments:
;   -Startup        used when starting with Windows; doesn't show the search window
;   -Reloaded       after reloading (saved settings etc.); doesn't show the search window
;   -Preferences PAGE [X Y]  after reloading, open that page of Preferences (its name, e.g. Hotkeys; at X, Y: "Apply" was clicked in Preferences)
;   -SendTo <path...>  Explorer's "Send to" menu: add the files / folders as custom commands (one opens the edit dialog)
;   -Update         check for updates and install a new version right away (asks first when ALTRun is already running and gets it handed over)
;   -Updated <version>  the new version started by a one-click update: says it was updated; doesn't show the search window
;
; Only one ALTRun runs: when one is already running, a new one hands it the command-line arguments (double-clicking the exe = show
; the search window, Send to = add commands, start with Windows = nothing) and exits, see HandOff
;
; Usage (from other modules):
;   App.Notify("...")              a short notice after an action (HUD)
;   App.Toast("...")               a Windows notification for a background event (the HUD while the tray icon is hidden)
;   App.FocusPreviousWindow()      back to the window that was active before ALTRun was shown (for pasting)
;   App.OpenPreferences() / App.EditSettingsFile() / App.Reload() / App.Restart(args)
;   App.Quit() / App.RebuildIndex()
;===============================================================================

class App {
    static Name    := "ALTRun"
    static Version := "2026.11.13.10"
    static RepoUrl := "https://github.com/zhugecaomao/ALTRun"
    static Website := "https://zhugecaomao.github.io/ALTRun/"
    static IconFile := A_ScriptDir "\Resources\ALTRun.ico"                    ; tray, windows, shortcuts (the compiled exe has the same icon)
    static PreviousWindow := 0
    ; Installed from the Microsoft Store (an MSIX package): its package family name, "" otherwise. The package folder changes with each
    ; update, so shortcuts and tasks that start ALTRun use its execution alias (ExePath); the Store updates it (see UpdateChecker)
    static PackageFamily := App._PackageFamily()
    static _settingsTime := ""
    static _watchTimer := ""
    static _handOffNote := ""

    static Start() {
        if (A_Args.Length && AdminHelper.IsHelperArg(A_Args[1]))            ; the helper for administrator windows; it does only that
            return AdminHelper.Run(A_Args[1])
        if App.HandOff()
            ExitApp()
        if A_IsAdmin
            App._AllowNormalPrograms()
        started := Logger.Ms()
        ; Window commands (WinActivate etc.) no longer wait 100 ms by default: while waiting AHK handles messages recursively,
        ; and when showing the window with a constant stream of paint messages it would wait forever (the UI hangs). Use WinWaitActive where a wait is needed
        SetWinDelay(-1)
        Logger.Rotate()
        App._MoveLegacyResources()
        AppSettings.Load()
        Logger.Enabled := AppSettings.General["SaveLog"] ? true : false
        Logger.Debug("===== " App.Name " " App.Version (A_IsAdmin ? " (administrator)" : "") " starting =====")
        if (App._handOffNote != "")
            Logger.Debug(App._handOffNote)
        Logger.Time("startup: settings", started)                          ; log the time of each start-up stage when "Write debug log" is on
        phase := Logger.Ms()
        AppSettings.General["Language"] := I18n.Normalize(AppSettings.General["Language"])   ; the former "zh" -> "zh-CN"
        I18n.Init(AppSettings.General["Language"])
        Knowledge.Load()
        Usage.Load()
        ThemeManager.Load(AppSettings.Appearance["Theme"], AppSettings.Appearance["Density"])
        Logger.Time("startup: learning, usage, theme", phase)

        for provider in [ClipboardProvider, ApplicationProvider, CustomCommandProvider, SnippetProvider, SystemProvider
                        , CalculatorProvider, WebSearchProvider, BookmarkProvider, BrowserHistoryProvider, WindowProvider, ScriptProvider, FileSearchProvider, TerminalProvider, HelpProvider, RecentProvider, SyncProvider]
            ProviderRegistry.Register(provider)
        ProviderRegistry.InitAll()

        phase := Logger.Ms()
        App._SetIcon()                                                      ; before the windows are created: the window icon follows the tray icon
        SearchWindow.Create()
        Logger.Time("startup: search window", phase), phase := Logger.Ms()
        App._CreateTrayMenu()
        App._RegisterHotkeys()
        Logger.Time("startup: tray, hotkeys", phase), phase := Logger.Ms()
        SnippetExpander.Init()
        QuickSwitch.Init(AppSettings.Extension("QuickSwitch"))
        AutoDate.Init(AppSettings.Extension("AutoDate"))
        PTToolsWindow.Load(AppSettings.Extension("PTTools"))
        Logger.Time("startup: snippets, extensions", phase), phase := Logger.Ms()
        App._UpdateShellShortcuts()
        OnExit((reason, *) => App._OnExit(reason))
        Logger.Time("startup: shortcuts", phase)

        if (ThemeManager.ReadError != "")
            App.Toast(App.ThemeErrorText(ThemeManager.ReadError), 8000)
        if (AppSettings.ImportedFrom != "")
            App.Toast(I18n.T("Settings.ImportedIni", AppSettings.ImportedFrom), 6000)
        else if AppSettings.MigratedFrom
            App.Toast(I18n.T("Settings.Migrated", AppSettings.MigratedFrom, SchemaMigration.BackupFile(AppSettings.File, AppSettings.MigratedFrom)), 5000)
        else if (AppSettings.MovedFrom != "")
            App.Toast(I18n.T("Settings.Moved", AppSettings.File), 5000)
        else if AppSettings.NewerVersion
            App.Toast(I18n.T("Settings.Newer", AppSettings.NewerVersion, AppSettings.CurrentVersion), 8000)
        UpdateChecker.CleanUp()
        SetTimer(() => (IconCache.LoadDisk(), ProviderRegistry.WarmUp(), SearchWindow._PrepareLater()), -500)   ; before the first input: load the icon cache, compute the search keys, prepare the empty search box
        SetTimer(() => IconCache.Prefetch(ProviderRegistry.IconSpecs()), -SearchWindow.IconPrefetchDelay)   ; then read the common items' icons in the background
        if AppSettings.General["CheckForUpdates"]
            UpdateChecker.Schedule()                                        ; checked in the background (see UpdateChecker); a new version shows in the search window
        DelayedStart.Schedule()                                             ; the delayed start tasks follow the list, 30 s later (only with entries)
        SyncJobs.Adopt()                                                    ; sync tasks a restart left running

        Logger.Time("startup: total", started)
        Logger.Flush(), Logger.Immediate := true                            ; from now on each log line is written at once (kept even if the program hangs and is killed)
        ErrorHandler.Install()                                              ; from now on an error nothing caught is logged and noted, without AutoHotkey's dialog
        Logger.TraceUntil := A_TickCount + 60000                            ; log finer steps for 1 minute after start-up (Logger.Trace)
        OnMessage(0x4A, (p*) => App._OnCopyData(p*))                        ; command-line arguments handed over by ALTRun instances started later (WM_COPYDATA)
        App._WatchWake()
        App._HandleCommandLine(A_Args)
    }

    ;---------------------------------------------------------------------------
    ; Single instance
    ;---------------------------------------------------------------------------
    ; When an ALTRun (the same program file) is already running, hand it the command-line arguments (WM_COPYDATA) and return true: this process exits.
    ; Unlike #SingleInstance Force it isn't closed, so Preferences being edited and open edit dialogs aren't lost.
    ; It is replaced only on a restart (/restart: applying settings, one-click update) or when it doesn't answer (hung, or an old version that doesn't know the message)
    static CopyDataId := 0x414C5452                                         ; "ALTR": tells it apart from the WM_COPYDATA of Everything / TC

    static HandOff() {
        running := App._RunningInstance()
        if !running
            return false
        if !App.IsRestart(DllCall("GetCommandLineW", "Str")) {
            DllCall("GetWindowThreadProcessId", "Ptr", running, "UInt*", &pid := 0)
            if App.TakesOver(A_IsAdmin, A_IsAdmin && Win.IsElevated(pid)) {
                App._handOffNote := "App: started as administrator, replacing the instance without administrator rights"
            } else {
                DllCall("AllowSetForegroundWindow", "UInt", pid)            ; windows it shows may come to the front (this process was just started by the user, so it has that right)
                if App.SendArgs(running, A_Args)
                    return true
                App._handOffNote := "App: running instance did not answer, replacing it"
            }
        }
        App._CloseInstance(running)
        return false
    }

    ; The (hidden) main window of another ALTRun from the same program file, 0 if none
    static _RunningInstance() {
        saved := A_DetectHiddenWindows
        DetectHiddenWindows(true)
        found := 0
        for hwnd in WinGetList("ahk_class AutoHotkey") {
            if (hwnd != A_ScriptHwnd && App.IsInstanceTitle(WinGetTitle(hwnd), A_ScriptFullPath)) {
                found := hwnd
                break
            }
        }
        DetectHiddenWindows(saved)
        return found
    }

    ; Main window title: "path - AutoHotkey v2..." when running the source, only the path for the compiled exe;
    ; case-insensitive (the path's case may differ when started from a shortcut)
    static IsInstanceTitle(title, scriptPath) {
        return (title = scriptPath || InStr(title, scriptPath " - AutoHotkey") = 1) ? true : false
    }

    ; started by App.Restart: the command line is "program" /restart ...
    ; This new one runs as administrator but the running one doesn't (right-click → Run as administrator, or the scheduled task started after the start-up shortcut):
    ; replace it, otherwise this one would hand over the arguments and exit, leaving normal rights. The other way round (normal to administrator) hands over as usual
    static TakesOver(selfAdmin, runningElevated) => (selfAdmin && !runningElevated)

    ; Running as administrator, Windows blocks messages from programs with normal rights (UIPI). Allow WM_COPYDATA: search results from Everything,
    ; the current folder Total Commander replies with (QuickSwitch), arguments from ALTRun instances started later; and TaskbarCreated, to recreate
    ; the tray icon after Explorer restarts. Each WM_COPYDATA handler checks its own identifier first and only handles the replies it waits for
    static _AllowNormalPrograms() {
        for msg in [0x4A, DllCall("RegisterWindowMessageW", "Str", "TaskbarCreated", "UInt")]
            DllCall("ChangeWindowMessageFilterEx", "Ptr", A_ScriptHwnd, "UInt", msg, "UInt", 1, "Ptr", 0)   ; MSGFLT_ALLOW
    }

    ; The Caps Lock on / off notice (CapsLockSwitch and the helper for administrator windows); not shown when Keyboard.CapsLockNotify is off
    static CapsLockNotice(on) {
        keyboard := AppSettings.Keyboard
        if (!keyboard.Has("CapsLockNotify") || keyboard["CapsLockNotify"])
            App.Notify(CapsLockSwitch.StateText(on), CapsLockSwitch.NotifyMs)
    }

    ; Shown after the name / version when running as administrator: " (Administrator)"
    static AdminSuffix() => A_IsAdmin ? " (" I18n.T("App.Admin") ")" : ""

    static IsRestart(commandLine) {
        return RegExMatch(commandLine, 'i)^\s*("[^"]*"|\S+)\s+/restart(\s|$)') ? true : false
    }

    ; The arguments are sent one per line; only a reply of 1 counts as received (false when it hangs for 3 s or an old version doesn't know it)
    ; reply: what it replied (the helper for administrator windows answers -Jump with a code, see AdminHelper.Refused)
    static SendArgs(hwnd, args, &reply := 0) {
        text := ""
        for arg in args
            text .= arg "`n"
        data := Buffer(StrPut(text, "UTF-16"))
        StrPut(text, data, "UTF-16")
        copyData := Buffer(A_PtrSize * 3, 0)                                ; COPYDATASTRUCT
        NumPut("UPtr", App.CopyDataId, copyData, 0)
        NumPut("UInt", data.Size, copyData, A_PtrSize)
        NumPut("Ptr", data.Ptr, copyData, A_PtrSize * 2)
        reply := 0
        sent := DllCall("SendMessageTimeoutW", "Ptr", hwnd, "UInt", 0x4A, "Ptr", A_ScriptHwnd, "Ptr", copyData.Ptr
                      , "UInt", 2, "UInt", 3000, "UPtr*", &reply)            ; 2 = SMTO_ABORTIFHUNG
        return (sent && reply = 1)
    }

    static _OnCopyData(wParam, lParam, msg, hwnd) {
        if (NumGet(lParam, 0, "UPtr") != App.CopyDataId)
            return                                                          ; not for us (replies from Everything, TC)
        args := []
        for arg in StrSplit(Win.CopyDataText(lParam), "`n")                ; within the length given, at most 32768 characters (the longest command line)
            if (arg != "")
                args.Push(arg)
        Logger.Debug("App: arguments from a new instance: " (args.Length ? args[1] " (" args.Length ")" : "none"))
        SetTimer(() => App._HandleCommandLine(args, true), -1)              ; reply first (the sender is waiting), then handle it: it may open an edit dialog
        return 1
    }

    ; Closes the other ALTRun (a normal exit, saving the learned ranking etc.); killed if it hasn't exited after 3 s
    static _CloseInstance(hwnd) {
        saved := A_DetectHiddenWindows
        DetectHiddenWindows(true)
        try {
            pid := WinGetPID(hwnd)
            WinClose(hwnd)
            if !WinWaitClose(hwnd, , 3)
                ProcessClose(pid)
        } catch as e
            Logger.Warn("App: could not close the other instance - " e.Message)
        DetectHiddenWindows(saved)
    }

    ;---------------------------------------------------------------------------
    ; Window focus / notifications
    ;---------------------------------------------------------------------------
    static RememberActiveWindow() {
        activeHwnd := WinExist("A")
        if (activeHwnd && !(IsObject(SearchWindow.Gui) && activeHwnd = SearchWindow.Gui.Hwnd))
            App.PreviousWindow := activeHwnd
    }

    static FocusPreviousWindow() {
        target := App.PreviousWindow
        if (!target || !WinExist("ahk_id " target))
            return false
        try {
            WinActivate("ahk_id " target)
            WinWaitActive("ahk_id " target, , 1)
        }
        return WinActive("ahk_id " target) ? true : false
    }

    ; A short notice after an action (copied, cleared...): a HUD in the theme's colors that disappears after duration ms, see Hud.ahk
    static Notify(text, duration := 1500) {
        Hud.Show(text, duration)
    }

    ; Something that happened in the background (update finished, settings imported, index rebuilt...): a Windows notification, still visible in the notification center later.
    ; Notifications belong to the tray icon; with the tray icon hidden the HUD is used instead
    static Toast(text, duration := 5000) {
        if A_IconHidden
            return Hud.Show(text, duration)
        try TrayTip(text, App.Name, 0x34)                                   ; 4 = tray icon, 0x10 = silent, 0x20 = large icon
        catch
            Hud.Show(text, duration)
    }

    ;---------------------------------------------------------------------------
    ; Commands (tray menu / system commands / hotkeys)
    ;---------------------------------------------------------------------------
    ; page: page number, or the page's key (e.g. "Prefs.Page.Advanced")
    static OpenPreferences(page := 1) {
        PreferencesWindow.Show(page)
    }

    ; Edit Data\ALTRun.json in Notepad directly; reloads after saving
    static EditSettingsFile() {
        App.Notify(I18n.T("Settings.EditHint"), 4000)
        App._settingsTime := FileGetTime(AppSettings.File, "M")
        if (App._watchTimer = "")
            App._watchTimer := Logger.Watch("settings watch", () => App._CheckSettingsChanged())
        WinTimer.Set(App._watchTimer, 1500)
        Run('notepad.exe "' AppSettings.File '"')
    }

    static _CheckSettingsChanged() {
        try {
            if (FileGetTime(AppSettings.File, "M") != App._settingsTime)
                App.Restart()
        }
    }

    ; Since 3.1 Res\ is called Resources\: files the user put in the old folder (e.g. DOSBox.exe for SPF2M)
    ; move to Resources\, old copies of files the new version ships (Kanji.txt) are deleted, and finally the empty Res\ is removed
    static _MoveLegacyResources() {
        legacyDir := A_ScriptDir "\Res", newDir := A_ScriptDir "\Resources"
        if !DirExist(legacyDir)
            return
        try {
            DirCreate(newDir)
            Loop Files, legacyDir "\*", "FD" {
                target := newDir "\" A_LoopFileName
                if InStr(A_LoopFileAttrib, "D") {
                    if !DirExist(target)
                        DirMove(A_LoopFileFullPath, target)
                } else if FileExist(target) {
                    FileDelete(A_LoopFileFullPath)
                } else {
                    FileMove(A_LoopFileFullPath, target)
                }
            }
            DirDelete(legacyDir)                                            ; only removes an empty folder; throws and keeps it if anything is left
            Logger.Debug("App: moved Res\ to Resources\")
        } catch as e {
            Logger.Error("App: cannot move Res\ to Resources\ - " e.Message)
        }
    }

    static RebuildIndex() {
        count := ApplicationProvider.Rebuild()
        App.Toast(I18n.T("Index.Done", count))
    }

    static Reload() {
        App.Restart()
    }

    ; Restarts ALTRun with command-line arguments (default -Reloaded: doesn't show the search window)
    static Restart(arguments := "-Reloaded") {
        App._exitBy := "restart " arguments
        App._LogExitRequest()
        App.WaitForEnterUp()
        if A_IsCompiled
            Run('"' A_ScriptFullPath '" /restart ' arguments)
        else
            Run('"' A_AhkPath '" /restart "' A_ScriptFullPath '" ' arguments)
        ExitApp()
    }

    ; The notice for a theme file that couldn't be read (e.g. a JSON error in a custom theme): which file and line. Light takes its place,
    ; whether it is the chosen theme or one of its "Base" themes
    static ThemeErrorText(readError) {
        SplitPath(readError["File"], &fileName)
        light := I18n.T("Theme.Light")
        return readError["Line"] ? I18n.T("Theme.ReadErrorLine", fileName, readError["Line"], light) : I18n.T("Theme.ReadError", fileName, light)
    }

    static Quit() {
        App._exitBy := "quit"
        App._LogExitRequest()
        App.WaitForEnterUp()
        ExitApp()
    }

    ; Waits until Enter is let go, at most maxMs, so its key-up comes to an ALTRun window still in front and not alone to the window that
    ; comes back once that one hides or ALTRun exits. That window is active before its focus is restored, and a key-up arriving while no
    ; window has the focus comes as WM_SYSKEYUP to the active window (probed on Windows CI, 2026-11-09): Total Commander's Synchronize
    ; directories window takes WM_SYSKEYUP with Enter as Alt+Enter and opens a properties window. So the search window waits before hiding
    ; for Quit and Reload (ResultItem.WaitKeysUp); Quit / Restart wait again for the other ways in (e.g. Apply in Preferences, whose window
    ; is still in front). Waiting after the search window hid didn't help. Usually it waits the time the key is held (tens of ms).
    ; Returns the milliseconds waited. isDown / wait are replaced in the tests
    static WaitForEnterUp(maxMs := 2000, isDown := "", wait := "") {
        isDown := IsObject(isDown) ? isDown : () => App.IsEnterDown()
        wait := IsObject(wait) ? wait : Sleep
        started := A_TickCount, waited := 0
        while (isDown() && waited < maxMs) {
            wait(10)
            waited := A_TickCount - started
        }
        if (waited && Logger.Enabled)
            Logger.Debug("App: waited " waited " ms for Enter to be released")
        return waited
    }

    ; Debug log: the keys held and the window in front when quitting / restarting is asked for, and how long after Enter ran a result
    ; (the exit note in _OnExit comes later, after the settings are saved)
    static _LogExitRequest() {
        if !Logger.Enabled
            return
        since := (SearchKeys.EnterAt && A_TickCount - SearchKeys.EnterAt < 5000) ? ", " (A_TickCount - SearchKeys.EnterAt) " ms after Enter" : ""
        Logger.Debug("App: " App._exitBy " asked" since "; keys held: " ((held := App.HeldKeys()) = "" ? "none" : held) "; foreground: " SearchWindow.ForegroundText())
    }

    ; For the log: why ALTRun exits, the keys still held down and the window in front (2026-10: a properties dialog opened
    ; in Total Commander after "quit", as if Alt+Enter reached it)
    static _exitBy := ""
    static ExitNote(reason, by, held, foreground) {
        return "App: exit (" reason (by = "" ? "" : ", " by) "); keys held: " (held = "" ? "none" : held) "; foreground: " foreground
    }

    ; Whether Enter is down, from the keyboard or from a program typing for the user. Not the physical state ("P"): AutoHotkey's hook
    ; counts only keys from a keyboard driver, and remote desktop tools (Chrome Remote Desktop, ...) inject every key with SendInput, so
    ; there Enter was never "physically" down and the search window hid at once (KeyDiag, 2026-11-09; probed on Windows CI)
    static IsEnterDown() => GetKeyState("Enter") || GetKeyState("NumpadEnter")

    ; Whether a key or mouse button is held: down as the system sees it (also when a remote desktop tool injects it, see IsEnterDown)
    ; or physically down (a hotkey that blocks the key leaves it logically up while held)
    static IsKeyDown(key) => GetKeyState(key) || GetKeyState(key, "P")

    ; The keys held down, e.g. "Enter+LAlt" ("" = none), from the keyboard or injected (see IsEnterDown). No NumpadEnter: it shares its
    ; virtual key with Enter, so it reads the same
    static HeldKeys() {
        held := ""
        for key in ["Enter", "LAlt", "RAlt", "LCtrl", "RCtrl", "LShift", "RShift", "LWin", "RWin"]
            if GetKeyState(key)
                held .= (held = "" ? "" : "+") key
        return held
    }

    ; About: the Advanced page of Preferences (icon, version, home page, check for updates)
    static About() {
        App.OpenPreferences("Prefs.Page.Advanced")
    }

    ; System information to attach to a problem report (Preferences → Advanced → Copy system info). No file paths and no search text
    static SystemInfo() {
        general := AppSettings.General, appearance := AppSettings.Appearance
        windows := A_OSVersion
        try windows .= "." RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "UBR")
        try windows .= " (" RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "DisplayVersion") ")"
        layouts := DllCall("GetKeyboardLayoutList", "Int", 0, "Ptr", 0, "Int")
        fileSearch := AppSettings.Feature("FileSearch")
        everythingState := (fileSearch.Has("UseEverything") && fileSearch["UseEverything"]) ? (Everything.IsRunning() ? "on, running" : "on, not running") : "off"
        dataDir := AppSettings.DataDir
        location := (dataDir != AppSettings.DefaultDataDir()) ? "custom" : (AppSettings.Portable ? "portable" : "AppData")
        lines := [
            App.Name " " App.Version " (" (A_IsCompiled ? "exe" : "source") ", AutoHotkey " A_AhkVersion ", " (A_PtrSize * 8) "-bit)",
            "Administrator: " (A_IsAdmin ? "yes" : "no"),
            "Windows: " windows,
            "Display: " MonitorGetCount() " monitor(s), " A_ScreenWidth "x" A_ScreenHeight ", scale " Round(A_ScreenDPI / 96 * 100) "%",
            "Language: " I18n.Lang " (setting " general["Language"] "), Windows " A_Language,
            Format("Input: {} keyboard layout(s), current {:08X}, CapsLock '{}', switch to English {}", layouts, Win.KeyboardLayout() & 0xFFFFFFFF, AppSettings.Keyboard["CapsLock"], AppSettings.Keyboard["SwitchToEnglishInput"]),
            "Theme: " appearance["Theme"] ", status bar " (AppSettings.Window["StatusBar"] ? "on" : "off"),
            "Everything: " everythingState,
            "Data folder: " location
        ]
        text := ""
        for line in lines
            text .= line "`n"
        return text
    }

    static OpenLog() {
        Logger.Flush()
        if FileExist(Logger.File)
            Path.OpenText(Logger.File)                                      ; the default program for .log (e.g. Notepad4), Notepad if there is none
    }

    ;---------------------------------------------------------------------------
    ; Startup helpers
    ;---------------------------------------------------------------------------
    static _CreateTrayMenu() {
        if !AppSettings.General["ShowTrayIcon"] {
            A_IconHidden := true
            return
        }
        tray := A_TrayMenu
        tray.Delete()
        tray.Add(I18n.T("Tray.Show"), (*) => SetTimer(() => SearchWindow.Show(), -100))
        tray.Add(I18n.T("Tray.Preferences"), (*) => App.OpenPreferences())
        tray.Add()
        tray.Add(I18n.T("Tray.RebuildIndex"), (*) => App.RebuildIndex())
        tray.Add(I18n.T("Tray.CheckUpdate"), (*) => UpdateChecker.Check())
        tray.Add()
        tray.Add(I18n.T("Tray.Reload"), (*) => App.Reload())
        tray.Add(I18n.T("Tray.Exit"), (*) => App.Quit())
        tray.Default := I18n.T("Tray.Show")
        tray.ClickCount := 1
        A_IconTip := App.Name App.AdminSuffix() " - " App._HotkeyText()
        A_IconHidden := false
    }

    static _HotkeyText() {
        return Win.HotkeyLabel(AppSettings.General["Hotkey"])
    }

    static _RegisterHotkeys() {
        for key in [AppSettings.General["Hotkey"], AppSettings.General["SecondaryHotkey"]] {
            if (key != "")
                App._TryHotkey(key, (*) => SearchWindow.Toggle(), true)
        }

        tap := AppSettings.General.Has("DoubleTap") ? AppSettings.General["DoubleTap"] : ""
        for key in App.TapKeys(tap) {                                       ; double-tap Ctrl / Shift (or only the left / right one) / left Alt to show
            App._TryHotkey("~" key, (*) => App._TapDown())
            App._TryHotkey("~" key " up", (*) => App._TapUp())
        }

        capsLock := AppSettings.Keyboard.Has("CapsLock") ? AppSettings.Keyboard["CapsLock"] : ""
        if CapsLockSwitch.IsMode(capsLock) {                                ; CapsLock switches the input language; Caps Lock only when held
            CapsLockSwitch.Mode := capsLock
            CapsLockSwitch.Register((key, fn) => App._TryHotkey(key, fn))
        }
        AdminHelper.Start()                                                 ; Caps Lock and Quick Switch in administrator windows, when those options are on

        if (AppSettings.General["SelectionHotkey"] != "")                  ; actions for the selection
            App._TryHotkey(AppSettings.General["SelectionHotkey"], (*) => SelectionActions.Run())

        clipboard := AppSettings.Feature("Clipboard")
        if (clipboard["Enabled"] && clipboard["Hotkey"] != "")
            App._TryHotkey(clipboard["Hotkey"], (*) => SearchWindow.Show(AppSettings.Feature("Clipboard")["Keyword"] " "))

        ; custom hotkeys: Key -> Action (a system command, Search typing its SearchText, "custom:<Id>" = a custom command, "script:<file name>"
        ; = a script); with a WinTitle only in that window
        for entry in AppSettings.Hotkeys {
            if !(entry is Map) || !entry.Has("Key") || !entry.Has("Action")
                continue
            winTitle := entry.Has("WinTitle") ? Trim(entry["WinTitle"]) : ""
            try {
                if (winTitle = "ALTRun")                                    ; only in ALTRun's search window (matching the title would include the Preferences window too)
                    HotIf(App._inSearchWindow)
                else if (winTitle != "")
                    HotIfWinActive(winTitle)
                App._TryHotkey(entry["Key"], App._HotkeyAction(App.HotkeyActionId(entry), App.HotkeySearchText(entry)))
            } catch as e {                                                  ; a wrong WinTitle makes HotIfWinActive fail too
                Logger.Error("App: cannot register hotkey " entry["Key"] " - " e.Message)
            }
            HotIf()
        }
    }

    static _inSearchWindow := (*) => SearchWindow.IsActive()               ; one function object: UnregisterHotkey turns off a hotkey of this context by it

    ; Turns off a custom hotkey registered by _RegisterHotkeys (its command was deleted), in the same context (WinTitle); one that
    ; wasn't registered is left alone
    static UnregisterHotkey(entry) {
        if !(entry is Map && entry.Has("Key") && entry["Key"] != "")
            return
        winTitle := entry.Has("WinTitle") ? Trim(entry["WinTitle"]) : ""
        try {
            if (winTitle = "ALTRun")
                HotIf(App._inSearchWindow)
            else if (winTitle != "")
                HotIfWinActive(winTitle)
            Hotkey(entry["Key"], "Off")
        } catch as e
            Logger.Warn("App: cannot turn off hotkey " entry["Key"] " - " e.Message)
        HotIf()
    }

    ; Reinstalls the keyboard hook after waking from sleep, unlocking or plugging in / unplugging a monitor (the same effect as reloading). With the PC on for days,
    ; the key states the hook remembers can drift from reality (keys released while locked never reach it), and Windows may silently remove a slow hook;
    ; after that hotkeys, hotstrings (snippet expansion), double-tap Ctrl and CapsLock fail or misbehave (e.g. Ctrl+D replacing a file name with "d")
    static WakeDelay := 1500, _wakeTimer := "", _wakeReason := ""
    static _WatchWake() {
        OnMessage(0x218, (wParam, *) => (wParam = 7 || wParam = 0x12) ? App._AfterWake("resume") : "")   ; WM_POWERBROADCAST: PBT_APMRESUMESUSPEND / AUTOMATIC
        OnMessage(0x2B1, (wParam, *) => (wParam = 8) ? App._AfterWake("unlock") : "")                  ; WM_WTSSESSION_CHANGE: WTS_SESSION_UNLOCK
        OnMessage(0x7E, (*) => App._AfterWake("display change"))                                        ; WM_DISPLAYCHANGE
        try DllCall("wtsapi32\WTSRegisterSessionNotification", "Ptr", A_ScriptHwnd, "UInt", 0)          ; NOTIFY_FOR_THIS_SESSION
    }

    ; these events often come in a row: wait a little after the last one and reinstall once
    static _AfterWake(reason) {
        App._wakeReason := reason
        if (App._wakeTimer = "")
            App._wakeTimer := Logger.Watch("reinstall hooks", () => App.ReinstallHooks(App._wakeReason))
        SetTimer(App._wakeTimer, -App.WakeDelay)
    }

    static ReinstallHooks(reason := "") {
        try InstallKeybdHook(true, true)
        if App.UsesMouseHotkeys(AppSettings.Hotkeys)
            try InstallMouseHook(true, true)
        Logger.Debug("App: keyboard hook reinstalled (" reason ")")
    }

    ; reinstall the mouse hook too when custom hotkeys use mouse buttons (e.g. ~MButton for PT Tools)
    static UsesMouseHotkeys(hotkeys) {
        for entry in hotkeys
            if (entry is Map && entry.Has("Key") && RegExMatch(entry["Key"], "i)(Button|Wheel)"))
                return true
        return false
    }

    ; Registers a hotkey and logs a failure; notify: also show a message (without the show hotkey the program is barely usable)
    static _TryHotkey(key, fn, notify := false) {
        try {
            Hotkey(key, fn)
            return true
        } catch as e {
            Logger.Error("App: cannot register hotkey " key " - " e.Message)
            if notify
                MsgBox(I18n.T("App.HotkeyFailed", Win.HotkeyLabel(key), e.Message), App.Name, "Icon! 0x40000")   ; 0x40000: always on top
            return false
        }
    }

    ; a separate method makes the closure, so each hotkey keeps its own Action
    ; Double-tap Ctrl / Shift / left Alt: two separate short presses of the same key (no other key in between, e.g. Ctrl+C), at most DoubleTapMs apart.
    ; Only the left Alt (#501): the right one is AltGr on many layouts (French, German, Polish...), pressed to type characters. A lone Alt
    ; press also opens the menu bar of some programs: the first tap opens it, the second closes it again
    ; Left Shift then right Shift isn't a double tap: typing capitals with alternating Shift keys opened the window (AutoHotkey forum)
    static DoubleTapMs := 400, TapHoldMs := 300
    static _tapLast := 0, _tapDownAt := 0, _tapHeld := false, _tapKey := ""

    ; The keys of the DoubleTap setting: "Ctrl" / "Shift" = either side, "LCtrl" / "RCtrl" / "LShift" / "RShift" / "LAlt" = only that one; [] = off
    static TapKeys(setting) {
        if (setting = "Ctrl" || setting = "Shift")
            return ["L" setting, "R" setting]
        return RegExMatch(setting, "i)^([LR](Ctrl|Shift)|LAlt)$") ? [setting] : []
    }

    static _TapDown() {
        if !App._tapHeld                                                    ; auto-repeat while held doesn't count
            App._tapHeld := true, App._tapDownAt := A_TickCount
    }

    static _TapUp() {
        App._tapHeld := false
        key := RegExReplace(A_ThisHotkey, "i)^~|\s+up$")                  ; "LCtrl" / "RShift"
        switch App.TapDecision(A_PriorKey, key, A_TickCount - App._tapDownAt, A_TickCount - App._tapLast, App._tapKey) {
            case "show":
                App._tapLast := 0
                SearchWindow.Toggle()
            case "first": App._tapLast := A_TickCount, App._tapKey := key
            default:      App._tapLast := 0
        }
    }

    ; priorKey: A_PriorKey (the last key pressed before the release); firstKey: the key of the tap before ("LShift"...).
    ; Returns "show" (second tap of the same key) / "first" (a first tap, also when the other side's key was tapped before) / "reset" (a combination or held too long)
    static TapDecision(priorKey, key, heldMs, sinceLastMs, firstKey := "") {
        expected := (key = "LCtrl") ? "LControl" : (key = "RCtrl") ? "RControl" : key
        if (!(priorKey = expected || priorKey = key) || heldMs > App.TapHoldMs)
            return "reset"
        return (sinceLastMs <= App.DoubleTapMs && (firstKey = "" || firstKey = key)) ? "show" : "first"
    }

    ; Remember the window the hotkey was pressed in first: commands such as TerminalHere work on that window, and {hwnd} / {pid}
    ; of a custom command become it. "custom:<Id>" runs the user's custom command with that Id, "script:<file name>" a script
    static _HotkeyAction(actionId, searchText := "") {
        if (SubStr(actionId, 1, 7) = "custom:")
            return (*) => (App.RememberActiveWindow(), CustomCommandProvider.RunFromHotkey(SubStr(actionId, 8)))
        if (SubStr(actionId, 1, 7) = "script:")
            return (*) => (App.RememberActiveWindow(), ScriptProvider.RunFromHotkey(SubStr(actionId, 8)))
        return (*) => (App.RememberActiveWindow(), SystemProvider.RunCommand(actionId, searchText))
    }

    ; A custom hotkey's action Id
    static HotkeyActionId(entry) => entry.Has("Action") ? entry["Action"] "" : ""

    ; The search text of a custom hotkey (for example "folder {selection}"): only Search types it into the search box; "" for any other
    static HotkeySearchText(entry) {
        action := entry.Has("Action") ? entry["Action"] "" : ""
        if (action = "Search")
            return entry.Has("SearchText") ? entry["SearchText"] "" : ""
        return ""
    }

    static _SetIcon() {
        if FileExist(App.IconFile)
            try TraySetIcon(App.IconFile)
    }

    ; function: replaced in the tests. Windows 7 has no GetCurrentPackageFamilyName (Windows 8 and later): not packaged then
    static _PackageFamily(function := "GetCurrentPackageFamilyName") {
        size := 0
        try code := DllCall(function, "UInt*", &size, "Ptr", 0)
        catch
            return ""
        if (code != 122)                                                    ; ERROR_INSUFFICIENT_BUFFER: there is a package (else APPMODEL_ERROR_NO_PACKAGE)
            return ""
        buf := Buffer(size * 2, 0)
        DllCall(function, "UInt*", &size, "Ptr", buf)
        return StrGet(buf)
    }

    ; The program that starts ALTRun, and its folder: packaged, the execution alias "altrun" (AppxManifest.xml), which stays the same
    ; across updates and starts it in its package; otherwise the running program
    static ExePath => App.PackageFamily != "" ? EnvGet("LOCALAPPDATA") "\Microsoft\WindowsApps\altrun.exe" : A_ScriptFullPath
    static ExeDir => App.PackageFamily != "" ? EnvGet("LOCALAPPDATA") "\Microsoft\WindowsApps" : A_ScriptDir

    ; The start-up / Explorer "Send to" / Start menu shortcuts, created or deleted as set.
    ; Only ALTRun's own are deleted (see IsOwnShortcut): a shortcut of the same name the user put there is left alone.
    ; Packaged, no Start menu shortcut of its own: the package has its Start menu entry
    static _UpdateShellShortcuts() {
        general := AppSettings.General
        target := A_IsCompiled ? App.ExePath : A_AhkPath
        prefix := A_IsCompiled ? "" : '"' A_ScriptFullPath '" '
        icon := (!A_IsCompiled && FileExist(App.IconFile)) ? App.IconFile : ""   ; when running the source, the shortcuts don't use the AutoHotkey icon
        sendTo := RegExReplace(A_StartMenu, "\\Start Menu$", "\SendTo") "\ALTRun.lnk"
        description := App.Name " - " I18n.T("App.Tagline")
        for shortcut in [
            [general["LaunchAtLogin"], A_Startup "\" App.StartupLinkName, "-Startup"],
            [general["SendToMenu"], sendTo, "-SendTo"],
            [general["StartMenuShortcut"] && App.PackageFamily = "", A_Programs "\ALTRun.lnk", ""]
        ] {
            try {
                arguments := Trim(prefix shortcut[3])
                if !shortcut[1]
                    App.RemoveOwnShortcut(shortcut[2], shortcut[3])
                else if !App.ShortcutMatches(shortcut[2], target, App.ExeDir, arguments, description, icon)
                    FileCreateShortcut(target, shortcut[2], App.ExeDir, arguments, description, icon)
            } catch as e {
                Logger.Error("App: shortcut " shortcut[2] " - " e.Message)
            }
        }
    }

    ; "Launch ALTRun at login" can also be turned off in Task Manager → Startup apps: the shortcut stays, and Windows marks it off in
    ; StartupApproved (see DelayedStart.SetEntryOff). Preferences shows the option as off then, and turns it on again there when ticked.
    ; ALTRun itself never turns it back on at start-up: what the user chose in Task Manager stands
    static StartupLinkName := "ALTRun.lnk"
    static StartupApprovedKey => "HKCU" DelayedStart.ApprovedKey "StartupFolder"

    static LaunchOffInTaskManager(name := App.StartupLinkName) {
        data := ""
        try data := RegRead(App.StartupApprovedKey, name)
        return DelayedStart.IsOffData(data)
    }

    ; Ticked in Preferences: on in Task Manager too (only when it is off there; nothing is written otherwise)
    static TurnOnInTaskManager(name := App.StartupLinkName) {
        if !App.LaunchOffInTaskManager(name)
            return false
        try {
            RegWrite(DelayedStart.ApprovedData(false), "REG_BINARY", App.StartupApprovedKey, name)
            return true
        } catch as e {
            Logger.Error("App: cannot turn on the start-up shortcut in Task Manager - " e.Message)
            return false
        }
    }

    ; A shortcut that is already right isn't rewritten: rewriting 3 shortcuts at every start-up calls the Windows shell and Explorer handles the file changes;
    ; at boot Explorer / OneDrive are busy, and it once took 11 s
    static ShortcutMatches(path, target, dir, arguments, description, icon := "") {
        if !FileExist(path)
            return false
        outTarget := "", outDir := "", outArgs := "", outDescription := "", outIcon := ""
        try FileGetShortcut(path, &outTarget, &outDir, &outArgs, &outDescription, &outIcon)
        catch
            return false
        return (outTarget = target && outDir = dir && outArgs == arguments && outDescription == description && outIcon = icon)
    }

    ; Deletes a shortcut ALTRun created; one the user created is kept. Returns whether it was deleted
    static RemoveOwnShortcut(path, flag) {
        if !FileExist(path)
            return false
        arguments := "", description := ""
        try FileGetShortcut(path, , , &arguments, &description)
        catch
            return false
        if !App.IsOwnShortcut(arguments, description, flag)
            return false
        FileDelete(path)
        return true
    }

    ; A shortcut ALTRun created: its comment is "ALTRun - <tagline>" (any language), or its arguments end with a switch ALTRun uses (-Startup / -SendTo)
    static IsOwnShortcut(arguments, description, flag) {
        if (SubStr(description, 1, StrLen(App.Name " - ")) = App.Name " - ")
            return true
        return (flag != "" && RegExMatch(arguments, "i)(^|\s)\Q" flag "\E$")) ? true : false
    }

    ; forwarded: arguments handed over by an ALTRun started later (no arguments = the exe was double-clicked: show the search window)
    static _HandleCommandLine(args, forwarded := false) {
        if (args.Length >= 2 && args[1] = "-SendTo") {
            paths := []                                                     ; with several files selected, all their paths come in at once
            Loop args.Length - 1 {
                target := args[A_Index + 1]
                if (SubStr(target, -4) = ".lnk") {
                    try {
                        FileGetShortcut(target, &linkTarget)
                        if (linkTarget != "")
                            target := linkTarget
                    }
                }
                paths.Push(target)
            }
            CustomCommandProvider.AddFromPaths(paths)
            return
        }
        if (args.Length >= 1 && (args[1] = "-Startup" || args[1] = "-Reloaded"))
            return
        if (args.Length >= 1 && args[1] = "-Updated") {
            App.Toast(I18n.T("Update.Done", App.Version), 5000)
            return
        }
        if (args.Length >= 2 && args[1] = "-CapsLockState") {               ; the helper in an administrator window toggled Caps Lock: show the notice here
            App.CapsLockNotice(args[2] = "1")
            return
        }
        if (args.Length >= 2 && args[1] = AdminHelper.HotkeyArg) {           ; the helper caught a Quick Switch hotkey in an administrator dialog: only which one
            AdminHelper.HotkeyPressed(args[2])
            return
        }
        if (args.Length >= 2 && args[1] = "-AdminHelperNote") {             ; a line for the log from the helper for administrator windows
            Logger.Debug("AdminHelper: " SubStr(args[2], 1, 300))
            return
        }
        if (args.Length >= 1 && args[1] = "-Update") {                     ; handed over to a running ALTRun: any program could have sent it (to one
            UpdateChecker.Check(!forwarded)                                 ; running as administrator too), so it asks first, like a check from the tray
            return
        }
        if (args.Length >= 1 && args[1] = "-Preferences") {
            opened := PreferencesWindow.ParseArgs(args)
            PreferencesWindow.Show(opened.Page, opened.X, opened.Y)
            return
        }
        if (!forwarded && AppSettings.MigratedFrom)
            return                                                          ; the upgrade notice is showing: don't show the window right away
        SearchWindow.Show()
        if (!forwarded && AppSettings.MovedFrom = "")                       ; don't cover the "settings file moved to Data" notice
            App.Notify(I18n.T("App.Running", App._HotkeyText()), 3000)
    }

    ; a non-zero return from an OnExit callback cancels the exit, so nothing is returned here
    static _OnExit(reason := "") {
        if Logger.Enabled
            Logger.Debug(App.ExitNote(reason, App._exitBy, App.HeldKeys(), SearchWindow.ForegroundText()))
        try ErrorHandler.WriteCounts()                                     ; uncaught errors counted but not written yet
        AdminHelper.Stop()
        try SyncPreview.SaveOnExit()                                       ; where an open sync window is (closing one saves it, quitting doesn't close them)
        catch as e
            Logger.Error("App: cannot save the sync window's place on exit - " e.Message)
        SyncJobs.Shutdown(App._exitBy)                                      ; quitting stops the sync tasks, a restart hands them over
        try ApplicationProvider.SavePending()                              ; indexes rebuilt in the background but not written yet
        catch as e
            Logger.Error("App: cannot save the app index on exit - " e.Message)
        try FileIndex.SavePending()
        catch as e
            Logger.Error("App: cannot save the file index on exit - " e.Message)
        Knowledge.Save()
        Usage.Save()
        ClipboardProvider.Save()
        IconCache.SaveDisk()
        Logger.Flush(), Logger.Close()
    }
}
