;===============================================================================
; App.ahk - 程序启动 / 托盘菜单 / 全局热键 / 退出 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; ALTRun.ahk 只负责 #Include 和调用 App.Start(), 启动顺序都在这里:
;   设置 -> 语言 -> 学习记录 -> 主题 -> 搜索功能 -> 搜索窗口 -> 托盘 -> 热键
;   -> 扩展 (片段自动展开 / QuickSwitch / AutoDate) -> 开机启动等快捷方式 -> 命令行参数
;
; 命令行参数:
;   -Startup        开机自启动时使用, 不弹出搜索窗口
;   -Reloaded       重新载入后 (保存设置等), 不弹出搜索窗口
;   -Preferences N [X Y]  重新载入后打开偏好设置的第 N 页 (在 X, Y 位置: 偏好设置里点了 "应用")
;   -SendTo <path...>  资源管理器 "发送到" 菜单: 把文件/文件夹添加为自定义命令 (1 个弹出编辑对话框)
;   -Update         检查更新, 有新版本就直接下载安装 (不询问)
;   -Updated <版本>  一键更新后启动的新版本: 提示已更新, 不弹出搜索窗口
;
; 用法 (其它模块里):
;   App.Notify("...")              操作后的简短提示 (HUD)
;   App.Toast("...")               后台事件的 Windows 通知 (隐藏托盘图标时用 HUD)
;   App.FocusPreviousWindow()      回到呼出 ALTRun 之前的窗口 (粘贴用)
;   App.OpenPreferences() / App.EditSettingsFile() / App.Reload() / App.Restart(args)
;   App.Quit() / App.RebuildIndex()
;===============================================================================

class App {
    static Name    := "ALTRun"
    static Version := "2026.10.09"
    static RepoUrl := "https://github.com/zhugecaomao/ALTRun"
    static Website := "https://zhugecaomao.github.io/ALTRun/"
    static IconFile := A_ScriptDir "\Resources\ALTRun.ico"                    ; 托盘、窗口、快捷方式 (编译后的 exe 里也有同一个图标)
    static PreviousWindow := 0
    static _settingsTime := ""
    static _watchTimer := ""

    static Start() {
        started := Logger.Ms()
        Logger.Rotate()
        App._MoveLegacyResources()
        AppSettings.Load()
        Logger.Enabled := AppSettings.General["SaveLog"] ? true : false
        Logger.Debug("===== " App.Name " " App.Version " starting =====")
        Logger.Time("startup: settings", started)                          ; 打开 "写入调试日志" 时记录启动各阶段的耗时
        phase := Logger.Ms()
        AppSettings.General["Language"] := I18n.Normalize(AppSettings.General["Language"])   ; 以前的 "zh" -> "zh-CN"
        I18n.Init(AppSettings.General["Language"])
        Knowledge.Load()
        Usage.Load()
        ThemeManager.Load(AppSettings.Appearance["Theme"])
        Logger.Time("startup: learning, usage, theme", phase)

        for provider in [ClipboardProvider, ApplicationProvider, CustomCommandProvider, SnippetProvider, SystemProvider
                        , CalculatorProvider, WebSearchProvider, BookmarkProvider, WindowProvider, ScriptProvider, FileSearchProvider, TerminalProvider, HelpProvider, RecentProvider]
            ProviderRegistry.Register(provider)
        ProviderRegistry.InitAll()

        phase := Logger.Ms()
        App._SetIcon()                                                      ; 在创建窗口之前: 窗口的图标跟随托盘图标
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
        OnExit((*) => App._OnExit())
        Logger.Time("startup: shortcuts", phase)

        if (AppSettings.ImportedFrom != "")
            App.Toast(I18n.T("Settings.ImportedIni", AppSettings.ImportedFrom), 6000)
        else if AppSettings.MigratedFrom
            App.Toast(I18n.T("Settings.Migrated", AppSettings.MigratedFrom, SchemaMigration.BackupFile(AppSettings.File, AppSettings.MigratedFrom)), 5000)
        else if (AppSettings.MovedFrom != "")
            App.Toast(I18n.T("Settings.Moved", AppSettings.File), 5000)
        UpdateChecker.CleanUp()
        SetTimer(() => ProviderRegistry.WarmUp(), -500)                    ; 第一次输入前算好搜索 Key 和图标
        if AppSettings.General["CheckForUpdates"]
            UpdateChecker.Schedule()                                        ; 后台定时检查 (见 UpdateChecker), 新版本显示在搜索窗口里

        Logger.Time("startup: total", started)
        App._HandleCommandLine()
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

    ; 操作后的简短提示 (已复制、已清空...): 跟随主题的 HUD, duration 毫秒后消失, 见 Hud.ahk
    static Notify(text, duration := 1500) {
        Hud.Show(text, duration)
    }

    ; 后台发生的事 (更新完成、导入设置、索引重建完成...): Windows 通知, 之后还能在通知中心看到。
    ; 通知挂在托盘图标上, 隐藏了托盘图标时改用 HUD
    static Toast(text, duration := 5000) {
        if A_IconHidden
            return Hud.Show(text, duration)
        try TrayTip(text, App.Name, 0x34)                                   ; 4 = 托盘图标, 0x10 = 不响, 0x20 = 大图标
        catch
            Hud.Show(text, duration)
    }

    ;---------------------------------------------------------------------------
    ; Commands (tray menu / system commands / hotkeys)
    ;---------------------------------------------------------------------------
    ; page: 页码, 或页面的键 (例如 "Prefs.Page.Advanced")
    static OpenPreferences(page := 1) {
        PreferencesWindow.Show(page)
    }

    ; 直接用记事本编辑 Data\ALTRun.json, 保存后自动重新载入
    static EditSettingsFile() {
        App.Notify(I18n.T("Settings.EditHint"), 4000)
        App._settingsTime := FileGetTime(AppSettings.File, "M")
        if (App._watchTimer = "")
            App._watchTimer := () => App._CheckSettingsChanged()
        SetTimer(App._watchTimer, 1500)
        Run('notepad.exe "' AppSettings.File '"')
    }

    static _CheckSettingsChanged() {
        try {
            if (FileGetTime(AppSettings.File, "M") != App._settingsTime)
                App.Restart()
        }
    }

    ; 3.1 起 Res\ 改名为 Resources\: 旧文件夹里用户自己放的文件 (例如 SPF2M 的 DOSBox.exe)
    ; 移到 Resources\, 新版本已经带有的文件 (Kanji.txt) 直接删掉旧的, 最后删除空的 Res\
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
            DirDelete(legacyDir)                                            ; 只删空文件夹, 还有内容时抛错保留
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

    ; 重新启动 ALTRun, 带上命令行参数 (默认 -Reloaded: 不弹出搜索窗口)
    static Restart(arguments := "-Reloaded") {
        if A_IsCompiled
            Run('"' A_ScriptFullPath '" /restart ' arguments)
        else
            Run('"' A_AhkPath '" /restart "' A_ScriptFullPath '" ' arguments)
        ExitApp()
    }

    static Quit() {
        ExitApp()
    }

    ; 关于: 偏好设置的高级页 (图标、版本、项目主页、检查更新)
    static About() {
        App.OpenPreferences("Prefs.Page.Advanced")
    }

    static OpenLog() {
        Logger.Flush()
        if FileExist(Logger.File)
            Run('notepad.exe "' Logger.File '"')
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
        A_IconTip := App.Name " - " App._HotkeyText()
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
        if (tap = "Ctrl" || tap = "Shift") {                                ; 双击 Ctrl / Shift 呼出
            for side in ["L", "R"] {
                App._TryHotkey("~" side tap, (*) => App._TapDown())
                App._TryHotkey("~" side tap " up", (*) => App._TapUp())
            }
        }

        if (AppSettings.General["SelectionHotkey"] != "")                  ; 选中内容的操作
            App._TryHotkey(AppSettings.General["SelectionHotkey"], (*) => SelectionActions.Run())

        clipboard := AppSettings.Feature("Clipboard")
        if (clipboard["Enabled"] && clipboard["Hotkey"] != "")
            App._TryHotkey(clipboard["Hotkey"], (*) => SearchWindow.Show(AppSettings.Feature("Clipboard")["Keyword"] " "))

        ; 自定义热键: Key -> 系统命令 Action, WinTitle 非空时只在该窗口里生效
        for entry in AppSettings.Hotkeys {
            if !(entry is Map) || !entry.Has("Key") || !entry.Has("Action")
                continue
            winTitle := entry.Has("WinTitle") ? Trim(entry["WinTitle"]) : ""
            try {
                if (winTitle = "ALTRun")                                    ; 只在 ALTRun 的搜索窗口里 (标题匹配会连偏好设置窗口也算上)
                    HotIf((*) => SearchWindow.IsActive())
                else if (winTitle != "")
                    HotIfWinActive(winTitle)
                App._TryHotkey(entry["Key"], App._HotkeyAction(entry["Action"]))
            } catch as e {                                                  ; WinTitle 写错时 HotIfWinActive 也会出错
                Logger.Error("App: cannot register hotkey " entry["Key"] " - " e.Message)
            }
            HotIf()
        }
    }

    ; 注册热键, 失败时写日志; notify: 同时弹窗提示 (呼出热键注册不上时程序基本没法用)
    static _TryHotkey(key, fn, notify := false) {
        try {
            Hotkey(key, fn)
            return true
        } catch as e {
            Logger.Error("App: cannot register hotkey " key " - " e.Message)
            if notify
                MsgBox("Cannot register hotkey " key ":`n" e.Message, App.Name, 48)
            return false
        }
    }

    ; 单独一个方法生成闭包, 每个热键各自记住自己的 Action
    ; 双击 Ctrl / Shift: 两次单独的短按 (中间没有按别的键, 例如 Ctrl+C), 间隔不超过 DoubleTapMs
    static DoubleTapMs := 400, TapHoldMs := 300
    static _tapLast := 0, _tapDownAt := 0, _tapHeld := false

    static _TapDown() {
        if !App._tapHeld                                                    ; 按住不放时的自动重复不算
            App._tapHeld := true, App._tapDownAt := A_TickCount
    }

    static _TapUp() {
        App._tapHeld := false
        key := RegExReplace(A_ThisHotkey, "i)^~|\s+up$")                  ; "LCtrl" / "RShift"
        switch App.TapDecision(A_PriorKey, key, A_TickCount - App._tapDownAt, A_TickCount - App._tapLast) {
            case "show":
                App._tapLast := 0
                SearchWindow.Toggle()
            case "first": App._tapLast := A_TickCount
            default:      App._tapLast := 0
        }
    }

    ; priorKey: A_PriorKey (松开前最后按下的键); 返回 "show" (第二下) / "first" (第一下) / "reset" (是组合键或按太久)
    static TapDecision(priorKey, key, heldMs, sinceLastMs) {
        expected := (key = "LCtrl") ? "LControl" : (key = "RCtrl") ? "RControl" : key
        if (!(priorKey = expected || priorKey = key) || heldMs > App.TapHoldMs)
            return "reset"
        return (sinceLastMs <= App.DoubleTapMs) ? "show" : "first"
    }

    static _HotkeyAction(actionId) {
        return (*) => SystemProvider.RunCommand(actionId)
    }

    static _SetIcon() {
        if FileExist(App.IconFile)
            try TraySetIcon(App.IconFile)
    }

    ; 开机启动 / 资源管理器 "发送到" / 开始菜单 三个快捷方式, 按设置创建或删除。
    ; 只删 ALTRun 自己建的 (见 IsOwnShortcut): 用户自己放在同一位置、同名的快捷方式不动
    static _UpdateShellShortcuts() {
        general := AppSettings.General
        target := A_IsCompiled ? A_ScriptFullPath : A_AhkPath
        prefix := A_IsCompiled ? "" : '"' A_ScriptFullPath '" '
        icon := (!A_IsCompiled && FileExist(App.IconFile)) ? App.IconFile : ""   ; 运行源码时快捷方式不用 AutoHotkey 的图标
        sendTo := RegExReplace(A_StartMenu, "\\Start Menu$", "\SendTo") "\ALTRun.lnk"
        for shortcut in [
            [general["LaunchAtLogin"], A_Startup "\ALTRun.lnk", "-Startup"],
            [general["SendToMenu"], sendTo, "-SendTo"],
            [general["StartMenuShortcut"], A_Programs "\ALTRun.lnk", ""]
        ] {
            try {
                if shortcut[1]
                    FileCreateShortcut(target, shortcut[2], A_ScriptDir, Trim(prefix shortcut[3]), App.Name " - " I18n.T("App.Tagline"), icon)
                else
                    App.RemoveOwnShortcut(shortcut[2], shortcut[3])
            } catch as e {
                Logger.Error("App: shortcut " shortcut[2] " - " e.Message)
            }
        }
    }

    ; 删掉 ALTRun 建的快捷方式; 不是 ALTRun 建的 (用户自己建的) 保留。返回是否删了
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

    ; ALTRun 建的快捷方式: 备注是 "ALTRun - 宣传语" (任何语言), 或者参数以 ALTRun 用的开关结尾 (-Startup / -SendTo)
    static IsOwnShortcut(arguments, description, flag) {
        if (SubStr(description, 1, StrLen(App.Name " - ")) = App.Name " - ")
            return true
        return (flag != "" && RegExMatch(arguments, "i)(^|\s)\Q" flag "\E$")) ? true : false
    }

    static _HandleCommandLine() {
        if (A_Args.Length >= 2 && A_Args[1] = "-SendTo") {
            paths := []                                                     ; 选中了几个文件, 就一次传进来几个路径
            Loop A_Args.Length - 1 {
                target := A_Args[A_Index + 1]
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
        if (A_Args.Length >= 1 && (A_Args[1] = "-Startup" || A_Args[1] = "-Reloaded"))
            return
        if (A_Args.Length >= 1 && A_Args[1] = "-Updated") {
            App.Toast(I18n.T("Update.Done", App.Version), 5000)
            return
        }
        if (A_Args.Length >= 1 && A_Args[1] = "-Update") {
            UpdateChecker.Check(true)
            return
        }
        if (A_Args.Length >= 1 && A_Args[1] = "-Preferences") {
            args := PreferencesWindow.ParseArgs(A_Args)
            PreferencesWindow.Show(args.Page, args.X, args.Y)
            return
        }
        if AppSettings.MigratedFrom
            return                                                          ; 升级提示显示中, 不马上弹出窗口
        SearchWindow.Show()
        if (AppSettings.MovedFrom = "")                                     ; 不盖掉 "设置文件已移到 Data" 的提示
            App.Notify(I18n.T("App.Running", App._HotkeyText()), 3000)
    }

    ; OnExit 回调返回非零值会取消退出, 所以这里不返回任何值
    static _OnExit() {
        Knowledge.Save()
        Usage.Save()
        ClipboardProvider.Save()
        Logger.Flush()
    }
}
