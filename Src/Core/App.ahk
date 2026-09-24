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
;   -SendTo <path>  资源管理器 "发送到" 菜单: 把文件/文件夹添加为自定义命令
;
; 用法 (其它模块里):
;   App.Notify("...")              屏幕上方短暂提示
;   App.FocusPreviousWindow()      回到呼出 ALTRun 之前的窗口 (粘贴用)
;   App.OpenPreferences() / App.EditSettingsFile() / App.Reload() / App.Restart(args)
;   App.Quit() / App.RebuildIndex()
;===============================================================================

class App {
    static Name    := "ALTRun"
    static Version := "2026.09.24"
    static RepoUrl := "https://github.com/zhugecaomao/ALTRun"
    static PreviousWindow := 0
    static _settingsTime := ""
    static _watchTimer := ""

    static Start() {
        Logger.Rotate()
        App._MoveLegacyResources()
        AppSettings.Load()
        Logger.Enabled := AppSettings.General["SaveLog"] ? true : false
        Logger.Debug("===== " App.Name " " App.Version " starting =====")
        I18n.Init(AppSettings.General["Language"])
        Knowledge.Load()
        ThemeManager.Load(AppSettings.Appearance["Theme"])

        for provider in [ClipboardProvider, ApplicationProvider, CustomCommandProvider, SnippetProvider, SystemProvider
                        , CalculatorProvider, WebSearchProvider, FileSearchProvider, TerminalProvider, HelpProvider]
            ProviderRegistry.Register(provider)
        ProviderRegistry.InitAll()

        SearchWindow.Create()
        App._CreateTrayMenu()
        App._RegisterHotkeys()
        SnippetExpander.Init()
        QuickSwitch.Init(AppSettings.Extension("QuickSwitch"))
        AutoDate.Init(AppSettings.Extension("AutoDate"))
        PTToolsWindow.Load(AppSettings.Extension("PTTools"))
        App._UpdateShellShortcuts()
        OnExit((*) => App._OnExit())

        if (AppSettings.ImportedFrom != "")
            App.Notify(I18n.T("Settings.ImportedIni", AppSettings.ImportedFrom), 6000)
        else if AppSettings.MigratedFrom
            App.Notify(I18n.T("Settings.Migrated", AppSettings.MigratedFrom, SchemaMigration.BackupFile(AppSettings.File, AppSettings.MigratedFrom)), 5000)
        if AppSettings.General["CheckForUpdates"]
            SetTimer(() => UpdateChecker.Check(true), -10000)

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

    ; 屏幕上方居中显示一条提示, duration 毫秒后消失
    static Notify(text, duration := 1500) {
        area := Win.WorkAreaAtMouse()
        CoordMode("ToolTip", "Screen")
        ToolTip(text, area.Left + (area.Right - area.Left) // 2 - 150, area.Top + Round((area.Bottom - area.Top) * 0.12), 20)
        SetTimer(() => ToolTip(, , , 20), -duration)
    }

    ;---------------------------------------------------------------------------
    ; Commands (tray menu / system commands / hotkeys)
    ;---------------------------------------------------------------------------
    static OpenPreferences(pageIndex := 1) {
        PreferencesWindow.Show(pageIndex)
    }

    ; 直接用记事本编辑 ALTRun.json, 保存后自动重新载入
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
        App.Notify(I18n.T("Index.Done", count))
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

    static About() {
        MsgBox(App.Name " " App.Version "`n" I18n.T("App.Tagline") "`n`n" App.RepoUrl, App.Name, 64)
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
        try TraySetIcon("imageres.dll", -100)
        tray := A_TrayMenu
        tray.Delete()
        tray.Add(I18n.T("Tray.Show"), (*) => SetTimer(() => SearchWindow.Show(), -100))
        tray.Add(I18n.T("Tray.Preferences"), (*) => App.OpenPreferences())
        tray.Add()
        tray.Add(I18n.T("Tray.RebuildIndex"), (*) => App.RebuildIndex())
        tray.Add(I18n.T("Tray.CheckUpdate"), (*) => UpdateChecker.Check(false))
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
            if (key = "")
                continue
            try {
                Hotkey(key, (*) => SearchWindow.Toggle())
            } catch as e {
                Logger.Error("App: cannot register hotkey " key " - " e.Message)
                MsgBox("Cannot register hotkey " key ":`n" e.Message, App.Name, 48)
            }
        }

        clipboard := AppSettings.Feature("Clipboard")
        if (clipboard["Enabled"] && clipboard["Hotkey"] != "") {
            try {
                Hotkey(clipboard["Hotkey"], (*) => SearchWindow.Show(AppSettings.Feature("Clipboard")["Keyword"] " "))
            } catch as e {
                Logger.Error("App: cannot register clipboard hotkey - " e.Message)
            }
        }

        ; 自定义热键: Key -> 系统命令 Action, WinTitle 非空时只在该窗口里生效
        for entry in AppSettings.Hotkeys {
            if !(entry is Map) || !entry.Has("Key") || !entry.Has("Action")
                continue
            winTitle := entry.Has("WinTitle") ? entry["WinTitle"] : ""
            try {
                if (winTitle != "")
                    HotIfWinActive(winTitle)
                Hotkey(entry["Key"], App._HotkeyAction(entry["Action"]))
            } catch as e {
                Logger.Error("App: cannot register hotkey " entry["Key"] " - " e.Message)
            }
            HotIfWinActive()
        }
    }

    ; 单独一个方法生成闭包, 每个热键各自记住自己的 Action
    static _HotkeyAction(actionId) {
        return (*) => SystemProvider.RunCommand(actionId)
    }

    ; 开机启动 / 资源管理器 "发送到" / 开始菜单 三个快捷方式, 按设置创建或删除
    static _UpdateShellShortcuts() {
        general := AppSettings.General
        target := A_IsCompiled ? A_ScriptFullPath : A_AhkPath
        prefix := A_IsCompiled ? "" : '"' A_ScriptFullPath '" '
        sendTo := RegExReplace(A_StartMenu, "\\Start Menu$", "\SendTo") "\ALTRun.lnk"
        for shortcut in [
            [general["LaunchAtLogin"], A_Startup "\ALTRun.lnk", "-Startup"],
            [general["SendToMenu"], sendTo, "-SendTo"],
            [general["StartMenuShortcut"], A_Programs "\ALTRun.lnk", ""]
        ] {
            try {
                if shortcut[1]
                    FileCreateShortcut(target, shortcut[2], A_ScriptDir, Trim(prefix shortcut[3]), App.Name " - " I18n.T("App.Tagline"))
                else if FileExist(shortcut[2])
                    FileDelete(shortcut[2])
            } catch as e {
                Logger.Error("App: shortcut " shortcut[2] " - " e.Message)
            }
        }
    }

    static _HandleCommandLine() {
        if (A_Args.Length >= 2 && A_Args[1] = "-SendTo") {
            target := A_Args[2]
            if (SubStr(target, -4) = ".lnk") {
                try {
                    FileGetShortcut(target, &linkTarget)
                    if (linkTarget != "")
                        target := linkTarget
                }
            }
            CustomCommandProvider.AddFromPath(target)
            return
        }
        if (A_Args.Length >= 1 && (A_Args[1] = "-Startup" || A_Args[1] = "-Reloaded"))
            return
        if (A_Args.Length >= 1 && A_Args[1] = "-Preferences") {
            args := PreferencesWindow.ParseArgs(A_Args)
            PreferencesWindow.Show(args.Page, args.X, args.Y)
            return
        }
        if AppSettings.MigratedFrom
            return                                                          ; 升级提示显示中, 不马上弹出窗口
        SearchWindow.Show()
        App.Notify(I18n.T("App.Running", App._HotkeyText()), 3000)
    }

    ; OnExit 回调返回非零值会取消退出, 所以这里不返回任何值
    static _OnExit() {
        Knowledge.Save()
        ClipboardProvider.Save()
        Logger.Flush()
    }
}
