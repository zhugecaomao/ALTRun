;===============================================================================
; SystemProvider.ahk - 系统命令 / Windows 工具 / ALTRun 命令 / 剪贴板文字工具 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 每条命令有一个固定的 Id, ALTRun.json 的 Hotkeys 列表用 Id 引用命令, 例如:
;   { "Key": "~MButton", "Action": "PTTools", "WinTitle": "ahk_exe RAPTW.exe" }
; 可用的 Id 见 Commands() 里的列表; 旧版本 (2.x) 的函数名通过 Aliases 兼容。
; 另外两个只给热键用的 Id: "ToggleWindow" (显示/隐藏搜索窗口), "Show"。
;
; 关机 / 重启 / 注销 / 休眠 / 清空回收站 执行前会确认 (Features.System.ConfirmActions)。
;
; 用法:
;   SystemProvider.RunCommand("Lock")
;===============================================================================

class SystemProvider {
    static Id := "System"
    static _commands := ""

    static Aliases := Map(
        "Options", "Preferences", "Reindex", "RebuildIndex", "RestartApp", "Reload", "Exit", "Quit",
        "Update", "CheckUpdate", "TurnMonitorOff", "MonitorOff", "MuteVolume", "Mute",
        "IncreaseVolume", "VolumeUp", "DecreaseVolume", "VolumeDown", "ShutdownMachine", "Shutdown",
        "RestartMachine", "Restart", "HibernateMachine", "Hibernate", "OpenTerminalHere", "TerminalHere",
        "ListProcess", "ListProcesses", "ListService", "ListServices", "ClipUpper", "TextUpper",
        "ClipLower", "TextLower", "ClipTitleCase", "TextTitle", "ClipReverse", "TextReverse",
        "ClipSortAsc", "TextSortAsc", "ClipSortDesc", "TextSortDesc", "ClipTrimLines", "TextTrimLines",
        "ClipRemoveBlankLines", "TextRemoveBlank", "ClipDedupeLines", "TextDedupe",
        "ClipToTraditional", "TextToTraditional", "ClipToSimplified", "TextToSimplified", "UrlEncode", "TextUrlEncode",
        "Activate", "Show"
    )

    static Init() {
        SystemProvider.Commands()
    }

    static Search(query) {
        results := []
        needle := StrLower(query.Text)
        hidden := SystemProvider._HiddenIds()
        options := AppSettings.Feature("System")
        settingsPages := !options.Has("SettingsPages") || options["SettingsPages"]
        for command in SystemProvider.Commands() {
            if (hidden.Has(command["Id"]) || !settingsPages && command.Has("IsSetting"))
                continue
            score := FuzzyMatcher.BestKey(needle, command["Keys"])
            if (score <= 0)
                continue
            results.Push(ResultItem(command["Title"], command["Subtitle"], {
                Icon: command["Icon"], Uid: "system:" command["Id"], Score: score, Source: command,
                OnRun: SystemProvider._Runner(command["Id"])
            }))
            if (command["Id"] = "CheckUpdate" && IsObject(update := UpdateChecker.PendingItem())) {
                update.Score := score + 1                                   ; 有新版本时, 搜索 "更新" 排在 "检查更新" 前面
                results.Push(update)
            }
        }
        return results
    }

    ;---------------------------------------------------------------------------
    ; 隐藏用不到的命令 (搜索结果里 Ctrl+Del / 右键 "删除"); 热键和别的地方仍然可以用 Id 调用
    ;---------------------------------------------------------------------------
    static DeleteItem(item) {
        return SystemProvider.Hide(item.Source["Id"])
    }

    static DeletePrompt(item) {
        return I18n.T("Sys.ConfirmHide", item.Title)
    }

    static Hide(id) {
        if SystemProvider._HiddenIds().Has(id)
            return true
        AppSettings.Feature("System")["Hidden"].Push(id)
        return AppSettings.Save()
    }

    ; 隐藏列表 -> Map (不区分大小写), 列表变化 (换了数组或条数变了) 时重新生成
    static _HiddenIds() {
        static cache := "", cacheFor := ""
        hidden := AppSettings.Feature("System")["Hidden"]
        signature := ObjPtr(hidden) ":" hidden.Length
        if (cacheFor != signature) {
            cache := Map()
            cache.CaseSense := "Off"
            for id in hidden
                cache[Trim(id)] := true
            cacheFor := signature
        }
        return cache
    }

    ; 空搜索框里的结果: 后台发现了新版本时, 显示 "发现新版本: ALTRun x" (和 Alfred 一样, 不弹窗)
    static EmptyResults() {
        update := UpdateChecker.PendingItem()
        return IsObject(update) ? [update] : []
    }

    static _Runner(id) {
        return (*) => SystemProvider.RunCommand(id)
    }

    static _PageOpener(pageKey) => () => App.OpenPreferences(pageKey)

    ; 剪贴板里的内容去掉格式 (字体、颜色、表格...) 粘贴, 剪贴板随后还原。可以在 自定义热键 里设成 Ctrl+Shift+V:
    ; 用热键时贴到当前窗口; 在搜索窗口里运行时贴到呼出之前的窗口
    static PastePlainText() {
        text := A_Clipboard
        if (text = "")
            return App.Notify(I18n.T("Sys.PastePlainEmpty"))
        active := WinExist("A")
        ActionCatalog.PasteText(text, !active || (IsObject(SearchWindow.Gui) && active = SearchWindow.Gui.Hwnd))
    }

    static RunCommand(id) {
        if SystemProvider.Aliases.Has(id)
            id := SystemProvider.Aliases[id]
        if (id = "ToggleWindow")
            return SearchWindow.Toggle()
        if (id = "Show")
            return SearchWindow.Show()
        for command in SystemProvider.Commands() {
            if (command["Id"] != id)
                continue
            if (command["Confirm"] && AppSettings.Feature("System")["ConfirmActions"]) {
                if (MsgBox(I18n.T("Sys.ConfirmTitle", command["Title"]), App.Name, "YesNo Icon?") != "Yes")
                    return
            }
            return command["Run"].Call()
        }
        Logger.Error("SystemProvider: unknown command " id)
    }

    ; 命令表只建一次 (标题跟随当前语言)
    static Commands() {
        if IsObject(SystemProvider._commands)
            return SystemProvider._commands
        list := []
        add(id, titleKey, icon, fn, confirm := false, subtitleKey := "Sys.Subtitle", title := "", english := "") {
            title := (title != "") ? title : I18n.T(titleKey)
            english := (english != "") ? english : I18n.Strings[titleKey], pinyinText := Pinyin.Initials(title)
            keys := [FuzzyMatcher.Key(title), FuzzyMatcher.Key(id), (english != title) ? FuzzyMatcher.Key(english) : "", (pinyinText != title) ? FuzzyMatcher.Key(pinyinText) : ""]
            list.Push(Map("Id", id, "Title", title, "English", english, "Keys", keys,
                "Subtitle", I18n.T(subtitleKey), "Icon", icon, "Run", fn, "Confirm", confirm))
        }
        tool(id, titleKey, target, arguments := "", icon := "") {
            add(id, titleKey, (icon != "") ? icon : target, () => Run(Trim(target " " arguments)), false, "Tool.Subtitle")
        }
        system32 := A_WinDir "\System32\"
        text(id, titleKey, fn) => add(id, titleKey, ClipboardProvider.Icon, () => SystemProvider.TransformClipboard(fn))

        ; --- ALTRun ---
        add("Preferences" , "Sys.Preferences" , "res:imageres.dll,-114" , () => App.OpenPreferences())
        for pageKey in PreferencesWindow.PageKeys                          ; 每个设置页一条, 直接打开那一页 ("ALTRun 偏好设置: 外观");
            add("Preferences." SubStr(pageKey, 12), "", "res:imageres.dll,-114", SystemProvider._PageOpener(pageKey), false, "Sys.Subtitle"   ; 标题比主项长, 同一档里排在它后面
              , I18n.T("Sys.PreferencesPage", I18n.T(pageKey)), Format(I18n.Strings["Sys.PreferencesPage"], I18n.Strings[pageKey]))
        add("Reload"      , "Sys.Reload"      , "res:imageres.dll,-5311", () => App.Reload())
        add("RebuildIndex", "Sys.RebuildIndex", "res:imageres.dll,-8"   , () => App.RebuildIndex())
        appIcon := FileExist(App.IconFile) ? App.IconFile : "res:imageres.dll,-81"      ; ALTRun 自己的事用程序图标
        add("CheckUpdate" , "Sys.CheckUpdate" , appIcon                 , () => UpdateChecker.Check())
        add("About"       , "Sys.About"       , appIcon                 , () => App.About())
        add("Log"         , "Sys.Log"         , "res:imageres.dll,-102" , () => App.OpenLog())
        add("Quit"        , "Sys.Quit"        , "res:imageres.dll,-98"  , () => App.Quit())

        ; --- System ---
        add("Lock"        , "Sys.Lock"        , "res:imageres.dll,-59"  , () => DllCall("LockWorkStation"))
        add("Sleep"       , "Sys.Sleep"       , "res:shell32.dll,-28"   , () => DllCall("PowrProf\SetSuspendState", "Int", 0, "Int", 0, "Int", 0))
        add("Hibernate"   , "Sys.Hibernate"   , "res:shell32.dll,-28"   , () => DllCall("PowrProf\SetSuspendState", "Int", 1, "Int", 0, "Int", 0), true)
        add("Shutdown"    , "Sys.Shutdown"    , "res:shell32.dll,-28"   , () => Shutdown(1 | 8), true)
        add("Restart"     , "Sys.Restart"     , "res:shell32.dll,-28"   , () => Shutdown(2), true)
        add("Logoff"      , "Sys.Logoff"      , "res:shell32.dll,-45"   , () => Shutdown(0), true)
        add("EmptyRecycle", "Sys.EmptyRecycle", "res:shell32.dll,-32"   , () => FileRecycleEmpty(), true)
        add("MonitorOff"  , "Sys.MonitorOff"  , system32 "DisplaySwitch.exe", () => SendMessage(0x112, 0xF170, 2, , "Program Manager"))
        add("Mute"        , "Sys.Mute"        , system32 "SndVol.exe" , () => SoundSetMute(-1))
        add("VolumeUp"    , "Sys.VolumeUp"    , system32 "SndVol.exe" , () => SoundSetVolume("+10"))
        add("VolumeDown"  , "Sys.VolumeDown"  , system32 "SndVol.exe" , () => SoundSetVolume("-10"))
        add("MediaPlayPause", "Sys.MediaPlayPause", system32 "SndVol.exe", () => Send("{Media_Play_Pause}"))
        add("MediaNext"   , "Sys.MediaNext"   , system32 "SndVol.exe" , () => Send("{Media_Next}"))
        add("MediaPrev"   , "Sys.MediaPrev"   , system32 "SndVol.exe" , () => Send("{Media_Prev}"))
        add("MediaStop"   , "Sys.MediaStop"   , system32 "SndVol.exe" , () => Send("{Media_Stop}"))
        add("ShowIP"      , "Sys.ShowIP"      , "res:imageres.dll,-25"  , () => SystemProvider.ShowIP())
        add("PastePlain"  , "Sys.PastePlain"  , ClipboardProvider.Icon, () => SystemProvider.PastePlainText())
        add("ScriptsFolder", "Sys.ScriptsFolder", "res:imageres.dll,-5323", () => ScriptProvider.OpenFolder())
        add("TerminalHere", "Sys.TerminalHere", "res:imageres.dll,-5323", () => TerminalProvider.OpenAtCurrentFolder())
        add("ListProcesses", "Sys.ListProcesses", system32 "taskmgr.exe", () => SystemProvider._ShowCommandOutput("tasklist", "ALTRun.Processes.txt"))
        add("ListServices", "Sys.ListServices", system32 "services.msc", () => SystemProvider._ShowCommandOutput("net start", "ALTRun.Services.txt"))
        add("PTTools"     , "Sys.PTTools"     , IconCache.Own("Calculator", "res:imageres.dll,-182"), () => PTToolsWindow.Show())
        add("SPF2M"       , "Sys.SPF2M"       , IconCache.Own("Calculator", "res:imageres.dll,-182"), () => PTToolsWindow.ShowSpf2m())

        ; --- Clipboard text tools ---
        text("TextUpper"        , "Text.Upper"        , (s) => TextTools.Upper(s))
        text("TextLower"        , "Text.Lower"        , (s) => TextTools.Lower(s))
        text("TextTitle"        , "Text.Title"        , (s) => TextTools.TitleCase(s))
        text("TextSortAsc"      , "Text.SortAsc"      , (s) => TextTools.SortLines(s))
        text("TextSortDesc"     , "Text.SortDesc"     , (s) => TextTools.SortLines(s, true))
        text("TextTrimLines"    , "Text.TrimLines"    , (s) => TextTools.TrimLines(s))
        text("TextRemoveBlank"  , "Text.RemoveBlank"  , (s) => TextTools.RemoveBlankLines(s))
        text("TextDedupe"       , "Text.Dedupe"       , (s) => TextTools.DedupeLines(s))
        text("TextReverse"      , "Text.Reverse"      , (s) => TextTools.Reverse(s))
        text("TextToTraditional", "Text.ToTraditional", (s) => Kanji.ToTraditional(s))
        text("TextToSimplified" , "Text.ToSimplified" , (s) => Kanji.ToSimplified(s))
        text("TextUrlEncode"    , "Text.UrlEncode"    , (s) => Url.Encode(s))

        ; --- Windows tools ---
        tool("TaskManager"       , "Tool.TaskManager"       , system32 "taskmgr.exe")
        tool("ControlPanel"      , "Tool.ControlPanel"      , system32 "control.exe")
        tool("Settings"          , "Tool.Settings"          , "ms-settings:", , "shell:AppsFolder\windows.immersivecontrolpanel_cw5n1h2txyewy!microsoft.windows.immersivecontrolpanel")
        tool("DeviceManager"     , "Tool.DeviceManager"     , system32 "devmgmt.msc")
        tool("Services"          , "Tool.Services"          , system32 "services.msc")
        tool("Registry"          , "Tool.Registry"          , A_WinDir "\regedit.exe")
        tool("EventViewer"       , "Tool.EventViewer"       , system32 "eventvwr.msc")
        tool("DiskManagement"    , "Tool.DiskManagement"    , system32 "diskmgmt.msc")
        tool("ComputerManagement", "Tool.ComputerManagement", system32 "compmgmt.msc")
        tool("TaskScheduler"     , "Tool.TaskScheduler"     , system32 "taskschd.msc")
        tool("Programs"          , "Tool.Programs"          , system32 "control.exe", "appwiz.cpl", system32 "appwiz.cpl")
        tool("SystemProperties"  , "Tool.SystemProperties"  , system32 "control.exe", "sysdm.cpl", system32 "sysdm.cpl")
        tool("Network"           , "Tool.Network"           , system32 "control.exe", "ncpa.cpl", system32 "ncpa.cpl")
        tool("Firewall"          , "Tool.Firewall"          , system32 "control.exe", "firewall.cpl", system32 "firewall.cpl")
        tool("ResourceMonitor"   , "Tool.ResourceMonitor"   , system32 "perfmon.exe", "/res")
        tool("DiskCleanup"       , "Tool.DiskCleanup"       , system32 "cleanmgr.exe")
        tool("SystemConfig"      , "Tool.SystemConfig"      , system32 "msconfig.exe")
        tool("GroupPolicy"       , "Tool.GroupPolicy"       , system32 "gpedit.msc")
        tool("CommandPrompt"     , "Tool.CommandPrompt"     , system32 "cmd.exe")
        tool("PowerShell"        , "Tool.PowerShell"        , system32 "WindowsPowerShell\v1.0\powershell.exe")
        tool("Explorer"          , "Tool.Explorer"          , A_WinDir "\explorer.exe")
        tool("RecycleBin"        , "Tool.RecycleBin"        , "::{645FF040-5081-101B-9F08-00AA002F954E}")
        tool("ThisPC"            , "Tool.ThisPC"            , "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}")
        tool("Printers"          , "Tool.Printers"          , system32 "control.exe", "printers", "::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{A8A91A66-3A7D-4424-8D24-04E180695C7A}")
        tool("Notepad"           , "Tool.Notepad"           , system32 "notepad.exe")
        tool("Calculator"        , "Tool.Calculator"        , system32 "calc.exe")
        tool("Paint"             , "Tool.Paint"             , system32 "mspaint.exe")
        tool("WinVer"            , "Tool.WinVer"            , system32 "winver.exe")

        ; --- Windows 设置的页面 (ms-settings:, Windows 10 / 11 都有), 可以在偏好设置里关掉 (Features.System.SettingsPages) ---
        setting(id, page) {
            add("Set" id, "Setting." id, "shell:AppsFolder\windows.immersivecontrolpanel_cw5n1h2txyewy!microsoft.windows.immersivecontrolpanel"
                , () => Run("ms-settings:" page), false, "Setting.Subtitle")
            list[list.Length]["IsSetting"] := true
        }
        setting("Display"          , "display")
        setting("NightLight"       , "nightlight")
        setting("Sound"            , "sound")
        setting("Notifications"    , "notifications")
        setting("Focus"            , "quiethours")
        setting("Power"            , "powersleep")
        setting("Battery"          , "batterysaver")
        setting("Storage"          , "storagesense")
        setting("Multitasking"     , "multitasking")
        setting("Clipboard"        , "clipboard")
        setting("About"            , "about")
        setting("Bluetooth"        , "bluetooth")
        setting("Mouse"            , "mousetouchpad")
        setting("Touchpad"         , "devices-touchpad")
        setting("Typing"           , "typing")
        setting("Network"          , "network-status")
        setting("Wifi"             , "network-wifi")
        setting("Vpn"              , "network-vpn")
        setting("Proxy"            , "network-proxy")
        setting("Airplane"         , "network-airplanemode")
        setting("Background"       , "personalization-background")
        setting("Colors"           , "colors")
        setting("LockScreen"       , "lockscreen")
        setting("Themes"           , "themes")
        setting("Taskbar"          , "taskbar")
        setting("Start"            , "personalization-start")
        setting("Fonts"            , "fonts")
        setting("Apps"             , "appsfeatures")
        setting("DefaultApps"      , "defaultapps")
        setting("StartupApps"      , "startupapps")
        setting("OptionalFeatures" , "optionalfeatures")
        setting("Account"          , "yourinfo")
        setting("SignIn"           , "signinoptions")
        setting("DateTime"         , "dateandtime")
        setting("Region"           , "regionformatting")
        setting("Language"         , "regionlanguage")
        setting("Update"           , "windowsupdate")
        setting("Security"         , "windowsdefender")
        setting("Privacy"          , "privacy")
        setting("Recovery"         , "recovery")
        setting("Activation"       , "activation")
        setting("Developers"       , "developers")
        setting("Backup"           , "backup")

        SystemProvider._commands := list
        return list
    }

    ;---------------------------------------------------------------------------
    ; Command implementations
    ;---------------------------------------------------------------------------
    static TransformClipboard(fn) {
        text := A_Clipboard
        if (text = "")
            return App.Notify(I18n.T("Text.Empty"))
        A_Clipboard := fn(text)
        App.Notify(I18n.T("Text.Done"))
    }

    static ShowIP() {
        addresses := ""
        for address in SysGetIPAddresses()
            addresses .= (addresses = "" ? "" : "`n") address
        if (addresses = "")
            return MsgBox(I18n.T("Sys.NoIP"), App.Name, 48)
        A_Clipboard := StrSplit(addresses, "`n")[1]
        MsgBox(addresses, I18n.T("Sys.IPCopied"), 64)
    }

    ; 运行控制台命令, 结果写到临时文件后用 .txt 的默认程序 (一般是记事本) 打开
    static _ShowCommandOutput(command, fileName) {
        outFile := A_Temp "\" fileName
        try FileDelete(outFile)
        RunWait(A_ComSpec ' /c ' command ' > "' outFile '"', A_Temp, "Hide")
        Path.OpenText(outFile)
    }
}
