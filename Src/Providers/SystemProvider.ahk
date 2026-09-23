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
        for command in SystemProvider.Commands() {
            score := FuzzyMatcher.Best(query.Text, [command["Title"], command["Id"], command["English"], command["Pinyin"]])
            if (score <= 0)
                continue
            results.Push(ResultItem(command["Title"], command["Subtitle"], {
                Icon: command["Icon"], Uid: "system:" command["Id"], Score: score,
                OnRun: SystemProvider._Runner(command["Id"])
            }))
        }
        return results
    }

    static _Runner(id) {
        return (*) => SystemProvider.RunCommand(id)
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
        add(id, titleKey, icon, fn, confirm := false, subtitleKey := "Sys.Subtitle") {
            title := I18n.T(titleKey)
            list.Push(Map("Id", id, "Title", title, "English", I18n.Strings[titleKey][1], "Pinyin", Pinyin.Initials(title),
                "Subtitle", I18n.T(subtitleKey), "Icon", icon, "Run", fn, "Confirm", confirm))
        }
        tool(id, titleKey, target, arguments := "", icon := "") {
            add(id, titleKey, (icon != "") ? icon : target, () => Run(Trim(target " " arguments)), false, "Tool.Subtitle")
        }
        system32 := A_WinDir "\System32\"
        text(id, titleKey, fn) => add(id, titleKey, "res:imageres.dll,-5314", () => SystemProvider.TransformClipboard(fn))

        ; --- ALTRun ---
        add("Preferences" , "Sys.Preferences" , "res:imageres.dll,-114" , () => App.OpenPreferences())
        add("Reload"      , "Sys.Reload"      , "res:imageres.dll,-5311", () => App.Reload())
        add("RebuildIndex", "Sys.RebuildIndex", "res:imageres.dll,-8"   , () => App.RebuildIndex())
        add("CheckUpdate" , "Sys.CheckUpdate" , "res:imageres.dll,-5338", () => UpdateChecker.Check(false))
        add("About"       , "Sys.About"       , "res:imageres.dll,-81"  , () => App.About())
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
        add("ShowIP"      , "Sys.ShowIP"      , "res:imageres.dll,-25"  , () => SystemProvider.ShowIP())
        add("TerminalHere", "Sys.TerminalHere", "res:imageres.dll,-5323", () => TerminalProvider.OpenAtCurrentFolder())
        add("ListProcesses", "Sys.ListProcesses", system32 "taskmgr.exe", () => SystemProvider._ShowCommandOutput("tasklist", "ALTRun.Processes.txt"))
        add("ListServices", "Sys.ListServices", system32 "services.msc", () => SystemProvider._ShowCommandOutput("net start", "ALTRun.Services.txt"))
        add("PTTools"     , "Sys.PTTools"     , "res:imageres.dll,-182" , () => PTToolsWindow.Show())
        add("SPF2M"       , "Sys.SPF2M"       , "res:imageres.dll,-182" , () => PTToolsWindow.ShowSpf2m())

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
        tool("Printers"          , "Tool.Printers"          , system32 "control.exe", "printers", "::{A8A91A66-3A7D-4424-8D24-04E180695C7A}")
        tool("Notepad"           , "Tool.Notepad"           , system32 "notepad.exe")
        tool("Calculator"        , "Tool.Calculator"        , system32 "calc.exe")
        tool("Paint"             , "Tool.Paint"             , system32 "mspaint.exe")
        tool("WinVer"            , "Tool.WinVer"            , system32 "winver.exe")

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

    ; 运行控制台命令, 结果写到临时文件后用记事本打开
    static _ShowCommandOutput(command, fileName) {
        outFile := A_Temp "\" fileName
        try FileDelete(outFile)
        RunWait(A_ComSpec ' /c ' command ' > "' outFile '"', A_Temp, "Hide")
        Run('notepad.exe "' outFile '"')
    }
}
