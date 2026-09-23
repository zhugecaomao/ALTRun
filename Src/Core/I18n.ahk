;===============================================================================
; I18n.ahk - 界面文字 (英文 / 中文) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 每条文字用一个语义化的键名, 值是 [英文, 中文]。{1} {2} ... 是参数占位符。
;
; 用法:
;   I18n.Init("auto")                          启动时调用, "auto" 按系统语言选择
;   I18n.T("Search.Placeholder")               取当前语言的文字
;   I18n.T("Web.SearchFor", "Google", "abc")   带参数
;===============================================================================

class I18n {
    static Lang := "en"

    static Init(setting := "auto") {
        if (setting = "en" || setting = "zh")
            I18n.Lang := setting
        else
            I18n.Lang := I18n.IsChineseSystem() ? "zh" : "en"
    }

    static IsChineseSystem() {
        ; 0804 简体(中国大陆) 0404 繁体(台湾) 0C04 香港 1004 新加坡 1404 澳门
        for code in ["0804", "0404", "0C04", "1004", "1404"]
            if (A_Language = code)
                return true
        return false
    }

    static T(key, args*) {
        if !I18n.Strings.Has(key)
            return key
        pair := I18n.Strings[key]
        text := (I18n.Lang = "zh") ? pair[2] : pair[1]
        for index, arg in args
            text := StrReplace(text, "{" index "}", arg)
        return text
    }

    static Strings := Map(
        ; --- App / tray ---
        "App.Tagline"            , ["An effective launcher for Windows", "高效的 Windows 启动器"],
        "App.Running"            , ["ALTRun is running. Press {1} to search.", "ALTRun 已在运行, 按 {1} 开始搜索。"],
        "Tray.Show"              , ["Show ALTRun", "显示 ALTRun"],
        "Tray.Preferences"       , ["Preferences...", "偏好设置..."],
        "Tray.RebuildIndex"      , ["Rebuild Index", "重建索引"],
        "Tray.CheckUpdate"       , ["Check for Updates", "检查更新"],
        "Tray.Reload"            , ["Reload", "重新载入"],
        "Tray.Exit"              , ["Quit", "退出"],

        ; --- Settings / migration ---
        "Settings.ParseError"    , ["ALTRun.json could not be read:`n`n{1}`n`nIt was renamed to {2} and the default settings are used.", "ALTRun.json 无法解析:`n`n{1}`n`n已改名为 {2}, 现在使用默认设置。"],
        "Settings.SaveError"     , ["Could not save ALTRun.json:`n`n{1}", "无法保存 ALTRun.json:`n`n{1}"],
        "Settings.Migrated"      , ["Settings were upgraded from version {1}. A backup was saved as {2}.", "设置已从版本 {1} 升级, 原文件备份为 {2}。"],
        "Settings.EditHint"      , ["Preferences are edited in ALTRun.json for now. ALTRun reloads when you save the file.", "目前在 ALTRun.json 中修改偏好设置, 保存文件后 ALTRun 会自动重新载入。"],

        ; --- Search window ---
        "Search.Placeholder"     , ["ALTRun Search", "ALTRun 搜索"],
        "Search.ActionsFor"      , ["Actions for {1}", "{1} 的操作"],
        "Search.Copied"          , ["Copied: {1}", "已复制: {1}"],

        ; --- Actions ---
        "Action.Open"            , ["Open", "打开"],
        "Action.Run"             , ["Run", "运行"],
        "Action.RunAsAdmin"      , ["Run as Administrator", "以管理员身份运行"],
        "Action.Reveal"          , ["Reveal in File Manager", "在文件管理器中显示"],
        "Action.CopyPath"        , ["Copy Path", "复制路径"],
        "Action.CopyName"        , ["Copy Name", "复制名称"],
        "Action.CopyUrl"         , ["Copy URL", "复制网址"],
        "Action.Copy"            , ["Copy to Clipboard", "复制到剪贴板"],
        "Action.Paste"           , ["Paste to Front Window", "粘贴到前台窗口"],
        "Action.LargeType"       , ["Show in Large Type", "大字显示"],
        "Action.OpenTerminal"    , ["Open Terminal Here", "在此处打开终端"],
        "Action.Properties"      , ["Properties", "属性"],
        "Action.AddCommand"      , ["Add to Custom Commands", "添加到自定义命令"],

        ; --- Providers ---
        "App.Subtitle.Store"     , ["Microsoft Store app", "应用商店应用"],
        "Calc.Subtitle"          , ["Copy result to clipboard", "复制结果到剪贴板"],
        "Calc.BeamWidth"         , ["Beam width {1} mm: {2} main bars @ {3} c/c", "梁宽 {1} mm: 主筋 {2} 根 @ {3} c/c"],
        "Calc.RebarArea"         , ["As = {1} mm²: {2}", "As = {1} mm²: {2}"],
        "Web.SearchFor"          , ["Search {1} for '{2}'", "用 {1} 搜索 '{2}'"],
        "Web.SearchEmpty"        , ["Search {1} for '...'", "用 {1} 搜索 '...'"],
        "Files.SearchFor"        , ["Search files for '{1}'", "搜索文件 '{1}'"],
        "Files.OpenEverything"   , ["Search '{1}' in Everything", "在 Everything 中搜索 '{1}'"],
        "Files.WindowsSearch"    , ["Search '{1}' with Windows Search", "用 Windows 搜索 '{1}'"],
        "Files.Keyword"          , ["Find files by name", "按文件名查找文件"],
        "Terminal.Run"           , ["Run '{1}' in terminal", "在终端运行 '{1}'"],
        "Terminal.Empty"         , ["Run a command in terminal", "在终端运行命令"],
        "Snippet.Subtitle"       , ["Paste snippet · {1}", "粘贴片段 · {1}"],
        "Clipboard.Subtitle"     , ["{1} · {2} · {3} characters", "{1} · {2} · {3} 个字符"],
        "Clipboard.Empty"        , ["Clipboard history is empty", "剪贴板历史为空"],
        "Clipboard.EmptyHint"    , ["Copied text will appear here ({1})", "复制过的文字会出现在这里 ({1})"],
        "Clipboard.Clear"        , ["Clear Clipboard History", "清空剪贴板历史"],
        "Clipboard.ClearHint"    , ["{1} items", "{1} 条"],
        "Clipboard.Cleared"      , ["Clipboard history cleared", "剪贴板历史已清空"],
        "Clipboard.Delete"       , ["Delete from History", "从历史中删除"],
        "Custom.Added"           , ["Added '{1}' to Custom Commands.", "已将 '{1}' 添加到自定义命令。"],
        "Index.Done"             , ["Index rebuilt: {1} applications.", "索引已重建: {1} 个应用。"],

        ; --- System commands ---
        "Sys.Preferences"        , ["ALTRun Preferences", "ALTRun 偏好设置"],
        "Sys.Reload"             , ["Reload ALTRun", "重新载入 ALTRun"],
        "Sys.RebuildIndex"       , ["Rebuild ALTRun Index", "重建 ALTRun 索引"],
        "Sys.Quit"               , ["Quit ALTRun", "退出 ALTRun"],
        "Sys.CheckUpdate"        , ["Check for ALTRun Updates", "检查 ALTRun 更新"],
        "Sys.About"              , ["About ALTRun", "关于 ALTRun"],
        "Sys.Log"                , ["Open ALTRun Log", "打开 ALTRun 日志"],
        "Sys.Lock"               , ["Lock Screen", "锁定屏幕"],
        "Sys.Sleep"              , ["Sleep", "睡眠"],
        "Sys.Hibernate"          , ["Hibernate", "休眠"],
        "Sys.Shutdown"           , ["Shut Down", "关机"],
        "Sys.Restart"            , ["Restart", "重启"],
        "Sys.Logoff"             , ["Log Off", "注销"],
        "Sys.EmptyRecycle"       , ["Empty Recycle Bin", "清空回收站"],
        "Sys.MonitorOff"         , ["Turn Off Monitor", "关闭显示器"],
        "Sys.Mute"               , ["Toggle Mute", "静音 / 取消静音"],
        "Sys.VolumeUp"           , ["Volume Up", "增大音量"],
        "Sys.VolumeDown"         , ["Volume Down", "减小音量"],
        "Sys.ShowIP"             , ["Show IP Address", "显示 IP 地址"],
        "Sys.TerminalHere"       , ["Open Terminal at Current Folder", "在当前文件夹打开终端"],
        "Sys.ListProcesses"      , ["List Running Processes", "列出运行中的进程"],
        "Sys.ListServices"       , ["List Running Services", "列出运行中的服务"],
        "Sys.PTTools"            , ["PT Tools (Rebar / BRC Calculator)", "PT 工具箱 (钢筋 / BRC 计算)"],
        "Sys.SPF2M"              , ["SPF2M Profile Calculator", "SPF2M 束线型计算"],
        "Sys.Subtitle"           , ["System command", "系统命令"],
        "Sys.ConfirmTitle"       , ["{1}?", "确定要{1}吗?"],
        "Sys.NoIP"               , ["No IP address found.", "没有找到 IP 地址。"],
        "Sys.IPCopied"           , ["IP address (the first one is copied)", "IP 地址 (第一个已复制)"],
        "Sys.NoFolder"           , ["No Explorer or Total Commander folder found.", "没有找到资源管理器或 Total Commander 的当前文件夹。"],
        "Text.Upper"             , ["Clipboard: UPPERCASE", "剪贴板: 转大写"],
        "Text.Lower"             , ["Clipboard: lowercase", "剪贴板: 转小写"],
        "Text.Title"             , ["Clipboard: Title Case", "剪贴板: 首字母大写"],
        "Text.SortAsc"           , ["Clipboard: Sort Lines A-Z", "剪贴板: 行排序 A-Z"],
        "Text.SortDesc"          , ["Clipboard: Sort Lines Z-A", "剪贴板: 行排序 Z-A"],
        "Text.TrimLines"         , ["Clipboard: Trim Each Line", "剪贴板: 去除行首尾空白"],
        "Text.RemoveBlank"       , ["Clipboard: Remove Blank Lines", "剪贴板: 删除空行"],
        "Text.Dedupe"            , ["Clipboard: Remove Duplicate Lines", "剪贴板: 删除重复行"],
        "Text.Reverse"           , ["Clipboard: Reverse Text", "剪贴板: 反转文字"],
        "Text.ToTraditional"     , ["Clipboard: Simplified → Traditional", "剪贴板: 简体转繁体"],
        "Text.ToSimplified"      , ["Clipboard: Traditional → Simplified", "剪贴板: 繁体转简体"],
        "Text.UrlEncode"         , ["Clipboard: URL Encode", "剪贴板: URL 编码"],
        "Text.Done"              , ["Clipboard converted", "剪贴板已转换"],
        "Text.Empty"             , ["Clipboard is empty", "剪贴板为空"],

        ; --- Windows tools ---
        "Tool.TaskManager"       , ["Task Manager", "任务管理器"],
        "Tool.ControlPanel"      , ["Control Panel", "控制面板"],
        "Tool.Settings"          , ["Windows Settings", "Windows 设置"],
        "Tool.DeviceManager"     , ["Device Manager", "设备管理器"],
        "Tool.Services"          , ["Services", "服务"],
        "Tool.Registry"          , ["Registry Editor", "注册表编辑器"],
        "Tool.EventViewer"       , ["Event Viewer", "事件查看器"],
        "Tool.DiskManagement"    , ["Disk Management", "磁盘管理"],
        "Tool.ComputerManagement", ["Computer Management", "计算机管理"],
        "Tool.TaskScheduler"     , ["Task Scheduler", "任务计划程序"],
        "Tool.Programs"          , ["Programs and Features", "程序和功能"],
        "Tool.SystemProperties"  , ["System Properties", "系统属性"],
        "Tool.Network"           , ["Network Connections", "网络连接"],
        "Tool.Firewall"          , ["Windows Defender Firewall", "Windows Defender 防火墙"],
        "Tool.ResourceMonitor"   , ["Resource Monitor", "资源监视器"],
        "Tool.DiskCleanup"       , ["Disk Cleanup", "磁盘清理"],
        "Tool.SystemConfig"      , ["System Configuration", "系统配置"],
        "Tool.GroupPolicy"       , ["Group Policy Editor", "组策略编辑器"],
        "Tool.CommandPrompt"     , ["Command Prompt", "命令提示符"],
        "Tool.PowerShell"        , ["PowerShell", "PowerShell"],
        "Tool.Explorer"          , ["File Explorer", "文件资源管理器"],
        "Tool.RecycleBin"        , ["Recycle Bin", "回收站"],
        "Tool.ThisPC"            , ["This PC", "此电脑"],
        "Tool.Printers"          , ["Devices and Printers", "设备和打印机"],
        "Tool.Notepad"           , ["Notepad", "记事本"],
        "Tool.Calculator"        , ["Calculator", "计算器"],
        "Tool.Paint"             , ["Paint", "画图"],
        "Tool.WinVer"            , ["About Windows", "关于 Windows"],
        "Tool.Subtitle"          , ["Windows tool", "Windows 工具"],

        ; --- Update checker ---
        "Update.Available"       , ["A new version {1} is available. Open the download page?", "发现新版本 {1}, 是否打开下载页面?"],
        "Update.Latest"          , ["You are running the latest version ({1}).", "当前已是最新版本 ({1})。"],
        "Update.Failed"          , ["Could not check for updates:`n`n{1}", "检查更新失败:`n`n{1}"],

        ; --- QuickSwitch (file dialog path sync) ---
        "QuickSwitch.Hint"       , ["{1}: Total Commander folder  {2}: Explorer folder", "{1}: 跳到 TC 目录  {2}: 跳到资源管理器目录"],
        "QuickSwitch.NoTC"       , ["Total Commander is not running.", "Total Commander 没有运行。"],
        "QuickSwitch.NoExplorer" , ["No File Explorer window is open.", "没有打开的资源管理器窗口。"]
    )
}
