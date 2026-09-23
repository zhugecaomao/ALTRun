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

    static Strings := I18n._Build()

    ; 每条: 键名 -> [英文, 中文]
    static _Build() {
        s := Map()
        ; --- App / tray ---
        s["App.Tagline"]               := ["An effective launcher for Windows", "高效的 Windows 启动器"]
        s["App.Running"]               := ["ALTRun is running. Press {1} to search.", "ALTRun 已在运行, 按 {1} 开始搜索。"]
        s["Tray.Show"]                 := ["Show ALTRun", "显示 ALTRun"]
        s["Tray.Preferences"]          := ["Preferences...", "偏好设置..."]
        s["Tray.RebuildIndex"]         := ["Rebuild Index", "重建索引"]
        s["Tray.CheckUpdate"]          := ["Check for Updates", "检查更新"]
        s["Tray.Reload"]               := ["Reload", "重新载入"]
        s["Tray.Exit"]                 := ["Quit", "退出"]

        ; --- Settings / migration ---
        s["Settings.ParseError"]       := ["ALTRun.json could not be read:`n`n{1}`n`nIt was renamed to {2} and the default settings are used.", "ALTRun.json 无法解析:`n`n{1}`n`n已改名为 {2}, 现在使用默认设置。"]
        s["Settings.SaveError"]        := ["Could not save ALTRun.json:`n`n{1}", "无法保存 ALTRun.json:`n`n{1}"]
        s["Settings.Migrated"]         := ["Settings were upgraded from version {1}. A backup was saved as {2}.", "设置已从版本 {1} 升级, 原文件备份为 {2}。"]
        s["Settings.EditHint"]         := ["ALTRun reloads when you save ALTRun.json.", "保存 ALTRun.json 后 ALTRun 会自动重新载入。"]

        ; --- Search window ---
        s["Search.Placeholder"]        := ["ALTRun Search", "ALTRun 搜索"]
        s["Search.ActionsFor"]         := ["Actions for {1}", "{1} 的操作"]
        s["Search.Copied"]             := ["Copied: {1}", "已复制: {1}"]
        s["Search.NotEditable"]        := ["This result cannot be edited.", "这一项不能编辑。"]
        s["Search.ConfirmDelete"]      := ["Delete '{1}'?", "确定删除 '{1}' 吗?"]

        ; --- Actions ---
        s["Action.Open"]               := ["Open", "打开"]
        s["Action.Run"]                := ["Run", "运行"]
        s["Action.RunAsAdmin"]         := ["Run as Administrator", "以管理员身份运行"]
        s["Action.Reveal"]             := ["Reveal in File Manager", "在文件管理器中显示"]
        s["Action.CopyPath"]           := ["Copy Path", "复制路径"]
        s["Action.CopyName"]           := ["Copy Name", "复制名称"]
        s["Action.CopyUrl"]            := ["Copy URL", "复制网址"]
        s["Action.Copy"]               := ["Copy to Clipboard", "复制到剪贴板"]
        s["Action.Paste"]              := ["Paste to Front Window", "粘贴到前台窗口"]
        s["Action.LargeType"]          := ["Show in Large Type", "大字显示"]
        s["Action.OpenTerminal"]       := ["Open Terminal Here", "在此处打开终端"]
        s["Action.Properties"]         := ["Properties", "属性"]
        s["Action.AddCommand"]         := ["Add to Custom Commands...", "添加到自定义命令..."]
        s["Action.Edit"]               := ["Edit...", "编辑..."]
        s["Action.Delete"]             := ["Delete", "删除"]

        ; --- Providers ---
        s["App.Subtitle.Store"]        := ["Microsoft Store app", "应用商店应用"]
        s["Calc.Subtitle"]             := ["Copy result to clipboard", "复制结果到剪贴板"]
        s["Calc.BeamWidth"]            := ["Beam width {1} mm: {2} main bars @ {3} c/c", "梁宽 {1} mm: 主筋 {2} 根 @ {3} c/c"]
        s["Calc.RebarArea"]            := ["As = {1} mm²: {2}", "As = {1} mm²: {2}"]
        s["Web.SearchFor"]             := ["Search {1} for '{2}'", "用 {1} 搜索 '{2}'"]
        s["Web.SearchEmpty"]           := ["Search {1} for '...'", "用 {1} 搜索 '...'"]
        s["Files.SearchFor"]           := ["Search files for '{1}'", "搜索文件 '{1}'"]
        s["Files.OpenEverything"]      := ["Search '{1}' in Everything", "在 Everything 中搜索 '{1}'"]
        s["Files.WindowsSearch"]       := ["Search '{1}' with Windows Search", "用 Windows 搜索 '{1}'"]
        s["Files.Keyword"]             := ["Find files by name", "按文件名查找文件"]
        s["Terminal.Run"]              := ["Run '{1}' in terminal", "在终端运行 '{1}'"]
        s["Terminal.Empty"]            := ["Run a command in terminal", "在终端运行命令"]
        s["Snippet.Subtitle"]          := ["Paste snippet · {1}", "粘贴片段 · {1}"]
        s["Clipboard.Subtitle"]        := ["{1} · {2} · {3} characters", "{1} · {2} · {3} 个字符"]
        s["Clipboard.Empty"]           := ["Clipboard history is empty", "剪贴板历史为空"]
        s["Clipboard.EmptyHint"]       := ["Copied text will appear here ({1})", "复制过的文字会出现在这里 ({1})"]
        s["Clipboard.Clear"]           := ["Clear Clipboard History", "清空剪贴板历史"]
        s["Clipboard.ClearHint"]       := ["{1} items", "{1} 条"]
        s["Clipboard.Cleared"]         := ["Clipboard history cleared", "剪贴板历史已清空"]
        s["Custom.Added"]              := ["Added '{1}' to Custom Commands.", "已将 '{1}' 添加到自定义命令。"]
        s["Index.Done"]                := ["Index rebuilt: {1} applications.", "索引已重建: {1} 个应用。"]

        ; --- System commands ---
        s["Sys.Preferences"]           := ["ALTRun Preferences", "ALTRun 偏好设置"]
        s["Sys.Reload"]                := ["Reload ALTRun", "重新载入 ALTRun"]
        s["Sys.RebuildIndex"]          := ["Rebuild ALTRun Index", "重建 ALTRun 索引"]
        s["Sys.Quit"]                  := ["Quit ALTRun", "退出 ALTRun"]
        s["Sys.CheckUpdate"]           := ["Check for ALTRun Updates", "检查 ALTRun 更新"]
        s["Sys.About"]                 := ["About ALTRun", "关于 ALTRun"]
        s["Sys.Log"]                   := ["Open ALTRun Log", "打开 ALTRun 日志"]
        s["Sys.Lock"]                  := ["Lock Screen", "锁定屏幕"]
        s["Sys.Sleep"]                 := ["Sleep", "睡眠"]
        s["Sys.Hibernate"]             := ["Hibernate", "休眠"]
        s["Sys.Shutdown"]              := ["Shut Down", "关机"]
        s["Sys.Restart"]               := ["Restart", "重启"]
        s["Sys.Logoff"]                := ["Log Off", "注销"]
        s["Sys.EmptyRecycle"]          := ["Empty Recycle Bin", "清空回收站"]
        s["Sys.MonitorOff"]            := ["Turn Off Monitor", "关闭显示器"]
        s["Sys.Mute"]                  := ["Toggle Mute", "静音 / 取消静音"]
        s["Sys.VolumeUp"]              := ["Volume Up", "增大音量"]
        s["Sys.VolumeDown"]            := ["Volume Down", "减小音量"]
        s["Sys.ShowIP"]                := ["Show IP Address", "显示 IP 地址"]
        s["Sys.TerminalHere"]          := ["Open Terminal at Current Folder", "在当前文件夹打开终端"]
        s["Sys.ListProcesses"]         := ["List Running Processes", "列出运行中的进程"]
        s["Sys.ListServices"]          := ["List Running Services", "列出运行中的服务"]
        s["Sys.PTTools"]               := ["PT Tools (Rebar / BRC Calculator)", "PT 工具箱 (钢筋 / BRC 计算)"]
        s["Sys.SPF2M"]                 := ["SPF2M Profile Calculator", "SPF2M 束线型计算"]
        s["Sys.Subtitle"]              := ["System command", "系统命令"]
        s["Sys.ConfirmTitle"]          := ["{1}?", "确定要{1}吗?"]
        s["Sys.NoIP"]                  := ["No IP address found.", "没有找到 IP 地址。"]
        s["Sys.IPCopied"]              := ["IP address (the first one is copied)", "IP 地址 (第一个已复制)"]
        s["Sys.NoFolder"]              := ["No Explorer or Total Commander folder found.", "没有找到资源管理器或 Total Commander 的当前文件夹。"]
        s["Text.Upper"]                := ["Clipboard: UPPERCASE", "剪贴板: 转大写"]
        s["Text.Lower"]                := ["Clipboard: lowercase", "剪贴板: 转小写"]
        s["Text.Title"]                := ["Clipboard: Title Case", "剪贴板: 首字母大写"]
        s["Text.SortAsc"]              := ["Clipboard: Sort Lines A-Z", "剪贴板: 行排序 A-Z"]
        s["Text.SortDesc"]             := ["Clipboard: Sort Lines Z-A", "剪贴板: 行排序 Z-A"]
        s["Text.TrimLines"]            := ["Clipboard: Trim Each Line", "剪贴板: 去除行首尾空白"]
        s["Text.RemoveBlank"]          := ["Clipboard: Remove Blank Lines", "剪贴板: 删除空行"]
        s["Text.Dedupe"]               := ["Clipboard: Remove Duplicate Lines", "剪贴板: 删除重复行"]
        s["Text.Reverse"]              := ["Clipboard: Reverse Text", "剪贴板: 反转文字"]
        s["Text.ToTraditional"]        := ["Clipboard: Simplified → Traditional", "剪贴板: 简体转繁体"]
        s["Text.ToSimplified"]         := ["Clipboard: Traditional → Simplified", "剪贴板: 繁体转简体"]
        s["Text.UrlEncode"]            := ["Clipboard: URL Encode", "剪贴板: URL 编码"]
        s["Text.Done"]                 := ["Clipboard converted", "剪贴板已转换"]
        s["Text.Empty"]                := ["Clipboard is empty", "剪贴板为空"]

        ; --- Windows tools ---
        s["Tool.TaskManager"]          := ["Task Manager", "任务管理器"]
        s["Tool.ControlPanel"]         := ["Control Panel", "控制面板"]
        s["Tool.Settings"]             := ["Windows Settings", "Windows 设置"]
        s["Tool.DeviceManager"]        := ["Device Manager", "设备管理器"]
        s["Tool.Services"]             := ["Services", "服务"]
        s["Tool.Registry"]             := ["Registry Editor", "注册表编辑器"]
        s["Tool.EventViewer"]          := ["Event Viewer", "事件查看器"]
        s["Tool.DiskManagement"]       := ["Disk Management", "磁盘管理"]
        s["Tool.ComputerManagement"]   := ["Computer Management", "计算机管理"]
        s["Tool.TaskScheduler"]        := ["Task Scheduler", "任务计划程序"]
        s["Tool.Programs"]             := ["Programs and Features", "程序和功能"]
        s["Tool.SystemProperties"]     := ["System Properties", "系统属性"]
        s["Tool.Network"]              := ["Network Connections", "网络连接"]
        s["Tool.Firewall"]             := ["Windows Defender Firewall", "Windows Defender 防火墙"]
        s["Tool.ResourceMonitor"]      := ["Resource Monitor", "资源监视器"]
        s["Tool.DiskCleanup"]          := ["Disk Cleanup", "磁盘清理"]
        s["Tool.SystemConfig"]         := ["System Configuration", "系统配置"]
        s["Tool.GroupPolicy"]          := ["Group Policy Editor", "组策略编辑器"]
        s["Tool.CommandPrompt"]        := ["Command Prompt", "命令提示符"]
        s["Tool.PowerShell"]           := ["PowerShell", "PowerShell"]
        s["Tool.Explorer"]             := ["File Explorer", "文件资源管理器"]
        s["Tool.RecycleBin"]           := ["Recycle Bin", "回收站"]
        s["Tool.ThisPC"]               := ["This PC", "此电脑"]
        s["Tool.Printers"]             := ["Devices and Printers", "设备和打印机"]
        s["Tool.Notepad"]              := ["Notepad", "记事本"]
        s["Tool.Calculator"]           := ["Calculator", "计算器"]
        s["Tool.Paint"]                := ["Paint", "画图"]
        s["Tool.WinVer"]               := ["About Windows", "关于 Windows"]
        s["Tool.Subtitle"]             := ["Windows tool", "Windows 工具"]

        ; --- Preferences window ---
        s["Prefs.Title"]               := ["ALTRun Preferences", "ALTRun 偏好设置"]
        s["Prefs.Save"]                := ["Save", "保存"]
        s["Prefs.Cancel"]              := ["Cancel", "取消"]
        s["Prefs.SaveHint"]            := ["Saving applies the changes and reloads ALTRun.", "保存后立即生效 (ALTRun 会重新载入)。"]
        s["Prefs.Add"]                 := ["Add...", "添加..."]
        s["Prefs.Edit"]                := ["Edit...", "编辑..."]
        s["Prefs.Delete"]              := ["Delete", "删除"]
        s["Prefs.Browse"]              := ["...", "..."]
        s["Prefs.ConfirmDelete"]       := ["Delete '{1}'?", "确定删除 '{1}' 吗?"]
        s["Prefs.Required"]            := ["'{1}' cannot be empty.", "'{1}' 不能为空。"]
        s["Prefs.HotkeyHint"]          := ["AutoHotkey syntax: ! Alt  ^ Ctrl  + Shift  # Win, e.g. !Space = Alt+Space", "AutoHotkey 写法: ! Alt  ^ Ctrl  + Shift  # Win, 例如 !Space = Alt+Space"]
        s["Prefs.ListHint"]            := ["One item per line", "每行一项"]
        s["Prefs.Page.General"]        := ["General", "通用"]
        s["Prefs.Page.Appearance"]     := ["Appearance", "外观"]
        s["Prefs.Page.Features"]       := ["Features", "功能"]
        s["Prefs.Page.Applications"]   := ["Applications", "应用搜索"]
        s["Prefs.Page.FileSearch"]     := ["File Search", "文件搜索"]
        s["Prefs.FileInDefault"]       := ["Show matching files and folders in normal results", "直接输入名称时也显示匹配的文件和文件夹"]
        s["Prefs.FileDefaultLimit"]    := ["Files shown in normal results", "普通结果里显示的文件数"]
        s["Prefs.FileMaxResults"]      := ["Results for ' / open search", "' / open 搜索的结果数"]
        s["Prefs.UseEverything"]       := ["Use Everything when it is running (whole disk)", "Everything 运行时用它搜索 (全盘)"]
        s["Prefs.EverythingFilter"]    := ["Everything filter", "Everything 过滤条件"]
        s["Prefs.ScopeFolders"]        := ["Folders indexed without Everything", "没有 Everything 时索引的文件夹"]
        s["Prefs.ScopeDepth"]          := ["Subfolder depth", "子文件夹深度"]
        s["Prefs.RebuildFileIndex"]    := ["Rebuild File Index", "重建文件索引"]
        s["Prefs.EverythingStatus"]    := ["Everything: {1}", "Everything: {1}"]
        s["Prefs.Running"]             := ["running - searching the whole disk", "正在运行 - 搜索全盘"]
        s["Prefs.NotRunning"]          := ["not running - using the built-in index", "没有运行 - 使用内置索引"]
        s["Prefs.Page.Commands"]       := ["Custom Commands", "自定义命令"]
        s["Prefs.Page.Snippets"]       := ["Snippets", "文字片段"]
        s["Prefs.Page.Clipboard"]      := ["Clipboard", "剪贴板历史"]
        s["Prefs.Page.WebSearch"]      := ["Web Search", "网页搜索"]
        s["Prefs.Page.Hotkeys"]        := ["Hotkeys", "自定义热键"]
        s["Prefs.Page.Extensions"]     := ["Extensions", "扩展功能"]
        s["Prefs.Page.Advanced"]       := ["Advanced", "高级"]
        ; General
        s["Prefs.Hotkey"]              := ["ALTRun hotkey", "呼出热键"]
        s["Prefs.SecondaryHotkey"]     := ["Second hotkey", "第二个呼出热键"]
        s["Prefs.Language"]            := ["Language", "界面语言"]
        s["Prefs.Language.auto"]       := ["Automatic", "自动"]
        s["Prefs.LaunchAtLogin"]       := ["Launch ALTRun at login", "开机自动启动"]
        s["Prefs.ShowTrayIcon"]        := ["Show tray icon", "显示托盘图标"]
        s["Prefs.HideOnDeactivate"]    := ["Hide the window when it loses focus", "失去焦点时隐藏窗口"]
        s["Prefs.EnglishInput"]        := ["Switch to English input when shown", "显示窗口时切换到英文输入法"]
        s["Prefs.SendToMenu"]          := ["Add to Explorer 'Send to' menu", "添加到资源管理器 '发送到' 菜单"]
        s["Prefs.StartMenu"]           := ["Add to Start menu", "添加到开始菜单"]
        s["Prefs.CheckUpdates"]        := ["Check for updates at startup", "启动时检查更新"]
        s["Prefs.SaveLog"]             := ["Write a debug log", "写入调试日志"]
        s["Prefs.FileManager"]         := ["File manager", "文件管理器"]
        s["Prefs.HistorySize"]         := ["Search history size", "搜索历史条数"]
        ; Appearance
        s["Prefs.Theme"]               := ["Theme", "主题"]
        s["Prefs.Width"]               := ["Window width (px)", "窗口宽度 (像素)"]
        s["Prefs.VisibleRows"]         := ["Visible results (1-9)", "显示结果行数 (1-9)"]
        s["Prefs.ThemeHint"]           := ["Custom themes: Themes\<Name>.json. Copy a theme below to start from it.", "自定义主题: Themes\<名称>.json。可以先用下面的按钮复制一个主题再修改。"]
        s["Prefs.CopyTheme"]           := ["Copy as Custom Theme...", "复制为自定义主题..."]
        s["Prefs.CopyThemePrompt"]     := ["Name of the new theme (saved in the Themes folder):", "新主题的名称 (保存在 Themes 文件夹):"]
        s["Prefs.ThemeExists"]         := ["Theme '{1}' already exists. Replace it?", "主题 '{1}' 已经存在, 要覆盖吗?"]
        s["Theme.System"]              := ["System (Light / Dark)", "跟随系统 (浅色 / 深色)"]
        s["Theme.Light"]               := ["Light", "浅色"]
        s["Theme.Dark"]                := ["Dark", "深色"]
        s["Theme.Classic"]             := ["Classic", "经典"]
        s["Theme.Midnight"]            := ["Midnight", "午夜"]
        s["Theme.Frost"]               := ["Frost", "霜白"]
        s["Theme.Graphite"]            := ["Graphite", "石墨"]
        s["Theme.Ocean"]               := ["Ocean", "海洋"]
        s["Theme.Paper"]               := ["Paper", "纸张"]
        s["Prefs.OpenThemes"]          := ["Open Themes Folder", "打开主题文件夹"]
        ; Features
        s["Prefs.EnabledFeatures"]     := ["Enabled features", "启用的功能"]
        s["Prefs.Feature.Applications"] := ["Applications", "应用"]
        s["Prefs.Feature.CustomCommands"] := ["Custom commands", "自定义命令"]
        s["Prefs.Feature.Snippets"]    := ["Snippets", "文字片段"]
        s["Prefs.Feature.Clipboard"]   := ["Clipboard history", "剪贴板历史"]
        s["Prefs.Feature.Calculator"] := ["Calculator", "计算器"]
        s["Prefs.Feature.WebSearch"]   := ["Web search", "网页搜索"]
        s["Prefs.Feature.FileSearch"] := ["File search", "文件搜索"]
        s["Prefs.Feature.Terminal"]    := ["Terminal", "终端"]
        s["Prefs.Feature.System"]      := ["System commands", "系统命令"]
        s["Prefs.StructuralCalc"]      := ["Calculator: add structural results (main bars / rebar area)", "计算器: 附带结构计算 (主筋 / 配筋面积)"]
        s["Prefs.ConfirmActions"]      := ["Confirm before shut down / restart / empty recycle bin", "关机 / 重启 / 清空回收站前确认"]
        s["Prefs.TerminalPrefix"]      := ["Terminal prefix", "终端前缀"]
        s["Prefs.TerminalShell"]       := ["Terminal shell", "终端程序"]
        s["Prefs.FileKeywords"]        := ["File search keywords", "文件搜索关键字"]
        s["Prefs.QuotePrefix"]         := ["' prefix searches files", "以 ' 开头搜索文件"]
        s["Prefs.EverythingPath"]      := ["Everything folder", "Everything 目录"]
        s["Prefs.CommaHint"]           := ["Separate with commas", "用逗号分隔"]
        ; Applications
        s["Prefs.AppFolders"]          := ["Folders to index", "索引的文件夹"]
        s["Prefs.AppFileTypes"]        := ["File types", "文件类型"]
        s["Prefs.AppDepth"]            := ["Subfolder depth", "子文件夹深度"]
        s["Prefs.AppExclude"]          := ["Exclude (regex)", "排除 (正则)"]
        s["Prefs.AppHidden"]           := ["Hidden apps", "已删除的应用"]
        s["Prefs.AppHiddenHint"]       := ["One per line, added by Ctrl+Del in search results. Delete a line to show the app again.", "每行一项, 在搜索结果里按 Ctrl+Del 添加。删掉一行即可恢复显示。"]
        s["App.ConfirmHide"]           := ["Remove '{1}' from search results?`n`nThe program is not uninstalled. You can restore it in Preferences > Applications.", "从搜索结果中删除 '{1}' 吗?`n`n不会卸载程序, 可以在 偏好设置 → 应用搜索 里恢复。"]
        s["Prefs.StoreApps"]           := ["Include Microsoft Store apps", "包含应用商店应用"]
        s["Prefs.MatchPinyin"]         := ["Match Chinese names by pinyin initials", "中文名称按拼音首字母匹配"]
        s["Prefs.RefreshMinutes"]      := ["Refresh index every (minutes)", "索引刷新间隔 (分钟)"]
        ; Lists
        s["Prefs.Col.Title"]           := ["Title", "名称"]
        s["Prefs.Col.Type"]            := ["Type", "类型"]
        s["Prefs.Col.Target"]          := ["Target", "目标"]
        s["Prefs.Col.Arguments"]       := ["Arguments", "参数"]
        s["Prefs.Col.Keyword"]         := ["Keyword", "关键字"]
        s["Prefs.Col.Name"]            := ["Name", "名称"]
        s["Prefs.Col.Text"]            := ["Text", "正文"]
        s["Prefs.Col.Url"]             := ["URL ({query} = search term)", "网址 ({query} = 搜索词)"]
        s["Prefs.Col.Id"]              := ["Id", "Id"]
        s["Prefs.Col.Key"]             := ["Hotkey", "热键"]
        s["Prefs.Col.Action"]          := ["Action", "执行的命令"]
        s["Prefs.Col.WinTitle"]        := ["Only in window (optional)", "只在此窗口生效 (可选)"]
        s["Prefs.Col.AutoExpand"]      := ["Expand automatically when typed", "输入关键字时自动展开"]
        s["Prefs.Type.File"]           := ["File / Program", "文件 / 程序"]
        s["Prefs.Type.Folder"]         := ["Folder", "文件夹"]
        s["Prefs.Type.Command"]        := ["Command line", "命令行"]
        s["Prefs.Type.Url"]            := ["Web address", "网址"]
        s["Prefs.SnippetAutoExpand"]   := ["Expand snippet keywords typed in any app", "在任何程序里输入关键字时自动展开"]
        s["Prefs.ExpandPrefix"]        := ["Keyword prefix", "关键字前缀"]
        s["Prefs.PasteMode"]           := ["Paste by", "粘贴方式"]
        s["Prefs.PasteMode.Clipboard"] := ["Clipboard + Ctrl+V", "剪贴板 + Ctrl+V"]
        s["Prefs.PasteMode.Type"]      := ["Typing the text", "逐字输入"]
        s["Prefs.SnippetKeyword"]      := ["Search keyword", "搜索关键字"]
        s["Prefs.ClipKeyword"]         := ["Keyword", "关键字"]
        s["Prefs.ClipHotkey"]          := ["Hotkey", "热键"]
        s["Prefs.ClipMaxItems"]        := ["Keep items", "保存条数"]
        s["Prefs.ClipPersist"]         := ["Keep history after ALTRun quits", "退出后保留历史"]
        s["Prefs.ClipIgnoreApps"]      := ["Never record from (process names)", "不记录这些程序 (进程名)"]
        s["Prefs.ClipClear"]           := ["Clear History Now", "立即清空历史"]
        s["Prefs.Fallbacks"]           := ["Fallback searches (engine Ids, 'files')", "兜底搜索 (引擎 Id, 'files')"]
        ; Extensions
        s["Prefs.QuickSwitch"]         := ["File dialog quick switch", "对话框快速跳转"]
        s["Prefs.QSExplorer"]          := ["Jump to Explorer folder", "跳到资源管理器目录"]
        s["Prefs.QSTotalCmd"]          := ["Jump to Total Commander folder", "跳到 Total Commander 目录"]
        s["Prefs.QSAuto"]              := ["Jump automatically when switching from Total Commander", "从 Total Commander 切换过来时自动跳转"]
        s["Prefs.AutoDate"]            := ["Ctrl+D adds date", "Ctrl+D 加日期"]
        s["Prefs.DateFormat"]          := ["Date format", "日期格式"]
        s["Prefs.RenameHotkey"]        := ["Rename hotkey", "重命名热键"]
        s["Prefs.AppendHotkey"]        := ["Append hotkey", "备注热键"]
        s["Prefs.EnableExtension"]     := ["Enabled", "启用"]
        ; Advanced
        s["Prefs.SettingsFile"]        := ["Settings file", "设置文件"]
        s["Prefs.EditJson"]            := ["Edit ALTRun.json...", "编辑 ALTRun.json..."]
        s["Prefs.OpenDataFolder"]      := ["Open Data Folder", "打开数据文件夹"]
        s["Prefs.ResetLearning"]       := ["Reset Learned Ranking", "重置学习排序"]
        s["Prefs.ResetDone"]           := ["Learned ranking and search history cleared.", "学习排序和搜索历史已清空。"]
        s["Prefs.Version"]             := ["Version {1}", "版本 {1}"]

        ; --- Update checker ---
        s["Update.Available"]          := ["A new version {1} is available. Open the download page?", "发现新版本 {1}, 是否打开下载页面?"]
        s["Update.Latest"]             := ["You are running the latest version ({1}).", "当前已是最新版本 ({1})。"]
        s["Update.Failed"]             := ["Could not check for updates:`n`n{1}", "检查更新失败:`n`n{1}"]

        ; --- QuickSwitch (file dialog path sync) ---
        s["QuickSwitch.Hint"]          := ["{1}: Total Commander folder  {2}: Explorer folder", "{1}: 跳到 TC 目录  {2}: 跳到资源管理器目录"]
        s["QuickSwitch.NoTC"]          := ["Total Commander is not running.", "Total Commander 没有运行。"]
        s["QuickSwitch.NoExplorer"]    := ["No File Explorer window is open.", "没有打开的资源管理器窗口。"]
        return s
    }
}
