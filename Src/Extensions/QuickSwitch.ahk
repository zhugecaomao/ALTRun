;===============================================================================
; QuickSwitch.ahk - 打开/保存对话框路径快速跳转 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 Listary 的 "Quick Switch" 一样: 在标准的打开/保存文件对话框里按热键, 把对话框
; 跳到 Total Commander (默认 Ctrl+G) 或资源管理器 (默认 Ctrl+E) 当前打开的文件夹。
; 对话框标题上会提示这两个热键。AutoSwitch = 1 时, 从 TC 切换到对话框会自动跳转。
; 文件夹菜单 (MenuHotkey, 默认 Ctrl+Shift+G, 和 Listary 的 Quick Switch 菜单一样): 列出所有 TC 窗口的
; 两个面板、打开的资源管理器窗口和最近用过的文件夹 (Windows 的 "最近使用的项目"), 选一个对话框就跳过去。
; 文件夹面板 (ShowPanel, 默认打开, 和 Listary 的 Quick Switch 窗口一样): 对话框一出现, 下面就贴一个搜索框 + 同样内容的
; 列表, 点一下就跳过去, 不用记热键; 在搜索框里输入还能搜索其它文件夹。对话框不在前台时隐藏, 回到对话框时刷新
; (TC 里换了目录也能马上看到)。
;
; 设置 (ALTRun.json -> Extensions.QuickSwitch):
;   Enabled / ExplorerHotkey / TotalCmdHotkey / MenuHotkey / RecentFolders / ShowPanel / AutoSwitch / DialogWindows / ExcludeWindows / AutoSwitchExclude
;   DialogWindows 等是逗号分隔的窗口条件 (ahk_class / ahk_exe / 标题)
;
; 用法:
;   QuickSwitch.Init(AppSettings.Extension("QuickSwitch"))    启动时
;   QuickSwitch.FolderOfWindow(hwnd)     TC / 资源管理器窗口当前的文件夹 (只读, 给 "在此处打开终端" 用)
;   QuickSwitch.MenuFolders()            文件夹菜单的内容 [{Path, Group, Tag}]
;   QuickSwitch.RecentFolders(n)         最近用过的 n 个文件夹 (打开过的文件取所在的文件夹)
;===============================================================================

class QuickSwitch {
    static Options := Map()
    static _titles := Map(), _hintedHwnd := 0, _lastWasTC := false
    static _panel := "", _panelFor := 0, _panelFolders := [], _panelPos := ""

    static Init(options) {
        QuickSwitch.Options := options
        if !options["Enabled"]
            return
        Loop Parse, options["DialogWindows"], ","
            GroupAdd("ALTRunDialogs", Trim(A_LoopField))
        Loop Parse, options["ExcludeWindows"], ","
            GroupAdd("ALTRunDialogExclude", Trim(A_LoopField))
        if options.Has("AutoSwitchExclude")
            Loop Parse, options["AutoSwitchExclude"], ","
                if (Trim(A_LoopField) != "")
                    GroupAdd("ALTRunAutoSwitchExclude", Trim(A_LoopField))

        HotIf((*) => QuickSwitch.IsFileDialog())
        try {
            if (options["ExplorerHotkey"] != "")
                Hotkey(options["ExplorerHotkey"], (*) => QuickSwitch.SyncExplorerPath())
            if (options["TotalCmdHotkey"] != "")
                Hotkey(options["TotalCmdHotkey"], (*) => QuickSwitch.SyncTotalCmdPath())
            if (options["MenuHotkey"] != "")
                Hotkey(options["MenuHotkey"], (*) => QuickSwitch.ShowMenu())
        } catch as e {
            Logger.Error("QuickSwitch: cannot register hotkeys - " e.Message)
        }
        HotIf((*) => QuickSwitch.PanelActive())
        for key in ["Up", "Down", "Enter", "NumpadEnter", "Escape"]
            Hotkey(key, QuickSwitch._KeyHandler(key = "NumpadEnter" ? "Enter" : key))
        HotIf()

        SetTimer(() => QuickSwitch._Watch(), 250)
    }

    static _KeyHandler(key) => (*) => QuickSwitch.PanelKey(key)

    ; 每 250 ms: 在对话框标题上显示热键提示; AutoSwitch 时从 TC 切到对话框自动跳转
    static _Watch() {
        if QuickSwitch.PanelActive()                                        ; 正在面板的搜索框里输入: 对话框和面板都保持原样
            return
        isDialog := QuickSwitch.IsFileDialog()
        if (isDialog && QuickSwitch.Options["AutoSwitch"] && QuickSwitch._lastWasTC && !WinActive("ahk_group ALTRunAutoSwitchExclude"))
            QuickSwitch.SyncTotalCmdPath(true)
        QuickSwitch._lastWasTC := WinActive("ahk_class TTOTAL_CMD") ? true : false
        QuickSwitch._UpdateHint(isDialog)
        if (QuickSwitch.Options.Has("ShowPanel") && QuickSwitch.Options["ShowPanel"])
            QuickSwitch._UpdatePanel(isDialog)
    }

    ;---------------------------------------------------------------------------
    ; 文件夹面板 (和 Listary 的 Quick Switch 窗口一样): 贴在对话框下面, 和对话框一样宽;
    ; 上面是搜索框 (输入文字搜索文件夹: 先过滤列表, 再用 Everything / 内置索引找文件夹), 下面是文件夹列表。
    ; 点一个文件夹或在搜索框里按 Enter 就跳过去; Esc 回到对话框
    ;---------------------------------------------------------------------------
    static _UpdatePanel(isDialog) {
        dialog := 0
        try dialog := isDialog ? WinGetID("A") : 0
        if !dialog
            return QuickSwitch.HidePanel()
        if (dialog != QuickSwitch._panelFor)                                ; 新的对话框, 或者从别的窗口切回来: 重新列出文件夹
            QuickSwitch._BuildPanel(dialog)
        QuickSwitch._PlacePanel(dialog)
    }

    ; 正在面板的搜索框里输入 (面板是前台窗口)
    static PanelActive(*) => IsObject(QuickSwitch._panel) && WinActive("ahk_id " QuickSwitch._panel.Hwnd)

    static HidePanel() {
        if IsObject(QuickSwitch._panel)
            try QuickSwitch._panel.Destroy()
        QuickSwitch._panel := "", QuickSwitch._panelFor := 0, QuickSwitch._panelPos := ""
    }

    static PanelRows := 8

    static _BuildPanel(dialog) {
        QuickSwitch.HidePanel()
        QuickSwitch._panelFor := dialog
        QuickSwitch._panelBase := QuickSwitch.MenuFolders()
        rect := QuickSwitch._FrameRect(dialog)
        width := IsObject(rect) ? Max(420, Min(rect.W, 1000)) : 600
        panel := Gui("-Caption +ToolWindow +AlwaysOnTop +Border", "ALTRun Quick Switch")
        panel.MarginX := 0, panel.MarginY := 0, panel.BackColor := "FFFFFF"
        panel.SetFont("s10", "Segoe UI")
        search := panel.AddEdit("x8 y7 w" (width - 16) " r1 -Multi")
        hint := (QuickSwitch.Options.Has("MenuHotkey") && QuickSwitch.Options["MenuHotkey"] != "") ? "  (" Win.HotkeyLabel(QuickSwitch.Options["MenuHotkey"]) ")" : ""
        DllCall("SendMessage", "Ptr", search.Hwnd, "UInt", 0x1501, "Ptr", 1, "WStr", " " I18n.T("QuickSwitch.SearchCue") hint)   ; EM_SETCUEBANNER: 灰色提示文字
        search.OnEvent("Change", (*) => SetTimer(QuickSwitch._searchTimer, -150))
        panel.SetFont("s9")
        list := panel.AddListView("x0 y+6 w" width " h100 -Hdr -Multi -E0x200 +LV0x10400 BackgroundFFFFFF", ["Name", "Folder", "From"])   ; 0x400 = 路径太长时悬停显示完整路径
        icons := IL_Create(1)
        IL_Add(icons, "shell32.dll", 4)
        list.SetImageList(icons)
        list.OnEvent("Click", (ctrl, row) => QuickSwitch._PanelJump(row))
        list.Add("Icon1", "X")                                              ; 先放一行才量得出行高 (LVM_APPROXIMATEVIEWRECT)
        QuickSwitch._rowTop := SendMessage(0x1040, 1, 0, list) >> 16
        QuickSwitch._rowHeight := (SendMessage(0x1040, 2, 0, list) >> 16) - QuickSwitch._rowTop
        list.Delete()
        QuickSwitch._panel := panel, QuickSwitch._panelList := list, QuickSwitch._panelSearch := search, QuickSwitch._panelWidth := width
        try DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", panel.Hwnd, "UInt", 33, "Int*", 3, "UInt", 4)   ; Win11: 小圆角
        QuickSwitch._ShowFolders(QuickSwitch._panelBase)
    }

    static _panelBase := [], _panelList := "", _panelSearch := "", _panelWidth := 600, _rowTop := 20, _rowHeight := 18
    static _searchTimer := () => QuickSwitch._SearchPanel()

    ; 列表: 文件夹名 | 完整路径 | 来源 (TC / 资源管理器 / 最近 / 搜索)
    static _ShowFolders(folders) {
        list := QuickSwitch._panelList
        if !IsObject(list)
            return
        QuickSwitch._panelFolders := folders
        list.Opt("-Redraw")
        list.Delete()
        for entry in folders
            list.Add("Icon1", QuickSwitch.FolderName(entry.Path), entry.Path, entry.Tag)
        list.ModifyCol(3, "Auto"), list.ModifyCol(1, "Auto")
        nameW := Min(SendMessage(0x101D, 0, 0, list), QuickSwitch._panelWidth // 3)   ; LVM_GETCOLUMNWIDTH
        list.ModifyCol(1, Max(nameW, 120))
        list.ModifyCol(2, QuickSwitch._panelWidth - Max(nameW, 120) - SendMessage(0x101D, 2, 0, list) - (folders.Length > QuickSwitch.PanelRows ? 22 : 4))
        if folders.Length
            list.Modify(1, "Select Focus")
        list.Opt("+Redraw")
        rows := Max(3, Min(folders.Length, QuickSwitch.PanelRows))           ; 高度跟着内容 (3 ~ PanelRows 行)
        list.Move(, , , QuickSwitch._rowTop + (rows - 1) * QuickSwitch._rowHeight + 4)
        if DllCall("IsWindowVisible", "Ptr", QuickSwitch._panel.Hwnd) {     ; 搜索结果变了: 重新调整大小和位置
            QuickSwitch._panelPos := ""
            QuickSwitch._PlacePanel(QuickSwitch._panelFor)
        }
    }

    ; "D:\Projects\Tower" -> "Tower"; 磁盘根目录显示 "D:\"
    static FolderName(folder) {
        trimmed := RTrim(folder, "\")
        if (trimmed ~= "^[A-Za-z]:$")
            return trimmed "\"
        SplitPath(trimmed, &name)
        return (name != "") ? name : folder
    }

    ; 搜索框里的文字: 先过滤列表 (路径包含每个词), 再加上搜到的文件夹 (Everything 或内置索引, 最多 20 个)
    static _SearchPanel() {
        if !IsObject(QuickSwitch._panelSearch)
            return
        term := Trim(QuickSwitch._panelSearch.Value)
        QuickSwitch._ShowFolders(QuickSwitch.FilterFolders(QuickSwitch._panelBase, term, term = "" ? [] : QuickSwitch._FindFolders(term)))
    }

    static _FindFolders(term) {
        found := []
        try {
            for item in FileSearchProvider.Query(term, 20, false, true)
                found.Push(item.Path)
        }
        return found
    }

    ; 过滤 base (不区分大小写, 每个词都要出现在路径里), 再把 extra 里没有重复的路径加在后面 (来源标 "搜索")
    static FilterFolders(base, term, extra) {
        result := [], seen := Map(), words := StrSplit(Trim(term), " ")
        for entry in base {
            matched := true
            for word in words
                if (word != "" && !InStr(entry.Path, word))
                    matched := false
            if matched {
                result.Push(entry)
                seen[StrLower(RTrim(entry.Path, "\"))] := true
            }
        }
        for folder in extra {
            if seen.Has(StrLower(RTrim(folder, "\")))
                continue
            seen[StrLower(RTrim(folder, "\"))] := true
            result.Push({Path: folder, Group: "search", Tag: I18n.T("QuickSwitch.TagSearch")})
        }
        return result
    }

    static _PanelJump(row) {
        if (row < 1 || row > QuickSwitch._panelFolders.Length || !QuickSwitch._panelFor)
            return
        folder := QuickSwitch._panelFolders[row].Path
        QuickSwitch._BackToDialog()
        QuickSwitch.SetDialogPath(RTrim(folder, "\") "\")
        if (IsObject(QuickSwitch._panelSearch) && QuickSwitch._panelSearch.Value != "") {   ; 跳过去之后清空搜索, 列表恢复原样
            QuickSwitch._panelSearch.Value := ""
            QuickSwitch._ShowFolders(QuickSwitch._panelBase)
        }
    }

    static _BackToDialog() {
        dialog := QuickSwitch._panelFor
        if (!dialog || WinActive("ahk_id " dialog))
            return
        try {
            WinActivate("ahk_id " dialog)
            WinWaitActive("ahk_id " dialog, , 1)
        }
    }

    ; 搜索框里的按键: ↑ ↓ 选择, Enter 跳转, Esc 回到对话框
    static PanelKey(key) {
        list := QuickSwitch._panelList
        if !IsObject(list)
            return
        count := list.GetCount(), row := list.GetNext()
        switch key {
            case "Up", "Down":
                if !count
                    return
                row := (key = "Down") ? Min(row + 1, count) : Max(row - 1, 1)
                list.Modify(0, "-Select"), list.Modify(row, "Select Focus Vis")
            case "Enter":
                QuickSwitch._PanelJump(row ? row : 1)
            case "Escape":
                QuickSwitch._BackToDialog()
        }
    }

    static _PlacePanel(dialog) {
        if !IsObject(QuickSwitch._panel)
            return
        rect := QuickSwitch._FrameRect(dialog)
        if !IsObject(rect)
            return QuickSwitch.HidePanel()
        if !QuickSwitch._panelPos                                           ; 第一次显示或大小变了: 先按内容算出大小
            QuickSwitch._panel.Show(DllCall("IsWindowVisible", "Ptr", QuickSwitch._panel.Hwnd) ? "NA AutoSize" : "NA Hide AutoSize")
        QuickSwitch._panel.GetPos(, , &panelW, &panelH)
        area := Win.WorkAreaAt(rect.X + rect.W // 2, rect.Y + rect.H // 2)
        pos := QuickSwitch.PanelPosition(rect.X, rect.Y, rect.W, rect.H, panelW, panelH, area)
        key := pos.X "," pos.Y
        if (key != QuickSwitch._panelPos) {                                 ; 对话框移动了才重新摆
            QuickSwitch._panel.Show("NA x" pos.X " y" pos.Y)
            QuickSwitch._panelPos := key
        }
    }

    ; 窗口看得见的边框 (Win10/11 窗口外面还有一圈看不见的调整大小的边框, WinGetPos 包括它)
    static _FrameRect(hwnd) {
        rect := Buffer(16, 0)
        if !DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", hwnd, "UInt", 9, "Ptr", rect, "UInt", 16) {   ; DWMWA_EXTENDED_FRAME_BOUNDS
            left := NumGet(rect, 0, "Int"), top := NumGet(rect, 4, "Int")
            if (NumGet(rect, 8, "Int") > left)
                return {X: left, Y: top, W: NumGet(rect, 8, "Int") - left, H: NumGet(rect, 12, "Int") - top}
        }
        try {
            WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
            return {X: x, Y: y, W: w, H: h}
        }
        return ""
    }

    ; 面板的位置 (和 Listary 一样): 对话框正下方, 左边对齐; 下面放不下时放在对话框上方;
    ; 上下都放不下时贴在对话框里面的底部
    static PanelPosition(x, y, w, h, panelW, panelH, area) {
        left := Max(area.Left, Min(x, area.Right - panelW))
        if (y + h + panelH <= area.Bottom)
            return {X: left, Y: y + h}
        if (y - panelH >= area.Top)
            return {X: left, Y: y - panelH}
        return {X: left, Y: Max(area.Top, area.Bottom - panelH)}
    }

    static _UpdateHint(isDialog) {
        titles := QuickSwitch._titles
        if isDialog {
            hwnd := WinGetID("A")
            if (QuickSwitch._hintedHwnd && QuickSwitch._hintedHwnd != hwnd)
                QuickSwitch._RestoreTitle(QuickSwitch._hintedHwnd)
            if !titles.Has(hwnd)
                titles[hwnd] := WinGetTitle("ahk_id " hwnd)
            title := titles[hwnd] " / " QuickSwitch.HintText()
            if (WinGetTitle("ahk_id " hwnd) != title)
                WinSetTitle(title, "ahk_id " hwnd)
            QuickSwitch._hintedHwnd := hwnd
        } else if QuickSwitch._hintedHwnd {
            QuickSwitch._RestoreTitle(QuickSwitch._hintedHwnd)
            QuickSwitch._hintedHwnd := 0
        }
    }

    ; "Ctrl+G: 跳到 TC 目录  Ctrl+E: 跳到资源管理器目录  Ctrl+Shift+G: 文件夹菜单", 没设置的热键不显示
    static HintText() {
        hint := ""
        for pair in [["TotalCmdHotkey", "QuickSwitch.HintTC"], ["ExplorerHotkey", "QuickSwitch.HintExplorer"], ["MenuHotkey", "QuickSwitch.HintMenu"]]
            if (QuickSwitch.Options.Has(pair[1]) && QuickSwitch.Options[pair[1]] != "")
                hint .= (hint = "" ? "" : "  ") Win.HotkeyLabel(QuickSwitch.Options[pair[1]]) ": " I18n.T(pair[2])
        return hint
    }

    static _RestoreTitle(hwnd) {
        if QuickSwitch._titles.Has(hwnd) {
            try WinSetTitle(QuickSwitch._titles[hwnd], "ahk_id " hwnd)
            QuickSwitch._titles.Delete(hwnd)
        }
    }

    ; 只认真正的打开/保存文件对话框 (有文件名输入框 + 文件列表 + 按钮或路径栏)
    static IsFileDialog(*) {
        if (!WinActive("ahk_group ALTRunDialogs") || WinActive("ahk_group ALTRunDialogExclude"))
            return false
        controls := []
        try controls := WinGetControls("A")
        if !controls.Length {
            title := ""
            try title := WinGetTitle("A")
            return RegExMatch(title, "i)(open|save|select|import|export|browse|打开|另存|选择|导入|导出)") > 0
        }
        hasNameEdit := QuickSwitch._HasControl(controls, "^Edit\d+$")
        hasFileView := QuickSwitch._HasControl(controls, "^DirectUIHWND\d+$", "^SHELLDLL_DefView\d*$", "^SysListView32\d*$")
        hasButtons  := QuickSwitch._HasControl(controls, "^Button\d+$", "^ToolbarWindow32\d+$", "^ComboBoxEx32\d+$", "^ComboBox\d+$")
        return hasNameEdit && hasFileView && hasButtons
    }

    static _HasControl(controls, patterns*) {
        for name in controls
            for pattern in patterns
                if RegExMatch(name, "i)" pattern)
                    return true
        return false
    }

    static SyncTotalCmdPath(silent := false) {
        folder := QuickSwitch.TotalCmdFolder()
        if (folder = "")
            return silent ? "" : MsgBox(I18n.T("QuickSwitch.NoTC"), App.Name, 48)
        QuickSwitch.SetDialogPath(folder "\")                               ; 结尾带反斜杠, AutoCAD 等程序才认
    }

    static SyncExplorerPath() {
        folder := QuickSwitch.ExplorerFolder()
        if (folder = "")
            return MsgBox(I18n.T("QuickSwitch.NoExplorer"), App.Name, 48)
        QuickSwitch.SetDialogPath(folder)
    }

    ; Total Commander 面板的文件夹 (TC 7 ~ 11: WM_USER+75, 2029 / 2030 = 当前 / 另一个面板的路径放到剪贴板)
    ; hwnd 为 0 时取最前面的 TC 窗口; 临时借用剪贴板, 这段时间剪贴板历史不记录
    static TotalCmdFolder(hwnd := 0, otherPanel := false) {
        if !(hwnd := hwnd ? hwnd : WinExist("ahk_class TTOTAL_CMD"))
            return ""
        ClipboardProvider.PauseRecording(1000)
        savedClipboard := ClipboardAll()
        A_Clipboard := ""
        folder := ""
        try {
            SendMessage(1075, otherPanel ? 2030 : 2029, 0, , "ahk_id " hwnd)
            if ClipWait(0.2)
                folder := RTrim(A_Clipboard, "\")
        }
        A_Clipboard := savedClipboard
        return (folder ~= "^([A-Za-z]:|\\\\)") ? folder : ""               ; 只要真正的路径 (不要 FTP、压缩包里面等)
    }

    ;---------------------------------------------------------------------------
    ; 文件夹菜单
    ;---------------------------------------------------------------------------
    static ShowMenu() {
        folders := QuickSwitch.MenuFolders()
        if !folders.Length
            return App.Notify(I18n.T("QuickSwitch.NoFolders"), 2000)
        folderMenu := Menu(), lastGroup := ""
        for index, entry in folders {
            if (index > 1 && entry.Group != lastGroup)
                folderMenu.Add()                                            ; 分隔线: 打开的窗口 / 最近的文件夹
            lastGroup := entry.Group
            label := QuickSwitch.MenuLabel(index, entry)
            folderMenu.Add(label, QuickSwitch._Jumper(entry.Path))
            try folderMenu.SetIcon(label, "shell32.dll", 4)
        }
        try {                                                               ; 显示在文件名输入框下面
            ControlGetPos(&x, &y, , &h, "Edit1", "A")
            return folderMenu.Show(x, y + h)
        }
        folderMenu.Show()
    }

    ; "&1  D:\Projects<Tab>Total Commander": 前 9 项按数字键直接选; 路径里的 & 要写成 &&
    static MenuLabel(index, entry) {
        return (index <= 9 ? "&" index "  " : "     ") StrReplace(entry.Path, "&", "&&") "`t" entry.Tag
    }

    static _Jumper(folder) {
        return (*) => QuickSwitch.SetDialogPath(RTrim(folder, "\") "\")
    }

    ; [{Path, Group: "open" / "recent", Tag}]: 每个 TC 窗口的当前面板和另一个面板、每个资源管理器窗口、
    ; 最近的文件夹 (RecentFolders 个), 去掉重复的
    static MenuFolders() {
        list := [], seen := Map()
        add(folder, group, tag) {
            if (folder = "" || seen.Has(StrLower(RTrim(folder, "\"))))
                return
            seen[StrLower(RTrim(folder, "\"))] := true
            list.Push({Path: folder, Group: group, Tag: tag})
        }
        for hwnd in WinGetList("ahk_class TTOTAL_CMD") {
            add(QuickSwitch.TotalCmdFolder(hwnd), "open", I18n.T("QuickSwitch.TagTC"))
            add(QuickSwitch.TotalCmdFolder(hwnd, true), "open", I18n.T("QuickSwitch.TagTCOther"))
        }
        try {
            for window in ComObject("Shell.Application").Windows
                try add(window.Document.Folder.Self.Path, "open", I18n.T("QuickSwitch.TagExplorer"))
        }
        count := QuickSwitch.Options.Has("RecentFolders") ? QuickSwitch.Options["RecentFolders"] : 10
        for folder in QuickSwitch.RecentFolders(count)
            add(folder, "recent", I18n.T("QuickSwitch.TagRecent"))
        return list
    }

    ; Windows "最近使用的项目" (Recent 文件夹里的快捷方式) 里最近的 limit 个文件夹; 打开过的文件取所在的文件夹。
    ; 网络上的路径不访问磁盘 (可能很慢), 按有没有扩展名判断是不是文件
    static RecentFolders(limit, recentDir := "") {
        folders := [], seen := Map()
        if (limit <= 0)
            return folders
        recentDir := (recentDir != "") ? recentDir : A_AppData "\Microsoft\Windows\Recent"
        lines := ""
        Loop Files, recentDir "\*.lnk"
            lines .= A_LoopFileTimeModified "`t" A_LoopFileFullPath "`n"
        for line in StrSplit(Sort(RTrim(lines, "`n"), "R"), "`n") {
            if (line = "" || A_Index > limit * 6)                           ; 只看最近的一部分, 不用读完几百个快捷方式
                break
            target := ""
            try FileGetShortcut(SubStr(line, InStr(line, "`t") + 1), &target)
            if !(target ~= "^([A-Za-z]:\\|\\\\)")
                continue
            if IconCache.IsRemote(target)
                folder := (IconCache._Extension(target) = "") ? target : QuickSwitch._Parent(target)
            else if (attributes := FileExist(target))
                folder := InStr(attributes, "D") ? target : QuickSwitch._Parent(target)
            else
                continue                                                    ; 已经删掉或移走了
            folder := RTrim(folder, "\")
            if (folder ~= "^[A-Za-z]:$")
                folder .= "\"
            if (folder = "" || seen.Has(StrLower(folder)))
                continue
            seen[StrLower(folder)] := true
            folders.Push(folder)
            if (folders.Length >= limit)
                break
        }
        return folders
    }

    static _Parent(target) {
        SplitPath(RTrim(target, "\"), , &dir)
        return dir
    }

    ; 资源管理器窗口当前的文件夹; hwnd 为 0 时取最前面的资源管理器窗口
    static ExplorerFolder(hwnd := 0) {
        if !hwnd
            hwnd := WinExist("ahk_class CabinetWClass")
        if !hwnd
            return ""
        try {
            for window in ComObject("Shell.Application").Windows
                if (window.HWND = hwnd)
                    return window.Document.Folder.Self.Path
        }
        return ""
    }

    ; 最前面的 TC / 资源管理器窗口当前的文件夹: {Path, Tag}; 没有时 "" (文件操作 "复制到 ..." 用)
    static FileManagerFolder() {
        for hwnd in WinGetList() {                                          ; 按窗口的前后顺序
            className := ""
            try className := WinGetClass("ahk_id " hwnd)
            switch className {
                case "TTOTAL_CMD"   : return {Path: QuickSwitch.TotalCmdFolder(hwnd), Tag: "Total Commander"}
                case "CabinetWClass": return {Path: QuickSwitch.ExplorerFolder(hwnd), Tag: I18n.T("QuickSwitch.TagExplorer")}
            }
        }
        return ""
    }

    static FolderOfWindow(hwnd) {
        if (!hwnd || !WinExist("ahk_id " hwnd))
            return ""
        switch WinGetClass("ahk_id " hwnd) {
            case "TTOTAL_CMD"   : return QuickSwitch.TotalCmdFolder()
            case "CabinetWClass": return QuickSwitch.ExplorerFolder(hwnd)
        }
        return ""
    }

    static SetDialogPath(folder) {
        if (folder = "" || !FileExist(folder))
            return
        Usage.Count("QuickSwitch")
        if (WinGetClass("A") = "Qt5QWindowIcon") {                          ; WPS 的对话框没有可用的 Edit 控件, 只能模拟输入
            SendText(folder)
            SendInput("{Enter}")
            return
        }
        try {
            ControlFocus("Edit1", "A")
            ControlSetText(folder, "Edit1", "A")
            ControlSend("{Enter}", "Edit1", "A")
        } catch {
            SendInput("!d")                                                 ; 没有 Edit1 的自绘对话框: 用地址栏
            Sleep(60)
            SendText(folder)
            SendInput("{Enter}")
        }
    }
}
