;===============================================================================
; QuickSwitch.ahk - 打开/保存对话框路径快速跳转 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 Listary 的 "Quick Switch" 一样: 在标准的打开/保存文件对话框里按热键, 把对话框
; 跳到 Total Commander (默认 Ctrl+G) 或资源管理器 (默认 Ctrl+E) 当前打开的文件夹。
; AutoSwitch = 1 时, 从 TC 切换到对话框会自动跳转。
; 文件夹面板 (ShowPanel, 默认打开, 和 Listary 的 Quick Switch 窗口一样): 对话框一出现, 下面就贴一个搜索框 +
; 文件夹列表 (所有 TC 窗口的两个面板、资源管理器窗口、最近用过的文件夹), 点一下就跳过去, 不用记热键;
; 搜索框还能搜索文件夹 (PanelSearch = all 时也搜文件, 选中文件跳到它所在的文件夹)。对话框不在前台时隐藏,
; 回到对话框时刷新 (TC 里换了目录也能马上看到)。颜色跟随 ALTRun 的主题。
; 热键 (MenuHotkey, 默认 Ctrl+Shift+G): 光标跳到面板的搜索框, 用键盘选择; 没有打开自动显示时临时显示面板。
;
; 设置 (ALTRun.json -> Extensions.QuickSwitch):
;   Enabled / ExplorerHotkey / TotalCmdHotkey / MenuHotkey / RecentFolders / ShowPanel / PanelSearch / AutoSwitch / DialogWindows / ExcludeWindows / AutoSwitchExclude
;   DialogWindows 等是逗号分隔的窗口条件 (ahk_class / ahk_exe / 标题)
;
; 用法:
;   QuickSwitch.Init(AppSettings.Extension("QuickSwitch"))    启动时
;   QuickSwitch.FolderOfWindow(hwnd)     TC / 资源管理器窗口当前的文件夹 (只读, 给 "在此处打开终端" 用)
;   QuickSwitch.MenuFolders()            面板里的文件夹 [{Path, Group, Tag}]
;   QuickSwitch.RecentFolders(n)         最近用过的 n 个文件夹 (打开过的文件取所在的文件夹)
;===============================================================================

class QuickSwitch {
    static Options := Map()
    static _lastWasTC := false
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
                Hotkey(options["MenuHotkey"], (*) => QuickSwitch.FocusPanel())
        } catch as e {
            Logger.Error("QuickSwitch: cannot register hotkeys - " e.Message)
        }
        OnMessage(0x100, (p*) => QuickSwitch._OnKeyDown(p*))                ; WM_KEYDOWN: 面板搜索框里的 ↑ ↓ Enter Esc

        SetTimer(() => QuickSwitch._Watch(), 250)
    }

    ; 用 WM_KEYDOWN 而不是热键: 中文输入法正在输入 (还没上屏) 时按键是 VK_PROCESSKEY, 不会被当成 Enter / Esc
    ; 只比较记下的窗口句柄: 面板关掉之后控件对象已经销毁, 不能再读它们的 Hwnd
    static _OnKeyDown(wParam, lParam, msg, hwnd) {
        if !(hwnd && (hwnd = QuickSwitch._searchHwnd || hwnd = QuickSwitch._listHwnd))
            return
        switch wParam {
            case 0x26: QuickSwitch.PanelKey("Up")
            case 0x28: QuickSwitch.PanelKey("Down")
            case 0x0D: QuickSwitch.PanelKey("Enter")
            case 0x1B: QuickSwitch.PanelKey("Escape")
            default:   return
        }
        return 0
    }

    ; 每 250 ms: AutoSwitch 时从 TC 切到对话框自动跳转; 显示 / 隐藏 / 摆放文件夹面板
    static _Watch() {
        if QuickSwitch.PanelActive() {                                      ; 正在面板的搜索框里输入: 对话框和面板都保持原样
            if !WinExist("ahk_id " QuickSwitch._panelFor)                   ; 对话框已经关掉 (程序自己关的): 面板也关掉
                QuickSwitch.HidePanel()
            return
        }
        isDialog := QuickSwitch.IsFileDialog()
        if (isDialog && QuickSwitch.Options["AutoSwitch"] && QuickSwitch._lastWasTC && !WinActive("ahk_group ALTRunAutoSwitchExclude"))
            QuickSwitch.SyncTotalCmdPath(true)
        QuickSwitch._lastWasTC := WinActive("ahk_class TTOTAL_CMD") ? true : false
        if (QuickSwitch._PanelEnabled() || IsObject(QuickSwitch._panel))    ; 关掉了自动显示时, 按热键临时打开的面板也要跟着对话框
            QuickSwitch._UpdatePanel(isDialog)
    }

    static _SearchesFiles() => QuickSwitch.Options.Has("PanelSearch") && QuickSwitch.Options["PanelSearch"] = "all"

    static _PanelEnabled() => QuickSwitch.Options.Has("ShowPanel") && QuickSwitch.Options["ShowPanel"]

    ; 热键 (MenuHotkey, 默认 Ctrl+Shift+G): 光标跳到面板的搜索框, 用键盘选择; 没有打开自动显示时临时显示面板
    static FocusPanel() {
        dialog := WinExist("A")
        if (!IsObject(QuickSwitch._panel) || QuickSwitch._panelFor != dialog) {
            QuickSwitch._BuildPanel(dialog)
            QuickSwitch._PlacePanel(dialog)
        }
        if !IsObject(QuickSwitch._panel)
            return
        try {
            WinActivate("ahk_id " QuickSwitch._panel.Hwnd)
            ControlFocus(QuickSwitch._panelSearch)
        }
    }

    ;---------------------------------------------------------------------------
    ; 文件夹面板 (和 Listary 的 Quick Switch 窗口一样): 贴在对话框下面, 和对话框一样宽;
    ; 上面是搜索框 (先过滤列表, 再用 Everything / 内置索引找文件夹和文件), 下面是列表。
    ; 点一项或在搜索框里按 Enter 就跳过去 (文件见 PickFile); Esc 回到对话框
    ;---------------------------------------------------------------------------
    static _UpdatePanel(isDialog) {
        dialog := 0
        try dialog := isDialog ? WinGetID("A") : 0
        if !dialog
            return QuickSwitch.HidePanel()
        if (dialog != QuickSwitch._panelFor) {                              ; 新的对话框, 或者从别的窗口切回来: 重新列出文件夹
            if !QuickSwitch._PanelEnabled()
                return QuickSwitch.HidePanel()
            QuickSwitch._BuildPanel(dialog)
        }
        QuickSwitch._PlacePanel(dialog)
    }

    ; 正在面板的搜索框里输入 (面板是前台窗口)
    static PanelActive(*) => IsObject(QuickSwitch._panel) && WinActive("ahk_id " QuickSwitch._panel.Hwnd)

    static HidePanel() {
        if IsObject(QuickSwitch._panel)
            try QuickSwitch._panel.Destroy()
        QuickSwitch._panel := "", QuickSwitch._panelFor := 0, QuickSwitch._panelPos := ""
        QuickSwitch._panelSearch := "", QuickSwitch._panelList := "", QuickSwitch._panelLine := ""   ; 控件已经销毁, 不再引用
        QuickSwitch._searchHwnd := 0, QuickSwitch._listHwnd := 0
    }

    static PanelRows := 8

    static _BuildPanel(dialog) {
        QuickSwitch.HidePanel()
        QuickSwitch._panelFor := dialog
        QuickSwitch._panelBase := QuickSwitch.MenuFolders()
        rect := QuickSwitch._FrameRect(dialog)
        width := QuickSwitch.PanelWidthFor(IsObject(rect) ? rect.W : 600)
        colors := QuickSwitch.PanelColors()                                 ; 跟随 ALTRun 的主题 (浅色 / 深色...)
        panel := Gui("-Caption +ToolWindow +AlwaysOnTop +Border", "ALTRun Quick Switch")
        panel.MarginX := 0, panel.MarginY := 0, panel.BackColor := colors.Background
        panel.SetFont("s10 c" colors.Text, QuickSwitch._FontName())
        search := panel.AddEdit("x10 y8 w" (width - 20) " r1 -Multi -E0x200 Background" colors.Background)
        hint := (QuickSwitch.Options.Has("MenuHotkey") && QuickSwitch.Options["MenuHotkey"] != "") ? "  (" Win.HotkeyLabel(QuickSwitch.Options["MenuHotkey"]) ")" : ""
        DllCall("SendMessage", "Ptr", search.Hwnd, "UInt", 0x1501, "Ptr", 1, "WStr", " " I18n.T(QuickSwitch._SearchesFiles() ? "QuickSwitch.SearchCue" : "QuickSwitch.SearchCueFolders") hint)   ; EM_SETCUEBANNER: 灰色提示文字
        search.OnEvent("Change", (*) => SetTimer(QuickSwitch._searchTimer, -150))
        line := panel.AddText("x0 y+6 w" width " h1 Background" colors.Separator)   ; 搜索框和列表之间的分隔线
        panel.SetFont("s9 c" colors.Text)
        list := panel.AddListView("x0 y+2 w" width " h100 -Hdr -Multi -E0x200 +LV0x10400 Background" colors.Background, ["Name", "Folder", "From"])   ; 0x400 = 路径太长时悬停显示完整路径
        icons := IL_Create(2)
        IL_Add(icons, "shell32.dll", 4)                                     ; 1 = 文件夹
        IL_Add(icons, "shell32.dll", 1)                                     ; 2 = 文件
        list.SetImageList(icons)
        list.OnEvent("Click", (ctrl, row) => QuickSwitch._PanelJump(row))
        list.OnNotify(-12, (ctrl, lParam) => QuickSwitch._OnCustomDraw(lParam))   ; NM_CUSTOMDRAW: 选中行和灰色文字用主题的颜色
        list.Add("Icon1", "X")                                              ; 先放一行才量得出行高 (LVM_APPROXIMATEVIEWRECT)
        QuickSwitch._rowTop := SendMessage(0x1040, 1, 0, list) >> 16
        QuickSwitch._rowHeight := (SendMessage(0x1040, 2, 0, list) >> 16) - QuickSwitch._rowTop
        list.Delete()
        QuickSwitch._panel := panel, QuickSwitch._panelList := list, QuickSwitch._panelSearch := search, QuickSwitch._panelWidth := width
        QuickSwitch._panelLine := line, QuickSwitch._panelDialogW := IsObject(rect) ? rect.W : 0, QuickSwitch._panelColors := colors
        QuickSwitch._searchHwnd := search.Hwnd, QuickSwitch._listHwnd := list.Hwnd
        if colors.Dark
            try DllCall("uxtheme\SetWindowTheme", "Ptr", list.Hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)   ; 深色主题: 滚动条和选中行也用深色
        try DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", panel.Hwnd, "UInt", 33, "Int*", 3, "UInt", 4)   ; Win11: 小圆角
        QuickSwitch._ShowFolders(QuickSwitch._panelBase)
    }

    static _searchHwnd := 0, _listHwnd := 0, _panelColors := "", _panelBase := [], _panelList := "", _panelSearch := "", _panelLine := "", _panelWidth := 600, _panelDialogW := 0, _rowTop := 20, _rowHeight := 18

    ; 面板宽度 = 对话框宽度 (420 ~ 1000)
    static PanelWidthFor(dialogWidth) => Max(420, Min(dialogWidth, 1000))

    ; 面板的颜色: 用 ALTRun 当前主题的背景 / 文字 / 分隔线颜色; 主题还没载入时用浅色
    static PanelColors() {
        get(key, fallback) => ((value := ThemeManager.Get(key)) != "") ? value : fallback
        background := get("Background", "FFFFFF")
        return {Background: background, Text: get("Title", "1F1F1F"), Subtitle: get("Subtitle", "808080"), Separator: get("Separator", "E4E4E4")
              , Selected: get("SelectedBackground", "DDE7F6"), SelectedText: get("SelectedTitle", "000000"), SelectedSubtitle: get("SelectedSubtitle", "4A5568")
              , Dark: QuickSwitch.IsDarkColor(background)}
    }

    ; "RRGGBB" -> Windows 的 COLORREF (0x00BBGGRR)
    static ColorRef(hex) {
        value := Integer("0x" hex)
        return ((value & 0xFF) << 16) | (value & 0xFF00) | ((value >> 16) & 0xFF)
    }

    ; 每个格子自己定颜色: 选中行用主题的选中色, 路径和来源用灰色 (NMLVCUSTOMDRAW)
    static _OnCustomDraw(lParam) {
        static CDDS_PREPAINT := 0x1, CDDS_ITEMPREPAINT := 0x10001, CDDS_SUBITEMPREPAINT := 0x30001
        static CDRF_NEWFONT := 0x2, CDRF_NOTIFYITEMDRAW := 0x20, CDRF_NOTIFYSUBITEMDRAW := 0x20
        x64 := (A_PtrSize = 8)
        stage := NumGet(lParam, A_PtrSize * 3, "UInt")
        if (stage = CDDS_PREPAINT)
            return CDRF_NOTIFYITEMDRAW
        if (stage = CDDS_ITEMPREPAINT)
            return CDRF_NOTIFYSUBITEMDRAW
        if (stage != CDDS_SUBITEMPREPAINT)
            return 0
        colors := QuickSwitch._panelColors
        row := NumGet(lParam, x64 ? 56 : 36, "UPtr")
        column := NumGet(lParam, x64 ? 88 : 56, "Int")
        selected := SendMessage(0x102C, row, 0x2, QuickSwitch._panelList) & 0x2   ; LVM_GETITEMSTATE / LVIS_SELECTED
        stateOffset := x64 ? 64 : 40
        NumPut("UInt", NumGet(lParam, stateOffset, "UInt") & ~0x11, lParam, stateOffset)   ; 去掉 CDIS_SELECTED / CDIS_FOCUS, 不画系统的蓝色高亮
        text := selected ? (column ? colors.SelectedSubtitle : colors.SelectedText) : (column ? colors.Subtitle : colors.Text)
        NumPut("UInt", QuickSwitch.ColorRef(text), lParam, x64 ? 80 : 48)
        NumPut("UInt", QuickSwitch.ColorRef(selected ? colors.Selected : colors.Background), lParam, x64 ? 84 : 52)
        return CDRF_NEWFONT
    }

    ; "1E1E1E" 这种颜色是不是深色 (亮度低于一半)
    static IsDarkColor(hex) {
        if !RegExMatch(hex, "i)^[0-9a-f]{6}$")
            return false
        value := Integer("0x" hex)
        return (((value >> 16) & 0xFF) * 299 + ((value >> 8) & 0xFF) * 587 + (value & 0xFF) * 114) / 1000 < 128
    }

    static _FontName() {
        try return ThemeManager.FontName()
        return "Segoe UI"
    }

    ; 对话框的宽度改了: 面板跟着改宽度
    static _ResizePanel(dialogWidth) {
        width := QuickSwitch.PanelWidthFor(dialogWidth)
        QuickSwitch._panelDialogW := dialogWidth
        if (width = QuickSwitch._panelWidth)
            return
        QuickSwitch._panelWidth := width
        QuickSwitch._panelSearch.Move(, , width - 20)
        QuickSwitch._panelLine.Move(, , width)
        QuickSwitch._panelList.Move(, , width)
        QuickSwitch._ShowFolders(QuickSwitch._panelFolders)                ; 重新算列宽, 并调整面板大小和位置
    }
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
            list.Add(entry.HasOwnProp("IsFile") && entry.IsFile ? "Icon2" : "Icon1", QuickSwitch.FolderName(entry.Path), entry.Path, entry.Tag)
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

    ; 搜索框里的文字: 先过滤列表 (路径包含每个词), 再加上搜到的文件夹和文件 (Everything 或内置索引, 最多 30 个)
    static _SearchPanel() {
        if !IsObject(QuickSwitch._panelSearch)
            return
        term := Trim(QuickSwitch._panelSearch.Value)
        QuickSwitch._ShowFolders(QuickSwitch.FilterFolders(QuickSwitch._panelBase, term, term = "" ? [] : QuickSwitch._FindItems(term)))
    }

    ; [{Path, IsFolder}], 文件夹在前。文件夹单独搜 (和文件一起搜时, 名称完全相同的文件夹可能因为不是最近修改的而排不进来);
    ; PanelSearch = "folders" 时只搜文件夹
    static _FindItems(term) {
        found := [], seen := Map()
        try {
            for item in FileSearchProvider.Query(term, 15, false, true)
                found.Push({Path: item.Path, IsFolder: true}), seen[StrLower(item.Path)] := true
            if !QuickSwitch._SearchesFiles()
                return found
            for item in FileSearchProvider.Query(term, 30, false)
                if (!item.IsFolder && !seen.Has(StrLower(item.Path)))
                    found.Push({Path: item.Path, IsFolder: false})
        }
        return found
    }

    ; 过滤 base (不区分大小写, 每个词都要出现在路径里), 再把 extra 里没有重复的加在后面:
    ; extra 的每一项是路径 (文件夹) 或 {Path, IsFolder}; 文件夹的来源标 "搜索", 文件标 "文件"
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
        for item in extra {
            itemPath := IsObject(item) ? item.Path : item
            isFile := IsObject(item) && !item.IsFolder
            if seen.Has(StrLower(RTrim(itemPath, "\")))
                continue
            seen[StrLower(RTrim(itemPath, "\"))] := true
            result.Push({Path: itemPath, Group: "search", IsFile: isFile, Tag: I18n.T(isFile ? "QuickSwitch.TagFile" : "QuickSwitch.TagSearch")})
        }
        return result
    }

    static _PanelJump(row) {
        if (row < 1 || row > QuickSwitch._panelFolders.Length || !QuickSwitch._panelFor)
            return
        entry := QuickSwitch._panelFolders[row], dialog := QuickSwitch._panelFor
        QuickSwitch._BackToDialog()
        if (entry.HasOwnProp("IsFile") && entry.IsFile)
            QuickSwitch.PickFile(entry.Path, dialog)
        else
            QuickSwitch.SetDialogPath(RTrim(entry.Path, "\") "\", dialog)
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
        if (QuickSwitch._panelDialogW && rect.W != QuickSwitch._panelDialogW)
            return QuickSwitch._ResizePanel(rect.W)
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
    ; 面板里的文件夹
    ;---------------------------------------------------------------------------
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

    ; 在对话框里选中一个文件: 打开和保存对话框都一样, 跳到它所在的文件夹 (不替用户打开或保存)
    static PickFile(filePath, dialog := 0) {
        SplitPath(filePath, , &dir)
        if (dir != "")
            QuickSwitch.SetDialogPath(RTrim(dir, "\") "\", dialog)
    }

    ; 文件名框里原来是普通的文件名时才填回去 (空的、路径、*.txt 这样的筛选条件不填)
    static ShouldRestoreName(saved) {
        saved := Trim(saved)
        return saved != "" && !RegExMatch(saved, "[\\/:*?]")
    }

    ; dialog: 要跳转的对话框 (默认是当前窗口)。模拟键盘输入之前确认它是前台窗口,
    ; 否则文字会打到别的窗口里 (例如切回对话框失败、面板还在前台)
    static SetDialogPath(folder, dialog := 0) {
        if (folder = "" || !FileExist(folder))
            return
        if !dialog
            dialog := WinExist("A")
        if !(dialog && WinExist("ahk_id " dialog))
            return
        Usage.Count("QuickSwitch")
        className := ""
        try className := WinGetClass("ahk_id " dialog)
        if (className = "Qt5QWindowIcon") {                                 ; WPS 的对话框没有可用的 Edit 控件, 只能模拟输入
            if !WinActive("ahk_id " dialog)
                return
            SendText(folder)
            SendInput("{Enter}")
            return
        }
        try {
            saved := ""
            try saved := ControlGetText("Edit1", "ahk_id " dialog)             ; 文件名框原来的内容 (例如另存为时程序预填的文件名)
            ControlFocus("Edit1", "ahk_id " dialog)
            ControlSetText(folder, "Edit1", "ahk_id " dialog)
            ControlSend("{Enter}", "Edit1", "ahk_id " dialog)
            if QuickSwitch.ShouldRestoreName(saved) {                       ; 跳过去之后填回去, 不丢掉预填的文件名
                Sleep(150)
                try ControlSetText(saved, "Edit1", "ahk_id " dialog)
            }
        } catch {
            if !WinActive("ahk_id " dialog)
                return
            SendInput("!d")                                                 ; 没有 Edit1 的自绘对话框: 用地址栏
            Sleep(60)
            SendText(folder)
            SendInput("{Enter}")
        }
    }
}
