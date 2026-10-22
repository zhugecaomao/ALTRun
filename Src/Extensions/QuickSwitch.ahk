;===============================================================================
; QuickSwitch.ahk - 打开/保存对话框路径快速跳转 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like Listary's "Quick Switch": in an Open / Save dialog, the hotkey (Ctrl+G by default) jumps the dialog to the
; folder open in the file manager used last: Total Commander, Directory Opus or File Explorer (see FileManagerFolder).
; AutoSwitch = 1: switching from one of them to a dialog jumps automatically.
; 文件夹面板 (ShowPanel, 默认打开, 和 Listary 的 Quick Switch 窗口一样): 对话框一出现, 下面就贴一个搜索框 +
; 文件夹列表 (所有 TC 窗口的两个面板、资源管理器窗口、最近用过的文件夹), 点一下就跳过去, 不用记热键;
; 搜索框还能搜索文件夹 (PanelSearch = all 时也搜文件, 选中文件跳到它所在的文件夹)。对话框不在前台时隐藏,
; 回到对话框时刷新 (TC 里换了目录也能马上看到)。颜色跟随 ALTRun 的主题。
; 热键 (MenuHotkey, 默认 Ctrl+Shift+G): 光标跳到面板的搜索框, 用键盘选择; 没有打开自动显示时临时显示面板。
;
; 设置 (ALTRun.json -> Extensions.QuickSwitch):
;   Enabled / TotalCmdHotkey (the jump hotkey, for all three) / MenuHotkey / RecentFolders / ShowPanel / PanelSearch / AutoSwitch / DialogWindows / ExcludeWindows / AutoSwitchExclude
;   (ExplorerHotkey is no longer used; renamed / removed with the next SchemaVersion, see #176)
;   DialogWindows 等是逗号分隔的窗口条件 (ahk_class / ahk_exe / 标题)
;
; 用法:
;   QuickSwitch.Init(AppSettings.Extension("QuickSwitch"))    启动时
;   QuickSwitch.FolderOfWindow(hwnd)     current folder of a TC / Opus / Explorer window (read only, for "terminal here")
;   QuickSwitch.MenuFolders()            面板里的文件夹 [{Path, Group, Tag}]
;   QuickSwitch.RecentFolders(n)         最近用过的 n 个文件夹 (打开过的文件取所在的文件夹)
;===============================================================================

class QuickSwitch {
    static Options := Map(), _hasExclude := false
    static _lastWasManager := false
    static _panel := "", _panelFor := 0, _panelFolders := [], _panelPos := ""

    static Init(options) {
        QuickSwitch.Options := options
        if !options["Enabled"]
            return
        Loop Parse, options["DialogWindows"], ","
            GroupAdd("ALTRunDialogs", Trim(A_LoopField))
        QuickSwitch._hasExclude := false
        Loop Parse, options["ExcludeWindows"], ","
            if (Trim(A_LoopField) != "")
                GroupAdd("ALTRunDialogExclude", Trim(A_LoopField)), QuickSwitch._hasExclude := true
        if options.Has("AutoSwitchExclude")
            Loop Parse, options["AutoSwitchExclude"], ","
                if (Trim(A_LoopField) != "")
                    GroupAdd("ALTRunAutoSwitchExclude", Trim(A_LoopField))

        HotIf((*) => QuickSwitch.IsFileDialog())
        try {
            if (options["TotalCmdHotkey"] != "")
                Hotkey(options["TotalCmdHotkey"], (*) => QuickSwitch.SyncFileManagerPath())
            if (options["MenuHotkey"] != "")
                Hotkey(options["MenuHotkey"], (*) => QuickSwitch.FocusPanel())
        } catch as e {
            Logger.Error("QuickSwitch: cannot register hotkeys - " e.Message)
        }
        OnMessage(0x100, (p*) => QuickSwitch._OnKeyDown(p*))                ; WM_KEYDOWN: 面板搜索框里的 ↑ ↓ Enter Esc

        SetTimer(Logger.Watch("quick switch watch", () => QuickSwitch._Watch()), 250)
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
        if (isDialog && QuickSwitch.Options["AutoSwitch"] && QuickSwitch._lastWasManager && !WinActive("ahk_group ALTRunAutoSwitchExclude"))
            QuickSwitch.SyncFileManagerPath(true)
        QuickSwitch._lastWasManager := QuickSwitch.IsFileManager(WinExist("A"))
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
        QuickSwitch._panelBase := QuickSwitch.MenuFolders(false)
        rect := QuickSwitch._FrameRect(dialog)
        width := QuickSwitch.PanelWidthFor(IsObject(rect) ? rect.W : 600)
        colors := QuickSwitch.PanelColors()                                 ; 跟随 ALTRun 的主题 (浅色 / 深色...)
        ; -DPIScale: 面板和对话框一样宽, 宽度是对话框的实际像素; 行高也是量出来的像素。让 AHK 再按 DPI 放大一次的话,
        ; 150% 缩放时面板会比对话框宽一半, 下面多出一片空白。固定的边距用 Win.Scale 换算
        panel := Gui("-Caption +ToolWindow +AlwaysOnTop +Border -DPIScale", "ALTRun Quick Switch")
        panel.MarginX := 0, panel.MarginY := 0, panel.BackColor := colors.Background
        panel.SetFont("s10 c" colors.Text, QuickSwitch._FontName())
        search := panel.AddEdit("x" Win.Scale(10) " y" Win.Scale(8) " w" (width - Win.Scale(20)) " r1 -Multi -E0x200 Background" colors.Background)
        hint := (QuickSwitch.Options.Has("MenuHotkey") && QuickSwitch.Options["MenuHotkey"] != "") ? "  (" Win.HotkeyLabel(QuickSwitch.Options["MenuHotkey"]) ")" : ""
        DllCall("SendMessage", "Ptr", search.Hwnd, "UInt", 0x1501, "Ptr", 1, "WStr", " " I18n.T(QuickSwitch._SearchesFiles() ? "QuickSwitch.SearchCue" : "QuickSwitch.SearchCueFolders") hint)   ; EM_SETCUEBANNER: 灰色提示文字
        search.OnEvent("Change", (*) => SetTimer(QuickSwitch._searchTimer, -150))
        line := panel.AddText("x0 y+" Win.Scale(6) " w" width " h1 Background" colors.Separator)   ; 搜索框和列表之间的分隔线
        panel.SetFont("s9 c" colors.Text)
        list := panel.AddListView("x0 y+" Win.Scale(2) " w" width " h100 -Hdr -Multi -E0x200 +LV0x10400 Background" colors.Background, ["Name", "Folder", "From"])   ; 0x400 = 路径太长时悬停显示完整路径
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
        if WinExist("ahk_class dopus.lister")                               ; reading Opus takes a moment: show the panel first
            SetTimer(QuickSwitch._AddDOpusFolders.Bind(QuickSwitch, dialog), -1)
    }

    static _AddDOpusFolders(dialog) {
        if (QuickSwitch._panelFor != dialog || !IsObject(QuickSwitch._panelList))
            return
        QuickSwitch._panelBase := QuickSwitch.MenuFolders()
        if (QuickSwitch._panelFor = dialog && IsObject(QuickSwitch._panelSearch) && QuickSwitch._panelSearch.Value = "")   ; not while searching
            QuickSwitch._ShowFolders(QuickSwitch._panelBase)
    }

    static _searchHwnd := 0, _listHwnd := 0, _panelColors := "", _panelBase := [], _panelList := "", _panelSearch := "", _panelLine := "", _panelWidth := 600, _panelDialogW := 0, _rowTop := 20, _rowHeight := 18

    ; 面板宽度 = 对话框宽度 (420 ~ 1000, 按 DPI 换算; 都是实际像素)
    static PanelWidthFor(dialogWidth) => Max(Win.Scale(420), Min(dialogWidth, Win.Scale(1000)))

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

    static IsDarkColor(hex) => ThemeManager.IsDarkColor(hex)

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
        QuickSwitch._panelSearch.Move(, , width - Win.Scale(20))
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
        nameW := Max(Min(SendMessage(0x101D, 0, 0, list), QuickSwitch._panelWidth // 3), Win.Scale(120))   ; LVM_GETCOLUMNWIDTH
        list.ModifyCol(1, nameW)
        list.ModifyCol(2, QuickSwitch._panelWidth - nameW - SendMessage(0x101D, 2, 0, list) - Win.Scale(folders.Length > QuickSwitch.PanelRows ? 22 : 4))
        if folders.Length
            list.Modify(1, "Select Focus")
        list.Opt("+Redraw")
        rows := Max(3, Min(folders.Length, QuickSwitch.PanelRows))           ; 高度跟着内容 (3 ~ PanelRows 行)
        list.Move(, , , QuickSwitch._rowTop + (rows - 1) * QuickSwitch._rowHeight + Win.Scale(4))
        list.ModifyCol(3, "AutoHdr")                                        ; 最后一列补满到右边, 选中行的颜色铺满整行; 改完高度 (该有滚动条时已经有了) 再算, 不会多出横向滚动条
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
        browse := QuickSwitch.Options.Has("BrowseDialogs") && QuickSwitch.Options["BrowseDialogs"]
        listed := WinActive("ahk_group ALTRunDialogs")
        if !(listed || browse && WinActive("ahk_class #32770"))
            return false
        if (QuickSwitch._hasExclude && WinActive("ahk_group ALTRunDialogExclude"))
            return false
        controls := [], title := "", winClass := ""                         ; read once for both kinds (this runs every 250 ms)
        try controls := WinGetControls("A")
        try title := WinGetTitle("A")
        try winClass := WinGetClass("A")
        if (browse && QuickSwitch.LooksLikeBrowseDialog(winClass, controls, QuickSwitch._TreeId(controls, "A")))
            return true
        return listed && QuickSwitch.LooksLikeFileDialog(winClass, controls, title)
    }

    ; 按窗口的控件判断 (和系统语言无关)。#32770 是系统对话框: 要有资源管理器的文件视图才算,
    ; 只有普通列表的 (例如 "系统信息" msinfo32 的主窗口, 有查找框、列表和按钮) 不算。
    ; 其它程序 (偏好设置里加的 "其它对话框") 有普通列表也算; 读不到控件的 (Qt 等) 按标题判断
    static LooksLikeFileDialog(winClass, controls, title := "") {
        if !controls.Length
            return RegExMatch(title, "i)(open|save|select|import|export|browse|打开|另存|选择|导入|导出)") > 0
        hasNameEdit := QuickSwitch._HasControl(controls, "^Edit\d+$")
        hasFileView := (winClass = "#32770")
            ? QuickSwitch._HasControl(controls, "^DirectUIHWND\d+$", "^SHELLDLL_DefView\d*$")
            : QuickSwitch._HasControl(controls, "^DirectUIHWND\d+$", "^SHELLDLL_DefView\d*$", "^SysListView32\d*$")
        hasButtons  := QuickSwitch._HasControl(controls, "^Button\d+$", "^ToolbarWindow32\d+$", "^ComboBoxEx32\d+$", "^ComboBox\d+$")
        return hasNameEdit && hasFileView && hasButtons
    }

    ; The old "Browse for Folder" dialog (SHBrowseForFolder): only a folder tree, no file list, so it is not one
    ; of the dialogs above. Told apart by fixed names, whatever the system language: the new style puts the tree
    ; in a "SHBrowseForFolder ShellNameSpace Control", the classic style has the tree itself with control ID 0x3741.
    ; Turned on by Works in → "Browse for Folder dialogs" (BrowseDialogs); ExcludeWindows applies too
    static IsBrowseWindow(hwnd) {
        controls := [], winClass := ""
        try controls := WinGetControls("ahk_id " hwnd), winClass := WinGetClass("ahk_id " hwnd)
        return QuickSwitch.LooksLikeBrowseDialog(winClass, controls, QuickSwitch._TreeId(controls, "ahk_id " hwnd))
    }

    static LooksLikeBrowseDialog(winClass, controls, treeId := 0) {
        return winClass = "#32770" && (treeId = 0x3741 || QuickSwitch._HasControl(controls, "^SHBrowseForFolder ShellNameSpace Control\d+$"))
    }

    ; Control ID of the dialog's first tree view (0 = none)
    static _TreeId(controls, win) {
        for name in controls
            if (name = "SysTreeView321")
                try return DllCall("GetDlgCtrlID", "Ptr", ControlGetHwnd(name, win), "Int")
        return 0
    }

    ; Select a folder in a "Browse for Folder" dialog's tree, like Listary: BFFM_SETSELECTIONW expands the tree to it.
    ; The message carries a pointer, so the path is first written into the dialog's own process. Doesn't press OK
    static SelectInBrowseDialog(folder, dialog) {
        folder := RTrim(folder, "\")
        if RegExMatch(folder, "^[A-Za-z]:$")
            folder .= "\"                                                   ; a drive needs its backslash
        pid := 0
        try pid := WinGetPID("ahk_id " dialog)
        if !pid || !(process := DllCall("OpenProcess", "UInt", 0x38, "Int", false, "UInt", pid, "Ptr"))   ; PROCESS_VM_OPERATION | VM_READ | VM_WRITE
            return false
        size := (StrLen(folder) + 1) * 2, ok := false
        remote := DllCall("VirtualAllocEx", "Ptr", process, "Ptr", 0, "UPtr", size, "UInt", 0x3000, "UInt", 0x04, "Ptr")   ; MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE
        if remote {
            DllCall("IsWow64Process", "Ptr", process, "Int*", &wow64 := 0)
            text := Buffer(size), StrPut(folder, text, "UTF-16")
            if (!(wow64 && remote > 0xFFFFFFFF)                             ; a 32-bit program only sees the low 4 GB
                && DllCall("WriteProcessMemory", "Ptr", process, "Ptr", remote, "Ptr", text, "UPtr", size, "Ptr", 0)) {
                try {
                    SendMessage(0x467, 1, remote, , "ahk_id " dialog, , , , 2000)   ; BFFM_SETSELECTIONW (WM_USER + 103); wParam TRUE = lParam is a path
                    ok := true
                }
            }
            DllCall("VirtualFreeEx", "Ptr", process, "Ptr", remote, "UPtr", 0, "UInt", 0x8000)   ; MEM_RELEASE
        }
        DllCall("CloseHandle", "Ptr", process)
        if !ok
            Logger.Error("QuickSwitch: cannot select " folder " in the Browse for Folder dialog")
        return ok
    }

    static _HasControl(controls, patterns*) {
        for name in controls
            for pattern in patterns
                if RegExMatch(name, "i)" pattern)
                    return true
        return false
    }

    ; The jump hotkey: the folder of the file manager used last (TC, Opus or Explorer, see FileManagerFolder)
    static SyncFileManagerPath(silent := false) {
        manager := QuickSwitch.FileManagerFolder()
        if !(IsObject(manager) && manager.Path != "")
            return silent ? "" : MsgBox(I18n.T("QuickSwitch.NoFileManager"), App.Name, 48)
        QuickSwitch.SetDialogPath(RTrim(manager.Path, "\") "\")             ; with a trailing backslash: AutoCAD and others need it
    }

    ; Total Commander 面板的文件夹: 用 TC 的 WM_COPYDATA 接口直接问 (TC 8.0 起支持), 不碰剪贴板。
    ; hwnd 为 0 时取最前面的 TC 窗口; otherPanel = true 时取另一个面板
    static TotalCmdFolder(hwnd := 0, otherPanel := false) {
        if !(hwnd := hwnd ? hwnd : WinExist("ahk_class TTOTAL_CMD"))
            return ""
        folder := RTrim(QuickSwitch.TotalCmdAsk(hwnd, otherPanel ? "TP" : "SP"), "\")
        return (folder ~= "^([A-Za-z]:|\\\\)") ? folder : ""               ; 只要真正的路径 (不要 FTP、压缩包里面等)
    }

    ; 发 WM_COPYDATA (dwData = "GW", 内容是 ANSI 的命令) 问 TC, TC 马上用 WM_COPYDATA (dwData = "RW", UTF-16)
    ; 回复给 wParam 里的窗口。命令: A = 当前是哪一边 (L / R); 两个字母: 第一个 L 左 / R 右 / S 当前 / T 另一个,
    ; 第二个 P 路径 / C 文件数 / I 光标位置 / N 光标下的文件名。TC 没有回复时返回 ""
    static _tcWaiting := false, _tcReply := "", _tcListening := false
    static TotalCmdAsk(hwnd, command, timeoutMs := 500) {
        if !QuickSwitch._tcListening {
            OnMessage(0x4A, (p*) => QuickSwitch._OnTotalCmdReply(p*))         ; WM_COPYDATA
            QuickSwitch._tcListening := true
        }
        request := Buffer(StrPut(command, "CP0"))
        StrPut(command, request, "CP0")
        copyData := Buffer(A_PtrSize * 3, 0)                                    ; COPYDATASTRUCT
        NumPut("UPtr", Ord("G") + 256 * Ord("W"), copyData, 0)
        NumPut("UInt", request.Size, copyData, A_PtrSize)
        NumPut("Ptr", request.Ptr, copyData, A_PtrSize * 2)
        QuickSwitch._tcReply := "", QuickSwitch._tcWaiting := true, accepted := 0
        DllCall("SendMessageTimeoutW", "Ptr", hwnd, "UInt", 0x4A, "Ptr", A_ScriptHwnd, "Ptr", copyData.Ptr
              , "UInt", 2, "UInt", timeoutMs, "UPtr*", &accepted)                 ; 2 = SMTO_ABORTIFHUNG: TC 卡住时不等
        QuickSwitch._tcWaiting := false                                         ; TC 在处理这条消息时就回复了, 不用另外等
        return QuickSwitch._tcReply
    }

    static _OnTotalCmdReply(wParam, lParam, msg, hwnd) {
        if (!QuickSwitch._tcWaiting || NumGet(lParam, 0, "UPtr") != Ord("R") + 256 * Ord("W"))
            return                                                              ; 不是在等的回复: 交给别的处理 (Everything)
        size := NumGet(lParam, A_PtrSize, "UInt"), data := NumGet(lParam, A_PtrSize * 2, "Ptr")
        reply := (data && size >= 2) ? StrGet(data, size // 2, "UTF-16") : ""  ; 到结尾的 0 为止
        QuickSwitch._tcReply := reply, QuickSwitch._tcWaiting := false
        return true
    }

    ;---------------------------------------------------------------------------
    ; 面板里的文件夹
    ;---------------------------------------------------------------------------
    ; [{Path, Group: "open" / "recent", Tag}]: 每个 TC 窗口的当前面板和另一个面板、每个资源管理器窗口、
    ; 最近的文件夹 (RecentFolders 个), 去掉重复的
    ; opus = false: without Directory Opus (the panel adds those a moment later, see _AddDOpusFolders)
    static MenuFolders(opus := true) {
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
        if opus                                                             ; the source tab of each lister, then the other side
            for state in [1, 2]
                for tab in QuickSwitch.DOpusTabs()
                    if (tab.State = state)
                        add(tab.Path, "open", I18n.T(state = 1 ? "QuickSwitch.TagOpus" : "QuickSwitch.TagOpusOther"))
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
            try FileGetShortcut(SubStr(line, InStr(line, "`t") + 1), &target)
            catch                                                           ; 损坏的快捷方式: 出错时 target 成了未赋值
                continue
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

    ; Current folder of the TC / Opus / Explorer window nearest the top: {Path, Tag, Hwnd}, "" when none.
    ; Windows are listed front to back, so with a dialog in front this is the file manager used last
    ; (the jump hotkey, auto jump, "Copy to ..." in the actions)
    static FileManagerFolder() {
        for hwnd in WinGetList() {
            className := ""
            try className := WinGetClass("ahk_id " hwnd)
            switch className {
                case "TTOTAL_CMD"   : return {Path: QuickSwitch.TotalCmdFolder(hwnd), Tag: "Total Commander", Hwnd: hwnd}
                case "dopus.lister" : return {Path: QuickSwitch.DOpusFolder(hwnd), Tag: "Directory Opus", Hwnd: hwnd}
                case "CabinetWClass": return {Path: QuickSwitch.ExplorerFolder(hwnd), Tag: I18n.T("QuickSwitch.TagExplorer"), Hwnd: hwnd}
            }
        }
        return ""
    }

    static IsFileManager(hwnd) {
        className := ""
        try className := WinGetClass("ahk_id " hwnd)
        return className = "TTOTAL_CMD" || className = "dopus.lister" || className = "CabinetWClass"
    }

    static FolderOfWindow(hwnd) {
        if (!hwnd || !WinExist("ahk_id " hwnd))
            return ""
        switch WinGetClass("ahk_id " hwnd) {
            case "TTOTAL_CMD"   : return QuickSwitch.TotalCmdFolder(hwnd)
            case "dopus.lister" : return QuickSwitch.DOpusFolder(hwnd)
            case "CabinetWClass": return QuickSwitch.ExplorerFolder(hwnd)
        }
        return ""
    }

    ;---------------------------------------------------------------------------
    ; Directory Opus: its own "dopusrt.exe /info <file>,paths" writes every lister's tabs to a small XML file
    ; (read only, no clipboard). It runs only when an Opus window is open and a folder is actually needed,
    ; and a result is reused for DOpusCacheMs
    ;---------------------------------------------------------------------------
    static DOpusCacheMs := 1500, _dopusTabs := [], _dopusAt := 0

    ; Source (focused) tab of a lister; hwnd 0 = the active lister
    static DOpusFolder(hwnd := 0) => QuickSwitch.PickDOpusFolder(QuickSwitch.DOpusTabs(), hwnd)

    static PickDOpusFolder(tabs, hwnd := 0) {
        for tab in tabs
            if (tab.State = 1 && (hwnd ? tab.Lister = hwnd : tab.ActiveLister))
                return tab.Path
        for tab in tabs                                                     ; lister not found (e.g. Opus just opened it): the active one
            if (tab.State = 1 && tab.ActiveLister)
                return tab.Path
        return ""
    }

    static DOpusTabs() {
        if !(lister := WinExist("ahk_class dopus.lister"))
            return []
        if (QuickSwitch._dopusAt && A_TickCount - QuickSwitch._dopusAt < QuickSwitch.DOpusCacheMs)
            return QuickSwitch._dopusTabs
        started := Logger.Ms(), tabs := []
        try {
            SplitPath(WinGetProcessPath("ahk_id " lister), , &dir)
            xmlFile := A_Temp "\ALTRun.dopus-paths.xml"
            try FileDelete(xmlFile)
            Run('"' dir '\dopusrt.exe" /info "' xmlFile '",paths', , "Hide", &pid)
            if ProcessWaitClose(pid, 3)                                     ; still running after 3 s (returns its PID): never hang the dialog on it
                try ProcessClose(pid)
            tabs := QuickSwitch.ParseDOpusPaths(FileRead(xmlFile, "UTF-8"))
            try FileDelete(xmlFile)
        } catch as e {
            Logger.Error("QuickSwitch: cannot read the Directory Opus folders - " e.Message)
        }
        QuickSwitch._dopusTabs := tabs, QuickSwitch._dopusAt := A_TickCount
        if Logger.Enabled
            Logger.Debug(Format("Perf: Directory Opus paths {} ms, {} tabs", Round(Logger.Ms() - started), tabs.Length))
        return tabs
    }

    ; <path active_lister="1" active_tab="1" tab_state="1" side="1" lister="0x70e2a" tab="0x50e1a" ...>C:\Folder</path>
    ; tab_state: 1 = the source (focused) tab of its lister, 2 = the tab shown on the other side of a dual lister
    static ParseDOpusPaths(xml) {
        tabs := [], pos := 1, attrs := ""
        value(name) => RegExMatch(attrs, "\b" name '="([^"]*)"', &a) ? a[1] : ""
        while (pos := RegExMatch(xml, "s)<path\b([^>]*)>(.*?)</path>", &m, pos)) {
            pos += m.Len
            attrs := m[1], folder := m[2]
            for entity in [["&lt;", "<"], ["&gt;", ">"], ["&quot;", '"'], ["&apos;", "'"], ["&amp;", "&"]]   ; &amp; last
                folder := StrReplace(folder, entity[1], entity[2])
            lister := value("lister"), state := value("tab_state")
            tabs.Push({Path: Trim(folder), Lister: IsInteger(lister) ? Integer(lister) : 0, State: IsInteger(state) ? Integer(state) : 0
                , ActiveLister: value("active_lister") = "1"})
        }
        return tabs
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
        if QuickSwitch.IsBrowseWindow(dialog)                               ; folder tree only: select it there, pressing Enter would mean OK
            return QuickSwitch.SelectInBrowseDialog(folder, dialog)
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
