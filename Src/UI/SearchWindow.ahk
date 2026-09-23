;===============================================================================
; SearchWindow.ahk - Alfred 式搜索窗口 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 界面: 无标题栏圆角窗口, 上面是大字号输入框, 下面是结果列表 (每行: 图标 +
; 标题 + 灰色副标题 + 右侧 Ctrl+N 提示), 窗口高度随结果数量伸缩。
;
; 结果列表用 ListView + NM_CUSTOMDRAW 自己画每一行; ListView 里只放当前可见
; 的几行 (最多 VisibleRows 行), 滚动由 Offset 控制, 所以永远没有滚动条。
;
; 键盘 (都在 _OnKeyDown 里处理, 通过 OnMessage 拦截 WM_KEYDOWN / WM_SYSKEYDOWN):
;   Enter / Ctrl+Enter / Alt+Enter   执行 / 显示位置或粘贴 / 复制   (见 ActionCatalog)
;   ↑ ↓  PgUp PgDn  Ctrl+P Ctrl+N    移动选择; 搜索框为空时 ↑ 调出历史搜索
;   Ctrl+1 ~ Ctrl+9                  直接执行可见的第 N 行
;   Tab                              自动补全
;   →  (光标在末尾时)                打开操作面板; ← / Esc 返回
;   空格 (搜索框为空时)              进入文件搜索模式 (只搜文件, 提示 "搜索文件..."); Backspace 返回
;   空格 (已输入文字, SpaceToRun)    执行选中项; Shift+空格输入空格
;   Ctrl+C (输入框没有选中文字时)    复制当前项
;   Ctrl+L                           大字显示
;   F3                               编辑当前项 (没有结果时: 用输入的文字新建自定义命令)
;   Ctrl+Del (光标在末尾时)          删除当前项 (确认后)
;   F2 / Ctrl+,                      偏好设置
;   F4                               用记事本编辑 ALTRun.json
;   Esc                              关闭操作面板 / 隐藏窗口
; 鼠标: 单击选择, 双击执行, 右键弹出这一项的操作菜单 (和操作面板相同)
;
; 用法:
;   SearchWindow.Create()             启动时调用一次
;   SearchWindow.Show([text])  Hide()  Toggle()
;===============================================================================

class SearchWindow {
    static Gui := "", Input := "", List := "", Separator := ""
    static Results := [], Selected := 0, Offset := 0
    static Mode := "results"                        ; results = 搜索结果; actions = 操作面板
    static FileMode := false                        ; 文件搜索模式: 空的搜索框里按空格进入, 只搜文件
    static ActionSource := "", SavedQuery := "", AllActions := []
    static HistoryIndex := 0
    static Width := 0, Padding := 0, InputHeight := 0, RowHeight := 0, IconSize := 0, VisibleRows := 8
    static SelectedRadius := 0
    static _gdi := Map()
    static _searchTimer := "", _hideTimer := ""
    static _posX := 0, _posY := 0
    static _keepOpen := false                        ; 右键菜单 / 删除确认期间不因失去焦点而隐藏

    ;---------------------------------------------------------------------------
    ; Create
    ;---------------------------------------------------------------------------
    static Create() {
        SearchWindow.Width       := Win.Scale(AppSettings.Appearance["Width"])
        SearchWindow.VisibleRows := Max(1, Min(AppSettings.Appearance["VisibleRows"], 9))
        SearchWindow.Padding     := Win.Scale(ThemeManager.Get("Padding"))
        SearchWindow.RowHeight   := Win.Scale(ThemeManager.Get("RowHeight"))
        SearchWindow.IconSize    := Win.Scale(ThemeManager.Get("IconSize"))
        SearchWindow.SelectedRadius := Win.Scale(ThemeManager.Get("SelectedRadius"))
        IconCache.Size := SearchWindow.IconSize
        IconCache.OnLoaded := () => SearchWindow._Repaint()
        SearchWindow._CreateGdiObjects()

        w := SearchWindow.Width, pad := SearchWindow.Padding
        background := ThemeManager.Get("Background")

        g := Gui("-Caption +AlwaysOnTop +ToolWindow -DPIScale" (SearchWindow._IsWindows11() ? "" : " +Border"), App.Name)
        g.BackColor := background
        g.MarginX := 0, g.MarginY := 0
        g.SetFont("s" ThemeManager.Get("InputFontSize") " c" ThemeManager.Get("InputText"), ThemeManager.FontName())

        searchBox := g.AddEdit("x" pad " y" pad " w" (w - 2 * pad) " r1 -Multi -E0x200 -VScroll -WantReturn Background" background)
        searchBox.GetPos(, , , &editH)
        SearchWindow.InputHeight := Max(editH, Win.Scale(40))
        searchBox.Move(pad, pad + (SearchWindow.InputHeight - editH) // 2)
        searchBox.OnEvent("Change", (*) => SearchWindow._OnInputChange())
        Win.SetCueBanner(searchBox.Hwnd, I18n.T("Search.Placeholder"))

        listTop := pad + SearchWindow.InputHeight + pad // 2
        separator := g.AddText("x0 y" listTop " w" w " h1 Hidden Background" ThemeManager.Get("Separator"))

        ; 0x2000 = LVS_NOSCROLL; LV0x10000 = 双缓冲, 避免闪烁
        list := g.AddListView("x0 y" (listTop + 1) " w" w " h0 Hidden -Hdr -Multi -E0x200 -TabStop +0x2000 +LV0x10000 Background" background, ["Result"])
        list.ModifyCol(1, w)
        rowImageList := DllCall("comctl32\ImageList_Create", "Int", 1, "Int", SearchWindow.RowHeight, "UInt", 0x21, "Int", 1, "Int", 1, "Ptr")
        list.SetImageList(rowImageList, 1)                                  ; 空白小图标列表只用来撑开行高
        list.OnNotify(-12, (ctrl, lParam) => SearchWindow._OnCustomDraw(lParam))    ; NM_CUSTOMDRAW
        list.OnEvent("Click", (ctrl, row) => SearchWindow._OnRowClick(row))
        list.OnEvent("DoubleClick", (ctrl, row) => SearchWindow._OnRowDoubleClick(row))
        list.OnEvent("ContextMenu", (ctrl, row, isRightClick, x, y) => SearchWindow._OnRowContextMenu(row, x, y))

        g.OnEvent("Close", (*) => SearchWindow.Hide())

        SearchWindow.Gui := g, SearchWindow.Input := searchBox, SearchWindow.List := list, SearchWindow.Separator := separator

        OnMessage(0x100, (p*) => SearchWindow._OnKeyDown(p*))               ; WM_KEYDOWN
        OnMessage(0x104, (p*) => SearchWindow._OnKeyDown(p*))               ; WM_SYSKEYDOWN (Alt+...)
        OnMessage(0x6,   (p*) => SearchWindow._OnActivate(p*))              ; WM_ACTIVATE
        OnMessage(0x20A, (p*) => SearchWindow._OnMouseWheel(p*))            ; WM_MOUSEWHEEL

        g.Show("Hide w" w " h" SearchWindow._WindowHeight(0))
        Win.SetCorner(g.Hwnd)
        Win.SetBorderColor(g.Hwnd, ThemeManager.Get("Border"))
        opacity := ThemeManager.Get("Opacity")
        if (IsNumber(opacity) && opacity >= 1 && opacity < 255) {
            DetectHiddenWindows(true)                                       ; 窗口此时还是隐藏的
            WinSetTransparent(Round(opacity), g.Hwnd)
        }
    }

    static _CreateGdiObjects() {
        font := ThemeManager.FontName()
        gdi := SearchWindow._gdi
        gdi["TitleFont"]    := SearchWindow._CreateFont(font, ThemeManager.Get("TitleFontSize"), 400)
        gdi["SubtitleFont"] := SearchWindow._CreateFont(font, ThemeManager.Get("SubtitleFontSize"), 400)
        gdi["ShortcutFont"] := SearchWindow._CreateFont(font, ThemeManager.Get("ShortcutFontSize"), 400)
        gdi["Background"]   := DllCall("CreateSolidBrush", "UInt", Win.ColorToBgr(ThemeManager.Get("Background")), "Ptr")
        gdi["Selected"]     := DllCall("CreateSolidBrush", "UInt", Win.ColorToBgr(ThemeManager.Get("SelectedBackground")), "Ptr")
        for key in ["Title", "Subtitle", "Shortcut", "SelectedTitle", "SelectedSubtitle", "SelectedShortcut"]
            gdi[key "Color"] := Win.ColorToBgr(ThemeManager.Get(key))
    }

    static _CreateFont(name, pointSize, weight) {
        height := -Round(pointSize * A_ScreenDPI / 72)
        return DllCall("CreateFontW", "Int", height, "Int", 0, "Int", 0, "Int", 0, "Int", weight
            , "UInt", 0, "UInt", 0, "UInt", 0, "UInt", 1, "UInt", 0, "UInt", 0, "UInt", 5, "UInt", 0   ; 1 = DEFAULT_CHARSET, 5 = CLEARTYPE_QUALITY
            , "WStr", name, "Ptr")
    }

    static _IsWindows11() {
        return VerCompare(A_OSVersion, "10.0.22000") >= 0
    }

    ;---------------------------------------------------------------------------
    ; Show / Hide
    ;---------------------------------------------------------------------------
    static IsVisible() {
        return IsObject(SearchWindow.Gui) && DllCall("IsWindowVisible", "Ptr", SearchWindow.Gui.Hwnd)
    }

    static Show(text := "") {
        if !IsObject(SearchWindow.Gui)
            return
        App.RememberActiveWindow()
        SearchWindow.Mode := "results"
        SearchWindow.FileMode := false
        SearchWindow._UpdateCueBanner()
        SearchWindow.HistoryIndex := 0
        SearchWindow._SetInput(text)

        area := Win.WorkAreaAtMouse()
        SearchWindow._posX := area.Left + (area.Right - area.Left - SearchWindow.Width) // 2
        SearchWindow._posY := area.Top + Round((area.Bottom - area.Top) * 0.2)
        SearchWindow.Gui.Show("x" SearchWindow._posX " y" SearchWindow._posY " w" SearchWindow.Width " h" SearchWindow._WindowHeight(SearchWindow._VisibleCount()))
        try WinActivate("ahk_id " SearchWindow.Gui.Hwnd)
        SearchWindow.Input.Focus()
        len := StrLen(SearchWindow.Input.Value)
        SendMessage(0xB1, len, len, SearchWindow.Input.Hwnd)                ; 获得焦点时 Edit 会全选, 把光标放回末尾
        if AppSettings.General["SwitchToEnglishInput"]
            Win.SwitchToEnglishIME()
    }

    static Hide() {
        if IsObject(SearchWindow.Gui)
            SearchWindow.Gui.Hide()
        LargeType.Close()
    }

    static Toggle() {
        if (SearchWindow.IsVisible() && WinActive("ahk_id " SearchWindow.Gui.Hwnd))
            SearchWindow.Hide()
        else
            SearchWindow.Show()
    }

    ;---------------------------------------------------------------------------
    ; Search / results
    ;---------------------------------------------------------------------------
    static _SetInput(text) {
        SearchWindow.Input.Value := text
        len := StrLen(text)
        SendMessage(0xB1, len, len, SearchWindow.Input.Hwnd)                ; EM_SETSEL: 光标放到末尾
        SearchWindow._RunSearch()
    }

    static _OnInputChange() {
        text := SearchWindow.Input.Value
        if RegExMatch(text, "[`r`n`t]") {                                   ; 粘贴了多行文字: 合并成一行, 搜索框始终只有一行
            SearchWindow.Input.Value := RegExReplace(text, "\s*[`r`n`t]+\s*", " ")
            len := StrLen(SearchWindow.Input.Value)
            SendMessage(0xB1, len, len, SearchWindow.Input.Hwnd)
        }
        SearchWindow.HistoryIndex := 0
        if (SearchWindow._searchTimer = "")
            SearchWindow._searchTimer := () => SearchWindow._RunSearch()
        SetTimer(SearchWindow._searchTimer, -1)                             ; 连续输入时只搜索最后一次
    }

    static _RunSearch() {
        text := SearchWindow.Input.Value
        if (SearchWindow.Mode = "actions") {
            filtered := []
            for action in SearchWindow.AllActions
                if (Trim(text) = "" || FuzzyMatcher.Score(text, action.Title) > 0)
                    filtered.Push(action)
            return SearchWindow.SetResults(filtered)
        }
        if SearchWindow.FileMode
            return SearchWindow.SetResults(ProviderRegistry.SearchFiles(text))
        SearchWindow.SetResults(ProviderRegistry.Search(text))
    }

    static SetResults(results) {
        SearchWindow.Results  := results
        SearchWindow.Selected := results.Length ? 1 : 0
        SearchWindow.Offset   := 0
        SearchWindow._Layout()
    }

    static _VisibleCount() {
        return Min(SearchWindow.Results.Length, SearchWindow.VisibleRows)
    }

    static _WindowHeight(visibleCount) {
        pad := SearchWindow.Padding
        if !visibleCount
            return pad + SearchWindow.InputHeight + pad
        ; 输入框 + 间距 + 分隔线 + 结果行 + 底部留白 (和 Create() 里的布局一致)
        return pad + SearchWindow.InputHeight + pad // 2 + 1 + visibleCount * SearchWindow.RowHeight + pad // 2
    }

    static _Layout() {
        visibleCount := SearchWindow._VisibleCount()
        list := SearchWindow.List
        list.Opt("-Redraw")
        if (list.GetCount() != visibleCount) {
            list.Delete()
            Loop visibleCount
                list.Add("", "")
        }
        list.Move(, , , visibleCount * SearchWindow.RowHeight)
        list.Visible := visibleCount > 0
        SearchWindow.Separator.Visible := visibleCount > 0
        list.Opt("+Redraw")
        SearchWindow._Repaint()
        if SearchWindow.IsVisible()
            SearchWindow.Gui.Show("NA w" SearchWindow.Width " h" SearchWindow._WindowHeight(visibleCount))
    }

    static _Repaint() {
        DllCall("InvalidateRect", "Ptr", SearchWindow.List.Hwnd, "Ptr", 0, "Int", 1)
    }

    static SelectedItem() {
        if (SearchWindow.Selected < 1 || SearchWindow.Selected > SearchWindow.Results.Length)
            return ""
        return SearchWindow.Results[SearchWindow.Selected]
    }

    static MoveSelection(step) {
        count := SearchWindow.Results.Length
        if !count
            return
        SearchWindow.Selected := Max(1, Min(count, SearchWindow.Selected + step))
        visible := SearchWindow.VisibleRows
        if (SearchWindow.Selected <= SearchWindow.Offset)
            SearchWindow.Offset := SearchWindow.Selected - 1
        else if (SearchWindow.Selected > SearchWindow.Offset + visible)
            SearchWindow.Offset := SearchWindow.Selected - visible
        SearchWindow._Repaint()
    }

    ;---------------------------------------------------------------------------
    ; Execute
    ;---------------------------------------------------------------------------
    static _Execute(modifier := "") {
        item := SearchWindow.SelectedItem()
        if !IsObject(item)
            return

        if (SearchWindow.Mode = "actions") {
            source := SearchWindow.ActionSource
            Knowledge.Record(SearchWindow.SavedQuery, source.Uid)
            SearchWindow.Hide()
            SearchWindow._SafeRun(() => item.OnRun.Call(source))
            return
        }

        if !item.Valid {
            if (item.AutoComplete != "")
                SearchWindow._SetInput(item.AutoComplete)
            return
        }
        Knowledge.Record(SearchWindow.Input.Value, item.Uid)
        SearchWindow.Hide()
        if (modifier = "")
            SearchWindow._SafeRun(() => ActionCatalog.RunDefault(item))
        else
            SearchWindow._SafeRun(() => ActionCatalog.RunModifier(item, modifier))
    }

    static _SafeRun(fn) {
        try {
            fn()
        } catch as e {
            Logger.Error("SearchWindow: action failed - " e.Message)
            MsgBox(e.Message, App.Name, 48)
        }
    }

    static _RunVisibleRow(row) {
        if (row > SearchWindow._VisibleCount())
            return
        SearchWindow.Selected := SearchWindow.Offset + row
        SearchWindow._Execute()
    }

    ;---------------------------------------------------------------------------
    ; Action panel
    ;---------------------------------------------------------------------------
    static _OpenActions() {
        item := SearchWindow.SelectedItem()
        if (!IsObject(item) || !item.Valid || SearchWindow.Mode = "actions")
            return false
        actions := ActionCatalog.ListFor(item)
        if !actions.Length
            return false
        SearchWindow.SavedQuery := SearchWindow.Input.Value
        SearchWindow.ActionSource := item
        SearchWindow.AllActions := actions
        SearchWindow.Mode := "actions"
        Win.SetCueBanner(SearchWindow.Input.Hwnd, I18n.T("Search.ActionsFor", item.Title))
        SearchWindow._SetInput("")
        return true
    }

    static _CloseActions() {
        SearchWindow.Mode := "results"
        SearchWindow._UpdateCueBanner()
        SearchWindow._SetInput(SearchWindow.SavedQuery)
    }

    ;---------------------------------------------------------------------------
    ; File search mode (空的搜索框里按空格, 和 Alfred 一样)
    ;---------------------------------------------------------------------------
    static _CanEnterFileMode() {
        return SearchWindow.Mode = "results" && !SearchWindow.FileMode && SearchWindow.Input.Value = ""
            && AppSettings.Feature("FileSearch")["SpacePrefix"] && ProviderRegistry.IsEnabled(FileSearchProvider)
    }

    static _SetFileMode(enabled) {
        SearchWindow.FileMode := enabled
        SearchWindow._UpdateCueBanner()
        SearchWindow._RunSearch()
    }

    ; 灰色提示文字: 普通搜索 / 文件搜索模式 (操作面板的提示在 _OpenActions 里设置)
    static _UpdateCueBanner() {
        Win.SetCueBanner(SearchWindow.Input.Hwnd, I18n.T(SearchWindow.FileMode ? "Search.FilesPlaceholder" : "Search.Placeholder"))
    }

    ;---------------------------------------------------------------------------
    ; Edit / delete (F3, Ctrl+Del, 右键菜单, 操作面板)
    ;---------------------------------------------------------------------------
    ; 当前操作的是哪一项: 操作面板里是打开面板的那一项, 否则是选中的结果
    static _TargetItem() {
        return (SearchWindow.Mode = "actions") ? SearchWindow.ActionSource : SearchWindow.SelectedItem()
    }

    static _CurrentQuery() {
        return (SearchWindow.Mode = "actions") ? SearchWindow.SavedQuery : SearchWindow.Input.Value
    }

    ; 按键消息处理完之后再弹对话框, 不在 OnMessage 回调里等待
    static _Later(fn) {
        SetTimer(fn, -1)
    }

    static _EditSelected() {
        item := SearchWindow._TargetItem()
        isFallback := IsObject(item) && item.Provider = ""                  ; 没有匹配时的兜底项 (网页搜索等)
        if (IsObject(item) && !isFallback && ActionCatalog.CanEdit(item))
            return SearchWindow._Later(() => SearchWindow.EditItem(item))
        query := Trim(SearchWindow._CurrentQuery())
        if (query != "" && (!IsObject(item) || isFallback))                 ; 没有结果: 用输入的文字新建一条命令
            return SearchWindow._Later(() => SearchWindow.NewCommand(query))
        App.Notify(I18n.T("Search.NotEditable"))
    }

    ; 编辑对话框关闭后回到原来的搜索, 马上能看到修改的结果
    static EditItem(item) {
        query := SearchWindow._CurrentQuery()
        SearchWindow.Hide()
        SearchWindow._SafeRun(() => ActionCatalog.EditItem(item))
        SearchWindow._Reopen(query)
    }

    static NewCommand(title) {
        SearchWindow.Hide()
        SearchWindow._SafeRun(() => CustomCommandProvider.Edit("", Map("Title", title)))
        SearchWindow._Reopen(title)
    }

    static DeleteItem(item) {
        query := SearchWindow._CurrentQuery()
        wasVisible := SearchWindow.IsVisible()
        SearchWindow._keepOpen := true
        answer := MsgBox(ActionCatalog.DeletePrompt(item), App.Name, "YesNo Icon! Default2" (wasVisible ? " Owner" SearchWindow.Gui.Hwnd : ""))
        SearchWindow._keepOpen := false
        if (answer = "Yes")
            SearchWindow._SafeRun(() => ActionCatalog.DeleteItem(item))
        selected := SearchWindow.Selected
        SearchWindow._Reopen(query)
        if (answer = "Yes" && selected > 1)
            SearchWindow.MoveSelection(selected - 2)                        ; 选中被删除那一行的上一行 (MoveSelection 相对第 1 行)
    }

    ; 重新显示搜索窗口, 不改变 "呼出前的窗口" (粘贴等操作还是回到那个窗口)
    static _Reopen(query) {
        previous := App.PreviousWindow
        SearchWindow.Show(query)
        App.PreviousWindow := previous
    }

    static _OnRowContextMenu(row, x, y) {
        if (row < 1 || SearchWindow.Offset + row > SearchWindow.Results.Length)
            return
        if (SearchWindow.Mode = "results") {
            SearchWindow.Selected := SearchWindow.Offset + row
            SearchWindow._Repaint()
        }
        item := SearchWindow._TargetItem()
        if (!IsObject(item) || !item.Valid && !ActionCatalog.CanEdit(item))
            return
        if (SearchWindow.Mode = "actions")                                  ; 操作面板里右键 = 执行这一项
            return SearchWindow._Execute()
        contextMenu := Menu()
        for action in ActionCatalog.ListFor(item) {
            label := StrReplace(action.Title, "&", "&&") (action.Subtitle != "" ? "`t" action.Subtitle : "")
            contextMenu.Add(label, SearchWindow._MenuHandler(action, item))
            if (A_Index = 1)
                contextMenu.Default := label
        }
        ; 菜单显示期间 AHK 来不及处理 NM_CUSTOMDRAW, 列表会按默认样式画 (蓝色选中条):
        ; 先去掉 ListView 自己的选中状态, 并在弹出菜单前按我们的样式画好
        Critical("Off")
        SearchWindow.List.Modify(0, "-Select -Focus")
        SearchWindow._Repaint()
        DllCall("UpdateWindow", "Ptr", SearchWindow.List.Hwnd)
        SearchWindow._keepOpen := true
        contextMenu.Show(x, y)
        SearchWindow._keepOpen := false
        if (SearchWindow.IsVisible() && !WinActive("ahk_id " SearchWindow.Gui.Hwnd))
            try WinActivate("ahk_id " SearchWindow.Gui.Hwnd)
    }

    ; 单独一个方法生成闭包, 每个菜单项记住自己的操作
    static _MenuHandler(action, item) {
        return (*) => SearchWindow._RunMenuAction(action, item)
    }

    static _RunMenuAction(action, item) {
        SearchWindow._keepOpen := false
        Knowledge.Record(SearchWindow.Input.Value, item.Uid)
        SearchWindow.SavedQuery := SearchWindow.Input.Value
        SearchWindow.Hide()
        SearchWindow._SafeRun(() => action.OnRun.Call(item))
    }

    ;---------------------------------------------------------------------------
    ; Keyboard / mouse
    ;---------------------------------------------------------------------------
    static _IsOwnWindow(hwnd) {
        return IsObject(SearchWindow.Gui) && (hwnd = SearchWindow.Gui.Hwnd || hwnd = SearchWindow.Input.Hwnd || hwnd = SearchWindow.List.Hwnd)
    }

    ; 返回 0 表示这个按键已经处理, 不再交给输入框; 返回空 = 按正常方式处理
    static _OnKeyDown(wParam, lParam, msg, hwnd) {
        if !SearchWindow._IsOwnWindow(hwnd)
            return
        vk := wParam
        ctrl  := GetKeyState("Ctrl")
        shift := GetKeyState("Shift")
        alt   := (msg = 0x104) || GetKeyState("Alt")
        actions := SearchWindow.Mode = "actions"

        switch vk {
            case 0x0D:                                                      ; Enter
                SearchWindow._Execute(ctrl ? "ctrl" : alt ? "alt" : "")
                return 0
            case 0x1B:                                                      ; Esc
                if actions
                    SearchWindow._CloseActions()
                else
                    SearchWindow.Hide()
                return 0
            case 0x20:                                                      ; 空格: 空的搜索框里进入文件搜索模式
                if (ctrl || alt)
                    return
                if (!shift && SearchWindow._CanEnterFileMode()) {
                    SearchWindow._SetFileMode(true)
                    return 0
                }
                ; SpaceToRun (2.x 的同名选项): 已经输入了文字时空格执行选中项, Shift+空格照常输入空格
                if (!shift && AppSettings.General["SpaceToRun"] && SearchWindow.Input.Value != "" && IsObject(SearchWindow.SelectedItem())) {
                    SearchWindow._Execute()
                    return 0
                }
                return
            case 0x26:                                                      ; ↑
                if (!actions && !SearchWindow.FileMode && SearchWindow._RecallHistory())
                    return 0
                SearchWindow.MoveSelection(-1)
                return 0
            case 0x28:                                                      ; ↓
                SearchWindow.MoveSelection(1)
                return 0
            case 0x21:                                                      ; PgUp
                SearchWindow.MoveSelection(-SearchWindow.VisibleRows)
                return 0
            case 0x22:                                                      ; PgDn
                SearchWindow.MoveSelection(SearchWindow.VisibleRows)
                return 0
            case 0x09:                                                      ; Tab
                SearchWindow._AutoComplete()
                return 0
            case 0x27:                                                      ; →
                if (SearchWindow._CaretAtEnd() && SearchWindow._OpenActions())
                    return 0
                return
            case 0x25:                                                      ; ←
                if (actions && SearchWindow.Input.Value = "") {
                    SearchWindow._CloseActions()
                    return 0
                }
                return
            case 0x72:                                                      ; F3
                SearchWindow._EditSelected()
                return 0
            case 0x71:                                                      ; F2
                SearchWindow.Hide()
                App.OpenPreferences()
                return 0
            case 0x73:                                                      ; F4
                SearchWindow.Hide()
                App.EditSettingsFile()
                return 0
            case 0x2E:                                                      ; Ctrl+Del
                if (ctrl && SearchWindow._CaretAtEnd() && IsObject(item := SearchWindow._TargetItem()) && ActionCatalog.CanDelete(item)) {
                    SearchWindow._Later(() => SearchWindow.DeleteItem(item))
                    return 0
                }
                return
            case 0x08:                                                      ; Backspace
                if (actions && SearchWindow.Input.Value = "") {
                    SearchWindow._CloseActions()
                    return 0
                }
                if (SearchWindow.FileMode && SearchWindow.Input.Value = "") {  ; 文件搜索模式下删空后再按: 回到普通搜索
                    SearchWindow._SetFileMode(false)
                    return 0
                }
                if ctrl {                                                   ; Ctrl+Backspace: 删除前一个词 (普通 Edit 控件不支持)
                    SearchWindow._DeletePreviousWord()
                    return 0
                }
                return
        }

        if !ctrl
            return
        switch vk {
            case 0x4E:                                                      ; Ctrl+N
                SearchWindow.MoveSelection(1)
                return 0
            case 0x50:                                                      ; Ctrl+P
                SearchWindow.MoveSelection(-1)
                return 0
            case 0x43:                                                      ; Ctrl+C
                if (shift || SearchWindow._HasTextSelection())
                    return
                SearchWindow._CopySelected()
                return 0
            case 0x4C:                                                      ; Ctrl+L
                SearchWindow._ShowLargeType()
                return 0
            case 0xBC:                                                      ; Ctrl+,
                SearchWindow.Hide()
                App.OpenPreferences()
                return 0
        }
        if (vk >= 0x31 && vk <= 0x39) {                                     ; Ctrl+1..9
            SearchWindow._RunVisibleRow(vk - 0x30)
            return 0
        }
    }

    static _RecallHistory() {
        history := Knowledge.History
        text := SearchWindow.Input.Value
        if (!history.Length)
            return false
        if (SearchWindow.HistoryIndex = 0 && text != "")
            return false
        if (SearchWindow.HistoryIndex > 0 && (SearchWindow.HistoryIndex > history.Length || text != history[SearchWindow.HistoryIndex]))
            return false
        if (SearchWindow.HistoryIndex >= history.Length)
            return true
        SearchWindow.HistoryIndex += 1
        index := SearchWindow.HistoryIndex
        SearchWindow._SetInput(history[index])
        SearchWindow.HistoryIndex := index                                  ; _SetInput 不经过 Change 事件, 这里保持历史位置
        return true
    }

    static _AutoComplete() {
        item := SearchWindow.SelectedItem()
        if (!IsObject(item) || SearchWindow.Mode = "actions")
            return
        text := (item.AutoComplete != "") ? item.AutoComplete : item.Title
        SearchWindow._SetInput(text)
    }

    static _CopySelected() {
        item := SearchWindow.SelectedItem()
        if !IsObject(item)
            return
        SearchWindow.Hide()
        ActionCatalog.CopyText(item.CopyText())
    }

    static _ShowLargeType() {
        item := SearchWindow.SelectedItem()
        text := IsObject(item) ? item.DisplayText() : SearchWindow.Input.Value
        SearchWindow.Hide()
        LargeType.Show(text)
    }

    static _CaretAtEnd() {
        selection := SendMessage(0xB0, 0, 0, SearchWindow.Input.Hwnd)        ; EM_GETSEL
        return ((selection >> 16) & 0xFFFF) >= StrLen(SearchWindow.Input.Value)
    }

    static _HasTextSelection() {
        selection := SendMessage(0xB0, 0, 0, SearchWindow.Input.Hwnd)
        return (selection & 0xFFFF) != ((selection >> 16) & 0xFFFF)
    }

    static _DeletePreviousWord() {
        text := SearchWindow.Input.Value
        caret := (SendMessage(0xB0, 0, 0, SearchWindow.Input.Hwnd) >> 16) & 0xFFFF
        before := RegExReplace(SubStr(text, 1, caret), "\S*\s*$")
        SearchWindow.Input.Value := before SubStr(text, caret + 1)
        len := StrLen(before)
        SendMessage(0xB1, len, len, SearchWindow.Input.Hwnd)
        SearchWindow._OnInputChange()
    }

    static _OnRowClick(row) {
        if (row >= 1 && SearchWindow.Offset + row <= SearchWindow.Results.Length) {
            SearchWindow.Selected := SearchWindow.Offset + row
            SearchWindow._Repaint()
        }
        SearchWindow.Input.Focus()
    }

    static _OnRowDoubleClick(row) {
        if (row >= 1 && SearchWindow.Offset + row <= SearchWindow.Results.Length) {
            SearchWindow.Selected := SearchWindow.Offset + row
            SearchWindow._Execute()
        }
    }

    static _OnMouseWheel(wParam, lParam, msg, hwnd) {
        if !SearchWindow._IsOwnWindow(hwnd)
            return
        delta := (wParam >> 16) & 0xFFFF
        SearchWindow.MoveSelection(delta >= 0x8000 ? 1 : -1)                ; 高位是有符号的滚动量
        return 0
    }

    static _OnActivate(wParam, lParam, msg, hwnd) {
        if (!IsObject(SearchWindow.Gui) || hwnd != SearchWindow.Gui.Hwnd)
            return
        if ((wParam & 0xFFFF) = 0 && AppSettings.General["HideOnDeactivate"]) {
            if (SearchWindow._hideTimer = "")
                SearchWindow._hideTimer := () => SearchWindow._HideIfInactive()
            SetTimer(SearchWindow._hideTimer, -50)
        }
    }

    static _HideIfInactive() {
        if SearchWindow._keepOpen
            return
        if (SearchWindow.IsVisible() && !WinActive("ahk_id " SearchWindow.Gui.Hwnd))
            SearchWindow.Hide()
    }

    ;---------------------------------------------------------------------------
    ; Painting (NM_CUSTOMDRAW)
    ;---------------------------------------------------------------------------
    static _OnCustomDraw(lParam) {
        static CDDS_PREPAINT := 0x1, CDDS_ITEMPREPAINT := 0x10001
        static CDRF_DODEFAULT := 0x0, CDRF_SKIPDEFAULT := 0x4, CDRF_NOTIFYITEMDRAW := 0x20
        stage := NumGet(lParam, A_PtrSize * 3, "UInt")
        if (stage = CDDS_PREPAINT)
            return CDRF_NOTIFYITEMDRAW
        if (stage != CDDS_ITEMPREPAINT)
            return CDRF_DODEFAULT

        hdc := NumGet(lParam, A_PtrSize = 8 ? 32 : 16, "Ptr")
        row := NumGet(lParam, A_PtrSize = 8 ? 56 : 36, "UPtr")             ; dwItemSpec, 0 起
        index := SearchWindow.Offset + row + 1
        if (index > SearchWindow.Results.Length)
            return CDRF_SKIPDEFAULT

        rect := Buffer(16, 0)                                               ; left = LVIR_BOUNDS (0)
        SendMessage(0x100E, row, rect.Ptr, SearchWindow.List.Hwnd)          ; LVM_GETITEMRECT
        SearchWindow._PaintRow(hdc, rect, SearchWindow.Results[index], index = SearchWindow.Selected, row + 1)
        return CDRF_SKIPDEFAULT
    }

    static _PaintRow(hdc, rect, item, selected, visibleRow) {
        static DT_RIGHT := 0x2, DT_VCENTER := 0x4, DT_BOTTOM := 0x8, DT_SINGLELINE := 0x20
        static DT_NOPREFIX := 0x800, DT_PATH_ELLIPSIS := 0x4000, DT_END_ELLIPSIS := 0x8000
        gdi := SearchWindow._gdi
        left := NumGet(rect, 0, "Int"), top := NumGet(rect, 4, "Int")
        right := NumGet(rect, 8, "Int"), bottom := NumGet(rect, 12, "Int")
        pad := SearchWindow.Padding, iconSize := SearchWindow.IconSize, rowH := bottom - top

        if (selected && SearchWindow.SelectedRadius > 0) {                 ; 圆角选中: 左右留一点边距, 画圆角矩形
            DllCall("FillRect", "Ptr", hdc, "Ptr", rect, "Ptr", gdi["Background"])
            inset := Win.Scale(6), gap := Win.Scale(2), diameter := SearchWindow.SelectedRadius * 2
            oldBrush := DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["Selected"], "Ptr")
            oldPen := DllCall("SelectObject", "Ptr", hdc, "Ptr", DllCall("GetStockObject", "Int", 8, "Ptr"), "Ptr")   ; NULL_PEN
            DllCall("RoundRect", "Ptr", hdc, "Int", left + inset, "Int", top + gap, "Int", right - inset + 1, "Int", bottom - gap + 1, "Int", diameter, "Int", diameter)
            DllCall("SelectObject", "Ptr", hdc, "Ptr", oldPen)
            DllCall("SelectObject", "Ptr", hdc, "Ptr", oldBrush)
        } else {
            DllCall("FillRect", "Ptr", hdc, "Ptr", rect, "Ptr", selected ? gdi["Selected"] : gdi["Background"])
        }
        DllCall("SetBkMode", "Ptr", hdc, "Int", 1)                          ; TRANSPARENT

        if (hIcon := IconCache.Get(item.Icon))
            DllCall("DrawIconEx", "Ptr", hdc, "Int", left + pad, "Int", top + (rowH - iconSize) // 2, "Ptr", hIcon, "Int", iconSize, "Int", iconSize, "UInt", 0, "Ptr", 0, "UInt", 3)

        textLeft := left + pad + iconSize + Win.Scale(12)
        shortcutW := Win.Scale(56)
        oldFont := DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["ShortcutFont"], "Ptr")

        if (visibleRow <= 9 && SearchWindow.Mode = "results") {             ; 右侧 Ctrl+N 提示
            DllCall("SetTextColor", "Ptr", hdc, "UInt", selected ? gdi["SelectedShortcutColor"] : gdi["ShortcutColor"])
            SearchWindow._DrawText(hdc, "Ctrl+" visibleRow, right - pad - shortcutW, top, right - pad, bottom, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX)
        } else if (item.Subtitle != "" && SearchWindow.Mode = "actions") {  ; 操作面板: 右侧显示快捷键
            DllCall("SetTextColor", "Ptr", hdc, "UInt", selected ? gdi["SelectedShortcutColor"] : gdi["ShortcutColor"])
            SearchWindow._DrawText(hdc, item.Subtitle, right - pad - Win.Scale(100), top, right - pad, bottom, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX)
        }
        textRight := right - pad - (SearchWindow.Mode = "actions" ? Win.Scale(108) : shortcutW + Win.Scale(8))

        DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["TitleFont"])
        DllCall("SetTextColor", "Ptr", hdc, "UInt", selected ? gdi["SelectedTitleColor"] : gdi["TitleColor"])
        hasSubtitle := (item.Subtitle != "" && SearchWindow.Mode = "results")
        middle := top + Round(rowH * 0.54)
        if hasSubtitle
            SearchWindow._DrawText(hdc, item.Title, textLeft, top, textRight, middle, DT_BOTTOM | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS)
        else
            SearchWindow._DrawText(hdc, item.Title, textLeft, top, textRight, bottom, DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS)

        if hasSubtitle {
            DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["SubtitleFont"])
            DllCall("SetTextColor", "Ptr", hdc, "UInt", selected ? gdi["SelectedSubtitleColor"] : gdi["SubtitleColor"])
            ellipsis := (item.Kind = "file" || item.Kind = "folder") ? DT_PATH_ELLIPSIS : DT_END_ELLIPSIS
            SearchWindow._DrawText(hdc, item.Subtitle, textLeft, middle + Win.Scale(2), textRight, bottom, DT_SINGLELINE | DT_NOPREFIX | ellipsis)
        }
        DllCall("SelectObject", "Ptr", hdc, "Ptr", oldFont)
    }

    static _DrawText(hdc, text, left, top, right, bottom, flags) {
        rect := Buffer(16)
        NumPut("Int", left, "Int", top, "Int", right, "Int", bottom, rect)
        DllCall("DrawTextW", "Ptr", hdc, "WStr", text, "Int", -1, "Ptr", rect, "UInt", flags)
    }
}
