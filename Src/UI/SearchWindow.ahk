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
;   Ctrl+C (输入框没有选中文字时)    复制当前项
;   Ctrl+L                           大字显示
;   Ctrl+,                           偏好设置
;   Esc                              关闭操作面板 / 隐藏窗口
;
; 用法:
;   SearchWindow.Create()             启动时调用一次
;   SearchWindow.Show([text])  Hide()  Toggle()
;===============================================================================

class SearchWindow {
    static Gui := "", Input := "", List := "", Separator := ""
    static Results := [], Selected := 0, Offset := 0
    static Mode := "results"                        ; results = 搜索结果; actions = 操作面板
    static ActionSource := "", SavedQuery := "", AllActions := []
    static HistoryIndex := 0
    static Width := 0, Padding := 0, InputHeight := 0, RowHeight := 0, IconSize := 0, VisibleRows := 8
    static _gdi := Map()
    static _searchTimer := "", _hideTimer := ""
    static _posX := 0, _posY := 0

    ;---------------------------------------------------------------------------
    ; Create
    ;---------------------------------------------------------------------------
    static Create() {
        SearchWindow.Width       := Win.Scale(AppSettings.Appearance["Width"])
        SearchWindow.VisibleRows := Max(1, Min(AppSettings.Appearance["VisibleRows"], 9))
        SearchWindow.Padding     := Win.Scale(ThemeManager.Get("Padding"))
        SearchWindow.RowHeight   := Win.Scale(ThemeManager.Get("RowHeight"))
        SearchWindow.IconSize    := Win.Scale(ThemeManager.Get("IconSize"))
        IconCache.Size := SearchWindow.IconSize
        SearchWindow._CreateGdiObjects()

        w := SearchWindow.Width, pad := SearchWindow.Padding
        background := ThemeManager.Get("Background")

        g := Gui("-Caption +AlwaysOnTop +ToolWindow -DPIScale" (SearchWindow._IsWindows11() ? "" : " +Border"), App.Name)
        g.BackColor := background
        g.MarginX := 0, g.MarginY := 0
        g.SetFont("s" ThemeManager.Get("InputFontSize") " c" ThemeManager.Get("InputText"), ThemeManager.FontName())

        searchBox := g.AddEdit("x" pad " y" pad " w" (w - 2 * pad) " r1 -E0x200 -VScroll -WantReturn Background" background)
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

        g.OnEvent("Close", (*) => SearchWindow.Hide())

        SearchWindow.Gui := g, SearchWindow.Input := searchBox, SearchWindow.List := list, SearchWindow.Separator := separator

        OnMessage(0x100, (p*) => SearchWindow._OnKeyDown(p*))               ; WM_KEYDOWN
        OnMessage(0x104, (p*) => SearchWindow._OnKeyDown(p*))               ; WM_SYSKEYDOWN (Alt+...)
        OnMessage(0x6,   (p*) => SearchWindow._OnActivate(p*))              ; WM_ACTIVATE
        OnMessage(0x20A, (p*) => SearchWindow._OnMouseWheel(p*))            ; WM_MOUSEWHEEL

        g.Show("Hide w" w " h" SearchWindow._WindowHeight(0))
        Win.SetCorner(g.Hwnd)
        Win.SetBorderColor(g.Hwnd, ThemeManager.Get("Border"))
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
        SearchWindow.HistoryIndex := 0
        SearchWindow._SetInput(text)

        area := Win.WorkAreaAtMouse()
        SearchWindow._posX := area.Left + (area.Right - area.Left - SearchWindow.Width) // 2
        SearchWindow._posY := area.Top + Round((area.Bottom - area.Top) * 0.2)
        SearchWindow.Gui.Show("x" SearchWindow._posX " y" SearchWindow._posY " w" SearchWindow.Width " h" SearchWindow._WindowHeight(SearchWindow._VisibleCount()))
        try WinActivate("ahk_id " SearchWindow.Gui.Hwnd)
        SearchWindow.Input.Focus()
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
        Win.SetCueBanner(SearchWindow.Input.Hwnd, I18n.T("Search.Placeholder"))
        SearchWindow._SetInput(SearchWindow.SavedQuery)
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
            case 0x26:                                                      ; ↑
                if (!actions && SearchWindow._RecallHistory())
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
            case 0x08:                                                      ; Backspace
                if (actions && SearchWindow.Input.Value = "") {
                    SearchWindow._CloseActions()
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

        DllCall("FillRect", "Ptr", hdc, "Ptr", rect, "Ptr", selected ? gdi["Selected"] : gdi["Background"])
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
