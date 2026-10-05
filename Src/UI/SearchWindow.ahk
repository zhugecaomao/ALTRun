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
;   ↑ ↓  PgUp PgDn  Ctrl+P Ctrl+N    只移动选择 (和 Alfred 一样, 不和翻历史混在一起)
;   Ctrl+↑ / Ctrl+↓                  上一条 / 下一条搜索记录; 翻回最新之后恢复原来输入的文字
;   Ctrl+1 ~ Ctrl+9                  直接执行可见的第 N 行
;   Tab                              自动补全; 文件夹: 进入文件夹浏览 (输入框变成 "路径\")
;   Insert                           标记 / 取消标记文件和文件夹 (可以标记多个), → 对所有标记的一起操作
;   →  (光标在末尾时)                打开操作面板; ← / Esc 返回
;   空格 (搜索框为空时)              进入文件搜索模式 (只搜文件, 提示 "搜索文件..."); Backspace 返回
;   空格 (已输入文字, SpaceToRun)    执行选中项; Shift+空格输入空格
;   Ctrl+C (输入框没有选中文字时)    复制当前项
;   Ctrl+L                           大字显示
;   F3                               编辑当前项 (没有结果时: 用输入的文字新建自定义命令)
;   Ctrl+Del (光标在末尾时)          删除当前项 (确认后)
;   F1                               关于 ALTRun (偏好设置的高级页)
;   F2 / Ctrl+,                      偏好设置
;   F4                               用记事本编辑 Data\ALTRun.json
;   Esc                              关闭操作面板 / 隐藏窗口
; 鼠标: 单击选择, 双击执行, 右键弹出这一项的操作菜单 (和操作面板相同);
;       按住输入框四周的空白处可以拖动窗口
;
; 位置: 显示在哪块屏幕由 Appearance.ShowOn 决定 (Mouse 鼠标所在 / Primary 主屏幕 /
; Active 当前窗口所在)。默认水平居中、离顶部 20%; 打开 Appearance.RememberPosition 后,
; 拖动过的位置按在屏幕里的相对位置 (千分比, Appearance.Position) 保存, 换一块屏幕也放在对应的地方。
;
; 用法:
;   SearchWindow.Create()             启动时调用一次
;   SearchWindow.Show([text])  Hide()  Toggle()
;===============================================================================

class SearchWindow {
    static Gui := "", Input := "", List := "", Separator := ""
    static Results := [], Selected := 0, Offset := 0
    static Mode := "results"                        ; results = 搜索结果; actions = 操作面板
    static HighlightText := ""                      ; 搜索结果标题里高亮和它匹配的字 (Highlight 颜色)
    static FileMode := false                        ; 文件搜索模式: 空的搜索框里按空格进入, 只搜文件
    static ActionSource := "", SavedQuery := "", AllActions := []
    static _actionsOnly := false                                            ; 直接打开的操作面板 (选中内容的操作): Esc / ← 关掉窗口
    static HistoryIndex := 0                         ; 正在看第几条搜索记录 (Knowledge.History, 1 = 最近), 0 = 没有在翻
    static Marked := Map()                           ; Insert 标记的文件 / 文件夹: 小写路径 -> 路径 (每次呼出时清空)
    static _historyDraft := ""                       ; 开始翻记录之前输入框里的文字, Ctrl+↓ 翻回来时恢复
    static Width := 0, Padding := 0, InputHeight := 0, RowHeight := 0, IconSize := 0, VisibleRows := 8
    static SelectedRadius := 0
    static _gdi := Map()
    static _searchTimer := "", _hideTimer := ""
    static _shownRows := -1                                                 ; 窗口当前按几行结果的高度显示
    static _tip := ""                                                       ; 这次显示时的使用提示 (HelpProvider.NextTip)
    static LastArea := "", HiddenAt := 0                                    ; 上次隐藏时所在屏幕的工作区和时间 (Hud 用)
    static _last := ""                                                      ; 上次隐藏时的搜索 {Text, FileMode, Selected} (KeepLastQuery)
    static _layoutBefore := 0                                               ; 呼出前前台窗口的输入法 (SwitchToEnglishInput 时隐藏后切回)
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
        if (IconCache.Size != SearchWindow.IconSize)                        ; 换了主题 / DPI: 缓存的图标是按原来的尺寸缩放的
            IconCache.Clear()
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
        OnMessage(0x201, (p*) => SearchWindow._OnLButtonDown(p*))           ; WM_LBUTTONDOWN: 拖动窗口
        OnMessage(0x232, (p*) => SearchWindow._OnMoved(p*))                 ; WM_EXITSIZEMOVE: 拖动结束

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
        for key in ["Title", "Subtitle", "Shortcut", "SelectedTitle", "SelectedSubtitle", "SelectedShortcut", "Highlight", "SelectedHighlight"]
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

    static IsActive() {
        return IsObject(SearchWindow.Gui) && WinActive("ahk_id " SearchWindow.Gui.Hwnd) ? true : false
    }

    ; text: 要搜索的文字; 不写时空白, 打开了 "保留上一次的搜索" (KeepLastQuery) 时恢复上次的
    ; 输入、文件搜索模式和选中的行, 文字全选: 按 Enter 再执行一次, 直接输入就开始新的搜索
    static _showing := false
    static Show(text := "") {
        if !IsObject(SearchWindow.Gui)
            return
        if SearchWindow._showing {                                          ; 上一次还没显示完 (例如显示过程中又按了热键): 不重入
            Logger.Debug("SearchWindow: show skipped, still showing")
            return
        }
        SearchWindow._showing := true
        try SearchWindow._Show(text)
        finally SearchWindow._showing := false
    }

    static _Show(text) {
        started := Logger.Ms()
        if Logger.Enabled
            Logger.Debug("SearchWindow: show")                              ; 和下面的 "Perf: show" 配对: 卡住时看得出卡在显示窗口里
        wasVisible := SearchWindow.IsVisible()
        if (text = "" && wasVisible)
            SearchWindow._RememberQuery()                                   ; 窗口还开着 (例如没有失去焦点就隐藏): 保留现在的输入
        App.RememberActiveWindow()
        SearchWindow._actionsOnly := false
        SearchWindow.Marked := Map()
        Usage.Count("Show")
        last := SearchWindow._last
        restore := (text = "" && AppSettings.General["KeepLastQuery"] && IsObject(last) && last.Text != "")
        ; 恢复上次的搜索, 而且输入框和结果还是隐藏时的样子: 原样显示, 不重新搜索 (输入新的文字时才搜索)。
        ; 搜索 (例如文件搜索) 要几十毫秒, 以前先搜索再显示, 这段时间里按下的第一个键还在原来的窗口里, 会丢失或延迟
        reuse := restore && SearchWindow.Mode = "results" && SearchWindow.Input.Value == last.Text && SearchWindow.FileMode = last.FileMode
        SearchWindow.Mode := "results"
        SearchWindow.FileMode := restore ? last.FileMode : false
        SearchWindow._tip := AppSettings.General["ShowTips"] ? HelpProvider.NextTip() : ""
        SearchWindow._UpdateCueBanner()
        SearchWindow.HistoryIndex := 0
        Logger.Trace("show: results")
        if !reuse {
            SearchWindow._SetInput(restore ? last.Text : text)
            if restore
                SearchWindow.MoveSelection(last.Selected - SearchWindow.Selected)
        }

        appearance := AppSettings.Appearance
        pos := SearchWindow.Place(SearchWindow._ScreenArea(appearance["ShowOn"]), SearchWindow.Width
            , SearchWindow._WindowHeight(SearchWindow.VisibleRows), appearance["RememberPosition"] ? appearance["Position"] : "")
        ; 窗口隐藏期间的重画请求会被丢掉, 显示出来后才重画的话, 结果列表会先白一下 / 闪一下。
        ; 先让 DWM 把窗口藏起来 (cloak), 显示并立即画好整个窗口之后再露出来
        hwnd := SearchWindow.Gui.Hwnd
        Logger.Trace("show: window")
        cloaked := !SearchWindow.IsVisible() && SearchWindow._Cloak(hwnd, true)
        try {
            SearchWindow.Gui.Show("x" pos.X " y" pos.Y " w" SearchWindow.Width " h" SearchWindow._WindowHeight(SearchWindow._VisibleCount()))
            SearchWindow._shownRows := SearchWindow._VisibleCount()
            DllCall("RedrawWindow", "Ptr", hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x185)   ; RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW
        } finally {
            if cloaked
                SearchWindow._Cloak(hwnd, false)
        }
        Logger.Trace("show: activate")
        try WinActivate("ahk_id " SearchWindow.Gui.Hwnd)                    ; 先拿到焦点, 之后按下的键都进搜索框
        SearchWindow.Input.Focus()
        len := StrLen(SearchWindow.Input.Value)
        if restore
            SendMessage(0xB1, 0, -1, SearchWindow.Input.Hwnd)               ; 恢复的文字全选: 直接输入就替换掉
        else
            SendMessage(0xB1, len, len, SearchWindow.Input.Hwnd)            ; 获得焦点时 Edit 会全选, 把光标放回末尾
        if AppSettings.General["SwitchToEnglishInput"] {
            Logger.Trace("show: input language")
            if (!wasVisible || !SearchWindow._layoutBefore)                     ; 隐藏时切回呼出前的输入法
                SearchWindow._layoutBefore := App.PreviousWindow ? Win.KeyboardLayout(App.PreviousWindow) : 0
            Win.SwitchToEnglishIME()
        }
        Logger.Time("show", started)
    }

    ; 呼出时切到了英文输入法: 隐藏时切回原来的。所有程序共用一个输入法时 (Windows 默认) 才需要;
    ; 打开了 "为每个应用窗口使用不同的输入法" 时, 只有 ALTRun 自己切换了, 别的程序不受影响
    static _RestoreInputLanguage() {
        layout := SearchWindow._layoutBefore, SearchWindow._layoutBefore := 0
        if (!layout || !AppSettings.General["SwitchToEnglishInput"] || !AppSettings.General["RestoreInput"]
            || Win.PerWindowInputMethod() || layout = Win.KeyboardLayout())
            return
        if SearchWindow.IsActive() {                                        ; 还在前台: 自己切回去, 对所有程序生效
            try DllCall("ActivateKeyboardLayout", "Ptr", layout, "UInt", 0)
        } else if (hwnd := WinExist("A")) {                                 ; 已经失去焦点: 请前台的窗口切换
            try PostMessage(0x50, 0, layout, , "ahk_id " hwnd)              ; WM_INPUTLANGCHANGEREQUEST
        }
    }

    ; DWMWA_CLOAK: 窗口照常显示、绘制, 只是 DWM 不把它画到屏幕上。返回 true = 成功 (Windows 8 起才有)
    static _Cloak(hwnd, on) {
        try return DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "UInt", 13, "Int*", on ? 1 : 0, "UInt", 4) = 0
        return false
    }

    ; 窗口放在 area 里的位置 {X, Y}。position: {X, Y} 千分比 (0 = 最左 / 最上, 1000 = 最右 / 最下,
    ; Y 是窗口上沿在工作区高度里的位置), "" = 默认 (水平居中, 离顶部 20%)。
    ; maxHeight: 结果最多时的窗口高度, 位置再靠下也保证整个窗口留在屏幕里
    static Place(area, width, maxHeight, position := "") {
        fx := 500, fy := 200
        if (position is Map && position.Has("X") && position.Has("Y") && IsNumber(position["X"]) && IsNumber(position["Y"]))
            fx := position["X"], fy := position["Y"]
        areaW := area.Right - area.Left, areaH := area.Bottom - area.Top
        x := area.Left + Round((areaW - width) * Max(0, Min(fx, 1000)) / 1000)
        y := area.Top + Round(areaH * Max(0, Min(fy, 1000)) / 1000)
        x := Max(area.Left, Min(x, area.Right - width))
        y := Max(area.Top, Min(y, area.Bottom - maxHeight))
        return {X: x, Y: y}
    }

    ; Place 的反过来: 窗口左上角 (x, y) 在 area 里的千分比位置
    static RelativePosition(area, width, x, y) {
        areaW := area.Right - area.Left, areaH := area.Bottom - area.Top
        fx := (areaW > width) ? Round((x - area.Left) * 1000 / (areaW - width)) : 500
        fy := (areaH > 0) ? Round((y - area.Top) * 1000 / areaH) : 200
        return Map("X", Max(0, Min(fx, 1000)), "Y", Max(0, Min(fy, 1000)))
    }

    ; showOn: Mouse 鼠标所在的屏幕 / Primary 主屏幕 / Active 呼出前的活动窗口所在的屏幕 (没有时用鼠标所在)
    static _ScreenArea(showOn) {
        switch showOn, false {
            case "Primary": return Win.WorkArea(MonitorGetPrimary())
            case "Active":
                area := Win.WorkAreaOfWindow(App.PreviousWindow)
                if IsObject(area)
                    return area
        }
        return Win.WorkAreaAtMouse()
    }

    ; 按住窗口本身 (输入框四周的空白, 不是输入框和结果列表) 拖动: 当作按住标题栏
    static _OnLButtonDown(wParam, lParam, msg, hwnd) {
        if (!IsObject(SearchWindow.Gui) || hwnd != SearchWindow.Gui.Hwnd)
            return
        PostMessage(0xA1, 2, , , "ahk_id " hwnd)                            ; WM_NCLBUTTONDOWN, HTCAPTION
        return 0
    }

    ; 拖动结束: 打开了 "记住位置" 时保存位置 (不保存时只在这一次显示期间有效)
    static _OnMoved(wParam, lParam, msg, hwnd) {
        if (!IsObject(SearchWindow.Gui) || hwnd != SearchWindow.Gui.Hwnd || !AppSettings.Appearance["RememberPosition"])
            return
        WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
        area := Win.WorkAreaAt(x + w // 2, y + SearchWindow.InputHeight // 2)
        AppSettings.Appearance["Position"] := SearchWindow.RelativePosition(area, w, x, y)
        AppSettings.Save()
    }

    static Hide() {
        if SearchWindow.IsVisible() {
            try {                                                           ; 刚隐藏后的提示 (例如 "已置顶") 显示在同一块屏幕上, 见 Hud
                WinGetPos(&x, &y, &w, &h, "ahk_id " SearchWindow.Gui.Hwnd)
                SearchWindow.LastArea := Win.WorkAreaAt(x + w // 2, y + SearchWindow.InputHeight // 2), SearchWindow.HiddenAt := A_TickCount
            }
            SearchWindow._RememberQuery()
            SearchWindow._RestoreInputLanguage()
            SearchWindow.Gui.Hide()
        }
        LargeType.Close()
    }

    ; 记住隐藏时的搜索 (KeepLastQuery 用); 操作面板里记住打开面板之前的搜索
    static _RememberQuery() {
        actions := (SearchWindow.Mode = "actions")
        SearchWindow._last := {Text: actions ? SearchWindow.SavedQuery : SearchWindow.Input.Value
                             , FileMode: SearchWindow.FileMode, Selected: actions ? 1 : Max(1, SearchWindow.Selected)}
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
        browse := FileSearchProvider.BrowsePath(text)                       ; 浏览文件夹: 只按最后一段 (过滤的文字) 高亮, 路径里的词不算
        SearchWindow.HighlightText := IsObject(browse) ? browse.Filter : text
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

    ; 每次按键都会调用: 行数没变时只重画列表, 不动窗口大小和控件 (调整窗口大小比重画慢得多)
    static _Layout() {
        visibleCount := SearchWindow._VisibleCount()
        list := SearchWindow.List
        if (list.GetCount() != visibleCount) {
            list.Opt("-Redraw")
            list.Delete()
            Loop visibleCount
                list.Add("", "")
            list.Move(, , , visibleCount * SearchWindow.RowHeight)
            list.Opt("+Redraw")
        }
        if (list.Visible != (visibleCount > 0)) {
            list.Visible := visibleCount > 0
            SearchWindow.Separator.Visible := visibleCount > 0
        }
        SearchWindow._Repaint()
        if (SearchWindow.IsVisible() && SearchWindow._shownRows != visibleCount) {
            SearchWindow.Gui.Show("NA w" SearchWindow.Width " h" SearchWindow._WindowHeight(visibleCount))
            SearchWindow._shownRows := visibleCount
        }
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
            RecentProvider.Remember(source)
            Usage.CountItem(source)
            SearchWindow.Marked := Map()
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
        RecentProvider.Remember(item)
        Usage.CountItem(item)
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
    ; 直接打开某一项的操作面板 (不在搜索结果里, 例如选中内容的操作): Esc / ← 关掉窗口
    static ShowActions(item) {
        SearchWindow.Show()
        if SearchWindow._OpenActionsFor(item)
            SearchWindow._actionsOnly := true
    }

    static _OpenActions() {
        return SearchWindow._OpenActionsFor(SearchWindow.SelectedItem())
    }

    static _OpenActionsFor(item) {
        if (!IsObject(item) || !item.Valid || SearchWindow.Mode = "actions")
            return false
        if SearchWindow.Marked.Count {                                      ; 有标记的文件: 对所有标记的一起操作
            paths := []
            for key, markedPath in SearchWindow.Marked
                paths.Push(markedPath)
            item := ResultItem(I18n.T("Search.MarkedItems", paths.Length), "", {Icon: "folder:"})
            actions := ActionCatalog.ListForMany(paths)
        } else
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
    ; 空搜索框里的灰色文字: 文件搜索模式 "搜索文件...", 否则是这次显示时轮到的使用提示 (关掉提示时 "ALTRun 搜索")
    static _UpdateCueBanner() {
        if SearchWindow.FileMode
            text := I18n.T("Search.FilesPlaceholder")
        else
            text := (SearchWindow._tip != "") ? SearchWindow._tip : I18n.T("Search.Placeholder")
        Win.SetCueBanner(SearchWindow.Input.Hwnd, text)
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
        RecentProvider.Remember(item)
        Usage.CountItem(item)
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
                if (actions && SearchWindow._actionsOnly)
                    SearchWindow.Hide()
                else if actions
                    SearchWindow._CloseActions()
                else
                    SearchWindow.Hide()
                return 0
            case 0x20:                                                      ; 空格: 空的搜索框里进入文件搜索模式
                if (ctrl || alt)
                    return
                ; 文字全选时 (恢复的上次搜索) 空格会替换掉它们, 当作空的搜索框处理
                if (!shift && !actions && !SearchWindow.FileMode && SearchWindow._AllTextSelected()) {
                    SearchWindow._SetInput("")
                    if SearchWindow._CanEnterFileMode() {
                        SearchWindow._SetFileMode(true)
                        return 0
                    }
                    return
                }
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
            case 0x26:                                                      ; ↑ / Ctrl+↑ 上一条搜索记录
                if ctrl
                    SearchWindow.RecallHistory(1)
                else
                    SearchWindow.MoveSelection(-1)
                return 0
            case 0x28:                                                      ; ↓ / Ctrl+↓ 下一条搜索记录
                if ctrl
                    SearchWindow.RecallHistory(-1)
                else
                    SearchWindow.MoveSelection(1)
                return 0
            case 0x21:                                                      ; PgUp
                SearchWindow.MoveSelection(-SearchWindow.VisibleRows)
                return 0
            case 0x22:                                                      ; PgDn
                SearchWindow.MoveSelection(SearchWindow.VisibleRows)
                return 0
            case 0x2D:                                                      ; Insert: 标记 / 取消标记 (和 Total Commander 一样), 移到下一行
                if (!actions && !ctrl && !shift && !alt) {
                    SearchWindow.ToggleMark()
                    return 0
                }
                return
            case 0x09:                                                      ; Tab
                SearchWindow._AutoComplete()
                return 0
            case 0x27:                                                      ; →
                if (SearchWindow._CaretAtEnd() && SearchWindow._OpenActions())
                    return 0
                return
            case 0x25:                                                      ; ←
                if (actions && SearchWindow.Input.Value = "") {
                    if SearchWindow._actionsOnly
                        SearchWindow.Hide()
                    else
                        SearchWindow._CloseActions()
                    return 0
                }
                return
            case 0x70:                                                      ; F1
                SearchWindow.Hide()
                App.About()
                return 0
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
                    if SearchWindow._actionsOnly
                        SearchWindow.Hide()
                    else
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

    ; step: 1 = 更早的一条 (Ctrl+↑), -1 = 更近的一条 (Ctrl+↓)。输入框里的文字一改就从最近的一条重新开始
    ; (_OnInputChange 把 HistoryIndex 清零)。和现在的文字一样的记录跳过, 按一下总有变化
    static RecallHistory(step) {
        if (SearchWindow.Mode = "actions" || SearchWindow.FileMode)
            return
        history := Knowledge.History
        current := SearchWindow.Input.Value
        index := SearchWindow.HistoryIndex
        if (index = 0)
            SearchWindow._historyDraft := current
        Loop {
            index += step
            if (index < 0 || index > history.Length)
                return                                                      ; 已经是最早 / 最新的一条
            text := index ? history[index] : SearchWindow._historyDraft
            if (text != current || index = 0)
                break
        }
        SearchWindow._SetInput(text)
        SearchWindow.HistoryIndex := index                                  ; _SetInput 不经过 Change 事件, 这里保持历史位置
    }

    ; 标记 / 取消标记选中的文件或文件夹, 然后移到下一行
    static ToggleMark() {
        item := SearchWindow.SelectedItem()
        if !(IsObject(item) && (item.Kind = "file" || item.Kind = "folder") && item.Arg != "")
            return false
        key := StrLower(item.Arg)
        if SearchWindow.Marked.Has(key)
            SearchWindow.Marked.Delete(key)
        else
            SearchWindow.Marked[key] := item.Arg
        if (SearchWindow.Selected < SearchWindow.Results.Length)
            SearchWindow.MoveSelection(1)
        else
            SearchWindow._Repaint()
        return true
    }

    static IsMarked(item) {
        return SearchWindow.Marked.Count && IsObject(item) && (item.Kind = "file" || item.Kind = "folder") && SearchWindow.Marked.Has(StrLower(item.Arg))
    }

    static _AutoComplete() {
        item := SearchWindow.SelectedItem()
        if (!IsObject(item) || SearchWindow.Mode = "actions")
            return
        text := (item.AutoComplete != "") ? item.AutoComplete : item.Title
        if (item.AutoComplete = "" && item.Kind = "folder" && RegExMatch(item.Arg, "^([A-Za-z]:\\|\\\\)"))
            text := RTrim(item.Arg, "\") "\"                                ; 文件夹: Tab 进入, 浏览里面的内容
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

    ; 输入框里的文字全部选中 (恢复上次的搜索之后): 这时按键会替换掉它们
    static _AllTextSelected() {
        len := StrLen(SearchWindow.Input.Value)
        selection := SendMessage(0xB0, 0, 0, SearchWindow.Input.Hwnd)        ; EM_GETSEL
        return len && (selection & 0xFFFF) = 0 && ((selection >> 16) & 0xFFFF) >= len
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

    ; 置顶标记 Resources\Icons\Pinned.ico, 按尺寸读取一次 (0 = 没有这个文件)
    static _PinBadge(size) {
        static icons := Map(), file := A_ScriptDir "\Resources\Icons\Pinned.ico"
        if !icons.Has(size)
            icons[size] := FileExist(file) ? DllCall("LoadImageW", "Ptr", 0, "WStr", file, "UInt", 1, "Int", size, "Int", size, "UInt", 0x10, "Ptr") : 0   ; IMAGE_ICON, LR_LOADFROMFILE
        return icons[size]
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
        if (SearchWindow.Mode = "results" && SearchWindow.IsMarked(item)) {  ; 标记的: 左边一条高亮色的竖线
            mark := Buffer(16)
            NumPut("Int", left + Win.Scale(2), "Int", top + Win.Scale(4), "Int", left + Win.Scale(6), "Int", bottom - Win.Scale(4), mark)
            brush := DllCall("CreateSolidBrush", "UInt", gdi["HighlightColor"], "Ptr")
            DllCall("FillRect", "Ptr", hdc, "Ptr", mark, "Ptr", brush)
            DllCall("DeleteObject", "Ptr", brush)
        }

        iconY := top + (rowH - iconSize) // 2
        if (hIcon := IconCache.Get(item.Icon))
            DllCall("DrawIconEx", "Ptr", hdc, "Int", left + pad, "Int", iconY, "Ptr", hIcon, "Int", iconSize, "Int", iconSize, "UInt", 0, "Ptr", 0, "UInt", 3)
        if (item.Pinned && (hBadge := SearchWindow._PinBadge(badge := Max(Win.Scale(13), Round(iconSize * 0.55))))) {   ; 置顶的: 右下角一个图钉
            offset := iconSize - badge + Win.Scale(3)
            DllCall("DrawIconEx", "Ptr", hdc, "Int", left + pad + offset, "Int", iconY + offset, "Ptr", hBadge, "Int", badge, "Int", badge, "UInt", 0, "Ptr", 0, "UInt", 3)
        }

        textLeft := left + pad + iconSize + Win.Scale(12)
        shortcutW := Win.Scale(56)
        oldFont := DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["ShortcutFont"], "Ptr")
        textRight := right - pad - shortcutW - Win.Scale(8)

        if (visibleRow <= 9 && SearchWindow.Mode = "results") {             ; 右侧 Ctrl+N 提示
            DllCall("SetTextColor", "Ptr", hdc, "UInt", selected ? gdi["SelectedShortcutColor"] : gdi["ShortcutColor"])
            SearchWindow._DrawText(hdc, "Ctrl+" visibleRow, right - pad - shortcutW, top, right - pad, bottom, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX)
        } else if (SearchWindow.Mode = "actions") {                         ; 操作面板: 右侧显示快捷键或说明, 按实际宽度留位置
            hintW := 0, gap := Win.Scale(12)
            if (item.Subtitle != "") {
                hintW := SearchWindow._TextWidth(hdc, item.Subtitle)
                DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["TitleFont"])
                titleW := SearchWindow._TextWidth(hdc, item.Title)
                DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["ShortcutFont"])
                hintW := SearchWindow.HintWidth(right - pad - textLeft, titleW, hintW, gap)
                DllCall("SetTextColor", "Ptr", hdc, "UInt", selected ? gdi["SelectedShortcutColor"] : gdi["ShortcutColor"])
                SearchWindow._DrawText(hdc, item.Subtitle, right - pad - hintW, top, right - pad, bottom, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS)
            }
            textRight := right - pad - (hintW ? hintW + gap : 0)
        }

        DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["TitleFont"])
        DllCall("SetTextColor", "Ptr", hdc, "UInt", selected ? gdi["SelectedTitleColor"] : gdi["TitleColor"])
        hasSubtitle := (item.Subtitle != "" && SearchWindow.Mode = "results")
        middle := top + Round(rowH * 0.54)
        ranges := (SearchWindow.Mode = "results") ? SearchWindow.HighlightRanges(item) : []
        if hasSubtitle
            SearchWindow._DrawTitle(hdc, item.Title, ranges, textLeft, top, textRight, middle, DT_BOTTOM | DT_SINGLELINE | DT_NOPREFIX, selected)
        else
            SearchWindow._DrawTitle(hdc, item.Title, ranges, textLeft, top, textRight, bottom, DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX, selected)

        if hasSubtitle {
            DllCall("SelectObject", "Ptr", hdc, "Ptr", gdi["SubtitleFont"])
            DllCall("SetTextColor", "Ptr", hdc, "UInt", selected ? gdi["SelectedSubtitleColor"] : gdi["SubtitleColor"])
            ellipsis := (item.Kind = "file" || item.Kind = "folder") ? DT_PATH_ELLIPSIS : DT_END_ELLIPSIS
            SearchWindow._DrawText(hdc, item.Subtitle, textLeft, middle + Win.Scale(2), textRight, bottom, DT_SINGLELINE | DT_NOPREFIX | ellipsis)
        }
        DllCall("SelectObject", "Ptr", hdc, "Ptr", oldFont)
    }

    ; 操作面板一行里右侧说明的宽度: 名称优先 (最多占一半), 说明用剩下的; 放得下时两者都完整显示
    static HintWidth(available, titleW, hintW, gap) {
        keep := Min(titleW, available // 2)
        return Max(0, Min(hintW, available - keep - gap))
    }

    ; 标题里要高亮的位置, 按搜索文字缓存在结果上 (重画、移动选中行时不重算)
    static HighlightRanges(item) {
        text := SearchWindow.HighlightText
        if (Trim(text) = "" || item.Title = "")
            return []
        if (!item.HasOwnProp("HighlightFor") || item.HighlightFor !== text)
            item.HighlightFor := text, item.HighlightRanges := FuzzyMatcher.MatchRanges(text, item.Title)
        return item.HighlightRanges
    }

    ; 画标题: 匹配的字用 Highlight 颜色, 放不下时自己截断并加 "…"。
    ; 每个字的位置都按整个标题一次量出来 (GetTextExtentExPoint), 各段用 ExtTextOut 按这些位置摆放;
    ; 没有高亮时也这样画。每段分开量再拼起来、或者有没有高亮用两种画法 (DrawText 会按字距调整),
    ; 都会差 1 像素左右, 输入时高亮的边界移动, 字距就会跟着跳动
    static _DrawTitle(hdc, title, ranges, left, top, right, bottom, flags, selected) {
        static DT_VCENTER := 0x4, DT_BOTTOM := 0x8, ETO_CLIPPED := 0x4
        if (title = "")
            return
        gdi := SearchWindow._gdi
        available := right - left
        extents := SearchWindow._Extents(hdc, title)
        if (extents.Length && extents[extents.Length] > available) {
            ellipsis := SearchWindow._TextWidth(hdc, "…"), fit := 0
            while (fit < extents.Length && extents[fit + 1] <= available - ellipsis)
                fit++
            title := SubStr(title, 1, fit) "…"
            extents := SearchWindow._Extents(hdc, title)
        }
        metrics := Buffer(60, 0)                                            ; TEXTMETRICW, tmHeight 在最前面
        DllCall("GetTextMetricsW", "Ptr", hdc, "Ptr", metrics)
        lineHeight := NumGet(metrics, 0, "Int")
        y := (flags & DT_BOTTOM) ? bottom - lineHeight : (flags & DT_VCENTER) ? top + (bottom - top - lineHeight) // 2 : top
        clip := Buffer(16)
        NumPut("Int", left, "Int", top, "Int", right, "Int", bottom, clip)
        normal := selected ? gdi["SelectedTitleColor"] : gdi["TitleColor"]
        highlight := selected ? gdi["SelectedHighlightColor"] : gdi["HighlightColor"]
        pos := 1
        for segment in FuzzyMatcher.Segments(title, ranges) {
            count := StrLen(segment[1])
            advances := Buffer(count * 4)                                   ; 每个字的宽度 (相邻两个累计宽度之差)
            Loop count
                NumPut("Int", extents[pos + A_Index - 1] - (pos + A_Index - 2 >= 1 ? extents[pos + A_Index - 2] : 0), advances, (A_Index - 1) * 4)
            x := left + (pos > 1 ? extents[pos - 1] : 0)
            DllCall("SetTextColor", "Ptr", hdc, "UInt", segment[2] ? highlight : normal)
            DllCall("ExtTextOutW", "Ptr", hdc, "Int", x, "Int", y, "UInt", ETO_CLIPPED, "Ptr", clip, "WStr", segment[1], "UInt", count, "Ptr", advances)
            pos += count
        }
        DllCall("SetTextColor", "Ptr", hdc, "UInt", normal)
    }

    ; 每个字结束处到开头的宽度 (像素, 按整段文字一次量出): extents[i] = 前 i 个字的宽度
    static _Extents(hdc, text) {
        count := StrLen(text), extents := []
        if !count
            return extents
        fit := 0, partial := Buffer(count * 4, 0), size := Buffer(8, 0)
        DllCall("GetTextExtentExPointW", "Ptr", hdc, "WStr", text, "Int", count, "Int", 0x7FFFFFFF, "Int*", &fit, "Ptr", partial, "Ptr", size)
        Loop count
            extents.Push(NumGet(partial, (A_Index - 1) * 4, "Int"))
        return extents
    }

    static _TextWidth(hdc, text) {
        size := Buffer(8, 0)
        DllCall("GetTextExtentPoint32W", "Ptr", hdc, "WStr", text, "Int", StrLen(text), "Ptr", size)
        return NumGet(size, 0, "Int")
    }

    static _DrawText(hdc, text, left, top, right, bottom, flags) {
        rect := Buffer(16)
        NumPut("Int", left, "Int", top, "Int", right, "Int", bottom, rect)
        DllCall("DrawTextW", "Ptr", hdc, "WStr", text, "Int", -1, "Ptr", rect, "UInt", flags)
    }
}
