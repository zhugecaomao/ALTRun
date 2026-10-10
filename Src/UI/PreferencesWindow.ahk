;===============================================================================
; PreferencesWindow.ahk - 偏好设置窗口 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 Alfred 的 Preferences 一样左边是分类, 右边是该分类的设置。
; 所有修改先作用在设置的一份副本 (Working) 上, 按钮和一般的 Windows 设置窗口一样:
;   确定   写回 ALTRun.json, 关闭窗口, ALTRun 在后台重新载入 (没有修改时直接关闭)
;   取消   丢弃修改 (有修改时先确认); Esc 和右上角的关闭按钮相同
;   应用   写回并重新载入, 偏好设置在同一页、同一位置重新打开; 没有修改时是灰色的
;   帮助   (F1) 打开当前页对应的 Wiki 说明
; 设置要在重新载入后才生效 (热键、主题、索引...), 所以 "应用" 时窗口会闪一下。
;
; 页面构建用几个小工具, 每个设置项只写一行:
;   _Check("General.LaunchAtLogin", "Prefs.LaunchAtLogin", , , "Prefs.Group.Startup")  复选框 (最后是左列的分组标签)
;   _Field("Features.Snippets.Keyword", "Prefs.SnippetKeyword", "M")   文本框, 宽度 "S" / "M" / "L"
;   _Field("General.Hotkey", "Prefs.Hotkey", "K", "hotkey")     录制热键的框 (HotkeyBox): 显示 Alt+Space, 保存 !Space
;   _Choice("General.Language", "Prefs.Language", 值数组, 显示文字数组)
;   _Lines("Features.Applications.Folders", "Prefs.AppFolders", 行数)   数组 <-> 每行一项
;   _Csv("Features.FileSearch.Keywords", "Prefs.FileKeywords", "M")   数组 <-> 逗号分隔
;   _Button("Prefs.ClipClear", fn) / _Section("Prefs.QuickSwitch") / _Info(标签, 文字)
;   _List(...)                                                  列表 + 添加/编辑/删除 (ItemEditor)
; 路径是 ALTRun.json 里的键, 用 "." 连接。
; 布局是两列表单: 左列是右对齐的标签 (LabelW), 所有控件从 _InputX() 开始, 这样每一页都整齐。
; 和 Alfred 一样, 每个设置下面有一行灰色的说明: I18n 里有 "<标签的键>.Desc" (例如
; Prefs.LaunchAtLogin.Desc) 时自动显示, 不用在页面里另写。
;
; 用法:
;   PreferencesWindow.Show([页码, x, y])
;   PreferencesWindow.AddCommands(paths)   "发送到" 时偏好设置开着: 加进 "命令" 页
;===============================================================================

class PreferencesWindow {
    static Gui := "", Working := "", Pages := [], Binds := [], PageList := ""
    static _lists := Map()                                                  ; 列表页的设置路径 (例如 "CustomCommands") -> 加入 / 编辑项目的方法
    ; 页面的顺序 (和 _BuildPages 一致, 有测试检查); 搜索里的 "ALTRun 偏好设置: 外观" 等由它生成, 不用打开窗口
    static PageKeys := ["Prefs.Page.General", "Prefs.Page.Window", "Prefs.Page.Appearance", "Prefs.Page.Features", "Prefs.Page.Applications"
        , "Prefs.Page.FileSearch", "Prefs.Page.FileIndex", "Prefs.Page.Commands", "Prefs.Page.Snippets", "Prefs.Page.Clipboard", "Prefs.Page.WebSearch"
        , "Prefs.Page.Calculator", "Prefs.Page.Scripts", "Prefs.Page.Hotkeys", "Prefs.Page.QuickSwitch", "Prefs.Page.QSPanel", "Prefs.Page.DateStamp"
        , "Prefs.Page.Usage", "Prefs.Page.Advanced"]
    static _page := 0, _y := 0
    static _dirty := false, _ready := false, ApplyButton := ""
    ; 两列表单 (和 Alfred / Listary 一样): 左列是右对齐的标签, 一页里的控件从同一条竖线 (_InputX) 开始,
    ; 左列宽度每页按标签长短定 (_BeginPage); 输入框只用三种宽度, 复选框按组在左列加标签,
    ; 按钮统一尺寸, 灰色说明和控件左对齐
    static ContentX := 190, ContentW := 560, LabelW := 190
    static WidthS := 70, WidthM := 220, ButtonW := 200, ButtonH := 26       ; WidthL = 整个控件列 (_InputW)
    static WidthK := 150                                                    ; 热键框 ("K")
    static _pairX := 0, _pairFirstW := 0, _pairSecondW := 0, _pairSizes := ["", ""]
    static ButtonY := 619                                                   ; 底部按钮的位置; 页面内容要在它上面 (y < ButtonY - 10)
    static LabelMax := 250                                                  ; 左列最宽多少, 再长的标签换行
    static _positionReset := false
    static _pageKey := "", _wider := Map()                                  ; 正在建的页; 标签放不下的页 -> 需要的左列宽度
    static _labelWidths := Map()                                            ; "语言 页" -> 加宽后的左列宽度 (打开过一次就记住)

    ; pageIndex: 页码, 或页面的键 (例如 "Prefs.Page.Advanced")
    static Show(pageIndex := 1, x := "", y := "") {
        if IsObject(PreferencesWindow.Gui) {
            index := IsInteger(pageIndex) ? pageIndex : PreferencesWindow._PageIndex(pageIndex)
            if (index >= 1 && index <= PreferencesWindow.Pages.Length)
                PreferencesWindow._SelectRow(index)
            WinActivate("ahk_id " PreferencesWindow.Gui.Hwnd)
            return
        }
        SearchWindow.Hide()
        ; 左列宽度每页写死了一个值 (按中英文的标签定); 别的语言的标签更长、放不下时,
        ; 按实际的文字宽度加宽那一页的左列, 重新建一次窗口 (记住宽度, 下次打开不用再建两次)
        Loop 2 {
            PreferencesWindow.Working := PreferencesWindow.DeepCopy(AppSettings.Data)
            PreferencesWindow.Pages := [], PreferencesWindow.Binds := [], PreferencesWindow._wider := Map(), PreferencesWindow._lists := Map()
            PreferencesWindow._dirty := false, PreferencesWindow._ready := false, PreferencesWindow._positionReset := false, PreferencesWindow._page := 0

            g := Gui("-MinimizeBox", I18n.T("Prefs.Title"))
            g.SetFont("s9", ThemeManager.FontName())
            g.OnEvent("Close", (*) => PreferencesWindow.Cancel())            ; 返回 true = 不关闭 (选择了继续编辑)
            g.OnEvent("Escape", (*) => PreferencesWindow.Cancel())
            PreferencesWindow.Gui := g

            ; 左边的页面列表最先建: Tab 键的顺序是 列表 -> 页面里的控件 -> 底部按钮 (页面名字建完页面再填)
            ; ListView 而不是 ListBox: 每页前面一个图标 (系统图标字体画的, 见 NavIcons); 0x20 整行选中, 0x10000 双缓冲
            pageList := g.AddListView("x12 y12 w160 h" (PreferencesWindow.ButtonY - 24) " -Hdr -Multi LV0x10020", [""])
            PreferencesWindow._BuildPages()
            if (!PreferencesWindow._wider.Count || A_Index = 2)
                break
            for pageKey, width in PreferencesWindow._wider
                PreferencesWindow._labelWidths[I18n.Lang " " pageKey] := width
            g.Destroy()
        }

        names := []
        for page in PreferencesWindow.Pages
            names.Push(page.Name)
        icons := NavIcons.ImageList(PreferencesWindow.PageKeys)             ; 图像列表的高度就是行高 (更宽松的侧边栏)
        if icons
            pageList.SetImageList(icons, 1)
        for index, name in names
            pageList.Add("Icon" index, name)
        pageList.ModifyCol(1, "AutoHdr")                                    ; 唯一的一列占满列表宽度 (列比列表窄时右边会多一条列分隔线)
        DllCall("uxtheme\SetWindowTheme", "Ptr", pageList.Hwnd, "Str", "Explorer", "Ptr", 0)   ; 和资源管理器一样的悬停 / 选中效果
        pageList.OnNotify(-101, (ctrl, lParam) => PreferencesWindow._OnPageRow(lParam))   ; LVN_ITEMCHANGED (ItemSelect 事件有时收不到)
        pageList.OnNotify(-12, (ctrl, lParam) => PreferencesWindow._OnPageListDraw(lParam))   ; NM_CUSTOMDRAW: 不画虚线焦点框
        for code in [-2, -3, -5, -6]                                        ; NM_CLICK / NM_DBLCLK / NM_RCLICK / NM_RDBLCLK (点得快时第二下算双击)
            pageList.OnNotify(code, (ctrl, lParam) => PreferencesWindow._OnPageClick(lParam))
        PreferencesWindow.PageList := pageList

        buttonY := " y" PreferencesWindow.ButtonY
        g.AddButton("x12" buttonY " w85", I18n.T("Prefs.Help")).OnEvent("Click", (*) => PreferencesWindow.Help())
        g.AddButton("x475" buttonY " w85 Default", I18n.T("Prefs.OK")).OnEvent("Click", (*) => PreferencesWindow.OK())
        g.AddButton("x570" buttonY " w85", I18n.T("Prefs.Cancel")).OnEvent("Click", (*) => PreferencesWindow.Cancel())
        apply := g.AddButton("x665" buttonY " w85 Disabled", I18n.T("Prefs.Apply"))
        apply.OnEvent("Click", (*) => PreferencesWindow.Apply())
        PreferencesWindow.ApplyButton := apply
        HotIfWinActive("ahk_id " g.Hwnd)
        Hotkey("F1", (*) => PreferencesWindow.Help())
        HotIfWinActive()

        if !IsInteger(pageIndex)
            pageIndex := PreferencesWindow._PageIndex(pageIndex)
        pageIndex := Max(1, Min(pageIndex, PreferencesWindow.Pages.Length))
        PreferencesWindow._SelectRow(pageIndex)
        g.Show((IsInteger(x) && IsInteger(y) ? "x" x " y" y " " : "") "w765 h" (PreferencesWindow.ButtonY + 41))
        pageList.Focus()                                                    ; 焦点在页面列表: ↑ ↓ 切换页面, 不会不小心改了第一页的设置
        ; 打开窗口时程序自己填的值不算修改, 等控件的通知都处理完再开始记录
        SetTimer(() => (PreferencesWindow._ready := true), -300)
    }

    static _SelectRow(index) {
        PreferencesWindow.SelectPage(index)
        PreferencesWindow.PageList.Modify(index, "Select Focus Vis")
    }

    ; 选中了一行: 显示那一页
    static _OnPageRow(lParam) {
        row := NumGet(lParam, 3 * A_PtrSize, "Int") + 1                     ; NMLISTVIEW: iItem (从 0 开始), uNewState, uOldState
        newState := NumGet(lParam, 3 * A_PtrSize + 8, "UInt"), oldState := NumGet(lParam, 3 * A_PtrSize + 12, "UInt")
        if !((newState ^ oldState) & 0x2)                                   ; LVIS_SELECTED 没变 (例如只是焦点变了)
            return
        if ((newState & 0x2) && row != PreferencesWindow._page)
            PreferencesWindow.SelectPage(row)
    }

    ; 选中行只用高亮表示, 不画虚线焦点框 (和 Windows 设置的导航一样)。Windows 按 "最近是不是用了键盘"
    ; 决定显示不显示焦点框, 所以有时第一次打开偏好设置会看到虚线、有时看不到; 这里在画每一行之前去掉 CDIS_FOCUS
    static _OnPageListDraw(lParam) {
        stage := NumGet(lParam, 3 * A_PtrSize, "UInt")                      ; NMCUSTOMDRAW.dwDrawStage
        if (stage = 0x1)                                                    ; CDDS_PREPAINT
            return 0x20                                                     ; CDRF_NOTIFYITEMDRAW
        if (stage = 0x10001) {                                              ; CDDS_ITEMPREPAINT
            offset := 6 * A_PtrSize + 16                                    ; NMCUSTOMDRAW.uItemState
            NumPut("UInt", NumGet(lParam, offset, "UInt") & ~0x10, lParam, offset)   ; CDIS_FOCUS
        }
        return 0                                                            ; CDRF_DODEFAULT
    }

    ; 点了列表下面的空白处 (没有点到行), ListView 会取消选中: 松开鼠标后选回当前页。
    ; 不在取消选中时马上选回: 点另一行时也会先取消选中上一行, 那样上一行会闪一下
    static _OnPageClick(lParam) {
        if (NumGet(lParam, 3 * A_PtrSize, "Int") < 0 && !PreferencesWindow.PageList.GetNext())   ; NMITEMACTIVATE.iItem = -1
            PreferencesWindow.PageList.Modify(PreferencesWindow._page, "Select Focus")
    }

    static _PageIndex(nameKey) {
        for index, page in PreferencesWindow.Pages
            if (page.Key = nameKey)
                return index
        return 1
    }

    ; 有控件被用户修改过: "应用" 变为可用, 取消时要确认; force: 刚打开窗口时也算 ("发送到" 加进来的命令)
    static MarkDirty(force := false) {
        if ((!PreferencesWindow._ready && !force) || PreferencesWindow._dirty)
            return
        PreferencesWindow._dirty := true
        try PreferencesWindow.ApplyButton.Enabled := true
    }

    ; 换页时只隐藏上一页、显示这一页的控件, 再只重画右边的页面区域:
    ; 不重画左边的页面列表和底部按钮, 切换时不闪 (不对整个窗口用 WM_SETREDRAW: 顶层窗口停止重画后子控件会画不出来)
    static SelectPage(pageIndex) {
        hwnd := PreferencesWindow.Gui.Hwnd
        previous := PreferencesWindow._page
        PreferencesWindow._page := pageIndex
        for index, page in PreferencesWindow.Pages {
            if (index = pageIndex || index = previous || previous = 0)      ; 刚建好时所有页面都要隐藏一次
                for ctrl in page.Controls
                    ctrl.Visible := (index = pageIndex)
        }
        ; 复选框和说明文字是透明背景: 不先擦掉背景就重画, 文字会叠两遍变成粗体
        DllCall("RedrawWindow", "Ptr", hwnd, "Ptr", PreferencesWindow._PageArea(), "Ptr", 0, "UInt", 0x185)   ; RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW
    }

    ; 页面区域 (窗口客户区里, 页面列表右边、底部按钮上面) 的 RECT
    static _PageArea() {
        list := Buffer(16), client := Buffer(16), area := Buffer(16)
        DllCall("GetWindowRect", "Ptr", PreferencesWindow.PageList.Hwnd, "Ptr", list)
        DllCall("MapWindowPoints", "Ptr", 0, "Ptr", PreferencesWindow.Gui.Hwnd, "Ptr", list, "UInt", 2)   ; 屏幕坐标 -> 客户区坐标
        DllCall("GetClientRect", "Ptr", PreferencesWindow.Gui.Hwnd, "Ptr", client)
        NumPut("Int", NumGet(list, 8, "Int"), "Int", 0, "Int", NumGet(client, 8, "Int"), "Int", NumGet(list, 12, "Int"), area)
        return area
    }

    static Close() {
        HotkeyBox.CancelActive()                                            ; 正在录制热键时关窗口: 先结束录制, 恢复 ALTRun 的热键
        if IsObject(PreferencesWindow.Gui)
            PreferencesWindow.Gui.Destroy()
        PreferencesWindow.Gui := "", PreferencesWindow._lists := Map()
    }

    ; "发送到" 时偏好设置开着: 切到 "命令" 页, 把文件 / 文件夹加进列表, 按 确定 / 应用 才保存
    ; (直接写进设置文件, 会被之后的保存盖掉)。和没开偏好设置时一样: 1 个打开编辑框
    ; (已经有这个命令时编辑它), 多个直接加上, 跳过已有的
    static AddCommands(paths) {
        PreferencesWindow.Show("Prefs.Page.Commands")
        list := PreferencesWindow._lists["CustomCommands"]
        if (paths.Length = 1) {
            existing := CustomCommandProvider.FindByTarget(paths[1], list.Items)
            if IsObject(existing) {
                App.Notify(I18n.T("Custom.Exists", existing["Title"]), 2500)
                list.Edit.Call(existing)
            } else {
                list.Insert.Call(CustomCommandProvider.FromPath(paths[1]))
            }
            return
        }
        added := 0, skipped := 0
        for target in paths {
            if IsObject(CustomCommandProvider.FindByTarget(target, list.Items)) {
                skipped += 1
                continue
            }
            list.Append.Call(CustomCommandProvider.FromPath(target))
            added += 1
        }
        App.Toast(I18n.T("Custom.AddedMany", added) (skipped ? " " I18n.T("Custom.SkippedExisting", skipped) : ""), 3000)
    }

    static OK() {
        if PreferencesWindow._dirty
            PreferencesWindow.Save(false)
        else
            PreferencesWindow.Close()
    }

    static Apply() {
        if PreferencesWindow._dirty
            PreferencesWindow.Save(true)
    }

    ; 返回 true 表示没有关闭 (有修改, 选择了继续编辑)
    static Cancel() {
        if (PreferencesWindow._dirty && MsgBox(I18n.T("Prefs.DiscardChanges"), I18n.T("Prefs.Title"), "YesNo Icon? Default2 Owner" PreferencesWindow.Gui.Hwnd) != "Yes")
            return true
        PreferencesWindow.Close()
        return false
    }

    static Help() {
        page := PreferencesWindow.Pages[Max(1, PreferencesWindow._page)]
        ActionCatalog.OpenUrl(HelpProvider.WikiPage(page.Wiki))
    }

    ; 每一页对应的 Wiki 页面 (帮助按钮 / F1)
    static WikiPage(nameKey) {
        static pages := Map("Prefs.Page.Window", "Usage", "Prefs.Page.Appearance", "Themes", "Prefs.Page.Features", "Usage", "Prefs.Page.FileSearch", "File-Search"
                          , "Prefs.Page.Commands", "Commands-and-Snippets", "Prefs.Page.Snippets", "Commands-and-Snippets"
                          , "Prefs.Page.Clipboard", "Commands-and-Snippets", "Prefs.Page.WebSearch", "Usage", "Prefs.Page.Calculator", "Extensions", "Prefs.Page.Scripts", "Extensions"
                          , "Prefs.Page.Hotkeys", "Extensions", "Prefs.Page.QuickSwitch", "Extensions", "Prefs.Page.QSPanel", "Extensions", "Prefs.Page.DateStamp", "Extensions"
                          , "Prefs.Page.FileIndex", "File-Search", "Prefs.Page.Usage", "Usage")
        return pages.Has(nameKey) ? pages[nameKey] : "Configuration"
    }

    ; 命令行 "-Preferences 页码 [x y]" (应用 之后重新打开) -> {Page, X, Y}
    static ParseArgs(args) {
        value(i) => (args.Length >= i && IsInteger(args[i])) ? Integer(args[i]) : ""
        page := value(2)
        return {Page: (page = "") ? 1 : page, X: value(3), Y: value(4)}
    }

    ; reopen: 应用 = 重新载入后在同一页、同一位置重新打开; 确定 = 只在后台重新载入
    static Save(reopen := true) {
        for bind in PreferencesWindow.Binds
            PreferencesWindow.SetPath(PreferencesWindow.Working, bind.Path, bind.Read.Call())
        general := PreferencesWindow.Working["General"]
        appearance := PreferencesWindow.Working["Appearance"]
        appearance["VisibleRows"] := Max(1, Min(appearance["VisibleRows"], 9))
        appearance["Width"] := Max(400, appearance["Width"])
        if (general["Hotkey"] = "")
            general["Hotkey"] := "!Space"
        clash := PreferencesWindow.DuplicateHotkey(PreferencesWindow._GlobalHotkeys(PreferencesWindow.Working))
        if (IsObject(clash) && MsgBox(I18n.T("Prefs.HotkeyClash", HotkeyBox.Label(clash.Key), clash.First, clash.Second)
                , I18n.T("Prefs.Title"), "YesNo Icon! Default2 Owner" PreferencesWindow.Gui.Hwnd) != "Yes")
            return
        if !PreferencesWindow._positionReset                                ; 打开偏好设置之后拖动过搜索窗口: 用新的位置
            appearance["Position"] := PreferencesWindow.DeepCopy(AppSettings.Appearance["Position"])

        AppSettings.Data := PreferencesWindow.Working
        if !AppSettings.Save()
            return
        page := PreferencesWindow._page
        WinGetPos(&x, &y, , , PreferencesWindow.Gui.Hwnd)
        PreferencesWindow.Close()
        App.Restart(reopen ? "-Preferences " page " " x " " y : "-Reloaded")
    }

    ;---------------------------------------------------------------------------
    ; Pages
    ;---------------------------------------------------------------------------
    static _BuildPages() {
        PreferencesWindow._BuildGeneral()
        PreferencesWindow._BuildWindow()
        PreferencesWindow._BuildAppearance()
        PreferencesWindow._BuildFeatures()
        PreferencesWindow._BuildApplications()
        PreferencesWindow._BuildFileSearch()
        PreferencesWindow._BuildFileIndex()
        PreferencesWindow._BuildCommands()
        PreferencesWindow._BuildSnippets()
        PreferencesWindow._BuildClipboard()
        PreferencesWindow._BuildWebSearch()
        PreferencesWindow._BuildCalculator()
        PreferencesWindow._BuildScripts()
        PreferencesWindow._BuildHotkeys()
        PreferencesWindow._BuildQuickSwitch()
        PreferencesWindow._BuildQuickSwitchPanel()
        PreferencesWindow._BuildDateStamp()
        PreferencesWindow._BuildUsage()
        PreferencesWindow._BuildAdvanced()
    }

    static _BuildGeneral() {
        PreferencesWindow._BeginPage("Prefs.Page.General", 120)
        PreferencesWindow._Pair(["General.Hotkey", "Prefs.Hotkey", "K", "hotkey"], ["General.SecondaryHotkey", "Prefs.SecondaryHotkey", "K", "hotkey"])
        PreferencesWindow._Field("General.SelectionHotkey", "Prefs.SelectionHotkey", "K", "hotkey")
        PreferencesWindow._Pair(["General.DoubleTap", "Prefs.DoubleTap", "K", "choice", ["", "Ctrl", "Shift"], [I18n.T("Prefs.DoubleTap.None"), I18n.T("Prefs.DoubleTap.Ctrl"), I18n.T("Prefs.DoubleTap.Shift")]]
            , ["General.Language", "Prefs.Language", "K", "choice", PreferencesWindow._LanguageValues(), PreferencesWindow._LanguageLabels()])
        PreferencesWindow._Gap()
        for row in [["LaunchAtLogin", "Prefs.LaunchAtLogin", "Prefs.Group.Startup"], ["ShowTrayIcon", "Prefs.ShowTrayIcon", ""]
                   , ["SendToMenu", "Prefs.SendToMenu", "Prefs.Group.Integration"], ["StartMenuShortcut", "Prefs.StartMenu", ""]
                   , ["CheckForUpdates", "Prefs.CheckUpdates", "Prefs.Group.Updates"], ["SaveLog", "Prefs.SaveLog", "Prefs.Group.Diagnostics"]]
            PreferencesWindow._Check("General." row[1], row[2], , , row[3])
        PreferencesWindow._Gap()
        PreferencesWindow._Field("General.FileManager", "Prefs.FileManager", "L", "file")
    }

    ; 搜索窗口的行为 (设置仍然在 General 里, 只是分到单独一页)
    static _BuildWindow() {
        PreferencesWindow._BeginPage("Prefs.Page.Window", 16)                ; 分节排列, 每节的内容缩进一点
        PreferencesWindow._Section("Prefs.Group.Window")
        PreferencesWindow._Check("General.HideOnDeactivate", "Prefs.HideOnDeactivate")
        PreferencesWindow._Check("General.KeepLastQuery", "Prefs.KeepLastQuery")
        PreferencesWindow._Section("Prefs.Group.Typing")
        english := PreferencesWindow._Check("General.SwitchToEnglishInput", "Prefs.EnglishInput")
        restore := PreferencesWindow._Check("General.RestoreInput", "Prefs.RestoreInput", PreferencesWindow._InputX() + 18)   ; 子选项, 缩进
        restore.Enabled := english.Value
        english.OnEvent("Click", (*) => restore.Enabled := english.Value)
        PreferencesWindow._Check("General.SpaceToRun", "Prefs.SpaceToRun")
        PreferencesWindow._Section("Prefs.Section.EmptyBox")
        PreferencesWindow._Check("General.ShowTips", "Prefs.ShowTips")
        PreferencesWindow._Gap()
        PreferencesWindow._InlineField("Features.Recent.RecentCount", "Prefs.RecentCount", "S", "number")
        PreferencesWindow._Section("Prefs.Section.History")
        PreferencesWindow._InlineField("General.HistorySize", "Prefs.HistorySize", "S", "number")
    }

    static _BuildAppearance() {
        PreferencesWindow._BeginPage("Prefs.Page.Appearance", 150)
        PreferencesWindow._Section("Prefs.Section.Theme")
        themes := ThemeManager.Names()
        gallery := PreferencesWindow._ThemeGallery("Appearance.Theme", themes, 188)
        PreferencesWindow._Buttons("Prefs.Group.CustomThemes"
            , ["Prefs.CopyTheme", (*) => PreferencesWindow._CopyTheme(gallery, themes)]
            , ["Prefs.OpenThemes", (*) => PreferencesWindow._OpenFolder(ThemeManager.UserDir)])
        PreferencesWindow._Section("Prefs.Section.Size")
        PreferencesWindow._Pair(["Appearance.Width", "Prefs.Width", "S", "number"], ["Appearance.VisibleRows", "Prefs.VisibleRows", "S", "number"])
        PreferencesWindow._Section("Prefs.WindowPosition")
        PreferencesWindow._Choice("Appearance.ShowOn", "Prefs.ShowOn", ["Mouse", "Primary", "Active"]
            , [I18n.T("Prefs.ShowOn.Mouse"), I18n.T("Prefs.ShowOn.Primary"), I18n.T("Prefs.ShowOn.Active")], "B")
        PreferencesWindow._Check("Appearance.RememberPosition", "Prefs.RememberPosition", , , "Prefs.Group.Dragging")
        PreferencesWindow._Button("Prefs.ResetPosition", (button, *) => PreferencesWindow._ResetPosition(button))
    }

    ; 记住的位置恢复为默认 (保存时生效)
    static _ResetPosition(button) {
        PreferencesWindow.Working["Appearance"]["Position"] := AppSettings.Defaults()["Appearance"]["Position"]
        PreferencesWindow._positionReset := true
        PreferencesWindow.MarkDirty()
        button.Enabled := false
    }

    ; 界面语言的选项: 自动 + Resources\Lang 里的语言 (每种语言用它自己的文字显示)
    static _LanguageValues() {
        values := ["auto"]
        for language in I18n.Languages()
            values.Push(language[1])
        return values
    }

    static _LanguageLabels() {
        labels := [I18n.T("Prefs.Language.auto")]
        for language in I18n.Languages()
            labels.Push(language[2])
        return labels
    }

    ; 内置主题显示翻译后的名称, 用户主题显示文件名
    static _ThemeLabel(themeName) {
        label := I18n.T("Theme." themeName)
        return (label = "Theme." themeName) ? themeName : label
    }

    ; 主题列表: 每个主题一张缩略图 (ThemePreview, 按主题的颜色画的迷你搜索窗口), 像 Alfred / Listary 一样点选。
    ; 用 ListView 的大图标视图: 自带滚动条、方向键选择, 主题再多也放得下。返回 ListView (行号 = themes 里的序号)
    static _ThemeGallery(path, themes, height) {
        thumbW := Win.Scale(112), thumbH := Win.Scale(76)
        list := PreferencesWindow._Add("ListView", "x" PreferencesWindow.ContentX " w" PreferencesWindow.ContentW " h" height
            . " Icon -Multi -Hdr", ["Theme"])
        DllCall("uxtheme\SetWindowTheme", "Ptr", list.Hwnd, "WStr", "Explorer", "Ptr", 0)   ; 选中项画浅色方框, 不把缩略图染成蓝色
        list.SetImageList(ThemePreview.ImageList(themes, thumbW, thumbH), 0)   ; 0 = 大图标
        SendMessage(0x1035, 0, (thumbW + Win.Scale(16)) | ((thumbH + Win.Scale(30)) << 16), list.Hwnd)   ; LVM_SETICONSPACING
        current := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        selected := 1
        for index, themeName in themes {
            list.Add("Icon" index, PreferencesWindow._ThemeLabel(themeName))
            if (themeName = current)
                selected := index
        }
        list.Modify(selected, "Select Focus Vis")
        list.OnNotify(-101, (ctrl, lParam) => PreferencesWindow._OnGalleryChange(lParam))   ; LVN_ITEMCHANGED (程序里选中的也算)
        PreferencesWindow._Bind(path, () => themes[list.GetNext() || selected])
        PreferencesWindow._Below(8, list)
        return list
    }

    ; 主题列表里选中了另一个主题: "应用" 可用
    static _OnGalleryChange(lParam) {
        static LVIS_SELECTED := 0x2
        offset := A_PtrSize * 3 + 4                                         ; NMLISTVIEW: hdr, iItem, iSubItem, uNewState
        newState := NumGet(lParam, offset + 4, "UInt"), oldState := NumGet(lParam, offset + 8, "UInt")
        if ((newState & LVIS_SELECTED) && !(oldState & LVIS_SELECTED))
            PreferencesWindow.MarkDirty()
    }

    ; 把选中的主题 (展开成完整的键) 复制到 Themes\<新名称>.json, 用记事本打开, 并在列表里选中它
    static _CopyTheme(gallery, themes) {
        source := themes[gallery.GetNext() || 1]
        if (source = "System")
            source := ThemeManager.SystemUsesDark() ? "Dark" : "Light"
        theme := ThemeManager.Resolve(source)
        if !IsObject(theme)
            return
        PreferencesWindow.Gui.Opt("+OwnDialogs")
        answer := InputBox(I18n.T("Prefs.CopyThemePrompt"), I18n.T("Prefs.CopyTheme"), "w340 h130", source " Custom")
        themeName := Trim(RegExReplace(answer.Value, '[\\/:*?"<>|]'))
        if (answer.Result != "OK" || themeName = "")
            return
        themeFile := ThemeManager.UserDir "\" themeName ".json"
        if (FileExist(themeFile) && MsgBox(I18n.T("Prefs.ThemeExists", themeName), App.Name, "YesNo Icon! Default2") != "Yes")
            return
        try {
            JSON.WriteFile(themeFile, theme)
        } catch as e {
            return MsgBox(e.Message, App.Name, 48)
        }
        index := 0
        for existing, value in themes
            if (value = themeName)
                index := existing
        if !index {
            themes.Push(themeName)
            imageList := SendMessage(0x1002, 0, 0, gallery.Hwnd)           ; LVM_GETIMAGELIST (LVSIL_NORMAL)
            DllCall("comctl32\ImageList_GetIconSize", "Ptr", imageList, "Int*", &thumbW := 0, "Int*", &thumbH := 0)
            gallery.Add("Icon" (ThemePreview.AddTo(imageList, themeName, thumbW, thumbH) + 1), themeName)
            index := themes.Length
        }
        gallery.Modify(0, "-Select")
        gallery.Modify(index, "Select Focus Vis")
        Run('notepad.exe "' themeFile '"')
    }

    static _BuildFeatures() {
        PreferencesWindow._BeginPage("Prefs.Page.Features", 140)
        PreferencesWindow._Section("Prefs.EnabledFeatures")
        features := ["Applications", "CustomCommands", "Snippets", "Clipboard", "Calculator", "WebSearch", "Bookmarks", "Windows", "Recent", "FileSearch", "Scripts", "Terminal", "System", "Help"]
        startY := PreferencesWindow._y, columnW := PreferencesWindow._InputW() // 2      ; 两列: 英文名称较长, 三列会换行
        for index, feature in features {
            column := Mod(index - 1, 2), row := (index - 1) // 2
            PreferencesWindow._y := startY + row * 20
            PreferencesWindow._Check("Features." feature ".Enabled", "Prefs.Feature." feature
                , PreferencesWindow._InputX() + column * columnW, , (index = 1) ? "Prefs.Group.SearchFeatures" : "", columnW)
        }
        PreferencesWindow._y := startY + (features.Length // 2) * 20          ; 最后一格: Windows 设置的页面 (系统命令的一部分)
        PreferencesWindow._Check("Features.System.SettingsPages", "Prefs.SettingsPages", PreferencesWindow._InputX() + columnW, , "", columnW)
        PreferencesWindow._y := startY + Ceil((features.Length + 1) / 2) * 20
        PreferencesWindow._Section("Prefs.Section.FeatureOptions")
        PreferencesWindow._Check("Features.System.ConfirmActions", "Prefs.ConfirmActions", , , "Prefs.Feature.System")
        PreferencesWindow._Lines("Features.System.Hidden", "Prefs.SysHidden", 2)
        PreferencesWindow._Gap()
        PreferencesWindow._Pair(["Features.Terminal.Prefix", "Prefs.TerminalPrefix", "S"]
            , ["Features.Terminal.Shell", "Prefs.TerminalShell", "M", "choice", ["cmd", "powershell", "pwsh", "wt"], ["Command Prompt (cmd)", "Windows PowerShell", "PowerShell 7 (pwsh)", "Windows Terminal (wt)"]])
    }

    static _BuildApplications() {
        PreferencesWindow._BeginPage("Prefs.Page.Applications", 130)
        PreferencesWindow._Section("Prefs.Section.AppIndex")
        PreferencesWindow._Lines("Features.Applications.Folders", "Prefs.AppFolders", 4)
        PreferencesWindow._Csv("Features.Applications.FileTypes", "Prefs.AppFileTypes", "L")
        PreferencesWindow._Field("Features.Applications.Exclude", "Prefs.AppExclude", "L")
        PreferencesWindow._Pair(["Features.Applications.Depth", "Prefs.AppDepth", "S", "number"], ["Features.Applications.RefreshMinutes", "Prefs.RefreshMinutes", "S", "number"])
        PreferencesWindow._Button("Prefs.RebuildIndex", (*) => App.RebuildIndex())
        PreferencesWindow._Section("Prefs.Section.AppResults")
        PreferencesWindow._Lines("Features.Applications.Hidden", "Prefs.AppHidden", 4)
        PreferencesWindow._Check("Features.Applications.StoreApps", "Prefs.StoreApps", , , "Prefs.Group.Options")
        PreferencesWindow._Check("Features.Applications.MatchPinyin", "Prefs.MatchPinyin")
    }

    static _BuildFileSearch() {
        PreferencesWindow._BeginPage("Prefs.Page.FileSearch", 130)
        PreferencesWindow._Section("Prefs.Section.FileStart")
        PreferencesWindow._Check("Features.FileSearch.SpacePrefix", "Prefs.SpacePrefix", , , "Prefs.Group.StartFileSearch")
        PreferencesWindow._Check("Features.FileSearch.QuotePrefix", "Prefs.QuotePrefix")
        PreferencesWindow._Pair(["Features.FileSearch.Keywords", "Prefs.FileKeywords", "M", "csv"], ["Features.FileSearch.FolderKeywords", "Prefs.FolderKeywords", "S", "csv"])
        PreferencesWindow._Lines("Features.FileSearch.TypeFilters", "Prefs.TypeFilters", 7)
        PreferencesWindow._Section("Prefs.Section.FileResults")
        PreferencesWindow._Check("Features.FileSearch.InDefaultResults", "Prefs.FileInDefault", , , "Prefs.Group.NormalSearch")
        PreferencesWindow._Pair(["Features.FileSearch.MaxResults", "Prefs.FileMaxResults", "S", "number"], ["Features.FileSearch.DefaultResultsLimit", "Prefs.FileDefaultLimit", "S", "number"])
        PreferencesWindow._Field("Features.FileSearch.MinQueryLength", "Prefs.FileMinLength", "S", "number")
        PreferencesWindow._Section("Prefs.Section.Browse")                  ; 浏览文件夹总是可用, 只有说明
    }

    ; 搜索来源: Everything (在运行时用它搜全盘) 和 Everything 没有运行时的内置文件索引
    static _BuildFileIndex() {
        PreferencesWindow._BeginPage("Prefs.Page.FileIndex", 130)
        PreferencesWindow._Section("Prefs.Section.Everything")
        status := I18n.T(Everything.IsRunning() ? "Prefs.Running" : "Prefs.NotRunning")
        PreferencesWindow._Check("Features.FileSearch.UseEverything", "Prefs.UseEverything", , I18n.T("Prefs.EverythingStatus", status))
        PreferencesWindow._Field("Features.FileSearch.EverythingFilter", "Prefs.EverythingFilter", "L")
        PreferencesWindow._Field("Features.FileSearch.EverythingPath", "Prefs.EverythingPath", "L", "folder")
        PreferencesWindow._Section("Prefs.Section.BuiltinIndex")
        PreferencesWindow._Lines("Features.FileSearch.ScopeFolders", "Prefs.ScopeFolders", 4)
        depthY := PreferencesWindow._y                                      ; 按钮和 "子文件夹深度" 放在同一行
        PreferencesWindow._Field("Features.FileSearch.ScopeDepth", "Prefs.ScopeDepth", "S", "number")
        PreferencesWindow._SideButton(depthY, "Prefs.RebuildFileIndex", (*) => FileIndex.Rebuild())
        PreferencesWindow._Field("Features.FileSearch.ScopeExclude", "Prefs.ScopeExclude", "L")
        PreferencesWindow._Pair(["Features.FileSearch.MaxEntries", "Prefs.MaxEntries", "S", "number"], ["Features.FileSearch.RefreshMinutes", "Prefs.RefreshMinutes", "S", "number"])
    }

    static _BuildCommands() {
        PreferencesWindow._BeginPage("Prefs.Page.Commands")
        PreferencesWindow._List("CustomCommands", 540
            , [["Prefs.Col.Title", "Title", 140], ["Prefs.Col.Type", "Type", 70], ["Prefs.Col.Target", "Target", 190], ["Prefs.Col.Keyword", "Keyword", 70], ["Prefs.Col.Status", "Status", 70]]
            , CustomCommandProvider.EditorFields(), () => CustomCommandProvider.NewCommand()
            , (item, roots) => CustomCommandProvider.CheckTarget(item, roots))
        PreferencesWindow._Note("Prefs.CommandsNote")
    }

    static _BuildSnippets() {
        PreferencesWindow._BeginPage("Prefs.Page.Snippets", 150)
        PreferencesWindow._List("Snippets", 220
            , [["Prefs.Col.Name", "Name", 150], ["Prefs.Col.Keyword", "Keyword", 90], ["Prefs.Col.Text", "Text", 300]]
            , SnippetProvider.EditorFields(), () => SnippetProvider.NewSnippet())
        PreferencesWindow._Section("Prefs.Section.Options")
        PreferencesWindow._Check("Features.Snippets.SearchText", "Prefs.SnippetSearchText", , , "Prefs.Group.Search")
        PreferencesWindow._Field("Features.Snippets.Keyword", "Prefs.SnippetKeyword", "M")
        PreferencesWindow._Check("Features.Snippets.AutoExpand", "Prefs.SnippetAutoExpand", , , "Prefs.Group.AutoExpand")
        PreferencesWindow._Field("Features.Snippets.ExpandPrefix", "Prefs.ExpandPrefix", "S")
        PreferencesWindow._Field("Features.Snippets.ExpandExclude", "Prefs.ExpandExclude", "L")
        PreferencesWindow._Pair(["Features.Snippets.PasteMode", "Prefs.PasteMode", "M", "choice", ["Clipboard", "Type"], [I18n.T("Prefs.PasteMode.Clipboard"), I18n.T("Prefs.PasteMode.Type")]]
            , ["Features.Snippets.PasteDelay", "Prefs.PasteDelay", "S", "number"])
    }

    static _BuildClipboard() {
        PreferencesWindow._BeginPage("Prefs.Page.Clipboard", 150)
        PreferencesWindow._Field("Features.Clipboard.Hotkey", "Prefs.ClipHotkey", "K", "hotkey")
        PreferencesWindow._Field("Features.Clipboard.Keyword", "Prefs.ClipKeyword", "M")
        PreferencesWindow._Pair(["Features.Clipboard.MaxItems", "Prefs.ClipMaxItems", "S", "number"], ["Features.Clipboard.MaxItemLength", "Prefs.ClipMaxLength", "S", "number"])
        PreferencesWindow._Check("Features.Clipboard.Persist", "Prefs.ClipPersist", , , "Prefs.Group.History")
        PreferencesWindow._Check("Features.Clipboard.Images", "Prefs.ClipImages")
        PreferencesWindow._Field("Features.Clipboard.MaxImages", "Prefs.ClipMaxImages", "S", "number")
        PreferencesWindow._Check("Features.Clipboard.MergeDoubleCopy", "Prefs.ClipMerge", , , "Prefs.Group.Copy")
        PreferencesWindow._Gap()
        PreferencesWindow._Lines("Features.Clipboard.IgnoreApps", "Prefs.ClipIgnoreApps", 4)
        PreferencesWindow._Button("Prefs.ClipClear", (*) => ClipboardProvider.Clear())
    }

    static _BuildWebSearch() {
        PreferencesWindow._BeginPage("Prefs.Page.WebSearch", 130)
        PreferencesWindow._List("Features.WebSearch.Engines", 320
            , [["Prefs.Col.Keyword", "Keyword", 70], ["Prefs.Col.Title", "Title", 130], ["Prefs.Col.Url", "Url", 340]]
            , WebSearchProvider.EditorFields(), () => WebSearchProvider.NewEngine())
        PreferencesWindow._Csv("Features.WebSearch.Fallbacks", "Prefs.Fallbacks", "L")
    }

    ; 计算器: 货币换算, 结构计算 (梁主筋 / 配筋面积) 和它的参数
    static _BuildCalculator() {
        base := "Features.Calculator."
        PreferencesWindow._BeginPage("Prefs.Page.Calculator", 180)
        PreferencesWindow._Section("Prefs.Section.Calculator")
        PreferencesWindow._Check(base "Currency", "Prefs.Currency")
        PreferencesWindow._Section("Prefs.Section.Structural")
        PreferencesWindow._Check(base "StructuralCalc", "Prefs.StructuralCalc")
        PreferencesWindow._Field(base "RebarCover", "Prefs.RebarCover", "S", "number")      ; 每项一行: 标签和输入框各自对齐
        PreferencesWindow._Field(base "MaxBarSpacing", "Prefs.MaxBarSpacing", "S", "number")
        PreferencesWindow._Csv(base "BarSizes", "Prefs.BarSizes", "M")
        PreferencesWindow._Field(base "BarPrefix", "Prefs.BarPrefix", "S")
    }

    ; 脚本扩展: 说明、Scripts 文件夹、找到的脚本 (双击用记事本编辑); 开关在 "功能" 页
    static _BuildScripts() {
        PreferencesWindow._BeginPage("Prefs.Page.Scripts", 110)
        PreferencesWindow._Section("Prefs.Section.Scripts")
        PreferencesWindow._Info("Prefs.Group.ScriptsFolder", ScriptProvider.Dir, " cGray")
        PreferencesWindow._y += 4
        PreferencesWindow._Button("Prefs.ScriptsFolder", (*) => (ScriptProvider.OpenFolder(), PreferencesWindow._FillScripts()))
        PreferencesWindow._Section("Prefs.Section.ScriptList")
        listView := PreferencesWindow._Add("ListView", "w" PreferencesWindow.ContentW " h170 -Multi NoSort Grid ReadOnly"
            , [I18n.T("Prefs.Col.Name"), I18n.T("Prefs.Col.Keyword"), I18n.T("Prefs.Col.Argument"), I18n.T("Prefs.Col.Mode"), I18n.T("Prefs.Col.File")])
        for index, width in [145, 70, 95, 135, 100]
            listView.ModifyCol(index, width)
        listView.OnEvent("DoubleClick", (ctrl, row) => (row && row <= ScriptProvider.Scripts.Length) ? Run('notepad.exe "' ScriptProvider.Scripts[row].Path '"') : "")
        PreferencesWindow._scriptList := listView
        PreferencesWindow._FillScripts()
        PreferencesWindow._Below(6, listView)
        PreferencesWindow._Note("Prefs.ScriptsNote")
    }

    static _scriptList := ""
    static _FillScripts() {
        static modes := Map("window", "Script.Mode.Window", "silent", "Script.Mode.Silent", "output", "Script.Mode.Output")
        listView := PreferencesWindow._scriptList
        if !IsObject(listView)
            return
        ScriptProvider.Load()
        listView.Delete()
        for script in ScriptProvider.Scripts
            listView.Add(, script.Title, script.Keyword, script.Argument, I18n.T(modes[script.Mode]), script.Name)
        if !ScriptProvider.Scripts.Length
            listView.Add(, I18n.T("Prefs.NoScripts"))
    }

    static _BuildHotkeys() {
        PreferencesWindow._BeginPage("Prefs.Page.Hotkeys", 100)                ; 标签列窄一些: 下面的列表从最左边开始, 上面的选项也靠左
        PreferencesWindow._Choice("General.CapsLock", "Prefs.CapsLock", ["", "Layout", "Mode"]     ; 按 CapsLock 切换输入法 (CapsLockSwitch)
            , [I18n.T("Prefs.CapsLock.Off"), I18n.T("Prefs.CapsLock.Layout"), I18n.T("Prefs.CapsLock.Mode")])
        PreferencesWindow._Gap()
        actions := [["ToggleWindow", "Show / hide ALTRun (ToggleWindow)"]]
        for command in SystemProvider.Commands()
            actions.Push([command["Id"], command["Title"] " (" command["Id"] ")"])
        PreferencesWindow._List("Hotkeys", 430
            , [["Prefs.Col.Key", "Key", 160], ["Prefs.Col.Action", "Action", 150], ["Prefs.Col.WinTitle", "WinTitle", 230]]
            , [ItemEditor.Field("Key", "Prefs.Col.Key", "hotkey", true, "", I18n.T("Prefs.HotkeyHint"))
             , ItemEditor.Field("Action", "Prefs.Col.Action", "choice", false, actions)
             , ItemEditor.Field("WinTitle", "Prefs.Col.WinTitle", "text", false, "", I18n.T("Prefs.WinTitleHint"))]
            , () => Map("Key", "", "Action", "ToggleWindow", "WinTitle", ""))
        PreferencesWindow._Note("Prefs.HotkeysNote")
    }

    ; 对话框快速跳转 (Listary 的 Quick Switch)
    static _BuildQuickSwitch() {
        base := "Extensions.QuickSwitch."
        PreferencesWindow._BeginPage("Prefs.Page.QuickSwitch", 170)
        PreferencesWindow._Section("Prefs.QuickSwitch")
        PreferencesWindow._Check(base "Enabled", "Prefs.EnableExtension", , , "Prefs.Group.Status")
        PreferencesWindow._Pair([base "TotalCmdHotkey", "Prefs.QSTotalCmd", "K", "hotkey"], [base "ExplorerHotkey", "Prefs.QSExplorer", "K", "hotkey"])

        ; DialogWindows 拆成 "标准对话框" 复选框 + 其它对话框的列表, 保存时再合成一个列表
        PreferencesWindow._Section("Prefs.Section.QSDialogs")
        windows := PreferencesWindow.SplitWindows(PreferencesWindow.GetPath(PreferencesWindow.Working, base "DialogWindows"))
        standardClass := "ahk_class #32770", others := []
        for window in windows
            if (window != standardClass)
                others.Push(window)
        PreferencesWindow._GroupLabel("Prefs.Group.DialogWindows")
        x := PreferencesWindow._InputX()
        standard := PreferencesWindow._Add("Checkbox", "x" x " w" (PreferencesWindow.ContentX + PreferencesWindow.ContentW - x), I18n.T("Prefs.QSStandard"))
        standard.Value := (others.Length < windows.Length) ? 1 : 0
        PreferencesWindow._Below(0, standard)                               ; 说明写在复选框的文字里 (这一页放了三个 4 行的列表)
        otherList := PreferencesWindow._WinList("", "Prefs.QSOtherDialogs", 4, others)
        PreferencesWindow._Bind(base "DialogWindows", () => PreferencesWindow.JoinWindows(standard.Value ? [standardClass] : [], PreferencesWindow.SplitLines(otherList.Value)))
        PreferencesWindow._WinList(base "ExcludeWindows", "Prefs.QSExclude", 4)

        PreferencesWindow._Section("Prefs.Section.QSAuto")
        PreferencesWindow._Check(base "AutoSwitch", "Prefs.QSAuto", , , "Prefs.Group.Options")
        PreferencesWindow._WinList(base "AutoSwitchExclude", "Prefs.QSAutoExclude", 4)
    }

    ; 对话框面板 (QuickSwitch 的一部分): 打开 / 保存对话框下面的文件夹列表和搜索框, 以及键盘用的文件夹菜单
    static _BuildQuickSwitchPanel() {
        base := "Extensions.QuickSwitch."
        PreferencesWindow._BeginPage("Prefs.Page.QSPanel", 150)
        PreferencesWindow._Section("Prefs.QSPanelSection")
        PreferencesWindow._Check(base "ShowPanel", "Prefs.QSPanel", , , "Prefs.Group.Status")
        PreferencesWindow._Choice(base "PanelSearch", "Prefs.QSPanelSearch", ["folders", "all"], [I18n.T("Prefs.QSPanelSearch.Folders"), I18n.T("Prefs.QSPanelSearch.All")], "K")
        PreferencesWindow._Field(base "RecentFolders", "Prefs.QSRecent", "S", "number")
        PreferencesWindow._Section("Prefs.QSMenuSection")
        PreferencesWindow._Field(base "MenuHotkey", "Prefs.QSMenu", "K", "hotkey")
    }

    ; 一键加日期 (AutoDate): 重命名文件 / 备注框里各自的热键和生效的窗口
    static _BuildDateStamp() {
        base := "Extensions.AutoDate."
        PreferencesWindow._BeginPage("Prefs.Page.DateStamp", 150)
        PreferencesWindow._Section("Prefs.AutoDate")
        PreferencesWindow._Check(base "Enabled", "Prefs.EnableExtension", , , "Prefs.Group.Status")
        PreferencesWindow._Field(base "DateFormat", "Prefs.DateFormat", "M")
        PreferencesWindow._Section("Prefs.Section.DateRename")
        PreferencesWindow._Field(base "RenameHotkey", "Prefs.RenameHotkey", "K", "hotkey")
        PreferencesWindow._WinList(base "RenameWindows", "Prefs.RenameWindows", 4)
        PreferencesWindow._Section("Prefs.Section.DateAppend")
        PreferencesWindow._Field(base "AppendHotkey", "Prefs.AppendHotkey", "K", "hotkey")
        PreferencesWindow._WinList(base "AppendWindows", "Prefs.AppendWindows", 4)
    }

    ; 上面是 "关于" (图标、名称、版本、主页、检查更新), 下面是设置和数据文件的位置和操作
    static _BuildAdvanced() {
        PreferencesWindow._BeginPage("Prefs.Page.Advanced", 110)
        textX := PreferencesWindow._InputX(), textW := PreferencesWindow._InputW(), top := PreferencesWindow._y + 6
        PreferencesWindow._y := top                                         ; 图标在左列, 文字和按钮和下面的控件列对齐
        if FileExist(App.IconFile)                                         ; 图标文件不在时 (例如只复制了 exe) 用 exe 里的图标
            PreferencesWindow._Add("Picture", "x" (textX - 64 - 18) " w64 h64", App.IconFile)
        else if A_IsCompiled
            PreferencesWindow._Add("Picture", "x" (textX - 64 - 18) " w64 h64 Icon1", A_ScriptFullPath)
        PreferencesWindow.Gui.SetFont("s15 bold")
        PreferencesWindow._Add("Text", "x" textX " w" textW, App.Name)
        PreferencesWindow.Gui.SetFont("s9 norm")
        PreferencesWindow._y := top + 32
        PreferencesWindow._Add("Text", "x" textX " w" textW, I18n.T("App.Tagline"))
        PreferencesWindow._y := top + 52
        copyright := "© 2013–" SubStr(App.Version, 1, 4) " zhugecaomao"         ; 版本号以发布年份开头
        PreferencesWindow._Add("Text", "x" textX " w" textW " cGray", I18n.T("Prefs.Version", App.Version) "   ·   " copyright "   ·   GPL-3.0")
        PreferencesWindow._y := top + 74
        links := ""
        for link in [[App.Website, I18n.T("Prefs.Homepage")], [App.RepoUrl, "GitHub"], [App.RepoUrl "/releases", I18n.T("Update.ReleaseNotes")]
                   , [App.RepoUrl "/issues/new/choose", I18n.T("Prefs.ReportIssue")]]
            links .= (links = "" ? "" : "   ·   ") '<a href="' link[1] '">' link[2] '</a>'
        PreferencesWindow._Add("Link", "x" textX " w" textW, links)
        PreferencesWindow._y := top + 102
        PreferencesWindow._Add("Button", "x" textX " w" PreferencesWindow.ButtonW " h" PreferencesWindow.ButtonH, I18n.T("Tray.CheckUpdate"))
            .OnEvent("Click", (*) => UpdateChecker.Check())
        PreferencesWindow._y := top + 102 + PreferencesWindow.ButtonH + 22

        PreferencesWindow._Section("Prefs.Section.Data")
        ; 设置文件就是数据文件夹里的 ALTRun.json, 不另写一行; 路径太长时中间用省略号 (SS_PATHELLIPSIS), 鼠标停留显示完整路径
        folder := PreferencesWindow._Info("Prefs.Group.DataFolder", AppSettings.DataDir, " r1 cGray 0x8100")
        Win.AddTooltip(folder.Hwnd, AppSettings.DataDir)
        PreferencesWindow._y += 4
        PreferencesWindow._Buttons(""
            , ["Prefs.EditJson", (*) => (PreferencesWindow.Cancel() || App.EditSettingsFile())]
            , ["Prefs.OpenDataFolder", (*) => PreferencesWindow._OpenFolder(AppSettings.DataDir)]
            , ["Prefs.ChangeDataFolder", (*) => PreferencesWindow._ChangeDataFolder()])

        PreferencesWindow._Section("Prefs.Section.Reset")
        PreferencesWindow._Button("Prefs.ResetLearning", (*) => PreferencesWindow._ResetLearning())
    }

    ; 使用统计 (和 Alfred 的 Usage 一样): 合计, 最近 30 天每天的柱状图, 每个功能的次数
    static _usage := ""
    static _BuildUsage() {
        PreferencesWindow._BeginPage("Prefs.Page.Usage")
        PreferencesWindow._Section("Usage.Title")
        w := PreferencesWindow.ContentW, x := PreferencesWindow.ContentX
        totals := PreferencesWindow._Add("Text", "w" w)
        PreferencesWindow._Below(4, totals)
        since := PreferencesWindow._Add("Text", "w" w " cGray")
        PreferencesWindow._Below(12, since)
        top := PreferencesWindow._y, barW := 14, step := 18, chartH := 100
        bars := []
        Loop 30 {
            PreferencesWindow._y := top
            bars.Push(PreferencesWindow._Add("Progress", "x" (x + (A_Index - 1) * step) " w" barW " h" chartH " Vertical c3B82F6 BackgroundE6EBF2"))
        }
        PreferencesWindow._y := top + chartH + 4
        PreferencesWindow.Gui.SetFont("s8")
        first := PreferencesWindow._Add("Text", "x" x " w100 cGray")
        chart := PreferencesWindow._Add("Text", "x" (x + 100) " w" (30 * step - barW - 200 + 10) " Center cGray")
        PreferencesWindow._Add("Text", "x" (x + 30 * step - barW - 90) " w" (90 + barW - 4) " Right cGray", I18n.T("Usage.Today"))
        PreferencesWindow.Gui.SetFont("s9")
        PreferencesWindow._y += 22
        headers := []
        for key in ["Feature", "Today", "Week", "Month", "All", "Share"]
            headers.Push(I18n.T("Usage.Col." key))
        list := PreferencesWindow._Add("ListView", "w" w " h272 -Multi NoSort Grid ReadOnly", headers)
        for index, width in [178, 66, 66, 66, 80, 80]
            list.ModifyCol(index, width (index > 1 ? " Right" : ""))
        PreferencesWindow._Below(8, list)
        PreferencesWindow._Button("Usage.Clear", (*) => PreferencesWindow._ClearUsage())
        PreferencesWindow._usage := {Totals: totals, Since: since, Bars: bars, First: first, Chart: chart, List: list}
        PreferencesWindow._RefreshUsage()
    }

    static _RefreshUsage() {
        ui := PreferencesWindow._usage
        num := (n) => Calc.Thousands(n)
        periods := [Usage.Summary(1), Usage.Summary(7), Usage.Summary(30), Usage.Summary(0)]
        ui.Totals.Value := I18n.T("Usage.Totals", num(Usage.Total(periods[1])), num(Usage.Total(periods[2])), num(Usage.Total(periods[3])), num(Usage.Total(periods[4])))
        all := periods[4]
        ui.Since.Value := I18n.T("Usage.Since", (Usage.Since != "") ? Usage.Since : FormatTime(, "yyyy-MM-dd"), num(all.Has("Show") ? all["Show"] : 0))
        daily := Usage.Daily(30)
        most := 0
        for day in daily
            most := Max(most, day[2])
        for index, bar in ui.Bars {
            bar.Opt("Range0-" Max(most, 1))
            bar.Value := daily[index][2]
        }
        ui.First.Value := SubStr(daily[1][1], 6)
        ui.Chart.Value := I18n.T("Usage.Chart", num(most))

        ; 次数多的在前 (相同时按 Usage.Features 的顺序), "呼出搜索窗口" 放最后
        rows := []
        for order, feature in Usage.Features
            rows.Push({Id: feature, Order: order, All: all.Has(feature) ? all[feature] : 0})
        Loop rows.Length - 1 {
            Loop rows.Length - A_Index {
                a := rows[A_Index], b := rows[A_Index + 1]
                if (b.All > a.All || (b.All = a.All && b.Order < a.Order))
                    rows[A_Index] := b, rows[A_Index + 1] := a
            }
        }
        rows.Push({Id: "Show", Order: 0, All: all.Has("Show") ? all["Show"] : 0})
        total := Usage.Total(all)
        list := ui.List
        list.Delete()
        for row in rows {
            counts := []
            for period in periods
                counts.Push(num(period.Has(row.Id) ? period[row.Id] : 0))
            share := (row.Id = "Show") ? "-" : (total ? Format("{:.1f}%", row.All * 100 / total) : "0%")
            list.Add("", I18n.T("Usage.F." row.Id), counts[1], counts[2], counts[3], counts[4], share)
        }
    }

    static _ClearUsage() {
        if (MsgBox(I18n.T("Usage.ClearConfirm"), App.Name, "YesNo Icon? Default2 Owner" PreferencesWindow.Gui.Hwnd) != "Yes")
            return
        Usage.Clear()
        PreferencesWindow._RefreshUsage()
    }

    static _ResetLearning() {
        Knowledge.Picks := Map(), Knowledge.QueryPicks := Map(), Knowledge.History := [], Knowledge.Recent := []
        Knowledge.Save()
        MsgBox(I18n.T("Prefs.ResetDone"), App.Name, 64)
    }

    ; 换一个数据文件夹 (例如 OneDrive 里的, 几台电脑共用): 把现在的设置和数据复制过去 (那里已经有 ALTRun 的
    ; 设置时可以直接用那里的), 位置记在默认位置的 ALTRun.json ("DataLocation"), 重新载入。原来的文件夹不动, 相当于留了一份备份
    static _ChangeDataFolder() {
        current := AppSettings.DataDir
        folder := DirSelect("*" current, 3, I18n.T("Prefs.DataFolderPrompt"))
        if (folder = "")
            return
        folder := RTrim(folder, "\")
        owner := " Owner" PreferencesWindow.Gui.Hwnd
        if (folder = current)
            return
        if (InStr(folder "\", current "\") = 1 || InStr(current "\", folder "\") = 1)
            return MsgBox(I18n.T("Prefs.DataFolderNested"), App.Name, "Icon!" owner)
        useExisting := false
        if FileExist(folder "\ALTRun.json") {
            answer := MsgBox(I18n.T("Prefs.DataFolderExisting", folder), App.Name, "YesNoCancel Icon?" owner)
            if (answer = "Cancel")
                return
            useExisting := (answer = "Yes")
        } else if (MsgBox(I18n.T("Prefs.DataFolderCopy", folder), App.Name, "OKCancel Icon?" owner) != "OK") {
            return
        }
        if PreferencesWindow.Cancel()                                       ; 有没保存的修改时先问要不要放弃
            return
        try {
            if !useExisting {
                AppSettings.Save(), Knowledge.Save(), Usage.Save(), ClipboardProvider.Save()
                DirCopy(current, folder, true)
            }
            AppSettings.SetDataLocation(folder)
        } catch as e {
            return MsgBox(I18n.T("Prefs.DataFolderFailed", folder, e.Message), App.Name, "Icon!")
        }
        App.Reload()
    }

    static _OpenFolder(folder) {
        DirCreate(folder)
        ActionCatalog.OpenFolder(folder)                                    ; 设置的文件管理器 (例如 Total Commander)
    }

    ;---------------------------------------------------------------------------
    ; Builders
    ;---------------------------------------------------------------------------
    ; labelW: 这一页左列标签的宽度 (按这一页最长的标签定, 内容不会整体偏右)
    static _BeginPage(nameKey, labelW := 190) {
        PreferencesWindow.Pages.Push({Key: nameKey, Name: I18n.T(nameKey), Wiki: PreferencesWindow.WikiPage(nameKey), Controls: []})
        PreferencesWindow._y := 14
        PreferencesWindow._pageKey := nameKey
        PreferencesWindow._pairX := 0                                       ; 这一页成对设置的右列位置, 见 _Pair
        PreferencesWindow.LabelW := Max(labelW, PreferencesWindow._labelWidths.Get(I18n.Lang " " nameKey, 0))
    }

    ; 左列的标签放不下 (会换行) 时记下这一页需要的宽度, 见 Show
    static _FitLabel(ctrl, text) {
        need := Ceil(Win.TextExtent(ctrl.Hwnd, text) * 96 / A_ScreenDPI) + 12 + 2
        if (need > PreferencesWindow.LabelW && PreferencesWindow.LabelW < PreferencesWindow.LabelMax) {
            key := PreferencesWindow._pageKey
            PreferencesWindow._wider[key] := Min(PreferencesWindow.LabelMax, Max(need, PreferencesWindow._wider.Get(key, 0)))
        }
    }

    static _Add(type, options, text := "") {
        pos := InStr(options, " x") || InStr(options, "x") = 1 ? "" : "x" PreferencesWindow.ContentX " "
        ctrl := PreferencesWindow.Gui.Add(type, pos "y" PreferencesWindow._y " " options " Hidden", text)
        PreferencesWindow.Pages[PreferencesWindow.Pages.Length].Controls.Push(ctrl)
        switch type, false {                                                ; 用户修改 -> "应用" 可用 (类型名不区分大小写)
            case "Edit", "DropDownList", "ComboBox": ctrl.OnEvent("Change", (*) => PreferencesWindow.MarkDirty())
            case "CheckBox", "Radio":                ctrl.OnEvent("Click", (*) => PreferencesWindow.MarkDirty())
        }
        return ctrl
    }

    static _Bind(path, readFn) {
        PreferencesWindow.Binds.Push({Path: path, Read: readFn})
    }

    ; 把 _y 移到这些控件最下面的位置之下 (控件的实际高度, 文字换行后也不会重叠)
    static _Below(gap, controls*) {
        bottom := PreferencesWindow._y
        for ctrl in controls {
            ctrl.GetPos(, &top, , &height)
            bottom := Max(bottom, top + height)
        }
        PreferencesWindow._y := bottom + gap
    }

    ; 分节标题: 粗体 + 下面一条细线; 不是页面第一行时和上面隔开一点
    static _Section(labelKey) {
        if (PreferencesWindow._y > 20)
            PreferencesWindow._y += 8
        PreferencesWindow.Gui.SetFont("bold")
        ctrl := PreferencesWindow._Add("Text", "w" PreferencesWindow.ContentW, I18n.T(labelKey))
        PreferencesWindow.Gui.SetFont("norm")
        PreferencesWindow._Below(3, ctrl)
        line := PreferencesWindow._Add("Text", "w" PreferencesWindow.ContentW " h1 0x10")   ; SS_ETCHEDHORZ
        PreferencesWindow._Below(PreferencesWindow._HasDesc(labelKey) ? 4 : 10, line)
        PreferencesWindow._Desc(labelKey, PreferencesWindow.ContentX)
    }

    static _HasDesc(labelKey) {
        return I18n.Strings.Has(labelKey ".Desc")
    }

    ; 设置下面的灰色小字说明 (I18n 里的 "<labelKey>.Desc"); 没有时什么也不加。x: 和上面的控件左边对齐
    static _Desc(labelKey, x, text := "") {
        if (text = "") {
            if !PreferencesWindow._HasDesc(labelKey)
                return
            text := I18n.T(labelKey ".Desc")
        }
        PreferencesWindow.Gui.SetFont("s8")
        ctrl := PreferencesWindow._Add("Text", "x" x " w" (PreferencesWindow.ContentX + PreferencesWindow.ContentW - x) " cGray", text)
        PreferencesWindow.Gui.SetFont("s9")
        PreferencesWindow._Below(10, ctrl)
    }

    ; 整行宽的灰色说明 (列表下面)
    static _Note(textKey) {
        PreferencesWindow._Desc("", PreferencesWindow.ContentX, I18n.T(textKey))
    }

    static _Gap() {
        PreferencesWindow._y += 10
    }

    ; 标签和输入框放在同一行: 标签往下挪 3 像素和输入框的文字对齐
    static _Label(labelKey) {
        PreferencesWindow._y += 3
        ctrl := PreferencesWindow._Add("Text", "w" (PreferencesWindow.LabelW - 12) " Right", I18n.T(labelKey))
        PreferencesWindow._y -= 3
        PreferencesWindow._FitLabel(ctrl, I18n.T(labelKey))
        return ctrl
    }

    static _InputX() {
        return PreferencesWindow.ContentX + PreferencesWindow.LabelW
    }

    ; 控件列的宽度 (长输入框 "L" 的宽度)
    static _InputW() {
        return PreferencesWindow.ContentW - PreferencesWindow.LabelW
    }

    ; 输入框宽度: "S" 数字 / "M" 热键、关键字 / "L" 路径、过滤条件 (整个控件列; 带浏览按钮时让出按钮的位置)
    static _Width(size, withBrowse := false) {
        switch size {
            case "S": return PreferencesWindow.WidthS
            case "M": return PreferencesWindow.WidthM
            case "K": return PreferencesWindow.WidthK
            case "B": return PreferencesWindow.ButtonW                      ; 和按钮一样宽 (外观页: 主题下拉框和下面的按钮对齐)
        }
        return PreferencesWindow._InputW() - (withBrowse ? 34 : 0)
    }

    ; 复选框放在控件列; group: 左列的分组标签 (一组里的第一个复选框才写)。
    ; desc: 代替 "<labelKey>.Desc" 显示的说明 (例如随状态变化的文字); width: 默认到页面右边
    static _Check(path, labelKey, x := 0, desc := "", group := "", width := 0) {
        x := x ? x : PreferencesWindow._InputX()
        if (group != "")
            PreferencesWindow._GroupLabel(group)
        ctrl := PreferencesWindow._Add("Checkbox", "x" x " w" (width ? width : PreferencesWindow.ContentX + PreferencesWindow.ContentW - x), I18n.T(labelKey))
        ctrl.Value := PreferencesWindow.GetPath(PreferencesWindow.Working, path) ? 1 : 0
        PreferencesWindow._Bind(path, () => ctrl.Value)
        hasDesc := (desc != "" || PreferencesWindow._HasDesc(labelKey))
        PreferencesWindow._Below(hasDesc ? 0 : 6, ctrl)
        PreferencesWindow._Desc(labelKey, x + 18, desc)                     ; 和复选框的文字对齐
        return ctrl
    }

    ; 左列的分组标签, 和右边第一个控件同一行
    static _GroupLabel(labelKey, dy := 1) {
        PreferencesWindow._y += dy
        ctrl := PreferencesWindow._Add("Text", "x" PreferencesWindow.ContentX " w" (PreferencesWindow.LabelW - 12) " Right", I18n.T(labelKey))
        PreferencesWindow._y -= dy
        PreferencesWindow._FitLabel(ctrl, I18n.T(labelKey))
        return ctrl
    }

    ; 左列标签 + 右边一段文字 (或链接), 例如 "版本: 2026.09.26"
    static _Info(labelKey, text, options := "", type := "Text") {
        label := PreferencesWindow._GroupLabel(labelKey)
        ctrl := PreferencesWindow._Add(type, "x" PreferencesWindow._InputX() " w" PreferencesWindow._InputW() options, text)
        PreferencesWindow._Below(6, label, ctrl)
        return ctrl
    }

    ; 标签后面紧跟输入框 (分节排列的页面用, 不用左列): "搜索历史条数 [30]"
    static _InlineField(path, labelKey, size, kind := "text") {
        x := PreferencesWindow._InputX()
        PreferencesWindow._y += 3
        label := PreferencesWindow._Add("Text", "x" x, I18n.T(labelKey))
        PreferencesWindow._y -= 3
        label.GetPos(, , &labelW)
        value := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        ctrl := PreferencesWindow._Add("Edit", "x" (x + labelW + 8) " w" PreferencesWindow._Width(size) " r1 -Multi" (kind = "number" ? " Number" : ""), value)
        if (kind = "number")
            PreferencesWindow._Bind(path, () => IsInteger(ctrl.Value) ? Integer(ctrl.Value) : 0)
        else
            PreferencesWindow._Bind(path, () => ctrl.Value)
        PreferencesWindow._Below(PreferencesWindow._HasDesc(labelKey) ? 3 : 8, label, ctrl)
        PreferencesWindow._Desc(labelKey, x)
    }

    ; kind: text / number / file / folder; size: "S" / "M" / "L" (见 _Width)
    static _Field(path, labelKey, size, kind := "text") {
        label := PreferencesWindow._Label(labelKey)
        if (kind = "hotkey")
            return PreferencesWindow._BelowInput(labelKey, label, PreferencesWindow._HotkeyBox(PreferencesWindow._InputX(), path))
        value := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        browseKind := (kind = "file" || kind = "folder")
        ctrl := PreferencesWindow._Add("Edit", "x" PreferencesWindow._InputX() " w" PreferencesWindow._Width(size, browseKind) " r1 -Multi" (kind = "number" ? " Number" : ""), value)
        if browseKind {
            browse := PreferencesWindow._Add("Button", "x+4 yp-1 w30", I18n.T("Prefs.Browse"))
            browseFn := ItemEditor._Browser(ctrl, kind, PreferencesWindow.Gui)
            browse.OnEvent("Click", (p*) => (browseFn(p*), PreferencesWindow.MarkDirty()))
        }
        if (kind = "number")
            PreferencesWindow._Bind(path, () => IsInteger(ctrl.Value) ? Integer(ctrl.Value) : 0)
        else
            PreferencesWindow._Bind(path, () => ctrl.Value)
        PreferencesWindow._BelowInput(labelKey, label, ctrl)
    }

    ; 标签 + 输入框一行之后: 有说明时说明紧贴在输入框下面
    static _BelowInput(labelKey, controls*) {
        hasDesc := PreferencesWindow._HasDesc(labelKey)
        PreferencesWindow._Below(hasDesc ? 3 : 8, controls*)
        PreferencesWindow._Desc(labelKey, PreferencesWindow._InputX())
    }

    static _Choice(path, labelKey, values, labels, size := "M") {
        label := PreferencesWindow._Label(labelKey)
        ctrl := PreferencesWindow._Add("DropDownList", "x" PreferencesWindow._InputX() " w" PreferencesWindow._Width(size), labels)
        current := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        ctrl.Value := 1
        for index, value in values
            if (value = current)
                ctrl.Value := index
        PreferencesWindow._Bind(path, () => values[ctrl.Value])
        PreferencesWindow._BelowInput(labelKey, label, ctrl)
        return ctrl
    }

    static _Lines(path, labelKey, rows) {
        label := PreferencesWindow._Label(labelKey)
        value := PreferencesWindow.JoinLines(PreferencesWindow.GetPath(PreferencesWindow.Working, path))
        ctrl := PreferencesWindow._Add("Edit", "x" PreferencesWindow._InputX() " w" PreferencesWindow._InputW() " r" rows " +Multi -Wrap +HScroll +VScroll", value)
        PreferencesWindow._Bind(path, () => PreferencesWindow.SplitLines(ctrl.Value))
        PreferencesWindow._Below(3, label, ctrl)
        PreferencesWindow._Desc(labelKey, PreferencesWindow._InputX())
    }

    ; 窗口列表 (设置里是逗号分隔的 "ahk_class ..., ahk_exe ...") 每行显示一个。
    ; path 为空时不绑定, value 直接给数组; 返回输入框
    static _WinList(path, labelKey, rows, value := "") {
        label := PreferencesWindow._Label(labelKey)
        if (path != "")
            value := PreferencesWindow.SplitWindows(PreferencesWindow.GetPath(PreferencesWindow.Working, path))
        ctrl := PreferencesWindow._Add("Edit", "x" PreferencesWindow._InputX() " w" PreferencesWindow._InputW() " r" rows " +Multi -Wrap +HScroll +VScroll", PreferencesWindow.JoinLines(value))
        if (path != "")
            PreferencesWindow._Bind(path, () => PreferencesWindow.JoinWindows(PreferencesWindow.SplitLines(ctrl.Value)))
        PreferencesWindow._Below(3, label, ctrl)
        PreferencesWindow._Desc(labelKey, PreferencesWindow._InputX())
        return ctrl
    }

    static _Csv(path, labelKey, size) {
        label := PreferencesWindow._Label(labelKey)
        value := PreferencesWindow.JoinCsv(PreferencesWindow.GetPath(PreferencesWindow.Working, path))
        ctrl := PreferencesWindow._Add("Edit", "x" PreferencesWindow._InputX() " w" PreferencesWindow._Width(size) " r1 -Multi", value)
        PreferencesWindow._Bind(path, () => PreferencesWindow.SplitCsv(ctrl.Value))
        PreferencesWindow._BelowInput(labelKey, label, ctrl)
    }

    ; 按钮放在控件列, 统一尺寸; group: 左列的分组标签
    static _Button(labelKey, fn, group := "") {
        if (group != "")
            PreferencesWindow._GroupLabel(group, 5)                         ; 和按钮上的文字对齐
        ctrl := PreferencesWindow._Add("Button", "x" PreferencesWindow._InputX() " w" PreferencesWindow.ButtonW " h" PreferencesWindow.ButtonH, I18n.T(labelKey))
        ctrl.OnEvent("Click", fn)
        PreferencesWindow._Below(PreferencesWindow._HasDesc(labelKey) ? 2 : 6, ctrl)
        PreferencesWindow._Desc(labelKey, PreferencesWindow._InputX())
        return ctrl
    }

    ; 一行并排几个按钮: specs = [labelKey, fn], ...; 说明用第一个按钮的 "<labelKey>.Desc"
    static _Buttons(group, specs*) {
        if (group != "")
            PreferencesWindow._GroupLabel(group, 5)
        x := PreferencesWindow._InputX(), controls := []
        buttonW := Min(PreferencesWindow.ButtonW, (PreferencesWindow._InputW() - 10 * (specs.Length - 1)) // specs.Length)   ; 左列加宽后放不下时按钮窄一点
        for spec in specs {
            ctrl := PreferencesWindow._Add("Button", "x" x " w" buttonW " h" PreferencesWindow.ButtonH, I18n.T(spec[1]))
            ctrl.OnEvent("Click", spec[2])
            controls.Push(ctrl)
            x += buttonW + 10
        }
        PreferencesWindow._Below(PreferencesWindow._HasDesc(specs[1][1]) ? 2 : 6, controls*)
        PreferencesWindow._Desc(specs[1][1], PreferencesWindow._InputX())
    }

    ; 同一行两个设置: 左边的标签在左列, 右边的标签紧跟在左边的输入框后面。
    ; left / right = [path, labelKey, size, kind := "text", values, labels] (kind: text / number / csv / choice)
    ; 说明用左边的 "<labelKey>.Desc"
    static _Pair(left, right) {
        label := PreferencesWindow._Label(left[2])
        first := PreferencesWindow._Input(PreferencesWindow._InputX(), left*)
        first.GetPos(&x, , &w)
        sameAsFirstRow := PreferencesWindow._pairX && left[3] = PreferencesWindow._pairSizes[1] && right[3] = PreferencesWindow._pairSizes[2]
        if sameAsFirstRow {                                                 ; 和这一页第一行一样的输入框: 用同样的宽度, 右边对齐到同一列
            w := PreferencesWindow._pairFirstW
            first.Move(, , w)
        }
        PreferencesWindow._y += 3
        label2 := PreferencesWindow._Add("Text", "x" (x + w + 18), I18n.T(right[2]))
        PreferencesWindow._y -= 3
        label2.GetPos(&x2, , &w2)
        ; 左列加宽后控件列变窄, 放不下时把左边的输入框缩短 (最短 "S"); 两边是同一种输入框时一起缩短, 保持一样宽
        secondW := sameAsFirstRow ? PreferencesWindow._pairSecondW : 0
        overflow := x2 + w2 + 8 + (secondW ? secondW : PreferencesWindow._Width(right[3])) - (PreferencesWindow.ContentX + PreferencesWindow.ContentW)
        if (!sameAsFirstRow && overflow > 0 && left[3] = right[3] && w > PreferencesWindow.WidthS) {
            w := Max(PreferencesWindow.WidthS, (PreferencesWindow.ContentX + PreferencesWindow.ContentW - x - 18 - w2 - 8) // 2)
            first.Move(, , w)
            label2.Move(x + w + 18)
            label2.GetPos(&x2)
            secondW := w, overflow := x2 + w2 + 8 + w - (PreferencesWindow.ContentX + PreferencesWindow.ContentW)
        }
        if (overflow > 0 && w > PreferencesWindow.WidthS) {
            shrink := Min(overflow, w - PreferencesWindow.WidthS)
            w -= shrink, overflow -= shrink
            first.Move(, , w)
            label2.Move(x + w + 18)
            label2.GetPos(&x2)
        }
        if (overflow > 0) {                                                 ; 还是放不下: 右边这一项换到下一行
            first.GetPos(, &firstY, , &firstH)
            PreferencesWindow._y := firstY + firstH + 6
            label2.Move(PreferencesWindow._InputX(), PreferencesWindow._y + 3)
            label2.GetPos(&x2)
        }
        secondX := x2 + w2 + 8
        ; 同一页的几行成对的设置, 左边输入框一样宽时: 右边的输入框对齐到第一行右边输入框的位置 (标签靠右贴着输入框)
        if (overflow <= 0) {
            target := PreferencesWindow._pairX
            if !target
                PreferencesWindow._pairX := secondX, PreferencesWindow._pairFirstW := w, PreferencesWindow._pairSizes := [left[3], right[3]]
                , PreferencesWindow._pairSecondW := secondW ? secondW : PreferencesWindow._Width(right[3])
            else if (w = PreferencesWindow._pairFirstW && target >= secondX && target + (secondW ? secondW : PreferencesWindow._Width(right[3])) <= PreferencesWindow.ContentX + PreferencesWindow.ContentW) {
                label2.Move(target - 8 - w2)
                secondX := target
            }
        }
        second := PreferencesWindow._Input(secondX, right*)
        if secondW
            second.Move(, , secondW)
        PreferencesWindow._BelowInput(left[2], label, first, label2, second)
    }

    ; 单独一个输入控件 (不带标签), 并绑定 path; kind 见 _Pair
    static _Input(x, path, labelKey, size, kind := "text", values := "", labels := "") {
        current := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        if (kind = "hotkey")
            return PreferencesWindow._HotkeyBox(x, path)
        if (kind = "choice") {
            ctrl := PreferencesWindow._Add("DropDownList", "x" x " w" PreferencesWindow._Width(size), labels)
            ctrl.Value := 1
            for index, value in values
                if (value = current)
                    ctrl.Value := index
            PreferencesWindow._Bind(path, () => values[ctrl.Value])
            return ctrl
        }
        text := (kind = "csv") ? PreferencesWindow.JoinCsv(current) : current
        ctrl := PreferencesWindow._Add("Edit", "x" x " w" PreferencesWindow._Width(size) " r1 -Multi" (kind = "number" ? " Number" : ""), text)
        switch kind {
            case "number": PreferencesWindow._Bind(path, () => IsInteger(ctrl.Value) ? Integer(ctrl.Value) : 0)
            case "csv":    PreferencesWindow._Bind(path, () => PreferencesWindow.SplitCsv(ctrl.Value))
            default:       PreferencesWindow._Bind(path, () => ctrl.Value)
        }
        return ctrl
    }

    ; 录制热键的框: 不经过 _Add (程序改框里的文字不算修改), 录到新热键时才标记 "有修改"
    static _HotkeyBox(x, path) {
        ctrl := HotkeyBox.Add(PreferencesWindow.Gui, "x" x " y" PreferencesWindow._y " w" PreferencesWindow.WidthK " Hidden"
            , PreferencesWindow.GetPath(PreferencesWindow.Working, path), false, () => PreferencesWindow.MarkDirty())
        PreferencesWindow.Pages[PreferencesWindow.Pages.Length].Controls.Push(ctrl)
        PreferencesWindow._Bind(path, () => HotkeyBox.Value(ctrl))
        return ctrl
    }

    ; 和上一行的短输入框 ("S") 同一行的按钮 (例如 "子文件夹深度 [4]  [重建文件索引]")
    static _SideButton(rowY, labelKey, fn) {
        x := PreferencesWindow._InputX() + PreferencesWindow.WidthS + 12
        ctrl := PreferencesWindow._Add("Button", "x" x " y" (rowY - 2) " w" PreferencesWindow.ButtonW " h" PreferencesWindow.ButtonH, I18n.T(labelKey))
        ctrl.OnEvent("Click", fn)
        return ctrl
    }

    ; 列表页: ListView 显示 path 指向的数组, 添加 / 编辑 / 删除都用 ItemEditor
    ; check: 可选, (item, roots) => "OK" / "Missing" / "Unavailable" / "Skipped"。
    ; 提供时多一个 "检查路径" 按钮, 结果显示在 "Status" 列 (列表里的项目本身没有这个键)
    ; 按钮右边的筛选框: 只显示含有所有输入的词的项目 (在各列的文字里找, 不分大小写)
    static _List(path, height, columns, fields, newItem, check := "") {
        items := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        headers := []
        for column in columns
            headers.Push(I18n.T(column[1]))
        listView := PreferencesWindow._Add("ListView", "w" PreferencesWindow.ContentW " h" (height - 40) " -Multi NoSort Grid", headers)
        for index, column in columns
            listView.ModifyCol(index, column[3])
        pageName := PreferencesWindow.Pages[PreferencesWindow.Pages.Length].Name
        statuses := Map()                                                   ; ObjPtr(item) -> 检查结果, 检查过才有
        shown := []                                                         ; 每一行对应 items 里的第几项 (筛选时只显示一部分)
        filterBox := ""

        Refresh(selectRow := 0) {
            listView.Opt("-Redraw")
            listView.Delete()
            shown.Length := 0
            filter := IsObject(filterBox) ? filterBox.Value : ""
            for index, item in items {
                cells := []
                for column in columns
                    cells.Push((column[2] = "Status") ? StatusText(item) : PreferencesWindow._Cell(item, column[2]))
                if !PreferencesWindow.FilterMatch(filter, cells)
                    continue
                listView.Add("", cells*)
                shown.Push(index)
            }
            if (selectRow && shown.Length)
                listView.Modify(Min(selectRow, shown.Length), "Select Focus Vis")
            listView.Opt("+Redraw")
        }
        RowOf(index) {                                                      ; items 里的第 index 项在第几行, 没有显示时 0
            for row, shownIndex in shown
                if (shownIndex = index)
                    return row
            return 0
        }
        StatusText(item) {
            key := ObjPtr(item)
            return statuses.Has(key) ? I18n.T("Prefs.Status." statuses[key]) : ""
        }
        Recheck(item) {                                                     ; 检查过之后, 新增 / 修改的项目也马上检查
            if statuses.Count
                statuses[ObjPtr(item)] := check(item, Map())
        }
        AddItem(*) => Insert(newItem())
        Insert(item) {                                                      ; 打开编辑框 (item 是预先填好的内容), 确定后加进列表
            edited := ItemEditor.Edit(PreferencesWindow.Gui, pageName, fields, item)
            if IsObject(edited)
                Append(edited)
        }
        Append(item) {
            items.Push(item)
            PreferencesWindow.MarkDirty(true)
            Recheck(item)
            if IsObject(filterBox)
                filterBox.Value := ""                                       ; 清掉筛选, 新加的一项一定看得到
            Refresh(items.Length)
        }
        EditExisting(item) {                                                ; 选中列表里的这一项并打开编辑框
            for index, existing in items {
                if (existing = item) {
                    if IsObject(filterBox)
                        filterBox.Value := ""
                    Refresh(index)
                    return EditItem()
                }
            }
        }
        EditItem(*) {
            row := listView.GetNext()
            if !row
                return
            index := shown[row]
            edited := ItemEditor.Edit(PreferencesWindow.Gui, pageName, fields, ItemEditor.WithDefaults(items[index], newItem()))
            if IsObject(edited) {
                if statuses.Has(ObjPtr(items[index]))
                    statuses.Delete(ObjPtr(items[index]))
                items[index] := edited
                PreferencesWindow.MarkDirty()
                Recheck(edited)
                Refresh(row)
            }
        }
        CheckItems(button, *) {
            button.Enabled := false
            statuses.Clear()
            roots := Map()
            counts := Map("OK", 0, "Missing", 0, "Unavailable", 0, "Skipped", 0)
            firstProblem := 0
            for index, item in items {
                result := check(item, roots)
                statuses[ObjPtr(item)] := result
                counts[result] += 1
                if (!firstProblem && (result = "Missing" || result = "Unavailable"))
                    firstProblem := index
            }
            Refresh(RowOf(firstProblem))
            button.Enabled := true
            if !firstProblem
                message := I18n.T("Prefs.CheckAllOk", counts["OK"])
            else
                message := I18n.T("Prefs.CheckResult", counts["Missing"], counts["Unavailable"])
            if counts["Skipped"]
                message .= "`n`n" I18n.T("Prefs.CheckSkipped", counts["Skipped"])
            MsgBox(message, pageName, (firstProblem ? "Icon!" : "Iconi") " Owner" PreferencesWindow.Gui.Hwnd)
        }
        DeleteItem(*) {
            row := listView.GetNext()
            if !row
                return
            if (MsgBox(I18n.T("Prefs.ConfirmDelete", listView.GetText(row, 1)), pageName, "YesNo Icon? Owner" PreferencesWindow.Gui.Hwnd) != "Yes")
                return
            items.RemoveAt(shown[row])
            PreferencesWindow.MarkDirty()
            Refresh(row)
        }

        listView.OnEvent("DoubleClick", EditItem)
        Refresh()
        PreferencesWindow._lists[path] := {Items: items, Insert: Insert, Append: Append, Edit: EditExisting}   ; 外面加进来的项目 (AddCommands)
        PreferencesWindow._y += height - 34
        PreferencesWindow._Add("Button", "w80", I18n.T("Prefs.Add")).OnEvent("Click", AddItem)
        PreferencesWindow._Add("Button", "x+8 yp w80", I18n.T("Prefs.Edit")).OnEvent("Click", EditItem)
        PreferencesWindow._Add("Button", "x+8 yp w80", I18n.T("Prefs.Delete")).OnEvent("Click", DeleteItem)
        if IsObject(check)
            PreferencesWindow._Add("Button", "x+24 yp w100", I18n.T("Prefs.CheckTargets")).OnEvent("Click", CheckItems)
        listView.GetPos(&listX, , &listW)
        filterBox := PreferencesWindow.Gui.Add("Edit", "x" (listX + listW - 160) " yp+1 w160 r1 -Multi Hidden")   ; 不用 _Add: 筛选不算修改设置
        PreferencesWindow.Pages[PreferencesWindow.Pages.Length].Controls.Push(filterBox)
        Win.SetCueBanner(filterBox.Hwnd, I18n.T("Prefs.Filter"))
        filterBox.OnEvent("Change", (*) => Refresh())
        PreferencesWindow._y += 38
    }

    ; 筛选框: filter 里的每个词 (空格分开) 都出现在某一列里才显示, 不分大小写; 空白时全部显示
    static FilterMatch(filter, cells) {
        filter := Trim(filter)
        if (filter = "")
            return true
        text := ""
        for cell in cells
            text .= cell "`n"
        for word in StrSplit(filter, " ")
            if (word != "" && !InStr(text, word))
                return false
        return true
    }

    static _Cell(item, key) {
        if !(item is Map) || !item.Has(key)
            return ""
        if (key = "Target" && item.Has("Arguments") && Trim(item["Arguments"]) != "")   ; 自定义命令: 目标后面接着显示参数
            return RegExReplace(item["Target"] " " item["Arguments"], "\s+", " ")
        if (key = "Key")                                                    ; 自定义热键: 显示 Ctrl+Alt+P、鼠标中键
            return HotkeyBox.Label(item[key])
        if (key = "Type") {                                                 ; 自定义命令的类型显示翻译后的名称
            label := I18n.T("Prefs.TypeShort." item[key])
            if (label != "Prefs.TypeShort." item[key])
                return label
        }
        return RegExReplace(item[key] "", "\s+", " ")
    }

    ;---------------------------------------------------------------------------
    ; Data helpers (纯函数, 有单元测试)
    ;---------------------------------------------------------------------------
    ; 全局热键 (不限定窗口的) -> [{Key, Label}...]: 呼出热键、第二个呼出热键、剪贴板历史、没有填窗口的自定义热键。
    ; 对话框跳转、一键加日期只在特定的窗口里生效 (两个 Ctrl+D 是正常的), 不算在内
    static _GlobalHotkeys(data) {
        list := []
        add(key, label) => (Trim(key) != "") ? list.Push({Key: key, Label: label}) : ""
        general := data["General"]
        add(general["Hotkey"], I18n.T("Prefs.Hotkey"))
        add(general["SecondaryHotkey"], I18n.T("Prefs.SecondaryHotkey"))
        add(general.Has("SelectionHotkey") ? general["SelectionHotkey"] : "", I18n.T("Prefs.SelectionHotkey"))
        clipboard := PreferencesWindow.GetPath(data, "Features.Clipboard")
        if (clipboard is Map && clipboard.Has("Enabled") && clipboard["Enabled"] && clipboard.Has("Hotkey"))
            add(clipboard["Hotkey"], I18n.T("Prefs.Page.Clipboard"))
        if (data.Has("Hotkeys") && data["Hotkeys"] is Array)
            for entry in data["Hotkeys"]
                if (entry is Map && entry.Has("Key") && (!entry.Has("WinTitle") || Trim(entry["WinTitle"]) = ""))
                    add(entry["Key"], I18n.T("Prefs.Page.Hotkeys") " (" (entry.Has("Action") ? entry["Action"] : "") ")")
        return list
    }

    ; 第一个重复的热键 -> {Key, First, Second} (两处的名称); 没有重复时 ""。~ * $ 不影响是不是同一个键
    static DuplicateHotkey(list) {
        seen := Map()
        for entry in list {
            key := StrLower(RegExReplace(Trim(entry.Key), "^[~*$]+"))
            if RegExMatch(key, "^([\^!+#]+)(.+)$", &m)                     ; 修饰符的顺序不同也是同一个键 (!^c = ^!c)
                key := StrReplace(Sort(RegExReplace(m[1], "(.)", "$1`n")), "`n") m[2]
            if seen.Has(key)
                return {Key: entry.Key, First: seen[key], Second: entry.Label}
            seen[key] := entry.Label
        }
        return ""
    }

    static DeepCopy(value) {
        return JSON.Parse(JSON.Stringify(value))
    }

    static GetPath(data, path) {
        node := data
        for key in StrSplit(path, ".") {
            if !(node is Map) || !node.Has(key)
                return ""
            node := node[key]
        }
        return node
    }

    static SetPath(data, path, value) {
        keys := StrSplit(path, ".")
        node := data
        Loop keys.Length - 1 {
            key := keys[A_Index]
            if !(node.Has(key) && node[key] is Map)
                node[key] := Map()
            node := node[key]
        }
        node[keys[keys.Length]] := value
    }

    static SplitLines(text) {
        items := []
        for line in StrSplit(text, "`n", "`r")
            if (Trim(line) != "")
                items.Push(Trim(line))
        return items
    }

    static JoinLines(items) {
        return (items is Array) ? TextTools.Join(items) : items
    }

    static SplitCsv(text) {
        items := []
        for part in StrSplit(text, ",")
            if (Trim(part) != "")
                items.Push(Trim(part))
        return items
    }

    ; 窗口条件列表 "ahk_class A, ahk_exe B" <-> 数组 (去掉空白和空项)
    static SplitWindows(text) {
        return PreferencesWindow.SplitCsv(text)
    }

    ; 几个数组合成 "ahk_class A, ahk_exe B", 去掉重复
    static JoinWindows(lists*) {
        all := [], seen := Map()
        seen.CaseSense := false
        for list in lists
            for window in list
                if (Trim(window) != "" && !seen.Has(Trim(window)))
                    seen[Trim(window)] := true, all.Push(Trim(window))
        return PreferencesWindow.JoinCsv(all)
    }

    static JoinCsv(items) {
        if !(items is Array)
            return items
        out := ""
        for index, item in items
            out .= (index = 1 ? "" : ", ") item
        return out
    }
}
