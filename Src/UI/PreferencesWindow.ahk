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
;   _Field("General.Hotkey", "Prefs.Hotkey", "M")               文本框, 宽度 "S" / "M" / "L"
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
;===============================================================================

class PreferencesWindow {
    static Gui := "", Working := "", Pages := [], Binds := [], PageList := ""
    static _page := 0, _y := 0
    static _dirty := false, _ready := false, ApplyButton := ""
    ; 两列表单 (和 Alfred / Listary 一样): 左列是右对齐的标签, 所有控件从同一条竖线 (_InputX) 开始;
    ; 输入框只用三种宽度, 复选框按组在左列加标签, 按钮统一尺寸, 灰色说明和控件左对齐
    static ContentX := 190, ContentW := 560, LabelW := 190
    static WidthS := 70, WidthM := 220, ButtonW := 200, ButtonH := 26       ; WidthL = 整个控件列 (_InputW)
    static ButtonY := 579                                                   ; 底部按钮的位置; 页面内容要在它上面 (y < ButtonY - 10)
    static _positionReset := false

    static Show(pageIndex := 1, x := "", y := "") {
        if IsObject(PreferencesWindow.Gui) {
            WinActivate("ahk_id " PreferencesWindow.Gui.Hwnd)
            return
        }
        SearchWindow.Hide()
        PreferencesWindow.Working := PreferencesWindow.DeepCopy(AppSettings.Data)
        PreferencesWindow.Pages := [], PreferencesWindow.Binds := []
        PreferencesWindow._dirty := false, PreferencesWindow._ready := false, PreferencesWindow._positionReset := false

        g := Gui("-MinimizeBox", I18n.T("Prefs.Title"))
        g.SetFont("s9", ThemeManager.FontName())
        g.OnEvent("Close", (*) => PreferencesWindow.Cancel())                ; 返回 true = 不关闭 (选择了继续编辑)
        g.OnEvent("Escape", (*) => PreferencesWindow.Cancel())
        PreferencesWindow.Gui := g

        PreferencesWindow._BuildPages()

        names := []
        for page in PreferencesWindow.Pages
            names.Push(page.Name)
        pageList := g.AddListBox("x12 y12 w160 h" (PreferencesWindow.ButtonY - 24) " AltSubmit", names)
        pageList.OnEvent("Change", (ctrl, *) => PreferencesWindow.SelectPage(ctrl.Value))
        SendMessage(0x1A0, 0, Round(26 * A_ScreenDPI / 96), pageList.Hwnd)  ; LB_SETITEMHEIGHT: 更宽松的侧边栏
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

        pageIndex := Max(1, Min(pageIndex, PreferencesWindow.Pages.Length))
        pageList.Value := pageIndex
        PreferencesWindow.SelectPage(pageIndex)
        g.Show((IsInteger(x) && IsInteger(y) ? "x" x " y" y " " : "") "w765 h" (PreferencesWindow.ButtonY + 41))
        ; 打开窗口时程序自己填的值不算修改, 等控件的通知都处理完再开始记录
        SetTimer(() => (PreferencesWindow._ready := true), -300)
    }

    ; 有控件被用户修改过: "应用" 变为可用, 取消时要确认
    static MarkDirty() {
        if (!PreferencesWindow._ready || PreferencesWindow._dirty)
            return
        PreferencesWindow._dirty := true
        try PreferencesWindow.ApplyButton.Enabled := true
    }

    static SelectPage(pageIndex) {
        PreferencesWindow._page := pageIndex
        for index, page in PreferencesWindow.Pages
            for ctrl in page.Controls
                ctrl.Visible := (index = pageIndex)
        ; 复选框和说明文字是透明背景: 不先擦掉背景就重画, 文字会叠两遍变成粗体
        DllCall("RedrawWindow", "Ptr", PreferencesWindow.Gui.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x185)   ; RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW
    }

    static Close() {
        if IsObject(PreferencesWindow.Gui)
            PreferencesWindow.Gui.Destroy()
        PreferencesWindow.Gui := ""
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
        ActionCatalog.OpenUrl(HelpProvider.WikiUrl page.Wiki)
    }

    ; 每一页对应的 Wiki 页面 (帮助按钮 / F1)
    static WikiPage(nameKey) {
        static pages := Map("Prefs.Page.Window", "Usage", "Prefs.Page.Appearance", "Themes", "Prefs.Page.Features", "Usage", "Prefs.Page.FileSearch", "File-Search"
                          , "Prefs.Page.Commands", "Commands-and-Snippets", "Prefs.Page.Snippets", "Commands-and-Snippets"
                          , "Prefs.Page.Clipboard", "Commands-and-Snippets", "Prefs.Page.WebSearch", "Usage"
                          , "Prefs.Page.Hotkeys", "Extensions", "Prefs.Page.Extensions", "Extensions", "Prefs.Page.Usage", "Usage")
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
        PreferencesWindow._BuildCommands()
        PreferencesWindow._BuildSnippets()
        PreferencesWindow._BuildClipboard()
        PreferencesWindow._BuildWebSearch()
        PreferencesWindow._BuildHotkeys()
        PreferencesWindow._BuildExtensions()
        PreferencesWindow._BuildUsage()
        PreferencesWindow._BuildAdvanced()
    }

    static _BuildGeneral() {
        PreferencesWindow._BeginPage("Prefs.Page.General")
        PreferencesWindow._Field("General.Hotkey", "Prefs.Hotkey", "M")
        PreferencesWindow._Field("General.SecondaryHotkey", "Prefs.SecondaryHotkey", "M")
        PreferencesWindow._Choice("General.Language", "Prefs.Language", ["auto", "en", "zh"], [I18n.T("Prefs.Language.auto"), "English", "中文"])
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
        PreferencesWindow._BeginPage("Prefs.Page.Window")
        for row in [["HideOnDeactivate", "Prefs.HideOnDeactivate", "Prefs.Group.Window"], ["KeepLastQuery", "Prefs.KeepLastQuery", ""]
                   , ["SwitchToEnglishInput", "Prefs.EnglishInput", "Prefs.Group.Typing"], ["SpaceToRun", "Prefs.SpaceToRun", ""]
                   , ["ShowTips", "Prefs.ShowTips", "Prefs.Group.Tips"]]
            PreferencesWindow._Check("General." row[1], row[2], , , row[3])
        PreferencesWindow._Gap()
        PreferencesWindow._Field("General.HistorySize", "Prefs.HistorySize", "S", "number")
    }

    static _BuildAppearance() {
        PreferencesWindow._BeginPage("Prefs.Page.Appearance")
        themes := ThemeManager.Names(), labels := []
        for themeName in themes
            labels.Push(PreferencesWindow._ThemeLabel(themeName))
        themeChoice := PreferencesWindow._Choice("Appearance.Theme", "Prefs.Theme", themes, labels)
        PreferencesWindow._Field("Appearance.Width", "Prefs.Width", "S", "number")
        PreferencesWindow._Field("Appearance.VisibleRows", "Prefs.VisibleRows", "S", "number")
        PreferencesWindow._Button("Prefs.CopyTheme", (*) => PreferencesWindow._CopyTheme(themeChoice, themes), "Prefs.Group.CustomThemes")
        PreferencesWindow._Button("Prefs.OpenThemes", (*) => PreferencesWindow._OpenFolder(ThemeManager.UserDir))
        PreferencesWindow._Section("Prefs.WindowPosition")
        PreferencesWindow._Choice("Appearance.ShowOn", "Prefs.ShowOn", ["Mouse", "Primary", "Active"]
            , [I18n.T("Prefs.ShowOn.Mouse"), I18n.T("Prefs.ShowOn.Primary"), I18n.T("Prefs.ShowOn.Active")])
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

    ; 内置主题显示翻译后的名称, 用户主题显示文件名
    static _ThemeLabel(themeName) {
        label := I18n.T("Theme." themeName)
        return (label = "Theme." themeName) ? themeName : label
    }

    ; 把选中的主题 (展开成完整的键) 复制到 Themes\<新名称>.json, 用记事本打开, 并在列表里选中它
    static _CopyTheme(themeChoice, themes) {
        source := themes[themeChoice.Value]
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
            DirCreate(ThemeManager.UserDir)
            FileOpen(themeFile, "w", "UTF-8").Write(JSON.Stringify(theme))
        } catch as e {
            return MsgBox(e.Message, App.Name, 48)
        }
        index := 0
        for existing, value in themes
            if (value = themeName)
                index := existing
        if !index {
            themes.Push(themeName)
            themeChoice.Add([themeName])
            index := themes.Length
        }
        themeChoice.Choose(index)
        Run('notepad.exe "' themeFile '"')
    }

    static _BuildFeatures() {
        PreferencesWindow._BeginPage("Prefs.Page.Features")
        PreferencesWindow._Section("Prefs.EnabledFeatures")
        features := ["Applications", "CustomCommands", "Snippets", "Clipboard", "Calculator", "WebSearch", "FileSearch", "Terminal", "System"]
        startY := PreferencesWindow._y, columnW := PreferencesWindow._InputW() // 2      ; 两列: 英文名称较长, 三列会换行
        for index, feature in features {
            column := Mod(index - 1, 2), row := (index - 1) // 2
            PreferencesWindow._y := startY + row * 24
            PreferencesWindow._Check("Features." feature ".Enabled", "Prefs.Feature." feature
                , PreferencesWindow._InputX() + column * columnW, , (index = 1) ? "Prefs.Group.SearchFeatures" : "", columnW)
        }
        PreferencesWindow._y := startY + Ceil(features.Length / 2) * 24 + 12
        PreferencesWindow._Check("Features.Calculator.StructuralCalc", "Prefs.StructuralCalc", , , "Prefs.Feature.Calculator")
        PreferencesWindow._Check("Features.System.ConfirmActions", "Prefs.ConfirmActions", , , "Prefs.Feature.System")
        PreferencesWindow._Gap()
        PreferencesWindow._Field("Features.Terminal.Prefix", "Prefs.TerminalPrefix", "S")
        PreferencesWindow._Choice("Features.Terminal.Shell", "Prefs.TerminalShell", ["cmd", "powershell", "pwsh", "wt"], ["Command Prompt (cmd)", "Windows PowerShell", "PowerShell 7 (pwsh)", "Windows Terminal (wt)"])
    }

    static _BuildApplications() {
        PreferencesWindow._BeginPage("Prefs.Page.Applications")
        PreferencesWindow._Lines("Features.Applications.Folders", "Prefs.AppFolders", 3)
        PreferencesWindow._Csv("Features.Applications.FileTypes", "Prefs.AppFileTypes", "L")
        PreferencesWindow._Field("Features.Applications.Depth", "Prefs.AppDepth", "S", "number")
        PreferencesWindow._Field("Features.Applications.Exclude", "Prefs.AppExclude", "L")
        PreferencesWindow._Lines("Features.Applications.Hidden", "Prefs.AppHidden", 3)
        refreshY := PreferencesWindow._y                                    ; 按钮和 "索引刷新间隔" 放在同一行, 页面才放得下
        PreferencesWindow._Field("Features.Applications.RefreshMinutes", "Prefs.RefreshMinutes", "S", "number")
        PreferencesWindow._SideButton(refreshY, "Prefs.RebuildIndex", (*) => App.RebuildIndex())
        PreferencesWindow._Check("Features.Applications.StoreApps", "Prefs.StoreApps", , , "Prefs.Group.Options")
        PreferencesWindow._Check("Features.Applications.MatchPinyin", "Prefs.MatchPinyin")
    }

    static _BuildFileSearch() {
        PreferencesWindow._BeginPage("Prefs.Page.FileSearch")
        PreferencesWindow._Check("Features.FileSearch.SpacePrefix", "Prefs.SpacePrefix", , , "Prefs.Group.StartFileSearch")
        PreferencesWindow._Check("Features.FileSearch.QuotePrefix", "Prefs.QuotePrefix")
        PreferencesWindow._Csv("Features.FileSearch.Keywords", "Prefs.FileKeywords", "M")
        PreferencesWindow._Csv("Features.FileSearch.FolderKeywords", "Prefs.FolderKeywords", "M")
        PreferencesWindow._Field("Features.FileSearch.MaxResults", "Prefs.FileMaxResults", "S", "number")
        PreferencesWindow._Check("Features.FileSearch.InDefaultResults", "Prefs.FileInDefault", , , "Prefs.Group.NormalSearch")
        PreferencesWindow._Field("Features.FileSearch.DefaultResultsLimit", "Prefs.FileDefaultLimit", "S", "number")
        status := I18n.T(Everything.IsRunning() ? "Prefs.Running" : "Prefs.NotRunning")
        PreferencesWindow._Check("Features.FileSearch.UseEverything", "Prefs.UseEverything", , I18n.T("Prefs.EverythingStatus", status), "Prefs.Group.Everything")
        PreferencesWindow._Field("Features.FileSearch.EverythingFilter", "Prefs.EverythingFilter", "L")
        PreferencesWindow._Field("Features.FileSearch.EverythingPath", "Prefs.EverythingPath", "L", "folder")
        PreferencesWindow._Lines("Features.FileSearch.ScopeFolders", "Prefs.ScopeFolders", 2)
        depthY := PreferencesWindow._y                                      ; 按钮和 "子文件夹深度" 放在同一行, 页面才放得下
        PreferencesWindow._Field("Features.FileSearch.ScopeDepth", "Prefs.ScopeDepth", "S", "number")
        PreferencesWindow._SideButton(depthY, "Prefs.RebuildFileIndex", (*) => FileIndex.Rebuild())
    }

    static _BuildCommands() {
        PreferencesWindow._BeginPage("Prefs.Page.Commands")
        PreferencesWindow._List("CustomCommands", 500
            , [["Prefs.Col.Title", "Title", 140], ["Prefs.Col.Type", "Type", 70], ["Prefs.Col.Target", "Target", 190], ["Prefs.Col.Keyword", "Keyword", 70], ["Prefs.Col.Status", "Status", 70]]
            , CustomCommandProvider.EditorFields(), () => CustomCommandProvider.NewCommand()
            , (item, roots) => CustomCommandProvider.CheckTarget(item, roots))
        PreferencesWindow._Note("Prefs.CommandsNote")
    }

    static _BuildSnippets() {
        PreferencesWindow._BeginPage("Prefs.Page.Snippets")
        PreferencesWindow._List("Snippets", 330
            , [["Prefs.Col.Name", "Name", 150], ["Prefs.Col.Keyword", "Keyword", 90], ["Prefs.Col.Text", "Text", 300]]
            , SnippetProvider.EditorFields(), () => SnippetProvider.NewSnippet())
        PreferencesWindow._Check("Features.Snippets.AutoExpand", "Prefs.SnippetAutoExpand", , , "Prefs.Group.AutoExpand")
        PreferencesWindow._Field("Features.Snippets.ExpandPrefix", "Prefs.ExpandPrefix", "S")
        PreferencesWindow._Field("Features.Snippets.Keyword", "Prefs.SnippetKeyword", "M")
        PreferencesWindow._Choice("Features.Snippets.PasteMode", "Prefs.PasteMode", ["Clipboard", "Type"], [I18n.T("Prefs.PasteMode.Clipboard"), I18n.T("Prefs.PasteMode.Type")])
    }

    static _BuildClipboard() {
        PreferencesWindow._BeginPage("Prefs.Page.Clipboard")
        PreferencesWindow._Field("Features.Clipboard.Hotkey", "Prefs.ClipHotkey", "M")
        PreferencesWindow._Field("Features.Clipboard.Keyword", "Prefs.ClipKeyword", "M")
        PreferencesWindow._Field("Features.Clipboard.MaxItems", "Prefs.ClipMaxItems", "S", "number")
        PreferencesWindow._Check("Features.Clipboard.Persist", "Prefs.ClipPersist", , , "Prefs.Group.History")
        PreferencesWindow._Gap()
        PreferencesWindow._Lines("Features.Clipboard.IgnoreApps", "Prefs.ClipIgnoreApps", 5)
        PreferencesWindow._Button("Prefs.ClipClear", (*) => ClipboardProvider.Clear())
    }

    static _BuildWebSearch() {
        PreferencesWindow._BeginPage("Prefs.Page.WebSearch")
        PreferencesWindow._List("Features.WebSearch.Engines", 420
            , [["Prefs.Col.Keyword", "Keyword", 70], ["Prefs.Col.Title", "Title", 130], ["Prefs.Col.Url", "Url", 340]]
            , WebSearchProvider.EditorFields(), () => WebSearchProvider.NewEngine())
        PreferencesWindow._Csv("Features.WebSearch.Fallbacks", "Prefs.Fallbacks", "L")
    }

    static _BuildHotkeys() {
        PreferencesWindow._BeginPage("Prefs.Page.Hotkeys")
        actions := [["ToggleWindow", "Show / hide ALTRun (ToggleWindow)"]]
        for command in SystemProvider.Commands()
            actions.Push([command["Id"], command["Title"] " (" command["Id"] ")"])
        PreferencesWindow._List("Hotkeys", 480
            , [["Prefs.Col.Key", "Key", 110], ["Prefs.Col.Action", "Action", 160], ["Prefs.Col.WinTitle", "WinTitle", 270]]
            , [ItemEditor.Field("Key", "Prefs.Col.Key", "text", true, "", I18n.T("Prefs.HotkeyHint"))
             , ItemEditor.Field("Action", "Prefs.Col.Action", "choice", false, actions)
             , ItemEditor.Field("WinTitle", "Prefs.Col.WinTitle", "text", false, "", "ahk_exe RAPTW.exe")]
            , () => Map("Key", "", "Action", "ToggleWindow", "WinTitle", ""))
        PreferencesWindow._Note("Prefs.HotkeysNote")
    }

    static _BuildExtensions() {
        PreferencesWindow._BeginPage("Prefs.Page.Extensions")
        PreferencesWindow._Section("Prefs.QuickSwitch")
        PreferencesWindow._Check("Extensions.QuickSwitch.Enabled", "Prefs.EnableExtension", , , "Prefs.Group.Status")
        PreferencesWindow._Field("Extensions.QuickSwitch.TotalCmdHotkey", "Prefs.QSTotalCmd", "M")
        PreferencesWindow._Field("Extensions.QuickSwitch.ExplorerHotkey", "Prefs.QSExplorer", "M")
        PreferencesWindow._Check("Extensions.QuickSwitch.AutoSwitch", "Prefs.QSAuto", , , "Prefs.Group.Options")
        PreferencesWindow._Section("Prefs.AutoDate")
        PreferencesWindow._Check("Extensions.AutoDate.Enabled", "Prefs.EnableExtension", , , "Prefs.Group.Status")
        PreferencesWindow._Field("Extensions.AutoDate.DateFormat", "Prefs.DateFormat", "M")
        PreferencesWindow._Field("Extensions.AutoDate.RenameHotkey", "Prefs.RenameHotkey", "M")
        PreferencesWindow._Field("Extensions.AutoDate.AppendHotkey", "Prefs.AppendHotkey", "M")
    }

    static _BuildAdvanced() {
        PreferencesWindow._BeginPage("Prefs.Page.Advanced")
        PreferencesWindow._Section("Prefs.SettingsFile")
        PreferencesWindow._Info("Prefs.Group.File", AppSettings.File, " cGray")
        PreferencesWindow._Button("Prefs.EditJson", (*) => (PreferencesWindow.Cancel() || App.EditSettingsFile()))
        PreferencesWindow._Button("Prefs.OpenDataFolder", (*) => PreferencesWindow._OpenFolder(AppSettings.DataDir))
        PreferencesWindow._Button("Prefs.ResetLearning", (*) => PreferencesWindow._ResetLearning())
        PreferencesWindow._Section("Sys.About")
        PreferencesWindow._Info("Prefs.Group.Program", App.Name " - " I18n.T("App.Tagline"))
        PreferencesWindow._Info("Prefs.Group.Version", App.Version)
        PreferencesWindow._Info("Prefs.Group.Homepage", '<a href="' App.RepoUrl '">' App.RepoUrl '</a>', "", "Link")
        PreferencesWindow._Button("Tray.CheckUpdate", (*) => UpdateChecker.Check(false))
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
        Knowledge.Picks := Map(), Knowledge.QueryPicks := Map(), Knowledge.History := []
        Knowledge.Save()
        MsgBox(I18n.T("Prefs.ResetDone"), App.Name, 64)
    }

    static _OpenFolder(folder) {
        DirCreate(folder)
        Run('explorer.exe "' folder '"')
    }

    ;---------------------------------------------------------------------------
    ; Builders
    ;---------------------------------------------------------------------------
    static _BeginPage(nameKey) {
        PreferencesWindow.Pages.Push({Name: I18n.T(nameKey), Wiki: PreferencesWindow.WikiPage(nameKey), Controls: []})
        PreferencesWindow._y := 14
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
    }

    ; 左列的分组标签, 和右边第一个控件同一行
    static _GroupLabel(labelKey, dy := 1) {
        PreferencesWindow._y += dy
        ctrl := PreferencesWindow._Add("Text", "x" PreferencesWindow.ContentX " w" (PreferencesWindow.LabelW - 12) " Right", I18n.T(labelKey))
        PreferencesWindow._y -= dy
        return ctrl
    }

    ; 左列标签 + 右边一段文字 (或链接), 例如 "版本: 2026.09.26"
    static _Info(labelKey, text, options := "", type := "Text") {
        label := PreferencesWindow._GroupLabel(labelKey)
        ctrl := PreferencesWindow._Add(type, "x" PreferencesWindow._InputX() " w" PreferencesWindow._InputW() options, text)
        PreferencesWindow._Below(6, label, ctrl)
    }

    ; kind: text / number / file / folder; size: "S" / "M" / "L" (见 _Width)
    static _Field(path, labelKey, size, kind := "text") {
        label := PreferencesWindow._Label(labelKey)
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

    static _Choice(path, labelKey, values, labels) {
        label := PreferencesWindow._Label(labelKey)
        ctrl := PreferencesWindow._Add("DropDownList", "x" PreferencesWindow._InputX() " w" PreferencesWindow.WidthM, labels)
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
        ctrl := PreferencesWindow._Add("Edit", "x" PreferencesWindow._InputX() " w" PreferencesWindow._InputW() " r" rows " +Multi -Wrap +HScroll", value)
        PreferencesWindow._Bind(path, () => PreferencesWindow.SplitLines(ctrl.Value))
        PreferencesWindow._Below(3, label, ctrl)
        PreferencesWindow._Desc(labelKey, PreferencesWindow._InputX())
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

        Refresh(selectRow := 0) {
            listView.Delete()
            for item in items {
                cells := []
                for column in columns
                    cells.Push((column[2] = "Status") ? StatusText(item) : PreferencesWindow._Cell(item, column[2]))
                listView.Add("", cells*)
            }
            if selectRow
                listView.Modify(Min(selectRow, listView.GetCount()), "Select Focus Vis")
        }
        StatusText(item) {
            key := ObjPtr(item)
            return statuses.Has(key) ? I18n.T("Prefs.Status." statuses[key]) : ""
        }
        Recheck(item) {                                                     ; 检查过之后, 新增 / 修改的项目也马上检查
            if statuses.Count
                statuses[ObjPtr(item)] := check(item, Map())
        }
        AddItem(*) {
            edited := ItemEditor.Edit(PreferencesWindow.Gui, pageName, fields, newItem())
            if IsObject(edited) {
                items.Push(edited)
                PreferencesWindow.MarkDirty()
                Recheck(edited)
                Refresh(items.Length)
            }
        }
        EditItem(*) {
            row := listView.GetNext()
            if !row
                return
            edited := ItemEditor.Edit(PreferencesWindow.Gui, pageName, fields, ItemEditor.WithDefaults(items[row], newItem()))
            if IsObject(edited) {
                if statuses.Has(ObjPtr(items[row]))
                    statuses.Delete(ObjPtr(items[row]))
                items[row] := edited
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
            Refresh(firstProblem)
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
            items.RemoveAt(row)
            PreferencesWindow.MarkDirty()
            Refresh(row)
        }

        listView.OnEvent("DoubleClick", EditItem)
        Refresh()
        PreferencesWindow._y += height - 34
        PreferencesWindow._Add("Button", "w80", I18n.T("Prefs.Add")).OnEvent("Click", AddItem)
        PreferencesWindow._Add("Button", "x+8 yp w80", I18n.T("Prefs.Edit")).OnEvent("Click", EditItem)
        PreferencesWindow._Add("Button", "x+8 yp w80", I18n.T("Prefs.Delete")).OnEvent("Click", DeleteItem)
        if IsObject(check)
            PreferencesWindow._Add("Button", "x+24 yp w100", I18n.T("Prefs.CheckTargets")).OnEvent("Click", CheckItems)
        PreferencesWindow._y += 38
    }

    static _Cell(item, key) {
        if !(item is Map) || !item.Has(key)
            return ""
        return RegExReplace(item[key] "", "\s+", " ")
    }

    ;---------------------------------------------------------------------------
    ; Data helpers (纯函数, 有单元测试)
    ;---------------------------------------------------------------------------
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

    static JoinCsv(items) {
        if !(items is Array)
            return items
        out := ""
        for index, item in items
            out .= (index = 1 ? "" : ", ") item
        return out
    }
}
