;===============================================================================
; PreferencesWindow.ahk - 偏好设置窗口 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 Alfred 的 Preferences 一样左边是分类, 右边是该分类的设置。
; 所有修改先作用在设置的一份副本 (Working) 上: "保存" 写回 ALTRun.json 并重新
; 载入 ALTRun (重新打开到同一页), "取消" 直接丢弃。
;
; 页面构建用几个小工具, 每个设置项只写一行:
;   _Check("General.LaunchAtLogin", "Prefs.LaunchAtLogin")      复选框
;   _Field("General.Hotkey", "Prefs.Hotkey", 160)               文本框
;   _Choice("General.Language", "Prefs.Language", 值数组, 显示文字数组)
;   _Lines("Features.Applications.Folders", "Prefs.AppFolders", 行数)   数组 <-> 每行一项
;   _Csv("Features.FileSearch.Keywords", "Prefs.FileKeywords")        数组 <-> 逗号分隔
;   _List(...)                                                  列表 + 添加/编辑/删除 (ItemEditor)
; 路径是 ALTRun.json 里的键, 用 "." 连接。
;
; 用法:
;   PreferencesWindow.Show([页码])
;===============================================================================

class PreferencesWindow {
    static Gui := "", Working := "", Pages := [], Binds := [], PageList := ""
    static _page := 0, _y := 0
    static ContentX := 190, ContentW := 560, LabelW := 215

    static Show(pageIndex := 1) {
        if IsObject(PreferencesWindow.Gui) {
            WinActivate("ahk_id " PreferencesWindow.Gui.Hwnd)
            return
        }
        SearchWindow.Hide()
        PreferencesWindow.Working := PreferencesWindow.DeepCopy(AppSettings.Data)
        PreferencesWindow.Pages := [], PreferencesWindow.Binds := []

        g := Gui("-MinimizeBox", I18n.T("Prefs.Title"))
        g.SetFont("s9", ThemeManager.FontName())
        g.OnEvent("Close", (*) => PreferencesWindow.Close())
        g.OnEvent("Escape", (*) => PreferencesWindow.Close())
        PreferencesWindow.Gui := g

        PreferencesWindow._BuildPages()

        names := []
        for page in PreferencesWindow.Pages
            names.Push(page.Name)
        pageList := g.AddListBox("x12 y12 w160 h470 AltSubmit", names)
        pageList.OnEvent("Change", (ctrl, *) => PreferencesWindow.SelectPage(ctrl.Value))
        SendMessage(0x1A0, 0, Round(26 * A_ScreenDPI / 96), pageList.Hwnd)  ; LB_SETITEMHEIGHT: 更宽松的侧边栏
        PreferencesWindow.PageList := pageList

        g.AddText("x" PreferencesWindow.ContentX " y500 w360 cGray", I18n.T("Prefs.SaveHint"))
        g.AddButton("x570 y494 w85 Default", I18n.T("Prefs.Save")).OnEvent("Click", (*) => PreferencesWindow.Save())
        g.AddButton("x665 y494 w85", I18n.T("Prefs.Cancel")).OnEvent("Click", (*) => PreferencesWindow.Close())

        pageIndex := Max(1, Min(pageIndex, PreferencesWindow.Pages.Length))
        pageList.Value := pageIndex
        PreferencesWindow.SelectPage(pageIndex)
        g.Show("w765 h535")
    }

    static SelectPage(pageIndex) {
        PreferencesWindow._page := pageIndex
        for index, page in PreferencesWindow.Pages
            for ctrl in page.Controls
                ctrl.Visible := (index = pageIndex)
    }

    static Close() {
        if IsObject(PreferencesWindow.Gui)
            PreferencesWindow.Gui.Destroy()
        PreferencesWindow.Gui := ""
    }

    static Save() {
        for bind in PreferencesWindow.Binds
            PreferencesWindow.SetPath(PreferencesWindow.Working, bind.Path, bind.Read.Call())
        general := PreferencesWindow.Working["General"]
        appearance := PreferencesWindow.Working["Appearance"]
        appearance["VisibleRows"] := Max(1, Min(appearance["VisibleRows"], 9))
        appearance["Width"] := Max(400, appearance["Width"])
        if (general["Hotkey"] = "")
            general["Hotkey"] := "!Space"

        AppSettings.Data := PreferencesWindow.Working
        if !AppSettings.Save()
            return
        page := PreferencesWindow._page
        PreferencesWindow.Close()
        App.Restart("-Preferences " page)
    }

    ;---------------------------------------------------------------------------
    ; Pages
    ;---------------------------------------------------------------------------
    static _BuildPages() {
        PreferencesWindow._BuildGeneral()
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
        PreferencesWindow._BuildAdvanced()
    }

    static _BuildGeneral() {
        PreferencesWindow._BeginPage("Prefs.Page.General")
        PreferencesWindow._Field("General.Hotkey", "Prefs.Hotkey", 160)
        PreferencesWindow._Field("General.SecondaryHotkey", "Prefs.SecondaryHotkey", 160)
        PreferencesWindow._Hint("Prefs.HotkeyHint")
        PreferencesWindow._Choice("General.Language", "Prefs.Language", ["auto", "en", "zh"], [I18n.T("Prefs.Language.auto"), "English", "中文"])
        PreferencesWindow._Gap()
        for pair in [["LaunchAtLogin", "Prefs.LaunchAtLogin"], ["ShowTrayIcon", "Prefs.ShowTrayIcon"]
                    , ["HideOnDeactivate", "Prefs.HideOnDeactivate"], ["SwitchToEnglishInput", "Prefs.EnglishInput"]
                    , ["SendToMenu", "Prefs.SendToMenu"], ["StartMenuShortcut", "Prefs.StartMenu"]
                    , ["CheckForUpdates", "Prefs.CheckUpdates"], ["SaveLog", "Prefs.SaveLog"]]
            PreferencesWindow._Check("General." pair[1], pair[2])
        PreferencesWindow._Gap()
        PreferencesWindow._Field("General.FileManager", "Prefs.FileManager", 300, "file")
        PreferencesWindow._Field("General.HistorySize", "Prefs.HistorySize", 60, "number")
    }

    static _BuildAppearance() {
        PreferencesWindow._BeginPage("Prefs.Page.Appearance")
        themes := ["Light", "Dark"]
        Loop Files, A_ScriptDir "\Themes\*.json" {
            SplitPath(A_LoopFileName, , , , &themeName)
            themes.Push(themeName)
        }
        PreferencesWindow._Choice("Appearance.Theme", "Prefs.Theme", themes, themes)
        PreferencesWindow._Hint("Prefs.ThemeHint")
        PreferencesWindow._Field("Appearance.Width", "Prefs.Width", 80, "number")
        PreferencesWindow._Field("Appearance.VisibleRows", "Prefs.VisibleRows", 80, "number")
        PreferencesWindow._Gap()
        PreferencesWindow._Button("Prefs.OpenThemes", (*) => PreferencesWindow._OpenFolder(A_ScriptDir "\Themes"))
    }

    static _BuildFeatures() {
        PreferencesWindow._BeginPage("Prefs.Page.Features")
        PreferencesWindow._Section("Prefs.EnabledFeatures")
        features := ["Applications", "CustomCommands", "Snippets", "Clipboard", "Calculator", "WebSearch", "FileSearch", "Terminal", "System"]
        startY := PreferencesWindow._y
        for index, feature in features {
            column := Mod(index - 1, 3), row := (index - 1) // 3
            PreferencesWindow._y := startY + row * 26
            PreferencesWindow._Check("Features." feature ".Enabled", "Prefs.Feature." feature, PreferencesWindow.ContentX + column * 185)
        }
        PreferencesWindow._y := startY + 3 * 26 + 10
        PreferencesWindow._Check("Features.Calculator.StructuralCalc", "Prefs.StructuralCalc")
        PreferencesWindow._Check("Features.System.ConfirmActions", "Prefs.ConfirmActions")
        PreferencesWindow._Gap()
        PreferencesWindow._Field("Features.Terminal.Prefix", "Prefs.TerminalPrefix", 60)
        PreferencesWindow._Choice("Features.Terminal.Shell", "Prefs.TerminalShell", ["cmd", "powershell", "pwsh", "wt"], ["Command Prompt (cmd)", "Windows PowerShell", "PowerShell 7 (pwsh)", "Windows Terminal (wt)"])
    }

    static _BuildApplications() {
        PreferencesWindow._BeginPage("Prefs.Page.Applications")
        PreferencesWindow._Lines("Features.Applications.Folders", "Prefs.AppFolders", 6)
        PreferencesWindow._Csv("Features.Applications.FileTypes", "Prefs.AppFileTypes", 300)
        PreferencesWindow._Field("Features.Applications.Depth", "Prefs.AppDepth", 60, "number")
        PreferencesWindow._Field("Features.Applications.Exclude", "Prefs.AppExclude", 300)
        PreferencesWindow._Field("Features.Applications.RefreshMinutes", "Prefs.RefreshMinutes", 60, "number")
        PreferencesWindow._Check("Features.Applications.StoreApps", "Prefs.StoreApps")
        PreferencesWindow._Check("Features.Applications.MatchPinyin", "Prefs.MatchPinyin")
        PreferencesWindow._Gap()
        PreferencesWindow._Button("Tray.RebuildIndex", (*) => App.RebuildIndex())
    }

    static _BuildFileSearch() {
        PreferencesWindow._BeginPage("Prefs.Page.FileSearch")
        status := I18n.T(Everything.IsRunning() ? "Prefs.Running" : "Prefs.NotRunning")
        PreferencesWindow._Below(8, PreferencesWindow._Add("Text", "w" PreferencesWindow.ContentW " cGray", I18n.T("Prefs.EverythingStatus", status)))
        PreferencesWindow._Check("Features.FileSearch.InDefaultResults", "Prefs.FileInDefault")
        PreferencesWindow._Field("Features.FileSearch.DefaultResultsLimit", "Prefs.FileDefaultLimit", 60, "number")
        PreferencesWindow._Csv("Features.FileSearch.Keywords", "Prefs.FileKeywords", 200)
        PreferencesWindow._Check("Features.FileSearch.QuotePrefix", "Prefs.QuotePrefix")
        PreferencesWindow._Field("Features.FileSearch.MaxResults", "Prefs.FileMaxResults", 60, "number")
        PreferencesWindow._Gap()
        PreferencesWindow._Check("Features.FileSearch.UseEverything", "Prefs.UseEverything")
        PreferencesWindow._Field("Features.FileSearch.EverythingFilter", "Prefs.EverythingFilter", 300)
        PreferencesWindow._Field("Features.FileSearch.EverythingPath", "Prefs.EverythingPath", 300, "folder")
        PreferencesWindow._Gap()
        PreferencesWindow._Lines("Features.FileSearch.ScopeFolders", "Prefs.ScopeFolders", 3)
        PreferencesWindow._Field("Features.FileSearch.ScopeDepth", "Prefs.ScopeDepth", 60, "number")
        PreferencesWindow._Button("Prefs.RebuildFileIndex", (*) => FileIndex.Rebuild())
    }

    static _BuildCommands() {
        PreferencesWindow._BeginPage("Prefs.Page.Commands")
        PreferencesWindow._List("CustomCommands", 440
            , [["Prefs.Col.Title", "Title", 150], ["Prefs.Col.Type", "Type", 70], ["Prefs.Col.Target", "Target", 230], ["Prefs.Col.Keyword", "Keyword", 90]]
            , CustomCommandProvider.EditorFields(), () => CustomCommandProvider.NewCommand())
    }

    static _BuildSnippets() {
        PreferencesWindow._BeginPage("Prefs.Page.Snippets")
        PreferencesWindow._List("Snippets", 300
            , [["Prefs.Col.Name", "Name", 150], ["Prefs.Col.Keyword", "Keyword", 90], ["Prefs.Col.Text", "Text", 300]]
            , SnippetProvider.EditorFields(), () => SnippetProvider.NewSnippet())
        PreferencesWindow._Check("Features.Snippets.AutoExpand", "Prefs.SnippetAutoExpand")
        PreferencesWindow._Field("Features.Snippets.ExpandPrefix", "Prefs.ExpandPrefix", 60)
        PreferencesWindow._Field("Features.Snippets.Keyword", "Prefs.SnippetKeyword", 100)
        PreferencesWindow._Choice("Features.Snippets.PasteMode", "Prefs.PasteMode", ["Clipboard", "Type"], [I18n.T("Prefs.PasteMode.Clipboard"), I18n.T("Prefs.PasteMode.Type")])
    }

    static _BuildClipboard() {
        PreferencesWindow._BeginPage("Prefs.Page.Clipboard")
        PreferencesWindow._Field("Features.Clipboard.Hotkey", "Prefs.ClipHotkey", 160)
        PreferencesWindow._Field("Features.Clipboard.Keyword", "Prefs.ClipKeyword", 100)
        PreferencesWindow._Field("Features.Clipboard.MaxItems", "Prefs.ClipMaxItems", 80, "number")
        PreferencesWindow._Check("Features.Clipboard.Persist", "Prefs.ClipPersist")
        PreferencesWindow._Gap()
        PreferencesWindow._Lines("Features.Clipboard.IgnoreApps", "Prefs.ClipIgnoreApps", 5)
        PreferencesWindow._Gap()
        PreferencesWindow._Button("Prefs.ClipClear", (*) => ClipboardProvider.Clear())
    }

    static _BuildWebSearch() {
        PreferencesWindow._BeginPage("Prefs.Page.WebSearch")
        PreferencesWindow._List("Features.WebSearch.Engines", 400
            , [["Prefs.Col.Keyword", "Keyword", 70], ["Prefs.Col.Title", "Title", 130], ["Prefs.Col.Url", "Url", 340]]
            , WebSearchProvider.EditorFields(), () => WebSearchProvider.NewEngine())
        PreferencesWindow._Csv("Features.WebSearch.Fallbacks", "Prefs.Fallbacks", 300)
    }

    static _BuildHotkeys() {
        PreferencesWindow._BeginPage("Prefs.Page.Hotkeys")
        actions := [["ToggleWindow", "Show / hide ALTRun (ToggleWindow)"]]
        for command in SystemProvider.Commands()
            actions.Push([command["Id"], command["Title"] " (" command["Id"] ")"])
        PreferencesWindow._List("Hotkeys", 400
            , [["Prefs.Col.Key", "Key", 110], ["Prefs.Col.Action", "Action", 160], ["Prefs.Col.WinTitle", "WinTitle", 270]]
            , [ItemEditor.Field("Key", "Prefs.Col.Key", "text", true, "", I18n.T("Prefs.HotkeyHint"))
             , ItemEditor.Field("Action", "Prefs.Col.Action", "choice", false, actions)
             , ItemEditor.Field("WinTitle", "Prefs.Col.WinTitle", "text", false, "", "ahk_exe RAPTW.exe")]
            , () => Map("Key", "", "Action", "ToggleWindow", "WinTitle", ""))
    }

    static _BuildExtensions() {
        PreferencesWindow._BeginPage("Prefs.Page.Extensions")
        PreferencesWindow._Section("Prefs.QuickSwitch")
        PreferencesWindow._Check("Extensions.QuickSwitch.Enabled", "Prefs.EnableExtension")
        PreferencesWindow._Field("Extensions.QuickSwitch.TotalCmdHotkey", "Prefs.QSTotalCmd", 100)
        PreferencesWindow._Field("Extensions.QuickSwitch.ExplorerHotkey", "Prefs.QSExplorer", 100)
        PreferencesWindow._Check("Extensions.QuickSwitch.AutoSwitch", "Prefs.QSAuto")
        PreferencesWindow._Gap()
        PreferencesWindow._Section("Prefs.AutoDate")
        PreferencesWindow._Check("Extensions.AutoDate.Enabled", "Prefs.EnableExtension")
        PreferencesWindow._Field("Extensions.AutoDate.DateFormat", "Prefs.DateFormat", 120)
        PreferencesWindow._Field("Extensions.AutoDate.RenameHotkey", "Prefs.RenameHotkey", 100)
        PreferencesWindow._Field("Extensions.AutoDate.AppendHotkey", "Prefs.AppendHotkey", 100)
    }

    static _BuildAdvanced() {
        PreferencesWindow._BeginPage("Prefs.Page.Advanced")
        PreferencesWindow._Section("Prefs.SettingsFile")
        PreferencesWindow._Below(8, PreferencesWindow._Add("Text", "w" PreferencesWindow.ContentW " cGray", AppSettings.File))
        PreferencesWindow._Button("Prefs.EditJson", (*) => (PreferencesWindow.Close(), App.EditSettingsFile()))
        PreferencesWindow._Button("Prefs.OpenDataFolder", (*) => PreferencesWindow._OpenFolder(AppSettings.DataDir))
        PreferencesWindow._Button("Prefs.ResetLearning", (*) => PreferencesWindow._ResetLearning())
        PreferencesWindow._Gap()
        PreferencesWindow._Section("Sys.About")
        PreferencesWindow._Below(6, PreferencesWindow._Add("Text", "w" PreferencesWindow.ContentW, App.Name " - " I18n.T("App.Tagline")))
        PreferencesWindow._Below(6, PreferencesWindow._Add("Text", "w" PreferencesWindow.ContentW, I18n.T("Prefs.Version", App.Version)))
        PreferencesWindow._Below(10, PreferencesWindow._Add("Link", "w" PreferencesWindow.ContentW, '<a href="' App.RepoUrl '">' App.RepoUrl '</a>'))
        PreferencesWindow._Button("Tray.CheckUpdate", (*) => UpdateChecker.Check(false))
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
        PreferencesWindow.Pages.Push({Name: I18n.T(nameKey), Controls: []})
        PreferencesWindow._y := 14
    }

    static _Add(type, options, text := "") {
        pos := InStr(options, " x") || InStr(options, "x") = 1 ? "" : "x" PreferencesWindow.ContentX " "
        ctrl := PreferencesWindow.Gui.Add(type, pos "y" PreferencesWindow._y " " options " Hidden", text)
        PreferencesWindow.Pages[PreferencesWindow.Pages.Length].Controls.Push(ctrl)
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

    static _Section(labelKey) {
        PreferencesWindow.Gui.SetFont("bold")
        ctrl := PreferencesWindow._Add("Text", "w" PreferencesWindow.ContentW, I18n.T(labelKey))
        PreferencesWindow.Gui.SetFont("norm")
        PreferencesWindow._Below(8, ctrl)
    }

    static _Hint(textKey) {
        PreferencesWindow._y -= 2
        ctrl := PreferencesWindow._Add("Text", "x" PreferencesWindow._InputX() " w" (PreferencesWindow.ContentW - PreferencesWindow.LabelW) " cGray", I18n.T(textKey))
        PreferencesWindow._Below(8, ctrl)
    }

    static _Gap() {
        PreferencesWindow._y += 10
    }

    ; 标签和输入框放在同一行: 标签往下挪 3 像素和输入框的文字对齐
    static _Label(labelKey) {
        PreferencesWindow._y += 3
        ctrl := PreferencesWindow._Add("Text", "w" (PreferencesWindow.LabelW - 8) " Right", I18n.T(labelKey))
        PreferencesWindow._y -= 3
        return ctrl
    }

    static _InputX() {
        return PreferencesWindow.ContentX + PreferencesWindow.LabelW
    }

    static _Check(path, labelKey, x := 0) {
        x := x ? x : PreferencesWindow.ContentX
        ctrl := PreferencesWindow._Add("Checkbox", "x" x " w" (PreferencesWindow.ContentW - (x - PreferencesWindow.ContentX)), I18n.T(labelKey))
        ctrl.Value := PreferencesWindow.GetPath(PreferencesWindow.Working, path) ? 1 : 0
        PreferencesWindow._Bind(path, () => ctrl.Value)
        PreferencesWindow._Below(6, ctrl)
    }

    ; kind: text / number / file / folder
    static _Field(path, labelKey, width, kind := "text") {
        label := PreferencesWindow._Label(labelKey)
        value := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        ctrl := PreferencesWindow._Add("Edit", "x" PreferencesWindow._InputX() " w" width " r1 -Multi" (kind = "number" ? " Number" : ""), value)
        if (kind = "file" || kind = "folder") {
            browse := PreferencesWindow._Add("Button", "x+4 yp-1 w30", I18n.T("Prefs.Browse"))
            browse.OnEvent("Click", ItemEditor._Browser(ctrl, kind, PreferencesWindow.Gui))
        }
        if (kind = "number")
            PreferencesWindow._Bind(path, () => IsInteger(ctrl.Value) ? Integer(ctrl.Value) : 0)
        else
            PreferencesWindow._Bind(path, () => ctrl.Value)
        PreferencesWindow._Below(8, label, ctrl)
    }

    static _Choice(path, labelKey, values, labels) {
        label := PreferencesWindow._Label(labelKey)
        ctrl := PreferencesWindow._Add("DropDownList", "x" PreferencesWindow._InputX() " w240", labels)
        current := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        ctrl.Value := 1
        for index, value in values
            if (value = current)
                ctrl.Value := index
        PreferencesWindow._Bind(path, () => values[ctrl.Value])
        PreferencesWindow._Below(8, label, ctrl)
    }

    static _Lines(path, labelKey, rows) {
        label := PreferencesWindow._Label(labelKey)
        value := PreferencesWindow.JoinLines(PreferencesWindow.GetPath(PreferencesWindow.Working, path))
        ctrl := PreferencesWindow._Add("Edit", "x" PreferencesWindow._InputX() " w" (PreferencesWindow.ContentW - PreferencesWindow.LabelW) " r" rows " +Multi", value)
        PreferencesWindow._Bind(path, () => PreferencesWindow.SplitLines(ctrl.Value))
        PreferencesWindow._Below(4, label, ctrl)
        PreferencesWindow._Hint("Prefs.ListHint")
    }

    static _Csv(path, labelKey, width) {
        label := PreferencesWindow._Label(labelKey)
        value := PreferencesWindow.JoinCsv(PreferencesWindow.GetPath(PreferencesWindow.Working, path))
        ctrl := PreferencesWindow._Add("Edit", "x" PreferencesWindow._InputX() " w" width " r1 -Multi", value)
        PreferencesWindow._Bind(path, () => PreferencesWindow.SplitCsv(ctrl.Value))
        PreferencesWindow._Below(8, label, ctrl)
    }

    static _Button(labelKey, fn) {
        ctrl := PreferencesWindow._Add("Button", "w220 h28", I18n.T(labelKey))
        ctrl.OnEvent("Click", fn)
        PreferencesWindow._Below(6, ctrl)
    }

    ; 列表页: ListView 显示 path 指向的数组, 添加 / 编辑 / 删除都用 ItemEditor
    static _List(path, height, columns, fields, newItem) {
        items := PreferencesWindow.GetPath(PreferencesWindow.Working, path)
        headers := []
        for column in columns
            headers.Push(I18n.T(column[1]))
        listView := PreferencesWindow._Add("ListView", "w" PreferencesWindow.ContentW " h" (height - 40) " -Multi NoSort Grid", headers)
        for index, column in columns
            listView.ModifyCol(index, column[3])
        pageName := PreferencesWindow.Pages[PreferencesWindow.Pages.Length].Name

        Refresh(selectRow := 0) {
            listView.Delete()
            for item in items {
                cells := []
                for column in columns
                    cells.Push(PreferencesWindow._Cell(item, column[2]))
                listView.Add("", cells*)
            }
            if selectRow
                listView.Modify(Min(selectRow, listView.GetCount()), "Select Focus Vis")
        }
        AddItem(*) {
            edited := ItemEditor.Edit(PreferencesWindow.Gui, pageName, fields, newItem())
            if IsObject(edited) {
                items.Push(edited)
                Refresh(items.Length)
            }
        }
        EditItem(*) {
            row := listView.GetNext()
            if !row
                return
            edited := ItemEditor.Edit(PreferencesWindow.Gui, pageName, fields, ItemEditor.WithDefaults(items[row], newItem()))
            if IsObject(edited) {
                items[row] := edited
                Refresh(row)
            }
        }
        DeleteItem(*) {
            row := listView.GetNext()
            if !row
                return
            if (MsgBox(I18n.T("Prefs.ConfirmDelete", listView.GetText(row, 1)), pageName, "YesNo Icon? Owner" PreferencesWindow.Gui.Hwnd) != "Yes")
                return
            items.RemoveAt(row)
            Refresh(row)
        }

        listView.OnEvent("DoubleClick", EditItem)
        Refresh()
        PreferencesWindow._y += height - 34
        PreferencesWindow._Add("Button", "w80", I18n.T("Prefs.Add")).OnEvent("Click", AddItem)
        PreferencesWindow._Add("Button", "x+8 yp w80", I18n.T("Prefs.Edit")).OnEvent("Click", EditItem)
        PreferencesWindow._Add("Button", "x+8 yp w80", I18n.T("Prefs.Delete")).OnEvent("Click", DeleteItem)
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
