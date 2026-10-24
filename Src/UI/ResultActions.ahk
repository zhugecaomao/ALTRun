;===============================================================================
; ResultActions.ahk - what the search window offers for a result: the actions panel, F3 edit, Ctrl+Del delete
;   (AutoHotkey v2)
;-------------------------------------------------------------------------------
;   →            the actions panel: every action available for the result (ListFor)
;   F3           edit the result (EditItem): custom commands / snippets / search engines are changed directly;
;                applications / files / folders / URLs get a new custom command (filled in)
;   Ctrl+Del     delete the result (DeleteItem): custom commands / snippets / search engines / clipboard history;
;                applications are hidden from the results
; The panel of a file or folder also has (as in Listary): copy / cut the file (to paste in TC or Explorer),
; Open with, copy / move to the folder open in TC (or Explorer), move to the Recycle Bin.
; The operations themselves are in ActionCatalog (Src\Core\ActionCatalog.ahk).
;
; Usage:
;   ResultActions.ListFor(item)           -> [ResultItem...] the actions panel
;   ResultActions.ListForMany(paths)      -> [ResultItem...] for several marked files
;   ResultActions.CanEdit(item) / EditItem(item) / CanDelete(item) / DeleteItem(item) / DeletePrompt(item)
;===============================================================================

class ResultActions {
    ; 操作面板: 第一项是默认操作, 然后是这一类结果的通用操作, 再加上结果自带的操作
    static ListFor(item) {
        list := []
        add(titleKey, icon, fn, hint := "") => list.Push(ResultItem(I18n.T(titleKey), hint, {Icon: icon, OnRun: fn}))
        target := item.Arg
        enter := "Enter"
        if (item.RunTitle != "" && IsObject(item.OnRun)) {                   ; 结果自己的默认操作 (例如剪贴板历史里的图片: 粘贴)
            list.Push(ResultItem(item.RunTitle, "Enter", {Icon: item.Icon, OnRun: (*) => item.OnRun.Call(item)}))
            enter := ""
        }

        switch item.Kind {
            case "file":
                add("Action.Open", target, (*) => ActionCatalog.OpenFile(target, item.Arguments), enter)
                if RegExMatch(target, "i)\.(exe|lnk|bat|cmd|msc|ps1)$")
                    add("Action.RunAsAdmin", "res:imageres.dll,-78", (*) => ActionCatalog.RunAsAdmin(target, item.Arguments))
                if IsObject(editor := ActionCatalog.EditorFor(target))      ; 这类文件登记了 "编辑" (例如 .ahk 用编辑器打开)
                    list.Push(ResultItem(I18n.T("Action.EditFile"), editor.Name (editor.Name != "" ? "  " : "") "Ctrl+E"
                        , {Icon: editor.Icon, OnRun: (*) => ActionCatalog.EditFile(target)}))
                if FileExist(target)
                    add("Action.OpenWith", "res:shell32.dll,-3", (*) => ActionCatalog.OpenWith(target))   ; 通用的程序图标
                add("Action.Reveal", "folder:", (*) => ActionCatalog.Reveal(target), "Ctrl+Enter")
                add("Action.CopyPath", "res:imageres.dll,-5314", (*) => ActionCatalog.CopyText(target), "Alt+Enter")
                add("Action.CopyName", "res:imageres.dll,-5314", (*) => ActionCatalog.CopyText(Path.Leaf(target)))
                ResultActions._FileActions(list, target)
                add("Action.OpenTerminal", "res:imageres.dll,-5323", (*) => TerminalProvider.OpenAt(ActionCatalog.ParentDir(target)))
                if !ActionCatalog.IsShellItem(target)                       ; 应用商店应用 (shell:AppsFolder\...) 没有属性窗口
                    add("Action.Properties", "res:imageres.dll,-81", (*) => ActionCatalog.ShowProperties(target))
                if FileExist(target)
                    add("Action.Recycle", "res:shell32.dll,-32", (*) => ActionCatalog.Recycle(target))
            case "folder":
                add("Action.Open", target, (*) => ActionCatalog.OpenFolder(target), enter)
                add("Action.OpenTerminal", "res:imageres.dll,-5323", (*) => TerminalProvider.OpenAt(target))
                add("Action.Reveal", "folder:", (*) => ActionCatalog.Reveal(target), "Ctrl+Enter")
                add("Action.CopyPath", "res:imageres.dll,-5314", (*) => ActionCatalog.CopyText(target), "Alt+Enter")
                ResultActions._FileActions(list, target)
                add("Action.Properties", "res:imageres.dll,-81", (*) => ActionCatalog.ShowProperties(target))
                if (FileExist(target) && !RegExMatch(RTrim(target, "\"), "^[A-Za-z]:$"))       ; 磁盘根目录不能删
                    add("Action.Recycle", "res:shell32.dll,-32", (*) => ActionCatalog.Recycle(target))
            case "url":
                add("Action.Open", "url:", (*) => ActionCatalog.OpenUrl(target), enter)
                add("Action.CopyUrl", "res:imageres.dll,-5314", (*) => ActionCatalog.CopyText(target), "Alt+Enter")
            case "text":
                add("Action.Copy", "res:imageres.dll,-5314", (*) => ActionCatalog.CopyText(target), enter)
                add("Action.Paste", "res:imageres.dll,-5314", (*) => ActionCatalog.PasteText(target), "Ctrl+Enter")
            default:
                add("Action.Run", item.Icon, (*) => ActionCatalog.RunDefault(item), enter)
        }

        for extra in item.Actions
            list.Push(extra)

        if (item.Provider = ClipboardProvider.Id && item.Source is Map && ClipboardProvider.TypeOf(item.Source) = "text")   ; 剪贴板历史里的文字: 用每个网页搜索引擎搜索 (打开 → 菜单时才生成)
            for extra in SelectionActions.SearchActions(item.Arg)
                list.Push(extra)

        if ResultActions._ProviderCan(item, "EditItem")
            add("Action.Edit", "res:imageres.dll,-5306", (*) => SearchWindow.EditItem(item), "F3")
        else if ResultActions.CanEdit(item)
            add("Action.AddCommand", "res:imageres.dll,-2", (*) => SearchWindow.EditItem(item), "F3")
        if ResultActions.CanDelete(item)
            add("Action.Delete", "res:shell32.dll,-240", (*) => SearchWindow.DeleteItem(item), "Ctrl+Del")
        if (item.Provider = RecentProvider.Id && item.HasOwnProp("Pinned") && item.Pinned)
            add("Action.Unpin", IconCache.Own("Pinned", "res:imageres.dll,-5303"), (*) => RecentProvider.Unpin(item.Uid))
        else if (RecentProvider.CanPin(item) && !RecentProvider.IsPinned(item))
            add("Action.Pin", IconCache.Own("Pinned", "res:imageres.dll,-5303"), (*) => RecentProvider.Pin(item))
        if LargeType.Enabled()
            add("Action.LargeType", "res:imageres.dll,-183", (*) => LargeType.Show(item.DisplayText()), "Ctrl+L")
        return list
    }

    ; 标记的多个文件 / 文件夹 (搜索窗口里 Insert 标记) 一起操作
    static ListForMany(paths) {
        list := []
        add(title, icon, fn, hint := "") => list.Push(ResultItem(title, hint, {Icon: icon, OnRun: fn}))
        count := paths.Length
        add(I18n.T("Action.OpenAll", count), "res:imageres.dll,-5302", (*) => ActionCatalog.OpenMany(paths), "Enter")
        add(I18n.T("Action.CopyFiles", count), "res:imageres.dll,-5314", (*) => ActionCatalog.CopyFiles(paths), I18n.T("Action.CopyFile.Hint"))
        add(I18n.T("Action.CutFiles", count), "res:imageres.dll,-5314", (*) => ActionCatalog.CopyFiles(paths, true), I18n.T("Action.CopyFile.Hint"))
        joined := TextTools.Join(paths)
        add(I18n.T("Action.CopyPaths", count), "res:imageres.dll,-5314", (*) => ActionCatalog.CopyText(joined))
        manager := QuickSwitch.FileManagerFolder()
        destination := IsObject(manager) ? manager.Path : ""
        if (destination != "" && DirExist(destination)) {
            add(I18n.T("Action.CopyTo", manager.Tag), "folder:", (*) => ActionCatalog.CopyMany(paths, destination), destination)
            add(I18n.T("Action.MoveTo", manager.Tag), "folder:", (*) => ActionCatalog.CopyMany(paths, destination, true), destination)
        }
        add(I18n.T("Action.AddAllCommands"), "res:imageres.dll,-2", (*) => CustomCommandProvider.AddFromPaths(paths))
        add(I18n.T("Action.RecycleAll", count), "res:shell32.dll,-32", (*) => ActionCatalog.RecycleMany(paths))
        return list
    }

    ;---------------------------------------------------------------------------
    ; 编辑 / 删除 (交给结果所属的 Provider 的 EditItem / DeleteItem)
    ;---------------------------------------------------------------------------
    static CanEdit(item) {
        if ResultActions._ProviderHas(item, "EditItem")                     ; 结果所属的 Provider 自己决定 (剪贴板历史里的图片不能编辑)
            return ResultActions._ProviderCan(item, "EditItem")
        return item.Valid && item.Arg != "" && RegExMatch(item.Kind, "^(file|folder|url)$")
    }

    static CanDelete(item) {
        return ResultActions._ProviderCan(item, "DeleteItem")
    }

    ; 返回 true = 已保存修改
    static EditItem(item) {
        if ResultActions._ProviderCan(item, "EditItem")
            return ProviderRegistry.ById(item.Provider).EditItem(item)
        if ResultActions.CanEdit(item)
            return ResultActions.AddAsCommand(item)
        return false
    }

    static DeleteItem(item) {
        if !ResultActions.CanDelete(item)
            return false
        return ProviderRegistry.ById(item.Provider).DeleteItem(item)
    }

    ; 删除前的确认文字; Provider 可以用 DeletePrompt(item) 说明删除的实际效果
    static DeletePrompt(item) {
        provider := ProviderRegistry.ById(item.Provider)
        if (IsObject(provider) && HasMethod(provider, "DeletePrompt"))
            return provider.DeletePrompt(item)
        return I18n.T("Search.ConfirmDelete", item.Title)
    }

    ; 应用 / 文件 / 文件夹 / 网址 -> 打开编辑对话框新建一条自定义命令 (可以再加关键字等)
    static AddAsCommand(item) {
        commandType := (item.Kind = "folder") ? "Folder" : (item.Kind = "url") ? "Url" : "File"
        return CustomCommandProvider.Edit("", Map("Title", item.Title, "Type", commandType, "Target", item.Arg, "Arguments", item.Arguments))
    }

    ; 结果带着 Source (设置里对应的那一条), 且所属 Provider 实现了 method 时才能编辑 / 删除
    ; Provider 还可以用 CanEditItem(item) / CanDeleteItem(item) 只让其中一部分结果编辑 / 删除
    static _ProviderCan(item, method) {
        if !ResultActions._ProviderHas(item, method)
            return false
        provider := ProviderRegistry.ById(item.Provider)
        return HasMethod(provider, "Can" method) ? (provider.%"Can" method%(item) ? true : false) : true
    }

    static _ProviderHas(item, method) {
        if (item.Provider = "" || !IsObject(item.Source))
            return false
        provider := ProviderRegistry.ById(item.Provider)
        return IsObject(provider) && HasMethod(provider, method)
    }

    ; 复制 / 剪切文件, 复制 / 移动到文件管理器当前的文件夹 (真实存在的本地或网络路径才有)
    static _FileActions(list, target) {
        if (!RegExMatch(target, "^([A-Za-z]:\\|\\\\)") || !FileExist(target))
            return
        list.Push(ResultItem(I18n.T("Action.CopyFile"), I18n.T("Action.CopyFile.Hint"), {Icon: "res:imageres.dll,-5314", OnRun: (*) => ActionCatalog.CopyFiles([target])}))
        list.Push(ResultItem(I18n.T("Action.CutFile"), I18n.T("Action.CopyFile.Hint"), {Icon: "res:imageres.dll,-5314", OnRun: (*) => ActionCatalog.CopyFiles([target], true)}))
        manager := QuickSwitch.FileManagerFolder()
        destination := IsObject(manager) ? manager.Path : ""
        if (destination = "" || !DirExist(destination))
            return
        parent := RTrim(ActionCatalog.ParentDir(RTrim(target, "\")), "\")
        inside := (InStr(RTrim(destination, "\") "\", RTrim(target, "\") "\") = 1)   ; 不能把文件夹复制到它自己里面
        if (StrLower(RTrim(destination, "\")) = StrLower(parent) || inside)
            return
        list.Push(ResultItem(I18n.T("Action.CopyTo", manager.Tag), destination, {Icon: "folder:", OnRun: (*) => ActionCatalog.CopyTo(target, destination)}))
        list.Push(ResultItem(I18n.T("Action.MoveTo", manager.Tag), destination, {Icon: "folder:", OnRun: (*) => ActionCatalog.CopyTo(target, destination, true)}))
    }
}
