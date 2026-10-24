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
    ; The actions panel: the default action first, then the general actions for this kind of result, then the result's own actions
    static ListFor(item) {
        list := []
        add(titleKey, icon, fn, hint := "") => list.Push(ResultItem(I18n.T(titleKey), hint, {Icon: icon, OnRun: fn}))
        target := item.Arg
        enter := "Enter"
        if (item.RunTitle != "" && IsObject(item.OnRun)) {                   ; the result's own default action (e.g. images in the clipboard history: paste)
            list.Push(ResultItem(item.RunTitle, "Enter", {Icon: item.Icon, OnRun: (*) => item.OnRun.Call(item)}))
            enter := ""
        }

        switch item.Kind {
            case "file":
                add("Action.Open", target, (*) => ActionCatalog.OpenFile(target, item.Arguments), enter)
                if RegExMatch(target, "i)\.(exe|lnk|bat|cmd|msc|ps1)$")
                    add("Action.RunAsAdmin", "res:imageres.dll,-78", (*) => ActionCatalog.RunAsAdmin(target, item.Arguments))
                if IsObject(editor := ActionCatalog.EditorFor(target))      ; the file type has an "Edit" verb (e.g. .ahk opens in an editor)
                    list.Push(ResultItem(I18n.T("Action.EditFile"), editor.Name (editor.Name != "" ? "  " : "") "Ctrl+E"
                        , {Icon: editor.Icon, OnRun: (*) => ActionCatalog.EditFile(target)}))
                if FileExist(target)
                    add("Action.OpenWith", "res:shell32.dll,-3", (*) => ActionCatalog.OpenWith(target))   ; the generic program icon
                add("Action.Reveal", "folder:", (*) => ActionCatalog.Reveal(target), "Ctrl+Enter")
                add("Action.CopyPath", "res:imageres.dll,-5314", (*) => ActionCatalog.CopyText(target), "Alt+Enter")
                add("Action.CopyName", "res:imageres.dll,-5314", (*) => ActionCatalog.CopyText(Path.Leaf(target)))
                ResultActions._FileActions(list, target)
                add("Action.OpenTerminal", "res:imageres.dll,-5323", (*) => TerminalProvider.OpenAt(ActionCatalog.ParentDir(target)))
                if !ActionCatalog.IsShellItem(target)                       ; Store apps (shell:AppsFolder\...) have no properties window
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
                if (FileExist(target) && !RegExMatch(RTrim(target, "\"), "^[A-Za-z]:$"))       ; drive roots can't be deleted
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

        if (item.Provider = ClipboardProvider.Id && item.Source is Map && ClipboardProvider.TypeOf(item.Source) = "text")   ; text in the clipboard history: search with each web search engine (built only when → opens the menu)
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

    ; Several marked files / folders (marked with Insert in the search window) together
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
    ; Edit / delete (handed to the EditItem / DeleteItem of the result's provider)
    ;---------------------------------------------------------------------------
    static CanEdit(item) {
        if ResultActions._ProviderHas(item, "EditItem")                     ; the result's provider decides (images in the clipboard history can't be edited)
            return ResultActions._ProviderCan(item, "EditItem")
        return item.Valid && item.Arg != "" && RegExMatch(item.Kind, "^(file|folder|url)$")
    }

    static CanDelete(item) {
        return ResultActions._ProviderCan(item, "DeleteItem")
    }

    ; true = the change was saved
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

    ; The confirmation text before deleting; a provider can explain what deleting really does with DeletePrompt(item)
    static DeletePrompt(item) {
        provider := ProviderRegistry.ById(item.Provider)
        if (IsObject(provider) && HasMethod(provider, "DeletePrompt"))
            return provider.DeletePrompt(item)
        return I18n.T("Search.ConfirmDelete", item.Title)
    }

    ; Application / file / folder / URL -> opens the edit dialog for a new custom command (a keyword etc. can be added)
    static AddAsCommand(item) {
        commandType := (item.Kind = "folder") ? "Folder" : (item.Kind = "url") ? "Url" : "File"
        return CustomCommandProvider.Edit("", Map("Title", item.Title, "Type", commandType, "Target", item.Arg, "Arguments", item.Arguments))
    }

    ; Only results that carry a Source (their entry in the settings) and whose provider implements method can be edited / deleted
    ; A provider can also let only some results be edited / deleted with CanEditItem(item) / CanDeleteItem(item)
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

    ; Copy / cut the file, copy / move to the file manager's current folder (only for real local or network paths)
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
        inside := (InStr(RTrim(destination, "\") "\", RTrim(target, "\") "\") = 1)   ; a folder can't be copied into itself
        if (StrLower(RTrim(destination, "\")) = StrLower(parent) || inside)
            return
        list.Push(ResultItem(I18n.T("Action.CopyTo", manager.Tag), destination, {Icon: "folder:", OnRun: (*) => ActionCatalog.CopyTo(target, destination)}))
        list.Push(ResultItem(I18n.T("Action.MoveTo", manager.Tag), destination, {Icon: "folder:", OnRun: (*) => ActionCatalog.CopyTo(target, destination, true)}))
    }
}
