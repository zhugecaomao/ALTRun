;===============================================================================
; ActionCatalog.ahk - 结果的操作: 打开 / 显示位置 / 复制 / 粘贴 ... (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 按 Alfred 的习惯:
;   Enter        默认操作  (文件/程序: 打开, 文件夹: 打开, 网址: 打开, 文字: 复制)
;   Ctrl+Enter   文件/文件夹: 在文件管理器中显示;  文字: 粘贴到前台窗口
;   Alt+Enter    复制 (路径 / 网址 / 文字)
;   →            打开操作面板, 列出这一项全部可用的操作 (ListFor)
;   F3           编辑这一项 (EditItem): 自定义命令 / 片段 / 搜索引擎 直接修改;
;                应用 / 文件 / 文件夹 / 网址 新建一条自定义命令 (预先填好)
;   Ctrl+Del     删除这一项 (DeleteItem): 自定义命令 / 片段 / 搜索引擎 / 剪贴板历史;
;                应用: 从搜索结果中隐藏
; 文件 / 文件夹的操作面板里还有 (和 Listary 一样): 复制 / 剪切文件 (到 TC、资源管理器里粘贴)、打开方式、
; 复制 / 移动到 TC (或资源管理器) 当前打开的文件夹、移到回收站
;
; 用法:
;   ActionCatalog.RunDefault(item)
;   ActionCatalog.RunModifier(item, "ctrl" | "alt")
;   ActionCatalog.ListFor(item)           -> [ResultItem...] 操作面板的内容
;   ActionCatalog.CanEdit(item) / EditItem(item) / CanDelete(item) / DeleteItem(item)
;   ActionCatalog.OpenFile(path) / OpenFolder(path) / Reveal(path) / CopyText(text) / PasteText(text) ...
;===============================================================================

class ActionCatalog {

    static RunDefault(item) {
        if IsObject(item.OnRun)
            return item.OnRun.Call(item)
        switch item.Kind {
            case "file"  : ActionCatalog.OpenFile(item.Arg, item.Arguments)
            case "folder": ActionCatalog.OpenFolder(item.Arg)
            case "url"   : ActionCatalog.OpenUrl(item.Arg)
            case "text"  : ActionCatalog.CopyText(item.Arg)
        }
    }

    static RunModifier(item, modifier) {
        if (modifier = "alt")
            return ActionCatalog.CopyText(item.CopyText())
        if (modifier = "ctrl") {
            switch item.Kind {
                case "file", "folder": return ActionCatalog.Reveal(item.Arg)
                case "text"          : return ActionCatalog.PasteText(item.Arg)
                case "url"           : return ActionCatalog.CopyText(item.Arg)
            }
        }
        ActionCatalog.RunDefault(item)
    }

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
                if FileExist(target)
                    add("Action.OpenWith", "res:shell32.dll,-16702", (*) => ActionCatalog.OpenWith(target))
                add("Action.Reveal", "folder:", (*) => ActionCatalog.Reveal(target), "Ctrl+Enter")
                add("Action.CopyPath", "res:imageres.dll,-5314", (*) => ActionCatalog.CopyText(target), "Alt+Enter")
                add("Action.CopyName", "res:imageres.dll,-5314", (*) => ActionCatalog.CopyText(Path.Leaf(target)))
                ActionCatalog._FileActions(list, target)
                add("Action.OpenTerminal", "res:imageres.dll,-5323", (*) => TerminalProvider.OpenAt(ActionCatalog._ParentDir(target)))
                if !ActionCatalog.IsShellItem(target)                       ; 应用商店应用 (shell:AppsFolder\...) 没有属性窗口
                    add("Action.Properties", "res:imageres.dll,-81", (*) => ActionCatalog.ShowProperties(target))
                if FileExist(target)
                    add("Action.Recycle", "res:shell32.dll,-32", (*) => ActionCatalog.Recycle(target))
            case "folder":
                add("Action.Open", target, (*) => ActionCatalog.OpenFolder(target), enter)
                add("Action.OpenTerminal", "res:imageres.dll,-5323", (*) => TerminalProvider.OpenAt(target))
                add("Action.Reveal", "folder:", (*) => ActionCatalog.Reveal(target), "Ctrl+Enter")
                add("Action.CopyPath", "res:imageres.dll,-5314", (*) => ActionCatalog.CopyText(target), "Alt+Enter")
                ActionCatalog._FileActions(list, target)
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

        if ActionCatalog._ProviderCan(item, "EditItem")
            add("Action.Edit", "res:imageres.dll,-5306", (*) => SearchWindow.EditItem(item), "F3")
        else if ActionCatalog.CanEdit(item)
            add("Action.AddCommand", "res:imageres.dll,-2", (*) => SearchWindow.EditItem(item), "F3")
        if ActionCatalog.CanDelete(item)
            add("Action.Delete", "res:shell32.dll,-240", (*) => SearchWindow.DeleteItem(item), "Ctrl+Del")
        if (item.Provider = RecentProvider.Id && item.HasOwnProp("Pinned") && item.Pinned)
            add("Action.Unpin", IconCache.Own("Pinned", "res:imageres.dll,-5303"), (*) => RecentProvider.Unpin(item.Uid))
        else if (RecentProvider.CanPin(item) && !RecentProvider.IsPinned(item))
            add("Action.Pin", IconCache.Own("Pinned", "res:imageres.dll,-5303"), (*) => RecentProvider.Pin(item))
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
        joined := ""
        for markedPath in paths
            joined .= (joined = "" ? "" : "`r`n") markedPath
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

    static OpenMany(paths) {
        for markedPath in paths
            InStr(FileExist(markedPath), "D") ? ActionCatalog.OpenFolder(markedPath) : ActionCatalog.OpenFile(markedPath)
    }

    ; 复制 / 移动到一个文件夹 (Windows 自己的复制, 有进度和重名提示); 不能复制到自己或自己的子文件夹里
    static CopyMany(paths, destination, move := false) {
        folder := ComObject("Shell.Application").NameSpace(RTrim(destination, "\") (RegExMatch(destination, "^[A-Za-z]:\\?$") ? "\" : ""))
        if !IsObject(folder)
            return false
        done := 0
        for markedPath in paths {
            if (InStr(RTrim(destination, "\") "\", RTrim(markedPath, "\") "\") = 1)
                continue
            move ? folder.MoveHere(markedPath, 0) : folder.CopyHere(markedPath, 0)
            done += 1
        }
        App.Notify(I18n.T(move ? "Action.MovingMany" : "Action.CopyingMany", done, destination), 2000)
        return done > 0
    }

    ; 一起移到回收站: 先确认一次 (可以从回收站还原)
    static RecycleMany(paths) {
        if (MsgBox(I18n.T("Action.ConfirmRecycleAll", paths.Length), App.Name, "YesNo Icon? Default2") != "Yes")
            return false
        failed := 0
        for markedPath in paths
            try FileRecycle(markedPath)
            catch
                failed += 1
        if failed
            App.Notify(I18n.T("Action.RecycleFailed", failed))
        return !failed
    }

    ;---------------------------------------------------------------------------
    ; 编辑 / 删除 (交给结果所属的 Provider 的 EditItem / DeleteItem)
    ;---------------------------------------------------------------------------
    static CanEdit(item) {
        if ActionCatalog._ProviderHas(item, "EditItem")                     ; 结果所属的 Provider 自己决定 (剪贴板历史里的图片不能编辑)
            return ActionCatalog._ProviderCan(item, "EditItem")
        return item.Valid && item.Arg != "" && RegExMatch(item.Kind, "^(file|folder|url)$")
    }

    static CanDelete(item) {
        return ActionCatalog._ProviderCan(item, "DeleteItem")
    }

    ; 返回 true = 已保存修改
    static EditItem(item) {
        if ActionCatalog._ProviderCan(item, "EditItem")
            return ProviderRegistry.ById(item.Provider).EditItem(item)
        if ActionCatalog.CanEdit(item)
            return ActionCatalog.AddAsCommand(item)
        return false
    }

    static DeleteItem(item) {
        if !ActionCatalog.CanDelete(item)
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
        if !ActionCatalog._ProviderHas(item, method)
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

    ;---------------------------------------------------------------------------
    ; 具体操作
    ;---------------------------------------------------------------------------
    static OpenFile(target, arguments := "") {
        resolved := Path.Resolve(target)
        workDir := ActionCatalog._ParentDir(resolved)
        if RegExMatch(resolved, "i)^(shell:|::\{|[a-z]+:\/\/)") || !FileExist(resolved)
            return Run(Trim(resolved " " arguments))                        ; shell: 路径 / 协议 / PATH 里的程序名
        Run(Trim('"' resolved '" ' arguments), DirExist(workDir) ? workDir : "")
    }

    static RunAsAdmin(target, arguments := "") {
        Run(Trim('*RunAs "' Path.Resolve(target) '" ' arguments))
    }

    static OpenFolder(target) {
        resolved := Path.Resolve(target)
        fileManager := AppSettings.General["FileManager"]
        if (fileManager = "" || InStr(fileManager, "explorer"))
            return Run('explorer.exe "' resolved '"')
        Run(fileManager ' "' resolved '"')
    }

    static Reveal(target) {
        resolved := Path.Resolve(target)
        fileManager := AppSettings.General["FileManager"]
        if (fileManager = "" || InStr(fileManager, "explorer"))
            return Run('explorer.exe /select,"' resolved '"')
        if InStr(fileManager, "totalcmd")
            return Run(fileManager ' /P "' resolved '"')                   ; Total Commander: 打开所在文件夹并选中这个文件
        Run(fileManager ' "' ActionCatalog._ParentDir(resolved) '"')        ; 其它文件管理器的参数各不相同: 打开所在文件夹
    }

    static OpenUrl(address) {
        if !RegExMatch(address, "i)^[a-z][a-z0-9+.\-]*:")
            address := "https://" address
        Run(address)
    }

    ; 和资源管理器里 "右键 -> 属性" 一样。Run 的 properties 动词对网络盘上的文件夹等会失败, 所以用 SHObjectProperties
    static ShowProperties(target) {
        target := Path.Resolve(target)
        if !DllCall("shell32\SHObjectProperties", "Ptr", 0, "UInt", 0x2, "WStr", target, "Ptr", 0)   ; 0x2 = SHOP_FILEPATH
            Run('properties "' target '"')
    }

    static IsShellItem(target) => RegExMatch(target, "i)^shell:") ? true : false

    static CopyText(text) {
        A_Clipboard := text
        App.Notify(I18n.T("Search.Copied", ActionCatalog._Preview(text)))
    }

    ; 把文字粘贴到呼出 ALTRun 之前的那个窗口 (focusPrevious = false 时粘贴到当前窗口)。
    ; 粘贴方式见 Features.Snippets.PasteMode: "Clipboard" = 临时借用剪贴板 + Ctrl+V
    ; (之后还原, 这段时间剪贴板历史不记录); "Type" = 逐字输入
    static PasteText(text, focusPrevious := true) {
        snippetSettings := AppSettings.Feature("Snippets")
        mode  := snippetSettings["PasteMode"]
        delay := snippetSettings["PasteDelay"]

        if focusPrevious
            App.FocusPreviousWindow()
        Win.WaitModifiersUp()
        if (mode = "Type") {
            SendInput("{Text}" text)
            return true
        }
        ClipboardProvider.PauseRecording(delay + 1000)
        savedClipboard := ClipboardAll()
        A_Clipboard := ""
        A_Clipboard := text
        if !ClipWait(1) {
            A_Clipboard := savedClipboard
            return false
        }
        SendInput("^v")
        Sleep(delay)
        A_Clipboard := savedClipboard
        return true
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
        parent := RTrim(ActionCatalog._ParentDir(RTrim(target, "\")), "\")
        inside := (InStr(RTrim(destination, "\") "\", RTrim(target, "\") "\") = 1)   ; 不能把文件夹复制到它自己里面
        if (StrLower(RTrim(destination, "\")) = StrLower(parent) || inside)
            return
        list.Push(ResultItem(I18n.T("Action.CopyTo", manager.Tag), destination, {Icon: "folder:", OnRun: (*) => ActionCatalog.CopyTo(target, destination)}))
        list.Push(ResultItem(I18n.T("Action.MoveTo", manager.Tag), destination, {Icon: "folder:", OnRun: (*) => ActionCatalog.CopyTo(target, destination, true)}))
    }

    ; 把文件放进剪贴板 (剪贴板历史照常记录), 在 TC / 资源管理器里 Ctrl+V 粘贴
    static CopyFiles(paths, cut := false) {
        if !ClipboardData.SetFiles(paths, cut)
            return false
        App.Notify(I18n.T(cut ? "Action.FileCut" : "Action.FileCopied", Path.Leaf(RTrim(paths[1], "\"))))
        return true
    }

    static OpenWith(target) {
        Run("rundll32.exe shell32.dll,OpenAs_RunDLL " Path.Resolve(target))    ; 路径不加引号 (OpenAs_RunDLL 取整个剩余的命令行)
    }

    ; 用 Windows 自己的复制 / 移动 (有进度窗口、重名时询问、可以撤销), 在后台进行, 不会卡住 ALTRun
    static CopyTo(target, destination, move := false) {
        folder := ComObject("Shell.Application").NameSpace(RTrim(destination, "\") (RegExMatch(destination, "^[A-Za-z]:\\?$") ? "\" : ""))
        if !IsObject(folder)
            return false
        move ? folder.MoveHere(target, 0) : folder.CopyHere(target, 0)
        App.Notify(I18n.T(move ? "Action.Moving" : "Action.Copying", Path.Leaf(RTrim(target, "\")), destination), 2000)
        return true
    }

    ; 移到回收站: 用 Windows 的 "删除" (会先确认, 可以从回收站还原)
    static Recycle(target) {
        SplitPath(RTrim(target, "\"), &name, &dir)
        try {
            folderItem := ComObject("Shell.Application").NameSpace(RegExMatch(dir, "^[A-Za-z]:$") ? dir "\" : dir).ParseName(name)
            folderItem.InvokeVerb("delete")
            return true
        }
        return false
    }

    static _ParentDir(target) {
        SplitPath(target, , &dir)
        return dir
    }

    static _Preview(text) {
        text := RegExReplace(text, "\s+", " ")
        return (StrLen(text) > 60) ? SubStr(text, 1, 60) "..." : text
    }
}
