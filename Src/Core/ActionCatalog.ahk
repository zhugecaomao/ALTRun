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

        switch item.Kind {
            case "file":
                add("Action.Open", target, (*) => ActionCatalog.OpenFile(target, item.Arguments), "Enter")
                if RegExMatch(target, "i)\.(exe|lnk|bat|cmd|msc|ps1)$")
                    add("Action.RunAsAdmin", "res:imageres.dll,-78", (*) => ActionCatalog.RunAsAdmin(target, item.Arguments))
                add("Action.Reveal", "folder:", (*) => ActionCatalog.Reveal(target), "Ctrl+Enter")
                add("Action.CopyPath", "res:imageres.dll,-5314", (*) => ActionCatalog.CopyText(target), "Alt+Enter")
                add("Action.CopyName", "res:imageres.dll,-5314", (*) => ActionCatalog.CopyText(Path.Leaf(target)))
                add("Action.OpenTerminal", "res:imageres.dll,-5323", (*) => TerminalProvider.OpenAt(ActionCatalog._ParentDir(target)))
                add("Action.Properties", "res:imageres.dll,-81", (*) => ActionCatalog.ShowProperties(target))
            case "folder":
                add("Action.Open", target, (*) => ActionCatalog.OpenFolder(target), "Enter")
                add("Action.OpenTerminal", "res:imageres.dll,-5323", (*) => TerminalProvider.OpenAt(target))
                add("Action.Reveal", "folder:", (*) => ActionCatalog.Reveal(target), "Ctrl+Enter")
                add("Action.CopyPath", "res:imageres.dll,-5314", (*) => ActionCatalog.CopyText(target), "Alt+Enter")
                add("Action.Properties", "res:imageres.dll,-81", (*) => ActionCatalog.ShowProperties(target))
            case "url":
                add("Action.Open", "url:", (*) => ActionCatalog.OpenUrl(target), "Enter")
                add("Action.CopyUrl", "res:imageres.dll,-5314", (*) => ActionCatalog.CopyText(target), "Alt+Enter")
            case "text":
                add("Action.Copy", "res:imageres.dll,-5314", (*) => ActionCatalog.CopyText(target), "Enter")
                add("Action.Paste", "res:imageres.dll,-5314", (*) => ActionCatalog.PasteText(target), "Ctrl+Enter")
            default:
                add("Action.Run", item.Icon, (*) => ActionCatalog.RunDefault(item), "Enter")
        }

        for extra in item.Actions
            list.Push(extra)

        if ActionCatalog._ProviderCan(item, "EditItem")
            add("Action.Edit", "res:imageres.dll,-5306", (*) => SearchWindow.EditItem(item), "F3")
        else if ActionCatalog.CanEdit(item)
            add("Action.AddCommand", "res:imageres.dll,-2", (*) => SearchWindow.EditItem(item), "F3")
        if ActionCatalog.CanDelete(item)
            add("Action.Delete", "res:shell32.dll,-240", (*) => SearchWindow.DeleteItem(item), "Ctrl+Del")
        add("Action.LargeType", "res:imageres.dll,-183", (*) => LargeType.Show(item.DisplayText()), "Ctrl+L")
        return list
    }

    ;---------------------------------------------------------------------------
    ; 编辑 / 删除 (交给结果所属的 Provider 的 EditItem / DeleteItem)
    ;---------------------------------------------------------------------------
    static CanEdit(item) {
        return ActionCatalog._ProviderCan(item, "EditItem") || (item.Valid && item.Arg != "" && RegExMatch(item.Kind, "^(file|folder|url)$"))
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
    static _ProviderCan(item, method) {
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
        Run(fileManager ' /P "' resolved '"')                               ; Total Commander 等: /P 打开所在文件夹
    }

    static OpenUrl(address) {
        if !RegExMatch(address, "i)^[a-z][a-z0-9+.\-]*:")
            address := "https://" address
        Run(address)
    }

    static ShowProperties(target) {
        Run('properties "' Path.Resolve(target) '"')
    }

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

    static _ParentDir(target) {
        SplitPath(target, , &dir)
        return dir
    }

    static _Preview(text) {
        text := RegExReplace(text, "\s+", " ")
        return (StrLen(text) > 60) ? SubStr(text, 1, 60) "..." : text
    }
}
