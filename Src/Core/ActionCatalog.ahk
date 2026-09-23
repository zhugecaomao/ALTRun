;===============================================================================
; ActionCatalog.ahk - 结果的操作: 打开 / 显示位置 / 复制 / 粘贴 ... (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 按 Alfred 的习惯:
;   Enter        默认操作  (文件/程序: 打开, 文件夹: 打开, 网址: 打开, 文字: 复制)
;   Ctrl+Enter   文件/文件夹: 在文件管理器中显示;  文字: 粘贴到前台窗口
;   Alt+Enter    复制 (路径 / 网址 / 文字)
;   →            打开操作面板, 列出这一项全部可用的操作 (ListFor)
;
; 用法:
;   ActionCatalog.RunDefault(item)
;   ActionCatalog.RunModifier(item, "ctrl" | "alt")
;   ActionCatalog.ListFor(item)           -> [ResultItem...] 操作面板的内容
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

        if ((item.Kind = "file" || item.Kind = "folder") && item.Provider != "CustomCommands")
            add("Action.AddCommand", "res:imageres.dll,-2", (*) => CustomCommandProvider.AddFromPath(target, item.Title))
        add("Action.LargeType", "res:imageres.dll,-183", (*) => LargeType.Show(item.DisplayText()), "Ctrl+L")
        return list
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

    ; 把文字粘贴到呼出 ALTRun 之前的那个窗口。
    ; mode: "Clipboard" = 临时借用剪贴板 + Ctrl+V (之后还原); "Type" = 逐字输入
    static PasteText(text, mode := "", delay := 0) {
        snippetSettings := AppSettings.Feature("Snippets")
        mode  := (mode != "") ? mode : snippetSettings["PasteMode"]
        delay := delay ? delay : snippetSettings["PasteDelay"]

        App.FocusPreviousWindow()
        Win.WaitModifiersUp()
        if (mode = "Type") {
            SendInput("{Text}" text)
            return true
        }
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
