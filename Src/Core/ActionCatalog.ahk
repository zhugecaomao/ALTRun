;===============================================================================
; ActionCatalog.ahk - what running a result does, and the operations behind it: open / reveal / copy / paste ...
;   (AutoHotkey v2)
;-------------------------------------------------------------------------------
; As in Alfred:
;   Enter        the default action (file / program / folder / URL: open; text: copy)
;   Ctrl+Enter   file / folder: reveal in the file manager; text: paste into the foreground window
;   Alt+Enter    copy (path / URL / text)
; The actions panel (→), F3 edit and Ctrl+Del delete are built on these in Src\UI\ResultActions.ahk.
;
; Usage:
;   ActionCatalog.RunDefault(item)
;   ActionCatalog.RunModifier(item, "ctrl" | "alt")
;   ActionCatalog.OpenFile(path) / OpenFolder(path) / Reveal(path) / CopyText(text) / PasteText(text) ...
;   ActionCatalog.OpenMany(paths) / CopyMany(paths, folder) / RecycleMany(paths)   several marked files
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
    ; 具体操作
    ;---------------------------------------------------------------------------
    static OpenFile(target, arguments := "") {
        resolved := Path.Resolve(target)
        workDir := ActionCatalog.ParentDir(resolved)
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
        Run(fileManager ' "' ActionCatalog.ParentDir(resolved) '"')        ; 其它文件管理器的参数各不相同: 打开所在文件夹
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
        if ActionCatalog.TypesText(text, mode) {
            SendInput("{Text}" text)
            return true
        }
        ClipboardData.Borrow(delay + 1000)
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

    ; 粘贴方式 (Features.Snippets.PasteMode) -> 这段文字是不是直接输入:
    ;   Auto (默认)  不长 (AutoTypeMax 个字符以内) 的单行文字直接输入: 马上出来, 不动剪贴板, 也不用等着还原剪贴板;
    ;                多行 (直接输入时换行是 Enter, 可能把聊天消息发出去)、带 Tab、很长、有 emoji 等补充平面字符的用剪贴板
    ;   Clipboard    总是剪贴板 + Ctrl+V
    ;   Type         总是逐字输入 (适合不接受粘贴的程序)
    static AutoTypeMax := 100

    static TypesText(text, mode) {
        if (mode = "Type")
            return true
        if (mode != "Auto")
            return false
        return StrLen(text) <= ActionCatalog.AutoTypeMax && !RegExMatch(text, "[\r\n\t]|[^\x{0}-\x{FFFF}]")
    }

    ; 把文件放进剪贴板 (剪贴板历史照常记录), 在 TC / 资源管理器里 Ctrl+V 粘贴
    static CopyFiles(paths, cut := false) {
        if !ClipboardData.SetFiles(paths, cut)
            return false
        App.Notify(I18n.T(cut ? "Action.FileCut" : "Action.FileCopied", Path.Leaf(RTrim(paths[1], "\"))))
        return true
    }

    ; 这类文件在 Windows 里登记的 "编辑" 程序 (右键菜单里的 "编辑") -> {Name, Icon}; 没有登记时 ""
    static EditorFor(target) {
        static ASSOCSTR_COMMAND := 1, ASSOCSTR_EXECUTABLE := 2, ASSOCSTR_FRIENDLYAPPNAME := 4
        SplitPath(target, , , &ext)
        attributes := FileExist(target)
        if (ext = "" || attributes = "" || InStr(attributes, "D"))
            return ""
        if (ActionCatalog.AssocString(ASSOCSTR_COMMAND, "." ext, "edit") = "")
            return ""
        exe := ActionCatalog.AssocString(ASSOCSTR_EXECUTABLE, "." ext, "edit")
        return {Name: ActionCatalog.AssocString(ASSOCSTR_FRIENDLYAPPNAME, "." ext, "edit")
              , Icon: (exe != "" && FileExist(exe)) ? exe : A_WinDir "\notepad.exe"}
    }

    ; AssocQueryStringW: 文件类型 (".ahk") 的某个动作 ("edit") 的命令、程序、程序名称; 没有时 ""
    static AssocString(kind, ext, verb) {
        static ASSOCF_NOTRUNCATE := 0x20, ASSOCF_INIT_IGNOREUNKNOWN := 0x400   ; 没登记的类型不要退到 "打开方式" 的设置
        size := 0
        DllCall("shlwapi\AssocQueryStringW", "UInt", ASSOCF_NOTRUNCATE | ASSOCF_INIT_IGNOREUNKNOWN, "UInt", kind, "WStr", ext, "WStr", verb
            , "Ptr", 0, "UInt*", &size)
        if (size <= 1)
            return ""
        buf := Buffer(size * 2, 0)
        if (DllCall("shlwapi\AssocQueryStringW", "UInt", ASSOCF_NOTRUNCATE | ASSOCF_INIT_IGNOREUNKNOWN, "UInt", kind, "WStr", ext, "WStr", verb
            , "Ptr", buf, "UInt*", &size) != 0)
            return ""
        return StrGet(buf)
    }

    ; 用登记的编辑程序打开 (和右键菜单的 "编辑" 一样)
    static EditFile(target) {
        Run('edit "' Path.Resolve(target) '"')
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

    static ParentDir(target) {
        SplitPath(target, , &dir)
        return dir
    }

    static _Preview(text) {
        text := RegExReplace(text, "\s+", " ")
        return (StrLen(text) > 60) ? SubStr(text, 1, 60) "..." : text
    }
}
