;===============================================================================
; SelectionActions.ahk - 选中内容直接调出操作 (Alfred 的 Universal Actions) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 在任何程序里选中文字、文件或网址, 按热键 (General.SelectionHotkey, 默认 Ctrl+Alt+\),
; ALTRun 直接打开这些内容的操作面板:
;   文字   复制、用各个搜索引擎搜索、存为文字片段、大字显示、计算 (是算式时)、
;          替换为转换后的文字 (大写 / 小写 / 排序 / 简繁转换...), 替换会粘贴回原来的程序
;   网址   打开、复制
;   文件   打开、以管理员身份运行、在文件管理器中显示、复制路径、在此处打开终端、属性、添加到自定义命令
;   多个文件  复制路径、全部添加到自定义命令
; 取选中内容的方法和 Alfred 一样: 模拟复制 (Ctrl+C), 读完马上还原剪贴板, 也不记进剪贴板历史。
; 命令行 / 终端窗口里 Ctrl+C 会中断正在运行的程序, 这些窗口改用 Ctrl+Insert (复制)。
;
; 用法:
;   SelectionActions.Run()                    热键调用
;   SelectionActions.ItemFor({Text: "..."})   -> ResultItem (操作面板的来源), 也可以是 {Files: [...]}
;===============================================================================

class SelectionActions {
    static ConsoleClasses := "i)^(ConsoleWindowClass|CASCADIA_HOSTING_WINDOW_CLASS|mintty|PuTTY|VirtualConsoleClass)$"

    static Run() {
        selection := SelectionActions.Capture()
        if !IsObject(selection)
            return App.Notify(I18n.T("Selection.None"), 2000)
        SearchWindow.ShowActions(SelectionActions.ItemFor(selection))
    }

    ; 前台窗口里选中的内容: {Files: [路径...]} / {Text: "..."}; 什么都没选中时 ""
    static Capture() {
        className := ""
        try className := WinGetClass("A")
        Win.WaitModifiersUp()                                               ; 热键的 Ctrl+Alt 还按着时, Ctrl+C 会变成 Ctrl+Alt+C
        ClipboardProvider.PauseRecording(2000)
        saved := ClipboardAll()
        A_Clipboard := ""
        Send(RegExMatch(className, SelectionActions.ConsoleClasses) ? "^{Insert}" : "^c")
        result := ""
        if ClipWait(0.8, 1) {
            if DllCall("IsClipboardFormatAvailable", "UInt", 15) {         ; CF_HDROP: 复制的是文件
                files := []
                for line in StrSplit(A_Clipboard, "`n", "`r")
                    if (line != "")
                        files.Push(line)
                if files.Length
                    result := {Files: files}
            } else if (Trim(A_Clipboard, " `t`r`n") != "") {
                result := {Text: A_Clipboard}
            }
        }
        A_Clipboard := saved
        return result
    }

    static ItemFor(selection) {
        if selection.HasOwnProp("Files") {
            files := selection.Files
            if (files.Length = 1)
                return SelectionActions._FileItem(files[1])
            return SelectionActions._FilesItem(files)
        }
        text := selection.Text
        trimmed := Trim(text, " `t`r`n")
        if RegExMatch(trimmed, "i)^https?://\S+$")
            return ResultItem(trimmed, I18n.T("Selection.Link"), {Kind: "url", Arg: trimmed, Icon: "url:", Provider: "Selection"})
        if (!InStr(trimmed, "`n") && RegExMatch(trimmed, "^(?:[A-Za-z]:\\|\\\\)") && FileExist(trimmed))   ; 选中的是一个路径
            return SelectionActions._FileItem(trimmed)
        return SelectionActions._TextItem(text)
    }

    static _FileItem(filePath) {
        isFolder := InStr(FileExist(filePath), "D") ? true : false
        SplitPath(RTrim(filePath, "\"), &name)
        return ResultItem((name != "") ? name : filePath, filePath, {Kind: isFolder ? "folder" : "file", Arg: filePath
            , Icon: isFolder ? IconCache.FolderIcon(filePath) : filePath, Provider: "Selection"})
    }

    static _FilesItem(files) {
        names := ""
        for filePath in files {
            SplitPath(RTrim(filePath, "\"), &name)
            names .= (names = "" ? "" : ", ") name
            if (StrLen(names) > 100)
                break
        }
        joined := ""
        for filePath in files
            joined .= (joined = "" ? "" : "`r`n") filePath
        actions := [ResultItem(I18n.T("Action.AddAllCommands"), "", {Icon: "res:imageres.dll,-2", OnRun: (*) => CustomCommandProvider.AddFromPaths(files)})]
        return ResultItem(I18n.T("Selection.Files", files.Length), names, {Kind: "text", Arg: joined, Icon: "folder:", Provider: "Selection", Actions: actions})
    }

    static _TextItem(text) {
        preview := Trim(RegExReplace(text, "\s+", " "))
        if (StrLen(preview) > 80)
            preview := SubStr(preview, 1, 80) "..."
        actions := []
        ; 算式: 直接给出结果
        expression := StrReplace(Trim(text), ",")
        if (Calc.Looks(expression) && IsNumber(value := Calc.Eval(expression))) {
            result := CalculatorProvider.Format(value)
            actions.Push(ResultItem("= " result, I18n.T("Action.CopyResult"), {Icon: IconCache.Own("Calculator", "res:imageres.dll,-182"), OnRun: (*) => ActionCatalog.CopyText(result)}))
        }
        ; 用每个搜索引擎搜索
        query := Trim(RegExReplace(text, "\s+", " "))
        if ProviderRegistry.IsEnabled(WebSearchProvider) {
            for engine in AppSettings.Feature("WebSearch")["Engines"] {
                searchUrl := StrReplace(engine["Url"], "{query}", Url.Encode(query))
                actions.Push(ResultItem(I18n.T("Action.SearchWith", engine["Title"]), "", {Icon: "url:", OnRun: SelectionActions._Opener(searchUrl)}))
            }
        }
        actions.Push(ResultItem(I18n.T("Action.SaveSnippet"), "", {Icon: IconCache.Own("Snippet", "res:imageres.dll,-102")
            , OnRun: (*) => SnippetProvider.Edit("", Map("Name", SubStr(preview, 1, 40), "Text", text))}))
        ; 替换为转换后的文字 (粘贴回原来的程序, 选中的文字还在那里)
        for transform in SelectionActions.Transforms() {
            name := RegExReplace(I18n.T(transform[1]), "^[^:：]+[:：]\s*")    ; 去掉 "剪贴板: " 前缀
            actions.Push(ResultItem(I18n.T("Action.ReplaceWith", name), "", {Icon: "res:imageres.dll,-5314", OnRun: SelectionActions._Replacer(text, transform[2])}))
        }
        return ResultItem(preview, I18n.T("Selection.Text", StrLen(text)), {Kind: "text", Arg: text, Icon: "res:imageres.dll,-5314"
            , LargeText: text, Provider: "Selection", Actions: actions})
    }

    ; [I18n 键, 转换函数]; 和系统命令里的剪贴板文字转换是同一套
    static Transforms() {
        return [["Text.Upper", (s) => TextTools.Upper(s)], ["Text.Lower", (s) => TextTools.Lower(s)], ["Text.Title", (s) => TextTools.TitleCase(s)]
              , ["Text.SortAsc", (s) => TextTools.SortLines(s)], ["Text.SortDesc", (s) => TextTools.SortLines(s, true)]
              , ["Text.TrimLines", (s) => TextTools.TrimLines(s)], ["Text.RemoveBlank", (s) => TextTools.RemoveBlankLines(s)]
              , ["Text.Dedupe", (s) => TextTools.DedupeLines(s)], ["Text.ToTraditional", (s) => Kanji.ToTraditional(s)]
              , ["Text.ToSimplified", (s) => Kanji.ToSimplified(s)], ["Text.UrlEncode", (s) => Url.Encode(s)]]
    }

    ; 每个动作各自记住自己的参数 (单独的方法生成闭包)
    static _Opener(target) {
        return (*) => ActionCatalog.OpenUrl(target)
    }

    static _Replacer(text, fn) {
        return (*) => ActionCatalog.PasteText(fn(text))
    }
}
