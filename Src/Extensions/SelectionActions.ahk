;===============================================================================
; SelectionActions.ahk - actions for the selection directly (Alfred's Universal Actions) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Select text, files or a URL in any program and press the hotkey (General.SelectionHotkey, default Ctrl+Alt+\):
; ALTRun opens the actions panel for them directly:
;   text       copy, search with each search engine, save as a snippet, Large Type, calculate (for an expression),
;              replace with converted text (upper / lower case / sort / Simplified-Traditional...); replacing pastes back into the program
;   URL        open, copy
;   file       open, run as administrator, reveal in the file manager, copy path, open terminal here, properties, add as custom command
;   files      copy paths, add all as custom commands
; The selection is read as in Alfred: simulate copying (Ctrl+C), restore the clipboard right after reading, and keep it out of the clipboard history.
; In command-line / terminal windows Ctrl+C would interrupt the running program, so those use Ctrl+Insert (copy).
;
; Usage:
;   SelectionActions.Run()                    called by the hotkey
;   SelectionActions.ItemFor({Text: "..."})   -> ResultItem (the source of the actions panel); can also be {Files: [...]}
;===============================================================================

class SelectionActions {
    static ConsoleClasses := "i)^(ConsoleWindowClass|CASCADIA_HOSTING_WINDOW_CLASS|mintty|PuTTY|VirtualConsoleClass)$"

    static Run() {
        selection := SelectionActions.Capture()
        if !IsObject(selection)
            return App.Notify(I18n.T("Selection.None"), 2000)
        SearchWindow.ShowActions(SelectionActions.ItemFor(selection))
    }

    ; The selection in the foreground window: {Files: [paths...]} / {Text: "..."}; "" when nothing is selected
    static Capture() {
        className := ""
        try className := WinGetClass("A")
        Win.WaitModifiersUp()                                               ; with the hotkey's Ctrl+Alt still held, Ctrl+C would become Ctrl+Alt+C
        ClipboardData.Borrow(2000)
        saved := ClipboardAll()
        A_Clipboard := ""
        Send(RegExMatch(className, SelectionActions.ConsoleClasses) ? "^{Insert}" : "^c")
        result := ""
        if ClipWait(0.8, 1) {
            if DllCall("IsClipboardFormatAvailable", "UInt", 15) {         ; CF_HDROP: files were copied
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
        if (!InStr(trimmed, "`n") && RegExMatch(trimmed, "^(?:[A-Za-z]:\\|\\\\)") && FileExist(trimmed))   ; the selection is one path
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
        ; an expression: give the result directly
        expression := StrReplace(Trim(text), ",")
        if (Calc.Looks(expression) && IsNumber(value := Calc.Eval(expression))) {
            result := CalculatorProvider.Format(value)
            actions.Push(ResultItem("= " result, I18n.T("Action.CopyResult"), {Icon: IconCache.Own("Calculator", "res:imageres.dll,-182"), OnRun: (*) => ActionCatalog.CopyText(result)}))
        }
        for action in SelectionActions.SearchActions(text)
            actions.Push(action)
        actions.Push(ResultItem(I18n.T("Action.SaveSnippet"), "", {Icon: IconCache.Own("Snippet", "res:imageres.dll,-102")
            , OnRun: (*) => SnippetProvider.Edit("", Map("Name", SubStr(preview, 1, 40), "Text", text))}))
        ; replace with converted text (pasted back into the program, where the text is still selected)
        for transform in SelectionActions.Transforms() {
            name := RegExReplace(I18n.T(transform[1]), "^[^:：]+[:：]\s*")    ; drop the "Clipboard: " prefix
            actions.Push(ResultItem(I18n.T("Action.ReplaceWith", name), "", {Icon: "res:imageres.dll,-5314", OnRun: SelectionActions._Replacer(text, transform[2])}))
        }
        return ResultItem(preview, I18n.T("Selection.Text", StrLen(text)), {Kind: "text", Arg: text, Icon: "res:imageres.dll,-5314"
            , LargeText: text, Provider: "Selection", Actions: actions})
    }

    ; Actions that search text with each search engine (shared by the selection and the clipboard history); none with web search off
    static SearchActions(text) {
        actions := []
        query := WebSearchProvider.Term(text)
        if (query = "" || !ProviderRegistry.IsEnabled(WebSearchProvider))
            return actions
        for engine in AppSettings.Feature("WebSearch")["Engines"] {
            searchUrl := WebSearchProvider.Fill(engine["Url"], query)
            actions.Push(ResultItem(I18n.T("Action.SearchWith", engine["Title"]), "", {Icon: WebSearchProvider.IconFor(engine), OnRun: SelectionActions._Opener(searchUrl)}))
        }
        return actions
    }

    ; [I18n key, conversion function]; the same set as the clipboard text conversions in the system commands
    static Transforms() {
        return [["Text.Upper", (s) => TextTools.Upper(s)], ["Text.Lower", (s) => TextTools.Lower(s)], ["Text.Title", (s) => TextTools.TitleCase(s)]
              , ["Text.SortAsc", (s) => TextTools.SortLines(s)], ["Text.SortDesc", (s) => TextTools.SortLines(s, true)]
              , ["Text.TrimLines", (s) => TextTools.TrimLines(s)], ["Text.RemoveBlank", (s) => TextTools.RemoveBlankLines(s)]
              , ["Text.Dedupe", (s) => TextTools.DedupeLines(s)], ["Text.ToTraditional", (s) => Kanji.ToTraditional(s)]
              , ["Text.ToSimplified", (s) => Kanji.ToSimplified(s)], ["Text.UrlEncode", (s) => Url.Encode(s)]]
    }

    ; each action keeps its own arguments (a separate method makes the closure)
    static _Opener(target) {
        return (*) => ActionCatalog.OpenUrl(target)
    }

    static _Replacer(text, fn) {
        return (*) => ActionCatalog.PasteText(fn(text))
    }
}
