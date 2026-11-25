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
; Content types and the user's own actions (custom commands shown for a selection):
;   What the selection is and which content types it has (built-in: Text, Url, File, Folder, Files; the user's own in "ContentTypes"):
;   ContentTypes.Classify. A File type may name the command that opens its files, as an "Open" of Total Commander's internal associations
;   ("OpenCommand": the command's Id; it counts for the type even when the command's "ContentTypes" doesn't name it). Files and folders in
;   the search results get the same commands in their actions panel (ContentTypes.ForResult), the open command first; Enter on such a file
;   runs it (OpenActionFor) instead of opening the file with Windows' association. Programs and shortcuts are never opened this way.
;   A custom command with "ContentTypes": "PO, PDF" (type names, any case) is shown first in the panel of a selection of any of those types:
;     {query} in Target / Arguments becomes the selected text or path (encoded in URLs); several files become "a" "b" (each quoted);
;     a File / Command without {query} gets the selection appended as its last argument, quoted (as "Open with").
;   Selected text can't break out of the command line: a command can't take the text when it has a line break or a double quote
;   (except URLs, which are encoded), nor when it runs a shell (cmd, PowerShell, a batch file...) and the text has that shell's special characters;
;   its row stays in the panel saying why, and Enter on it only says so. PowerShell started directly takes any of them where the selection
;   is only text to it: an argument of the script after -File, or a single-quoted string of -Command (its single quotes doubled).
;   Scripts with "@altrun.selection PO, PDF" (or "@altrun.types") in their header come after the custom commands, with the same checks (ScriptProvider.ActionsFor).
;   When no command or script is set for the selection's types, the panel ends with "Add a Command for This Content..." (the command
;   edit dialog with "Content types" filled in) and, when none of the user's types matched, "Make a Content Type from It..." (the type
;   edit dialog with a pattern guessed from the selection, GuessType; then the command dialog for the new type). With Preferences open
;   both go to its pages, so the next Apply doesn't overwrite them.
;
; Usage:
;   SelectionActions.Run()                    called by the hotkey
;   SelectionActions.ItemFor({Text: "..."})   -> ResultItem (the source of the actions panel); can also be {Files: [...]}
;   SelectionActions.GuessType(info)          -> {Name, Kind, Pattern} Map for "Make a Content Type from It...", "" when there is no good one
;   SelectionActions.TypeEditorFields([sample, typeName]) the fields of a content type in the edit dialog (sample: a line saying whether
;                                             it matches; typeName: its name when the dialog opens: its commands come first in "Open (Enter)")
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
        ClipboardData.Borrow(SelectionActions.WaitMs + SelectionActions.RetryMs + 500)
        saved := ClipboardAll(), result := ""
        try {                                                               ; the user's clipboard comes back even when this fails half way
            A_Clipboard := ""
            Send(RegExMatch(className, SelectionActions.ConsoleClasses) ? "^{Insert}" : "^c")
            result := SelectionActions.ReadCopied(className)
        } finally
            A_Clipboard := saved
        return result
    }

    ; Reads what the copy just put on the emptied clipboard: {Files: [...]} for files (CF_HDROP), {Text: "..."} for text, "" for nothing.
    ; Excel (#513) doesn't put text there at once: its formats come one by one and the text is made only when a program reads it,
    ; so the first format can come before the text, and the first read can be empty. So: wait up to WaitMs for files or text
    ; (CF_UNICODETEXT, which Windows also offers for CF_TEXT), then read empty text again every 50 ms for up to RetryMs.
    ; With nothing selected this takes WaitMs. With the debug log on, a failed read logs the formats seen and the time waited
    static WaitMs := 2000, RetryMs := 500
    static ReadCopied(className := "") {
        start := A_TickCount
        while !(hasFiles := ClipboardData.HasFiles()) && !DllCall("IsClipboardFormatAvailable", "UInt", 13) {
            if (A_TickCount - start >= SelectionActions.WaitMs)
                return SelectionActions._LogNothing(className, start, "no text or files")
            Sleep(20)
        }
        if hasFiles {
            files := []
            for line in StrSplit(A_Clipboard, "`n", "`r")
                if (line != "")
                    files.Push(line)
            return files.Length ? {Files: files} : SelectionActions._LogNothing(className, start, "no file paths")
        }
        retryUntil := A_TickCount + SelectionActions.RetryMs
        Loop {
            text := A_Clipboard
            if (Trim(text, " `t`r`n") != "")
                return {Text: text}
            if (A_TickCount >= retryUntil)
                return SelectionActions._LogNothing(className, start, "the text is empty")
            Sleep(50)
        }
    }

    static _LogNothing(className, start, reason) {
        if Logger.Enabled
            Logger.Debug("Selection: nothing read from [" className "] after " (A_TickCount - start) " ms, " reason
                "; clipboard formats: " SelectionActions.ClipboardFormats())
        return ""
    }

    ; The formats on the clipboard, "13 CF_UNICODETEXT, 49161 DataObject, ..." (at most 20; "none" when empty)
    static ClipboardFormats() {
        static names := Map(1, "CF_TEXT", 2, "CF_BITMAP", 3, "CF_METAFILEPICT", 4, "CF_SYLK", 5, "CF_DIF", 7, "CF_OEMTEXT", 8, "CF_DIB"
            , 13, "CF_UNICODETEXT", 14, "CF_ENHMETAFILE", 15, "CF_HDROP", 16, "CF_LOCALE", 17, "CF_DIBV5")
        if !ClipboardData._Open()
            return "(the clipboard is held by another program)"
        list := "", count := 0, formatId := 0
        try {
            while (formatId := DllCall("EnumClipboardFormats", "UInt", formatId, "UInt")) {
                if (++count > 20)
                    return list ", ..."
                name := names.Has(formatId) ? names[formatId] : ""
                if (name = "") {
                    buf := Buffer(256 * 2, 0)
                    if DllCall("GetClipboardFormatName", "UInt", formatId, "Ptr", buf, "Int", 256)
                        name := StrGet(buf)
                }
                list .= (list = "" ? "" : ", ") formatId (name != "" ? " " name : "")
            }
        } finally
            DllCall("CloseClipboard")
        return (list = "") ? "none" : list
    }

    static ItemFor(selection) {
        info := ContentTypes.Classify(selection)
        switch info.Kind {
            case "files"        : item := SelectionActions._FilesItem(info.Files)
            case "file", "folder": item := SelectionActions._FileItem(info.Path)
            case "url"          : item := ResultItem(info.Text, I18n.T("Selection.Link"), {Kind: "url", Arg: info.Text, Icon: "url:", Provider: "Selection"})
            default             : item := SelectionActions._TextItem(info.Text)
        }
        item.LeadActions := SelectionActions.CommandActions(info)
        item.TrailActions := SelectionActions.TrailActions(info)
        return item
    }

    ; The rows at the end of the panel when no command or script is set for any of the selection's types: "Add a Command for This
    ; Content..." for its most specific type (the first of the user's that matched, else the built-in one), and with none of the
    ; user's types matching, "Make a Content Type from It..." when a pattern can be guessed. None with custom commands turned off
    static TrailActions(info) {
        rows := []
        if (!ProviderRegistry.IsEnabled(CustomCommandProvider) || SelectionActions.HasCommands(info))
            return rows
        typeName := (info.Types.Length > 1) ? info.Types[2] : info.Types[1]
        rows.Push(ResultItem(I18n.T("Selection.AddCommand"), I18n.T("Selection.AddCommand.Sub", typeName)
            , {Icon: "res:imageres.dll,-2", OnRun: (*) => SelectionActions.AddCommandFor(typeName)}))
        if (info.Types.Length = 1 && IsObject(guess := SelectionActions.GuessType(info)))
            rows.Push(ResultItem(I18n.T("Selection.NewType"), SelectionActions.KindText(guess["Kind"]) " · " guess["Pattern"]
                , {Icon: "res:imageres.dll,-2", OnRun: (*) => SelectionActions.NewTypeFrom(info, guess)}))
        return rows
    }

    ; Whether a custom command (with custom commands on) or a script (with scripts on) is set for one of the selection's types,
    ; shown or not (a command left out for the text's shell characters still counts: adding another wouldn't help)
    static HasCommands(info) {
        if ProviderRegistry.IsEnabled(CustomCommandProvider)
            for command in AppSettings.CustomCommands
                if ContentTypes.CommandFor(command, info.Types)
                    return true
        if ProviderRegistry.IsEnabled(ScriptProvider)
            for script in ScriptProvider.Current()
                if ScriptProvider.ForTypes(script, info.Types)
                    return true
        return false
    }

    ; A content type guessed from a selection (Classify's result) -> Map {Name, Kind, Pattern}, "" when there is none:
    ;   text on one line, at most 100 characters: digits only ^\d{N}$; letters, then digits (an optional - _ or space between)
    ;   ^PO\d+$ named after the letters; an email address (a general pattern); anything else the text itself, escaped
    ;   a URL: its host, ^https?://(www\.)?github\.com(/|$), named after the host
    ;   one file or several: their extension when they all have the same one, named after it in capitals; a folder: none
    ; A suggested name already taken by one of the user's types is left out (the user names it)
    static GuessType(info, userTypes := "") {
        guess := ""
        switch info.Kind {
            case "url":
                if RegExMatch(info.Term, "i)^https?://(?:www\.)?([^/?#:]+)", &m)
                    guess := Map("Name", StrLower(m[1]), "Kind", "Text", "Pattern", "^https?://(www\.)?" SelectionActions.RegexEscape(StrLower(m[1])) "(/|$)")
            case "text":
                text := info.Term
                if (text = "" || InStr(text, "`n") || StrLen(text) > 100)
                    return ""
                if RegExMatch(text, "^\d+$")
                    guess := Map("Name", "", "Kind", "Text", "Pattern", "^\d{" StrLen(text) "}$")
                else if RegExMatch(text, "^([A-Za-z]+)([-_ ]?)\d+$", &m)
                    guess := Map("Name", m[1], "Kind", "Text", "Pattern", "^" SelectionActions.RegexEscape(m[1] m[2]) "\d+$")
                else if RegExMatch(text, "^[^@\s]+@[^@\s]+\.[^@\s]+$")
                    guess := Map("Name", "", "Kind", "Text", "Pattern", "^[\w.+-]+@[\w-]+(\.[\w-]+)+$")
                else
                    guess := Map("Name", "", "Kind", "Text", "Pattern", "^" SelectionActions.RegexEscape(text) "$")
            case "file", "files":
                ext := ""
                for filePath in ((info.Kind = "file") ? [info.Path] : info.Files) {
                    SplitPath(filePath, , , &each)
                    if (each = "" || InStr(FileExist(filePath), "D") || (ext != "" && each != ext))
                        return ""
                    ext := each
                }
                guess := Map("Name", StrUpper(ext), "Kind", "File", "Pattern", StrLower(ext))
        }
        if !IsObject(guess)
            return ""
        if (guess["Name"] != "")
            for each in (IsObject(userTypes) ? userTypes : ContentTypes.UserTypes())
                if (each.Name = guess["Name"])
                    guess["Name"] := ""
        return guess
    }

    ; Text matched literally by a regular expression
    static RegexEscape(text) => RegExReplace(text, "[\\.*?+\[\]{}()|^$]", "\$0")

    ; "Add a Command for This Content...": the command edit dialog with "Content types" set to typeName (in Preferences when it is open)
    static AddCommandFor(typeName) {
        prefill := Map("ContentTypes", typeName)
        if IsObject(PreferencesWindow.Gui) {
            PreferencesWindow.Show("Prefs.Page.Commands")
            command := CustomCommandProvider.NewCommand()
            command["ContentTypes"] := typeName
            return PreferencesWindow._lists["CustomCommands"].Insert.Call(command)
        }
        if CustomCommandProvider.Edit("", prefill)
            App.Toast(I18n.T("Custom.Added", AppSettings.CustomCommands[-1]["Title"]), 2500)
    }

    ; "Make a Content Type from It...": the type edit dialog with the guess and a line saying whether the selection matches; on OK the
    ; type is saved, then the command dialog for it opens (cancelling that keeps the type). In Preferences when it is open
    static NewTypeFrom(info, guess) {
        if IsObject(PreferencesWindow.Gui) {
            PreferencesWindow.Show("Prefs.Page.Selection")
            list := PreferencesWindow._lists["ContentTypes"], before := list.Items.Length
            list.Insert.Call(guess.Clone())
            if (list.Items.Length > before)
                SelectionActions.AddCommandFor(list.Items[-1]["Name"])
            return
        }
        edited := ItemEditor.Edit(ItemEditor.Owner(), I18n.T("Prefs.Page.Selection"), SelectionActions.TypeEditorFields(SelectionActions.SampleOf(info)), guess.Clone())
        if !IsObject(edited)
            return
        if !(AppSettings.Data.Has("ContentTypes") && AppSettings.Data["ContentTypes"] is Array)
            AppSettings.Data["ContentTypes"] := []
        AppSettings.Data["ContentTypes"].Push(edited)
        if !AppSettings.Save()
            return
        App.Toast(I18n.T("SelType.Added", edited["Name"]), 2500)
        SelectionActions.AddCommandFor(edited["Name"])
    }

    ; The text or path a type is tested with in its edit dialog: the selected text, the path, or the first of several files
    static SampleOf(info) {
        switch info.Kind {
            case "files":          return info.Files[1]
            case "file", "folder": return info.Path
        }
        return info.Term
    }

    ; The edit dialog's line about the sample: "PO12345678 matches" / "... doesn't match" (a type without a name is tried as one)
    static SampleText(sample, edited) {
        entry := edited.Clone()
        if (Trim(entry.Has("Name") ? entry["Name"] : "") = "")
            entry["Name"] := "?"
        shown := (StrLen(sample) > 60) ? TextTools.Head(sample, 59) "…" : sample
        return I18n.T((ContentTypes.SampleTypes(sample, [entry]).Length > 1) ? "SelType.Sample.Match" : "SelType.Sample.NoMatch", shown)
    }

    ;---------------------------------------------------------------------------
    ; The user's commands for the selection's content types (the types themselves: ContentTypes)
    ;---------------------------------------------------------------------------
    ; The custom commands for a selection (Classify's result), in the order of the command list (none with custom commands turned
    ; off), then the scripts for it (ScriptProvider.ActionsFor). The file type's open command (OpenCommand) comes first, so Enter in
    ; the panel does what Enter does on the file; it is there even when its "ContentTypes" doesn't name the type
    static CommandActions(info) {
        actions := []
        if ProviderRegistry.IsEnabled(CustomCommandProvider) {
            openCommand := ContentTypes.OpenCommand(info)
            if (IsObject(openCommand) && IsObject(action := SelectionActions._CommandAction(openCommand, info, true)))
                actions.Push(action)
            for command in AppSettings.CustomCommands
                if (command != openCommand && ContentTypes.CommandFor(command, info.Types) && IsObject(action := SelectionActions._CommandAction(command, info, true)))
                    actions.Push(action)
        }
        for action in ScriptProvider.ActionsFor(info)
            actions.Push(action)
        return actions
    }

    ; Enter on a file in the search results: the row of its type's open command (run with its OnRun), "" to open it as usual.
    ; Only the file name is looked at until one of the user's types with an open command has its extension, so Enter on other
    ; files (or with no open commands set) never touches the disk here
    static OpenActionFor(item) {
        if (!IsObject(item) || item.Kind != "file" || ContentTypes._IsProgram(item.Arg))
            return ""
        SplitPath(item.Arg, , , &ext)
        found := false
        for userType in ContentTypes.UserTypes()
            if (userType.OpenCommand != "" && ContentTypes._HasExt(userType, ext)) {
                found := true
                break
            }
        if (!found || !IsObject(info := ContentTypes.ForResult(item)) || !IsObject(command := ContentTypes.OpenCommand(info)))
            return ""
        return SelectionActions._CommandAction(command, info)
    }

    ; The titles of the commands (a list being edited in Preferences) that are shown for these types, then of the scripts
    ; ("Backup (script)"; scripts: ScriptProvider.Scripts-like objects, the Scripts folder's when not given)
    static CommandTitles(types, commands, scripts := "") {
        titles := []
        for command in commands
            if (command is Map && ContentTypes.CommandFor(command, types))
                titles.Push(command.Has("Title") ? command["Title"] : "")
        for script in (scripts is Array) ? scripts : ScriptProvider.Current()
            if ScriptProvider.ForTypes(script, types)
                titles.Push(I18n.T("Selection.ScriptTitle", script.Title))
        return titles
    }

    ;---------------------------------------------------------------------------
    ; Editing content types (Preferences → Selection)
    ;---------------------------------------------------------------------------
    ; The example commands (defaults with "ContentTypes") into commands: a missing one is added (a copy); one with the same title or keyword
    ; that isn't set for any type gets the example's types. Returns {Added, Linked}
    static AddExampleCommands(commands, examples) {
        added := 0, linked := 0
        for example in examples {
            if !(example.Has("ContentTypes") && example["ContentTypes"] != "")
                continue
            existing := ""
            for command in commands
                if (command is Map && SelectionActions._SameCommand(command, example)) {
                    existing := command
                    break
                }
            if !IsObject(existing) {
                commands.Push(example.Clone())
                added += 1
            } else if !ContentTypes.CommandTypes(existing).Length {
                existing["ContentTypes"] := example["ContentTypes"]
                linked += 1
            }
        }
        return {Added: added, Linked: linked}
    }

    static _SameCommand(command, example) {
        for key in PrefsPages.CommandKeys(command)
            for exampleKey in PrefsPages.CommandKeys(example)
                if (key != "" && key = exampleKey)
                    return true
        return false
    }

    ; The type names offered in a command's "Content types" menu: the built-in ones, then the user's (being edited in Preferences when it is open)
    static TypeNamesForEditor() {
        names := ContentTypes.BuiltInTypes.Clone()
        userTypes := ContentTypes.UserTypes(SelectionActions._EditingTypes())
        if userTypes.Length
            names.Push("-")
        for each in userTypes
            names.Push(each.Name)
        return names
    }

    ; The ContentTypes being edited in Preferences, "" when it isn't open (then the saved ones count)
    static _EditingTypes() {
        if (IsObject(PreferencesWindow.Gui) && IsObject(PreferencesWindow.Working) && PreferencesWindow.Working.Has("ContentTypes"))
            return PreferencesWindow.Working["ContentTypes"]
        return ""
    }

    ; "" when every name in a command's "Content types" is a known type, else a warning naming the others (a typo would never match)
    static UnknownTypesWarning(value) {
        known := Map(), known.CaseSense := false
        for each in SelectionActions.TypeNamesForEditor()
            known[each] := true
        unknown := ""
        for each in StrSplit(ItemEditor.TagsText(value), ",")
            if ((each := Trim(each)) != "" && !known.Has(each))
                unknown .= (unknown = "" ? "" : ", ") each
        return (unknown = "") ? "" : I18n.T("Cmd.Selection.Unknown", unknown)
    }

    ; "Open (Enter)": Windows' association, or a command that runs a program (IsProgramCommand; the commands being edited in Preferences
    ; when it is open): the ones set for this type (its name when the dialog opened) first, then the others. The command's Id is saved,
    ; so renaming the command keeps it; choosing a command is enough, its "ContentTypes" needn't name the type
    static _OpenCommandField(typeName) {
        choices := [["", I18n.T("SelType.OpenCommand.None")]], others := []
        commands := (IsObject(PreferencesWindow.Gui) && IsObject(PreferencesWindow.Working) && PreferencesWindow.Working.Has("CustomCommands"))
            ? PreferencesWindow.Working["CustomCommands"] : AppSettings.CustomCommands
        for command in commands {
            if !ContentTypes.IsProgramCommand(command)
                continue
            choice := [command["Id"], (command.Has("Title") && command["Title"] != "") ? command["Title"] : command["Id"]]
            if (Trim(typeName) != "" && ContentTypes.CommandFor(command, [Trim(typeName)]))
                choices.Push(choice)
            else
                others.Push(choice)
        }
        choices.Push(others*)
        field := ItemEditor.Field("OpenCommand", "SelType.OpenCommand", "choice", false, choices, I18n.T("SelType.Field.OpenCommand"))
        field.OmitEmpty := true                                             ; most types have none: no "OpenCommand": "" in the settings
        field.Check := (value, edited) => (value != "" && !(edited.Has("Kind") && edited["Kind"] = "File")) ? I18n.T("SelType.OpenCommand.TextKind") : ""
        return field
    }

    ; The title of the command with this Id in commands, "" when there is none (the types list's "Open (Enter)" column)
    static CommandTitle(id, commands) {
        if (id = "")
            return ""
        for command in (commands is Array) ? commands : []
            if (command is Map && command.Has("Id") && command["Id"] = id)
                return command.Has("Title") ? command["Title"] : id
        return ""
    }

    static NewType() {
        return Map("Name", "", "Kind", "Text", "Pattern", "")
    }

    static TypeEditorFields(sample := "", typeName := "") {
        kinds := [["Text", I18n.T("Prefs.SelKind.Text")], ["File", I18n.T("Prefs.SelKind.File")]]
        name := ItemEditor.Field("Name", "Prefs.Col.Name", "text", true, "", I18n.T("SelType.Field.Name"))
        name.Check := (value, *) => ContentTypes.IsBuiltInType(Trim(value)) ? I18n.T("SelType.BuiltInName", Trim(value)) : ""
        pattern := ItemEditor.Field("Pattern", "Prefs.Col.Pattern", "text", true, "", I18n.T("SelType.Field.Pattern"))
        pattern.Check := (value, edited) => ContentTypes.PatternProblem(edited.Has("Kind") ? edited["Kind"] : "Text", value)
        fields := [name, ItemEditor.Field("Kind", "Prefs.Col.Kind", "choice", false, kinds, I18n.T("SelType.Field.Kind")), pattern
            , SelectionActions._OpenCommandField(typeName)]
        if (sample != "") {                                                 ; opened from the selection: whether it matches, while editing
            line := ItemEditor.Field("Sample", "SelType.Sample", "info", false, "", "")
            line.Text := (edited) => SelectionActions.SampleText(sample, edited)
            fields.Push(line)
        }
        return fields
    }

    ; The text of a type's Kind in the list
    static KindText(kind) => I18n.T((kind = "File") ? "Prefs.SelKind.File" : "Prefs.SelKind.Text")

    ; The panel row of a custom command for the selection, "" when the selection can't be put into its command line safely
    ; (refusedRow: then a row saying why, RefusedRow)
    static _CommandAction(command, info, refusedRow := false) {
        commandType := command.Has("Type") ? command["Type"] : "File"
        takesQuery := CustomCommandProvider.TakesQuery(command)
        appends := !takesQuery && (commandType = "File" || commandType = "Command")
        if (commandType = "Url") {
            if (takesQuery && info.Kind = "files")                          ; several paths make no URL
                return ""
        } else if ((takesQuery || appends) && IsObject(problem := SelectionActions.ArgumentProblem(info, command))) {
            return refusedRow ? SelectionActions.RefusedRow(command["Title"], CustomCommandProvider._ToItem(command, 0).Icon, problem) : ""
        }
        term := info.Term
        if (takesQuery && SelectionActions.PowerShellMode(command) = "quoted")
            term := RegExReplace(term, "['‘’‚‛]", "$0$0")                  ; inside PowerShell's single quotes: a quote doubled is one quote
        item := CustomCommandProvider._ToItem(command, 0, takesQuery ? term : "")
        item.Title := command["Title"]
        item.CommandId := command.Has("Id") ? command["Id"] : ""            ; which command the row runs (the open command is found by it)
        if RegExMatch(item.Arg, "\\\.\.?\\") {                            ; A_ScriptDir\..\Typora\Typora.exe: shown and run as C:\Typora\Typora.exe
            shown := CustomCommandProvider._Resolve(item.Arg)               ; what the subtitle shows (A_ and %variables% expanded)
            if RegExMatch(shown, "^(?:[A-Za-z]:\\|\\\\)") {
                full := Path.Full(shown)
                item.Subtitle := StrReplace(item.Subtitle, shown, full), item.Arg := full
            }
        }
        item.Keyword := command.Has("Keyword") ? command["Keyword"] : ""     ; found by it when typing in the panel
        item.AutoComplete := ""
        if appends {
            item.Arguments := Trim(item.Arguments " " info.Quoted)
            item.Subtitle := Trim(item.Subtitle " " info.Quoted)
        }
        item.OnRun := SelectionActions._Runner(item.Kind, item.Arg, item.Arguments, item.WindowArgs)   ; the actions panel runs a row through its OnRun
        return item
    }

    ; each row keeps its own target (a separate method makes the closure)
    static _Runner(kind, target, arguments, windowArgs := false) {
        runItem := ResultItem("", "", {Kind: kind, Arg: target, Arguments: arguments, WindowArgs: windowArgs})
        return (*) => ActionCatalog.RunDefault(runItem)
    }

    ; Whether the selection can go into the command's command line (ArgumentProblem finds nothing)
    static SafeArgument(info, command) => !IsObject(SelectionActions.ArgumentProblem(info, command))

    ; Why the selection can't go into the command's command line: {Chars: the characters refused, Shell: the shells that would read
    ; them as code ("" for a line break or a double quote, which break any command line)}; "" when it can.
    ; Paths have no double quotes; text can't have a line break or a double quote; shells also refuse their special characters. A shell
    ; is looked for in the target after expanding it (%ComSpec%, A_WinDir, a name found on PATH) and anywhere in the arguments, so a
    ; shell started through another program ("wt.exe cmd /k ...", "conhost.exe powershell ...") is caught too. PowerShell started
    ; directly takes them where the selection is only text to it (PowerShellMode). quoted: ALTRun itself quotes the selection as one
    ; argument (Win.QuoteArg, a script started directly), so a line break or a double quote arrives as it is; the shell checks still apply
    static ArgumentProblem(info, command, quoted := false) {
        if (!quoted && (info.Kind = "text" || info.Kind = "url") && RegExMatch(info.Term, '[\r\n"]'))
            return {Chars: SelectionActions._CharsIn(info.Term, '[\r\n"]'), Shell: ""}
        target := SearchQuery.PlainPlaceholder(command.Has("Target") ? command["Target"] : "")   ; a {query|default} isn't read as part of the command line
        arguments := SearchQuery.PlainPlaceholder(command.Has("Arguments") ? command["Arguments"] : "")
        line := target " " CustomCommandProvider._Resolve(Trim(target, '" ')) " " Path.ExpandEnv(arguments)
        switch (shell := SelectionActions.ShellIn(line)) {
            case "cmd": special := "[&|<>^%!]"
            case "":    return ""
            default:    special := "[&|<>^%!``$;(){}@']"                  ; PowerShell / Unix shells: also cmd's characters (cmd may be in between)
        }
        if (!RegExMatch(info.Term, special) || shell = "powershell" && SelectionActions._TextToPowerShell(info, command))
            return ""
        names := ""
        for name in SelectionActions.ShellNames(line)
            names .= (names = "" ? "" : " / ") name
        return {Chars: SelectionActions._CharsIn(info.Term, special), Shell: names}
    }

    ; PowerShell started directly reads the selection as text (PowerShellMode). In -Command it goes into single quotes: not several
    ; quoted paths, and not text ending in \ right before a " of the command line (that would escape the quote)
    static _TextToPowerShell(info, command) {
        switch SelectionActions.PowerShellMode(command) {
            case "file":
                return true
            case "quoted":
                arguments := SearchQuery.PlainPlaceholder(command.Has("Arguments") ? command["Arguments"] : "")
                return info.Kind != "files" && !(SubStr(info.Term, -1) = "\" && RegExMatch(arguments, 'i)\{query\}"'))
        }
        return false
    }

    ; How PowerShell started directly (the target is powershell.exe or pwsh.exe) reads the selection, by where {query} is (a command
    ; without {query} gets it as its last argument): "file" when it's an argument of the script after -File (pwsh: also a script named
    ; without -File, its default), "quoted" when every {query} is inside a single-quoted string of -Command (powershell.exe: also code
    ; written without -Command, its default), else "" (as code, or a command line not understood: the strict check).
    ; {query|default} counts as {query} (its default could have spaces or quotes)
    static PowerShellMode(command) {
        target := Trim(SearchQuery.PlainPlaceholder(command.Has("Target") ? command["Target"] : ""), '" ')
        if InStr(target, "{query}")
            return ""
        SplitPath(CustomCommandProvider._Resolve(target), &name)
        if !RegExMatch(name, "i)^(powershell|pwsh)(\.exe)?$", &program)
            return ""
        arguments := SearchQuery.PlainPlaceholder(command.Has("Arguments") ? command["Arguments"] : "")
        args := Win.Args(Path.ExpandEnv(arguments) (CustomCommandProvider.TakesQuery(command) ? "" : " {query}"))
        index := 0
        while (++index <= args.Length) {
            arg := args[index]
            if !RegExMatch(arg, "^-(.+)$", &parameter)                      ; the first word that isn't a parameter
                return (program[1] = "pwsh") ? SelectionActions._AfterScript(args, index) : SelectionActions._QuotedCode(args, index)
            if InStr(arg, "{query}")                                        ; in one of PowerShell's own parameters
                return ""
            switch SelectionActions._PowerShellParameter(parameter[1]) {
                case "file":    return SelectionActions._AfterScript(args, index + 1)
                case "command": return SelectionActions._QuotedCode(args, index + 1)
                case "value":   index++                                     ; its value is checked as the next word would be
                    if (index <= args.Length && InStr(args[index], "{query}"))
                        return ""
                case "switch":  continue
                default:        return ""                                   ; -EncodedCommand, -CommandWithArgs, an unknown abbreviation...
            }
        }
        return ""
    }

    ; "file" when args[index] is the script and {query} comes only after it: PowerShell passes the script its arguments as text
    static _AfterScript(args, index) => (index <= args.Length && !InStr(args[index], "{query}")) ? "file" : ""

    ; "quoted" when the code (args from index on, joined by spaces as PowerShell does) has every {query} in single quotes, and
    ; nothing in it runs a string as code again (Invoke-Expression, a script block, another PowerShell or a process started from it)
    static _QuotedCode(args, index) {
        code := ""
        while (index <= args.Length)
            code .= (code = "" ? "" : " ") args[index++]
        if RegExMatch(code, "i)(?<![\w-])(iex|invoke-expression|invoke-command|icm|scriptblock|expandstring|start-process|saps|start|powershell|pwsh)(?![\w-])")
            return ""
        return SelectionActions.InSingleQuotes(code) ? "quoted" : ""
    }

    ; A PowerShell command-line parameter (without its -): "file", "command", "switch" (takes no value), "value" (takes the next word)
    ; or "" (not known: the strict check). Abbreviations of 3 letters or more and the usual short names are known
    static _PowerShellParameter(name) {
        static switches := ["noprofile", "nologo", "noexit", "noninteractive", "noprofileloadtime", "sta", "mta", "login", "interactive"]
        static values := ["executionpolicy", "windowstyle", "workingdirectory", "version", "inputformat", "outputformat", "psconsolefile"
            , "configurationname", "configurationfile", "settingsfile", "custompipename"]
        static shortValues := Map("ep", 1, "ex", 1, "w", 1, "wd", 1)
        name := StrLower(name)
        if (name = "f" || name = "file")
            return "file"
        if (name = "c" || name = "command")
            return "command"
        if shortValues.Has(name)
            return "value"
        if (StrLen(name) >= 3) {
            for each in switches
                if (InStr(each, name) = 1)
                    return "switch"
            for each in values
                if (InStr(each, name) = 1)
                    return "value"
        }
        return ""
    }

    ; Whether PowerShell code has {query} and every {query} is inside a single-quoted string ('...', also with PowerShell's other single
    ; quotes ‘ ’ ‚ ‛), where nothing but a quote is special. Double-quoted strings ($ and ` work there) and comments end the scan: false
    static InSingleQuotes(code) {
        static single := "'‘’‚‛", double := '"“”„'
        state := "", found := false, index := 0, length := StrLen(code)
        while (++index <= length) {
            char := SubStr(code, index, 1)
            if (SubStr(code, index, 7) = "{query}") {
                if (state != "single")
                    return false
                found := true, index += 6
                continue
            }
            switch state {
                case "single":
                    if InStr(single, char)
                        if (index < length && InStr(single, SubStr(code, index + 1, 1)))
                            index++                                         ; doubled: a quote inside the string
                        else
                            state := ""
                case "double":
                    if (char = "``")
                        index++
                    else if InStr(double, char)
                        if (index < length && InStr(double, SubStr(code, index + 1, 1)))
                            index++
                        else
                            state := ""
                default:
                    if (char = "``")
                        index++
                    else if InStr(single, char)
                        state := "single"
                    else if InStr(double, char)
                        state := "double"
                    else if (char = "#")                                    ; a comment (# or <# #>)
                        return false
            }
        }
        return found && state = ""
    }

    ; The characters of text matching pattern (a character class), each once, separated by spaces; a line break is named
    static _CharsIn(text, pattern) {
        chars := "", seen := Map()
        Loop Parse, text {
            char := (A_LoopField = "`r" || A_LoopField = "`n") ? I18n.T("Selection.LineBreak") : A_LoopField
            if (!seen.Has(char) && !TextTools.IsLoneSurrogate(A_LoopField) && RegExMatch(A_LoopField, pattern) > 0) {   ; half an emoji: RegExMatch throws
                seen[char] := true
                chars .= (chars = "" ? "" : " ") char
            }
        }
        return chars
    }

    ; The panel row of a command or script that can't take the selection (ArgumentProblem's problem): its subtitle says why, and Enter
    ; says it again (for PowerShell, with how to write the command so it can)
    static RefusedRow(title, icon, problem) {
        text := (problem.Shell = "") ? I18n.T("Selection.Refused", problem.Chars) : I18n.T("Selection.Refused.Shell", problem.Chars, problem.Shell)
        notice := (problem.Shell = "PowerShell") ? text "`n" I18n.T("Selection.Refused.PowerShell") : text
        return ResultItem(title, text, {Icon: icon, Refused: true, OnRun: (*) => App.Notify(notice, 5000)})
    }

    ; The shells a command line names: PowerShell, a Unix shell (by its name), cmd. Whole names only (not "-Command");
    ; a false match only refuses more selections
    static ShellNames(commandLine) {
        names := []
        if RegExMatch(commandLine, "i)(?<![\w.-])(powershell|powershell_ise|pwsh)(\.exe)?(?![\w.-])|\.ps1(?![\w.])")
            names.Push("PowerShell")
        if RegExMatch(commandLine, "i)(?<![\w.-])(bash|sh|wsl|zsh)(\.exe)?(?![\w.-])", &unix)
            names.Push(StrLower(unix[1]))
        if RegExMatch(commandLine, "i)(?<![\w.-])(cmd|command)(\.exe|\.com)?(?![\w.-])|\.(bat|cmd)(?![\w.])")
            names.Push("cmd")
        return names
    }

    ; "powershell" when a command line names PowerShell and no other shell, "cmd" when it names cmd or a batch file and no other
    ; shell, "posix" when it names a Unix shell or more than one shell (the strictest check), else ""
    static ShellIn(commandLine) {
        names := SelectionActions.ShellNames(commandLine)
        if (names.Length = 0)
            return ""
        return (names.Length = 1 && (names[1] = "PowerShell" || names[1] = "cmd")) ? StrLower(names[1]) : "posix"
    }

    ; One command-line argument in double quotes; trailing backslashes are doubled so "C:\" doesn't escape the closing quote
    static Quote(value) {
        if RegExMatch(value, "\\+$", &m)
            value .= m[0]
        return '"' value '"'
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
            preview := TextTools.Head(preview, 80) "..."
        actions := []
        ; an expression: give the result directly
        expression := StrReplace(StrReplace(StrReplace(Trim(text), ","), Chr(0xD7), "*"), Chr(0xF7), "/")    ; × ÷
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
            searchUrl := WebSearchProvider.FillAll(engine["Url"], query)          ; several addresses: one per line, all opened
            actions.Push(ResultItem(I18n.T("Action.SearchWith", engine["Title"]), "", {Icon: WebSearchProvider.IconFor(engine), OnRun: SelectionActions._Opener(searchUrl)
                , Keyword: engine.Has("Keyword") ? engine["Keyword"] : ""}))
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
