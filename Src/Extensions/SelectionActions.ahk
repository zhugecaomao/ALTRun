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
;   Built-in types: Text, Url, File, Folder, Files (several files / folders). The user's own types are in ALTRun.json:
;     "SelectionTypes": [ { "Name": "PO", "Kind": "Text", "Pattern": "^PO\\d{8}$" },    Text: a regular expression, matched against the trimmed text
;                         { "Name": "PDF", "Kind": "File", "Pattern": "pdf, docx" } ]   File: extensions; several files match when all of them do
;   A text selection is Text plus every Text type it matches; a URL is Url plus the Text types it matches; one file is File (a folder: Folder)
;   plus the File types of its extension; several files are Files plus the File types all of them match.
;   A custom command with "Selection": "PO, PDF" (type names, any case) is shown first in the panel of a selection of any of those types:
;     {query} in Target / Arguments becomes the selected text or path (encoded in URLs); several files become "a" "b" (each quoted);
;     a File / Command without {query} gets the selection appended as its last argument, quoted (as "Open with").
;   Selected text can't break out of the command line: a command isn't shown when the text has a line break or a double quote
;   (except URLs, which are encoded), nor when it runs a shell (cmd, PowerShell, a batch file...) and the text has that shell's special characters.
;   Scripts with "@altrun.selection PO, PDF" in their header come after the custom commands, with the same checks (ScriptProvider.ActionsFor).
;   When no command or script is set for the selection's types, the panel ends with "Add a Command for This Content..." (the command
;   edit dialog with "Selection types" filled in) and, when none of the user's types matched, "Make a Content Type from It..." (the type
;   edit dialog with a pattern guessed from the selection, GuessType; then the command dialog for the new type). With Preferences open
;   both go to its pages, so the next Apply doesn't overwrite them.
;
; Usage:
;   SelectionActions.Run()                    called by the hotkey
;   SelectionActions.ItemFor({Text: "..."})   -> ResultItem (the source of the actions panel); can also be {Files: [...]}
;   SelectionActions.Classify(selection)      -> {Kind, Types, Term, ...}: what the selection is and which content types it has
;   SelectionActions.SampleTypes(text, list)  the types of a sample text or path (the test box in Preferences → Selection), list = SelectionTypes being edited
;   SelectionActions.GuessType(info)          -> {Name, Kind, Pattern} Map for "Make a Content Type from It...", "" when there is no good one
;   SelectionActions.TypeEditorFields([sample]) the fields of a content type in the edit dialog (sample: a line saying whether it matches)
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
        saved := ClipboardAll(), result := ""
        try {                                                               ; the user's clipboard comes back even when this fails half way
            A_Clipboard := ""
            Send(RegExMatch(className, SelectionActions.ConsoleClasses) ? "^{Insert}" : "^c")
            if ClipWait(0.8, 1) {
                if DllCall("IsClipboardFormatAvailable", "UInt", 15) {     ; CF_HDROP: files were copied
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
        } finally
            A_Clipboard := saved
        return result
    }

    static ItemFor(selection) {
        info := SelectionActions.Classify(selection)
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
                if SelectionActions.CommandFor(command, info.Types)
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
            for each in (IsObject(userTypes) ? userTypes : SelectionActions.UserTypes())
                if (each.Name = guess["Name"])
                    guess["Name"] := ""
        return guess
    }

    ; Text matched literally by a regular expression
    static RegexEscape(text) => RegExReplace(text, "[\\.*?+\[\]{}()|^$]", "\$0")

    ; "Add a Command for This Content...": the command edit dialog with "Selection types" set to typeName (in Preferences when it is open)
    static AddCommandFor(typeName) {
        prefill := Map("Selection", typeName)
        if IsObject(PreferencesWindow.Gui) {
            PreferencesWindow.Show("Prefs.Page.Commands")
            command := CustomCommandProvider.NewCommand()
            command["Selection"] := typeName
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
            list := PreferencesWindow._lists["SelectionTypes"], before := list.Items.Length
            list.Insert.Call(guess.Clone())
            if (list.Items.Length > before)
                SelectionActions.AddCommandFor(list.Items[-1]["Name"])
            return
        }
        edited := ItemEditor.Edit(ItemEditor.Owner(), I18n.T("Prefs.Page.Selection"), SelectionActions.TypeEditorFields(SelectionActions.SampleOf(info)), guess.Clone())
        if !IsObject(edited)
            return
        if !(AppSettings.Data.Has("SelectionTypes") && AppSettings.Data["SelectionTypes"] is Array)
            AppSettings.Data["SelectionTypes"] := []
        AppSettings.Data["SelectionTypes"].Push(edited)
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
        shown := (StrLen(sample) > 60) ? SubStr(sample, 1, 59) "…" : sample
        return I18n.T((SelectionActions.SampleTypes(sample, [entry]).Length > 1) ? "SelType.Sample.Match" : "SelType.Sample.NoMatch", shown)
    }

    ;---------------------------------------------------------------------------
    ; Content types and the user's commands for them
    ;---------------------------------------------------------------------------
    static BuiltInTypes := ["Text", "Url", "File", "Folder", "Files"]

    ; What the selection is: {Kind: text / url / file / folder / files, Types: [type names], Term: what {query} becomes,
    ; Quoted: Term as one quoted argument (several files: each quoted), Text / Path / Files}
    ; userTypes: UserTypes() of the saved settings when not given
    static Classify(selection, userTypes := "") {
        if !IsObject(userTypes)
            userTypes := SelectionActions.UserTypes()
        if selection.HasOwnProp("Files") {
            files := selection.Files
            if (files.Length = 1)
                return SelectionActions._PathInfo(files[1], userTypes)
            types := ["Files"]
            for userType in userTypes
                if (userType.Kind = "File" && SelectionActions._AllMatch(userType, files))
                    types.Push(userType.Name)
            quoted := ""
            for filePath in files
                quoted .= (quoted = "" ? "" : " ") SelectionActions.Quote(filePath)
            return {Kind: "files", Types: types, Term: quoted, Quoted: quoted, Files: files}
        }
        text := selection.Text
        trimmed := Trim(text, " `t`r`n")
        if RegExMatch(trimmed, "i)^https?://\S+$")
            return {Kind: "url", Types: SelectionActions._TextTypes(trimmed, "Url", userTypes), Term: trimmed, Quoted: SelectionActions.Quote(trimmed), Text: trimmed}
        if (!InStr(trimmed, "`n") && RegExMatch(trimmed, "^(?:[A-Za-z]:\\|\\\\)") && FileExist(trimmed))   ; the selection is one path
            return SelectionActions._PathInfo(trimmed, userTypes)
        return {Kind: "text", Types: SelectionActions._TextTypes(trimmed, "Text", userTypes), Term: trimmed, Quoted: SelectionActions.Quote(trimmed), Text: text}
    }

    static _PathInfo(filePath, userTypes) {
        isFolder := (InStr(FileExist(filePath), "D") || SubStr(filePath, -1) = "\") ? true : false
        types := [isFolder ? "Folder" : "File"]
        if !isFolder
            for userType in userTypes
                if (userType.Kind = "File" && SelectionActions._ExtMatches(userType, filePath))
                    types.Push(userType.Name)
        return {Kind: isFolder ? "folder" : "file", Types: types, Term: filePath, Quoted: SelectionActions.Quote(filePath), Path: filePath}
    }

    ; builtIn first, then each Text type whose pattern matches (a pattern that isn't a valid regular expression is skipped, logged once)
    static _TextTypes(text, builtIn, userTypes) {
        static badPatterns := Map()
        types := [builtIn]
        for userType in userTypes {
            if (userType.Kind != "Text")
                continue
            try {
                if (RegExMatch(text, userType.Pattern) > 0)
                    types.Push(userType.Name)
            } catch as err {
                if !badPatterns.Has(userType.Pattern) {
                    badPatterns[userType.Pattern] := true
                    Logger.Warn("Selection type " userType.Name ": bad pattern " userType.Pattern " (" err.Message ")")
                }
            }
        }
        return types
    }

    static _AllMatch(userType, files) {
        for filePath in files
            if !SelectionActions._ExtMatches(userType, filePath)
                return false
        return true
    }

    ; a file (not a folder) whose extension is one of the type's ("pdf, .docx" - separated by commas, semicolons, | or spaces)
    static _ExtMatches(userType, filePath) {
        SplitPath(filePath, , , &ext)
        if (ext = "")
            return false
        for each in StrSplit(userType.Pattern, [",", ";", "|", " ", "`t"])
            if ((each := LTrim(Trim(each), ".")) != "" && each = ext)
                return !InStr(FileExist(filePath), "D")                    ; the disk is asked only when the extension matches
        return false
    }

    ; The user's content types from ALTRun.json "SelectionTypes": [{Name, Kind: Text | File, Pattern}]; entries without a name or pattern,
    ; or named like a built-in type, are left out. list: the entries to read (the copy being edited in Preferences), the saved ones when not given
    static UserTypes(list := "") {
        types := []
        if !IsObject(list)
            list := (AppSettings.Data.Has("SelectionTypes") && AppSettings.Data["SelectionTypes"] is Array) ? AppSettings.Data["SelectionTypes"] : []
        for entry in list {
            if !(entry is Map)
                continue
            name := Trim(entry.Has("Name") ? entry["Name"] : "")
            pattern := entry.Has("Pattern") ? entry["Pattern"] : ""
            if (name = "" || Trim(pattern) = "" || SelectionActions.IsBuiltInType(name))
                continue
            kind := (entry.Has("Kind") && entry["Kind"] = "File") ? "File" : "Text"
            types.Push({Name: name, Kind: kind, Pattern: pattern})
        }
        return types
    }

    static IsBuiltInType(name) {
        for each in SelectionActions.BuiltInTypes
            if (each = name)
                return true
        return false
    }

    ; The type names of a custom command's "Selection" ("PO, PDF", or an array)
    static CommandTypes(command) {
        names := []
        if !(command is Map && command.Has("Selection"))
            return names
        value := command["Selection"]
        for each in (value is Array) ? value : StrSplit(value, ",")
            if (!IsObject(each) && (each := Trim(each)) != "")
                names.Push(each)
        return names
    }

    ; The custom commands for a selection (Classify's result), in the order of the command list (none with custom commands turned
    ; off), then the scripts for it (ScriptProvider.ActionsFor)
    static CommandActions(info) {
        actions := []
        if ProviderRegistry.IsEnabled(CustomCommandProvider)
            for command in AppSettings.CustomCommands
                if (SelectionActions.CommandFor(command, info.Types) && IsObject(action := SelectionActions._CommandAction(command, info)))
                    actions.Push(action)
        for action in ScriptProvider.ActionsFor(info)
            actions.Push(action)
        return actions
    }

    ; Whether a custom command is shown for a selection of these types (any of its "Selection" names, any case)
    static CommandFor(command, types) {
        for name in SelectionActions.CommandTypes(command)
            for each in types
                if (each = name)
                    return true
        return false
    }

    ; The titles of the commands (a list being edited in Preferences) that are shown for these types, then of the scripts
    ; ("Backup (script)"; scripts: ScriptProvider.Scripts-like objects, the Scripts folder's when not given)
    static CommandTitles(types, commands, scripts := "") {
        titles := []
        for command in commands
            if (command is Map && SelectionActions.CommandFor(command, types))
                titles.Push(command.Has("Title") ? command["Title"] : "")
        for script in (scripts is Array) ? scripts : ScriptProvider.Current()
            if ScriptProvider.ForTypes(script, types)
                titles.Push(I18n.T("Selection.ScriptTitle", script.Title))
        return titles
    }

    ;---------------------------------------------------------------------------
    ; Editing content types (Preferences → Selection)
    ;---------------------------------------------------------------------------
    ; The types of a sample in the test box: a path (C:\..., \\server\...) counts as a selected file without looking at the disk
    ; (a folder when it ends with \ or exists as one), anything else as selected text; list = the SelectionTypes being edited
    static SampleTypes(text, list) {
        userTypes := SelectionActions.UserTypes(list)
        trimmed := Trim(text, " `t`r`n")
        if (trimmed = "")
            return []
        if (!InStr(trimmed, "`n") && RegExMatch(trimmed, "^(?:[A-Za-z]:\\|\\\\)"))
            return SelectionActions._PathInfo(trimmed, userTypes).Types
        return SelectionActions.Classify({Text: text}, userTypes).Types
    }

    ; The example commands (defaults with "Selection") into commands: a missing one is added (a copy); one with the same title or keyword
    ; that isn't set for any type gets the example's types. Returns {Added, Linked}
    static AddExampleCommands(commands, examples) {
        added := 0, linked := 0
        for example in examples {
            if !(example.Has("Selection") && example["Selection"] != "")
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
            } else if !SelectionActions.CommandTypes(existing).Length {
                existing["Selection"] := example["Selection"]
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

    ; The type names offered in a command's "Selection types" menu: the built-in ones, then the user's (being edited in Preferences when it is open)
    static TypeNamesForEditor() {
        names := SelectionActions.BuiltInTypes.Clone()
        userTypes := SelectionActions.UserTypes(SelectionActions._EditingTypes())
        if userTypes.Length
            names.Push("-")
        for each in userTypes
            names.Push(each.Name)
        return names
    }

    ; The SelectionTypes being edited in Preferences, "" when it isn't open (then the saved ones count)
    static _EditingTypes() {
        if (IsObject(PreferencesWindow.Gui) && IsObject(PreferencesWindow.Working) && PreferencesWindow.Working.Has("SelectionTypes"))
            return PreferencesWindow.Working["SelectionTypes"]
        return ""
    }

    ; "" when every name in a command's "Selection types" is a known type, else a warning naming the others (a typo would never match)
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

    static NewType() {
        return Map("Name", "", "Kind", "Text", "Pattern", "")
    }

    static TypeEditorFields(sample := "") {
        kinds := [["Text", I18n.T("Prefs.SelKind.Text")], ["File", I18n.T("Prefs.SelKind.File")]]
        name := ItemEditor.Field("Name", "Prefs.Col.Name", "text", true, "", I18n.T("SelType.Field.Name"))
        name.Check := (value, *) => SelectionActions.IsBuiltInType(Trim(value)) ? I18n.T("SelType.BuiltInName", Trim(value)) : ""
        pattern := ItemEditor.Field("Pattern", "Prefs.Col.Pattern", "text", true, "", I18n.T("SelType.Field.Pattern"))
        pattern.Check := (value, edited) => SelectionActions.PatternProblem(edited.Has("Kind") ? edited["Kind"] : "Text", value)
        fields := [name, ItemEditor.Field("Kind", "Prefs.Col.Kind", "choice", false, kinds, I18n.T("SelType.Field.Kind")), pattern]
        if (sample != "") {                                                 ; opened from the selection: whether it matches, while editing
            line := ItemEditor.Field("Sample", "SelType.Sample", "info", false, "", "")
            line.Text := (edited) => SelectionActions.SampleText(sample, edited)
            fields.Push(line)
        }
        return fields
    }

    ; "" when the pattern can be used, else the warning: a Text type's pattern must be a valid regular expression
    static PatternProblem(kind, pattern) {
        if (kind = "File" || Trim(pattern) = "")
            return ""
        try RegExMatch("", pattern)
        catch as err
            return I18n.T("SelType.BadPattern", err.Message)
        return ""
    }

    ; The text of a type's Kind in the list
    static KindText(kind) => I18n.T((kind = "File") ? "Prefs.SelKind.File" : "Prefs.SelKind.Text")

    ; The panel row of a custom command for the selection, "" when the selection can't be put into its command line safely
    static _CommandAction(command, info) {
        commandType := command.Has("Type") ? command["Type"] : "File"
        takesQuery := CustomCommandProvider.TakesQuery(command)
        appends := !takesQuery && (commandType = "File" || commandType = "Command")
        if (commandType = "Url") {
            if (takesQuery && info.Kind = "files")                          ; several paths make no URL
                return ""
        } else if ((takesQuery || appends) && !SelectionActions.SafeArgument(info, command)) {
            return ""
        }
        item := CustomCommandProvider._ToItem(command, 0, takesQuery ? info.Term : "")
        item.Title := command["Title"]
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

    ; Whether the selection can go into the command's command line: paths always can (they have no double quotes);
    ; text can't have a line break or a double quote; shells also refuse their special characters. A shell is looked for
    ; in the target after expanding it (%ComSpec%, A_WinDir, a name found on PATH) and anywhere in the arguments, so a
    ; shell started through another program ("wt.exe cmd /k ...", "conhost.exe powershell ...") is caught too
    static SafeArgument(info, command) {
        if (info.Kind = "text" || info.Kind = "url") {
            if RegExMatch(info.Term, '[\r\n"]')
                return false
        }
        target := command.Has("Target") ? command["Target"] : ""
        arguments := command.Has("Arguments") ? command["Arguments"] : ""
        switch SelectionActions.ShellIn(target " " CustomCommandProvider._Resolve(Trim(target, '" ')) " " Path.ExpandEnv(arguments)) {
            case "posix": return !RegExMatch(info.Term, "[&|<>^%!``$;(){}@']")   ; PowerShell / Unix shells: also cmd's characters (cmd may be in between)
            case "cmd":   return !RegExMatch(info.Term, "[&|<>^%!]")
        }
        return true
    }

    ; "posix" when a command line names PowerShell or a Unix shell, "cmd" when it names cmd or a batch file, else "".
    ; Whole names only (not "-Command"); a false match only refuses more selections
    static ShellIn(commandLine) {
        if RegExMatch(commandLine, "i)(?<![\w.-])(powershell|powershell_ise|pwsh|bash|sh|wsl|zsh)(\.exe)?(?![\w.-])|\.ps1(?![\w.])")
            return "posix"
        if RegExMatch(commandLine, "i)(?<![\w.-])(cmd|command)(\.exe|\.com)?(?![\w.-])|\.(bat|cmd)(?![\w.])")
            return "cmd"
        return ""
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
