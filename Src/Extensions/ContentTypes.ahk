;===============================================================================
; ContentTypes.ahk - what a selection or a file is: the built-in and the user's content types (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Built-in types: Text, Url, File, Folder, Files (several files / folders). The user's own types are in ALTRun.json:
;   "ContentTypes": [ { "Name": "PO", "Kind": "Text", "Pattern": "^PO\\d{8}$" },    Text: a regular expression, matched against the trimmed text
;                     { "Name": "PDF", "Kind": "File", "Pattern": "pdf, docx", "OpenCommand": "<command Id>" } ]   File: extensions
; A text selection is Text plus every Text type it matches; a URL is Url plus the Text types it matches; one file is File (a folder: Folder)
; plus the File types of its extension; several files are Files plus the File types all of them match.
; A custom command is shown for the types in its "ContentTypes" ("PO, PDF", type names, any case); a script for the ones in its
; "@altrun.selection" (or "@altrun.types") header line (ScriptProvider). A File type's OpenCommand is the command Enter opens its files
; with (OpenCommand); programs and shortcuts (ProgramTypes) are never opened that way.
; Up to settings version 7 the list was "SelectionTypes" and a command's field "Selection" (AppSettings.RenameContentTypes).
; The panel rows and the editing in Preferences that use these: SelectionActions.
;
; Usage:
;   ContentTypes.Classify(selection)          -> {Kind, Types, Term, ...}: what the selection ({Text} / {Files}) is and which types it has
;   ContentTypes.ForResult(item)              the same for a file or folder in the search results
;   ContentTypes.UserTypes([list])            the user's types, as objects
;   ContentTypes.CommandFor(command, types)   whether a custom command is shown for these types
;   ContentTypes.OpenCommand(info)            the command Enter opens this file with, ""
;   ContentTypes.SampleTypes(text, list)      the types of a sample text or path (the test box in Preferences → Selection)
;===============================================================================

class ContentTypes {
    static BuiltInTypes := ["Text", "Url", "File", "Folder", "Files"]

    ; What the selection is: {Kind: text / url / file / folder / files, Types: [type names], Term: what {query} becomes,
    ; Quoted: Term as one quoted argument (several files: each quoted), Text / Path / Files}
    ; userTypes: UserTypes() of the saved settings when not given
    static Classify(selection, userTypes := "") {
        if !IsObject(userTypes)
            userTypes := ContentTypes.UserTypes()
        if selection.HasOwnProp("Files") {
            files := selection.Files
            if (files.Length = 1)
                return ContentTypes._PathInfo(files[1], userTypes)
            types := ["Files"]
            for userType in userTypes
                if (userType.Kind = "File" && ContentTypes._AllMatch(userType, files))
                    types.Push(userType.Name)
            quoted := ""
            for filePath in files
                quoted .= (quoted = "" ? "" : " ") SelectionActions.Quote(filePath)
            return {Kind: "files", Types: types, Term: quoted, Quoted: quoted, Files: files}
        }
        text := selection.Text
        trimmed := Trim(text, " `t`r`n")
        if TextTools.Match(trimmed, "i)^https?://\S+$")
            return {Kind: "url", Types: ContentTypes._TextTypes(trimmed, "Url", userTypes), Term: trimmed, Quoted: SelectionActions.Quote(trimmed), Text: trimmed}
        if (!InStr(trimmed, "`n") && TextTools.Match(trimmed, "^(?:[A-Za-z]:\\|\\\\)") && FileExist(trimmed))   ; the selection is one path
            return ContentTypes._PathInfo(trimmed, userTypes)
        return {Kind: "text", Types: ContentTypes._TextTypes(trimmed, "Text", userTypes), Term: trimmed, Quoted: SelectionActions.Quote(trimmed), Text: text}
    }

    static _PathInfo(filePath, userTypes) {
        isFolder := (InStr(FileExist(filePath), "D") || SubStr(filePath, -1) = "\") ? true : false
        types := [isFolder ? "Folder" : "File"]
        if !isFolder
            for userType in userTypes
                if (userType.Kind = "File" && ContentTypes._ExtMatches(userType, filePath))
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
                if (TextTools.Match(text, userType.Pattern) > 0)
                    types.Push(userType.Name)
            } catch as err {
                if !badPatterns.Has(userType.Pattern) {
                    badPatterns[userType.Pattern] := true
                    Logger.Warn("Content type " userType.Name ": bad pattern " userType.Pattern " (" err.Message ")")
                }
            }
        }
        return types
    }

    static _AllMatch(userType, files) {
        for filePath in files
            if !ContentTypes._ExtMatches(userType, filePath)
                return false
        return true
    }

    ; a file (not a folder) whose extension is one of the type's ("pdf, .docx" - separated by commas, semicolons, | or spaces)
    static _ExtMatches(userType, filePath) {
        SplitPath(filePath, , , &ext)
        return ContentTypes._HasExt(userType, ext) && !InStr(FileExist(filePath), "D")   ; the disk is asked only when the extension matches
    }

    ; The user's content types from ALTRun.json "ContentTypes": [{Name, Kind: Text | File, Pattern, OpenCommand}] (OpenCommand: a
    ; command Id, File types only, "" when not set); entries without a name or pattern,
    ; or named like a built-in type, are left out. list: the entries to read (the copy being edited in Preferences), the saved ones when not given
    static UserTypes(list := "") {
        types := []
        if !IsObject(list)
            list := (AppSettings.Data.Has("ContentTypes") && AppSettings.Data["ContentTypes"] is Array) ? AppSettings.Data["ContentTypes"] : []
        for entry in list {
            if !(entry is Map)
                continue
            name := Trim(entry.Has("Name") ? entry["Name"] : "")
            pattern := entry.Has("Pattern") ? entry["Pattern"] : ""
            if (name = "" || Trim(pattern) = "" || ContentTypes.IsBuiltInType(name))
                continue
            kind := (entry.Has("Kind") && entry["Kind"] = "File") ? "File" : "Text"
            openCommand := (kind = "File" && entry.Has("OpenCommand") && !IsObject(entry["OpenCommand"])) ? Trim(entry["OpenCommand"]) : ""
            types.Push({Name: name, Kind: kind, Pattern: pattern, OpenCommand: openCommand})
        }
        return types
    }

    static IsBuiltInType(name) {
        for each in ContentTypes.BuiltInTypes
            if (each = name)
                return true
        return false
    }

    ; The type names of a custom command's "ContentTypes" ("PO, PDF", or an array)
    static CommandTypes(command) {
        names := []
        if !(command is Map && command.Has("ContentTypes"))
            return names
        value := command["ContentTypes"]
        for each in (value is Array) ? value : StrSplit(value, ",")
            if (!IsObject(each) && (each := Trim(each)) != "")
                names.Push(each)
        return names
    }

    ; Whether a custom command is shown for a selection of these types (any of its "ContentTypes" names, any case)
    static CommandFor(command, types) {
        for name in ContentTypes.CommandTypes(command)
            for each in types
                if (each = name)
                    return true
        return false
    }

    ; Programs and shortcuts: never opened with a type's open command, so a type can't take over starting programs
    static ProgramTypes := Map("exe", 1, "com", 1, "lnk", 1, "url", 1, "msc", 1, "cpl", 1, "scr", 1, "pif", 1, "appref-ms", 1)

    ; Commands a type can open its files with: custom commands that run a program and have an Id: a Command, or a File whose target is a
    ; program, script or shortcut (or a bare name such as notepad); not a File that opens a document (Budget.xlsx), a folder or a web page
    static IsProgramCommand(command) {
        if !(command is Map && command.Has("Id") && command["Id"] != "")
            return false
        commandType := command.Has("Type") ? command["Type"] : "File"
        if (commandType = "Command")
            return true
        target := Trim(command.Has("Target") ? command["Target"] : "", '" ')
        SplitPath(target, &name, , &ext)
        return commandType = "File" && (ext = "" || TextTools.Match(ext, "i)^(exe|com|bat|cmd|ps1|lnk|ahk|py|pyw|vbs|js)$"))
    }

    ; The custom command that opens a file of this selection (Classify's result) with Enter: the OpenCommand of the first of the
    ; user's File types it has, when that command exists and runs a program; "" otherwise, for programs and shortcuts, or with
    ; custom commands turned off. userTypes / commands: the saved ones when not given
    static OpenCommand(info, userTypes := "", commands := "") {
        if (info.Kind != "file" || !ProviderRegistry.IsEnabled(CustomCommandProvider) || ContentTypes._IsProgram(info.Path))
            return ""
        for userType in IsObject(userTypes) ? userTypes : ContentTypes.UserTypes() {
            if (userType.OpenCommand = "" || !ContentTypes.CommandFor(Map("ContentTypes", userType.Name), info.Types))
                continue
            for command in IsObject(commands) ? commands : AppSettings.CustomCommands
                if (ContentTypes.IsProgramCommand(command) && command["Id"] = userType.OpenCommand)
                    return command
            return ""                                                       ; the first type with an open command decides, even when the command is gone
        }
        return ""
    }

    static _IsProgram(filePath) {
        SplitPath(filePath, , , &ext)
        return ContentTypes.ProgramTypes.Has(StrLower(ext))
    }

    ; A file or folder in the search results (search, pinned, recent, browsing) -> Classify's result for its path, as if it were selected;
    ; "" for other results, programs started with parameters (a custom command), results with their own action, and paths that aren't
    ; there (a path on a network drive counts as there without asking the network: Path.ExistsOrRemote)
    static ForResult(item) {
        if (!IsObject(item) || !(item.Kind = "file" || item.Kind = "folder") || item.Provider = "Selection" || IsObject(item.OnRun)
            || item.Arguments != "" || item.WindowArgs || !TextTools.Match(item.Arg, "^(?:[A-Za-z]:\\|\\\\)") || !Path.ExistsOrRemote(item.Arg))
            return ""
        return ContentTypes.Classify({Files: [item.Arg]})
    }

    ; Whether one of a File type's extensions is ext (the name only, no disk)
    static _HasExt(userType, ext) {
        if (ext = "" || userType.Kind != "File")
            return false
        for each in StrSplit(userType.Pattern, [",", ";", "|", " ", "`t"])
            if ((each := LTrim(Trim(each), ".")) != "" && each = ext)
                return true
        return false
    }

    ; The types of a sample in the test box: a path (C:\..., \\server\...) counts as a selected file without looking at the disk
    ; (a folder when it ends with \ or exists as one), anything else as selected text; list = the ContentTypes being edited
    static SampleTypes(text, list) {
        userTypes := ContentTypes.UserTypes(list)
        trimmed := Trim(text, " `t`r`n")
        if (trimmed = "")
            return []
        if (!InStr(trimmed, "`n") && TextTools.Match(trimmed, "^(?:[A-Za-z]:\\|\\\\)"))
            return ContentTypes._PathInfo(trimmed, userTypes).Types
        return ContentTypes.Classify({Text: text}, userTypes).Types
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
}
