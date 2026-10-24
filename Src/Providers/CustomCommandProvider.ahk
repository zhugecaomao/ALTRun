;===============================================================================
; CustomCommandProvider.ahk - the user's custom commands (AutoHotkey v2)
;-------------------------------------------------------------------------------
; ALTRun.json -> CustomCommands, each entry:
;   { "Title": "display name", "Type": "File|Folder|Command|Url",
;     "Target": "path / program / URL", "Arguments": "command-line arguments", "Keyword": "optional keyword" }
; Target can start with a built-in variable such as A_Desktop / A_ScriptDir, or use environment variables such as %AppData%.
; An exactly matching Keyword comes first; several can be given, separated by commas ("np, ㅜㅔ"), and any of them works.
; Target / Arguments can contain {query}: typing "keyword text" replaces it with the text (encoded in URLs), e.g.
;   { "Title": "Jira", "Type": "Url", "Target": "https://jira.example.com/browse/{query}", "Keyword": "jira" }
;   typing "jira ABC-123" opens .../browse/ABC-123; when the command is found but no text typed yet, Enter completes to "jira "
; Searched: name, keyword, pinyin initials of the name; for File / Folder also the target's file name
; (without extension) or folder name, e.g. Target "D:\Projects\2026 - Annual Report" is found by "annual" too.
; At the same match level, name matches come before file / folder name matches.
;
; Usage:
;   CustomCommandProvider.AddFromPaths(paths)            "Send to" menu: one opens the edit dialog, several are added directly
;   CustomCommandProvider.Edit(command [, prefill])      opens the edit dialog; command = "" creates a new one
;   F3 / right-click "Edit" in the results calls EditItem(), Ctrl+Del calls DeleteItem()
;===============================================================================

class CustomCommandProvider {
    static Id := "CustomCommands"
    static _lastNeedle := "", _lastMatches := [], _lastList := ""

    static Init() {
    }

    ; All commands are scored first, and results are only made for the top MaxResults (with many commands one letter can
    ; match hundreds, and making a result for each would make typing lag). The learned bonus counts when taking the top ones, so often chosen commands aren't pushed out.
    ; When typing on (the new input starts with the previous one, at least 3 letters) only the previous matches are searched,
    ; as in ApplicationProvider; after commands are added, changed or deleted, all commands are searched again
    static Search(query) {
        needle := StrLower(query.Text)
        commands := AppSettings.CustomCommands
        list := ObjPtr(commands) "|" commands.Length
        candidates := commands
        if (list = CustomCommandProvider._lastList && StrLen(CustomCommandProvider._lastNeedle) >= 3
            && InStr(needle, CustomCommandProvider._lastNeedle) = 1) {
            candidates := Map()
            for index in CustomCommandProvider._lastMatches
                candidates[index] := commands[index]
        }
        scores := Map(), ranks := Map(), matches := [], withQuery := []
        for index, command in candidates {
            if !(command is Map) || !command.Has("Title") || !command.Has("Target")
                continue
            keyword := command.Has("Keyword") ? command["Keyword"] : ""
            exact := (keyword != "") && query.MatchKeyword(SearchQuery.Keywords(keyword), &rest, &typed)
            if (exact && query.HasRest && CustomCommandProvider.TakesQuery(command)) {
                item := CustomCommandProvider._ToItem(command, 150, rest, typed)  ; "jira ABC-123": only commands that take arguments
                item.Exclusive := true
                withQuery.Push(item)
                continue
            }
            commandType := command.Has("Type") ? command["Type"] : "File"
            score := FuzzyMatcher.BestKey(needle, CustomCommandProvider._KeysFor(command["Title"], keyword))
            targetScore := FuzzyMatcher.BestKey(needle, CustomCommandProvider._TargetKeysFor(commandType, command["Target"]))
            if (targetScore > 0)
                score := Max(score, targetScore - 5)                        ; the name first
            if exact
                score := Max(score, 100)
            if (score <= 0)
                continue
            scores[index] := score + 15                                     ; the user's own commands rank above indexed applications
            ranks[index] := scores[index] + Knowledge.Boost(query.Text, CustomCommandProvider._Uid(command))
            matches.Push(index)
        }
        CustomCommandProvider._lastNeedle := needle, CustomCommandProvider._lastMatches := matches, CustomCommandProvider._lastList := list
        if withQuery.Length {
            CustomCommandProvider._lastNeedle := ""                          ; the arguments differ every time: don't narrow down with the previous matches
            return withQuery
        }
        results := []
        for index in FuzzyMatcher.TopIndexes(ranks, ProviderRegistry.MaxResults)
            results.Push(CustomCommandProvider._ToItem(commands[index], scores[index]))
        return results
    }

    ; While idle after start-up, compute each command's search keys (name / keyword / pinyin / target name) and target path in advance,
    ; so the first input doesn't wait (hundreds of commands take hundreds of milliseconds the first time). WarmBatch at a time, so input can be handled in between
    static WarmBatch := 40
    static Warm(start := 1) {
        commands := AppSettings.CustomCommands
        last := Min(commands.Length, start + CustomCommandProvider.WarmBatch - 1)
        Loop Max(0, last - start + 1) {
            command := commands[start + A_Index - 1]
            if !(command is Map) || !command.Has("Title") || !command.Has("Target")
                continue
            commandType := command.Has("Type") ? command["Type"] : "File"
            CustomCommandProvider._KeysFor(command["Title"], command.Has("Keyword") ? command["Keyword"] : "")
            CustomCommandProvider._TargetKeysFor(commandType, command["Target"])
            if (commandType = "Folder")
                IconCache.FolderIcon(CustomCommandProvider._Resolve(command["Target"]), true)
        }
        if (last < commands.Length)
            SetTimer(() => CustomCommandProvider.Warm(last + 1), -10)
    }

    ; after commands changed, the next search looks at all commands
    static _ResetNarrowing() {
        CustomCommandProvider._lastNeedle := ""
    }

    ; For the icon prefetch after start-up (see ProviderRegistry.IconSpecs)
    static IconItems() {
        items := []
        for command in AppSettings.CustomCommands {
            try {
                item := CustomCommandProvider._ToItem(command, 0)
                items.Push({Icon: item.Icon, Uid: item.Uid})
            }
        }
        return items
    }

    static _Uid(command) {
        arguments := command.Has("Arguments") ? command["Arguments"] : ""
        commandType := command.Has("Type") ? command["Type"] : "File"
        return "custom:" StrLower(commandType "|" command["Target"] "|" arguments)
    }

    ; Search keys of the name / each keyword / pinyin initials, cached by "name|keyword" (a changed command gets new keys naturally)
    static _KeysFor(title, keyword) {
        static cache := Map()
        cacheKey := title "|" keyword
        if !cache.Has(cacheKey) {
            pinyinText := Pinyin.Initials(title)
            keys := [FuzzyMatcher.Key(title), (pinyinText != title) ? FuzzyMatcher.Key(pinyinText) : ""]
            for each in SearchQuery.Keywords(keyword)
                keys.Push(FuzzyMatcher.Key(each))
            cache[cacheKey] := keys
        }
        return cache[cacheKey]
    }

    ; The target's file name (without extension) or folder name of File / Folder and its pinyin initials, cached by "type|target"
    static _TargetKeysFor(commandType, target) {
        static cache := Map()
        cacheKey := commandType "|" target
        if !cache.Has(cacheKey)
            cache[cacheKey] := CustomCommandProvider._TargetKeys(commandType, target)
        return cache[cacheKey]
    }

    static _TargetKeys(commandType, target) {
        if !(commandType = "Folder" || commandType = "File")
            return []
        SplitPath(RTrim(CustomCommandProvider._Resolve(target), "\/"), &fileName, , , &nameNoExt)
        name := (commandType = "File") ? nameNoExt : fileName
        if (name = "" || InStr(name, ":"))                                  ; a drive root "Q:" isn't a name
            return []
        pinyinText := Pinyin.Initials(name)
        return [FuzzyMatcher.Key(name), (pinyinText != name) ? FuzzyMatcher.Key(pinyinText) : ""]
    }

    ; Target / Arguments contain {query}
    static TakesQuery(command) {
        return InStr(command["Target"], "{query}") || (command.Has("Arguments") && InStr(command["Arguments"], "{query}"))
    }

    ; term: the typed argument, replacing {query} (encoded in URLs); a {query} command without an argument yet: with a keyword Enter completes "keyword ", otherwise it becomes empty
    ; keyword: the keyword that was typed (completion keeps it when there are several), the first if none
    static _ToItem(command, score, term := "", keyword := "") {
        target := command["Target"]
        arguments := command.Has("Arguments") ? command["Arguments"] : ""
        commandType := command.Has("Type") ? command["Type"] : "File"
        if (keyword = "") {
            keywords := SearchQuery.Keywords(command.Has("Keyword") ? command["Keyword"] : "")
            keyword := keywords.Length ? keywords[1] : ""
        }
        title := command["Title"], extra := {}
        if CustomCommandProvider.TakesQuery(command) {
            if (term = "" && keyword != "")
                extra := {Valid: false, AutoComplete: keyword " "}
            else if (term != "")
                title .= ": " term, extra := {AutoComplete: keyword " " term}
            target := StrReplace(target, "{query}", (commandType = "Url") ? Url.Encode(term) : term)
            arguments := StrReplace(arguments, "{query}", term)
        }
        switch commandType, false {
            case "Folder": kind := "folder", icon := CustomCommandProvider._FolderIcon(CustomCommandProvider._Resolve(target))
            case "Url"   : kind := "url",    icon := RegExMatch(target, "i)^ms-settings:") ? "res:imageres.dll,-114" : "url:"   ; system settings links use the gear icon
            default      : kind := "file",   icon := CustomCommandProvider._Resolve(target)
        }
        displayTarget := (kind = "url") ? target : CustomCommandProvider._Resolve(target)
        subtitle := Trim(displayTarget " " arguments)
        if (extra.HasOwnProp("Valid") && !extra.Valid)                     ; waiting for an argument: say what to type next
            subtitle := I18n.T("Search.TypeAfter", keyword) " · " subtitle
        item := ResultItem(title, subtitle, {
            Kind: kind, Arg: target, Arguments: arguments, Icon: icon, Score: score, Source: command,
            Uid: CustomCommandProvider._Uid(command)
        })
        for name, value in extra.OwnProps()
            item.%name% := value
        return item
    }

    ; folders on network locations use the generic folder icon directly (no network access, and a dot in the name doesn't make it a file)
    static _FolderIcon(folder) {
        return IconCache.FolderIcon(folder)
    }

    ; Path.Resolve needs the disk and PATH for targets given only as a program name ("cmd.exe"): cache the result
    static _Resolve(target) {
        static cache := Map()
        if !cache.Has(target)
            cache[target] := Path.Resolve(target)
        return cache[target]
    }

    static EditorFields() {
        types := [["File", I18n.T("Prefs.Type.File")], ["Folder", I18n.T("Prefs.Type.Folder")]
                , ["Command", I18n.T("Prefs.Type.Command")], ["Url", I18n.T("Prefs.Type.Url")]]
        hint := (name) => I18n.T("Cmd.Field." name)                        ; the gray description below each field
        target := ItemEditor.Field("Target", "Prefs.Col.Target", "file", true, "", hint("Target"))
        target.FolderWhen := ["Type", "Folder"]                             ; with the type set to "Folder", the browse button chooses a folder
        return [ItemEditor.Field("Title", "Prefs.Col.Title", "text", true, "", hint("Title"))
              , ItemEditor.Field("Type", "Prefs.Col.Type", "choice", false, types, hint("Type"))
              , target
              , ItemEditor.Field("Arguments", "Prefs.Col.Arguments", "text", false, "", hint("Arguments"))
              , ItemEditor.Field("Keyword", "Prefs.Col.Keyword", "text", false, "", hint("Keyword"))]
    }

    ; Checks whether a command's target still exists (renamed folders and moved files break commands); returns:
    ;   "OK"          exists
    ;   "Missing"     not found
    ;   "Unavailable" its drive or network location can't be reached now (e.g. a network drive not connected), so it can't be told
    ;   "Skipped"     targets that aren't files (URLs, shell: etc.) aren't checked
    ; roots: a Map shared by one check run that remembers whether each drive / network share is reachable, so a disconnected network drive is waited for once
    static CheckTarget(command, roots := "") {
        commandType := command.Has("Type") ? command["Type"] : "File"
        if (commandType = "Url")
            return "Skipped"
        raw := Trim(command.Has("Target") ? command["Target"] : "", "`" `t")
        if (raw = "")
            return "Missing"
        if InStr(raw, "{query}")                                            ; the path contains an argument, known only once typed
            return "Skipped"
        target := Path.Resolve(raw)
        if RegExMatch(target, "i)^([a-z][a-z0-9+.-]+:|::\{)")               ; shell:, ms-settings:, http: etc. (not "C:")
            return "Skipped"
        if !InStr(target, "\")                                              ; a bare file name not in PATH
            return CustomCommandProvider._InAppPaths(target) ? "OK" : "Missing"
        SplitPath(target, , , , , &drive)
        if (drive != "") {
            if !IsObject(roots)
                roots := Map()
            if !roots.Has(drive)
                roots[drive] := DirExist(drive "\") != ""
            if !roots[drive]
                return "Unavailable"
        }
        found := (commandType = "Folder") ? DirExist(target) : FileExist(target)
        return (found != "") ? "OK" : "Missing"
    }

    ; programs registered under "App Paths" (winword, chrome...), which Run can open directly too
    static _InAppPaths(name) {
        SplitPath(name, , , &ext)
        if (ext = "")
            name .= ".exe"
        for root in ["HKCU", "HKLM"]
            try if (RegRead(root "\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" name, "") != "")
                return true
        return false
    }

    static NewCommand() {
        return Map("Title", "", "Type", "File", "Target", "", "Arguments", "", "Keyword", "")
    }

    ; command: an entry of AppSettings.CustomCommands (changed in place), or "" for a new one (prefill = the prefilled fields)
    static Edit(command := "", prefill := "") {
        isNew := !IsObject(command)
        base := isNew ? CustomCommandProvider.NewCommand() : ItemEditor.WithDefaults(command, CustomCommandProvider.NewCommand())
        if IsObject(prefill)
            for key, value in prefill
                base[key] := value
        edited := ItemEditor.Edit(ItemEditor.Owner(), I18n.T("Prefs.Page.Commands"), CustomCommandProvider.EditorFields(), base)
        if !IsObject(edited)
            return false
        if isNew
            AppSettings.CustomCommands.Push(edited)
        else
            for key, value in edited
                command[key] := value
        CustomCommandProvider._ResetNarrowing()
        return AppSettings.Save()
    }

    static EditItem(item) {
        return CustomCommandProvider.Edit(item.Source)
    }

    static DeleteItem(item) {
        for index, command in AppSettings.CustomCommands {
            if (ObjPtr(command) = ObjPtr(item.Source)) {
                AppSettings.CustomCommands.RemoveAt(index)
                CustomCommandProvider._ResetNarrowing()
                return AppSettings.Save()
            }
        }
        return false
    }

    ; Explorer's "Send to → ALTRun" (all selected paths come in at once):
    ;   one: opens the edit dialog with name / type / target filled in, added only on OK; if the command already exists it opens that for editing
    ;   several: all added directly (a dialog each would be too many), existing ones skipped
    static AddFromPaths(paths) {
        if !paths.Length
            return
        if IsObject(PreferencesWindow.Gui)                                  ; Preferences is open: add to its Commands page, so the next save doesn't overwrite them
            return PreferencesWindow.AddCommands(paths)
        if (paths.Length = 1) {
            existing := CustomCommandProvider.FindByTarget(paths[1])
            if IsObject(existing) {
                App.Notify(I18n.T("Custom.Exists", existing["Title"]), 2500)
                CustomCommandProvider.Edit(existing)
            } else if CustomCommandProvider.Edit("", CustomCommandProvider.FromPath(paths[1])) {
                App.Toast(I18n.T("Custom.Added", AppSettings.CustomCommands[-1]["Title"]), 2500)
            }
            return
        }
        added := 0, skipped := 0
        for target in paths {
            if IsObject(CustomCommandProvider.FindByTarget(target)) {
                skipped += 1
                continue
            }
            AppSettings.CustomCommands.Push(CustomCommandProvider.FromPath(target))
            added += 1
        }
        if added {
            CustomCommandProvider._ResetNarrowing()
            AppSettings.Save()
        }
        App.Toast(I18n.T("Custom.AddedMany", added) (skipped ? " " I18n.T("Custom.SkippedExisting", skipped) : ""), 3000)
    }

    ; File / folder path -> a new command (name = file name without extension, or folder name)
    static FromPath(target) {
        target := RTrim(target, "\/")
        if RegExMatch(target, "^[A-Za-z]:$")                                ; a whole drive "D:"
            target .= "\"
        isFolder := DirExist(target) != ""
        SplitPath(RTrim(target, "\"), &fileName, , , &nameNoExt)
        title := isFolder ? fileName : nameNoExt
        if (title = "")
            title := target
        return Map("Title", title, "Type", isFolder ? "Folder" : "File", "Target", target, "Arguments", "", "Keyword", "")
    }

    ; The command whose target is the same file / folder (comparing paths with variables expanded, case-insensitive), "" if none; commands: the list to search (the one being edited in Preferences)
    static FindByTarget(target, commands := AppSettings.CustomCommands) {
        wanted := StrLower(RTrim(Path.Resolve(target), "\/"))
        for command in commands {
            if !(command is Map) || !command.Has("Target")
                continue
            if (command.Has("Type") && command["Type"] = "Url")
                continue
            if (StrLower(RTrim(CustomCommandProvider._Resolve(command["Target"]), "\/")) = wanted)
                return command
        }
        return ""
    }
}
