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
; {hwnd} / {pid} become the handle / process id of the window in front before ALTRun was shown, when the command runs
;   (e.g. Arguments "C:\Tools\WindowInfo.ahk {hwnd}"; see ActionCatalog.ForWindow)
; Searched: name, keyword, pinyin initials of the name; for File / Folder also the target's file name
; (without extension) or folder name, e.g. Target "D:\Projects\2026 - Annual Report" is found by "annual" too.
; At the same match level, name matches come before file / folder name matches.
;
; Usage:
;   CustomCommandProvider.AddFromPaths(paths)            "Send to" menu: one opens the edit dialog, several are added directly
;   CustomCommandProvider.Edit(command [, prefill])      opens the edit dialog; command = "" creates a new one
;   CustomCommandProvider.EditDialog(owner, title, item, settings)   the edit dialog with its "Triggers" line and "Set Hotkey..."
;       (Preferences and Edit use it; the hotkeys set there are saved with the command on OK)
;   CustomCommandProvider.HotkeyActions(command) / SetHotkey(command, entry)   the actions panel's "Set Hotkey..." / "Change Hotkey ..."
;   CustomCommandProvider.RunFromHotkey(id)              a custom hotkey with Action "custom:<Id>" (see App._HotkeyAction)
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

    ; "custom:<Id>": the learned ranking and recent items stay with the command when its name or target changes (a command
    ; without an Id yet, added since loading, gets one when saving: until then the Uid of versions before 7)
    static _Uid(command) {
        if (command.Has("Id") && command["Id"] != "")
            return "custom:" command["Id"]
        return Knowledge.LegacyUid("custom", command)
    }

    ; The command with this Id ("custom:<Id>" in a hotkey), "" if it no longer exists
    static FindById(id) => AppSettings.ItemById(AppSettings.CustomCommands, id)

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

    ; The text for {query} in a URL: encoded; right after "mailto:" an address keeps its @ (mail programs don't all decode %40 there)
    static UrlTerm(target, term) {
        encoded := Url.Encode(term)
        return RegExMatch(target, "i)^mailto:\{query\}") ? StrReplace(encoded, "%40", "@") : encoded
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
            target := StrReplace(target, "{query}", (commandType = "Url") ? CustomCommandProvider.UrlTerm(target, term) : term)
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
            Uid: CustomCommandProvider._Uid(command), WindowArgs: ActionCatalog.UsesWindow(target, arguments)
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
              , ItemEditor.Field("Keyword", "Prefs.Col.Keyword", "text", false, "", hint("Keyword"))
              , CustomCommandProvider._SelectionField(hint("Selection"))]
    }

    ; "Selection types": the content types the command comes first for (Ctrl+Alt+\), picked from a menu; saved only when set
    static _SelectionField(hintText) {
        field := ItemEditor.Field("Selection", "Prefs.Col.Selection", "tags", false, () => SelectionActions.TypeNamesForEditor(), hintText)
        field.ButtonLabel := I18n.T("Cmd.Field.Selection.Button")
        field.OmitEmpty := true
        field.Check := (value, *) => SelectionActions.UnknownTypesWarning(value)
        return field
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
        if (InStr(raw, "{query}") || ActionCatalog.UsesWindow(raw))         ; the path contains an argument, known only once typed or run
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

    ; Placeholders in the target / arguments that won't work (a read-only check: the Commands page's Status column), in this order:
    ;   {Kind: "Selection"}        {selection} is written: a command gets the selected text through {query}
    ;   {Kind: "Query"}            {query} gets no text: no keyword, no content type and no hotkey runs the command
    ;   {Kind: "Unknown", Name}    any other {word}, which stays as written; only {query} {hwnd} {pid} are replaced (any case).
    ;                              GUIDs ("shell:::{20D04FE0-...}") have hyphens, so they aren't taken for placeholders
    ; hotkeys: the custom hotkeys (the ones with Action "custom:<the command's Id>" count). Returns [] when all is well
    static PlaceholderProblems(command, hotkeys) {
        problems := [], unknown := [], seen := Map(), selection := false
        seen.CaseSense := false
        text := (command.Has("Target") ? command["Target"] : "") " " (command.Has("Arguments") ? command["Arguments"] : "")
        pos := 1
        while (pos := RegExMatch(text, "\{([A-Za-z_]\w*)\}", &m, pos)) {
            pos += m.Len
            name := m[1]
            if seen.Has(name)
                continue
            seen[name] := true
            if (name = "selection")
                selection := true
            else if !(name = "query" || name = "hwnd" || name = "pid")
                unknown.Push({Kind: "Unknown", Name: m[0]})
        }
        if selection
            problems.Push({Kind: "Selection"})
        if (CustomCommandProvider.TakesQuery(command) && !CustomCommandProvider._CanTakeQuery(command, hotkeys))
            problems.Push({Kind: "Query"})
        for each in unknown
            problems.Push(each)
        return problems
    }

    ; A keyword ("ping 10.0.0.1"), content types (the selection panel) or a hotkey (the selection) fills in {query}
    static _CanTakeQuery(command, hotkeys) {
        if (SearchQuery.Keywords(command.Has("Keyword") ? command["Keyword"] : "").Length || SelectionActions.CommandTypes(command).Length)
            return true
        for entry in CustomCommandProvider.HotkeysFor(command, hotkeys)
            if (entry.Has("Key") && LTrim(entry["Key"], "~") != "")
                return true
        return false
    }

    ; The Status column's short text for the first problem of PlaceholderProblems ("" = none)
    static PlaceholderStatus(command, hotkeys) {
        problems := CustomCommandProvider.PlaceholderProblems(command, hotkeys)
        if !problems.Length
            return ""
        switch problems[1].Kind {
            case "Selection": return I18n.T("Prefs.Status.UseQuery")
            case "Query"    : return I18n.T("Prefs.Status.QueryEmpty")
            default         : return I18n.T("Prefs.Status.Unknown", problems[1].Name)
        }
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
        return Map("Title", "", "Type", "File", "Target", "", "Arguments", "", "Keyword", "")   ; the Id is given when the edit dialog opens (EditDialog) or when saving (AppSettings.EnsureIds)
    }

    ; command: an entry of AppSettings.CustomCommands (changed in place), or "" for a new one (prefill = the prefilled fields).
    ; Hotkeys set in the dialog are saved with it; when they changed, ALTRun restarts so they work (only the command changed: no restart)
    static Edit(command := "", prefill := "") {
        isNew := !IsObject(command)
        base := isNew ? CustomCommandProvider.NewCommand() : ItemEditor.WithDefaults(command, CustomCommandProvider.NewCommand())
        if IsObject(prefill)
            for key, value in prefill
                base[key] := value
        if !(AppSettings.Data.Has("Hotkeys") && AppSettings.Hotkeys is Array)
            AppSettings.Data["Hotkeys"] := []
        result := CustomCommandProvider.EditDialog(ItemEditor.Owner(), I18n.T("Prefs.Page.Commands"), base
            , {Commands: AppSettings.CustomCommands, Hotkeys: AppSettings.Hotkeys, Value: (path) => PreferencesWindow.GetPath(AppSettings.Data, path)})
        if !IsObject(result)
            return false
        if isNew
            AppSettings.CustomCommands.Push(result.Command)
        else
            for key, value in result.Command
                command[key] := value
        CustomCommandProvider._ResetNarrowing()
        if !AppSettings.Save()
            return false
        if result.HotkeysChanged                                            ; hotkeys are registered at start-up (as after Apply in Preferences)
            App.Restart("-Reloaded")
        return true
    }

    ; The edit dialog with a read-only "Triggers" line (keyword, content types, hotkeys) and "Set Hotkey...": with no hotkey it opens the
    ; hotkey edit window for a new one, otherwise a menu (edit / delete each, add another). The command gets its Id when the dialog opens,
    ; so a new one can have a hotkey before it is saved; the hotkeys are changed in a draft, written to settings.Hotkeys only on OK.
    ; settings: {Commands, Hotkeys, Value(path)}: the commands and hotkeys being edited (Preferences: Working; otherwise AppSettings),
    ; Value reads the other hotkey settings for the clash check.
    ; Returns {Command: the edited Map, HotkeysChanged}, "" when cancelled
    static EditDialog(owner, title, item, settings) {
        item := item.Clone()
        if !(item.Has("Id") && AppSettings.IsId(item["Id"])) {
            taken := Map()
            for each in settings.Commands
                if (each is Map && each.Has("Id"))
                    taken[each["Id"] ""] := true
            item["Id"] := AppSettings.NewId(taken)
        }
        id := item["Id"], action := "custom:" id
        draft := []                                                         ; {Entry: the hotkey as edited, Original: its Map in settings.Hotkeys ("" = new)}
        for entry in CustomCommandProvider.HotkeysFor(item, settings.Hotkeys)
            draft.Push({Entry: entry.Clone(), Original: entry})

        fields := CustomCommandProvider.EditorFields()
        triggers := ItemEditor.Field("Triggers", "Prefs.Col.Triggers", "info", false, "", I18n.T("Cmd.Field.Triggers"))
        triggers.Text := (edited) => CustomCommandProvider.TriggerText(edited, Entries())
        triggers.Button := {Label: I18n.T("Cmd.SetHotkey"), OnClick: SetHotkey}
        fields.Push(triggers)
        edited := ItemEditor.Edit(owner, title, fields, item)
        if !IsObject(edited)
            return ""
        return {Command: edited, HotkeysChanged: CustomCommandProvider.MergeHotkeys(settings.Hotkeys, action, draft)}

        Entries() {
            list := []
            for each in draft
                list.Push(each.Entry)
            return list
        }
        SetHotkey(edited, g, button, refresh) {
            if !draft.Length
                return EditHotkey(edited, g, 0)
            m := Menu()                                                     ; an item runs after the menu closes: it updates the line itself
            for index, each in draft
                m.Add(I18n.T("Cmd.Hotkey.Edit", HotkeyBox.Label(LTrim(each.Entry["Key"], "~"))), ((index, *) => (EditHotkey(edited, g, index), refresh())).Bind(index))
            m.Add(I18n.T("Cmd.Hotkey.Add"), (*) => (EditHotkey(edited, g, 0), refresh()))
            m.Add()
            for index, each in draft
                m.Add(I18n.T("Cmd.Hotkey.Delete", HotkeyBox.Label(LTrim(each.Entry["Key"], "~"))), ((index, *) => (draft.RemoveAt(index), refresh())).Bind(index))
            ControlGetPos(&x, &y, , &h, button), WinGetClientPos(&left, &top, , , g)   ; pixels, whatever the scaling
            CoordMode("Menu", "Screen")
            m.Show(left + x, top + y + h)                                   ; below the button
        }
        EditHotkey(edited, g, index) {                                      ; index 0 = a new hotkey
            entry := index ? draft[index].Entry : Map("Key", "", "Action", action, "SearchText", "", "WinTitle", "")
            commands := [edited]                                            ; the command as the dialog shows it now (its name in the hotkey window)
            for each in settings.Commands
                if !(each is Map && each.Has("Id") && each["Id"] == id)
                    commands.Push(each)
            hotkeys := []                                                   ; the others as they will be: this command's from the draft
            for each in settings.Hotkeys
                if !(each is Map && each.Has("Action") && each["Action"] == action)
                    hotkeys.Push(each)
            for each in draft
                hotkeys.Push(each.Entry)
            inUse := PreferencesWindow.AllHotkeys(Map("CustomCommands", commands, "Hotkeys", hotkeys), settings.Value)
            result := HotkeyEditor.Edit(g, entry, index ? entry : "", commands, inUse, true)
            if !IsObject(result)
                return
            if index
                draft[index].Entry := result
            else
                draft.Push({Entry: result, Original: ""})
        }
    }

    ; The actions panel's hotkey items of a command -> [{Title, Entry}]: "Set Hotkey..." (Entry "") without one, otherwise
    ; "Change Hotkey Ctrl+Alt+N..." for each (another one is added in the edit window, F3)
    static HotkeyActions(command, hotkeys := "") {
        actions := []
        for entry in CustomCommandProvider.HotkeysFor(command, (hotkeys is Array) ? hotkeys : CustomCommandProvider._SavedHotkeys())
            if (entry.Has("Key") && LTrim(entry["Key"], "~") != "")
                actions.Push({Title: I18n.T("Action.ChangeHotkey", HotkeyBox.Label(LTrim(entry["Key"], "~"))), Entry: entry})
        if !actions.Length
            actions.Push({Title: I18n.T("Action.SetHotkey"), Entry: ""})
        return actions
    }

    static _SavedHotkeys() => (AppSettings.Data.Has("Hotkeys") && AppSettings.Hotkeys is Array) ? AppSettings.Hotkeys : []

    ; The hotkey window for a saved command (entry: one of its hotkeys, "" = a new one); on OK saved, and ALTRun restarts when it
    ; changed, so the hotkey works (as Edit does)
    static SetHotkey(command, entry := "") {
        if !(AppSettings.Data.Has("Hotkeys") && AppSettings.Hotkeys is Array)
            AppSettings.Data["Hotkeys"] := []
        if !(command.Has("Id") && AppSettings.IsId(command["Id"]))       ; saved commands have one (EnsureIds at loading)
            AppSettings.EnsureIds(AppSettings.Data)
        action := "custom:" command["Id"]
        item := IsObject(entry) ? entry : Map("Key", "", "Action", action, "SearchText", "", "WinTitle", "")
        result := HotkeyEditor.Edit(ItemEditor.Owner(), item, entry, AppSettings.CustomCommands, "", true)
        if !IsObject(result)
            return false
        draft := []                                                         ; the command's hotkeys with this one changed or added
        for each in CustomCommandProvider.HotkeysFor(command, AppSettings.Hotkeys)
            draft.Push({Entry: (IsObject(entry) && ObjPtr(each) = ObjPtr(entry)) ? result : each.Clone(), Original: each})
        if !IsObject(entry)
            draft.Push({Entry: result, Original: ""})
        changed := CustomCommandProvider.MergeHotkeys(AppSettings.Hotkeys, action, draft)
        if !changed
            return true
        if !AppSettings.Save()
            return false
        App.Restart("-Reloaded")                                            ; hotkeys are registered at start-up
        return true
    }

    ; The "Triggers" line: "Keyword ping · Selection IP · Hotkey Ctrl+Alt+P"; "Only found by name in the search" without any
    static TriggerText(command, hotkeys) {
        parts := []
        keywords := ""
        for each in SearchQuery.Keywords(command.Has("Keyword") ? command["Keyword"] : "")
            keywords .= (keywords = "" ? "" : ", ") each
        if (keywords != "")
            parts.Push(I18n.T("Cmd.Triggers.Keyword", keywords))
        types := ItemEditor.TagsText(command.Has("Selection") ? command["Selection"] : "")
        if (types != "")
            parts.Push(I18n.T("Cmd.Triggers.Selection", types))
        keys := ""
        for entry in hotkeys
            if (entry is Map && entry.Has("Key") && LTrim(entry["Key"], "~") != "")
                keys .= (keys = "" ? "" : ", ") HotkeyBox.Label(LTrim(entry["Key"], "~"))
        if (keys != "")
            parts.Push(I18n.T("Cmd.Triggers.Hotkey", keys))
        if !parts.Length
            return I18n.T("Cmd.Triggers.None")
        text := ""
        for part in parts
            text .= (text = "" ? "" : " · ") part
        return text
    }

    ; Writes the hotkeys of the edit dialog's draft ({Entry, Original}...) to hotkeys (in place: the list Preferences shows): the ones
    ; edited change in place, the ones deleted go, new ones are added at the end. Returns whether the hotkeys running action changed
    static MergeHotkeys(hotkeys, action, draft) {
        before := "", after := "", kept := Map()
        for each in draft
            if IsObject(each.Original)
                kept[ObjPtr(each.Original)] := true
        index := hotkeys.Length
        while (index >= 1) {
            entry := hotkeys[index]
            if (entry is Map && entry.Has("Action") && entry["Action"] == action) {
                before := JSON.Stringify(entry) "|" before
                if !kept.Has(ObjPtr(entry))
                    hotkeys.RemoveAt(index)
            }
            index--
        }
        for each in draft {
            if IsObject(each.Original) {
                each.Original.Clear()
                for key, value in each.Entry
                    each.Original[key] := value
            } else
                hotkeys.Push(each.Entry)
        }
        for entry in hotkeys
            if (entry is Map && entry.Has("Action") && entry["Action"] == action)
                after .= JSON.Stringify(entry) "|"
        return before != after
    }

    ; A custom hotkey whose Action is "custom:<Id>". A command that takes something (it has {query}, or content types) gets the
    ; selection as from the selection panel (SelectionActions._CommandAction): {query} becomes the selected text or path; a File /
    ; Command without {query} gets it as its last argument. The panel's checks keep selected text out of a shell; the content types
    ; aren't checked (the hotkey was set for this command on purpose). Nothing selected: a {query} command with a keyword opens the
    ; search window with "keyword " to type the rest, one without a keyword says so, one with only content types runs as it is.
    ; Any other command runs directly, without copying the selection (the clipboard isn't touched).
    ; capture: replaces SelectionActions.Capture in the tests
    static RunFromHotkey(id, capture := "") {
        command := CustomCommandProvider.FindById(id)
        if !IsObject(command) {
            Logger.Warn("CustomCommandProvider: the hotkey's command custom:" id " no longer exists")
            return App.Notify(I18n.T("Hotkey.CommandMissing"), 3000)
        }
        takesQuery := CustomCommandProvider.TakesQuery(command), item := ""
        if (takesQuery || SelectionActions.CommandTypes(command).Length) {
            selection := IsObject(capture) ? capture.Call() : SelectionActions.Capture()
            if IsObject(selection) {
                item := SelectionActions._CommandAction(command, SelectionActions.Classify(selection))
                if !IsObject(item)
                    return App.Notify(I18n.T("Hotkey.SelectionRefused", command["Title"]), 3000)
            } else if takesQuery {
                keywords := SearchQuery.Keywords(command.Has("Keyword") ? command["Keyword"] : "")
                return keywords.Length ? SearchWindow.Show(keywords[1] " ") : App.Notify(I18n.T("Selection.None"), 2000)
            }
        }
        if !IsObject(item)
            item := CustomCommandProvider._ToItem(command, 0)
        Usage.Count(CustomCommandProvider.Id)
        DllCall("AllowSetForegroundWindow", "UInt", 0xFFFFFFFF)            ; ASFW_ANY: the program started may come to the front
        ActionCatalog.RunDefault(item)
    }

    static EditItem(item) {
        return CustomCommandProvider.Edit(item.Source)
    }

    ; Ctrl+Del in the results: the hotkeys that run the command are deleted with it (the question says so) and stop at once
    static DeleteItem(item) {
        for index, command in AppSettings.CustomCommands {
            if (ObjPtr(command) = ObjPtr(item.Source)) {
                AppSettings.CustomCommands.RemoveAt(index)
                for entry in CustomCommandProvider.RemoveHotkeys(command, AppSettings.Hotkeys)
                    App.UnregisterHotkey(entry)
                CustomCommandProvider._ResetNarrowing()
                return AppSettings.Save()
            }
        }
        return false
    }

    static DeletePrompt(item) {
        return CustomCommandProvider.DeleteQuestion(item.Source, AppSettings.Hotkeys) || I18n.T("Search.ConfirmDelete", item.Title)
    }

    ; The custom hotkeys (entries of hotkeys) whose Action runs this command
    static HotkeysFor(command, hotkeys) {
        found := []
        if !(command is Map && command.Has("Id") && command["Id"] != "" && hotkeys is Array)
            return found
        for entry in hotkeys
            if (entry is Map && entry.Has("Action") && entry["Action"] == "custom:" command["Id"])
                found.Push(entry)
        return found
    }

    ; "Delete 'Notepad'? Hotkeys that run it are deleted too: Ctrl+Alt+N", "" when no hotkey runs it (the usual question then)
    static DeleteQuestion(command, hotkeys) {
        keys := ""
        for entry in CustomCommandProvider.HotkeysFor(command, hotkeys)
            keys .= (keys = "" ? "" : ", ") HotkeyBox.Label(entry.Has("Key") ? entry["Key"] : "")
        return (keys = "") ? "" : I18n.T("Custom.ConfirmDeleteHotkeys", command["Title"], keys)
    }

    ; Removes the hotkeys that run this command from hotkeys (in place: the list Preferences shows) and returns them
    static RemoveHotkeys(command, hotkeys) {
        removed := CustomCommandProvider.HotkeysFor(command, hotkeys)
        for entry in removed
            for index, existing in hotkeys
                if (ObjPtr(existing) = ObjPtr(entry)) {
                    hotkeys.RemoveAt(index)
                    break
                }
        return removed
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
