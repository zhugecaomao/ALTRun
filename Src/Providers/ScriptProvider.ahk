;===============================================================================
; ScriptProvider.ahk - Scripts extension: scripts in the Scripts\ folder become commands (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like Raycast's Script Commands and RunZ's plugins: put a script in Scripts\ (in the program folder, or
; %APPDATA%\ALTRun\Scripts when that isn't writable), and it can be found by name or keyword and run. .ahk .ps1 .bat .cmd .py are supported.
; The comments at the top of a script can say (any comment marker ; # REM :: // works; all are optional):
;   @altrun.title     name              (the file name by default)
;   @altrun.keyword   keyword           comes first when typed exactly
;   @altrun.argument  hint text         takes an argument: type "keyword text", and the text is passed to the script as its first argument
;   @altrun.mode      window | silent | output | filter
;                     window (default) runs normally with a window; silent runs in the background and shows the last line of output as a notification when done;
;                     output runs in the background and opens all output with the default program for .txt (usually Notepad) when done;
;                     filter gives live results: typing "keyword text" runs it with the text while typing and shows what it prints as
;                     results (needs a keyword; .ahk .py .ps1; see ScriptFilter.ahk)
;   @altrun.filter    altrun            a filter script runs once on entering the keyword; ALTRun then filters what it printed while typing
;   @altrun.selection Text, PDF         content type names (built-in or Preferences → Selection): the script is offered in the actions
;                     panel of such a selection (Ctrl+Alt+\, after the custom commands), which is passed to it as its first argument
;                     (several files: one argument each), with the panel's command-line checks (SelectionActions.SafeArgument)
; Read again when the folder changes (modified time), without reloading ALTRun. F3 edits the script in Notepad.
; A custom hotkey with Action "script:<file name>" runs a script (RunFromHotkey); renaming the file breaks it (shown in red in Preferences).
; A filter script isn't offered for a selection; its hotkey opens the search window with "keyword " and the selected line of text.
;
; Settings (ALTRun.json -> Features.Scripts): Enabled / Keyword ("script backup" searches scripts only, "script " lists all,
; like the snippets' "snip"; just "script" without the space still shows the commands with that word)
;===============================================================================

class ScriptProvider {
    static Id := "Scripts"
    static Extensions := Map("ahk", 1, "ps1", 1, "bat", 1, "cmd", 1, "py", 1)
    static Scripts := []
    static _lastStamp := "", _checked := 0
    static TimeoutMs := 120000                                              ; how long to wait for a background script at most

    static HintIcon := "res:imageres.dll,-5323"                             ; icon of the keyword hint row
    static Folder := ""                                                     ; when not empty, this folder is used (for tests)
    static Dir => (ScriptProvider.Folder != "") ? ScriptProvider.Folder : AppSettings.Portable ? A_ScriptDir "\Scripts" : AppSettings.UserDir "\Scripts"

    static Init() {
        ScriptProvider.Load()
    }

    static Search(query) {
        if (A_TickCount - ScriptProvider._checked > 2000)                   ; check whether the folder changed at most every 2 s
            ScriptProvider._RefreshIfChanged()
        results := [], keyword := AppSettings.Feature("Scripts")["Keyword"], full := ""
        onlyScripts := (keyword != "" && query.HasRest && query.MatchKeyword([keyword], &term))
        needle := StrLower(onlyScripts ? term : query.Text)
        if (!onlyScripts && query.HasRest)                                  ; "keyword text" of a filter script: its live results only
            for script in ScriptProvider.Scripts
                if (script.Mode = "filter" && query.Keyword = script.Keyword && ScriptFilter.Supports(script))
                    return ScriptFilter.Results(script, query.Rest)
        ScriptFilter.Leave()                                                ; not (or no longer) in one: a run still going isn't needed
        if (keyword != "" && !query.HasRest && query.MatchKeyword([keyword], &term))   ; "script" without the space: a hint, like the other keywords (ResultItem.KeywordHintScore); Enter completes "script "
            results.Push(ResultItem.KeywordHint(I18n.T("Scripts.Keyword"), keyword, ScriptProvider.HintIcon))
        else if (keyword != "" && query.StartsKeyword([keyword], &full))       ; "scrip": the same hint at the end of the list
            results.Push(ResultItem.KeywordHint(I18n.T("Scripts.Keyword"), full, ScriptProvider.HintIcon, true))
        for script in ScriptProvider.Scripts {
            if onlyScripts {                                                ; "script " / "script name": scripts only, all of them for no name
                score := (needle = "") ? 50 : FuzzyMatcher.BestKey(needle, script.Keys)
                if (score > 0) {
                    item := ScriptProvider.ItemFor(script, score + 40)
                    item.Exclusive := true
                    results.Push(item)
                }
                continue
            }
            if (script.Keyword != "" && query.Keyword = script.Keyword && script.Argument != "" && script.Mode != "filter" && query.HasRest) {
                item := ScriptProvider.ItemFor(script, 150, query.Rest)
                item.Exclusive := true
                return [item]
            }
            score := FuzzyMatcher.BestKey(needle, script.Keys)
            if (script.Keyword != "" && query.Keyword = script.Keyword)
                score := Max(score, 100)
            if (score > 0)
                results.Push(ScriptProvider.ItemFor(script, score + 10))
        }
        return results
    }

    ; "script:<file name>" (the Scripts folder isn't searched below its top level, so the name is unique): the same in a portable
    ; and an installed ALTRun, and after the data folder moves. Before settings version 7 it was the full path (see Knowledge.ConvertUids)
    static Uid(script) => "script:" StrLower(script.Name)

    ; The scripts as the folder has them now: read again when it changed (looked at most every 2 s, as Search does)
    static Current() {
        if (A_TickCount - ScriptProvider._checked > 2000)
            ScriptProvider._RefreshIfChanged()
        return ScriptProvider.Scripts
    }

    ; The script with this file name (any case; "script:<file name>" in a hotkey), "" when the folder has none
    static FindByName(name) {
        for script in ScriptProvider.Current()
            if (script.Name = name)
                return script
        return ""
    }

    ; Whether "@altrun.selection" names one of these content types (any case)
    static ForTypes(script, types) {
        for name in script.Selection
            for each in types
                if (each = name)
                    return true
        return false
    }

    ; The actions panel rows of the scripts for a selection (SelectionActions.Classify's result), in file name order; none with
    ; scripts turned off. A script whose command line can't take the selection safely isn't shown (as the custom commands)
    static ActionsFor(info) {
        actions := []
        if !ProviderRegistry.IsEnabled(ScriptProvider)
            return actions
        for script in ScriptProvider.Current() {
            if (script.Mode = "filter" || !(ScriptProvider.ForTypes(script, info.Types) && SelectionActions.SafeArgument(info, ScriptProvider.ShellCommand(script))))
                continue
            actions.Push(ResultItem(script.Title, I18n.T("Script.Subtitle", script.Name)
                , {Icon: ScriptProvider.IconOf(script), OnRun: ScriptProvider._Runner(script, ScriptProvider.SelectionArgument(info))}))
        }
        return actions
    }

    static _Runner(script, argument) => (*) => ScriptProvider.Run(script, argument)   ; each row keeps its own argument

    ; What a script gets for a selection: the text, the path, or several paths (an argument each)
    static SelectionArgument(info) {
        switch info.Kind {
            case "files":          return info.Files.Clone()
            case "file", "folder": return info.Path
        }
        return info.Term
    }

    ; The command line the selection checks look at (SelectionActions.SafeArgument): a script whose output is captured runs through cmd too
    ; (named as cmd.exe: only looked at, and %ComSpec% may be any path)
    static ShellCommand(script) {
        line := ScriptProvider.CommandLine(script)
        return Map("Target", ScriptProvider.ThroughCmd(script) ? "cmd.exe /c " line : line, "Arguments", "")
    }

    ; A custom hotkey whose Action is "script:<file name>", by the rules of CustomCommandProvider.RunFromHotkey: a script that takes
    ; something (@altrun.argument or @altrun.selection) gets the selection as its argument (several files: one each), if it passes
    ; the panel's checks; the content types aren't checked. Nothing selected: one with an argument and a keyword opens the search
    ; window with "keyword ", one with an argument and no keyword says so, one with only content types runs without an argument.
    ; Any other script runs directly, without copying the selection. capture: replaces SelectionActions.Capture in the tests
    static RunFromHotkey(name, capture := "") {
        ScriptProvider._RefreshIfChanged()                                  ; just renamed or added: the folder as it is now
        script := ScriptProvider.FindByName(name)
        if !IsObject(script) {
            Logger.Warn("ScriptProvider: the hotkey's script " name " no longer exists")
            return App.Notify(I18n.T("Hotkey.ScriptMissing", name), 3000)
        }
        if (script.Mode = "filter") {                                      ; the search window with "keyword " and the selected line of text, if any
            text := ""
            if (ScriptFilter.Problem(script) = "" && IsObject(selection := IsObject(capture) ? capture.Call() : SelectionActions.Capture())) {
                info := SelectionActions.Classify(selection)
                if (info.Kind = "text" || info.Kind = "url") && StrLen(info.Term) <= 200 && !RegExMatch(info.Term, "[\r\n]")
                    text := info.Term
            }
            return ScriptProvider.Run(script, text)
        }
        takesArgument := (script.Argument != ""), argument := ""
        if (takesArgument || script.Selection.Length) {
            selection := IsObject(capture) ? capture.Call() : SelectionActions.Capture()
            if IsObject(selection) {
                info := SelectionActions.Classify(selection)
                if !SelectionActions.SafeArgument(info, ScriptProvider.ShellCommand(script))
                    return App.Notify(I18n.T("Hotkey.SelectionRefused", script.Title), 3000)
                argument := ScriptProvider.SelectionArgument(info)
            } else if takesArgument {
                return (script.Keyword != "") ? SearchWindow.Show(script.Keyword " ") : App.Notify(I18n.T("Selection.None"), 2000)
            }
        }
        Usage.Count(ScriptProvider.Id)
        DllCall("AllowSetForegroundWindow", "UInt", 0xFFFFFFFF)            ; ASFW_ANY: the window it opens may come to the front
        return ScriptProvider.Run(script, argument)
    }

    static ItemFor(script, score, argument := "") {
        props := {Kind: "file", Arg: script.Path, Icon: ScriptProvider.IconOf(script), Score: score, Source: script
            , Uid: ScriptProvider.Uid(script), OnRun: (*) => ScriptProvider.Run(script, argument)}
        title := script.Title, subtitle := I18n.T("Script.Subtitle", script.Name)
        if (script.Argument != "" || script.Mode = "filter") {
            if (argument = "") {
                props.Valid := (script.Keyword = "")                        ; with a keyword: Enter completes "keyword " so the argument can be typed next
                props.AutoComplete := (script.Keyword != "") ? script.Keyword " " : ""
                hint := ScriptProvider.HintText(script)
                subtitle := ((script.Keyword != "") ? I18n.T("Search.TypeArgAfter", script.Keyword, hint) : hint) " · " subtitle
            } else
                title .= ": " argument
        }
        return ResultItem(title, subtitle, props)
    }

    ; What to type after the keyword: @altrun.argument, or for a filter script without one "text to search for"
    static HintText(script) => (script.Argument != "") ? script.Argument : (script.Mode = "filter") ? I18n.T("Script.FilterHint") : ""

    static IconOf(script) {
        switch script.Ext {
            case "ahk": return A_IsCompiled ? A_ScriptFullPath : A_AhkPath
            case "ps1": return A_WinDir "\System32\WindowsPowerShell\v1.0\powershell.exe"
            default:    return A_ComSpec
        }
    }

    ; The command line that runs the script; argument: one, or an array of them (several selected files)
    static CommandLine(script, argument := "") {
        quoted := '"' script.Path '"'
        for each in ScriptProvider._Arguments(argument)
            quoted .= " " Win.QuoteArg(each)
        switch script.Ext {
            case "ahk": return '"' A_AhkPath '"' (A_IsCompiled ? " /script " : " ") quoted
            case "ps1": return 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File ' quoted
            case "py":  return 'py ' quoted
            default:    return A_ComSpec ' /c "' quoted '"'
        }
    }

    ; Whether the command line goes through cmd.exe: batch files, and every script whose output is captured (cmd redirects it).
    ; cmd doesn't know \" and ends the quoted argument there, so & | < > after it would run as commands: a quote in the argument is refused
    static ThroughCmd(script) => script.Mode != "window" || !RegExMatch(script.Ext, "^(ahk|ps1|py)$")

    static _Arguments(argument) => (argument is Array) ? argument : (argument != "") ? [argument] : []

    static Run(script, argument := "") {
        if (script.Mode = "filter") {                                      ; runs only from the search box: open it in the keyword mode
            if ((problem := ScriptFilter.Problem(script)) != "")
                return App.Notify(I18n.T((problem = "NeedsKeyword") ? "Script.Filter.NeedsKeyword" : "Script.Filter.NotForType", script.Title), 4000)
            return SearchWindow.Show(script.Keyword " " (argument is Array ? "" : argument))
        }
        for each in ScriptProvider._Arguments(argument)
            if (InStr(each, '"') && ScriptProvider.ThroughCmd(script))
                return App.Notify(I18n.T("Script.QuoteInArgument"), 4000)
        SplitPath(script.Path, , &dir)
        command := ScriptProvider.CommandLine(script, argument)
        if (script.Mode = "window")
            return Run(command, dir)
        ScriptProvider.CleanOutputs()
        output := A_Temp "\ALTRun-script-" A_TickCount ".txt"              ; runs in the background, output goes to a temporary file
        Run(A_ComSpec ' /c "' command ' > "' output '" 2>&1"', dir, "Hide", &pid)
        started := A_TickCount
        check() {
            if (ProcessExist(pid) && A_TickCount - started < ScriptProvider.TimeoutMs)
                return
            SetTimer(check, 0)
            ScriptProvider._Finished(script, output)
        }
        SetTimer(check, 300)
    }

    ; The output files of earlier runs, at least a day old (the "output" mode leaves its file for the editor showing it)
    static CleanOutputs(dir := A_Temp, hours := 24) {
        Loop Files, dir "\ALTRun-script-*.txt"
            if (DateDiff(A_Now, A_LoopFileTimeModified, "Hours") >= hours)
                try FileDelete(A_LoopFileFullPath)
    }

    static _Finished(script, output) {
        text := ""
        try text := FileRead(output)
        if (script.Mode = "output") {
            Path.OpenText(output)                                           ; the default program for .txt (usually Notepad); the file stays a day (CleanOutputs)
            return
        }
        try FileDelete(output)
        last := ""
        Loop Parse, text, "`n", "`r"
            if (Trim(A_LoopField) != "")
                last := Trim(A_LoopField)
        message := (last != "") ? last : I18n.T("Script.Done", script.Title)
        App.Notify(message, 3000)
        return message
    }

    ; F3: edit the script in Notepad (also on a filter script's results)
    static EditItem(item) => ScriptProvider.EditScript(item.Source)

    static EditScript(script) {
        Run('notepad.exe "' script.Path '"')
        return false
    }

    static Hidden() => ScriptFilter.Stop()                                  ; the search window hid: end a filter script's run, forget its results

    static Load() {
        scripts := []
        if DirExist(ScriptProvider.Dir) {
            Loop Files, ScriptProvider.Dir "\*.*", "F" {
                ext := StrLower(A_LoopFileExt)
                if !ScriptProvider.Extensions.Has(ext)
                    continue
                if !IsObject(script := ScriptProvider.ReadScript(A_LoopFileFullPath))
                    continue
                position := scripts.Length + 1                              ; by file name (NTFS lists them so, other file systems may not)
                while (position > 1 && StrCompare(scripts[position - 1].Name, script.Name, "Logical") > 0)
                    position -= 1
                scripts.InsertAt(position, script)
            }
        }
        ScriptProvider.Scripts := scripts
        ScriptProvider._lastStamp := ScriptProvider._Stamp()
        ScriptProvider._checked := A_TickCount
    }

    ; Script file -> {Path, Name, Ext, Title, Keyword, Argument, Mode, Selection: [type names], FilterBy: script / altrun, Keys}; reads @altrun.xxx in the first 30 lines
    static ReadScript(path) {
        SplitPath(path, &name, , &ext, &baseName)
        meta := Map("title", baseName, "keyword", "", "argument", "", "mode", "window", "selection", "", "filter", "")
        try {
            Loop Read, path {
                if (A_Index > 30)
                    break
                if RegExMatch(A_LoopReadLine, "i)^\s*(?:;|#|//|::|rem\b)\s*@altrun\.(\w+)\s+(.*?)\s*$", &m)
                    meta[StrLower(m[1])] := m[2]
            }
        } catch
            return ""
        mode := StrLower(meta["mode"])
        if !(mode = "silent" || mode = "output" || mode = "filter")
            mode := "window"
        title := meta["title"], keyword := StrLower(Trim(meta["keyword"]))
        types := []
        for each in StrSplit(meta["selection"], [",", ";"])
            if ((each := Trim(each)) != "")
                types.Push(each)
        pinyinText := Pinyin.Initials(title)
        return {Path: path, Name: name, Ext: StrLower(ext), Title: title, Keyword: keyword, Argument: Trim(meta["argument"]), Mode: mode, Selection: types
              , FilterBy: (StrLower(Trim(meta["filter"])) = "altrun") ? "altrun" : "script"
              , Keys: [FuzzyMatcher.Key(title), (keyword != "") ? FuzzyMatcher.Key(keyword) : "", (pinyinText != title) ? FuzzyMatcher.Key(pinyinText) : ""]}
    }

    ; the modified times of the folder and each file in it joined: read again when it changes
    static _Stamp() {
        stamp := ""
        if DirExist(ScriptProvider.Dir)
            Loop Files, ScriptProvider.Dir "\*.*", "F"
                stamp .= A_LoopFileName A_LoopFileTimeModified "|"
        return stamp
    }

    static _RefreshIfChanged() {
        ScriptProvider._checked := A_TickCount
        if (ScriptProvider._Stamp() != ScriptProvider._lastStamp)
            ScriptProvider.Load()
    }

    ; Example scripts, each showing one form: background with a notification (ahk, silent), a keyword with an argument (bat, window; the keyword mustn't clash with the default commands),
    ; all output in Notepad (ps1, output), live results (ahk, filter; one result per line). ASCII only: Windows PowerShell 5 reads files without BOM as ANSI
    static ExampleScripts() {
        return Map(
            "Environment Variables.ahk",
                "; @altrun.title     Example: Environment Variables`r`n"
              . "; @altrun.keyword   envvar`r`n"
              . "; @altrun.mode      filter`r`n"
              . "; @altrun.argument  part of a name`r`n"
              . '; Live results: type "envvar path" in ALTRun. While you type, ALTRun runs this script with the text (A_Args[1]) and shows' "`r`n"
              . "; what it prints: one result per line, title<Tab>subtitle<Tab>arg<Tab>type (or JSON, see the Scripts page in the wiki).`r`n"
              . "; Type text: Enter copies the value.`r`n"
              . 'query := A_Args.Length ? A_Args[1] : ""' "`r`n"
              . 'block := DllCall("GetEnvironmentStringsW", "Ptr"), at := block, out := ""' "`r`n"
              . 'while ((entry := StrGet(at, "UTF-16")) != "") {' "`r`n"
              . '    at += (StrLen(entry) + 1) * 2' "`r`n"
              . '    split := InStr(entry, "=", , 2)' "`r`n"
              . '    name := SubStr(entry, 1, split - 1), value := StrReplace(SubStr(entry, split + 1), "``t", " ")' "`r`n"
              . '    if (split > 1 && SubStr(name, 1, 7) != "ALTRUN_" && (query = "" || InStr(name, query)))' "`r`n"
              . '        out .= name "``t" value "``t" value "``ttext``n"' "`r`n"
              . '}' "`r`n"
              . 'DllCall("FreeEnvironmentStringsW", "Ptr", block)' "`r`n"
              . 'FileAppend(out, "*", "UTF-8-RAW")' "`r`n",
            "Today.ahk",
                "; @altrun.title  Example: Today's Date`r`n"
              . "; @altrun.mode   silent`r`n"
              . "; Runs in the background; the last line it prints is shown as a notification.`r`n"
              . "; Change the mode to output to read everything it prints in your text editor, or remove it to run normally.`r`n"
              . 'FileAppend(FormatTime(, "dddd, d MMMM yyyy") ", week " SubStr(FormatTime(, "YWeek"), 5), "*")' "`r`n",
            "Port.bat",
                "@echo off`r`n"
              . "rem @altrun.title     Example: Who Uses a Port`r`n"
              . "rem @altrun.keyword   port`r`n"
              . "rem @altrun.argument  a port number`r`n"
              . 'rem Type "port 8080" in ALTRun: the text after the keyword is passed to the script as %1.' "`r`n"
              . "rem No mode line: it runs normally in its own window.`r`n"
              . 'netstat -ano | findstr /c:":%~1 "' "`r`n"
              . "pause`r`n",
            "IP Addresses.ps1",
                "# @altrun.title  Example: IP Addresses`r`n"
              . "# @altrun.mode   output`r`n"
              . "# Runs in the background, then opens everything it prints in your text editor.`r`n"
              . "Get-NetIPAddress -AddressFamily IPv4 | Where-Object IPAddress -ne '127.0.0.1' | Format-Table InterfaceAlias, IPAddress -AutoSize`r`n")
    }

    ; Opens the Scripts folder; creates a few example scripts first if it doesn't exist yet (ExampleScripts)
    static OpenFolder() {
        dir := ScriptProvider.Dir
        if !DirExist(dir) {
            DirCreate(dir)
            for name, text in ScriptProvider.ExampleScripts()
                FileAppend(text, dir "\" name, "UTF-8-RAW")                 ; no BOM: cmd would take a BOM as part of a command on the first line and fail
        }
        ActionCatalog.OpenFolder(dir)                                       ; the file manager in the settings (e.g. Total Commander)
    }
}
