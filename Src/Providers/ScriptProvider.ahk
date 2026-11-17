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
;                     output runs in the background and opens all output with the default program for .txt (usually Notepad) when done
;                     (both: BackgroundLine, Finished);
;                     filter gives live results: typing "keyword text" runs it with the text while typing and shows what it prints as
;                     results (needs a keyword; .ahk .py .ps1; see ScriptFilter.ahk)
;   @altrun.filter    altrun            a filter script runs once on entering the keyword; ALTRun then filters what it printed while typing
;   @altrun.selection Text, PDF         content type names (built-in or Preferences → Selection): the script is offered in the actions
;                     panel of such a selection (Ctrl+Alt+\, after the custom commands), which is passed to it as its first argument
;                     (several files: one argument each), with the panel's command-line checks (SelectionActions.SafeArgument)
; Read again when the folder changes (modified time), without reloading ALTRun. F3 edits the script in Notepad.
; A custom hotkey with Action "script:<file name>" runs a script (RunFromHotkey); renaming the file breaks it (shown in red in Preferences).
; A filter script isn't offered for a selection; its hotkey opens the search window with "keyword " and the selected line of text.
; Background runs (silent / output) start the script itself with ChildProcess (no cmd.exe, so & | > " in an argument are just text;
; batch files still go through cmd, ThroughCmd), with the variables of a filter script's run (Environment). What the script starts
; keeps running (KeepStarted); the script itself is ended after TimeoutMs. AutoHotkey writes ANSI unless told otherwise, so an .ahk's
; output that isn't UTF-8 is read as ANSI, the others' as OEM (a console's default).
;
; Settings (ALTRun.json -> Features.Scripts): Enabled / Keyword ("script backup" searches scripts only, "script " lists all,
; like the snippets' "snip"; just "script" without the space still shows the commands with that word)
;===============================================================================

class ScriptProvider {
    static Id := "Scripts"
    static Extensions := Map("ahk", 1, "ps1", 1, "bat", 1, "cmd", 1, "py", 1)
    static Scripts := []
    static _lastStamp := "", _checked := 0
    static TimeoutMs := 120000                                              ; how long a background script may run at most
    static MaxOutput := 16777216                                            ; the most an "output" script's text can be (16 MB; more is cut)
    static StartFn := (line, done, options) => ChildProcess.Start(line, done, options)   ; replaced in tests
    static NotifyFn := (text, ms) => App.Notify(text, ms)
    static OpenFn := (file) => Path.OpenText(file)

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
                , {Icon: ScriptProvider.IconOf(script), OnRun: ScriptProvider._Runner(script, ScriptProvider.SelectionArgument(info)), Keyword: script.Keyword}))
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

    ; The command line the selection checks look at (SelectionActions.SafeArgument): a batch file runs through cmd
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

    ; Whether the command line goes through cmd.exe: batch files (started directly, Windows still has cmd read the arguments).
    ; cmd doesn't know \" and ends the quoted argument there, so & | < > after it would run as commands: a quote in the argument is refused
    static ThroughCmd(script) => !RegExMatch(script.Ext, "^(ahk|ps1|py)$")

    ; The command line of a background run (silent / output): CommandLine, with /ErrorStdOut for an .ahk (a syntax error is printed
    ; instead of shown in a dialog; an error while it runs still is)
    static BackgroundLine(script, argument := "") {
        if (script.Ext != "ahk")
            return ScriptProvider.CommandLine(script, argument)
        line := '"' A_AhkPath '"' (A_IsCompiled ? " /script" : "") ' /ErrorStdOut "' script.Path '"'
        for each in ScriptProvider._Arguments(argument)
            line .= " " Win.QuoteArg(each)
        return line
    }

    ; The variables of a background run, as a filter script's run of a script item gets them (ScriptFilter.RunItem): ALTRUN_ACTION=run,
    ; ALTRUN_ARG (the first argument), ALTRUN_KEYWORD, ALTRUN_VERSION and the window in front before (ALTRUN_HWND / _PID / _PROCESS)
    static Environment(script, argument := "") {
        arguments := ScriptProvider._Arguments(argument)
        env := Map("ALTRUN_ACTION", "run", "ALTRUN_ARG", arguments.Length ? arguments[1] : "", "ALTRUN_KEYWORD", script.Keyword
            , "ALTRUN_VERSION", App.Version, "PYTHONIOENCODING", "utf-8")
        for name, value in ScriptFilter.WindowVariables(App.PreviousWindow)
            env[name] := value
        return env
    }

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
        if (script.Mode = "window")
            return Run(ScriptProvider.CommandLine(script, argument), dir)
        ScriptProvider.CleanOutputs()
        output := (script.Mode = "output")
        ScriptProvider.StartFn.Call(ScriptProvider.BackgroundLine(script, argument), (result) => ScriptProvider.Finished(script, result)
            , {Dir: dir, Env: ScriptProvider.Environment(script, argument), TimeoutMs: ScriptProvider.TimeoutMs, KeepStarted: true
            , MaxStdout: output ? ScriptProvider.MaxOutput : "", MaxStderr: output ? 65536 : ""
            , Codepage: (script.Ext = "ahk") ? 0 : ""})
    }

    ; The output files of earlier runs, at least a day old (the "output" mode leaves its file for the editor showing it)
    static CleanOutputs(dir := A_Temp, hours := 24) {
        Loop Files, dir "\ALTRun-script-*.txt"
            if (DateDiff(A_Now, A_LoopFileTimeModified, "Hours") >= hours)
                try FileDelete(A_LoopFileFullPath)
    }

    ; A background run ended (ChildProcess result). silent: exit code 0 shows the last line printed (none: "Title: done"); another exit
    ; code the first line of standard error, else the last line printed, else the exit code; can't start / time-out the reason.
    ; output: what it printed, then standard error, opened as text (OpenOutput). Returns the text shown (output: the file)
    static Finished(script, result) {
        if Logger.Enabled
            Logger.Debug("Perf: script " script.Name " " result.Ms " ms, exit " result.ExitCode (result.TimedOut ? ", timed out" : "") (result.Error != "" ? ", can't start" : ""))
        timedOut := I18n.T("Script.TimedOut", Round(ScriptProvider.TimeoutMs / 1000))
        if (result.Error != "")
            return ScriptProvider._Notify(script.Title ": " I18n.T("ScriptFilter.CantStart", result.Error), 4000)
        if (script.Mode = "output")
            return ScriptProvider.OpenOutput(result, timedOut)
        if result.TimedOut
            return ScriptProvider._Notify(script.Title ": " timedOut, 4000)
        last := ScriptProvider._LineOf(result.Stdout)
        if (result.ExitCode != 0) {
            if ((first := ScriptProvider._LineOf(result.Stderr, false)) != "")
                return ScriptProvider._Notify(script.Title ": " first, 4000)
            return (last != "") ? ScriptProvider._Notify(last, 3000) : ScriptProvider._Notify(script.Title ": " I18n.T("ScriptFilter.ExitCode", result.ExitCode), 4000)
        }
        return ScriptProvider._Notify((last != "") ? last : I18n.T("Script.Done", script.Title), 3000)
    }

    ; An "output" script's text: what it printed, then standard error after an empty line, a note when more was printed than kept or
    ; it was stopped (timedOut: that note). Written as UTF-8 to a temporary file opened with the default program for .txt (usually
    ; Notepad); the file stays a day (CleanOutputs). Returns the file
    static OpenOutput(result, timedOut := "") {
        text := result.Stdout
        add(part) => text := (Trim(text, " `t`r`n") = "") ? part : RTrim(text, " `t`r`n") "`r`n`r`n" part
        if (Trim(result.Stderr, " `t`r`n") != "")
            add(result.Stderr)
        if result.Truncated
            add(I18n.T("Script.OutputCut", Round(ScriptProvider.MaxOutput / 1048576)))
        if result.TimedOut
            add(timedOut)
        textFile := A_Temp "\ALTRun-script-" A_TickCount ".txt"
        while FileExist(textFile)
            textFile := A_Temp "\ALTRun-script-" A_TickCount "-" A_Index ".txt"
        out := FileOpen(textFile, "w", "UTF-8")                             ; with a BOM, so any editor reads it as UTF-8
        out.Write(text), out.Close()
        ScriptProvider.OpenFn.Call(textFile)
        return textFile
    }

    ; The last (or first) line of text that isn't blank, trimmed; "" for none
    static _LineOf(text, last := true) {
        found := ""
        Loop Parse, text, "`n", "`r"
            if ((line := Trim(A_LoopField)) != "") {
                found := line
                if !last
                    break
            }
        return found
    }

    static _Notify(message, ms) {
        ScriptProvider.NotifyFn.Call(message, ms)
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

    ; Opens the Scripts folder, created first if it doesn't exist yet (empty: ready-made scripts are in the script library on the website)
    static OpenFolder() {
        dir := ScriptProvider.Dir
        if !DirExist(dir)
            DirCreate(dir)
        ActionCatalog.OpenFolder(dir)                                       ; the file manager in the settings (e.g. Total Commander)
    }
}
