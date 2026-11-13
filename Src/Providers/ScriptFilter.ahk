;===============================================================================
; ScriptFilter.ahk - live results from a script: "@altrun.mode filter" (part of the Scripts extension) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like Alfred's Script Filter: typing "keyword text" runs the script in the background with the text, and what it prints becomes
; the results (files, folders, web addresses, text), which run, open the actions panel (→) and F3-edit the script as usual.
;   - Only in the script's keyword mode ("keyword " with the space); never for scripts without a keyword. .ahk .py .ps1 only:
;     a batch file always goes through cmd.exe, which would take & | > in the text as commands
;   - The text arrives as the first argument (empty: "") and in the variable ALTRUN_QUERY (also ALTRUN_KEYWORD, ALTRUN_VERSION).
;     PowerShell gets it only from the variable (-Command, never put into the command's text; a text starting with - isn't a parameter)
;   - The first run starts at once; later ones once typing paused DebounceMs. A new run ends the one still running (with what it
;     started, ChildProcess); so do hiding the window and leaving the keyword. A run takes at most TimeoutMs
;   - Results are kept per text until the window hides (Backspace back to a text shows them again without running)
;   - While it runs, the last results stay; with none yet, the script's own row, after WaitRowMs a "Running..." row
;   - Output: JSON {"items": [{title, subtitle, arg, type, icon, uid, valid, autocomplete, match, text: {largetype}}]} (a subset of
;     Alfred's, names the same) or one result per line: title<Tab>subtitle<Tab>arg<Tab>type<Tab>icon (all but the title optional).
;     type: file / folder / url / text, guessed from arg without touching the disk when missing. icon: a local file or an IconCache
;     form, never a web address. uid: the item takes part in learned ranking; without it the script's order is kept
;   - "@altrun.filter altrun": run once on entering the keyword (with an empty text), then ALTRun filters what it printed while
;     typing (fuzzy, on "match" or the title, pinyin initials too) without running it again: for lists that don't depend on the
;     text (projects, hosts...) and for slow starters such as PowerShell
;   - Nothing usable printed: one error row (the first line of standard error, the exit code, the JSON problem, the time-out);
;     Enter edits the script, Ctrl+L shows all of standard error. Exit code 0 and no items: a "No results" row
;   - The log gets the script's file name, time, item count and exit code, never the text typed or the output
; Design: notes design/script-filter.md (F1-F10).
;
; Usage (ScriptProvider):
;   ScriptFilter.Supports(script)        whether it can give live results (keyword, .ahk / .py / .ps1)
;   ScriptFilter.Results(script, text)   the rows to show now; starts a run when needed (ScriptFilter.OnResults is called when it ends)
;   ScriptFilter.Leave()                 the search is no longer in a filter script's keyword mode: end the run
;   ScriptFilter.Stop()                  the window hid: end the run and forget the results
;===============================================================================

class ScriptFilter {
    static DebounceMs := 150, WaitRowMs := 300, TimeoutMs := 10000, MaxStdout := 524288, MaxItems := 200, CacheSize := 20
    static Extensions := Map("ahk", 1, "py", 1, "ps1", 1)
    static OnResults := () => SearchWindow.LiveResultsArrived()             ; results arrived: search again (replaced in tests)
    static StartFn := (line, done, options) => ChildProcess.Start(line, done, options)   ; replaced in tests
    static KillFn := (child) => ChildProcess.Kill(child)

    static _cache := Map(), _order := []                                     ; script name `n text -> {Items, Error, Detail}, oldest key first
    static _last := Map()                                                    ; script name -> its latest entry (shown while the next one runs)
    static _child := "", _runningKey := "", _gen := 0
    static _wanted := "", _wantedAt := 0, _session := ""
    static _runTimer := "", _waitTimer := ""

    static Supports(script) => script.Keyword != "" && ScriptFilter.Extensions.Has(script.Ext)

    ; Why a filter script gives no live results ("" when it does): for the Scripts page and running it from a hotkey
    static Problem(script) {
        if (script.Mode != "filter")
            return ""
        if (script.Keyword = "")
            return "NeedsKeyword"
        return ScriptFilter.Extensions.Has(script.Ext) ? "" : "NotForType"
    }

    static Results(script, text) {
        filterText := text
        if (script.FilterBy = "altrun")                                    ; one run for every text: ALTRun filters its items
            text := ""
        key := ScriptFilter._Key(script, text)
        if (ScriptFilter._session != script.Name) {                        ; just entered this keyword: the first run starts at once
            ScriptFilter.Leave()
            ScriptFilter._session := script.Name
            first := true
        } else
            first := false
        if ScriptFilter._cache.Has(key) {
            ScriptFilter._CancelWanted()
            if (ScriptFilter._runningKey != "" && ScriptFilter._runningKey != key)
                ScriptFilter._EndRun()
            return ScriptFilter.Rows(script, ScriptFilter._cache[key], filterText)
        }
        ScriptFilter._Want(script, text, key, first)
        if ScriptFilter._last.Has(script.Name)
            return ScriptFilter.Rows(script, ScriptFilter._last[script.Name], filterText)
        if (A_TickCount - ScriptFilter._wantedAt >= ScriptFilter.WaitRowMs - 20)   ; (its timer may come a tick early)
            return [ScriptFilter._Row(script, I18n.T("ScriptFilter.Running", script.Title), "", {Valid: false})]
        return [ScriptFilter._Row(script, script.Title, ScriptProvider.HintText(script), {Valid: false})]
    }

    static Leave() {
        if (ScriptFilter._session = "" && !IsObject(ScriptFilter._wanted) && !ScriptFilter.Running)
            return
        ScriptFilter._CancelWanted()
        ScriptFilter._EndRun()
        ScriptFilter._session := ""
    }

    static Stop() {
        ScriptFilter.Leave()
        ScriptFilter._cache := Map(), ScriptFilter._order := [], ScriptFilter._last := Map()
    }

    static Running => ScriptFilter._runningKey != ""

    ; The command line, by the script's type (Supports); text: the typed text, passed as the first argument
    static CommandLine(script, text) {
        switch script.Ext {
            case "ahk": return '"' A_AhkPath '"' (A_IsCompiled ? " /script" : "") ' /ErrorStdOut "' script.Path '" ' Win.QuoteArg(text)
            case "py":  return 'py "' script.Path '" ' Win.QuoteArg(text)
            case "ps1": return '"' A_WinDir '\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "'
                . "[Console]::OutputEncoding=[Text.Encoding]::UTF8; & '" StrReplace(script.Path, "'", "''") "' $env:ALTRUN_QUERY; exit $LASTEXITCODE" '"'
        }
        throw ValueError("not a filter script type: " script.Ext)
    }

    static Environment(script, text) {
        return Map("ALTRUN_QUERY", text, "ALTRUN_KEYWORD", script.Keyword, "ALTRUN_VERSION", App.Version, "PYTHONIOENCODING", "utf-8")
    }

    ; A run's result (ChildProcess) -> {Items: [...], Error: "" or the error row's text, Detail: all of standard error}
    static Read(result) {
        entry := {Items: [], Error: "", Detail: Trim(result.Stderr, " `t`r`n")}
        if (result.Error != "")
            return (entry.Error := I18n.T("ScriptFilter.CantStart", result.Error), entry)
        if result.TimedOut
            return (entry.Error := I18n.T("ScriptFilter.TimedOut", Round(ScriptFilter.TimeoutMs / 1000)), entry)
        problem := ""
        try entry.Items := ScriptFilter.Parse(result.Stdout)
        catch as e
            problem := e.Message
        if (entry.Items.Length && problem = "")
            return entry
        firstLine := ""
        Loop Parse, entry.Detail, "`n", "`r"
            if ((firstLine := Trim(A_LoopField)) != "")
                break
        if (firstLine != "")
            entry.Error := firstLine
        else if (result.ExitCode != 0)
            entry.Error := I18n.T("ScriptFilter.ExitCode", result.ExitCode)
        else if (problem != "")
            entry.Error := I18n.T("ScriptFilter.BadJson", problem)
        return entry
    }

    ; What a script printed -> [{Title, Subtitle, Arg, Type, Icon, Uid, Valid, AutoComplete, Match, LargeText}], at most MaxItems;
    ; throws when JSON can't be read. JSON when it starts with "{", else one result per line (Tab-separated)
    static Parse(text) {
        text := LTrim(text, " `t`r`n")
        items := []
        if (SubStr(text, 1, 1) = "{") {
            data := JSON.Parse(text)
            if !(data is Map && data.Has("items") && data["items"] is Array)
                throw ValueError(I18n.T("ScriptFilter.NoItems"))
            for each in data["items"] {
                if !(each is Map)
                    continue
                field := ScriptFilter._Field.Bind(ScriptFilter, each)
                icon := each.Has("icon") ? each["icon"] : ""
                if (icon is Map)
                    icon := icon.Has("path") && !IsObject(icon["path"]) ? icon["path"] : ""
                large := (each.Has("text") && each["text"] is Map && each["text"].Has("largetype")) ? String(each["text"]["largetype"]) : ""
                item := ScriptFilter._Item(field("title"), field("subtitle"), field("arg"), field("type"), IsObject(icon) ? "" : icon
                    , field("uid"), !(each.Has("valid") && (each["valid"] = 0 || each["valid"] = "false")), field("autocomplete"), field("match"), large)
                if IsObject(item) && (items.Push(item), items.Length >= ScriptFilter.MaxItems)
                    break
            }
            return items
        }
        Loop Parse, text, "`n", "`r" {
            columns := StrSplit(A_LoopField, "`t")
            column := (n) => columns.Length >= n ? Trim(columns[n]) : ""
            item := ScriptFilter._Item(column(1), column(2), column(3), column(4), column(5))
            if IsObject(item) && (items.Push(item), items.Length >= ScriptFilter.MaxItems)
                break
        }
        return items
    }

    static _Field(fields, name) => fields.Has(name) && !IsObject(fields[name]) ? String(fields[name]) : ""

    static _Item(title, subtitle, arg, type, icon, uid := "", valid := true, autoComplete := "", match := "", large := "") {
        if ((title := Trim(title)) = "")
            return ""
        if (arg = "")
            arg := title
        return {Title: title, Subtitle: subtitle, Arg: arg, Type: ScriptFilter.KindOf(type, arg), Icon: ScriptFilter._LocalIcon(icon)
            , Uid: Trim(uid), Valid: valid, AutoComplete: autoComplete, Match: match, LargeText: large}
    }

    ; file / folder / url / text: as given, else guessed from arg by its form only (a network path would make a check wait)
    static KindOf(type, arg) {
        type := StrLower(Trim(type))
        if RegExMatch(type, "^(file|folder|url|text)$")
            return type
        if RegExMatch(arg, "i)^([a-z][a-z0-9+.\-]*://|mailto:)")
            return "url"
        if RegExMatch(arg, "^([A-Za-z]:\\|\\\\[^\\])")
            return (SubStr(arg, -1) = "\") ? "folder" : "file"
        return "text"
    }

    static _LocalIcon(icon) => RegExMatch(Trim(icon), "i)^[a-z][a-z0-9+.\-]*://") ? "" : Trim(icon)   ; no web addresses: nothing is downloaded for an icon

    ; The rows for an entry (Read): the items, the error row or "No results"; new ResultItems each time (the search changes their scores).
    ; filterText: for a script ALTRun filters (@altrun.filter altrun), only the items matching it, best first
    static Rows(script, entry, filterText := "") {
        if (entry.Error != "")
            return [ScriptFilter._Row(script, I18n.T("ScriptFilter.Error", script.Title), entry.Error
                , {LargeText: (entry.Detail != "") ? entry.Detail : entry.Error, RunTitle: I18n.T("ScriptFilter.EditScript"), OnRun: (*) => ScriptProvider.EditScript(script)})]
        items := entry.Items, scores := ""
        if (script.FilterBy = "altrun" && filterText != "")
            items := ScriptFilter.Matching(items, filterText, &scores)
        if !items.Length
            return [ScriptFilter._Row(script, I18n.T("ScriptFilter.NoResults"), script.Title, {Valid: false})]
        rows := []
        for index, each in items {
            props := {Kind: each.Type, Arg: each.Arg, Valid: each.Valid, Score: IsObject(scores) ? 1000 + scores[index] : 1000 - index, Exclusive: true, Source: script
                , Icon: (each.Icon != "") ? each.Icon : ScriptFilter._KindIcon(script, each)
                , Uid: (each.Uid != "") ? ScriptProvider.Uid(script) ":" each.Uid : "", LargeText: each.LargeText}
            if (each.AutoComplete != "")
                props.AutoComplete := script.Keyword " " each.AutoComplete
            rows.Push(ResultItem(each.Title, each.Subtitle, props))
        }
        return rows
    }

    ; The items matching text (FuzzyMatcher on "match", else the title, and its pinyin initials), in their order; scores: each one's score
    static Matching(items, text, &scores) {
        needle := StrLower(text), found := [], scores := []
        for each in items {
            if !each.HasOwnProp("Keys") {
                name := (each.Match != "") ? each.Match : each.Title, initials := Pinyin.Initials(name)
                each.Keys := [FuzzyMatcher.Key(name), (initials != name) ? FuzzyMatcher.Key(initials) : ""]
            }
            if ((score := FuzzyMatcher.BestKey(needle, each.Keys)) > 0)
                found.Push(each), scores.Push(score)
        }
        return found
    }

    static _KindIcon(script, each) {
        switch each.Type {
            case "file", "folder": return each.Arg
            case "url":            return "url:"
        }
        return ScriptProvider.IconOf(script)
    }

    static _Row(script, title, subtitle, props) {
        props.Icon := ScriptProvider.IconOf(script), props.Score := 1000, props.Exclusive := true, props.Source := script
        if !props.HasOwnProp("Kind")
            props.Kind := "command"
        return ResultItem(title, subtitle, props)
    }

    static _Key(script, text) => StrLower(script.Name) "`n" text

    static _Want(script, text, key, first) {
        if (ScriptFilter._runningKey = key || IsObject(ScriptFilter._wanted) && ScriptFilter._wanted.Key = key)
            return
        if !(IsObject(ScriptFilter._wanted) || ScriptFilter.Running) {          ; waiting since the first text that had nothing to show
            ScriptFilter._wantedAt := A_TickCount
            if (ScriptFilter._waitTimer = "")
                ScriptFilter._waitTimer := () => ScriptFilter._WaitRowDue()
            if !ScriptFilter._last.Has(script.Name)
                SetTimer(ScriptFilter._waitTimer, -ScriptFilter.WaitRowMs)
        }
        ScriptFilter._wanted := {Script: script, Text: text, Key: key}
        if (ScriptFilter._runTimer = "")
            ScriptFilter._runTimer := () => ScriptFilter._Run()
        SetTimer(ScriptFilter._runTimer, first ? -1 : -ScriptFilter.DebounceMs)   ; -1: after this search has been shown
    }

    static _WaitRowDue() {                                                  ; still nothing to show: the "Running..." row
        if ScriptFilter.Running || IsObject(ScriptFilter._wanted)
            ScriptFilter.OnResults.Call()
    }

    static _CancelWanted() {
        ScriptFilter._wanted := ""
        for fn in [ScriptFilter._runTimer, ScriptFilter._waitTimer]
            if (fn != "")
                SetTimer(fn, 0)
    }

    static _EndRun() {
        if IsObject(ScriptFilter._child)
            ScriptFilter.KillFn.Call(ScriptFilter._child)
        ScriptFilter._child := "", ScriptFilter._runningKey := ""
        ScriptFilter._gen++
    }

    static _Run() {
        want := ScriptFilter._wanted
        if !IsObject(want)
            return
        ScriptFilter._wanted := ""
        ScriptFilter._EndRun()
        gen := ScriptFilter._gen
        SplitPath(want.Script.Path, , &dir)
        ScriptFilter._runningKey := want.Key
        child := ScriptFilter.StartFn.Call(ScriptFilter.CommandLine(want.Script, want.Text), (result) => ScriptFilter._Done(want, gen, result)
            , {Dir: dir, Env: ScriptFilter.Environment(want.Script, want.Text), TimeoutMs: ScriptFilter.TimeoutMs, MaxStdout: ScriptFilter.MaxStdout})
        if (gen = ScriptFilter._gen && ScriptFilter._runningKey = want.Key)    ; (a start that failed may already have reported)
            ScriptFilter._child := child
    }

    static _Done(want, gen, result) {
        entry := ScriptFilter.Read(result)
        if Logger.Enabled
            Logger.Debug("Perf: script filter " want.Script.Name " " result.Ms " ms, " entry.Items.Length " items, exit " result.ExitCode
                . (entry.Error != "" ? ", error" : "") (result.TimedOut ? ", timed out" : ""))
        if (gen != ScriptFilter._gen || ScriptFilter._session != want.Script.Name)
            return                                                          ; ended or left meanwhile: the window may be hidden, nothing to keep
        ScriptFilter._child := "", ScriptFilter._runningKey := ""
        if ScriptFilter._cache.Has(want.Key)
            ScriptFilter._order.RemoveAt(ScriptFilter._IndexOf(want.Key))
        ScriptFilter._cache[want.Key] := entry, ScriptFilter._order.Push(want.Key)
        while (ScriptFilter._order.Length > ScriptFilter.CacheSize)
            ScriptFilter._cache.Delete(ScriptFilter._order.RemoveAt(1))
        ScriptFilter._last[want.Script.Name] := entry
        ScriptFilter.OnResults.Call()
    }

    static _IndexOf(key) {
        for index, each in ScriptFilter._order
            if (each == key)
                return index
        return 0
    }
}
