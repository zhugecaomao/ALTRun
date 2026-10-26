;===============================================================================
; ScriptProvider.ahk - Scripts extension: scripts in the Scripts\ folder become commands (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like Raycast's Script Commands and RunZ's plugins: put a script in Scripts\ (in the program folder, or
; %APPDATA%\ALTRun\Scripts when that isn't writable), and it can be found by name or keyword and run. .ahk .ps1 .bat .cmd .py are supported.
; The comments at the top of a script can say (any comment marker ; # REM :: // works; all are optional):
;   @altrun.title     name              (the file name by default)
;   @altrun.keyword   keyword           comes first when typed exactly
;   @altrun.argument  hint text         takes an argument: type "keyword text", and the text is passed to the script as its first argument
;   @altrun.mode      window | silent | output
;                     window (default) runs normally with a window; silent runs in the background and shows the last line of output as a notification when done;
;                     output runs in the background and opens all output with the default program for .txt (usually Notepad) when done
; Read again when the folder changes (modified time), without reloading ALTRun. F3 edits the script in Notepad.
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
            if (script.Keyword != "" && query.Keyword = script.Keyword && script.Argument != "" && query.HasRest) {
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

    static ItemFor(script, score, argument := "") {
        props := {Kind: "file", Arg: script.Path, Icon: ScriptProvider._Icon(script), Score: score, Source: script
            , Uid: "script:" StrLower(script.Path), OnRun: (*) => ScriptProvider.Run(script, argument)}
        title := script.Title, subtitle := I18n.T("Script.Subtitle", script.Name)
        if (script.Argument != "") {
            if (argument = "") {
                props.Valid := (script.Keyword = "")                        ; with a keyword: Enter completes "keyword " so the argument can be typed next
                props.AutoComplete := (script.Keyword != "") ? script.Keyword " " : ""
                subtitle := ((script.Keyword != "") ? I18n.T("Search.TypeArgAfter", script.Keyword, script.Argument) : script.Argument) " · " subtitle
            } else
                title .= ": " argument
        }
        return ResultItem(title, subtitle, props)
    }

    static _Icon(script) {
        switch script.Ext {
            case "ahk": return A_IsCompiled ? A_ScriptFullPath : A_AhkPath
            case "ps1": return A_WinDir "\System32\WindowsPowerShell\v1.0\powershell.exe"
            default:    return A_ComSpec
        }
    }

    ; The command line that runs the script
    static CommandLine(script, argument := "") {
        quoted := '"' script.Path '"' ((argument != "") ? ' "' StrReplace(argument, '"', '\"') '"' : "")
        switch script.Ext {
            case "ahk": return '"' A_AhkPath '"' (A_IsCompiled ? " /script " : " ") quoted
            case "ps1": return 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File ' quoted
            case "py":  return 'py ' quoted
            default:    return A_ComSpec ' /c "' quoted '"'
        }
    }

    static Run(script, argument := "") {
        SplitPath(script.Path, , &dir)
        command := ScriptProvider.CommandLine(script, argument)
        if (script.Mode = "window")
            return Run(command, dir)
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

    static _Finished(script, output) {
        text := ""
        try text := FileRead(output)
        if (script.Mode = "output") {
            Path.OpenText(output)                                           ; the default program for .txt (usually Notepad); the temporary file stays in %Temp%
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

    ; F3: edit the script in Notepad
    static EditItem(item) {
        Run('notepad.exe "' item.Source.Path '"')
        return false
    }

    static Load() {
        scripts := []
        if DirExist(ScriptProvider.Dir) {
            Loop Files, ScriptProvider.Dir "\*.*", "F" {
                ext := StrLower(A_LoopFileExt)
                if !ScriptProvider.Extensions.Has(ext)
                    continue
                if IsObject(script := ScriptProvider.ReadScript(A_LoopFileFullPath))
                    scripts.Push(script)
            }
        }
        ScriptProvider.Scripts := scripts
        ScriptProvider._lastStamp := ScriptProvider._Stamp()
        ScriptProvider._checked := A_TickCount
    }

    ; Script file -> {Path, Name, Ext, Title, Keyword, Argument, Mode, Keys}; reads @altrun.xxx in the first 30 lines
    static ReadScript(path) {
        SplitPath(path, &name, , &ext, &baseName)
        meta := Map("title", baseName, "keyword", "", "argument", "", "mode", "window")
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
        if !(mode = "silent" || mode = "output")
            mode := "window"
        title := meta["title"], keyword := StrLower(Trim(meta["keyword"]))
        pinyinText := Pinyin.Initials(title)
        return {Path: path, Name: name, Ext: StrLower(ext), Title: title, Keyword: keyword, Argument: Trim(meta["argument"]), Mode: mode
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
    ; all output in Notepad (ps1, output). ASCII only: Windows PowerShell 5 reads files without BOM as ANSI
    static ExampleScripts() {
        return Map(
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
