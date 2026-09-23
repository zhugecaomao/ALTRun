;===============================================================================
; CommandRunner.ahk - 搜索命令 + 运行命令 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 用法:
;   CommandRunner.Search(query)        按输入框文字搜索, 结果写入 g_MATCHED 并交给
;                                      MainWindow.ShowResults() 显示。支持:
;                                        "+" / " " / ">" 前缀  -> FallbackCommand 1/2/3
;                                        "/" 前缀              -> 内置 Func 命令面板
;                                        数学表达式            -> 计算结果 (+ 结构计算)
;                                        其它                  -> 模糊匹配 (空格分隔多个关键词)
;   CommandRunner.Execute(cmdLine)     运行一条命令行 "Type | Path | Desc"
;   CommandRunner.GetField(cmdLine, n) 取命令行的第 n 段 (1 = 类型, 2 = 路径, 3 = 描述)
;   CommandRunner.DisplayPath(cmdLine) 第 2 段的可读形式 (Clip 类型显示单行预览)
;   CommandRunner.OpenContainer()      用文件管理器定位当前命令所在目录
;
; 注意: RunCurrentCommand()/OpenContainer() 仍以裸的全局函数外壳留在 ALTRun.ahk
; 里 - 它们登记在 FuncList (可绑定自定义热键) 或是内置 Func 命令, 函数名以
; 字符串形式保存在 ALTRun.json 里。Execute() 里的 %cmdPath%() 按名字动态调用,
; 只认裸的全局函数名, 不认 Class.Method。
;
; 方法名用 Execute 而不是 Run: 方法内部要调用内置的 Run() 函数, 同名容易看混。
;===============================================================================

Class CommandRunner {

    ;---------------------------------------------------------------------------
    ; Search
    ;---------------------------------------------------------------------------
    static Search(query := "") {
        Global g_MATCHED

        g_MATCHED := Array()
        g_RUNTIME["CurrentCommand"] := ""
        listLimit := g_GUI["ListRows"]
        prefix := SubStr(query, 1, 1)
        isExpr := Calc.Looks(query)

        ; Prefix-based fallback shortcuts: "+" / " " / ">"
        if CommandRunner._IsFallbackPrefix(prefix) {
            if (g_FALLBACK.Length = 0)
                return MainWindow.ShowResults(g_MATCHED)
            fallbackIndex := (prefix = "+") ? 1 : (prefix = " ") ? 2 : 3
            g_RUNTIME["CurrentCommand"] := g_FALLBACK[Min(fallbackIndex, g_FALLBACK.Length)]
            g_MATCHED.Push(g_RUNTIME["CurrentCommand"])
            return MainWindow.ShowResults(g_MATCHED)
        }

        ; "/" (optionally followed by more text): command palette - lists every
        ; built-in Func command with its description, live-filtered by whatever
        ; comes after "/". A self-documenting "what can I even type" list.
        if (prefix = "/")
            return CommandRunner._SearchFuncPalette(SubStr(query, 2), listLimit)

        ; Search precomputed command index.
        if (!isExpr) {
            pattern := CommandRunner._FuzzyPattern(query)
            if (pattern = "") {
                Loop Min(listLimit, g_COMMANDS.Length)
                    g_MATCHED.Push(g_COMMANDS[A_Index])
            } else {
                regexPattern := g_RUNTIME["RegEx"] . pattern
                for cmdIndex, searchableText in g_CMDINDEX {
                    if RegExMatch(searchableText, regexPattern) {
                        g_MATCHED.Push(g_COMMANDS[cmdIndex])
                        if g_MATCHED.Length >= listLimit
                            break
                    }
                }
            }
        }

        ; No command match: try expression evaluation, otherwise fallback list.
        if (g_MATCHED.Length > 0) {
            g_RUNTIME["CurrentCommand"] := g_MATCHED[1]
            g_RUNTIME["UseFallback"] := False
        } else {
            if (isExpr) {
                evalResult := Calc.Eval(query)
                if (IsNumber(evalResult)) {
                    g_RUNTIME["UseFallback"] := False
                    g_RUNTIME["CurrentCommand"] := ""
                    g_MATCHED := CommandRunner._StructuralCalc(Round(evalResult, 6)) ; normalize float precision
                    return MainWindow.ShowResults(g_MATCHED, True)
                }
            }

            g_RUNTIME["UseFallback"] := True
            g_MATCHED := g_FALLBACK
            g_RUNTIME["CurrentCommand"] := g_FALLBACK.Length ? g_FALLBACK[1] : ""
        }

        return MainWindow.ShowResults(g_MATCHED)
    }

    ; "/" command palette: filters g_COMMANDS (already rank-sorted) down to Func-type
    ; entries only, then applies the same fuzzy matching Search() uses for
    ; everything else against each one's function name + description.
    static _SearchFuncPalette(remainder, listLimit) {
        funcCmds := []
        for _, cmdLine in g_COMMANDS {
            parts := StrSplit(cmdLine, " | ")
            if (parts.Length >= 1 && parts[1] = "Func")
                funcCmds.Push(cmdLine)
        }

        remainder := Trim(remainder)
        if (remainder = "") {
            Loop Min(listLimit, funcCmds.Length)
                g_MATCHED.Push(funcCmds[A_Index])
        } else {
            regexPattern := g_RUNTIME["RegEx"] . CommandRunner._FuzzyPattern(remainder)
            for _, cmdLine in funcCmds {
                parts := StrSplit(cmdLine, " | ")
                searchable := (parts.Length >= 2 ? parts[2] : "") " " (parts.Length >= 3 ? parts[3] : "")
                if RegExMatch(searchable, regexPattern) {
                    g_MATCHED.Push(cmdLine)
                    if g_MATCHED.Length >= listLimit
                        break
                }
            }
        }

        g_RUNTIME["CurrentCommand"] := g_MATCHED.Length ? g_MATCHED[1] : ""
        g_RUNTIME["UseFallback"] := False
        return MainWindow.ShowResults(g_MATCHED)
    }

    ; "wi ex" -> "wi.*ex": every space-separated token must appear, in order. Regex metacharacters are escaped.
    static _FuzzyPattern(needle) {
        needle := Trim(RegExReplace(needle, "[\s\\]+", " "))
        if !InStr(needle, " ")
            return RegExReplace(needle, "([\\\^\$\.\|\?\*\+\(\)\[\]\{\}])", "\\$1")

        pattern := ""
        for _, token in StrSplit(needle, " ") {
            if (token = "")
                continue
            pattern .= (pattern = "" ? "" : ".*") . RegExReplace(token, "([\\\^\$\.\|\?\*\+\(\)\[\]\{\}])", "\\$1")
        }
        return pattern
    }

    static _IsFallbackPrefix(prefix) {
        if (prefix = "")
            return false
        return InStr("+ >", prefix, 0)
    }

    ; Result rows for a math expression: the value itself, plus (when StruCalc is on)
    ; main-bar count/spacing for a beam of that width and bar counts for that As.
    static _StructuralCalc(evalResult) {
        result    := []
        formatVal := Calc.Thousands(evalResult)
        result.Push("Eval | " formatVal)

        if !g_CONFIG["StruCalc"]
            return result

        result.Push(" | ")  ; 空行分隔
        ; 主筋计算
        rebarNum := Ceil((evalResult - 80) / 300 + 1)
        spacing  := Max(Round((evalResult - 80) / (rebarNum - 0.999)), 0)   ; Use 0.999 to avoid division by zero error
        result.Push("Eval | With beam width = " formatVal " mm")
        result.Push(" | Main bar number = " rebarNum " (" spacing " C/C)")
        result.Push(" | ")  ; 空行分隔
        ; 配筋面积计算
        result.Push("Eval | With As = " formatVal " mm2")
        result.Push(" | Rebar = " Ceil(evalResult / 132.7) "H13 / "
                            . Ceil(evalResult / 201.1) "H16 / "
                            . Ceil(evalResult / 314.2) "H20 / "
                            . Ceil(evalResult / 490.9) "H25 / "
                            . Ceil(evalResult / 804.2) "H32")

        return result
    }

    ;---------------------------------------------------------------------------
    ; Command line fields
    ;---------------------------------------------------------------------------
    ; Field fieldNo of "Type | Path | Desc" ("" if missing). The last split is cached,
    ; callers usually ask for several fields of the same command in a row.
    static GetField(cmdLine, fieldNo) {
        static lastCmd := "", lastParts := ""
        if (cmdLine != lastCmd) {
            lastCmd := cmdLine
            lastParts := StrSplit(cmdLine, " | ")
        }
        return lastParts.Length >= fieldNo ? lastParts[fieldNo] : ""
    }

    ; Field 2 of a command, made readable (Clip text gets a short preview)
    static DisplayPath(cmdLine) {
        return (CommandRunner.GetField(cmdLine, 1) = "Clip")
            ? Clip.ClipPreview(CommandRunner.GetField(cmdLine, 2))
            : CommandRunner.GetField(cmdLine, 2)
    }

    ;---------------------------------------------------------------------------
    ; Run
    ;---------------------------------------------------------------------------
    static Execute(originCmd) {
        if (originCmd = "")
            return

        if (g_RUNTIME["UseDisplay"]) {
            g_LOG.Debug("Execute: blocked in display mode, cmd=" originCmd)
            return
        }

        executed := false
        CommandRunner._ParseArg(MainWindow.Input.Value)                     ; Before Hide(): with KeepInput off it clears the input box and re-runs Search("")
        MainWindow.Hide()
        g_LOG.Debug("Execute: Execute request=" originCmd)

        parts := StrSplit(originCmd, " | ")
        cmdType := parts.Length >= 1 ? parts[1] : ""
        rawPath := parts.Length >= 2 ? parts[2] : ""
        ; Clip payload is plain text, never run it through Path.Resolve().
        cmdPath := (rawPath != "" && cmdType != "Clip") ? Path.Resolve(rawPath, True) : rawPath

        if (cmdType = "") {
            return
        } else if (cmdType = "Clip") {
            executed := Clip.PasteClipText(rawPath)
        } else if (cmdType = "DIR") {
            executed := CommandRunner._OpenDir(cmdPath)
        } else if (cmdType = "FUNC") {
            try {
                %cmdPath%()                                                 ; Resolves a bare global function by name, see header
                executed := true
            } catch as e {
                MsgBox("Could not find function: " cmdPath "`n`nError message: " e.Message, g_TITLE, 48)
            }
        } else {
            try {
                Run(cmdPath)
                executed := true
            } catch as e {
                MsgBox("Could not run command: " cmdPath "`n`nError message: " e.Message, g_TITLE, 48)
            }
        }

        if (executed) {
            CommandStore.UpdateRunCount()
            CommandStore.UpdateRank(originCmd)                              ; Saves by itself only when SmartRank is on
            CommandStore.UpdateHistory(originCmd)
            AppData.SaveAppData()                                           ; Guarantees RunCount/History persist either way, in one write
            g_LOG.Debug("Execute: Execute success, RunCount=" g_CONFIG["RunCount"] ", cmd=" originCmd)
        } else {
            g_LOG.Debug("Execute: Execute failed, cmd=" originCmd)
        }
    }

    ; Sets g_RUNTIME["Arg"] from the input box text: whatever follows a fallback
    ; prefix or the first space, or the whole text when a fallback command runs.
    static _ParseArg(inputText) {
        commandPrefix := SubStr(inputText, 1, 1)
        spacePos := InStr(inputText, " ")

        if CommandRunner._IsFallbackPrefix(commandPrefix) {
            return g_RUNTIME["Arg"] := SubStr(inputText, 2)
        }

        if (spacePos && !g_RUNTIME["UseFallback"]) {
            g_RUNTIME["Arg"] := SubStr(inputText, spacePos + 1)
        } else if (g_RUNTIME["UseFallback"]) {
            g_RUNTIME["Arg"] := inputText
        } else {
            g_RUNTIME["Arg"] := ""
        }
    }

    static _OpenDir(dirPath) {                                              ; Named dirPath, not Path - that's the Util.ahk class
        dirPath := Path.Resolve(dirPath)

        try {
            Run(g_CONFIG["FileMgr"] ' `"' dirPath '`"')
            g_LOG.Debug("_OpenDir: Using=" g_CONFIG["FileMgr"] " to open dir=" dirPath "...OK")
            return true
        } catch as e {
            g_LOG.Debug("_OpenDir: Failed to open dir=" dirPath " Error=" e.Message)
            MsgBox("Could not open dir: " dirPath "`n`nError message: " e.Message, g_TITLE, 48)
            return false
        }
    }

    ; Ctrl+D / list context menu / "Func | OpenContainer": show the current command's file in the file manager
    static OpenContainer() {
        currentCmd := g_RUNTIME["CurrentCommand"]
        if (CommandRunner.GetField(currentCmd, 1) = "Clip")                 ; Clip has no container folder
            return
        cmdPath := CommandRunner.GetField(currentCmd, 2)
        if (cmdPath = "") {
            return MsgBox("No valid file to open container folder.", g_TITLE, 48)
        }
        containerPath := Path.Resolve(cmdPath)                              ; Named containerPath, not Path - that's the Util.ahk class

        try {
            runArg := (g_CONFIG["FileMgr"] = "Explorer.exe") ? ' /Select, `"' containerPath '`"' : ' /P `"' containerPath '`"' ; /P Parent folder
            Run(g_CONFIG["FileMgr"] runArg)
            g_LOG.Debug("OpenContainer: Using=" g_CONFIG["FileMgr"] " to open container dir for file=" containerPath "...OK")
        } catch as e {
            g_LOG.Debug("OpenContainer: Failed to open container dir for file=" containerPath " Error=" e.Message)
            MsgBox("Failed to open container dir for file: " . containerPath "`n`nError message: " . e.Message, g_TITLE, 48)
        }
    }
}
