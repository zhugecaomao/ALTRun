;===================================================
; ALTRun - An effective launcher for Windows
; https://github.com/zhugecaomao/ALTRun
;===================================================
#Requires AutoHotkey v2.0
#SingleInstance Force
#NoTrayIcon
#Warn All, OutputDebug
; Project layout (every #Include is explicit - auto-include only reliably covers
; ClassName(...) construction calls, not ClassName.Method(...) static calls like
; JSON.parse(), so relying on it is asking for an "unassigned variable" failure):
;   Lib\            General-purpose libraries, nothing ALTRun-specific
;   Src\Core\       Data and state: ALTRun.json, command store, UI text
;   Src\UI\         Windows: main window, settings, command manager, PT Tools / SPF2M
;   Src\Features\   Self-contained features wired in as commands/hotkeys
;   Res\            Data files and binaries (not code)

; --- Lib: general-purpose libraries ---
#Include Lib\JSON.ahk                    ; JSON.parse()/JSON.stringify()
#Include Lib\Logger.ahk                  ; g_LOG - see just below
#Include Lib\Util.ahk                    ; Path / Fonts / Win / Pinyin / Calc / Arr
#Include Lib\Dialogs.ahk                 ; FontDialog / ColorDialog

; --- Src\Core: data and state ---
#Include Src\Core\Language.ahk           ; Language.Load() - builds g_LNG, the UI text table
#Include Src\Core\IniMigration.ahk       ; IniMigration.MigrateFromIni() - one-off ALTRun.ini -> ALTRun.json
#Include Src\Core\AppData.ahk            ; AppData.LoadAppData()/SaveAppData() - reads/writes ALTRun.json
#Include Src\Core\CommandStore.ahk       ; CommandStore.LoadCommands() etc. - command cache/rank/usage/history

; --- Src\UI: windows ---
#Include Src\UI\MainWindow.ahk           ; MainWindow - search box, result list, tray/right-click menus, hotkeys
#Include Src\UI\OptionsWindow.ahk        ; OptionsWindow.Show() - settings window, see Options() below
#Include Src\UI\CommandManager.ahk       ; CommandManager.Open()/Edit()/Delete() - see OpenCommandManager() below
#Include Src\UI\PTToolsWindow.ahk        ; PTToolsWindow - Rebar/BRC calculator + SPF2M, see PTTools() below

; --- Src\Features: self-contained features ---
#Include Src\Features\Clip.ahk           ; Clip snippet command + clipboard text transforms
#Include Src\Features\Kanji.ahk          ; Kanji.ToSimplified()/ToTraditional() - see ClipToSimplified() below
#Include Src\Features\Listary.ahk        ; Listary.Init() - open/save dialog path quick-switch
#Include Src\Features\AutoDate.ahk       ; AutoDate.Init() - Ctrl+D auto-date
#Include Src\Features\SystemActions.ahk  ; shutdown/volume/process list/search engines/etc
#Include Src\Features\UpdateChecker.ahk  ; UpdateChecker.Check() - see AutoCheckUpdate()/Update() below
#Include Src\Features\Indexer.ahk        ; startup shortcuts + file index rebuild, see Reindex() below
SetWorkingDir(A_ScriptDir)
FileEncoding("UTF-8")

;@Ahk2Exe-SetName ALTRun
;@Ahk2Exe-SetDescription ALTRun - An effective launcher
;@Ahk2Exe-SetVersion 2026.08.12
;@Ahk2Exe-SetCopyright Copyright (c) 2013-2026
;@Ahk2Exe-SetOrigFilename ALTRun.ahk


;===================================================
; 声明全局变量, 默认情况下, 函数是假定-局部的
; 在函数内访问或创建的变量默认为局部的, 但以下情况除外:
; - 全局变量只能被函数读取, 不能被赋值或使用引用运算符(&).
; - 嵌套函数可以引用由闭合它的函数创建的局部或静态变量.
; 内置类, 如 Object; 它们被预定义为全局变量
;===================================================
Global g_LOG   := Logger                                                ; Logger is a static class - this points g_LOG at the class
                                                                         ; itself (no parens/instance), so every existing g_LOG.Debug(...)
                                                                         ; call below still works unchanged, same as calling Logger.Debug(...).
Logger.Rotate()                                                        ; Truncate ALTRun.log to a fresh .old copy if it's grown too big.
Global g_JSON  := A_ScriptDir . "\ALTRun.json"    ; Commands live here, ini sections are capped at 64 KB by the Windows API
Global g_TITLE := "ALTRun - v2026.08.12"

Global g_CMDDATA  := Map()           ; Command store: section name -> Map(commandLine -> rank), FallbackCommand -> Array
Global g_COMMANDS := Array()         ; All commands
Global g_CMDINDEX := Array()         ; Searchable text for All commands
Global g_FALLBACK := Array()         ; Fallback commands
Global g_HISTORY := Array()         ; Execution history
Global g_MATCHED  := Array()         ; Matched commands
Global g_DELUNDO  := Array()         ; Ctrl+Z undo stack for DelCommand: {Section, CmdLine, Rank}, most recent last

Global g_CONFIG := Map(
    "AutoStartup"    , 1,
    "EnableSendTo"   , 1,
    "InStartMenu"    , 1,
    "ShowTrayIcon"   , 1,
    "HideOnLostFocus", 1,
    "AlwaysOnTop"    , 1,
    "ShowCaption"    , 0,
    "XPthemeBg"      , 1,
    "EscClearInput"  , 1,
    "KeepInput"      , 1,
    "ShowIcon"       , 1,
    "LargeIcons"     , 1,
    "SendToGetLnk"   , 1,
    "SaveHistory"    , 1,
    "SaveLog"        , 1,
    "MatchPath"      , 0,
    "ShowGrid"       , 0,
    "ShowHdr"        , 0,
    "ShowSN"         , 1,
    "ShowBorder"     , 1,
    "SmartRank"      , 1,
    "SmartMatch"     , 1,
    "MatchBeginning" , 0,
    "ShowHint"       , 1,
    "ShowRunCount"   , 1,
    "ShowStatusBar"  , 1,
    "ShowBtnRun"     , 1,
    "ShowBtnOpt"     , 1,
    "DoubleBuffer"   , 1,
    "StruCalc"       , 0,
    "ShortenPath"    , 1,
    "Chinese"        , 0,
    "MatchPinyin"    , 1,
    "MidScrollSwitch", 0,
    "MidClickRun"    , 0,
    "SpaceToRun"     , 0,
    "AutoEngIME"     , 0,
    "AutoUpdateCheck", 1,
    "RoundCorner"    , 1,
    "ClipSendMode"   , 1,                                               ; Clip paste mode, 1 = Clipboard + Ctrl+V (default), 2 = SendInput raw keystrokes
    "ClipPasteDelay" , 300,                                             ; ms to wait after Ctrl+V before restoring the old clipboard
    "HistoryLen"     , 10,
    "RunCount"       , 0,
    "AutoSwitchDir"  , 0,
    "FileMgr"        , "Explorer.exe",
    "IndexDir"       , "A_ProgramsCommon,A_StartMenu,C:\Path\IndexLocation",
    "IndexType"      , "*.lnk,*.exe",
    "IndexDepth"     , 2,
    "IndexExclude"   , "Uninstall *",
    "IndexStoreApp"  , 1,
    "Everything"     , "C:\Apps\Everything.exe",
    "DialogWin"      , "ahk_class #32770",
    "FileMgrID"      , "ahk_class CabinetWClass, ahk_class TTOTAL_CMD",
    "ExcludeWin"     , "ahk_class SysListView32, ahk_exe Explorer.exe"
)

g_LOG.Debug("///// ALTRun is starting... /////`n")
Global g_LNG := Language.Load()                                            ; UI text table (English/Chinese), see Src\Core\Language.ahk

Global g_CONFIG_LABELS := Map(
    "AutoStartup"    , g_LNG[101],
    "EnableSendTo"   , g_LNG[102],
    "InStartMenu"    , g_LNG[103],
    "ShowTrayIcon"   , g_LNG[104],
    "HideOnLostFocus", g_LNG[105],
    "AlwaysOnTop"    , g_LNG[106],
    "ShowCaption"    , g_LNG[107],
    "XPthemeBg"      , g_LNG[108],
    "EscClearInput"  , g_LNG[109],
    "KeepInput"      , g_LNG[110],
    "ShowIcon"       , g_LNG[111],
    "SendToGetLnk"   , g_LNG[112],
    "SaveHistory"    , g_LNG[113],
    "SaveLog"        , g_LNG[114],
    "MatchPath"      , g_LNG[115],
    "ShowGrid"       , g_LNG[116],
    "ShowHdr"        , g_LNG[117],
    "ShowSN"         , g_LNG[118],
    "ShowBorder"     , g_LNG[119],
    "SmartRank"      , g_LNG[120],
    "SmartMatch"     , g_LNG[121],
    "MatchBeginning" , g_LNG[122],
    "ShowHint"       , g_LNG[123],
    "ShowRunCount"   , g_LNG[124],
    "ShowStatusBar"  , g_LNG[125],
    "ShowBtnRun"     , g_LNG[126],
    "ShowBtnOpt"     , g_LNG[127],
    "DoubleBuffer"   , g_LNG[128],
    "StruCalc"       , g_LNG[129],
    "ShortenPath"    , g_LNG[130],
    "Chinese"        , g_LNG[131],
    "MatchPinyin"    , g_LNG[132],
    "MidScrollSwitch", g_LNG[133],
    "MidClickRun"    , g_LNG[134],
    "SpaceToRun"     , g_LNG[138],
    "AutoEngIME"     , g_LNG[139],
    "AutoUpdateCheck", g_LNG[135],
    "LargeIcons"     , g_LNG[136],
    "RoundCorner"    , g_LNG[137]
)

Global g_HOTKEY := Map(
    "Hotkey1"       , "F1",
    "Trigger1"      , "About",
    "Hotkey2"       , "F2",
    "Trigger2"      , "Options",
    "Hotkey3"       , "F3", 
    "Trigger3"      , "EditCommand",
    "Hotkey4"       , "F4", 
    "Trigger4"      , "UserCommand",
    "Hotkey5"       , "^n", 
    "Trigger5"      , "NewCommand",
    "Hotkey6"       , "None", 
    "Trigger6"      , "Unset",
    "Hotkey7"       , "None", 
    "Trigger7"      , "Unset",
    "CondTitle"     , "ahk_exe RAPTW.exe",
    "CondHotkey"    , "~Mbutton",
    "CondAction"    , "PTTools",
    "GlobalHotkey1" , "~!Space",  ; ~ allows system to also process this hotkey
    "GlobalHotkey2" , "!r",
    "TotalCMDDir"   , "^g",
    "ExplorerDir"   , "^e",
    "AutoDateAtEnd" , "ahk_class TCmtEditForm,ahk_exe Notepad4.exe",
    "AutoDateAEHKey", "^d",
    "AutoDateBefExt", "ahk_class CabinetWClass,ahk_class Progman,ahk_class WorkerW,ahk_class #32770",
    "AutoDateBEHKey", "^d"
)

Global g_GUI := Map( ; GUI related variables
    "ListRows"      , 9,
    "ColWidth"      , "36,0,300,AutoHdr",
    "MainGUIFont"   , "Microsoft YaHei, norm s10.0",
    "OptGUIFont"    , "Microsoft YaHei, norm s9.0",
    "MainSBFont"    , "Microsoft YaHei, norm s9.0",
    "WinX"          , 600,
    "WinY"          , 400,
    "MainGUIColor"  , "0xFFFFFF",
    "CMDListColor"  , "0xFFFFFF",
    "Background"    , "Default",
    "Transparency"  , 255
)

Global g_RUNTIME := Map( ; Runtime variables, not saved to ini
    "CurrentCommand", "",
    "UseDisplay"    , 0,
    "UseFallback"   , 0,
    "Arg"           , "",
    "OneDrive"      , EnvGet("OneDrive"),
    "RegEx"         , "imS)",
    "Max"           , 1,
    "LastWin"       , 0                                                 ; Hwnd of the window that was active before ALTRun popped up (used by Clip)
)

Global g_USAGE := Map(A_YYYY . A_MM . A_DD, 1)

Global g_BENCH := Map(  ; Last Benchmark() result, purely informational
    "AvgMs"          , 0,
    "P50Ms"          , 0,
    "P95Ms"          , 0,
    "ElapsedTotalMs" , 0,
    "LastTime"       , ""
)

AppData.LoadAppData()    ; Loads (and migrates/creates) ALTRun.json; fills Config/Gui/Hotkey/Usage/History/Benchmark/commands

OnExit((p*) => AppData.OnAppExit(p*))                                           ; Flush any Usage bump / buffered log lines on Reload()/ExitApp()

CommandStore.LoadCommands()
CommandStore.LoadHistory()
Indexer.UpdateSendTo()
Indexer.UpdateStartup()
Indexer.UpdateStartMenu()
MainWindow.CreateTrayMenu()               ; Tray menu before the window: the GUI window uses the tray icon that was in effect at the time the window was created
MainWindow.Create()                       ; Create and set main GUI
MainWindow.RegisterHotkeys()
Listary.Init()
AutoDate.Init()
AutoCheckUpdate()
return
;;==================== Autorun until here =========================

; SetMainGUI()/MoveWindow()/SetTrayMenu()/SetMenuItemIcons()/RegisterHotkey()/
; RunFunctionCommand()/Activate()/Input_Change()/ListResult()/SyncCurrentCommandByRow()/
; GetIconIndex()/GetIcon()/GotoCommand()/RunSelectedCommand()/ChangeCommand()/
; LV_Click()/LV_ContextMenu()/SetListViewContextMenu()/LVRunCommand()/MainGUI_*()/
; SetMainGUIContextMenu()/SetStatusBar()/MonitorFocus()/ExtractRes() and the
; MainGUI/myInputBox/myListView/myImageList/myIconMap globals used to live here;
; all moved into the MainWindow class in Src\UI\MainWindow.ahk (see the #Include
; list at the top of this file and the MainWindow.Create() call in the autorun
; section). The six below stay bare wrappers - all of them are listed in
; FuncList for custom hotkeys (stored by name in ALTRun.json), and RunCommand()
; calls those by name via %cmdPath%(), which only resolves plain global
; function names, not Class.Method.
ToggleWindow(*) {
    MainWindow.Toggle()
}
TabFunc(*) {                                                            ; Tab only switches focus between the input box and the list
    MainWindow.ToggleFocus()
}
PrevCommand(*) {
    MainWindow.MoveSelection(-1)
}
NextCommand(*) {
    MainWindow.MoveSelection(1)
}
CopyCommand(*) {                                                        ; Ctrl+C / list context menu
    MainWindow.CopyFocusedCommand()
}
ClearInput(*) {
    MainWindow.ClearInput()
}

SearchCommand(command := "") {
    Global g_MATCHED, g_RUNTIME, g_FALLBACK, g_COMMANDS, g_CMDINDEX

    g_MATCHED := Array()
    g_RUNTIME["CurrentCommand"] := ""
    listLimit := g_GUI["ListRows"]
    prefix := SubStr(command, 1, 1)
    isExpr := Calc.Looks(command)

    ; Prefix-based fallback shortcuts: "+" / " " / ">"
    if IsFallbackPrefix(prefix) {
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
        return SearchFuncPalette(SubStr(command, 2), listLimit)

    ; Search precomputed command index.
    if (!isExpr) {
        pattern := BuildFuzzyPattern(command)
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
            evalResult := Calc.Eval(command)
            if (IsNumber(evalResult)) {
                g_RUNTIME["UseFallback"] := False
                g_RUNTIME["CurrentCommand"] := ""
                g_MATCHED := StruCalc(Round(evalResult, 6)) ; normalize float precision
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
; entries only, then applies the same fuzzy matching SearchCommand() uses for
; everything else against each one's function name + description.
SearchFuncPalette(remainder, listLimit) {
    Global g_MATCHED, g_RUNTIME, g_COMMANDS

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
        regexPattern := g_RUNTIME["RegEx"] . BuildFuzzyPattern(remainder)
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

GetCmdPart(command, fieldNo) {
    static lastCmd := "", lastParts := ""
    if (command != lastCmd) {
        lastCmd := command
        lastParts := StrSplit(command, " | ")
    }
    parts := lastParts
    return parts.Length >= fieldNo ? parts[fieldNo] : ""
}

GetCmdDisplayPath(command) {                                            ; Field 2 of a command, made readable (Clip text gets a short preview)
    return (GetCmdPart(command, 1) = "Clip")
        ? Clip.ClipPreview(GetCmdPart(command, 2))
        : GetCmdPart(command, 2)
}

; AbsPath()/RelativePath() used to live here; both are now Path.Resolve()/
; Path.Shorten() in Lib/Util.ahk. RelativePath() had no callers left in this
; file, so it's simply gone rather than moved - Path.Shorten() covers the
; same job if something needs it again.

RunCommand(originCmd) {
    if (originCmd = "")
        return

    if (g_RUNTIME["UseDisplay"]) {
        g_LOG.Debug("RunCommand: blocked in display mode, cmd=" originCmd)
        return
    }

    executed := false
    MainWindow.Hide()
    ParseArg()
    g_LOG.Debug("RunCommand: Execute request=" originCmd)

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
        executed := OpenDir(cmdPath)
    } else if (cmdType = "FUNC") {
        try {
            %cmdPath%()
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
        CommandStore.UpdateRank(originCmd)                                          ; Saves by itself only when SmartRank is on
        CommandStore.UpdateHistory(originCmd)
        AppData.SaveAppData()                                                  ; Guarantees RunCount/History persist either way, in one write
        g_LOG.Debug("RunCommand: Execute success, RunCount=" g_CONFIG["RunCount"] ", cmd=" originCmd)
    } else {
        g_LOG.Debug("RunCommand: Execute failed, cmd=" originCmd)
    }
    return
}

Exit(*) {
    ExitApp()
}

RestartApp(*) {
    Reload()
}

RunCurrentCommand(*) {
    RunCommand(g_RUNTIME["CurrentCommand"])
}

ParseArg() {
    inputVal := MainWindow.Input.Value
    commandPrefix := SubStr(inputVal, 1, 1)
    spacePos := InStr(inputVal, " ")

    if IsFallbackPrefix(commandPrefix) {
        return g_RUNTIME["Arg"] := SubStr(inputVal, 2)
    }

    if (spacePos && !g_RUNTIME["UseFallback"]) {
        g_RUNTIME["Arg"] := SubStr(inputVal, spacePos + 1)
    } else if (g_RUNTIME["UseFallback"]) {
        g_RUNTIME["Arg"] := inputVal
    } else {
        g_RUNTIME["Arg"] := ""
    }
}

BuildFuzzyPattern(needle) {
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

IsFallbackPrefix(prefix) {
    if (prefix = "")
        return false
    return InStr("+ >", prefix, 0)
}

; UpdateRank()/UpdateUsage()/UpdateRunCount()/UpdateHistory()/LoadCommands()/
; LoadHistory() used to live here; all moved into the CommandStore class in
; Src\Core\CommandStore.ahk (see the #Include list at the top of this file).
;
; RankUp()/RankDown() stay bare global functions (not CommandStore methods):
; the Options window's FuncList lets you bind them to a custom hotkey by
; storing the function name as a string in g_HOTKEY[Trigger*], and RunCommand()
; then calls it by name via %cmdPath%(), which only resolves plain global
; function names, not Class.Method.
RankUp(*) {
    CommandStore.UpdateRank(g_RUNTIME["CurrentCommand"], true)
}

RankDown(*) {
    CommandStore.UpdateRank(g_RUNTIME["CurrentCommand"], true, -1)
}

; GetCmdOutput()/GetRunResult() used to live here; GetCmdOutput() is now the
; private SystemActions._GetCmdOutput() in Src\Features\SystemActions.ahk (its only
; caller, _ShowCmdOutputInNotepad(), moved there with it). GetRunResult() was
; an unused alternative implementation ("方式2") with no callers anywhere in
; the codebase, so it was dropped rather than moved.

OpenDir(dirPath) {                                                      ; Named dirPath, not Path - that's the Util.ahk class now
    dirPath := Path.Resolve(dirPath)

    Try{
        Run(g_CONFIG["FileMgr"] ' `"' dirPath '`"')
        g_LOG.Debug("OpenDir: Using=" g_CONFIG["FileMgr"] " to open dir=" dirPath "...OK")
        return true
    } catch as e {
        g_LOG.Debug("OpenDir: Failed to open dir=" dirPath " Error=" e.Message)
        MsgBox("Could not open dir: " dirPath "`n`nError message: " e.Message, g_TITLE, 48)
        return false
    }
}

OpenContainer(*) {
    if (GetCmdPart(g_RUNTIME["CurrentCommand"], 1) = "Clip")            ; Clip has no container folder
        return
    cmdPath := GetCmdPart(g_RUNTIME["CurrentCommand"], 2)
    if (cmdPath = "") {
        return MsgBox("No valid file to open container folder.", g_TITLE, 48)
    }
    containerPath := Path.Resolve(cmdPath)                             ; Named containerPath, not Path - that's the Util.ahk class now

    try {
        runArg := (g_CONFIG["FileMgr"] = "Explorer.exe") ? ' /Select, `"' containerPath '`"' : ' /P `"' containerPath '`"' ; /P Parent folder
        Run(g_CONFIG["FileMgr"] runArg)

    g_LOG.Debug("OpenContainer: Using=" g_CONFIG["FileMgr"] " to open container dir for file=" containerPath "...OK")
    } catch as e {
        g_LOG.Debug("OpenContainer: Failed to open container dir for file=" containerPath " Error=" e.Message)
        MsgBox("Failed to open container dir for file: " . containerPath "`n`nError message: " . e.Message, g_TITLE, 48)
    }
}

; UpdateSendTo()/UpdateStartup()/UpdateStartMenu()/Reindex() used to live
; here; all moved into the Indexer class in Src\Features\Indexer.ahk (see the
; #Include list at the top of this file, and Indexer.UpdateSendTo() etc.
; called directly from the autorun section above - only Reindex() needs a
; bare wrapper below, since it's the one of the four that's a built-in Func
; command and listed in FuncList / bound to the tray and right-click menus.
Reindex(*) {
    Indexer.Rebuild()
}

About(*) {
    Options(8)
}

Usage(*) {
    Options(7)
}

; AutoCheckUpdate()/Update()/CheckUpdate()/CompareVersion() used to live here;
; all moved into the UpdateChecker class in Src\Features\UpdateChecker.ahk (see the
; #Include list at the top of this file). These three stay bare wrappers -
; Update() is bound directly as a Menu.Add() callback (tray/right-click menu)
; and listed in FuncList for custom hotkeys, and CheckUpdate() is passed bare
; to SetTimer() - none of that accepts Class.Method, only plain function names.
AutoCheckUpdate(*) {
    UpdateChecker.AutoCheck()
}
Update(*) {
    CheckUpdate(False)
}
CheckUpdate(Silent := True) {
    UpdateChecker.Check(Silent)
}

; Listary()/ShowListaryHint()/GetListaryHintText()/IsQuickSwitchDialog()/
; HasAnyCtrlMatch()/IsLikelyFileDialogTitle()/SyncTCPath()/SyncExplorerPath()/
; SetDialogPath() used to live here; all moved into the Listary class in
; Src\Features\Listary.ahk (see the #Include list at the top of this file and the
; "Listary.Init()" call in the autorun section).

UserCommand(*) {                                                        ; F4 - edit the command database directly
    Run("Notepad.exe " . g_JSON)
}

; From command "New Command" or GUI context menu "New Command"
NewCommand(*) {
    CommandManager.Open("UserCommand", , , g_RUNTIME["Arg"], 1, "")
}

; OpenCommandManager()/PickCommandTarget()/SaveCommandFromManager()/
; CloseCommandManager()/EditCommand()/DelCommand()/UndoDelCommand() used to
; live here; all moved into the CommandManager class in Src\UI\CommandManager.ahk
; (see the #Include list at the top of this file). The five below stay bare
; wrappers - all of them are listed in FuncList for custom hotkeys, and
; EditCommand/NewCommand/DelCommand are also built-in Func commands - Func-type
; dispatch and FuncList only resolve plain global function names, never
; Class.Method.
OpenCommandManager(Section := "UserCommand", Type := "File", Path := "", Desc := "", Rank := 1, OriginCmd := "") {
    CommandManager.Open(Section, Type, Path, Desc, Rank, OriginCmd)
}
EditCommand(*) {
    CommandManager.Edit()
}
DelCommand(*) {
    CommandManager.Delete()
}
UndoDelCommand(*) {
    CommandManager.UndoDelete()
}

; The old Plugins()/RenameWithDate()/LineEndAddDate()/NameAddDate() used to live here;
; all moved into the AutoDate class in Src\Features\AutoDate.ahk (see the #Include list
; at the top of this file and the "AutoDate.Init()" call in the autorun section).

; GetArrayIndex() used to live here; it's now Arr.IndexOf() in Lib\Util.ahk -
; moved there instead of into CommandManager since it's also used by
; Src\UI\OptionsWindow.ahk for the custom-hotkey trigger dropdowns.

; LoadAppData()/SaveAppData()/MergeIntoDefaults()/OnAppExit()/ParseCommandBlock()/
; DefaultCommandText()/UserCommandText()/FallbackCommandText() used to live
; here; all moved into the AppData class in Src\Core\AppData.ahk (called as
; AppData.LoadAppData() etc. - see the #Include list at the top of this file
; and the comment block at the top of that file for the ALTRun.json layout).
; ALTRun.ini and its migration code (MigrateFromIni/ReadIniMapLikeDefaults/
; FinishIniMigration/JoinArray/ReadIniSectionRaw/UnescapeCommandKey/g_SECTION/
; g_INI) are retired - this was always a single-machine, single-user install,
; so once the one-off migration off the ini ran there was no reason to keep
; that code around for a scenario that will never come up again.

; =========================================================================
; === Clip (Snippet) support ==============================================
; EscapeClipText()/UnescapeClipText()/ClipPreview()/ExpandClipPlaceholders()/
; FocusLastWindow()/PasteClipText()/EditClipText()/CloseClipEditor() used to
; live here; all moved into the Clip class in Src\Features\Clip.ahk (see the #Include
; list at the top of this file and the comment block at the top of that file
; for the command format / placeholder syntax).
; =========================================================================

; NewClip() must stay a bare global function (not a Clip class method): the
; built-in "Func | NewClip | New Clip (text snippet)" command calls it by
; name via %cmdPath%() in RunCommand(), which only resolves plain global
; function names, not Class.Method - see FuncList in Options()/DefaultCommandText().
NewClip(*) {                                                            ; Command "New Clip", opens the manager pre-set to type Clip
    OpenCommandManager("UserCommand", "Clip", Clip.EscapeClipText(g_RUNTIME["Arg"]), "", 1, "")
}

; One-shot clipboard text transforms (see Src\Features\Clip.ahk) - each is a bare wrapper
; for the same reason NewClip() above is: Func-type dispatch only resolves
; plain global function names, never Class.Method.
ClipUpper() {
    Clip.ToUpper()
}
ClipLower() {
    Clip.ToLower()
}
ClipTitleCase() {
    Clip.ToTitleCase()
}
ClipReverse() {
    Clip.Reverse()
}
ClipSortAsc() {
    Clip.SortAsc()
}
ClipSortDesc() {
    Clip.SortDesc()
}
ClipTrimLines() {
    Clip.TrimLines()
}
ClipRemoveBlankLines() {
    Clip.RemoveBlankLines()
}
ClipDedupeLines() {
    Clip.DedupeLines()
}
ClipToTraditional() {
    Clip.ToTraditional()
}
ClipToSimplified() {
    Clip.ToSimplified()
}

; Opens a cmd.exe window at whatever folder Total Commander/Explorer was
; browsing right before ALTRun was invoked - reuses the same read-only path
; detection Listary.ahk already has for the dialog-box quick-switch feature.
OpenTerminalHere() {
    hwnd := g_RUNTIME["LastWin"]
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return MsgBox(g_LNG[840], g_TITLE, 48)

    winClass := WinGetClass("ahk_id " hwnd)
    if (winClass = "TTOTAL_CMD")
        path := Listary.TCCurrentPath()
    else if (winClass = "CabinetWClass")
        path := Listary.ExplorerCurrentPath()
    else
        path := ""

    if (path = "" || !FileExist(path))
        return MsgBox(g_LNG[840], g_TITLE, 48)

    try {
        Run(A_ComSpec, path)
    } catch as e {
        MsgBox("Could not open a terminal at: " path "`n`n" e.Message, g_TITLE, 48)
    }
}

PTTools() {
    PTToolsWindow.Show()                                                  ; Src\UI\PTToolsWindow.ahk - Rebar/BRC calculator
}

SPF2M() {
    PTToolsWindow.ShowSpf2m()                                             ; Src\UI\PTToolsWindow.ahk - SPF2M profile calculator automation
}

StruCalc(evalResult) {
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

; Options()/ResetHotkey()/SelectFont()/PickCMDListColor()/PickMainGUIColor()/
; SelectBackground()/OPTButtonOK()/OPTGuiClose()/ToggleGlobalHotkeys()/
; SaveConfig()/GetOptCtrlValue()/CoerceLikeCurrent() used to live here; all
; moved into the OptionsWindow class in Src\UI\OptionsWindow.ahk (see the
; #Include list at the top of this file).

; Options() must stay a bare global function (not an OptionsWindow class
; method): F2 / the tray menu / the right-click menu all call it directly,
; and "Func | Options | ..." plus the custom-hotkey FuncList in
; OptionsWindow.Show() call it by name via %cmdPath%() in RunCommand(),
; which only resolves plain global function names, not Class.Method.
Options(ActTab := 1) {
    OptionsWindow.Show(ActTab)
}

; ==================== Built-in Functions =========================
; AhkRun()/TurnMonitorOff()/EmptyRecycle()/MuteVolume()/Google()/Bing()/
; Everything()/Baidu()/Taobao()/JD()/ShowIP()/UrlEncode()/Logoff()/
; ShutdownMachine()/RestartMachine()/HibernateMachine()/IncreaseVolume()/
; DecreaseVolume()/ListProcess()/ListService() used to live here; all moved
; into the SystemActions class in Src\Features\SystemActions.ahk (see the #Include
; list at the top of this file). Each stays a bare wrapper below for the
; same reason NewClip()/PTTools() etc. do - Func-type dispatch and FuncList
; custom hotkeys only resolve plain global function names, never Class.Method.
AhkRun() {
    SystemActions.AhkRun()
}
TurnMonitorOff() {
    SystemActions.TurnMonitorOff()
}
EmptyRecycle() {
    SystemActions.EmptyRecycle()
}
MuteVolume() {
    SystemActions.MuteVolume()
}
Google() {
    SystemActions.Google()
}
Bing() {
    SystemActions.Bing()
}
Everything() {
    SystemActions.Everything()
}
Baidu() {
    SystemActions.Baidu()
}
Taobao() {
    SystemActions.Taobao()
}
JD() {
    SystemActions.JD()
}
ShowIP() {
    SystemActions.ShowIP()
}
UrlEncode() {
    SystemActions.UrlEncode()
}
Logoff() {
    SystemActions.Logoff()
}
ShutdownMachine() {
    SystemActions.ShutdownMachine()
}
RestartMachine() {
    SystemActions.RestartMachine()
}
HibernateMachine() {
    SystemActions.HibernateMachine()
}
IncreaseVolume() {
    SystemActions.IncreaseVolume()
}
DecreaseVolume() {
    SystemActions.DecreaseVolume()
}
ListProcess() {
    SystemActions.ListProcess()
}
ListService() {
    SystemActions.ListService()
}

; SetLanguage()/ReadChineseFlag() used to live here; both are now Language.Load()/
; Language.IsChinese() in Src\Core\Language.ahk (see the #Include list at the top of
; this file and the "Global g_LNG := Language.Load()" call near the top).
; Eval()/EvalSimple() used to live here; both are now Calc.Eval() in Lib/Util.ahk.

;;==================== Performance Test Only =========================
; Not wired to any hotkey/menu/command - run BenchmarkRun() manually from an
; editor/debugger when you want a search-performance snapshot.

BenchmarkRun(rounds := 10) {
    Global g_LOG, g_COMMANDS
    rounds := Max(10, rounds)

    static queries := [
        "n", "no", "note", "core", "kanji", "a t", "wi ex", "sys in"
        , "plugin", "open", "update", "help", "xyz_not_found"
        , "2+3*5", "5+5", "12345*6789", "nir", "control panel"
        , "new", "edit", "delete", "reload", "reindex", "history", "option", "usage"
        , "google test", "bing test", "notepad", "explorer", "cmd", "powershell", "service"
        , "disk", "device", "event", "task", "reg", "calc", "paint", "startup"
    ]

    MainWindow.Show()
    MainWindow.Input.Focus()

    for _, q in queries
    {
        MainWindow.Input.Value := q
        Sleep(10)
        SearchCommand(q)
    }

    sampleLines := ""
    sampleCount := 0
    sumMs := 0.0
    minMs := 999999.0
    maxMs := 0.0
    totalStartMs := A_TickCount
    Loop rounds {
        for _, q in queries {
            MainWindow.Input.Value := q
            Sleep(10)
            t0 := A_TickCount
            SearchCommand(q)
            ms := A_TickCount - t0
            sampleCount += 1
            sumMs += ms
            minMs := Min(minMs, ms)
            maxMs := Max(maxMs, ms)
            sampleLines .= ms "`n"
        }
    }
    elapsedTotalMs := A_TickCount - totalStartMs

    avgMs := Round(sumMs / sampleCount, 3)
    minMs := Round(minMs, 3)
    maxMs := Round(maxMs, 3)

    sortedLines := Sort(sampleLines, "N")
    p50Idx := Ceil(sampleCount * 0.50)
    p95Idx := Ceil(sampleCount * 0.95)
    p50Ms := 0.0
    p95Ms := 0.0
    i := 0
    for _, line in StrSplit(sortedLines, "`n", "`r") {
        if (line = "")
            continue
        i += 1
        if (i = p50Idx)
            p50Ms := line + 0
        if (i = p95Idx) {
            p95Ms := line + 0
            break
        }
    }
    p50Ms := Round(p50Ms, 3)
    p95Ms := Round(p95Ms, 3)

    prevElapsed := g_BENCH["ElapsedTotalMs"]
    if (prevElapsed = "" || prevElapsed = 0) {
        deltaText := "N/A (first run)"
    } else {
        d := Round(elapsedTotalMs - prevElapsed, 2)
        p := (prevElapsed != 0) ? Round((d / prevElapsed) * 100, 2) : 0
        deltaText := (d > 0 ? "+" : "") d " ms (" (p > 0 ? "+" : "") p "%) - " (d < 0 ? "Faster" : d > 0 ? "Slower" : "No change")
    }

    nowText := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    g_LOG.Debug("Benchmark: Time=" nowText
        . ", AvgMs=" avgMs
        . ", P50Ms=" p50Ms
        . ", P95Ms=" p95Ms
        . ", MinMs=" minMs
        . ", MaxMs=" maxMs
        . ", Ops=" sampleCount
        . ", Rounds=" rounds
        . ", QueriesPerRound=" queries.Length
        . ", Commands=" g_COMMANDS.Length
        . ", ElapsedTotalMs=" Round(elapsedTotalMs, 2))

    g_BENCH["AvgMs"] := avgMs
    g_BENCH["P50Ms"] := p50Ms
    g_BENCH["P95Ms"] := p95Ms
    g_BENCH["ElapsedTotalMs"] := Round(elapsedTotalMs, 2)
    g_BENCH["LastTime"] := nowText
    AppData.SaveAppData()

    report := "ALTRun Benchmark`n`n"
    report .= "Time: " nowText "`n"
    report .= "Commands: " g_COMMANDS.Length "`n"
    report .= "Rounds: " rounds ", Queries/Round: " queries.Length ", Ops: " sampleCount "`n`n"
    report .= "Avg: " avgMs " ms/query`n"
    report .= "P50: " p50Ms " ms/query`n"
    report .= "P95: " p95Ms " ms/query`n"
    report .= "Min/Max: " minMs " / " maxMs " ms`n"
    report .= "Total elapsed: " Round(elapsedTotalMs, 2) " ms`n`n"
    report .= "Vs last Total elapsed: " deltaText
    MsgBox(report, "ALTRun Benchmark")
}

; FontSelect()/ColorSelect()/ColorSwapRGBBGR()/ColorHex() used to live here;
; they're now FontDialog.Choose()/ColorDialog.Choose()/ColorDialog.RgbBgr()/
; ColorDialog.Hex() in Lib/Dialogs.ahk.

; GetFirstChar() used to live here; it's now Pinyin.Initials() in Lib/Util.ahk.

