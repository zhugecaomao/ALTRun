;===================================================
; ALTRun - An effective launcher for Windows
; https://github.com/zhugecaomao/ALTRun
;===================================================
#Requires AutoHotkey v2.0
#SingleInstance Force
#NoTrayIcon
#Warn All, OutputDebug
#Include Lib\JSON.ahk                                                  ; JSON.parse()/JSON.stringify() - used by the ALTRun.json store below.
#Include Lib\Logger.ahk                                                ; g_LOG - see just below.
#Include Lib\Util.ahk                                                  ; Path / Fonts / Win / Pinyin / Calc - see each call site below.
#Include Lib\Dialogs.ahk                                               ; FontDialog / ColorDialog - see the Options-window font/color pickers.
#Include Lib\Language.ahk                                              ; Lang.Load()/Lang.IsChinese() - builds g_LNG (the UI text table) below.
#Include Lib\Listary.ahk                                               ; Listary.Init() - open/save dialog path quick-switch, called in the autorun section below.
#Include Lib\Plugins.ahk                                               ; Plugins.Init() - Ctrl+D auto-date plugin, called in the autorun section below.
#Include Lib\Clip.ahk                                                  ; Clip.PasteClipText()/ClipPreview()/EditClipText() - the "Clip" snippet command.
#Include Lib\IniMigration.ahk                                          ; IniMigration.MigrateFromIni() - one-off ALTRun.ini -> ALTRun.json conversion.
#Include Lib\AppData.ahk                                               ; AppData.LoadAppData()/AppData.SaveAppData() - reads/writes ALTRun.json.
#Include Lib\CommandStore.ahk                                          ; CommandStore.LoadCommands() etc. - in-memory command cache/rank/usage/history.
#Include Lib\PTTools.ahk                                               ; PTToolsWindow - Rebar/BRC calculator + SPF2M automation (see PTTools() below).
                                                                         ; All explicit: auto-include only reliably covers ClassName(...)
                                                                         ; construction calls, not ClassName.Method(...) static calls like
                                                                         ; JSON.parse(), so relying on it for every Lib class is asking for
                                                                         ; the same "unassigned variable" failure JSON.ahk hit without this.
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
Global g_HISTORYS := Array()         ; Execution history
Global g_MATCHED  := Array()         ; Matched commands

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
Global g_LNG := Lang.Load()                                            ; UI text table (English/Chinese), see Lib\Language.ahk

Global g_CONFIG_P1 := Map(
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

; Global variables which are only read by the function, not assigned or used with the reference operator (&).
Global MainGUI
Global myListView
Global myInputBox
Global OptGUI
Global OptListView
Global g_CmdMgrGui
Global g_ClipEditGui
Global myImageList := IL_Create(10, 5, g_CONFIG["LargeIcons"])          ; Create an ImageList so that the ListView can display some icons, 3rd param is 1: large icons, 0: small icons
Global myIconMap   := Map("DIR", IL_Add(myImageList,"imageres.dll",-3)  ; Icon cache index, IconIndex=1/2/3/4 for type dir/func/url/eval/cmd
                        ,"FUNC", IL_Add(myImageList,"imageres.dll",-100)
                        ,"URL" , IL_Add(myImageList,"imageres.dll",-144)
                        ,"EVAL", IL_Add(myImageList,"imageres.dll",-182)
                        ,"CMD" , IL_Add(myImageList,"imageres.dll",-100)
                        ,"CLIP", IL_Add(myImageList,"imageres.dll",-102)) ; "imageres.dll",-5323 is cmd.exe icon

OnExit((p*) => AppData.OnAppExit(p*))                                           ; Flush any Usage bump / buffered log lines on Reload()/ExitApp()

CommandStore.LoadCommands()
CommandStore.LoadHistory()
UpdateSendTo()
UpdateStartup()
UpdateStartMenu()
SetTrayMenu()               ; SetTrayMenu before SetMainGUI, GUI window uses the tray icon that was in effect at the time the window was created
SetMainGUI()                ; Create and set main GUI
RegisterHotkey()
Listary.Init()
Plugins.Init()
AutoCheckUpdate()
return
;;==================== Autorun until here =========================

SetMainGUI() {
    Global MainGUI, myInputBox, myListView, myStatus, myImageList

    Run_W   := g_CONFIG["ShowBtnRun"] * 80
    Run_X   := g_CONFIG["ShowBtnRun"] * 10
    Run_H   := !g_CONFIG["ShowBtnRun"]
    Opt_W   := g_CONFIG["ShowBtnOpt"] * 80
    Opt_X   := g_CONFIG["ShowBtnOpt"] * 10
    Opt_H   := !g_CONFIG["ShowBtnOpt"]
    List_W  := g_GUI["WinX"] - 24
    List_H  := g_GUI["WinY"] - 95
    Input_W := List_W - Run_X - Run_W - Opt_X - Opt_W
    TopMost := g_CONFIG["AlwaysOnTop"] ? "AlwaysOnTop" : ""
    Caption := g_CONFIG["ShowCaption"] ? "" : " -Caption"
    Theme   := g_CONFIG["XPthemeBg"] ? "" : " -Theme"
    DBuffer := g_CONFIG["DoubleBuffer"] ? " +LV0x10000" : ""
    Header  := g_CONFIG["ShowHdr"] ? "" : " -Hdr"
    Grid    := g_CONFIG["ShowGrid"] ? " Grid" : ""
    Border  := g_CONFIG["ShowBorder"] ? "" : " -E0x200"

    MainGUI := Gui(TopMost Caption Theme " +MinSize300x160", g_TITLE)
    MainGUI.OnEvent("Close" , MainGUI_Close)
    MainGUI.OnEvent("Escape", MainGUI_Escape)
    MainGUI.OnEvent("Size"  , MainGUI_Size)
    MainGUI.OnEvent("ContextMenu", MainGUI_ContextMenu)
    MainGUI.BackColor := g_GUI["MainGUIColor"]
    mainGuiFont := Fonts.Spec(g_GUI["MainGUIFont"], "Microsoft YaHei", "norm s10.0")
    MainGUI.SetFont(mainGuiFont.opt, mainGuiFont.name)
    myInputBox := MainGUI.AddEdit("x12 y10 r1 -WantReturn border -E0x200 W" Input_W, g_LNG[13])
    myInputBox.Opt("Background" g_GUI["CMDListColor"])
    myInputBox.OnEvent("Change", Input_Change)
    myRunBtn := MainGUI.AddButton("x+" Run_X " yp W" Run_W " hp Default Hidden" Run_H, g_LNG[11])
    myRunBtn.OnEvent("Click", RunCurrentCommand)
    myOptBtn := MainGUI.AddButton("x+" Opt_X " yp W" Opt_W " hp Hidden" Opt_H, g_LNG[12])
    myOptBtn.OnEvent("Click", (*) => Options())
    myListView := MainGUI.AddListView("x12 yp+36 W" List_W " H" List_H " -Multi", g_LNG[10])
    myListView.Opt(DBuffer Header Grid Border " Background" g_GUI["CMDListColor"] " +Report") ; ListView View Modes: Report, Icon, Tile, IconSmall, List
    myListView.OnEvent("Click", LV_Click)
    myListView.OnEvent("ContextMenu", LV_ContextMenu)
    myListView.OnEvent("DoubleClick", LVRunCommand)
    if (g_CONFIG["LargeIcons"]) {
        myListView.SetImageList(myImageList, 1)                         ; Attach the ImageList to the ListView, 2nd param is 1: large icons, 0: small icons, 2: state icons (AHK Doc incorrect)
    } else {
        myListView.SetImageList(myImageList)
    }

    colWidths := StrSplit(g_GUI["ColWidth"], ",")
    Loop 4 {
        if (colWidths.Length >= A_Index) {
            if (colWidths[A_Index] != "")
                myListView.ModifyCol(A_Index, colWidths[A_Index])
        }
    }

    myStatus := MainGUI.AddEdit("x12 y+10 r1 -WantReturn ReadOnly -E0x200 border W" List_W " Hidden" (!g_CONFIG["ShowStatusBar"]), )
    myStatus.Opt("Background" g_GUI["CMDListColor"])
    mainSbFont := Fonts.Spec(g_GUI["MainSBFont"], "Microsoft YaHei", "norm s9.0")
    myStatus.SetFont(mainSbFont.opt, mainSbFont.name)

    if FileExist(Path.Resolve(g_GUI["Background"])) {
        try MainGUI.AddPic("x0 y0 0x4000000", Path.Resolve(g_GUI["Background"]))
    } else if (g_GUI["Background"] = "Default") {
        try MainGUI.AddPic("x0 y0 0x4000000", ExtractRes())
    }

    ListResult(g_LNG[50])

    ;===================================================
    ; Resolve command line arguments, A_Args[1] A_Args[2]
    ;===================================================
    local HideWin := ""
    for value in A_Args{
        g_LOG.Debug("Resolving command line args" A_Index " = " value)
        if (A_Index = 1) {
            if (value = "-Startup" || value = "-StartMenu")
                HideWin := "Hide "

            if (value = "-SendTo" && A_Args.Length >= 2) {
                HideWin := "Hide "
                sendToPath := A_Args[2]                                 ; Not "Path" - that's the Util.ahk class now, and this whole
                                                                         ; function also calls Path.Resolve() earlier (see below);
                                                                         ; a local assigned anywhere in a function shadows the
                                                                         ; global/class of the same name for the WHOLE function.

                SplitPath(sendToPath, &Desc, , &fileExt)                ; Extra name from _Path (if _Type is dir and has "." in path, nameNoExt will not get full folder name)

                fileType := InStr(FileExist(sendToPath), "D") ? "Dir" : "File" ; Default Type is File, Set Type is Dir only if the file exists and is a directory

                if (fileExt = "lnk" && g_CONFIG["SendToGetLnk"]) {
                    FileGetShortcut(sendToPath, sendToPath, , &fileArg, &Desc)
                    sendToPath .= " " fileArg
                }
                OpenCommandManager("UserCommand", fileType, sendToPath, Desc, 1, "")   ; Add new command to database
            }
        }
    }

    if (g_GUI["Transparency"] < 250){
        WinSetTransparent(g_GUI["Transparency"], MainGUI.Hwnd)          ; By default, hidden windows are not detected. however, when using pure HWNDs, hidden windows are always detected regardless of DetectHiddenWindows.
    }

    Win.SetCorner(MainGUI.Hwnd, g_CONFIG["RoundCorner"])

    MainGUI.Show(HideWin "w" g_GUI["WinX"] " h" g_GUI["WinY"] " Center")

    ; Enable dragging for captionless window
    OnMessage(0x201, MoveWindow)

    if g_CONFIG["HideOnLostFocus"] {
        ; 方案 1 - OnMessage(0x0006, WM_ACTIVATE)
        ; 事件驱动, 高效无延迟, 无资源占用
        ; 某些情况有窗口闪烁和托盘菜单右键点击"显示"窗口闪退问题
        ; 方案 2 - SetTimer(MonitorFocus, 30)
        ; Ahk原生方式, 代码简单可靠, 但有轻微性能开销, 有稍许延迟 30ms
        ; 方案 3 - Control.OnEvent("LoseFocus", MonitorFocus)
        ; Ahk原生方式, 事件驱动, 高效无延迟, 无资源占用
        ; 因 Gui 本身没有 LoseFocus 事件, 需要注册主界面所有控件的 LoseFocus 事件
        ; 修复托盘菜单右键点击"显示"窗口闪退问题
        myInputBox.OnEvent("LoseFocus", MonitorFocus)
        myListView.OnEvent("LoseFocus", MonitorFocus)
        myRunBtn.OnEvent("LoseFocus", MonitorFocus)
        myOptBtn.OnEvent("LoseFocus", MonitorFocus)
        myStatus.OnEvent("LoseFocus", MonitorFocus)
    }
    return
}

MoveWindow(wParam, lParam, msg, hwnd) {                                                         ; Allow moving a captionless window by mouse-drag
    if (hwnd != MainGUI.Hwnd)
        return
    DllCall("ReleaseCapture")
    SendMessage(0xA1, 2, 0, MainGUI.Hwnd)  ; WM_NCLBUTTONDOWN, HTCAPTION
}

; 创建任务栏托盘程序图标
SetTrayMenu() {
    if !g_CONFIG["ShowTrayIcon"] {
        A_IconHidden := 1
        return
    }

    static myTrayMenu := ""  ; 只在第一次调用时创建

    if !IsObject(myTrayMenu) {
        myTrayMenu := A_TrayMenu
        try {
            TraySetIcon("imageres.dll", -100)
            
            myTrayMenu.Delete() ; 删除默认项
            myTrayMenu.Add(g_LNG[300], ToggleWindow)
            myTrayMenu.Add(g_LNG[301], (*) => Options())
            myTrayMenu.Add(g_LNG[310], UserCommand)
            myTrayMenu.Add()  ; 分隔线
            myTrayMenu.Add(g_LNG[302], Reindex)
            myTrayMenu.Add(g_LNG[303], Usage)
            myTrayMenu.Add(g_LNG[305], (*) => ListLines())
            myTrayMenu.Add()  ; 分隔线
            myTrayMenu.Add(g_LNG[309], Update)
            myTrayMenu.Add(g_LNG[304], About)
            myTrayMenu.Add()
            myTrayMenu.Add(g_LNG[307], RestartApp)
            myTrayMenu.Add(g_LNG[308], Exit)

            SetMenuItemIcons(myTrayMenu)

            myTrayMenu.Default    := g_LNG[300]
            myTrayMenu.ClickCount := 1
            A_IconTip             := g_TITLE
            A_IconHidden          := 0

            g_LOG.Debug("SetTrayMenu: Create myTrayMenu...OK")
        } catch as e {
            g_LOG.Debug("SetTrayMenu: Error creating myTrayMenu: " . e.Message)
        }
    }
    return
}

SetMenuItemIcons(menu, isContext := false) {
    if !isContext {
        menu.SetIcon(g_LNG[300], "imageres.dll", -100)
        menu.SetIcon(g_LNG[301], "imageres.dll", -114)
    } else {
        menu.SetIcon(g_LNG[301], "imageres.dll", -114)
    }
    menu.SetIcon(g_LNG[310], "imageres.dll", -88)
    menu.SetIcon(g_LNG[302], "imageres.dll", -8)
    menu.SetIcon(g_LNG[303], "imageres.dll", -150)
    menu.SetIcon(g_LNG[305], "imageres.dll", -165)
    menu.SetIcon(g_LNG[309], "imageres.dll", -5338)
    menu.SetIcon(g_LNG[304], "imageres.dll", -81)
    menu.SetIcon(g_LNG[307], "imageres.dll", -5311)
    menu.SetIcon(g_LNG[308], "imageres.dll", -98)
}

RegisterHotkey() {
    ; 注册全局热键
    HotIfWinActive
    try {
        Hotkey(g_HOTKEY["GlobalHotkey1"], ToggleWindow)
        Hotkey(g_HOTKEY["GlobalHotkey2"], ToggleWindow)

        g_LOG.Debug("RegisterHotkey: Set global activate hotkeys...OK")
    } catch as e {
        g_LOG.Debug("RegisterHotkey: Failed to set global activate hotkeys..." e.Message)
    }

    ; 注册主窗口热键, 使用 ahk_id Hwnd 增强可靠性
    HotIfWinActive("ahk_id " MainGUI.Hwnd)
    try {
        Hotkey("Tab"        , TabFunc)
        Hotkey("F1"         , About)
        Hotkey("F2"         , Options)
        Hotkey("F3"         , EditCommand)
        Hotkey("F4"         , UserCommand)
        Hotkey("^d"         , OpenContainer)
        Hotkey("^c"         , CopyCommand)
        Hotkey("^n"         , NewCommand)
        Hotkey("^Del"       , DelCommand)
        Hotkey("Down"       , NextCommand)
        Hotkey("Up"         , PrevCommand)
        Hotkey("^NumpadAdd" , RankUp)
        Hotkey("^NumpadSub" , RankDown)

        if (g_CONFIG["MidScrollSwitch"]) {
            Hotkey("WheelDown"  , NextCommand)
            Hotkey("WheelUp"    , PrevCommand)
        }

        if (g_CONFIG["MidClickRun"]) {
            Hotkey("MButton"    , RunCurrentCommand)
        }

        if (g_CONFIG["SpaceToRun"]) {
            Hotkey("Space"      , RunCurrentCommand)
        }

        g_LOG.Debug("RegisterHotkey: Set local hotkeys (F1-F4)...OK")
    } catch as e {
        g_LOG.Debug("RegisterHotkey: Failed to set local hotkeys (F1-F4)..." e.Message)
    }
    
    Loop g_GUI["ListRows"] {
        try {
            Hotkey("!" . A_Index, RunSelectedCommand)                   ; 通过热键选择并运行指定命令 = Alt + index (1-9)
            Hotkey("^" . A_Index, GotoCommand)                          ; 通过热键选择指定命令 = Ctrl + index (1-9)

            g_LOG.Debug("RegisterHotkey: Set command list local hotkey " A_Index "...OK")
        } catch as e {
            g_LOG.Debug("RegisterHotkey: Failed to set command list local hotkey " A_Index . e.Message)
        }
    }

    Loop 7 {
        KeyName := "Hotkey"  . A_Index
        Trigger := "Trigger" . A_Index
        if (g_HOTKEY.Has(KeyName) && g_HOTKEY[KeyName] != "" && g_HOTKEY[Trigger] != "") { ; 自定义热键执行指定功能 = Hotkey + Trigger
            try {
                Hotkey(g_HOTKEY[KeyName], RunFunctionCommand.Bind(, g_HOTKEY[Trigger], A_Index))
                g_LOG.Debug("RegisterHotkey: Set customized function list local hotkey " A_Index " " g_HOTKEY[KeyName] " <-> " g_HOTKEY[Trigger] "...OK")
            } catch as e {
                g_LOG.Debug("RegisterHotkey: Failed to set customized function list local hotkey..."  e.Message)
            }
        }
    }

    ; 注册条件热键, 执行指定功能
    HotIfWinActive(g_HOTKEY["CondTitle"])
    if g_HOTKEY.Has("CondTitle") && g_HOTKEY.Has("CondHotkey") && g_HOTKEY.Has("CondAction") {
        try {
            Hotkey(g_HOTKEY["CondHotkey"], RunFunctionCommand.Bind(, g_HOTKEY["CondAction"], 1))
            g_LOG.Debug("RegisterHotkey: Set conditional hotkey " g_HOTKEY["CondHotkey"] " <-> " g_HOTKEY["CondAction"] " for " g_HOTKEY["CondTitle"] "...OK")
        } catch as e {
            g_LOG.Debug("RegisterHotkey: Failed to set conditional hotkey..." e.Message)
        }
    }
    HotIfWinActive                                                      ; Turn off context, make subsequent hotkeys global again
    return
}

RunFunctionCommand(HotkeyName, FuncName, Index) {
    RunCommand("FUNC | " FuncName)
    g_LOG.Debug("RunFunctionCommand: Execute function...=" FuncName)
}

Activate() {
    ; Remember the window that is active right now, so a Clip command knows where to paste.
    try {
        activeHwnd := WinExist("A")
        if (activeHwnd && activeHwnd != MainGUI.Hwnd)
            g_RUNTIME["LastWin"] := activeHwnd
    }

    MainGUI.Show()

    if (WinWaitActive("ahk_id " MainGUI.Hwnd, , 3)) {                   ; Wait for the window to be active, ahk_id is more reliable than g_TITLE
        if (g_CONFIG["AutoEngIME"]) {
            Win.SwitchToEnglishIME()
        }
        myInputBox.Focus()
        SendMessage(0xB1, 0, -1, myInputBox.Hwnd)                       ; EM_SETSEL (0xB1)
    }
}

ToggleWindow(*) {
    WinActive("ahk_id " MainGUI.Hwnd) ? MainGUI_Close() : Activate()
}

Input_Change(*) {
    SearchCommand(myInputBox.Value)
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
            return ListResult(g_MATCHED)
        fallbackIndex := (prefix = "+") ? 1 : (prefix = " ") ? 2 : 3
        g_RUNTIME["CurrentCommand"] := g_FALLBACK[Min(fallbackIndex, g_FALLBACK.Length)]
        g_MATCHED.Push(g_RUNTIME["CurrentCommand"])
        return ListResult(g_MATCHED)
    }

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
                return ListResult(g_MATCHED, True)
            }
        }

        g_RUNTIME["UseFallback"] := True
        g_MATCHED := g_FALLBACK
        g_RUNTIME["CurrentCommand"] := g_FALLBACK.Length ? g_FALLBACK[1] : ""
    }

    return ListResult(g_MATCHED)
}

ListResult(rows := [], useDisplay := false) {
    myListView.Opt("-Redraw")
    myListView.Delete()
    g_RUNTIME["UseDisplay"] := useDisplay
    showSN := g_CONFIG["ShowSN"], shortenPath := g_CONFIG["ShortenPath"]

    for rowIndex, rowCommand in rows {
        parts := StrSplit(rowCommand, " | ")
        cmdType := parts.Length >= 1 ? parts[1] : ""
        cmdPath := parts.Length >= 2 ? parts[2] : ""
        cmdDesc := parts.Length >= 3 ? parts[3] : ""
        displayNo := showSN ? rowIndex : ""
        iconIndex := GetIconIndex(cmdPath, cmdType)

        if (cmdType = "Clip") {
            ; Clip text is not a path, show a single-line preview instead.
            cmdPath := Clip.ClipPreview(cmdPath)
        } else if (shortenPath && cmdType != "URL") {
            ; Keep URL full text, shorten other command paths for list readability.
            SplitPath(cmdPath, &cmdPath)
        }

        myListView.Add("Icon" iconIndex, displayNo, cmdType, cmdPath, cmdDesc)
    }
    rowCount := myListView.GetCount()
    statusBarText := (g_RUNTIME["CurrentCommand"] != "")
        ? GetCmdDisplayPath(g_RUNTIME["CurrentCommand"])
        : (rowCount ? myListView.GetText(1, 3) : "")

    if (rowCount) myListView.Modify(1, "Select Focus Vis")
    myListView.Opt("+Redraw")
    SetStatusBar(statusBarText)
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

SyncCurrentCommandByRow(rowNumber, updateStatus := true) {
    if (g_MATCHED.Length >= rowNumber) {
        g_RUNTIME["CurrentCommand"] := g_MATCHED[rowNumber]
        if updateStatus
            SetStatusBar(GetCmdDisplayPath(g_RUNTIME["CurrentCommand"]))
        return true
    }
    if updateStatus
        SetStatusBar(myListView.GetText(rowNumber, 3))
    return false
}

GetIconIndex(filePath, type) {                                          ; Get file's icon index (TO-DO: Prepare to omit the file type)
                                                                         ; Named filePath, not path - "path" and "Path" (the Util.ahk
                                                                         ; class) are the same identifier to AHK, case-insensitive.
    Global myIconMap
    if not g_CONFIG["ShowIcon"]                                         ; ShowIcon disabled, return 0
        Return 0

    if (type = "") {
        return 0
    } else if (type = "DIR") {
        return 1
    } else if InStr("FUNC,TIP,提示,CMD", type, 0) {
        return 2
    } else if (type = "URL") {
        return 3
    } else if (type = "EVAL") {
        return 4
    } else if (type = "Clip") {
        return myIconMap.Has("CLIP") ? myIconMap["CLIP"] : 2
    } else if (type = "FILE") {
        filePath := Path.Resolve(filePath)                              ; Must store in var for afterward use, trim space (in Path.Resolve)
        SplitPath(filePath, , , &fileExt)                               ; Get the file's extension.
        if (fileExt ~= "^(?i:EXE|ICO|ANI|CUR|LNK)$") {                  ; File types that have their own icon
            IconIndex := myIconMap.Has(filePath) ? myIconMap[filePath] : GetIcon(filePath, filePath) ; File path exist in ImageList, get the index, several calls can be avoided and performance is greatly improved
        } else {                                                        ; Some other extension/file-type like pdf or xlsx
            IconIndex := myIconMap.Has(fileExt) ? myIconMap[fileExt] : GetIcon(filePath, fileExt)
        }
        Return IconIndex
    } else if (type = "App") {
        Return 2
    }
}

GetIcon(path, ExtOrPath) {                                             ; Get file's icon
    Global myImageList, myIconMap
    sfi_size := A_PtrSize + 688
    sfi      := Buffer(sfi_size)                                        ; Calculate buffer size required for SHFILEINFO structure. VarSetStrCapacity change to Buffer
    iconSize := g_CONFIG["LargeIcons"] ? 0x100 : 0x101                 ; 0x100 is SHGFI_ICON+SHGFI_LARGEICON, 0x101 is SHGFI_ICON+SHGFI_SMALLICON

    try {
        if not DllCall("Shell32\SHGetFileInfoW", "Str", path, "UInt", 0, "Ptr", sfi, "UInt", sfi_size, "UInt", iconSize)
            IconIndex := 2                                                  ; Use default function icon instead of out-of-bounds index
        else {                                                              ; Icon successfully loaded. Extract the hIcon member from the structure
            hIcon := NumGet(sfi, 0, "Ptr")                                  ; Add the HICON directly to the small-icon lists.
            IconIndex := DllCall("ImageList_ReplaceIcon", "ptr", myImageList, "int", -1, "ptr", hIcon) + 1 ; Uses +1 to convert the returned index from zero-based to one-based:
            DllCall("DestroyIcon", "Ptr", hIcon)                            ; Now that it's been copied into the ImageLists, the original should be destroyed
            myIconMap[ExtOrPath] := IconIndex                               ; Cache the icon based on file type (xlsx, pdf) or path (exe, lnk) to save memory and improve loading performance
        }
    } catch as e {
        IconIndex := 2                                                      ; Use default function icon for error cases
        g_LOG.Debug("GetIcon: Error getting icon for " path ": " e.Message)
    }
    Return IconIndex
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
    MainGUI_Close()
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

TabFunc(*) {                                                            ; Limit tab to switch focused control between myInputBox & ListView only
    if (MainGUI.FocusedCtrl.ClassNN = "Edit1") {                        ; MainGUI.FocusedCtrl.ClassNN: Edit1 or SysListView321
        myListView.Focus()
    } else {
         myInputBox.Focus()
    }
}

PrevCommand(*) {
    ChangeCommand(-1, False)
}

NextCommand(*) {
    ChangeCommand(1, False)
}

GotoCommand(*) {
    index := SubStr(A_ThisHotkey, 2, 1)                                 ; Get index from hotkey (select specific command = Shift + index)

    if (index <= g_MATCHED.Length)
        ChangeCommand(index, True)
}

RunSelectedCommand(*) {
    GotoCommand()
    RunCommand(g_RUNTIME["CurrentCommand"])
}
ChangeCommand(Step := 1, ResetSelRow := False) {
    rowCount := myListView.GetCount()
    if (rowCount = 0)
        return
    selectedRow := ResetSelRow ? Step : myListView.GetNext() + Step     ; Get target row no. to be selected
    selectedRow := selectedRow > rowCount ? 1 : selectedRow              ; Listview cycle selection (Mod has bug on upward cycle)
    selectedRow := selectedRow < 1 ? rowCount : selectedRow

    SyncCurrentCommandByRow(selectedRow)                                ; Get current command from selected row

    myListView.Modify(selectedRow, "Select Focus Vis")                  ; make new index row selected, Focused, and Visible
}

LV_Click(myListView, rowNumber) {
    if (!rowNumber)     ; 如果用户左键点击了列表行以外的地方
        Return

    SyncCurrentCommandByRow(rowNumber)                                  ; Get current command from focused row
}

LV_ContextMenu(GuiCtrlObj, rowNumber, IsRightClick, X, Y) {             ; On ListView ContextMenu
    Global myListView, g_MATCHED, g_RUNTIME

    if (rowNumber = 0) {    ; 如果用户右键点击了列表行以外的地方
        SetMainGUIContextMenu("", GuiCtrlObj, rowNumber, IsRightClick, X, Y)
        return
    }

    if SyncCurrentCommandByRow(rowNumber, false) {
        SetListViewContextMenu(X, Y)
    } else {                ; For cases like first hint page
        SyncCurrentCommandByRow(rowNumber)
        SetMainGUIContextMenu("", GuiCtrlObj, rowNumber, IsRightClick, X, Y)
    }
}

SetListViewContextMenu(X, Y) {
    static myListViewContextMenu := ""  ; 只在第一次调用右键菜单时创建

    if !IsObject(myListViewContextMenu) {
        myListViewContextMenu := Menu()

        try {
            myListViewContextMenu.Add(g_LNG[400], LVRunCommand)
            myListViewContextMenu.Add(g_LNG[401], OpenContainer)
            myListViewContextMenu.Add(g_LNG[402], CopyCommand)
            myListViewContextMenu.Add()
            myListViewContextMenu.Add(g_LNG[403], NewCommand)
            myListViewContextMenu.Add(g_LNG[404], EditCommand)
            myListViewContextMenu.Add(g_LNG[405], DelCommand)

            myListViewContextMenu.SetIcon(g_LNG[400], "imageres.dll", -100)
            myListViewContextMenu.SetIcon(g_LNG[401], "imageres.dll", -3)
            myListViewContextMenu.SetIcon(g_LNG[402], "imageres.dll", -5314)
            myListViewContextMenu.SetIcon(g_LNG[403], "imageres.dll", -2)
            myListViewContextMenu.SetIcon(g_LNG[404], "imageres.dll", -5306)
            myListViewContextMenu.SetIcon(g_LNG[405], "imageres.dll", -5305)

            g_LOG.Debug("SetListViewContextMenu: Create myListViewContextMenu...OK")
        } catch as e {
            g_LOG.Debug("SetListViewContextMenu: Error creating myListViewContextMenu: " . e.Message)
        }
    }
    myListViewContextMenu.Show(X, Y)
    return
}

LVRunCommand(*) {                                                       ; On ListView double click action
    focusedRow := myListView.GetNext(0, "Focused")                      ; Check focused row, only operate focusd row instead of all selected rows
    if (!focusedRow)                                                    ; Return if no focused row is found
        Return

    if SyncCurrentCommandByRow(focusedRow, false) {                     ; Get current command from focused row
        RunCommand(g_RUNTIME["CurrentCommand"])                         ; Execute the command if the user selected "Run Enter"
    }
}

CopyCommand(*) {                                                        ; ListView ContextMenu
    focusedRow := myListView.GetNext(0, "Focused")                      ; Check focused row, only operate focusd row instead of all selected rows
    if (!focusedRow)                                                    ; Return if no focused row is found
        Return

    SyncCurrentCommandByRow(focusedRow, false)                          ; Get current command from focused row

    if (MainGUI.FocusedCtrl.ClassNN = "SysListView321") {
        A_Clipboard := myListView.GetText(focusedRow, 3)                ; Get the text from the focusedRow's 3rd field.
    } else {
        SendInput("^c")                                                 ; If input box or status box is focused
    }
}

; Get text from 1st part of StatusBar
CopyStatusBarText(*) {
    A_Clipboard := StatusBarGetText(1, "ahk_id " MainGUI.Hwnd)
    ToolTip(g_LNG[408] A_Clipboard)
    SetTimer(() => ToolTip(""), -2000)    ; Hide tooltip after 2 seconds
}

MainGUI_Escape(*) {
    (g_CONFIG["EscClearInput"] && myInputBox.Value) ? ClearInput() : MainGUI_Close()
}

MainGUI_Close(*) {
    if (!g_CONFIG["KeepInput"]) {
        ClearInput()
    }

    ; Animate hide if possible
    ;try DllCall("AnimateWindow", "Ptr", MainGUI.Hwnd, "Int", 90, "UInt", 0x90000)

    MainGUI.Hide()
    ; CommandStore.UpdateUsage() only mutates g_USAGE in memory. MainGUI_Close() fires on
    ; every dismiss, including a plain Esc/Alt+Space with nothing run, so it
    ; must NOT trigger a full ALTRun.json save here. The bumped count rides
    ; along on the next real save instead (a command run, a settings change,
    ; or app exit - see the OnExit handler near the top of the script).
    CommandStore.UpdateUsage()
    SetStatusBar("TIP")                                                 ; Update StatusBar tip information after GUI hide
}

MainGUI_Size(GuiObj, MinMax, Width, Height) {
    ; g_GUI["WinX"] := Width
    ; g_GUI["WinY"] := Height

    ; g_GUI.ListX := g_GUI.WinX - 24
    ; g_GUI.ListY := g_GUI.WinY - 76
    ; g_GUI.Input_W := g_GUI.ListX - g_CONFIG.ShowBtnRun * 90 - g_CONFIG.ShowBtnOpt * 90
    ; GuiControl, Main:Move, MyListView, % "W" g_GUI.ListX " H" g_GUI.ListY ; Resize ListView to fit new window size
    ; GuiControl, Main:Move, MyInput, % "W" g_GUI.Input_W                 ; Resize Input to fit new window size
    ; LV_ModifyCol(4, "AutoHdr")                                          ; Auto adjust column width to fit new window size
    ; SB_SetParts(g_GUI.WinX - 90 * g_CONFIG.ShowRunCount)
    ; SetStatusBar("窗口大小已经改变, 如需保留窗口尺寸, 请进入选项设置页面进行保存")
    ; OutputDebug("MainGUI_Size - " MinMax "-" Width "-" Height)
}

; 主界面右键菜
MainGUI_ContextMenu(GuiObj, GuiCtrlObj, Item, IsRightClick, X, Y) {

    GuiCtrlType := IsObject(GuiCtrlObj) ? GuiCtrlObj.Type : ""

    if (GuiCtrlType = "ListView" || GuiCtrlType = "StatusBar")
        return

    SetMainGUIContextMenu(GuiObj, GuiCtrlObj, Item, IsRightClick, X, Y)
    return
}

SetMainGUIContextMenu(GuiObj, GuiCtrlObj, Item, IsRightClick, X, Y) {
    static myContextMenu := ""  ; 只在第一次调用时创建

    if !IsObject(myContextMenu) {
        myContextMenu := Menu()
        try {
            myContextMenu.Add(g_LNG[301], (*) => Options())
            myContextMenu.Add(g_LNG[310], UserCommand)
            myContextMenu.Add()  ; 分隔线
            myContextMenu.Add(g_LNG[302], Reindex)
            myContextMenu.Add(g_LNG[303], Usage)
            myContextMenu.Add(g_LNG[305], (*) => ListLines())
            myContextMenu.Add()  ; 分隔线
            myContextMenu.Add(g_LNG[309], Update)
            myContextMenu.Add(g_LNG[304], About)
            myContextMenu.Add()
            myContextMenu.Add(g_LNG[307], RestartApp)
            myContextMenu.Add(g_LNG[308], Exit)

            SetMenuItemIcons(myContextMenu, true)

            g_LOG.Debug("SetMainGUIContextMenu: Create myContextMenu...OK")
        } catch as e {
            g_LOG.Debug("SetMainGUIContextMenu: Error creating myContextMenu: " . e.Message)
        }
    }
    myContextMenu.Show(X, Y)
    return
}

ClearInput() {
    myInputBox.Focus()
    myInputBox.Value := ""
    Input_Change()                                                      ; v1 no need, v2 需要手动调用绑定的事件处理函数
}

Exit(*) {
    ExitApp()
}

RestartApp(*) {
    Reload()
}

SetStatusBar(strToShow) {                                               ; Set StatusBar text, Mode 1: Current command (default), 2: Hint, 3: Any text
    if (strToShow = "TIP")
        strToShow := g_LNG[51] g_LNG[Random(52, 71)]                    ; Randomly select a tip from hint list g_LNG 52~71

    myStatus.value := strToShow
}

RunCurrentCommand(*) {
    RunCommand(g_RUNTIME["CurrentCommand"])
}

ParseArg() {
    inputVal := myInputBox.Value
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
; Lib\CommandStore.ahk (see the #Include list at the top of this file).
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

GetCmdOutput(command) {
    TempFile    := A_Temp . "\ALTRun.stdout"
    FullCommand := A_ComSpec " /C " command " > " TempFile

    RunWait(FullCommand, A_Temp, "Hide")
    Result := FileRead(TempFile)
    try FileDelete(TempFile)
    Return RTrim(Result, "`r`n")                                        ; Remove result rightmost/last "`r`n"
}

GetRunResult(command) {                                                 ; 运行CMD并取返回结果方式2
    shell := ComObject("WScript.Shell")                                 ; WshShell object: https://msdn.microsoft.com/en-us/library/aew9yb99
    exec := shell.Exec(A_ComSpec " /C " command)                        ; Execute a single command via cmd.exe
    Return exec.StdOut.ReadAll()                                        ; Read and Return the command's output
}

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

; 监听窗口失去焦点时自动关闭
MonitorFocus(*) {
    if (!WinExist("ahk_id " MainGUI.Hwnd) || g_RUNTIME["UseDisplay"])
        return

    if (!WinActive("ahk_id " MainGUI.Hwnd)) {
        MainGUI_Close()
        g_LOG.Debug("MonitorFocus: ALTRun lose focus, auto closing...")
    }
}

; WM_ACTIVATE(wParam, lParam, msg, hwnd){                                 ; Close on lose focus, OnMessage is far more efficient than SetTimer + WinActive check
;     if (hwnd != MainGUI.Hwnd) {                                         ; Ignore messages from other windows
;         g_LOG.Debug("WM_ACTIVATE: Ignored message from hwnd (" hwnd ") != MainGUI.Hwnd (" MainGUI.Hwnd ")")
;         return 0
;     }

;     if (!WinExist("ahk_id " MainGUI.Hwnd)) {                            ; Ignore when MainGUI does not exist, to avoid flahshing issue
;         g_LOG.Debug("WM_ACTIVATE: Ignored message, MainGUI does not exist...")
;         return 0
;     }

;     isActivated := (wParam > 0)                                         ; wParam > 0 means the window is being activated
;     g_LOG.Debug("WM_ACTIVATE: Window is " (isActivated ? "activated..." : "de-activated..."))

;     if (!isActivated && WinExist("ahk_id " MainGUI.Hwnd) && !g_RUNTIME["UseDisplay"]) {
;         MainGUI_Close()
;         g_LOG.Debug("WM_ACTIVATE: Window lose focus, auto closing...")
;     }
;     return 0
; }

UpdateSendTo() {                 ; the lnk in SendTo must point to a exe
    lnkPath := StrReplace(A_StartMenu, "\Start Menu", "\SendTo\") "ALTRun.lnk"
    if (!g_CONFIG["EnableSendTo"]) {
        try FileDelete(lnkPath)
        g_LOG.Debug("UpdateSendTo: Update SendTo shortcut...Disabled")
        return
    }

    if (A_IsCompiled)
        FileCreateShortcut(A_ScriptFullPath, lnkPath, ,"-SendTo", "Send command to ALTRun User Command list")
    else
        FileCreateShortcut(A_AhkPath, lnkPath, , A_ScriptFullPath " -SendTo", "Send command to ALTRun User Command list")

    g_LOG.Debug("UpdateSendTo: Update SendTo shortcut...OK")
    return
}

UpdateStartup() {
    lnkPath := A_Startup "\ALTRun.lnk"

    if (!g_CONFIG["AutoStartup"]) {
        try FileDelete(lnkPath)
        g_LOG.Debug("UpdateStartup: Update Startup shortcut...Disabled")
        return
    }

    FileCreateShortcut(A_ScriptFullPath, lnkPath, A_ScriptDir, "-startup", "ALTRun - An effective launcher")

    g_LOG.Debug("UpdateStartup: Update Startup shortcut...OK")
    return
}

UpdateStartMenu() {
    lnkPath := A_Programs "\ALTRun.lnk"

    if (!g_CONFIG["InStartMenu"]) {
        try FileDelete(lnkPath)
        g_LOG.Debug("UpdateStartMenu: Update StartMenu shortcut...Disabled")
        return
    }

    FileCreateShortcut(A_ScriptFullPath, lnkPath, A_ScriptDir, "-StartMenu", "ALTRun - An effective launcher")
    g_LOG.Debug("UpdateStartMenu: Update StartMenu shortcut...OK")
    return
}

Reindex(*) {                                                            ; Re-create Index section
    ; Collect every indexed entry into a fresh map, then store it in one go
    AppData.LoadAppData()
    indexMap := Map()

    ; Create ProgressGui at the start
    ProgressGui := Gui("-MinimizeBox +AlwaysOnTop", "Reindex")
    ProgressGui.Add("Text", , "ReIndexing...")
    ProgressGui.Add("Progress", "vMyProgress w200", 0)
    ProgressGui.Add("Text", "vMyFileName w200", "Starting...")
    ProgressGui.Show()

    ; Move repeated config queries outside loop
    maxDepth := g_CONFIG["IndexDepth"]
    shouldCheckExclude := g_CONFIG["IndexExclude"] != ""
    excludePattern := g_CONFIG["IndexExclude"]

    for dirIndex, dir in StrSplit(g_CONFIG["IndexDir"], ",") {
        searchPath := RegExReplace(Path.Resolve(Trim(dir)), "\\+$")     ; Remove trailing backslashes
        if !DirExist(searchPath)
            continue

        for extIndex, ext in StrSplit(g_CONFIG["IndexType"], ",") {
            ext := Trim(ext)
            if (ext = "")
                continue
            Loop Files, searchPath "\" ext, "R" {                       ; Calculate path relative to searchPath and count subdir levels
                rel := SubStr(A_LoopFileFullPath, StrLen(searchPath) + 2) ; +2 to skip the backslash
                seps := (rel = "") ? 0 : StrLen(rel) - StrLen(StrReplace(rel, "\", "")) ; Count backslashes to determine depth

                if (seps > maxDepth)                                    ; If file is deeper than allowed depth, skip it.
                    continue

                if (shouldCheckExclude && RegExMatch(A_LoopFileFullPath, excludePattern))
                    continue                                            ; Skip this file and move on to the next loop.

                indexMap["File | " . A_LoopFileFullPath] := 1   ; Collect file entry

                ; Update ProgressGui (throttled to reduce UI overhead)
                if (!Mod(A_Index, 20))
                    ProgressGui["MyProgress"].Value := Mod(A_Index, 100), ProgressGui["MyFileName"].Text := A_LoopFileName
            }
        }
    }

    ; Index Windows Store Apps
    if (g_CONFIG["IndexStoreApp"]) {
        try {
            ProgressGui["MyFileName"].Text  := "Indexing Store Apps..."
            tempFile := A_Temp . "\ALTRun_StoreApps.csv"
            RunWait('powershell -Command "Get-StartApps | Select-Object Name, AppID | ConvertTo-Csv -NoTypeInformation" > "' . tempFile . '"', , "Hide")
            if FileExist(tempFile) {
                output := FileRead(tempFile)
                FileDelete(tempFile)
                lines := StrSplit(output, "`n", "`r")
                for line in lines {
                    if (A_Index == 1 or Trim(line) == "")  ; Skip header
                        continue
                    fields := StrSplit(line, '","')
                    if (fields.Length >= 2) {
                        name := StrReplace(fields[1], '"', '')
                        appid := StrReplace(fields[2], '"', '')
                        indexMap["App | shell:AppsFolder\" . appid . " | " . name] := 1 ; Collect app entry
                    }
                    ProgressGui["MyProgress"].Value := A_Index
                    ProgressGui["MyFileName"].Text  := name ? name : "Unknown App"
                    Sleep 10  ; Small delay to show progress
                }
                g_LOG.Debug("Reindex: Indexed Windows Store apps successfully")
            } else {
                g_LOG.Debug("Reindex: Temp file not found for Store apps")
            }
        } catch as e {
            g_LOG.Debug("Reindex: Error indexing Store apps: " . e.Message)
        }
    }

    ; Destroy ProgressGui at the end
    ProgressGui.Destroy()

    ; Keep the rank a command already earned, so reindexing does not reset SmartRank
    for cmdLine, _ in indexMap {
        if (g_CMDDATA["Index"].Has(cmdLine) && IsInteger(g_CMDDATA["Index"][cmdLine]))
            indexMap[cmdLine] := g_CMDDATA["Index"][cmdLine]
    }
    g_CMDDATA["Index"] := indexMap
    AppData.SaveAppData()

    g_LOG.Debug("Reindex: Indexing search database...OK")
    TrayTip("ReIndex database finish successfully.", g_TITLE, 8)
    CommandStore.LoadCommands()
}

About(*) {
    Options(8)
}

Usage(*) {
    Options(7)
}

AutoCheckUpdate(*) {
    if (!g_CONFIG["AutoUpdateCheck"])
        return

    ; One-time timer to check for updates
    SetTimer(CheckUpdate, -1000)
}

Update(*) {
    ; Manual update check, show message box
    CheckUpdate(False)
}

; Main update check function, called from tray menu (*) or with silent flag
CheckUpdate(Silent := True) {
    RepoAPI     := "https://api.github.com/repos/zhugecaomao/ALTRun/releases/latest"
    ReleasePage := "https://github.com/zhugecaomao/ALTRun/releases"

    try {
        tmpFile := A_Temp "\ALTRun_latest.json"
        Download(RepoAPI, tmpFile)
        json := FileRead(tmpFile, "UTF-8")

        ; Extract latest version from tag_name
        if !RegExMatch(json, '"tag_name"\s*:\s*"([^"]+)"', &verMatch)
            throw Error("Cannot find 'tag_name' in GitHub API response.")

        latestVersion  := Trim(verMatch[1], "vV ")
        currentVersion := Trim(g_TITLE, "ALTRun - v ")

        ; Compare versions
        if (CompareVersion(latestVersion, currentVersion) > 0) {
            MsgBox(g_LNG[805] latestVersion g_LNG[806], g_Title, 64)
            Run ReleasePage
        } else if (!Silent) {
            ; Only show "up-to-date" message for manual checks
            MsgBox(g_LNG[807] currentVersion g_LNG[808], g_Title, 64)
        }

    } catch as e {
        if (!Silent)
            MsgBox(g_LNG[809] e.Message, g_Title, 48)
        else
            g_LOG.Debug("CheckUpdate: Update check failed: " e.Message)
    }
}

CompareVersion(v1, v2) {
    v1Parts := StrSplit(v1, ".")
    v2Parts := StrSplit(v2, ".")
    Loop Max(v1Parts.Length, v2Parts.Length) {
        diff := (v1Parts[A_Index] + 0) - (v2Parts[A_Index] + 0)
        if diff
            return diff
    }
    return 0
}

; Listary()/ShowListaryHint()/GetListaryHintText()/IsQuickSwitchDialog()/
; HasAnyCtrlMatch()/IsLikelyFileDialogTitle()/SyncTCPath()/SyncExplorerPath()/
; SetDialogPath() used to live here; all moved into the Listary class in
; Lib\Listary.ahk (see the #Include list at the top of this file and the
; "Listary.Init()" call in the autorun section).

UserCommand(*) {                                                        ; F4 - edit the command database directly
    Run("Notepad.exe " . g_JSON)
}

; From command "New Command" or GUI context menu "New Command"
NewCommand(*) {
    OpenCommandManager("UserCommand", , , g_RUNTIME["Arg"], 1, "")
}

EditCommand(*) {
    Global g_RUNTIME  ; 明确声明全局变量

    currentCmd := g_RUNTIME["CurrentCommand"]
    if !currentCmd
        return MsgBox(g_LNG[810], g_TITLE, 64)                          ; 64 = Info icon

    AppData.LoadAppData()

    for _, section in ["DefaultCommand", "UserCommand", "Index"] {
        if !g_CMDDATA[section].Has(currentCmd)
            continue
        rank := g_CMDDATA[section][currentCmd]

        if IsInteger(rank) {
            parts := StrSplit(currentCmd, " | ")
            type := parts.Length >= 1 ? parts[1] : ""
            path := parts.Length >= 2 ? parts[2] : ""
            desc := parts.Length >= 3 ? parts[3] : ""

            g_Log.Debug("EditCommand: Editing command=" currentCmd)
            OpenCommandManager(section, type, path, desc, rank, currentCmd)
            break
        }
    }
}

DelCommand(*) {
    currentCmd := g_RUNTIME["CurrentCommand"]
    if !currentCmd
        return

    AppData.LoadAppData()

    for _, section in ["DefaultCommand", "UserCommand", "Index"] {
        if !g_CMDDATA[section].Has(currentCmd)
            continue

        result := MsgBox(g_LNG[800] section "]`n`n" currentCmd, g_LNG[801], 52) ; 52 = Yes/No + Question icon

        if result = "YES" {
            try {
                g_CMDDATA[section].Delete(currentCmd)
                AppData.SaveAppData()
                MsgBox(g_LNG[802] "`n`n" currentCmd, g_TITLE, 64)       ; 64 = Info icon
            } catch as e {
                MsgBox(g_LNG[803] "`n`n" currentCmd, g_TITLE, 48)       ; 48 = Error icon
            }
            break
        }
    }
    CommandStore.LoadCommands()
}


OpenCommandManager(Section := "UserCommand", Type := "File", Path := "", Desc := "", Rank := 1, OriginCmd := "") { ; 命令管理窗口
    Global g_CmdMgrGui
    Local  typeList := Array("File", "Dir", "CMD", "URL", "Func", "Clip")
    chooseIndex := GetArrayIndex(Type, typeList)
    chooseIndex := chooseIndex ? chooseIndex : 1

    g_LOG.Debug("Starting Command Manager... Args=" Section "|" Type "|" Path "|" Desc "|" Rank)

    g_CmdMgrGui := Gui(, g_LNG[700])
    g_CmdMgrGui.SetFont("S9 Norm", "Microsoft Yahei")
    g_CmdMgrGui.AddGroupBox("w600 h260", g_LNG[701])
    g_CmdMgrGui.Add("Text", "x25 yp+30", g_LNG[702])
    g_CmdMgrGui.AddDropDownList("x160 yp-5 w130 vType Choose" chooseIndex, typeList)
    g_CmdMgrGui.Add("Text", "x315 yp+5", g_LNG[705])
    g_CmdMgrGui.Add("Edit", "x435 yp-5 w130 Disabled vSection", Section)
    g_CmdMgrGui.Add("Text", "x25 yp+60", g_LNG[703])
    g_CmdMgrGui.Add("Edit", "x160 yp-5 w405 -WantReturn vPath", Path).Focus()
    g_CmdMgrGui.AddButton("x575 yp w30 hp", "...").OnEvent("Click", (*) => PickCommandTarget(g_CmdMgrGui["Type"].Text))
    g_CmdMgrGui.Add("Text", "x25 yp+80", g_LNG[704])
    g_CmdMgrGui.AddEdit("x160 yp-5 w405 -WantReturn vDesc", Desc)
    g_CmdMgrGui.AddText("x25 yp+60", g_LNG[706])
    g_CmdMgrGui.AddEdit("x160 yp-5 w405 +Number vRank", Rank)
    g_CmdMgrGui.AddButton("Default x420 w90", g_LNG[7]).OnEvent("Click", (*) => SaveCommandFromManager(Section, g_CmdMgrGui["Type"].Text, g_CmdMgrGui["Path"].Text, g_CmdMgrGui["Desc"].Text, g_CmdMgrGui["Rank"].Text, OriginCmd))
    g_CmdMgrGui.AddButton("x521 yp w90", g_LNG[8]).OnEvent("Click", CloseCommandManager)
    g_CmdMgrGui.OnEvent("Close", CloseCommandManager)
    g_CmdMgrGui.OnEvent("Escape", CloseCommandManager)
    g_CmdMgrGui.Show("Center")
}

PickCommandTarget(cmdType) {
    g_CmdMgrGui.Opt("+OwnDialogs")                                      ; Make open dialog Modal

    if (cmdType = "Dir")
        cmdPath := DirSelect(, 3, 'Please select directory')
    else if (cmdType = "File")
        cmdPath := FileSelect(3, , , 'All Files (*.*)')
    else if (cmdType = "Clip")
        return Clip.EditClipText()                                           ; Clip uses a multi-line text editor instead of a file picker
    else
        return MsgBox("Path picker only supports File/Dir/Clip type.", g_LNG[700], 64)

    if (cmdPath != "")
        g_CmdMgrGui["Path"].Value := cmdPath
}

SaveCommandFromManager(section, cmdType, cmdPath, cmdDesc, cmdRank, originCmd) {
    g_CmdMgrGui.Submit()
    validType := Map("File", 1, "Dir", 1, "CMD", 1, "URL", 1, "Func", 1, "Clip", 1)
    section := Trim(section)
    cmdType := Trim(cmdType)
    cmdPath := Trim(cmdPath)
    cmdDesc := Trim(cmdDesc)
    cmdRank := Trim(cmdRank)

    if !validType.Has(cmdType)
        return MsgBox("Invalid command type: " cmdType, g_LNG[820], 48)

    if (cmdPath = "") {
        return MsgBox(g_LNG[821], g_LNG[820], 64)
    }

    if (cmdType = "Clip") {
        ; The Path field already holds the escaped single-line form (EditClipText produced it),
        ; so only guard against stray real line breaks - never re-escape, that would double the backslashes.
        cmdPath := StrReplace(StrReplace(StrReplace(cmdPath, "`r`n", "\n"), "`n", "\n"), "`r", "\n")
        if (cmdDesc = "")
            return MsgBox("A Clip command needs a short name in the Description field, that is what you type to call it.", g_LNG[820], 48)
    }

    if (!IsInteger(cmdRank) || cmdRank <= 0)
        cmdRank := 1

    cmdLine := cmdType " | " cmdPath (cmdDesc != "" ? " | " cmdDesc : "")
    try {
        AppData.LoadAppData()
        if !g_CMDDATA.Has(section)
            section := "UserCommand"
        if (originCmd != "" && originCmd != cmdLine) {                  ; Drop the old key only when editing changed the command line
            for _, sec in ["DefaultCommand", "UserCommand", "Index"]
                if g_CMDDATA[sec].Has(originCmd)
                    g_CMDDATA[sec].Delete(originCmd)
        }
        g_CMDDATA[section][cmdLine] := cmdRank + 0
        AppData.SaveAppData()
    } catch as e {
        MsgBox(g_LNG[822] e.Message, g_LNG[820], 64)
        return
    }
    MsgBox(g_LNG[823] section " ]`n`n" cmdLine " = " cmdRank, g_LNG[820], 64)
    CommandStore.LoadCommands()
}

CloseCommandManager(*) {
    g_CmdMgrGui.Destroy()
}

; Plugins()/RenameWithDate()/LineEndAddDate()/NameAddDate() used to live here;
; all moved into the Plugins class in Lib\Plugins.ahk (see the #Include list
; at the top of this file and the "Plugins.Init()" call in the autorun section).

GetArrayIndex(searchValue, Array){
    for index, element in Array
    {
        if (element = searchValue)
            return index
    }
    return 0
}

; LoadAppData()/SaveAppData()/MergeIntoDefaults()/OnAppExit()/ParseCommandBlock()/
; DefaultCommandText()/UserCommandText()/FallbackCommandText() used to live
; here; all moved into the AppData class in Lib\AppData.ahk (called as
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
; live here; all moved into the Clip class in Lib\Clip.ahk (see the #Include
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

PTTools() {
    PTToolsWindow.Show()                                                  ; Lib\PTTools.ahk - Rebar/BRC calculator + SPF2M automation
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

Options(ActTab := 1) {
    Global OptGUI, OptListView
    static FuncList := ["Unset", "Active", "ToggleWindow", "Google", "Bing"
        , "Everything", "TabFunc", "PrevCommand", "NextCommand", "CopyCommand"
        , "ClearInput", "RunCurrentCommand", "RankUp", "RankDown", "Reindex"
        , "About", "Usage", "Update", "UserCommand", "NewCommand", "EditCommand"
        , "DelCommand", "OpenCommandManager", "Options", "TurnMonitorOff", "EmptyRecycle"
        , "MuteVolume", "RestartApp", "Exit", "PTTools"]

    g_LOG.Debug("Options: Opening Options window... Tab=" ActTab)
    MainGUI_Close()
    if WinExist(g_LNG[2]) {
        return WinActivate(g_LNG[2])
    }

    t := A_TickCount
    ActTab := IsNumber(ActTab) ? ActTab : 1                             ; Convert ActTab to number, default is 1 (for case like [Option`tF2])
    optFont := Fonts.Spec(g_GUI["OptGUIFont"], "Microsoft YaHei", "norm s9.0")
    mainFont := Fonts.Spec(g_GUI["MainGUIFont"], "Microsoft YaHei", "norm s10.0")
    sbFont := Fonts.Spec(g_GUI["MainSBFont"], "Microsoft YaHei", "norm s9.0")
    OptGUI := Gui("+Owner" MainGUI.hwnd, g_LNG[2])                      ; +Owner MainGUI.hwnd fix GUI flicking issue
    OptGUI.SetFont(optFont.opt, optFont.name)
    OptTab := OptGUI.AddTab3("Choose" ActTab, g_LNG[100])

    OptTab.UseTab(1) ; CONFIG Tab
    OptListView := OptGUI.AddListView("w500 h300 Checked -Hdr", ["Settings"])
    for key, description in g_CONFIG_P1 {
        OptListView.Add("Check" g_CONFIG[key], description)
    }
    OptListView.ModifyCol(1, "AutoHdr")

    OptGUI.AddText("x24 yp+320", g_LNG[150])
    OptGUI.AddComboBox("x130 yp-5 w394 vFileMgr Choose1", [g_CONFIG["FileMgr"], "Explorer.exe", "C:\Apps\TotalCMD.exe /O /T /S"])
    OptGUI.AddText("x24 yp+40", g_LNG[151])
    OptGUI.AddComboBox("x130 yp-5 w394 vEverything Choose1", [g_CONFIG["Everything"], "C:\Apps\Everything.exe"])
    OptGUI.AddText("x24 yp+40", g_LNG[152])
    OptGUI.AddDDL("x130 yp-5 w394 Sort vHistoryLen Choose" g_CONFIG["HistoryLen"]*0.1, [10,20,30,40,50,60])

    OptTab.UseTab(2) ; GUI Tab
    OptGUI.AddGroupBox("w500 h420", g_LNG[170])
    OptGUI.AddText("x33 yp+25", g_LNG[171])
    OptGUI.AddDDL("x183 yp-5 w330 vListRows Choose" g_GUI["ListRows"], [1,2,3,4,5,6,7,8,9]) ; ListRows limit <= 9
    OptGUI.AddText("x33 yp+45", g_LNG[172])
    OptGUI.AddComboBox("x183 yp-5 w330 vColWidth Choose1", [g_GUI["ColWidth"], "20,0,460,AutoHdr", "30,46,460,AutoHdr"])
    OptGUI.AddText("x33 yp+45", g_LNG[176])
    OptGUI.AddEdit("x183 yp-5 w120 +Number vWinX", g_GUI["WinX"])
    OptGUI.AddText("x345 yp", "x")
    OptGUI.AddEdit("x393 yp w120 +Number vWinY", g_GUI["WinY"])

    OptGUI.AddText("x33 yp+45", g_LNG[173])
    OptGUI.AddEdit("x183 yp w240 r1 -E0x200 +ReadOnly vMainGUIFont", g_GUI["MainGUIFont"]).SetFont(mainFont.opt, mainFont.name)
    OptGUI.AddButton("x433 yp-5 w80", g_LNG[182]).OnEvent("Click", (*) => SelectFont("MainGUIFont"))
    OptGUI.AddText("x33 yp+45", g_LNG[174])
    OptGUI.AddEdit("x183 yp w240 r1 -E0x200 +ReadOnly vOptGUIFont", g_GUI["OptGUIFont"])
    OptGUI.AddButton("x433 yp-5 w80", g_LNG[182]).OnEvent("Click", (*) => SelectFont("OptGUIFont"))
    OptGUI.AddText("x33 yp+45", g_LNG[175])
    OptGUI.AddEdit("x183 yp w240 r1 -E0x200 +ReadOnly vMainSBFont", g_GUI["MainSBFont"]).SetFont(sbFont.opt, sbFont.name)
    OptGUI.AddButton("x433 yp-5 w80", g_LNG[182]).OnEvent("Click", (*) => SelectFont("MainSBFont"))

    OptGUI.AddText("x33 yp+45", g_LNG[179])
    OptGUI.AddEdit("x183 yp w240 r1 -E0x200 +ReadOnly vMainGUIColor", g_GUI["MainGUIColor"])
    OptGUI.AddButton("x433 yp-5 w80", g_LNG[183]).OnEvent("Click", PickMainGUIColor)
    OptGUI.AddText("x33 yp+45", g_LNG[178])
    OptGUI.AddEdit("x183 yp w240 r1 -E0x200 +ReadOnly vCMDListColor", g_GUI["CMDListColor"])
    OptGUI.AddButton("x433 yp-5 w80", g_LNG[183]).OnEvent("Click", PickCMDListColor)

    OptGUI.AddText("x33 yp+45", g_LNG[180])
    OptGUI.AddComboBox("x183 yp-5 w240 vBackground Choose1", [g_GUI["Background"], "Default", "None", "ALTRun.jpg", "C:\Path\Picture.jpg"])
    OptGUI.AddButton("x433 yp-2 w80 vSelectBackground", g_LNG[184]).OnEvent("Click", SelectBackground)
    OptGUI.AddText("x33 yp+45", g_LNG[181])
    OptGUI.AddSlider("x183 yp-5 w330 Range50-255 TickInterval5 Tooltip vTransparency", g_GUI["Transparency"])

    OptTab.UseTab(3) ; Hotkey Tab
    OptGUI.AddGroupBox("w500 h115", g_LNG[191])
    OptGUI.AddText("x33 yp+25", g_LNG[192])
    OptGUI.AddHotkey("x285 yp-4 w230 vGlobalHotkey1", g_HOTKEY["GlobalHotkey1"])
    OptGUI.AddText("x33 yp+35", g_LNG[193])
    OptGUI.AddHotkey("x285 yp-4 w230 vGlobalHotkey2", g_HOTKEY["GlobalHotkey2"])
    OptGUI.AddText("x33 yp+35", g_LNG[194])
    OptGUI.AddLink("x285 yp w230", "<a>" g_LNG[195] "</a>").OnEvent("Click", ResetHotkey)

    OptGUI.Add("GroupBox", "x24 yp+38 w500 h290", g_LNG[200])
    Loop 7 {
        OptGUI.AddText("x33 yp+40", g_LNG[201])
        OptGUI.AddHotkey("x143 yp-5 w120 vHotkey" A_Index, g_HOTKEY["Hotkey" A_Index])
        OptGUI.AddText("x285 yp+5", g_LNG[202])
        OptGUI.AddDDL("x395 yp-5 w120 vTrigger" A_Index " Choose" GetArrayIndex(g_HOTKEY["Trigger" A_Index], FuncList), FuncList)
    }

    ToggleGlobalHotkeys("Off", "Options")                                ; Turn off global hotkeys in options

    OptTab.UseTab(4) ; INDEX Tab
    OptGUI.AddGroupBox("w500 h220", g_LNG[160])
    OptGUI.AddText("x33 yp+25", g_LNG[161])
    OptGUI.AddComboBox("x183 yp-5 w330 vIndexDir Choose1", [g_CONFIG["IndexDir"], "A_ProgramsCommon,A_StartMenu"])
    OptGUI.AddText("x33 yp+45", g_LNG[162])
    OptGUI.AddComboBox("x183 yp-5 w330 vIndexType Choose1", [g_CONFIG["IndexType"], "*.lnk,*.exe"])
    OptGUI.AddText("x33 yp+45", g_LNG[164])
    OptGUI.AddDropDownList("x183 yp-5 w330 vIndexDepth Choose" g_CONFIG["IndexDepth"], [1,2,3,4,5,6,7,8,9])
    OptGUI.AddText("x33 yp+45", g_LNG[163])
    OptGUI.AddComboBox("x183 yp-5 w330 vIndexExclude Choose1", [g_CONFIG["IndexExclude"], "Uninstall *"])
    OptGUI.AddCheckBox("x33 yp+45 vIndexStoreApp Checked" g_CONFIG["IndexStoreApp"], g_LNG[165])

    OptTab.UseTab(5) ; LISTARY Tab
    OptGUI.AddGroupBox("w500 h145", g_LNG[211])
    OptGUI.AddText("x33 yp+30", g_LNG[212])
    OptGUI.AddComboBox("x183 yp-5 w330 vFileMgrID Choose1", [g_CONFIG["FileMgrID"], "ahk_class CabinetWClass", "ahk_class CabinetWClass, ahk_class TTOTAL_CMD"])
    OptGUI.AddText("x33 yp+45", g_LNG[213])
    OptGUI.AddComboBox("x183 yp-5 w330 vDialogWin Choose1", [g_CONFIG["DialogWin"], "ahk_class #32770"])
    OptGUI.AddText("x33 yp+45", g_LNG[214])
    OptGUI.AddComboBox("x183 yp-5 w330 vExcludeWin Choose1", [g_CONFIG["ExcludeWin"], "ahk_class SysListView32, ahk_exe Explorer.exe"])
    OptGUI.AddGroupBox("x24 yp+50 w500 h145", g_LNG[215])
    OptGUI.AddText("x33 yp+30", g_LNG[216])
    OptGUI.AddHotkey("x183 yp-5 w330 vTotalCMDDir", g_HOTKEY["TotalCMDDir"])
    OptGUI.AddText("x33 yp+45", g_LNG[217])
    OptGUI.AddHotkey("x183 yp-5 w330 vExplorerDir", g_HOTKEY["ExplorerDir"])
    OptGUI.AddCheckBox("x33 yp+45 vAutoSwitchDir Checked" g_CONFIG["AutoSwitchDir"], g_LNG[218])

    OptTab.UseTab(6) ; PLUGINS Tab
    OptGUI.AddGroupBox("w500 h110", g_LNG[251])
    OptGUI.AddText("x33 yp+30", g_LNG[252])
    OptGUI.AddComboBox("x183 yp-5 w330 vAutoDateAtEnd Choose1", [g_HOTKEY["AutoDateAtEnd"], "ahk_class TCmtEditForm,ahk_exe Notepad4.exe"])
    OptGUI.AddText("x33 yp+45", g_LNG[253])
    OptGUI.AddHotkey("x183 yp-5 w80 vAutoDateAEHKey", g_HOTKEY["AutoDateAEHKey"])
    OptGUI.AddText("x300 yp+5", g_LNG[254])
    OptGUI.AddDDL("x395 yp-5 w120 vAutoDateAEFormat Choose1", ["- dd.MM.yyyy"])

    OptGUI.AddGroupBox("x24 y+30 w500 h110", g_LNG[255])
    OptGUI.AddText("x33 yp+30", g_LNG[252])
    OptGUI.AddComboBox("x183 yp-5 w330 vAutoDateBefExt Choose1", [g_HOTKEY["AutoDateBefExt"], "ahk_class CabinetWClass,ahk_class Progman,ahk_class WorkerW,ahk_class #32770"])
    OptGUI.AddText("x33 yp+45", g_LNG[253])
    OptGUI.AddHotkey("x183 yp-5 w80 vAutoDateBEHKey", g_HOTKEY["AutoDateBEHKey"])
    OptGUI.AddText("x300 yp+5", g_LNG[254])
    OptGUI.AddDDL("x395 yp-5 w120 vAutoDateBEFormat Choose1", ["- dd.MM.yyyy"])

    OptGUI.AddGroupBox("x24 y+30 w500 h110", g_LNG[259])
    OptGUI.AddText("x33 yp+30", g_LNG[260])
    OptGUI.AddComboBox("x183 yp-5 w330 vCondTitle Choose1", [g_HOTKEY["CondTitle"]])
    OptGUI.AddText("x33 yp+45", g_LNG[261])
    OptGUI.AddComboBox("x183 yp-5 w80 vCondHotkey Choose1", [g_HOTKEY["CondHotkey"]])
    OptGUI.AddText("x300 yp+5", g_LNG[262])
    OptGUI.AddDDL("x395 yp-5 w120 vCondAction Choose" GetArrayIndex(g_HOTKEY["CondAction"], FuncList), FuncList)

    OptTab.UseTab(7) ; USAGE Tab
    OptGUI.AddGroupBox("x66 y80 w445 h300", )

    g_USAGE[A_YYYY . A_MM . A_DD] := g_USAGE.Has(A_YYYY . A_MM . A_DD) ? g_USAGE[A_YYYY . A_MM . A_DD] : 1
    for date, count in g_USAGE { ; Draw usage graph
        OptGUI.AddProgress("c94DD88 Vertical y96 w14 h280 xm+" 50+A_Index*14 " Range0-" g_RUNTIME["Max"]+10, count)
    }

    OptGUI.AddText("x24 yp-5 cGray",g_RUNTIME["Max"])
    OptGUI.AddText("x24 yp+140 cGray", Round(g_RUNTIME["Max"]/2))
    OptGUI.AddText("x24 yp+140 cGray", 0)
    OptGUI.AddText("x66 yp+15 cGray", g_LNG[500])
    OptGUI.AddText("x476 yp cGray", g_LNG[501])
    OptGUI.AddText("x66 yp+33", g_LNG[502])
    OptGUI.AddEdit("x400 yp-5 w100 r1 -E0x200 +ReadOnly Right vRunCount", g_CONFIG["RunCount"])
    OptGUI.AddText("x66 yp+35", g_LNG[503])
    OptGUI.AddEdit("x400 yp-5 w100 r1 -E0x200 +ReadOnly Right", g_USAGE[A_YYYY . A_MM . A_DD])

    OptTab.UseTab(8) ; ABOUT Tab
    OptGUI.AddPic("x33 y+20 w48 h-1 Icon-100", "imageres.dll")
    OptGUI.AddText("x96 yp+5 w400", g_TITLE).SetFont("S11")
    OptGUI.AddLink("xp yp+45 w400", g_LNG[601])

    OptTab.UseTab()  ; 后续添加的控件将不属于前面的选项卡控件
    OptGUI.AddButton("Default x278 w80", g_LNG[7]).OnEvent("Click", OPTButtonOK)
    OptGUI.AddButton("x368 yp w80", g_LNG[8]).OnEvent("Click", OPTGuiClose)
    OptGUI.AddButton("x458 yp w80", g_LNG[9]).OnEvent("Click", (*) => Run("https://github.com/zhugecaomao/ALTRun/wiki"))
    OptGUI.OnEvent("Close", OPTGuiClose)
    OptGUI.OnEvent("Escape", OPTGuiClose)

    g_LOG.Debug("Options: Load options window...OK, elapsed time=" A_TickCount - t "ms")
    OutputDebug("Options: Load options window...OK, elapsed time=" A_TickCount - t "ms")
    OptGUI.Show("Center")
    return
}

ResetHotkey(*) {
    OptGUI["GlobalHotkey1"].Value := "!Space"
    OptGUI["GlobalHotkey2"].Value := "!r"
    return
}

SelectFont(TargetVar := "MainGUIFont") {
    ; Set the fontObj (optional) - only set the ones you want to pre-select
	; fontObj := Map("name","Terminal","size",14,"color",0xFF0000,"strike",1,"underline",1,"italic",1,"bold",1)
    initFont := StrSplit(g_GUI[TargetVar], ",")[1]
    fontObj  := Map("name", initFont)
    fontObj  := FontDialog.Choose(fontObj, OptGUI.hwnd)
    if (!fontObj)
        return

    OptGUI[TargetVar].Text := fontObj["name"] ", " fontObj["str"]       ; 更新控件字体并设置显示文本
    OptGUI[TargetVar].SetFont(fontObj["str"], fontObj["name"])
    g_LOG.Debug("SelectFont: OptGUI[" TargetVar "] font set to=" fontObj["str"] ", " fontObj["name"])
}

PickCMDListColor(*) {
    color := ColorDialog.Choose(g_GUI["CMDListColor"], OptGUI.hwnd, , "full")  ; hwnd and custColorObj are optional
    if (color = -1)
        return

    ;g_GUI["CMDListColor"]        := color
    OptGUI["CMDListColor"].Value := color                               ; 更新选项窗口控件并设置控件颜色
    ;OptGUI["CMDListColor"].Opt("c" color)
}

PickMainGUIColor(*) {
    color := ColorDialog.Choose(g_GUI["MainGUIColor"], OptGUI.hwnd, , "full")
    if (color = -1)
        return

    ;g_GUI["MainGUIColor"]        := color
    OptGUI["MainGUIColor"].Value := color
    ;OptGUI["MainGUIColor"].Opt("c" color)
}

SelectBackground(*) {
    OptGUI.Opt("+OwnDialogs")                                           ; Make open dialog Modal

    file := FileSelect(3, , , 'Image Files (*.jpg; *.png; *.bmp; *.gif)')
    if (file = "")
        return

    OptGUI["Background"].Text := file
    g_LOG.Debug("SelectBackground: Background image selected=" file)
}

OPTButtonOK(*) {
    SaveConfig()
    Reload
}

OPTGuiClose(*) {
    g_LOG.Debug("OPTGuiClose: Closing Options window...")

    ToggleGlobalHotkeys("On", "OPTGuiClose")                             ; Turn on global hotkeys

    OptGUI.Hide()
    g_LOG.Debug("OPTGuiClose: OptGUI.Hide...OK")
    return
}

ToggleGlobalHotkeys(mode, caller := "") {
    HotIfWinActive
    for _, hk in ["GlobalHotkey1", "GlobalHotkey2"] {
        if (g_HOTKEY[hk] = "")
            continue
        try {
            Hotkey(g_HOTKEY[hk], ToggleWindow, mode)
            g_LOG.Debug(caller ": Turn " mode " " hk "...OK")
        } catch as e {
            g_LOG.Debug(caller ": Turn " mode " " hk "...Failed: " e.Message)
        }
    }
}

; NOTE: there is no more LoadConfig() - AppData.LoadAppData() (see the JSON command
; storage section) loads Config/Gui/Hotkey/Usage/History/Benchmark from
; ALTRun.json at startup, in one pass alongside the commands.

SaveConfig() {
    Global OptListView

    OptGUI.Submit()
    checkedRows := Map(), row := 0
    while (row := OptListView.GetNext(row, "C"))
        checkedRows[row] := 1

    ; Tab1 checklist values (plain booleans, no type coercion needed).
    for key, _ in g_CONFIG_P1
        g_CONFIG[key] := checkedRows.Has(A_Index) ? 1 : 0

    static configKeys := Array("FileMgr", "Everything", "HistoryLen", "RunCount"
        , "AutoSwitchDir", "IndexDir", "IndexType", "IndexDepth"
        , "IndexExclude", "IndexStoreApp", "DialogWin", "FileMgrID", "ExcludeWin")

    for _, key in configKeys
        g_CONFIG[key] := CoerceLikeCurrent(g_CONFIG[key], GetOptCtrlValue(OptGUI[key]))

    for key, _ in g_GUI
        g_GUI[key] := CoerceLikeCurrent(g_GUI[key], GetOptCtrlValue(OptGUI[key]))

    for key, _ in g_HOTKEY
        g_HOTKEY[key] := GetOptCtrlValue(OptGUI[key])                   ; Hotkeys/window titles are always text

    AppData.SaveAppData()

    g_LOG.Debug("SaveConfig: Save config...OK")
    return
}

GetOptCtrlValue(ctrl) {
    return InStr(",CheckBox,Slider,Hotkey,", "," ctrl.Type ",") ? ctrl.Value : ctrl.Text
}

CoerceLikeCurrent(currentVal, newVal) {                                 ; Keep a setting's number-vs-text type stable across saves,
    return (Type(currentVal) = "Integer" || Type(currentVal) = "Float") ; so JSON.stringify writes e.g. 300 instead of "300", while
        && IsNumber(newVal) ? newVal + 0 : newVal                       ; a hex color string like "0xFFFFFF" is left alone.
}

; ==================== Built-in Functions =========================
AhkRun() {
    try {
        Run(g_RUNTIME["Arg"])
    } catch as e {
        g_LOG.Debug("AhkRun: Error occur=" . e.Message)
    }
    return
}

TurnMonitorOff() {                                                      ; 关闭显示器:
    SendMessage(0x112, 0xF170, 2, , "Program Manager")                  ; 0x112 is WM_SYSCOMMAND, 0xF170 is SC_MONITORPOWER, 使用 -1 代替 2 来打开显示器, 使用 1 代替 2 来激活显示器的节能模式.
}

EmptyRecycle() {
    local Result := MsgBox("Do you really want to empty the Recycle Bin?", , "YesNo")
    if (Result = "Yes")
    {
        FileRecycleEmpty
    }
    return
}

MuteVolume() {
    SoundSetMute(true)
}

Google() {
    word := g_RUNTIME["Arg"] = "" ? A_Clipboard : g_RUNTIME["Arg"]
    Run("https://www.google.com/search?q=" word "&newwindow=1")
}

Bing() {
    word := g_RUNTIME["Arg"] = "" ? A_Clipboard : g_RUNTIME["Arg"]
    Run("https://cn.bing.com/search?q=" word)
}

Everything() {
    try {
        Run(g_CONFIG["Everything"] . ' -s `"' g_RUNTIME["Arg"] '`"')
    } catch as e {
        MsgBox("Everything software not found.`n`nPlease check ALTRun setting and Everything program file.`n`nError message=" . e.Message)
    }
    return
}
; SetLanguage()/ReadChineseFlag() used to live here; both are now Lang.Load()/
; Lang.IsChinese() in Lib\Language.ahk (see the #Include list at the top of
; this file and the "Global g_LNG := Lang.Load()" call near the top).
; Eval()/EvalSimple() used to live here; both are now Calc.Eval() in Lib/Util.ahk.

;;==================== Performance Test Only =========================
; Not wired to any hotkey/menu/command - run BenchmarkRun() manually from an
; editor/debugger when you want a search-performance snapshot.

BenchmarkRun(rounds := 10) {
    Global g_LOG, g_COMMANDS, myInputBox
    rounds := Max(10, rounds)

    static queries := [
        "n", "no", "note", "core", "kanji", "a t", "wi ex", "sys in"
        , "plugin", "open", "update", "help", "xyz_not_found"
        , "2+3*5", "5+5", "12345*6789", "nir", "control panel"
        , "new", "edit", "delete", "reload", "reindex", "history", "option", "usage"
        , "google test", "bing test", "notepad", "explorer", "cmd", "powershell", "service"
        , "disk", "device", "event", "task", "reg", "calc", "paint", "startup"
    ]

    Activate()
    myInputBox.Focus()

    for _, q in queries
    {
        myInputBox.Value := q
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
            myInputBox.Value := q
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

ExtractRes() {
    static file := A_Temp "\ALTRun_" A_ScriptHwnd ".jpg"
    if FileExist(file)
        return file

    base64 := "
    (
    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"
    )"
    bin := Buffer(StrLen(base64) * 3 // 4)
    size := bin.Size
    ok := DllCall("Crypt32.dll\CryptStringToBinary", "Str", base64, "UInt", 0, "UInt", 1, "Ptr", bin, "UIntP", size, "Ptr", 0, "Ptr", 0, "Int")
    if (!ok || size <= 0)
        return ""

    f := FileOpen(file, "w", "CP0")
    if !f
        return ""

    f.RawWrite(bin, size)
    f.Close()
    return file
}

