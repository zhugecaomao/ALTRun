;===================================================
; ALTRun - An effective launcher for Windows
; https://github.com/zhugecaomao/ALTRun
;===================================================
#Requires AutoHotkey v2.0
#SingleInstance Force
#NoTrayIcon
#Warn All, OutputDebug
SetWorkingDir(A_ScriptDir)
FileEncoding("UTF-8")

;@Ahk2Exe-SetName ALTRun
;@Ahk2Exe-SetDescription ALTRun - An effective launcher for Windows
;@Ahk2Exe-SetVersion g_TITLE
;@Ahk2Exe-SetCopyright Copyright (c) 2025 zhugecaomao
;@Ahk2Exe-SetOrigFilename ALTRun.ahk


;===================================================
; Declare Global Variables
; Built-in classes such as Object; they are predefined as global variables.
; Variables accessed or created inside an assume-local function are local by default
; Global variables which are only read by the function, not assigned or used with the reference operator (&).
;===================================================
Global g_LOG   := Logger(A_Temp . "\ALTRun.log")
Global g_INI   := A_ScriptDir "\ALTRun.ini"
Global g_VER   := "2025.09.22"
Global g_TITLE := "ALTRun - Ver " g_VER

Global g_COMMANDS := Array()         ; All commands
Global g_CMDINDEX := Array()         ; Searchable text for All commands
Global g_FALLBACK := Array()         ; Fallback commands
Global g_HISTORYS := Array()         ; Execution history
Global g_MATCHED  := Array()         ; Matched commands
Global g_TYPELST  := Array("File", "Dir", "Cmd", "URL", "Func")
Global g_SECTION  := Map(
    "CONFIG"    , "Config",
    "GUI"       , "Gui",
    "DFTCMD"    , "DefaultCommand",
    "USERCMD"   , "UserCommand",
    "FALLBACK"  , "FallbackCommand",
    "HOTKEY"    , "Hotkey",
    "HISTORY"   , "History",
    "INDEX"     , "Index",
    "USAGE"     , "Usage"
)

Global g_CONFIG := Map(
    "AutoStartup"    , 1,
    "EnableSendTo"   , 1,
    "InStartMenu"    , 1,
    "ShowTrayIcon"   , 1,
    "HideOnLostFocus", 1,
    "AlwaysOnTop"    , 1,
    "ShowCaption"    , 1,
    "XPthemeBg"      , 1,
    "EscClearInput"  , 1,
    "KeepInput"      , 1,
    "ShowIcon"       , 1,
    "SendToGetLnk"   , 1,
    "SaveHistory"    , 1,
    "SaveLog"        , 1,
    "MatchPath"      , 0,
    "ShowGrid"       , 0,
    "ShowHdr"        , 1,
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
    "Chinese"        , InStr("7804,0004,0804,1004", A_Language) ? 1 : 0,
    "MatchPinyin"    , 1,
    "RunCount"       , 0,
    "HistoryLen"     , 10,
    "AutoSwitchDir"  , 0,
    "FileMgr"        , "Explorer.exe",
    "IndexDir"       , "A_ProgramsCommon,A_StartMenu,C:\Other\Index\Location",
    "IndexType"      , "*.lnk,*.exe",
    "IndexDepth"     , 2,
    "IndexExclude"   , "Uninstall *",
    "Everything"     , "C:\Apps\Everything.exe",
    "DialogWin"      , "ahk_class #32770",
    "FileMgrID"      , "ahk_class CabinetWClass, ahk_class TTOTAL_CMD",
    "ExcludeWin"     , "ahk_class SysListView32, ahk_exe Explorer.exe"
)

g_LOG.Debug("///// ALTRun is starting /////")
SetLanguage()

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
    "MatchPinyin"    , g_LNG[132]
)

Global g_CONFIG_P2 :=Map(
    "FileMgr"        , "Explorer.exe",
    "Everything"     , "C:\Apps\Everything.exe",
    "HistoryLen"     , 10,
    "RunCount"       , 0,
    "AutoSwitchDir"  , 0,
    "IndexDir"       , "A_ProgramsCommon,A_StartMenu,C:\Other\Index\Location",
    "IndexType"      , "*.lnk,*.exe",
    "IndexDepth"     , 2,
    "IndexExclude"   , "Uninstall *",
    "DialogWin"      , "ahk_class #32770",
    "FileMgrID"      , "ahk_class CabinetWClass, ahk_class TTOTAL_CMD",
    "ExcludeWin"     , "ahk_class SysListView32, ahk_exe Explorer.exe",
)

Global g_HOTKEY := Map(
    "Hotkey1"       , "F1",
    "Trigger1"      , "Help",
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
    "GlobalHotkey1" , "!Space",
    "GlobalHotkey2" , "!r",
    "TotalCMDDir"   , "^g",
    "ExplorerDir"   , "^e",
    "AutoDateAtEnd" , "ahk_class TCmtEditForm,ahk_exe Notepad4.exe",
    "AutoDateAEHKey", "^d",
    "AutoDateBefExt", "ahk_class CabinetWClass,ahk_class Progman,ahk_class WorkerW,ahk_class #32770",
    "AutoDateBEHKey", "^d"
)

Global g_GUI := Map(
    "ListRows"      , 9,
    "ColWidth"      , "36,0,300,AutoHdr",
    "MainGUIFont"   , "Microsoft YaHei,norm s9",
    "OptGUIFont"    , "Microsoft YaHei,norm s9",
    "MainSBFont"    , "Microsoft YaHei,norm s8",
    "WinX"          , 660,
    "WinY"          , 300,
    "CtrlColor"     , "Default",
    "WinColor"      , "Silver",
    "Background"    , "ALTRun.jpg",
    "Transparency"  , 230
)

Global g_RUNTIME := Map(
    "CurrentCommand", "",
    "UseDisplay"    , 0,
    "UseFallback"   , 0,
    "Arg"           , "",
    "OneDrive"      , EnvGet("OneDrive"),
    "RegEx"         , "imS)",
    "Max"           , 1,
    "AppDate"       , A_YYYY . A_MM . A_DD
)

Global g_USAGE := Map(A_YYYY . A_MM . A_DD, 1)

Global MainGUI                                                          ; Global variables which are only read by the function, not assigned or used with the reference operator (&).
Global myListView
Global myInputBox
Global OptGUI
Global OptListView
Global CmdMgrGUI
Global myImageList := IL_Create(10, 5, 0)                               ; Create an ImageList so that the ListView can display some icons
Global myIconMap   := Map("DIR", IL_Add(myImageList,"imageres.dll",-3)  ; Icon cache index, IconIndex=1/2/3/4 for type dir/func/url/eval
                        ,"FUNC", IL_Add(myImageList,"imageres.dll",-100)
                        ,"URL",  IL_Add(myImageList,"imageres.dll",-144)
                        ,"EVAL", IL_Add(myImageList,"imageres.dll",-182))

LoadConfig("initialize")                                                ; Load ini config, iniWrite create ini whenever not exist
LoadCommands()
LoadHistory()
UpdateSendTo()
UpdateStartup()
UpdateStartMenu()
SetTrayMenu()                                                           ; SetTrayMenu before SetMainGUI, GUI window uses the tray icon that was in effect at the time the window was created
SetMainGUI()                                                            ; Create and set main GUI
RegisterHotkey()
Listary()
Plugins()
return
;;==================== Autorun until here =========================

SetMainGUI() {
    Global MainGUI, myInputBox, myListView, myStatusBar, myImageList

    Input_W := g_GUI["WinX"] - 24 - g_CONFIG["ShowBtnRun"] * 90 - g_CONFIG["ShowBtnOpt"] * 90
    Enter_W := g_CONFIG["ShowBtnRun"] * 80
    Enter_X := g_CONFIG["ShowBtnRun"] * 10
    Opt_W   := g_CONFIG["ShowBtnOpt"] * 80
    Opt_X   := g_CONFIG["ShowBtnOpt"] * 10

    MainGUI := Gui("+Owner" (g_CONFIG["AlwaysOnTop"] ? " AlwaysOnTop" : ""), g_TITLE)
    MainGUI.Hide()  ; 尝试解决屏幕左上角一闪而过的窗口
    MainGUI.OnEvent("Close", MainGuiClose)
    MainGUI.OnEvent("Escape", MainGuiEscape)
    MainGUI.BackColor := g_GUI["WinColor"]
    MainGUI.Opt((g_CONFIG["ShowCaption"] ? "" : "-Caption") (g_CONFIG["XPthemeBg"] ? "" : " -Theme"))
    MainGUI.SetFont(StrSplit(g_GUI["MainGUIFont"], ",")[2], StrSplit(g_GUI["MainGUIFont"], ",")[1])
    myInputBox := MainGUI.AddEdit("x12 y10 -WantReturn W" Input_W, g_LNG[13])
    myInputBox.OnEvent("Change", OnSearchInput)
    MainGUI.AddButton("x+" Enter_X " yp W" Enter_W " hp Default Hidden" (!g_CONFIG["ShowBtnRun"]), g_LNG[11]).OnEvent("Click", RunCurrentCommand)
    MainGUI.AddButton("x+" Opt_X " yp W" Opt_W " hp Hidden" (!g_CONFIG["ShowBtnOpt"]), g_LNG[12]).OnEvent("Click", (*) => Options(1))
    myListView := MainGUI.AddListView("x12 yp+36 W" (g_GUI["WinX"] - 24) " H" (g_GUI["WinY"] - 76) " -Multi", g_LNG[10])
    myListView.Opt((g_CONFIG["DoubleBuffer"] ? " +LV0x10000" : "") (g_CONFIG["ShowHdr"] ? "" : " -Hdr") (g_CONFIG["ShowGrid"] ? " Grid" : "") (g_CONFIG["ShowBorder"] ? "" : " -E0x200"))
    myListView.OnEvent("Click", OnListviewClick)
    myListView.OnEvent("ContextMenu", OnListViewContextMenu)
    myListView.OnEvent("DoubleClick", LVRunCommand)
    myListView.SetImageList(myImageList)                                ; Attach the ImageList to the ListView so that it can later display the icons
    Loop 4 {
        if (StrSplit(g_GUI["ColWidth"], ",").Length >= A_Index) {
            if (StrSplit(g_GUI["ColWidth"], ",")[A_Index] != "")
                myListView.ModifyCol(A_Index, StrSplit(g_GUI["ColWidth"], ",")[A_Index])
        }
    }

    if FileExist(g_GUI["Background"])
        try MainGUI.AddPic("x0 y0 0x4000000", g_GUI["Background"])

    myStatusBar := MainGUI.AddStatusBar("Hidden" (!g_CONFIG["ShowStatusBar"]))
    myStatusBar.OnEvent("Click", OnStatusBarClick)
    myStatusBar.OnEvent("ContextMenu", OnStatusBarContextMenu)
    myStatusBar.SetFont(StrSplit(g_GUI["MainSBFont"], ",")[2], StrSplit(g_GUI["MainSBFont"], ",")[1])
    myStatusBar.SetParts(g_GUI["WinX"] - 90 * g_CONFIG["ShowRunCount"])
    myStatusBar.SetIcon("imageres.dll",-150, 2)

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
                Path := A_Args[2]

                SplitPath(Path, &Desc, , &fileExt)                      ; Extra name from _Path (if _Type is dir and has "." in path, nameNoExt will not get full folder name)

                fileType := InStr(FileExist(Path), "D") ? "Dir" : "File" ; Default Type is File, Set Type is Dir only if the file exists and is a directory
                
                if (fileExt = "lnk" && g_CONFIG["SendToGetLnk"]) {
                    FileGetShortcut(Path, Path, , &fileArg, &Desc)
                    Path .= " " fileArg
                }
                CmdMgr(g_SECTION["USERCMD"], fileType, Path, Desc, 1, "")              ; Add new command to database
            }
        }
    }

    if (g_GUI["Transparency"] < 250){
        WinSetTransparent(g_GUI["Transparency"], MainGUI.Hwnd)          ; By default, hidden windows are not detected. however, when using pure HWNDs, hidden windows are always detected regardless of DetectHiddenWindows.
    }

    MainGUI.Show(HideWin "w" g_GUI["WinX"] " h" g_GUI["WinY"] " Center")

    if g_CONFIG["HideOnLostFocus"]
        OnMessage(0x06, WM_ACTIVATE)

    return
}

; 创建任务栏托盘程序图标
SetTrayMenu() {
    static myTrayMenu := ""  ; 只在第一次调用时创建

    if !IsObject(myTrayMenu) {
        myTrayMenu := A_TrayMenu
        try {
            TraySetIcon("imageres.dll", -100)
            
            myTrayMenu.Delete() ; 删除默认项
            myTrayMenu.Add(g_LNG[300], ToggleWindow)
            myTrayMenu.Add(g_LNG[301], (*) => Options(1))
            myTrayMenu.Add(g_LNG[302], Reindex)
            myTrayMenu.Add(g_LNG[303], Usage)
            myTrayMenu.Add(g_LNG[309], Update)
            myTrayMenu.Add()  ; 分隔线
            myTrayMenu.Add(g_LNG[304], Help)
            myTrayMenu.Add(g_LNG[305], (*) => ListLines())
            myTrayMenu.Add()
            myTrayMenu.Add(g_LNG[307], (*) => Reload())
            myTrayMenu.Add(g_LNG[308], (*) => ExitApp())

            myTrayMenu.SetIcon(g_LNG[300], "imageres.dll", -100)
            myTrayMenu.SetIcon(g_LNG[301], "imageres.dll", -114)
            myTrayMenu.SetIcon(g_LNG[302], "imageres.dll", -8)
            myTrayMenu.SetIcon(g_LNG[303], "imageres.dll", -150)
            myTrayMenu.SetIcon(g_LNG[309], "imageres.dll", -5338)
            myTrayMenu.SetIcon(g_LNG[304], "imageres.dll", -99)
            myTrayMenu.SetIcon(g_LNG[305], "imageres.dll", -165)
            myTrayMenu.SetIcon(g_LNG[307], "imageres.dll", -5311)
            myTrayMenu.SetIcon(g_LNG[308], "imageres.dll", -98)

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

RegisterHotkey() {
    try {
        Hotkey(g_HOTKEY["GlobalHotkey1"], ToggleWindow)                 ; 全局热键
        Hotkey(g_HOTKEY["GlobalHotkey2"], ToggleWindow)

        g_LOG.Debug("RegisterHotkey: Set GlobalHotkeys...OK")
    } catch as e {
        g_LOG.Debug("RegisterHotkey: Failed to set GlobalHotkeys..." e.Message)
    }

    HotIfWinActive(g_TITLE)                                             ; 窗口特定热键
    try {
        Hotkey("!F4"        , (*) => ExitApp())
        Hotkey("Tab"        , TabFunc)
        Hotkey("F1"         , Help)
        Hotkey("F2"         , (*) => Options(1))
        Hotkey("F3"         , EditCommand)
        Hotkey("F4"         , UserCommand)
        Hotkey("^q"         , ReStart)
        Hotkey("^d"         , OpenContainer)
        Hotkey("^c"         , CopyCommand)
        Hotkey("^n"         , NewCommand)
        Hotkey("Del"        , DelCommand)
        Hotkey("^i"         , Reindex)
        Hotkey("Down"       , NextCommand)
        Hotkey("Up"         , PrevCommand)
        Hotkey("^NumpadAdd" , RankUp)
        Hotkey("^NumpadSub" , RankDown)

        g_LOG.Debug("RegisterHotkey: Set ALTRun window Hotkey...OK")
    } catch as e {
        g_LOG.Debug("RegisterHotkey: Failed to set hotkey..." e.Message)
    }
    
    Loop g_GUI["ListRows"] {
        try {
            Hotkey("!" . A_Index, RunSelectedCommand)                   ; 通过热键选择并运行指定命令 = Alt + index (1-9)
            Hotkey("^" . A_Index, GotoCommand)                          ; 通过热键选择指定命令 = Ctrl + index (1-9)

            g_LOG.Debug("RegisterHotkey: Set CommandList Hotkey" A_Index "  finish")
        } catch as e {
            g_LOG.Debug("RegisterHotkey: Failed to set CommandList hotkey " A_Index . e.Message)
        }
    }

    Loop 7 {
        KeyName := "Hotkey"  . A_Index
        Trigger := "Trigger" . A_Index
        if (g_HOTKEY.Has(KeyName) && g_HOTKEY[Trigger] != "") {         ; 自定义热键执行指定功能 = Hotkey + Trigger
            try {
                Hotkey(g_HOTKEY[KeyName], ExecuteFunc.Bind(, g_HOTKEY[Trigger], A_Index))
                g_LOG.Debug("RegisterHotkey: Set MainGUI FuncList hotkey finish: " g_HOTKEY[KeyName] " <-> " g_HOTKEY[Trigger])
            } catch as e {
                g_LOG.Debug("RegisterHotkey: Set MainGUI FuncList hotkey failed, error="  e.Message)
            }
        }
    }

    HotIfWinActive g_HOTKEY["CondTitle"]                                ; 条件热键执行指定功能 = CondHotkey + CondAction
    if g_HOTKEY.Has("CondTitle") && g_HOTKEY.Has("CondHotkey") && g_HOTKEY.Has("CondAction") {
        try {
            Hotkey g_HOTKEY["CondHotkey"], g_HOTKEY["CondAction"]
            g_LOG.Debug("RegisterHotkey: Conditional hotkey " g_HOTKEY["CondHotkey"] " set for " g_HOTKEY["CondTitle"])
        } catch as e {
            g_LOG.Debug("RegisterHotkey: Failed to set conditional hotkey..." e.Message)
        }
    }
    HotIfWinActive                                                      ; Turn off context, make subsequent hotkeys global again
    return
}

ExecuteFunc(HotkeyName, TriggerName, Index) {
    RunCommand("FUNC | " TriggerName)
    OutputDebug("ExecutFunc: Through Hotkey execute function=" TriggerName)
}

Activate() {
    MainGUI.Show()

    if (WinWaitActive(g_TITLE, , 3)) {                                  ; Wait for the window to be active
        myInputBox.Focus()
        SendMessage(0xB1, 0, -1, myInputBox.Hwnd)                       ; EM_SETSEL (0xB1)
    }
    UpdateUsage()
}

ToggleWindow(*) {
    WinActive(g_TITLE) ? MainGuiClose() : Activate()
}

OnSearchInput(*) {
    SearchCommand(myInputBox.Value)
}

SearchCommand(command := "") {
    Global g_MATCHED, g_RUNTIME, g_FALLBACK, g_COMMANDS, g_CMDINDEX, g_GUI, g_CONFIG
    g_MATCHED := Array()
    Prefix    := SubStr(command, 1, 1)

    ; Handle fallback commands
    if (Prefix = "+" or Prefix = " " or Prefix = ">") {
        idx := (Prefix = "+") ? 1 : (Prefix = " ") ? 2 : 3
        g_RUNTIME["CurrentCommand"] := g_FALLBACK[idx]                  ; Corresponding to fallback commands position no. 1, 2 & 3
        g_MATCHED.Push(g_RUNTIME["CurrentCommand"])
        return ListResult(g_MATCHED)
    }

    ; Search commands using precomputed searchable text
    for index, strToSearch in g_CMDINDEX {
        if FuzzyMatch(strToSearch, command) {
            g_MATCHED.Push(g_COMMANDS[index])
            g_RUNTIME["CurrentCommand"] := g_MATCHED[1]

            if g_MATCHED.Length >= g_GUI["ListRows"]
                break
        }
    }

    if (g_MATCHED.Length = 0) {                                         ; No matched command found
        evalResult := Eval(command)
        if (IsNumber(evalResult) && evalResult != 0) {
            g_MATCHED := StruCalcResult(evalResult)
            return ListResult(g_MATCHED, True)
        }

        g_RUNTIME["UseFallback"]    := True
        g_MATCHED                   := g_FALLBACK
        g_RUNTIME["CurrentCommand"] := g_FALLBACK[1]
    } Else {
        g_RUNTIME["UseFallback"]    := False
    }
    return ListResult(g_MATCHED)
}

ListResult(ArrayToList := "", UseDisplay := false) {
    myListView.Opt("-Redraw")                                           ; Improve performance by disabling redrawing during load.
    myListView.Delete()
    g_RUNTIME["UseDisplay"] := UseDisplay

    for index, command in ArrayToList {
        splitResult := StrSplit(command, " | ")
        _Type := splitResult.Length >= 1 ? splitResult[1] : ""          ; Ensure _Type has default
        _Path := splitResult.Length >= 2 ? splitResult[2] : ""          ; Ensure _Path has default
        _Desc := splitResult.Length >= 3 ? splitResult[3] : ""          ; Ensure _Desc has default (fix for missing 3rd element)
        IconIndex := GetIconIndex(_Path, _Type)
        
        SplitPath(_Path, &fileName)                                     ; Extra name from _Path (if _Type is Dir and has "." in path, nameNoExt will not get full folder name)
        PathToShow := (g_CONFIG["ShortenPath"]) ? fileName : _Path      ; Show Full path / Shorten path

        myListView.Add("Icon" IconIndex, (g_CONFIG["ShowSN"] ? A_Index : ""), _Type, PathToShow, _Desc)
    }

    statusBarText := myListView.GetText(1, 3)
    if (g_RUNTIME["CurrentCommand"] != "") {
        statusBarText := StrSplit(g_RUNTIME["CurrentCommand"], " | ")[2]
    }

    myListView.Modify(1, "Select Focus Vis")                            ; Select 1st row
    myListView.Opt("+Redraw")                                           ; Re-enable redrawing.
    SetStatusBar(statusBarText)
}

GetIconIndex(_Path, _Type) {                                            ; Get file's icon index
    Global myIconMap
    if not g_CONFIG["ShowIcon"]
        Return 0                                                        ; No icon to show, return 0

    if (_Type = "") {
        return 0
    } else if (_Type = "DIR") {
        return 1
    } else if InStr("FUNC,CMD,TIP,提示", _Type, 0) {
        return 2
    } else if (_Type = "URL") {
        return 3
    } else if (_Type = "EVAL") {
        return 4
    } else if (_Type = "FILE") {
        _Path := AbsPath(_Path)                                         ; Must store in var for afterward use, trim space (in AbsPath)
        SplitPath(_Path, , , &fileExt)                                  ; Get the file's extension.
        if (fileExt ~= "^(?i:EXE|ICO|ANI|CUR|LNK)$") {                  ; File types that have their own icon
            IconIndex := myIconMap.Has(_Path) ? myIconMap[_Path] : GetIcon(_Path, _Path) ; File path exist in ImageList, get the index, several calls can be avoided and performance is greatly improved
        } else {                                                        ; Some other extension/file-type like pdf or xlsx
            IconIndex := myIconMap.Has(fileExt) ? myIconMap[fileExt] : GetIcon(_Path, fileExt)
        }
        Return IconIndex
    }
}

GetIcon(_Path, ExtOrPath) {                                             ; Get file's icon
    Global myImageList, myIconMap
    sfi_size := A_PtrSize + 688
    sfi      := Buffer(sfi_size)                                        ; Calculate buffer size required for SHFILEINFO structure. VarSetStrCapacity change to Buffer

    if not DllCall("Shell32\SHGetFileInfoW", "Str", _Path, "UInt", 0, "Ptr", sfi, "UInt", sfi_size, "UInt", 0x101) ; 0x101 is SHGFI_ICON+SHGFI_SMALLICON
        IconIndex := 9999999                                            ; Set it out of bounds to display a blank icon.
    else {                                                              ; Icon successfully loaded. Extract the hIcon member from the structure
        hIcon := NumGet(sfi, 0, "Ptr")                                  ; Add the HICON directly to the small-icon lists.
        IconIndex := DllCall("ImageList_ReplaceIcon", "ptr", myImageList, "int", -1, "ptr", hIcon) + 1 ; Uses +1 to convert the returned index from zero-based to one-based:
        DllCall("DestroyIcon", "Ptr", hIcon)                            ; Now that it's been copied into the ImageLists, the original should be destroyed
        myIconMap[ExtOrPath] := IconIndex                               ; Cache the icon based on file type (xlsx, pdf) or path (exe, lnk) to save memory and improve loading performance
    }
    Return IconIndex
}

AbsPath(Path, KeepRunAs := False) {                                     ; Convert path to absolute path
    Path := Trim(Path)

    if (!KeepRunAs)
        Path := StrReplace(Path,  "*RunAs ", "")                        ; Remove *RunAs (Admin Run) to get absolute path

    if (InStr(Path, "A_") = 1 && InStr(Path, "\")) {                    ; Resolve path like A_ScriptDir, some server path has "Plot A_IGLS" in it, so InStr must be 1
        SubParts := StrSplit(Path, " ", " `t")
        if (SubParts.Length > 1 && InStr(SubParts[1], "A_") = 1) {
            VarName  := SubParts[1]
            VarValue := %VarName%
            Path := VarValue . StrReplace(SubParts[2], "`"", "")
        }
    } else if (InStr(Path, "A_") = 1) {
        Path := %Path%
    }

    Path := StrReplace(Path, "%Temp%", A_Temp)
    Path := StrReplace(Path, "%OneDrive%", g_RUNTIME["OneDrive"])       ; Convert OneDrive to absolute path due to #NoEnv
    Return Path
}

RelativePath(Path) {                                                    ; Convert path to relative path
    Path := StrReplace(Path, A_Temp, "%Temp%")
    Path := StrReplace(Path, g_RUNTIME["OneDrive"], "%OneDrive%")
    Return Path
}

RunCommand(originCmd) {
    MainGuiClose()
    ParseArg()
    g_RUNTIME["UseDisplay"] := false

    splitResult := StrSplit(originCmd, " | ")

    _Type := splitResult.Length >= 1 ? splitResult[1] : ""              ; Ensure _Type has default
    _Path := splitResult.Length >= 2 ? AbsPath(splitResult[2], True) : "" ; Ensure _Path has default

    if (_Type = "") {
        return
    } else if (_Type = "DIR") {
        OpenDir(_Path)
    } else if (_Type = "FUNC") {
        try {
            %_Path%()
        } catch as e {
            MsgBox("Could not find function: " _Path "`n`nError message: " e.Message, g_TITLE, 4096)
        }
    } else {                                                            ; For _Type = "FILE","URL","CMD" and other Unknown type
        try {
            Run(_Path)
        } catch as e {
            MsgBox("Could not run command: " _Path "`n`nError message: " e.Message, g_TITLE, 4096)
        }
    }

    if (g_CONFIG["SaveHistory"]) {
        g_HISTORYS.InsertAt(1, originCmd " Arg=" g_RUNTIME["Arg"])      ; Adjust command history

        if (g_HISTORYS.Length > g_CONFIG["HistoryLen"]) {
            g_HISTORYS.Pop()
        }

        IniDelete(g_INI, g_SECTION["HISTORY"])
        for index, element in g_HISTORYS
            IniWrite(element, g_INI, g_SECTION["HISTORY"], index) ; Save command history
    }

    UpdateRunCount()
    if (g_CONFIG["SmartRank"]) {
        UpdateRank(originCmd)
    }
    g_LOG.Debug("RunCommand: Execute " g_CONFIG["RunCount"] " = " originCmd)
    OutputDebug("RunCommand: Execute " g_CONFIG["RunCount"] " = " originCmd)
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

    if (index <= g_MATCHED.Length) {
        ChangeCommand(index, True)
        g_RUNTIME["CurrentCommand"] := g_MATCHED[index]
    }
}

RunSelectedCommand(*) {
    GotoCommand()
    RunCommand(g_RUNTIME["CurrentCommand"])
}

ChangeCommand(Step := 1, ResetSelRow := False) {
    selectedRow := ResetSelRow ? Step : myListView.GetNext() + Step     ; Get target row no. to be selected
    selectedRow := selectedRow > myListView.GetCount() ? 1 : selectedRow ; Listview cycle selection (Mod has bug on upward cycle)
    selectedRow := selectedRow < 1 ? myListView.GetCount() : selectedRow

    if (g_MATCHED.Length >= selectedRow) {
        g_RUNTIME["CurrentCommand"] := g_MATCHED[selectedRow]           ; Get current command from selected row
        SetStatusBar(StrSplit(g_RUNTIME["CurrentCommand"], " | ")[2])
    } else {
        SetStatusBar(myListView.GetText(selectedRow, 3))
    }

    myListView.Modify(selectedRow, "Select Focus Vis")                  ; make new index row selected, Focused, and Visible
    OutputDebug("ChangeCommand: Selecting Row=" selectedRow)
}

OnListviewClick(myListView, rowNumber) {                                ; On ListView left click
    if (!rowNumber)                                                     ; 没有焦点行
        Return

    if (g_MATCHED.Length >= rowNumber) {
        g_RUNTIME["CurrentCommand"] := g_MATCHED[rowNumber]             ; Get current command from focused row
        SetStatusBar(StrSplit(g_RUNTIME["CurrentCommand"], " | ")[2])
    } else {
        SetStatusBar(myListView.GetText(rowNumber, 3))
    }
}

OnListViewContextMenu(GuiCtrlObj, rowNumber, IsRightClick, X, Y) {      ; On ListView ContextMenu
    Global myListView, g_MATCHED, g_RUNTIME, g_LNG, g_LOG
    if (!rowNumber)                                                     ; If the user clicked somewhere other than a row
        Return

    if (g_MATCHED.Length >= rowNumber) {
        g_RUNTIME["CurrentCommand"] := g_MATCHED[rowNumber]
        SetListViewContextMenu(X, Y)                                        ; Show ContextMenu on right click
    } else {                                                            ; For cases like first hint page
        SetStatusBar(myListView.GetText(rowNumber, 3))
    }
}

SetListViewContextMenu(X, Y) {
    static myListViewContextMenu := ""  ; 只在第一次调用时创建

    if !IsObject(myListViewContextMenu) {
        myListViewContextMenu := Menu()
        myListViewContextMenu.Add(g_LNG[400], LVRunCommand)
        myListViewContextMenu.Add(g_LNG[401], OpenContainer)
        myListViewContextMenu.Add(g_LNG[402], CopyCommand)
        myListViewContextMenu.Add()
        myListViewContextMenu.Add(g_LNG[403], NewCommand)
        myListViewContextMenu.Add(g_LNG[404], EditCommand)
        myListViewContextMenu.Add(g_LNG[405], DelCommand)
        myListViewContextMenu.Add(g_LNG[406], UserCommand)

        myListViewContextMenu.SetIcon(g_LNG[400], "imageres.dll", -100)
        myListViewContextMenu.SetIcon(g_LNG[401], "imageres.dll", -3)
        myListViewContextMenu.SetIcon(g_LNG[402], "imageres.dll", -5314)
        myListViewContextMenu.SetIcon(g_LNG[403], "imageres.dll", -2)
        myListViewContextMenu.SetIcon(g_LNG[404], "imageres.dll", -5306)
        myListViewContextMenu.SetIcon(g_LNG[405], "imageres.dll", -5305)
        myListViewContextMenu.SetIcon(g_LNG[406], "imageres.dll", -88)
        myListViewContextMenu.Default := g_LNG[400]

        g_LOG.Debug("SetListViewContextMenu: Create myListViewContextMenu...OK")
        OutputDebug("SetListViewContextMenu: Create myListViewContextMenu...OK")
    }

    myListViewContextMenu.Show(X, Y)
}

LVRunCommand(*) {                                                       ; On ListView double click action
    focusedRow := myListView.GetNext(0, "Focused")                      ; Check focused row, only operate focusd row instead of all selected rows
    if (!focusedRow)                                                    ; Return if no focused row is found
        Return

    if (g_MATCHED.Length >= focusedRow) {
        g_RUNTIME["CurrentCommand"] := g_MATCHED[focusedRow]            ; Get current command from focused row
        RunCommand(g_RUNTIME["CurrentCommand"])                         ; Execute the command if the user selected "Run Enter"
    }
}

CopyCommand(*) {                                                        ; ListView ContextMenu
    focusedRow := myListView.GetNext(0, "Focused")                      ; Check focused row, only operate focusd row instead of all selected rows
    if (!focusedRow)                                                    ; Return if no focused row is found
        Return

    if (g_MATCHED.Length >= focusedRow) {
        g_RUNTIME["CurrentCommand"] := g_MATCHED[focusedRow]            ; Get current command from focused row
        A_Clipboard := myListView.GetText(focusedRow, 3)                       ; Get the text from the focusedRow's 3rd field.
    }
}

OnStatusBarClick(GuiCtrlObj, partNumber) {
    if (partNumber = 1) {
        A_Clipboard := StatusBarGetText(1, g_TITLE) ; Get text from 1st part of StatusBar
        SetStatusBar(g_LNG[407] " : " A_Clipboard)
    } else if (partNumber = 2) {
        Usage()
    }
}

OnStatusBarContextMenu(GuiCtrlObj, partNumber, RightClick, X, Y) {
    if (partNumber = 1) {
        SB_ContextMenu1 := Menu()
        SB_ContextMenu1.Add(g_LNG[407], (*) => SetStatusBar(g_LNG[407] " : " A_Clipboard := StatusBarGetText(1, g_TITLE))) ; Get text from 1st part of StatusBar
        SB_ContextMenu1.SetIcon(g_LNG[407], "imageres.dll", -5314)
        SB_ContextMenu1.Show()
    } else if (partNumber = 2) {
        SB_ContextMenu2 := Menu()
        SB_ContextMenu2.Add(g_LNG[408], Usage)
        SB_ContextMenu2.SetIcon(g_LNG[408], "imageres.dll", -150)
        SB_ContextMenu2.Show()
    }
}

MainGuiEscape(*) {
    (g_CONFIG["EscClearInput"] and myInputBox.Value) ? ClearInput() : MainGuiClose()
}

MainGuiClose(*) {
    if (!g_CONFIG["KeepInput"]) {
        ClearInput()
    }
    MainGUI.Hide()
    SetStatusBar("TIP")                                                 ; Update StatusBar tip information after GUI hide
}

ClearInput() {
    myInputBox.Focus()
    myInputBox.Text := ""
    OnSearchInput()                                                     ; v1 no need, v2 需要手动调用绑定的事件处理函数
}

Exit(*) {
    ExitApp()
}

ReStart(*) {
    Reload()
}

SetStatusBar(strToShow) {                                               ; Set StatusBar text, Mode 1: Current command (default), 2: Hint, 3: Any text
    if (strToShow = "TIP") {
        strToShow := g_LNG[51] g_LNG[Random(52, 71)]                   ; Randomly select a tip from hint list g_LNG 52~71
    } else {
        strToShow := strToShow
    }

    myStatusBar.SetText(strToShow, 1)
    myStatusBar.SetText(g_CONFIG["RunCount"], 2)
}

RunCurrentCommand(*) {
    RunCommand(g_RUNTIME["CurrentCommand"])
}

ParseArg() {
    commandPrefix := SubStr(myInputBox.Value, 1, 1)

    if (commandPrefix = "+" || commandPrefix = " " || commandPrefix = ">") {
        Return g_RUNTIME["Arg"] := SubStr(myInputBox.Value, 2)        ; 直接取命令为参数
    }

    if (InStr(myInputBox.Value, " ") && !g_RUNTIME["UseFallback"]) {  ; 用空格来判断参数
        g_RUNTIME["Arg"] := SubStr(myInputBox.Value, InStr(myInputBox.Value, " ") + 1)
    }
    else if (g_RUNTIME["UseFallback"]) {
        g_RUNTIME["Arg"] := myInputBox.Value
    }
    else {
        g_RUNTIME["Arg"] := ""
    }
}

FuzzyMatch(Haystack, Needle) {
    Needle := StrReplace(Needle, "+", "\+")                             ; For Eval (preceded by a backslash to be seen as literal)
    Needle := StrReplace(Needle, "*", "\*")                             ; For Eval (eg. 25+5 or 6*5 will show Eval result instead of match file with "30")
    Needle := StrReplace(Needle, "\", ".*")
    Needle := StrReplace(Needle, " ", ".*")                             ; 空格直接替换为匹配任意字符
    Return RegExMatch(Haystack, g_RUNTIME["RegEx"] . Needle)
}

UpdateRank(originCmd, showRank := false, inc := 1) {
    RANKSEC := g_SECTION["DFTCMD"] "|" g_SECTION["USERCMD"] "|" g_SECTION["INDEX"]
    Loop Parse, RANKSEC, "|"                                            ; Update Rank for related sections
    {
        Rank := IniRead(g_INI, A_LoopField, originCmd, "KeyNotFound")

        if (Rank = "KeyNotFound" or Rank = "ERROR" or originCmd = "")   ; If originCmd not exist in this section, then check next section
            continue                                                    ; Skips the rest of a loop and begins a new one.
        else if IsInteger(Rank)                                         ; If originCmd exist in this section, then update it's rank.
            Rank += inc
        else
            Rank := inc

        if (Rank < 0)                                                   ; 如果降到负数,都设置成 -1,然后屏蔽/排除
            Rank := -1

        IniWrite(Rank, g_INI, A_LoopField, originCmd)                   ; Update new Rank for originCmd

        if (showRank) {
            SetStatusBar("UpdateRank: Rank for current command : " Rank)
        }
    }
    LoadCommands()                                                      ; New rank will take effect in real-time by LoadCommands again
}

UpdateUsage() {
    currDate := A_YYYY . A_MM . A_DD
    if (g_RUNTIME["AppDate"] != currDate) {
        g_RUNTIME["AppDate"] := currDate
        g_USAGE[currDate] := 1
    } else {
        g_USAGE[currDate]++
    }
    g_RUNTIME["Max"] := Max(g_RUNTIME["Max"], g_USAGE[currDate])
    IniWrite(g_USAGE[currDate], g_INI, g_SECTION["USAGE"], currDate)
}

UpdateRunCount() {
    g_CONFIG["RunCount"]++
    IniWrite(g_CONFIG["RunCount"], g_INI, g_SECTION["CONFIG"], "RunCount") ; Update RunCount in g_CONFIG
}

RankUp(*) {
    UpdateRank(g_RUNTIME["CurrentCommand"], true)
}

RankDown(*) {
    UpdateRank(g_RUNTIME["CurrentCommand"], true, -1)
}

LoadCommands() {
    Global g_COMMANDS := Array()                                        ; Clear g_COMMANDS, g_FALLBACK, g_CMDINDEX (searchable text for all commands)
    Global g_CMDINDEX := Array()
    Global g_FALLBACK := Array()
    Local  RankString := ""

    for index, line in StrSplit(LoadConfig("commands"), "`n", "`r")     ; Read commands sections (built-in, user & index), read each line, separate key and value
    {
        if (!Trim(line) || SubStr(line, 1, 1) = ";")                    ; Skip empty line or comment line
            continue
    
        command := StrSplit(line, "=")[1]
        rank    := StrSplit(line, "=")[2]

        if (command != "" and rank > 0) {
            splitResult := StrSplit(command, " | ")
            type := splitResult[1]
            path := splitResult[2]
            desc := splitResult.Has(3) ? splitResult[3] : ""
            SplitPath(path, &filename)

            strToSearch := g_CONFIG["MatchPath"] ? path " " desc : filename " " desc ; search file name include extension, and desc (For MatchBeginning option, exclude "type")
            strToSearch := g_CONFIG["MatchPinyin"] ? GetFirstChar(strToSearch) : strToSearch ; 中文转为拼音首字母

            RankString .= rank "`t" command "`t" strToSearch "`n"
        }
    }

    ; Sort commands by rank (reverse numerical)
    RankString := Sort(RankString, "R N")
    for index, line in StrSplit(RankString, "`n", "`r")
    {
        if !Trim(line)
            continue  ; 跳过空行

        command     := StrSplit(line, "`t")[2]
        strToSearch := StrSplit(line, "`t")[3]
        g_COMMANDS.Push(command)
        g_CMDINDEX.Push(strToSearch)
    }

    ; Read FALLBACK section
    FALLBACKCMDSEC := ""
    Try FALLBACKCMDSEC := IniRead(g_INI, g_SECTION["FALLBACK"])
    if (FALLBACKCMDSEC = "") {
        IniWrite("
        (
        ; Fallback Commands show when search result is empty
        ; Commands in order, modify as desired
        ; Format: Command Type | Command | Description
        ; Command Type: File, Dir, CMD, URL
        ;
        Func | NewCommand | New Command
        Func | Everything | Search by Everything
        Func | Google | Search Clipboard or Input by Google
        Func | AhkRun | Run Command use AutoHotkey Run
        Func | Bing | Search Clipboard or Input by Bing
        Cmd | Calc.exe | Calculator
        )", g_INI, g_SECTION["FALLBACK"])
        FALLBACKCMDSEC := IniRead(g_INI, g_SECTION["FALLBACK"])
    }
    for line in StrSplit(FALLBACKCMDSEC, "`n") {
        if (line != "")
            g_FALLBACK.Push(line)
    }

    OutputDebug("LoadCommands: Loaded g_COMMANDS.Length=" g_COMMANDS.Length ", g_FALLBACK.Length=" g_FALLBACK.Length)
    g_LOG.Debug("LoadCommands: Loaded g_COMMANDS.Length=" g_COMMANDS.Length ", g_FALLBACK.Length=" g_FALLBACK.Length)
    return
}

LoadHistory() {
    if (g_CONFIG["SaveHistory"]) {
        Loop g_CONFIG["HistoryLen"] {
            Try history := IniRead(g_INI, g_SECTION["HISTORY"], A_Index, "")
            g_HISTORYS.Push(history)
        }
    } else
        Try IniDelete(g_INI, g_SECTION["HISTORY"])

    OutputDebug("LoadHistory: History length=" g_HISTORYS.Length)
    return
}

GetCmdOutput(command) {
    TempFile    := A_Temp . "\ALTRun.stdout"
    FullCommand := A_ComSpec " /C " command " > " TempFile

    RunWait(FullCommand, A_Temp, "Hide")
    Result := FileRead(TempFile)
    FileDelete(TempFile)
    Return RTrim(Result, "`r`n")                                        ; Remove result rightmost/last "`r`n"
}

GetRunResult(command) {                                                 ; 运行CMD并取返回结果方式2
    shell := ComObject("WScript.Shell")                                 ; WshShell object: https://msdn.microsoft.com/en-us/library/aew9yb99
    exec := shell.Exec(A_ComSpec " /C " command)                        ; Execute a single command via cmd.exe
    Return exec.StdOut.ReadAll()                                        ; Read and Return the command's output
}

OpenDir(Path) {
    Path := AbsPath(Path)

    Try{
        Run(g_CONFIG["FileMgr"] ' `"' Path '`"')
    } catch as e {
        MsgBox("Could not open dir: " Path "`n`nError message: " e.Message, g_TITLE, 4096)
    }
    g_LOG.Debug("OpenDir: Opening path=" Path " ,using FileMgr=" g_CONFIG["FileMgr"])
}

OpenContainer(*) {
    Path := AbsPath(StrSplit(g_RUNTIME["CurrentCommand"], " | ")[2])

    try {
        if (g_CONFIG["FileMgr"] = "Explorer.exe")
            Run(g_CONFIG["FileMgr"] ' /Select, `"' Path '`"')
        else
            Run(g_CONFIG["FileMgr"] ' /P `"' Path '`"')                 ; /P Parent folder
    } catch as e {
        MsgBox("Error finding parent directory: " . e.Message, g_TITLE, 4096)
    }
    g_LOG.Debug("OpenContainer: Path=" . Path)
}

WM_ACTIVATE(wParam, lParam, msg, hwnd){                                 ; Close on lose focus
    Global MainGUI, g_RUNTIME, g_LOG, g_TITLE

    if hwnd != MainGUI.Hwnd
        return 0  ; 忽略非主窗口的消息

    isActivated := (wParam > 0)
    g_LOG.Debug("WM_ACTIVATE: " (isActivated ? "Activated" : "Deactivated") ", hwnd=" hwnd)

    if (!isActivated && WinExist("ahk_id " MainGUI.Hwnd) && !g_RUNTIME["UseDisplay"]) {
        g_LOG.Debug("WM_ACTIVATE: MainGUI lost focus, closing...")
        MainGuiClose()
    }
    return 0
}

UpdateSendTo() {                 ; the lnk in SendTo must point to a exe
    lnkPath := StrReplace(A_StartMenu, "\Start Menu", "\SendTo\") "ALTRun.lnk"
    if (!g_CONFIG["EnableSendTo"]) {
        FileDelete(lnkPath)
        g_LOG.Debug("UpdateSendTo: Update SendTo shortcut=Disabled")
        return
    }

    if (A_IsCompiled)
        FileCreateShortcut(A_ScriptFullPath, lnkPath, ,"-SendTo", "Send command to ALTRun User Command list")
    else
        FileCreateShortcut(A_AhkPath, lnkPath, , A_ScriptFullPath " -SendTo", "Send command to ALTRun User Command list")

    g_LOG.Debug("UpdateSendTo: Update SendTo shortcut=OK")
    return
}

UpdateStartup() {
    lnkPath := A_Startup "\ALTRun.lnk"

    if (!g_CONFIG["AutoStartup"]) {
        FileDelete(lnkPath)
        g_LOG.Debug("UpdateStartup: Update Startup shortcut=Disabled")
        return
    }

    FileCreateShortcut(A_ScriptFullPath, lnkPath, A_ScriptDir, "-startup", "ALTRun - An effective launcher")

    g_LOG.Debug("UpdateStartup: Update Startup shortcut=OK")
    return
}

UpdateStartMenu() {
    lnkPath := A_Programs "\ALTRun.lnk"

    if (!g_CONFIG["InStartMenu"]) {
        FileDelete(lnkPath)
        g_LOG.Debug("UpdateStartMenu: Update StartMenu shortcut=Disabled")
        return
    }

    FileCreateShortcut(A_ScriptFullPath, lnkPath, A_ScriptDir, "-StartMenu", "ALTRun - An effective launcher")
    g_LOG.Debug("UpdateStartMenu: Update StartMenu shortcut=OK")
    return
}

Reindex(*) {                                                            ; Re-create Index section
    IniDelete(g_INI, g_SECTION["INDEX"])                     ; Clear old index section
    for dirIndex, dir in StrSplit(g_CONFIG["IndexDir"], ",")
    {
        searchPath := AbsPath(dir)
        searchPath := RegExReplace(searchPath, "\\+$")                  ; Remove trailing backslashes

        for extIndex, ext in StrSplit(g_CONFIG["IndexType"], ",") {
            Loop Files, searchPath "\" ext, "R"
            {
                ; Calculate path relative to searchPath and count subdir levels
                rel := SubStr(A_LoopFileFullPath, StrLen(searchPath) + 2) ; +2 to skip the backslash
                seps := (rel = "") ? 0 : StrLen(rel) - StrLen(StrReplace(rel, "\", "")) ; Count backslashes to determine depth

                if (seps >= g_CONFIG["IndexDepth"])                     ; If file is deeper than allowed depth, skip it.
                    continue

                if (g_CONFIG["IndexExclude"] != "" && RegExMatch(A_LoopFileFullPath, g_CONFIG["IndexExclude"]))
                    continue                                            ; Skip this file and move on to the next loop.

                IniWrite("1", g_INI, g_SECTION["INDEX"], "File | " A_LoopFileFullPath) ; Store file type for later use

                static ProgressGui := ""                                ; Static to persist GUI across loop iterations
                if (!ProgressGui) {                                     ; Create GUI only once
                    ProgressGui := Gui("-MinimizeBox +AlwaysOnTop", "Reindex")
                    ProgressGui.Add("Text", , "ReIndexing...")
                    ProgressGui.Add("Progress", "vMyProgress w200", A_Index)
                    ProgressGui.Add("Text", "vMyFileName w200", A_LoopFileName)
                    ProgressGui.Show()
                } else {                                                ; Update existing GUI
                    ProgressGui["MyProgress"].Value := A_Index
                    ProgressGui["MyFileName"].Text := A_LoopFileName
                    Sleep 10
                }
            }
        }
        if (ProgressGui) {                                              ; Destroy GUI after loop
            ProgressGui.Destroy()
            ProgressGui := ""
        }
    }

    g_LOG.Debug("Reindex: Indexing search database...")
    TrayTip("ReIndex database finish successfully.", g_TITLE, 8)
    LoadCommands()
}

Help(*) {
    Options(8)
}

Usage(*) {
    Options(7)
}

Update(*) {
    Run("https://github.com/zhugecaomao/ALTRun/releases")
}

Listary() {                                                             ; Listary 快速更换保存/打开对话框路径
    g_LOG.Debug("Listary: Initializing...")

    Loop Parse, g_CONFIG["FileMgrID"], ","                              ; File Manager Class, default is Windows Explorer & Total Commander
        GroupAdd("FileMgrID", A_LoopField)

    Loop Parse, g_CONFIG["DialogWin"], ","                              ; 需要QuickSwith的窗口, 包括打开/保存对话框等
        GroupAdd("DialogBox", A_LoopField)

    Loop Parse, g_CONFIG["ExcludeWin"], ","                             ; 排除特定窗口,避免被 Auto-QuickSwitch 影响
        GroupAdd("ExcludeWin", A_LoopField)

    if (g_CONFIG["AutoSwitchDir"]) {
        g_LOG.Debug("Listary: Auto-QuickSwitch enabled, starting monitoring thread...")
        Loop {
            ; WinWaitActive ahk_class TTOTAL_CMD
            ;     ThisHWND := WinGetID("A")
            ; WinWaitNotActive
            TcHwnd := WinWaitActive("ahk_class TTOTAL_CMD")
            WinWaitNotActive()

            ; 检测当前窗口是否符合打开保存对话框条件
            If(WinActive("ahk_group DialogBox") && !WinActive("ahk_group ExcludeWin")) {
                Title := WinGetTitle("A")
                ProcessName := WinGetProcessName("A")
                g_LOG.Debug("Listary: Dialog detected, active window ahk_title=" Title ", ahk_exe=" ProcessName)
                SyncTCPath()                                            ; NO Return, as will terimate loop (AutoSwitchDir)
            }
            Sleep 100  ; Reduce CPU usage
        }
    }
    HotIfWinActive("ahk_group DialogBox")                               ; 设置对话框路径定位热键,为了不影响其他程序热键,设置只对打开/保存对话框生效
    try {
        Hotkey(g_HOTKEY["ExplorerDir"], SyncExplorerPath)               ; Ctrl+E 把打开/保存对话框的路径定位到资源管理器当前浏览的目录
        Hotkey(g_HOTKEY["TotalCMDDir"], SyncTCPath)                     ; Ctrl+G 把打开/保存对话框的路径定位到TC当前浏览的目录
    } catch as e {
        g_LOG.Debug("Listary: Failed to set hotkey with error:" e.Message)
        OutputDebug("Listary: Failed to set hotkey with error:" e.Message)
    }
    HotIfWinActive                                                      ; Turn off context, make subsequent hotkeys global again
    return
}

SyncTCPath(*) { ; Sync dialog box to Total Commander path
    ClipSaved   := ClipboardAll()
    A_Clipboard := ""
    hwnd := WinGetID("ahk_class TTOTAL_CMD")                            ; Get the HWND of the active TC
    if (!hwnd) {
        g_LOG.Debug("SyncTCPath: No Total Commander window found")
        OutputDebug("SyncTCPath: No Total Commander window found")
        return
    }
    try {
        SendMessage(1075, 2029, 0, , "ahk_class TTOTAL_CMD")            ; TC: WM_USER + 75, TC_GETCURRENTPATH = 2029
    } catch as e {
        g_LOG.Debug("SyncTCPath: SendMessage failed, exception - " . e.Message)
        OutputDebug("SyncTCPath: SendMessage failed, exception - " . e.Message)
        A_Clipboard := ClipSaved
        return
    }
    ClipWait(0.1)                                                       ; Wait up to 0.1 seconds for the clipboard to contain data
    if (A_Clipboard = "") {
        A_Clipboard := ClipSaved
        g_LOG.Debug("SyncTCPath: A_Clipboard is empty")
        return
    }
    OutDir := RTrim(A_Clipboard, "\") . "\" ; Normalize path with single trailing \ 解决AutoCAD不识别路径问题
    A_Clipboard := ClipSaved
    ChangePath(OutDir)
}

SyncExplorerPath(*) {  ; Sync dialog box to Explorer path (Win7~11)
    hwnd := WinGetID("ahk_class CabinetWClass")                         ; Get the HWND of the active Explorer window
    if (!hwnd) {
        g_LOG.Debug("SyncExplorerPath: No Explorer window found")
        OutputDebug("SyncExplorerPath: No Explorer window found")
        return
    }
    try {
        for window in ComObject("Shell.Application").Windows
            if (window.HWND = hwnd) {
                Dir := window.Document.Folder.Self.Path
                ChangePath(Dir)
                return
            }
        g_LOG.Debug("SyncExplorerPath: No matching Explorer window")
        OutputDebug("SyncExplorerPath: No matching Explorer window")
    } catch as e {
        g_LOG.Debug("SyncExplorerPath: COM error - " e.Message)
        OutputDebug("SyncExplorerPath: COM error - " e.Message)
    }
}

ChangePath(Dir) {                                                       ; Set dialog box path to specified directory
    if (!Dir || !FileExist(Dir)) {
        g_LOG.Debug("ChangePath: Invalid directory :" Dir)
        return
    }
    ActiveClass := WinGetClass("A")
    if (ActiveClass = "Qt5QWindowIcon") {
        ; WPS dialog: Its Edit control has no valid id, try simulate keyboard input (unreliable)
        Send "{Dir}"
        SendInput "{Enter}"
        g_LOG.Debug("ChangePath: Set WPS path to " Dir " (keyboard fallback)")
    } else {
        ; For standard dialog (#32770)
        dialogControl := ControlGetHwnd("Edit1", "A")
        if (dialogControl) {
            currentText := ControlGetText("Edit1", "A")
            ControlClick("Edit1", "A")
            ControlSetText(Dir, "Edit1", "A")
            ControlSend("{Enter}", "Edit1", "A")
            g_LOG.Debug("ChangePath: Set path to " Dir " (Edit1)")
        } else {
            g_LOG.Debug("ChangePath: No editable control found")
        }
    }
}

UserCommand(*) {
    Run("Notepad.exe " . g_INI)
}

NewCommand(*) {                                                         ; From command "New Command" or GUI context menu "New Command"
    CmdMgr(g_SECTION["USERCMD"], , , g_RUNTIME["Arg"], 1, "")
}

EditCommand(*) {
    Global g_RUNTIME, g_SECTION, g_INI  ; 明确声明全局变量

    currentCmd := g_RUNTIME["CurrentCommand"]
    if !currentCmd
        return

    for index, section in StrSplit(g_SECTION["DFTCMD"] "," g_SECTION["USERCMD"] "," g_SECTION["INDEX"], ",") {
        rank := IniRead(g_INI, section, currentCmd, "KeyNotFound")

        ; If currentCmd not exist in this section, skips the rest of a loop and begin to check next section
        if (rank = "KeyNotFound" || rank = "ERROR")
            continue

        if IsInteger(rank) {
            parts := StrSplit(currentCmd, " | ")
            type := parts.Length >= 1 ? parts[1] : ""
            path := parts.Length >= 2 ? parts[2] : ""
            desc := parts.Length >= 3 ? parts[3] : ""

            g_Log.Debug("EditCommand: Editing command=" currentCmd)
            CmdMgr(section, type, path, desc, rank, currentCmd)
            break
        }
    }
}

DelCommand(*) {
    Global g_RUNTIME, g_SECTION, g_INI, g_LNG, g_TITLE

    currentCmd := g_RUNTIME["CurrentCommand"]
    if !currentCmd
        return

    for index, section in StrSplit(g_SECTION["DFTCMD"] "," g_SECTION["USERCMD"] "," g_SECTION["INDEX"], ",") {
        rank := IniRead(g_INI, section, currentCmd, "KeyNotFound")
        if (rank = "KeyNotFound" || rank = "ERROR")
            continue

        prompt := g_LNG[800] " [ " section " ] " g_LNG[801] "`n`n" currentCmd
        result := MsgBox(prompt, "Confirm want to delete?", 52)         ; 52 = Yes/No + Question icon

        if result = "YES" {
            try {
                IniDelete(g_INI, section, currentCmd)
                MsgBox(g_LNG[802] "`n`n" currentCmd, g_TITLE, 64)       ; 64 = Info icon
            } catch as e {
                MsgBox(g_LNG[803] "`n`n" currentCmd, g_TITLE, 48)       ; 48 = Error icon
            }
            break
        }
    }
    LoadCommands()
}


CmdMgr(Section := "UserCommand", Type := "File", Path := "", Desc := "", Rank := 1, OriginCmd := "") { ; 命令管理窗口
    Global CmdMgrGUI

    g_LOG.Debug("Starting Command Manager... Args=" Section "|" Type "|" Path "|" Desc "|" Rank)

    CmdMgrGUI := Gui(, g_LNG[700])
    CmdMgrGUI.SetFont("S9 Norm", "Microsoft Yahei")
    CmdMgrGUI.AddGroupBox("w600 h260", g_LNG[701])
    CmdMgrGUI.Add("Text", "x25 yp+30", g_LNG[702])
    CmdMgrGUI.AddDropDownList("x160 yp-5 w130 vType Choose" GetArrayIndex(Type, g_TYPELST), g_TYPELST)
    CmdMgrGUI.Add("Text", "x315 yp+5", g_LNG[705])
    CmdMgrGUI.Add("Edit", "x435 yp-5 w130 Disabled vSection", Section)
    CmdMgrGUI.Add("Text", "x25 yp+60", g_LNG[703])
    CmdMgrGUI.Add("Edit", "x160 yp-5 w405 -WantReturn vPath", Path).Focus()
    CmdMgrGUI.AddButton("x575 yp w30 hp", "...").OnEvent("Click", (*) => SelectCmdPath(CmdMgrGUI["Type"].Text))
    CmdMgrGUI.Add("Text", "x25 yp+80", g_LNG[704])
    CmdMgrGUI.AddEdit("x160 yp-5 w405 -WantReturn vDesc", Desc)
    CmdMgrGUI.AddText("x25 yp+60", g_LNG[706])
    CmdMgrGUI.AddEdit("x160 yp-5 w405 +Number vRank", Rank)
    CmdMgrGUI.AddButton("Default x420 w90", g_LNG[7]).OnEvent("Click", (*) => CmdMgrButtonOK(Section, CmdMgrGUI["Type"].Text, CmdMgrGUI["Path"].Text, CmdMgrGUI["Desc"].Text, CmdMgrGUI["Rank"].Text, OriginCmd))
    CmdMgrGUI.AddButton("x521 yp w90", g_LNG[8]).OnEvent("Click", CmdMgrGuiClose)
    CmdMgrGUI.OnEvent("Close", CmdMgrGUIClose)
    CmdMgrGUI.OnEvent("Escape", CmdMgrGUIClose)
    CmdMgrGUI.Show("Center")
}

SelectCmdPath(Type) {
    CmdMgrGUI.Opt("+OwnDialogs")                                        ; Make open dialog Modal

    if (Type = "Dir")
        Path := DirSelect(, 3, 'Please select directory')
    else
        Path := FileSelect(3, , , 'All Files (*.*)')

    if (Path != "")
        CmdMgrGUI["Path"].Value := Path
}

CmdMgrButtonOK(Section, Type, Path, Desc, Rank, OriginCmd) {
    CmdMgrGUI.Submit()
    Desc := Desc ? " | " Desc : Desc

    if (Path = "") {
        return MsgBox("Command Path is empty, please input correct command path!", "Command Manager", 64)
    } else {
        try {
            IniDelete(g_INI, Section, OriginCmd)                       ; Delete old command if path or desc changed
            IniWrite(Rank, g_INI, Section, Type " | " Path Desc)
        } catch as e {
            MsgBox("CmdMgr: Add command error occur, error info=" e.Message)
            return
        }
        MsgBox("The following command added / modified successfully!`n`n[ " Section " ]`n`n" Type " | " Path " " Desc " = " Rank, "Command Manager", 64)
    }
    LoadCommands()
}

CmdMgrGuiClose(*) {
    CmdMgrGUI.Destroy()
}

Plugins() {                                                             ; Plugins (Ctrl+D 自动添加日期)
    Loop Parse, g_HOTKEY["AutoDateBefExt"], ","
        GroupAdd("FileListMangr", A_LoopField)

    Loop Parse, g_HOTKEY["AutoDateAtEnd"], ","
        GroupAdd("TextBox", A_LoopField)

    HotIfWinActive("ahk_group FileListMangr")                           ; 针对所有设定好的程序 按Ctrl+D自动在文件(夹)名之后添加日期
    Hotkey(g_HOTKEY["AutoDateBEHKey"], RenameWithDate)


    HotIfWinActive("ahk_group TextBox")
    Hotkey(g_HOTKEY["AutoDateAEHKey"], LineEndAddDate)
    HotIfWinActive

    g_LOG.Debug("Plugins: Loaded AutoDate plugins...")
    return
}

RenameWithDate(*) {                                                     ; 针对所有设定好的程序 按Ctrl+D自动在文件(夹)名之后添加日期
    FocusedHwnd  := ControlGetFocus("A")                                ; 获取当前激活的窗口中的聚焦的控件名称
    FocusedClassNN := ControlGetClassNN(FocusedHwnd)

    if (InStr(FocusedClassNN, "Edit") or InStr(FocusedClassNN, "Scintilla")) ; 如果当前激活的控件为Edit类或者Scintilla1(Notepad2),则Ctrl+D功能生效
        NameAddDate("FileListMangr", FocusedClassNN)
    Else
        SendInput "^D"                                                  ; 如果不是,则发送原始的Ctrl+D

    OutputDebug("RenameWithDate: Current control=" FocusedClassNN)
    Return
}

LineEndAddDate(*) {                                                      ; 针对TC File Comment对话框　按Ctrl+D自动在备注文字之后添加日期
    CurrentDate := FormatTime(, "dd.MM.yyyy")
    SendInput "{End}"
    Sleep 10
    SendInput "{Blind}{Text} - " CurrentDate
    g_LOG.Debug("LineEndAddDate: Add date at end= - " CurrentDate)
}

NameAddDate(WinName, CurrCtrl) {                                        ; 在文件（夹）名编辑框中添加日期,CurrCtrl为当前控件(名称编辑框Edit)
    EditCtrlText := ControlGetText(CurrCtrl, "A")
    SplitPath(EditCtrlText, &fileName, &fileDir, &fileExt, &nameNoExt)
    CurrentDate := FormatTime(, "dd.MM.yyyy")

    if (StrLen(fileExt) < 5 && !RegExMatch(fileExt,"^\d+$")) {          ; 如果有真实文件后缀名,才加日期在后缀名之前
        if RegExMatch(nameNoExt, " - \d{2}\.\d{2}\.\d{4}$") {
            baseName := RegExReplace(nameNoExt, " - \d{2}\.\d{2}\.\d{4}$", "")
        }
        else if RegExMatch(nameNoExt, "-\d{2}\.\d{2}\.\d{4}$") {
            baseName := RegExReplace(nameNoExt, "-\d{2}\.\d{2}\.\d{4}$", "")
        }
        else {
            baseName := nameNoExt
        }
        NameWithDate := baseName " - " CurrentDate "." fileExt
    }
    else if (RegExMatch(fileName, " - \d{2}\.\d{2}\.\d{4}$")) {         ; 如果无后缀, 文件(夹)名最后有日期,则更新为当前日期
        NameWithDate := RegExReplace(fileName, " - \d{2}\.\d{2}\.\d{4}$", " - " CurrentDate)
    }
    else if (RegExMatch(nameNoExt, "-\d{2}\.\d{2}\.\d{4}$")) {
        NameWithDate := RegExReplace(fileName, "-\d{2}\.\d{2}\.\d{4}$", " - " CurrentDate)
    }
    else {
        NameWithDate := EditCtrlText " - " CurrentDate
    }
    ControlClick(CurrCtrl, "A")
    ControlSetText(NameWithDate, CurrCtrl, "A")
    SendInput "{Blind}{End}"
    g_LOG.Debug("NameAddDate: Add date to name= " NameWithDate)
}

GetArrayIndex(searchValue, Array){
    for index, element in Array
    {
        if (element = searchValue)
            return index
    }
}

PTTools() {
    if not WinExist("PT Tools"){
        Try Run(A_ScriptDir "\PTTools.ahk")
    }
    else{
        WinActivate("PT Tools")
    }
}

FormatThousand(Number) {                                                ; Function to add thousand separator
    Return RegExReplace(Number, "\G\d+?(?=(\d{3})+(?:\D|$))", "$0" ",")
}

StruCalcResult(evalResult) {
    result := []

    formatVal := FormatThousand(evalResult)
    result.Push("Eval | " formatVal)

    if !g_CONFIG["StruCalc"]
        return result

    result.Push("")  ; 空行分隔

    ; 主筋计算
    rebarNum := Ceil((evalResult - 80) / 300 + 1)
    spacing  := Max(Round((evalResult - 80) / (rebarNum - 0.999)), 0)   ; Use 0.999 to avoid division by zero error
    result.Push("Eval | With beam width = " formatVal " mm")
    result.Push(" | Main bar number = " rebarNum " (" spacing " C/C)")

    result.Push("")  ; 空行分隔

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
    Local FuncList := ["Unset", "Active", "ToggleWindow", "Google", "Bing"
        , "Everything", "TabFunc", "PrevCommand", "NextCommand", "CopyCommand"
        , "ClearInput", "RunCurrentCommand", "RankUp", "RankDown", "Reindex"
        , "Help", "Usage", "Update", "UserCommand", "NewCommand", "EditCommand"
        , "DelCommand", "CmdMgr", "Options", "TurnMonitorOff", "EmptyRecycle"
        , "MuteVolume", "ReStart", "Exit"]

    t := A_TickCount
    ActTab := IsNumber(ActTab) ? ActTab : 1                             ; Convert ActTab to number, default is 1 (for case like [Option`tF2])
    OptGUI := Gui("+Owner", g_LNG[2])
    OptGUI.Hide()
    OptGUI.SetFont(StrSplit(g_GUI["OptGUIFont"], ",")[2], StrSplit(g_GUI["OptGUIFont"], ",")[1])
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
    OptGUI.AddEdit("x183 yp w240 r1 -E0x200 +ReadOnly vMainGUIFont", g_GUI["MainGUIFont"]).SetFont(StrSplit(g_GUI["MainGUIFont"], ",")[2], StrSplit(g_GUI["MainGUIFont"], ",")[1])
    OptGUI.AddButton("x433 yp-5 w80 vSelectMainGUIFont", g_LNG[182]).OnEvent("Click", SelectMainGUIFont)
    OptGUI.AddText("x33 yp+45", g_LNG[174])
    OptGUI.AddEdit("x183 yp w240 r1 -E0x200 +ReadOnly vOptGUIFont", g_GUI["OptGUIFont"])
    OptGUI.AddButton("x433 yp-5 w80 vSelectOptGUIFont", g_LNG[182]).OnEvent("Click", SelectOptGUIFont)
    OptGUI.AddText("x33 yp+45", g_LNG[175])
    OptGUI.AddEdit("x183 yp w240 r1 -E0x200 +ReadOnly vMainSBFont", g_GUI["MainSBFont"]).SetFont(StrSplit(g_GUI["MainSBFont"], ",")[2], StrSplit(g_GUI["MainSBFont"], ",")[1])
    OptGUI.AddButton("x433 yp-5 w80 vSelectMainSBFont", g_LNG[182]).OnEvent("Click", SelectMainSBFont)

    OptGUI.AddText("x33 yp+45", g_LNG[178])
    OptGUI.AddEdit("x183 yp w240 r1 -E0x200 +ReadOnly vCtrlColor c" g_GUI["CtrlColor"], g_GUI["CtrlColor"])
    OptGUI.AddButton("x433 yp-5 w80 vSelectCtrlColor", g_LNG[183]).OnEvent("Click", SelectCtrlColor)
    OptGUI.AddText("x33 yp+45", g_LNG[179])
    OptGUI.AddEdit("x183 yp w240 r1 -E0x200 +ReadOnly vWinColor c" g_GUI["WinColor"], g_GUI["WinColor"])
    OptGUI.AddButton("x433 yp-5 w80 vSelectWinColor", g_LNG[183]).OnEvent("Click", SelectWinColor)

    OptGUI.AddText("x33 yp+45", g_LNG[180])
    OptGUI.AddComboBox("x183 yp-5 w330 vBackground Choose1", [g_GUI["Background"], "ALTRun.jpg", "None", "C:\Path\BG.jpg"])
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
    try {
        Hotkey(g_HOTKEY["GlobalHotkey1"], "Off")
        Hotkey(g_HOTKEY["GlobalHotkey2"], "Off")
    } catch as e {
        g_LOG.Debug("Options: Set GlobalHotkey Off=" . e.Message)
    }

    OptTab.UseTab(4) ; INDEX Tab
    OptGUI.AddGroupBox("w500 h190", g_LNG[160])
    OptGUI.AddText("x33 yp+25", g_LNG[161])
    OptGUI.AddComboBox("x183 yp-5 w330 vIndexDir Choose1", [g_CONFIG["IndexDir"], "A_ProgramsCommon,A_StartMenu"])
    OptGUI.AddText("x33 yp+45", g_LNG[162])
    OptGUI.AddComboBox("x183 yp-5 w330 vIndexType Choose1", [g_CONFIG["IndexType"], "*.lnk,*.exe"])
    OptGUI.AddText("x33 yp+45", g_LNG[164])
    OptGUI.AddDropDownList("x183 yp-5 w330 vIndexDepth Choose" g_CONFIG["IndexDepth"], [1,2,3,4,5,6,7,8,9])
    OptGUI.AddText("x33 yp+45", g_LNG[163])
    OptGUI.AddComboBox("x183 yp-5 w330 vIndexExclude Choose1", [g_CONFIG["IndexExclude"], "Uninstall *"])

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
    OptGUI.AddGroupBox("w500 h110", g_LNG[221])
    OptGUI.AddText("x33 yp+30", g_LNG[222])
    OptGUI.AddComboBox("x183 yp-5 w330 vAutoDateAtEnd Choose1", [g_HOTKEY["AutoDateAtEnd"], "ahk_class TCmtEditForm,ahk_exe Notepad4.exe"])
    OptGUI.AddText("x33 yp+45", g_LNG[223])
    OptGUI.AddHotkey("x183 yp-5 w80 vAutoDateAEHKey", g_HOTKEY["AutoDateAEHKey"])
    OptGUI.AddText("x300 yp+5", g_LNG[224])
    OptGUI.AddDDL("x395 yp-5 w120 vAutoDateAEFormat Choose1", ["- dd.MM.yyyy"])

    OptGUI.AddGroupBox("x24 y+30 w500 h110", g_LNG[225])
    OptGUI.AddText("x33 yp+30", g_LNG[222])
    OptGUI.AddComboBox("x183 yp-5 w330 vAutoDateBefExt Choose1", [g_HOTKEY["AutoDateBefExt"], "ahk_class CabinetWClass,ahk_class Progman,ahk_class WorkerW,ahk_class #32770"])
    OptGUI.AddText("x33 yp+45", g_LNG[223])
    OptGUI.AddHotkey("x183 yp-5 w80 vAutoDateBEHKey", g_HOTKEY["AutoDateBEHKey"])
    OptGUI.AddText("x300 yp+5", g_LNG[224])
    OptGUI.AddDDL("x395 yp-5 w120 vAutoDateBEFormat Choose1", ["- dd.MM.yyyy"])

    OptGUI.AddGroupBox("x24 y+30 w500 h110", g_LNG[229])
    OptGUI.AddText("x33 yp+30", g_LNG[230])
    OptGUI.AddComboBox("x183 yp-5 w330 vCondTitle Choose1", [g_HOTKEY["CondTitle"]])
    OptGUI.AddText("x33 yp+45", g_LNG[231])
    OptGUI.AddComboBox("x183 yp-5 w80 vCondHotkey Choose1", [g_HOTKEY["CondHotkey"]])
    OptGUI.AddText("x300 yp+5", g_LNG[232])
    OptGUI.AddDDL("x395 yp-5 w120 vCondAction Choose" GetArrayIndex(g_HOTKEY["CondAction"], FuncList), FuncList)

    OptTab.UseTab(7) ; USAGE Tab
    OptGUI.AddGroupBox("x66 y80 w445 h300", )
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
    OptGUI.AddText("x96 yp+5", g_TITLE)
    OptGUI.AddLink("xp yp+45 w400", g_LNG[601])

    OptTab.UseTab()  ; 后续添加的控件将不属于前面的选项卡控件
    OptGUI.AddButton("Default x278 w80", g_LNG[7]).OnEvent("Click", OPTButtonOK)
    OptGUI.AddButton("x368 yp w80", g_LNG[8]).OnEvent("Click", OPTGuiClose)
    OptGUI.AddButton("x458 yp w80", g_LNG[9]).OnEvent("Click", (*) => Run("https://github.com/zhugecaomao/ALTRun/wiki"))
    OptGUI.OnEvent("Close", OPTGuiClose)
    OptGUI.OnEvent("Escape", OPTGuiClose)

    g_LOG.Debug("Options: Loading options window... ActTab=" ActTab ", elapsed time=" A_TickCount - t "ms")
    OutputDebug("Options: Loading options window... ActTab=" ActTab ", elapsed time=" A_TickCount - t "ms")
    OptGUI.Show("Center")
    return
}

ResetHotkey(*) {
    OptGUI["g_GlobalHotkey1"].Value := "!Space"
    OptGUI["g_GlobalHotkey2"].Value := "!r"
    return
}

SelectMainGUIFont(*) {
    SelectFont("MainGUIFont")
}

SelectOptGUIFont(*) {
    SelectFont("OptGUIFont")
}

SelectMainSBFont(*) {
    SelectFont("MainSBFont")
}

SelectFont(TargetVar := "MainGUIFont") {
    ; 从 g_GUI 读取当前字体名作为初始选择
    ; Set the fontObj (optional) - only set the ones you want to pre-select
	; fontObj := Map("name","Terminal","size",14,"color",0xFF0000,"strike",1,"underline",1,"italic",1,"bold",1)
    initFont := StrSplit(g_GUI[TargetVar], ",")[1]
    fontObj  := Map("name", initFont)
    fontObj  := FontSelect(fontObj, OptGUI.hwnd)
    If (!fontObj)
        return

    OptGUI[TargetVar].Text := fontObj["name"] ", " fontObj["str"]       ; 更新控件字体并设置显示文本
    OptGUI[TargetVar].SetFont(fontObj["str"], fontObj["name"])
    OutputDebug("SelectFont: OptGUI[" TargetVar " font set to Option=" fontObj["str"] ", font name=" fontObj["name"])
}

SelectCtrlColor(*) {
    color := ColorSelect(, OptGUI.hwnd, , "full")                       ; hwnd and custColorObj are optional
    OptGUI["CtrlColor"].Value := color                                  ; 更新选项窗口控件并设置控件颜色
    OptGUI["CtrlColor"].Opt("c" . color)
}

SelectWinColor(*) {
    color := ColorSelect(, OptGUI.hwnd, , "full")
    OptGUI["WinColor"].Value := color
    OptGUI["WinColor"].Opt("c" . color)
}

OPTButtonOK(*) {
    SaveConfig()
    Reload
}

OPTGuiClose(*) {
    try {
        Hotkey(g_HOTKEY["GlobalHotkey1"], "On")
        Hotkey(g_HOTKEY["GlobalHotkey2"], "On")

        OutputDebug("OPTGuiClose: Set back on GlobalHotkey...OK")
        g_LOG.Debug("OPTGuiClose: Set back on GlobalHotkey...OK")
    } catch as e {
        OutputDebug("OPTGuiClose: Set back on GlobalHotkey error=" . e.Message)
        g_LOG.Debug("OPTGuiClose: Set back on GlobalHotkey error=" . e.Message)
    }
    OptGUI.Destroy()

    OutputDebug("OPTGuiClose: OptGUI.Destroy()...OK")
    g_LOG.Debug("OPTGuiClose: OptGUI.Destroy()...OK")
    return
}

LoadConfig(Arg) {
    g_LOG.Debug("LoadConfig: Loading configuration...Arg=" . Arg)

    if (Arg = "config" || Arg = "initialize" || Arg = "all") {
        for key, value in g_CONFIG {  ; Read [Config] to Map
            g_CONFIG[key] := IniRead(g_INI, g_SECTION["CONFIG"], key, value)
        }

        for key, value in g_HOTKEY {  ; Read [Hotkey] section
            g_HOTKEY[key] := IniRead(g_INI, g_SECTION["HOTKEY"], key, value)
        }

        for key, value in g_GUI {  ; Read [GUI] section
            g_GUI[key] := IniRead(g_INI, g_SECTION["GUI"], key, value)
        }

        g_RUNTIME["RegEx"] := g_CONFIG["MatchBeginning"] ? "imS)^" : "imS)"

        OffsetDate := DateAdd(A_Now, -30, "Days")

        USAGESEC := ""
        Try USAGESEC := IniRead(g_INI, g_SECTION["USAGE"])
        if (USAGESEC != "") {
            for line in StrSplit(USAGESEC, "`n") {
                if (!line)  ; Skip empty lines
                    continue
                split := StrSplit(line, "=")
                Date  := split[1]
                Count := split[2]

                if (Date <= SubStr(OffsetDate, 1, 8)) {  ; Clean up usage record before 30 days
                    IniDelete(g_INI, g_SECTION["USAGE"], Date)
                    continue
                }

                g_USAGE[Date] := Count
                g_RUNTIME["Max"] := Max(g_RUNTIME["Max"], Count)
            }
        }

        Loop 30 {
            OffsetDate := DateAdd(OffsetDate, 1, "Days")
            Date := SubStr(OffsetDate, 1, 8)
            g_USAGE[Date] := g_USAGE.Has(Date) ? g_USAGE[Date] : 0
        }
    }

    if (Arg = "commands" || Arg = "initialize" || Arg = "all") {  ; Built-in command initialize
        DFTCMDSEC := ""
        Try DFTCMDSEC := IniRead(g_INI, g_SECTION["DFTCMD"])
        if (DFTCMDSEC = "") {
            IniWrite("
            (
            ; This section is Built-In commands with high priority
            ; App will auto generate this section while it is empty
            ; Please make sure App is not running before modifying.
            ;
            Func | Help | ALTRun Help & About (F1)=99
            Func | Options | ALTRun Options Preference Settings (F2)=99
            Func | Reload | Reload ALTRun=99
            Func | EditCommand | Edit current command (F3)=99
            Func | UserCommand | ALTRun User-defined command (F4)=99
            Func | NewCommand | New Command=99
            Func | OpenContainer | Locate cmd's dir with File Manager=99
            Func | Usage | ALTRun Usage Status=99
            Func | Reindex | Reindex search database=99
            Func | Everything | Search by Everything=99
            Func | PTTools | PT Tools (AHK)=99
            Func | AhkRun | Run Command use AutoHotkey Run=99
            Func | Google | Search Clipboard or Input by Google=99
            Func | Bing | Search Clipboard or Input by Bing=99
            Func | EmptyRecycle | Empty Recycle Bin=99
            Func | TurnMonitorOff | Turn off Monitor, Close Monitor=99
            Func | MuteVolume | Mute Volume=99
            File | %Temp%\ALTRun.log | ALTRun Log File=99
            Dir | A_ScriptDir | ALTRun Program Dir=99
            Dir | A_Startup | Current User Startup Dir=99
            Dir | A_StartupCommon | All User Startup Dir=99
            Dir | A_ProgramsCommon | Windows Search.Index.Cortana Dir=99
            Dir | A_Desktop=99
            Dir | %AppData%\Microsoft\Windows\SendTo | Windows SendTo Dir=99
            Dir | %OneDrive% | OneDrive=99
            Cmd | explorer.exe | Windows File Explorer=99
            Cmd | cmd.exe | DOS / CMD=99
            Cmd | cmd.exe /k ipconfig | Check IP Address=99
            Cmd | Shell:AppsFolder | AppsFolder Applications=66
            Cmd | ::{645FF040-5081-101B-9F08-00AA002F954E} | Recycle Bin=66
            Cmd | ::{20D04FE0-3AEA-1069-A2D8-08002B30309D} | This PC=66
            Cmd | Notepad.exe | Notepad=66
            Cmd | WF.msc | Windows Defender Firewall with Advanced Security=66
            Cmd | TaskSchd.msc | Task Scheduler=66
            Cmd | DevMgmt.msc | Device Manager=66
            Cmd | EventVwr.msc | Event Viewer=66
            Cmd | CompMgmt.msc | Computer Manager=66
            Cmd | TaskMgr.exe | Task Manager=66
            Cmd | Calc.exe | Calculator=66
            Cmd | MsPaint.exe | Paint=66
            Cmd | Regedit.exe | Registry Editor=66
            Cmd | CleanMgr.exe | Disk Space Clean-up Manager=66
            Cmd | GpEdit.msc | Group Policy=66
            Cmd | DiskMgmt.msc | Disk Management=66
            Cmd | DxDiag.exe | Directx Diagnostic Tool=66
            Cmd | LusrMgr.msc | Local Users and Groups=66
            Cmd | MsConfig.exe | System Configuration=66
            Cmd | PerfMon.exe /Res | Resources Monitor=66
            Cmd | PerfMon.exe | Performance Monitor=66
            Cmd | WinVer.exe | About Windows=66
            Cmd | Services.msc | Services=66
            Cmd | NetPlWiz | User Accounts=66
            Cmd | Control | Control Panel=66
            Cmd | Control Intl.cpl | Region and Language Options=66
            Cmd | Control Firewall.cpl | Windows Defender Firewall=66
            Cmd | Control Access.cpl | Ease of Access Centre=66
            Cmd | Control AppWiz.cpl | Programs and Features=66
            Cmd | Control Sticpl.cpl | Scanners and Cameras=66
            Cmd | Control Sysdm.cpl | System Properties=66
            Cmd | Control Mouse | Mouse Properties=66
            Cmd | Control Desk.cpl | Display=66
            Cmd | Control Mmsys.cpl | Sound=66
            Cmd | Control Ncpa.cpl | Network Connections=66
            Cmd | Control Powercfg.cpl | Power Options=66
            Cmd | Control AdminTools | Windows Tools=66
            Cmd | Control Desktop | Personalisation=66
            Cmd | Control Inetcpl.cpl,,4 | Internet Properties=66
            Cmd | Control Printers | Devices and Printers=66
            Cmd | Control UserPasswords | User Accounts=66
            )", g_INI, g_SECTION["DFTCMD"])
            DFTCMDSEC := IniRead(g_INI, g_SECTION["DFTCMD"])
        }

        USERCMDSEC := ""
        Try USERCMDSEC := IniRead(g_INI, g_SECTION["USERCMD"])
        if (USERCMDSEC = "") {
            IniWrite("
            (
            ; This section is User-Defined commands, modify as desired
            ; Format: Command Type | Command | Description=Rank
            ; Command type: File, Dir, CMD, URL, some sample below
            ; Please make sure App is not running before modifying.
            ;
            Dir | A_AppData "\Microsoft\Windows\SendTo" | Windows SendTo Dir=9
            Dir | %OneDrive% | OneDrive=9
            Dir | A_ScriptDir | ALTRun Program Dir=9
            Cmd | cmd.exe /k ipconfig | Check IP Address=9
            Cmd | explorer /select,C:\Program Files | Open C: and locate to Program Files=9
            Cmd | Control TimeDate.cpl | Date and Time=9
            Cmd | ::{20D04FE0-3AEA-1069-A2D8-08002B30309D} | This PC=9
            URL | www.google.com | Google=9
            File | C:\Apps\TotalCMD64\TOTALCMD64.exe=9
            )", g_INI, g_SECTION["USERCMD"])
            USERCMDSEC := IniRead(g_INI, g_SECTION["USERCMD"])
        }

        INDEXSEC := ""
        Try INDEXSEC := IniRead(g_INI, g_SECTION["INDEX"])
        if (INDEXSEC = "") {
            msgText := g_CONFIG["Chinese"] ? "索引数据库为空, 请点击`n`n'确定'重新建立索引`n`n'取消'退出程序`n`n(请确保程序目录有写入权限)" 
                                         : "Index database is empty, please click`n`n'OK' to rebuild the index`n`n'Cancel' to exit the program`n`n(Please ensure the program directory is writable)"
            if (MsgBox(msgText, g_TITLE, 4161) = "Cancel") {
                ExitApp()
            }
            Reindex()
        }

        return DFTCMDSEC . "`n" . USERCMDSEC . "`n" . INDEXSEC
    }
    return
}

SaveConfig() {
    Global OptListView

    OptGUI.Submit()
    For key, description in g_CONFIG_P1 {
        g_CONFIG[key] := (A_Index = OptListView.GetNext(A_Index-1, "C")) ? 1 : 0 ; for Options - General page - Check Listview
        IniWrite(g_CONFIG[key], g_INI, g_SECTION["CONFIG"], key)
    }

    For key, value in g_CONFIG_P2 {
        if (OptGUI[key].Type = "CheckBox") {
            g_CONFIG[key] := OptGUI[key].Value
        }
        else {
            g_CONFIG[key] := OptGUI[key].Text
        }
        IniWrite(g_CONFIG[key], g_INI, g_SECTION["CONFIG"], key)
    }

    For key, value in g_GUI {
        if (OptGUI[key].Type = "Slider") {
            g_GUI[key] := OptGUI[key].Value
        }
        else {
            g_GUI[key] := OptGUI[key].Text
        }
        IniWrite(g_GUI[key], g_INI, g_SECTION["GUI"], key)
    }

    For key, value in g_HOTKEY {
        if (OptGUI[key].Type = "Hotkey") {
            g_HOTKEY[key] := OptGUI[key].Value
        }
        else {
            g_HOTKEY[key] := OptGUI[key].Text
        }
        IniWrite(g_HOTKEY[key], g_INI, g_SECTION["HOTKEY"], key)
    }

    g_LOG.Debug("SaveConfig: Save config...OK")
    return
}

;;==================== Built-in Functions =========================
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
    local Result := MsgBox("Do you really want to empty the Recycle Bin?", g_TITLE, "YesNo")
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
        Run(g_CONFIG["Everything"] . ' -s `"' g_RUNTIME["Arg"] '`"',,)
    } catch as e {
        MsgBox("Everything software not found.`n`nPlease check ALTRun setting and Everything program file.`n`nError message=" . e.Message)
    }
    return
}

;;==================== Language Setting =========================

SetLanguage() {
    ENG     := Map()
    CHN     := Map()

    ENG[1]  := "English"                                                ; 1~9 Reserved
    ENG[2]  := "Options"
    ENG[7]  := "OK"
    ENG[8]  := "Cancel"
    ENG[9]  := "Help"
    ENG[10] := ["No.", "Type", "Command", "Description"]                ; 10~49 Main GUI
    ENG[11] := "Run"
    ENG[12] := "Options"
    ENG[13] := "Type anything here to search..."
    ENG[50] := ["Tip | F1 | Help & About", "Tip | F2 | Options and settings", "Tip | F3 | Edit current command", "Tip | F4 | User-defined commands", "Tip | ALT+SPACE / ALT+R | Activate ALTRun", "Tip | ALT+SPACE / ESC / LOSE FOCUS | Deactivate ALTRun", "Tip | ENTER / ALT+NO. | Run selected command", "Tip | ARROW UP or DOWN | Select previous / next command", "Tip | CTRL+D | Locate cmd's dir with File Manager"]
    ENG[51] := "Tips: "
    ENG[52] := "It's better to activate ALTRun by hotkey (ALT + Space)" ; 50~99 Tips
    ENG[53] := "Smart Rank - Auto adjusts command priority (rank) based on frequency of use."
    ENG[54] := "Arrow Up / Down = Move to previous / next command"
    ENG[55] := "Esc = Clear input / close window"
    ENG[56] := "Enter = Run current command"
    ENG[57] := "Alt + No. = Run specific command"
    ENG[58] := "Start with + = New Command"
    ENG[59] := "F3 = Edit current command"
    ENG[60] := "F2 = Options setting"
    ENG[61] := "Ctrl+I = Reindex file search database"
    ENG[62] := "F1 = ALTRun Help & About"
    ENG[63] := "ALT + Space = Show / Hide Window"
    ENG[64] := "Ctrl+Q = ReStart ALTRun"
    ENG[65] := "Ctrl + No. = Select specific command"
    ENG[66] := "Alt + F4 = Exit"
    ENG[67] := "Ctrl+D = Open current command's dir with File Manager"
    ENG[68] := "F4 = Edit user-defined commands (.ini) directly"
    ENG[69] := "Start with space = Search file by Everything"
    ENG[70] := "Ctrl+'+' = Increase rank of current command"
    ENG[71] := "Ctrl+'-' = Decrease rank of current command"
    ENG[100] := ["General", "GUI", "Hotkey", "Index", "Listary", "Plugins", "Usage", "About"] ; 100~149 Options window (General - Check Listview)
    ENG[101] := "Launch on Windows startup"
    ENG[102] := "Enable SendTo - Create commands conveniently using Windows SendTo"
    ENG[103] := "Enable ALTRun shortcut in the Windows Start menu"
    ENG[104] := "Show tray icon in the system taskbar"
    ENG[105] := "Close window on losing focus"
    ENG[106] := "Always stay on top"
    ENG[107] := "Show window caption"
    ENG[108] := "Use Windows XP Theme instead of Classic Theme"
    ENG[109] := "Press [ESC] to clear input, press again to close window (Untick: Close directly)"
    ENG[110] := "Keep last input and matching result on close"
    ENG[111] := "Show command icon in the result list"
    ENG[112] := "SendToGetLnk - Retrieve .lnk target on SendTo"
    ENG[113] := "Save commands execution history"
    ENG[114] := "Save application log"
    ENG[115] := "Match full path on search"
    ENG[116] := "Show Grid - Provides boundary lines between list's rows and columns"
    ENG[117] := "Show Header - Show list's header (top row contains column titles)"
    ENG[118] := "Show Serial Number in command list"
    ENG[119] := "Show border line around the command list"
    ENG[120] := "Smart Rank - Auto adjust command priority (rank) based on use frequency"
    ENG[121] := "Smart Match - Fuzzy and Smart matching and filtering result"
    ENG[122] := "Match beginning of the string (Untick: Match from any position)"
    ENG[123] := "Show hints/tips in the bottom status bar"
    ENG[124] := "Show command executed RunCount in the status bar"
    ENG[125] := "Show status bar at the bottom of the window"
    ENG[126] := "Show [Run] button on main window"
    ENG[127] := "Show [Options] button on main window"
    ENG[128] := "Double Buffer - Paints via double-buffering, reduces flicker (WinXP+)"
    ENG[129] := "Enable express structure calculation"
    ENG[130] := "Shorten Path - Show file/folder/app name only instead of full path in result"
    ENG[131] := "Set language to Chinese Simplified (简体中文)"
    ENG[132] := "Match Chinese Pinyin first characters"
    ENG[150] := "File Manager"                                          ; 150~159 Options window (Other than Check Listview)
    ENG[151] := "Everything"
    ENG[152] := "History length"
    ENG[160] := "Index"                                                 ; 160~169 Index
    ENG[161] := "Index location"
    ENG[162] := "Index file type"
    ENG[163] := "Index exclude"
    ENG[164] := "Index depth"
    ENG[170] := "GUI"                                                   ; 170~189 GUI
    ENG[171] := "Search result number"
    ENG[172] := "Width of each column"
    ENG[173] := "Font (Main GUI)"
    ENG[174] := "Font (Options)"
    ENG[175] := "Font (Status Bar)"
    ENG[176] := "Window size (W x H)"
    ENG[177] := "Cmd list size (W x H)"
    ENG[178] := "Control color"
    ENG[179] := "Background color"
    ENG[180] := "Background picture"
    ENG[181] := "Transparency"
    ENG[182] := "Select font"
    ENG[183] := "Select color"
    ENG[190] := "Hotkey"                                                ; 190~209 Hotkey
    ENG[191] := "Activate Hotkey (Global)"
    ENG[192] := "Primary Hotkey"
    ENG[193] := "Secondary Hotkey"
    ENG[194] := "Two hotkeys can be set simultaneously"
    ENG[195] := "Reset hotkey"
    ENG[200] := "Actions and Hotkeys (Non-Global)"
    ENG[201] := "Hotkey"
    ENG[202] := "Trigger action"
    ENG[203] := "Hotkey"
    ENG[204] := "Hotkey 2"
    ENG[206] := "Hotkey 3"
    ENG[210] := "Listary"                                               ; 210~219 Listary
    ENG[211] := "Dir Quick-Switch"
    ENG[212] := "File Manager ID"
    ENG[213] := "Open/Save Dialog ID"
    ENG[214] := "Exclude Windows ID"
    ENG[215] := "Hotkey"
    ENG[216] := "QuickSwitch to TC's dir"
    ENG[217] := "QuickSwitch to Explorer"
    ENG[218] := "Auto switch dir on open/save dialog"
    ENG[220] := "Plugins"                                               ; 220~299 Plugins
    ENG[221] := "Auto-date at end of text"
    ENG[222] := "Apply to window id"
    ENG[223] := "Hotkey"
    ENG[224] := "Date format"
    ENG[225] := "Auto-date before file extension"
    ENG[229] := "Conditional action"
    ENG[230] := "If window id contains"
    ENG[231] := "Hotkey (Editable)"
    ENG[232] := "Trigger action"
    ENG[300] := "Show"                                                  ; 300+ TrayMenu
    ENG[301] := "Options`tF2"
    ENG[302] := "ReIndex`tCtrl+I"
    ENG[303] := "Usage"
    ENG[304] := "About`tF1"
    ENG[305] := "Script Info"
    ENG[307] := "ReStart`tCtrl+Q"
    ENG[308] := "Exit`tAlt+F4"
    ENG[309] := "Update"
    ENG[400] := "Run`tEnter"                                            ; 400+ LV_ContextMenu (Right-click)
    ENG[401] := "Locate`tCtrl+D"
    ENG[402] := "Copy`tCtrl+C"
    ENG[403] := "New`tCtrl+N"
    ENG[404] := "Edit`tF3"
    ENG[405] := "Delete`tDelete"
    ENG[406] := "User Command`tF4"
    ENG[407] := "Copy statusbar text"
    ENG[408] := "Show usage status"
    ENG[500] := "30 days ago"                                           ; 500+ Usage Status
    ENG[501] := "Now"
    ENG[502] := "Total number of times the command was executed"
    ENG[503] := "Number of times the program was activated today"
    ENG[600] := "About"
    ENG[601] := "An effective launcher for Windows by ZhugeCaomao, an <a href=`"https://www.autohotkey.com/docs/v1/`">AutoHotkey</a> open-source project. "
        . "It provides a streamlined and efficient way to find anything on your system and launch any application in your way."
        . "`n`nSetting file:`n" g_INI "`n`nProgram file:`n" A_ScriptFullPath
        . "`n`nCheck for Updates"
        . "`n<a href=`"https://github.com/zhugecaomao/ALTRun/releases`">https://github.com/zhugecaomao/ALTRun/releases</a>"
        . "`n`nSource code at GitHub"
        . "`n<a href=`"https://github.com/zhugecaomao/ALTRun`">https://github.com/zhugecaomao/ALTRun</a>"
        . "`n`nSee Help and Wiki page for more details"
        . "`n<a href=`"https://github.com/zhugecaomao/ALTRun/wiki`">https://github.com/zhugecaomao/ALTRun/wiki</a>"
    ENG[700] := "Commander Manager"
    ENG[701] := "Command"
    ENG[702] := "Command type"
    ENG[703] := "Command line"
    ENG[704] := "ShortCut/Description"
    ENG[705] := "Command Section"
    ENG[706] := "Command Rank"
    ENG[800] := "Do you really want to delete the following command from section"
    ENG[801] := "?"
    ENG[802] := "Command has been deleted successfully!"
    ENG[803] := "Error occur when delete the command!"

    CHN[1]  := "简体中文"                                               ; 1~9 Reserved
    CHN[2]  := "配置"
    CHN[7]  := "确定"
    CHN[8]  := "取消"
    CHN[9]  := "帮助"
    CHN[10] := ["序号", "类型", "命令", "描述"]                          ; 10~49 Main GUI
    CHN[11] := "运行"
    CHN[12] := "配置"
    CHN[13] := "在此输入搜索内容..."
    CHN[50] := ["提示 | F1 | 帮助&关于", "提示 | F2 | 配置选项", "提示 | F3 | 编辑当前命令", "提示 | F4 | 用户定义命令", "提示 | ALT+空格 / ALT+R | 激活 ALTRun", "提示 | 热键 / ESC / 失去焦点 | 关闭 ALTRun", "提示 | 回车 / ALT+序号 | 运行命令", "提示 | 上下箭头键 | 选择上一个或下一个命令", "提示 | CTRL+D | 使用文件管理器定位命令所在目录"]
    CHN[51] := "提示: "                                                 ; 50~99 Tips
    CHN[52] := "推荐使用热键激活 (ALT + 空格)"
    CHN[53] := "智能排序 - 根据使用频率自动调整命令优先级 (排序)"
    CHN[54] := "上/下箭头 = 上/下一个命令"
    CHN[55] := "Esc = 清除输入 / 关闭窗口"
    CHN[56] := "回车 = 运行当前命令"
    CHN[57] := "Alt + 序号 = 运行指定的命令"
    CHN[58] := "以 + 开头 = 新建命令"
    CHN[59] := "F3 = 直接编辑当前命令 (.ini)"
    CHN[60] := "F2 = 配置选项设置"
    CHN[61] := "Ctrl+I = 重建文件搜索数据库"
    CHN[62] := "F1 = ALTRun 帮助&关于"
    CHN[63] := "ALT + 空格 = 显示 / 隐藏窗口"
    CHN[64] := "Ctrl+Q = 重新加载 ALTRun"
    CHN[65] := "Ctrl + 序号 = 选择指定的命令"
    CHN[66] := "Alt + F4 = 退出"
    CHN[67] := "Ctrl+D = 使用文件管理器定位当前命令所在目录"
    CHN[68] := "F4 = 直接编辑用户定义命令 (.ini)"
    CHN[69] := "以空格开头 = 使用 Everything 搜索文件"
    CHN[70] := "Ctrl+'+' = 增加当前命令的优先级"
    CHN[71] := "Ctrl+'-' = 减少当前命令的优先级"
    CHN[100] := ["常规", "界面", "热键", "索引", "Listary", "插件", "状态统计", "关于"] ; 100~149 Options window (General - Check Listview)
    CHN[101] := "随系统自动启动"
    CHN[102] := "添加到“发送到”菜单"
    CHN[103] := "添加到“开始”菜单"
    CHN[104] := "显示托盘图标 (系统任务栏中)"
    CHN[105] := "失去焦点时关闭窗口"
    CHN[106] := "窗口置顶"
    CHN[107] := "显示窗口标题栏"
    CHN[108] := "使用 Windows XP 主题"
    CHN[109] := "按下 [ESC] 清除输入, 再次按下关闭窗口 (取消勾选: 直接关闭窗口)"
    CHN[110] := "保留最近一次输入和匹配结果"
    CHN[111] := "显示命令图标"
    CHN[112] := "使用“发送到”时, 追溯 .lnk 目标文件"
    CHN[113] := "保存历史记录"
    CHN[114] := "保存运行日志"
    CHN[115] := "搜索时匹配完整路径"
    CHN[116] := "显示网格 - 在列表的行和列之间提供边界线"
    CHN[117] := "显示标题 - 显示列表的标题 (顶部行包含列标题)"
    CHN[118] := "显示命令列表序号"
    CHN[119] := "显示命令列表边框线"
    CHN[120] := "智能排序 - 根据使用频率自动调整命令优先级 (排序)"
    CHN[121] := "智能匹配 - 模糊和智能匹配和过滤结果"
    CHN[122] := "搜索时匹配字符串开头 (取消勾选: 匹配任意位置)"
    CHN[123] := "显示提示信息 (状态栏)"
    CHN[124] := "显示命令执行次数 (状态栏)"
    CHN[125] := "显示状态栏 (窗口底部)"
    CHN[126] := "显示主窗口 [运行] 按钮"
    CHN[127] := "显示主窗口 [选项] 按钮"
    CHN[128] := "双缓冲绘图, 改善窗口闪烁 (Win XP+)"
    CHN[129] := "启用快速结构计算"
    CHN[130] := "简化路径 - 仅显示文件/文件夹/应用程序名称, 而非完整路径"
    CHN[131] := "设置语言为简体中文 (Simplified Chinese)"
    CHN[132] := "搜索时匹配拼音首字母"
    CHN[150] := "文件管理器"                                            ; 150~159 Options window (Other than Check Listview)
    CHN[151] := "Everything"
    CHN[152] := "历史命令数量"
    CHN[160] := "索引"                                                  ; 160~169 Index
    CHN[161] := "索引位置"
    CHN[162] := "索引文件类型"
    CHN[163] := "索引排除项"
    CHN[164] := "索引目录深度"
    CHN[170] := "界面"                                                  ; 170~189 GUI
    CHN[171] := "搜索结果数量"
    CHN[172] := "每列宽度"
    CHN[173] := "字体 (主界面)"
    CHN[174] := "字体 (选项页)"
    CHN[175] := "字体 (状态栏)"
    CHN[176] := "主窗口尺寸 (宽 x 高)"
    CHN[177] := "命令列表尺寸 (宽 x 高)"
    CHN[178] := "控件颜色"
    CHN[179] := "背景颜色"
    CHN[180] := "背景图片"
    CHN[181] := "透明度"
    CHN[182] := "选择字体"
    CHN[183] := "选择颜色"
    CHN[190] := "热键"                                                  ; 190~209 Hotkey
    CHN[191] := "激活热键 (全局)"
    CHN[192] := "主热键"
    CHN[193] := "辅热键"
    CHN[194] := "可以同时设置两个热键"
    CHN[195] := "重置激活热键"
    CHN[200] := "快捷操作和热键 (非全局)"
    CHN[201] := "快捷键"
    CHN[202] := "触发操作"
    CHN[203] := "热键 1"
    CHN[204] := "热键 2"
    CHN[206] := "热键 3"
    CHN[210] := "Listary"                                               ; 210~219 Listary
    CHN[211] := "目录快速切换"
    CHN[212] := "文件管理器 ID"
    CHN[213] := "打开/保存对话框 ID"
    CHN[214] := "排除窗口 ID"
    CHN[215] := "热键"
    CHN[216] := "快速切换到 TC 目录"
    CHN[217] := "快速切换到 资源管理器"
    CHN[218] := "自动切换路径"
    CHN[220] := "插件"                                                  ; 220~299 Plugins
    CHN[221] := "文本末尾自动添加日期"
    CHN[222] := "应用到窗口 ID"
    CHN[223] := "热键"
    CHN[224] := "日期格式"
    CHN[225] := "扩展名前自动添加日期"
    CHN[229] := "条件触发快捷操作"
    CHN[230] := "如果窗口 ID 包含"
    CHN[231] := "热键 (自由定制)"
    CHN[232] := "触发操作"
    CHN[300] := "显示"                                                  ; 300+ 托盘菜单
    CHN[301] := "配置选项`tF2"
    CHN[302] := "重建索引`tCtrl+I"
    CHN[303] := "状态统计"
    CHN[304] := "关于`tF1"
    CHN[305] := "脚本信息"
    CHN[307] := "重新加载`tCtrl+Q"
    CHN[308] := "退出`tAlt+F4"
    CHN[309] := "检查更新"
    CHN[400] := "运行命令`tEnter"                                       ; 400+ 列表右键菜单
    CHN[401] := "定位命令`tCtrl+D"
    CHN[402] := "复制命令`tCtrl+C"
    CHN[403] := "新建命令`tCtrl+N"
    CHN[404] := "编辑命令`tF3"
    CHN[405] := "删除命令`tDelete"
    CHN[406] := "用户命令`tF4"
    CHN[407] := "复制状态栏信息"
    CHN[408] := "显示状态统计"
    CHN[500] := "30天前"
    CHN[501] := "当前"                                                  ; 500+ 状态统计
    CHN[502] := "运行过的命令总次数"
    CHN[503] := "今天激活程序的次数"
    CHN[600] := "关于"
    CHN[601] := "ALTRun 是由诸葛草帽开发的一款高效 Windows 启动器，是一款基于 <a href=`"https://www.autohotkey.com/docs/v1/`">AutoHotkey</a> 的开源项目。 "
        . "它提供了一种简洁高效的方式，让你能够快速查找系统中的任何内容，并以自己的方式启动任意应用程序。"
        . "`n`n配置文件`n" g_INI "`n`n程序文件`n" A_ScriptFullPath
        . "`n`n版本更新"
        . "`n<a href=`"https://github.com/zhugecaomao/ALTRun/releases`">https://github.com/zhugecaomao/ALTRun/releases</a>"
        . "`n`n源代码开源在 GitHub"
        . "`n<a href=`"https://github.com/zhugecaomao/ALTRun`">https://github.com/zhugecaomao/ALTRun</a>"
        . "`n`n有关更多详细信息，请参阅帮助和 Wiki 页面"
        . "`n<a href=`"https://github.com/zhugecaomao/ALTRun/wiki`">https://github.com/zhugecaomao/ALTRun/wiki</a>"
    CHN[700] := "命令管理器"
    CHN[701] := "命令"
    CHN[702] := "命令类型"
    CHN[703] := "命令行"
    CHN[704] := "快捷项/描述 (可搜索)"
    CHN[705] := "储存节段"
    CHN[706] := "命令权重"
    CHN[800] := "您确定要从命令节段"
    CHN[801] := "中删除以下命令吗?"
    CHN[802] := "命令已成功删除!"
    CHN[803] := "删除命令时发生错误!"

    Global g_LNG := IniRead(g_INI, "Config", "Chinese", 0) ? CHN : ENG
    g_LOG.Debug("SetLanguage: Set language to=" g_LNG[1])
    return
}

;;==================== Expression Eval =========================
Eval(expression) {
    ; 移除所有空格
    expression := StrReplace(expression, " ")
    
    ; 检查非法字符（只允许数字、运算符、括号、小数点）
    if (!RegExMatch(expression, "^[\d+\-*/^().]*$"))
        return 0

    ; 递归处理括号
    while RegExMatch(expression, "\(([^()]*)\)", &match) {
        result := EvalSimple(match[1])  ; 计算括号内的内容
        expression := StrReplace(expression, (match&&match[0]), result)
    }

    ; 计算最终无括号表达式
    Return EvalSimple(expression)
}

EvalSimple(expression) {            ; 计算不含括号的简单数学表达式
    ; 处理幂运算符 ^
    while RegExMatch(expression, "(-?\d+(\.\d+)?)([\^])(-?\d+(\.\d+)?)", &match) {
        base := match[1], exponent := match[4]
        result := base ** exponent  ; 执行幂运算
        expression := StrReplace(expression, (match&&match[0]), result)
    }

    ; 支持 ** 作为幂运算符替代
    while RegExMatch(expression, "(-?\d+(\.\d+)?)(\*\*)(-?\d+(\.\d+)?)", &match) {
        base := match[1], exponent := match[4]
        result := base ** exponent
        expression := StrReplace(expression, (match&&match[0]), result)
    }

    ; 处理乘除法运算
    while RegExMatch(expression, "(-?\d+(\.\d+)?)([*/])(-?\d+(\.\d+)?)", &match) {
        operand1 := match[1], operator := match[3], operand2 := match[4]
        result := (operator = "*") ? operand1 * operand2 : operand1 / operand2
        expression := StrReplace(expression, (match&&match[0]), result)
    }

    ; 处理加减法运算
    while RegExMatch(expression, "(-?\d+(\.\d+)?)([+\-])(-?\d+(\.\d+)?)", &match) {
        operand1 := match[1], operator := match[3], operand2 := match[4]
        result := (operator = "+") ? operand1 + operand2 : operand1 - operand2
        expression := StrReplace(expression, (match&&match[0]), result)
    }

    ; 返回最终结果
    Return expression
}

;;==================== Performance Test Only =========================

Test() {
    t := A_TickCount
    Loop 50
    {
        chr1 := Random(Ord("a"),Ord("z"))
        chr2 := Random(Ord("A"),Ord("Z")) ;65,90
        chr3 := Random(Ord("a"),Ord("z")) ;97,122

        Activate()
        myInputBox.Value := chr(chr1)
        OnSearchInput()
        Sleep 10
        myInputBox.Value := chr(chr1) " " chr(chr2)
        OnSearchInput()
        Sleep 10
        myInputBox.Value := chr(chr1) " " chr(chr2) " " chr(chr3)
        OnSearchInput()
    }
    t := A_TickCount - t
    g_LOG.Debug("mock test search ' " chr(chr1) " " chr(chr2) " " chr(chr3) " ' 50 times, elapsed time=" t)
    MsgBox "Search '" chr(chr1) " " chr(chr2) " " chr(chr3) "' elapsed time=" t
}

;;==================== Logger Class =========================

class Logger {
    __New(filename) {
        this.filename := filename
    }

    Debug(Msg) {
        ;OutputDebug(Msg)
        if (g_CONFIG["SaveLog"])
            FileAppend("[" . A_Now . "] " . Msg . "`n", this.filename)
    }
}

;;==================== Font Select Dialog =========================
; FontSelect() - Display the standard Windows font selection dialog

; AHK v2
; originally posted by maestrith 
; https://autohotkey.com/board/topic/94083-ahk-11-font-and-color-dialogs/

; ======================================================================
; Example
; ======================================================================

; Global fontObj

; oGui := Gui.New("","Change Font Example")
; oGui.OnEvent("close","gui_exit")
; ctl := oGui.AddEdit("w500 h200 vMyEdit1","Sample Text")
; ctl.SetFont("bold underline italic strike c0xFF0000")
; oGui.AddEdit("w500 h200 vMyEdit2","Sample Text")
; oGui.AddButton("Default","Change Font").OnEvent("click","button_click")
; oGui.Show()

; button_click(ctl,info) {
	; If (!isSet(fontObj))
		; fontObj := ""
	; fontObj := Map("name","Terminal","size",14,"color",0xFF0000,"strike",1,"underline",1,"italic",1,"bold",1) ; init font obj (optional)
	
	; fontObj := FontSelect(fontObj,ctl.gui.hwnd) ; shows the user the font selection dialog
	
	; If (!fontObj)
		; return ; to get info from fontObj use ... bold := fontObj["bold"], fontObj["name"], etc.
	
	; ctl.gui["MyEdit1"].SetFont(fontObj["str"],fontObj["name"]) ; apply font+style in one line, or...
	; ctl.gui["MyEdit2"].SetFont(fontObj["str"],fontObj["name"])
; }

; gui_exit(oGui) {
	; ExitApp
; }

; ======================================================================
; END Example
; ======================================================================

; to initialize fontObj (not required):
; ============================================
; fontObj := Map("name","Tahoma","size",14,"color",0xFF0000,"strike",1,"underline",1,"italic",1,"bold",1)

; ==================================================================
; fntName		= name of var to store selected font
; fontObj	    = name of var to store fontObj object
; hwnd			= parent gui hwnd for modal, leave blank for not modal
; effects		= allow selection of underline / strike out / italic
; ==================================================================
; fontObj output:
;
;	fontObj["str"]	= string to use with AutoHotkey to set GUI values - see examples
;	fontObj["hwnd"]	= handle of the font object to use with SendMessage - see examples
; ==================================================================
FontSelect(fontObject:="",hwnd:=0,Effects:=1) {
	fontObject := (fontObject="") ? Map() : fontObject
	logfont := Buffer((A_PtrSize = 4) ? 60 : 92, 0)
	uintVal := DllCall("GetDC","uint",0)
	LogPixels := DllCall("GetDeviceCaps","uint",uintVal,"uint",90)
	Effects := 0x041 + (Effects ? 0x100 : 0)
	
	fntName := fontObject.Has("name") ? fontObject["name"] : ""
	fontBold := fontObject.Has("bold") ? fontObject["bold"] : 0
	fontBold := fontBold ? 700 : 400
	fontItalic := fontObject.Has("italic") ? fontObject["italic"] : 0
	fontUnderline := fontObject.Has("underline") ? fontObject["underline"] : 0
	fontStrikeout := fontObject.Has("strike") ? fontObject["strike"] : 0
	fontSize := fontObject.Has("size") ? fontObject["size"] : 10
	fontSize := fontSize ? Floor(fontSize*LogPixels/72) : 16
	c := fontObject.Has("color") ? fontObject["color"] : 0
	
	c1 := Format("0x{:02X}",(c&255)<<16)	; convert RGB colors to BGR for input
	c2 := Format("0x{:02X}",c&65280)
	c3 := Format("0x{:02X}",c>>16)
	fontColor := Format("0x{:06X}",c1|c2|c3)
	
	NumPut "uint", fontSize, logfont
	NumPut "uint", fontBold, "char", fontItalic, "char", fontUnderline, "char", fontStrikeout, logfont, 16
	
	choosefont := Buffer(A_PtrSize=8?104:60,0), cap := choosefont.size
	NumPut "UPtr", hwnd, choosefont, A_PtrSize
	offset1 := (A_PtrSize = 8) ? 24 : 12
	offset2 := (A_PtrSize = 8) ? 36 : 20
	offset3 := (A_PtrSize = 4) ? 6 * A_PtrSize : 5 * A_PtrSize
	
	fontArray := Array([cap,0,"Uint"],[logfont.ptr,offset1,"UPtr"],[effects,offset2,"Uint"],[fontColor,offset3,"Uint"])
	
	for index,value in fontArray
		NumPut value[3], value[1], choosefont, value[2]
	
	if (A_PtrSize=8) {
		strput(fntName,logfont.ptr+28,"UTF-16")
		r := DllCall("comdlg32\ChooseFont","UPtr",CHOOSEFONT.ptr) ; cdecl 
		fntName := strget(logfont.ptr+28,"UTF-16")
	} else {
		strput(fntName,logfont.ptr+28,32,"UTF-8")
		r := DllCall("comdlg32\ChooseFontA","UPtr",CHOOSEFONT.ptr) ; cdecl
		fntName := strget(logfont.ptr+28,32,"UTF-8")
	}
	
	if !r
		return false
	
	fontObj := Map("bold",16,"italic",20,"underline",21,"strike",22)
	for a,b in fontObj
		fontObject[a] := NumGet(logfont,b,"UChar")
	
	fontObject["bold"] := (fontObject["bold"] < 188) ? 0 : 1
	
	c := NumGet(choosefont,(A_PtrSize=4)?6*A_PtrSize:5*A_PtrSize,"UInt") ; convert from BGR to RBG for output
	c1 := Format("0x{:02X}",(c&255)<<16), c2 := Format("0x{:02X}",c&65280), c3 := Format("0x{:02X}",c>>16)
	c := Format("0x{:06X}",c1|c2|c3)
	fontObject["color"] := c
	
	fontSize := NumGet(choosefont,A_PtrSize=8?32:16,"UInt") / 10 ; 32:16
	fontObject["size"] := fontSize
	fontHwnd := DllCall("CreateFontIndirect","uptr",logfont.ptr) ; last param "cdecl"
	fontObject["name"] := fntName
	
	logfont := "", choosefont := ""
	
	If (!fontHwnd) {
		fontObject := ""
		return 0
	} Else {
		fontObject["hwnd"] := fontHwnd
		b := fontObject["bold"] ? "bold" : ""
		i := fontObject["italic"] ? "italic" : ""
		s := fontObject["strike"] ? "strike" : ""
		c := fontObject["color"] ? "c" fontObject["color"] : ""
		z := fontObject["size"] ? "s" fontObject["size"] : ""
		u := fontObject["underline"] ? "underline" : ""
		fullStr := b "|" i "|" s "|" c "|" z "|" u
		str := ""
		Loop Parse fullStr, "|"
			If (A_LoopField)
				str .= A_LoopField " "
		fontObject["str"] := "norm " Trim(str)
		
		return fontObject
	}
}

; typedef struct tagLOGFONTW {
  ; LONG  lfHeight;                 |4        / 0
  ; LONG  lfWidth;                  |4        / 4
  ; LONG  lfEscapement;             |4        / 8
  ; LONG  lfOrientation;            |4        / 12
  ; LONG  lfWeight;                 |4        / 16
  ; BYTE  lfItalic;                 |1        / 20
  ; BYTE  lfUnderline;              |1        / 21
  ; BYTE  lfStrikeOut;              |1        / 22
  ; BYTE  lfCharSet;                |1        / 23
  ; BYTE  lfOutPrecision;           |1        / 24
  ; BYTE  lfClipPrecision;          |1        / 25
  ; BYTE  lfQuality;                |1        / 26
  ; BYTE  lfPitchAndFamily;         |1        / 27
  ; WCHAR lfFaceName[LF_FACESIZE];  |[32|64]  / 28  ---> size [60|92] -- 32 TCHARs [UTF-8|UTF-16]
; } LOGFONTW, *PLOGFONTW, *NPLOGFONTW, *LPLOGFONTW;


; typedef struct tagCHOOSEFONTW {
  ; DWORD        lStructSize;               |4        / 0
  ; HWND         hwndOwner;                 |[4|8]    / [ 4| 8]  A_PtrSize * 1
  ; HDC          hDC;                       |[4|8]    / [ 8|16]  A_PtrSize * 2
  ; LPLOGFONTW   lpLogFont;                 |[4|8]    / [12|24]  A_PtrSize * 3
  ; INT          iPointSize;                |4        / [16|32]  A_PtrSize * 4
  ; DWORD        Flags;                     |4        / [20|36]
  ; COLORREF     rgbColors;                 |4        / [24|40]
  ; LPARAM       lCustData;                 |[4|8]    / [28|48]
  ; LPCFHOOKPROC lpfnHook;                  |[4|8]    / [32|56]
  ; LPCWSTR      lpTemplateName;            |[4|8]    / [36|64]
  ; HINSTANCE    hInstance;                 |[4|8]    / [40|72]
  ; LPWSTR       lpszStyle;                 |[4|8]    / [44|80]
  ; WORD         nFontType;                 |2        / [48|88]
  ; WORD         ___MISSING_ALIGNMENT__;    |2        / [50|92]
  ; INT          nSizeMin;                  |4        / [52|96]
  ; INT          nSizeMax;                  |4        / [56|100] -- len: 60 / 104
; } CHOOSEFONTW;

;;==================== Color Select Dialog =========================
; ColorSelect() - Display the standard Windows color selection dialog

; AHK v2
; originally posted by maestrith 
; https://autohotkey.com/board/topic/94083-ahk-11-font-and-color-dialogs/

; #SingleInstance,Force

; Global defColor

; === optional input color object, max 16 indexes number 1-16 ===
; === can be Array(), [], or Map()
; defColor := Array(0xFF0000,0,0x00FF00,0,0x0000FF)
; ===============================================================
; Example
; ===============================================================

; global cc, defColor
; cc := 0x00FF00 ; green
; defColor := [0xAA0000,0x00AA00,0x0000AA]

; oGui := Gui.New("-MinimizeBox -MaximizeBox","Choose Color")
; oGui.OnEvent("close","close_event")
; oGui.OnEvent("escape","close_event")
; oGui.AddButton("w150","Choose Color").OnEvent("click","choose_event")
; oGui.BackColor := cc
; oGui.Show("")
; return


; choose_event(ctl,info) {
	; hwnd := ctl.gui.hwnd ; grab hwnd
	; cc := "0x" ctl.gui.BackColor ; pre-select color from gui background (optional)
	
	; cc := ColorSelect(cc,hwnd,defColor,0) ; hwnd and defColor are optional
	
	; If (cc = -1)
		; return
	
	; colorList := ""
	; For k, v in defColor {
		; If v {
			; colorList .= "Index: " k " / Color: " v "`r`n"
		; }
	; }
		
	; If cc
		; msgbox "Output color: " cc "`r`n`r`nCustom colors saved:`r`n`r`n" Trim(colorList,"`r`n")
	
	; ctl.gui.BackColor := cc ; set gui background color
; }

; close_event(guiObj) {
	; ExitApp
; }

; ===============================================================
; END Example
; ===============================================================

; =============================================================================================
; Color			= Start color
; hwnd			= Parent window
; custColorObj	= Use for input to init custom colors, or output to save custom colors, or both.
;                 ... custColorObj can be Array() or Map().
; disp			= full / basic ... full displays custom colors panel, basic does not
; =============================================================================================
; All params are optional.  With no hwnd dialog will show at top left of screen.  User must
; parse output custColorObj and decide how to save custom colors... no more automatic ini file.
; =============================================================================================

ColorSelect(Color := 0, hwnd := 0, &custColorObj := "",disp:=1) {
	Color := (Color = "") ? 0 : Color ; fix silly user "oops"
	disp := disp ? 0x3 : 0x1 ; init disp / 0x3 = full panel / 0x1 = basic panel
	
	c1 := Format("0x{:02X}",(Color&255)<<16)	; convert RGB colors to BGR for input
	c2 := Format("0x{:02X}",Color&65280)		; init start Color
	c3 := Format("0x{:02X}",Color>>16)
	Color := Format("0x{:06X}",c1|c2|c3)
	
	CUSTOM := Buffer(16 * A_PtrSize,0) ; init custom colors obj
	
	CHOOSECOLOR := Buffer(9 * A_PtrSize,0) ; init dialog
	size := CHOOSECOLOR.size
	
	If (IsObject(custColorObj)) {
		Loop 16 {
			If (custColorObj.Has(A_Index)) {
				col := custColorObj[A_Index] = "" ? 0 : custColorObj[A_Index] ; init "" to 0 (black)
				
				c4 := Format("0x{:02X}",(col&255)<<16)	; convert RGB colors to BGR for input
				c5 := Format("0x{:02X}",col&65280)		; 
				c6 := Format("0x{:02X}",col>>16)
				custCol := Format("0x{:06X}",c4|c5|c6)
				NumPut "UInt", custCol, CUSTOM, ((A_Index-1) * 4) ; type, number, target, offset
			}
		}
	}
	
	NumPut "UInt", size, CHOOSECOLOR, 0
	NumPut "UPtr", hwnd, CHOOSECOLOR, A_PtrSize
	NumPut "UInt", Color, CHOOSECOLOR, 3 * A_PtrSize
	NumPut "UInt", disp, CHOOSECOLOR, 5 * A_PtrSize
	NumPut "UPtr", CUSTOM.ptr, CHOOSECOLOR, 4 * A_PtrSize
	
	ret := DllCall("comdlg32\ChooseColor", "UPtr", CHOOSECOLOR.ptr, "UInt")
	
	if !ret
		return -1
	
	custColorObj := Array()
	Loop 16 {
		newCustCol := NumGet(CUSTOM, (A_Index-1) * 4, "UInt")
		c7 := Format("0x{:02X}",(newCustCol&255)<<16)	; convert RGB colors to BGR for input
		c8 := Format("0x{:02X}",newCustCol&65280)
		c9 := Format("0x{:02X}",newCustCol>>16)
		newCustCol := Format("0x{:06X}",c7|c8|c9)
		custColorObj.InsertAt(A_Index, newCustCol)
	}
	
	Color := NumGet(CHOOSECOLOR, 3 * A_PtrSize, "UInt")
	
	c1 := Format("0x{:02X}",(Color&255)<<16)	; convert RGB colors to BGR for input
	c2 := Format("0x{:02X}",Color&65280)
	c3 := Format("0x{:02X}",Color>>16)
	Color := Format("0x{:06X}",c1|c2|c3)
	
	CUSTOM := "", CHOOSECOLOR := ""
	
	return Color
}

;; 获取拼音首字母
GetFirstChar(str) {
    ; GBK编码区间对应的拼音首字母
	static array := [ [-20319,-20284,"A"], [-20283,-19776,"B"], [-19775,-19219,"C"]
        , [-19218,-18711,"D"], [-18710,-18527,"E"], [-18526,-18240,"F"], [-18239,-17923,"G"]
        , [-17922,-17418,"H"], [-17417,-16475,"J"], [-16474,-16213,"K"], [-16212,-15641,"L"]
        , [-15640,-15166,"M"], [-15165,-14923,"N"], [-14922,-14915,"O"], [-14914,-14631,"P"]
        , [-14630,-14150,"Q"], [-14149,-14091,"R"], [-14090,-13319,"S"], [-13318,-12839,"T"]
        , [-12838,-12557,"W"], [-12556,-11848,"X"], [-11847,-11056,"Y"], [-11055,-10247,"Z"] ]
	
	; 如果不包含中文字符，则直接返回原字符
	if !RegExMatch(str, "[^\x{00}-\x{ff}]")
		Return str

    out := ""
    for char in StrSplit(str) {
        code := Ord(char)
        if (code >= 0x2E80 && code <= 0x9FFF) {
            buf := Buffer(4)
            StrPut(char, buf, "CP936")
            nGBKCode := (NumGet(buf, 0, "UChar") << 8) + NumGet(buf, 1, "UChar") - 65536
            for i, a in array {
                if (nGBKCode >= a[1] && nGBKCode <= a[2]) {
                    out .= a[3]
                    break
                }
            }
        } else {
            out .= char
        }
    }
    return out
}