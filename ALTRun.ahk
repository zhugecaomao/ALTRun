;===================================================
; ALTRun - An effective launcher for Windows
; https://github.com/zhugecaomao/ALTRun
;===================================================
#Requires AutoHotkey v1.1+
#NoEnv
#SingleInstance Force
#NoTrayIcon
#Persistent
#Warn All, OutputDebug

FileEncoding, UTF-8
SendMode, Input
SetWorkingDir %A_ScriptDir%

;===================================================
; Declare global variables
;===================================================
Global g_LOG:= New Logger(A_Temp "\ALTRun.log")
, g_COMMANDS:= {}           ; All commands
, g_FALLBACK:= {}           ; Fallback commands
, g_HISTORYS:= {}           ; Execution history
, g_MATCHED := {}           ; Matched commands
, g_SECTION := {CONFIG      : "Config"
            ,GUI            : "Gui"
            ,DFTCMD         : "DefaultCommand"
            ,USERCMD        : "UserCommand"
            ,FALLBACK       : "FallbackCommand"
            ,HOTKEY         : "Hotkey"
            ,HISTORY        : "History"
            ,INDEX          : "Index"
            ,USAGE          : "Usage"}
, g_CONFIG  := {AutoStartup : 1
            ,EnableSendTo   : 1
            ,InStartMenu    : 1
            ,ShowTrayIcon   : 1
            ,HideOnLostFocus: 1
            ,AlwaysOnTop    : 1
            ,ShowCaption    : 1
            ,XPthemeBg      : 1
            ,EscClearInput  : 1
            ,KeepInput      : 1
            ,ShowIcon       : 1
            ,SendToGetLnk   : 1
            ,SaveHistory    : 1
            ,SaveLog        : 1
            ,MatchPath      : 0
            ,MatchPinyin    : 1
            ,ShowGrid       : 0
            ,ShowHdr        : 1
            ,ShowSN         : 1
            ,ShowBorder     : 1
            ,SmartRank      : 1
            ,SmartMatch     : 1
            ,MatchBeginning : 0
            ,ShowHint       : 1
            ,ShowRunCount   : 1
            ,ShowStatusBar  : 1
            ,ShowBtnRun     : 1
            ,ShowBtnOpt     : 1
            ,ShortenPath    : 1
            ,StruCalc       : 0
            ,RunCount       : 0
            ,HistoryLen     : 10
            ,DoubleBuffer   : 1
            ,AutoSwitchDir  : 0
            ,FileMgr        : "Explorer.exe"
            ,IndexDir       : "A_ProgramsCommon,A_StartMenu,C:\Other\Index\Location"
            ,IndexType      : "*.lnk,*.exe"
            ,IndexDepth     : 2
            ,IndexExclude   : "Uninstall *"
            ,Everything     : "C:\Apps\Everything.exe"
            ,DialogWin      : "ahk_class #32770"
            ,FileMgrID      : "ahk_class CabinetWClass, ahk_class TTOTAL_CMD"
            ,ExcludeWin     : "ahk_class SysListView32, ahk_exe Explorer.exe"
            ,Chinese        : InStr("7804,0004,0804,1004", A_Language) ? 1 : 0}
, g_HOTKEY  := {Hotkey1     : "F1"      ,Trigger1   : "Help"
            ,Hotkey2        : "F2"      ,Trigger2   : "Options"
            ,Hotkey3        : "F3"      ,Trigger3   : "EditCommand"
            ,Hotkey4        : "F4"      ,Trigger4   : "UserCommand"
            ,Hotkey5        : "^d"      ,Trigger5   : "OpenContainer"
            ,Hotkey6        : "^n"      ,Trigger6   : "NewCommand"
            ,Hotkey7        : ""        ,Trigger7   : "None"
            ,CondTitle      : "ahk_exe RAPTW.exe"
            ,CondHotkey     : "~Mbutton"
            ,CondAction     : "PTTools"
            ,GlobalHotkey1  : "!Space"
            ,GlobalHotkey2  : "!r"
            ,TotalCMDDir    : "^g"
            ,ExplorerDir    : "^e"
            ,AutoDateAtEnd  : "ahk_class TCmtEditForm,ahk_exe Notepad4.exe" ; TC file comment window, Notepad4
            ,AutoDateAEHKey : "^d"
            ,AutoDateBefExt : "ahk_class CabinetWClass,ahk_class Progman,ahk_class WorkerW,ahk_class #32770" ; Windows 资源管理器文件列表, 桌面文件, 文件保存对话框,TC 文件列表
            ,AutoDateBEHKey : "^d"}
, g_GUI     := {ListRows    : 9
            ,ColWidth       : "36,0,300,AutoHdr"
            ,Font           : "Microsoft YaHei,norm s9"
            ,OptsFont       : "Microsoft YaHei,norm s9"
            ,SBFont         : "Microsoft YaHei,norm s8"
            ,WinX           : 660
            ,WinY           : 300
            ,ListX          : 636
            ,ListY          : 230
            ,CtrlColor      : "Default"
            ,WinColor       : "Silver"
            ,Background     : "ALTRun.jpg"
            ,Transparency   : 230}
, g_RUNTIME := {Ini         : A_ScriptDir "\" A_ComputerName ".ini"     ; 程序运行需要的临时全局变量, 不需要用户参与修改, 不读写入ini
            ,WinName        : "ALTRun - Ver 2025.08.23"
            ,WinHide        : ""
            ,UseDisplay     : 0
            ,UseFallback    : 0
            ,ActiveCommand  : ""
            ,Input          : ""
            ,Arg            : ""                                        ; 用来调用管道的完整参数
            ,OneDrive       : EnvGet("OneDrive")                        ; Due to #NoEnv
            ,LV_ContextMenu : []
            ,TrayMenu       : []
            ,FuncList       : ""
            ,RegEx          : "imS)"                                    ; 字符匹配的正则表达式
            ,Max            : 1
            ,AppDate        : A_YYYY A_MM A_DD}
, g_USAGE   := {A_YYYY A_MM A_DD : 1}

g_LOG.Debug("///// ALTRun is starting /////")
LoadConfig("initialize")                                                ; Load ini config, iniWrite create ini whenever not exist
SetLanguage()

Global g_CHKLV      := {AutoStartup : g_LNG.101                         ; Options - General - CheckedListview
    ,EnableSendTo   : g_LNG.102 ,InStartMenu    : g_LNG.103
    ,ShowTrayIcon   : g_LNG.104 ,HideOnLostFocus: g_LNG.105
    ,AlwaysOnTop    : g_LNG.106 ,ShowCaption    : g_LNG.107
    ,XPthemeBg      : g_LNG.108 ,EscClearInput  : g_LNG.109
    ,KeepInput      : g_LNG.110 ,ShowIcon       : g_LNG.111
    ,SendToGetLnk   : g_LNG.112 ,SaveHistory    : g_LNG.113
    ,SaveLog        : g_LNG.114 ,MatchPath      : g_LNG.115
    ,ShowGrid       : g_LNG.116 ,ShowHdr        : g_LNG.117
    ,ShowSN         : g_LNG.118 ,ShowBorder     : g_LNG.119
    ,SmartRank      : g_LNG.120 ,SmartMatch     : g_LNG.121
    ,MatchBeginning : g_LNG.122 ,ShowHint       : g_LNG.123
    ,ShowRunCount   : g_LNG.124 ,ShowStatusBar  : g_LNG.125
    ,ShowBtnRun     : g_LNG.126 ,ShowBtnOpt     : g_LNG.127
    ,DoubleBuffer   : g_LNG.128 ,StruCalc       : g_LNG.129
    ,ShortenPath    : g_LNG.130 ,Chinese        : g_LNG.131
    ,MatchPinyin    : g_LNG.132}

;===================================================
; Create ContextMenu and TrayMenu
;===================================================
g_RUNTIME.LV_ContextMenu := [g_LNG.400 ",LVRunCommand,imageres.dll,-100"
    ,g_LNG.401 ",OpenContainer,imageres.dll,-3"
    ,g_LNG.402 ",CopyCommand,imageres.dll,-5314",""
    ,g_LNG.403 ",NewCommand,imageres.dll,-2"
    ,g_LNG.404 ",EditCommand,imageres.dll,-5306"
    ,g_LNG.405 ",DelCommand,imageres.dll,-5305"
    ,g_LNG.406 ",UserCommand,imageres.dll,-88"]
g_RUNTIME.TrayMenu := [g_LNG.300 ",ToggleWindow,imageres.dll,-100",""
    ,g_LNG.301 ",Options,imageres.dll,-114"
    ,g_LNG.302 ",Reindex,imageres.dll,-8"
    ,g_LNG.303 ",Usage,imageres.dll,-150"
    ,g_LNG.309 ",Update,imageres.dll,-5338"
    ,g_LNG.304 ",Help,imageres.dll,-99",""
    ,g_LNG.305 ",ScriptInfo,imageres.dll,-165"
    ,""
    ,g_LNG.307 ",Reload,imageres.dll,-5311"
    ,g_LNG.308 ",Exit,imageres.dll,-98"]

Menu, Tray, UseErrorLevel
for index, MenuItem in g_RUNTIME.LV_ContextMenu {
    if (MenuItem = "") {
        Menu, LV_ContextMenu, Add
        continue
    }
    Item := StrSplit(MenuItem, ",")
    Menu, LV_ContextMenu, Add, % Item[1], % Item[2]
    Menu, LV_ContextMenu, Icon, % Item[1], % Item[3], % Item[4]
}

if (g_CONFIG.ShowTrayIcon) {
    Menu, Tray, NoStandard
    Menu, Tray, Icon
    Menu, Tray, Icon, imageres.dll, -100                                ; Index of icon changes between Windows versions, refer to the icon by resource ID for consistency
    For Index, MenuItem in g_RUNTIME.TrayMenu
    {
        Item := StrSplit(MenuItem, ",")                                 ; Item[1,2,3,4] <-> Name,Func,Icon,IconNo
        Menu, Tray, Add , % Item[1], % Item[2]
        Menu, Tray, Icon, % Item[1], % Item[3], % Item[4]
    }
    Menu, Tray, Tip, % g_RUNTIME.WinName
    Menu, Tray, Default, % g_LNG.300
    Menu, Tray, Click, 1
}
;===================================================
; Load commands database and command history
; Update "SendTo", "Startup", "StartMenu" lnk
;===================================================
LoadCommands()
LoadHistory()

g_LOG.Debug("Updating 'SendTo' setting..." UpdateSendTo(g_CONFIG.EnableSendTo))
g_LOG.Debug("Updating 'Startup' setting..." UpdateStartup(g_CONFIG.AutoStartup))
g_LOG.Debug("Updating 'StartMenu' setting..." UpdateStartMenu(g_CONFIG.InStartMenu))

;===================================================
; 主窗口配置代码
;===================================================
Input_W := g_GUI.ListX - g_CONFIG.ShowBtnRun * 90 - g_CONFIG.ShowBtnOpt * 90
Enter_W := g_CONFIG.ShowBtnRun * 80
Enter_X := g_CONFIG.ShowBtnRun * 10
Opt_W   := g_CONFIG.ShowBtnOpt * 80
Opt_X   := g_CONFIG.ShowBtnOpt * 10

Gui, Main:Color, % g_GUI.WinColor, % g_GUI.CtrlColor
Gui, Main:Font, % StrSplit(g_GUI.Font, ",")[2], % StrSplit(g_GUI.Font, ",")[1]
Gui, % "Main:+HwndMainGuiHwnd" (g_CONFIG.AlwaysOnTop ? " +AlwaysOnTop" : " -AlwaysOnTop") (g_CONFIG.ShowCaption ? " +Caption" : " -Caption") (g_CONFIG.XPthemeBg ? " +Theme" : " -Theme")
Gui, Main:Default ; Set default GUI before any ListView / statusbar update
Gui, Main:Add, Edit, x12 y10 W%Input_W% -WantReturn vMyInput gOnSearchInput, % g_LNG.13
Gui, Main:Add, Button, % "x+"Enter_X " yp W" Enter_W " hp Default gRunCurrentCommand Hidden" !g_CONFIG.ShowBtnRun, % g_LNG.11
Gui, Main:Add, Button, % "x+"Opt_X " yp W" Opt_W " hp gOptions Hidden" !g_CONFIG.ShowBtnOpt, % g_LNG.12
Gui, Main:Add, ListView, % "x12 yp+36 W" g_GUI.ListX " H" g_GUI.ListY " vMyListView AltSubmit gOnClickListview -Multi" (g_CONFIG.DoubleBuffer ? " +LV0x10000" : "") (g_CONFIG.ShowHdr ? "" : " -Hdr") (g_CONFIG.ShowGrid ? " Grid" : "") (g_CONFIG.ShowBorder ? "" : " -E0x200"), % g_LNG.10 ; LV0x10000 Paints via double-buffering, which reduces flicker
Gui, Main:Add, Picture, x0 y0 0x4000000, % g_GUI.Background
Gui, Main:Font,,
Gui, Main:Font, % StrSplit(g_GUI.SBFont, ",")[2], % StrSplit(g_GUI.SBFont, ",")[1]
Gui, Main:Add, StatusBar, % "gOnClickStatusBar Hidden" !g_CONFIG.ShowStatusBar,

Loop, 4 {
    LV_ModifyCol(A_Index, StrSplit(g_GUI.ColWidth, ",")[A_Index])
}

SB_SetParts(g_GUI.WinX - 90 * g_CONFIG.ShowRunCount)
SB_SetIcon("imageres.dll",-150, 2)

ListResult(g_LNG.50)

if (g_CONFIG.ShowIcon) {
    Global ImageListID := IL_Create(10, 5, 0)                           ; Create an ImageList so that the ListView can display some icons
    Global IconMap     := {"DIR":IL_Add(ImageListID,"imageres.dll",-3)  ; Icon cache index, IconIndex=1/2/3/4 for type dir/func/url/eval
                       ,"FUNC":IL_Add(ImageListID,"imageres.dll",-100)
                       ,"URL":IL_Add(ImageListID,"imageres.dll",-144)
                       ,"EVAL":IL_Add(ImageListID,"imageres.dll",-182)}
    LV_SetImageList(ImageListID)                                        ; Attach the ImageLists to the ListView so that it can later display the icons
}

;===================================================
; Resolve command line arguments, %1% %2% or A_Args[1] A_Args[2]
;===================================================
g_LOG.Debug("Resolving command line args=" A_Args[1] " " A_Args[2])
if (A_Args[1] = "-Startup")
    g_RUNTIME.WinHide := " Hide"

if (A_Args[1] = "-SendTo") {
    g_RUNTIME.WinHide := " Hide"
    Path := A_Args[2]

    SplitPath Path, Desc,, fileExt,,                                   ; Extra name from _Path (if _Type is dir and has "." in path, nameNoExt will not get full folder name)

    Type := InStr(FileExist(Path), "D") ? "Dir" : "File"                ; Default Type is File, Set Type is Dir only if the file exists and is a directory
    
    if (fileExt = "lnk" && g_CONFIG.SendToGetLnk) {
        FileGetShortcut, %Path%, Path,, fileArg, Desc
        Path .= " " fileArg
    }
    CmdMgr(g_SECTION.USERCMD, Type, Path, Desc, 1, "")                  ; Add new command to database
}

Gui, Main:Show, % "w" g_GUI.WinX " h" g_GUI.WinY " Center" g_RUNTIME.WinHide, % g_RUNTIME.WinName

if (g_GUI.Transparency < 250)
    WinSet, Transparent, % g_GUI.Transparency, % g_RUNTIME.WinName

(g_CONFIG.HideOnLostFocus) ? OnMessage(0x06, "WM_ACTIVATE")

;===================================================
; Set Hotkey
;===================================================
Hotkey, % g_HOTKEY.GlobalHotkey1, ToggleWindow, UseErrorLevel
Hotkey, % g_HOTKEY.GlobalHotkey2, ToggleWindow, UseErrorLevel

Hotkey, IfWinActive, % g_RUNTIME.WinName                                ; Hotkey take effect only when ALTRun actived
Hotkey, !F4, Exit, UseErrorLevel
Hotkey, Tab, TabFunc, UseErrorLevel
Hotkey, F1, Help, UseErrorLevel
Hotkey, F2, Options, UseErrorLevel
Hotkey, F3, EditCommand, UseErrorLevel
Hotkey, F4, UserCommand, UseErrorLevel
Hotkey, ^q, Reload, UseErrorLevel
Hotkey, ^d, OpenContainer, UseErrorLevel
Hotkey, ^c, CopyCommand, UseErrorLevel
Hotkey, ^N, NewCommand, UseErrorLevel
Hotkey, Del, DelCommand, UseErrorLevel
Hotkey, ^i, Reindex, UseErrorLevel
Hotkey, ^NumpadAdd, RankUp, UseErrorLevel
Hotkey, ^NumpadSub, RankDown, UseErrorLevel
Hotkey, Down, NextCommand, UseErrorLevel
Hotkey, Up, PrevCommand, UseErrorLevel

;===================================================
; Run or locate command shortcut: Ctrl Alt Shift + No.
;===================================================
Loop, % g_GUI.ListRows
{
    Hotkey, !%A_Index%, RunSelectedCommand, UseErrorLevel ; ALT + No. run command
    Hotkey, ^%A_Index%, GotoCommand, UseErrorLevel        ; Ctrl + No. locate command
}

Loop, 7
    Hotkey, % g_HOTKEY["Hotkey"A_Index], % g_HOTKEY["Trigger"A_Index], UseErrorLevel ; Set Hotkey <-> Trigger, UseErrorLevel to Skips any warning dialogs

Hotkey, IfWinActive, % g_HOTKEY.CondTitle                 ; Conditional hotkey-action, mainly for workflow RAPT-MButton-PTTools
Hotkey, % g_HOTKEY.CondHotkey, % g_HOTKEY.CondAction, UseErrorLevel
Hotkey, IfWinActive                                       ; Reset hotkey condition

Listary()
AppControl()                                                            ; Set Listary Dir QuickSwitch, Set AppControl
Return

Activate() {
    Gui, Main:Show,, % g_RUNTIME.WinName

    WinWaitActive, % g_RUNTIME.WinName,, 3                              ; Use WinWaitActive 3s instead of previous Loop method
    {
        GuiControl, Main:Focus, MyInput
        ControlSend, Edit1, ^a, % g_RUNTIME.WinName                     ; Select all content in Input Box
    }
    UpdateUsage()
}

ToggleWindow() {
    WinActive(g_RUNTIME.WinName) ? MainGuiClose() : Activate()
}

OnSearchInput() {
    GuiControlGet, TempValue, Main:, MyInput
    SearchCommand(g_RUNTIME.Input := TempValue)                         ; Assign Input value to g_RUNTIME.Input and in the meantime, search it
}

SearchCommand(command := "") {
    g_MATCHED := {}
    Prefix    := SubStr(command, 1, 1)

    ; Handle fallback commands
    if (Prefix = "+" or Prefix = " " or Prefix = ">") {
        g_RUNTIME.ActiveCommand := g_FALLBACK[InStr("+ >", Prefix)]     ; Corresponding to fallback commands position no. 1, 2 & 3
        g_MATCHED.Push(g_RUNTIME.ActiveCommand)
        Return ListResult(g_RUNTIME.ActiveCommand)
    }

    ; Search commands
    for index, element in g_COMMANDS
    {
        splitResult := StrSplit(element, " | ")
        _Type := splitResult[1]
        _Path := splitResult[2]
        _Desc := splitResult[3]
        SplitPath, _Path, fileName                                      ; Extra name from _Path (if _Type is Dir and has "." in path, nameNoExt will not get full folder name)

        elementToSearch := g_CONFIG.MatchPath ? _Path " " _Desc : fileName " " _Desc ; search file name include extension, and desc (not search type for MatchBeginning option)
        if (g_CONFIG.MatchPinyin) {
            elementToSearch := GetFirstChar(elementToSearch)            ; 中文转为拼音首字母
        }

        if FuzzyMatch(elementToSearch, command) {
            g_MATCHED.Push(element)
            If (g_MATCHED.Length() = 1) {
                g_RUNTIME.ActiveCommand := element
            } else If (g_MATCHED.Length() >= g_GUI.ListRows)
                Break
        }
    }

    ; No matched command found
    if (g_MATCHED.Length() = 0) {
        if (EvalResult := Eval(command)) {
            RebarQty   := Ceil((EvalResult-40*2) / 300) + 1
            EvalResFmt := FormatThousand(EvalResult)
            g_MATCHED.Push("Eval | " EvalResFmt)

            if (g_CONFIG.StruCalc) {
                g_MATCHED.Push("`nEval | Beam width = " EvalResFmt " mm")
                g_MATCHED.Push(" | Main bar no. = " RebarQty " (" Round((EvalResult-40*2) / (RebarQty - 1)) " c/c)")
                g_MATCHED.Push("`nEval | As = " EvalResFmt " mm2")
                g_MATCHED.Push(" | Rebar = " Ceil(EvalResult/132.7) "H13 / " Ceil(EvalResult/201.1) "H16 / " Ceil(EvalResult/314.2) "H20 / " Ceil(EvalResult/490.9) "H25 / " Ceil(EvalResult/804.2) "H32")
            }
            Return ListResult(JoinResult(g_MATCHED), True)
        }

        g_RUNTIME.UseFallback   := True
        g_MATCHED               := g_FALLBACK
        g_RUNTIME.ActiveCommand := g_FALLBACK[1]
    } Else {
        g_RUNTIME.UseFallback   := False
    }
    ListResult(JoinResult(g_MATCHED))
}

JoinResult(commands) {
    Result := ""
    For index, command in commands {
        Result .= (index = 1 ? "" : "`n") command
    }
    Return Result
}

ListResult(text := "", UseDisplay := false)
{
    Gui, Main:Default                                                   ; Set default GUI before update any listview or statusbar
    GuiControl, Main:-Redraw, MyListView                                ; Improve performance by disabling redrawing during load.
    LV_Delete()
    g_RUNTIME.UseDisplay := UseDisplay

    Loop Parse, text, `n, `r
    {
        splitResult := StrSplit(A_LoopField, " | ")
        _Type       := splitResult[1]
        _Path       := splitResult[2]
        _Desc       := splitResult[3]
        IconIndex   := GetIconIndex(_Path, _Type)

        SplitPath, _Path, fileName                                      ; Extra name from _Path (if _Type is Dir and has "." in path, nameNoExt will not get full folder name)
        PathToShow  := (g_CONFIG.ShortenPath) ? fileName : _Path        ; Show Full path / Shorten path

        LV_Add("Icon" IconIndex, (g_CONFIG.ShowSN ? A_Index : ""), _Type, PathToShow, _Desc)
    }

    LV_Modify(1, "Select Focus Vis")                                    ; Select 1st row
    GuiControl, Main:+Redraw, MyListView
    SetStatusBar()

    ; For index, value in IconMap
    ;     elements .= index " = " value "`n"
    ; OutputDebug, % "IconMap length is " IconMap.Count() ", elements are `n" elements
}

GetIconIndex(_Path, _Type)                                              ; Get file's icon index
{
    if g_CONFIG.ShowIcon = 0
        Return 0                                                        ; No icon to show, return 0
    Switch (_Type) {
    Case "DIR" : Return 1
    Case "FUNC","CMD","TIP","提示": Return 2
    Case "URL" : Return 3
    Case "EVAL": Return 4
    Case "FILE": {
            _Path := AbsPath(_Path)                                     ; Must store in var for afterward use, trim space (in AbsPath)
            SplitPath, _Path,,, FileExt                                 ; Get the file's extension.
            if FileExt in EXE,ICO,ANI,CUR,LNK
                IconIndex := IconMap.HasKey(_Path) ? IconMap[_Path] : GetIcon(_Path, _Path) ; File path exist in ImageList, get the index, several calls can be avoided and performance is greatly improved
            else                                                        ; Some other extension/file-type like pdf or xlsx
                IconIndex := IconMap.HasKey(FileExt) ? IconMap[FileExt] : GetIcon(_Path, FileExt)
            Return IconIndex
        }
    }
}

GetIcon(_Path, ExtOrPath) {                                             ; Get file's icon
    VarSetCapacity(sfi, sfi_size := 698)                                ; Calculate buffer size required for SHFILEINFO structure.
    if (!DllCall("Shell32\SHGetFileInfoW", "Str", _Path, "UInt", 0, "Ptr", &sfi, "UInt", sfi_size, "UInt", 0x101)) ; 0x101 is SHGFI_ICON+SHGFI_SMALLICON
        IconIndex = 9999999                                             ; Set it out of bounds to display a blank icon.
    else {                                                              ; Icon successfully loaded. Extract the hIcon member from the structure
        hIcon := NumGet(sfi, 0)                                         ; Add the HICON directly to the small-icon lists.
        IconIndex := DllCall("ImageList_ReplaceIcon", "ptr", ImageListID, "int", -1, "ptr", hIcon) + 1 ; Uses +1 to convert the returned index from zero-based to one-based:
        DllCall("DestroyIcon", "ptr", hIcon)                            ; Now that it's been copied into the ImageLists, the original should be destroyed
        IconMap[ExtOrPath] := IconIndex                                 ; Cache the icon based on file type (xlsx, pdf) or path (exe, lnk) to save memory and improve loading performance
    }
    Return IconIndex
}

AbsPath(Path, KeepRunAs := False)                                       ; Convert path to absolute path
{
    Path := Trim(Path)

    if (!KeepRunAs)
        Path := StrReplace(Path,  "*RunAs ", "")                        ; Remove *RunAs (Admin Run) to get absolute path

    if (InStr(Path, "A_") = 1)                                          ; Resolve path like A_ScriptDir, some server path has "Plot A_IGLS" in it, so InStr must be 1
        Path := %Path%

    Path := StrReplace(Path, "%Temp%", A_Temp)
    Path := StrReplace(Path, "%OneDrive%", g_RUNTIME.OneDrive)          ; Convert OneDrive to absolute path due to #NoEnv
    Return Path
}

RelativePath(Path)                                                      ; Convert path to relative path
{
    Path := StrReplace(Path, A_Temp, "%Temp%")
    Path := StrReplace(Path, g_RUNTIME.OneDrive, "%OneDrive%")
    Return Path
}

EnvGet(EnvVar) {
    EnvGet, OutputVar, %EnvVar%
    Return OutputVar
}

RunCommand(originCmd)
{
    MainGuiClose()
    ParseArg()
    g_RUNTIME.UseDisplay := false

    _Type := StrSplit(originCmd, " | ")[1]
    _Path := AbsPath(StrSplit(originCmd, " | ")[2], True)

    switch (_Type) {
        case "FILE","URL","CMD":
            Run % _Path,, UseErrorLevel
            if ErrorLevel
                MsgBox Could not open "%_Path%"
        case "DIR":
            OpenDir(_Path)
        case "FUNC":
            IsFunc(_Path) ? %_Path%() : MsgBox Could not find function "%_Path%"
        Default:                                                        ; For all other un-defined type
            Run % _Path,, UseErrorLevel
    }

    if (g_CONFIG.SaveHistory) {
        g_HISTORYS.InsertAt(1, originCmd " Arg=" g_RUNTIME.Arg)        ; Adjust command history

        (g_HISTORYS.Length() > g_CONFIG.HistoryLen) ? g_HISTORYS.Pop()

        IniDelete, % g_RUNTIME.Ini, % g_SECTION.HISTORY
        for index, element in g_HISTORYS
            IniWrite(g_SECTION.HISTORY, index, element)                 ; Save command history
    }

    UpdateRunCount()
    (g_CONFIG.SmartRank) ? UpdateRank(originCmd)
    g_LOG.Debug("Execute:" g_CONFIG.RunCount " = " originCmd)
}

TabFunc()
{
    GuiControlGet, CurrCtrl, Main:FocusV                                ; Limit tab to switch between Edit1 & ListView only
    GuiControl, Main:Focus, % (CurrCtrl = "MyInput") ? "MyListView" : "MyInput"
}

PrevCommand() {
    ChangeCommand(-1, False)
}

NextCommand() {
    ChangeCommand(1, False)
}

GotoCommand() {
    index := SubStr(A_ThisHotkey, 0, 1)                                 ; Get index from hotkey (select specific command = Shift + index)

    if (index <= g_MATCHED.Count()) {
        ChangeCommand(index, True)
        g_RUNTIME.ActiveCommand := g_MATCHED[index]
    }
}

ChangeCommand(Step = 1, ResetSelRow = False) {
    Gui, Main:Default                                                   ; Use it before any LV update

    SelRow := ResetSelRow ? Step : LV_GetNext() + Step                  ; Get target row no. to be selected
    SelRow := SelRow > LV_GetCount() ? 1 : SelRow                       ; Listview cycle selection (Mod has bug on upward cycle)
    SelRow := SelRow < 1 ? LV_GetCount() : SelRow
    g_RUNTIME.ActiveCommand := g_MATCHED[SelRow]                        ; Get current command from selected row

    LV_Modify(SelRow, "Select Focus Vis")                               ; make new index row selected, Focused, and Visible
    SetStatusBar()
}

OnClickListview() {                                                     ; ListView g label actions (left / double click) behavior
    Gui, Main:Default                                                   ; Use it before any LV update
    focusedRow := LV_GetNext(0, "Focused")                              ; 查找焦点行, 仅对焦点行进行操作而不是所有选择的行:
    if (!focusedRow)                                                    ; 没有焦点行
        Return

    g_RUNTIME.ActiveCommand := g_MATCHED[focusedRow]                    ; Get current command from focused row

    if (A_GuiEvent = "RightClick")
        Menu, LV_ContextMenu, Show

    else if (A_GuiEvent = "DoubleClick" and g_RUNTIME.ActiveCommand)    ; It will clear g_SECTION.USERCMD/Index/DftCMD when g_RUNTIME.ActiveCommand = "" eg. first tip page
        RunCommand(g_RUNTIME.ActiveCommand)

    else if (A_GuiEvent = "Normal")                                     ; Left click behavior
        SetStatusBar()
}

LVRunCommand() {                                                        ; ListView ContextMenu (right click & its menu) actions
    Gui, Main:Default                                                   ; Use it before any LV update
    focusedRow := LV_GetNext(0, "Focused")                              ; Check focused row, only operate focusd row instead of all selected rows
    if (!focusedRow)                                                    ; Return if no focused row is found
        Return

    g_RUNTIME.ActiveCommand := g_MATCHED[focusedRow]                    ; Get current command from focused row
    RunCommand(g_RUNTIME.ActiveCommand)                                 ; Execute the command if the user selected "Run Enter"
}

CopyCommand() {                                                         ; ListView ContextMenu (right click & its menu) actions
    Gui, Main:Default                                                   ; Use it before any LV update
    focusedRow := LV_GetNext(0, "Focused")                              ; Check focused row, only operate focusd row instead of all selected rows
    if (!focusedRow)                                                    ; Return if no focused row is found
        Return

    g_RUNTIME.ActiveCommand := g_MATCHED[focusedRow]                    ; Get current command from focused row
    LV_GetText(Text, focusedRow, 3) ? (A_Clipboard := Text)             ; Get the text from the focusedRow's 3rd field.
}

OnClickStatusBar() {
    if (A_GuiEvent = "RightClick" and A_EventInfo = 1) {
        Menu, SB_ContextMenu, Add, Copy, SBContextMenu
        Menu, SB_ContextMenu, Icon, Copy, imageres.dll, -5314
        Menu, SB_ContextMenu, Show
    } else if (A_EventInfo = 2) ; Apply for normal click and rightclick
        Usage()
}

SBContextMenu() {
    StatusBarGetText, A_Clipboard, 1, % g_RUNTIME.WinName
}

ScriptInfo() {
    ListLines
}

MainGuiEscape() {
    (g_CONFIG.EscClearInput and g_RUNTIME.Input) ? ClearInput() : MainGuiClose()
}

MainGuiClose() {                                                        ; If GuiClose is a function, the GUI is hidden by default
    g_CONFIG.KeepInput ? "" : ClearInput()
    Gui, Main:Hide
    SetStatusBar("TIP")                                                 ; Update StatusBar tip information after GUI hide
}

Exit() {
    ExitApp
}

Reload() {
    Reload
}

Test() {
    t := A_TickCount
    Loop 50
    {
        random,chr1,asc("a"),asc("z")
        random,chr2,asc("A"),asc("Z") ;65,90
        random,chr3,asc("a"),asc("z") ;97,122

        Activate()
        GuiControl, Main:Text, MyInput, % chr(chr1)
        Sleep, 10
        GuiControl, Main:Text, MyInput, % chr(chr1) " " chr(chr2)
        Sleep, 10
        GuiControl, Main:Text, MyInput, % chr(chr1) " " chr(chr2) " " chr(chr3)
    }
    t := A_TickCount - t
    g_LOG.Debug("mock test search ' " chr(chr1) " " chr(chr2) " " chr(chr3) " ' 50 times, elapsed time=" t)
    MsgBox % "Search '" chr(chr1) " " chr(chr2) " " chr(chr3) "' elapsed time=" t
}

ClearInput() {
    GuiControl, Main:Text, MyInput,
    GuiControl, Main:Focus, MyInput
}

SetStatusBar(Mode := "CMD") {                                           ; Set StatusBar text, Mode 1: Current command (default), 2: Hint, 3: Any text
    Global
    Gui, Main:Default                                                   ; Set default GUI window before any ListView / StatusBar operate
    Switch (Mode) {
        Case "CMD" : SBText := StrSplit(g_RUNTIME.ActiveCommand, " | ")[2]
        Case "TIP" : {
            Random, index, 52, 71                                       ; Randomly select a tip from hint list g_LNG 52~71
            SBText := g_LNG.51 g_LNG[index]
        }
        Default    : SBText := Mode
    }
    SB_SetText(SBText, 1), SB_SetText(g_CONFIG.RunCount, 2)
}

RunCurrentCommand() {
    RunCommand(g_RUNTIME.ActiveCommand)
}

ParseArg() {
    Global
    commandPrefix := SubStr(g_RUNTIME.Input, 1, 1)

    if (commandPrefix = "+" || commandPrefix = " " || commandPrefix = ">") {
        Return g_RUNTIME.Arg := SubStr(g_RUNTIME.Input, 2)              ; 直接取命令为参数
    }

    if (InStr(g_RUNTIME.Input, " ") && !g_RUNTIME.UseFallback) {        ; 用空格来判断参数
        g_RUNTIME.Arg := SubStr(g_RUNTIME.Input, InStr(g_RUNTIME.Input, " ") + 1)
    }
    else if (g_RUNTIME.UseFallback) {
        g_RUNTIME.Arg := g_RUNTIME.Input
    }
    else {
        g_RUNTIME.Arg := ""
    }
}

FuzzyMatch(Haystack, Needle) {
    Needle := StrReplace(Needle, "+", "\+")                             ; For Eval (preceded by a backslash to be seen as literal)
    Needle := StrReplace(Needle, "*", "\*")                             ; For Eval (eg. 25+5 or 6*5 will show Eval result instead of match file with "30")
    Needle := StrReplace(Needle, " ", ".*")                             ; 空格直接替换为匹配任意字符
    Return RegExMatch(Haystack, g_RUNTIME.RegEx . Needle)
}

UpdateRank(originCmd, showRank := false, inc := 1) {
    RANKSEC := g_SECTION.DFTCMD "|" g_SECTION.USERCMD "|" g_SECTION.INDEX
    Loop Parse, RANKSEC, |                                              ; Update Rank for related sections
    {
        IniRead, Rank, % g_RUNTIME.Ini, %A_LoopField%, %originCmd%, KeyNotFound

        if (Rank = "KeyNotFound" or Rank = "ERROR" or originCmd = "")   ; If originCmd not exist in this section, then check next section
            continue                                                    ; Skips the rest of a loop and begins a new one.
        else if Rank is integer                                         ; If originCmd exist in this section, then update it's rank.
            Rank += inc
        else
            Rank := inc

        if (Rank < 0)                                                   ; 如果降到负数,都设置成 -1,然后屏蔽/排除
            Rank := -1

        IniWrite(A_LoopField, originCmd, Rank)                          ; Update new Rank for originCmd

        showRank ? SetStatusBar("Rank for current command : " Rank)
    }
    LoadCommands()                                                      ; New rank will take effect in real-time by LoadCommands again
}

UpdateUsage() {
    currDate := A_YYYY . A_MM . A_DD
    if (g_RUNTIME.AppDate != currDate) {
        g_RUNTIME.AppDate := currDate
        g_USAGE[currDate] := 1
    } else {
        g_USAGE[currDate]++
    }
    g_RUNTIME.Max := Max(g_RUNTIME.Max, g_USAGE[currDate])
    IniWrite(g_SECTION.USAGE, currDate, g_USAGE[currDate])
}

UpdateRunCount() {
    g_CONFIG.RunCount++
    IniWrite(g_SECTION.CONFIG, "RunCount", g_CONFIG.RunCount)           ; Update RunCount in g_CONFIG
    SetStatusBar()
}

RunSelectedCommand() {
    index := SubStr(A_ThisHotkey, 0, 1)
    RunCommand(g_MATCHED[index])
}

RankUp() {
    UpdateRank(g_RUNTIME.ActiveCommand, true)
}

RankDown() {
    UpdateRank(g_RUNTIME.ActiveCommand, true, -1)
}

LoadCommands() {
    g_COMMANDS          := {}                                           ; Clear g_COMMANDS and g_FALLBACK list
    g_FALLBACK          := {}
    RankString          := ""
    g_RUNTIME.FuncList  := ""                                           ; Clear FuncList, FuncList is used to store all functions for Options window

    Loop Parse, % LoadConfig("commands"), `n                            ; Read commands sections (built-in, user & index), read each line, separate key and value
    {
        command := StrSplit(A_LoopField, "=")[1]                        ; pass first string (key) to command
        rank    := StrSplit(A_LoopField, "=")[2]                        ; pass second string (value) to rank

        if (command != "" and rank > 0)
        {
            RankString .= rank "`t" command "`n"

            splitResult := StrSplit(command, " | ")
            g_RUNTIME.FuncList .= (splitResult[1] = "Func" and IsFunc(splitResult[2])) ? splitResult[2] "|" : ""
        }
    }
    Sort, RankString, R N
    Loop Parse, RankString, `n
    {
        command := StrSplit(A_LoopField, "`t")[2]
        g_COMMANDS.Push(command)
    }
    
    FALLBACKCMDSEC := IniRead(g_SECTION.FALLBACK)                       ; Read FALLBACK section, initialize it if section not exist
    if (FALLBACKCMDSEC = "") {
        IniWrite, 
        (Ltrim
        ; Fallback Commands show when search result is empty
        ; Commands in order, modify as desired
        ; Format: Command Type | Command | Description
        ; Command type: File, Dir, CMD, URL
        ;
        Func | NewCommand | New Command
        Func | Everything | Search by Everything
        Func | Google | Search Clipboard or Input by Google
        Func | AhkRun | Run Command use AutoHotkey Run
        Func | Bing | Search Clipboard or Input by Bing
        ), % g_RUNTIME.Ini, % g_SECTION.FALLBACK
        FALLBACKCMDSEC := IniRead(g_SECTION.FALLBACK)
    }
    Loop Parse, FALLBACKCMDSEC, `n                                      ; Get and verify each FBCommand (Rank not necessary) and push it to g_FALLBACK
    {
        g_FALLBACK.Push(A_LoopField)
    }
    Return g_LOG.Debug("Loading commands list...OK")
}

LoadHistory() {
    if (g_CONFIG.SaveHistory) {
        Loop % g_CONFIG.HistoryLen
        {
            g_HISTORYS.Push(IniRead(g_SECTION.HISTORY, A_Index, A_Space))
        }
    } else
        IniDelete, % g_RUNTIME.Ini, % g_SECTION.HISTORY
}

GetCmdOutput(command) {
    TempFile   := A_Temp "\ALTRun.stdout"
    FullCommand = %ComSpec% /C "%command% > %TempFile%"

    RunWait, %FullCommand%, %A_Temp%, Hide
    FileRead, Result, %TempFile%
    FileDelete, %TempFile%
    Return RTrim(Result, "`r`n")                                        ; Remove result rightmost/last "`r`n"
}

GetRunResult(command) {                                                 ; 运行CMD并取返回结果方式2
    shell := ComObjCreate("WScript.Shell")                              ; WshShell object: https://msdn.microsoft.com/en-us/library/aew9yb99
    exec := shell.Exec(ComSpec " /C " command)                          ; Execute a single command via cmd.exe
    Return exec.StdOut.ReadAll()                                        ; Read and Return the command's output
}

OpenDir(Path) {
    Path := AbsPath(Path)

    Run, % g_CONFIG.FileMgr " """ Path """",, UseErrorLevel

    if ErrorLevel
        MsgBox, 4096, % g_RUNTIME.WinName, Error found, error code : %A_LastError%

    g_LOG.Debug("Open Dir=" Path)
}

OpenContainer() {
    Path := AbsPath(StrSplit(g_RUNTIME.ActiveCommand, " | ")[2])

    if (g_CONFIG.FileMgr = "Explorer.exe")
        Run, % g_CONFIG.FileMgr " /Select`, """ Path """",, UseErrorLevel
    else
        Run, % g_CONFIG.FileMgr " /P """ Path """",, UseErrorLevel    ; /P Parent folder

    if ErrorLevel
        MsgBox, 4096, % g_RUNTIME.WinName, Error found, error code : %A_LastError%

    g_LOG.Debug("Open Container=" Path)
}

WM_ACTIVATE(wParam, lParam) {                                           ; Close on lose focus
    if (wParam > 0)                                                     ; wParam > 0: window being activated
        Return
    else if (wParam <= 0 && WinExist(g_RUNTIME.WinName) && !g_RUNTIME.UseDisplay) ; wParam <= 0: window being deactivated (lost focus)
        MainGuiClose()
}

UpdateSendTo(create := true) {                                          ; the lnk in SendTo must point to a exe
    lnkPath := StrReplace(A_StartMenu, "\Start Menu", "\SendTo\") "ALTRun.lnk"

    if (!create) {
        FileDelete, %lnkPath%
        Return "Disabled"
    }

    if (A_IsCompiled)
        FileCreateShortcut, "%A_ScriptFullPath%", %lnkPath%, ,-SendTo
        , Send command to ALTRun User Command list, imageres.dll, , -100
    else
        FileCreateShortcut, "%A_AhkPath%", %lnkPath%, , "%A_ScriptFullPath%" -SendTo
        , Send command to ALTRun User Command list, imageres.dll, , -100
    Return "OK"
}

UpdateStartup(create := true) {
    lnkPath := A_Startup "\ALTRun.lnk"

    if (!create) {
        FileDelete, %lnkPath%
        Return "Disabled"
    }

    FileCreateShortcut, %A_ScriptFullPath%, %lnkPath%, %A_ScriptDir%
        , -startup, ALTRun - An effective launcher, imageres.dll, , -100
    Return "OK"
}

UpdateStartMenu(create := true) {
    lnkPath := A_Programs "\ALTRun.lnk"

    if (!create) {
        FileDelete, %lnkPath%
        Return "Disabled"
    }

    FileCreateShortcut, %A_ScriptFullPath%, %lnkPath%, %A_ScriptDir%
        , -StartMenu, ALTRun, imageres.dll, , -100
    Return "OK"
}

Reindex() {                                                             ; Re-create Index section
    IniDelete, % g_RUNTIME.Ini, % g_SECTION.INDEX
    for dirIndex, dir in StrSplit(g_CONFIG.IndexDir, ",")
    {
        searchPath := AbsPath(dir)
        searchPath := RegExReplace(searchPath, "\\+$")                  ; Remove trailing backslashes

        for extIndex, ext in StrSplit(g_CONFIG.IndexType, ",")
        {
            Loop Files, %searchPath%\%ext%, R
            {
                ; Calculate path relative to searchPath and count subdir levels
                rel := SubStr(A_LoopFileLongPath, StrLen(searchPath) + 2) ; +2 to skip the backslash
                seps := (rel = "") ? 0 : StrLen(rel) - StrLen(StrReplace(rel, "\", "")) ; Count backslashes to determine depth

                if (seps >= g_CONFIG.IndexDepth)                        ; If file is deeper than allowed depth, skip it.
                    continue

                if (g_CONFIG.IndexExclude != "" && RegExMatch(A_LoopFileLongPath, g_CONFIG.IndexExclude))
                    continue                                            ; Skip this file and move on to the next loop.

                IniWrite(g_SECTION.INDEX, "File | " A_LoopFileLongPath, "1") ; Store file type for later use
                Progress, %A_Index%, %A_LoopFileName%, ReIndexing..., Reindex
            }
        }
        Progress, Off
    }

    g_LOG.Debug("Indexing search database...")
    TrayTip, % g_RUNTIME.WinName, ReIndex database finish successfully. , 8
    LoadCommands()
}

Help() {
    Options(8)
}

Usage() {
    Options(7)
}

Update() {
    Run, https://github.com/zhugecaomao/ALTRun/releases
}

Listary() {                                                             ; Listary 快速更换保存/打开对话框路径
    g_LOG.Debug("Listary function starting...")

    Loop Parse, % g_CONFIG.FileMgrID, `,                                ; File Manager Class, default is Windows Explorer & Total Commander
        GroupAdd, FileMgrID, %A_LoopField%

    Loop Parse, % g_CONFIG.DialogWin, `,                                ; 需要QuickSwith的窗口, 包括打开/保存对话框等
        GroupAdd, DialogBox, %A_LoopField%

    Loop Parse, % g_CONFIG.ExcludeWin, `,                               ; 排除特定窗口,避免被 Auto-QuickSwitch 影响
        GroupAdd, ExcludeWin, %A_LoopField%

    if (g_CONFIG.AutoSwitchDir) {
        g_LOG.Debug("Listary Auto-QuickSwitch Enabled.")
        Loop
        {
            WinWaitActive ahk_class TTOTAL_CMD
                WinGet, ThisHWND, ID, A
            WinWaitNotActive

            If(WinActive("ahk_group DialogBox") && !WinActive("ahk_group ExcludeWin")) ; 检测当前窗口是否符合打开保存对话框条件
            {
                WinGetActiveTitle, Title
                WinGet, ActiveProcess, ProcessName, A

                g_LOG.Debug("Listary dialog detected, active window ahk_title=" Title ", ahk_exe=" ActiveProcess)
                LocateTC()                                              ; NO Return, as will terimate loop (AutoSwitchDir)
            }
        }
    }
    Hotkey, IfWinActive, ahk_group DialogBox                            ; 设置对话框路径定位热键,为了不影响其他程序热键,设置只对打开/保存对话框生效
    Hotkey, % g_HOTKEY.ExplorerDir, LocateExplorer, UseErrorLevel       ; Ctrl+E 把打开/保存对话框的路径定位到资源管理器当前浏览的目录
    Hotkey, % g_HOTKEY.TotalCMDDir, LocateTC, UseErrorLevel             ; Ctrl+G 把打开/保存对话框的路径定位到TC当前浏览的目录
    Hotkey, IfWinActive
}

LocateTC() {                                                            ; Get TC current dir path, and change dialog box path to it
    ClipSaved := ClipboardAll 
    Clipboard :=
    SendMessage 1075, 2029, 0, , ahk_class TTOTAL_CMD
    ClipWait, 200
    OutDir=%Clipboard%\                                                 ; 结尾添加\ 符号,变为路径,试图解决AutoCAD不识别路径问题
    Clipboard := ClipSaved 
    ClipSaved := 

    ChangePath(OutDir)
}

LocateExplorer() {                                                      ; Get Explorer current dir path, and change dialog box path to it
    Loop 9
    {
        ControlGetText, Dir, ToolbarWindow32%A_Index%, ahk_class CabinetWClass
    } until (InStr(Dir,"Address"))
 
    Dir:=StrReplace(Dir,"Address: ","")
    if (Dir="Computer")
        Dir:="C:\"

    If (SubStr(Dir,2,2) != ":\")                                        ; then Explorer lists it as one of the library directories such as Music or Pictures
        Dir:=% "C:\Users\" A_UserName "\" Dir

    ChangePath(Dir)
}

ChangePath(Dir) {
    ControlGetText, w_Edit1Text, Edit1, A
    ControlClick, Edit1, A
    ControlSetText, Edit1, %Dir%, A
    ControlSend, Edit1, {Enter}, A
    ;Sleep,100
    ;ControlSetText, Edit1, %w_Edit1Text%, A                            ; 还原之前的窗口 File Name 内容, 在选择文件的对话框时没有问题, 但是在选择文件夹的对话框有Bug,暂时注释掉
    g_LOG.Debug("Listary change path=" Dir)
}

UserCommand() {
    Run, % "Notepad.exe " . g_RUNTIME["Ini"]                            ; TO-DO: To use build-in command manger to manage commands
}

NewCommand() {                                                          ; From command "New Command" or GUI context menu "New Command"
    CmdMgr(g_SECTION.USERCMD, , , g_RUNTIME.Arg, 1, "")
}

EditCommand() {
    if (g_RUNTIME.ActiveCommand) {
        RANKSEC := g_SECTION.DFTCMD "|" g_SECTION.USERCMD "|" g_SECTION.INDEX
        Loop Parse, RANKSEC, |                                          ; Update Rank for related sections
        {
            IniRead, Rank, % g_RUNTIME.Ini, %A_LoopField%, % g_RUNTIME.ActiveCommand, KeyNotFound

            if (Rank = "KeyNotFound" or Rank = "ERROR")                 ; If g_RUNTIME.ActiveCommand not exist in this section, then check next section
                continue                                                ; Skips the rest of a loop and begins a new one.
            else if Rank is integer                                     ; If g_RUNTIME.ActiveCommand exist in this section, then update it's rank.
            {
                Section := A_LoopField
                Type    := StrSplit(g_RUNTIME.ActiveCommand, " | ")[1]
                Path    := StrSplit(g_RUNTIME.ActiveCommand, " | ")[2]
                Desc    := StrSplit(g_RUNTIME.ActiveCommand, " | ")[3]
                CmdMgr(Section, Type, Path, Desc, Rank, g_RUNTIME.ActiveCommand)
                Break
            }
        }
    }
}

DelCommand() {                                                          ; Delete current command
    Global
    if (g_RUNTIME.ActiveCommand) {
        RANKSEC := g_SECTION.DFTCMD "|" g_SECTION.USERCMD "|" g_SECTION.INDEX
        Loop Parse, RANKSEC, |
        {
            IniRead, Rank, % g_RUNTIME.Ini, %A_LoopField%, % g_RUNTIME.ActiveCommand, KeyNotFound

            if (Rank = "KeyNotFound" or Rank = "ERROR")                 ; If g_RUNTIME.ActiveCommand not exist in this section, then check next section
                continue                                                ; Skips the rest of a loop and begins a new one.
            else
            {
                MsgBox, 52, % g_RUNTIME.WinName, % g_LNG.800 " [ " A_LoopField " ] " g_LNG.801 "`n`n" g_RUNTIME.ActiveCommand
                IfMsgBox Yes
                {
                    IniDelete, % g_RUNTIME.Ini, %A_LoopField%, % g_RUNTIME.ActiveCommand
                    if (!ErrorLevel)
                        MsgBox,64, % g_RUNTIME.WinName, % g_LNG.802 "`n`n" g_RUNTIME.ActiveCommand
                    Break
                }
            }
        }
        LoadCommands()
    }
}

CmdMgr(Section := "UserCommand", Type := "File", Path := "", Desc := "", Rank := 1, OriginCmd := "") { ; 命令管理窗口
    Global
    g_LOG.Debug("Starting Command Manager... Args=" Section "|" Type "|" Path "|" Desc "|" Rank)

    _Section  := Section
    _Type     := Type
    _Path     := RelativePath(Path)
    _Desc     := Desc
    _Rank     := Rank
    _OriginCmd:= OriginCmd

    Gui, CmdMgr:New
    Gui, CmdMgr:Font, S9 Norm, Microsoft Yahei
    Gui, CmdMgr:Add, GroupBox, w600 h260, % g_LNG.701
    Gui, CmdMgr:Add, Text, x25 yp+30, % g_LNG.702
    Gui, CmdMgr:Add, DropDownList, x160 yp-5 w130 v_Type, % StrReplace("File|Dir|Cmd|URL|Func|", _Type, _Type . "|",, 1)
    Gui, CmdMgr:Add, Text, x315 yp+5, % g_LNG.705
    Gui, CmdMgr:Add, Edit, x435 yp-5 w130 Disabled v_Section, %_Section%
    Gui, CmdMgr:Add, Text, x25 yp+60, % g_LNG.703
    Gui, CmdMgr:Add, Edit, x160 yp-5 w405 -WantReturn v_Path, %_Path%
    Gui, CmdMgr:Add, Button, x575 yp w30 hp gSelectCmdPath, ...
    Gui, CmdMgr:Add, Text, x25 yp+80, % g_LNG.704
    Gui, CmdMgr:Add, Edit, x160 yp-5 w405 -WantReturn v_Desc, %_Desc%
    Gui, CmdMgr:Add, Text, x25 yp+60, % g_LNG.706
    Gui, CmdMgr:Add, Edit, x160 yp-5 w405 +Number v_Rank, %_Rank%
    Gui, CmdMgr:Add, Button, Default x420 w90 gCmdMgrButtonOK, % g_LNG.7
    Gui, CmdMgr:Add, Button, x521 yp w90 gCmdMgrButtonCancel, % g_LNG.8
    Gui, CmdMgr:Show, Center, % g_LNG.700
    GuiControl, Focus, _Path
}

SelectCmdPath() {
    Global
    Gui, CmdMgr:+OwnDialogs                                             ; Make open dialog Modal
    Gui, CmdMgr:Submit, NoHide
    if (_Type = "Dir")
        FileSelectFolder, _Path, , 3
    else
        FileSelectFile, _Path, 3, , Select, All File (*.*)

    if (_Path != "")
        GuiControl,, _Path, %_Path%
}

CmdMgrButtonOK() {
    Global
    Gui, CmdMgr:Submit
    _Desc := _Desc ? " | " _Desc : _Desc

    if (_Path = "")
    {
        MsgBox,64, Command Manager, Command Path is empty`, please input correct command path!
        Return
    } else {
        IniDelete, % g_RUNTIME.Ini, % _Section, % _OriginCmd
        IniWrite(_Section, _Type " | " _Path _Desc, _Rank)
        if (!ErrorLevel)
            MsgBox,64, Command Manager, The following command added / modified successfully!`n`n[ %_Section% ]`n`n%_Type% | %_Path% %_Desc% = %_Rank%
    }
    LoadCommands()
}

CmdMgrGuiEscape() {
    Gui, CmdMgr:Destroy
}

CmdMgrButtonCancel() {
    Gui, CmdMgr:Destroy
}

CmdMgrGuiClose() {
    Gui, CmdMgr:Destroy
}

AppControl() {                                                          ; AppControl (Ctrl+D 自动添加日期)
    Loop Parse, % g_HOTKEY.AutoDateBefExt, `,
        GroupAdd, FileListMangr, %A_LoopField%
    Hotkey, IfWinActive, ahk_group FileListMangr                        ; 针对所有设定好的程序 按Ctrl+D自动在文件(夹)名之后添加日期
    Hotkey, % g_HOTKEY.AutoDateBEHKey, RenameWithDate, UseErrorLevel

    Loop Parse, % g_HOTKEY.AutoDateAtEnd, `,
        GroupAdd, TextBox, %A_LoopField%
    Hotkey, IfWinActive, ahk_group TextBox
    Hotkey, % g_HOTKEY.AutoDateAEHKey, LineEndAddDate, UseErrorLevel
    Hotkey, IfWinActive
}

PTTools()
{
    IfWinNotExist, PT Tools
        Run, %A_ScriptDir%\PTTools.ahk,, UseErrorLevel
    else
        WinActivate
}

RenameWithDate() {                                                      ; 针对所有设定好的程序 按Ctrl+D自动在文件(夹)名之后添加日期
    ControlGetFocus, CurrCtrl, A                                        ; 获取当前激活的窗口中的聚焦的控件名称
    if (InStr(CurrCtrl, "Edit") or InStr(CurrCtrl, "Scintilla"))        ; 如果当前激活的控件为Edit类或者Scintilla1(Notepad2),则Ctrl+D功能生效
        NameAddDate("FileListMangr", CurrCtrl)
    Else
        SendInput ^D
    Return
}

LineEndAddDate() {                                                      ; 针对TC File Comment对话框　按Ctrl+D自动在备注文字之后添加日期
    FormatTime, CurrentDate,, dd.MM.yyyy
    SendInput {End}
    Sleep, 10
    SendInput {Blind}{Text} - %CurrentDate%
    g_LOG.Debug("Add Date At End= - " CurrentDate)
}

NameAddDate(WinName, CurrCtrl, isFile:= True) {                         ; 在文件（夹）名编辑框中添加日期,CurrCtrl为当前控件(名称编辑框Edit)
    ControlGetText, EditCtrlText, %CurrCtrl%, A
    SplitPath, EditCtrlText, fileName, fileDir, fileExt, nameNoExt
    FormatTime, CurrentDate,, dd.MM.yyyy

    if (isFile && fileExt != "" && StrLen(fileExt) < 5 && !RegExMatch(fileExt,"^\d+$")) ; 如果是文件,而且有真实文件后缀名,才加日期在后缀名之前
    {
        if RegExMatch(nameNoExt, " - \d{2}\.\d{2}\.\d{4}$") {
                baseName := RegExReplace(nameNoExt, " - \d{2}\.\d{2}\.\d{4}$", "")
            }
            else {
                baseName := nameNoExt
            }
            NameWithDate := baseName " - " CurrentDate "." fileExt
    }
    else {
        NameWithDate := EditCtrlText " - " CurrentDate
    }
    ControlClick, %CurrCtrl%, A
    ControlSetText, %CurrCtrl%, %NameWithDate%, A
    SendInput {Blind}{End}
    g_LOG.Debug(WinName ", RenameWithDate=" NameWithDate)
}

FormatThousand(Number) {                                                ; Function to add thousand separator
    Return RegExReplace(Number, "\G\d+?(?=(\d{3})+(?:\D|$))", "$0" ",")
}

Options(ActTab := 1) {
    Global                                                              ; Assume-global mode
    t := A_TickCount
    ActTab := ActTab+0 ? ActTab : 1                                     ; Convert ActTab to number, default is 1 (for case like [Option`tF2])
    Gui, Setting:New, +AlwaysOnTop +HwndOptsHwnd, % g_LNG.1 ; Omit +OwnerMain, +OwnDialogs, lagging window
    Gui, Setting:Font, % StrSplit(g_GUI.OptsFont, ",")[2], % StrSplit(g_GUI.OptsFont, ",")[1]
    Gui, Setting:Add, Tab3, vCurrTab Choose%ActTab%, % g_LNG.100
    Gui, Setting:Tab, 1 ; CONFIG Tab
    Gui, Setting:Add, ListView, w500 h300 Checked -Multi AltSubmit -Hdr vOptListView, % g_LNG.1
    
    For key, description in g_CHKLV
        LV_Add("Check" g_CONFIG[key], description)
    LV_ModifyCol(1, "AutoHdr")

    Gui, Setting:Add, Text, x24 yp+320, % g_LNG.150
    Gui, Setting:Add, ComboBox, x130 yp-5 w394 vg_FileMgr, % g_CONFIG.FileMgr "||Explorer.exe|C:\Apps\TotalCMD.exe /O /T /S"
    Gui, Setting:Add, Text, x24 yp+40, % g_LNG.151
    Gui, Setting:Add, ComboBox, x130 yp-5 w394 vg_Everything, % g_CONFIG.Everything "||C:\Apps\Everything.exe"
    Gui, Setting:Add, Text, x24 yp+40, % g_LNG.152
    Gui, Setting:Add, DropDownList, % "x130 yp-5 w394 Sort vg_HistoryLen Choose" g_CONFIG.HistoryLen * 0.1, 10|20|30|40|50|60

    Gui, Setting:Tab, 2 ; GUI Tab
    Gui, Setting:Add, GroupBox, w500 h420, % g_LNG.170
    Gui, Setting:Add, Text, x33 yp+25 , % g_LNG.171
    Gui, Setting:Add, DropDownList, % "x183 yp-5 w330 vg_ListRows Choose" g_GUI.ListRows, 1||2|3|4|5|6|7|8|9| ; ListRows limit <= 9
    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.172
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_ColWidth, % g_GUI.ColWidth "||20,0,460,AutoHdr|30,46,460,AutoHdr"
    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.176
    Gui, Setting:Add, Edit, x183 yp-5 w120 +Number vg_WinX, % g_GUI.WinX
    Gui, Setting:Add, Text, x345 yp, x
    Gui, Setting:Add, Edit, x393 yp w120 +Number vg_WinY, % g_GUI.WinY

    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.173
    Gui, Setting:Font,,
    Gui, Setting:Font, % StrSplit(g_GUI.Font, ",")[2], % StrSplit(g_GUI.Font, ",")[1]
    Gui, Setting:Add, Edit, x183 yp w240 r1 -E0x200 +ReadOnly vg_Font, % g_GUI.Font
    Gui, Setting:Font,,
    Gui, Setting:Font, % StrSplit(g_GUI.OptsFont, ",")[2], % StrSplit(g_GUI.OptsFont, ",")[1]
    Gui, Setting:Add, Button, x433 yp-5 w80 vSelectFont gSelectFont, % g_LNG.182
    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.174
    Gui, Setting:Add, Edit, x183 yp w240 r1 -E0x200 +ReadOnly vg_OptsFont, % g_GUI.OptsFont
    Gui, Setting:Add, Button, x433 yp-5 w80 vSelectOptsFont gSelectOptsFont, % g_LNG.182
    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.175
    Gui, Setting:Font,,
    Gui, Setting:Font, % StrSplit(g_GUI.SBFont, ",")[2], % StrSplit(g_GUI.SBFont, ",")[1]
    Gui, Setting:Add, Edit, x183 yp w240 r1 -E0x200 +ReadOnly vg_SBFont, % g_GUI.SBFont
    Gui, Setting:Font,,
    Gui, Setting:Font, % StrSplit(g_GUI.OptsFont, ",")[2], % StrSplit(g_GUI.OptsFont, ",")[1]
    Gui, Setting:Add, Button, x433 yp-5 w80 vSelectSBFont gSelectSBFont, % g_LNG.182

    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.178
    Gui, Setting:Add, Edit, % "x183 yp w240 r1 -E0x200 +ReadOnly vg_CtrlColor c" g_GUI.CtrlColor, % g_GUI.CtrlColor
    Gui, Setting:Add, Button, x433 yp-5 w80 vSelectCtrlColor gSelectCtrlColor, % g_LNG.183
    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.179
    Gui, Setting:Add, Edit, % "x183 yp w240 r1 -E0x200 +ReadOnly vg_WinColor c" g_GUI.WinColor, % g_GUI.WinColor
    Gui, Setting:Add, Button, x433 yp-5 w80 vSelectWinColor gSelectWinColor, % g_LNG.183

    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.180
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_Background, % g_GUI.Background "||ALTRun.jpg|None|C:\Path\BG.jpg"
    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.181
    Gui, Setting:Add, Slider, x183 yp-5 w330 Range50-255 TickInterval5 Tooltip vg_Transparency, % g_GUI.Transparency

    Gui, Setting:Tab, 3 ; Hotkey Tab
    Gui, Setting:Add, GroupBox, w500 h115, % g_LNG.191
    Gui, Setting:Add, Text, x33 yp+25 , % g_LNG.192
    Gui, Setting:Add, Hotkey, x285 yp-4 w230 vg_GlobalHotkey1, % g_HOTKEY.GlobalHotkey1
    Gui, Setting:Add, Text, x33 yp+35 , % g_LNG.193
    Gui, Setting:Add, Hotkey, x285 yp-4 w230 vg_GlobalHotkey2,% g_HOTKEY.GlobalHotkey2
    Gui, Setting:Add, Text, x33 yp+35, % g_LNG.194
    Gui, Setting:Add, Link, x285 yp w230 gResetHotkey, % "<a>" g_LNG.195 "</a>"

    Gui, Setting:Add, GroupBox, x24 yp+38 w500 h290, % g_LNG.200
    loop 7
    {
        Gui, Setting:Add, Text, x33 yp+40 , % g_LNG.201
        Gui, Setting:Add, Hotkey, x143 yp-5 w120 vg_Hotkey%A_Index%, % g_HOTKEY["Hotkey" A_Index]
        Gui, Setting:Add, Text, x285 yp+5, % g_LNG.202
        Gui, Setting:Add, DropDownList, x395 yp-5 w120 vg_Trigger%A_Index%, % SetFuncList(g_HOTKEY["Trigger" A_Index])
    }

    Gui, Setting:Tab, 4 ; INDEX Tab
    Gui, Setting:Add, GroupBox, w500 h190, % g_LNG.160
    Gui, Setting:Add, Text, x33 yp+25, % g_LNG.161
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_IndexDir, % g_CONFIG.IndexDir "||A_ProgramsCommon,A_StartMenu"
    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.162
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_IndexType, % g_CONFIG.IndexType "||*.lnk,*.exe"
    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.164
    Gui, Setting:Add, DropDownList, % "x183 yp-5 w330 vg_IndexDepth Choose" g_CONFIG.IndexDepth, 1|2|3||4|5|6|7|8|9|
    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.163
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_IndexExclude, % g_CONFIG.IndexExclude "||Uninstall *"

    Gui, Setting:Tab, 5 ; LISTARTY TAB
    Gui, Setting:Add, GroupBox, w500 h145, % g_LNG.211
    Gui, Setting:Add, Text, x33 yp+30 , % g_LNG.212
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_FileMgrID, % g_CONFIG.FileMgrID "||ahk_class CabinetWClass|ahk_class CabinetWClass, ahk_class TTOTAL_CMD"
    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.213
    Gui, Setting:Add, Combobox, x183 yp-5 w330 vg_DialogWin, % g_CONFIG.DialogWin "||ahk_class #32770"
    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.214
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_ExcludeWin, % g_CONFIG.ExcludeWin "||ahk_class SysListView32|ahk_class SysListView32, ahk_exe Explorer.exe"
    Gui, Setting:Add, GroupBox, x24 yp+50 w500 h145, % g_LNG.215
    Gui, Setting:Add, Text, x33 yp+30, % g_LNG.216
    Gui, Setting:Add, Hotkey, x183 yp-5 w330 vg_TotalCMDDir, % g_HOTKEY.TotalCMDDir
    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.217
    Gui, Setting:Add, Hotkey, x183 yp-5 w330 vg_ExplorerDir, % g_HOTKEY.ExplorerDir
    Gui, Setting:Add, CheckBox, % "x33 yp+45 vg_AutoSwitchDir checked" g_CONFIG.AutoSwitchDir, % g_LNG.218

    Gui, Setting:Tab, 6 ; PLUGINS TAB
    Gui, Setting:Add, GroupBox, w500 h110, % g_LNG.221
    Gui, Setting:Add, Text, x33 yp+30, % g_LNG.222
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_AutoDateAtEnd, % g_HOTKEY.AutoDateAtEnd "||ahk_class TCmtEditForm,ahk_exe Notepad4.exe"
    Gui, Setting:Add, Text, x33 yp+45 , % g_LNG.223
    Gui, Setting:Add, Hotkey, x183 yp-5 w80 vg_AutoDateAEHKey, % g_HOTKEY.AutoDateAEHKey
    Gui, Setting:Add, Text, x300 yp+5, % g_LNG.224
    Gui, Setting:Add, DropDownList, x395 yp-5 w120, - dd.MM.yyyy||

    Gui, Setting:Add, GroupBox, x24 y+30 w500 h110, % g_LNG.225
    Gui, Setting:Add, Text, x33 yp+30, % g_LNG.222
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_AutoDateBefExt, % g_HOTKEY.AutoDateBefExt "||ahk_class CabinetWClass,ahk_class Progman,ahk_class WorkerW,ahk_class #32770"
    Gui, Setting:Add, Text, x33 yp+45 , % g_LNG.223
    Gui, Setting:Add, Hotkey, x183 yp-5 w80 vg_AutoDateBEHKey, % g_HOTKEY.AutoDateBEHKey
    Gui, Setting:Add, Text, x300 yp+5, % g_LNG.224
    Gui, Setting:Add, DropDownList, x395 yp-5 w120, - dd.MM.yyyy||

    Gui, Setting:Add, GroupBox, x24 y+30 w500 h110, % g_LNG.229
    Gui, Setting:Add, Text, x33 yp+30 , % g_LNG.230
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_CondTitle, % g_HOTKEY.CondTitle "||"
    Gui, Setting:Add, Text, x33 yp+45 , % g_LNG.231
    Gui, Setting:Add, ComboBox, x183 yp-5 w80 vg_CondHotkey, % g_HOTKEY.CondHotkey "||"
    Gui, Setting:Add, Text, x300 yp+5, % g_LNG.232
    Gui, Setting:Add, DropDownList, x395 yp-5 w120 vg_CondAction, % SetFuncList(g_HOTKEY.CondAction)

    Gui, Setting:Tab, 7 ; USAGE TAB
    Gui, Setting:Add, GroupBox, x66 y80 w445 h300,

    for Date, Count in g_USAGE {
        Gui, Setting:Add, Progress, % "c94DD88 Vertical y96 w14 h280 xm+" 50+A_Index*14 " Range0-" g_RUNTIME.Max+10, % Count
    }

    Gui, Setting:Add, Text, x24 yp-5 cGray, % g_RUNTIME.Max
    Gui, Setting:Add, Text, x24 yp+140 cGray, % Round(g_RUNTIME.Max/2)
    Gui, Setting:Add, Text, x24 yp+140 cGray, 0
    Gui, Setting:Add, Text, x66 yp+15 cGray, % g_LNG.500
    Gui, Setting:Add, Text, x476 yp cGray, % g_LNG.501
    Gui, Setting:Add, Text, x66 yp+33, % g_LNG.502
    Gui, Setting:Add, Edit, x400 yp-5 w100 r1 -E0x200 +ReadOnly Right vg_RunCount, % g_CONFIG.RunCount
    Gui, Setting:Add, Text, x66 yp+35, % g_LNG.503
    Gui, Setting:Add, Edit, x400 yp-5 w100 r1 -E0x200 +ReadOnly Right, % g_USAGE[A_YYYY . A_MM . A_DD]

    Gui, Setting:Tab, 8 ; ABOUT TAB
    Gui, Setting:Add, Picture, x33 y+20 w48 h-1 Icon-100, imageres.dll
    Gui, Setting:Add, Text, x96 yp+5, % g_RUNTIME.WinName
    Gui, Setting:Add, Link, xp yp+45 w400, % g_LNG.601

    Gui, Setting:Tab                                                    ; 后续添加的控件将不属于前面的选项卡控件
    Gui, Setting:Add, Button, Default x278 w80 vSettingButtonOK gSettingButtonOK, % g_LNG.7
    Gui, Setting:Add, Button, x368 yp w80 gSettingButtonCancel, % g_LNG.8
    Gui, Setting:Add, Button, x458 yp w80 gSettingButtonHelp, % g_LNG.9
    Gui, Setting:Show, Center, % g_LNG.1

    Hotkey, % g_HOTKEY.GlobalHotkey1, Off, UseErrorLevel
    Hotkey, % g_HOTKEY.GlobalHotkey2, Off, UseErrorLevel
    t := A_TickCount - t
    g_LOG.Debug("Loading options window... ActTab=" ActTab ", elapsed time=" t "ms")
    OutputDebug, % "Loading options window... ActTab=" ActTab ", elapsed time=" t "ms"
}

ResetHotkey() {
    GuiControl, Setting:, g_GlobalHotkey1, !Space
    GuiControl, Setting:, g_GlobalHotkey2, !r
}

SetFuncList(FuncName) {                                                 ; Set the DropDownList items for FuncName
    Global
    return StrReplace("None|" g_RUNTIME.FuncList, FuncName, FuncName . "|",, 1)
}

SelectFont() {
    Global
	; Set the fontObj (optional) - only set the ones you want to pre-select
	; fontObj := Object("name","Terminal","size",14,"color",0xFF0000,"strike",1,"underline",1,"italic",1,"bold",1)
	fontObj := Object("name", StrSplit(g_GUI.Font, ",")[1])
	fontObj := FontSelect(fontObj,OptsHwnd) ; shows the font selection dialog
	If (!fontObj)
		return

    Gui, Setting:Font,,
    Gui, Setting:Font, % fontObj["str"], % fontObj["name"]
	GuiControl, Setting:Font, g_Font
    GuiControl, Setting:, g_Font, % fontObj["name"] "," fontObj["Str"] ; fontObj["str"] = AHK compatible string to set all options

}

SelectOptsFont() {
    Global
	fontObj := Object("name", StrSplit(g_GUI.OptsFont, ",")[1])
	fontObj := FontSelect(fontObj,OptsHwnd)
	If (!fontObj)
		return

    Gui, Setting:Font,,
    Gui, Setting:Font, % fontObj["str"], % fontObj["name"]
    GuiControl, Setting:Font, g_OptsFont
    GuiControl, Setting:, g_OptsFont, % fontObj["name"] "," fontObj["Str"]
}

SelectSBFont() {
    Global
	fontObj := Object("name", StrSplit(g_GUI.SBFont, ",")[1])
	fontObj := FontSelect(fontObj,OptsHwnd)
	If (!fontObj)
		return

    Gui, Setting:Font,,
    Gui, Setting:Font, % fontObj["str"], % fontObj["name"]
    GuiControl, Setting:Font, g_SBFont
    GuiControl, Setting:, g_SBFont, % fontObj["name"] "," fontObj["Str"]
}

SelectCtrlColor() {
    Global
    custColorObj := Array(g_GUI.CtrlColor,g_GUI.WinColor,0xFF0000)
	color := ColorSelect(g_GUI.CtrlColor,OptsHwnd,custColorObj,"full")            ; hwnd and custColorObj are optional

    GuiControl, Setting:, g_CtrlColor, % color
    GuiControl, Setting:+c%color%, g_CtrlColor
}

SelectWinColor() {
    Global
    custColorObj := Array(g_GUI.CtrlColor,g_GUI.WinColor,0xFF0000)
	color := ColorSelect(g_GUI.WinColor,OptsHwnd,custColorObj,"full")            ; hwnd and custColorObj are optional

    GuiControl, Setting:, g_WinColor, % color
    GuiControl, Setting:+c%color%, g_WinColor
}

SettingButtonOK() {
    SAVECONFIG()
    Reload()
}

SettingGuiEscape() {
    SettingGuiClose()
}

SettingButtonCancel() {
    SettingGuiClose()
}

SettingButtonHelp() {
    Run, https://github.com/zhugecaomao/ALTRun/wiki
}

SettingGuiClose() {
    Hotkey, % g_HOTKEY.GlobalHotkey1, On, UseErrorLevel
    Hotkey, % g_HOTKEY.GlobalHotkey2, On, UseErrorLevel
    Gui, Setting:Destroy
}

LoadConfig(Arg) {                                                       ; 加载主配置文件
    g_LOG.Debug("Loading configuration...Arg=" Arg)

    if (Arg = "config" or Arg = "initialize" or Arg = "all") {
        for key, value in g_CONFIG                                      ; Read [Config] to Object
            g_CONFIG[key] := IniRead(g_SECTION.CONFIG, key, value)

        for key, value in g_HOTKEY                                      ; Read [Hotkey] section
            g_HOTKEY[key] := IniRead(g_SECTION.HOTKEY, key, value)

        for key, value in g_GUI                                         ; Read [GUI] section
            g_GUI[key]    := IniRead(g_SECTION.GUI, key, value)

        g_GUI.ListX     := g_GUI.WinX - 24
        g_GUI.ListY     := g_GUI.WinY - 76
        g_RUNTIME.RegEx := g_CONFIG.MatchBeginning ? "imS)^" : "imS)"

        OffsetDate := A_Now
        EnvAdd, OffsetDate, -30, Days

        Loop, Parse, % IniRead(g_SECTION.USAGE), `n
        {
            Date  := StrSplit(A_LoopField, "=")[1]
            Count := StrSplit(A_LoopField, "=")[2]

            if (Date <= SubStr(OffsetDate, 1, 8)) {                     ; Clean up usage record before 30 days (YYYYMMDD format)
                IniDelete, % g_RUNTIME.Ini, % g_SECTION.USAGE, %Date%
                Continue
            }

            g_USAGE[Date] := Count
            g_RUNTIME.Max := Max(g_RUNTIME.Max, Count)
        }

        Loop, 30
        {
            EnvAdd, OffsetDate, +1, Days
            Date := SubStr(OffsetDate, 1, 8)
            g_USAGE[Date] := g_USAGE.HasKey(Date) ? g_USAGE[Date] : 0
        }
    }

    if (Arg = "commands" or Arg = "initialize" or Arg = "all") {        ; Built-in command initialize
        DFTCMDSEC := IniRead(g_SECTION.DFTCMD)
        if (DFTCMDSEC = "") {
            IniWrite,
            (Ltrim
            ; Please make sure ALTRun is not running before modifying this file.
            ; Built-in commands, high priority, recommended to maintain as it is
            ; App will auto generate [DefaultCommand] section while it is empty
            ;
            Func | Help | ALTRun Help & About (F1)=99
            Func | Options | ALTRun Options Preference Settings (F2)=99
            Func | Reload | ALTRun Reload=99
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
            File | `%Temp`%\ALTRun.log | ALTRun Log File=99
            Dir | A_ScriptDir | ALTRun Program Dir=99
            Dir | A_Startup | Current User Startup Dir=99
            Dir | A_StartupCommon | All User Startup Dir=99
            Dir | A_ProgramsCommon | Windowns Search.Index.Cortana Dir=99
            Dir | A_Desktop=99
            Dir | `%AppData`%\Microsoft\Windows\SendTo | Windows SendTo Dir=99
            Dir | `%OneDrive`% | OneDrive=99
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
            ), % g_RUNTIME.Ini, % g_SECTION.DFTCMD
            DFTCMDSEC := IniRead(g_SECTION.DFTCMD)
        }

        USERCMDSEC := IniRead(g_SECTION.USERCMD)
        if (USERCMDSEC = "") {
            IniWrite,
            (Ltrim
            ; User-Defined commands, high priority, modify as desired
            ; Format: Command Type | Command | Description=Rank
            ; Command type: File, Dir, CMD, URL
            ;
            Dir | `%AppData`%\Microsoft\Windows\SendTo | Windows SendTo Dir=9
            Dir | `%OneDrive`% | OneDrive=9
            Dir | A_ScriptDir | ALTRun Program Dir=9
            Cmd | cmd.exe /k ipconfig | Check IP Address=9
            Cmd | explorer /select,C:\Program Files | Open C: and locate to Program Files=9
            Cmd | Control TimeDate.cpl | Date and Time=9
            Cmd | ::{20D04FE0-3AEA-1069-A2D8-08002B30309D} | This PC=9
            URL | www.google.com | Google=9
            File | C:\Apps\TotalCMD64\TOTALCMD64.exe=9
            ), % g_RUNTIME.Ini, % g_SECTION.USERCMD
            USERCMDSEC := IniRead(g_SECTION.USERCMD)
        }

        INDEXSEC := IniRead(g_SECTION.INDEX)                            ; Read whole section of Index database
        if (INDEXSEC = "") {
            MsgBox, 4161, % g_RUNTIME.WinName, % (g_CONFIG.Chinese ? "索引数据库为空, 请点击`n`n'确定'重新建立索引`n`n'取消'退出程序`n`n(请确保程序目录有写入权限)" 
                : "Index database is empty, please click`n`n'OK' to rebuild the index`n`n'Cancel' to exit the program`n`n(Please ensure the program directory is writable)")
            IfMsgBox, Cancel
                Exit()
            Reindex()
        }
        Return DFTCMDSEC "`n" USERCMDSEC "`n" INDEXSEC
    }
    Return
}

IniRead(Section, Key := "", Default := "") {
    IniRead, Value, % g_RUNTIME.Ini, % Section, % Key, % Default
    Return Value
}

IniWrite(Section, Key := "", Value := "") {
    IniWrite, % Value, % g_RUNTIME.Ini, % Section, % Key
}

SAVECONFIG() {
    Gui, Setting:Submit

    For key, description in g_CHKLV
        g_%key% := (A_Index = LV_GetNext(A_Index-1, "C")) ? 1 : 0       ; for Options - General page - Check Listview

    For key, value in g_CONFIG
        IniWrite(g_SECTION.CONFIG, key, g_%key%)                        ; For all ini file - [Config] g_ 变量从控件v变量和上一步Check Listview取得

    For key, value in g_GUI
        IniWrite(g_SECTION.GUI, key, g_%key%)

    For key, value in g_HOTKEY
        IniWrite(g_SECTION.HOTKEY, key, g_%key%)

    Return g_LOG.Debug("Saving config...")
}

;===================================================
; Some Built-in Functions
;===================================================
AhkRun() {
    Run, % g_RUNTIME.Arg,, UseErrorLevel
}

TurnMonitorOff() {                                                      ; 关闭显示器:
    SendMessage, 0x112, 0xF170, 2,, Program Manager                     ; 0x112 is WM_SYSCOMMAND, 0xF170 is SC_MONITORPOWER, 使用 -1 代替 2 来打开显示器, 使用 1 代替 2 来激活显示器的节能模式.
}

EmptyRecycle() {
    MsgBox, 4, % g_RUNTIME.WinName, Do you really want to empty the Recycle Bin?
    IfMsgBox Yes
    {
        FileRecycleEmpty,
    }
}

MuteVolume() {
    SoundSet, MUTE
}

Google() {
    word := g_RUNTIME.Arg = "" ? clipboard : g_RUNTIME.Arg
    Run, https://www.google.com/search?q=%word%&newwindow=1
}

Bing() {
    word := g_RUNTIME.Arg = "" ? clipboard : g_RUNTIME.Arg
    Run, https://cn.bing.com/search?q=%word%
}

Everything() {
    Run, % g_CONFIG.Everything " -s """ g_RUNTIME.Arg """",, UseErrorLevel
    if ErrorLevel
        MsgBox, % "Everything software not found.`n`nPlease check ALTRun setting and Everything program file."
}

SetLanguage() {                                                         ; Max string length that can pass to the array initially is 8192
    ENG     := {}
    CHN     := {}

    ENG.1   := "Options"                                                ; 1~9 Reserved
    ENG.7   := "OK"
    ENG.8   := "Cancel"
    ENG.9   := "Help"
    ENG.10  := "No.|Type|Command|Description"                           ; 10~49 Main GUI
    ENG.11  := "Run"
    ENG.12  := "Options"
    ENG.13  := "Type anything here to search..."
    ENG.50  := "Tip | F1 | Help & About`nTip | F2 | Options and settings`nTip | F3 | Edit current command`nTip | F4 | User-defined commands`nTip | ALT+SPACE / ALT+R | Activate ALTRun`nTip | ALT+SPACE / ESC / LOSE FOCUS | Deactivate ALTRun`nTip | ENTER / ALT+NO. | Run selected command`nTip | ARROW UP or DOWN | Select previous / next command`nTip | CTRL+D | Locate cmd's dir with File Manager" ; Initial tips
    ENG.51  := "Tips: "
    ENG.52  := "It's better to activate ALTRun by hotkey (ALT + Space)" ; 50~99 Tips
    ENG.53  := "Smart Rank - Auto adjusts command priority (rank) based on frequency of use."
    ENG.54  := "Arrow Up / Down = Move to previous / next command"
    ENG.55  := "Esc = Clear input / close window"
    ENG.56  := "Enter = Run current command"
    ENG.57  := "Alt + No. = Run specific command"
    ENG.58  := "Start with + = New Command"
    ENG.59  := "F3 = Edit current command"
    ENG.60  := "F2 = Options setting"
    ENG.61  := "Ctrl+I = Reindex file search database"
    ENG.62  := "F1 = ALTRun Help & About"
    ENG.63  := "ALT + Space = Show / Hide Window"
    ENG.64  := "Ctrl+Q = Reload ALTRun"
    ENG.65  := "Ctrl + No. = Select specific command"
    ENG.66  := "Alt + F4 = Exit"
    ENG.67  := "Ctrl+D = Open current command's dir with File Manager"
    ENG.68  := "F4 = Edit user-defined commands (.ini) directly"
    ENG.69  := "Start with space = Search file by Everything"
    ENG.70  := "Ctrl+'+' = Increase rank of current command"
    ENG.71  := "Ctrl+'-' = Decrease rank of current command"
    ENG.100 := "General||GUI|Hotkey|Index|Listary|Plugins|Usage|About"  ; 100~149 Options window (General - Check Listview)
    ENG.101 := "Launch on Windows startup"
    ENG.102 := "Enable SendTo - Create commands conveniently using Windows SendTo"
    ENG.103 := "Enable ALTRun shortcut in the Windows Start menu"
    ENG.104 := "Show tray icon in the system taskbar"
    ENG.105 := "Close window on losing focus"
    ENG.106 := "Always stay on top"
    ENG.107 := "Show window caption"
    ENG.108 := "Use Windows XP Theme instead of Classic Theme"
    ENG.109 := "Press [ESC] to clear input, press again to close window (Untick: Close directly)"
    ENG.110 := "Keep last input and matching result on close"
    ENG.111 := "Show command icon in the result list"
    ENG.112 := "SendToGetLnk - Retrieve .lnk target on SendTo"
    ENG.113 := "Save commands execution history"
    ENG.114 := "Save application log"
    ENG.115 := "Match full path on search"
    ENG.116 := "Show Grid - Provides boundary lines between list's rows and columns"
    ENG.117 := "Show Header - Show list's header (top row contains column titles)"
    ENG.118 := "Show Serial Number in command list"
    ENG.119 := "Show border line around the command list"
    ENG.120 := "Smart Rank - Auto adjust command priority (rank) based on use frequency"
    ENG.121 := "Smart Match - Fuzzy and Smart matching and filtering result"
    ENG.122 := "Match beginning of the string (Untick: Match from any position)"
    ENG.123 := "Show hints/tips in the bottom status bar"
    ENG.124 := "Show RunCount - Show command executed times in the status bar"
    ENG.125 := "Show status bar at the bottom of the window"
    ENG.126 := "Show [Run] button on main window"
    ENG.127 := "Show [Options] button on main window"
    ENG.128 := "Double Buffer - Paints via double-buffering, reduces flicker (WinXP+)"
    ENG.129 := "Enable express structure calculation"
    ENG.130 := "Shorten Path - Show file/folder/app name only instead of full path in result"
    ENG.131 := "Set language to Chinese Simplified (简体中文)"
    ENG.132 := "Match Chinese Pinyin first characters"
    ENG.150 := "File Manager"                                           ; 150~159 Options window (Other than Check Listview)
    ENG.151 := "Everything"
    ENG.152 := "History length"
    ENG.160 := "Index"                                                  ; 160~169 Index
    ENG.161 := "Index location"
    ENG.162 := "Index file type"
    ENG.163 := "Index exclude"
    ENG.164 := "Index depth"
    ENG.170 := "GUI"                                                    ; 170~189 GUI
    ENG.171 := "Search result number"
    ENG.172 := "Width of each column"
    ENG.173 := "Font (Main GUI)"
    ENG.174 := "Font (Options)"
    ENG.175 := "Font (Status Bar)"
    ENG.176 := "Window size (W x H)"
    ENG.177 := "Cmd list size (W x H)"
    ENG.178 := "Control color"
    ENG.179 := "Background color"
    ENG.180 := "Background picture"
    ENG.181 := "Transparency"
    ENG.182 := "Select font"
    ENG.183 := "Select color"
    ENG.190 := "Hotkey"                                                 ; 190~209 Hotkey
    ENG.191 := "Activate Hotkey (Global)"
    ENG.192 := "Primary Hotkey"
    ENG.193 := "Secondary Hotkey"
    ENG.194 := "Two hotkeys can be set simultaneously"
    ENG.195 := "Reset hotkey"
    ENG.200 := "Actions and Hotkeys (Non-Global)"
    ENG.201 := "Hotkey"
    ENG.202 := "Trigger action"
    ENG.203 := "Hotkey"
    ENG.204 := "Hotkey 2"
    ENG.206 := "Hotkey 3"
    ENG.210 := "Listary"                                                ; 210~219 Listary
    ENG.211 := "Dir Quick-Switch"
    ENG.212 := "File manager id"
    ENG.213 := "Open/Save dialog id"
    ENG.214 := "Exclude windows id"
    ENG.215 := "Hotkey for Switch Open/Save dialog path to"
    ENG.216 := "Total Commander's dir"
    ENG.217 := "Windows Explorer's dir"
    ENG.218 := "Auto switch dir on open/save dialog"
    ENG.220 := "Plugins"                                                ; 220~299 Plugins
    ENG.221 := "Auto-date at end of text"
    ENG.222 := "Apply to window id"
    ENG.223 := "Hotkey"
    ENG.224 := "Date format"
    ENG.225 := "Auto-date before file extension"
    ENG.229 := "Conditional action"
    ENG.230 := "If window id contains"
    ENG.231 := "Hotkey (Editable)"
    ENG.232 := "Trigger action"
    ENG.300 := "Show"                                                   ; 300+ TrayMenu
    ENG.301 := "Options`tF2"
    ENG.302 := "ReIndex`tCtrl+I"
    ENG.303 := "Usage"
    ENG.304 := "About`tF1"
    ENG.305 := "Script Info"
    ENG.307 := "Reload`tCtrl+Q"
    ENG.308 := "Exit`tAlt+F4"
    ENG.309 := "Update"
    ENG.400 := "Run`tEnter"                                             ; 400+ LV_ContextMenu (Right-click)
    ENG.401 := "Locate`tCtrl+D"
    ENG.402 := "Copy`tCtrl+C"
    ENG.403 := "New`tCtrl+N"
    ENG.404 := "Edit`tF3"
    ENG.405 := "Delete`tDelete"
    ENG.406 := "User Command`tF4"
    ENG.500 := "30 days ago"                                            ; 500+ Usage Status
    ENG.501 := "Now"
    ENG.502 := "Total number of times the command was executed"
    ENG.503 := "Number of times the program was activated today"
    ENG.600 := "About"                                                  ; 600+ About
    ENG.601 := "An effective launcher for Windows by ZhugeCaomao, an <a href=""https://www.autohotkey.com/docs/v1/"">AutoHotkey</a> open-source project. "
        . "It provides a streamlined and efficient way to find anything on your system and launch any application in your way."
        . "`n`nSetting file:`n" g_RUNTIME.Ini "`n`nProgram file:`n" A_ScriptFullPath
        . "`n`nCheck for Updates"
        . "`n<a href=""https://github.com/zhugecaomao/ALTRun/releases"">https://github.com/zhugecaomao/ALTRun/releases</a>"
        . "`n`nSource code at GitHub"
        . "`n<a href=""https://github.com/zhugecaomao/ALTRun"">https://github.com/zhugecaomao/ALTRun</a>"
        . "`n`nSee Help and Wiki page for more details"
        . "`n<a href=""https://github.com/zhugecaomao/ALTRun/wiki"">https://github.com/zhugecaomao/ALTRun/wiki</a>"
    
    ENG.700 := "Commander Manager"                                      ; 700+ Commander Manager
    ENG.701 := "Command"
    ENG.702 := "Command type"
    ENG.703 := "Command line"
    ENG.704 := "ShortCut/Description"
    ENG.705 := "Command Section"
    ENG.706 := "Command Rank"
    ENG.800 := "Do you really want to delete the following command from section" ; 800+ Msgbox
    ENG.801 := "?"
    ENG.802 := "Command has been deleted successfully!"

    CHN.1   := "配置"                                                   ; 1~9 Reserved
    CHN.7   := "确定"
    CHN.8   := "取消"
    CHN.9   := "帮助"
    CHN.10  := "序号|类型|命令|描述"                                     ; 10~49 Main GUI
    CHN.11  := "运行"
    CHN.12  := "配置"
    CHN.13  := "在此输入搜索内容..."
    CHN.50  := "提示 | F1 | 帮助&关于`n提示 | F2 | 配置选项`n提示 | F3 | 编辑当前命令`n提示 | F4 | 用户定义命令`n提示 | ALT+空格 / ALT+R | 激活 ALTRun`n提示 | 热键 / ESC / 失去焦点 | 关闭 ALTRun`n提示 | 回车 / ALT+序号 | 运行命令`n提示 | 上下箭头键 | 选择上一个或下一个命令`n提示 | CTRL+D | 使用文件管理器定位命令所在目录"
    CHN.51  := "提示: "                                                 ; 50~99 Tips
    CHN.52  := "推荐使用热键激活 (ALT + 空格)"
    CHN.53  := "智能排序 - 根据使用频率自动调整命令优先级 (排序)"
    CHN.54  := "上/下箭头 = 上/下一个命令"
    CHN.55  := "Esc = 清除输入 / 关闭窗口"
    CHN.56  := "回车 = 运行当前命令"
    CHN.57  := "Alt + 序号 = 运行指定的命令"
    CHN.58  := "以 + 开头 = 新建命令"
    CHN.59  := "F3 = 直接编辑当前命令 (.ini)"
    CHN.60  := "F2 = 配置选项设置"
    CHN.61  := "Ctrl+I = 重建文件搜索数据库"
    CHN.62  := "F1 = ALTRun 帮助&关于"
    CHN.63  := "ALT + 空格 = 显示 / 隐藏窗口"
    CHN.64  := "Ctrl+Q = 重新加载 ALTRun"
    CHN.65  := "Ctrl + 序号 = 选择指定的命令"
    CHN.66  := "Alt + F4 = 退出"
    CHN.67  := "Ctrl+D = 使用文件管理器定位当前命令所在目录"
    CHN.68  := "F4 = 直接编辑用户定义命令 (.ini)"
    CHN.69  := "以空格开头 = 使用 Everything 搜索文件"
    CHN.70  := "Ctrl+'+' = 增加当前命令的优先级"
    CHN.71  := "Ctrl+'-' = 减少当前命令的优先级"
    CHN.100 := "常规||界面|热键|索引|Listary|插件|状态统计|关于"           ; 100~149 Options window (General - Check Listview)
    CHN.101 := "随系统自动启动"
    CHN.102 := "添加到“发送到”菜单"
    CHN.103 := "添加到“开始”菜单"
    CHN.104 := "显示托盘图标 (系统任务栏中)"
    CHN.105 := "失去焦点时关闭窗口"
    CHN.106 := "窗口置顶"
    CHN.107 := "显示窗口标题栏"
    CHN.108 := "使用 Windows XP 主题"
    CHN.109 := "按下 [ESC] 清除输入, 再次按下关闭窗口 (取消勾选: 直接关闭窗口)"
    CHN.110 := "保留最近一次输入和匹配结果"
    CHN.111 := "显示命令图标"
    CHN.112 := "使用“发送到”时, 追溯 .lnk 目标文件"
    CHN.113 := "保存历史记录"
    CHN.114 := "保存运行日志"
    CHN.115 := "搜索时匹配完整路径"
    CHN.116 := "显示网格 - 在列表的行和列之间提供边界线"
    CHN.117 := "显示标题 - 显示列表的标题 (顶部行包含列标题)"
    CHN.118 := "显示命令列表序号"
    CHN.119 := "显示命令列表边框线"
    CHN.120 := "智能排序 - 根据使用频率自动调整命令优先级 (排序)"
    CHN.121 := "智能匹配 - 模糊和智能匹配和过滤结果"
    CHN.122 := "搜索时匹配字符串开头 (取消勾选: 匹配任意位置)"
    CHN.123 := "显示提示信息 - 在底部状态栏中显示技巧提示信息"
    CHN.124 := "显示运行次数 - 在底部状态栏中显示命令执行次数"
    CHN.125 := "显示状态栏 (窗口底部)"
    CHN.126 := "显示主窗口 [运行] 按钮"
    CHN.127 := "显示主窗口 [选项] 按钮"
    CHN.128 := "双缓冲绘图, 改善窗口闪烁 (Win XP+)"
    CHN.129 := "启用快速结构计算"
    CHN.130 := "简化路径 - 仅显示文件/文件夹/应用程序名称, 而不是完整路径"
    CHN.131 := "设置语言为简体中文 (Simplified Chinese)"
    CHN.132 := "搜索时匹配拼音首字母"
    CHN.150 := "文件管理器"                                              ; 150~159 Options window (Other than Check Listview)
    CHN.151 := "Everything"
    CHN.152 := "历史命令数量"
    CHN.160 := "索引"                                                   ; 160~169 Index
    CHN.161 := "索引位置"
    CHN.162 := "索引文件类型"
    CHN.163 := "索引排除项"
    CHN.164 := "索引目录深度"
    CHN.170 := "界面"                                                   ; 170~189 GUI
    CHN.171 := "搜索结果数量"
    CHN.172 := "每列宽度"
    CHN.173 := "字体 (主界面)"
    CHN.174 := "字体 (选项页)"
    CHN.175 := "字体 (状态栏)"
    CHN.176 := "主窗口尺寸 (宽 x 高)"
    CHN.177 := "命令列表尺寸 (宽 x 高)"
    CHN.178 := "控件颜色"
    CHN.179 := "背景颜色"
    CHN.180 := "背景图片"
    CHN.181 := "透明度"
    CHN.182 := "选择字体"
    CHN.183 := "选择颜色"
    CHN.190 := "热键"                                                   ; 190~209 Hotkey
    CHN.191 := "激活热键 (全局)"
    CHN.192 := "主热键"
    CHN.193 := "辅热键"
    CHN.194 := "可以同时设置两个热键"
    CHN.195 := "重置激活热键"
    CHN.200 := "快捷操作和热键 (非全局)"
    CHN.201 := "快捷键"
    CHN.202 := "触发操作"
    CHN.203 := "热键 1"
    CHN.204 := "热键 2"
    CHN.206 := "热键 3"
    CHN.210 := "Listary"                                                ; 210~219 Listary
    CHN.211 := "目录快速切换"
    CHN.212 := "文件管理器 ID"
    CHN.213 := "打开/保存对话框 ID"
    CHN.214 := "排除窗口 ID"
    CHN.215 := "切换打开/保存对话框路径的热键"
    CHN.216 := "Total Commander 路径"
    CHN.217 := "资源管理器路径"
    CHN.218 := "自动切换路径"
    CHN.220 := "插件"                                                   ; 220~299 Plugins
    CHN.221 := "文本末尾自动添加日期"
    CHN.222 := "应用到窗口 ID"
    CHN.223 := "热键"
    CHN.224 := "日期格式"
    CHN.225 := "扩展名前自动添加日期"
    CHN.229 := "条件触发快捷操作"
    CHN.230 := "如果窗口 ID 包含"
    CHN.231 := "热键 (可编辑)"
    CHN.232 := "触发操作"
    CHN.300 := "显示"                                                   ; 300+ 托盘菜单
    CHN.301 := "配置选项`tF2"
    CHN.302 := "重建索引`tCtrl+I"
    CHN.303 := "状态统计"
    CHN.304 := "关于`tF1"
    CHN.305 := "脚本信息"
    CHN.307 := "重新加载`tCtrl+Q"
    CHN.308 := "退出`tAlt+F4"
    CHN.309 := "检查更新"
    CHN.400 := "运行命令`tEnter"                                        ; 400+ 列表右键菜单
    CHN.401 := "定位命令`tCtrl+D"
    CHN.402 := "复制命令`tCtrl+C"
    CHN.403 := "新建命令`tCtrl+N"
    CHN.404 := "编辑命令`tF3"
    CHN.405 := "删除命令`tDelete"
    CHN.406 := "用户命令`tF4"
    CHN.500 := "30天前"                                                 ; 500+ 状态统计
    CHN.501 := "当前"
    CHN.502 := "运行过的命令总次数"
    CHN.503 := "今天激活程序的次数"
    CHN.600 := "关于"                                                   ; 600+ 关于
    CHN.601 := "ALTRun 是由诸葛草帽开发的一款高效 Windows 启动器，是一款基于 <a href=""https://www.autohotkey.com/docs/v1/"">AutoHotkey</a> 的开源项目。 "
        . "它提供了一种简洁高效的方式，让你能够快速查找系统中的任何内容，并以自己的方式启动任意应用程序。"
        . "`n`n配置文件`n" g_RUNTIME.Ini "`n`n程序文件`n" A_ScriptFullPath
        . "`n`n版本更新"
        . "`n<a href=""https://github.com/zhugecaomao/ALTRun/releases"">https://github.com/zhugecaomao/ALTRun/releases</a>"
        . "`n`n源代码开源在 GitHub"
        . "`n<a href=""https://github.com/zhugecaomao/ALTRun"">https://github.com/zhugecaomao/ALTRun</a>"
        . "`n`n有关更多详细信息，请参阅帮助和 Wiki 页面"
        . "`n<a href=""https://github.com/zhugecaomao/ALTRun/wiki"">https://github.com/zhugecaomao/ALTRun/wiki</a>"

    CHN.700 := "命令管理器"                                              ; 700+ 命令管理器
    CHN.701 := "命令"
    CHN.702 := "命令类型"
    CHN.703 := "命令行"
    CHN.704 := "快捷项/描述 (可搜索)"
    CHN.705 := "储存节段"
    CHN.706 := "命令权重"
    CHN.800 := "您确定要从命令节段"                                      ; 800+ 消息内容
    CHN.801 := "中删除以下命令吗?"
    CHN.802 := "命令已成功删除!"

    Global g_LNG := g_CONFIG.Chinese ? CHN : ENG
}

;===================================================
; Eval - Calculate a math expression
; Support +, -, *, /, ^ (or **), and ()
; Return 0 if input contains illegal characters
;===================================================
Eval(expression) {
    ; 移除所有空格
    expression := StrReplace(expression, " ")
    
    ; 检查非法字符（只允许数字、运算符、括号、小数点）
    if (!RegExMatch(expression, "^[\d+\-*/^().]*$"))
        return 0

    ; 递归处理括号
    while RegExMatch(expression, "\(([^()]*)\)", match) {
        result := EvalSimple(match1)  ; 计算括号内的内容
        expression := StrReplace(expression, match, result)
    }

    ; 计算最终无括号表达式
    Return EvalSimple(expression)
}

EvalSimple(expression) {            ; 计算不含括号的简单数学表达式
    ; 处理幂运算符 ^
    while RegExMatch(expression, "(-?\d+(\.\d+)?)([\^])(-?\d+(\.\d+)?)", match) {
        base := match1, exponent := match4
        result := base ** exponent  ; 执行幂运算
        expression := StrReplace(expression, match, result)
    }

    ; 支持 ** 作为幂运算符替代
    while RegExMatch(expression, "(-?\d+(\.\d+)?)(\*\*)(-?\d+(\.\d+)?)", match) {
        base := match1, exponent := match4
        result := base ** exponent
        expression := StrReplace(expression, match, result)
    }

    ; 处理乘除法运算
    while RegExMatch(expression, "(-?\d+(\.\d+)?)([*/])(-?\d+(\.\d+)?)", match) {
        operand1 := match1, operator := match3, operand2 := match4
        result := (operator = "*") ? operand1 * operand2 : operand1 / operand2
        expression := StrReplace(expression, match, result)
    }

    ; 处理加减法运算
    while RegExMatch(expression, "(-?\d+(\.\d+)?)([+\-])(-?\d+(\.\d+)?)", match) {
        operand1 := match1, operator := match3, operand2 := match4
        result := (operator = "+") ? operand1 + operand2 : operand1 - operand2
        expression := StrReplace(expression, match, result)
    }

    ; 返回最终结果
    Return expression
}

Class Logger
{
    __New(filename) {
        this.filename := filename
    }

    Debug(Msg) {
        if (g_CONFIG.SaveLog)
            FileAppend, % "[" A_Now "] " Msg "`n", % this.filename
    }
}

; originally posted by maestrith 
; https://autohotkey.com/board/topic/94083-ahk-11-font-and-color-dialogs/

; to initialize fontObject object (not required):
; ============================================
; fontObject := Object("name","Tahoma","size",14,"color",0xFF0000,"strike",1,"underline",1,"italic",1,"bold",1)

; ==================================================================
; fntName		= name of var to store selected font
; fontObject	= name of var to store fontObject
; hwnd			= parent gui hwnd for modal, leave blank for not modal
; effects		= allow selection of underline / strike out / italic
; ==================================================================
; fontObject output:
;
;	fontObject["str"]	= string to use with AutoHotkey to set GUI values - see examples
;	fontObject["hwnd"]	= handle of the font object to use with SendMessage - see examples
; ==================================================================
FontSelect(fontObject:="",hwnd:=0,effects:=1) {
	fontObject := (fontObject="") ? Object() : fontObject
	VarSetCapacity(logfont,60)
	uintVal := DllCall("GetDC","uint",0)
	LogPixels := DllCall("GetDeviceCaps","uint",uintVal,"uint",90)
	Effects := 0x041 + (Effects ? 0x100 : 0)
	
	fntName := fontObject.HasKey("name") ? fontObject["name"] : ""
	fontBold := fontObject.HasKey("bold") ? fontObject["bold"] : 0
	fontBold := fontBold ? 700 : 400
	fontItalic := fontObject.HasKey("italic") ? fontObject["italic"] : 0
	fontUnderline := fontObject.HasKey("underline") ? fontObject["underline"] : 0
	fontStrikeout := fontObject.HasKey("strike") ? fontObject["strike"] : 0
	fontSize := fontObject.HasKey("size") ? fontObject["size"] : 10
	fontSize := fontSize ? Floor(fontSize*LogPixels/72) : 16
	c := fontObject.HasKey("color") ? fontObject["color"] : 0
	
	c1 := Format("0x{:02X}",(c&255)<<16)	; convert RGB colors to BGR for input
	c2 := Format("0x{:02X}",c&65280)
	c3 := Format("0x{:02X}",c>>16)
	fontColor := Format("0x{:06X}",c1|c2|c3)
	
	fontval := Object(16,fontBold,20,fontItalic,21,fontUnderline,22,fontStrikeout,0,fontSize)
	
	for a,b in fontval
		NumPut(b,logfont,a)
	
	cap:=VarSetCapacity(choosefont,A_PtrSize=8?103:60,0)
	NumPut(hwnd,choosefont,A_PtrSize)
	offset1 := (A_PtrSize = 8) ? 24 : 12
	offset2 := (A_PtrSize = 8) ? 36 : 20
	offset3 := (A_PtrSize = 4) ? 6 * A_PtrSize : 5 * A_PtrSize
	
	fontArray := Array([cap,0,"Uint"],[&logfont,offset1,"Uptr"],[effects,offset2,"Uint"],[fontColor,offset3,"Uint"])
	
	for index,value in fontArray
		NumPut(value[1],choosefont,value[2],value[3])
	
	if (A_PtrSize=8) {
		strput(fntName,&logfont+28)
		r := DllCall("comdlg32\ChooseFont","uptr",&CHOOSEFONT,"cdecl")
		fntName := strget(&logfont+28)
	} else {
		strput(fntName,&logfont+28,32,"utf-8")
		r := DllCall("comdlg32\ChooseFontA","uptr",&CHOOSEFONT,"cdecl")
		fntName := strget(&logfont+28,32,"utf-8")
	}
	
	if !r
		return false
	
	fontObj := Object("bold",16,"italic",20,"underline",21,"strike",22)
	for a,b in fontObj
		fontObject[a] := NumGet(logfont,b,"UChar")
	
	fontObject["bold"] := (fontObject["bold"] < 188) ? 0 : 1
	
	c := NumGet(choosefont,A_PtrSize=4?6*A_PtrSize:5*A_PtrSize) ; convert from BGR to RBG for output
	c1 := Format("0x{:02X}",(c&255)<<16)
	c2 := Format("0x{:02X}",c&65280)
	c3 := Format("0x{:02X}",c>>16)
	c := Format("0x{:06X}",c1|c2|c3)
	fontObject["color"] := c
	
	fontObject["size"] := NumGet(choosefont,A_PtrSize=8?32:16,"UInt")//10
	fontHwnd := DllCall("CreateFontIndirect","uptr",&logfont) ; last param "cdecl"
	fontObject["name"] := fntName
	
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
		Loop Parse, fullStr, |
			If (A_LoopField) 
				str .= A_LoopField " "
		fontObject["str"] := "norm " Trim(str)
		
		return fontObject
	}
}

; =============================================================================================
; Color			= Start color
; hwnd			= Parent window
; custColorObj	= Use for input to init custom colors, or output to save custom colors, or both.
;                 ... custColorObj can be Array() or Object().
; disp			= full / basic ... full displays custom colors panel, basic does not
; =============================================================================================
; All params are optional.  With no hwnd dialog will show at top left of screen.  User must
; parse output custColorObj and decide how to save custom colors... no more automatic ini file.
; =============================================================================================

ColorSelect(Color := 0, hwnd := 0, ByRef custColorObj := "",disp:="full") {
	disp := (disp = "basic" ? 0x1 : 0x3)
	
	c1 := Format("0x{:02X}",(Color&255)<<16)	; convert RGB colors to BGR for input
	c2 := Format("0x{:02X}",Color&65280)		; init start Color
	c3 := Format("0x{:02X}",Color>>16)
	Color := Format("0x{:06X}",c1|c2|c3)
	
	VarSetCapacity(CUSTOM, 16 * A_PtrSize,0) ; init custom colors obj
	size := VarSetCapacity(CHOOSECOLOR, 9 * A_PtrSize,0) ; init dialog
	
	If (IsObject(custColorObj)) {
		Loop 16 {
			If (custColorObj.HasKey(A_Index)) {
				col := custColorObj[A_Index]
				c4 := Format("0x{:02X}",(col&255)<<16)	; convert RGB colors to BGR for input
				c5 := Format("0x{:02X}",col&65280)		; 
				c6 := Format("0x{:02X}",col>>16)
				custCol := Format("0x{:06X}",c4|c5|c6)
				NumPut(custCol, CUSTOM, (A_Index-1) * 4, "UInt")
			}
		}
	}
	
	NumPut(size, CHOOSECOLOR, 0, "UInt")
	NumPut(hwnd, CHOOSECOLOR, A_PtrSize, "UPtr")
	NumPut(Color, CHOOSECOLOR, 3 * A_PtrSize, "UInt")
	NumPut(disp, CHOOSECOLOR, 5 * A_PtrSize, "UInt") ; flags? - original = 3 (0x1 and 0x2)
	NumPut(&CUSTOM, CHOOSECOLOR, 4 * A_PtrSize, "UPtr")
	
	ret := DllCall("comdlg32\ChooseColor", "UPtr", &CHOOSECOLOR, "UInt")
	
	if !ret
		Return
	
	custColorObj := Array()
	Loop 16 {
		newCustCol := NumGet(custom, (A_Index-1) * 4, "UInt")
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

GetFirstChar(str) {
    static array := [ [-20319,-20284,"A"], [-20283,-19776,"B"], [-19775,-19219,"C"], [-19218,-18711,"D"], [-18710,-18527,"E"], [-18526,-18240,"F"], [-18239,-17923,"G"], [-17922,-17418,"H"], [-17417,-16475,"J"], [-16474,-16213,"K"], [-16212,-15641,"L"], [-15640,-15166,"M"], [-15165,-14923,"N"], [-14922,-14915,"O"], [-14914,-14631,"P"], [-14630,-14150,"Q"], [-14149,-14091,"R"], [-14090,-13319,"S"], [-13318,-12839,"T"], [-12838,-12557,"W"], [-12556,-11848,"X"], [-11847,-11056,"Y"], [-11055,-10247,"Z"] ]
    out          := ""
    ; 如果不包含中文字符，则直接返回原字符
    if !RegExMatch(str, "[^\x{00}-\x{ff}]")
        Return str
    Loop, Parse, str
    {
        if ( Asc(A_LoopField) >= 0x2E80 and Asc(A_LoopField) <= 0x9FFF )
        {
            VarSetCapacity(var, 2)
            StrPut(A_LoopField, &var, "CP936")
            nGBKCode := (NumGet(var, 0, "UChar") << 8) + NumGet(var, 1, "UChar") - 65536
            For i, a in array
                if nGBKCode between % a.1 and % a.2
                {
                    out .= a.3
                    Break
                }
        }
        else
            out .= A_LoopField
    }
    Return out
}