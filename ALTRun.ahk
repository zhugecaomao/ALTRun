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
            ,HistoryLen     : 15
            ,DoubleBuffer   : 1
            ,AutoSwitchDir  : 0
            ,Editor         : "Notepad.exe"
            ,FileMgr        : "Explorer.exe"
            ,IndexDir       : "A_ProgramsCommon,A_StartMenu,C:\Other\Index\Location"
            ,IndexType      : "*.lnk,*.exe"
            ,IndexExclude   : "Uninstall *"
            ,Everything     : "C:\Apps\Everything.exe"
            ,DialogWin      : "ahk_class #32770"
            ,FileMgrID      : "ahk_class CabinetWClass, ahk_class TTOTAL_CMD"
            ,ExcludeWin     : "ahk_class SysListView32, ahk_exe Explorer.exe, AutoCAD"
            ,Chinese        : (A_Language = "7804" or A_Language = "0004" or A_Language = "0804" or A_Language = "1004") ? 1 : 0}
, g_HOTKEY  := {Hotkey1     : "^o"
            ,Trigger1       : "Options"
            ,Hotkey2        : ""
            ,Trigger2       : "---"
            ,Hotkey3        : ""
            ,Trigger3       : "---"
            ,CondTitle      : "ahk_exe RAPTW.exe"
            ,CondHotkey     : "~Mbutton"
            ,CondAction     : "PTTools"
            ,GlobalHotkey1  : "!Space"
            ,GlobalHotkey2  : "!R"
            ,TotalCMDDir    : "^g"
            ,ExplorerDir    : "^e"
            ,AutoDateAtEnd  : "ahk_class TCmtEditForm,ahk_class Notepad4" ; TC File Comment 对话框, Notepad4
            ,AutoDateAEHKey : "^d"
            ,AutoDateBefExt : "ahk_class TTOTAL_CMD,ahk_class CabinetWClass,ahk_class Progman,ahk_class WorkerW,ahk_class TSTDTREEDLG,ahk_class #32770,ahk_class TCOMBOINPUT" ; TC 文件列表重命名, Windows 资源管理器文件列表重命名, Windows 桌面文件重命名 (WinXP to Win10, Win11), TC 新建其他格式文件如txt, rtf, docx..., 资源管理器 文件保存对话框, TC F7 创建新文件夹对话框
            ,AutoDateBEHKey : "^d"}
, g_GUI     := {ListRows    : 9
            ,ColWidth       : "36,40,300,AutoHdr"
            ,FontName       : "Microsoft YaHei"
            ,FontSize       : 9
            ,FontColor      : "Default"
            ,WinWidth       : 660
            ,WinHeight      : 300
            ,CtrlColor      : "Default"
            ,WinColor       : "Silver"
            ,Background     : "DEFAULT"
            ,Transparency   : 230}
, g_RUNTIME := {Ini         : A_ScriptDir "\" A_ComputerName ".ini"     ; 程序运行需要的临时全局变量, 不需要用户参与修改, 不读写入ini
            ,WinName        : "ALTRun - Ver 2025.02.20"
            ,BGPic          : ""
            ,WinHide        : ""
            ,UseDisplay     : 0
            ,UseFallback    : 0
            ,ActiveCommand  : ""
            ,Input          : ""
            ,Arg            : ""                                        ; Arg: 用来调用管道的完整参数
            ,UsageToday     : 0
            ,UsageCountMax  : 0
            ,AppStartDate   : A_YYYY A_MM A_DD
            ,OneDrive       : EnvGet("OneDrive")                        ; OneDrive var due to #NoEnv
            ,OneDriveConsumer:EnvGet("OneDriveConsumer")
            ,OneDriveCommercial:EnvGet("OneDriveCommercial")
            ,LV_ContextMenu : []
            ,TrayMenu       : []
            ,FuncList       : ""}

g_LOG.Debug("///// ALTRun is starting /////")
LoadConfig("initialize")                                                ; Load ini config, IniWrite will create it if not exist
SetLanguage()                                                           ; Set language

; For key, value in g_RUNTIME
;   OutputDebug, % key " = " g_RUNTIME[key]

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
    ,ShortenPath    : g_LNG.130 ,Chinese        : g_LNG.131}

;===================================================
; Create ContextMenu and TrayMenu
;===================================================
g_RUNTIME.LV_ContextMenu := [g_LNG.400 ",LVRunCommand,imageres.dll,-100"
    ,g_LNG.401 ",OpenContainer,imageres.dll,-3"
    ,g_LNG.402 ",LVCopyCommand,imageres.dll,-5314",""
    ,g_LNG.403 ",CmdMgr,imageres.dll,-2"
    ,g_LNG.404 ",EditCommand,imageres.dll,-5306"
    ,g_LNG.405 ",UserCommand,imageres.dll,-88"]
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
LV_H         := g_GUI.WinHeight - 43 - 3 * g_GUI.FontSize
LV_W         := g_GUI.WinWidth - 24
Input_W      := LV_W - g_CONFIG.ShowBtnRun * 90 - g_CONFIG.ShowBtnOpt * 90
Enter_W      := g_CONFIG.ShowBtnRun * 80
Enter_X      := g_CONFIG.ShowBtnRun * 10
Options_W    := g_CONFIG.ShowBtnOpt * 80
Options_X    := g_CONFIG.ShowBtnOpt * 10

Gui, Main:Color, % g_GUI.WinColor, % g_GUI.CtrlColor
Gui, Main:Font, % "c" g_GUI.FontColor " s" g_GUI.FontSize, % g_GUI.FontName
Gui, % "Main:" (g_CONFIG.AlwaysOnTop ? "+AlwaysOnTop" : "-AlwaysOnTop")
Gui, % "Main:" (g_CONFIG.ShowCaption ? "+Caption" : "-Caption")
Gui, % "Main:" (g_CONFIG.XPthemeBg ? "+Theme" : "-Theme")
Gui, Main:+HwndMainGuiHwnd
Gui, Main:Default ; Set default GUI before any ListView / statusbar update
Gui, Main:Add, Edit, x12 W%Input_W% -WantReturn vMyInput gOnSearchInput, % g_LNG.13
Gui, Main:Add, Button, % "x+"Enter_X " yp W" Enter_W " hp Default gRunCurrentCommand Hidden" !g_CONFIG.ShowBtnRun, % g_LNG.11
Gui, Main:Add, Button, % "x+"Options_X " yp W" Options_W " hp gOptions Hidden" !g_CONFIG.ShowBtnOpt, % g_LNG.12
Gui, Main:Add, ListView, % "x12 ys+35 W" LV_W " H" LV_H " vMyListView AltSubmit gOnClickListview -Multi" (g_CONFIG.DoubleBuffer ? " +LV0x10000" : "") (g_CONFIG.ShowHdr ? "" : " -Hdr") (g_CONFIG.ShowGrid ? " Grid" : "") (g_CONFIG.ShowBorder ? "" : " -E0x200"), % g_LNG.10 ; LV0x10000 Paints via double-buffering, which reduces flicker
Gui, Main:Add, Picture, x0 y0 0x4000000, % g_RUNTIME.BGPic
Gui, Main:Add, StatusBar, % "gOnClickStatusBar Hidden" !g_CONFIG.ShowStatusBar,

Loop, 4 {
    LV_ModifyCol(A_Index, StrSplit(g_GUI.ColWidth, ",")[A_Index])
}

SB_SetParts(g_GUI.WinWidth - 90 * g_CONFIG.ShowRunCount)
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

g_LOG.Debug("Resolving command line args=" A_Args[1] " " A_Args[2])     ; Command line args, Args are %1% %2% or A_Args[1] A_Args[2]
if (A_Args[1] = "-Startup")
    g_RUNTIME.WinHide := " Hide"

if (A_Args[1] = "-SendTo") {
    g_RUNTIME.WinHide := " Hide"
    CmdMgr(A_Args[2])
}

Gui, Main:Show, % "w" g_GUI.WinWidth " h" g_GUI.WinHeight " Center " g_RUNTIME.WinHide, % g_RUNTIME.WinName

if (g_GUI.Transparency != "OFF" or g_GUI.Transparency != "255")
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

Loop, 3
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
    Result    := ""
    Order     := 1
    g_MATCHED := {}
    Prefix    := SubStr(command, 1, 1)

    if (Prefix = "+" or Prefix = " " or Prefix = ">") {
        g_RUNTIME.ActiveCommand := g_FALLBACK[InStr("+ >", Prefix)]     ; Corresponding to fallback commands position no. 1, 2 & 3
        g_MATCHED.Push(g_RUNTIME.ActiveCommand)
        Return ListResult(g_RUNTIME.ActiveCommand)
    }

    for index, element in g_COMMANDS
    {
        splitResult := StrSplit(element, " | ")
        _Type := splitResult[1]
        _Path := splitResult[2]
        _Desc := splitResult[3]
        SplitPath, _Path, fileName                                      ; Extra name from _Path (if _Type is Dir and has "." in path, nameNoExt will not get full folder name)

        elementToSearch := g_CONFIG.MatchPath ? element : _Type " " fileName " " _Desc ; search type, file name include extension, and desc

        if FuzzyMatch(elementToSearch, command) {
            g_MATCHED.Push(element)

            if (Order = 1) {
                g_RUNTIME.ActiveCommand := element
                Result .= element
            } else {
                Result .= "`n" element
            }
            Order++
            if (Order > g_GUI.ListRows)
                Break
        }
    }

    if (Result = "") {
        if (EvalResult := Eval(command))
        {
            RebarQty := Ceil((EvalResult-40*2) / 300) + 1
            EvalResultTho := FormatThousand(EvalResult)

            Result := "Eval | " EvalResultTho

            if (g_CONFIG.StruCalc) {
                Result .= "`n | ------------------------------------------------------"
                Result .= "`n | Beam width = " EvalResultTho " mm"
                Result .= "`n | Main bar no. = " RebarQty " (" Round((EvalResult-40*2) / (RebarQty - 1)) " c/c), " RebarQty + 1 " (" Round((EvalResult-40*2) / (RebarQty+1-1)) " c/c), " RebarQty - 1 " (" Round((EvalResult-40*2) / (RebarQty-1-1)) " c/c)"
                Result .= "`n | ------------------------------------------------------"
                Result .= "`n | As = " EvalResultTho " mm2"
                Result .= "`n | Rebar = " Ceil(EvalResult/132.7) "H13 / " Ceil(EvalResult/201.1) "H16 / " Ceil(EvalResult/314.2) "H20 / " Ceil(EvalResult/490.9) "H25 / " Ceil(EvalResult/804.2) "H32"
            }
            Return ListResult(Result, True)
        }

        g_RUNTIME.UseFallback   := true
        g_MATCHED               := g_FALLBACK
        g_RUNTIME.ActiveCommand := g_FALLBACK[1]

        for i, cmd in g_FALLBACK
            Result .= (i = 1) ? g_FALLBACK[i] : "`n" g_FALLBACK[i]
    } else {
        g_RUNTIME.UseFallback := false
    }
    ListResult(Result)
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
        IconIndex   := g_CONFIG.ShowIcon ? GetIconIndex(_Path, _Type) : 0

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
    Path := StrReplace(Path, "%Desktop%", A_Desktop)
    Path := StrReplace(Path, "%OneDrive%", g_RUNTIME.OneDrive)          ; Convert OneDrive to absolute path due to #NoEnv
    Path := StrReplace(Path, "%OneDriveConsumer%", g_RUNTIME.OneDriveConsumer)
    Path := StrReplace(Path, "%OneDriveCommercial%", g_RUNTIME.OneDriveCommercial)
    Return Path
}

RelativePath(Path)                                                      ; Convert path to relative path
{
    Path := StrReplace(Path, A_Temp, "%Temp%")
    Path := StrReplace(Path, A_Desktop, "%Desktop%")
    Path := StrReplace(Path, g_RUNTIME.OneDrive, "%OneDrive%")
    Path := StrReplace(Path, g_RUNTIME.OneDriveConsumer, "%OneDriveConsumer%")
    Path := StrReplace(Path, g_RUNTIME.OneDriveCommercial, "%OneDriveCommercial%")
    Return Path
}

EnvGet(EnvVar) {
    EnvGet, OutputVar, %EnvVar%
    Return OutputVar
}

RunCommand(originCmd)
{
    MainGuiClose()                                                      ; 先关闭窗口,避免出现延迟的感觉
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
        g_HISTORYS.InsertAt(1, originCmd " /Arg=" g_RUNTIME.Arg)        ; Adjust command history

        (g_HISTORYS.Length() > g_CONFIG.HistoryLen) ? g_HISTORYS.Pop()

        IniDelete, % g_RUNTIME.Ini, % g_SECTION.HISTORY
        for index, element in g_HISTORYS
            IniWrite, %element%, % g_RUNTIME.Ini, % g_SECTION.HISTORY, %index% ; Save command history
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

LVCopyCommand() {                                                       ; ListView ContextMenu (right click & its menu) actions
    Gui, Main:Default                                                   ; Use it before any LV update
    focusedRow := LV_GetNext(0, "Focused")                              ; Check focused row, only operate focusd row instead of all selected rows
    if (!focusedRow)                                                    ; Return if no focused row is found
        Return

    g_RUNTIME.ActiveCommand := g_MATCHED[focusedRow]                    ; Get current command from focused row
    LV_GetText(Text, focusedRow, 3) ? (A_Clipboard := Text)         ; Get the text from the focusedRow's 3rd field.
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
    g_LOG.Debug("mock test search ' " chr(chr1) " " chr(chr2) " " chr(chr3) " ' 50 times, use time = " t)
    MsgBox % "Search '" chr(chr1) " " chr(chr2) " " chr(chr3) "' use Time =  " t
}

UserCommand() {
    if (g_CONFIG.Editor = "notepad.exe")
        Run, % g_CONFIG.Editor " " g_RUNTIME.Ini,, UseErrorLevel
    Else
        Run, % g_CONFIG.Editor " /m [" g_SECTION.USERCMD "] """ g_RUNTIME.Ini """",, UseErrorLevel ; /m Match text
}

EditCommand() {
    if (g_CONFIG.Editor = "notepad.exe")
        Run, % g_CONFIG.Editor " " g_RUNTIME.Ini,, UseErrorLevel
    else
        Run, % g_CONFIG.Editor " /m " """" g_RUNTIME.ActiveCommand "=""" " """ g_RUNTIME.Ini """",, UseErrorLevel ; /m Match text, locate to current command, add = at end to filter out [history] commands
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
    Return RegExMatch(Haystack, "imS)" Needle)
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

        IniWrite, %Rank%, % g_RUNTIME.Ini, %A_LoopField%, %originCmd% ; Update new Rank for originCmd

        showRank ? SetStatusBar("Rank for current command : " Rank)
    }
    LoadCommands()                                                      ; New rank will take effect in real-time by LoadCommands again
}

UpdateUsage() {
    if (g_RUNTIME.AppStartDate != A_YYYY . A_MM . A_DD) {
        g_RUNTIME.AppStartDate := A_YYYY . A_MM . A_DD
        g_RUNTIME.UsageToday   := 0
    }
    g_RUNTIME.UsageToday++
    IniWrite, % g_RUNTIME.UsageToday, % g_RUNTIME.Ini, % g_SECTION.USAGE, % A_YYYY . A_MM . A_DD
}

UpdateRunCount() {
    g_CONFIG.RunCount++
    IniWrite, % g_CONFIG.RunCount, % g_RUNTIME.Ini, % g_SECTION.CONFIG, RunCount ; Record run counting
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
    
    IniRead, FALLBACKCMDSEC, % g_RUNTIME.Ini, % g_SECTION.FALLBACK      ; Read FALLBACK section, initialize it if section not exist
    if (FALLBACKCMDSEC = "") {
        IniWrite, 
        (Ltrim
        ; Fallback Commands show when search result is empty
        ; Commands in order, modify as desired
        ; Format: Command Type | Command | Description
        ; Command type: File, Dir, CMD, URL
        ;
        Func | CmdMgr | New Command
        Func | Everything | Search by Everything
        Func | Google | Search Clipboard or Input by Google
        Func | AhkRun | Run Command use AutoHotkey Run
        Func | Bing | Search Clipboard or Input by Bing
        ), % g_RUNTIME.Ini, % g_SECTION.FALLBACK
        IniRead, FALLBACKCMDSEC, % g_RUNTIME.Ini, % g_SECTION.FALLBACK
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
            IniRead, History, % g_RUNTIME.Ini, % g_SECTION.HISTORY, % A_Index, % A_Space
            g_HISTORYS.Push(History)
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

        for extIndex, ext in StrSplit(g_CONFIG.IndexType, ",")
        {
            Loop Files, %searchPath%\%ext%, R
            {
                if (g_CONFIG.IndexExclude != "" && RegExMatch(A_LoopFileLongPath, g_CONFIG.IndexExclude))
                    continue                                            ; Skip this file and move on to the next loop.

                IniWrite, 1, % g_RUNTIME.Ini, % g_SECTION.INDEX, File | %A_LoopFileLongPath% ; Assign initial rank to 1
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
    Options("Help", 8)                                                  ; Open Options window 8th tab (help tab)
}

Usage() {
    Options("Usage", 7)
}

Update() {
    Run, https://github.com/zhugecaomao/ALTRun/releases
}

Listary() {                                                             ; Listary Dir QuickSwitch Function (快速更换保存/打开对话框路径)
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

CmdMgr(Path := "") {                                                    ; 命令管理窗口
    Global
    g_LOG.Debug("Starting Command Manager... Args=" Path)

    SplitPath Path, _Desc,, fileExt,,                                   ; Extra name from _Path (if _Type is dir and has "." in path, nameNoExt will not get full folder name) 
    
    if InStr(FileExist(Path), "D")                                      ; True only if the file exists and is a directory.
        _Type := 2                                                      ; It is a normal folder
    else                                                                ; From command "New Command" or GUI context menu "New Command"
        _Desc := g_RUNTIME.Arg
    
    if (fileExt = "lnk" && g_CONFIG.SendToGetLnk) {
        FileGetShortcut, %Path%, Path,, fileArg, _Desc
        Path .= " " fileArg
    }

    Gui, CmdMgr:New
    Gui, CmdMgr:Font, S9 Norm, Microsoft Yahei
    Gui, CmdMgr:Add, GroupBox, w600 h260, % g_LNG.701
    Gui, CmdMgr:Add, Text, x25 yp+30, % g_LNG.702
    Gui, CmdMgr:Add, DropDownList, x145 yp-5 w130 v_Type Choose%_Type%, File||Dir|Cmd|URL
    Gui, CmdMgr:Add, Text, x300 yp+5, Command Section
    Gui, CmdMgr:Add, DropDownList, x420 yp-5 w130 v_Section, UserCommand||DefaultCommand|Index|FallbackCommand
    Gui, CmdMgr:Add, Text, x25 yp+60, % g_LNG.703
    Gui, CmdMgr:Add, Edit, x145 yp-5 w405 -WantReturn v_Path, % RelativePath(Path)
    Gui, CmdMgr:Add, Button, x560 yp w30 hp gSelectCmdPath, ...
    Gui, CmdMgr:Add, Text, x25 yp+80, % g_LNG.704
    Gui, CmdMgr:Add, Edit, x145 yp-5 w405 -WantReturn v_Desc, %_Desc%
    Gui, CmdMgr:Add, Text, x25 yp+60, 命令权重
    Gui, CmdMgr:Add, ComboBox, x145 yp-5 w405, 1||2|3|4|5|6|7|8|9|10
    Gui, CmdMgr:Add, Button, Default x420 w90 gCmdMgrButtonOK, % g_LNG.8
    Gui, CmdMgr:Add, Button, x521 yp w90 gCmdMgrButtonCancel, % g_LNG.9
    Gui, CmdMgr:Show, AutoSize, % g_LNG.700
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
    _Desc := _Desc ? "| " _Desc : _Desc

    if (_Path = "")
    {
        MsgBox,64, Command Manager, Command Path is empty`, please input correct command path!
        Return
    } else {
        IniWrite, 1, % g_RUNTIME.Ini, % g_SECTION.USERCMD, %_Type% | %_Path% %_Desc% ; initial rank = 1
        if (!ErrorLevel)
            MsgBox,64, Command Manager, Command added successfully!`n`n%_Path%
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

RenameWithDate()                                                        ; 针对所有设定好的程序 按Ctrl+D自动在文件(夹)名之后添加日期
{
    ControlGetFocus, CurrCtrl, A                                        ; 获取当前激活的窗口中的聚焦的控件名称
    if (InStr(CurrCtrl, "Edit") or InStr(CurrCtrl, "Scintilla"))        ; 如果当前激活的控件为Edit类或者Scintilla1(Notepad2),则Ctrl+D功能生效
        NameAddDate("FileListMangr", CurrCtrl)
    Else
        SendInput ^D
    Return
}

LineEndAddDate()                                                        ; 针对TC File Comment对话框　按Ctrl+D自动在备注文字之后添加日期
{
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
        NameWithDate := nameNoExt " - " CurrentDate "." fileExt
    }
    else
    {
        NameWithDate := EditCtrlText " - " CurrentDate
    }
    ControlClick, %CurrCtrl%, A
    ControlSetText, %CurrCtrl%, %NameWithDate%, A
    SendInput {Blind}{End}
    g_LOG.Debug(WinName ", RenameWithDate=" NameWithDate)
}

FormatThousand(Number)                                                  ; Function to add thousand separator
{
    Return RegExReplace(Number, "\G\d+?(?=(\d{3})+(?:\D|$))", "$0" ",")
}

Options(Arg := "", ActTab := 1)                                         ; Options settings, 1st parameter is to avoid menu like [Option`tF2] disturb ActTab
{
    Global                                                              ; Assume-global mode
    Gui, Setting:New, +OwnDialogs +AlwaysOnTop, % g_LNG.1               ; +OwnerMain: (omit due to lug options window)
    Gui, Setting:Font, S9 Norm, Microsoft Yahei
    Gui, Setting:Add, Tab3, vCurrTab Choose%ActTab%, % g_LNG.100
    Gui, Setting:Tab, 1 ; CONFIG Tab
    Gui, Setting:Add, ListView, w500 h300 Checked -Multi AltSubmit -Hdr vOptListView, % g_LNG.1

    For key, description in g_CHKLV
        LV_Add("Check" g_CONFIG[key], description)

    LV_ModifyCol(1, "AutoHdr")

    Gui, Setting:Add, Text, x24 yp+320, % g_LNG.150
    Gui, Setting:Add, ComboBox, x130 yp-5 w394 Sort vg_Editor, % g_CONFIG.Editor "||Notepad.exe|C:\Apps\Notepad4.exe"
    Gui, Setting:Add, Text, x24 yp+40, % g_LNG.151
    Gui, Setting:Add, ComboBox, x130 yp-5 w394 Sort vg_Everything, % g_CONFIG.Everything "||C:\Apps\Everything.exe"
    Gui, Setting:Add, Text, x24 yp+40, % g_LNG.152
    Gui, Setting:Add, ComboBox, x130 yp-5 w394 Sort vg_FileMgr, % g_CONFIG.FileMgr "||Explorer.exe|C:\Apps\TotalCMD.exe /O /T /S"

    Gui, Setting:Tab, 2 ; INDEX Tab
    Gui, Setting:Add, GroupBox, w500 h130, % g_LNG.160
    Gui, Setting:Add, Text, x33 yp+25, % g_LNG.161
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_IndexDir, % g_CONFIG.IndexDir "||A_ProgramsCommon,A_StartMenu"
    Gui, Setting:Add, Text, x33 yp+40, % g_LNG.162
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_IndexType, % g_CONFIG.IndexType "||*.lnk,*.exe"
    Gui, Setting:Add, Text, x33 yp+40, % g_LNG.163
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_IndexExclude, % g_CONFIG.IndexExclude "||Uninstall *"
    Gui, Setting:Add, GroupBox, x24 yp+45 w500 h270, % g_LNG.164
    Gui, Setting:Add, Text, x33 yp+25, % g_LNG.165
    Gui, Setting:Add, DropDownList, x183 yp-5 w330 Sort vg_HistoryLen, % StrReplace("0|10|15|20|25|30|50|90|", g_CONFIG.HistoryLen, g_CONFIG.HistoryLen . "|",, 1)

    Gui, Setting:Tab, 3 ; GUI Tab
    Gui, Setting:Add, GroupBox, w500 h420, % g_LNG.170
    Gui, Setting:Add, Text, x33 yp+25 , % g_LNG.171
    Gui, Setting:Add, DropDownList, x183 yp-5 w330 vg_ListRows, % StrReplace("3|4|5|6|7|8|9|", g_GUI.ListRows, g_GUI.ListRows . "|",, 1) ; ListRows limit <= 9, not using Choose%g_ListRows% as list start from 3
    Gui, Setting:Add, Text, x33 yp+40, % g_LNG.172
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 Sort vg_ColWidth, % g_GUI.ColWidth "||33,46,460,AutoHdr|40,45,430,340|40,0,475,340|23,0,460,AutoHdr"
    Gui, Setting:Add, Text, x33 yp+40, % g_LNG.173
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 Sort vg_FontName, % g_GUI.FontName "||Default|Segoe UI Semibold|Microsoft Yahei"
    Gui, Setting:Add, Text, x33 yp+40, % g_LNG.174
    Gui, Setting:Add, ComboBox, x183 yp-5 r1 w330 vg_FontSize, % g_GUI.FontSize "||8|9|10|11|12"
    Gui, Setting:Add, Text, x33 yp+40, % g_LNG.175
    Gui, Setting:Add, ComboBox, x183 yp-5 r1 w330 vg_FontColor, % g_GUI.FontColor "||Default|Black|Blue|DCDCDC|000000"
    Gui, Setting:Add, Text, x33 yp+40, % g_LNG.176
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_WinWidth, % g_GUI.WinWidth "||920"
    Gui, Setting:Add, Text, x33 yp+40, % g_LNG.177
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_WinHeight, % g_GUI.WinHeight "||313"
    Gui, Setting:Add, Text, x33 yp+40, % g_LNG.178
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_CtrlColor, % g_GUI.CtrlColor "||Default|White|Blue|202020|FFFFFF"
    Gui, Setting:Add, Text, x33 yp+40, % g_LNG.179
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_WinColor, % g_GUI.WinColor "||Default|White|Blue|202020|FFFFFF"
    Gui, Setting:Add, Text, x33 yp+40, % g_LNG.180
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_Background, % g_GUI.Background "||NO PICTURE|DEFAULT|C:\Path\BackgroundPicture.jpg"
    Gui, Setting:Add, Text, x33 yp+40, % g_LNG.181
    Gui, Setting:Add, DropDownList, x183 yp-5 w330 vg_Transparency, % StrReplace("OFF|50|75|100|125|150|175|200|210|220|230|240|250|255|", g_GUI.Transparency, g_GUI.Transparency . "|",, 1)

    Gui, Setting:Tab, 4 ; Hotkey Tab
    Gui, Setting:Add, GroupBox, w500 h115, % g_LNG.191
    Gui, Setting:Add, Text, x33 yp+25 , % g_LNG.192
    Gui, Setting:Add, Hotkey, x285 yp-4 w230 vg_GlobalHotkey1, % g_HOTKEY.GlobalHotkey1
    Gui, Setting:Add, Text, x33 yp+35 , % g_LNG.193
    Gui, Setting:Add, Hotkey, x285 yp-4 w230 vg_GlobalHotkey2,% g_HOTKEY.GlobalHotkey2
    Gui, Setting:Add, Text, x33 yp+35, % g_LNG.194
    Gui, Setting:Add, Link, x285 yp w230 gResetHotkey, % "<a>" g_LNG.195 "</a>"

    Gui, Setting:Add, GroupBox, x24 yp+35 w500 h55, % g_LNG.196
    Gui, Setting:Add, Text, x33 yp+25 , % g_LNG.197
    Gui, Setting:Add, Text, x183 yp , % g_LNG.198
    Gui, Setting:Add, Text, x285 yp, % g_LNG.199
    Gui, Setting:Add, Text, x395 yp, % g_LNG.200

    Gui, Setting:Add, GroupBox, x24 yp+38 w500 h140, % g_LNG.201
    Gui, Setting:Add, Text, x33 yp+30 , % g_LNG.202
    Gui, Setting:Add, Hotkey, x183 yp-5 w80 vg_Hotkey1, % g_HOTKEY.Hotkey1
    Gui, Setting:Add, Text, x285 yp+5, % g_LNG.203
    Gui, Setting:Add, DropDownList, x395 yp-5 w120 Sort vg_Trigger1, % StrReplace("---|" g_RUNTIME.FuncList, g_HOTKEY.Trigger1, g_HOTKEY.Trigger1 . "|",, 1)
    Gui, Setting:Add, Text, x33 yp+40 , % g_LNG.204
    Gui, Setting:Add, Hotkey, x183 yp-5 w80 vg_Hotkey2, % g_HOTKEY.Hotkey2
    Gui, Setting:Add, Text, x285 yp+5, % g_LNG.203
    Gui, Setting:Add, DropDownList, x395 yp-5 w120 Sort vg_Trigger2, % StrReplace("---|" g_RUNTIME.FuncList, g_HOTKEY.Trigger2, g_HOTKEY.Trigger2 . "|",, 1)
    Gui, Setting:Add, Text, x33 yp+40 , % g_LNG.206
    Gui, Setting:Add, Hotkey, x183 yp-5 w80 vg_Hotkey3, % g_HOTKEY.Hotkey3
    Gui, Setting:Add, Text, x285 yp+5, % g_LNG.203
    Gui, Setting:Add, DropDownList, x395 yp-5 w120 Sort vg_Trigger3, % StrReplace("---|" g_RUNTIME.FuncList, g_HOTKEY.Trigger3, g_HOTKEY.Trigger3 . "|",, 1)

    Gui, Setting:Tab, 5 ; LISTARTY TAB
    Gui, Setting:Add, GroupBox, w500 h145, % g_LNG.211
    Gui, Setting:Add, Text, x33 yp+30 , % g_LNG.212
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 Sort vg_FileMgrID, % g_CONFIG.FileMgrID "||ahk_class CabinetWClass|ahk_class CabinetWClass, ahk_class TTOTAL_CMD"
    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.213
    Gui, Setting:Add, Combobox, x183 yp-5 w330 Sort vg_DialogWin, % g_CONFIG.DialogWin "||ahk_class #32770"
    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.214
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 Sort vg_ExcludeWin, % g_CONFIG.ExcludeWin "||ahk_class SysListView32|ahk_class SysListView32, ahk_exe Explorer.exe|ahk_class SysListView32, ahk_exe Explorer.exe, ahk_exe Totalcmd64.exe, AutoCAD LT Alert"
    Gui, Setting:Add, GroupBox, x24 yp+50 w500 h145, % g_LNG.215
    Gui, Setting:Add, Text, x33 yp+30, % g_LNG.216
    Gui, Setting:Add, Hotkey, x183 yp-5 w330 vg_TotalCMDDir, % g_HOTKEY.TotalCMDDir
    Gui, Setting:Add, Text, x33 yp+45, % g_LNG.217
    Gui, Setting:Add, Hotkey, x183 yp-5 w330 vg_ExplorerDir, % g_HOTKEY.ExplorerDir
    Gui, Setting:Add, CheckBox, % "x33 yp+45 vg_AutoSwitchDir checked" g_CONFIG.AutoSwitchDir, % g_LNG.218

    Gui, Setting:Tab, 6 ; Plugins TAB
    Gui, Setting:Add, GroupBox, w500 h110, % g_LNG.221
    Gui, Setting:Add, Text, x33 yp+30, % g_LNG.222
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_AutoDateAtEnd, % g_HOTKEY.AutoDateAtEnd "||ahk_class TCmtEditForm,ahk_class Notepad4|"
    Gui, Setting:Add, Text, x33 yp+45 , % g_LNG.223
    Gui, Setting:Add, Hotkey, x183 yp-5 w80 vg_AutoDateAEHKey, % g_HOTKEY.AutoDateAEHKey
    Gui, Setting:Add, Text, x285 yp+5, % g_LNG.224
    Gui, Setting:Add, DropDownList, x395 yp-5 w120, - dd.MM.yyyy||

    Gui, Setting:Add, GroupBox, x24 y+30 w500 h110, % g_LNG.225
    Gui, Setting:Add, Text, x33 yp+30, % g_LNG.222
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_AutoDateBefExt, % g_HOTKEY.AutoDateBefExt "||ahk_class TTOTAL_CMD,ahk_class CabinetWClass,ahk_class Progman,ahk_class WorkerW,ahk_class #32770|"
    Gui, Setting:Add, Text, x33 yp+45 , % g_LNG.223
    Gui, Setting:Add, Hotkey, x183 yp-5 w80 vg_AutoDateBEHKey, % g_HOTKEY.AutoDateBEHKey
    Gui, Setting:Add, Text, x285 yp+5, % g_LNG.224
    Gui, Setting:Add, DropDownList, x395 yp-5 w120, - dd.MM.yyyy||

    Gui, Setting:Add, GroupBox, x24 y+30 w500 h110, % g_LNG.229
    Gui, Setting:Add, Text, x33 yp+30 , % g_LNG.230
    Gui, Setting:Add, ComboBox, x183 yp-5 w330 vg_CondTitle, % g_HOTKEY.CondTitle "||"
    Gui, Setting:Add, Text, x33 yp+45 , % g_LNG.231
    Gui, Setting:Add, ComboBox, x183 yp-5 w80 vg_CondHotkey, % g_HOTKEY.CondHotkey "||"
    Gui, Setting:Add, Text, x285 yp+5, % g_LNG.232
    Gui, Setting:Add, DropDownList, x395 yp-5 w120 Sort vg_CondAction, % StrReplace("---|" g_RUNTIME.FuncList, g_HOTKEY.CondAction, g_HOTKEY.CondAction . "|",, 1)

    Gui, Setting:Tab, 7 ; USAGE TAB
    Gui, Setting:Add, GroupBox, x66 y80 w445 h300,

    OffsetDate := A_Now
    EnvAdd, OffsetDate, -30, Days                                       ; Subtract 30 days

    Loop, 30
    {
        EnvAdd, OffsetDate, +1, Days
        FormatTime, OffsetDate, %OffsetDate%, yyyyMMdd

        IniRead, OutputVar, % g_RUNTIME.Ini, % g_SECTION.USAGE, % OffsetDate, 0
        Gui, Setting:Add, Progress, % "c94DD88 BackgroundF9F9F9 Vertical y96 w14 h280 xm+" 50+A_Index*14 " Range0-" g_RUNTIME.UsageCountMax+10, %OutputVar%
    }
    Gui, Setting:Add, Text, x24 yp-5 cGray, % g_RUNTIME.UsageCountMax+10
    Gui, Setting:Add, Text, x24 yp+140 cGray, % Round(g_RUNTIME.UsageCountMax/2)+5
    Gui, Setting:Add, Text, x24 yp+140 cGray, 0
    Gui, Setting:Add, Text, x66 yp+15 cGray, % g_LNG.500
    Gui, Setting:Add, Text, x476 yp cGray, % g_LNG.501
    Gui, Setting:Add, Text, x66 yp+33, % g_LNG.502
    Gui, Setting:Add, Edit, x410 yp-5 w100 Disabled Right vg_RunCount, % g_CONFIG.RunCount
    Gui, Setting:Add, Text, x66 yp+35, % g_LNG.503
    Gui, Setting:Add, Edit, x410 yp-5 w100 Disabled Right, % g_RUNTIME.UsageToday

    Gui, Setting:Tab, 8 ; ABOUT TAB
    Gui, Setting:Add, Picture, x33 y+20 w48 h-1 Icon-100, imageres.dll
    Gui, Setting:Font, S10,
    Gui, Setting:Add, Text, x96 yp+5, % g_RUNTIME.WinName
    Gui, Setting:Font, S9,
    Gui, Setting:Add, Link, xp yp+30 w400, % g_LNG.601

    Gui, Setting:Tab                                                    ; 后续添加的控件将不属于前面的选项卡控件
    Gui, Setting:Add, Button, Default x355 w80 gSettingButtonOK, % g_LNG.8
    Gui, Setting:Add, Button, x445 yp w80 gSettingButtonCancel, % g_LNG.9
    Gui, Setting:Show,, % g_LNG.1
    Hotkey, % g_HOTKEY.GlobalHotkey1, Off, UseErrorLevel
    Hotkey, % g_HOTKEY.GlobalHotkey2, Off, UseErrorLevel
    g_LOG.Debug("Loading options window...Arg=" Arg ", ActTab=" ActTab)
}

ResetHotkey() {
    GuiControl, Setting:, g_GlobalHotkey1, !Space
    GuiControl, Setting:, g_GlobalHotkey2, !R
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

SettingGuiClose() {
    Hotkey, % g_HOTKEY.GlobalHotkey1, On, UseErrorLevel
    Hotkey, % g_HOTKEY.GlobalHotkey2, On, UseErrorLevel
    Gui, Setting:Destroy
}

LoadConfig(Arg) {                                                       ; 加载主配置文件
    g_LOG.Debug("Loading configuration...Arg=" Arg)

    if (Arg = "config" or Arg = "initialize" or Arg = "all") {
        For key, value in g_CONFIG                                      ; Read [Config] to Object
        {
            IniRead, tempValue, % g_RUNTIME.Ini, % g_SECTION.CONFIG, %key%, %value%
            g_CONFIG[key] := tempValue
        }

        For key, value in g_HOTKEY                                      ; Read Hotkey section
        {
            IniRead, tempValue, % g_RUNTIME.Ini, % g_SECTION.HOTKEY, %key%, %value%
            g_HOTKEY[key] := tempValue
        }

        For key, value in g_GUI                                         ; Read GUI section
        {
            IniRead, tempValue, % g_RUNTIME.Ini, % g_SECTION.GUI, %key%, %value%
            g_GUI[key] := tempValue
        }

        g_RUNTIME.BGPic := (g_GUI.Background = "DEFAULT") ? Extract_BG(A_Temp "\ALTRun.jpg") : g_GUI.Background

        IniRead, tempValue, % g_RUNTIME.Ini, % g_SECTION.USAGE, % A_YYYY . A_MM . A_DD, 0 ; For app usage
        g_RUNTIME.UsageToday := tempValue

        OffsetDate := A_Now
        EnvAdd, OffsetDate, -30, Days                                   ; - 30 days
        FormatTime, OffsetDate, %OffsetDate%, yyyyMMdd
    
        IniRead, USAGE, % g_RUNTIME.Ini, % g_SECTION.USAGE              ; Clean up usage record before 30 days
        Loop, Parse, USAGE, `n
        {
            UsageDate  := StrSplit(A_LoopField, "=")[1]
            UsageCount := StrSplit(A_LoopField, "=")[2]
            
            if (UsageDate <= OffsetDate) {
                IniDelete, % g_RUNTIME.Ini, % g_SECTION.USAGE, %UsageDate%
                Continue
            }

            g_RUNTIME.UsageCountMax := Max(g_RUNTIME.UsageCountMax, UsageCount)
        }
    }

    if (Arg = "commands" or Arg = "initialize" or Arg = "all") {        ; Built-in command initialize
        IniRead, DFTCMDSEC, % g_RUNTIME.Ini, % g_SECTION.DFTCMD
        if (DFTCMDSEC = "") {
            IniWrite,
            (Ltrim
            ; Built-in commands, high priority, recommended to maintain as it is
            ; App will auto generate [DefaultCommnd] section while it is empty
            ;
            Func | Help | ALTRun Help Index (F1)=99
            Func | Options | ALTRun Options Preference Settings (F2)=99
            Func | Reload | ALTRun Reload=99
            Func | CmdMgr | New Command=99
            Func | UserCommand | ALTRun User-defined command (F4)=99
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
            Dir | `%Desktop`%=99
            Dir | `%AppData`%\Microsoft\Windows\SendTo | Windows SendTo Dir=99
            Dir | `%OneDriveConsumer`% | OneDrive Personal Dir=99
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
            Cmd | Control TimeDate.cpl | Date and Time=66
            Cmd | Control AdminTools | Windows Tools=66
            Cmd | Control Desktop | Personalisation=66
            Cmd | Control Inetcpl.cpl,,4 | Internet Properties=66
            Cmd | Control Printers | Devices and Printers=66
            Cmd | Control UserPasswords | User Accounts=66
            ), % g_RUNTIME.Ini, % g_SECTION.DFTCMD
            IniRead, DFTCMDSEC, % g_RUNTIME.Ini, % g_SECTION.DFTCMD
        }

        IniRead, USERCMDSEC, % g_RUNTIME.Ini, % g_SECTION.USERCMD
        if (USERCMDSEC = "") {
            IniWrite,
            (Ltrim
            ; User-Defined commands, high priority, modify as desired
            ; Format: Command Type | Command | Description=Rank
            ; Command type: File, Dir, CMD, URL
            ;
            Dir | `%AppData`%\Microsoft\Windows\SendTo | Windows SendTo Dir=1
            Dir | `%OneDriveConsumer`% | OneDrive Personal Dir=1
            Dir | A_ScriptDir | ALTRun Program Dir=99
            Cmd | cmd.exe /k ipconfig | Check IP Address=1
            Cmd | Control TimeDate.cpl | Date and Time=66
            Cmd | ::{20D04FE0-3AEA-1069-A2D8-08002B30309D} | This PC=66
            URL | www.google.com | Google=1
            File | C:\OneDrive\Apps\TotalCMD64\TOTALCMD64.exe=1
            ), % g_RUNTIME.Ini, % g_SECTION.USERCMD
            IniRead, USERCMDSEC, % g_RUNTIME.Ini, % g_SECTION.USERCMD
        }

        IniRead, INDEXSEC, % g_RUNTIME.Ini, % g_SECTION.INDEX           ; Read whole section of Index database
        if (INDEXSEC = "") {
            MsgBox, 4160, % g_RUNTIME.WinName, % (g_CONFIG.Chinese ? "索引数据库为空，请点击确定重新建立索引?" : "Index database is empty, please click OK to rebuild index.")
            Reindex()
        }
        Return DFTCMDSEC "`n" USERCMDSEC "`n" INDEXSEC
    }
    Return
}

SAVECONFIG() {
    Gui, Setting:Submit

    For key, description in g_CHKLV
        g_%key% := (A_Index = LV_GetNext(A_Index-1, "C")) ? 1 : 0       ; for Options - General page - Check Listview

    For key, value in g_CONFIG
        IniWrite, % g_%key%, % g_RUNTIME.Ini, % g_SECTION.CONFIG, %key% ; For all ini file - [Config] g_ 变量从控件v变量和上一步Check Listview取得

    For key, value in g_GUI
        IniWrite, % g_%key%, % g_RUNTIME.Ini, % g_SECTION.GUI, %key%

    For key, value in g_HOTKEY
        IniWrite, % g_%key%, % g_RUNTIME.Ini, % g_SECTION.HOTKEY, %key%

    Return g_LOG.Debug("Saving config...")
}

;===================================================
; Resources File - Background picture
;===================================================
Extract_BG(_Filename)
{
	Static Out_Data
    VarSetCapacity(TD, 4206 * 2)
    TD :="/9j/4AAQSkZJRgABAQAAAQABAAD/2wEEEAANAA0ADQANAA4ADQAOABAAEAAOABQAFgATABYAFAAeABsAGQAZABsAHgAtACAAIgAgACIAIAAtAEQAKgAyACoAKgAyACoARAA8AEkAOwA3ADsASQA8AGwAVQBLAEsAVQBsAH0AaQBjAGkAfQCXAIcAhwCXAL4AtQC+APkA+QFOEQANAA0ADQANAA4ADQAOABAAEAAOABQAFgATABYAFAAeABsAGQAZABsAHgAtACAAIgAgACIAIAAtAEQAKgAyACoAKgAyACoARAA8AEkAOwA3ADsASQA8AGwAVQBLAEsAVQBsAH0AaQBjAGkAfQCXAIcAhwCXAL4AtQC+APkA+QFO/8IAEQgBXAOYAwEiAAIRAQMRAf/EADAAAQACAwEBAAAAAAAAAAAAAAABBQIDBAcGAQEBAQEAAAAAAAAAAAAAAAAAAQID/9oADAMBAAIQAxAAAACqHTmAAAAAmB2WNF1pc58nRZsRICAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAInAnUs10Wcs6AAAAAAAA8+FAAAAAAAdVlR9KXefJvs2olAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAENSy2W8uveKEAAAAAAAAefCgAAAAAAAOizpN6Xmzi6bNqJAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFYxrI7d3dLEkoAAAAAAAAAHnwoAAAAAAAADdaUu1L7Zwddm1EgIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIVhjsMLbbnNBAAAAAAAAAAAHnwoAAAAAAAAADZZ1GaX+2v67NyJAQAAAAAAAAAAAAAAAAAAAAAAAAAAAjEnVFnLz2spoAAAAAAAAAAAADz4UAAAAAAAAAABnZ1OaX22u67OhjkAgAAAAAAAAAAAAAAAAAAAAAAAACGtZxzuJdfQKEAAAAAAAAAAAAAefCgAAAAAAAAAAAMrGsyL7dWdlz0scgEAAAAAAAAAAAAAAAAAAAAAABWMax1b++WJJQAAAAAAAAAAAAAAPPhQAAAAAAAAAAAAE2NbKX26q77OlhmAgAAAAAAAAAAAAAAAAAAAAhWCTC13bJoIAAAAAAAAAAAAAAAA8+FAAAAAAAAAAAAAAT3cEl7vqLC56mGYCAAAAAAAAAAAAAAAAAADAnXFgui1mc6AAAAAAAAAAAAAAAAAA8+FAAAAAAAAAAAAAAAO3iF9up7FnqnXnUgBAAAAAAAAAAAAAAAEMFYZ3EauoaCAAAAAAAAAAAAAAAAAAPPhQAAAAAAAAAAAAAAADq5ZS830lnZ1zrzJCAAAAAAAAAAAAAArFqJ377CWMiUAAAAAAAAAAAAAAAAAAADz4UAAAAAAAAAAAAAAAAA6OcXfRR2bPbOrZUgBAAAAAAAAAABCsAwtdu6aCAAAAAAAAAAAAAAAAAAAAAPPhQAAAAAAAAAAAAAAAAADdpF100NpZ3Tq2JIQAAAAAAAAAYGWqO6XRbzLQQAAAAAAAAAAAAAAAAAAAAAB58KAAAAAAAAAAAAAAAAAAAbdQuOqhs7O/LRuSQAgAAAAACGtZwyuJdXWSgAAAAAAAAAAAAAAAAAAAAAAAefCgAAAAAAAAAAAAAAAAAAAGeAtuygs2bDLRtrOBAAAAAGLUs7d9lLGRKAAAAAAAAAAAAAAAAAAAAAAAAB58KAAAAAAAAAAAAAAAAAAAAAZ4EtO2gsUs507ayACACBjELjZ7t80EAAAAAAAAAAAAAAAAAAAAAAAAAAefCgAAAAAAAAAAAAAAAAAAAAAGWIsu6gsGbXLm3VmSQYpOqOxdNvlM0EAAAAAAAAAAAAAAAAAAAAAAAAAAAefCgAAAAAAAAAAAAAAAAAAAAAAEwLDvoe6y2y5tiZ6sreNXaTQAAAAAAAAAAAAAAAAAAAAAAAAAAAAHnwoAAAAAAAAAAAAAAAAAAAAAAAAjA7rbn+mI2QiUSAAAAAAAAAAAAAAAAAAAAAAAAAAAAAefCgAAAAAAAAAAAAAAAAAAAAAAERtNX0/Z3QABMxIRIAAAAAAAAAAAAAAAAAAAAAAAAAAAB58KAAAAAAAAAAAAAAAAAAAAAEE4RdHJ9btygAAACZgSiQAAAAAAAAAAAAAAAAAAAAAAAAAADz4UAAAAAAAAAAAAAAAAAAAAIGOP1hhdkAAAAAASCQAAAAAAAAAAAAAAAAAAAAAAAAAAf/EAC8QAAICAQIFAQcEAwEAAAAAAAECAAMRBCExQVBRYBIFIjBAYXGxIIGRocHR4UL/2gAIAQEAAT8A+NTfggMf3iPA0Hh5aM00+jJw9v7LB8/TeRhWP2MR4D4cWnvOwVRknlNPpFr9593/AB0Km8r7rcIj5gMB8KJhaV12XN6UH3PaU6dKV23bmeiU3FNjw/ER9hAYD4QTGYSjTveQTsneJWta+lBgDo1NxQ4O6xHzzitB4MTGeafRl8Pbw5LAABgdIquKHfhK7AQDFMB8EJm7MFUZJ5TT6QV4d93/AKHS6rSh+krsDAHMVoD4CTC0RHub0oPue0o06Ug43Y8T02u1qz9O0rtDAERWgPXyYWEp073nPBO8rrStfSo26fXYazsf2lVoYAgxWgPXSYzTT6Qvh7Nl5DvAAAAP66ijshyP4lVoYZEVoD1smEliFAySeU0+jCYazBbtyHVEco2RKbg4itAesEiFoqva3pQf8lGmSkd25nqyuyHIlNwcf4itAerExmlNFl57JzMrqSpfSo6wrFCCDKbgwitAeqExmE0+kNmHsyF7d4AFAAGw60pZSCJTd6x9YrQHqRMJJIA4zT6P04a3du3broJByJTf6tucVhAenkwtFV7WCoN5RpkpGeLd+vg4OZReG2PH8xXgPTSRGaU02XnbZeZlVSVL6VH/AHwKi/Ox4xHgbpZMZpp9K1uGfZO3eKoUAAYAHglN+cAneI8DQdIJhJJwBuTNPo8Ye3j27eD0X/8Alv5iPAYOi5haKHsb0oN5p9MlIzxfv4TRfg+lj9jEeBpnoZMLCVU2Xttw5mVUpSuFH3Pfwum/0+638xHzA0HQSYzCafStbhnyE/sxVVAAoAA7eG03FDgnaI8DTPz5aFuQmn0fB7ePJe3iFNxQgHh+Ij5EDQH50tFDWMFQZM0+lWrc7v37eJU3FDg8IlgPOK0HzRMZpVS97YXhzMppSlcL+57+KVWlDjlK7ARFaA/MExmlGla73m2T8xVVFCqMAcvFqrTWfpK7AQCDFaA/LEwt24zT6M7Pb+y+M12ms/TtK7QwBBitAfkyYWg9TsFQZJmn0q1e827+N12FGyP4lVoYZBimA/ImM0qqe9sKNuZ7SmlKVwo35nv46jshyJVcHAIitmA/HJhaUaZrj6jskRFRQqjAHj6OyHIlNwcZitAfikxmmn0ecPaPssAx5CrlDkSm4MP8RWgPwy0HqdgqjJPKafSCv33wX/HkisVORKbgw+vaK0B+ATC0rre5sKPue0poSkYHHmfJlJByJTf6x9e0VhAf1ExmlGme45Oyd+8RFrUKowB5QCQciUXerY8YrwH9BMZpp9GWw9vDksAwPKgcEGU3+rY8YrwGZhabuQqjJPKafSCvDPu/48tH0lN+djxivPXER7m9Kj7ntKdOlI23bmfMKtTyYzS6azUbnKp37/aV1pWoVBgeX5hM0HslnxbqAQvFU/3AAoAA2A8vJiq9jBEUsxOwE0HslKMW3YazkOS+YEzT6e7VWeipfueQmi0FOkXb3nPFz5gWmi9n3as53WocW/1KNPVp6xXUuF8vMJmg9ktbi3UAqnJOZiqqKFUAADYDy8mAM7BEUsxOABPZ/slacW34Z+S8l8wMM9kaalNMlwX334nwb//EABoRAQEAAgMAAAAAAAAAAAAAAAFQAGARcID/2gAIAQIBAT8AoLZXFsLZWwtlbK2V3FbK2V6qWyu5c+Bv/8QAHBEBAAMBAQEBAQAAAAAAAAAAAQBAUBEwIBCA/9oACAEDAQE/APtMkKSRxgqJihXcEIFhI3wtpOXQupG2F9JywEDCSsGKlPkDY5AyU9yBlp6hAzeeYaCeIaSfRA1EnP0IGvyB/A//2Q=="
    
    VarSetCapacity(Out_Data, Bytes := 3070, 0)
    DllCall("Crypt32.dll\CryptStringToBinary" "W", "Ptr", &TD, "UInt", 0, "UInt", 1, "Ptr", &Out_Data, "UIntP", Bytes, "Int", 0, "Int", 0, "CDECL Int")
	
    FileExist(_Filename)
		FileDelete, %_Filename%
	
	h := DllCall("CreateFile", "Ptr", &_Filename, "Uint", 0x40000000, "Uint", 0, "UInt", 0, "UInt", 4, "Uint", 0, "UInt", 0)
	, DllCall("WriteFile", "Ptr", h, "Ptr", &Out_Data, "UInt", 3070, "UInt", 0, "UInt", 0)
	, DllCall("CloseHandle", "Ptr", h)
    Return _Filename
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
    ENG := {1:"Options"                                                 ; 1~10 Reserved
        ,8 :"OK"
        ,9 :"Cancel"
        ,10:"No.|Type|Command|Description"                              ; 10~50 Main GUI
        ,11:"Run"
        ,12:"Options"
        ,13:"Type anything here to search..."

        ,50:"Tip | F1 | Help`nTip | F2 | Options and settings`nTip | F3 | Edit current command`nTip | F4 | User-defined commands`nTip | ALT+SPACE / ALT+R | Activative ALTRun`nTip | ALT+SPACE / ESC / LOSE FOCUS | Deactivate ALTRun`nTip | ENTER / ALT+NO. | Run selected command`nTip | ARROW UP or DOWN | Select previous or next command`nTip | CTRL+D | Locate cmd's dir with File Manager" ; Initial tips
        ,51:"Tips: "                                                    ; 50~100 Tips
        ,52:"It's better to activate ALTRun by hotkey (ALT + Space)"
        ,53:"Smart Rank - Atuo adjusts command priority (rank) based on frequency of use."
        ,54:"Arrow Up / Down = Move to previous / next command"
        ,55:"Esc = Clear input / close window"
        ,56:"Enter = Run current command"
        ,57:"Alt + No. = Run specific command"
        ,58:"Start with + = New Command"
        ,59:"F3 = Edit current command"
        ,60:"F2 = Options setting"
        ,61:"Ctrl+I = Reindex file search database"
        ,62:"F1 = ALTRun Help Index"
        ,63:"ALT + Space = Show / Hide Window"
        ,64:"Ctrl+Q = Reload ALTRun"
        ,65:"Ctrl + No. = Select specific command"
        ,66:"Alt + F4 = Exit"
        ,67:"Ctrl+D = Open current command's dir with File Manager"
        ,68:"F4 = Edit user-defined commands (.ini) directly"
        ,69:"Start with space = Search file by Everything"
        ,70:"Ctrl+'+' = Increase rank of current command"
        ,71:"Ctrl+'-' = Decrease rank of current command"

        ,100:"General|Index|GUI|Hotkey|Listary|Plugins|Usage|About"     ; 100~149 Options window (General - Check Listview)
        ,101:"Launch on Windows startup"
        ,102:"Enable SendTo - Create commands conveniently using Windows SendTo"
        ,103:"Enable ALTRun shortcut in the Windows Start menu"
        ,104:"Show tray icon in the system taskbar"
        ,105:"Close window on losing focus"
        ,106:"Always stay on top"
        ,107:"Show Caption - Show window title bar"
        ,108:"XP Theme - Use Windows Theme instead of Classic Theme (WinXP+)"
        ,109:"[ESC] to clear input, press again to close window (Untick: Close directly)"
        ,110:"Keep last input and search result on close"
        ,111:"Show Icon - Show file/folder/app icon in result"
        ,112:"SendToGetLnk - Retrieve .lnk target on SendTo"
        ,113:"Save History - Commands executed with arg"
        ,114:"Save Log - App running and debug information"
        ,115:"Match full path on search"
        ,116:"Show Grid - Provides boundary lines between list's rows and columns"
        ,117:"Show Header - Show list's header (top row contains column titles)"
        ,118:"Show Serial Number in command list"
        ,119:"Show border line around the command list"
        ,120:"Smart Rank - Auto adjust command priority (rank) based on use frequency"
        ,121:"Smart Match - Fuzzy and Smart matching and filtering result"
        ,122:"Match beginning of the string (Untick: Match from any position)"
        ,123:"Show hints/tips in the bottom status bar"
        ,124:"Show RunCount - Show command executed times in the status bar"
        ,125:"Show status bar at the bottom of the window"
        ,126:"Show [Run] button on main window"
        ,127:"Show [Options] button on main window"
        ,128:"Double Buffer - Paints via double-buffering, reduces flicker (WinXP+)"
        ,129:"Enable express structure calculation"
        ,130:"Shorten Path - Show file/folder/app name only instead of full path in result"
        ,131:"Set language to Chinese Simplified (简体中文)"
        ,150:"Text Editor"                                              ; 150~159 Options window (Other than Check Listview)
        ,151:"Everything"
        ,152:"File Manager"
        ,160:"Index"                                                    ; 160~169 Index
        ,161:"Index location"
        ,162:"Index file type"
        ,163:"Index exclude"
        ,164:"Others"
        ,165:"Command history length"
        ,170:"GUI"                                                      ; 170~189 GUI
        ,171:"Search result number"
        ,172:"Width of each column"
        ,173:"Font name"
        ,174:"Font size"
        ,175:"Font color"
        ,176:"Window width"
        ,177:"Window height"
        ,178:"Control color"
        ,179:"Background color"
        ,180:"Background picture"
        ,181:"Transparency (0-255)"
        ,190:"Hotkey"                                                   ; 190~209 Hotkey
        ,191:"Activate"
        ,192:"Primary Hotkey"
        ,193:"Secondary Hotkey"
        ,194:"Two hotkeys can be set simultaneously"
        ,195:"Reset hotkey"
        ,196:"Commands"
        ,197:"Execute command"
        ,198:"Alt + No."
        ,199:"Select command"
        ,200:"Ctrl + No."
        ,201:"Actions and Hotkeys (non-global)"
        ,202:"Hotkey 1"
        ,203:"Trigger action"
        ,204:"Hotkey 2"
        ,206:"Hotkey 3"
        ,210:"Listary"                                                  ; 210~219 Listary
        ,211:"Dir Quick-Switch"
        ,212:"File manager id"
        ,213:"Open/Save dialog id"
        ,214:"Exclude windows id"
        ,215:"Hotkey for Swith open/save dialog path to"
        ,216:"Total Commander's dir"
        ,217:"Windows Explorer's dir"
        ,218:"Auto switch dir on open/save dialog"
        ,220:"Plugins"                                                  ; 220~299 Plugins
        ,221:"Auto-date at end of text"
        ,222:"Apply to window id"
        ,223:"Hotkey"
        ,224:"Date format"
        ,225:"Auto-date before file extension"
        ,229:"Conditional action"
        ,230:"If window id contains"
        ,231:"Hoktey (Editable)"
        ,232:"Trigger action"
        ,300:"Show"                                                     ; 300+ TrayMenu
        ,301:"Options`tF2"
        ,302:"ReIndex`tCtrl+I"
        ,303:"Usage"
        ,304:"Help`tF1"
        ,305:"Script Info"
        ,307:"Reload`tCtrl+Q"
        ,308:"Exit`tAlt+F4"
        ,309:"Update"}

    ENG.400 := "Run`tEnter"                                             ; 400+ LV_ContextMenu (Right-click)
    ENG.401 := "Locate`tCtrl+D"
    ENG.402 := "Copy"
    ENG.403 := "New"
    ENG.404 := "Edit`tF3"
    ENG.405 := "User Command`tF4"

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
    ENG.701 := "New command"
    ENG.702 := "Command type"
    ENG.703 := "Command path"
    ENG.704 := "Description"

    CHN := {1:"配置"                                                    ; 1~10 Reserved
        ,8 :"确定"
        ,9 :"取消"
        ,10:"序号|类型|命令|描述"                                        ; 10~50 Main GUI
        ,11:"运行"
        ,12:"配置"
        ,13:"在此输入搜索内容..."

        ,50:"提示 | F1 | 帮助`n提示 | F2 | 配置选项`n提示 | F3 | 编辑当前命令`n提示 | F4 | 用户定义命令`n提示 | ALT+空格 / ALT+R | 激活 ALTRun`n提示 | 热键 / Esc / 失去焦点 | 关闭 ALTRun`n提示 | 回车 / ALT+序号 | 运行命令`n提示 | 上下箭头键 | 选择上一个或下一个命令`n提示 | CTRL+D | 使用文件管理器定位命令所在目录"
        ,51:"提示: "                                                    ; 50~100 Tips
        ,52:"推荐使用热键激活 (ALT + 空格)"
        ,53:"智能排序 - 根据使用频率自动调整命令优先级 (排序)"
        ,54:"上/下箭头 = 上/下一个命令"
        ,55:"Esc = 清除输入 / 关闭窗口"
        ,56:"回车 = 运行当前命令"
        ,57:"Alt + 序号 = 运行特定命令"
        ,58:"以 + 开头 = 新建命令"
        ,59:"F3 = 直接编辑当前命令 (.ini)"
        ,60:"F2 = 配置选项设置"
        ,61:"Ctrl+I = 重建文件搜索数据库"
        ,62:"F1 = ALTRun 帮助索引"
        ,63:"ALT + 空格 = 显示 / 隐藏窗口"
        ,64:"Ctrl+Q = 重新加载 ALTRun"
        ,65:"Ctrl + 序号 = 选择特定命令"
        ,66:"Alt + F4 = 退出"
        ,67:"Ctrl+D = 使用文件管理器定位当前命令所在目录"
        ,68:"F4 = 直接编辑用户定义命令 (.ini)"
        ,69:"以空格开头 = 使用 Everything 搜索文件"
        ,70:"Ctrl+'+' = 增加当前命令的优先级"
        ,71:"Ctrl+'-' = 减少当前命令的优先级"

        ,100:"常规|索引|界面|热键|Listary|插件|状态统计|关于"              ; 100~149 Options window (General - Check Listview)
        ,101:"随系统自动启动"
        ,102:"添加到“发送到”菜单"
        ,103:"添加到“开始”菜单中"
        ,104:"在系统任务栏中显示托盘图标"
        ,105:"失去焦点时关闭窗口"
        ,106:"窗口置顶"
        ,107:"显示窗口标题栏"
        ,108:"XP 主题 - 使用 Windows 主题 (WinXP+)"
        ,109:"[Esc] 清除输入, 再次按下关闭窗口 (取消勾选: 直接关闭窗口)"
        ,110:"保留上次输入和搜索结果关闭"
        ,111:"显示图标 - 在结果中显示文件/文件夹/应用程序图标"
        ,112:"使用“发送到”时, 追溯 .lnk 目标文件"
        ,113:"保存历史记录 - 使用参数执行的命令"
        ,114:"保存日志 - 应用程序运行和调试信息"
        ,115:"搜索时匹配完整路径"
        ,116:"显示网格 - 在列表的行和列之间提供边界线"
        ,117:"显示标题 - 显示列表的标题 (顶部行包含列标题)"
        ,118:"在命令列表中显示序号"
        ,119:"在命令列表周围显示边框线"
        ,120:"智能排序 - 根据使用频率自动调整命令优先级 (排序)"
        ,121:"智能匹配 - 模糊和智能匹配和过滤结果"
        ,122:"匹配字符串开头 (取消勾选: 匹配字符串任意位置)"
        ,123:"在底部状态栏显示提示/提示"
        ,124:"显示运行次数 - 在状态栏中显示命令执行次数"
        ,125:"在窗口底部显示状态栏"
        ,126:"在主窗口上显示 [运行] 按钮"
        ,127:"在主窗口上显示 [选项] 按钮"
        ,128:"双缓冲绘图, 改善窗口闪烁 (Win XP+)"
        ,129:"启用快速结构计算"
        ,130:"缩短路径 - 仅显示文件/文件夹/应用程序名称, 而不是完整路径"
        ,131:"设置语言为简体中文(Simplified Chinese)"
        ,150:"文本编辑器"                                                ; 150~159 Options window (Other than Check Listview)
        ,151:"Everything"
        ,152:"文件管理器"
        ,160:"索引"                                                     ; 160~169 Index
        ,161:"索引位置"
        ,162:"索引文件类型"
        ,163:"索引排除项"
        ,164:"其他"
        ,165:"历史命令数量"
        ,170:"界面"                                                     ; 170~189 GUI
        ,171:"搜索结果数量"
        ,172:"每列宽度"
        ,173:"字体名称"
        ,174:"字体大小"
        ,175:"字体颜色"
        ,176:"窗口宽度"
        ,177:"窗口高度"
        ,178:"控件颜色"
        ,179:"背景颜色"
        ,180:"背景图片"
        ,181:"透明度 (0-255)"
        ,190:"热键"                                                     ; 190~209 Hotkey
        ,191:"激活"
        ,192:"主热键"
        ,193:"辅热键"
        ,194:"可以同时设置两个热键"
        ,195:"重置热键"
        ,196:"命令"
        ,197:"执行命令"
        ,198:"Alt + 序号"
        ,199:"选择命令"
        ,200:"Ctrl + 序号"
        ,201:"快捷操作和热键 (非全局)"
        ,202:"热键 1"
        ,203:"触发操作"
        ,204:"热键 2"
        ,206:"热键 3"
        ,210:"Listary"                                                  ; 210~219 Listary
        ,211:"目录快速切换"
        ,212:"文件管理器 ID"
        ,213:"打开/保存对话框 ID"
        ,214:"排除窗口 ID"
        ,215:"切换打开/保存对话框路径的热键"
        ,216:"Total Commander 路径"
        ,217:"资源管理器路径"
        ,218:"自动切换路径"
        ,220:"插件"                                                     ; 220~299 Plugins
        ,221:"文本末尾自动添加日期"
        ,222:"应用到窗口 ID"
        ,223:"热键"
        ,224:"日期格式"
        ,225:"扩展名前自动添加日期"
        ,229:"条件触发快捷操作"
        ,230:"如果窗口 ID 包含"
        ,231:"热键 (可编辑)"
        ,232:"触发操作"
        ,300:"显示"                                                     ; 300+ 托盘菜单
        ,301:"配置选项`tF2"
        ,302:"重建索引`tCtrl+I"
        ,303:"状态统计"
        ,304:"帮助`tF1"
        ,305:"脚本信息"
        ,307:"重新加载`tCtrl+Q"
        ,308:"退出`tAlt+F4"
        ,309:"检查更新"}

    CHN.400 := "运行命令`tEnter"                                         ; 400+ 列表右键菜单
    CHN.401 := "定位命令`tCtrl+D"
    CHN.402 := "复制命令"
    CHN.403 := "新建命令"
    CHN.404 := "编辑命令`tF3"
    CHN.405 := "用户命令`tF4"

    CHN.500 := "30天前"                                                 ; 500+ 状态统计
    CHN.501 := "当前"
    CHN.502 := "运行过的命令总次数"
    CHN.503 := "今天激活程序的次数"

    CHN.600 := "关于"                                                   ; 600+ 关于
    CHN.601 := "ALTRun 是由诸葛草帽开发的一款高效 Windows 启动器，是一款基于 <a href=""https://www.autohotkey.com/docs/v1/"">AutoHotkey</a> 的开源项目。 "
        . "它提供了一种简洁高效的方式，让你能够快速查找系统中的任何内容，并以自己的方式启动任意应用程序。"
        . "`n`n设置文件:`n" g_RUNTIME.Ini "`n`n程序文件:`n" A_ScriptFullPath
        . "`n`n检查更新"
        . "`n<a href=""https://github.com/zhugecaomao/ALTRun/releases"">https://github.com/zhugecaomao/ALTRun/releases</a>"
        . "`n`n源代码开源在 GitHub"
        . "`n<a href=""https://github.com/zhugecaomao/ALTRun"">https://github.com/zhugecaomao/ALTRun</a>"
        . "`n`n有关更多详细信息，请参阅帮助和 Wiki 页面"
        . "`n<a href=""https://github.com/zhugecaomao/ALTRun/wiki"">https://github.com/zhugecaomao/ALTRun/wiki</a>"

    CHN.700 := "命令管理器"                                              ; 700+ 命令管理器
    CHN.701 := "新建命令"
    CHN.702 := "命令类型"
    CHN.703 := "命令路径"
    CHN.704 := "命令描述"

    Global g_LNG := g_CONFIG.Chinese ? CHN : ENG
}

;===================================================
; Eval - Calculate a math expression
; Support +, -, *, /, ^ (or **), and ()
; Return 0 if input contains illegal characters
;===================================================
Eval(expression) {
    StringReplace, expression, expression, %A_Space%, , All             ; 移除所有空格

    If !RegExMatch(expression, "^[\d+\-*/^().]*$") {                    ; 检查非法字符（只允许数字、运算符、括号、小数点）
        Return 0
    }

    While RegExMatch(expression, "\(([^()]*)\)", match) {               ; 递归处理括号
        result := EvalSimple(match1) ; 计算括号内的内容
        StringReplace, expression, expression, %match%, %result%, All
    }

    Return EvalSimple(expression)    ; 计算最终无括号表达式
}

EvalSimple(expression) {                                                ; 计算不含括号的简单数学表达式
    While RegExMatch(expression, "(-?\d+(\.\d+)?)([\^])(-?\d+(\.\d+)?)", match) { ; 处理幂运算
        base := match1, exponent := match4
        result := base ** exponent ; 执行幂运算
        StringReplace, expression, expression, %match%, %result%, All
    }

    While RegExMatch(expression, "(-?\d+(\.\d+)?)([\*][\*])(-?\d+(\.\d+)?)", match) { ; 支持 ** 作为幂运算符的替代
        base := match1, exponent := match4
        result := base ** exponent
        StringReplace, expression, expression, %match%, %result%, All
    }

    While RegExMatch(expression, "(-?\d+(\.\d+)?)([*/])(-?\d+(\.\d+)?)", match) { ; 处理乘除运算
        operand1 := match1, operator := match3, operand2 := match4
        result := (operator = "*") ? operand1 * operand2 : operand1 / operand2
        StringReplace, expression, expression, %match%, %result%, All
    }

    While RegExMatch(expression, "(-?\d+(\.\d+)?)([+\-])(-?\d+(\.\d+)?)", match) { ; 处理加减运算
        operand1 := match1, operator := match3, operand2 := match4
        result := (operator = "+") ? operand1 + operand2 : operand1 - operand2
        StringReplace, expression, expression, %match%, %result%, All
    }

    Return expression    ; 返回最终结果
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