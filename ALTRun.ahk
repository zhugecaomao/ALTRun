;==============================================================
; ALTRun - An effective launcher for Windows.
; https://github.com/zhugecaomao/ALTRun
;==============================================================
#Requires AutoHotkey v1.1
#NoEnv
#SingleInstance, Force
#NoTrayIcon
#Persistent
#Warn All, OutputDebug

FileEncoding, UTF-8
SendMode, Input
SetWorkingDir %A_ScriptDir%

Global g_IniFile := A_ScriptDir "\" A_ComputerName ".ini"
, Log            := New Logger(A_Temp "\ALTRun.log")
, SEC_CONFIG     := "Config"
, SEC_GUI        := "Gui"
, SEC_DFTCMD     := "DefaultCommand"
, SEC_USERCMD    := "UserCommand"
, SEC_FALLBACK   := "FallbackCommand"
, SEC_HOTKEY     := "Hotkey"
, SEC_HISTORY    := "History"
, SEC_INDEX      := "Index"
, KEYLIST_CONFIG := "IndexDir,IndexType,IndexExclude,HistoryLen,Editor,FileMgr,Everything,RunCount,AutoSwitchDir,FileMgrID,DialogWin,ExcludeWin"
, KEYS_CONFIG    := {AutoStartup: "Launch on Windows startup", EnableSendTo: "Enable the SendTo menu", InStartMenu: "Enable the Start menu"
                    , ShowTrayIcon: "Show Tray Icon in the system taskbar", HideOnLostFocus: "Close window on losing focus", AlwaysOnTop: "Always stay on top"
                    , EscClearInput: "Press [ESC] to clear input, press again to close window (untick: close directly)", KeepInput: "Keep last input and search result on close"
                    , ShowIcon: "Show Icon of file, folder or apps in the command result list", SendToGetLnk: "Retrieve .lnk target on SendTo"
                    , SaveHistory: "Save command history", Logging: "Enable Logging function", MatchPath: "Match full path on search"
                    , ListGrid: "Show Grid in command list", ListHdr: "Show Header in command list", SmartRank: "Smart Rank - Auto adjust command priority (rank) based on use frequency"
                    , SmartMatch: "Smart Match - Fuzzy and Smart matching and filtering result", MatchAny: "Match from any position of the string"
                    , ShowTheme: "Show Theme - Software skin and background picture", ShowHint: "Show Hints and Tips in the bottom status bar"
                    , ShowRunCount: "Show RunCount - Command running times in the status bar", ShowStatusBar: "Show Status Bar"
                    , ShowBtnRun: "Show [Run] Button on main window", ShowBtnOpt: "Show [Options] Button on main window"}
, KEYLIST_GUI    := "ListRows,ColWidth,FontName,FontSize,FontColor,WinWidth,WinHeight,CtrlColor,WinColor,Background"
, KEYLIST_HOTKEY := "GlobalHotkey1,GlobalHotkey2,Hotkey1,Trigger1,Hotkey2,Trigger2,Hotkey3,Trigger3,TotalCMDDir,ExplorerDir"
, TRAYMENUS      := ["Show,ToggleWindow,Shell32.dll,-25","","Options `tF2,Options,Shell32.dll,-16826","ReIndex `tCtrl+I,Reindex,Shell32.dll,-16776"
                    ,"Help `tF1,Help,Shell32.dll,-24","","Script Info,TrayMenu,imageres.dll,-150","AHK Manual,TrayMenu,Shell32.dll,-512",""
                    ,"Reload `tCtrl+Q,Reload,imageres.dll,-5311","Exit `tAlt+F4,Exit,imageres.dll,-98"]

, g_AutoStartup   := 1   , g_IndexDir        := "A_ProgramsCommon,A_StartMenu,C:\Other\Index\Location"
, g_EnableSendTo  := 1   , g_IndexType       := "*.lnk,*.exe"
, g_InStartMenu   := 1   , g_IndexExclude    := "Uninstall *"
, g_MatchPath     := 0   , g_FileMgr         := "Explorer.exe"
, g_ShowIcon      := 1   , g_HideOnLostFocus := 1
, g_ShowBtnRun    := 1   , g_ShowBtnOpt      := 1
, g_KeepInput     := 1   , g_Editor          := "Notepad.exe"
, g_AlwaysOnTop   := 1   , g_HistoryLen      := 15
, g_SaveHistory   := 1   , g_Everything      := "C:\Apps\Everything\Everything.exe"
, g_RunCount      := 0   , g_FileMgrID       := "ahk_class CabinetWClass, ahk_class TTOTAL_CMD"
, g_Logging       := 1   , g_DialogWin       := "ahk_class #32770"
, g_EscClearInput := 1   , g_ExcludeWin      := "ahk_class SysListView32, ahk_exe Explorer.exe, AutoCAD"
, g_SendToGetLnk  := 1   , g_FontName        := "Segoe UI"
, g_AutoSwitchDir := 0   , g_FontColor       := "Default"
, g_ShowTrayIcon  := 1   , g_CtrlColor       := "Default"
, g_ListRows      := 9   , g_WinColor        := "Silver"
, g_ListGrid      := 0   , g_ColWidth        := "40,60,430,340"
, g_ListHdr       := 1   , g_SmartMatch      := 1
, g_SmartRank     := 1   , g_MatchAny        := 1
, g_ShowTheme     := 1   , g_ShowHint        := 1
, g_GlobalHotkey1   := "!Space"
, g_GlobalHotkey2   := "!R"
, g_FontSize      := 10  , g_TotalCMDDir     := "^g" ; Hotkey for Listary quick-switch dir    
, g_WinWidth      := 900 , g_ExplorerDir     := "^e"
, g_WinHeight     := 330 , g_Background      := "Default"
, g_ShowRunCount  := 1   , g_ShowStatusBar   := 1
, g_Hotkey1       := "^s", g_Trigger1        := "Everything"
, g_Hotkey2       := "^p", g_Trigger2        := "RunPTTools"
, g_Hotkey3       := ""  , g_Trigger3        := ""
, g_BGPicture, OneDrive, OneDriveConsumer, OneDriveCommercial
, g_Hints := ["It's better to show me by press hotkey (Default is ALT + Space)"
    , "ALT + Space = Show / Hide window", "Alt + F4 = Exit"
    , "Esc = Clear input / Close window", "Enter = Run current command"
    , "Alt + No. = Run specific command", "Start with + = New Command"
    , "Ctrl + No. = Select specific command"
    , "F1 = ALTRun Help Index", "F2 = Open Setting Config window"
    , "F3 = Edit current command (.ini) directly"
    , "F4 = Edit user-defined commands (.ini) directly"
    , "Arrow Up / Down = Move to Previous / Next command"
    , "Ctrl+Q = Reload ALTRun"
    , "Ctrl+'+' = Increase rank of current command"
    , "Ctrl+'-' = Decrease rank of current command"
    , "Ctrl+I = Reindex file search database"
    , "Ctrl+D = Open current command dir with TC / File Explorer"
    , "Command priority (rank) will auto adjust based on frequency"
    , "Start with space = Search file by Everything"]

EnvGet, OneDrive, OneDrive                                              ; OneDrive Environment Variables (due to #NoEnv)
EnvGet, OneDriveConsumer, OneDriveConsumer                              ; OneDrive for Personal
EnvGet, OneDriveCommercial, OneDriveCommercial                          ; OneDrive for Business

;=============================================================
; 声明全局变量
;=============================================================
Global Arg                              ; 用来调用管道的完整参数（所有列）
, g_WinName := "ALTRun - Ver 2024.11"   ; 主窗口标题
, g_OptionsWinName := "Options"         ; 选项窗口标题
, g_Commands                            ; 所有命令
, g_Fallback                            ; 当搜索无结果时使用的命令
, g_History := Object()                 ; 历史命令
, g_Input                               ; 编辑框当前内容, 也作为 ControlID 使用
, g_CurrentCommand := ""                ; 当前匹配到的第一条命令
, g_CurrentCommandList := Object()      ; 当前匹配到的所有命令
, g_UseDisplay                          ; 命令使用了显示框
, g_UseFallback                         ; 使用备用的命令

Log.Debug("●●●●● ALTRun is starting ●●●●●")
LOADCONFIG("initialize")                                                ; Load ini config, IniWrite will create it if not exist

;=============================================================
; Create ContextMenu and TrayMenu
;=============================================================
Menu, LV_ContextMenu, Add, Run`tEnter, LVContextMenu                    ; ListView ContextMenu
Menu, LV_ContextMenu, Add, Open Container`tCtrl+D, OpenCurrentFileDir
Menu, LV_ContextMenu, Add, Copy Command, LVContextMenu
Menu, LV_ContextMenu, Add
Menu, LV_ContextMenu, Add, New Command, CmdMgr
Menu, LV_ContextMenu, Add, Edit Command`tF3, EditCurrentCommand
Menu, LV_ContextMenu, Add, User Commands`tF4, UserCommandList

Menu, LV_ContextMenu, Icon, Run`tEnter, Shell32.dll, -25
Menu, LV_ContextMenu, Icon, Open Container`tCtrl+D, Shell32.dll, -4
Menu, LV_ContextMenu, Icon, Copy Command, Shell32.dll, -243
Menu, LV_ContextMenu, Icon, New Command, Shell32.dll, -1
Menu, LV_ContextMenu, Icon, Edit Command`tF3, Shell32.dll, -16775
Menu, LV_ContextMenu, Icon, User Commands`tF4, Shell32.dll, -44
Menu, LV_ContextMenu, Default, Run`tEnter                               ; 让 "Run" 粗体显示表示双击时会执行相同的操作.

if (g_ShowTrayIcon)
{
    Menu, Tray, NoStandard
    Menu, Tray, Icon
    Menu, Tray, Icon, Shell32.dll, -25                                  ; Index of icon changes between Windows versions, refer to the icon by resource ID for consistency
    For Index, MenuItem in TRAYMENUS
    {
        Item := StrSplit(MenuItem, ",")
        Name := Item[1], Func := Item[2], Icon := Item[3], IconNo := Item[4]
        Menu, Tray, Add, %Name%, %Func%
        Menu, Tray, Icon, %Name%, %Icon%, %IconNo%
    }
    Menu, Tray, Tip, %g_WinName%
    Menu, Tray, Default, Show
    Menu, Tray, Click, 1
}
;=============================================================
; Load commands database and command history
; Update "SendTo", "Startup", "StartMenu" lnk
;=============================================================
LoadCommands()
LoadHistory()

Log.Debug("Updating 'SendTo' setting..." UpdateSendTo(g_EnableSendTo))
Log.Debug("Updating 'Startup' setting..." UpdateStartup(g_AutoStartup))
Log.Debug("Updating 'StartMenu' setting..." UpdateStartMenu(g_InStartMenu))

;=============================================================
; 主窗口配置代码
;=============================================================
AlwaysOnTop  := g_AlwaysOnTop ? "+AlwaysOnTop" : ""
ListGrid     := g_ListGrid ? "Grid" : ""
ListHdr      := g_ListHdr ? "" : "-Hdr"
LV_H         := g_WinHeight - 43 - 3 * g_FontSize
LV_W         := g_WinWidth - 24
HideWin      := ""
Input_W      := LV_W - g_ShowBtnRun * 90 - g_ShowBtnOpt * 90
Enter_W      := g_ShowBtnRun * 80
Enter_X      := g_ShowBtnRun * 10
Options_W    := g_ShowBtnOpt * 80
Options_X    := g_ShowBtnOpt * 10

Gui, Main:Color, %g_WinColor%, %g_CtrlColor%
Gui, Main:Font, c%g_FontColor% s%g_FontSize%, %g_FontName%
Gui, Main:%AlwaysOnTop%
Gui, Main:Add, Edit, xm W%Input_W% -WantReturn vg_Input gGetInput, Type anything here to search...
Gui, Main:Add, Button, % "X+"Enter_X " yp W" Enter_W " hp Default Hidden" !g_ShowBtnRun " gRunCurrentCommand", Enter
Gui, Main:Add, Button, % "X+"Options_X " yp W" Options_W " hp Hidden" !g_ShowBtnOpt " gOptions", Options
Gui, Main:Add, ListView, xm ys+35 W%LV_W% H%LV_H% vMyListView AltSubmit gLVActions %ListHdr% +LV0x10000 %ListGrid% -Multi, No.|Type|Command|Description ; LV0x10000 Paints via double-buffering, which reduces flicker
Gui, Main:Add, Picture, X0 Y0 0x4000000, %g_BGPicture%
Gui, Main:Add, StatusBar, % "Hidden" !g_ShowStatusBar " gSBActions",
Gui, Main:Default                                                       ; Set default GUI before any ListView / statusbar update

SB_SetParts(g_WinWidth-90*g_ShowRunCount)
Loop, 4
{
    LV_ModifyCol(A_Index, StrSplit(g_ColWidth, ",")[A_Index])
}

ListResult("Tip | F1 | Help`nTip | F2 | Options and settings`n"         ; List initial tips
    . "Tip | F3 | Edit current command`nTip | F4 | User-defined commands`n"
    . "Tip | ALT+SPACE / ALT+R | Activative ALTRun`n"
    . "Tip | ALT+SPACE / ESC / LOSE FOCUS | Deactivate ALTRun`n"
    . "Tip | ENTER / ALT+NO. | Run selected command`n"
    . "Tip | ARROW UP or DOWN | Select previous or next command`n"
    . "Tip | CTRL+D | Open selected cmd's dir with File Manager")

if (g_ShowIcon)
{
    Global ImageListID1 := IL_Create(10, 5)                             ; Create an ImageList so that the ListView can display some icons
    IL_Add(ImageListID1, "shell32.dll", -4)                             ; Add folder icon for dir type (IconIndex=1)
    IL_Add(ImageListID1, "shell32.dll", -25)                            ; Add app default icon for Func type (IconIndex=2)
    IL_Add(ImageListID1, "shell32.dll", -512)                           ; Add Browser icon for url type (IconIndex=3)
    IL_Add(ImageListID1, "shell32.dll", -22)                            ; Add control panel icon for ctrl type (IconIndex=4)
    IL_Add(ImageListID1, "Calc.exe", -1)                                ; Add calculator icon for Eval type (IconIndex=5)
    LV_SetImageList(ImageListID1)                                       ; Attach the ImageLists to the ListView so that it can later display the icons
}

Log.Debug("Resolving command line args=" A_Args[1] " " A_Args[2])       ; Command line args, Args are %1% %2% or A_Args[1] A_Args[2]
if (A_Args[1] = "-Startup")
    HideWin := " Hide"

if (A_Args[1] = "-SendTo")
{
    HideWin := " Hide"
    CmdMgr(A_Args[2])
}

Gui, Main:Show, Center w%g_WinWidth% h%g_WinHeight% %HideWin%, %g_WinName%

if (g_HideOnLostFocus)
{
    OnMessage(0x06, "WM_ACTIVATE")
}

;=============================================================
; Set Hotkey
;=============================================================
Hotkey, %g_GlobalHotkey1%, ToggleWindow                                 ; Set Global Hotkeys
Hotkey, %g_GlobalHotkey2%, ToggleWindow

Hotkey, IfWinActive, %g_WinName%                                        ; Hotkey take effect only when ALTRun actived
Hotkey, !F4, Exit
Hotkey, Tab, TabFunc
Hotkey, F1, Help
Hotkey, F2, Options
Hotkey, F3, EditCurrentCommand
Hotkey, F4, UserCommandList
Hotkey, ^q, Reload
Hotkey, ^d, OpenCurrentFileDir
Hotkey, ^i, Reindex
Hotkey, ^NumpadAdd, RankUp
Hotkey, ^NumpadSub, RankDown
Hotkey, Down, NextCommand
Hotkey, Up, PrevCommand

;=============================================================
; Run or locate command shortcut: Ctrl Alt Shift + No.
;=============================================================
Loop, %g_ListRows%                                                      ; ListRows limit <= 9
{
    Hotkey, !%A_Index%, RunSelectedCommand                              ; ALT + No. run command
    Hotkey, ^%A_Index%, GotoCommand                                     ; Ctrl + No. locate command
}

Loop, 3                                                                 ; Set Trigger <-> Hotkey
{
    Hotkey  := % g_Hotkey%A_Index%
    Trigger := % g_Trigger%A_Index%

    if (Hotkey != "" and IsFunc(Trigger))
        Hotkey, %Hotkey%, %Trigger%
}
Hotkey, IfWinActive

Listary()
AppControl()                                                            ; Set Listary Dir QuickSwitch, Set AppControl
Return

Activate()
{
    Gui, Main:Show,,%g_WinName%

    WinWaitActive, %g_WinName%,, 3                                      ; Use WinWaitActive 3s instead of previous Loop method
    {
        GuiControl, Main:Focus, g_Input
        ControlSend, Edit1, ^a, %g_WinName%                             ; Select all content in Input Box
    }
}

ToggleWindow()
{
    WinActive(g_WinName) ? MainGuiClose() : Activate()
}

GetInput()
{
    GuiControlGet, g_Input, Main:, g_Input                              ; Gui, Main:Submit, NoHide
    SearchCommand(g_Input)
}

SearchCommand(command := "")
{
    Result := ""
    Order  := 1
    g_CurrentCommandList := Object()
    Prefix := SubStr(command, 1, 1)

    if (Prefix = "+" or Prefix = " " or Prefix = ">")
    {
        g_CurrentCommand := g_Fallback[InStr("+ >", Prefix)]            ; Corresponding to fallback commands position no. 1, 2 & 3
        g_CurrentCommandList.Push(g_CurrentCommand)
        ListResult(g_CurrentCommand)
        Return
    }

    for index, element in g_Commands
    {
        splitResult := StrSplit(element, " | ")
        _Type := splitResult[1]
        _Path := splitResult[2]
        _Desc := splitResult[3]
        SplitPath, _Path, fileName                                      ; Extra name from _Path (if _Type is Dir and has "." in path, nameNoExt will not get full folder name)

        if (_Type = "dir")
        {
            elementToShow := _Type " | " fileName " | " _Desc           ; Show folder name only
        } 
        else
        {
            elementToShow := _Type " | " _Path " | " _Desc              ; Use _Path to show file icons (file type), and all other types
        }

        elementToSearch := g_MatchPath ? _Path " " _Desc : fileName " " _Desc ; search file name include extension & desc, search dir type + folder name + desc

        if (FuzzyMatch(elementToSearch, command))
        {
            g_CurrentCommandList.Push(element)

            if (Order = 1)
            {
                g_CurrentCommand := element
                Result .= elementToShow
            }
            else
            {
                Result .= "`n" elementToShow
            }
            Order++
            if (Order > g_ListRows)
                Break
        }
    }

    if (Result = "") {
        if Eval(command)
        {
            EvalResult := Eval(command)
            RebarQty := Ceil((EvalResult-40*2) / 300) + 1

            Result1 := "Eval | = " FormatThousand(EvalResult)
            Result2 := "`n | ------------------------------------------------------"
            Result3 := "`n | Beam width = " FormatThousand(EvalResult) " mm"
            Result4 := "`n | Main bar no. = " RebarQty " (" Round((EvalResult-40*2) / (RebarQty - 1)) " c/c), " RebarQty + 1 " (" Round((EvalResult-40*2) / (RebarQty+1-1)) " c/c), " RebarQty - 1 " (" Round((EvalResult-40*2) / (RebarQty-1-1)) " c/c)"
            Result5 := "`n | ------------------------------------------------------"
            Result6 := "`n | As = " FormatThousand(EvalResult) " mm2"
            Result7 := "`n | Rebar = " Ceil(EvalResult/132.7) "H13 / " Ceil(EvalResult/201.1) "H16 / " Ceil(EvalResult/314.2) "H20 / " Ceil(EvalResult/490.9) "H25 / " Ceil(EvalResult/804.2) "H32"
            Return ListResult(Result1 Result2 Result3 Result4 Result5 Result6 Result7, True)
        }

        g_UseFallback        := true
        g_CurrentCommandList := g_Fallback
        g_CurrentCommand     := g_Fallback[1]
        Result               := g_Fallback[1]
    
        Loop, % g_Fallback.MaxIndex() {
            if (A_Index = 1)
                continue
            Result .= "`n" g_Fallback[A_Index]
        }
    } else {
        g_UseFallback := false
    }
    
    ListResult(Result)
}

ListResult(text := "", UseDisplay := false)
{
    g_UseDisplay := UseDisplay
    IconIndex    := ""

    Gui, Main:Default                                                   ; Set default GUI before update any listview or statusbar
    GuiControl, Main:-Redraw, MyListView                                ; Improve performance by disabling redrawing during load.
    LV_Delete()
    VarSetCapacity(sfi, sfi_size := 698)                                ; Calculate buffer size required for SHFILEINFO structure.
    
    Loop Parse, text, `n, `r
    {
        splitResult := StrSplit(A_LoopField, " | ")
        _Type := splitResult[1]
        _Path := AbsPath(splitResult[2])                                ; Must store in var for afterward use, trim space (in AbsPath)
        _Desc := splitResult[3]

        if (g_ShowIcon)
        {
            ; Build a unique extension ID to avoid characters that are illegal in variable names, such as dashes. 
            ; This unique ID method also performs better because finding an item, in the array does not require search-loop.
            SplitPath, _Path,,, FileExt                                 ; Get the file's extension.

            if (_Type = "Dir")
            {
                IconIndex := 1
            }
            else if _Type contains Func,CMD,Tip
            {
                IconIndex := 2
            }
            else if (_Type = "URL")
            {
                IconIndex := 3
            }
            else if (_Type = "Ctrl")
            {
                IconIndex := 4
            }
            else if (_Type = "Eval")
            {
                IconIndex := 5
            }
            else if FileExt in EXE,ICO,ANI,CUR,LNK
            {
                ExtID := FileExt
                IconIndex := 0                                          ; Flag it as not found so that these types can each have a unique icon
            }
            else                                                        ; Some other extension/file-type, so calculate its unique ID
            {
                ExtID := 0                                              ; Initialize to handle extensions that are shorter than others
                Loop Parse, FileExt
                    ExtID .= Format("{:02X}", Asc(A_LoopField))         ; Derive a Unique ID by convert FileExt to HEX

                IconIndex := IconArray%ExtID%                           ; Check if this file extension already has an icon in the ImageLists. If it does, several calls can be avoided and loading performance is greatly improved, especially for a folder containing hundreds of files
            }

            if (!IconIndex)                                             ; There is not yet any icon for this extension, so load it.
            {
                if (!DllCall("Shell32\SHGetFileInfoW", "Str", _Path, "UInt", 0, "Ptr", &sfi, "UInt", sfi_size, "UInt", 0x101)) ; 0x101 is SHGFI_ICON+SHGFI_SMALLICON
                    IconIndex = 9999999                                 ; Set it out of bounds to display a blank icon.
                else                                                    ; Icon successfully loaded. Extract the hIcon member from the structure
                {
                    hIcon := NumGet(sfi, 0)                             ; Add the HICON directly to the small-icon and large-icon lists.
                    IconIndex := DllCall("ImageList_ReplaceIcon", "ptr", ImageListID1, "int", -1, "ptr", hIcon) + 1 ; Uses +1 to convert the returned index from zero-based to one-based:
                    DllCall("DestroyIcon", "ptr", hIcon)                ; Now that it's been copied into the ImageLists, the original should be destroyed
                    IconArray%ExtID% := IconIndex                       ; Cache the icon to save memory and improve loading performance
                }
            }
        }
        LV_Add("Icon"IconIndex, A_Index, _Type, _Path, _Desc)
    }

    LV_Modify(1, "Select Focus Vis")                                    ; Select 1st row
    GuiControl, Main:+Redraw, MyListView
    SetStatusBar()
}

AbsPath(Path, KeepRunAs := False)                                       ; Convert path to absolute path
{
    Path := Trim(Path)
    
    if (!KeepRunAs)
        Path := StrReplace(Path,  "*RunAs ", "")                        ; Remove *RunAs (Admin Run) to get absolute path

    if (InStr(Path, "A_"))                                              ; Resolve path like A_ScriptDir
        Path := %Path%

    Path := StrReplace(Path, "%Temp%", A_Temp)
    Path := StrReplace(Path, "%Desktop%", A_Desktop)
    Path := StrReplace(Path, "%OneDrive%", OneDrive)                    ; Convert OneDrive to absolute path due to #NoEnv
    Path := StrReplace(Path, "%OneDriveConsumer%", OneDriveConsumer)    ; Convert OneDrive to absolute path due to #NoEnv
    Path := StrReplace(Path, "%OneDriveCommercial%", OneDriveCommercial) ; Convert OneDrive to absolute path due to #NoEnv
    Return Path
}

RelativePath(Path)                                                      ; Convert path to relative path
{
    Path := StrReplace(Path, A_Temp, "%Temp%")
    Path := StrReplace(Path, A_Desktop, "%Desktop%")
    Path := StrReplace(Path, OneDrive, "%OneDrive%")
    Path := StrReplace(Path, OneDriveConsumer, "%OneDriveConsumer%")
    Path := StrReplace(Path, OneDriveCommercial, "%OneDriveCommercial%")
    Return Path
}

RunCommand(originCmd)
{
    MainGuiClose()                                                      ; 先关闭窗口,避免出现延迟的感觉
    ParseArg()
    g_UseDisplay := false

    _Type := StrSplit(originCmd, " | ")[1]
    _Path := AbsPath(StrSplit(originCmd, " | ")[2], True)

    switch (_Type)
    {
        case "file":
            Run, %_Path%,, UseErrorLevel
            if ErrorLevel
                MsgBox Could not open "%_Path%"
        case "dir":
            OpenDir(_Path)
        case "func":
            if IsFunc(_Path)
                %_Path%()
        case "cmd":
            RunWithCmd(_Path)
        default:                                                        ; for type: url, control & all other un-defined type
            Run, %_Path%
    }

    if (g_SaveHistory)
    {
        g_History.InsertAt(1, originCmd " /arg=" Arg)                   ; Adjust command history

        if (g_History.Length() > g_HistoryLen)
            g_History.Pop()

        for index, element in g_History
            IniWrite, %element%, %g_IniFile%, %SEC_HISTORY%, %index%    ; Save command history
    }

    g_RunCount++
    IniWrite, %g_RunCount%, %g_IniFile%, %SEC_CONFIG%, RunCount         ; Record running number
    UpdateRank(originCmd)
    Log.Debug("Execute(" g_RunCount ")=" originCmd)
}

TabFunc()
{
    GuiControlGet, CurrCtrl, Main:FocusV                                ; Limit tab to switch between Edit1 & ListView only
    GuiControl, Main:Focus, % (CurrCtrl = "g_Input") ? "MyListView" : "g_Input"
}

PrevCommand()
{
    ChangeCommand(-1, False)
}

NextCommand()
{
    ChangeCommand(1, False)
}

GotoCommand()
{
    index := SubStr(A_ThisHotkey, 0, 1)                                 ; Get index from hotkey (select specific command = Shift + index)
    g_CurrentCommand := g_CurrentCommandList[index]

    if (g_CurrentCommand != "")
        ChangeCommand(index, True)
}

ChangeCommand(Step = 1, ResetSelRow = False)
{
    Gui, Main:Default                                                   ; Use it before any LV update

    SelRow := ResetSelRow ? Step : LV_GetNext() + Step                  ; Get target row no. to be selected
    SelRow := SelRow > LV_GetCount() ? 1 : SelRow                       ; Listview cycle selection (Mod has bug on upward cycle)
    SelRow := SelRow < 1 ? LV_GetCount() : SelRow
    g_CurrentCommand := g_CurrentCommandList[SelRow]                    ; Get current command from selected row

    LV_Modify(SelRow, "Select Focus Vis")                               ; make new index row selected, Focused, and Visible
    SetStatusBar()
}

LVActions()                                                             ; ListView g label actions (left / double click) behavior
{
    Gui, Main:Default                                                   ; Use it before any LV update
    focusedRow := LV_GetNext(0, "Focused")                              ; 查找焦点行, 仅对焦点行进行操作而不是所有选择的行:
    if (!focusedRow)                                                    ; 没有焦点行
        Return

    g_CurrentCommand := g_CurrentCommandList[focusedRow]                ; Get current command from focused row

    if (A_GuiEvent = "RightClick")
    {
        Menu, LV_ContextMenu, Show
    }
    else if (A_GuiEvent = "DoubleClick" and g_CurrentCommand)           ; Double click behavior, if g_CurrentCommand = "" eg. first tip page, run it will clear SEC_USERCMD, SEC_INDEX, SEC_DFTCMD
    {
        RunCommand(g_CurrentCommand)
    }
    else if (A_GuiEvent = "Normal")                                     ; left click behavior
    {
        SetStatusBar()
    }
}

LVContextMenu()                                                         ; ListView ContextMenu (right click & its menu) actions
{
    Gui, Main:Default                                                   ; Use it before any LV update
    focusedRow := LV_GetNext(0, "Focused")                              ; Check focused row, only operate focusd row instead of all selected rows
    if (!focusedRow)                                                    ; if not found
        Return

    g_CurrentCommand := g_CurrentCommandList[focusedRow]                ; Get current command from focused row
    If (A_ThisMenuItem = "Run`tEnter")                                  ; User selected "Run`tEnter"
        RunCommand(g_CurrentCommand)
    else if (A_ThisMenuItem = "Copy Command")
    {
        LV_GetText(Text, focusedRow, 3)                                 ; Get the text from the focusedRow's 3rd field.
        A_Clipboard := Text
    }
}

SBActions()
{
    if (A_GuiEvent = "RightClick" and A_EventInfo = 1)
    {
        Menu, SB_ContextMenu, Add, Copy, SBContextMenu
        Menu, SB_ContextMenu, Icon, Copy, Shell32.dll, -243
        Menu, SB_ContextMenu, Show
    }
    else if (A_GuiEvent = "Normal" and A_EventInfo = 2)
    {
        MsgBox, 64, %g_WinName%, Congraduations! You have run shortcut %g_RunCount% times by now!
    }
}

SBContextMenu()
{
    StatusBarGetText, A_Clipboard, 1, %g_WinName%
}

TrayMenu()
{
    If ( A_ThisMenuItem = "Script Info" )
        ListLines
    If ( A_ThisMenuItem = "AHK Manual" )
        Run, https://www.autohotkey.com/docs/v1/
}

MainGuiEscape()
{
    (g_EscClearInput and g_Input) ? ClearInput() : MainGuiClose()
}

MainGuiClose()                                                          ; If GuiClose is a function, the GUI is hidden by default
{
    (!g_KeepInput) ? ClearInput()
    Gui, Main:Hide
    SetStatusBar("Hint")                                                ; Update StatusBar hint information after GUI hide (move code from Activate() to here for better performance)
}

Exit()
{
    ExitApp
}

Reload()
{
    Reload
}

Test()
{
    t := A_TickCount
    Loop 50
    {
        random,chr1,asc("a"),asc("z")
        random,chr2,asc("A"),asc("Z") ;65,90
        random,chr3,asc("a"),asc("z") ;97,122

        Activate()
        GuiControl, Main:Text, g_Input, % chr(chr1)
        Sleep, 10
        GuiControl, Main:Text, g_Input, % chr(chr1) " " chr(chr2)
        Sleep, 10
        GuiControl, Main:Text, g_Input, % chr(chr1) " " chr(chr2) " " chr(chr3)
    }
    t := A_TickCount - t
    Log.Debug("mock test search ' " chr(chr1) " " chr(chr2) " " chr(chr3) " ' 50 times, use time = " t)
    MsgBox % "Search '" chr(chr1) " " chr(chr2) " " chr(chr3) "' use Time =  " t
    Return
}

UserCommandList()
{
    if (g_Editor != "")
    {
        Run, % g_Editor " /m [" SEC_USERCMD "] """ g_IniFile """"       ; /m Match text
    }
    else
    {
        Run, % g_IniFile
    }
}

ClearInput()
{
    GuiControl, Main:Text, g_Input,
    GuiControl, Main:Focus, g_Input
}

SetStatusBar(Mode := "Command")                                         ; Set StatusBar text, Mode 1: Current command (default), 2: Hint, 3: Any text
{
    Gui, Main:Default                                                   ; Set default GUI window before any ListView / StatusBar operate
    if (Mode = "Command")
    {
        SBText := StrSplit(g_CurrentCommand, " | ")[2]
    }
    else if (Mode = "Hint")
    {
        Random, index, 1, g_Hints.Length()
        SBText := g_Hints[index]
    }
    else
    {
        SBText := Mode
    }
    SB_SetText(SBText, 1)
    SB_SetText("RC: "g_RunCount, 2)
}

RunCurrentCommand()
{
    RunCommand(g_CurrentCommand)
}

ParseArg()
{
    Global
    commandPrefix := SubStr(g_Input, 1, 1)

    if (commandPrefix = "+" || commandPrefix = " " || commandPrefix = ">")
    {
        Arg := SubStr(g_Input, 2)                                       ; 直接取命令为参数
        Return
    }

    if (InStr(g_Input, " ") && !g_UseFallback)                          ; 用空格来判断参数
    {
        Arg := SubStr(g_Input, InStr(g_Input, " ") + 1)
    }
    else if (g_UseFallback)
    {
        Arg := g_Input
    }
    else
    {
        Arg := ""
    }
}

FuzzyMatch(Haystack, Needle)
{
    Needle := StrReplace(Needle, "+", "\+")                             ; for Eval (preceded by a backslash to be seen as literal)
    Needle := StrReplace(Needle, "*", "\*")                             ; for Eval (preceded by a backslash to be seen as literal)
    Needle := StrReplace(Needle, " ", ".*")                             ; RegExMatch should able to search with space as & separater, but not sure, use this way for now
    Return RegExMatch(Haystack, "imS)" Needle)
}

UpdateRank(originCmd, showRank := false, inc := 1)
{
    RANKSEC := SEC_DFTCMD "|" SEC_USERCMD "|" SEC_INDEX
    Loop Parse, RANKSEC, |                                              ; Update Rank for related sections
    {
        IniRead, Rank, %g_IniFile%, %A_LoopField%, %originCmd%, KeyNotFound

        if (Rank = "KeyNotFound" or Rank = "ERROR" or originCmd = "")   ; If originCmd not exist in this section, then check next section
            continue                                                    ; Skips the rest of a loop and begins a new one.
        else if Rank is integer                                         ; If originCmd exist in this section, then update it's rank.
            Rank += inc
        else
            Rank := inc

        if (Rank < 0)                                                   ; 如果降到负数,都设置成 -1,然后屏蔽/排除
            Rank := -1

        IniWrite, %Rank%, %g_IniFile%, %A_LoopField%, %originCmd%       ; Update new Rank for originCmd

        if (showRank)
        {
            SetStatusBar("Rank for current command : " Rank)
        }
    }
    LoadCommands()                                                      ; New rank will take effect in real-time by LoadCommands again
}

RunSelectedCommand()
{
    index := SubStr(A_ThisHotkey, 0, 1)
    RunCommand(g_CurrentCommandList[index])
}

RankUp()
{
    UpdateRank(g_CurrentCommand, true)
}

RankDown()
{
    UpdateRank(g_CurrentCommand, true, -1)
}

LoadCommands()
{
    g_Commands  := Object()                                             ; Clear g_Commands list
    g_Fallback  := Object()                                             ; Clear g_Fallback list
    RankString  := ""

    Loop Parse, % LOADCONFIG("commands"), `n                            ; Read commands sections (built-in, user & index), read each line, separate key and value
    {
        command := StrSplit(A_LoopField, "=")[1]                        ; pass first string (key) to command
        rank    := StrSplit(A_LoopField, "=")[2]                        ; pass second string (value) to rank

        if (command != "" and rank > 0)
            RankString .= rank "`t" command "`n"
    }
    Sort, RankString, R N
    Loop Parse, RankString, `n
    {
        command := StrSplit(A_LoopField, "`t")[2]
        g_Commands.Push(command)
    }
    
    IniRead, FALLBACKCMDSEC, %g_IniFile%, %SEC_FALLBACK%                ;read FALLBACK section, initialize it if section not exist
    if (FALLBACKCMDSEC = "")
    {
        IniWrite, 
        (Ltrim
        ;===========================================================
        ; Fallback Commands show when search result is empty
        ;
        Func | CmdMgr | New Command
        Func | Everything | Search by Everything
        Func | CmdRun | Run Command use CMD
        Func | Google | Search Clipboard or Input by Google
        Func | AhkRun | Run Command use AutoHotkey Run
        Func | Bing | Search Clipboard or Input by Bing
        ), %g_IniFile%, %SEC_FALLBACK%
        IniRead, FALLBACKCMDSEC, %g_IniFile%, %SEC_FALLBACK%
    }
    Loop Parse, FALLBACKCMDSEC, `n
    {
        if (IsFunc(StrSplit(A_LoopField, " | ")[2]))                    ;read each line, get each FBCommand (Rank not necessary)
        {
            g_Fallback.Push(A_LoopField)
        }
    }
    Return Log.Debug("Loading commands list...OK")
}

LoadHistory()
{
    if (g_SaveHistory)
    {
        Loop %g_HistoryLen%
        {
            IniRead, History, %g_IniFile%, %SEC_HISTORY%, %A_Index%
            g_History.Push(History)
        }
    }
    else
    {
        IniDelete, %g_IniFile%, %SEC_HISTORY%
    }
}

GetCmdOutput(command)
{
    TempFile   := A_Temp "\ALTRun.stdout"
    FullCommand = %ComSpec% /C "%command% > %TempFile%"

    RunWait, %FullCommand%, %A_Temp%, Hide
    FileRead, Result, %TempFile%
    FileDelete, %TempFile%
    Return RTrim(Result, "`r`n")                                        ; Remove result rightmost/last "`r`n"
}

GetRunResult(command)                                                   ;运行CMD并取返回结果方式2
{
    shell := ComObjCreate("WScript.Shell")                              ; WshShell object: http://msdn.microsoft.com/en-us/library/aew9yb99
    exec := shell.Exec(ComSpec " /C " command)                          ; Execute a single command via cmd.exe
    Return exec.StdOut.ReadAll()                                        ; Read and Return the command's output
}

RunWithCmd(command)
{
    Run, % ComSpec " /C " command " & pause"
}

OpenDir(Path, OpenContainer := False)
{
    Path := AbsPath(Path)

    if (OpenContainer)
    {
        if (g_FileMgr = "Explorer.exe")
            Run, %g_FileMgr% /select`, "%Path%",, UseErrorLevel
        else
            Run, %g_FileMgr% " /P " "%Path%",, UseErrorLevel             ; /P Parent folder
    }
    else
        Run, %g_FileMgr% "%Path%",, UseErrorLevel

    if ErrorLevel
        MsgBox, 4096, %g_WinName%, Error found, error code : %A_LastError%

    Log.Debug("Open dir="Path)
}

OpenCurrentFileDir()
{
    OpenDir(StrSplit(g_CurrentCommand, " | ")[2], True)
}

EditCurrentCommand()
{
    if (g_Editor != "")
    {
        Run, % g_Editor " /m " """" g_CurrentCommand "=""" " """ g_IniFile """" ; /m Match text, locate to current command, add = at end to filter out [history] commands
    }
    else
    {
        Run, % g_IniFile
    }
}

WM_ACTIVATE(wParam, lParam)                                             ; Close on lose focus
{
    if (wParam > 0)                                                     ; wParam > 0: window being activated
        Return

    else if (wParam <= 0 && WinExist(g_WinName) && !g_UseDisplay)       ; wParam <= 0: window being deactivated (lost focus)
        MainGuiClose()
}

UpdateSendTo(create := true)                                            ; the lnk in SendTo must point to a exe
{
    lnkPath := StrReplace(A_StartMenu, "\Start Menu", "\SendTo\") "ALTRun.lnk"

    if (!create)
    {
        FileDelete, %lnkPath%
        Return "Disabled"
    }

    if (A_IsCompiled)
        FileCreateShortcut, "%A_ScriptFullPath%", %lnkPath%, ,-SendTo
        , Send command to ALTRun User Command list, Shell32.dll, , -25
    else
        FileCreateShortcut, "%A_AhkPath%", %lnkPath%, , "%A_ScriptFullPath%" -SendTo
        , Send command to ALTRun User Command list, Shell32.dll, , -25
    Return "OK"
}

UpdateStartup(create := true)
{
    lnkPath := A_Startup "\ALTRun.lnk"

    if (!create)
    {
        FileDelete, %lnkPath%
        Return "Disabled"
    }

    FileCreateShortcut, %A_ScriptFullPath%, %lnkPath%, %A_ScriptDir%
        , -startup, ALTRun - An effective launcher, Shell32.dll, , -25
    Return "OK"
}

UpdateStartMenu(create := true)
{
    lnkPath := A_Programs "\ALTRun.lnk"

    if (!create)
    {
        FileDelete, %lnkPath%
        Return "Disabled"
    }

    FileCreateShortcut, %A_ScriptFullPath%, %lnkPath%, %A_ScriptDir%
        , -StartMenu, ALTRun, Shell32.dll, , -25
    Return "OK"
}

Reindex()                                                               ; Re-create Index section
{
    IniDelete, %g_IniFile%, %SEC_INDEX%
    for dirIndex, dir in StrSplit(g_IndexDir, ",")
    {
        searchPath := AbsPath(dir)

        for extIndex, ext in StrSplit(g_IndexType, ",")
        {
            Loop Files, %searchPath%\%ext%, R
            {
                if (g_IndexExclude != "" && RegExMatch(A_LoopFileLongPath, g_IndexExclude))
                    continue                                            ; Skip this file and move on to the next loop.

                IniWrite, 1, %g_IniFile%, %SEC_INDEX%, File | %A_LoopFileLongPath% ; Assign initial rank to 1
                Progress, %A_Index%, %A_LoopFileName%, ReIndexing..., Reindex
            }
        }
        Progress, Off
    }

    Log.Debug("Indexing search database...")
    TrayTip, %g_WinName%, ReIndex database finish successfully. , 8
    LoadCommands()
}

Help()
{
    Options(Arg, 6)                                                     ; Open Options window 7th tab (help tab)
}

Listary()                                                               ; Listary Dir QuickSwitch Function (快速更换保存/打开对话框路径)
{
    Log.Debug("Listary function starting...")

    Loop Parse, g_FileMgrID, `,                                       ; File Manager Class, default is Windows Explorer & Total Commander
        GroupAdd, FileMgrID, %A_LoopField%

    Loop Parse, g_DialogWin, `,                                         ; 需要QuickSwith的窗口, 包括打开/保存对话框等
        GroupAdd, DialogBox, %A_LoopField%

    Loop Parse, g_ExcludeWin, `,                                        ; 排除特定窗口,避免被 Auto-QuickSwitch 影响
        GroupAdd, ExcludeWin, %A_LoopField%

    if (g_AutoSwitchDir)
    {
        Log.Debug("Listary Auto-QuickSwitch Enabled.")
        Loop
        {
            WinWaitActive ahk_class TTOTAL_CMD
                WinGet, ThisHWND, ID, A
            WinWaitNotActive

            If(WinActive("ahk_group DialogBox") && !WinActive("ahk_group ExcludeWin")) ; 检测当前窗口是否符合打开保存对话框条件
            {
                WinGetActiveTitle, Title
                WinGet, ActiveProcess, ProcessName, A

                Log.Debug("Listary dialog detected, active window ahk_title=" Title ", ahk_exe=" ActiveProcess)
                LocateTC()                                              ; NO Return, as will terimate loop (AutoSwitchDir)
            }
        }
    }
    Hotkey, IfWinActive, ahk_group DialogBox                            ; 设置对话框路径定位热键,为了不影响其他程序热键,设置只对打开/保存对话框生效
    Hotkey, %g_ExplorerDir%, LocateExplorer                             ; Ctrl+E 把打开/保存对话框的路径定位到资源管理器当前浏览的目录
    Hotkey, %g_TotalCMDDir%, LocateTC                                   ; Ctrl+G 把打开/保存对话框的路径定位到TC当前浏览的目录
    Hotkey, IfWinActive
}

LocateTC()                                                              ; Get TC current dir path, and change dialog box path to it
{
    ClipSaved := ClipboardAll 
    Clipboard :=
    SendMessage 1075, 2029, 0, , ahk_class TTOTAL_CMD
    ClipWait, 200
    OutDir=%Clipboard%\                                                 ; 结尾添加\ 符号,变为路径,试图解决AutoCAD不识别路径问题
    Clipboard := ClipSaved 
    ClipSaved := 

    ChangePath(OutDir)
}

LocateExplorer()                                                        ; Get Explorer current dir path, and change dialog box path to it
{
    Loop,9
    {
        ControlGetText, Dir, ToolbarWindow32%A_Index%, ahk_class CabinetWClass
    } until (InStr(Dir,"Address"))
 
    Dir:=StrReplace(Dir,"Address: ","")
    if (Dir="Computer")
        Dir:="C:\"

    If (SubStr(Dir,2,2) != ":\")                                             ; then Explorer lists it as one of the library directories such as Music or Pictures
        Dir:=% "C:\Users\" A_UserName "\" Dir

    ChangePath(Dir)
}

ChangePath(Dir)
{
    ControlGetText, w_Edit1Text, Edit1, A
    ControlClick, Edit1, A
    ControlSetText, Edit1, %Dir%, A
    ControlSend, Edit1, {Enter}, A
    ;Sleep,100
    ;ControlSetText, Edit1, %w_Edit1Text%, A                            ; 还原之前的窗口 File Name 内容, 在选择文件的对话框时没有问题, 但是在选择文件夹的对话框有Bug,暂时注释掉
    Log.Debug("Listary change path=" Dir)
}

CmdMgr(Path := "")                                                      ; 命令管理窗口
{
    Global
    Log.Debug("Starting Command Manager... Args=" Path)

    SplitPath Path, _Desc, fileDir, fileExt, nameNoExt, fileDrive       ; Extra name from _Path (if _Type is dir and has "." in path, nameNoExt will not get full folder name) 
    
    if InStr(FileExist(Path), "D")                                      ; True only if the file exists and is a directory.
        _Type := 5                                                      ; It is a normal folder
    else                                                                ; From command "New Command" or GUI context menu "New Command"
        _Desc := Arg
    
    if (fileExt = "lnk" && g_SendToGetLnk)
    {
        FileGetShortcut, %Path%, Path, fileDir, fileArg, _Desc
        Path .= " " fileArg
    }

    Gui, CmdMgr:New
    Gui, CmdMgr:Font, S8 W400, Century Gothic
    Gui, CmdMgr:Margin, 5, 5
    Gui, CmdMgr:Add, GroupBox, w550 h230, New Command
    Gui, CmdMgr:Add, Text, xp+20 yp+35, Command Type: 
    Gui, CmdMgr:Add, DropDownList, xp+120 yp-5 w150 v_Type Choose%_Type%, Func|URL|Command|File||Dir
    Gui, CmdMgr:Add, Text, xp-120 yp+50, Command Path: 
    Gui, CmdMgr:Add, Edit, xp+120 yp-5 w350 v_Path, % RelativePath(Path)
    Gui, CmdMgr:Add, Button, xp+355 yp w30 hp gSelectCmdPath, ...
    Gui, CmdMgr:Add, Text, xp-475 yp+100, Description: 
    Gui, CmdMgr:Add, Edit, xp+120 yp-5 w350 v_Desc, %_Desc%
    Gui, CmdMgr:Add, Button, Default x415 w65, OK
    Gui, CmdMgr:Add, Button, xp+75 yp w65, Cancel
    Gui, CmdMgr:Show, AutoSize, Commander Manager
}

SelectCmdPath()
{
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

CmdMgrButtonOK()
{
    Global
    Gui, CmdMgr:Submit
    _Desc := _Desc ? "| " _Desc : _Desc

    if (_Path = "")
    {
        MsgBox, Please input correct command path!
        Return
    }
    else
    {
        IniWrite, 1, %g_IniFile%, %SEC_USERCMD%, %_Type% | %_Path% %_Desc% ; initial rank = 1
        if (!ErrorLevel)
            MsgBox,, Command Manager, Command added successfully!
    }
    LoadCommands()
}

CmdMgrGuiEscape()
{
    Gui, CmdMgr:Destroy
}

CmdMgrButtonCancel()
{
    Gui, CmdMgr:Destroy
}

CmdMgrGuiClose()
{
    Gui, CmdMgr:Destroy
}

AppControl()                                                            ; AppControl (Ctrl+D 自动添加日期, 鼠标中间激活PT Tools)
{
    GroupAdd, FileListMangr, ahk_class TTOTAL_CMD                       ; TC 文件列表重命名
    GroupAdd, FileListMangr, ahk_class CabinetWClass                    ; Windows 资源管理器文件列表重命名
    GroupAdd, FileListMangr, ahk_class Progman                          ; Windows 桌面文件重命名 (WinXP to Win10)
    GroupAdd, FileListMangr, ahk_class WorkerW                          ; Windows 桌面文件重命名 (Win11)
    GroupAdd, FileListMangr, ahk_class TSTDTREEDLG                      ; TC 新建其他格式文件如txt, rtf, docx...
    GroupAdd, FileListMangr, ahk_class #32770                           ; 资源管理器 文件保存对话框
    GroupAdd, FileListMangr, ahk_class TCOMBOINPUT                      ; TC F7 创建新文件夹对话框（可单独出来用isFile:= True来控制不考虑后缀的影响）

    GroupAdd, TextBox, ahk_class TCmtEditForm                           ; TC File Comment 对话框 按Ctrl+D自动在备注文字之后添加日期
    GroupAdd, TextBox, ahk_class Notepad2                               ; Notepad2 (原Ctrl+D 为重复当前行)

    Hotkey, IfWinActive, ahk_group FileListMangr                        ; 针对所有设定好的程序 按Ctrl+D自动在文件(夹)名之后添加日期
    Hotkey, ^D, RenameWithDate
    Hotkey, IfWinActive, ahk_group TextBox
    Hotkey, ^D, LineEndAddDate
    Hotkey, IfWinActive, ahk_exe RAPTW.exe                              ; 如果正在使用RAPT,鼠标中间激活PT Tools
    Hotkey, ~MButton, RunPTTools
    Hotkey, IfWinActive
}

RunPTTools()
{
    IfWinNotExist, PT Tools
        Run, %A_ScriptDir%\PTTools.ahk
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
    Log.Debug("Add Date At End= - " CurrentDate)
}

NameAddDate(WinName, CurrCtrl, isFile:= True) {                         ; 在文件（夹）名编辑框中添加日期,CurrCtrl为当前控件(名称编辑框Edit),isFile是可选参数,默认为真
    ControlGetText, EditCtrlText, %CurrCtrl%, A
    SplitPath, EditCtrlText, fileName, fileDir, fileExt, nameNoExt
    FormatTime, CurrentDate,, dd.MM.yyyy

    if (isFile && fileExt != "" && StrLen(fileExt) < 5 && !RegExMatch(fileExt,"^\d+$")) ; 如果是文件,而且有真实文件后缀名,才加日期在后缀名之前, another way is use if fileExt in %TrgExtList% but can not check isFile at the same time
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
    Log.Debug(WinName ", RenameWithDate=" NameWithDate)
}

FormatThousand(Number)                                                  ; Function to add thousand separator
{
    Return RegExReplace(Number, "\G\d+?(?=(\d{3})+(?:\D|$))", "$0" ",")
}

Options(Arg := "", ActTab := 1)                                         ; Options / Settings Library, 1st parameter is to avoid menu like [Option `tF2] disturb ActTab
{
    Global                                                              ; Assume-global mode
    Log.Debug("Loading options window...Arg=" Arg ", ActTab=" ActTab)
    
    Gui, Setting:New, -SysMenu, %g_OptionsWinName%                      ;-SysMenu: omit the system menu and icon in the window's upper left corner
    Gui, Setting:Font, s9, Segoe UI
    Gui, Setting:Margin, 5, 5
    Gui, Setting:Add, Tab3,xm ym vCurrTab Choose%ActTab% -Wrap, General|Index|GUI|Hotkey|Listary|About

    Gui, Setting:Tab, 1 ; CONFIG Tab
    Gui, Setting:Add, ListView, w500 h300 Checked -Multi AltSubmit -Hdr vOptListView, Options

    For key, description in KEYS_CONFIG
    {
        LV_Add("Check"g_%key%, description)
    }
    LV_ModifyCol(1, "Auto")

    Gui, Setting:Add, Text, yp+320, Text Editor:
    Gui, Setting:Add, ComboBox, xp+100 yp-5 w400 Sort vg_Editor, %g_Editor%||Notepad.exe|C:\Apps\Notepad4\Notepad4.exe
    Gui, Setting:Add, Text, xp-100 yp+40, Everything.exe:
    Gui, Setting:Add, ComboBox, xp+100 yp-5 w400 Sort vg_Everything, %g_Everything%||C:\Apps\Everything\Everything.exe
    Gui, Setting:Add, Text, xp-100 yp+40, File Manager:
    Gui, Setting:Add, ComboBox, xp+100 yp-5 w400 Sort vg_FileMgr, %g_FileMgr%||Explorer.exe|C:\Apps\TotalCMD64\TotalCMD.exe /O /T /S
    
    Gui, Setting:Tab, 2 ; INDEX Tab
    Gui, Setting:Add, GroupBox, w500 h420, Index Options
    Gui, Setting:Add, Text, xp+10 yp+25, Index Locations: 
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 vg_IndexDir, %g_IndexDir%||A_ProgramsCommon,A_StartMenu
    Gui, Setting:Add, Text, xp-150 yp+40, Index File Type: 
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 vg_IndexType, %g_IndexType%||*.lnk,*.exe
    Gui, Setting:Add, Text, xp-150 yp+40, Index File Exclude: 
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 vg_IndexExclude, %g_IndexExclude%||Uninstall *
    Gui, Setting:Add, Text, xp-150 yp+40, Command History Length: 
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 Sort vg_HistoryLen, %g_HistoryLen%||0|10|20|30

    Gui, Setting:Tab, 3 ; GUI Tab
    Gui, Setting:Add, GroupBox, w500 h420, GUI
    Gui, Setting:Add, Text, xp+10 yp+25 , Command result limit
    Gui, Setting:Add, DropDownList, xp+150 yp-5 w330 vg_ListRows, % StrReplace("3|4|5|6|7|8|9|", g_ListRows, g_ListRows . "|") ; Not Choose%g_ListRows% as list start from 3
    Gui, Setting:Add, Text, xp-150 yp+40, Width of each column:
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 Sort vg_ColWidth, %g_ColWidth%||40,45,430,340
    Gui, Setting:Add, Text, xp-150 yp+40, Font Name:
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 Sort vg_FontName, %g_FontName%||Default|Segoe UI Semibold|Microsoft Yahei
    Gui, Setting:Add, Text, xp-150 yp+40, Font Size: 
    Gui, Setting:Add, ComboBox, xp+150 yp-5 r1 w330 vg_FontSize, %g_FontSize%||
    Gui, Setting:Add, Text, xp-150 yp+40, Font Color: 
    Gui, Setting:Add, ComboBox, xp+150 yp-5 r1 w330 vg_FontColor, %g_FontColor%||
    Gui, Setting:Add, Text, xp-150 yp+40, Window Width: 
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 vg_WinWidth, %g_WinWidth%||
    Gui, Setting:Add, Text, xp-150 yp+40, Window Height:
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 vg_WinHeight, %g_WinHeight%||
    Gui, Setting:Add, Text, xp-150 yp+40, Controls' Color:
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 vg_CtrlColor, %g_CtrlColor%||
    Gui, Setting:Add, Text, xp-150 yp+40, Background Color:
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 vg_WinColor, %g_WinColor%||
    Gui, Setting:Add, Text, xp-150 yp+40, Background Picture:
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 vg_Background, %g_Background%||NO PICTURE|DEFAULT PICTURE|C:\Path\ExamplePicture.jpg

    Gui, Setting:Tab, 4 ; Hotkey Tab
    Gui, Setting:Add, GroupBox, w500 h120, Activate
    Gui, Setting:Add, Text, xp+10 yp+25 , Primary Hotkey :
    Gui, Setting:Add, Hotkey, xp+250 yp-4 w230 vg_GlobalHotkey1,%g_GlobalHotkey1%
    Gui, Setting:Add, Text, xp-250 yp+35 , Secondary Hotkey :
    Gui, Setting:Add, Hotkey, xp+250 yp-4 w230 vg_GlobalHotkey2,%g_GlobalHotkey2%
    Gui, Setting:Add, Link, xp-250 yp+35 gResetHotkey, You can set another hotkey as a secondary hotkey (<a id="Reset">Reset to Default</a>)
    Gui, Setting:Add, GroupBox, xp-10 yp+40 w500 h55, Command Hotkey:
    Gui, Setting:Add, Text, xp+10 yp+25 , Execute Command:
    Gui, Setting:Add, Text, xp+150 yp , ALT + No.
    Gui, Setting:Add, Text, xp+105 yp, Select Command: 
    Gui, Setting:Add, Text, xp+130 yp, Ctrl + No.
    Gui, Setting:Add, GroupBox, xp-395 yp+40 w500 h130, Action Hotkey:
    Gui, Setting:Add, Text, xp+10 yp+25 , Hotkey 1: 
    Gui, Setting:Add, Hotkey, xp+150 yp-5 w80 vg_Hotkey1, %g_Hotkey1%
    Gui, Setting:Add, Text, xp+100 yp+5, Toggle Action: 
    Gui, Setting:Add, Edit, xp+110 yp-5 r1 w120 vg_Trigger1, %g_Trigger1%
    Gui, Setting:Add, Text, xp-360 yp+40 , Hotkey 2: 
    Gui, Setting:Add, Hotkey, xp+150 yp-5 w80 vg_Hotkey2, %g_Hotkey2%
    Gui, Setting:Add, Text, xp+100 yp+5, Toggle Action: 
    Gui, Setting:Add, Edit, xp+110 yp-5 r1 w120 vg_Trigger2, %g_Trigger2%
    Gui, Setting:Add, Text, xp-360 yp+40 , Hotkey 3: 
    Gui, Setting:Add, Hotkey, xp+150 yp-5 w80 vg_Hotkey3, %g_Hotkey3%
    Gui, Setting:Add, Text, xp+100 yp+5, Toggle Action: 
    Gui, Setting:Add, Edit, xp+110 yp-5 r1 w120 vg_Trigger3, %g_Trigger3%

    Gui, Setting:Tab, 5 ; LISTARTY TAB
    Gui, Setting:Add, GroupBox, w500 h125, Listary Quick-Switch
    Gui, Setting:Add, Text, xp+10 yp+25 , File Manager Title:
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 Sort vg_FileMgrID, %g_FileMgrID%||ahk_class CabinetWClass|ahk_class CabinetWClass, ahk_class TTOTAL_CMD
    Gui, Setting:Add, Text, xp-150 yp+40, Open/Save Dialog Title:
    Gui, Setting:Add, Combobox, xp+150 yp-5 w330 Sort vg_DialogWin, %g_DialogWin%||ahk_class #32770
    Gui, Setting:Add, Text, xp-150 yp+40, Exclude Windows Title: 
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 Sort vg_ExcludeWin, %g_ExcludeWin%||ahk_class SysListView32|ahk_class SysListView32, ahk_exe Explorer.exe|ahk_class SysListView32, ahk_exe Explorer.exe, ahk_exe Totalcmd64.exe, AutoCAD LT Alert
    Gui, Setting:Add, GroupBox, xp-160 yp+45 w500 h125, Hotkey for Switch Open/Save dialog path to
    Gui, Setting:Add, Text, xp+10 yp+30, Total Commander's dir:
    Gui, Setting:Add, Hotkey, xp+150 yp-5 w330 vg_TotalCMDDir, %g_TotalCMDDir%
    Gui, Setting:Add, Text, xp-150 yp+40, Windows Explorer's dir:
    Gui, Setting:Add, Hotkey, xp+150 yp-5 w330 vg_ExplorerDir, %g_ExplorerDir%
    Gui, Setting:Add, CheckBox, xp-150 yp+40 vg_AutoSwitchDir checked%g_AutoSwitchDir%, Auto Switch Dir

    Gui, Setting:Tab, 6 ; ABOUT TAB
    Gui, Setting:Font, S10, Segoe UI Semibold
    Gui, Setting:Add, Edit, w500 h420 ReadOnly -WantReturn -Wrap,
    (Ltrim
    ALTRun
    Website: https://github.com/zhugecaomao/ALTRun
    Setting Notes: https://github.com/zhugecaomao/ALTRun/wiki#notes-for-altrun-options

    An effective launcher for Windows, an AutoHotkey open-source project. 
    It provides a streamlined and efficient way to find anything on your 
    system and launch any application in your way.

    1. Pure portable software, not write anything into Registry.
    2. Small size (< 100KB), low resource usage (< 5MB RAM), and high performance.
    3. User-friendly interface, highly customizable from the Options menu
    4. SendTo Menu allows you to create commands quickly and easily.
    5. Multi-Hotkey setup allowed.
    6. Integrated with Total Commander and Everything
    7. Smart Rank - Atuo adjusts command priority (rank) based on frequency of use.
    8. Smart Match - Fuzzy and Smart matching and filtering result
    9. Listary Quick Switch Dir function
    Many more functions...

    ------------------------------------------------------------------------
    F1        		ALTRun Help Index
    F2        		Open Setting Config window
    F3        		Edit current command (.ini) directly
    F4        		Edit user-defined commands (.ini) directly
    Enter   		Run current command
    Esc     		Clear input / Close window
    Up / Down   	Move to Previous / Next command
    Alt + Space 	Show / Hide window
    Alt + F4		Exit
    Alt + No.  		Run specific command
    Ctrl + No.   	Select specific command
    Ctrl + +		Increase rank of current command
    Ctrl + -		Decrease rank of current command
    Ctrl + I		Reindex file search database
    Ctrl + Q		Reload ALTRun
    Ctrl + D		Open current command dir with TC / File Explorer
    Space   		Start with Space - Search with Everything
    >        		Start with ">" - Run with CMD
    +        		Start with "+" - Create new Command
    No Result		Enter to add as a new command
    )
    
    Gui, Setting:Tab                                                    ; 后续添加的控件将不属于前面那个选项卡控件

    Hotkey, %g_GlobalHotkey1%, Off
    Hotkey, %g_GlobalHotkey2%, Off

    Gui, Setting:Add, Button, Default x350 w80, OK
    Gui, Setting:Add, Button, xp+90 yp w80, Cancel
    Gui, Setting:Show,, %g_OptionsWinName%
}

ResetHotkey()
{
    GuiControl,, g_GlobalHotkey1, !Space
    GuiControl,, g_GlobalHotkey2, !R
}

SettingButtonOK()                                                       ; 设置选项窗口 - 按钮动作
{
    SAVECONFIG()
    Reload
}

SettingGuiEscape()
{
    SettingGuiClose()
}

SettingButtonCancel()
{
    SettingGuiClose()
}

SettingGuiClose()
{
    Hotkey, %g_GlobalHotkey1%, On
    Hotkey, %g_GlobalHotkey2%, On
    Gui, Setting:Destroy
}

LOADCONFIG(Arg)                                                         ; 加载主配置文件
{
    Log.Debug("Loading configuration...Arg=" Arg)
    
    if (Arg = "config" or Arg = "initialize" or Arg = "all")
    {
        For key, description in KEYS_CONFIG                              ; Read config section
        {
            IniRead, g_%key%, %g_IniFile%, %SEC_CONFIG%, %key%, % g_%key%
        }
    
        Loop Parse, KEYLIST_CONFIG, `,                                  ; Read config section
        {
            IniRead, g_%A_LoopField%, %g_IniFile%, %SEC_CONFIG%, %A_LoopField%, % g_%A_LoopField%
        }

        Loop Parse, KEYLIST_HOTKEY, `,                                  ; Read Hotkey section
        {
            IniRead, g_%A_LoopField%, %g_IniFile%, %SEC_HOTKEY%, %A_LoopField%, % g_%A_LoopField%
        }

        Loop Parse, KEYLIST_GUI, `,                                     ; Read GUI section
        {
            IniRead, g_%A_LoopField%, %g_IniFile%, %SEC_GUI%, %A_LoopField%, % g_%A_LoopField%
        }
        
        if (g_Background = "Default")
        {
            Extract_BG(A_Temp "\ALTRun.jpg")
            g_BGPicture := A_Temp "\ALTRun.jpg"
        }
        else
        {
            g_BGPicture := g_Background
        }
    }

    if (Arg = "commands" or Arg = "initialize" or Arg = "all")          ; Built-in command initialize
    {
        IniRead, DFTCMDSEC, %g_IniFile%, %SEC_DFTCMD%
        if (DFTCMDSEC = "")
        {
            IniWrite, 
            (Ltrim
            ; Build-in Commands (High Priority, DO NOT Edit)
            ;
            Func | Help | ALTRun Help Index (F1)=100
            Func | Options | ALTRun Options Preference Settings (F2)=100
            Func | Reload | ALTRun Reload=100
            Func | CmdMgr | New Command=100
            Func | UserCommandList | ALTRun User-defined command (F4)=100
            Func | Reindex | Reindex search database=100
            Func | Everything | Search by Everything=100
            Func | RunPTTools | PT Tools (AHK)=100
            Func | AhkRun | Run Command use AutoHotkey Run=100
            Func | CmdRun | Run Command use CMD=100
            Func | Google | Search Clipboard or Input by Google=100
            Func | Bing | Search Clipboard or Input by Bing=100
            Func | EmptyRecycle | Empty Recycle Bin=100
            Func | TurnMonitorOff | Turn off Monitor, Close Monitor=100
            Func | MuteVolume | Mute Volume=100
            Dir | A_ScriptDir | ALTRun Program Dir=100
            Dir | A_Startup | Current User Startup Dir=100
            Dir | A_StartupCommon | All User Startup Dir=100
            Dir | A_ProgramsCommon | Windowns Search.Index.Cortana Dir=100
            File | %Temp%\ALTRun.log | ALTRun Log File=100
            ;
            ; Control Panel Commands
            ;
            Ctrl | Control | Control Panel=66
            Ctrl | wf.msc | Windows Defender Firewall with Advanced Security=66
            Ctrl | Control intl.cpl | Region and Language Options=66
            Ctrl | Control firewall.cpl | Windows Defender Firewall=66
            Ctrl | Control access.cpl | Ease of Access Centre=66
            Ctrl | Control appwiz.cpl | Programs and Features=66
            Ctrl | Control sticpl.cpl | Scanners and Cameras=66
            Ctrl | Control sysdm.cpl | System Properties=66
            Ctrl | Control joy.cpl | Game Controllers=66
            Ctrl | Control Mouse | Mouse Properties=66
            Ctrl | Control desk.cpl | Display=66
            Ctrl | Control mmsys.cpl | Sound=66
            Ctrl | Control ncpa.cpl | Network Connections=66
            Ctrl | Control powercfg.cpl | Power Options=66
            Ctrl | Control timedate.cpl | Date and Time=66
            Ctrl | Control admintools | Windows Tools=66
            Ctrl | Control desktop | Personalisation=66
            Ctrl | Control folders | File Explorer Options=66
            Ctrl | Control fonts | Fonts=66
            Ctrl | Control inetcpl.cpl,,4 | Internet Properties=66
            Ctrl | Control printers | Devices and Printers=66
            Ctrl | Control userpasswords | User Accounts=66
            Ctrl | taskschd.msc | Task Scheduler=66
            Ctrl | devmgmt.msc | Device Manager=66
            Ctrl | eventvwr.msc | Event Viewer=66
            Ctrl | compmgmt.msc | Computer Manager=66
            Ctrl | taskmgr.exe | Task Manager=66
            Ctrl | calc.exe | Calculator=66
            Ctrl | mspaint.exe | Paint=66
            Ctrl | cmd.exe | DOS / CMD=66
            Ctrl | regedit.exe | Registry Editor=66
            Ctrl | write.exe | Write=66
            Ctrl | cleanmgr.exe | Disk Space Clean-up Manager=66
            Ctrl | gpedit.msc | Group Policy=66
            Ctrl | comexp.msc | Component Services=66
            Ctrl | diskmgmt.msc | Disk Management=66
            Ctrl | dxdiag.exe | Directx Diagnostic Tool=66
            Ctrl | lusrmgr.msc | Local Users and Groups=66
            Ctrl | msconfig.exe | System Configuration=66
            Ctrl | perfmon.exe /Res | Resources Monitor=66
            Ctrl | perfmon.exe | Performance Monitor=66
            Ctrl | winver.exe | About Windows=66
            Ctrl | services.msc | Services=66
            Ctrl | netplwiz | User Accounts=66
            ), %g_IniFile%, %SEC_DFTCMD%
            IniRead, DFTCMDSEC, %g_IniFile%, %SEC_DFTCMD%
        }

        IniRead, USERCMDSEC, %g_IniFile%, %SEC_USERCMD%
        if (USERCMDSEC = "")
        {
            IniWrite, 
            (Ltrim
            ; User-Defined Commands (High priority, edit as desired)
            ; Command type: File, Dir, CMD, Func, URL
            ; Type | Command | Comments=Rank
            ;
            Dir | `%AppData`%\Microsoft\Windows\SendTo | Windows SendTo Dir=100
            Dir | `%OneDriveConsumer`% | OneDrive Personal Dir=100
            Dir | `%OneDriveCommercial`% | OneDrive Business Dir=100
            CMD | ipconfig | Show IP Address(CMD type will run with cmd.exe, auto pause after run)=100
            URL | www.google.com | Google=100
            File | C:\OneDrive\Apps\TotalCMD64\TOTALCMD64.exe
            ), %g_IniFile%, %SEC_USERCMD%
            IniRead, USERCMDSEC, %g_IniFile%, %SEC_USERCMD%
        }

        IniRead, INDEXSEC, %g_IniFile%, %SEC_INDEX%                     ; Read whole section SEC_INDEX (Index database)
        if (INDEXSEC = "")
        {
            MsgBox, 4160, %g_WinName%, ALTRun is going to initialize for the first time running...`n`nConfig software and build the index database for search.`n`nAuto initialize in 30 seconds or click OK now., 30
            Reindex()
        }
        Return DFTCMDSEC "`n" USERCMDSEC "`n" INDEXSEC
    }
    Return
}

SAVECONFIG() {
    Gui, Setting:Submit

    For key, description in KEYS_CONFIG
    {
        Checked := (A_Index = LV_GetNext(A_Index-1, "C")) ? 1 : 0
        IniWrite, %Checked%, %g_IniFile%, %SEC_CONFIG%, %key%
    }

    Loop Parse, KEYLIST_CONFIG, `,
        IniWrite, % g_%A_LoopField%, %g_IniFile%, %SEC_CONFIG%, %A_LoopField%
    
    Loop Parse, KEYLIST_GUI, `,
        IniWrite, % g_%A_LoopField%, %g_IniFile%, %SEC_GUI%, %A_LoopField%
    
    Loop Parse, KEYLIST_HOTKEY, `,
        IniWrite, % g_%A_LoopField%, %g_IniFile%, %SEC_HOTKEY%, %A_LoopField%
    
    Log.Debug("Saving config...")
    Return
}

;=============================================================
; Resources File - Background picture
;=============================================================
Extract_BG(_Filename)
{
	Static Out_Data
    VarSetCapacity(TD, 4206 * 2)
    TD :="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"
    
    VarSetCapacity(Out_Data, Bytes := 3070, 0)
    DllCall("Crypt32.dll\CryptStringToBinary" "W", "Ptr", &TD, "UInt", 0, "UInt", 1, "Ptr", &Out_Data, "UIntP", Bytes, "Int", 0, "Int", 0, "CDECL Int")
	
    FileExist(_Filename)
		FileDelete, %_Filename%
	
	h := DllCall("CreateFile", "Ptr", &_Filename, "Uint", 0x40000000, "Uint", 0, "UInt", 0, "UInt", 4, "Uint", 0, "UInt", 0)
	, DllCall("WriteFile", "Ptr", h, "Ptr", &Out_Data, "UInt", 3070, "UInt", 0, "UInt", 0)
	, DllCall("CloseHandle", "Ptr", h)
}

;=============================================================
; Some Built-in Functions
;=============================================================
CmdRun()
{
    RunWithCmd(Arg)
}

AhkRun()
{
    Global
    Run, %Arg%
}

TurnMonitorOff()                                                        ; 关闭显示器:
{
    SendMessage, 0x112, 0xF170, 2,, Program Manager                     ; 0x112 is WM_SYSCOMMAND, 0xF170 is SC_MONITORPOWER, 使用 -1 代替 2 来打开显示器, 使用 1 代替 2 来激活显示器的节能模式.
}

EmptyRecycle()
{
    MsgBox, 4, %g_WinName%, Do you really want to empty the Recycle Bin?
    IfMsgBox Yes
    {
        FileRecycleEmpty,
    }
}

MuteVolume()
{
    SoundSet, MUTE
}

Google()
{
    Global
    word := Arg = "" ? clipboard : Arg
    Run, https://www.google.com/search?q=%word%&newwindow=1
}

Bing()
{
    Global
    word := Arg = "" ? clipboard : Arg
    Run, http://cn.bing.com/search?q=%word%
}

Everything()
{
    Run, %g_Everything% -s "%Arg%",, UseErrorLevel
    if ErrorLevel
        MsgBox, % "Everything software not found.`n`nPlease check ALTRun setting and Everything program file."
}

;=======================================================================
; Library - Eval (Math Expression)
;=======================================================================

/* -test cases
MsgBox % Eval("1e1")                                               ; 10
MsgBox % Eval("0x1E")                                              ; 30
MsgBox % Eval("ToBin(35)")                                         ; 100011
MsgBox % Eval("$b 35")                                             ; 0100011
MsgBox % Eval("'10010")                                            ; -14
MsgBox % Eval("2>3 ? 9 : 7")                                       ; 7
MsgBox % Eval("$2E 1e3 -50.0e+0 + 100.e-1")                        ; 9.60E+002
MsgBox % Eval("fact(x) := x < 2 ? 1 : x*fact(x-1); fact(5)")       ; 120
MsgBox % Eval("f(ab):=sqrt(ab)/ab; y:=f(2); ff(y):=y*(y-1)/2/x; x := 2; y+ff(3)/f(16)") ; 6.70711
MsgBox % Eval("x := qq:1; x := 5*x; y := x+1")                     ; 6 [if y empty, x := 1...]
MsgBox % Eval("x:=-!0; x<0 ? 2*x : sqrt(x)")                       ; -2
MsgBox % Eval("tan(atan(atan(tan(1))))-exp(sqrt(1))")              ; -1.71828
MsgBox % Eval("---2+++9 + ~-2 --1 -2*-3")                          ; 15
MsgBox % Eval("x1:=1; f1:=sin(x1)/x1; y:=2; f2:=sin(y)/y; f1/f2")  ; 1.85082
MsgBox % Eval("Round(fac(10)/fac(5)**2) - (10 choose 5) + Fib(8)") ; 21
MsgBox % Eval("1 min-1 min-2 min 2")                               ; -2
MsgBox % Eval("(-1>>1<=9 && 3>2)<<2>>1")                           ; 2
MsgBox % Eval("(1 = 1) + (2<>3 || 2 < 1) + (9>=-1 && 3>2)")        ; 3
MsgBox % Eval("$b6 -21/3")                                         ; 111001
MsgBox % Eval("$b ('1001 << 5) | '01000")                          ; 100101000
MsgBox % Eval("$0 194*lb/1000")                                    ; 88 [Kg]
MsgBox % Eval("$x ~0xfffffff0 & 7 | 0x100 << 2")                   ; 0x407
MsgBox % Eval("- 1 * (+pi -((3%5))) +pi+ 1-2 + e-ROUND(abs(sqrt(floor(2)))**2)-e+pi $9") ; 3.141592654
t := A_TickCount
Loop 1000
   r := Eval("x:=" A_Index/1000 ";atan(x)-exp(sqrt(x))")           ; simulated plot
t := A_TickCount - t
MsgBox Result = %r%`nTime = %t%                                    ; -1.93288: ~400 ms [on Inspiron 9300]
*/

Eval(x) {                              ; non-recursive PRE/POST PROCESSING: I/O forms, numbers, ops, ";"
   Local FORM, FormF, FormI, i, W, y, y1, y2, y3, y4
   FormI := A_FormatInteger, FormF := A_FormatFloat

   SetFormat Integer, D                ; decimal intermediate results!
   RegExMatch(x, "\$(b|h|x|)(\d*[eEgG]?)", y)
   FORM := y1, W := y2                 ; HeX, Bin, .{digits} output format
   SetFormat FLOAT, 0.16e              ; Full intermediate float precision
   StringReplace x, x, %y%             ; remove $..
   Loop
      If RegExMatch(x, "i)(.*)(0x[a-f\d]*)(.*)", y)
         x := y1 . y2+0 . y3           ; convert hex numbers to decimal
      Else Break
   x := RegExReplace(x,"(^|[^.\d])(\d+)(e|E)","$1$2.$3")                ; add missing '.' before E (1e3 -> 1.e3)
   x := RegExReplace(x,"(\d*\.\d*|\d)([eE][+-]?\d+)","‘$1$2’")          ; literal scientific numbers between ‘ and ’ chars

   StringReplace x, x,`%, \, All            ; %  -> \ (= MOD)
   StringReplace x, x, **,@, All            ; ** -> @ for easier process
   StringReplace x, x, ^,@, All             ; ^ -> @ for easier process
   StringReplace x, x, +, ±, All            ; ± is addition
   x := RegExReplace(x,"(‘[^’]*)±","$1+")   ; ...not inside literal numbers
   StringReplace x, x, -, ¬, All            ; ¬ is subtraction
   x := RegExReplace(x,"(‘[^’]*)¬","$1-")   ; ...not inside literal numbers

   Loop Parse, x, `;
      y := Eval1(A_LoopField)          ; work on pre-processed sub expressions
                                       ; return result of last sub-expression (numeric)
   If FORM = b                         ; convert output to binary
      y := W ? ToBinW(Round(y),W) : ToBin(Round(y))
   Else If (FORM="h" or FORM="x") {
      SetFormat Integer, Hex           ; convert output to hex
      y := Round(y) + 0
   }
   Else {
      W := W="" ? "0.6g" : "0." . W    ; Set output form, Default = 6 decimal places
      SetFormat FLOAT, %W%
      y += 0.0
   }
   SetFormat Integer, %FormI%          ; restore original formats
   SetFormat FLOAT,   %FormF%
   Return y
}

Eval1(x) {                             ; recursive PREPROCESSING of :=, vars, (..) [decimal, no ";"]
   Local i, y, y1, y2, y3
   If RegExMatch(x, "(\S*?)\((.*?)\)\s*:=\s*(.*)", y) {                 ; save function definition: f(x) := expr
      f%y1%__X := y2, f%y1%__F := y3
      Return
   }
                                       ; execute leftmost ":=" operator of a := b := ...
   If RegExMatch(x, "(\S*?)\s*:=\s*(.*)", y) {
      y := "x" . y1                    ; user vars internally start with x to avoid name conflicts
      Return %y% := Eval1(y2)
   }
                                       ; here: no variable to the left of last ":="
   x := RegExReplace(x,"([\)’.\w]\s+|[\)’])([a-z_A-Z]+)","$1«$2»")  ; op -> «op»

   x := RegExReplace(x,"\s+")          ; remove spaces, tabs, newlines

   x := RegExReplace(x,"([a-z_A-Z]\w*)\(","'$1'(") ; func( -> 'func'( to avoid atan|tan conflicts

   x := RegExReplace(x,"([a-z_A-Z]\w*)([^\w'»’]|$)","%x$1%$2") ; VAR -> %xVAR%
   x := RegExReplace(x,"(‘[^’]*)%x[eE]%","$1e") ; in numbers %xe% -> e
   x := RegExReplace(x,"‘|’")          ; no more need for number markers
   Transform x, Deref, %x%             ; dereference all right-hand-side %var%-s

   Loop {                              ; find last innermost (..)
      If RegExMatch(x, "(.*)\(([^\(\)]*)\)(.*)", y)
         x := y1 . Eval@(y2) . y3      ; replace (x) with value of x
      Else Break
   }
   Return Eval@(x)
}

Eval@(x) {                             ; EVALUATE PRE-PROCESSED EXPRESSIONS [decimal, NO space, vars, (..), ";", ":="]
   Local i, y, y1, y2, y3, y4

   If x is number                      ; no more operators left
      Return x
                                       ; execute rightmost ?,: operator
   RegExMatch(x, "(.*)(\?|:)(.*)", y)
   IfEqual y2,?,  Return Eval@(y1) ? Eval@(y3) : ""
   IfEqual y2,:,  Return ((y := Eval@(y1)) = "" ? Eval@(y3) : y)

   StringGetPos i, x, ||, R            ; execute rightmost || operator
   IfGreaterOrEqual i,0, Return Eval@(SubStr(x,1,i)) || Eval@(SubStr(x,3+i))
   StringGetPos i, x, &&, R            ; execute rightmost && operator
   IfGreaterOrEqual i,0, Return Eval@(SubStr(x,1,i)) && Eval@(SubStr(x,3+i))
                                       ; execute rightmost =, <> operator
   RegExMatch(x, "(.*)(?<![\<\>])(\<\>|=)(.*)", y)
   IfEqual y2,=,  Return Eval@(y1) =  Eval@(y3)
   IfEqual y2,<>, Return Eval@(y1) <> Eval@(y3)
                                       ; execute rightmost <,>,<=,>= operator
   RegExMatch(x, "(.*)(?<![\<\>])(\<=?|\>=?)(?![\<\>])(.*)", y)
   IfEqual y2,<,  Return Eval@(y1) <  Eval@(y3)
   IfEqual y2,>,  Return Eval@(y1) >  Eval@(y3)
   IfEqual y2,<=, Return Eval@(y1) <= Eval@(y3)
   IfEqual y2,>=, Return Eval@(y1) >= Eval@(y3)
                                       ; execute rightmost user operator (low precedence)
   RegExMatch(x, "i)(.*)«(.*?)»(.*)", y)
   If IsFunc(y2)
      Return %y2%(Eval@(y1),Eval@(y3)) ; predefined relational ops

   StringGetPos i, x, |, R             ; execute rightmost | operator
   IfGreaterOrEqual i,0, Return Eval@(SubStr(x,1,i)) | Eval@(SubStr(x,2+i))
   StringGetPos i, x, ^, R             ; execute rightmost ^ operator
   IfGreaterOrEqual i,0, Return Eval@(SubStr(x,1,i)) ^ Eval@(SubStr(x,2+i))
   StringGetPos i, x, &, R             ; execute rightmost & operator
   IfGreaterOrEqual i,0, Return Eval@(SubStr(x,1,i)) & Eval@(SubStr(x,2+i))
                                       ; execute rightmost <<, >> operator
   RegExMatch(x, "(.*)(\<\<|\>\>)(.*)", y)
   IfEqual y2,<<, Return Eval@(y1) << Eval@(y3)
   IfEqual y2,>>, Return Eval@(y1) >> Eval@(y3)
                                       ; execute rightmost +- (not unary) operator
   RegExMatch(x, "(.*[^!\~±¬\@\*/\\])(±|¬)(.*)", y) ; lower precedence ops already handled
   IfEqual y2,±,  Return Eval@(y1) + Eval@(y3)
   IfEqual y2,¬,  Return Eval@(y1) - Eval@(y3)
                                       ; execute rightmost */% operator
   RegExMatch(x, "(.*)(\*|/|\\)(.*)", y)
   IfEqual y2,*,  Return Eval@(y1) * Eval@(y3)
   IfEqual y2,/,  Return Eval@(y1) / Eval@(y3)
   IfEqual y2,\,  Return Mod(Eval@(y1),Eval@(y3))
                                       ; execute rightmost power
   StringGetPos i, x, @, R
   IfGreaterOrEqual i,0, Return Eval@(SubStr(x,1,i)) ** Eval@(SubStr(x,2+i))
                                       ; execute rightmost function, unary operator
   If !RegExMatch(x,"(.*)(!|±|¬|~|'(.*)')(.*)", y)
      Return x                         ; no more function (y1 <> "" only at multiple unaries: --+-)
   IfEqual y2,!,Return Eval@(y1 . !y4) ; unary !
   IfEqual y2,±,Return Eval@(y1 .  y4) ; unary +
   IfEqual y2,¬,Return Eval@(y1 . -y4) ; unary - (they behave like functions)
   IfEqual y2,~,Return Eval@(y1 . ~y4) ; unary ~
   If IsFunc(y3)
      Return Eval@(y1 . %y3%(y4))      ; built-in and predefined functions(y4)
   Return Eval@(y1 . Eval1(RegExReplace(f%y3%__F, f%y3%__X, y4))) ; LAST: user defined functions
}

ToBin(n) {      ; Binary representation of n. 1st bit is SIGN: -8 -> 1000, -1 -> 1, 0 -> 0, 8 -> 01000
   if (n == "")
   {
       return 0
   }
   Return n=0||n=-1 ? -n : ToBin(n>>1) . n&1
}
ToBinW(n,W=8) { ; LS W-bits of Binary representation of n
   Loop %W%     ; Recursive (slower): Return W=1 ? n&1 : ToBinW(n>>1,W-1) . n&1
      b := n&1 . b, n >>= 1
   Return b
}

Class Logger                                                            ; Logger library
{
    __New(filename)
    {
        this.filename := filename
    }

    Debug(Msg)
    {
        if (g_Logging) 
        {
            FileAppend, % "[" A_Now "] " Msg "`n", % this.filename
        }
    }
}