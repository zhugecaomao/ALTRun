;==============================================================
; ALTRun - An effective launcher for Windows.
; https://github.com/zhugecaomao/ALTRun
;==============================================================
#Requires AutoHotkey v1.1+
#NoEnv
#SingleInstance, Force
#NoTrayIcon
#Persistent
#Warn All, OutputDebug

FileEncoding, UTF-8
SendMode, Input
SetWorkingDir %A_ScriptDir%

;===================================================
; 声明全局变量
;===================================================
Global Arg, Log:= New Logger(A_Temp "\ALTRun.log") ; Arg: 用来调用管道的完整参数（所有列）
, g_CurrentCommandList := Object()                 ; 当前匹配到的所有命令
, g_Commands:= Object()                            ; 所有命令
, g_Fallback:= Object()                            ; 当搜索无结果时使用的命令
, g_History := Object()                            ; 历史命令
, g_SEC     := {Config:"Config",Gui:"Gui",DftCMD:"DefaultCommand",UserCMD:"UserCommand",Fallback:"FallbackCommand",Hotkey:"Hotkey",History:"History",Index:"Index",Usage:"Usage"}
, g_CONFIG  := {AutoStartup:1,EnableSendTo:1,InStartMenu:1,ShowTrayIcon:1,HideOnLostFocus:1,AlwaysOnTop:1,EscClearInput:1,KeepInput:1,ShowIcon:1
            ,SendToGetLnk:1,SaveHistory:1,SaveLog:1,MatchPath:0,ShowGrid:0,ShowHdr:1,SmartRank:1,SmartMatch:1,MatchAny:1,ShowTheme:1,ShowHint:1
            ,ShowRunCount:1,ShowStatusBar:1,ShowBtnRun:1,ShowBtnOpt:1,ShowDirName:1,RunCount:0,HistoryLen:15,AutoSwitchDir:0,Editor:"Notepad.exe"
            ,FileMgr:"Explorer.exe",IndexDir:"A_ProgramsCommon,A_StartMenu,C:\Other\Index\Location",IndexType:"*.lnk,*.exe",IndexExclude:"Uninstall *"
            ,Everything:"C:\Apps\Everything\Everything.exe",DialogWin:"ahk_class #32770",FileMgrID:"ahk_class CabinetWClass, ahk_class TTOTAL_CMD"
            ,ExcludeWin:"ahk_class SysListView32, ahk_exe Explorer.exe, AutoCAD"}
, g_HOTKEY  := {GlobalHotkey1:"!Space",GlobalHotkey2:"!R",Hotkey1:"^o",Trigger1:"Options",Hotkey2:"",Trigger2:"---",Hotkey3:"",Trigger3:"---",TotalCMDDir:"^g",ExplorerDir:"^e"}
, g_GUI     := {ListRows:9,ColWidth:"40,60,430,340",FontName:"Segoe UI",FontSize:10,FontColor:"Default",WinWidth:900,WinHeight:330,CtrlColor:"Default",WinColor:"Silver",Background:"DEFAULT"}
, g_CHKLV   := {AutoStartup:"Launch on Windows startup",EnableSendTo:"Enable the SendTo menu",InStartMenu:"Enable the Start menu"
            ,ShowTrayIcon:"Show Tray Icon in the system taskbar",HideOnLostFocus:"Close window on losing focus",AlwaysOnTop:"Always stay on top"
            ,EscClearInput:"Press [ESC] to clear input, press again to close window (Untick:close directly)",KeepInput:"Keep last input and search result on close"
            ,ShowIcon:"Show Icon of file, folder or apps in the command result list",SendToGetLnk:"Retrieve .lnk target on SendTo"
            ,SaveHistory:"Save History - Commands executed with arg",SaveLog:"Save Log - App running and debug information",MatchPath:"Match full path on search"
            ,ShowGrid:"Show Grid in command list",ShowHdr:"Show Header in command list",SmartRank:"Smart Rank - Auto adjust command priority (rank) based on use frequency"
            ,SmartMatch:"Smart Match - Fuzzy and Smart matching and filtering result",MatchAny:"Match from any position of the string"
            ,ShowTheme:"Show Theme - Software skin and background picture",ShowHint:"Show Hints and Tips in the bottom status bar"
            ,ShowRunCount:"Show RunCount - Command running times in the status bar",ShowStatusBar:"Show Status Bar"
            ,ShowBtnRun:"Show [Run] Button on main window",ShowBtnOpt:"Show [Options] Button on main window"
            ,ShowDirName:"Show Shorten Dir - Show dir name only instead of full path in the result"} ; Options - General - CheckedListview
, g_RUNTIME := {Ini:A_ScriptDir "\" A_ComputerName ".ini",WinName:"ALTRun - Ver 2024.12",BGPic:"",WinHide:"",UseDisplay:"",UseFallback:"" ; 程序运行需要的临时全局变量, 不需要用户参与修改, 不读写入ini
            ,CurrentCMD:"",Input:"",FuncList:"",OneDrive:EnvGet("OneDrive"),OneDriveConsumer:EnvGet("OneDriveConsumer"),UsageToday:0,AppStartDate: A_YYYY . A_MM . A_DD
            ,OneDriveCommercial:EnvGet("OneDriveCommercial"),MaxVal:0} ; OneDrive Personal/Business Environment Variables (due to #NoEnv)
, g_Hints   := ["It's better to show me by press hotkey (Default is ALT + Space)","ALT + Space = Show / Hide window","Alt + F4 = Exit"
            ,"Esc = Clear input / Close window","Enter = Run current command","Alt + No. = Run specific command","Start with + = New Command"
            ,"Ctrl + No. = Select specific command","F1 = ALTRun Help Index","F2 = Open Setting Config window"
            ,"F3 = Edit current command (.ini) directly","F4 = Edit user-defined commands (.ini) directly"
            ,"Arrow Up / Down = Move to Previous / Next command","Ctrl+Q = Reload ALTRun","Ctrl+'+' = Increase rank of current command"
            ,"Ctrl+'-' = Decrease rank of current command","Ctrl+I = Reindex file search database","Start with space = Search file by Everything"
            ,"Ctrl+D = Open current command dir with TC / File Explorer","Command priority (rank) will auto adjust based on frequency"]

Log.Debug("///// ALTRun is starting /////")
LOADCONFIG("initialize")                                                ; Load ini config, IniWrite will create it if not exist

; For key, value in g_RUNTIME
;     {
;         OutputDebug, % key " = " g_RUNTIME[key]
;     }
;=============================================================
; Create ContextMenu and TrayMenu
;=============================================================
ContextMenu := ["Run`tEnter,LVContextMenu,shell32.dll,-25","Locate`tCtrl+D,OpenCurrentFileDir,shell32.dll,-4","Copy,LVContextMenu,shell32.dll,-243"
    ,"","New,CmdMgr,shell32.dll,-1","Edit`tF3,EditCurrentCommand,shell32.dll,-16775","User Defined`tF4,UserCommandList,shell32.dll,-44"]

For index, MenuItem in ContextMenu {
    If (MenuItem = "")
        Menu, LV_ContextMenu, Add
    Else {
        Item := StrSplit(MenuItem, ",")
        Menu, LV_ContextMenu, Add, % Item[1], % Item[2]
        Menu, LV_ContextMenu, Icon, % Item[1], % Item[3], % Item[4]
    }
}

if (g_CONFIG.ShowTrayIcon)
{
    TrayMenu := ["Show,ToggleWindow,Shell32.dll,-25","","Options `tF2,Options,Shell32.dll,-16826"
    ,"ReIndex `tCtrl+I,Reindex,Shell32.dll,-16776","Help `tF1,Help,Shell32.dll,-24",""
    ,"Script Info,ScriptInfo,imageres.dll,-150","AHK Manual,AHKManual,Shell32.dll,-512",""
    ,"Reload `tCtrl+Q,Reload,imageres.dll,-5311","Exit `tAlt+F4,Exit,imageres.dll,-98"]

    Menu, Tray, NoStandard
    Menu, Tray, Icon
    Menu, Tray, Icon, Shell32.dll, -25                                  ; Index of icon changes between Windows versions, refer to the icon by resource ID for consistency
    For Index, MenuItem in TrayMenu
    {
        Item := StrSplit(MenuItem, ",") ; Item[1,2,3,4] <-> Name,Func,Icon,IconNo
        Menu, Tray, Add, % Item[1], % Item[2]
        Menu, Tray, Icon, % Item[1], % Item[3], % Item[4]
    }
    Menu, Tray, Tip, % g_RUNTIME.WinName
    Menu, Tray, Default, Show
    Menu, Tray, Click, 1
}
;=============================================================
; Load commands database and command history
; Update "SendTo", "Startup", "StartMenu" lnk
;=============================================================
LoadCommands()
LoadHistory()

Log.Debug("Updating 'SendTo' setting..." UpdateSendTo(g_CONFIG.EnableSendTo))
Log.Debug("Updating 'Startup' setting..." UpdateStartup(g_CONFIG.AutoStartup))
Log.Debug("Updating 'StartMenu' setting..." UpdateStartMenu(g_CONFIG.InStartMenu))

;=============================================================
; 主窗口配置代码
;=============================================================
AlwaysOnTop  := g_CONFIG.AlwaysOnTop ? "+AlwaysOnTop" : ""
ShowGrid     := g_CONFIG.ShowGrid ? "Grid" : ""
ShowHdr      := g_CONFIG.ShowHdr ? "" : "-Hdr"
LV_H         := g_GUI.WinHeight - 43 - 3 * g_GUI.FontSize
LV_W         := g_GUI.WinWidth - 24
Input_W      := LV_W - g_CONFIG.ShowBtnRun * 90 - g_CONFIG.ShowBtnOpt * 90
Enter_W      := g_CONFIG.ShowBtnRun * 80
Enter_X      := g_CONFIG.ShowBtnRun * 10
Options_W    := g_CONFIG.ShowBtnOpt * 80
Options_X    := g_CONFIG.ShowBtnOpt * 10

Gui, Main:Color, % g_GUI.WinColor, % g_GUI.CtrlColor
Gui, Main:Font, % "c" g_GUI.FontColor " s" g_GUI.FontSize, % g_GUI.FontName
Gui, Main:%AlwaysOnTop%
Gui, Main:Add, Edit, x12 W%Input_W% -WantReturn vMyInput gOnSearchInput, Type anything here to search...
Gui, Main:Add, Button, % "x+"Enter_X " yp W" Enter_W " hp Default gRunCurrentCommand Hidden" !g_CONFIG.ShowBtnRun, Enter
Gui, Main:Add, Button, % "x+"Options_X " yp W" Options_W " hp gOptions Hidden" !g_CONFIG.ShowBtnOpt, Options
Gui, Main:Add, ListView, x12 ys+35 W%LV_W% H%LV_H% vMyListView AltSubmit gLVActions %ShowHdr% +LV0x10000 %ShowGrid% -Multi, No.|Type|Command|Description ; LV0x10000 Paints via double-buffering, which reduces flicker
Gui, Main:Add, Picture, x0 y0 0x4000000, % g_RUNTIME.BGPic
Gui, Main:Add, StatusBar, % "gSBActions Hidden" !g_CONFIG.ShowStatusBar,
Gui, Main:Default                                                       ; Set default GUI before any ListView / statusbar update

SB_SetParts(g_GUI.WinWidth - 90 * g_CONFIG.ShowRunCount)
SB_SetIcon("shell32.dll",-16752, 2)
Loop, 4
{
    LV_ModifyCol(A_Index, StrSplit(g_GUI.ColWidth, ",")[A_Index])
}

ListResult("Tip | F1 | Help`nTip | F2 | Options and settings`n"         ; List initial tips
    . "Tip | F3 | Edit current command`nTip | F4 | User-defined commands`n"
    . "Tip | ALT+SPACE / ALT+R | Activative ALTRun`n"
    . "Tip | ALT+SPACE / ESC / LOSE FOCUS | Deactivate ALTRun`n"
    . "Tip | ENTER / ALT+NO. | Run selected command`n"
    . "Tip | ARROW UP or DOWN | Select previous or next command`n"
    . "Tip | CTRL+D | Open selected cmd's dir with File Manager")

if (g_CONFIG.ShowIcon) {
    Global ImageListID := IL_Create(10, 5, 0)                           ; Create an ImageList so that the ListView can display some icons
    , IconMap := {"dir":IL_Add(ImageListID,"shell32.dll",-4)            ; Icon cache index, IconIndex=1/2/3/4/5 for type dir/func/url/eval
                ,"func":IL_Add(ImageListID,"shell32.dll",-25)
                ,"url":IL_Add(ImageListID,"shell32.dll",-512)
                ,"eval":IL_Add(ImageListID,"Calc.exe",-1)}
    LV_SetImageList(ImageListID)                                       ; Attach the ImageLists to the ListView so that it can later display the icons
}

Log.Debug("Resolving command line args=" A_Args[1] " " A_Args[2])       ; Command line args, Args are %1% %2% or A_Args[1] A_Args[2]
if (A_Args[1] = "-Startup")
    g_RUNTIME.WinHide := " Hide"

if (A_Args[1] = "-SendTo") {
    g_RUNTIME.WinHide := " Hide"
    CmdMgr(A_Args[2])
}

Gui, Main:Show, % "w" g_GUI.WinWidth " h" g_GUI.WinHeight " Center " g_RUNTIME.WinHide, % g_RUNTIME.WinName

if (g_CONFIG.HideOnLostFocus)
    OnMessage(0x06, "WM_ACTIVATE")

;=============================================================
; Set Hotkey
;=============================================================
Hotkey, % g_HOTKEY.GlobalHotkey1, ToggleWindow, UseErrorLevel
Hotkey, % g_HOTKEY.GlobalHotkey2, ToggleWindow, UseErrorLevel

Hotkey, IfWinActive, % g_RUNTIME.WinName                                ; Hotkey take effect only when ALTRun actived
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
Loop, % g_GUI.ListRows                      ; ListRows limit <= 9
{
    Hotkey, !%A_Index%, RunSelectedCommand  ; ALT + No. run command
    Hotkey, ^%A_Index%, GotoCommand         ; Ctrl + No. locate command
}

Loop, 3
    Hotkey, % g_HOTKEY["Hotkey"A_Index], % g_HOTKEY["Trigger"A_Index], UseErrorLevel ; Set Trigger <-> Hotkey, UseErrorLevel to Skips any warning dialogs

Hotkey, IfWinActive

Listary()
AppControl()                                                            ; Set Listary Dir QuickSwitch, Set AppControl
Return

Activate()
{
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
    Result := ""
    Order  := 1
    g_CurrentCommandList := Object()
    Prefix := SubStr(command, 1, 1)

    if (Prefix = "+" or Prefix = " " or Prefix = ">") {
        g_RUNTIME.CurrentCMD := g_Fallback[InStr("+ >", Prefix)]        ; Corresponding to fallback commands position no. 1, 2 & 3
        g_CurrentCommandList.Push(g_RUNTIME.CurrentCMD)
        Return ListResult(g_RUNTIME.CurrentCMD)
    }

    for index, element in g_Commands
    {
        splitResult := StrSplit(element, " | ")
        _Type := splitResult[1]
        _Path := splitResult[2]
        _Desc := splitResult[3]
        SplitPath, _Path, fileName                                      ; Extra name from _Path (if _Type is Dir and has "." in path, nameNoExt will not get full folder name)

        elementToShow   := (_Type = "dir" and g_CONFIG.ShowDirName) ? _Type " | " fileName " | " _Desc : _Type " | " _Path " | " _Desc ; Use _Path to show file icons (file type), and all other types, Show folder name only for dir type
        elementToSearch := g_CONFIG.MatchPath ? _Type " " _Path " " _Desc : _Type " " fileName " " _Desc ; search file name include extension & desc, search dir type + folder name + desc

        if FuzzyMatch(elementToSearch, command) {
            g_CurrentCommandList.Push(element)

            if (Order = 1) {
                g_RUNTIME.CurrentCMD := element
                Result .= elementToShow
            } else {
                Result .= "`n" elementToShow
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

            Result := "Eval | = " EvalResultTho
            Result .= "`n | ------------------------------------------------------"
            Result .= "`n | Beam width = " EvalResultTho " mm"
            Result .= "`n | Main bar no. = " RebarQty " (" Round((EvalResult-40*2) / (RebarQty - 1)) " c/c), " RebarQty + 1 " (" Round((EvalResult-40*2) / (RebarQty+1-1)) " c/c), " RebarQty - 1 " (" Round((EvalResult-40*2) / (RebarQty-1-1)) " c/c)"
            Result .= "`n | ------------------------------------------------------"
            Result .= "`n | As = " EvalResultTho " mm2"
            Result .= "`n | Rebar = " Ceil(EvalResult/132.7) "H13 / " Ceil(EvalResult/201.1) "H16 / " Ceil(EvalResult/314.2) "H20 / " Ceil(EvalResult/490.9) "H25 / " Ceil(EvalResult/804.2) "H32"
            Return ListResult(Result, True)
        }

        g_RUNTIME.UseFallback:= true
        g_CurrentCommandList := g_Fallback
        g_RUNTIME.CurrentCMD := g_Fallback[1]
    
        for i, cmd in g_Fallback
            Result .= (i = 1) ? g_Fallback[i] : "`n" g_Fallback[i]
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
        _Path       := AbsPath(splitResult[2])                          ; Must store in var for afterward use, trim space (in AbsPath)
        _Desc       := splitResult[3]
        IconIndex   := g_CONFIG.ShowIcon ? GetIconIndex(_Path, _Type) : 0

        LV_Add("Icon" . IconIndex, A_Index, _Type, _Path, _Desc)
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
    Switch (_Type)
    {
        Case "Dir" : Return 1
        Case "Func","Cmd","Tip": Return 2
        Case "URL" : Return 3
        Case "Eval": Return 4
        Case "File": {
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
        IconIndex = 9999999                         ; Set it out of bounds to display a blank icon.
    else {                                          ; Icon successfully loaded. Extract the hIcon member from the structure
        hIcon := NumGet(sfi, 0)                     ; Add the HICON directly to the small-icon lists.
        IconIndex := DllCall("ImageList_ReplaceIcon", "ptr", ImageListID, "int", -1, "ptr", hIcon) + 1 ; Uses +1 to convert the returned index from zero-based to one-based:
        DllCall("DestroyIcon", "ptr", hIcon)        ; Now that it's been copied into the ImageLists, the original should be destroyed
        IconMap[ExtOrPath] := IconIndex             ; Cache the icon based on file type (xlsx, pdf) or path (exe, lnk) to save memory and improve loading performance
    }
    Return IconIndex
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
    Path := StrReplace(Path, "%OneDrive%", g_RUNTIME.OneDrive)                     ; Convert OneDrive to absolute path due to #NoEnv
    Path := StrReplace(Path, "%OneDriveConsumer%", g_RUNTIME.OneDriveConsumer)     ; Convert OneDrive to absolute path due to #NoEnv
    Path := StrReplace(Path, "%OneDriveCommercial%", g_RUNTIME.OneDriveCommercial) ; Convert OneDrive to absolute path due to #NoEnv
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

EnvGet(EnvVar)
{
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

    switch (_Type)
    {
        case "FILE","URL","CMD":
            Run % _Path,, UseErrorLevel
            if ErrorLevel
                MsgBox Could not open "%_Path%"
        case "DIR":
            OpenDir(_Path)
        case "FUNC":
            if IsFunc(_Path)
                %_Path%()
        Default:                                                        ; For all other un-defined type
            Run % _Path,, UseErrorLevel
    }

    if (g_CONFIG.SaveHistory) {
        g_History.InsertAt(1, originCmd " /arg=" Arg)                   ; Adjust command history

        (g_History.Length() > g_CONFIG.HistoryLen) ? g_History.Pop()

        IniDelete, % g_RUNTIME.Ini, % g_SEC.History
        for index, element in g_History
            IniWrite, %element%, % g_RUNTIME.Ini, % g_SEC.History, %index% ; Save command history
    }

    UpdateRunCount()
    (g_CONFIG.SmartRank) ? UpdateRank(originCmd)
    Log.Debug("Execute(" g_CONFIG.RunCount ")=" originCmd)
}

TabFunc()
{
    GuiControlGet, CurrCtrl, Main:FocusV                                ; Limit tab to switch between Edit1 & ListView only
    GuiControl, Main:Focus, % (CurrCtrl = "MyInput") ? "MyListView" : "MyInput"
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

    if (index <= g_CurrentCommandList.Count()) {
        ChangeCommand(index, True)
        g_RUNTIME.CurrentCMD := g_CurrentCommandList[index]
    }
}

ChangeCommand(Step = 1, ResetSelRow = False)
{
    Gui, Main:Default                                                   ; Use it before any LV update

    SelRow := ResetSelRow ? Step : LV_GetNext() + Step                  ; Get target row no. to be selected
    SelRow := SelRow > LV_GetCount() ? 1 : SelRow                       ; Listview cycle selection (Mod has bug on upward cycle)
    SelRow := SelRow < 1 ? LV_GetCount() : SelRow
    g_RUNTIME.CurrentCMD := g_CurrentCommandList[SelRow]                ; Get current command from selected row

    LV_Modify(SelRow, "Select Focus Vis")                               ; make new index row selected, Focused, and Visible
    SetStatusBar()
}

LVActions()                                                             ; ListView g label actions (left / double click) behavior
{
    Gui, Main:Default                                                   ; Use it before any LV update
    focusedRow := LV_GetNext(0, "Focused")                              ; 查找焦点行, 仅对焦点行进行操作而不是所有选择的行:
    if (!focusedRow)                                                    ; 没有焦点行
        Return

    g_RUNTIME.CurrentCMD := g_CurrentCommandList[focusedRow]            ; Get current command from focused row

    if (A_GuiEvent = "RightClick")
    {
        Menu, LV_ContextMenu, Show
    }
    else if (A_GuiEvent = "DoubleClick" and g_RUNTIME.CurrentCMD)       ; Double click behavior, if g_RUNTIME.CurrentCMD = "" eg. first tip page, run it will clear g_SEC.UserCMD, g_SEC.Index, g_SEC.DftCMD
    {
        RunCommand(g_RUNTIME.CurrentCMD)
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

    g_RUNTIME.CurrentCMD := g_CurrentCommandList[focusedRow]            ; Get current command from focused row
    If (A_ThisMenuItem = "Run`tEnter")                                  ; User selected "Run`tEnter"
        RunCommand(g_RUNTIME.CurrentCMD)
    else if (A_ThisMenuItem = "Copy")
        LV_GetText(Text, focusedRow, 3) ? (A_Clipboard := Text)         ; Get the text from the focusedRow's 3rd field.
}

SBActions() {
    if (A_GuiEvent = "RightClick" and A_EventInfo = 1) {
        Menu, SB_ContextMenu, Add, Copy, SBContextMenu
        Menu, SB_ContextMenu, Icon, Copy, Shell32.dll, -243
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

AHKManual() {
    Run, https://www.autohotkey.com/docs/v1/
}

MainGuiEscape() {
    (g_CONFIG.EscClearInput and g_RUNTIME.Input) ? ClearInput() : MainGuiClose()
}

MainGuiClose() {                                                        ; If GuiClose is a function, the GUI is hidden by default
    g_CONFIG.KeepInput ? "" : ClearInput()
    Gui, Main:Hide
    SetStatusBar("Hint")                                                ; Update StatusBar hint information after GUI hide
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
    Log.Debug("mock test search ' " chr(chr1) " " chr(chr2) " " chr(chr3) " ' 50 times, use time = " t)
    MsgBox % "Search '" chr(chr1) " " chr(chr2) " " chr(chr3) "' use Time =  " t
}

UserCommandList() {
    if (g_CONFIG.Editor = "notepad.exe")
        Run, % g_CONFIG.Editor " " g_RUNTIME.Ini,, UseErrorLevel
    Else
        Run, % g_CONFIG.Editor " /m [" g_SEC.UserCMD "] """ g_RUNTIME.Ini """",, UseErrorLevel ; /m Match text
}

EditCurrentCommand() {
    if (g_CONFIG.Editor = "notepad.exe")
        Run, % g_CONFIG.Editor " " g_RUNTIME.Ini,, UseErrorLevel
    else
        Run, % g_CONFIG.Editor " /m " """" g_RUNTIME.CurrentCMD "=""" " """ g_RUNTIME.Ini """",, UseErrorLevel ; /m Match text, locate to current command, add = at end to filter out [history] commands
}

ClearInput() {
    GuiControl, Main:Text, MyInput,
    GuiControl, Main:Focus, MyInput
}

SetStatusBar(Mode := "Command") {                                       ; Set StatusBar text, Mode 1: Current command (default), 2: Hint, 3: Any text
    Gui, Main:Default                                                   ; Set default GUI window before any ListView / StatusBar operate
    Switch (Mode) {
        Case "Command": SBText := StrSplit(g_RUNTIME.CurrentCMD, " | ")[2]
        Case "Hint": {
            Random, index, 1, g_Hints.Count()
            SBText := g_Hints[index]
        }
        Default: SBText := Mode
    }
    SB_SetText(SBText, 1), SB_SetText(g_CONFIG.RunCount, 2)
}

RunCurrentCommand() {
    RunCommand(g_RUNTIME.CurrentCMD)
}

ParseArg() {
    Global
    commandPrefix := SubStr(g_RUNTIME.Input, 1, 1)

    if (commandPrefix = "+" || commandPrefix = " " || commandPrefix = ">") 
    {
        Return Arg := SubStr(g_RUNTIME.Input, 2)                        ; 直接取命令为参数
    }

    if (InStr(g_RUNTIME.Input, " ") && !g_RUNTIME.UseFallback)          ; 用空格来判断参数
    {
        Arg := SubStr(g_RUNTIME.Input, InStr(g_RUNTIME.Input, " ") + 1)
    }
    else if (g_RUNTIME.UseFallback)
    {
        Arg := g_RUNTIME.Input
    }
    else
    {
        Arg := ""
    }
}

FuzzyMatch(Haystack, Needle) {
    Needle := StrReplace(Needle, "+", "\+")                             ; For Eval (preceded by a backslash to be seen as literal)
    Needle := StrReplace(Needle, "*", "\*")                             ; For Eval (eg. 25+5 or 6*5 will show Eval result instead of match file with "30")
    Needle := StrReplace(Needle, " ", ".*")                             ; 空格直接替换为匹配任意字符
    Return RegExMatch(Haystack, "imS)" Needle)
}

UpdateRank(originCmd, showRank := false, inc := 1) {
    RANKSEC := g_SEC.DftCMD "|" g_SEC.UserCMD "|" g_SEC.Index
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
        g_RUNTIME.UsageToday := 0
    }
    g_RUNTIME.UsageToday++
    IniWrite, % g_RUNTIME.UsageToday, % g_RUNTIME.Ini, % g_SEC.Usage, % A_YYYY . A_MM . A_DD
}

UpdateRunCount() {
    g_CONFIG.RunCount++
    IniWrite, % g_CONFIG.RunCount, % g_RUNTIME.Ini, % g_SEC.Config, RunCount ; Record run counting
    SetStatusBar()
}

RunSelectedCommand() {
    index := SubStr(A_ThisHotkey, 0, 1)
    RunCommand(g_CurrentCommandList[index])
}

RankUp() {
    UpdateRank(g_RUNTIME.CurrentCMD, true)
}

RankDown() {
    UpdateRank(g_RUNTIME.CurrentCMD, true, -1)
}

LoadCommands() {
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
    
    IniRead, FALLBACKCMDSEC, % g_RUNTIME.Ini, % g_SEC.Fallback          ; Read FALLBACK section, initialize it if section not exist
    if (FALLBACKCMDSEC = "") {
        IniWrite, 
        (Ltrim
        ;===========================================================
        ; Fallback Commands show when search result is empty
        ;
        Func | CmdMgr | New Command
        Func | Everything | Search by Everything
        Func | Google | Search Clipboard or Input by Google
        Func | AhkRun | Run Command use AutoHotkey Run
        Func | Bing | Search Clipboard or Input by Bing
        ), % g_RUNTIME.Ini, % g_SEC.Fallback
        IniRead, FALLBACKCMDSEC, % g_RUNTIME.Ini, % g_SEC.Fallback
    }
    Loop Parse, FALLBACKCMDSEC, `n                                      ; Get and verify each FBCommand (Rank not necessary) and push it to g_Fallback
    {
        IsFunc(StrSplit(A_LoopField, " | ")[2]) ? g_Fallback.Push(A_LoopField)
    }
    Return Log.Debug("Loading commands list...OK")
}

LoadHistory() {
    if (g_CONFIG.SaveHistory) {
        Loop % g_CONFIG.HistoryLen
        {
            IniRead, History, % g_RUNTIME.Ini, % g_SEC.History, % A_Index, % A_Space
            g_History.Push(History)
        }
    } else
        IniDelete, % g_RUNTIME.Ini, % g_SEC.History
}

GetCmdOutput(command) {
    TempFile   := A_Temp "\ALTRun.stdout"
    FullCommand = %ComSpec% /C "%command% > %TempFile%"

    RunWait, %FullCommand%, %A_Temp%, Hide
    FileRead, Result, %TempFile%
    FileDelete, %TempFile%
    Return RTrim(Result, "`r`n")                                        ; Remove result rightmost/last "`r`n"
}

GetRunResult(command) {                                                 ;运行CMD并取返回结果方式2
    shell := ComObjCreate("WScript.Shell")                              ; WshShell object: http://msdn.microsoft.com/en-us/library/aew9yb99
    exec := shell.Exec(ComSpec " /C " command)                          ; Execute a single command via cmd.exe
    Return exec.StdOut.ReadAll()                                        ; Read and Return the command's output
}

OpenDir(Path, OpenContainer := False) {
    Path := AbsPath(Path)

    if (OpenContainer) {
        if (g_CONFIG.FileMgr = "Explorer.exe")
            Run, % g_CONFIG.FileMgr " /Select`, """ Path """",, UseErrorLevel
        else
            Run, % g_CONFIG.FileMgr " /P """ Path """",, UseErrorLevel    ; /P Parent folder
    } else
        Run, % g_CONFIG.FileMgr " """ Path """",, UseErrorLevel

    if ErrorLevel
        MsgBox, 4096, % g_RUNTIME.WinName, Error found, error code : %A_LastError%

    Log.Debug("Open Dir=" Path)
}

OpenCurrentFileDir() {
    OpenDir(StrSplit(g_RUNTIME.CurrentCMD, " | ")[2], True)
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
        , Send command to ALTRun User Command list, Shell32.dll, , -25
    else
        FileCreateShortcut, "%A_AhkPath%", %lnkPath%, , "%A_ScriptFullPath%" -SendTo
        , Send command to ALTRun User Command list, Shell32.dll, , -25
    Return "OK"
}

UpdateStartup(create := true) {
    lnkPath := A_Startup "\ALTRun.lnk"

    if (!create) {
        FileDelete, %lnkPath%
        Return "Disabled"
    }

    FileCreateShortcut, %A_ScriptFullPath%, %lnkPath%, %A_ScriptDir%
        , -startup, ALTRun - An effective launcher, Shell32.dll, , -25
    Return "OK"
}

UpdateStartMenu(create := true) {
    lnkPath := A_Programs "\ALTRun.lnk"

    if (!create) {
        FileDelete, %lnkPath%
        Return "Disabled"
    }

    FileCreateShortcut, %A_ScriptFullPath%, %lnkPath%, %A_ScriptDir%
        , -StartMenu, ALTRun, Shell32.dll, , -25
    Return "OK"
}

Reindex() {                                                             ; Re-create Index section
    IniDelete, % g_RUNTIME.Ini, % g_SEC.Index
    for dirIndex, dir in StrSplit(g_CONFIG.IndexDir, ",")
    {
        searchPath := AbsPath(dir)

        for extIndex, ext in StrSplit(g_CONFIG.IndexType, ",")
        {
            Loop Files, %searchPath%\%ext%, R
            {
                if (g_CONFIG.IndexExclude != "" && RegExMatch(A_LoopFileLongPath, g_CONFIG.IndexExclude))
                    continue                                            ; Skip this file and move on to the next loop.

                IniWrite, 1, % g_RUNTIME.Ini, % g_SEC.Index, File | %A_LoopFileLongPath% ; Assign initial rank to 1
                Progress, %A_Index%, %A_LoopFileName%, ReIndexing..., Reindex
            }
        }
        Progress, Off
    }

    Log.Debug("Indexing search database...")
    TrayTip, % g_RUNTIME.WinName, ReIndex database finish successfully. , 8
    LoadCommands()
}

Help() {
    Options(, 7)                                                        ; Open Options window 7th tab (help tab)
}

Usage() {
    Options(, 6)
}

Listary() {                                                             ; Listary Dir QuickSwitch Function (快速更换保存/打开对话框路径)
    Log.Debug("Listary function starting...")

    Loop Parse, % g_CONFIG.FileMgrID, `,                                ; File Manager Class, default is Windows Explorer & Total Commander
        GroupAdd, FileMgrID, %A_LoopField%

    Loop Parse, % g_CONFIG.DialogWin, `,                                ; 需要QuickSwith的窗口, 包括打开/保存对话框等
        GroupAdd, DialogBox, %A_LoopField%

    Loop Parse, % g_CONFIG.ExcludeWin, `,                               ; 排除特定窗口,避免被 Auto-QuickSwitch 影响
        GroupAdd, ExcludeWin, %A_LoopField%

    if (g_CONFIG.AutoSwitchDir) {
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
    Log.Debug("Listary change path=" Dir)
}

CmdMgr(Path := "") {                                                    ; 命令管理窗口
    Global
    Log.Debug("Starting Command Manager... Args=" Path)

    SplitPath Path, _Desc,, fileExt,,                                   ; Extra name from _Path (if _Type is dir and has "." in path, nameNoExt will not get full folder name) 
    
    if InStr(FileExist(Path), "D")                                      ; True only if the file exists and is a directory.
        _Type := 2                                                      ; It is a normal folder
    else                                                                ; From command "New Command" or GUI context menu "New Command"
        _Desc := Arg
    
    if (fileExt = "lnk" && g_CONFIG.SendToGetLnk) {
        FileGetShortcut, %Path%, Path,, fileArg, _Desc
        Path .= " " fileArg
    }

    Gui, CmdMgr:New
    Gui, CmdMgr:Font, S10, Segoe UI
    Gui, CmdMgr:Margin, 5, 5
    Gui, CmdMgr:Add, GroupBox, w550 h230, New Command
    Gui, CmdMgr:Add, Text, xp+20 yp+35, Command Type: 
    Gui, CmdMgr:Add, DropDownList, xp+120 yp-5 w150 v_Type Choose%_Type%, File||Dir|Cmd|URL
    Gui, CmdMgr:Add, Text, xp-120 yp+50, Command Path: 
    Gui, CmdMgr:Add, Edit, xp+120 yp-5 w350 v_Path, % RelativePath(Path)
    Gui, CmdMgr:Add, Button, xp+355 yp w30 hp gSelectCmdPath, ...
    Gui, CmdMgr:Add, Text, xp-475 yp+100, Name/Description: 
    Gui, CmdMgr:Add, Edit, xp+120 yp-5 w350 v_Desc, %_Desc%
    Gui, CmdMgr:Add, Button, Default x415 w65, OK
    Gui, CmdMgr:Add, Button, xp+75 yp w65, Cancel
    Gui, CmdMgr:Show, AutoSize, Commander Manager
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
        IniWrite, 1, % g_RUNTIME.Ini, % g_SEC.UserCMD, %_Type% | %_Path% %_Desc% ; initial rank = 1
        if (!ErrorLevel)
            MsgBox,64, Command Manager, %_Path% `n`nCommand added successfully!
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
    Gui, Setting:New, -SysMenu +AlwaysOnTop, Options                    ;-SysMenu: omit the system menu icons, +OwnerMain: (omit due to lug options window)
    Gui, Setting:Font, s9, Segoe UI
    Gui, Setting:Add, Tab3,vCurrTab Choose%ActTab%, General|Index|GUI|Hotkey|Listary|Usage|About
    Gui, Setting:Tab, 1 ; CONFIG Tab
    Gui, Setting:Add, ListView, w500 h300 Checked -Multi AltSubmit -Hdr vOptListView, Options

    For key, description in g_CHKLV
        LV_Add("Check" g_CONFIG[key], description)

    LV_ModifyCol(1, "AutoHdr")                                          ; AutoHdr: Automatically add a header to the column

    Gui, Setting:Add, Text, yp+320, Text Editor:
    Gui, Setting:Add, ComboBox, xp+100 yp-5 w400 Sort vg_Editor, % g_CONFIG.Editor "||Notepad.exe|C:\Apps\Notepad4\Notepad4.exe"
    Gui, Setting:Add, Text, xp-100 yp+40, Everything.exe:
    Gui, Setting:Add, ComboBox, xp+100 yp-5 w400 Sort vg_Everything, % g_CONFIG.Everything "||C:\Apps\Everything\Everything.exe"
    Gui, Setting:Add, Text, xp-100 yp+40, File Manager:
    Gui, Setting:Add, ComboBox, xp+100 yp-5 w400 Sort vg_FileMgr, % g_CONFIG.FileMgr "||Explorer.exe|C:\Apps\TotalCMD64\TotalCMD.exe /O /T /S"
    
    Gui, Setting:Tab, 2 ; INDEX Tab
    Gui, Setting:Add, GroupBox, w500 h130, Index
    Gui, Setting:Add, Text, xp+10 yp+25, Index Locations: 
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 vg_IndexDir, % g_CONFIG.IndexDir "||A_ProgramsCommon,A_StartMenu"
    Gui, Setting:Add, Text, xp-150 yp+40, Index File Type: 
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 vg_IndexType, % g_CONFIG.IndexType "||*.lnk,*.exe"
    Gui, Setting:Add, Text, xp-150 yp+40, Index File Exclude: 
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 vg_IndexExclude, % g_CONFIG.IndexExclude "||Uninstall *"
    Gui, Setting:Add, GroupBox, xp-160 yp+45 w500 h270, Others
    Gui, Setting:Add, Text, xp+10 yp+25, Command history length: 
    Gui, Setting:Add, DropDownList, xp+150 yp-5 w330 Sort vg_HistoryLen, % StrReplace("0|10|15|20|25|30|50|90|", g_CONFIG.HistoryLen, g_CONFIG.HistoryLen . "|")
    Gui, Setting:Add, Text, xp-150 yp+40, Config (.ini) file location:
    Gui, Setting:Add, Edit, xp+150 yp-5 w330 Disabled vg_Ini, % g_RUNTIME.Ini

    Gui, Setting:Tab, 3 ; GUI Tab
    Gui, Setting:Add, GroupBox, w500 h420, GUI
    Gui, Setting:Add, Text, xp+10 yp+25 , Command result limit
    Gui, Setting:Add, DropDownList, xp+150 yp-5 w330 vg_ListRows, % StrReplace("3|4|5|6|7|8|9|", g_GUI.ListRows, g_GUI.ListRows . "|") ; Not Choose%g_ListRows% as list start from 3
    Gui, Setting:Add, Text, xp-150 yp+40, Width of each column:
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 Sort vg_ColWidth, % g_GUI.ColWidth "||40,45,430,340"
    Gui, Setting:Add, Text, xp-150 yp+40, Font Name:
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 Sort vg_FontName, % g_GUI.FontName "||Default|Segoe UI Semibold|Microsoft Yahei"
    Gui, Setting:Add, Text, xp-150 yp+40, Font Size: 
    Gui, Setting:Add, ComboBox, xp+150 yp-5 r1 w330 vg_FontSize, % g_GUI.FontSize "||8|9|10|11|12"
    Gui, Setting:Add, Text, xp-150 yp+40, Font Color: 
    Gui, Setting:Add, ComboBox, xp+150 yp-5 r1 w330 vg_FontColor, % g_GUI.FontColor "||Default|Black|Blue|DCDCDC|000000"
    Gui, Setting:Add, Text, xp-150 yp+40, Window Width: 
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 vg_WinWidth, % g_GUI.WinWidth "||920"
    Gui, Setting:Add, Text, xp-150 yp+40, Window Height:
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 vg_WinHeight, % g_GUI.WinHeight "||313"
    Gui, Setting:Add, Text, xp-150 yp+40, Controls' Color:
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 vg_CtrlColor, % g_GUI.CtrlColor "||Default|White|Blue|202020|FFFFFF"
    Gui, Setting:Add, Text, xp-150 yp+40, Background Color:
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 vg_WinColor, % g_GUI.WinColor "||Default|White|Blue|202020|FFFFFF"
    Gui, Setting:Add, Text, xp-150 yp+40, Background Picture:
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 vg_Background, % g_GUI.Background "||NO PICTURE|DEFAULT|C:\Path\BackgroundPicture.jpg"

    g_RUNTIME.FuncList := Object()
    for index, element in g_Commands                                    ; Load all Func name, for Options window
    {
        splitResult := StrSplit(element, " | ")
        g_RUNTIME.FuncList .= (splitResult[1] = "Func" and IsFunc(splitResult[2])) ? splitResult[2] "|" : ""
    }

    Gui, Setting:Tab, 4 ; Hotkey Tab
    Gui, Setting:Add, GroupBox, w500 h120, Activate
    Gui, Setting:Add, Text, xp+10 yp+25 , Primary Hotkey :
    Gui, Setting:Add, Hotkey, xp+250 yp-4 w230 vg_GlobalHotkey1, % g_HOTKEY.GlobalHotkey1
    Gui, Setting:Add, Text, xp-250 yp+35 , Secondary Hotkey :
    Gui, Setting:Add, Hotkey, xp+250 yp-4 w230 vg_GlobalHotkey2,% g_HOTKEY.GlobalHotkey2
    Gui, Setting:Add, Link, xp-250 yp+35 gResetHotkey, You can set another hotkey as a secondary hotkey (<a id="Reset">Reset to Default</a>)
    Gui, Setting:Add, GroupBox, xp-10 yp+40 w500 h55, Command Hotkey:
    Gui, Setting:Add, Text, xp+10 yp+25 , Execute Command:
    Gui, Setting:Add, Text, xp+150 yp , ALT + No.
    Gui, Setting:Add, Text, xp+105 yp, Select Command: 
    Gui, Setting:Add, Text, xp+130 yp, Ctrl + No.
    Gui, Setting:Add, GroupBox, xp-395 yp+40 w500 h130, Action Hotkey:
    Gui, Setting:Add, Text, xp+10 yp+25 , Hotkey 1: 
    Gui, Setting:Add, Hotkey, xp+150 yp-5 w80 vg_Hotkey1, % g_HOTKEY.Hotkey1
    Gui, Setting:Add, Text, xp+100 yp+5, Toggle Action: 
    Gui, Setting:Add, DropDownList, xp+110 yp-5 w120 Sort vg_Trigger1, % GetFuncList(g_HOTKEY.Trigger1)
    Gui, Setting:Add, Text, xp-360 yp+40 , Hotkey 2: 
    Gui, Setting:Add, Hotkey, xp+150 yp-5 w80 vg_Hotkey2, % g_HOTKEY.Hotkey2
    Gui, Setting:Add, Text, xp+100 yp+5, Toggle Action: 
    Gui, Setting:Add, DropDownList, xp+110 yp-5 w120 Sort vg_Trigger2, % GetFuncList(g_HOTKEY.Trigger2)
    Gui, Setting:Add, Text, xp-360 yp+40 , Hotkey 3: 
    Gui, Setting:Add, Hotkey, xp+150 yp-5 w80 vg_Hotkey3, % g_HOTKEY.Hotkey3
    Gui, Setting:Add, Text, xp+100 yp+5, Toggle Action: 
    Gui, Setting:Add, DropDownList, xp+110 yp-5 w120 Sort vg_Trigger3, % GetFuncList(g_HOTKEY.Trigger3)

    Gui, Setting:Tab, 5 ; LISTARTY TAB
    Gui, Setting:Add, GroupBox, w500 h125, Listary Quick-Switch
    Gui, Setting:Add, Text, xp+10 yp+25 , File Manager Title:
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 Sort vg_FileMgrID, % g_CONFIG.FileMgrID "||ahk_class CabinetWClass|ahk_class CabinetWClass, ahk_class TTOTAL_CMD"
    Gui, Setting:Add, Text, xp-150 yp+40, Open/Save Dialog Title:
    Gui, Setting:Add, Combobox, xp+150 yp-5 w330 Sort vg_DialogWin, % g_CONFIG.DialogWin "||ahk_class #32770"
    Gui, Setting:Add, Text, xp-150 yp+40, Exclude Windows Title: 
    Gui, Setting:Add, ComboBox, xp+150 yp-5 w330 Sort vg_ExcludeWin, % g_CONFIG.ExcludeWin "||ahk_class SysListView32|ahk_class SysListView32, ahk_exe Explorer.exe|ahk_class SysListView32, ahk_exe Explorer.exe, ahk_exe Totalcmd64.exe, AutoCAD LT Alert"
    Gui, Setting:Add, GroupBox, xp-160 yp+45 w500 h125, Hotkey for Switch Open/Save dialog path to
    Gui, Setting:Add, Text, xp+10 yp+30, Total Commander's dir:
    Gui, Setting:Add, Hotkey, xp+150 yp-5 w330 vg_TotalCMDDir, % g_HOTKEY.TotalCMDDir
    Gui, Setting:Add, Text, xp-150 yp+40, Windows Explorer's dir:
    Gui, Setting:Add, Hotkey, xp+150 yp-5 w330 vg_ExplorerDir, % g_HOTKEY.ExplorerDir
    Gui, Setting:Add, CheckBox, % "xp-150 yp+40 vg_AutoSwitchDir checked" g_CONFIG.AutoSwitchDir, Auto Switch Dir

    Gui, Setting:Tab, 6 ; USAGE TAB
    Gui, Setting:Add, GroupBox, w500 h420, Program usage status
    Gui, Setting:Add, GroupBox, xp+40 yp+20 w445 h300,

    OffsetDate := A_Now
    EnvAdd, OffsetDate, -30, Days                                       ; 减去 30 天

    IniRead, USAGE, % g_RUNTIME.Ini, % g_SEC.Usage                      ; read all key & value in whole section
    Loop, Parse, USAGE, `n                                              ; read each line to get each key and value
    {
        g_RUNTIME.MaxVal := Max(g_RUNTIME.MaxVal, StrSplit(A_LoopField, "=")[2])
    }
    
    Loop, 30
    {
        EnvAdd, OffsetDate, +1, Days
        FormatTime, OffsetDate, %OffsetDate%, yyyyMMdd

        IniRead, OutputVar, % g_RUNTIME.Ini, % g_SEC.Usage, % OffsetDate, 0
        Gui, Setting:Add, Progress, % "c94DD88 BackgroundF9F9F9 Vertical y90 w14 h280 xm+" 50+A_Index*14 " Range0-" g_RUNTIME.MaxVal+10, %OutputVar%
    }
    Gui, Setting:Add, Text, xp-450 yp-5, % g_RUNTIME.MaxVal+10
    Gui, Setting:Add, Text, xp yp+140, % Round(g_RUNTIME.MaxVal/2)+5
    Gui, Setting:Add, Text, xp yp+140, 0
    Gui, Setting:Add, Text, xp+35 yp+15, 30 days ago 
    Gui, Setting:Add, Text, xp+410 yp, Now
    Gui, Setting:Add, Text, x66 yp+35, Total number of times the command was executed:
    Gui, Setting:Add, Edit, xp+343 yp-5 w100 Disabled vg_RunCount, % g_CONFIG.RunCount
    Gui, Setting:Add, Text, x66 yp+35, Number of times the program was activated today:
    Gui, Setting:Add, Edit, xp+343 yp-5 w100 Disabled, % g_RUNTIME.UsageToday

    Gui, Setting:Tab, 7 ; ABOUT TAB
    Gui, Setting:Font, S12, Segoe UI Semibold
    Gui, Setting:Add, Text, W450 Center, ALTRun
    Gui, Setting:Font, S10, Segoe UI Semibold
    Gui, Setting:Add, Edit, yp+25 w500 h370 ReadOnly -WantReturn -Wrap,
    (Ltrim
    An effective launcher for Windows, an AutoHotkey open-source project. 
    It provides a streamlined and efficient way to find anything on your 
    system and launch any application in your way.

    1. Pure portable software, not write anything into Registry.
    2. Small file size (< 100KB), low resource usage (< 10MB RAM), and high performance.
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
    Gui, Setting:Add, Link, yp+375, Check more details from <a href="https://github.com/zhugecaomao/ALTRun">Github Homepage</a> and <a href="https://github.com/zhugecaomao/ALTRun/wiki#notes-for-altrun-options">App Setting Notes</a>
    Gui, Setting:Tab                                                    ; 后续添加的控件将不属于前面的选项卡控件
    Gui, Setting:Add, Button, Default x365 w80, OK
    Gui, Setting:Add, Button, xp+90 yp w80, Cancel
    Gui, Setting:Show,, Options
    Hotkey, % g_HOTKEY.GlobalHotkey1, Off, UseErrorLevel
    Hotkey, % g_HOTKEY.GlobalHotkey2, Off, UseErrorLevel
    Log.Debug("Loading options window...Arg=" Arg ", ActTab=" ActTab)
}

GetFuncList(Trigger)
{
    g_RUNTIME.FuncList := ""
    for index, element in g_Commands                                    ; Load all Func name, for Options window
    {
        splitResult := StrSplit(element, " | ")
        g_RUNTIME.FuncList .= (splitResult[1] = "Func" and IsFunc(splitResult[2])) ? splitResult[2] "|" : ""
    }
    Return StrReplace("---|" g_RUNTIME.FuncList, Trigger, Trigger . "|")
}

ResetHotkey() {
    GuiControl, Setting:, g_GlobalHotkey1, !Space
    GuiControl, Setting:, g_GlobalHotkey2, !R
}

SettingButtonOK() {
    SAVECONFIG()
    Reload
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

LOADCONFIG(Arg)                                                         ; 加载主配置文件
{
    Log.Debug("Loading configuration...Arg=" Arg)

    if (Arg = "config" or Arg = "initialize" or Arg = "all") {
        For key, value in g_CONFIG                                      ; Read [Config] to Object
        {
            IniRead, tempValue, % g_RUNTIME.Ini, % g_SEC.Config, %key%, %value%
            g_CONFIG[key] := tempValue
            ; OutputDebug, % key " = " g_CONFIG[key]
        }

        For key, value in g_HOTKEY                                      ; Read Hotkey section
        {
            IniRead, tempValue, % g_RUNTIME.Ini, % g_SEC.Hotkey, %key%, %value%
            g_HOTKEY[key] := tempValue
        }

        For key, value in g_GUI                                         ; Read GUI section
        {
            IniRead, tempValue, % g_RUNTIME.Ini, % g_SEC.Gui, %key%, %value%
            g_GUI[key] := tempValue
        }

        g_RUNTIME.BGPic := (g_GUI.Background = "DEFAULT") ? Extract_BG(A_Temp "\ALTRun.jpg") : g_GUI.Background

        IniRead, tempValue, % g_RUNTIME.Ini, % g_SEC.Usage, % A_YYYY . A_MM . A_DD, 0
        g_RUNTIME.UsageToday := tempValue

        OffsetDate := A_Now
        EnvAdd, OffsetDate, -30, Days                                   ; 减去 30 天
        FormatTime, OffsetDate, %OffsetDate%, yyyyMMdd
    
        IniRead, USAGE, % g_RUNTIME.Ini, % g_SEC.Usage                  ; Clean up usage record before 30 days
        Loop, Parse, USAGE, `n
        {
            key := StrSplit(A_LoopField, "=")[1]
            if (key < OffsetDate)
            {
                IniDelete, % g_RUNTIME.Ini, % g_SEC.Usage, %key%
            }
        }
    }

    if (Arg = "commands" or Arg = "initialize" or Arg = "all") {        ; Built-in command initialize
        DFTCMDSEC := "
        (LTrim
        ; Built-in commands, high priority, do not modify
        ; Will be automatically overwritten by the program.
        ;
        Func | Help | ALTRun Help Index (F1)=99
        Func | Options | ALTRun Options Preference Settings (F2)=99
        Func | Reload | ALTRun Reload=99
        Func | CmdMgr | New Command=99
        Func | UserCommandList | ALTRun User-defined command (F4)=99
        Func | Usage | ALTRun Usage Status=99
        Func | Reindex | Reindex search database=99
        Func | Everything | Search by Everything=99
        Func | RunPTTools | PT Tools (AHK)=99
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
        Dir | `%AppData`%\Microsoft\Windows\SendTo | Windows SendTo Dir=99
        Dir | `%OneDriveConsumer`% | OneDrive Personal Dir=99
        Cmd | explorer.exe | Windows File Explorer=99
        Cmd | cmd.exe | DOS / CMD=99
        Cmd | cmd.exe /k ipconfig | Check IP Address=99
        Cmd | Shell:AppsFolder | AppsFolder Applications=99
        Cmd | ::{645FF040-5081-101B-9F08-00AA002F954E} | Recycle Bin=99
        Cmd | ::{20D04FE0-3AEA-1069-A2D8-08002B30309D} | This PC=99
        Cmd | WF.msc | Windows Defender Firewall with Advanced Security=99
        Cmd | TaskSchd.msc | Task Scheduler=99
        Cmd | DevMgmt.msc | Device Manager=99
        Cmd | EventVwr.msc | Event Viewer=99
        Cmd | CompMgmt.msc | Computer Manager=99
        Cmd | TaskMgr.exe | Task Manager=99
        Cmd | Calc.exe | Calculator=99
        Cmd | MsPaint.exe | Paint=99
        Cmd | Regedit.exe | Registry Editor=99
        Cmd | Write.exe | Write=99
        Cmd | CleanMgr.exe | Disk Space Clean-up Manager=99
        Cmd | GpEdit.msc | Group Policy=99
        Cmd | DiskMgmt.msc | Disk Management=99
        Cmd | DxDiag.exe | Directx Diagnostic Tool=99
        Cmd | LusrMgr.msc | Local Users and Groups=99
        Cmd | MsConfig.exe | System Configuration=99
        Cmd | PerfMon.exe /Res | Resources Monitor=99
        Cmd | PerfMon.exe | Performance Monitor=99
        Cmd | WinVer.exe | About Windows=99
        Cmd | Services.msc | Services=99
        Cmd | NetPlWiz | User Accounts=99
        Cmd | Control | Control Panel=99
        Cmd | Control Intl.cpl | Region and Language Options=99
        Cmd | Control Firewall.cpl | Windows Defender Firewall=99
        Cmd | Control Access.cpl | Ease of Access Centre=99
        Cmd | Control AppWiz.cpl | Programs and Features=99
        Cmd | Control Sticpl.cpl | Scanners and Cameras=99
        Cmd | Control Sysdm.cpl | System Properties=99
        Cmd | Control Mouse | Mouse Properties=99
        Cmd | Control Desk.cpl | Display=99
        Cmd | Control Mmsys.cpl | Sound=99
        Cmd | Control Ncpa.cpl | Network Connections=99
        Cmd | Control Powercfg.cpl | Power Options=99
        Cmd | Control TimeDate.cpl | Date and Time=99
        Cmd | Control AdminTools | Windows Tools=99
        Cmd | Control Desktop | Personalisation=99
        Cmd | Control Folders | File Explorer Options=99
        Cmd | Control Fonts | Fonts=99
        Cmd | Control Inetcpl.cpl,,4 | Internet Properties=99
        Cmd | Control Printers | Devices and Printers=99
        Cmd | Control UserPasswords | User Accounts=99
        )"
        IniWrite, % DFTCMDSEC, % g_RUNTIME.Ini, % g_SEC.DftCMD

        IniRead, USERCMDSEC, % g_RUNTIME.Ini, % g_SEC.UserCMD
        if (USERCMDSEC = "") {
            IniWrite,
            (Ltrim
            ; User-Defined commands, high priority, edit as desired
            ; Format: Command Type | Command | Description=Rank
            ; Command type: File, Dir, CMD, URL
            ;
            Dir | `%AppData`%\Microsoft\Windows\SendTo | Windows SendTo Dir=99
            Dir | `%OneDriveConsumer`% | OneDrive Personal Dir=99
            Dir | `%OneDriveCommercial`% | OneDrive Business Dir=99
            Cmd | ipconfig | Show IP Address(CMD type will run with cmd.exe, auto pause after run)=99
            URL | www.google.com | Google=99
            File | C:\OneDrive\Apps\TotalCMD64\TOTALCMD64.exe=99
            ), % g_RUNTIME.Ini, % g_SEC.UserCMD
            IniRead, USERCMDSEC, % g_RUNTIME.Ini, % g_SEC.UserCMD
        }

        IniRead, INDEXSEC, % g_RUNTIME.Ini, % g_SEC.Index               ; Read whole section g_SEC.Index (Index database)
        if (INDEXSEC = "") {
            MsgBox, 4160, % g_RUNTIME.WinName, ALTRun is going to initialize for the first time running...`n`nConfig software and build the index database for search.`n`nAuto initialize in 30 seconds or click OK now., 30
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
        IniWrite, % g_%key%, % g_RUNTIME.Ini, % g_SEC.Config, %key%     ; For all ini file - [Config] g_ 变量从控件v变量和上一步Check Listview取得

    For key, value in g_GUI
        IniWrite, % g_%key%, % g_RUNTIME.Ini, % g_SEC.Gui, %key%

    For key, value in g_HOTKEY
        IniWrite, % g_%key%, % g_RUNTIME.Ini, % g_SEC.Hotkey, %key%

    Return Log.Debug("Saving config...")
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
    Return _Filename
}

;=============================================================
; Some Built-in Functions
;=============================================================
AhkRun() {
    Run, %Arg%,, UseErrorLevel
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
    word := Arg = "" ? clipboard : Arg
    Run, https://www.google.com/search?q=%word%&newwindow=1
}

Bing() {
    word := Arg = "" ? clipboard : Arg
    Run, http://cn.bing.com/search?q=%word%
}

Everything() {
    Run, % g_CONFIG.Everything " -s """ Arg """",, UseErrorLevel
    if ErrorLevel
        MsgBox, % "Everything software not found.`n`nPlease check ALTRun setting and Everything program file."
}

;=======================================================================
; Eval - Calculate a math expression
; 计算数学表达式，支持 +, -, *, /, ^ (或 **), 和括号 ()
; 如果输入包含非法字符，返回 0
;=======================================================================
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

Class Logger                                                            ; Logger library
{
    __New(filename) {
        this.filename := filename
    }

    Debug(Msg) {
        if (g_CONFIG.SaveLog)
            FileAppend, % "[" A_Now "] " Msg "`n", % this.filename
    }
}