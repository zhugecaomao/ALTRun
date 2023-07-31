;==============================================================
; ALTRun - An effective launcher for Windows.
; https://github.com/zhugecaomao/ALTRun
;==============================================================
#Requires AutoHotkey v1.1
#NoEnv                                                                  ; Recommended for performance and compatibility.
#Warn All, StdOut ; OutputDebug / MsgBox                                ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force
#NoTrayIcon
#Persistent

FileEncoding, UTF-8
SendMode Input                                                          ; Recommended due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%                                             ; Ensures a consistent starting directory.
SetKeyDelay 0

;=============================================================
; 声明全局变量,定义配置文件设置
;=============================================================
Global g_IniFile    := A_ScriptDir "\" A_ComputerName ".ini"
    ,  g_LogFile    := A_Temp "\ALTRun.log"
    ,  SEC_CONFIG   := "Config"
    ,  SEC_GUI      := "Gui"
    ,  SEC_DFTCMD   := "DefaultCommand"
    ,  SEC_USERCMD  := "UserCommand"
    ,  SEC_FALLBACK := "FallbackCommand"
    ,  SEC_HOTKEY   := "Hotkey"
    ,  SEC_HISTORY  := "History"
    ,  SEC_INDEX    := "Index"

, KEYLIST_CONFIG := "AutoStartup,EnableSendTo,InStartMenu,IndexDir,IndexFileType,IndexExclude,SearchFullPath,ShowFileExt,ShowIcon,KeepInputText,RunIfOnlyOne,HideOnDeactivate,AlwaysOnTop,SaveHistory,HistorySize,isLogging,AutoRank,SwitchToEngIME,EscClearInput,SendToGetLnk,Editor,TCPath,Everything,RunCount,EnableScheduler,ShutdownTime,AutoSwitchDir,FileMgrClass,DialogClass,ExcludeWinClass"
, KEYLIST_GUI    := "HideTitle,ShowTrayIcon,HideCol2,LVGrid,DisplayRows,Col3Width,Col4Width,FontName,FontSize,FontColor,WinWidth,EditHeight,ListHeight,CtrlColor,WinColor,BackgroundPicture"
, KEYLIST_HOTKEY := "GlobalHotkey1,GlobalHotkey1Win,GlobalHotkey2,GlobalHotkey2Win,Hotkey1,Trigger1,Hotkey2,Trigger2,Hotkey3,Trigger3,RunCmdAlt,EnableCapsLockIME"

, g_AutoStartup := 1                        ; 是否添加快捷方式到开机启动
, g_EnableSendTo := 1                       ; 是否创建“发送到”菜单
, g_InStartMenu := 1                        ; 是否添加快捷方式到开始菜单中
, g_IndexDir := "A_ProgramsCommon | A_StartMenu" ; 搜索的目录,可以使用 全路径 或以 A_ 开头的AHK变量, 以 " | " 分隔, 路径可包含空格, 无需加引号
, g_IndexFileType := "*.lnk | *.exe"        ; 搜索的文件类型, 以 " | " 分隔
, g_IndexExclude := "Uninstall | 卸载"      ; 排除的文件,正则表达式
, g_SearchFullPath := 0                     ; 搜索完整路径,否则只搜文件名
, g_ShowFileExt := 1                        ; 在界面显示文件扩展名
, g_ShowIcon := 1                           ; Show Icon in File ListView
, g_KeepInputText := 1                      ; 窗口隐藏时不清空编辑框内容
, g_TCPath := A_Space                       ; TotalCommander 路径,如果为空则使用资源管理器打开
, g_RunIfOnlyOne := 0                       ; 如果结果中只有一个则直接运行
, g_HideOnDeactivate := 1                   ; 窗口失去焦点后关闭窗口
, g_AlwaysOnTop := 1                        ; 窗口置顶显示
, g_SaveHistory := 1                        ; 记录历史
, g_HistorySize := 15                       ; 记录历史的数量
, g_RunCount := 0                           ; Record command execution times
, g_isLogging := 1                          ; Enable log record
, g_AutoRank := 1                           ; 自动根据使用频率调节顺序
, g_SwitchToEngIME := 0                     ; 每次激活窗口自动切换到英文输入法
, g_EscClearInput := 1                      ; 输入 Esc 时,如果输入框有内容则清空,无内容才关闭窗口
, g_Editor := A_Space                       ; 用来打开配置文件的编辑器,推荐Notepad2,默认为资源管理器关联的编辑器,可以右键->属性->打开方式修改
, g_SendToGetLnk := 1                       ; 如果使用发送到菜单发送的文件是 .lnk 的快捷方式,从文件读取路径后添加目标文件
, g_Everything := A_Space                   ; Everything.exe 文件路径
, g_EnableScheduler := 0                    ; Task Scheduler for PC auto shutdown
, g_ShutdownTime := "22:30"                 ; Set timing for PC auto shutdown
, g_AutoSwitchDir := 0
, g_FileMgrClass := "TTOTAL_CMD | CabinetWClass"
, g_DialogClass := "#32770"                 ; Class name of the Dialog Box which Listary Switch Dir will take effect
, g_ExcludeWinClass := "ahk_exe 7zG.exe | ahk_class SysListView32 | AutoCAD LT | RAPT | ahk_exe Totalcmd64.exe" ; Exclude those windows that not want Listary Switch Dir take effect
, g_HideTitle := 0                          ; 隐藏标题栏
, g_ShowTrayIcon := 1                       ; 是否显示托盘图标
, g_DisplayRows := 9                        ; 在列表中显示的行数,如果超过9行,定位到该行的快捷键将无效
, g_Col3Width := 415                        ; 在列表中第三列的宽度
, g_Col4Width := 360                        ; 在列表中第四列的宽度
, g_HideCol2 := 0                           ; 隐藏第二列,即显示 文件、功能 的一列 `n如果隐藏, 第二列的空间会分给第三列
, g_LVGrid := 0                             ; Show Grid in ListView
, g_FontName := "Segoe UI"                  ; Font Name (eg. Default, Segoe UI, Microsoft Yahei)
, g_FontSize := 10                          ; Font Size, Default is 10
, g_FontColor := "Default"                  ; Font Color, (eg. cRed, cFFFFAA, cDefault)
, g_WinWidth := 900
, g_EditHeight := 24
, g_ListHeight := 260                       ; Command List Height
, g_CtrlColor := "Default"
, g_WinColor := "Default"
, g_BackgroundPicture := "Default"
, g_BGPicture                               ; Real path of the BackgroundPicture
, g_Hints := ["It's better to show me by press hotkey (Default is ALT + Space)"
    , "ALT + Space = Show / Hide window"
    , "Alt + F4 = Exit"
    , "Esc = Clear Input / Close window"
    , "Enter = Run current command"
    , "Alt + No. = Run that specific command"
    , "Ctrl + No. = Select that specific command"
    , "F1 = Show Help"
    , "F2 = Open Setting Config window"
    , "F3 = Edit config file (ALTRun.ini) directly"
    , "Down Arrow = Move to next command"
    , "Up Arrow = Move to previous command"
    , "Ctrl+H = Show command history"
    , "Ctrl+'+' = Increase rank of current command"
    , "Ctrl+'-' = Decrease rank of current command"
    , "Ctrl+I = Reindex file search database"
    , "Ctrl+Q = Reload ALTRun"
    , "Ctrl+D = Open current command dir with TC / File Explorer"
    , "Command priority (rank) will auto adjust based on frequency"
    , "Start with www or http = Open website"
    , "Start with + = Add New Command"
    , "Start with space = Search by Everything"]

, g_Hotkey1 := "^s", g_Trigger1 := "Everything"
, g_Hotkey2 := "^p", g_Trigger2 := "RunPTTools"
, g_Hotkey3 := "", g_Trigger3 := ""
, g_GlobalHotkey1 := "!Space", g_GlobalHotkey1Win := 0
, g_GlobalHotkey2 := "!R"    , g_GlobalHotkey2Win := 0
, g_RunCmdAlt := 0           , g_EnableCapsLockIME := 0
, OneDrive, OneDriveConsumer , OneDriveCommercial                       ; Environment Variables (Due to #NoEnv, some need to get from EnvGet)
EnvGet, OneDrive, OneDrive                                              ; OneDrive (Default)
EnvGet, OneDriveConsumer, OneDriveConsumer                              ; OneDrive for Personal
EnvGet, OneDriveCommercial, OneDriveCommercial                          ; OneDrive for Business

;=============================================================
; 声明全局变量
; 当前输入命令的参数,数组,为了方便没有添加 g_ 前缀
;=============================================================
global Arg                                    ; 用来调用管道的完整参数（所有列）
, FullPipeArg                                 ; 不能是 ALTRun.ahk 的子串,否则按键绑定会有问题
, g_WinName := "ALTRun - Ver 07.2023"         ; 主窗口标题
, g_OptionsWinName := "ALTRun Options"        ; 选项窗口标题
, g_CmdMgrWinName := "Command Manager"        ; 选项窗口标题
, g_Commands                                  ; 所有命令
, g_FallbackCommands                          ; 当搜索无结果时使用的命令
, g_CurrentInput                              ; 编辑框当前内容
, g_CurrentCommand                            ; 当前匹配到的第一条命令
, g_CurrentCommandList := Object()            ; 当前匹配到的所有命令
, g_UseDisplay                                ; 命令使用了显示框
, g_HistoryCommands := Object()               ; 历史命令
, g_UseFallback                               ; 使用备用的命令
, g_ExcludedCommands                          ; 排除的命令
, g_PipeArg                                   ; 用来调用管道的参数（结果第三列）
, g_InputBox  := "Edit1"
, g_ListView  := "MyListView"
, g_StatusBar := "Edit2"

;=============================================================
; 显示各个控件的ToolTip
;=============================================================
g_EnableSendTo_TT := "Whether to create a 'send to' menu"
g_AutoStartup_TT := "Whether to add a shortcut to boot"
g_InStartMenu_TT := "Whether to add a shortcut to the start menu"
g_IndexDir_TT := "Index location, you can use full path or AHK variable starting with A_, must be separated by ' | ', the path can contain spaces, without quotation marks"
g_IndexFileType_TT := "The index file types must be separated by '|'"
g_IndexExclude_TT := "excluded files, regular expression"
g_SearchFullPath_TT := "Search full path of the file or command, otherwise only search file name"
g_ShowFileExt_TT := "Show file extension on interface"
g_ShowIcon_TT := "Show icon in file ListView"
g_KeepInputText_TT := "Do not clear the content of the edit box when the window is hidden"
g_TCPath_TT := "TotalCommander path with parameters, eg: C:\OneDrive\Apps\TotalCMD64\Totalcmd64.exe /O /T /S /L=, if empty, use explorer to open"
g_SelectTCPath_TT := "Select Total Commander file path"
g_RunIfOnlyOne_TT := "Run directly if there is only one result"
g_HideOnDeactivate_TT := "The window closes after the window loses focus, and the window stay-on-top display function fails after activation"
g_AlwaysOnTop_TT := "Window on top most display"
g_SaveHistory_TT := "Save command history"
g_HistorySize_TT := "Number of recorded history"
g_isLogging_TT := "Enable or Disable log record"
g_AutoRank_TT := "Automatically adjust the command rank priority according to the frequency of use."
g_SwitchToEngIME_TT := "Automatically switch to English input method every time the window is activated"
g_EscClearInput_TT := "When inputting Esc, if there is content in the input box, it will be cleared, and if there is no content, the window will be closed"
g_Editor_TT := "The editor used to open the configuration file, the default is the editor associated with the resource manager, you can right-click->properties->open method to modify"
g_SendToGetLnk_TT := "If the file sent using the Send To menu is a .lnk shortcut, add the target file after reading the path from the file"
g_Everything_TT := "Everything.exe file path"

g_HideTitle_TT := "Hide Title Bar"
g_ShowTrayIcon_TT := "Whether to show the tray icon"
g_DisplayRows_TT := "The number of rows displayed in the list, if more than 9 rows, the shortcut key to locate this row will be invalid."
g_Col3Width_TT := "The width of the third column in the list"
g_Col4Width_TT := "Width of the fourth column in the list"
g_HideCol2_TT := "Hide the second column, that is, display a column of file and function `nIf hidden, the space of the second column will be allocated to the third column"
g_LVGrid_TT := "Show Grid in command ListView"
g_FontName_TT := "Font Name, eg. Default, Segoe UI, Microsoft Yahei"
g_FontSize_TT := "Font Size, Default is 10"
g_FontColor_TT := "Font Color, eg. cRed, cFFFFAA, cDefault"
g_WinWidth_TT := "Width of ALTRun app window"
g_EditHeight_TT := "Height of input box and detail box"
g_ListHeight_TT := "Command List Height"
g_CtrlColor_TT := "Set Color for Controls in Window"
g_WinColor_TT := "Window background color, including border color, current command detail box color, value can be like: White, Default, EBFFEB, 0xEBFFEB"
g_BackgroundPicture_TT := "Background picture, the background picture can only be displayed in the border part.`nIf there is a splash screen after using the picture, first adjust the size of the picture to solve the window size and improve the loading speed.`nIf the splash screen is still obvious, please Hollow and fill the position of the text box on the picture with a color similar to the text background, or modify it to the transparent color of png"

g_Hotkey1_TT := "Shortcut key 1`nkey=Default can cancel the key mapping in the code`n Note that the priority is higher than the default Alt + alphanumeric series keys, do not modify the Alt mapping without special reasons"
g_Trigger1_TT := "Function to be triggered by Hotkey 1"
g_Hotkey2_TT := "Shortcut key 2`nkey=Default can cancel the key mapping in the code`n Note that the priority is higher than the default Alt + alphanumeric series keys, if there is no special reason, do not modify the Alt mapping"
g_Trigger2_TT := "Function to be triggered by Hotkey 2"
g_Hotkey3_TT := "Shortcut key 3`nkey=Default can cancel the key mapping in the code`n Note that the priority is higher than the default Alt + alphanumeric series keys, do not modify the Alt mapping without special reasons"
g_Trigger3_TT := "Function to be triggered by Hotkey 3"
g_RunCmdAlt_TT := "Press Alt + command number to run the command, untick means Press command number to run the command"
g_EnableCapsLockIME_TT := "Use CapsLock to switch input methods (similar to macOS)"
g_GlobalHotkey1Win_TT := "Enable Win key"
g_GlobalHotkey1_TT := "Activate ALTRun global hotkey"
g_GlobalHotkey2Win_TT := "Enable Win key"
g_GlobalHotkey2_TT := "Activate ALTRun global hotkey"

g_AutoSwitchDir_TT := "Listary - Auto Switch Dir"
g_FileMgrClass_TT := "Takeover File Manager"
g_DialogClass_TT := "Class name of the Dialog Box which Listary Switch Dir will take effect"
g_ExcludeWinClass_TT := "Exclude those windows that not want Listary Switch Dir take effect"
g_EnableScheduler_TT := "Enable shutdown scheduled task"
g_ShutdownTime_TT := "Set timing for PC auto shutdown"

OK_TT := "Save and Apply the changes"
Cancel_TT := "Discard the changes"

;=============================================================
; Load Config and Set Logger
;=============================================================
Global Log := New Logger(g_LogFile)                                     ; Global Log so that can use in other Lib
Log.Msg("==== ALTRun is starting ====")

LoadConfig("initialize")                                                ; Load ini config, IniWrite will create it if not exist

;=============================================================
; Create ContextMenu and TrayMenu
;=============================================================
Menu, LV_ContextMenu, Add, Run`tEnter, ContextMenu                      ; ListView ContextMenu
Menu, LV_ContextMenu, Add, Open Path`tCtrl+D, OpenCurrentFileDir
Menu, LV_ContextMenu, Add, Copy Command, ContextMenu
Menu, LV_ContextMenu, Add
Menu, LV_ContextMenu, Add, New Command, AddCommand
Menu, LV_ContextMenu, Add, Edit Command`tF3, EditCurrentCommand
Menu, LV_ContextMenu, Add, User-defined Command`tF4, UserCommandList
Menu, LV_ContextMenu, Add
Menu, LV_ContextMenu, Add, History `tCtrl+H, ShowCmdHistory
Menu, LV_ContextMenu, Add, Options `tF2, Options
Menu, LV_ContextMenu, Add, Help `tF1, Help

Menu, LV_ContextMenu, Icon, Run`tEnter, Shell32.dll, -25
Menu, LV_ContextMenu, Icon, Open Path`tCtrl+D, Shell32.dll, -4
Menu, LV_ContextMenu, Icon, Copy Command, Shell32.dll, -243
Menu, LV_ContextMenu, Icon, New Command, Shell32.dll, -1
Menu, LV_ContextMenu, Icon, Edit Command`tF3, Shell32.dll, -16775
Menu, LV_ContextMenu, Icon, User-defined Command`tF4, Shell32.dll, -44
Menu, LV_ContextMenu, Icon, History `tCtrl+H, Shell32.dll, -16741
Menu, LV_ContextMenu, Icon, Options `tF2, Shell32.dll, -16826
Menu, LV_ContextMenu, Icon, Help `tF1, Shell32.dll, -24
Menu, LV_ContextMenu, Default, Run`tEnter                               ; 让 "Run" 粗体显示表示双击时会执行相同的操作.

if (g_ShowTrayIcon)
{
    Menu, Tray, Icon
    Menu, Tray, Icon, Shell32.dll, -25                                  ; if the index of an icon changes between Windows versions but the resource ID is consistent, refer to the icon by ID instead of index
    Menu, Tray, NoStandard
    Menu, Tray, Tip, %g_WinName%

    Menu, Tray, Add, Show, ActivateALTRun
    Menu, Tray, Add
    Menu, Tray, Add, Options `tF2, Options
    Menu, Tray, Add, ReIndex `tCtrl+I, ReindexFiles
    Menu, Tray, Add, Help `tF1, Help
    Menu, Tray, Add
    Menu, SubTray, Add, Script Info, TrayMenu                           ; Create one menu destined to become a submenu of the above menu.
    Menu, SubTray, Add, Script Help, TrayMenu
    Menu, SubTray, Add, Window Spy, TrayMenu
    Menu, Tray, Add, AutoHotkey, :SubTray                               ; Create a submenu in the first menu (a right-arrow indicator)
    Menu, Tray, Add,
    Menu, Tray, Add, Reload `tCtrl+Q, ALTRun_Reload                     ; Call ALTRun_Reload function with Arg=Reload `tCtrl+Q
    Menu, Tray, Add, Exit `tAlt+F4, ExitALTRun

    Menu, Tray, Icon, Show, Shell32.dll, -25
    Menu, Tray, Icon, Options `tF2, Shell32.dll, -16826
    Menu, Tray, Icon, ReIndex `tCtrl+I, Shell32.dll, -16776
    Menu, Tray, Icon, Help `tF1, Shell32.dll, -24
    Menu, Tray, Icon, AutoHotkey, %A_AhkPath%, -160
    Menu, Tray, Icon, Reload `tCtrl+Q, Shell32.dll, -16739
    Menu, Tray, Icon, Exit `tAlt+F4, Imageres.dll, -5102
    Menu, Tray, Default, Show
    Menu, Tray, Click, 1                                                ; Sets the number of clicks to activate the tray menu's default menu item.
}
;=============================================================
; Load commands database and command history
; Update "SendTo", "Startup", "StartMenu" lnk
;=============================================================
LoadCommands()

if (g_SaveHistory)
{
    LoadHistoryCommands()
}

Log.Msg("Updating 'SendTo' setting...Re=" UpdateSendTo(g_EnableSendTo))
Log.Msg("Updating 'Startup' setting...Re=" UpdateStartup(g_AutoStartup))
Log.Msg("Updating 'StartMenu' setting...Re=" UpdateStartMenu(g_InStartMenu))

;=============================================================
; 主窗口配置代码
;=============================================================
HideTitle    := g_HideTitle ? "-Caption" : ""                           ; Store "-Caption" in HideTitle if g_HideTitle is True, otherwise store ""
AlwaysOnTop  := g_AlwaysOnTop ? "+AlwaysOnTop" : ""                     ; Check Win AlwaysOnTop status
Col2Width    := g_HideCol2 ? 0 : 60                                     ; Check ListView column 2 hide status
LVGrid       := g_LVGrid ? "Grid" : ""                                  ; Check ListView Grid option
WinHeight    := g_EditHeight + g_ListHeight + 30 + 23                   ; Original WinHeight
WinY         := (A_ScreenHeight - WinHeight) / 3                        ; Set Win location
ListWidth    := g_WinWidth - 20
HideWin      := ""

Gui, Main:Color, %g_WinColor%, %g_CtrlColor%
Gui, Main:Font, c%g_FontColor% s%g_FontSize%, %g_FontName%
Gui, Main:%HideTitle% %AlwaysOnTop%
Gui, Main:Add, Picture, x0 y0 0x4000000, %g_BGPicture%                  ; If the picture cannot be loaded or displayed, the control is left empty and its W&H are set to zero. So FileExist() is not necessary.
Gui, Main:Add, Edit, x10 y10 w%ListWidth% h%g_EditHeight% -WantReturn v%g_InputBox% gSearchCommand, Type anything here to search...
Gui, Main:Add, ListView, Count15 y+10 w%ListWidth% h%g_ListHeight% v%g_ListView% gLVAction +LV0x00010000 %LVGrid% -Multi AltSubmit, No.|Type|Command|Description ; LVS_EX_DOUBLEBUFFER Avoids flickering.
Gui, Main:Add, StatusBar, v%g_StatusBar%, Nice! You have run shortcut %g_RunCount% times by now!
Gui, Main:Add, Button, x0 y0 w0 h0 Hidden Default gRunCurrentCommand
Gui, Main:Default                                                       ; Set default GUI before any ListView update

LV_ModifyCol(1, "40 Integer")                                           ; set ListView column width and format, Integer can use for sort
LV_ModifyCol(2, Col2Width)
LV_ModifyCol(3, g_Col3Width)
LV_ModifyCol(4, g_Col4Width)
LV_Modify(0, "-Select")                                                 ; De-select all.
LV_Modify(1, "Select Focus Vis")                                        ; select 1st row
ListResult("Function | F1 | ALTRun Help Index`n"                        ; Show initial list (hints, help, statusbar) on firstRun
         . "Function | F2 | ALTRun Options Settings`n"
         . "Function | F3 | ALTRun Edit current command`n"
         . "Function | F4 | ALTRun User-defined command`n"
         . "Function | ALT+SPACE or ALT+R | Activative ALTRun`n"
         . "Function | Lose Focus or Hotkey or ESC | Close ALTRun`n"
         . "Function | Enter or ALT+No. | Run selected command`n"
         . "Function | UP or DOWN | Select previous or next command`n"
         . "Function | CTRL+D | Open cmd dir with TC or File Explorer"
         , False, False, False)
SB_SetParts(g_WinWidth-120)

;=============================================================
; Command line args, Args are %1% %2% or A_Args[1] A_Args[2]
;=============================================================
Log.Msg("Resolving command line Args=" A_Args[1] " " A_Args[2] " " A_Args[3])

if (A_Args[1] = "-startup" or A_Args[2]="-hide")                        ; App run from startup
{
    HideWin := " Hide"
    Log.Msg("Starting ALTRun from Startup lnk or in Silent Mode...")
}

if (A_Args[1] = "-SendTo")
{
    HideWin := " Hide"
    Log.Msg("Starting ALTRun SendTo Mode...")
    CmdMgr(A_Args[2], A_Args[3])
}

Gui, Main:Show, xCenter y%WinY% w%g_WinWidth% h%WinHeight% %HideWin%, %g_WinName%

if (g_SwitchToEngIME)
{
    SwitchToEngIME()
}

if (g_HideOnDeactivate)
{
    OnMessage(0x06, "WM_ACTIVATE")
}

OnMessage(0x0200, "WM_MOUSEMOVE")
OnExit("ExitFunc")

;=============================================================
; Set Hotkey for %g_WinName% only
;=============================================================
Hotkey, IfWinActive, %g_WinName%                                        ; Hotkey take effect only when ALTRun actived

Hotkey, !F4, ExitALTRun
Hotkey, Tab, TabFunc
Hotkey, F1, Help
Hotkey, F2, Options
Hotkey, F3, EditCurrentCommand
Hotkey, F4, UserCommandList
Hotkey, ^q, ALTRun_Reload
Hotkey, ^d, OpenCurrentFileDir
Hotkey, ^i, ReindexFiles
Hotkey, ^H, ShowCmdHistory
Hotkey, ^NumpadAdd, IncreaseRank
Hotkey, ^NumpadSub, DecreaseRank
Hotkey, Down, NextCommand
Hotkey, Up, PrevCommand

;=============================================================
; Run or locate command shortcut: Ctrl Alt Shift + No.
;=============================================================
Loop, % Min(g_DisplayRows, 9)                                           ; Not set hotkey for DisplayRows > 9
{
    if (g_RunCmdAlt)
        Hotkey, !%A_Index%, RunSelectedCommand                          ; ALT + No. run command
    else
        Hotkey, %A_Index%, RunSelectedCommand                           ; 数字键直接启动,小键盘不影响数字输入
    Hotkey, ^%A_Index%, GotoCommand                                     ; Ctrl + No. locate command
}

;=============================================================
; Set Trigger <-> Hotkey and Global Hotkey
;=============================================================
Loop, 3
{
    Hotkey  := % g_Hotkey%A_Index%
    Trigger := % g_Trigger%A_Index%

    if (Hotkey != "" and IsFunc(Trigger))
        Hotkey, %Hotkey%, %Trigger%
}

Hotkey, IfWinActive                                                     ; Omit the parameters to turn off context sensitivity, to make subsequently-created hotkeys work in all windows
if (g_GlobalHotkey1 != "")
{
    Win := g_GlobalHotkey1Win ? "#" : ""
    Hotkey, %Win%%g_GlobalHotkey1%, ToggleWindow
}

if (g_GlobalHotkey2 != "")
{
    Win := g_GlobalHotkey2Win ? "#" : ""
    Hotkey, %Win%%g_GlobalHotkey2%, ToggleWindow
}

;=============================================================
; Don't use many SetTimer or GoSub, will affect each other
; Set TaskScheduler to turn off computer
;=============================================================
Listary()                                                               ; 启动Listary快速切换文件夹功能
AppControl()                                                            ; 启动AppControl功能
TaskScheduler(g_EnableScheduler)                                        ; Task Scheduler for PC Shutdown
;=============================================================

ActivateALTRun()
{
    Gui, Main:Show,, %g_WinName%
    Gui, Main:Default                                                   ; Set default GUI window before any ListView / StatusBar operate

    WinWaitActive, %g_WinName%,, 2                                      ; Use WinWaitActive 2s instead of previous Loop method
    {
        if (g_SwitchToEngIME)
        {
            SwitchToEngIME()
        }

        GuiControl, Main:Focus, %g_InputBox%
        if (g_KeepInputText)
        {
            ControlSend, %g_InputBox%, ^a, %g_WinName%                  ; 如设置为保留输入框内容,则全选
        }
        else
        {
            GuiControl, Main:Text, %g_InputBox%,                        ; 如设置为清空输入框内容
        }
        SetStatusBar(False)                                             ; 每次有效激活窗口之后StatusBar展示提示信息
    }
    if ErrorLevel
    {
        MsgBox, 4096, Warning, WinWaitActive timed out.
    }
}

ToggleWindow()
{
    if WinActive(g_WinName)
    {
        MainGuiClose()
    }
    else
    {
        ActivateALTRun()
    }
}

SearchCommand(command := "", firstRun := false)
{
    Global
    GuiControlGet, g_CurrentInput, Main:,%g_InputBox%                   ; Get input command string
    
    g_UseDisplay    := false
    result          := ""
    fullResult      := ""                                               ; 供去重使用
    commandPrefix   := SubStr(g_CurrentInput, 1, 1)
    
    if (commandPrefix = "+" or commandPrefix = " " or commandPrefix = ">")
    {
        g_PipeArg := ""

        if (commandPrefix = "+")
        {
            g_CurrentCommand := g_FallbackCommands[1]
        }
        else if (commandPrefix = " ")
        {
            g_CurrentCommand := g_FallbackCommands[2]
        }
        else if (commandPrefix = ">")
        {
            g_CurrentCommand := g_FallbackCommands[5]
        }
        g_CurrentCommandList := Object()
        g_CurrentCommandList.Push(g_CurrentCommand)
        result .= g_CurrentCommand
        ListResult(result, false, false)
        Return result
    }

    g_CurrentCommandList := Object()
    order := 1

    for index, element in g_Commands
    {
        if (InStr(fullResult, element "`n") or inStr(g_ExcludedCommands, element "`n"))
        {
            continue
        }

        splitedElement := StrSplit(element, " | ")
        _Type := splitedElement[1]
        _Path := splitedElement[2]
        _Desc := splitedElement[3]

        if (_Type = "file")                                             ; Equal (=), case-sensitive-equal (==)
        {
            SplitPath, _Path, fileName, fileDir, , nameNoExt

            elementToShow   := _Type " | " _Path " | " _Desc            ; Use _Path to show file icons
            elementToSearch := _Type " " fileName " " _Desc             ; search file name include extension

            if (g_SearchFullPath)
            {
                elementToSearch := StrReplace(fileDir, "\", " ") " " elementToSearch ; 搜索路径时强行将 \ 转成空格
            }
        }
        else if _Type in dir,tender,project
        {
            SplitPath, _Path, fileName, Dir, Ext, nameNoExt, Drive      ; Extra name from _Path (if _Type is Dir and has "." in path, nameNoExt will not get full folder name) 

            elementToShow   := _Type " | " fileName " | " _Desc         ; Show folder name only
            elementToSearch := _Type " " fileName " " _Desc             ; Search dir type + folder name + desc

            if (g_SearchFullPath)
            {
                elementToSearch := StrReplace(_Path, "\", " ") " " elementToSearch ; 搜索路径时强行将 \ 转成空格
            }
        }
        else
        {
            elementToShow   := _Type " | " _Path " | " _Desc
            elementToSearch := StrReplace(_Path, "/", " ")
            elementToSearch := StrReplace(_Path, "\", " ")
            elementToSearch := _type " " elementToSearch " " _Desc
        }

        if (g_CurrentInput = "" or FuzzyMatch(elementToSearch, g_CurrentInput))
        {
            fullResult .= element "`n"
            g_CurrentCommandList.Push(element)

            if (order = 1)
            {
                g_CurrentCommand := element
                result .= elementToShow
                order++
            }
            else
            {
                result .= "`n" elementToShow
                order++
            }

            if (order > g_DisplayRows)
            {
                break
            }
        }
    }

    if (result = "")
    {
        EvalResult := Eval(g_CurrentInput)
        if (EvalResult != "")
        {
            ListResult("Eval | " EvalResult, false, true)
            Return
        }

        g_UseFallback        := true
        g_CurrentCommandList := g_FallbackCommands
        g_CurrentCommand     := g_FallbackCommands[1]

        for index, element in g_FallbackCommands
        {
            if (index = 1)
            {
                result .= element
            }
            else
            {
                result .= "`n" element
            }
        }
    }
    else
    {
        g_UseFallback := false
    }

    ListResult(result, false, false)
}

ListResult(text := "", ActWin := false, UseDisplay := false, UpdateSB := true) ; 用来显示控制界面 & 用来显示命令结果
{
    if (ActWin)
    {
        ActivateALTRun()                                                ; 会导致快捷计算器失效
    }
    g_UseDisplay := UseDisplay
    
    Gui, Main:Default                                                   ; Set default GUI before any LV update methods
    GuiControl, Main:-Redraw, %g_ListView%                              ; 在加载时禁用重绘来提升性能.
    LV_Delete()

    if (g_ShowIcon)
    {
        ImageListID1 := IL_Create(10, 5)                                ; Create an ImageList so that the ListView can display some icons
        IL_Add(ImageListID1, "shell32.dll", -4)                         ; Add folder icon for dir type (IconNumber=1)
        IL_Add(ImageListID1, "shell32.dll", -25)                        ; Add app default icon for function type (IconNumber=2)
        IL_Add(ImageListID1, "shell32.dll", -512)                       ; Add Browser icon for url type (IconNumber=3)
        IL_Add(ImageListID1, "shell32.dll", -22)                        ; Add control panel icon for control type (IconNumber=4)
        LV_SetImageList(ImageListID1)                                   ; Attach the ImageLists to the ListView so that it can later display the icons
        IconNumber  := ""
        sfi_size    := A_PtrSize + 8 + 680                              ; 计算 SHFILEINFO 结构需要的缓存大小
        VarSetCapacity(sfi, sfi_size)
    }
    
    Loop Parse, text, `n, `r
    {
        if (!InStr(A_LoopField, " | "))                                 ; If do not have " | " then Return result and next line
        {
            _Type    := "Display"
            _Path    := A_LoopField
            _Desc    := ""
        }        
        else
        {
            _Type    := Trim(StrSplit(A_LoopField, " | ")[1])
            _Path    := Trim(StrSplit(A_LoopField, " | ")[2])           ; Must store in var for future use, trim space
            _Desc    := Trim(StrSplit(A_LoopField, " | ")[3])
        }

        _AbsPath := StrReplace(_Path,  "*RunAs ", "")                   ; Remove *RunAs (Admin Run) to get absolute path
        _AbsPath := StrReplace(_AbsPath, "%OneDrive%", OneDrive)                     ; Convert OneDrive to absolute path due to #NoEnv
        _AbsPath := StrReplace(_AbsPath, "%OneDriveConsumer%", OneDriveConsumer)     ; Convert OneDrive to absolute path due to #NoEnv
        _AbsPath := StrReplace(_AbsPath, "%OneDriveCommercial%", OneDriveCommercial) ; Convert OneDrive to absolute path due to #NoEnv
        
        ; 建立唯一的扩展 ID 以避免变量名中的非法字符, 例如破折号. 这种使用唯一 ID 的方法也会执行地更好, 因为在数组中查找项目不需要进行搜索循环.
        SplitPath, _AbsPath,,, FileExt                                  ; 获取文件扩展名.

        if (g_ShowIcon)
        {
            if _Type contains Dir,Tender,Project
            {
                ExtID := "dir", IconNumber := 1
            }
            else if _Type contains Function,CMD
            {
                ExtID := "cmd", IconNumber := 2
            }
            else if _Type contains URL
            {
                ExtID := "url", IconNumber := 3
            }
            else if _Type contains Control
            {
                ExtID := "cpl", IconNumber := 4
            }
            else if FileExt in EXE,ICO,ANI,CUR
            {
                ExtID := FileExt, IconNumber := 0                       ; ExtID 特殊 ID 作为占位符, IconNumber 进行标记这样每种类型就含有唯一的图标.
            }
            else                                                        ; 其他的扩展名/文件类型, 计算它们的唯一 ID.
            {
                ExtID := 0                                              ; 进行初始化来处理比其他更短的扩展名.
                Loop 4                                                  ; 限制扩展名为 4 个字符, 这样之后计算的结果才能存放到 64 位值 (use 4 due to some short folder name has dot)
                {
                    ExtChar := SubStr(FileExt, A_Index, 1)
                    if not ExtChar                                      ; 没有更多字符了.
                        break
                    ExtID := ExtID | (Ord(ExtChar) << (8 * (A_Index - 1))) ; 把每个字符与不同的位位置进行运算来得到唯一 ID
                }
                IconNumber := IconList%ExtID%                           ; 检查此文件扩展名的图标是否已经在图像列表中. 如果是, 可以避免多次调用并极大提高性能, 尤其对于包含数以百计文件的文件夹而言
            }

            if (!IconNumber)                                            ; 此扩展名还没有相应的图标, 所以进行加载.
            {
                ; 获取与此文件扩展名关联的高质量小图标:
                if (!DllCall("Shell32\SHGetFileInfoW", "Str", _AbsPath, "UInt", 0, "Ptr", &sfi, "UInt", sfi_size, "UInt", 0x101))  ; 0x101 为 SHGFI_ICON+SHGFI_SMALLICON
                    IconNumber = 9999999                                ; 如果未成功加载到图标, 把它设置到范围外来显示空图标.
                else                                                    ; 成功加载图标.
                {
                    hIcon := NumGet(sfi, 0)                             ; 从结构中提取 hIcon 成员
                    IconNumber := DllCall("ImageList_ReplaceIcon", "ptr", ImageListID1, "int", -1, "ptr", hIcon) + 1 ; 直接添加 HICON 到图标列表, 下面加上 1 来把返回的索引从基于零转换到基于1
                    DllCall("DestroyIcon", "ptr", hIcon)                ; 现在已经把它复制到图像列表, 所以应销毁原来的
                    IconList%ExtID% := IconNumber                       ; 缓存图标来节省内存并提升加载性能:
                }
            }
            LV_Add("Icon" . IconNumber, A_Index, _Type, _Path, _Desc)
        }
        else
        {
            LV_Add(, A_Index, _Type, _Path, _Desc)
        }
    }

    LV_Modify(0, "-Select")                                             ; De-select all.
    LV_Modify(1, "Select Focus Vis")                                    ; select 1st row
    GuiControl, Main:+Redraw, %g_ListView%                              ; 重新启用重绘 (上面把它禁用了)

    if (g_CurrentCommandList.Length() = 1 and g_RunIfOnlyOne)
    {
        RunCommand(g_CurrentCommand)
    }

    if (UpdateSB)
        SetStatusBar(True)
}

RunCommand(originCmd)
{
    MainGuiClose()                                                      ; 先隐藏或者关闭窗口,防止出现延迟的感觉
    ParseArg()

    g_UseDisplay := false

    _Type := StrSplit(originCmd, " | ")[1]
    _Path := StrSplit(originCmd, " | ")[2]
    _Path := StrReplace(_Path, "%OneDrive%", OneDrive)                  ; Convert OneDrive to absolute path (due to #NoEnv)
    _Path := StrReplace(_Path, "%OneDriveConsumer%", OneDriveConsumer)  ; Convert OneDrive to absolute path (due to #NoEnv)
    _Path := StrReplace(_Path, "%OneDriveCommercial%", OneDriveCommercial) ; Convert OneDrive to absolute path (due to #NoEnv)
    Log.Msg("Execute(" g_RunCount ")=" originCmd)

    if (_Type = "file")
    {
        SplitPath, _Path, , WorkingDir, ,
        if (Arg = "")
        {
            Run, %_Path%, %WorkingDir%, UseErrorLevel
        }
        else
        {
            Run, %_Path% "%Arg%", %WorkingDir%, UseErrorLevel
        }
        if ErrorLevel
            MsgBox Could not open "%_Path%"
    }
    else if _Type in dir,tender,project
    {
        OpenPath(_Path)
    }
    else if (_Type = "function" and IsFunc(_Path))
    {
        %_Path%()
    }
    else if (_Type = "cmd")
    {
        RunWithCmd(_Path)
    }
    else                                                                ; for type: url, control & all other un-defined type
    {
        Run, %_Path%
    }

    if (g_SaveHistory && _Path != "ShowCmdHistory")                     ; Save command history
    {
        if (Arg != "")
        {
            g_HistoryCommands.InsertAt(1, originCmd " | " Arg)
        }
        else
        {
            g_HistoryCommands.InsertAt(1, originCmd)
        }

        if (g_HistoryCommands.Length() > g_HistorySize)
        {
            g_HistoryCommands.Pop()
        }
    }

    if (g_AutoRank)
    {
        ChangeRank(originCmd)
    }

    g_PipeArg := ""
    FullPipeArg := ""
}

TabFunc()
{
    GuiControlGet, CurrCtrl, Main:FocusV                                ; Limit tab to switch between g_InputBox & ListView only
    if (CurrCtrl = g_InputBox)
    {
        GuiControl, Main:Focus, %g_ListView%
    }
    else
    {
        GuiControl, Main:Focus, %g_InputBox%
    }
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
    {
        ChangeCommand(index - 1, True)
    }
}

ChangeCommand(Step = 1, ResetSelRow = false)
{
    Gui, Main:Default                                                   ; Use it before any LV update

    if (ResetSelRow)
        SelRow := 1
    else
        SelRow := LV_GetNext()                                          ; Get selected row no.

    SelRow += Step
    if (SelRow > LV_GetCount())                                          ; Listview cycle selection
        SelRow := 1
    else if (SelRow < 1)
        SelRow := LV_GetCount()

    LV_Modify(0, "-Select")
    LV_Modify(SelRow, "Select Focus Vis")                               ; make new index row selected, Focused, and Visible

    g_CurrentCommand := g_CurrentCommandList[SelRow]                    ; Get current command from selected row
    SetStatusBar(True)
}

;=============================================================
; GuiContextMenu right click on GUI Control
;=============================================================
MainGuiContextMenu()                                                    ; 运行此标签来响应右键点击或按下 Appskey, 指定响应窗口为Main
{
    if (A_GuiControl = g_ListView)                                      ; 仅在 ListView 中点击时才显示菜单
        Menu, LV_ContextMenu, Show, %A_GuiX%, %A_GuiY%                  ; 在提供的坐标处显示菜单, 应该使用 A_GuiX & A_GuiY,因为即使用户按下 Appskey 它们也会提供正确的坐标
}

ContextMenu()                                                           ; ListView ContextMenu actions
{
    Gui, Main:Default                                                   ; Use it before any LV update
    focusedRow := LV_GetNext(0, "Focused")                              ; Check focused row, only operate focusd row instead of all selected rows
    if (!focusedRow)                                                    ; if not found
        return

    g_CurrentCommand := g_CurrentCommandList[focusedRow]                ; Get current command from focused row
    If (A_ThisMenuItem = "Run`tEnter")                                  ; 用户在上下文菜单中选择了 "Run`tEnter"
    {
        RunCommand(g_CurrentCommand)
    }
    else if (A_ThisMenuItem = "Copy Command")
    {
        A_Clipboard := StrSplit(g_CurrentCommand, " | ")[2]
    }
}

LVAction()                                                              ; Double click and normal left click on ListView behavior
{
    focusedRow := LV_GetNext(0, "Focused")                              ; 查找焦点行, 仅对焦点行进行操作而不是所有选择的行:
    if (!focusedRow)                                                    ; 没有焦点行
        return

    g_CurrentCommand := g_CurrentCommandList[focusedRow]                ; Get current command from focused row
    
    if (A_GuiEvent = "DoubleClick" and g_CurrentCommand)                ; Double click behavior, if g_CurrentCommand = "" eg. first tip page, run it will clear SEC_USERCMD, SEC_INDEX, SEC_DFTCMD
    {
        RunCommand(g_CurrentCommand)
    }
    else if (A_GuiEvent = "Normal")                                     ; left click behavior
    {
        SetStatusBar(True)
    }
}

TrayMenu()                                                              ;AutoHotkey标准托盘菜单
{
    If ( A_ThisMenuItem = "Script Info" )
        ListVars
    If ( A_ThisMenuItem = "Script Help" )
        Run % A_AhkPath
    If ( A_ThisMenuItem = "Window Spy" )
        Run, % StrReplace(A_AhkPath, "\AutoHotkey.exe", "\WindowSpy.ahk")
}

MainGuiEscape()
{
    ToolTip
    if (g_EscClearInput and g_CurrentInput)
    {
        ClearInput()
    }
    else
    {
        MainGuiClose()
    }
}

MainGuiClose()                                                          ; If GuiClose is a function, the GUI is hidden by default
{
    if (!g_KeepInputText)
    {
        ClearInput()
    }
    Gui, Main:Hide
}

ExitALTRun()
{
    ExitApp
}

ALTRun_Reload(Mode := "")
{
    Log.Msg("ALTRun Reloading... Mode="Mode)
    If (Mode = "Silent")
    {
        Run "%A_ScriptFullPath%" /restart -hide
    }
    Else
    {
        Reload
    }
}

ExitFunc(exitReason, exitCode)
{
    SaveConfig("History")                                               ; Save ini in OnExit function, ExitApp/Reload will auto call OnExit.
    Log.Msg("Exiting ALTRun...Reason=" exitReason)
}

ALTRun_Log()
{
    if (g_Editor != "")
    {
        Run, % g_Editor " /m " A_Now " """ g_LogFile """"               ; /m Match text, /g Jump to specified position, /g -1 means end of file.
    }
    else
    {
        Run, % g_LogFile
    }
}

UserCommandList()
{
    if (g_Editor != "")
    {
        Run, % g_Editor " /m " SEC_USERCMD " """ g_IniFile """"         ; /m Match text
    }
    else
    {
        Run, % g_IniFile
    }
}


AddCommand()
{
    Path := """Command Path to specific..."""
    Desc := """" Arg """"
    CmdMgr(Path, Desc)
}

ClearInput()
{
    GuiControl, Main:Text, %g_InputBox%,
    GuiControl, Main:Focus, %g_InputBox%
}

SetStatusBar(currentCommandMode := True)                                ; StatusBar shows currentCommand (true) or hints (false)
{
    if (currentCommandMode)
    {
        SBText :=StrSplit(g_CurrentCommand, " | ")[2]
    }
    else
    {
        g_RunCount ++
        Random, HintIndex, 1, g_Hints.Length()                          ; 随机抽出一条提示信息
        SBText := g_Hints[HintIndex]                                    ; 每次有效激活窗口之后StatusBar展示提示信息
    }
    SB_SetText(SBText, 1)                                               ; Omite SB_SetIcon for better performance
    SB_SetText("RunCount: "g_RunCount, 2)

}

RunCurrentCommand()
{
    RunCommand(g_CurrentCommand)
}

ParseArg()
{
    global
    if (g_PipeArg != "")
    {
        Arg := g_PipeArg
        Return
    }

    commandPrefix := SubStr(g_CurrentInput, 1, 1)

    if (commandPrefix = "+" || commandPrefix = " " || commandPrefix = ">") ; 分号或者冒号的情况,直接取命令为参数
    {
        Arg := SubStr(g_CurrentInput, 2)
        Return
    }

    if (InStr(g_CurrentInput, " ") && !g_UseFallback)                   ; 用空格来判断参数
    {
        Arg := SubStr(g_CurrentInput, InStr(g_CurrentInput, " ") + 1)
    }
    else if (g_UseFallback)
    {
        Arg := g_CurrentInput
    }
    else
    {
        Arg := ""
    }
}

FuzzyMatch(Haystack, Needle)
{
    Needle := StrReplace(Needle, " ", ".*")                             ; RegExMatch should able to search with space as & separater, but do not know now, use this way first.
    Return RegExMatch(Haystack, "im)" Needle)                           ; Use RegExMatch replace InStr & TCMatch
}

ChangeRank(originCmd, showToolTip := false, inc := 1)
{
    RANKSEC := SEC_DFTCMD "|" SEC_USERCMD "|" SEC_INDEX
    Loop Parse, RANKSEC, |                                              ; Update Rank for related sections
    {
        IniRead, Rank, %g_IniFile%, %A_LoopField%, %originCmd%, KeyNotFound

        if (Rank = "KeyNotFound" or Rank = "ERROR" or originCmd = "")   ; If originCmd not exist in this section, then check next section
        {
            continue                                                    ; Skips the rest of a loop and begins a new one.
        }
        else if Rank is integer                                         ; If originCmd exist in this section, then update it's rank.
        {
            Rank += inc
        }
        else
        {
            Rank := inc
        }

        if (Rank < 0)                                                   ; 如果降到负数,都设置成 -1,然后屏蔽
        {
            Rank := -1
            g_ExcludedCommands .= originCmd "`n"
        }
        IniWrite, %Rank%, %g_IniFile%, %A_LoopField%, %originCmd%       ; Update new Rank for originCmd

        if (showToolTip)
        {
            ToolTip, Adjust Command Rank: `n%originCmd% : %Rank%, 100, 150
            SetTimer, RemoveToolTip, -800
        }
    }
    LoadCommands()                                                      ; New rank will take effect in real-time by LoadCommands again
}

RunSelectedCommand()
{
    if (SubStr(A_ThisHotkey, 1, 1) = "~")
    {
        GuiControlGet, CurrCtrl, Main:FocusV
        if (CurrCtrl = g_InputBox)
        {
            Return
        }
    }

    index := SubStr(A_ThisHotkey, 0, 1)
    RunCommand(g_CurrentCommandList[index])
}

IncreaseRank()
{
    ChangeRank(g_CurrentCommand, true)
}

DecreaseRank()
{
    ChangeRank(g_CurrentCommand, true, -1)
}

LoadCommands(LoadRank := True)
{
    Log.Msg("Loading commands list...")
    g_Commands := Object()                                              ; Clear g_Commands list
    g_FallbackCommands := Object()                                      ; Clear g_FallbackCommands list
    RankString := ""
    g_ExcludedCommands := ""

    RANKSEC := LoadConfig("CMD_SECS")                                   ; Read built-in command & user commands and index commands whole sections
    Loop Parse, RANKSEC, `n                                             ; read each line, separate key and value
    {
        command := StrSplit(A_LoopField, "=")[1]                        ; pass first string (key) to command
        rank    := StrSplit(A_LoopField, "=")[2]                        ; pass second string (value) to rank

        if (StrLen(command) > 0)
        {
            if (rank >= 1)
            {
                RankString .= rank "`t" command "`n"
            }
            else
            {
                g_ExcludedCommands .= command "`n"
            }
        }
    }

    if (RankString != "")
    {
        Sort, RankString, R N
        Loop Parse, RankString, `n
        {
            if (A_LoopField = "")
            {
                continue
            }
            g_Commands.Push(StrSplit(A_LoopField, "`t")[2])
        }
    }

    FALLBACKCMDSEC := LoadConfig("FBCMD")                               ;read whole section, initialize it if section not exist
    Loop Parse, FALLBACKCMDSEC, `n                                      ;read each line, get each FBCommand (Rank not necessary)
    {
        FBCommand  := StrSplit(A_LoopField, " | ")[2]
        if (IsFunc(FBCommand))
        {
            g_FallbackCommands.Push(A_LoopField)
        }
    }
    Return "Loaded"
}

LoadHistoryCommands()
{
    Loop %g_HistorySize%
    {
        IniRead, HistoryCommand, %g_IniFile%, %SEC_HISTORY%, %A_Index%
        g_HistoryCommands.Push(HistoryCommand)
    }
}

ShowCmdHistory()
{
    result := ""
    g_CurrentCommandList := Object()

    for i, element in g_HistoryCommands
    {
        if (i = 1)
        {
            g_CurrentCommand := element
        }
        else
        {
            result .= "`n"
        }

        _Type   := StrSplit(element, " | ")[1]
        _Path   := StrSplit(element, " | ")[2]
        _Desc   := StrSplit(element, " | ")[3]
        _Arg    := StrSplit(element, " | ")[4]
        result  .= _Type " | " _Path " | " _Desc " #Arg： " _Arg

        g_CurrentCommandList.Push(element)
    }
    ListResult(result, true, true)
}

GetCmdOutput(command)
{
    TempFileName := "ALTRun.stdout.log"
    FullCommand = %ComSpec% /C "%command% > %TempFileName%"

    RunWait, %FullCommand%, %A_Temp%, Hide
    FileRead, result, %A_Temp%\%TempFileName%
    FileDelete, %A_Temp%\%TempFileName%
    result := RTrim(result, "`r`n")                                     ; Remove result rightmost/last "`r`n"
    Return result
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

OpenPath(filePath)
{
    filePath := StrReplace(filePath, "*RunAs ", "")                     ; 如果是以管理员权限运行,则去掉开头的 *RunAs 以得到正确的路径

    if (InStr(filePath, "A_"))                                          ; Condier path like A_ScriptDir
    {
        Path := %filePath%
    }
    else
    {
        Path := filePath
    }

    if (g_TCPath)
    {
        Run, %g_TCPath% "%Path%",, UseErrorLevel                        ; /S switch TC /L as Source, /R as Target. /O: If TC is running, active it. /T: open in new tab
    }
    else
    {
        Run, Explorer.exe /select`, "%Path%",, UseErrorLevel
    }
    if ErrorLevel
    {
        MsgBox, 4096, %g_WinName%, Error found, error code : %A_LastError%
    }
    Log.Msg("Opening path="Path)
}

OpenCurrentFileDir()
{
    filePath := StrSplit(g_CurrentCommand, " | ")[2]
    OpenPath(filePath)
}

EditCurrentCommand()
{
    if (g_Editor != "")
    {
        Run, % g_Editor " /m " """" g_CurrentCommand """" " """ g_IniFile """" ; /m Match text, locate to current command
    }
    else
    {
        Run, % g_IniFile
    }
}

;=============================================================
; MouseMove Behavior
;=============================================================
WM_MOUSEMOVE(wParam, lParam)
{
    if (wParam = 1)                                                     ; LButton
    {
        PostMessage, 0xA1, 2, , , A                                     ; WM_NCLBUTTONDOWN
    }

    if WinActive(g_OptionsWinName)                                      ; to show tooltip for controls, limit to Options Win
    {
        static CurrControl := "", PrevControl := "", _TT := ""          ; _TT is kept blank for use by the ToolTip command below.
        CurrControl := A_GuiControl
        if (CurrControl != PrevControl)                                 ; ToolTip for each control
        {
            ToolTip                                                     ; Turn off any previous tooltip.
            Try                                                         ; Using try will guard "CurrControl" against all possible errors
                ToolTip, % %CurrControl%_TT                             ; The leading percent sign tell it to use an expression.
            Catch
                ToolTip
            SetTimer, RemoveToolTip, -10000
            PrevControl := CurrControl
        }
    }
}

RemoveToolTip()
{
    ToolTip
    SetTimer, RemoveToolTip, Off
}

;=============================================================
; 窗口失去焦点后隐藏
;=============================================================
WM_ACTIVATE(wParam, lParam)
{
    if (wParam >= 1)                                                    ; 窗口激活
    {
        Return
    }
    else if (wParam <= 0)                                               ; 窗口非激活,这样有可能第一次显示主界面时,窗口失去焦点后不关闭
    {
        if (WinExist(g_WinName) && !g_UseDisplay)
        {
            MainGuiClose()
        }
    }
}

UpdateSendTo(create := true)                                            ; the lnk in SendTo must point to a exe
{
    lnkPath := StrReplace(A_StartMenu, "\Start Menu", "\SendTo\") "ALTRun.lnk"

    if (!create)
    {
        FileDelete, %lnkPath%
        Return "SendTo lnk cleaned up"
    }

    if (A_IsCompiled)
        FileCreateShortcut, "%A_ScriptFullPath%", %lnkPath%, ,-SendTo
        , Send command to ALTRun User Command list, Shell32.dll, , -25
    else
        FileCreateShortcut, "%A_AhkPath%", %lnkPath%, , "%A_ScriptFullPath%" -SendTo
        , Send command to ALTRun User Command list, Shell32.dll, , -25
    Return "updated"
}

UpdateStartup(create := true)
{
    lnkPath := A_Startup "\ALTRun.lnk"

    if (!create)
    {
        FileDelete, %lnkPath%
        Return "Startup lnk cleaned up"
    }

    FileCreateShortcut, %A_ScriptFullPath%, %lnkPath%, %A_ScriptDir%
        , -startup, ALTRun - An effective launcher, Shell32.dll, , -25
    Return "updated"
}

UpdateStartMenu(create := true)
{
    lnkPath := A_Programs "\ALTRun.lnk"

    if (!create)
    {
        FileDelete, %lnkPath%
        Return "Start Menu lnk cleaned up"
    }

    FileCreateShortcut, %A_ScriptFullPath%, %lnkPath%, %A_ScriptDir%
        , -StartMenu, ALTRun, Shell32.dll, , -25
    Return "updated"
}

TaskScheduler(SchEnable = False)
{

    TaskSch_Clean = SchTasks /delete /TN AHK_Shutdown /F                ; Delete old scheduler, /TN: TaskName
    Log.Msg("Cleaning task scheduler...Re=" GetCmdOutput(TaskSch_Clean)) ; Run and get output, record into log
    
    if (SchEnable)                                                      ; If enable task scheduler
    {
        TaskSch_Add = SchTasks /create /TN AHK_Shutdown /ST %g_ShutdownTime% /SC once /TR "shutdown /s /t 60" /F
        Log.Msg("Adding task scheduler(" g_ShutdownTime " shutdown)...Re=" GetCmdOutput(TaskSch_Add))
    }
}

ReindexFiles()                                                          ; Re-create Index section
{
    IniDelete, %g_IniFile%, %SEC_INDEX%
    for dirIndex, dir in StrSplit(g_IndexDir, " | ")
    {
        if (InStr(dir, "A_"))
        {
            searchPath := %dir%
        }
        else
        {
            searchPath := dir
        }

        for extIndex, ext in StrSplit(g_IndexFileType, " | ")
        {
            Loop Files, %searchPath%\%ext%, R
            {
                if (g_IndexExclude != "" && RegExMatch(A_LoopFileLongPath, g_IndexExclude))
                {
                    continue                                            ; Skip this file and move on to the next loop.
                }
                IniWrite, 1, %g_IniFile%, %SEC_INDEX%, File | %A_LoopFileLongPath% ; Assign initial rank to 1
                Progress, %A_Index%, %A_LoopFileName%, ReIndexing..., ReindexFiles
            }
        }
        Progress, Off
    }

    Log.Msg("Indexing search database...")
    TrayTip, %g_WinName%, ReIndex database finish successfully. , 8
    LoadCommands()
}

Help()
{
    Options(Arg, 7)                                                     ; Open Options window 7th tab (help tab)
}

;===========================================================================================
; Library - Listary ( QuickSwitch Function )
; 仿照Listary快速切换文件夹功能, 可以独立成单独的AHK文件
; 在保存/打开对话框中点击菜单项,可以更换对话框到相应路径, 增加自动切换路径功能
; Ctrl+G 将对话框路径切换到TC的路径, Ctrl+E 将对话框路径切换到资源管理器的路径
;===========================================================================================
Listary()
{
    Log.Msg("Starting Listary Function...")

    GroupAdd, FileManager, ahk_class TTOTAL_CMD                         ; QuickSwitch File Manager Class - Total Commander
    GroupAdd, FileManager, ahk_class CabinetWClass                      ; QuickSwitch File Manager Class - Windows Explorer

    GroupAdd, DialogBox, ahk_class %g_DialogClass%                      ; 需要QuickSwith的窗口, 包括打开/保存对话框等

    GroupAdd, ExcludeWin, ahk_exe 7zG.exe                               ; 排除特定窗口,避免被 Auto-QuickSwitch 影响
    GroupAdd, ExcludeWin, ahk_exe Explorer.exe                          ; Folder/File Properties Dialog
    GroupAdd, ExcludeWin, ahk_class SysListView32
    GroupAdd, ExcludeWin, ahk_exe Totalcmd64.exe
    GroupAdd, ExcludeWin, ADAPT Licensing
    GroupAdd, ExcludeWin, AutoCAD LT Alert
    GroupAdd, ExcludeWin, Open - Foreign DWG File
    GroupAdd, ExcludeWin, AutoCAD LT Error Report
    GroupAdd, ExcludeWin, AutoCAD LT
    GroupAdd, ExcludeWin, RAPT
    GroupAdd, ExcludeWin, Progress
    GroupAdd, ExcludeWin, Autodesk Self-Extract

    if(g_AutoSwitchDir)
    {
        Log.Msg("Listary Auto-QuickSwitch Enabled.")
        Loop
        {
            WinWaitActive ahk_class TTOTAL_CMD
                WinGet, ThisHWND, ID, A
            WinWaitNotActive

            If(WinActive("ahk_group DialogBox") && !WinActive("ahk_group ExcludeWin")) ; 检测当前窗口是否符合打开保存对话框条件
            {
                WinGetActiveTitle, Title
                WinGet, ActiveProcess, ProcessName, A

                TrayTip, Listary Function,  Dialog detected`, Active window info `nahk_title = %Title% `nahk_exe = %ActiveProcess%
                Log.Msg("Dialog detected, active window ahk_title = " Title ", ahk_exe = " ActiveProcess )
                ChangePath(GetTC())                                     ; NO Return, as will terimate loop (AutoSwitchDir)
            }
        }
    }
    
    Hotkey, IfWinActive, ahk_group DialogBox                            ; 设置对话框路径定位热键,为了不影响其他程序热键,设置只对打开/保存对话框生效
    Hotkey, ^e, LocateExplorer                                          ; Ctrl+E 把打开/保存对话框的路径定位到资源管理器当前浏览的目录
    Hotkey, ^g, LocateTC                                                ; Ctrl+G 把打开/保存对话框的路径定位到TC当前浏览的目录
    Hotkey, IfWinActive
}

LocateExplorer()
{
    ChangePath(GetExplorer())
}
LocateTC()
{
    ChangePath(GetTC())
}

GetTC()                                                                 ; 获取TC 当前文件夹路径
{
    ClipSaved := ClipboardAll 
    Clipboard =
    SendMessage 1075, 2029, 0, , ahk_class TTOTAL_CMD
    ClipWait, 200
    OutDir=%Clipboard%\                                                 ; 结尾添加\ 符号,变为路径,试图解决AutoCAD不识别路径问题
    Clipboard := ClipSaved 
    ClipSaved = 
    Return OutDir
}

GetExplorer()                                                           ; 获取Explorer路径
{
    Loop,9
    {
        ControlGetText, Dir, ToolbarWindow32%A_Index%, ahk_class CabinetWClass
    } until (InStr(Dir,"Address"))
 
    Dir:=StrReplace(Dir,"Address: ","")
    if (Dir="Computer" )
        Dir:="C:\"

    InitialAdd:=SubStr(Dir,2,2)
    If (InitialAdd != ":\")                                             ; then Explorer lists it as one of the library directories such as Music or Pictures
        Dir:=% "C:\Users\" A_UserName "\" Dir

    Return Dir
}

ChangePath(Dir)
{
    ControlGetText, w_Edit1Text, Edit1, A
    ControlClick, Edit1, A
    ControlSetText, Edit1, %Dir%, A
    ControlSend, Edit1, {Enter}, A
    ;Sleep,100
    ;ControlSetText, Edit1, %w_Edit1Text%, A                            ;还原之前的窗口 File Name 内容, 在选择文件的对话框时没有问题, 但是在选择文件夹的对话框有Bug,所以暂时注释掉
    Log.Msg("Listary Change Path=" Dir)
}

;=========================================================
; 命令管理可视化功能
;=========================================================
CmdMgr(Arg1, Arg2)
{
    global
    Log.Msg("Starting Command Manager... Args=" Arg1 " " Arg2)

    _Path   := Arg1                                                     ; _Path: cmd type eg. file, url, cmd, dir
    _Desc   := Arg2                                                     ; _Desc: cmd name or description, searchable

    if (_Desc = "")                                                     ; Normally this type is from file right click "SendTo"
    {
        SplitPath _Path, _Desc, fileDir, fileExt, nameNoExt, fileDrive  ; Extra name from _Path (if _Type is dir and has "." in path, nameNoExt will not get full folder name) 
        
        if InStr(FileExist(_Path), "D")                                 ; True only if the file exists and is a directory.
        {
            if InStr(_Path, "PROPOSALS & TENDERS")                      ; Check if the path contain "PROPOSALS & TENDERS"
            {
                _Type := 6
                _Desc := ""
            }
            else if InStr(_Path, "DESIGN PROJECTS")                     ; Check if the path contain "DESIGN PROJECTS"
            {
                _Type := 7
                _Desc := ""
            }
            else                                                        ; It is a normal folder
            {
                _Type := 5
            }
        }
        
        if (fileExt = "lnk" && g_SendToGetLnk)
        {
            FileGetShortcut, %_Path%, _Path, fileDir, targetArg, fileDesc

            if (fileDesc = _Path)
            {
                fileDesc := ""
            }

            _Path .= " " targetArg
        }
    }
    else                                                                ; Normally this type is from ALTRun AddCommand
    {
        ;_Desc := _Desc
    }

    _Path := StrReplace(_Path, OneDriveConsumer, "%OneDriveConsumer%")  ; Convert absolute path to relative path, but problem with FileExist()
    _Path := StrReplace(_Path, OneDriveCommercial, "%OneDriveCommercial%") ; Convert absolute path to relative path, but problem with FileExist()

    Gui, CmdMgr:New
    Gui, CmdMgr:Font, s8, Century Gothic, wRegular
    Gui, CmdMgr:+AlwaysOnTop +LastFound +OwnDialogs
    Gui, CmdMgr:Margin, 5, 5
    Gui, CmdMgr:Add, GroupBox, w550 h220, Add Command
    Gui, CmdMgr:Add, Text, xp+20 yp+35, Command Type: 
    Gui, CmdMgr:Add, DropDownList, xp+120 yp-5 w150 v_Type Choose%_Type%, Function|URL|Command|File||Dir|Tender|Project|
    Gui, CmdMgr:Add, Text, xp-120 yp+50, Command Path: 
    Gui, CmdMgr:Add, Edit, xp+120 yp-5 w350 v_Path, %_Path%
    Gui, CmdMgr:Add, Button, xp+355 yp w30 hp gSelectCmdPath, ...
    Gui, CmdMgr:Add, Text, xp-475 yp+100, Description: 
    Gui, CmdMgr:Add, Edit, xp+120 yp-5 w350 v_Desc, %_Desc%
    Gui, CmdMgr:Add, Button, Default x415 w65, OK
    Gui, CmdMgr:Add, Button, xp+75 yp w65, Cancel

    Gui, Main:Hide                                                      ; 隐藏主窗口
    Gui, CmdMgr:Show, AutoSize, %g_CmdMgrWinName%
    Return
}

SelectCmdPath()
{
    Global
    Gui +OwnDialogs                                                     ; Make open dialog Modal
    Gui, CmdMgr:Submit, Nohide                                          ; 保存每个控件的内容到其关联变量中.
    if(_Type = "Dir" )
    {
        FileSelectFolder, _Path, , 3
        if ( _Path != "")
        {
            GuiControl,, _Path, %_Path%
        }
    }
    else
    {
        FileSelectFile, _Path, 3, , Select Path, All File (*.*)
        if (_Path != "")
        {
            GuiControl,, _Path, %_Path%
        }
    }
    SplitPath _Path, fileName, fileDir, fileExt, nameNoExt, drive       ; Extra name from _Path (if _Type is dir and has "." in path, nameNoExt will not get full folder name) 
    GuiControl,, _Desc, %fileName%
}

CmdMgrButtonOK()
{
    global
    Gui +OwnDialogs                                                     ; Make open dialog Modal
    Gui, CmdMgr:Submit, Nohide                                          ; 保存每个控件的内容到其关联变量中, 不隐藏窗口

    if (_Path = "")
    {
        MsgBox, Please input correct command path!
        Return
    }
    else
    {
        IniWrite, 1, %g_IniFile%, %SEC_USERCMD%, %_Type% | %_Path% | %_Desc% ; Assign initial Rank=1
    }

    Gui, CmdMgr:Hide
    MsgBox, ALTRun command manager, ALTRun command added successfully!
    ALTRun_Reload("Silent")
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

;=========================================================
; AppControl 功能代码 (可以独立成单独的AHK文件)
;=========================================================

AppControl()
{
    ;=============== Run Total Commander =================
    /*
    #z::
        DetectHiddenWindows, on
        IfWinNotExist ahk_class TTOTAL_CMD
            Run, "%g_TCPath%"
        Else
            IfWinNotActive ahk_class TTOTAL_CMD
            WinActivate
        Else
            WinMinimize
    Return
    */

    ;================================================
    ; Ctrl+D 自动添加日期 生效的应用程序
    ;================================================
    GroupAdd, FileListMangr, ahk_class TTOTAL_CMD                       ; 针对TC文件列表重命名
    GroupAdd, FileListMangr, ahk_class CabinetWClass                    ; 针对Windows 资源管理器文件列表重命名
    GroupAdd, FileListMangr, ahk_class Progman                          ; 针对Windows 桌面文件重命名
    ;GroupAdd, FileListMangr, ahk_class TCOMBOINPUT                     ; 针对TC F7创建新文件夹对话框（可单独出来用isFile:= True来控制不考虑后缀的影响）
    GroupAdd, FileListMangr, ahk_class TSTDTREEDLG                      ; 针对TC 新建其他格式文件如txt, rtf, docx...
    GroupAdd, FileListMangr, ahk_class #32770                           ; 针对资源管理器文件保存对话框
    
    Hotkey, IfWinActive, ahk_group FileListMangr                        ; 针对所有设定好的程序 按Ctrl+D自动在文件(夹)名之后添加日期
    Hotkey, ^D, RenameWithDate
    Hotkey, IfWinActive, ahk_class TCmtEditForm                         ; 针对TC File Comment对话框 按Ctrl+D自动在备注文字之后添加日期
    Hotkey, ^D, LineEndAddDate
    Hotkey, IfWinActive, ahk_class Notepad2                             ; 针对Notepad2 (原Ctrl+D 为重复当前行)
    Hotkey, ^D, LineEndAddDate
    Hotkey, IfWinActive, ahk_class TCOMBOINPUT                          ; 针对TC F7创建新文件夹对话框（可单独出来用isFile:= True来控制不考虑后缀的影响）
    Hotkey, ^D, LineEndAddDate
    Hotkey, IfWinActive, ahk_exe Evernote.exe                           ; 针对Evernote 按Ctrl+D自动在光标处添加日期
    Hotkey, ^D, EvernoteDate
    
    Hotkey, IfWinActive, ahk_exe RAPTW.exe                              ; 如果正在使用RAPT,鼠标中间激活PT Tools
    Hotkey, ~MButton, RunPTTools
    Hotkey, IfWinActive
}

RunPTTools()                                                            ; 如果正在使用RAPT,鼠标中间激活PT Tools
{
    DetectHiddenWindows, on
    IfWinNotExist, PT Tools
        Run % A_ScriptDir "\PTTools.ahk"
    else IfWinNotActive, PT Tools
        WinActivate
    else
        WinMinimize
}

RenameWithDate()                                                        ; 针对所有设定好的程序 按Ctrl+D自动在文件(夹)名之后添加日期
{
    ControlGetFocus, CurrCtrl, A                                        ; 获取当前激活的窗口中的聚焦的控件名称
    if (InStr(CurrCtrl, "Edit") or InStr(CurrCtrl, "Scintilla"))        ; 如果当前激活的控件为Edit类或者Scintilla1(Notepad2),则Ctrl+D功能生效
        NameAddDate("FileListMangr", CurrCtrl)
    Else
        SendInput ^D
}

LineEndAddDate()                                                        ; 针对TC File Comment对话框　按Ctrl+D自动在备注文字之后添加日期
{
    SendInput {End}
    Sleep, 10
    SendInput {Space}- %A_DD%.%A_MM%.%A_YYYY%
    Log.Msg("AddDateAtEnd, Add= - " A_DD "." A_MM "." A_YYYY)
}

EvernoteDate()                                                          ; 针对Evernote 按Ctrl+D自动在光标处添加日期
{
    SendInput {Space}- %A_DD%.%A_MM%.%A_YYYY%
    Log.Msg("EvernoteDate, Add= - " A_DD "." A_MM "." A_YYYY)
}

NameAddDate(WinName, CurrCtrl, isFile:= True)                           ; 在文件（夹）名编辑框中添加日期,CurrCtrl为当前控件(名称编辑框Edit),isFile是可选参数,默认为真
{
    ControlGetText, EditCtrlText, %CurrCtrl%, A
    SplitPath, % EditCtrlText, fileName, fileDir, fileExt, nameNoExt
    
    if (isFile && fileExt!="" && StrLen(fileExt)<5 && !RegExMatch(fileExt,"^\d+$")) ; 如果是文件,而且有真实文件后缀名,才加日期在后缀名之前, another way is use if fileExt in %TrgExtList% but can not check isFile at the same time
    {
        NameWithDate = %nameNoExt% - %A_DD%.%A_MM%.%A_YYYY%.%fileExt%
    }
    else
    {
        NameWithDate = %EditCtrlText% - %A_DD%.%A_MM%.%A_YYYY%
    }
    ControlClick, %CurrCtrl%, A
    ControlSetText, %CurrCtrl%, %NameWithDate%, A
    Sleep,10
    SendInput {End}
    Log.Msg(WinName ", RenameWithDate=" NameWithDate)
}
;=============================================================
; 使用 CapsLock 切换输入法
; 为了让 Mac/Win 下的体验稍微一致些, 写了一个 ahk 脚本针对 CapsLock 键:
; 单击 切换输入法(本质是调用 win + sapce), 双击/长按 切换 CapsLock 状态
;=============================================================
switchCapsLockState() 
{
    state := GetKeyState("CapsLock", "T")
    nextState := !state
    SetCapsLockState % nextState
    
    return nextState
}

showTip(isOn) 
{
    title := isOn ? "CapsLock: ON" : "CapsLock: OFF"
    text := isOn ? "已打开" : "已关闭"

    TrayTip, %title%, %text%, 1, 16
    return 
}

ToggleAndShowTip()
{
    nextState :=switchCapsLockState()
    showTip(nextState)

    return
}

CapsLock::
    if (g_EnableCapsLockIME)
    {
        KeyWait, CapsLock, T0.3

        if (ErrorLevel) {
            ; long click
            ToggleAndShowTip()
        } else {
            KeyWait, CapsLock, D T0.1

            if (ErrorLevel) {
                ; single click
                SendInput #{Space} 
                ; 切换为中文时经常是中文输入法的英文模式,因为调用Win同时会调用Ctrl,Win10中设置：输入法模式右键-Key Configuration-Chinese/English mode switch-untick Ctrl+Space
                ; 对当前窗口激活下一输入法,会在中文（中文模式）,中文（英文模式）,英文循环切换
                ;DllCall("SendMessage", UInt, WinActive("A"), UInt, 80, UInt, 1, UInt, DllCall("ActivateKeyboardLayout", UInt, 1, UInt, 256))
            } else {                                                    ; double click
                ToggleAndShowTip()
            }
        }

        KeyWait, CapsLock
    }
    else
    {
        switchCapsLockState()
    }
Return

;=======================================================================
; Library - Logger
; Logs := New Logger("Logger.log")
; Logs.Debug(A_LineNumber, "Test Log Message")
;=======================================================================
Class Logger
{
    __New(filename)
    {
        this.filename := filename
        this.enable := True

        if (!g_isLogging)                                               ; If disable logging, clean up logs
        {
            this.enable := False
            FileDelete, %g_LogFile%
            Return
        }
    }

    Msg(Msg)
    {
        if (this.enable)
        {
            FileAppend, % "[" A_Now "] " Msg "`n", % this.filename
        }
    }
}

;============================================================
; Options / Settings Library
;============================================================
Options(Arg := "", ActTab := 1)                                         ; 1st parameter is to avoid menu like [Option `tF2] disturb ActTab
{
    Global                                                              ; Assume-global mode
    Log.Msg("Loading options window...Arg=" Arg ", ActTab=" ActTab)
    
    Gui, Setting:New
    Gui, Setting:Font, s8, Century Gothic, wRegular
    Gui, Setting:+LastFound +OwnDialogs +AlwaysOnTop
    Gui, Setting:Margin, 5, 5

    Gui, Setting:Add, Tab3,xm ym vCurrTab Choose%ActTab% -Wrap, General|Index|GUI|Command|Hotkey|Plugins|Help

    Gui, Setting:Tab, 1                                                 ; Config Tab
    Gui, Setting:Add, GroupBox, w520 h420, General Settings

    Gui, Setting:Add, CheckBox, xp+10 yp+25 vg_AutoStartup checked%g_AutoStartup%, Startup with Windows
    Gui, Setting:Add, CheckBox, xp+250 yp vg_EnableSendTo checked%g_EnableSendTo%, Create SendTo menu
    Gui, Setting:Add, CheckBox, xp-250 yp+30 vg_InStartMenu checked%g_InStartMenu%, Add into Start Menu
    Gui, Setting:Add, CheckBox, xp+250 yp vg_ShowIcon checked%g_ShowIcon%, Show icon in file list

    Gui, Setting:Add, CheckBox, xp-250 yp+30 vg_SaveHistory checked%g_SaveHistory%, Save command history
    Gui, Setting:Add, CheckBox, xp+250 yp vg_KeepInputText checked%g_KeepInputText%, Keep input text when close window
    Gui, Setting:Add, CheckBox, xp-250 yp+30 vg_AutoRank checked%g_AutoRank%, Auto Rank as per frequency
    Gui, Setting:Add, CheckBox, xp+250 yp vg_RunIfOnlyOne checked%g_RunIfOnlyOne%, Run if only one result
    
    Gui, Setting:Add, CheckBox, xp-250 yp+30 vg_SearchFullPath checked%g_SearchFullPath%, Search Full Path
    Gui, Setting:Add, CheckBox, xp+250 yp vg_HideOnDeactivate checked%g_HideOnDeactivate%, Close window when deactivate
    Gui, Setting:Add, CheckBox, xp-250 yp+30 vg_AlwaysOnTop checked%g_AlwaysOnTop%, Window Always-On-Top
    Gui, Setting:Add, CheckBox, xp+250 yp vg_EnableCapsLockIME checked%g_EnableCapsLockIME%, Enable CapsLock Switch IME

    Gui, Setting:Add, CheckBox, xp-250 yp+30 vg_SwitchToEngIME checked%g_SwitchToEngIME%, Switch To Eng IME
    Gui, Setting:Add, CheckBox, xp+250 yp vg_isLogging checked%g_isLogging%, Enable Logging
    Gui, Setting:Add, CheckBox, xp-250 yp+30 vg_EscClearInput checked%g_EscClearInput%, Esc clear input then close window
    Gui, Setting:Add, CheckBox, xp+250 yp, #SendTo Menu Simple Mode

    Gui, Setting:Add, CheckBox, xp-250 yp+30 vg_SendToGetLnk checked%g_SendToGetLnk%, SendTo Menu Retrieves Lnk Target
    Gui, Setting:Add, CheckBox, xp+250 yp vg_ShowFileExt checked%g_ShowFileExt%, Show file extension

    Gui, Setting:Add, Text, xp-250 yp+40, Text Editor (eg. Notepad2): 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w330 vg_Editor, %g_Editor%

    Gui, Setting:Add, Text, xp-150 yp+40, Everything.exe Path: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w330 vg_Everything, %g_Everything%

    Gui, Setting:Add, Text, xp-150 yp+40, Total Commander Path: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w330 vg_TCPath, %g_TCPath%
    
    Gui, Setting:Tab, 2                                                 ; Index Tab
    Gui, Setting:Add, GroupBox, w520 h420, Index Options

    Gui, Setting:Add, Text, xp+10 yp+40, Index Locations: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w330 vg_IndexDir, %g_IndexDir%
    Gui, Setting:Add, Text, xp-150 yp+40, Index File Type: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w330 vg_IndexFileType, %g_IndexFileType%
    Gui, Setting:Add, Text, xp-150 yp+40, Index File Exclude: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w330 vg_IndexExclude, %g_IndexExclude%

    Gui, Setting:Add, Text, xp-150 yp+40, HistorySize: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w330 vg_HistorySize, %g_HistorySize%

    Gui, Setting:Tab, 3                                                 ; GUI Tab
    Gui, Setting:Add, GroupBox, w520 h420, GUI Details

    Gui, Setting:Add, CheckBox, xp+10 yp+25 vg_HideTitle checked%g_HideTitle%, Hide Title Bar
    Gui, Setting:Add, CheckBox, xp+250 yp vg_ShowTrayIcon checked%g_ShowTrayIcon%, Show Tray Icon
    Gui, Setting:Add, CheckBox, xp-250 yp+30, #Show status bar
    Gui, Setting:Add, CheckBox, xp+250 yp vg_HideCol2 checked%g_HideCol2%, Hide 2nd Column
    Gui, Setting:Add, CheckBox, xp-250 yp+30 vg_LVGrid checked%g_LVGrid%, Show Grid in Command ListView
    Gui, Setting:Add, CheckBox, xp+250 yp, #
    Gui, Setting:Add, Text, xp-250 yp+40 , Display Rows (1-9): 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80 vg_DisplayRows, %g_DisplayRows%
    Gui, Setting:Add, Text, xp+100 yp+5, 3rd column width: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80 vg_Col3Width, %g_Col3Width%
    Gui, Setting:Add, Text, xp-400 yp+40, 4th column width: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80 vg_Col4Width, %g_Col4Width%
    Gui, Setting:Add, Text, xp+100 yp+5, Font Name: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80 vg_FontName, %g_FontName%
    Gui, Setting:Add, Text, xp-400 yp+40, Font Size: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80 vg_FontSize, %g_FontSize%
    Gui, Setting:Add, Text, xp+100 yp+5, Font Color: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80 vg_FontColor, %g_FontColor%
    Gui, Setting:Add, Text, xp-400 yp+40, Window Width: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80 vg_WinWidth, %g_WinWidth%
    Gui, Setting:Add, Text, xp+100 yp+5, Edit Height: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80 vg_EditHeight, %g_EditHeight%
    Gui, Setting:Add, Text, xp-400 yp+40, Command List Height: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80 vg_ListHeight, %g_ListHeight%
    Gui, Setting:Add, Text, xp+100 yp+5, #Transparency (100-255):
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80,
    Gui, Setting:Add, Text, xp-400 yp+40, Controls' Color:
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80 vg_CtrlColor, %g_CtrlColor%
    Gui, Setting:Add, Text, xp+100 yp+5, Window's Color:
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80 vg_WinColor, %g_WinColor%
    Gui, Setting:Add, Text, xp-400 yp+40, Background Picture: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80 vg_BackgroundPicture, %g_BackgroundPicture%
    Gui, Setting:Add, Text, xp+100 yp+5, #Border Size: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80,

    Gui, Setting:Tab, 5                                                 ; Hotkey Tab

    Gui, Setting:Add, GroupBox, w520 h85, Active ALTRun Global Hotkey:
    Gui, Setting:Add, Text, xp+10 yp+25 , Global Hotkey 1:
    Gui, Setting:Add, CheckBox, xp+150 yp w100 vg_GlobalHotkey1Win checked%g_GlobalHotkey1Win%, Win         +
    Gui, Setting:Add, Hotkey, xp+100 yp-4 w230 vg_GlobalHotkey1, %g_GlobalHotkey1%
    Gui, Setting:Add, Text, xp-250 yp+35 , Global Hotkey 2:
    Gui, Setting:Add, CheckBox, xp+150 yp w100 vg_GlobalHotkey2Win checked%g_GlobalHotkey2Win%, Win         +
    Gui, Setting:Add, Hotkey, xp+100 yp-4 w230 vg_GlobalHotkey2, %g_GlobalHotkey2%
    Gui, Setting:Add, GroupBox, xp-260 yp+40 w520 h55, Run Command Hotkey:
    Gui, Setting:Add, Text, xp+10 yp+25 , Run Command Hotkey:
    Gui, Setting:Add, CheckBox, xp+150 yp w50 vg_RunCmdAlt checked%g_RunCmdAlt%, ALT + 
    Gui, Setting:Add, Text, xp+50 yp, No.
    Gui, Setting:Add, Text, xp+40 yp, Select Command Hotkey: 
    Gui, Setting:Add, Text, xp+150 yp, Ctrl + No.
    Gui, Setting:Add, GroupBox, xp-400 yp+40 w520 h120, Action Hotkey:
    Gui, Setting:Add, Text, xp+10 yp+20 , Hotkey 1: 
    Gui, Setting:Add, Hotkey, xp+150 yp-5 w80 vg_Hotkey1, %g_Hotkey1%
    Gui, Setting:Add, Text, xp+100 yp+5, Toggle Action: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80 vg_Trigger1, %g_Trigger1%
    Gui, Setting:Add, Text, xp-400 yp+40 , Hotkey 2: 
    Gui, Setting:Add, Hotkey, xp+150 yp-5 w80 vg_Hotkey2, %g_Hotkey2%
    Gui, Setting:Add, Text, xp+100 yp+5, Toggle Action: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80 vg_Trigger2, %g_Trigger2%
    Gui, Setting:Add, Text, xp-400 yp+40 , Hotkey 3: 
    Gui, Setting:Add, Hotkey, xp+150 yp-5 w80 vg_Hotkey3, %g_Hotkey3%
    Gui, Setting:Add, Text, xp+100 yp+5, Toggle Action: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80 vg_Trigger3, %g_Trigger3%

    Gui, Setting:Tab, 6                                                 ; Plugins / Listary / Scheduler Tab
    Gui, Setting:Add, GroupBox, w520 h170, Listary Function
    Gui, Setting:Add, Text, xp+10 yp+35 , File Manager Class: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w330 vg_FileMgrClass, %g_FileMgrClass%
    Gui, Setting:Add, Text, xp-150 yp+40, Open/Save Dialog Class: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w330 vg_DialogClass, %g_DialogClass%
    Gui, Setting:Add, Text, xp-150 yp+40, Exclude Windows Class: 
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w330 vg_ExcludeWinClass, %g_ExcludeWinClass%
    Gui, Setting:Add, CheckBox, xp-150 yp+40 vg_AutoSwitchDir checked%g_AutoSwitchDir%, Auto Switch Dir
    Gui, Setting:Add, GroupBox, xp-10 yp+35 w520 h250, Scheduler
    Gui, Setting:Add, CheckBox, xp+10 yp+20 vg_EnableScheduler checked%g_EnableScheduler%, Shutdown Scheduler
    Gui, Setting:Add, Text, xp+250 yp, Shutdown Time:
    Gui, Setting:Add, Edit, xp+150 yp-5 r1 w80 vg_ShutdownTime, %g_ShutdownTime%

    Gui, Setting:Tab, 7                                                 ; Help Tab
    Gui, Setting:Add, GroupBox, w520 h420, Help Information:

    AllFunctions := GetAllFunctions()

    Gui, Setting:Add, Edit, xp+10 yp+20 w500 h380 ReadOnly -WantReturn vHelpText,
    (Ltrim
    # ALTRun
    ALTRun - An effective launcher for Windows, open source project
    Similar to RunZ(AutoHotkey) and ALTRun(Pascal).
    
    1. Pure portable software, will not write anything into Registry.
    2. Small size (< 2MB) and low resources usage (< 10MB RAM). 
    3. Highly customizable with GUI (Main Window, Options, Command Manager)
    2. Listary Quick Switch Dir function added
    3. AppControl func added, eg. Press middle mouse button to open PT Tools

    # GUI
    ![GUI](https://github.com/zhugecaomao/ALTRun/releases)

    # [Download](https://github.com/zhugecaomao/ALTRun/releases)

    ------------------------------------------------------------------------
    Congraduations! You have run shortcut %g_RunCount% times by now! 
    
    shortcut list:-

    F1        		打开帮助页面
    F2        		打开配置选项
    F3        		编辑配置文件
    F4        		用户自定义索引列表
    F5        		编辑当前命令
    Shift + F1		显示置顶的按键提示
    Alt + F4		退出置顶的按键提示
    Enter   		执行当前命令
    Esc     		关闭窗口
    Alt +    		加每列行首字符执行
    Tab +    		再按每列行首字符执行
    Tab +    		再按 Shift + 行首字符 定位
    Win + J  		显示或隐藏窗口
    Ctrl + F		在输出结果中翻到下一页
    Ctrl + B		在输出结果中翻到上一页
    Ctrl + H		显示历史记录
    Ctrl + +		可增加当前功能的权重
    Ctrl + -		可减少当前功能的权重
    Ctrl + L		清除编辑框内容
    Ctrl + R		重新创建待搜索文件列表
    Ctrl + Q		重启
    Ctrl + D		用默认文件管理器打开当前命令所在目录
    Ctrl + S		显示并复制当前文件的完整路径
    Shift + Del		删除当前文件
    Ctrl + Del		删除当前命令
    Ctrl + I		移动光标当行首
    Ctrl + O		移动光标当行尾
    Input Web   	可直接输入 www 或 http 开头的网址
    ;        		以分号开头命令,用 ahk 运行
    :        		以冒号开头的命令,用 cmd 运行
    No Result		搜索无结果,回车用 ahk 运行
    Space       	输入空格后,搜索内容锁定
    ------------------------------------------------------------------------
    All other functions:-

    %AllFunctions%
    )

    Gui, Setting:Tab                                                    ; 后续添加的控件将不属于前面那个选项卡控件
    Gui, Setting:Add, Button, Default x320 w65, OK
    Gui, Setting:Add, Button, xp+75 yp w65, Cancel
    Gui, Setting:Add, Button, xp+75 yp w65, Help

    Gui, Main:Hide                                                      ;隐藏主窗口
    Gui, Setting:Show, AutoSize, %g_OptionsWinName%
}

;=============== 设置选项窗口 - 按钮动作 =================
GetAllFunctions()
{
    result := ""
    for index, element in g_Commands
    {
        if (InStr(element, "function | ") && !InStr(result, element "`n"))
        {
            result .= element "`n"
        }
    }
    result := StrReplace(result, "function | ", "")
    Return result
}

SettingButtonOK()
{
    SaveConfig("main")
    ALTRun_Reload("Silent")
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
    RemoveToolTip()
    Gui, Setting:Destroy
}

SettingButtonHelp()
{
    GuiControl , Setting:Choose, CurrTab, 7
}

;=============================================================
; 加载主配置文件
;=============================================================
LoadConfig(Arg)
{
    Log.Msg("Loading config...Arg=" Arg)
    
    if (Arg = "config" or Arg = "initialize" or Arg = "all")
    {
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
        
        if (FileExist(g_BackgroundPicture))
        {
            g_BGPicture := g_BackgroundPicture
        }
        else
        {
            Extract_BG(A_Temp "\ALTRunBG.jpg")
            g_BGPicture := A_Temp "\ALTRunBG.jpg"
        }
    }

    if (Arg = "CMD_SECS" or Arg = "initialize" or Arg = "all")          ; Built-in command initialize
    {
        IniRead, DFTCMDSEC, %g_IniFile%, %SEC_DFTCMD%
        if (DFTCMDSEC = "")
        {
            IniWrite, 
            (Ltrim
            ;
            ; Build-in Commands (High Priority, DO NOT Edit)
            ; Command type: File, Dir, Command, Function, URL, Project, Tender
            ; Type | Command | Comments=Rank
            ; 
            ; File | notepad.exe | Notepad (File type will run with AHK's Run)=1
            ; URL | www.google.com | Google=1
            ; CMD | ipconfig | Check IP (CMD type will run with cmd.exe, auto pause after run)=1
            ;
            Function | Help | ALTRun Help Index (F1)=100
            Function | ALTRun_Log | ALTRun Log File=100
            Function | ALTRun_Reload | ALTRun Reload=100
            Function | ShowCmdHistory | History Commands=100
            Function | AddCommand | Add new command=100
            Function | UserCommandList | ALTRun User-defined command (F4)=100
            Function | ReindexFiles | Reindex search database=100
            Function | Everything | Search by Everything=100
            Function | Options | ALTRun Options Preference Settings (F2)=100
            Function | RunPTTools | PT Tools (AHK)=100
            Dir | A_ScriptDir | ALTRun Program Dir=100
            Dir | A_Startup | Current User Startup Dir=100
            Dir | A_StartupCommon | All User Startup Dir=100
            Dir | A_ProgramsCommon | Windowns Search.Index.Cortana Dir=100
            Function | AhkRun | Run Command use AutoHotkey Run=100
            Function | CmdRun | Run Command use CMD=100
            Function | RunAndDisplay | Run by CMD and display the result=100
            Function | RunClipboard | Run Clipboard content with AHK's Run=100
            Function | ShowArg | Show Arguments=100
            Function | SearchOnGoogle | Search Clipboard or Input by Google=100
            Function | SearchOnBing | Search Clipboard or Input by Bing=100
            Function | ShowIp | Show IP Address=100
            Function | Clip | Show clipboard content=100
            Function | EmptyRecycle | Empty Recycle Bin=100
            Function | TurnMonitorOff | Turn off Monitor, Close Monitor=100
            Function | MuteVolume | Mute Volume=100
            ;
            ; Control Panel Commands
            ;
            Control | Control | Control Panel=66
            Control | wf.msc | Windows Defender Firewall with Advanced Security=66
            Control | Control intl.cpl | Region and Language Options=66
            Control | Control firewall.cpl | Windows Defender Firewall=66
            Control | Control access.cpl | Ease of Access Centre=66
            Control | Control appwiz.cpl | Programs and Features=66
            Control | Control sticpl.cpl | Scanners and Cameras=66
            Control | Control sysdm.cpl | System Properties=66
            Control | Control joy.cpl | Game Controllers=66
            Control | Control Mouse | Mouse Properties=66
            Control | Control desk.cpl | Display=66
            Control | Control mmsys.cpl | Sound=66
            Control | Control ncpa.cpl | Network Connections=66
            Control | Control powercfg.cpl | Power Options=66
            Control | Control timedate.cpl | Date and Time=66
            Control | Control admintools | Windows Tools=66
            Control | Control desktop | Personalisation=66
            Control | Control folders | File Explorer Options=66
            Control | Control fonts | Fonts=66
            Control | Control inetcpl.cpl,,4 | Internet Properties=66
            Control | Control printers | Devices and Printers=66
            Control | Control userpasswords | User Accounts=66
            Control | taskschd.msc | Task Scheduler=66
            Control | devmgmt.msc | Device Manager=66
            Control | eventvwr.msc | Event Viewer=66
            Control | compmgmt.msc | Computer Manager=66
            Control | taskmgr.exe | Task Manager=66
            Control | calc.exe | Calculator=66
            Control | mspaint.exe | Paint=66
            Control | cmd.exe | DOS / CMD=66
            Control | regedit.exe | Registry Editor=66
            Control | write.exe | Write=66
            Control | cleanmgr.exe | Disk Space Clean-up Manager=66
            Control | gpedit.msc | Group Policy=66
            Control | comexp.msc | Component Services=66
            Control | diskmgmt.msc | Disk Management=66
            Control | dxdiag.exe | Directx Diagnostic Tool=66
            Control | lusrmgr.msc | Local Users and Groups=66
            Control | msconfig.exe | System Configuration=66
            Control | perfmon.exe /Res | Resources Monitor=66
            Control | perfmon.exe | Performance Monitor=66
            Control | winver.exe | About Windows=66
            Control | services.msc | Services=66
            Control | netplwiz | User Accounts=66
            ), %g_IniFile%, %SEC_DFTCMD%
            IniRead, DFTCMDSEC, %g_IniFile%, %SEC_DFTCMD%
        }

        IniRead, USERCMDSEC, %g_IniFile%, %SEC_USERCMD%
        if (USERCMDSEC = "")
        {
            IniWrite, 
            (Ltrim
            ;
            ; User-Defined Commands (High priority, edit command as desired)
            ; Command type: File, Dir, CMD, Function, URL, Project, Tender
            ; Type | Command | Comments=Rank
            ; 
            ; File | notepad.exe | Notepad (File type will run with AHK's Run)=1
            ; URL | www.google.com | Google=1
            ; CMD | ipconfig | Check IP (CMD type will run with cmd.exe, auto pause after run)=1
            ;
            Dir | `%AppData`%\Microsoft\Windows\SendTo | Windows SendTo Dir=100
            Dir | `%OneDriveConsumer`% | OneDrive Personal Dir=100
            Dir | `%OneDriveCommercial`% | OneDrive Business Dir=100
            File | C:\OneDrive\Apps\TotalCMD64\Tools\Notepad2.exe=100
            CMD | ipconfig | Show IP Address=100
            URL | www.google.com | Google=100
            Project | Q:\DESIGN PROJECTS | Design Folder=100
            Tender | Q:\PROPOSALS & TENDERS | Tender Folder=100
            ), %g_IniFile%, %SEC_USERCMD%
            IniRead, USERCMDSEC, %g_IniFile%, %SEC_USERCMD%
        }

        IniRead, INDEXSEC, %g_IniFile%, %SEC_INDEX%                     ; Read whole section SEC_INDEX (Index database)
        if (INDEXSEC = "")
        {
            MsgBox, 4096, %g_WinName%, ALTRun initialize for first time running.`n`nAuto initialize in 30 seconds or click OK now., 30
            ReindexFiles()
        }
        Return DFTCMDSEC "`n" USERCMDSEC "`n" INDEXSEC
    }

    if (Arg = "FBCMD")
    {
        IniRead, FALLBACKCMDSEC, %g_IniFile%, %SEC_FALLBACK%
        if (FALLBACKCMDSEC = "")
        {
            IniWrite, 
            (Ltrim
            ;======================================================================
            ; Fallback Commands show when search result is empty
            ; Commands in order, Press Enter to run first command
            ;
            Function | AddCommand | Add new command
            Function | Everything | Search by Everything
            Function | SearchOnGoogle | Search Clipboard or Input by Google
            Function | AhkRun | Run Command use AutoHotkey Run
            Function | CmdRun | Run Command use CMD
            Function | RunAndDisplay | Run by CMD and display the result
            Function | SearchOnBing | Search Clipboard or Input by Bing
            ), %g_IniFile%, %SEC_FALLBACK%
            IniRead, FALLBACKCMDSEC, %g_IniFile%, %SEC_FALLBACK%
        }
        Return %FALLBACKCMDSEC%
    }
    Return
}

;=============================================================
; 保存主配置文件
;=============================================================
SaveConfig(Arg)
{
    Log.Msg("Saving config...Arg=" Arg)
    if (Arg = "History")                                                ; 记录窗口激活次数RunCount,History,考虑效率,不能写入太频繁也不能写入太少
    {
        IniWrite, %g_RunCount%, %g_IniFile%, %SEC_CONFIG%, RunCount

        if (g_SaveHistory)
        {
            for index, element in g_HistoryCommands
            {
                IniWrite, %element%, %g_IniFile%, %SEC_HISTORY%, %index%
            }
        }
    }
    else if (Arg = "Main")
    {
        Gui, Setting:Submit                                             ; Submit and Hide, avoid delay feeling

        Loop Parse, KEYLIST_CONFIG, `,                                  ; Save Config Section settings
        {
            IniWrite, % g_%A_LoopField%, %g_IniFile%, %SEC_CONFIG%, %A_LoopField%
        }
        
        Loop Parse, KEYLIST_GUI, `,
        {
            IniWrite, % g_%A_LoopField%, %g_IniFile%, %SEC_GUI%, %A_LoopField%
        }
        
        Loop Parse, KEYLIST_HOTKEY, `,
        {
            IniWrite, % g_%A_LoopField%, %g_IniFile%, %SEC_HOTKEY%, %A_LoopField%
        }
    }
    Return
}

;=======================================================================
; Libraries
;=======================================================================

SwitchIME(dwLayout)
{
    HKL := DllCall("LoadKeyboardLayout", Str, dwLayout, UInt, 1)
    ControlGetFocus, ctl, A
    SendMessage, 0x50, 0, HKL, %ctl%, A
}

SwitchToEngIME()
{
    ; 下方代码可只保留一个
    SwitchIME(0x04090409) ; 英语(美国) 美式键盘
    SwitchIME(0x08040804) ; 中文(中国) 简体中文-美式键盘 / 中文(简体,新加坡)
}

; 0：英文 1：中文
GetInputState(WinTitle = "A")
{
    ControlGet, hwnd, HWND, , , %WinTitle%
    if (A_Cursor = "IBeam")
        return 1
    if (WinActive(WinTitle))
    {
        ptrSize := !A_PtrSize ? 4 : A_PtrSize
        VarSetCapacity(stGTI, cbSize := 4 + 4 + (PtrSize * 6) + 16, 0)
        NumPut(cbSize, stGTI, 0, "UInt")   ;   DWORD   cbSize;
        hwnd := DllCall("GetGUIThreadInfo", Uint, 0, Uint, &stGTI)
                        ? NumGet(stGTI, 8 + PtrSize, "UInt") : hwnd
    }
    return DllCall("SendMessage"
        , UInt, DllCall("imm32\ImmGetDefaultIMEWnd", Uint, hwnd)
        , UInt, 0x0283  ;Message : WM_IME_CONTROL
        , Int, 0x0005  ;wParam  : IMC_GETOPENSTATUS
        , Int, 0)      ;lParam  : 0
}

;=============================================================
; Resources File - Background picture
;=============================================================
Extract_BG(_Filename)
{
	Static Out_Data
    VarSetCapacity(TD, 13026)
	TD := "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"

    VarSetCapacity(Out_Data, Bytes := 4754, 0)
    DllCall("Crypt32.dll\CryptStringToBinary" "W", "Ptr", &TD, "UInt", 0, "UInt", 1, "Ptr", &Out_Data, "UIntP", Bytes, "Int", 0, "Int", 0, "CDECL Int")
	
	IfExist, %_Filename%
		FileDelete, %_Filename%
	
	h := DllCall("CreateFile", "Ptr", &_Filename, "Uint", 0x40000000, "Uint", 0, "UInt", 0, "UInt", 4, "Uint", 0, "UInt", 0)
	, DllCall("WriteFile", "Ptr", h, "Ptr", &Out_Data, "UInt", 4754, "UInt", 0, "UInt", 0)
	, DllCall("CloseHandle", "Ptr", h)
}

;=============================================================
; Some Built-in Functions
;=============================================================
CmdRun()
{
    global
    RunWithCmd(Arg)
}

AhkRun()
{
    global
    Run, %Arg%
}

ShowArg()
{
    args := StrSplit(Arg, " ")
    result := "Function | Total " args.Length() " Args`n"

    for index, argument in args
    {
        result .= "Function | No. " index " argument is: | " argument "`n"
    }

    ListResult(result, true, false)
}

RunClipboard()
{
    global
    Run, %clipboard%
}

RunAndDisplay()
{
    ListResult(GetCmdOutput(Arg), true, false)
}

Clip()
{
    ListResult("Display | Clipboard content length is " StrLen(clipboard) ", content is : | " clipboard, true, false)
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

ShowIp()
{    
    ListResult("IP 1 | " A_IPAddress1 
            . "`r`n IP 2 | " . A_IPAddress2
            . "`r`n IP 3 | " . A_IPAddress3
            . "`r`n IP 4 | " . A_IPAddress4, true, true)
}

SearchOnGoogle()
{
    global
    word := Arg == "" ? clipboard : Arg
    Run, https://www.google.com/search?q=%word%&newwindow=1
}

SearchOnBing()
{
    global
    word := Arg == "" ? clipboard : Arg
    Run, http://cn.bing.com/search?q=%word%
}

Everything()
{
    Run, %g_Everything% -s "%Arg%",, UseErrorLevel
    if ErrorLevel
        MsgBox, % "Everything software not found.`n`nPlease check ALTRun setting and Everything program file."
}

;=======================================================================
; Library - Eval
;=======================================================================
Eval(Expr, Format := FALSE)
{
    static obj := ""
    if ( !obj )
        obj := ComObjCreate("HTMLfile")

    Expr := StrReplace( RegExReplace(Expr, "\s") , ",", ".")
  , Expr := RegExReplace(StrReplace(Expr, "**", "^"), "(\w+(\.*\d+)?)\^(\w+(\.*\d+)?)", "pow($1,$3)")    ; 2**3 -> 2^3 -> pow(2,3)
  , Expr := RegExReplace(Expr, "=+", "==")    ; = -> ==  |  === -> ==  |  ==== -> ==  |  ..
  , Expr := RegExReplace(Expr, "\b(E|LN2|LN10|LOG2E|LOG10E|PI|SQRT1_2|SQRT2)\b", "Math.$1")
  , Expr := RegExReplace(Expr, "\b(abs|acos|asin|atan|atan2|ceil|cos|exp|floor|log|max|min|pow|random|round|sin|sqrt|tan)\b\(", "Math.$1(")

  , obj.write("<body><script>document.body.innerText=eval('" . Expr . "').toFixed(2);</script>")
  , Expr := obj.body.innerText
    obj := ""
    return InStr(Expr, "d") ? "" : InStr(Expr, "false") ? FALSE    ; d = body | undefined
                                 : InStr(Expr, "true")  ? TRUE
                                 : ( Format && InStr(Expr, "e") ? Format("{:f}",Expr) : Expr )
}
