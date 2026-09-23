;===============================================================================
; MainWindow.ahk - 主窗口: 输入框 + 结果列表 + 状态栏, 托盘/右键菜单, 热键 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 用法 (启动流程里按这个顺序调用一次, 见 ALTRun.ahk 的 autorun 段):
;   MainWindow.CreateTrayMenu()    托盘图标和菜单 (必须在 Create() 之前, 窗口
;                                  创建时会沿用当时生效的托盘图标)
;   MainWindow.Create()            创建并显示主窗口, 处理命令行参数 (-Startup/
;                                  -StartMenu/-SendTo)
;   MainWindow.RegisterHotkeys()   全局热键 + 主窗口内热键 + 自定义/条件热键
;
; 其它模块常用的接口:
;   MainWindow.Show() / Hide() / Toggle()     显示 / 隐藏 / 切换主窗口
;   MainWindow.ShowResults(rows, useDisplay)  把一组命令行填进结果列表
;   MainWindow.SetStatus(text)                设置状态栏文字, "TIP" = 随机提示
;   MainWindow.Gui / Input / ListView         窗口和控件对象 (原 Global MainGUI/
;                                             myInputBox/myListView)
;
; 注意: ToggleWindow()/TabFunc()/PrevCommand()/NextCommand()/CopyCommand()/
; ClearInput() 仍以裸的全局函数外壳留在 ALTRun.ahk 里 - 它们登记在
; OptionsWindow 的 FuncList 数组里 (可绑定自定义热键, 函数名以字符串形式保存
; 在 ALTRun.json 里), RunCommand() 靠 %cmdPath%() 按名字动态调用, 只认裸的
; 全局函数名, 不认 Class.Method。只在本类内部用到的事件回调 (Gui/控件事件、
; 菜单项、Alt/Ctrl+数字热键) 都是 _OnXxx 私有方法, 用 (p*) => 包一层再绑定。
;===============================================================================

Class MainWindow {
    static Gui       := ""    ; 主窗口 Gui 对象 (原 Global MainGUI)
    static Input     := ""    ; 搜索输入框 (原 Global myInputBox)
    static ListView  := ""    ; 结果列表 (原 Global myListView)
    static Status    := ""    ; 底部状态栏 - 其实是一个只读 Edit 控件 (原 myStatus)
    static ImageList := ""    ; 结果列表的图标 ImageList (原 Global myImageList)
    static IconMap   := ""    ; 图标缓存: 类型/扩展名/路径 -> ImageList 序号 (原 Global myIconMap)

    ;---------------------------------------------------------------------------
    ; Window creation
    ;---------------------------------------------------------------------------
    static Create() {
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

        MainWindow._CreateImageList()

        g := MainWindow.Gui := Gui(TopMost Caption Theme " +MinSize300x160", g_TITLE)
        g.OnEvent("Close"      , (p*) => MainWindow.Hide())
        g.OnEvent("Escape"     , (p*) => MainWindow._OnEscape())
        g.OnEvent("ContextMenu", (p*) => MainWindow._OnContextMenu(p*))
        g.OnEvent("DropFiles"  , (p*) => MainWindow._OnDropFiles(p*))
        g.BackColor := g_GUI["MainGUIColor"]
        mainGuiFont := Fonts.Spec(g_GUI["MainGUIFont"], "Microsoft YaHei", "norm s10.0")
        g.SetFont(mainGuiFont.opt, mainGuiFont.name)

        inputBox := MainWindow.Input := g.AddEdit("x12 y10 r1 -WantReturn border -E0x200 W" Input_W)
        Win.SetCueBanner(inputBox.Hwnd, g_LNG[13])                          ; "Type anything here to search..." as a real placeholder, not literal text
        inputBox.Opt("Background" g_GUI["CMDListColor"])
        inputBox.OnEvent("Change", (p*) => MainWindow._OnInputChange())

        runBtn := g.AddButton("x+" Run_X " yp W" Run_W " hp Default Hidden" Run_H, g_LNG[11])
        runBtn.OnEvent("Click", RunCurrentCommand)
        optBtn := g.AddButton("x+" Opt_X " yp W" Opt_W " hp Hidden" Opt_H, g_LNG[12])
        optBtn.OnEvent("Click", (*) => Options())

        lv := MainWindow.ListView := g.AddListView("x12 yp+36 W" List_W " H" List_H " -Multi", g_LNG[10])
        lv.Opt(DBuffer Header Grid Border " Background" g_GUI["CMDListColor"] " +Report") ; ListView View Modes: Report, Icon, Tile, IconSmall, List
        lv.OnEvent("Click"      , (p*) => MainWindow._OnListClick(p*))
        lv.OnEvent("ContextMenu", (p*) => MainWindow._OnListContextMenu(p*))
        lv.OnEvent("DoubleClick", (p*) => MainWindow.RunFocusedRow())
        if (g_CONFIG["LargeIcons"]) {
            lv.SetImageList(MainWindow.ImageList, 1)                        ; Attach the ImageList to the ListView, 2nd param is 1: large icons, 0: small icons, 2: state icons (AHK Doc incorrect)
        } else {
            lv.SetImageList(MainWindow.ImageList)
        }

        colWidths := StrSplit(g_GUI["ColWidth"], ",")
        Loop 4 {
            if (colWidths.Length >= A_Index) {
                if (colWidths[A_Index] != "")
                    lv.ModifyCol(A_Index, colWidths[A_Index])
            }
        }

        status := MainWindow.Status := g.AddEdit("x12 y+10 r1 -WantReturn ReadOnly -E0x200 border W" List_W " Hidden" (!g_CONFIG["ShowStatusBar"]), )
        status.Opt("Background" g_GUI["CMDListColor"])
        mainSbFont := Fonts.Spec(g_GUI["MainSBFont"], "Microsoft YaHei", "norm s9.0")
        status.SetFont(mainSbFont.opt, mainSbFont.name)

        if FileExist(Path.Resolve(g_GUI["Background"])) {
            try g.AddPic("x0 y0 0x4000000", Path.Resolve(g_GUI["Background"]))
        } else if (g_GUI["Background"] = "Default") {
            try g.AddPic("x0 y0 0x4000000", MainWindow._DefaultBackgroundFile())
        }

        MainWindow.ShowResults(g_LNG[50])

        hideWin := MainWindow._HandleCommandLineArgs() ? "Hide " : ""

        if (g_GUI["Transparency"] < 250) {
            WinSetTransparent(g_GUI["Transparency"], g.Hwnd)                ; By default, hidden windows are not detected. however, when using pure HWNDs, hidden windows are always detected regardless of DetectHiddenWindows.
        }

        Win.SetCorner(g.Hwnd, g_CONFIG["RoundCorner"])

        g.Show(hideWin "w" g_GUI["WinX"] " h" g_GUI["WinY"] " Center")

        ; Enable dragging for captionless window
        OnMessage(0x201, (p*) => MainWindow._OnLButtonDown(p*))

        if g_CONFIG["HideOnLostFocus"] {
            ; 方案 1 - OnMessage(0x0006, WM_ACTIVATE)
            ; 事件驱动, 高效无延迟, 无资源占用
            ; 某些情况有窗口闪烁和托盘菜单右键点击"显示"窗口闪退问题
            ; 方案 2 - SetTimer(_OnLoseFocus, 30)
            ; Ahk原生方式, 代码简单可靠, 但有轻微性能开销, 有稍许延迟 30ms
            ; 方案 3 - Control.OnEvent("LoseFocus", _OnLoseFocus)   <- 当前采用
            ; Ahk原生方式, 事件驱动, 高效无延迟, 无资源占用
            ; 因 Gui 本身没有 LoseFocus 事件, 需要注册主界面所有控件的 LoseFocus 事件
            ; 修复托盘菜单右键点击"显示"窗口闪退问题
            for _, ctrl in [inputBox, lv, runBtn, optBtn, status]
                ctrl.OnEvent("LoseFocus", (*) => MainWindow._OnLoseFocus())
        }
    }

    ; Icon cache: the first six entries are the fixed per-type icons, file icons get appended on demand by _LoadIcon().
    static _CreateImageList() {
        il := MainWindow.ImageList := IL_Create(10, 5, g_CONFIG["LargeIcons"]) ; 3rd param is 1: large icons, 0: small icons
        MainWindow.IconMap := Map("DIR" , IL_Add(il, "imageres.dll", -3)   ; IconIndex=1/2/3/4 for type dir/func/url/eval
                                 ,"FUNC", IL_Add(il, "imageres.dll", -100)
                                 ,"URL" , IL_Add(il, "imageres.dll", -144)
                                 ,"EVAL", IL_Add(il, "imageres.dll", -182)
                                 ,"CMD" , IL_Add(il, "imageres.dll", -100)
                                 ,"CLIP", IL_Add(il, "imageres.dll", -102)) ; "imageres.dll",-5323 is cmd.exe icon
    }

    ; Resolves A_Args[1] / A_Args[2]. Returns true when the window should start hidden.
    static _HandleCommandLineArgs() {
        if (A_Args.Length < 1)
            return false

        for value in A_Args
            g_LOG.Debug("Resolving command line args" A_Index " = " value)

        firstArg := A_Args[1]
        if (firstArg = "-Startup" || firstArg = "-StartMenu")
            return true

        if (firstArg = "-SendTo" && A_Args.Length >= 2) {
            sendToPath := A_Args[2]                                         ; Not "Path" - that's the Util.ahk class, and a local assigned
                                                                            ; anywhere in a function shadows the class for the WHOLE function.
            SplitPath(sendToPath, &Desc, , &fileExt)                        ; Extra name from _Path (if _Type is dir and has "." in path, nameNoExt will not get full folder name)

            fileType := InStr(FileExist(sendToPath), "D") ? "Dir" : "File"  ; Default Type is File, Set Type is Dir only if the file exists and is a directory

            if (fileExt = "lnk" && g_CONFIG["SendToGetLnk"]) {
                FileGetShortcut(sendToPath, &sendToPath, , &fileArg, &Desc)
                sendToPath .= " " fileArg
            }
            CommandManager.Open("UserCommand", fileType, sendToPath, Desc, 1, "")   ; Add new command to database
            return true
        }
        return false
    }

    ;---------------------------------------------------------------------------
    ; Tray / right-click menus
    ;---------------------------------------------------------------------------
    ; 创建任务栏托盘程序图标
    static CreateTrayMenu() {
        if !g_CONFIG["ShowTrayIcon"] {
            A_IconHidden := 1
            return
        }

        static trayMenu := ""                                               ; 只在第一次调用时创建

        if !IsObject(trayMenu) {
            trayMenu := A_TrayMenu
            try {
                TraySetIcon("imageres.dll", -100)

                trayMenu.Delete()                                           ; 删除默认项
                trayMenu.Add(g_LNG[300], ToggleWindow)
                MainWindow._AddCommonMenuItems(trayMenu)
                MainWindow._SetMenuItemIcons(trayMenu)

                trayMenu.Default    := g_LNG[300]
                trayMenu.ClickCount := 1
                A_IconTip           := g_TITLE
                A_IconHidden        := 0

                g_LOG.Debug("CreateTrayMenu: Create tray menu...OK")
            } catch as e {
                g_LOG.Debug("CreateTrayMenu: Error creating tray menu: " . e.Message)
            }
        }
    }

    ; Items shared by the tray menu and the main window's right-click menu.
    static _AddCommonMenuItems(menuObj) {
        menuObj.Add(g_LNG[301], (*) => Options())
        menuObj.Add(g_LNG[310], UserCommand)
        menuObj.Add()                                                       ; 分隔线
        menuObj.Add(g_LNG[302], Reindex)
        menuObj.Add(g_LNG[303], Usage)
        menuObj.Add(g_LNG[305], (*) => ListLines())
        menuObj.Add()                                                       ; 分隔线
        menuObj.Add(g_LNG[309], Update)
        menuObj.Add(g_LNG[304], About)
        menuObj.Add()
        menuObj.Add(g_LNG[307], RestartApp)
        menuObj.Add(g_LNG[308], Exit)
    }

    static _SetMenuItemIcons(menuObj, isContext := false) {
        if !isContext
            menuObj.SetIcon(g_LNG[300], "imageres.dll", -100)
        menuObj.SetIcon(g_LNG[301], "imageres.dll", -114)
        menuObj.SetIcon(g_LNG[310], "imageres.dll", -88)
        menuObj.SetIcon(g_LNG[302], "imageres.dll", -8)
        menuObj.SetIcon(g_LNG[303], "imageres.dll", -150)
        menuObj.SetIcon(g_LNG[305], "imageres.dll", -165)
        menuObj.SetIcon(g_LNG[309], "imageres.dll", -5338)
        menuObj.SetIcon(g_LNG[304], "imageres.dll", -81)
        menuObj.SetIcon(g_LNG[307], "imageres.dll", -5311)
        menuObj.SetIcon(g_LNG[308], "imageres.dll", -98)
    }

    ; 主界面右键菜单 (点在列表以外的地方)
    static _OnContextMenu(GuiObj, GuiCtrlObj, Item, IsRightClick, X, Y) {
        ctrlType := IsObject(GuiCtrlObj) ? GuiCtrlObj.Type : ""
        if (ctrlType = "ListView" || ctrlType = "StatusBar")               ; ListView has its own ContextMenu event
            return
        MainWindow._ShowMainMenu(X, Y)
    }

    static _ShowMainMenu(X, Y) {
        static mainMenu := ""                                               ; 只在第一次调用时创建

        if !IsObject(mainMenu) {
            mainMenu := Menu()
            try {
                MainWindow._AddCommonMenuItems(mainMenu)
                MainWindow._SetMenuItemIcons(mainMenu, true)
                g_LOG.Debug("_ShowMainMenu: Create main window context menu...OK")
            } catch as e {
                g_LOG.Debug("_ShowMainMenu: Error creating main window context menu: " . e.Message)
            }
        }
        mainMenu.Show(X, Y)
    }

    static _OnListContextMenu(GuiCtrlObj, rowNumber, IsRightClick, X, Y) {
        if (rowNumber = 0) {                                                ; 如果用户右键点击了列表行以外的地方
            MainWindow._ShowMainMenu(X, Y)
            return
        }

        if MainWindow.SyncCurrentCommand(rowNumber, false) {
            MainWindow._ShowListMenu(X, Y)
        } else {                                                            ; For cases like first hint page
            MainWindow.SyncCurrentCommand(rowNumber)
            MainWindow._ShowMainMenu(X, Y)
        }
    }

    static _ShowListMenu(X, Y) {
        static listMenu := ""                                               ; 只在第一次调用右键菜单时创建

        if !IsObject(listMenu) {
            listMenu := Menu()
            try {
                listMenu.Add(g_LNG[400], (*) => MainWindow.RunFocusedRow())
                listMenu.Add(g_LNG[401], OpenContainer)
                listMenu.Add(g_LNG[402], CopyCommand)
                listMenu.Add()
                listMenu.Add(g_LNG[403], NewCommand)
                listMenu.Add(g_LNG[404], EditCommand)
                listMenu.Add(g_LNG[405], DelCommand)

                listMenu.SetIcon(g_LNG[400], "imageres.dll", -100)
                listMenu.SetIcon(g_LNG[401], "imageres.dll", -3)
                listMenu.SetIcon(g_LNG[402], "imageres.dll", -5314)
                listMenu.SetIcon(g_LNG[403], "imageres.dll", -2)
                listMenu.SetIcon(g_LNG[404], "imageres.dll", -5306)
                listMenu.SetIcon(g_LNG[405], "imageres.dll", -5305)

                g_LOG.Debug("_ShowListMenu: Create list context menu...OK")
            } catch as e {
                g_LOG.Debug("_ShowListMenu: Error creating list context menu: " . e.Message)
            }
        }
        listMenu.Show(X, Y)
    }

    ;---------------------------------------------------------------------------
    ; Hotkeys
    ;---------------------------------------------------------------------------
    static RegisterHotkeys() {
        ; 注册全局热键
        HotIfWinActive
        try {
            Hotkey(g_HOTKEY["GlobalHotkey1"], ToggleWindow)
            Hotkey(g_HOTKEY["GlobalHotkey2"], ToggleWindow)

            g_LOG.Debug("RegisterHotkeys: Set global activate hotkeys...OK")
        } catch as e {
            g_LOG.Debug("RegisterHotkeys: Failed to set global activate hotkeys..." e.Message)
        }

        ; 注册主窗口热键, 使用 ahk_id Hwnd 增强可靠性
        HotIfWinActive("ahk_id " MainWindow.Gui.Hwnd)
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
            Hotkey("^z"         , UndoDelCommand)
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

            g_LOG.Debug("RegisterHotkeys: Set local hotkeys (F1-F4)...OK")
        } catch as e {
            g_LOG.Debug("RegisterHotkeys: Failed to set local hotkeys (F1-F4)..." e.Message)
        }

        Loop g_GUI["ListRows"] {
            try {
                Hotkey("!" . A_Index, (*) => MainWindow._OnRunRowHotkey())    ; 通过热键选择并运行指定命令 = Alt + index (1-9)
                Hotkey("^" . A_Index, (*) => MainWindow._OnSelectRowHotkey()) ; 通过热键选择指定命令 = Ctrl + index (1-9)

                g_LOG.Debug("RegisterHotkeys: Set command list local hotkey " A_Index "...OK")
            } catch as e {
                g_LOG.Debug("RegisterHotkeys: Failed to set command list local hotkey " A_Index . e.Message)
            }
        }

        ; ObjBindMethod() fixes the function name per hotkey right now - a closure
        ; here would capture the loop variables themselves, so every hotkey would
        ; end up running the last trigger.
        Loop 7 {
            keyName := "Hotkey"  . A_Index
            trigger := "Trigger" . A_Index
            if (g_HOTKEY.Has(keyName) && g_HOTKEY[keyName] != "" && g_HOTKEY[trigger] != "") { ; 自定义热键执行指定功能 = Hotkey + Trigger
                try {
                    Hotkey(g_HOTKEY[keyName], ObjBindMethod(MainWindow, "_RunBoundFunction", g_HOTKEY[trigger]))
                    g_LOG.Debug("RegisterHotkeys: Set customized function list local hotkey " A_Index " " g_HOTKEY[keyName] " <-> " g_HOTKEY[trigger] "...OK")
                } catch as e {
                    g_LOG.Debug("RegisterHotkeys: Failed to set customized function list local hotkey..."  e.Message)
                }
            }
        }

        ; 注册条件热键, 执行指定功能
        HotIfWinActive(g_HOTKEY["CondTitle"])
        if g_HOTKEY.Has("CondTitle") && g_HOTKEY.Has("CondHotkey") && g_HOTKEY.Has("CondAction") {
            try {
                Hotkey(g_HOTKEY["CondHotkey"], ObjBindMethod(MainWindow, "_RunBoundFunction", g_HOTKEY["CondAction"]))
                g_LOG.Debug("RegisterHotkeys: Set conditional hotkey " g_HOTKEY["CondHotkey"] " <-> " g_HOTKEY["CondAction"] " for " g_HOTKEY["CondTitle"] "...OK")
            } catch as e {
                g_LOG.Debug("RegisterHotkeys: Failed to set conditional hotkey..." e.Message)
            }
        }
        HotIfWinActive                                                      ; Turn off context, make subsequent hotkeys global again
    }

    ; Custom / conditional hotkey callback (the hotkey name is appended by Hotkey()).
    static _RunBoundFunction(funcName, *) {
        RunCommand("FUNC | " funcName)
        g_LOG.Debug("_RunBoundFunction: Execute function...=" funcName)
    }

    ; Alt + index: select that row and run it
    static _OnRunRowHotkey() {
        MainWindow._OnSelectRowHotkey()
        RunCommand(g_RUNTIME["CurrentCommand"])
    }

    ; Ctrl + index: select that row
    static _OnSelectRowHotkey() {
        index := SubStr(A_ThisHotkey, 2, 1)                                 ; Get index from hotkey ("!3" / "^3" -> 3)
        if (index <= g_MATCHED.Length)
            MainWindow.MoveSelection(index, true)
    }

    ;---------------------------------------------------------------------------
    ; Show / hide
    ;---------------------------------------------------------------------------
    static Show() {
        ; Remember the window that is active right now, so a Clip command knows where to paste.
        try {
            activeHwnd := WinExist("A")
            if (activeHwnd && activeHwnd != MainWindow.Gui.Hwnd)
                g_RUNTIME["LastWin"] := activeHwnd
        }

        MainWindow.Gui.Show()

        if (WinWaitActive("ahk_id " MainWindow.Gui.Hwnd, , 3)) {            ; Wait for the window to be active, ahk_id is more reliable than g_TITLE
            if (g_CONFIG["AutoEngIME"]) {
                Win.SwitchToEnglishIME()
            }
            MainWindow.Input.Focus()
            SendMessage(0xB1, 0, -1, MainWindow.Input.Hwnd)                 ; EM_SETSEL (0xB1)
        }
    }

    static Hide() {
        if (!g_CONFIG["KeepInput"]) {
            MainWindow.ClearInput()
        }

        MainWindow.Gui.Hide()
        ; CommandStore.UpdateUsage() only mutates g_USAGE in memory. Hide() fires on
        ; every dismiss, including a plain Esc/Alt+Space with nothing run, so it
        ; must NOT trigger a full ALTRun.json save here. The bumped count rides
        ; along on the next real save instead (a command run, a settings change,
        ; or app exit - see the OnExit handler near the top of ALTRun.ahk).
        CommandStore.UpdateUsage()
        MainWindow.SetStatus("TIP")                                         ; Update StatusBar tip information after GUI hide
    }

    static Toggle() {
        WinActive("ahk_id " MainWindow.Gui.Hwnd) ? MainWindow.Hide() : MainWindow.Show()
    }

    static _OnEscape() {
        (g_CONFIG["EscClearInput"] && MainWindow.Input.Value) ? MainWindow.ClearInput() : MainWindow.Hide()
    }

    ; 监听窗口失去焦点时自动关闭 (HideOnLostFocus)
    static _OnLoseFocus() {
        if (!WinExist("ahk_id " MainWindow.Gui.Hwnd) || g_RUNTIME["UseDisplay"])
            return

        if (!WinActive("ahk_id " MainWindow.Gui.Hwnd)) {
            MainWindow.Hide()
            g_LOG.Debug("_OnLoseFocus: ALTRun lose focus, auto closing...")
        }
    }

    ; Allow moving a captionless window by mouse-drag (WM_LBUTTONDOWN)
    static _OnLButtonDown(wParam, lParam, msg, hwnd) {
        if (hwnd != MainWindow.Gui.Hwnd)
            return
        DllCall("ReleaseCapture")
        SendMessage(0xA1, 2, 0, MainWindow.Gui.Hwnd)                        ; WM_NCLBUTTONDOWN, HTCAPTION
    }

    ; Drag a file/folder/shortcut onto the main window: pre-fill the Command Manager
    ; with it (File/Dir type + a guessed description) so the user can confirm/edit
    ; before it's actually saved - dropping never adds a command by itself.
    static _OnDropFiles(GuiObj, GuiCtrlObj, FileArray, X, Y) {
        if (!FileArray.Length)
            return

        droppedPath := FileArray[1]
        targetPath  := droppedPath
        if (SubStr(droppedPath, -3) = ".lnk") {
            try {
                FileGetShortcut(droppedPath, &target)
                if (target != "")
                    targetPath := target
            } catch as e {
                g_LOG.Debug("_OnDropFiles: FileGetShortcut failed on " droppedPath " - " e.Message)
            }
        }

        cmdType := DirExist(targetPath) ? "Dir" : "File"
        SplitPath(targetPath, , , , &nameNoExt)
        CommandManager.Open("UserCommand", cmdType, targetPath, nameNoExt, 1, "")
    }

    ;---------------------------------------------------------------------------
    ; Input box
    ;---------------------------------------------------------------------------
    static _OnInputChange() {
        SearchCommand(MainWindow.Input.Value)
    }

    static ClearInput() {
        MainWindow.Input.Focus()
        MainWindow.Input.Value := ""
        MainWindow._OnInputChange()                                         ; Setting .Value from code does not raise the Change event in v2
    }

    ; Tab only switches focus between the input box and the list
    static ToggleFocus() {
        if (MainWindow.Gui.FocusedCtrl.ClassNN = "Edit1") {                 ; FocusedCtrl.ClassNN: Edit1 or SysListView321
            MainWindow.ListView.Focus()
        } else {
            MainWindow.Input.Focus()
        }
    }

    ;---------------------------------------------------------------------------
    ; Result list
    ;---------------------------------------------------------------------------
    static ShowResults(rows := [], useDisplay := false) {
        lv := MainWindow.ListView
        lv.Opt("-Redraw")
        lv.Delete()
        g_RUNTIME["UseDisplay"] := useDisplay
        showSN := g_CONFIG["ShowSN"], shortenPath := g_CONFIG["ShortenPath"]

        for rowIndex, rowCommand in rows {
            parts := StrSplit(rowCommand, " | ")
            cmdType := parts.Length >= 1 ? parts[1] : ""
            cmdPath := parts.Length >= 2 ? parts[2] : ""
            cmdDesc := parts.Length >= 3 ? parts[3] : ""
            displayNo := showSN ? rowIndex : ""
            iconIndex := MainWindow._IconIndex(cmdPath, cmdType)

            if (cmdType = "Clip") {
                ; Clip text is not a path, show a single-line preview instead.
                cmdPath := Clip.ClipPreview(cmdPath)
            } else if (shortenPath && cmdType != "URL") {
                ; Keep URL full text, shorten other command paths for list readability.
                SplitPath(cmdPath, &cmdPath)
            }

            lv.Add("Icon" iconIndex, displayNo, cmdType, cmdPath, cmdDesc)
        }
        rowCount := lv.GetCount()
        statusText := (g_RUNTIME["CurrentCommand"] != "")
            ? GetCmdDisplayPath(g_RUNTIME["CurrentCommand"])
            : (rowCount ? lv.GetText(1, 3) : "")

        if (rowCount)
            lv.Modify(1, "Select Focus Vis")
        lv.Opt("+Redraw")
        MainWindow.SetStatus(statusText)
    }

    ; Makes g_MATCHED[rowNumber] the current command. Returns false for rows that
    ; aren't real commands (e.g. the tips page shown at startup).
    static SyncCurrentCommand(rowNumber, updateStatus := true) {
        if (g_MATCHED.Length >= rowNumber) {
            g_RUNTIME["CurrentCommand"] := g_MATCHED[rowNumber]
            if updateStatus
                MainWindow.SetStatus(GetCmdDisplayPath(g_RUNTIME["CurrentCommand"]))
            return true
        }
        if updateStatus
            MainWindow.SetStatus(MainWindow.ListView.GetText(rowNumber, 3))
        return false
    }

    ; step is relative to the selected row, or the target row itself when absolute is true.
    static MoveSelection(step := 1, absolute := false) {
        lv := MainWindow.ListView
        rowCount := lv.GetCount()
        if (rowCount = 0)
            return
        selectedRow := absolute ? step : lv.GetNext() + step                ; Get target row no. to be selected
        selectedRow := selectedRow > rowCount ? 1 : selectedRow             ; Listview cycle selection (Mod has bug on upward cycle)
        selectedRow := selectedRow < 1 ? rowCount : selectedRow

        MainWindow.SyncCurrentCommand(selectedRow)                          ; Get current command from selected row

        lv.Modify(selectedRow, "Select Focus Vis")                          ; make new index row selected, Focused, and Visible
    }

    static _OnListClick(GuiCtrlObj, rowNumber) {
        if (!rowNumber)                                                     ; 如果用户左键点击了列表行以外的地方
            return
        MainWindow.SyncCurrentCommand(rowNumber)                            ; Get current command from clicked row
    }

    ; List double click / "Run" context menu item
    static RunFocusedRow() {
        focusedRow := MainWindow.ListView.GetNext(0, "Focused")             ; Only operate the focused row instead of all selected rows
        if (!focusedRow)
            return

        if MainWindow.SyncCurrentCommand(focusedRow, false) {
            RunCommand(g_RUNTIME["CurrentCommand"])
        }
    }

    ; Ctrl+C / "Copy" context menu item: copies the focused row's command path,
    ; or does a normal copy when the input box has the focus.
    static CopyFocusedCommand() {
        focusedRow := MainWindow.ListView.GetNext(0, "Focused")             ; Only operate the focused row instead of all selected rows
        if (!focusedRow)
            return

        MainWindow.SyncCurrentCommand(focusedRow, false)

        if (MainWindow.Gui.FocusedCtrl.ClassNN = "SysListView321") {
            A_Clipboard := MainWindow.ListView.GetText(focusedRow, 3)       ; The focused row's 3rd field (path)
        } else {
            SendInput("^c")                                                 ; Input box or status box is focused
        }
    }

    ;---------------------------------------------------------------------------
    ; Status bar
    ;---------------------------------------------------------------------------
    static SetStatus(text) {
        if (text = "TIP")
            text := g_LNG[51] g_LNG[Random(52, 71)]                         ; Randomly select a tip from hint list g_LNG 52~71

        MainWindow.Status.Value := text
    }

    ;---------------------------------------------------------------------------
    ; Icons
    ;---------------------------------------------------------------------------
    ; ImageList index for a result row. filePath, not path - "path" and "Path"
    ; (the Util.ahk class) are the same identifier to AHK, case-insensitive.
    static _IconIndex(filePath, type) {
        if not g_CONFIG["ShowIcon"]
            return 0

        iconMap := MainWindow.IconMap
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
            return iconMap.Has("CLIP") ? iconMap["CLIP"] : 2
        } else if (type = "FILE") {
            filePath := Path.Resolve(filePath)                              ; Must store in var for afterward use, trim space (in Path.Resolve)
            SplitPath(filePath, , , &fileExt)
            if (fileExt ~= "^(?i:EXE|ICO|ANI|CUR|LNK)$") {                  ; File types that have their own icon: cache by path
                return iconMap.Has(filePath) ? iconMap[filePath] : MainWindow._LoadIcon(filePath, filePath)
            }
            return iconMap.Has(fileExt) ? iconMap[fileExt] : MainWindow._LoadIcon(filePath, fileExt) ; Other types (pdf, xlsx...): cache by extension
        } else if (type = "App") {
            return 2
        }
    }

    ; Loads a file's shell icon into the ImageList and caches its index under cacheKey (extension or path).
    static _LoadIcon(filePath, cacheKey) {
        sfi_size := A_PtrSize + 688
        sfi      := Buffer(sfi_size)                                        ; SHFILEINFO structure
        iconSize := g_CONFIG["LargeIcons"] ? 0x100 : 0x101                  ; 0x100 is SHGFI_ICON+SHGFI_LARGEICON, 0x101 is SHGFI_ICON+SHGFI_SMALLICON

        try {
            if not DllCall("Shell32\SHGetFileInfoW", "Str", filePath, "UInt", 0, "Ptr", sfi, "UInt", sfi_size, "UInt", iconSize)
                return 2                                                    ; Use default function icon instead of out-of-bounds index

            hIcon := NumGet(sfi, 0, "Ptr")                                  ; Icon successfully loaded. Extract the hIcon member from the structure
            iconIndex := DllCall("ImageList_ReplaceIcon", "ptr", MainWindow.ImageList, "int", -1, "ptr", hIcon) + 1 ; +1: zero-based -> one-based
            DllCall("DestroyIcon", "Ptr", hIcon)                            ; Now that it's been copied into the ImageList, the original should be destroyed
            MainWindow.IconMap[cacheKey] := iconIndex
            return iconIndex
        } catch as e {
            g_LOG.Debug("_LoadIcon: Error getting icon for " filePath ": " e.Message)
            return 2                                                        ; Use default function icon for error cases
        }
    }

    ;---------------------------------------------------------------------------
    ; Default background picture (g_GUI["Background"] = "Default")
    ;---------------------------------------------------------------------------
    ; Writes the embedded JPG to %Temp% once and returns its path ("" on failure).
    static _DefaultBackgroundFile() {
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
}
