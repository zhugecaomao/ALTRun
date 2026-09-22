;===============================================================================
; OptionsWindow.ahk - 设置窗口 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 8 个 Tab 的设置界面: CONFIG / GUI / Hotkey / Index / Listary / Plugins / Usage / About。
;
; 用法:
;   OptionsWindow.Show(ActTab := 1)   ; 打开设置窗口, ActTab 指定初始 Tab 序号
;
; 注意: 全局函数 Options(ActTab := 1) (F2 菜单 / 托盘菜单 / 右键菜单 / "Func | Options | ..."
; 内置命令用它的函数名做动态调用 %cmdPath%(), 必须是裸的全局函数, 且 "Options" 本身也在
; FuncList 数组里用于自定义热键的下拉列表) 仍留在 ALTRun.ahk 里, 只是内部改成调用
; OptionsWindow.Show() - 详见那边的注释。
;
; OptGUI/OptListView 原来是整个脚本里的 Global 变量, 但实际只在这个窗口的相关函数里
; 用到, 所以改成本类的 static 属性, 不再需要 Global 声明。
;===============================================================================

Class OptionsWindow {

    static G := ""                                                         ; Options 窗口的 Gui 对象 (原 Global OptGUI)
    static ListView := ""                                                  ; CONFIG Tab 里的设置勾选列表 (原 Global OptListView)

    static Show(ActTab := 1) {
        static FuncList := ["Unset", "Active", "ToggleWindow", "Google", "Bing"
            , "Everything", "TabFunc", "PrevCommand", "NextCommand", "CopyCommand"
            , "ClearInput", "RunCurrentCommand", "RankUp", "RankDown", "Reindex"
            , "About", "Usage", "Update", "UserCommand", "NewCommand", "EditCommand"
            , "DelCommand", "OpenCommandManager", "Options", "TurnMonitorOff", "EmptyRecycle"
            , "MuteVolume", "RestartApp", "Exit", "PTTools", "SPF2M"]

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
        g := OptionsWindow.G := Gui("+Owner" MainGUI.hwnd, g_LNG[2])         ; +Owner MainGUI.hwnd fix GUI flicking issue
        g.SetFont(optFont.opt, optFont.name)
        OptTab := g.AddTab3("Choose" ActTab, g_LNG[100])

        OptTab.UseTab(1) ; CONFIG Tab
        OptionsWindow.ListView := g.AddListView("w500 h300 Checked -Hdr", ["Settings"])
        for key, description in g_CONFIG_P1 {
            OptionsWindow.ListView.Add("Check" g_CONFIG[key], description)
        }
        OptionsWindow.ListView.ModifyCol(1, "AutoHdr")

        g.AddText("x24 yp+320", g_LNG[150])
        g.AddComboBox("x130 yp-5 w394 vFileMgr Choose1", [g_CONFIG["FileMgr"], "Explorer.exe", "C:\Apps\TotalCMD.exe /O /T /S"])
        g.AddText("x24 yp+40", g_LNG[151])
        g.AddComboBox("x130 yp-5 w394 vEverything Choose1", [g_CONFIG["Everything"], "C:\Apps\Everything.exe"])
        g.AddText("x24 yp+40", g_LNG[152])
        g.AddDDL("x130 yp-5 w394 Sort vHistoryLen Choose" g_CONFIG["HistoryLen"]*0.1, [10,20,30,40,50,60])

        OptTab.UseTab(2) ; GUI Tab
        g.AddGroupBox("w500 h420", g_LNG[170])
        g.AddText("x33 yp+25", g_LNG[171])
        g.AddDDL("x183 yp-5 w330 vListRows Choose" g_GUI["ListRows"], [1,2,3,4,5,6,7,8,9]) ; ListRows limit <= 9
        g.AddText("x33 yp+45", g_LNG[172])
        g.AddComboBox("x183 yp-5 w330 vColWidth Choose1", [g_GUI["ColWidth"], "20,0,460,AutoHdr", "30,46,460,AutoHdr"])
        g.AddText("x33 yp+45", g_LNG[176])
        g.AddEdit("x183 yp-5 w120 +Number vWinX", g_GUI["WinX"])
        g.AddText("x345 yp", "x")
        g.AddEdit("x393 yp w120 +Number vWinY", g_GUI["WinY"])

        g.AddText("x33 yp+45", g_LNG[173])
        g.AddEdit("x183 yp w240 r1 -E0x200 +ReadOnly vMainGUIFont", g_GUI["MainGUIFont"]).SetFont(mainFont.opt, mainFont.name)
        g.AddButton("x433 yp-5 w80", g_LNG[182]).OnEvent("Click", (*) => OptionsWindow.SelectFont("MainGUIFont"))
        g.AddText("x33 yp+45", g_LNG[174])
        g.AddEdit("x183 yp w240 r1 -E0x200 +ReadOnly vOptGUIFont", g_GUI["OptGUIFont"])
        g.AddButton("x433 yp-5 w80", g_LNG[182]).OnEvent("Click", (*) => OptionsWindow.SelectFont("OptGUIFont"))
        g.AddText("x33 yp+45", g_LNG[175])
        g.AddEdit("x183 yp w240 r1 -E0x200 +ReadOnly vMainSBFont", g_GUI["MainSBFont"]).SetFont(sbFont.opt, sbFont.name)
        g.AddButton("x433 yp-5 w80", g_LNG[182]).OnEvent("Click", (*) => OptionsWindow.SelectFont("MainSBFont"))

        g.AddText("x33 yp+45", g_LNG[179])
        g.AddEdit("x183 yp w240 r1 -E0x200 +ReadOnly vMainGUIColor", g_GUI["MainGUIColor"])
        g.AddButton("x433 yp-5 w80", g_LNG[183]).OnEvent("Click", (p*) => OptionsWindow.PickMainGUIColor(p*))
        g.AddText("x33 yp+45", g_LNG[178])
        g.AddEdit("x183 yp w240 r1 -E0x200 +ReadOnly vCMDListColor", g_GUI["CMDListColor"])
        g.AddButton("x433 yp-5 w80", g_LNG[183]).OnEvent("Click", (p*) => OptionsWindow.PickCMDListColor(p*))

        g.AddText("x33 yp+45", g_LNG[180])
        g.AddComboBox("x183 yp-5 w240 vBackground Choose1", [g_GUI["Background"], "Default", "None", "ALTRun.jpg", "C:\Path\Picture.jpg"])
        g.AddButton("x433 yp-2 w80 vSelectBackground", g_LNG[184]).OnEvent("Click", (p*) => OptionsWindow.SelectBackground(p*))
        g.AddText("x33 yp+45", g_LNG[181])
        g.AddSlider("x183 yp-5 w330 Range50-255 TickInterval5 Tooltip vTransparency", g_GUI["Transparency"])

        OptTab.UseTab(3) ; Hotkey Tab
        g.AddGroupBox("w500 h115", g_LNG[191])
        g.AddText("x33 yp+25", g_LNG[192])
        g.AddHotkey("x285 yp-4 w230 vGlobalHotkey1", g_HOTKEY["GlobalHotkey1"])
        g.AddText("x33 yp+35", g_LNG[193])
        g.AddHotkey("x285 yp-4 w230 vGlobalHotkey2", g_HOTKEY["GlobalHotkey2"])
        g.AddText("x33 yp+35", g_LNG[194])
        g.AddLink("x285 yp w230", "<a>" g_LNG[195] "</a>").OnEvent("Click", (p*) => OptionsWindow.ResetHotkey(p*))

        g.Add("GroupBox", "x24 yp+38 w500 h290", g_LNG[200])
        Loop 7 {
            g.AddText("x33 yp+40", g_LNG[201])
            g.AddHotkey("x143 yp-5 w120 vHotkey" A_Index, g_HOTKEY["Hotkey" A_Index])
            g.AddText("x285 yp+5", g_LNG[202])
            g.AddDDL("x395 yp-5 w120 vTrigger" A_Index " Choose" GetArrayIndex(g_HOTKEY["Trigger" A_Index], FuncList), FuncList)
        }

        OptionsWindow.ToggleGlobalHotkeys("Off", "Options")                 ; Turn off global hotkeys in options

        OptTab.UseTab(4) ; INDEX Tab
        g.AddGroupBox("w500 h220", g_LNG[160])
        g.AddText("x33 yp+25", g_LNG[161])
        g.AddComboBox("x183 yp-5 w330 vIndexDir Choose1", [g_CONFIG["IndexDir"], "A_ProgramsCommon,A_StartMenu"])
        g.AddText("x33 yp+45", g_LNG[162])
        g.AddComboBox("x183 yp-5 w330 vIndexType Choose1", [g_CONFIG["IndexType"], "*.lnk,*.exe"])
        g.AddText("x33 yp+45", g_LNG[164])
        g.AddDropDownList("x183 yp-5 w330 vIndexDepth Choose" g_CONFIG["IndexDepth"], [1,2,3,4,5,6,7,8,9])
        g.AddText("x33 yp+45", g_LNG[163])
        g.AddComboBox("x183 yp-5 w330 vIndexExclude Choose1", [g_CONFIG["IndexExclude"], "Uninstall *"])
        g.AddCheckBox("x33 yp+45 vIndexStoreApp Checked" g_CONFIG["IndexStoreApp"], g_LNG[165])

        OptTab.UseTab(5) ; LISTARY Tab
        g.AddGroupBox("w500 h145", g_LNG[211])
        g.AddText("x33 yp+30", g_LNG[212])
        g.AddComboBox("x183 yp-5 w330 vFileMgrID Choose1", [g_CONFIG["FileMgrID"], "ahk_class CabinetWClass", "ahk_class CabinetWClass, ahk_class TTOTAL_CMD"])
        g.AddText("x33 yp+45", g_LNG[213])
        g.AddComboBox("x183 yp-5 w330 vDialogWin Choose1", [g_CONFIG["DialogWin"], "ahk_class #32770"])
        g.AddText("x33 yp+45", g_LNG[214])
        g.AddComboBox("x183 yp-5 w330 vExcludeWin Choose1", [g_CONFIG["ExcludeWin"], "ahk_class SysListView32, ahk_exe Explorer.exe"])
        g.AddGroupBox("x24 yp+50 w500 h145", g_LNG[215])
        g.AddText("x33 yp+30", g_LNG[216])
        g.AddHotkey("x183 yp-5 w330 vTotalCMDDir", g_HOTKEY["TotalCMDDir"])
        g.AddText("x33 yp+45", g_LNG[217])
        g.AddHotkey("x183 yp-5 w330 vExplorerDir", g_HOTKEY["ExplorerDir"])
        g.AddCheckBox("x33 yp+45 vAutoSwitchDir Checked" g_CONFIG["AutoSwitchDir"], g_LNG[218])

        OptTab.UseTab(6) ; PLUGINS Tab
        g.AddGroupBox("w500 h110", g_LNG[251])
        g.AddText("x33 yp+30", g_LNG[252])
        g.AddComboBox("x183 yp-5 w330 vAutoDateAtEnd Choose1", [g_HOTKEY["AutoDateAtEnd"], "ahk_class TCmtEditForm,ahk_exe Notepad4.exe"])
        g.AddText("x33 yp+45", g_LNG[253])
        g.AddHotkey("x183 yp-5 w80 vAutoDateAEHKey", g_HOTKEY["AutoDateAEHKey"])
        g.AddText("x300 yp+5", g_LNG[254])
        g.AddDDL("x395 yp-5 w120 vAutoDateAEFormat Choose1", ["- dd.MM.yyyy"])

        g.AddGroupBox("x24 y+30 w500 h110", g_LNG[255])
        g.AddText("x33 yp+30", g_LNG[252])
        g.AddComboBox("x183 yp-5 w330 vAutoDateBefExt Choose1", [g_HOTKEY["AutoDateBefExt"], "ahk_class CabinetWClass,ahk_class Progman,ahk_class WorkerW,ahk_class #32770"])
        g.AddText("x33 yp+45", g_LNG[253])
        g.AddHotkey("x183 yp-5 w80 vAutoDateBEHKey", g_HOTKEY["AutoDateBEHKey"])
        g.AddText("x300 yp+5", g_LNG[254])
        g.AddDDL("x395 yp-5 w120 vAutoDateBEFormat Choose1", ["- dd.MM.yyyy"])

        g.AddGroupBox("x24 y+30 w500 h110", g_LNG[259])
        g.AddText("x33 yp+30", g_LNG[260])
        g.AddComboBox("x183 yp-5 w330 vCondTitle Choose1", [g_HOTKEY["CondTitle"]])
        g.AddText("x33 yp+45", g_LNG[261])
        g.AddComboBox("x183 yp-5 w80 vCondHotkey Choose1", [g_HOTKEY["CondHotkey"]])
        g.AddText("x300 yp+5", g_LNG[262])
        g.AddDDL("x395 yp-5 w120 vCondAction Choose" GetArrayIndex(g_HOTKEY["CondAction"], FuncList), FuncList)

        OptTab.UseTab(7) ; USAGE Tab
        g.AddGroupBox("x66 y80 w445 h300", )

        g_USAGE[A_YYYY . A_MM . A_DD] := g_USAGE.Has(A_YYYY . A_MM . A_DD) ? g_USAGE[A_YYYY . A_MM . A_DD] : 1
        for date, count in g_USAGE { ; Draw usage graph
            g.AddProgress("c94DD88 Vertical y96 w14 h280 xm+" 50+A_Index*14 " Range0-" g_RUNTIME["Max"]+10, count)
        }

        g.AddText("x24 yp-5 cGray",g_RUNTIME["Max"])
        g.AddText("x24 yp+140 cGray", Round(g_RUNTIME["Max"]/2))
        g.AddText("x24 yp+140 cGray", 0)
        g.AddText("x66 yp+15 cGray", g_LNG[500])
        g.AddText("x476 yp cGray", g_LNG[501])
        g.AddText("x66 yp+33", g_LNG[502])
        g.AddEdit("x400 yp-5 w100 r1 -E0x200 +ReadOnly Right vRunCount", g_CONFIG["RunCount"])
        g.AddText("x66 yp+35", g_LNG[503])
        g.AddEdit("x400 yp-5 w100 r1 -E0x200 +ReadOnly Right", g_USAGE[A_YYYY . A_MM . A_DD])

        OptTab.UseTab(8) ; ABOUT Tab
        g.AddPic("x33 y+20 w48 h-1 Icon-100", "imageres.dll")
        g.AddText("x96 yp+5 w400", g_TITLE).SetFont("S11")
        g.AddLink("xp yp+45 w400", g_LNG[601])

        OptTab.UseTab()  ; 后续添加的控件将不属于前面的选项卡控件
        g.AddButton("Default x278 w80", g_LNG[7]).OnEvent("Click", (p*) => OptionsWindow.OPTButtonOK(p*))
        g.AddButton("x368 yp w80", g_LNG[8]).OnEvent("Click", (p*) => OptionsWindow.OPTGuiClose(p*))
        g.AddButton("x458 yp w80", g_LNG[9]).OnEvent("Click", (*) => Run("https://github.com/zhugecaomao/ALTRun/wiki"))
        g.OnEvent("Close", (p*) => OptionsWindow.OPTGuiClose(p*))
        g.OnEvent("Escape", (p*) => OptionsWindow.OPTGuiClose(p*))

        g_LOG.Debug("Options: Load options window...OK, elapsed time=" A_TickCount - t "ms")
        OutputDebug("Options: Load options window...OK, elapsed time=" A_TickCount - t "ms")
        g.Show("Center")
        return
    }

    static ResetHotkey(*) {
        OptionsWindow.G["GlobalHotkey1"].Value := "!Space"
        OptionsWindow.G["GlobalHotkey2"].Value := "!r"
        return
    }

    static SelectFont(TargetVar := "MainGUIFont") {
        ; Set the fontObj (optional) - only set the ones you want to pre-select
		; fontObj := Map("name","Terminal","size",14,"color",0xFF0000,"strike",1,"underline",1,"italic",1,"bold",1)
        initFont := StrSplit(g_GUI[TargetVar], ",")[1]
        fontObj  := Map("name", initFont)
        fontObj  := FontDialog.Choose(fontObj, OptionsWindow.G.hwnd)
        if (!fontObj)
            return

        OptionsWindow.G[TargetVar].Text := fontObj["name"] ", " fontObj["str"]  ; 更新控件字体并设置显示文本
        OptionsWindow.G[TargetVar].SetFont(fontObj["str"], fontObj["name"])
        g_LOG.Debug("SelectFont: OptGUI[" TargetVar "] font set to=" fontObj["str"] ", " fontObj["name"])
    }

    static PickCMDListColor(*) {
        color := ColorDialog.Choose(g_GUI["CMDListColor"], OptionsWindow.G.hwnd, , "full")  ; hwnd and custColorObj are optional
        if (color = -1)
            return

        ;g_GUI["CMDListColor"]        := color
        OptionsWindow.G["CMDListColor"].Value := color                      ; 更新选项窗口控件并设置控件颜色
        ;OptGUI["CMDListColor"].Opt("c" color)
    }

    static PickMainGUIColor(*) {
        color := ColorDialog.Choose(g_GUI["MainGUIColor"], OptionsWindow.G.hwnd, , "full")
        if (color = -1)
            return

        ;g_GUI["MainGUIColor"]        := color
        OptionsWindow.G["MainGUIColor"].Value := color
        ;OptGUI["MainGUIColor"].Opt("c" color)
    }

    static SelectBackground(*) {
        OptionsWindow.G.Opt("+OwnDialogs")                                  ; Make open dialog Modal

        file := FileSelect(3, , , 'Image Files (*.jpg; *.png; *.bmp; *.gif)')
        if (file = "")
            return

        OptionsWindow.G["Background"].Text := file
        g_LOG.Debug("SelectBackground: Background image selected=" file)
    }

    static OPTButtonOK(*) {
        OptionsWindow.SaveConfig()
        Reload
    }

    static OPTGuiClose(*) {
        g_LOG.Debug("OPTGuiClose: Closing Options window...")

        OptionsWindow.ToggleGlobalHotkeys("On", "OPTGuiClose")              ; Turn on global hotkeys

        OptionsWindow.G.Hide()
        g_LOG.Debug("OPTGuiClose: OptGUI.Hide...OK")
        return
    }

    static ToggleGlobalHotkeys(mode, caller := "") {
        HotIfWinActive
        for _, hk in ["GlobalHotkey1", "GlobalHotkey2"] {
            if (g_HOTKEY[hk] = "")
                continue
            try {
                Hotkey(g_HOTKEY[hk], (p*) => ToggleWindow(p*), mode)
                g_LOG.Debug(caller ": Turn " mode " " hk "...OK")
            } catch as e {
                g_LOG.Debug(caller ": Turn " mode " " hk "...Failed: " e.Message)
            }
        }
    }

    ; NOTE: there is no more LoadConfig() - AppData.LoadAppData() (see the JSON command
    ; storage section) loads Config/Gui/Hotkey/Usage/History/Benchmark from
    ; ALTRun.json at startup, in one pass alongside the commands.

    static SaveConfig() {
        g := OptionsWindow.G
        g.Submit()
        checkedRows := Map(), row := 0
        while (row := OptionsWindow.ListView.GetNext(row, "C"))
            checkedRows[row] := 1

        ; Tab1 checklist values (plain booleans, no type coercion needed).
        for key, _ in g_CONFIG_P1
            g_CONFIG[key] := checkedRows.Has(A_Index) ? 1 : 0

        static configKeys := Array("FileMgr", "Everything", "HistoryLen", "RunCount"
            , "AutoSwitchDir", "IndexDir", "IndexType", "IndexDepth"
            , "IndexExclude", "IndexStoreApp", "DialogWin", "FileMgrID", "ExcludeWin")

        for _, key in configKeys
            g_CONFIG[key] := OptionsWindow.CoerceLikeCurrent(g_CONFIG[key], OptionsWindow.GetOptCtrlValue(g[key]))

        for key, _ in g_GUI
            g_GUI[key] := OptionsWindow.CoerceLikeCurrent(g_GUI[key], OptionsWindow.GetOptCtrlValue(g[key]))

        for key, _ in g_HOTKEY
            g_HOTKEY[key] := OptionsWindow.GetOptCtrlValue(g[key])          ; Hotkeys/window titles are always text

        AppData.SaveAppData()

        g_LOG.Debug("SaveConfig: Save config...OK")
        return
    }

    static GetOptCtrlValue(ctrl) {
        return InStr(",CheckBox,Slider,Hotkey,", "," ctrl.Type ",") ? ctrl.Value : ctrl.Text
    }

    static CoerceLikeCurrent(currentVal, newVal) {                         ; Keep a setting's number-vs-text type stable across saves,
        return (Type(currentVal) = "Integer" || Type(currentVal) = "Float") ; so JSON.stringify writes e.g. 300 instead of "300", while
            && IsNumber(newVal) ? newVal + 0 : newVal                       ; a hex color string like "0xFFFFFF" is left alone.
    }
}
