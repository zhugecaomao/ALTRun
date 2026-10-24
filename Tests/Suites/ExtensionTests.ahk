;===============================================================================
; ExtensionTests.ahk - Extensions: CapsLock switch, date stamp, snippet expansion, quick switch
;   folders, PT tools, self-update and selection actions
;-------------------------------------------------------------------------------
; Included by Tests\RunTests.ahk. Every static method whose name doesn't start with "_" is a test;
; TestRunner runs them in name order, and each one must put back the shared state it changes.
;===============================================================================
class ExtensionTests {
    ; 按 CapsLock 切换输入法: 设置的取值、下一个输入法的选择
    static CapsLockSwitch() {
        eq := (n, a, e) => TestRunner.Equal("CapsLockSwitch." n, a, e)
        eq("off by default", AppSettings.Defaults()["General"]["CapsLock"], "")
        eq("modes", (CapsLockSwitch.IsMode("Layout") ? 1 : 0) (CapsLockSwitch.IsMode("Mode") ? 1 : 0) (CapsLockSwitch.IsMode("") ? 1 : 0) (CapsLockSwitch.IsMode("x") ? 1 : 0), "1100")
        eq("hold time", CapsLockSwitch.HoldMs, 300)
        eq("notification text", CapsLockSwitch.StateText(true) "|" CapsLockSwitch.StateText(false), "Caps Lock On|Caps Lock Off")
        eq("uses the keyboard hook", InStr(CapsLockSwitch.Key, "$") ? 1 : 0, 1)    ; 不用钩子时拦不住大写切换, 按住时也会不停切换
        eq("with any modifier", InStr(CapsLockSwitch.Key, "*") ? 1 : 0, 1)        ; 远程桌面补发的 Alt+CapsLock 等也要拦下
        registered := [], savedInterval := A_HotkeyInterval
        CapsLockSwitch.Register((key, fn) => registered.Push(key))
        eq("down and up hotkeys", registered.Length = 2 ? registered[1] "|" registered[2] : "", "*$CapsLock|*$CapsLock up")
        eq("no too-many-hotkeys warning", A_HotkeyInterval, 0)
        A_HotkeyInterval := savedInterval

        ; 大写锁定的提示可以关掉 (General.CapsLockNotify)
        eq("notice on by default", AppSettings.Defaults()["General"]["CapsLockNotify"], 1)
        notices := [], savedNotify := App.GetOwnPropDesc("Notify")
        App.DefineProp("Notify", {Call: (this, text, *) => notices.Push(text)})
        savedNotice := AppSettings.General.Has("CapsLockNotify") ? AppSettings.General["CapsLockNotify"] : 1
        AppSettings.General["CapsLockNotify"] := 1, App.CapsLockNotice(true)
        AppSettings.General["CapsLockNotify"] := 0, App.CapsLockNotice(false)
        AppSettings.General["CapsLockNotify"] := savedNotice
        App.DefineProp("Notify", savedNotify)
        eq("notice shown only when on", notices.Length = 1 ? notices[1] : notices.Length, "Caps Lock On")

        ; 管理员窗口里的小助手 (CapsLockHelper, #158)
        eq("helper off by default", AppSettings.Defaults()["General"]["CapsLockElevated"], 0)
        eq("helper wanted", (CapsLockHelper.Wanted(Map("CapsLock", "Mode", "CapsLockElevated", 1)) ? 1 : 0)
            . (CapsLockHelper.Wanted(Map("CapsLock", "", "CapsLockElevated", 1)) ? 1 : 0)
            . (CapsLockHelper.Wanted(Map("CapsLock", "Layout", "CapsLockElevated", 0)) ? 1 : 0)
            . (CapsLockHelper.Wanted(Map("CapsLock", "Layout")) ? 1 : 0), "1000")
        settingsFile := A_Temp "\ALTRunHelperTest.json"
        try FileDelete(settingsFile)
        FileAppend('{"General": {"CapsLock": "Layout", "CapsLockElevated": 1}}', settingsFile, "UTF-8")
        eq("helper reads the mode", CapsLockHelper.ModeFromSettings(settingsFile), "Layout")
        FileDelete(settingsFile)
        FileAppend('{"General": {"CapsLock": "Layout", "CapsLockElevated": 0}}', settingsFile, "UTF-8")
        eq("helper: option off", CapsLockHelper.ModeFromSettings(settingsFile), "")
        FileDelete(settingsFile)
        eq("helper: no settings file", CapsLockHelper.ModeFromSettings(settingsFile), "")
        xml := CapsLockHelper.TaskXml("C:\A & B\ALTRun.exe", "-CapsLockHelper", "C:\A & B", "PC\me")
        eq("task: highest privileges", InStr(xml, "<RunLevel>HighestAvailable</RunLevel>") ? 1 : 0, 1)
        eq("task: interactive user", InStr(xml, "<UserId>PC\me</UserId><LogonType>InteractiveToken</LogonType>") ? 1 : 0, 1)
        eq("task: no time limit, runs on battery", (InStr(xml, "<ExecutionTimeLimit>PT0S</ExecutionTimeLimit>") && InStr(xml, "<DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>")) ? 1 : 0, 1)
        eq("task: escaped path", InStr(xml, "<Command>C:\A &amp; B\ALTRun.exe</Command><Arguments>-CapsLockHelper</Arguments>") ? 1 : 0, 1)
        eq("task: no trigger (started by ALTRun)", InStr(xml, "<Triggers>") ? 1 : 0, 0)
        shown := []                                                         ; 大写锁定的提示交给 Notify (小助手里换成告诉 ALTRun)
        savedNotify := CapsLockSwitch.Notify, CapsLockSwitch.Notify := (on) => shown.Push(on)
        before := GetKeyState("CapsLock", "T")
        CapsLockSwitch.ToggleCapsLock()
        eq("notify on toggle", shown.Length = 1 && shown[1] = !before ? 1 : 0, 1)
        SetCapsLockState(before ? "On" : "Off"), CapsLockSwitch.Notify := savedNotify

        ; 按下 / 松开的顺序 (假的时钟, 定时器手动触发)
        saved := Map()
        for name in ["Clock", "ToggleCapsLock", "Switch", "ShiftDown"]
            saved[name] := CapsLockSwitch.GetOwnPropDesc(name)
        clock := {Now: 0}, calls := {Text: ""}
        CapsLockSwitch.Clock := () => clock.Now
        CapsLockSwitch.DefineProp("ToggleCapsLock", {Call: (*) => calls.Text .= "caps;"})
        CapsLockSwitch.DefineProp("Switch", {Call: (*) => calls.Text .= "switch;"})
        shift := {Down: false}
        CapsLockSwitch.DefineProp("ShiftDown", {Call: (*) => shift.Down})
        at := (ms) => clock.Now := ms
        try {
            CapsLockSwitch._down := false
            at(1000), CapsLockSwitch.Down(), at(1100), CapsLockSwitch.Up()
            eq("tap: switch once", calls.Text, "switch;")
            calls.Text := ""
            at(2000), CapsLockSwitch.Down()
            Loop 20                                                         ; 按住时的自动重复, 每 30 ms 一下
                at(2000 + A_Index * 30), CapsLockSwitch.Down()
            CapsLockSwitch.Hold()                                           ; 定时器: 按住满 0.3 秒
            Loop 10
                at(2600 + A_Index * 30), CapsLockSwitch.Down()
            at(3000), CapsLockSwitch.Up()
            eq("hold: caps lock once, never switches", calls.Text, "caps;")
            calls.Text := ""
            at(4000), CapsLockSwitch.Down(), at(4400), CapsLockSwitch.Up()  ; 定时器没来得及运行
            eq("late timer: still a hold", calls.Text, "caps;")
            calls.Text := ""
            at(5000), CapsLockSwitch.Down(), at(7000), CapsLockSwitch.Down(), at(7050), CapsLockSwitch.Up()   ; 上一次的松开丢了
            eq("lost key-up: new press", calls.Text, "switch;")
            calls.Text := ""
            CapsLockSwitch.Up()
            eq("extra key-up ignored", calls.Text, "")
            at(8000), CapsLockSwitch.Down(), at(8000.4), CapsLockSwitch.Up()  ; Chrome 远程桌面补发来对齐大写锁定的
            eq("remote sync tap ignored", calls.Text, "")
            at(9000), CapsLockSwitch.Down(), at(9040), CapsLockSwitch.Up()
            eq("quick real tap still switches", calls.Text, "switch;")
            calls.Text := ""
            shift.Down := true
            at(10000), CapsLockSwitch.Down(), at(10080), CapsLockSwitch.Up()
            eq("shift+caps lock: caps lock as usual", calls.Text, "caps;")
            calls.Text := ""
            at(11000), CapsLockSwitch.Down(), at(11000.3), CapsLockSwitch.Up()   ; 打 Shift+字母时远程桌面补发的
            eq("remote sync tap with shift ignored", calls.Text, "")
            shift.Down := false
            CapsLockSwitch.Clock := saved["Clock"].Value
            ; Sleep 按 15.6 ms 一档的 A_TickCount 判断, 实际可能少睡将近一档: 睡 100 ms, 只要求量到 50 ms 以上
            first := CapsLockSwitch.Clock.Call(), Sleep(100), second := CapsLockSwitch.Clock.Call()
            TestRunner.True("CapsLockSwitch.precise clock", second - first >= 50 && second - first < 5000)
        } finally {
            SetTimer(CapsLockSwitch._holdTimer, 0)
            CapsLockSwitch._down := false
            for name, desc in saved
                CapsLockSwitch.DefineProp(name, desc)
        }
        eq("next layout", Win.NextLayout([0x4090409, 0x8040804], 0x4090409), 0x8040804)
        eq("wraps around", Win.NextLayout([0x4090409, 0x8040804], 0x8040804), 0x4090409)
        eq("unknown current: first", Win.NextLayout([0x4090409, 0x8040804], 0x4110411), 0x4090409)
        eq("no layouts", Win.NextLayout([], 0x4090409), 0)
    }

    static AutoDate() {
        eq := (n, a, e) => TestRunner.Equal("AutoDate." n, a, e)
        today := "23.09.2026"
        ; Ctrl+D 只在改名的编辑框里接管 (其它时候按键原样交给程序, 不补发)
        eq("edit box", AutoDate.IsEditControl("Edit1"), true)
        eq("TC rename box", AutoDate.IsEditControl("TInEdit1"), true)
        eq("scintilla", AutoDate.IsEditControl("Scintilla1"), true)
        eq("file list", AutoDate.IsEditControl("DirectUIHWND3"), false)
        eq("no focus", AutoDate.IsEditControl(""), false)
        eq("with ext", AutoDate.AddDateToName("Report.docx", today), "Report - 23.09.2026.docx")
        eq("update date", AutoDate.AddDateToName("Report - 01.01.2020.docx", today), "Report - 23.09.2026.docx")
        eq("update dash date", AutoDate.AddDateToName("Report-01.01.2020.docx", today), "Report - 23.09.2026.docx")
        eq("folder", AutoDate.AddDateToName("Project", today), "Project - 23.09.2026")
        eq("folder update", AutoDate.AddDateToName("Project - 01.01.2020", today), "Project - 23.09.2026")
        eq("numeric ext", AutoDate.AddDateToName("Version 1.2", today), "Version 1.2 - 23.09.2026")
        eq("dot space", AutoDate.AddDateToName("1. DWG", today), "1. DWG - 23.09.2026")
        ; 设置里换了日期格式: 末尾那种格式的日期照样更新, 不会再加一个
        eq("iso update", AutoDate.AddDateToName("Report - 2020-01-01.docx", "2026-09-23", "yyyy-MM-dd"), "Report - 2026-09-23.docx")
        eq("iso folder", AutoDate.AddDateToName("Project - 2020-01-01", "2026-09-23", "yyyy-MM-dd"), "Project - 2026-09-23")
        eq("dotted iso folder", AutoDate.AddDateToName("Project - 2020.01.01", "2026.09.23", "yyyy.MM.dd"), "Project - 2026.09.23")
        eq("compact", AutoDate.AddDateToName("Plan_20200101.pdf", "20260923", "yyyyMMdd"), "Plan_20200101 - 20260923.pdf")
        eq("compact update", AutoDate.AddDateToName("Plan - 20200101.pdf", "20260923", "yyyyMMdd"), "Plan - 20260923.pdf")
        eq("month name", AutoDate.AddDateToName("Memo - 1 Jan 2020", "23 Sep 2026", "d MMM yyyy"), "Memo - 23 Sep 2026")
        eq("other format kept", AutoDate.AddDateToName("Report - 01.01.2020.docx", "2026-09-23", "yyyy-MM-dd"), "Report - 01.01.2020 - 2026-09-23.docx")
        eq("pattern", AutoDate.DatePattern("dd.MM.yyyy"), "\d{2}\.\d{2}\.\d{4}")
    }

    ; 粘贴方式 Auto: 短的单行文字直接输入 (不动剪贴板), 多行 / Tab / 很长 / emoji 用剪贴板
    static SnippetPasteMode() {
        eq := (n, a, e) => TestRunner.Equal("SnippetPasteMode." n, a, e)
        types := (text, mode) => ActionCatalog.TypesText(text, mode)
        long := StrReplace(Format("{:" ActionCatalog.AutoTypeMax "}", ""), " ", "x")
        eq("default", AppSettings.Defaults()["Features"]["Snippets"]["PasteMode"], "Auto")
        eq("auto: short line typed", types("Best regards, Bench", "Auto"), true)
        eq("auto: CJK typed", types("你好, 谢谢", "Auto"), true)
        eq("auto: at the limit typed", types(long, "Auto"), true)
        eq("auto: longer pasted", types(long "x", "Auto"), false)
        eq("auto: several lines pasted", types("Dear Sir,`nThanks", "Auto"), false)
        eq("auto: CR pasted", types("a`rb", "Auto"), false)
        eq("auto: tab pasted", types("a`tb", "Auto"), false)
        eq("auto: emoji pasted", types("ok " Chr(0x1F600), "Auto"), false)
        eq("clipboard: never typed", types("hi", "Clipboard"), false)
        eq("type: always typed", types("Dear Sir,`nThanks", "Type"), true)
        eq("unknown mode: clipboard", types("hi", "Other"), false)

        ; 真的输入到一个编辑框: 文字出来了, 剪贴板没动
        snippets := AppSettings.Feature("Snippets")
        savedMode := snippets["PasteMode"], savedClip := ClipboardAll()
        g := Gui("+AlwaysOnTop", "ALTRun paste test")
        box := g.Add("Edit", "w300 h40")
        g.Show()
        try {
            WinActivate("ahk_id " g.Hwnd)
            if WinWaitActive("ahk_id " g.Hwnd, , 2) {
                ControlFocus(box)
                A_Clipboard := "keep me"
                snippets["PasteMode"] := "Auto"
                eq("auto: pasted", ActionCatalog.PasteText("typed text", false), true)
                Sleep(100)
                eq("auto: text typed", box.Value, "typed text")
                eq("auto: clipboard untouched", A_Clipboard, "keep me")
            } else {
                FileAppend("NOTE SnippetPasteMode: test window not active, typing test skipped`n", "*")
            }
        } finally {
            snippets["PasteMode"] := savedMode
            g.Destroy()
            A_Clipboard := savedClip
        }
    }

    static SnippetExpander() {
        eq := (n, a, e) => TestRunner.Equal("SnippetExpander." n, a, e)
        abbr := (snippet, prefix) => TestUtil.Join(SnippetExpander.Abbreviations(snippet, prefix))
        eq("prefix", abbr(Map("Keyword", "sig", "Text", "x"), ";"), ";sig")
        eq("no keyword", abbr(Map("Keyword", "", "Text", "x"), ";"), "")
        eq("disabled", abbr(Map("Keyword", "sig", "Text", "x", "AutoExpand", 0), ";"), "")
        eq("space", abbr(Map("Keyword", "a b", "Text", "x"), ";"), "")
        eq("no prefix", abbr(Map("Keyword", "sig", "Text", "x"), ""), "sig")
        eq("several keywords", abbr(Map("Keyword", " sig, ㄴ ,，sg ", "Text", "x"), ";"), ";sig,;ㄴ,;sg")
        eq("several: bad one skipped", abbr(Map("Keyword", "a b, ab", "Text", "x"), ";"), ";ab")
        options := AppSettings.Feature("Snippets"), saved := options["ExpandExclude"]
        options["ExpandExclude"] := ["ahk_exe mstsc.exe", "ahk_class KeePass"]
        SnippetExpander.Refresh()
        eq("exclude list", SnippetExpander._exclude.Length "|" SnippetExpander._exclude[2], "2|ahk_class KeePass")
        eq("default exclude", AppSettings.Defaults()["Features"]["Snippets"]["ExpandExclude"].Length > 0, true)
        options["ExpandExclude"] := saved
        SnippetExpander.Refresh()
    }

    ; 对话框的文件夹菜单: 最近的文件夹 (Windows 的 "最近使用的项目")、菜单文字
    static FolderMenu() {
        eq := (n, a, e) => TestRunner.Equal("FolderMenu." n, a, e)
        root := A_Temp "\ALTRun-test-recent"
        try DirDelete(root, true)
        DirCreate(root "\Recent"), DirCreate(root "\Projects\Tower"), DirCreate(root "\Docs")
        longPath := Buffer(2048)                                            ; A_Temp 可能是 8.3 短路径 (C:\Users\RUNNER~1), 快捷方式读回来的是长路径
        if DllCall("GetLongPathNameW", "Str", root, "Ptr", longPath, "UInt", 1024)
            root := StrGet(longPath)
        FileAppend("x", root "\Docs\report.pdf")
        FileCreateShortcut(root "\Projects\Tower", root "\Recent\Tower.lnk")
        FileSetTime("20260101000000", root "\Recent\Tower.lnk")
        FileCreateShortcut(root "\Docs\report.pdf", root "\Recent\report.pdf.lnk")    ; 文件: 取所在的文件夹
        FileSetTime("20260201000000", root "\Recent\report.pdf.lnk")
        FileCreateShortcut(root "\Docs", root "\Recent\Docs.lnk")                     ; 和上一条同一个文件夹: 只出现一次
        FileSetTime("20260115000000", root "\Recent\Docs.lnk")
        FileCreateShortcut(root "\Gone", root "\Recent\Gone.lnk")                     ; 已经不存在
        FileSetTime("20260301000000", root "\Recent\Gone.lnk")
        FileAppend("not a shortcut", root "\Recent\Broken.lnk")                     ; 损坏的快捷方式: 跳过, 不报错
        FileSetTime("20260401000000", root "\Recent\Broken.lnk")
        folders := QuickSwitch.RecentFolders(10, root "\Recent")
        eq("count", folders.Length, 2)
        eq("newest first", folders.Length ? folders[1] : "", root "\Docs")
        eq("second", folders.Length > 1 ? folders[2] : "", root "\Projects\Tower")
        eq("limit", QuickSwitch.RecentFolders(1, root "\Recent").Length, 1)
        eq("zero", QuickSwitch.RecentFolders(0, root "\Recent").Length, 0)
        eq("default hotkey", AppSettings.Defaults()["Extensions"]["QuickSwitch"]["MenuHotkey"], "^+g")
        eq("panel on by default", AppSettings.Defaults()["Extensions"]["QuickSwitch"]["ShowPanel"], 1)
        area := {Left: 0, Top: 0, Right: 1920, Bottom: 1040}
        pos := QuickSwitch.PanelPosition(400, 200, 800, 500, 800, 300, area)
        eq("panel below", pos.X "," pos.Y, "400,700")
        pos := QuickSwitch.PanelPosition(100, 600, 800, 500, 800, 300, area)
        eq("panel above", pos.X "," pos.Y, "100,300")
        pos := QuickSwitch.PanelPosition(1500, 100, 800, 900, 800, 300, area)
        eq("panel kept on screen", pos.X "," pos.Y, "1120,740")
        eq("folder name", QuickSwitch.FolderName("D:\Projects\Tower A\"), "Tower A")
        eq("drive name", QuickSwitch.FolderName("D:\"), "D:\")
        ; 文件对话框: #32770 要有资源管理器的文件视图; "系统信息" (msinfo32) 的主窗口只有普通列表, 不算
        fileDialog := ["DUIViewWndClassName1", "DirectUIHWND1", "SHELLDLL_DefView1", "DirectUIHWND2", "ComboBoxEx321", "ComboBox1", "Edit1", "Button1", "Button2", "ToolbarWindow323"]
        oldDialog := ["ToolbarWindow321", "ComboBox1", "SHELLDLL_DefView1", "SysListView321", "Edit1", "Button1", "Button2"]
        msinfo := ["SysTreeView321", "SysListView321", "Static1", "Edit1", "Button1", "Button2", "Button3", "Button4"]
        eq("file dialog", QuickSwitch.LooksLikeFileDialog("#32770", fileDialog), true)
        eq("old-style file dialog", QuickSwitch.LooksLikeFileDialog("#32770", oldDialog), true)
        eq("msinfo32 main window", QuickSwitch.LooksLikeFileDialog("#32770", msinfo), false)
        eq("other app with a plain list", QuickSwitch.LooksLikeFileDialog("TMyOpenDialog", msinfo), true)
        eq("no controls: title", QuickSwitch.LooksLikeFileDialog("Qt5QWindowIcon", [], "Open File") " " QuickSwitch.LooksLikeFileDialog("Qt5QWindowIcon", [], "Writer"), "1 0")
        ; "Browse for Folder" (folder tree only): its own kind, never a file dialog (pressing Enter there means OK)
        newStyle := ["Static1", "SHBrowseForFolder ShellNameSpace Control1", "SysTreeView321", "Static2", "Edit1", "Button1", "Button2", "Button3"]
        classic := ["SysTreeView321", "Static1", "Static2", "Static3", "Edit1", "Button1", "Button2", "Button3"]
        eq("browse dialog: new style", QuickSwitch.LooksLikeBrowseDialog("#32770", newStyle), true)
        eq("browse dialog: classic tree ID", QuickSwitch.LooksLikeBrowseDialog("#32770", classic, 0x3741), true)
        eq("other dialog with a tree", QuickSwitch.LooksLikeBrowseDialog("#32770", classic, 0x3F2), false)
        eq("browse dialog: other class", QuickSwitch.LooksLikeBrowseDialog("TMyDialog", newStyle), false)
        eq("file dialog is not a browse dialog", QuickSwitch.LooksLikeBrowseDialog("#32770", fileDialog), false)
        eq("browse dialog is not a file dialog", QuickSwitch.LooksLikeFileDialog("#32770", newStyle) QuickSwitch.LooksLikeFileDialog("#32770", classic), "00")
        eq("browse dialogs on by default", AppSettings.Defaults()["Extensions"]["QuickSwitch"]["BrowseDialogs"], 1)
        ; How the folder goes in: the file name box when there is one; Alt+D (address bar) only in system dialogs;
        ; dialogs that programs draw themselves get the path typed into the focused box, no Alt+D
        eq("jump: file name box", QuickSwitch.JumpMethod("#32770", true) " " QuickSwitch.JumpMethod("TMyOpenDialog", true), "edit edit")
        eq("jump: system dialog without the box", QuickSwitch.JumpMethod("#32770", false), "address")
        eq("jump: drawn dialogs", QuickSwitch.JumpMethod("Qt5QWindowIcon", false) " " QuickSwitch.JumpMethod("SunAwtDialog", false) " " QuickSwitch.JumpMethod("gdkWindowToplevel", false), "type type type")
        ; Directory Opus: "dopusrt /info <file>,paths" output -> tabs; the jump uses the source tab of the lister
        xml := '<?xml version="1.0"?>`n<results command="paths" result="1">`n'
            . '  <path active_lister="1" active_tab="1" tab_state="2" side="1" lister="0x70e2a" tab="0x50e1a" display_path="C:\Program Files">C:\Program Files</path>`n'
            . '  <path active_lister="1" active_tab="2" tab_state="1" side="2" lister="0x70e2a" tab="0xf0f30">D:\R&amp;D &lt;2026&gt;</path>`n'
            . '  <path active_lister="1" active_tab="0" tab_state="0" side="1" lister="0x70e2a" tab="0x1">E:\Background tab</path>`n'
            . '  <path active_lister="0" active_tab="1" tab_state="1" side="1" lister="0x9a0b4" tab="0x2">F:\Second lister</path>`n</results>'
        tabs := QuickSwitch.ParseDOpusPaths(xml)
        eq("opus tabs", tabs.Length, 4)
        eq("opus tab fields", tabs[1].Path "|" tabs[1].Lister "|" tabs[1].State "|" tabs[1].ActiveLister, "C:\Program Files|" 0x70e2a "|2|1")
        eq("opus entities", tabs[2].Path, "D:\R&D <2026>")
        eq("opus: active lister's source tab", QuickSwitch.PickDOpusFolder(tabs), "D:\R&D <2026>")
        eq("opus: a given lister", QuickSwitch.PickDOpusFolder(tabs, 0x9a0b4), "F:\Second lister")
        eq("opus: unknown lister -> active", QuickSwitch.PickDOpusFolder(tabs, 0x1234), "D:\R&D <2026>")
        eq("opus: nothing", QuickSwitch.PickDOpusFolder(QuickSwitch.ParseDOpusPaths("")), "")
        eq("opus not running: no tabs, no dopusrt", WinExist("ahk_class dopus.lister") ? "skip" : QuickSwitch.DOpusTabs().Length, WinExist("ahk_class dopus.lister") ? "skip" : 0)
        eq("file manager windows", QuickSwitch.IsFileManager(0), false)
        ; A real one: AutoHotkey's DirSelect is SHBrowseForFolder (option 4 = classic style). Detected, the folder is
        ; selected in its tree, and OK returns it
        target := A_WinDir "\System32"
        for options in [0, 4] {
            seen := {Detected: "", Selected: ""}, tries := 0
            probe() {
                if !(dlg := SearchWindow.InSystemDialog()) {
                    if (++tries < 50)
                        SetTimer(probe, -100)
                    return
                }
                seen.Detected := QuickSwitch.IsBrowseWindow(dlg)
                seen.Selected := QuickSwitch.SelectInBrowseDialog(target "\", dlg)
                Sleep(300)
                PostMessage(0x111, 1, 0, , "ahk_id " dlg)                   ; WM_COMMAND IDOK
                SetTimer(() => WinExist("ahk_id " dlg) && WinClose("ahk_id " dlg), -3000)   ; never left open
            }
            SetTimer(probe, -300)
            picked := DirSelect("C:\", options)
            eq("DirSelect " options ": detected", seen.Detected, true)
            eq("DirSelect " options ": selected", seen.Selected, true)
            eq("DirSelect " options ": OK returns it", StrLower(RTrim(picked, "\")), StrLower(target))
        }
        base := [{Path: "D:\Projects\Tower A", Tag: "TC"}, {Path: "C:\Docs", Tag: "Recent"}]
        filtered := QuickSwitch.FilterFolders(base, "tower", ["D:\Projects\Tower A", "E:\Archive\Tower B"])
        eq("filter count", filtered.Length, 2)
        eq("filter keeps listed first", filtered[1].Tag, "TC")
        eq("filter adds found", filtered[2].Path "|" filtered[2].Tag, "E:\Archive\Tower B|" I18n.T("QuickSwitch.TagSearch"))
        eq("filter words", QuickSwitch.FilterFolders(base, "proj tow", []).Length, 1)
        eq("filter empty", QuickSwitch.FilterFolders(base, "", []).Length, 2)
        withFile := QuickSwitch.FilterFolders(base, "tower", [{Path: "E:\Tower", IsFolder: true}, {Path: "E:\Tower\plan.dwg", IsFolder: false}])
        eq("file entry", withFile[3].IsFile "|" withFile[3].Tag, "1|" I18n.T("QuickSwitch.TagFile"))
        eq("folder entry", withFile[2].IsFile, false)
        eq("restore file name", QuickSwitch.ShouldRestoreName("Report 2026.docx"), true)
        eq("no restore empty", QuickSwitch.ShouldRestoreName("  "), false)
        eq("no restore filter", QuickSwitch.ShouldRestoreName("*.txt"), false)
        eq("no restore path", QuickSwitch.ShouldRestoreName("C:\Docs\a.txt"), false)
        ; TC 的路径用 WM_COPYDATA 问 (另一个进程模拟 TC 的回复), 不碰剪贴板
        A_Clipboard := "clipboard before"
        Run('"' A_AhkPath '" "' A_ScriptDir '\Fixtures\FakeTotalCmd.ahk"', , , &fakePid)
        fakeHwnd := WinWait("ALTRun Fake Total Commander ahk_pid " fakePid, , 5)
        eq("tc source panel", QuickSwitch.TotalCmdFolder(fakeHwnd), "C:\Work\Current")
        eq("tc other panel", QuickSwitch.TotalCmdFolder(fakeHwnd, true), "\\server\share\Other")
        eq("tc active side", QuickSwitch.TotalCmdAsk(fakeHwnd, "A"), "L")
        eq("tc not a folder", QuickSwitch.TotalCmdAsk(fakeHwnd, "LP"), "ftp://example.com/pub/")
        eq("tc clipboard untouched", A_Clipboard, "clipboard before")
        ProcessClose(fakePid), ProcessWaitClose(fakePid, 5)
        eq("tc closed", QuickSwitch.TotalCmdFolder(fakeHwnd), "")
        closed := Gui(), closedHwnd := closed.Hwnd
        closed.Destroy()
        threw := false
        try QuickSwitch.SetDialogPath(A_Temp "\", closedHwnd)               ; 对话框已经关掉: 什么也不做 (不往别的窗口里打字)
        catch
            threw := true
        eq("closed dialog ignored", threw, false)
        QuickSwitch._BuildPanel(A_ScriptHwnd)                              ; 真的建一次面板 (Gui 选项写错时这里就会报错)
        TestRunner.True("FolderMenu.panel built", IsObject(QuickSwitch._panel) && IsObject(QuickSwitch._panelList))
        QuickSwitch._ShowFolders([])                                       ; nothing open (no TC / Explorer, no recent folders)
        eq("empty panel: no horizontal scrollbar", ControlGetStyle(QuickSwitch._panelList) & 0x100000, 0)   ; WS_HSCROLL
        eq("empty panel: columns fit", QuickSwitch.ColumnsWidth(QuickSwitch._panelList) <= QuickSwitch.ClientWidth(QuickSwitch._panelList), true)
        QuickSwitch._ShowFolders(base)
        eq("panel rows", QuickSwitch._panelList.GetCount(), 2)
        eq("panel: no horizontal scrollbar", ControlGetStyle(QuickSwitch._panelList) & 0x100000, 0)
        eq("panel name column", QuickSwitch._panelList.GetText(1, 1), "Tower A")
        QuickSwitch._panelSearch.Value := "docs"
        QuickSwitch._panelBase := base
        QuickSwitch._SearchPanel()
        eq("panel search filters", QuickSwitch._panelList.GetText(1, 2), "C:\Docs")
        eq("resize before", QuickSwitch._panelWidth, QuickSwitch.PanelWidthFor(QuickSwitch._panelDialogW ? QuickSwitch._panelDialogW : 600))
        QuickSwitch._ResizePanel(900)                                      ; 对话框变宽: 面板跟着变宽
        QuickSwitch._panelList.GetPos(, , &listW)
        eq("resize width", QuickSwitch._panelWidth "|" listW, "900|900")
        eq("key other window", QuickSwitch._OnKeyDown(0x28, 0, 0x100, A_ScriptHwnd), "")
        eq("key down handled", QuickSwitch._OnKeyDown(0x28, 0, 0x100, QuickSwitch._panelSearch.Hwnd), 0)
        eq("ime key passes", QuickSwitch._OnKeyDown(0xE5, 0, 0x100, QuickSwitch._panelSearch.Hwnd), "")   ; 输入法正在输入: 不拦截
        searchHwnd := QuickSwitch._searchHwnd
        QuickSwitch.HidePanel()
        eq("panel hidden", QuickSwitch._panel, "")
        eq("controls released", QuickSwitch._panelSearch "|" QuickSwitch._searchHwnd, "|0")
        eq("key after hide", QuickSwitch._OnKeyDown(0x0D, 0, 0x100, searchHwnd), "")   ; 面板关掉之后: 不报错, 也不处理
        eq("key no panel", QuickSwitch._OnKeyDown(0x0D, 0, 0x100, 0), "")
        eq("width min", QuickSwitch.PanelWidthFor(300), 420)
        eq("width max", QuickSwitch.PanelWidthFor(1600), 1000)
        eq("dark color", QuickSwitch.IsDarkColor("1E1E1E"), true)
        eq("light color", QuickSwitch.IsDarkColor("FAFAFA"), false)
        eq("bad color", QuickSwitch.IsDarkColor("abc"), false)
        savedTheme := ThemeManager.Current
        ThemeManager.Current := Map("Background", "202020", "Title", "EEEEEE", "Separator", "333333")
        colors := QuickSwitch.PanelColors()
        eq("theme colors", colors.Background "|" colors.Text "|" colors.Dark, "202020|EEEEEE|1")
        ThemeManager.Current := Map()
        eq("no theme", QuickSwitch.PanelColors().Background, "FFFFFF")
        ThemeManager.Current := savedTheme
        eq("panel search default", AppSettings.Defaults()["Extensions"]["QuickSwitch"]["PanelSearch"], "folders")
        try DirDelete(root, true)
    }

    ; 束线型计算和原来的 SPF2M.EXE 逐个比对: Tests\Fixtures\SPF2M-Reference.json 是在 DOSBox 里运行
    ; SPF2M 得到的结果 (4 种线型 x 6 种钢绞线, 上升 / 下降, 非整米跨度, 自定义半径 / 反弯点 /
    ; 支架间距, 以及 SPF2M 报错或崩溃的情况); 生成方法见 Tests\Tools\SPF2M
    static TendonProfileVsSpf2m() {
        data := JSON.Parse(FileRead(A_ScriptDir "\Fixtures\SPF2M-Reference.json", "UTF-8"))
        checked := 0, failed := 0
        for ref in data["Cases"] {
            input := Map()
            for key in ["Profile", "Tendon", "Start", "End", "Distance", "Radius", "Contraflexure", "Intervals"]
                if ref.Has(key)
                    input[key] := ref[key]
            result := TendonProfile.Calc(input)
            name := "TendonProfile." ref["Profile"] "/" ref["Tendon"] " " ref["Start"] "->" ref["End"] " L" ref["Distance"]
                . (ref.Has("Radius") ? " R" ref["Radius"] : "") . (ref.Has("Contraflexure") ? " C" ref["Contraflexure"] : "")
            if ref.Has("Expect") {
                TestRunner.True(name " error", result.Error != "")
                continue
            }
            if (result.Error != "") {
                TestRunner.Fail(name, "unexpected error: " result.Error)
                continue
            }
            got := "", want := ""
            for row in result.Rows
                got .= row[1] "," row[2] "," row[3] "," row[4] "," row[5] ";"
            for row in ref["Rows"]
                want .= row[1] "," row[2] "," row[3] "," row[4] "," row[5] ";"
            if (got != want) {
                failed += 1
                TestRunner.Equal(name " rows", got, want)
            }
            checked += ref["Rows"].Length
            TestRunner.Equal(name " contraflexure", Integer(result.Contraflexure), ref["ShownContraflexure"])
            if ref.Has("ShownRadius")
                TestRunner.Equal(name " radius", Integer(result.Radius), ref["ShownRadius"])
        }
        TestRunner.True("TendonProfile.rows checked (" checked ")", checked > 1000 && !failed)
    }

    ; 输入检查、支架间距、SPF2M 没有处理好的情况 (崩溃 / 不显示任何东西)
    ; PT Tools 窗口: 结果框不停留 Tab 焦点 (Area Expression 之后直接到 Rebar Diameter), 单位写成 mm², 点进输入框时全选
    static PTToolsWindowUi() {
        eq := (n, a, e) => TestRunner.Equal("PTToolsWindowUi." n, a, e)
        saved := [PTToolsWindow.Settings["WinLeft"], PTToolsWindow.Settings["WinTop"]]
        PTToolsWindow.Settings["WinLeft"] := -3000, PTToolsWindow.Settings["WinTop"] := -3000
        PTToolsWindow.Show()
        g := PTToolsWindow.G
        tabbed := "", oldUnit := ""
        for ctrl in g {
            style := ControlGetStyle(ctrl)
            if (ctrl.Type = "Edit" && (style & 0x800) && (style & 0x10000))     ; ES_READONLY 还有 WS_TABSTOP
                tabbed .= ctrl.Name " "
            if InStr(ctrl.Text, "mm2")
                oldUnit .= ctrl.Text " | "
        }
        eq("read-only fields skip tab", tabbed, "")
        eq("units", oldUnit, "")
        next := DllCall("GetNextDlgTabItem", "Ptr", g.Hwnd, "Ptr", g["RequiredAreaExpression"].Hwnd, "Int", 0, "Ptr")
        eq("expression -> diameter", next = g["RequiredBarDiameter"].Hwnd, true)
        span := g["RebarSpanWidth"]
        SendMessage(0xB1, 1, 1, span)                                       ; 光标放在中间
        span.Focus()
        Sleep(100)                                                          ; 处理 PostMessage 的全选
        start := 0, finish := 0
        DllCall("SendMessage", "Ptr", span.Hwnd, "UInt", 0xB0, "UInt*", &start, "UInt*", &finish)   ; EM_GETSEL
        eq("select all on focus", start "-" finish, "0-" StrLen(span.Value))
        eq("right margin", (SendMessage(0xD4, 0, 0, span) >> 16) >= 4, true)   ; EM_GETMARGINS: 光标不紧贴数字
        g["RebarSpanWidth"].Value := 5000                                   ; 第一组的 Span Width -> BRC 的 Span Width -> 钢筋网面积一起重算
        PTToolsWindow.OnRebarFieldChanged()
        eq("brc span follows", g["BrcSpanWidth"].Value, "5000")
        eq("top mesh area recalculated", g["BrcTopArea"].Value, PTToolsWindow.Fmt2(5 * PTToolsWindow.CalculateMeshArea(g["BrcTopMeshMark"].Value)))
        eq("bottom mesh area recalculated", g["BrcBotArea"].Value, PTToolsWindow.Fmt2(5 * PTToolsWindow.CalculateMeshArea(g["BrcBotMeshMark"].Value)))
        safety := g["SafetyFactorEnabled"]
        eq("safety factor checkbox skips tab", (ControlGetStyle(safety) & 0x10000) != 0, false)
        ; 钢筋直径下拉框里的数字右对齐: 第一个字的位置 (EM_POSFROMCHAR) 在输入框的右半边 (Wine 不支持建好后改对齐, 只在 Windows 上看位置)
        diaEdit := DllCall("FindWindowEx", "Ptr", g["RebarBarDiameter"].Hwnd, "Ptr", 0, "Str", "Edit", "Ptr", 0, "Ptr")
        eq("diameter diaEdit has ES_RIGHT", (ControlGetStyle(diaEdit) & 0x2) != 0, true)
        if !DllCall("GetProcAddress", "Ptr", DllCall("GetModuleHandle", "Str", "ntdll", "Ptr"), "AStr", "wine_get_version", "Ptr") {
            WinGetClientPos(, , &ew, , diaEdit)
            firstX := SendMessage(0xD6, 0, 0, diaEdit) & 0xFFFF                     ; EM_POSFROMCHAR: 第 0 个字的 x
            TestRunner.True("PTToolsWindowUi.diameter text on the right (first char at " firstX " of " ew ")", firstX > ew // 2)
        }
        g.Destroy(), PTToolsWindow.G := ""
        PTToolsWindow.Settings["WinLeft"] := saved[1], PTToolsWindow.Settings["WinTop"] := saved[2]
    }

    static TendonProfileInputs() {
        eq := (n, a, e) => TestRunner.Equal("TendonProfileInputs." n, a, e)
        profileOf := (profile, start, finish, l, extra := "") => TendonProfile.Calc(TendonProfileInputsMap(profile, start, finish, l, extra))
        TendonProfileInputsMap(profile, start, finish, l, extra) {
            input := Map("Profile", profile, "Tendon", 3, "Start", start, "End", finish, "Distance", l)
            if IsObject(extra)
                for key, value in extra
                    input[key] := value
            return input
        }
        intervals := (l) => TendonProfileJoin(TendonProfile.DefaultIntervals(l))
        TendonProfileJoin(list) {
            text := ""
            for v in list
                text .= (text = "" ? "" : ",") v
            return text
        }
        eq("intervals whole metres", intervals(8000), "1000,1000,1000,1000,1000,1000,1000,1000")
        eq("intervals single split", intervals(9750), "750,1000,1000,1000,1000,1000,1000,1000,1000,1000")
        eq("intervals double split", intervals(7350), "650,700,1000,1000,1000,1000,1000,1000")
        eq("intervals double split 2", intervals(8420), "720,700,1000,1000,1000,1000,1000,1000,1000")
        eq("intervals short span", intervals(450), "450")
        eq("intervals 1300", intervals(1300), "600,700")
        eq("parse intervals", TendonProfileJoin(TendonProfile.ParseIntervals("500, 1500 800;550")), "500,1500,800,550")
        eq("parse empty = auto", TendonProfile.ParseIntervals("  "), "")
        eq("intervals must add up", InStr(profileOf(1, 600, 100, 8000, Map("Intervals", [1000, 1000])).Error, "add up") > 0, true)
        eq("interval not a number", profileOf(1, 600, 100, 8000, Map("Intervals", TendonProfile.ParseIntervals("1000 abc"))).Error != "", true)
        eq("missing distance", profileOf(1, 600, 100, "").Error != "", true)
        eq("missing level", profileOf(1, "", 100, 8000).Error != "", true)
        ; SPF2M 崩溃的情况: 这里给出错误说明
        eq("psp not achievable", InStr(profileOf(2, 900, 100, 3000, Map("Tendon", 6)).Error, "NOT ACHIEVABLE") > 0, true)
        ; SPF2M 什么都不显示的情况: 反弯点超过半跨
        eq("contraflexure beyond half span", InStr(profileOf(1, 600, 100, 9000, Map("Contraflexure", 5000)).Error, "half") > 0, true)
        eq("contraflexure with equal levels", profileOf(1, 300, 300, 6000, Map("Contraflexure", 500)).Error != "", true)
        result := profileOf(1, 600, 100, 9000, Map("Contraflexure", 1200))
        eq("contraflexure -> radius", Integer(result.Radius) "|" result.RadiusChanged, "10800|1")
        result := profileOf(1, 600, 100, 8000, Map("Contraflexure", 525))
        eq("contraflexure equal to default keeps radius", result.Radius "|" result.RadiusChanged, "4200|0")
        eq("default radius 22s", TendonProfile.DefaultRadius(5), 5700)
        eq("default duct dia", TendonProfile.DefaultDuctDiameter(1) "|" TendonProfile.DefaultDuctDiameter(3) "|" TendonProfile.DefaultDuctDiameter(6), "25|90|130")
        eq("default tendon 12S", PTToolsWindow.Defaults["TendonType"], 3)
    }

    ; 一键更新: 解析 GitHub 的 Release、SHA256、替换文件 (用临时文件夹里的假程序, 不下载)
    static SelfUpdate() {
        eq := (n, a, e) => TestRunner.Equal("SelfUpdate." n, a, e)
        ok := (n, c) => TestRunner.True("SelfUpdate." n, c)
        ; GitHub's hourly limit (per IP address) gets a message people understand; other errors stay as they are
        reason := (message) => StrSplit(UpdateChecker.FailureText(message), "`n")[1]
        for pair in [["HTTP 403 rate limit exceeded", "Update.RateLimited"], ["HTTP 429 Too Many Requests", "Update.RateLimited"]
                   , ["Network error 12007", "Update.Offline"], ["Timed out after 30 s", "Update.Offline"], ["HTTP 502 Bad Gateway", "Update.ServerError"]
                   , ["HTTP 404 Not Found", "Update.NotFound"], ["SHA256 does not match (expected a, got b).", "Update.Corrupt"]
                   , ["Could not extract the downloaded package (exit code 1).", "Update.BadPackage"], ["Cannot find 'tag_name' in the GitHub response.", "Update.BadResponse"]
                   , ["(5) Access is denied.", "Update.Unknown"]]
            eq("failure " pair[1] ": " pair[2], reason(pair[1]), I18n.T(pair[2]))
        eq("failure keeps the raw error as details", SubStr(UpdateChecker.FailureText("HTTP 404 Not Found"), -StrLen(I18n.T("Update.Detail", "HTTP 404 Not Found"))), I18n.T("Update.Detail", "HTTP 404 Not Found"))
        digest := "3779b0aedbd58a7bf235eab4c0cc52d3b4c9f3188cb98be38319837605cef017"
        response := '{"tag_name": "v2026.10.01", "html_url": "https://github.com/zhugecaomao/ALTRun/releases/tag/2026.10.01", "assets": ['
            . '{"name": "notes.txt", "browser_download_url": "https://x/notes.txt", "digest": null},'
            . '{"name": "ALTRun_v2026.10.01.zip", "browser_download_url": "https://github.com/zhugecaomao/ALTRun/releases/download/2026.10.01/ALTRun_v2026.10.01.zip", "digest": "sha256:' StrUpper(digest) '"}]}'
        release := UpdateChecker.ParseRelease(response)
        eq("version", release.Version, "2026.10.01")
        eq("page", release.Page, "https://github.com/zhugecaomao/ALTRun/releases/tag/2026.10.01")
        eq("zip", release.ZipUrl, "https://github.com/zhugecaomao/ALTRun/releases/download/2026.10.01/ALTRun_v2026.10.01.zip")
        eq("sha256", release.Sha256, digest)
        release := UpdateChecker.ParseRelease('{"tag_name": "2026.10.01", "assets": [{"name": "ALTRun_v2026.10.01.zip", "browser_download_url": "https://x/a.zip", "digest": null}]}')
        eq("no digest", release.Sha256 "|" release.Page, "|" UpdateChecker.ReleasePage)
        ok("no digest -> no install", !UpdateChecker.CanInstall(release))
        ; Only addresses in this repository are opened / downloaded (Run would start anything)
        other := UpdateChecker.ParseRelease('{"tag_name": "2026.10.01", "html_url": "calc.exe", "assets": [{"name": "ALTRun_v2026.10.01.zip"'
            . ', "browser_download_url": "https://evil.example/ALTRun_v2026.10.01.zip", "digest": "sha256:' digest '"}]}')
        eq("foreign page ignored", other.Page, UpdateChecker.ReleasePage)
        eq("foreign zip ignored", other.ZipUrl, "")
        ok("foreign zip -> no install", !UpdateChecker.CanInstall(other))
        eq("in repo", UpdateChecker.InRepo("https://github.com/zhugecaomao/ALTRun/releases/tag/2026.10.01"), true)
        eq("other repo", UpdateChecker.InRepo("https://github.com/zhugecaomao/ALTRunX/releases"), false)
        eq("tricks rejected", UpdateChecker.InRepo('https://github.com/zhugecaomao/ALTRun/../../x') || UpdateChecker.InRepo('https://github.com/zhugecaomao/ALTRun/" & calc'), false)
        eq("download folder", UpdateChecker.InRepo("https://github.com/zhugecaomao/ALTRun/releases/tag/x.zip", "releases/download/"), false)
        failed := false
        try UpdateChecker.ParseRelease('{"message": "API rate limit exceeded"}')
        catch
            failed := true
        ok("no tag throws", failed)

        root := A_Temp "\ALTRun-selfupdate-test"
        try DirDelete(root, true)
        src := root "\new", dest := root "\app"
        write := (path, text) => (DirCreate(RegExReplace(path, "\\[^\\]+$")), FileAppend(text, path, "UTF-8-RAW"))
        read := (path) => FileExist(path) ? FileRead(path, "UTF-8") : "<missing>"
        write(root "\abc.txt", "abc")
        eq("sha256 file", UpdateChecker.Sha256File(root "\abc.txt"), "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad")
        ok("writable", UpdateChecker.IsWritable(root))
        tar := A_WinDir "\System32\tar.exe"                                ; 解压 (Windows 10 起自带 tar.exe; Wine 里没有就跳过)
        if FileExist(tar) {
            write(root "\pkg\ALTRun.exe", "exe"), write(root "\pkg\Resources\Lang\zh-CN.json", "{}")
            RunWait('"' tar '" -a -cf "' root '\pkg.zip" -C "' root '\pkg" ALTRun.exe Resources', , "Hide")
            UpdateChecker.Extract(root "\pkg.zip", root "\out")
            eq("extract", read(root "\out\ALTRun.exe") "|" read(root "\out\Resources\Lang\zh-CN.json"), "exe|{}")
        }

        write(src "\ALTRun.exe", "new exe")
        write(src "\Resources\Kanji.txt", "new kanji")
        write(src "\Resources\Themes\Dark.json", "{}")
        write(src "\Resources\SDL.dll", "still shipped")
        write(src "\README.md", "readme")
        write(dest "\ALTRun.exe", "old exe")
        write(dest "\ALTRun.exe.old", "stale")
        write(dest "\Data\ALTRun.json", "settings")
        write(dest "\Data\Knowledge.json", "learned")
        write(dest "\Resources\Kanji.txt", "old kanji")
        write(dest "\Resources\DOSBox.exe", "obsolete")
        write(dest "\Resources\Mine.txt", "user file")
        UpdateChecker.Apply(src, dest, "ALTRun.exe")
        eq("exe replaced", read(dest "\ALTRun.exe"), "new exe")
        eq("old exe kept as .old", read(dest "\ALTRun.exe.old"), "old exe")
        eq("resources updated", read(dest "\Resources\Kanji.txt") "|" read(dest "\Resources\Themes\Dark.json"), "new kanji|{}")
        eq("settings untouched", read(dest "\Data\ALTRun.json") "|" read(dest "\Data\Knowledge.json"), "settings|learned")
        eq("obsolete removed, user file kept", read(dest "\Resources\DOSBox.exe") "|" read(dest "\Resources\Mine.txt"), "<missing>|user file")
        eq("obsolete but still in the package: kept", read(dest "\Resources\SDL.dll"), "still shipped")
        eq("top-level file copied", read(dest "\README.md"), "readme")
        extra := ""
        Loop Files dest "\*", "D"
            if !(A_LoopFileName = "Resources" || A_LoopFileName = "Data")
                extra .= A_LoopFileName " "
        eq("no stray folders", extra, "")

        write(dest "\Launcher.exe", "renamed old")                          ; 用户把 ALTRun.exe 改过名
        FileDelete(dest "\ALTRun.exe")
        UpdateChecker.Apply(src, dest, "Launcher.exe")
        eq("renamed exe replaced", read(dest "\Launcher.exe") "|" read(dest "\Launcher.exe.old"), "new exe|renamed old")
        eq("no extra ALTRun.exe", read(dest "\ALTRun.exe"), "<missing>")

        FileDelete(src "\ALTRun.exe")
        FileDelete(dest "\Launcher.exe.old")
        failed := false
        try UpdateChecker.Apply(src, dest, "Launcher.exe")
        catch
            failed := true
        ok("package without exe throws", failed)
        eq("package without exe changes nothing", read(dest "\Launcher.exe") "|" read(dest "\Launcher.exe.old"), "new exe|<missing>")

        ; 复制到一半失败 (目标文件被别的程序独占打开): 已经覆盖的改回去, 新增的文件和文件夹删掉, exe 不动
        try DirDelete(root, true)
        write(src "\ALTRun.exe", "new exe")
        write(src "\README.md", "new readme")
        write(src "\Resources\Kanji.txt", "new kanji")
        write(src "\Resources\New.txt", "new file")
        write(src "\Resources\Fonts\a.ttf", "font")
        write(dest "\ALTRun.exe", "old exe")
        write(dest "\README.md", "old readme")
        write(dest "\Resources\Kanji.txt", "old kanji")
        locked := FileOpen(dest "\Resources\Kanji.txt", "rw -rwd")
        failed := false
        try UpdateChecker.Apply(src, dest, "ALTRun.exe")
        catch
            failed := true
        locked.Close()
        ok("copy failure throws", failed)
        eq("copy failure: files restored", read(dest "\README.md") "|" read(dest "\Resources\Kanji.txt") "|" read(dest "\ALTRun.exe"), "old readme|old kanji|old exe")
        eq("copy failure: new files removed", read(dest "\Resources\New.txt") "|" read(dest "\Resources\Fonts\a.ttf") "|" (InStr(FileExist(dest "\Resources\Fonts"), "D") ? "dir" : "no dir"), "<missing>|<missing>|no dir")
        eq("copy failure: no .old, no backup left", read(dest "\ALTRun.exe.old") "|" (FileExist(src ".backup") ? "backup" : "none"), "<missing>|none")
        try DirDelete(root, true)
    }

    ; 后台检查更新: 每天一次、跳过的版本、搜索窗口里的更新提示 (不联网: 直接设置 Pending)
    static UpdateNotice() {
        eq := (n, a, e) => TestRunner.Equal("UpdateNotice." n, a, e)
        ok := (n, c) => TestRunner.True("UpdateNotice." n, c)
        ok("due: never checked", UpdateChecker.IsDue("", "20260928120000"))
        ok("not due: 5 h", !UpdateChecker.IsDue("20260928070000", "20260928120000"))
        ok("due: 6 h", UpdateChecker.IsDue("20260928060000", "20260928120000"))
        ok("startup: due after 1 h", UpdateChecker.IsDue("20260928110000", "20260928120000", UpdateChecker.StartupHours))
        ok("startup: not due within 1 h", !UpdateChecker.IsDue("20260928113000", "20260928120000", UpdateChecker.StartupHours))
        ok("due: clock moved back", UpdateChecker.IsDue("20260929120000", "20260928120000"))
        ok("due: bad value", UpdateChecker.IsDue("garbage", "20260928120000"))
        ok("wanted: newer", UpdateChecker.IsWanted("2026.10.01", "2026.09.28", ""))
        ok("not wanted: same", !UpdateChecker.IsWanted("2026.09.28", "2026.09.28", ""))
        ok("not wanted: skipped", !UpdateChecker.IsWanted("2026.10.01", "2026.09.28", "2026.10.01"))
        ok("wanted: newer than skipped", UpdateChecker.IsWanted("2026.10.02", "2026.09.28", "2026.10.01"))

        ; 状态文件
        savedFile := UpdateChecker.StateFile
        root := A_Temp "\ALTRun-test-update-state"
        try DirDelete(root, true)
        UpdateChecker.StateFile := root "\Data\Update.json"
        savedProviders := ProviderRegistry.Providers
        ProviderRegistry.Providers := [SystemProvider]
        eq("state: empty", UpdateChecker.LoadState()["LastCheck"] "|" UpdateChecker.LoadState()["Skip"], "|")
        UpdateChecker.SaveState(Map("LastCheck", "20260928120000", "Skip", "2026.10.01"))
        state := UpdateChecker.LoadState()
        eq("state: saved", state["LastCheck"] "|" state["Skip"], "20260928120000|2026.10.01")

        ; 没有新版本: 空搜索框没有结果
        UpdateChecker.Pending := ""
        eq("no update: empty query", ProviderRegistry.Search("").Length, 0)
        eq("no update: item", UpdateChecker.PendingItem(), "")

        ; 有新版本: 空搜索框和搜索 "update" 都显示, 排在 "检查更新" 前面
        release := {Version: "2026.10.02", Page: "https://x/notes", ZipUrl: "https://x/a.zip", Sha256: "ab"}
        UpdateChecker.SetPending(release)
        eq("mode (tests run the source)", release.Mode, "source")
        results := ProviderRegistry.Search("")
        eq("empty query: one item", results.Length, 1)
        eq("empty query: title", results[1].Title, "Update Available: ALTRun 2026.10.02")
        eq("empty query: provider", results[1].Provider, "System")
        ok("source hint", InStr(results[1].Subtitle, "git pull"))
        eq("actions: notes + skip", results[1].Actions.Length, 2)
        titles := ""
        for action in ActionCatalog.ListFor(results[1])
            titles .= action.Title "|"
        ok("action panel", InStr(titles, "Release Notes|Skip This Version|"))
        results := ProviderRegistry.Search("update")
        ok("search update: first", results.Length >= 2 && results[1].Title = "Update Available: ALTRun 2026.10.02")
        ; 操作面板右侧的说明: 放得下时完整显示, 放不下时名称优先 (最多一半), 说明用剩下的
        eq("hint fits", SearchWindow.HintWidth(600, 120, 300, 12), 300)
        eq("hint shares the row", SearchWindow.HintWidth(400, 120, 600, 12), 268)
        eq("long title keeps half", SearchWindow.HintWidth(400, 500, 600, 12), 188)
        eq("no room", SearchWindow.HintWidth(20, 100, 50, 12), 0)
        for mode, text in Map("install", "Install Update", "page", "Download Page", "scoop", "scoop update altrun", "winget", "winget upgrade")
            release.Mode := mode, ok("subtitle " mode, InStr(UpdateChecker.PendingItem().Subtitle, text))

        ; 跳过这个版本: 不再显示, 记在状态文件里
        UpdateChecker.SkipVersion("2026.10.02")
        eq("skipped: no item", ProviderRegistry.Search("").Length, 0)
        eq("skipped: saved", UpdateChecker.LoadState()["Skip"], "2026.10.02")
        Hud.Hide()
        UpdateChecker.StateFile := savedFile
        ProviderRegistry.Providers := savedProviders
        try DirDelete(root, true)
    }

    ; 片段的占位符
    static SnippetPlaceholders() {
        eq := (n, a, e) => TestRunner.Equal("Placeholders." n, a, e)
        ok := (n, c) => TestRunner.True("Placeholders." n, c)
        eq("date format", SnippetProvider.Expand("{date:yyyy}"), FormatTime(, "yyyy"))
        eq("date plus", SnippetProvider.Expand("{date+7:yyyyMMdd}"), FormatTime(DateAdd(A_Now, 7, "Days"), "yyyyMMdd"))
        eq("date minus", SnippetProvider.Expand("x {date-1:yyyyMMdd} y"), "x " FormatTime(DateAdd(A_Now, -1, "Days"), "yyyyMMdd") " y")
        eq("default date", SnippetProvider.Expand("{date}"), FormatTime(, AppSettings.Extension("AutoDate")["DateFormat"]))
        eq("time format", StrLen(SnippetProvider.Expand("{time:HH:mm:ss}")), 8)
        eq("cursor untouched", SnippetProvider.Expand("a{cursor}b"), "a{cursor}b")
        uuids := SnippetProvider.Expand("{uuid} {uuid}")
        ok("uuid format", RegExMatch(uuids, "^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12} [0-9a-f-]{36}$"))
        ok("uuid different", StrSplit(uuids, " ")[1] != StrSplit(uuids, " ")[2])
        saved := ClipboardProvider.Entries
        ClipboardProvider.Entries := [Map("Text", "newest"), Map("Text", "previous"), Map("Text", "older")]
        eq("clipboard:1", SnippetProvider.Expand("[{clipboard:1}] [{clipboard:2}] [{clipboard:9}]"), "[previous] [older] []")
        ClipboardProvider.Entries := saved
    }

    ; 文字片段的导入 / 导出: 合并规则、表头、自动展开
    static SnippetImport() {
        eq := (n, a, e) => TestRunner.Equal("SnippetImport." n, a, e)
        snippets := [Map("Name", "Mail", "Keyword", "em", "Text", "old", "AutoExpand", 1)
            , Map("Name", "Address", "Keyword", "", "Text", "old address", "AutoExpand", 1)]
        rows := [["名称", "关键字", "正文", "自动展开"]
            , ["Mail (work)", "EM", "new@b.com", ""]                       ; 关键字相同 (不分大小写): 更新
            , ["Address", "", "new address"]                               ; 没有关键字, 名称相同: 更新
            , ["", "sig", "Best regards`nCaomao", "0"]                     ; 没有名称: 用关键字; 0 = 不自动展开
            , ["", "", "  Line one  `nLine two"]                           ; 没有名称和关键字: 用正文第一行
            , ["Spaced", "two words", "x"]                                 ; 关键字有空格 (自动展开用不了): 自动展开关掉
            , ["Empty", "e", "   "]]                                       ; 没有正文: 跳过
        result := SnippetProvider.ImportRows(rows, snippets, ";")
        eq("counts", result.Added "/" result.Updated "/" result.Skipped, "3/2/1")
        eq("updated by keyword", snippets[1]["Name"] "|" snippets[1]["Text"] "|" snippets[1]["Keyword"], "Mail (work)|new@b.com|EM")
        eq("updated by name", snippets[2]["Text"], "new address")
        eq("name from keyword", snippets[3]["Name"] "|" snippets[3]["AutoExpand"], "sig|0")
        eq("name from text", snippets[4]["Name"], "Line one")
        eq("keyword with a space", snippets[5]["AutoExpand"], 0)
        eq("total", snippets.Length, 5)
        eq("header only on the first row", SnippetProvider.ImportRows([["Mail", "m", "x"], ["Name", "k", "Text"]], [], ";").Added, 2)
        eq("header english", SnippetProvider.IsHeaderRow(["Name", "Keyword", "Text"]), true)
        eq("header japanese", SnippetProvider.IsHeaderRow(["名前", "キーワード", "本文", "自動"]), true)
        eq("not header", SnippetProvider.IsHeaderRow(["Mail", "em", "a@b.com"]), false)
        exported := SnippetProvider.ExportRows(snippets)
        eq("export header", SnippetProvider.IsHeaderRow(exported[1]), true)
        eq("export rows", exported.Length, 6)
        eq("export auto expand", exported[4][4], 0)
        again := []                                                         ; 导出再导入: 内容不变
        SnippetProvider.ImportRows(DelimitedText.Parse(DelimitedText.ToCsv(exported)), again, ";")
        same := again.Length = snippets.Length
        for index, snippet in snippets
            for key in ["Name", "Keyword", "Text", "AutoExpand"]
                same := same && again[index][key] == snippet[key]
        eq("round trip", same, true)
    }

    ; 选中内容的操作: 按内容类型生成操作面板的来源
    static SelectionItems() {
        eq := (n, a, e) => TestRunner.Equal("Selection." n, a, e)
        ok := (n, c) => TestRunner.True("Selection." n, c)
        titles(item) {
            text := ""
            for action in ActionCatalog.ListFor(item)
                text .= action.Title "|"
            return text
        }
        item := SelectionActions.ItemFor({Text: "hello world"})
        eq("text kind", item.Kind "|" item.Arg "|" item.Provider, "text|hello world|Selection")
        list := titles(item)
        ok("text: copy first", InStr(list, "Copy to Clipboard|") = 1)
        ok("text: search engines", InStr(list, "Search Google|") && InStr(list, "Search GitHub|"))
        ok("text: save as snippet", InStr(list, "Save as Snippet...|"))
        ok("text: replace with uppercase", InStr(list, "Replace with: UPPERCASE|"))
        ok("text: no calculation", !InStr(list, "= "))
        ; 剪贴板历史里的文字条目: → 菜单里也有搜索操作 (图片、文件没有)
        clipItem := ResultItem("hello", "", {Kind: "text", Arg: "hello world", Provider: ClipboardProvider.Id, Source: Map("Text", "hello world", "Time", A_Now)})
        ok("clipboard history: search engines", InStr(titles(clipItem), "Search Google|"))
        clipItem.Source := Map("Type", "image", "Time", A_Now)
        ok("clipboard history: image has no search", !InStr(titles(clipItem), "Search Google|"))
        list := titles(SelectionActions.ItemFor({Text: "12*(3+4)"}))
        ok("math: result first extra", InStr(list, "= 84|"))
        item := SelectionActions.ItemFor({Text: "  https://github.com/zhugecaomao/ALTRun  "})
        eq("link", item.Kind "|" item.Arg, "url|https://github.com/zhugecaomao/ALTRun")
        item := SelectionActions.ItemFor({Text: A_WinDir})
        eq("path text -> folder", item.Kind "|" item.Arg, "folder|" A_WinDir)
        item := SelectionActions.ItemFor({Files: [A_WinDir "\notepad.exe"]})
        eq("one file", item.Kind "|" item.Title, "file|notepad.exe")
        item := SelectionActions.ItemFor({Files: ["C:\a\one.txt", "C:\b\two.txt"]})
        eq("files", item.Title "|" item.Subtitle, "2 files|one.txt, two.txt")
        ok("files: add all", InStr(titles(item), "Add All to Custom Commands|"))
        eq("transform", SelectionActions.Transforms()[1][2].Call("abc"), "ABC")
    }
}
