;===============================================================================
; ExtensionTests.ahk - Extensions: CapsLock switch, date stamp, snippet expansion, quick switch
;   folders, PT tools, self-update and selection actions
;-------------------------------------------------------------------------------
; Included by Tests\RunTests.ahk. Every static method whose name doesn't start with "_" is a test;
; TestRunner runs them in name order, and each one must put back the shared state it changes.
;===============================================================================
class ExtensionTests {
    ; CapsLock switches the input language: setting values, choosing the next input language
    static CapsLockSwitch() {
        eq := (n, a, e) => TestRunner.Equal("CapsLockSwitch." n, a, e)
        eq("off by default", AppSettings.Defaults()["General"]["CapsLock"], "")
        eq("modes", (CapsLockSwitch.IsMode("Layout") ? 1 : 0) (CapsLockSwitch.IsMode("Mode") ? 1 : 0) (CapsLockSwitch.IsMode("") ? 1 : 0) (CapsLockSwitch.IsMode("x") ? 1 : 0), "1100")
        eq("hold time", CapsLockSwitch.HoldMs, 300)
        eq("notification text", CapsLockSwitch.StateText(true) "|" CapsLockSwitch.StateText(false), "Caps Lock On|Caps Lock Off")
        eq("uses the keyboard hook", InStr(CapsLockSwitch.Key, "$") ? 1 : 0, 1)    ; without the hook the Caps Lock toggle can't be stopped, and holding would keep switching
        eq("with any modifier", InStr(CapsLockSwitch.Key, "*") ? 1 : 0, 1)        ; Alt+CapsLock etc. sent extra by remote desktops must be caught too
        registered := [], savedInterval := A_HotkeyInterval
        CapsLockSwitch.Register((key, fn) => registered.Push(key))
        eq("down and up hotkeys", registered.Length = 2 ? registered[1] "|" registered[2] : "", "*$CapsLock|*$CapsLock up")
        eq("no too-many-hotkeys warning", A_HotkeyInterval, 0)
        A_HotkeyInterval := savedInterval

        ; the Caps Lock notice can be turned off (General.CapsLockNotify)
        eq("notice on by default", AppSettings.Defaults()["General"]["CapsLockNotify"], 1)
        notices := [], savedNotify := App.GetOwnPropDesc("Notify")
        App.DefineProp("Notify", {Call: (this, text, *) => notices.Push(text)})
        savedNotice := AppSettings.General.Has("CapsLockNotify") ? AppSettings.General["CapsLockNotify"] : 1
        AppSettings.General["CapsLockNotify"] := 1, App.CapsLockNotice(true)
        AppSettings.General["CapsLockNotify"] := 0, App.CapsLockNotice(false)
        AppSettings.General["CapsLockNotify"] := savedNotice
        App.DefineProp("Notify", savedNotify)
        eq("notice shown only when on", notices.Length = 1 ? notices[1] : notices.Length, "Caps Lock On")

        ; the helper for administrator windows (CapsLockHelper, #158)
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
        shown := []                                                         ; the Caps Lock notice goes to Notify (in the helper, ALTRun is told instead)
        savedNotify := CapsLockSwitch.Notify, CapsLockSwitch.Notify := (on) => shown.Push(on)
        before := GetKeyState("CapsLock", "T")
        CapsLockSwitch.ToggleCapsLock()
        eq("notify on toggle", shown.Length = 1 && shown[1] = !before ? 1 : 0, 1)
        SetCapsLockState(before ? "On" : "Off"), CapsLockSwitch.Notify := savedNotify

        ; order of presses / releases (a fake clock, timers fired by hand)
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
            Loop 20                                                         ; auto-repeat while held, every 30 ms
                at(2000 + A_Index * 30), CapsLockSwitch.Down()
            CapsLockSwitch.Hold()                                           ; timer: held for 0.3 s
            Loop 10
                at(2600 + A_Index * 30), CapsLockSwitch.Down()
            at(3000), CapsLockSwitch.Up()
            eq("hold: caps lock once, never switches", calls.Text, "caps;")
            calls.Text := ""
            at(4000), CapsLockSwitch.Down(), at(4400), CapsLockSwitch.Up()  ; the timer didn't get to run
            eq("late timer: still a hold", calls.Text, "caps;")
            calls.Text := ""
            at(5000), CapsLockSwitch.Down(), at(7000), CapsLockSwitch.Down(), at(7050), CapsLockSwitch.Up()   ; the last release was lost
            eq("lost key-up: new press", calls.Text, "switch;")
            calls.Text := ""
            CapsLockSwitch.Up()
            eq("extra key-up ignored", calls.Text, "")
            at(8000), CapsLockSwitch.Down(), at(8000.4), CapsLockSwitch.Up()  ; sent by Chrome Remote Desktop to align Caps Lock
            eq("remote sync tap ignored", calls.Text, "")
            at(9000), CapsLockSwitch.Down(), at(9040), CapsLockSwitch.Up()
            eq("quick real tap still switches", calls.Text, "switch;")
            calls.Text := ""
            shift.Down := true
            at(10000), CapsLockSwitch.Down(), at(10080), CapsLockSwitch.Up()
            eq("shift+caps lock: caps lock as usual", calls.Text, "caps;")
            calls.Text := ""
            at(11000), CapsLockSwitch.Down(), at(11000.3), CapsLockSwitch.Up()   ; sent by a remote desktop while typing Shift+letter
            eq("remote sync tap with shift ignored", calls.Text, "")
            shift.Down := false
            CapsLockSwitch.Clock := saved["Clock"].Value
            ; Sleep goes by A_TickCount in 15.6 ms steps and may sleep almost a step short: sleeps 100 ms, only requires over 50 ms measured
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
        ; Ctrl+D is only taken in a rename box (at other times the key goes to the program unchanged, not sent again)
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
        ; the date format changed in the settings: a date in that format at the end is still updated, not added again
        eq("iso update", AutoDate.AddDateToName("Report - 2020-01-01.docx", "2026-09-23", "yyyy-MM-dd"), "Report - 2026-09-23.docx")
        eq("iso folder", AutoDate.AddDateToName("Project - 2020-01-01", "2026-09-23", "yyyy-MM-dd"), "Project - 2026-09-23")
        eq("dotted iso folder", AutoDate.AddDateToName("Project - 2020.01.01", "2026.09.23", "yyyy.MM.dd"), "Project - 2026.09.23")
        eq("compact", AutoDate.AddDateToName("Plan_20200101.pdf", "20260923", "yyyyMMdd"), "Plan_20200101 - 20260923.pdf")
        eq("compact update", AutoDate.AddDateToName("Plan - 20200101.pdf", "20260923", "yyyyMMdd"), "Plan - 20260923.pdf")
        eq("month name", AutoDate.AddDateToName("Memo - 1 Jan 2020", "23 Sep 2026", "d MMM yyyy"), "Memo - 23 Sep 2026")
        eq("other format kept", AutoDate.AddDateToName("Report - 01.01.2020.docx", "2026-09-23", "yyyy-MM-dd"), "Report - 01.01.2020 - 2026-09-23.docx")
        eq("pattern", AutoDate.DatePattern("dd.MM.yyyy"), "\d{2}\.\d{2}\.\d{4}")
    }

    ; Paste mode Auto: short single-line text is typed (clipboard untouched); multi-line / Tab / very long / emoji go through the clipboard
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

        ; really typed into an edit box: the text appears and the clipboard is unchanged
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
        abbr := (snippet, prefix) => TextTools.JoinList(SnippetExpander.Abbreviations(snippet, prefix), ",")
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

    ; The dialogs' folder menu: recent folders (Windows' "Recent items"), menu text
    static FolderMenu() {
        eq := (n, a, e) => TestRunner.Equal("FolderMenu." n, a, e)
        root := A_Temp "\ALTRun-test-recent"
        try DirDelete(root, true)
        DirCreate(root "\Recent"), DirCreate(root "\Projects\Tower"), DirCreate(root "\Docs")
        longPath := Buffer(2048)                                            ; A_Temp may be an 8.3 short path (C:\Users\RUNNER~1), while shortcuts read back long paths
        if DllCall("GetLongPathNameW", "Str", root, "Ptr", longPath, "UInt", 1024)
            root := StrGet(longPath)
        FileAppend("x", root "\Docs\report.pdf")
        FileCreateShortcut(root "\Projects\Tower", root "\Recent\Tower.lnk")
        FileSetTime("20260101000000", root "\Recent\Tower.lnk")
        FileCreateShortcut(root "\Docs\report.pdf", root "\Recent\report.pdf.lnk")    ; a file: its folder is taken
        FileSetTime("20260201000000", root "\Recent\report.pdf.lnk")
        FileCreateShortcut(root "\Docs", root "\Recent\Docs.lnk")                     ; the same folder as the previous one: listed once
        FileSetTime("20260115000000", root "\Recent\Docs.lnk")
        FileCreateShortcut(root "\Gone", root "\Recent\Gone.lnk")                     ; no longer exists
        FileSetTime("20260301000000", root "\Recent\Gone.lnk")
        FileAppend("not a shortcut", root "\Recent\Broken.lnk")                     ; a broken shortcut: skipped without an error
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
        ; File dialogs: #32770 needs Explorer's file view; the main window of "System Information" (msinfo32) only has a plain list and doesn't count
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
        eq("jump: system dialog with OK, as messages", QuickSwitch.JumpMethod("#32770", true, true) " " QuickSwitch.JumpMethod("TMyOpenDialog", true, true), "message edit")
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
        eq("transient windows while switching", QuickSwitch.IsTransientWindow("Shell_TrayWnd") QuickSwitch.IsTransientWindow("XamlExplorerHostIslandWindow"), "11")
        eq("not transient", QuickSwitch.IsTransientWindow("TTOTAL_CMD") QuickSwitch.IsTransientWindow("#32770"), "00")
        ; Windows 11's Alt+Tab window is already hidden when its foreground event is handled: still read its class, and pass over it
        passing := Gui()
        eq("class of a hidden window", QuickSwitch.ClassOf(passing.Hwnd), "AutoHotkeyGUI")
        eq("a hidden window in front is passing", QuickSwitch.IsPassingWindow(passing.Hwnd, "AutoHotkeyGUI"), true)
        passing.Show("NA w100 h50")
        eq("a shown window is not", QuickSwitch.IsPassingWindow(passing.Hwnd, "AutoHotkeyGUI"), false)
        gone := passing.Hwnd
        passing.Destroy()
        eq("a window gone is passing", QuickSwitch.ClassOf(gone) "|" QuickSwitch.IsPassingWindow(gone, ""), "|1")
        ; A dialog can be in front before it is shown: it still counts as a possible dialog (polling goes on until it shows)
        savedOptions := QuickSwitch.Options
        QuickSwitch.Options := AppSettings.Extension("QuickSwitch")
        GroupAdd("ALTRunDialogs", "QuickSwitch hidden dialog probe")
        otherWindow := Gui(, "QuickSwitch other window"), otherWindow.Show("w200 h100")
        hiddenDialog := Gui(, "QuickSwitch hidden dialog probe"), hiddenDialog.Show("Hide w200 h100")
        DllCall("SetForegroundWindow", "Ptr", hiddenDialog.Hwnd), Sleep(200)
        DetectHiddenWindows(false)
        if (DllCall("GetForegroundWindow", "Ptr") = hiddenDialog.Hwnd) {
            eq("hidden dialog in front: not seen by default", WinActive("ahk_group ALTRunDialogs") != 0, false)
            eq("hidden dialog in front: may be a dialog", QuickSwitch._MayBeDialog(), true)
            eq("hidden windows setting restored", A_DetectHiddenWindows, false)
        }
        WinActivate("ahk_id " otherWindow.Hwnd), Sleep(200)
        eq("other window in front: not a dialog", QuickSwitch._MayBeDialog(), false)
        hiddenDialog.Destroy(), otherWindow.Destroy()
        QuickSwitch.Options := savedOptions
        ; Dragging another window than the dialog: nothing to follow
        savedDrag := [QuickSwitch._placeTimer, QuickSwitch._panelFor]
        QuickSwitch._placeTimer := () => 0, QuickSwitch._panelFor := 0x1234
        QuickSwitch._OnEvent(0, 0xA, 0x5678, 0, 0, 0, 0)
        eq("drag: another window", QuickSwitch._dragHook, 0)
        QuickSwitch._placeTimer := savedDrag[1], QuickSwitch._panelFor := savedDrag[2]
        ; Foreground changes and a possible dialog's controls arrive as events, with no polling (Wine doesn't deliver WinEvents, so only checked on Windows)
        if !DllCall("GetProcAddress", "Ptr", DllCall("GetModuleHandle", "Str", "ntdll", "Ptr"), "AStr", "wine_get_version", "Ptr") {
            saved := QuickSwitch._checkTimer
            counts := {Check: 0}
            QuickSwitch._checkTimer := () => counts.Check++
            first := Gui("+AlwaysOnTop", "QuickSwitch event A"), first.Show("w200 h100 x100 y100")
            second := Gui("+AlwaysOnTop", "QuickSwitch event B"), second.Show("w200 h100 x400 y100")
            hook := QuickSwitch._HookEvent(0x3, 0x3)
            WinActivate("ahk_id " first.Hwnd), Sleep(300)
            WinActivate("ahk_id " second.Hwnd), Sleep(300)
            switched := counts.Check
            QuickSwitch._WatchWindow(first.Hwnd, "candidate")              ; a possible dialog: object events of its UI thread
            hooks := QuickSwitch._threadHooks.Length
            counts.Check := 0
            first.Add("Button", , "Late control"), Sleep(300)               ; a control created after the window came to the front
            shownLater := counts.Check
            QuickSwitch._WatchWindow(first.Hwnd, "dialog")                 ; recognized: only moving / resizing
            hooks .= " " QuickSwitch._threadHooks.Length
            counts.Check := 0
            first.Add("Button", , "Another control"), Sleep(300)
            shownAfter := counts.Check
            QuickSwitch._WatchWindow(first.Hwnd, "settle")                 ; another window just in front: its title only
            hooks .= " " QuickSwitch._threadHooks.Length
            counts.Check := 0
            first.Title := "QuickSwitch event A renamed", Sleep(300)
            renamed := counts.Check
            QuickSwitch._WatchWindow(first.Hwnd, "panel")                  ; the caret in the panel: the dialog's thread, its windows going
            hooks .= " " QuickSwitch._threadHooks.Length
            ; Dragging the dialog: its location changes move the panel along until the drag ends, no timer
            savedDrag := [QuickSwitch._placeTimer, QuickSwitch._panelFor, QuickSwitch._dragTimer]
            places := {Count: 0}
            QuickSwitch._placeTimer := () => places.Count++, QuickSwitch._dragTimer := () => 0, QuickSwitch._panelFor := first.Hwnd
            QuickSwitch._OnEvent(0, 0xA, first.Hwnd, 0, 0, 0, 0)
            dragHooked := QuickSwitch._dragHook != 0
            Loop 5
                WinMove(100 + A_Index * 20, 100, , , "ahk_id " first.Hwnd), Sleep(60)
            followed := places.Count
            QuickSwitch._OnEvent(0, 0xB, first.Hwnd, 0, 0, 0, 0), Sleep(100)
            afterEnd := places.Count
            WinMove(300, 100, , , "ahk_id " first.Hwnd), Sleep(200)
            dragUnhooked := QuickSwitch._dragHook = 0
            QuickSwitch._placeTimer := savedDrag[1], QuickSwitch._panelFor := savedDrag[2], QuickSwitch._dragTimer := savedDrag[3]
            QuickSwitch._UnwatchWindow()
            DllCall("UnhookWinEvent", "Ptr", hook)
            first.Destroy(), second.Destroy()
            QuickSwitch._checkTimer := saved
            TestRunner.True("QuickSwitch.foreground events (" switched ")", hook && switched >= 2)
            eq("thread events hooked: candidate dialog settle panel", hooks, "2 1 1 1")
            TestRunner.True("QuickSwitch.a control shown later is checked (" shownLater ")", shownLater >= 1)
            eq("controls of a recognized dialog aren't watched", shownAfter, 0)
            TestRunner.True("QuickSwitch.settle: a title set later is checked (" renamed ")", renamed >= 1)
            TestRunner.True("QuickSwitch.drag: the panel follows the location changes (" followed ")", dragHooked && followed >= 3)
            eq("drag: unhooked when it ends, placed once more", dragUnhooked " " (afterEnd - followed >= 1) " " (places.Count = afterEnd), "1 1 1")
            ; The whole way: an Open dialog in another process gets the panel, nothing polls while it is open, and the panel closes with it
            QuickSwitch.Options := AppSettings.Extension("QuickSwitch")
            GroupAdd("ALTRunDialogs", "ahk_class #32770")
            QuickSwitch._StartWatching()
            helper := A_Temp "\ALTRunTestOpenDialog.ahk"
            try FileDelete(helper)
            FileAppend('#NoTrayIcon`nFileSelect(, A_WinDir, "Open")`n', helper, "UTF-8")
            Run('"' A_AhkPath '" "' helper '"', , , &dialogPid)
            dialog := WinWait("Open ahk_class #32770 ahk_pid " dialogPid, , 10)
            panelShown := false
            Loop 50 {
                if (panelShown := IsObject(QuickSwitch._panel) && DllCall("IsWindowVisible", "Ptr", QuickSwitch._panel.Hwnd))
                    break
                Sleep(100)
            }
            Sleep(QuickSwitch.SettleMs + 300)
            pollingWhileOpen := QuickSwitch._polling, phaseWhileOpen := QuickSwitch._phase
            try WinClose("ahk_id " dialog)
            if !ProcessWaitClose(dialogPid, 3)
                try ProcessClose(dialogPid)
            Sleep(300)
            panelGone := !IsObject(QuickSwitch._panel)
            Sleep(QuickSwitch.SettleMs + 300)                               ; the window in front afterwards: watched for its title only for SettleMs
            idleAfter := QuickSwitch._polling " " (QuickSwitch._phase = "settle")
            DllCall("UnhookWinEvent", "Ptr", QuickSwitch._foregroundHook), QuickSwitch._foregroundHook := 0
            QuickSwitch._UnwatchWindow()
            SetTimer(QuickSwitch._pollTimer, 0), SetTimer(QuickSwitch._checkTimer, 0), QuickSwitch._polling := false
            QuickSwitch.HidePanel()
            TestRunner.True("QuickSwitch.panel below another process's Open dialog", dialog && panelShown)
            eq("no polling while the dialog is open", pollingWhileOpen " " phaseWhileOpen, "0 dialog")
            eq("afterwards: no polling, not watched for longer than SettleMs", idleAfter, "0 0")
            TestRunner.True("QuickSwitch.panel closed with the dialog", panelGone)
        }
        ; A real one: AutoHotkey's DirSelect is SHBrowseForFolder (option 4 = classic style). Detected, the folder is
        ; selected in its tree, and OK returns it
        target := A_WinDir "\System32"
        for options in [0, 4] {
            seen := {Detected: "", Selected: ""}, tries := 0
            probe() {
                if !(dlg := SearchKeys.InSystemDialog()) {
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
        ; TC's path is asked for with WM_COPYDATA (another process imitates TC's reply), without touching the clipboard
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
        try QuickSwitch.SetDialogPath(A_Temp "\", closedHwnd)               ; the dialog has closed: do nothing (don't type into another window)
        catch
            threw := true
        eq("closed dialog ignored", threw, false)
        ; A real Save As in another process: the jump goes in as messages (no keys), also while Alt is held as right after Alt+Tab; the name
        ; Save As suggested is put back, and OK then saves in the folder jumped to (Wine's dialogs differ, so only checked on Windows)
        if !DllCall("GetProcAddress", "Ptr", DllCall("GetModuleHandle", "Str", "ntdll", "Ptr"), "AStr", "wine_get_version", "Ptr") {
            target := A_Temp "\ALTRun jump test"
            try DirDelete(target, true)
            DirCreate(target)
            target := TestRunner.LongPath(target)
            helper := A_Temp "\ALTRunTestSaveDialog.ahk", result := A_Temp "\ALTRunTestSaveDialog.txt"
            for held in ["Alt", "Ctrl"] {                                   ; Alt: AutoSwitch right after Alt+Tab; Ctrl: the Ctrl+G hotkey
                try FileDelete(helper)
                try FileDelete(result)
                try FileDelete(target "\Report 2026*")
                FileAppend('#NoTrayIcon`nFileAppend(FileSelect("S", A_WinDir "\\Report 2026.txt", "Save probe"), "' result '", "UTF-8")`n', helper, "UTF-8")
                Run('"' A_AhkPath '" "' helper '"', , , &dialogPid)
                dialog := WinWait("Save probe ahk_class #32770 ahk_pid " dialogPid, , 10)
                jumpMethod := dialog ? QuickSwitch.DialogJumpMethod(dialog) : "", nameAfter := "no dialog", nameBefore := ""
                if dialog {
                    WinActivate("ahk_id " dialog), WinWaitActive("ahk_id " dialog, , 3), Sleep(300)
                    nameBefore := ControlGetText("Edit1", "ahk_id " dialog)    ; "Report 2026", or with ".txt" when extensions are shown
                    Send("{" held " down}")
                    try QuickSwitch.SetDialogPath(target "\", dialog)
                    finally Send("{" held " up}")
                    Sleep(500)
                    shown := WinExist("ahk_id " dialog) ? QuickSwitch.DialogFolder(dialog) : "the dialog closed"
                    nameAfter := WinExist("ahk_id " dialog) ? ControlGetText("Edit1", "ahk_id " dialog) : "the dialog closed"
                    DllCall("SendMessageTimeoutW", "Ptr", dialog, "UInt", 0x111, "Ptr", 1, "Ptr", 0, "UInt", 0x2, "UInt", 3000, "Ptr", 0)   ; OK
                }
                if !ProcessWaitClose(dialogPid, 5)
                    try ProcessClose(dialogPid)
                saved := ""
                try saved := FileRead(result, "UTF-8")
                eq("message jump: a system Save As", jumpMethod, "message")
                eq("message jump (" held " held): the address bar shows the folder", dialog ? shown : "no dialog", target)
                eq("message jump (" held " held): suggested name put back", nameAfter, nameBefore)
                eq("message jump (" held " held): saved in the folder jumped to", RegExReplace(saved, "\.txt$"), target "\Report 2026")
                try FileDelete(A_WinDir "\Report 2026*")                    ; only if a jump went wrong and saved in the old folder
            }
            try DirDelete(target, true)
            try FileDelete(helper)
            try FileDelete(result)
        }
        QuickSwitch._BuildPanel(A_ScriptHwnd)                              ; really build the panel once (a wrong Gui option would throw here)
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
        QuickSwitch._ResizePanel(900)                                      ; the dialog gets wider: the panel follows
        QuickSwitch._panelList.GetPos(, , &listW)
        eq("resize width", QuickSwitch._panelWidth "|" listW, "900|900")
        eq("key other window", QuickSwitch._OnKeyDown(0x28, 0, 0x100, A_ScriptHwnd), "")
        eq("key down handled", QuickSwitch._OnKeyDown(0x28, 0, 0x100, QuickSwitch._panelSearch.Hwnd), 0)
        eq("ime key passes", QuickSwitch._OnKeyDown(0xE5, 0, 0x100, QuickSwitch._panelSearch.Hwnd), "")   ; the IME is composing: not intercepted
        searchHwnd := QuickSwitch._searchHwnd
        QuickSwitch.HidePanel()
        eq("panel hidden", QuickSwitch._panel, "")
        eq("controls released", QuickSwitch._panelSearch "|" QuickSwitch._searchHwnd, "|0")
        eq("key after hide", QuickSwitch._OnKeyDown(0x0D, 0, 0x100, searchHwnd), "")   ; after the panel closes: no error, nothing handled
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

    ; The tendon profile calculation compared value by value with the original SPF2M.EXE: Tests\Fixtures\SPF2M-Reference.json holds the results of running
    ; SPF2M in DOSBox (4 profile types x 6 strand types, rising / falling, spans that aren't whole meters, custom radius / contraflexure /
    ; support spacing, and the cases where SPF2M reports an error or crashes); how it was made: see Tests\Tools\SPF2M
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

    ; Input checks, support spacing, cases SPF2M didn't handle well (crashes / shows nothing)
    ; PT Tools window: result boxes take no Tab stop (from Area expression straight to Bar diameter), units written as mm², even margins,
    ; inputs are combo boxes with their common values, the span width follows into the mesh area
    ; Only inputs are saved: Load fills in the defaults and drops keys no longer used (renamed fields, saved results)
    static PTToolsLoad() {
        eq := (n, a, e) => TestRunner.Equal("PTToolsLoad." n, a, e)
        saved := PTToolsWindow.Settings
        try {
            data := PTToolsWindow.Load(Map("RebarSpanWidth", "5700", "BrcSpanWidth", 3800, "RebarTotalArea", 5969.04))
            eq("kept", data["RebarSpanWidth"], "5700")
            eq("default filled in", data["MeshSpanWidth"], 3000)
            eq("old and result keys dropped", data.Has("BrcSpanWidth") || data.Has("RebarTotalArea"), false)
            stray := ""
            for name in PTToolsWindow.RebarFieldNames
                stray .= PTToolsWindow.Defaults.Has(name) ? "" : name " "
            eq("every saved field has a default", stray, "")
        } finally
            PTToolsWindow.Settings := saved
    }

    static PTToolsWindowUi() {
        eq := (n, a, e) => TestRunner.Equal("PTToolsWindowUi." n, a, e)
        saved := [PTToolsWindow.Settings["WinLeft"], PTToolsWindow.Settings["WinTop"]]
        PTToolsWindow.Settings["WinLeft"] := -3000, PTToolsWindow.Settings["WinTop"] := -3000
        PTToolsWindow.Show()
        g := PTToolsWindow.G
        tabbed := "", oldUnit := ""
        for ctrl in g {
            style := ControlGetStyle(ctrl)
            if (ctrl.Type = "Edit" && (style & 0x800) && (style & 0x10000))     ; ES_READONLY still with WS_TABSTOP
                tabbed .= ctrl.Name " "
            if InStr(ctrl.Text, "mm2")
                oldUnit .= ctrl.Text " | "
        }
        eq("even margins", PTToolsWindowUi_EvenMargins(g), true)
        eq("results calculated on opening", g["RebarBarCount"].Value != "" && g["RequiredBarCount"].Value != "" && g["MeshTopCombinedArea"].Value != "" && g["ExprResult"].Value != "", true)
        eq("read-only fields skip tab", tabbed, "")
        eq("units", oldUnit, "")
        next := DllCall("GetNextDlgTabItem", "Ptr", g.Hwnd, "Ptr", g["RequiredAreaExpression"].Hwnd, "Int", 0, "Ptr")
        eq("expression -> diameter", next = g["RequiredBarDiameter"].Hwnd, true)
        editable := ""                                                      ; every input is a combo box (native: the first click selects all); Edit only for results
        for ctrl in g
            if (ctrl.Type = "Edit" && !(ControlGetStyle(ctrl) & 0x800))
                editable .= ctrl.Name " "
        eq("inputs are combo boxes", editable, "")
        types := ""
        for name, list in PTToolsWindow.ComboFields
            types .= (g[name].Type = "ComboBox" && g[name].Text = PTToolsWindow.Settings[name]) ? "" : name " "
        eq("combo boxes with the saved text", types, "")
        spanEdit := DllCall("FindWindowEx", "Ptr", g["RebarSpanWidth"].Hwnd, "Ptr", 0, "Str", "Edit", "Ptr", 0, "Ptr")
        eq("span width digits only", (ControlGetStyle(spanEdit) & 0x2000) != 0, true)
        meshEdit := DllCall("FindWindowEx", "Ptr", g["MeshTopMark"].Hwnd, "Ptr", 0, "Str", "Edit", "Ptr", 0, "Ptr")
        eq("mesh mark in capitals", (ControlGetStyle(meshEdit) & 0x8) != 0, true)
        g["RebarSpanWidth"].Text := "5700"                                  ; a list item: .Value would be its index, the handlers read .Text
        g["RebarBarSpacing"].Text := "150"
        PTToolsWindow.OnRebarFieldChanged()
        eq("bar count from the combo texts", g["RebarBarCount"].Value, "38.00")
        eq("mesh span follows", g["MeshSpanWidth"].Text, "5700")
        eq("top mesh area recalculated", g["MeshTopArea"].Value, PTToolsWindow.Fmt2(5.7 * PTToolsWindow.CalculateMeshArea(g["MeshTopMark"].Text)))
        eq("bottom mesh area recalculated", g["MeshBotArea"].Value, PTToolsWindow.Fmt2(5.7 * PTToolsWindow.CalculateMeshArea(g["MeshBotMark"].Text)))
        g["RequiredAreaExpression"].Text := "5000"
        PTToolsWindow.OnRequiredExpressionChanged()
        eq("area expression from the combo text", g["RequiredTotalArea"].Value, "5000.00")
        g["ExprInput"].Text := "(6000/200+1)*201.06"
        PTToolsWindow.OnExpressionChanged()
        eq("calculator from the combo text", g["ExprResult"].Value, "6232.86")
        examples := ""
        for expr in PTToolsWindow.Expressions
            examples .= IsNumber(Calc.Eval(expr)) ? "" : expr " "
        eq("calculator examples all work", examples, "")
        safety := g["SafetyFactorEnabled"]
        eq("safety factor checkbox skips tab", (ControlGetStyle(safety) & 0x10000) != 0, false)
        ; Numbers in the bar diameter combo box are right-aligned: the first character's position (EM_POSFROMCHAR) is in the right half of the box (Wine doesn't support changing alignment after creation, so only checked on Windows)
        diaEdit := DllCall("FindWindowEx", "Ptr", g["RebarBarDiameter"].Hwnd, "Ptr", 0, "Str", "Edit", "Ptr", 0, "Ptr")
        eq("diameter diaEdit has ES_RIGHT", (ControlGetStyle(diaEdit) & 0x2) != 0, true)
        if !DllCall("GetProcAddress", "Ptr", DllCall("GetModuleHandle", "Str", "ntdll", "Ptr"), "AStr", "wine_get_version", "Ptr") {
            WinGetClientPos(, , &ew, , diaEdit)
            firstX := SendMessage(0xD6, 0, 0, diaEdit) & 0xFFFF                     ; EM_POSFROMCHAR: x of character 0
            TestRunner.True("PTToolsWindowUi.diameter text on the right (first char at " firstX " of " ew ")", firstX > ew // 2)
        }
        g.Destroy(), PTToolsWindow.G := ""
        PTToolsWindow.Settings["WinLeft"] := saved[1], PTToolsWindow.Settings["WinTop"] := saved[2]

        saved := [PTToolsWindow.Settings["Spf2mWinLeft"], PTToolsWindow.Settings["Spf2mWinTop"]]   ; SPF2M: the same margins
        PTToolsWindow.Settings["Spf2mWinLeft"] := -3000, PTToolsWindow.Settings["Spf2mWinTop"] := -3000
        PTToolsWindow.ShowSpf2m()
        eq("SPF2M even margins", PTToolsWindowUi_EvenMargins(PTToolsWindow.SpfG), true)
        PTToolsWindow.SpfG.Destroy(), PTToolsWindow.SpfG := ""
        PTToolsWindow.Settings["Spf2mWinLeft"] := saved[1], PTToolsWindow.Settings["Spf2mWinTop"] := saved[2]

        ; the same margin left, right and below the group boxes (the title bar is above)
        PTToolsWindowUi_EvenMargins(g) {
            left := 99999, right := 0, bottom := 0
            for ctrl in g
                if (ctrl.Type = "GroupBox") {
                    ControlGetPos(&cx, &cy, &cw, &ch, ctrl)
                    left := Min(left, cx), right := Max(right, cx + cw), bottom := Max(bottom, cy + ch)
                }
            WinGetClientPos(, , &clientW, &clientH, g)
            return Abs(clientW - right - left) <= 1 && Abs(clientH - bottom - left) <= 1
        }
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
        ; where SPF2M crashes: an error description here
        eq("psp not achievable", InStr(profileOf(2, 900, 100, 3000, Map("Tendon", 6)).Error, "NOT ACHIEVABLE") > 0, true)
        ; where SPF2M shows nothing: the contraflexure point beyond half the span
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

    ; One-click update: parsing GitHub's Release, SHA256, replacing files (with a fake program in a temporary folder, no download)
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
        tar := A_WinDir "\System32\tar.exe"                                ; extraction (tar.exe is built in since Windows 10; skipped when Wine lacks it)
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

        write(dest "\Launcher.exe", "renamed old")                          ; the user renamed ALTRun.exe
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

        ; copying fails halfway (the target file is held exclusively by another program): overwritten files are put back, new files and folders deleted, the exe untouched
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

    ; Background update checks: once a day, skipped versions, the update notice in the search window (no network: Pending is set directly)
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

        ; the state file
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

        ; no new version: no results in the empty search box
        UpdateChecker.Pending := ""
        eq("no update: empty query", ProviderRegistry.Search("").Length, 0)
        eq("no update: item", UpdateChecker.PendingItem(), "")

        ; a new version: shown in the empty search box and for "update", before "Check for updates"
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
        for action in ResultActions.ListFor(results[1])
            titles .= action.Title "|"
        ok("action panel", InStr(titles, "Release Notes|Skip This Version|"))
        results := ProviderRegistry.Search("update")
        ok("search update: first", results.Length >= 2 && results[1].Title = "Update Available: ALTRun 2026.10.02")
        ; Description on the right of the actions panel: in full when it fits, otherwise the name first (at most half) and the description gets the rest
        eq("hint fits", ResultRows.HintWidth(600, 120, 300, 12), 300)
        eq("hint shares the row", ResultRows.HintWidth(400, 120, 600, 12), 268)
        eq("long title keeps half", ResultRows.HintWidth(400, 500, 600, 12), 188)
        eq("no room", ResultRows.HintWidth(20, 100, 50, 12), 0)
        for mode, text in Map("install", "Install Update", "page", "Download Page", "scoop", "scoop update altrun", "winget", "winget upgrade")
            release.Mode := mode, ok("subtitle " mode, InStr(UpdateChecker.PendingItem().Subtitle, text))

        ; skip this version: no longer shown, noted in the state file
        UpdateChecker.SkipVersion("2026.10.02")
        eq("skipped: no item", ProviderRegistry.Search("").Length, 0)
        eq("skipped: saved", UpdateChecker.LoadState()["Skip"], "2026.10.02")
        Hud.Hide()
        UpdateChecker.StateFile := savedFile
        ProviderRegistry.Providers := savedProviders
        try DirDelete(root, true)
    }

    ; snippet placeholders
    static SnippetPlaceholders() {
        eq := (n, a, e) => TestRunner.Equal("Placeholders." n, a, e)
        ok := (n, c) => TestRunner.True("Placeholders." n, c)
        eq("date format", SnippetProvider.Expand("{date:yyyy}"), FormatTime(, "yyyy"))
        eq("date plus", SnippetProvider.Expand("{date+7:yyyyMMdd}"), FormatTime(DateAdd(A_Now, 7, "Days"), "yyyyMMdd"))
        eq("date minus", SnippetProvider.Expand("x {date-1:yyyyMMdd} y"), "x " FormatTime(DateAdd(A_Now, -1, "Days"), "yyyyMMdd") " y")
        eq("default date", SnippetProvider.Expand("{date}"), FormatTime(, AppSettings.Extension("AutoDate")["DateFormat"]))
        eq("time format", StrLen(SnippetProvider.Expand("{time:HH:mm:ss}")), 8)
        eq("cursor untouched", SnippetProvider.Expand("a{cursor}b"), "a{cursor}b")
        text := "Hi {cursor},`r`n`r`nThanks.`r`n`r`nBest regards,"
        eq("caret back: a line break is one step", SnippetProvider.CaretBack(&text) "|" text, "25|Hi ,`r`n`r`nThanks.`r`n`r`nBest regards,")
        text := "a{cursor}`nb"
        eq("caret back: LF only", SnippetProvider.CaretBack(&text) "|" text, "2|a`nb")
        text := "no cursor"
        eq("caret back: none", SnippetProvider.CaretBack(&text) "|" text, "0|no cursor")
        text := "a{cursor}b{cursor}c"
        eq("caret back: only the first {cursor}", SnippetProvider.CaretBack(&text) "|" text, "2|abc")
        text := "end{cursor}"
        eq("caret back: at the end", SnippetProvider.CaretBack(&text), 0)
        uuids := SnippetProvider.Expand("{uuid} {uuid}")
        ok("uuid format", RegExMatch(uuids, "^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12} [0-9a-f-]{36}$"))
        ok("uuid different", StrSplit(uuids, " ")[1] != StrSplit(uuids, " ")[2])
        saved := ClipboardProvider.Entries
        ClipboardProvider.Entries := [Map("Text", "newest"), Map("Text", "previous"), Map("Text", "older")]
        eq("clipboard:1", SnippetProvider.Expand("[{clipboard:1}] [{clipboard:2}] [{clipboard:9}]"), "[previous] [older] []")
        ClipboardProvider.Entries := saved
    }

    ; Snippet import / export: merge rules, header, auto expansion
    static SnippetImport() {
        eq := (n, a, e) => TestRunner.Equal("SnippetImport." n, a, e)
        snippets := [Map("Name", "Mail", "Keyword", "em", "Text", "old", "AutoExpand", 1)
            , Map("Name", "Address", "Keyword", "", "Text", "old address", "AutoExpand", 1)]
        rows := [["名称", "关键字", "正文", "自动展开"]
            , ["Mail (work)", "EM", "new@b.com", ""]                       ; same keyword (case-insensitive): updated
            , ["Address", "", "new address"]                               ; no keyword, same name: updated
            , ["", "sig", "Best regards`nCaomao", "0"]                     ; no name: the keyword is used; 0 = no auto expansion
            , ["", "", "  Line one  `nLine two"]                           ; no name or keyword: the first line of the text
            , ["Spaced", "two words", "x"]                                 ; a keyword with a space (auto expansion can't use it): auto expansion off
            , ["Empty", "e", "   "]]                                       ; no text: skipped
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
        again := []                                                         ; export then import: content unchanged
        SnippetProvider.ImportRows(DelimitedText.Parse(DelimitedText.ToCsv(exported)), again, ";")
        same := again.Length = snippets.Length
        for index, snippet in snippets
            for key in ["Name", "Keyword", "Text", "AutoExpand"]
                same := same && again[index][key] == snippet[key]
        eq("round trip", same, true)
    }

    ; Actions for the selection: the source of the actions panel by content type
    static SelectionItems() {
        eq := (n, a, e) => TestRunner.Equal("Selection." n, a, e)
        ok := (n, c) => TestRunner.True("Selection." n, c)
        titles(item) {
            text := ""
            for action in ResultActions.ListFor(item)
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
        ; text entries in the clipboard history: the → menu has search actions too (images and files don't)
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
