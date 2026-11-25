;===============================================================================
; FileDialogs.ahk - recognizing file dialogs and jumping them to a folder (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Which window is a real open / save file dialog or a "Browse for Folder" dialog, and how to send one to a folder:
; the folder box, a message to the dialog, the address bar or typing (DialogJumpMethod); selecting the folder in a
; Browse for Folder dialog's tree; reading a dialog's current folder. Used by QuickSwitch and its panel.
;
; Usage:
;   FileDialogs.IsFileDialog()                      the front window is a file dialog (also the hotkeys' condition)
;   FileDialogs.SetDialogPath(folder, dialog := 0)  jump a dialog to a folder
;   FileDialogs.PickFile(path, dialog := 0)         jump a dialog to a file's folder
; The extract windows of WinRAR and 7-Zip (ExtractDialogs) count too, with a jump of their own: the folder only goes into the
; destination box (Enter or OK there would start extracting), and AutoSwitch / AutoSwitchOnOpen leave them alone (see QuickSwitch._Check)
;===============================================================================

class FileDialogs {
    ; Only real open / save file dialogs count (a file name box + a file list + buttons or an address bar)
    static IsFileDialog(*) => FileDialogs._IsFileDialogWindow(WinExist("A"))   ; also the hotkeys' condition (HotIf)

    static _IsFileDialogWindow(hwnd) {
        if !(hwnd && DllCall("IsWindowVisible", "Ptr", hwnd) && QuickSwitch._MayBeDialog(hwnd))
            return false
        target := "ahk_id " hwnd, controls := [], title := "", winClass := ""  ; read once for both kinds
        try controls := WinGetControls(target)
        try title := WinGetTitle(target)
        try winClass := WinGetClass(target)
        browse := QuickSwitch.Options.Has("BrowseDialogs") && QuickSwitch.Options["BrowseDialogs"]
        if (browse && FileDialogs.LooksLikeBrowseDialog(winClass, controls, FileDialogs._TreeId(controls, target)))
            return true
        if (QuickSwitch._ExtractEnabled() && QuickSwitch._Matches("ahk_group ALTRunExtractDialogs", hwnd) && FileDialogs.ExtractPathBox(hwnd))
            return true
        return QuickSwitch._Matches("ahk_group ALTRunDialogs", hwnd) && FileDialogs.LooksLikeFileDialog(winClass, controls, title)
    }

    ; Decided by the window's controls (independent of the system language). #32770 is a system dialog: it needs Explorer's file view to count;
    ; one with only a plain list (e.g. the main window of "System Information" msinfo32, with a find box, a list and buttons) doesn't.
    ; Other programs ("other dialogs" added in Preferences) count with a plain list too; those whose controls can't be read (Qt etc.) are decided by title
    static LooksLikeFileDialog(winClass, controls, title := "") {
        if !controls.Length
            return TextTools.Match(title, "i)(open|save|select|import|export|browse|打开|另存|选择|导入|导出)") > 0
        hasNameEdit := FileDialogs._HasControl(controls, "^Edit\d+$")
        hasFileView := (winClass = "#32770")
            ? FileDialogs._HasControl(controls, "^DirectUIHWND\d+$", "^SHELLDLL_DefView\d*$")
            : FileDialogs._HasControl(controls, "^DirectUIHWND\d+$", "^SHELLDLL_DefView\d*$", "^SysListView32\d*$")
        hasButtons  := FileDialogs._HasControl(controls, "^Button\d+$", "^ToolbarWindow32\d+$", "^ComboBoxEx32\d+$", "^ComboBox\d+$")
        return hasNameEdit && hasFileView && hasButtons
    }

    ; The old "Browse for Folder" dialog (SHBrowseForFolder): only a folder tree, no file list, so it is not one
    ; of the dialogs above. Told apart by fixed names, whatever the system language: the new style puts the tree
    ; in a "SHBrowseForFolder ShellNameSpace Control", the classic style has the tree itself with control ID 0x3741.
    ; Turned on by Works in → "Browse for Folder dialogs" (BrowseDialogs); ExcludeWindows applies too
    static IsBrowseWindow(hwnd) {
        controls := [], winClass := ""
        try controls := WinGetControls("ahk_id " hwnd), winClass := WinGetClass("ahk_id " hwnd)
        return FileDialogs.LooksLikeBrowseDialog(winClass, controls, FileDialogs._TreeId(controls, "ahk_id " hwnd))
    }

    static LooksLikeBrowseDialog(winClass, controls, treeId := 0) {
        return winClass = "#32770" && (treeId = 0x3741 || FileDialogs._HasControl(controls, "^SHBrowseForFolder ShellNameSpace Control\d+$"))
    }

    ; Control ID of the dialog's first tree view (0 = none)
    static _TreeId(controls, win) {
        for name in controls
            if (name = "SysTreeView321")
                try return DllCall("GetDlgCtrlID", "Ptr", ControlGetHwnd(name, win), "Int")
        return 0
    }

    ; Select a folder in a "Browse for Folder" dialog's tree, like Listary: BFFM_SETSELECTIONW expands the tree to it.
    ; The message carries a pointer, so the path is first written into the dialog's own process. Doesn't press OK.
    ; A dialog of a program running as administrator can't be reached from ALTRun without administrator rights (its process can't
    ; be opened, and Windows blocks the message, UIPI): SetDialogPath hands those to AdminHelper; one that gets here anyway is logged, with a hint
    static SelectInBrowseDialog(folder, dialog) {
        folder := RTrim(folder, "\")
        if TextTools.Match(folder, "^[A-Za-z]:$")
            folder .= "\"                                                   ; a drive needs its backslash
        pid := 0
        try pid := WinGetPID("ahk_id " dialog)
        if !pid
            return FileDialogs._BrowseFailed(folder, dialog, 0, "its process isn't known")
        if !(process := DllCall("OpenProcess", "UInt", 0x38, "Int", false, "UInt", pid, "Ptr")) {   ; PROCESS_VM_OPERATION | VM_READ | VM_WRITE
            err := A_LastError
            return FileDialogs._BrowseFailed(folder, dialog, pid, "cannot open its process, error " err, err = 5)   ; ERROR_ACCESS_DENIED
        }
        size := (StrLen(folder) + 1) * 2, ok := false, sent := true, reason := "cannot reserve memory in its process"
        remote := DllCall("VirtualAllocEx", "Ptr", process, "Ptr", 0, "UPtr", size, "UInt", 0x3000, "UInt", 0x04, "Ptr")   ; MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE
        if remote {
            DllCall("IsWow64Process", "Ptr", process, "Int*", &wow64 := 0)
            text := Buffer(size), StrPut(folder, text, "UTF-16")
            if (wow64 && remote > 0xFFFFFFFF)                               ; a 32-bit program only sees the low 4 GB
                reason := "the memory is out of reach of a 32-bit program"
            else if !DllCall("WriteProcessMemory", "Ptr", process, "Ptr", remote, "Ptr", text, "UPtr", size, "Ptr", 0)
                reason := "cannot write into its process"
            else {
                try {
                    SendMessage(0x467, 1, remote, , "ahk_id " dialog, , , , 2000)   ; BFFM_SETSELECTIONW (WM_USER + 103); wParam TRUE = lParam is a path
                    ok := true
                } catch
                    sent := false, reason := "the message didn't get through in 2 s"
            }
            if sent                                                         ; not after a timeout: the dialog may still read the path later,
                DllCall("VirtualFreeEx", "Ptr", process, "Ptr", remote, "UPtr", 0, "UInt", 0x8000)   ; from freed memory (a few hundred bytes are left instead). MEM_RELEASE
        }
        DllCall("CloseHandle", "Ptr", process)
        if !ok
            return FileDialogs._BrowseFailed(folder, dialog, pid, reason)
        Logger.Debug("QuickSwitch: selected [" folder "] in the Browse for Folder dialog")
        return true
    }

    ; Logs why the folder couldn't be selected; when the dialog's program runs as administrator and ALTRun doesn't, also tells the user.
    ; denied: opening the process was refused (a normal process can't always ask an elevated one whether it is elevated)
    static _BrowseFailed(folder, dialog, pid, reason, denied := false) {
        Logger.Error("QuickSwitch: cannot select " folder " in the Browse for Folder dialog - " reason)
        if (!A_IsAdmin && pid && (denied || Win.IsElevated(pid)))
            FileDialogs.ElevatedHint(dialog)
        return false
    }

    ; The dialog's program runs as administrator (and ALTRun doesn't, see SetDialogPath): Windows blocks ALTRun's messages and keys to it.
    ; A normal process can't always ask an elevated one whether it is elevated, so being refused the access a jump needs counts too (ERROR_ACCESS_DENIED)
    static IsElevatedDialog(dialog) {
        pid := 0
        try pid := WinGetPID("ahk_id " dialog)
        if !pid
            return false
        if Win.IsElevated(pid)
            return true
        if (process := DllCall("OpenProcess", "UInt", 0x38, "Int", false, "UInt", pid, "Ptr"))   ; PROCESS_VM_OPERATION | VM_READ | VM_WRITE
            return (DllCall("CloseHandle", "Ptr", process), false)
        return A_LastError = 5
    }

    ; "This dialog's program runs as administrator" and where to turn on the jump for it; once per dialog (AutoSwitch would repeat it at every switch)
    static _hinted := Map()
    static ElevatedHint(dialog) {
        if FileDialogs._hinted.Has(dialog)
            return false
        if (FileDialogs._hinted.Count > 50)                                 ; window handles of closed dialogs
            FileDialogs._hinted := Map()
        FileDialogs._hinted[dialog] := true
        App.Notify(I18n.T("QuickSwitch.ElevatedDialog"), 6000)
        return false
    }

    ; The destination box (a combo box) of a WinRAR / 7-Zip extract window, 0 if the window isn't one. Told apart by control IDs, whatever the
    ; language (control trees read on Windows CI, WinRAR 7 / 7-Zip 25): see ExtractBoxIn
    static ExtractPathBox(dialog) {
        exe := "", controls := []
        if DllCall("GetWindowThreadProcessId", "Ptr", dialog, "UInt*", &pid := 0, "UInt")   ; WinGetProcessName doesn't see hidden windows
            try exe := ProcessGetName(pid)
        if !(TextTools.Match(exe, "i)^(7zG|7zFM|WinRAR)\.exe$") && QuickSwitch.ClassOf(dialog) = "#32770")
            return 0
        try for control in WinGetControlsHwnd("ahk_id " dialog)
            controls.Push({Hwnd: control, Class: QuickSwitch.ClassOf(control), Id: DllCall("GetDlgCtrlID", "Ptr", control, "Int"), Parent: DllCall("GetParent", "Ptr", control, "Ptr")})
        return FileDialogs.ExtractBoxIn(exe, dialog, controls)
    }

    ; controls: {Hwnd, Class, Id, Parent} of every control in the window
    ;   7-Zip "Extract" (7zG, or 7zFM's Extract button): path combo box 100, overwrite mode 103. "Add to Archive" has those IDs too, and its
    ;     archive format box 104, which Extract lacks
    ;   7-Zip File Manager "Copy" / "Move" (F5 / F6, also how it extracts from an archive open there): label 100, path combo box 101, "..." 102
    ;   WinRAR "Extraction path and options": a property sheet; its General page holds the path combo box 101 next to the folder tree 104
    static ExtractBoxIn(exe, dialog, controls) {
        find(parent, id, wanted := "") {
            for control in controls
                if (control.Parent = parent && control.Id = id)
                    return (wanted = "" || control.Class = wanted) ? control.Hwnd : 0
            return 0
        }
        if (exe = "7zG.exe" || exe = "7zFM.exe") {
            if ((box := find(dialog, 100, "ComboBox")) && find(dialog, 103, "ComboBox") && !find(dialog, 104))
                return box
            if (exe = "7zFM.exe" && (box := find(dialog, 101, "ComboBox")) && find(dialog, 102, "Button") && find(dialog, 100, "Static"))
                return box
        } else if (exe = "WinRAR.exe") {
            for control in controls
                if (control.Id = 101 && control.Class = "ComboBox" && find(control.Parent, 104, "SysTreeView32"))
                    return control.Hwnd
        }
        return 0
    }

    ; An extract window: the folder into its destination box, as typing does (the program reads the box when OK is clicked). No Enter and
    ; no OK: that would start extracting, which the user does after checking the other options
    static FillExtractPath(folder, dialog) {
        if !(box := FileDialogs.ExtractPathBox(dialog))
            return false
        editBox := DllCall("GetWindow", "Ptr", box, "UInt", 5, "Ptr")       ; GW_CHILD: the combo box's edit box
        try ok := editBox && FileDialogs._SetBoxText(editBox, folder)
        catch
            ok := false
        if !ok
            Logger.Error("QuickSwitch: the extract window's destination box didn't take " folder)
        else if Logger.Enabled
            Logger.Debug("QuickSwitch: filled [" folder "] into the extract window's destination box")
        return ok
    }

    static _HasControl(controls, patterns*) {
        for name in controls
            for pattern in patterns
                if TextTools.Match(name, "i)" pattern)
                    return true
        return false
    }

    ; Choosing a file in a dialog: for open and save dialogs alike, jump to its folder (without opening or saving for the user)
    static PickFile(filePath, dialog := 0) {
        SplitPath(filePath, , &dir)
        if (dir != "")
            FileDialogs.SetDialogPath(RTrim(dir, "\") "\", dialog)
    }

    ; fill the old name back only when the file name box held a normal file name (not when empty, a path or a filter like *.txt)
    static ShouldRestoreName(saved) {
        saved := Trim(saved)
        return saved != "" && !TextTools.Match(saved, "[\\/:*?]")
    }

    ; dialog: the dialog to jump in (the current window by default). A dialog of a program running as administrator, with ALTRun not
    ; running as one, is jumped by the helper with administrator rights (AdminHelper.Jump): Windows blocks ALTRun's messages and keys there
    ; A network folder must answer within ExplorerTabs.NetworkTimeout: asking an offline server waits ~21 s on this thread
    static SetDialogPath(folder, dialog := 0) {
        if (folder = "")
            return
        if !ExplorerTabs.Reachable(folder)
            return Logger.Debug("QuickSwitch: the network folder [" folder "] didn't answer within " ExplorerTabs.NetworkTimeout " ms")
        if !FileExist(folder)
            return
        if !dialog
            dialog := WinExist("A")
        if !(dialog && WinExist("ahk_id " dialog))
            return
        Usage.Count("QuickSwitch")
        if (!A_IsAdmin && FileDialogs.IsElevatedDialog(dialog))
            return AdminHelper.Jump(folder, dialog)
        return FileDialogs.Jump(folder, dialog)
    }

    ; The jump itself, with the method that suits the dialog (DialogJumpMethod); also run by the helper for administrator windows.
    ; Before simulating keyboard input, make sure it is the foreground window, otherwise the text would go to another window (e.g.
    ; switching back to the dialog failed and the panel is still in front). false when it surely didn't go there
    ; One jump at a time (Critical): a jump that types takes a moment, longer when SendInput falls back to single key events (another
    ; program's keyboard hook), and a second click in the panel meanwhile started a second jump whose path was typed into the middle
    ; of the first (#600). Now it waits until this one is done, then runs
    static Jump(folder, dialog) {
        wasCritical := A_IsCritical
        Critical
        try return FileDialogs._Jump(folder, dialog)
        finally Critical(wasCritical)
    }

    static LastTyped := false                                               ; whether the last jump typed the folder (for the tests)

    static _Jump(folder, dialog) {
        FileDialogs.LastTyped := false
        method := FileDialogs.DialogJumpMethod(dialog)
        if (method = "fill")                                                ; an extract window: Enter there would extract
            return FileDialogs.FillExtractPath(folder, dialog)
        if (method = "browse")                                              ; folder tree only: select it there, pressing Enter would mean OK
            return FileDialogs.SelectInBrowseDialog(folder, dialog)
        if (method = "message") {
            if FileDialogs._JumpByMessage(folder, dialog)
                return
            method := "address"                                             ; the box didn't take the folder: as if there were none
        }
        className := ""
        try className := WinGetClass("ahk_id " dialog)
        if (method = "edit") {
            try {
                saved := ""
                try saved := ControlGetText("Edit1", "ahk_id " dialog)         ; what the file name box held (e.g. a name Save As filled in)
                ControlFocus("Edit1", "ahk_id " dialog)
                ControlSetText(folder, "Edit1", "ahk_id " dialog)
                if (ControlGetText("Edit1", "ahk_id " dialog) != folder)    ; Enter with anything else in the box would open / save that
                    throw Error("the file name box didn't take the folder")
                ControlSend("{Enter}", "Edit1", "ahk_id " dialog)
                if FileDialogs.ShouldRestoreName(saved) {                   ; put it back after the jump, so the suggested name isn't lost
                    Sleep(150)
                    try ControlSetText(saved, "Edit1", "ahk_id " dialog)
                }
                return
            }
            method := (className = "#32770") ? "address" : "type"           ; the box refused: as if there were none
        }
        if !WinActive("ahk_id " dialog)                                     ; typing goes to the foreground window: only into the dialog
            return false
        if (method = "address") {                                           ; a system dialog without the box: its address bar
            SendInput("!d")
            Sleep(60)
            if FileDialogs._FillFocusedBox(folder, dialog) {                ; the whole folder at once, nothing typed
                SendInput("{Enter}")
                return FileDialogs._WaitForFolder(folder, dialog)
            }
        }
        FileDialogs.LastTyped := true
        SendText(folder)                                                    ; dialogs a program draws itself (WPS / Qt, GTK, Java...): the
        SendInput("{Enter}")                                                ; file name box has the focus; a folder path + Enter goes there
        if (method = "address")
            return FileDialogs._WaitForFolder(folder, dialog)
    }

    ; After Enter in the address bar: wait until the dialog shows the folder (at most WaitForFolderMs), so the jump is over when Jump returns.
    ; A jump right after it (a second click in the panel) otherwise ran while the dialog was still going to this folder: when it got there
    ; it put this folder back into the address bar over the next one, and the dialog stayed here (measured on Windows CI)
    static WaitForFolderMs := 1000

    static _WaitForFolder(folder, dialog) {
        folder := RTrim(folder, "\"), started := A_TickCount
        while (A_TickCount - started < FileDialogs.WaitForFolderMs) {
            if !WinExist("ahk_id " dialog)
                return
            if (FileDialogs.DialogFolder(dialog) = folder)
                return
            Sleep(30)
        }
        Logger.Debug("QuickSwitch: the dialog didn't show [" folder "] within " FileDialogs.WaitForFolderMs " ms after Enter in the address bar")
    }

    ; Puts the folder into the dialog's focused edit box (the address bar after Alt+D) with WM_SETTEXT; false if the focus isn't on an
    ; edit box or it doesn't hold exactly the folder afterwards (then the folder is typed)
    static _FillFocusedBox(folder, dialog) {
        try {
            box := ControlGetFocus("ahk_id " dialog)
            if !(box && WinGetClass("ahk_id " box) = "Edit")
                return false
            ControlSetText(folder, box)
            return ControlGetText(box) = folder
        }
        return false
    }

    ; A system dialog's file name box gets the folder and the dialog its OK command, all as messages, no keys: the jump works while Alt or a
    ; mouse button is still held (AutoSwitch right after Alt+Tab or a click on the title bar, the Ctrl of the hotkey), during the title bar's
    ; move loop and wherever the focus is. Keys were not safe there: with Alt held, a Save As took the posted Enter as OK for the name it
    ; already had and saved that file in its old folder (measured on Windows Server 2025).
    ;   1. WM_SETTEXT puts the folder into the box, and it is read back: OK with anything else in it would open or save that
    ;   2. the box's parent is told the text changed, as typing does (CBN_EDITCHANGE from the combo box of an Explorer-style dialog,
    ;      EN_CHANGE from a plain edit box): the Explorer-style dialog keeps the name it last saw and would otherwise use that one
    ;   3. WM_COMMAND IDOK, what Enter in a dialog becomes (IsDialogMessage): with a folder in the box the dialog goes there, for open and
    ;      save dialogs alike. Sent, not posted, so the old name goes back only after the dialog has read the box
    static _JumpByMessage(folder, dialog) {
        target := "ahk_id " dialog, saved := ""
        if Logger.Enabled
            Logger.Debug("QuickSwitch: jump to [" folder "] by message" FileDialogs._DialogState(dialog))
        try {
            box := ControlGetHwnd("Edit1", target)
            try saved := ControlGetText(box)
            if !FileDialogs._SetBoxText(box, folder) {
                Logger.Debug("QuickSwitch: the file name box didn't take the folder, or the dialog didn't take the change notice")
                return false
            }
        } catch as e {
            Logger.Debug("QuickSwitch: no file name box - " e.Message)
            return false
        }
        started := A_TickCount
        if !(answered := FileDialogs._SendCommand(dialog, DllCall("GetDlgItem", "Ptr", dialog, "Int", 1, "Ptr"), 0, 1))   ; BN_CLICKED from IDOK
            Logger.Warn("QuickSwitch: the dialog didn't answer the OK command in 2 s")   ; it may still go there: no second try with keys
        else if Logger.Enabled
            Logger.Debug("QuickSwitch: OK command answered in " (A_TickCount - started) " ms" FileDialogs._DialogState(dialog))
        if (answered && FileDialogs.ShouldRestoreName(saved)) {             ; put back a name Save As suggested, so it isn't lost
            Loop 6 {                                                        ; the dialog may put it back itself after going there: leave that
                Sleep(50)
                text := ""
                try text := ControlGetText(box)
                if (text != folder && text != "")
                    return true
            }
            try FileDialogs._SetBoxText(box, saved)
        }
        return true
    }

    ; The folder a system dialog shows, from its address bar ("Address: C:\Docs", the label in the UI language); "" if it can't be read.
    ; The tests check a jump with it
    static DialogFolder(dialog) {
        text := ""
        if (box := FileDialogs.ExtractPathBox(dialog)) {                   ; an extract window: its destination, if it is a folder already
            try text := RTrim(ControlGetText(box), "\")
            return (text != "" && DirExist(text)) ? text : ""
        }
        try text := WinGetText("ahk_id " dialog)
        Loop Parse, text, "`n", "`r"
            if TextTools.Match(A_LoopField, "^[^:：\\]{1,40}[:：]\s*((?:[A-Za-z]:\\|\\\\)[^\r\n]*)$", &m)
                return RTrim(m[1], "\")
        return ""
    }

    ; For the debug log: the file name box, the focused control, the dialog thread's state (GUITHREADINFO flags) and whether it is in front
    static _DialogState(dialog) {
        box := "", focus := "", flags := -1
        try box := ControlGetText("Edit1", "ahk_id " dialog)
        info := Buffer(A_PtrSize = 8 ? 72 : 48, 0)
        NumPut("UInt", info.Size, info)
        if ((tid := DllCall("GetWindowThreadProcessId", "Ptr", dialog, "Ptr", 0, "UInt")) && DllCall("GetGUIThreadInfo", "UInt", tid, "Ptr", info)) {
            flags := NumGet(info, 4, "UInt")
            if (focused := NumGet(info, 8 + A_PtrSize, "Ptr"))             ; hwndFocus
                try focus := WinGetClass("ahk_id " focused)
        }
        return Format(" (0x{:X}), box [{}], focus {}, thread flags 0x{:X}, in front {}", dialog, box, focus, flags, DllCall("GetForegroundWindow", "Ptr") = dialog)
    }

    ; Puts text into a dialog's file name box and tells its parent, as typing does (see _JumpByMessage, step 2); false if the box
    ; doesn't hold exactly that text afterwards or the dialog didn't take the notice
    static _SetBoxText(box, text) {
        ControlSetText(text, box)
        if (ControlGetText(box) != text)
            return false
        parent := DllCall("GetParent", "Ptr", box, "Ptr")
        if (WinGetClass("ahk_id " parent) = "ComboBox")
            return FileDialogs._SendCommand(DllCall("GetParent", "Ptr", parent, "Ptr"), parent, 5)   ; CBN_EDITCHANGE
        return FileDialogs._SendCommand(parent, box, 0x300)                                       ; EN_CHANGE
    }

    ; WM_COMMAND (code, the control's id) from control to window, sent with a 2 s limit; false if it didn't get through
    static _SendCommand(window, control, code, id := "") {
        id := (id != "") ? id : DllCall("GetDlgCtrlID", "Ptr", control, "Int")
        return window && DllCall("SendMessageTimeoutW", "Ptr", window, "UInt", 0x111, "Ptr", (code << 16) | (id & 0xFFFF), "Ptr", control
            , "UInt", 0x2, "UInt", 2000, "Ptr", 0, "Ptr") != 0              ; SMTO_ABORTIFHUNG
    }

    ; How the folder goes into this dialog, see JumpMethod; "browse" = a folder tree dialog (SelectInBrowseDialog), "fill" = an extract
    ; window (FillExtractPath)
    static DialogJumpMethod(dialog) {
        if FileDialogs.ExtractPathBox(dialog)
            return "fill"
        if FileDialogs.IsBrowseWindow(dialog)
            return "browse"
        className := "", hasEdit := false
        try className := WinGetClass("ahk_id " dialog)
        if (className = "#32770" && QuickSwitch._hasAddressBar && QuickSwitch._Matches("ahk_group ALTRunAddressBarDialogs", dialog))
            return "address"
        try hasEdit := ControlGetHwnd("Edit1", "ahk_id " dialog) != 0
        return FileDialogs.JumpMethod(className, hasEdit, DllCall("GetDlgItem", "Ptr", dialog, "Int", 1, "Ptr") != 0)
    }

    ; How to put the folder into a dialog:
    ;   message = a system dialog (#32770) with a file name box and an OK button (IDOK): the folder and OK as messages (see _JumpByMessage)
    ;   edit    = into its file name box (Edit1) + Enter: programs with real edit controls that aren't system dialogs
    ;   address = a system dialog (#32770) without that box: Alt+D (address bar), the folder put in (typed if that fails), Enter.
    ;             Also the standard dialogs of the programs in AddressBarDialogs (DialogJumpMethod): they take a folder in the file
    ;             name box + OK as the item chosen, not as a place to go (WeCom WXWork.exe sends that folder to the chat, #600);
    ;             Enter in the address bar only goes there
    ;   type    = a dialog a program draws itself (WPS / Qt, GTK, Java...): type into the focused file name box, Enter.
    ;             No Alt+D there: it isn't the address bar and may press some button of the dialog
    static JumpMethod(className, hasEdit, hasOk := false) {
        if (hasEdit && hasOk && className = "#32770")
            return "message"
        if hasEdit
            return "edit"
        return (className = "#32770") ? "address" : "type"
    }
}
