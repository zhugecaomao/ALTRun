;===============================================================================
; QuickSwitch.ahk - quick folder jumps in open / save dialogs (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like Listary's "Quick Switch": in an Open / Save dialog, the hotkey (Ctrl+G by default) jumps the dialog to the
; folder open in the file manager used last: Total Commander, Directory Opus or File Explorer (see FileManagerFolder).
; AutoSwitch = 1: switching from one of them to a dialog jumps automatically.
; The folder panel (ShowPanel, on by default, like Listary's Quick Switch window): as soon as a dialog appears, a search box +
; folder list (favorite folders, both panels of every TC window, Explorer windows, recently used folders) sits below it, as wide as the dialog; one click jumps there, no hotkey to remember;
; the search box also searches folders (with PanelSearch = all, files too: choosing a file jumps to its folder). Hidden when the dialog isn't in front,
; refreshed when back in the dialog (a folder change in TC shows right away). The colors follow ALTRun's theme.
; Hotkey (MenuHotkey, default Ctrl+Shift+G): the caret jumps to the panel's search box to choose with the keyboard; shows the panel for the moment when automatic showing is off.
; In a dialog of a program running as administrator ALTRun's keyboard hook gets no keys (UIPI): the helper with administrator rights catches both
; hotkeys there and tells ALTRun which was pressed (AdminHelper.HotkeyPressed).
; Favorite folders (Favorites): the ☆ button right of the panel's search box shows them in a menu, each with one level of its subfolders;
; the search box finds them too. Right-click a row to add it, or add the dialog's current folder (also Ctrl+D in the panel).
;
; Settings (ALTRun.json -> Extensions.QuickSwitch):
;   Enabled / Hotkey (the jump hotkey, for all three) / MenuHotkey / Favorites (folder paths, environment variables allowed) / RecentFolders / ShowPanel / PanelSearch / AutoSwitch /
;   StandardDialogs (ahk_class #32770) / BrowseDialogs / DialogWindows / ExcludeWindows / AutoSwitchExclude
;   The window lists are arrays of window conditions (ahk_class / ahk_exe / part of the title)
;
; This file: start-up, the hotkeys and watching which window is in front (Windows events). The other parts, each its own class:
;   QuickSwitchPanel.ahk      the folder panel below the dialog (search box + folder list)
;   QuickSwitchFavorites.ahk  the folders in the panel, the favorite folders and the ☆ menu
;   FileDialogs.ahk           recognizing open / save / "Browse for Folder" dialogs and jumping them to a folder
;   FileManagers.ahk          the current folder of Total Commander, Directory Opus and File Explorer; recent folders
;
; Usage:
;   QuickSwitch.Init(AppSettings.Extension("QuickSwitch"))    at start-up
;   QuickSwitch.SyncFileManagerPath()    the jump hotkey: the dialog goes to the folder of the file manager used last
;===============================================================================

class QuickSwitch {
    static Options := Map(), _hasExclude := false

    static Init(options) {
        QuickSwitch.Options := options
        if !options["Enabled"]
            return
        QuickSwitch.AddGroups(options)

        HotIf((*) => FileDialogs.IsFileDialog())
        try {
            if (options["Hotkey"] != "")
                Hotkey(options["Hotkey"], (*) => QuickSwitch.SyncFileManagerPath())
            if (options["MenuHotkey"] != "")
                Hotkey(options["MenuHotkey"], (*) => QuickSwitch.FocusPanel())
        } catch as e {
            Logger.Error("QuickSwitch: cannot register hotkeys - " e.Message)
        }
        OnMessage(0x100, (p*) => QuickSwitch._OnKeyDown(p*))                ; WM_KEYDOWN: ↑ ↓ Enter Esc F4 in the panel's search box
        OnMessage(0x104, (p*) => QuickSwitch._OnKeyDown(p*))                ; WM_SYSKEYDOWN: Alt+↓ there

        QuickSwitch._StartWatching()
    }

    ; The window groups of the dialog lists (FileDialogs.IsFileDialog reads them); also used by the helper for administrator windows (AdminHelper)
    static AddGroups(options) {
        if options["StandardDialogs"]
            GroupAdd("ALTRunDialogs", "ahk_class #32770")
        for window in options["DialogWindows"]
            GroupAdd("ALTRunDialogs", window)
        QuickSwitch._hasExclude := false
        for window in options["ExcludeWindows"]
            GroupAdd("ALTRunDialogExclude", window), QuickSwitch._hasExclude := true
        for window in options["AutoSwitchExclude"]
            GroupAdd("ALTRunAutoSwitchExclude", window)
    }

    ;---------------------------------------------------------------------------
    ; Watching dialogs with Windows events only, no polling (no CPU while idle; the panel as soon as the dialog is ready):
    ;   EVENT_SYSTEM_FOREGROUND, every process: another window came to the front
    ;   One window at a time, with events of its UI thread only, for what it is now (Phases):
    ;     candidate  the window in front may be a dialog (DialogWindows / #32770, also while still hidden): EVENT_OBJECT_DESTROY / SHOW /
    ;                HIDE (the dialog or its controls appear or go: some dialogs come to the front before they are shown, or build their
    ;                file list after that) and EVENT_OBJECT_NAMECHANGE (dialogs drawn by their program, Qt..., are told by their title,
    ;                which may be set later); a burst (a dialog creating its controls) is checked once, CheckDelayMs after it
    ;     dialog     recognized: only EVENT_SYSTEM_MOVESIZESTART / END. Object events aren't needed then, and browsing folders raises a
    ;                great many of them (every item of the file list). While it is dragged or resized, EVENT_OBJECT_LOCATIONCHANGE of
    ;                that thread moves the panel along (see _Drag)
    ;     settle     any other window, for SettleMs after it came to the front: EVENT_OBJECT_NAMECHANGE (a window listed by a title
    ;                that is set a moment later)
    ;     panel      the caret is in the panel's search box: EVENT_OBJECT_DESTROY of the dialog's thread (closed by its program)
    ; Only if a hook can't be set, it checks every PollMs instead (all the time, or while that window is watched)
    ;---------------------------------------------------------------------------
    static PollMs := 250, SettleMs := 1500, CheckDelayMs := 50, DragMs := 15
    static Phases := Map("candidate", [[0x8001, 0x8003], [0x800C, 0x800C]], "dialog", [[0xA, 0xB]], "settle", [[0x800C, 0x800C]], "panel", [[0x8001, 0x8001]])
    static _foregroundHook := 0, _threadHooks := [], _watched := 0, _phase := "", _eventCallback := 0, _dragHook := 0
    static _alwaysPoll := false, _threadPoll := false, _polling := false
    static _foreground := 0, _foregroundAt := 0, _fromManager := false
    ; Clicking a dialog may bring the program that owns it (e.g. Word) to the front for a moment first: when the dialog follows within
    ; OwnerPassMs, the window before the program still counts as the window before the dialog
    static OwnerPassMs := 500
    ; AutoSwitch jumps at once when the dialog takes a message jump (see SetDialogPath); a jump that types waits until the switch is over
    ; (see _InputSettled), at most AutoSwitchWaitMs
    static AutoSwitchWaitMs := 3000
    static _checkTimer := "", _settleTimer := "", _pollTimer := "", _placeTimer := "", _dragTimer := ""

    static _StartWatching() {
        QuickSwitch._checkTimer := Logger.Watch("quick switch check", () => QuickSwitch._Check())   ; one-shot (events), the end of settle
        QuickSwitch._settleTimer := Logger.Watch("quick switch settle", () => QuickSwitch._Check())  ; and repeating (polling) need separate timers
        QuickSwitch._pollTimer := Logger.Watch("quick switch watch", () => QuickSwitch._Check())
        QuickSwitch._placeTimer := () => QuickSwitchPanel._panelFor && QuickSwitchPanel._PlacePanel(QuickSwitchPanel._panelFor)
        QuickSwitch._dragTimer := () => QuickSwitchPanel._panelFor && QuickSwitchPanel._PlacePanel(QuickSwitchPanel._panelFor)   ; only if the drag hook can't be set
        if !(QuickSwitch._foregroundHook := QuickSwitch._HookEvent(0x3, 0x3)) {   ; EVENT_SYSTEM_FOREGROUND, every process
            Logger.Error("QuickSwitch: cannot watch foreground changes, checking every " QuickSwitch.PollMs " ms")
            QuickSwitch._alwaysPoll := true
        }
        SetTimer(QuickSwitch._checkTimer, -1)                               ; the window in front now
    }

    static _HookEvent(eventMin, eventMax, pid := 0, tid := 0) {
        if !QuickSwitch._eventCallback
            QuickSwitch._eventCallback := CallbackCreate((p*) => QuickSwitch._OnEvent(p*), , 7)
        return DllCall("SetWinEventHook", "UInt", eventMin, "UInt", eventMax, "Ptr", 0, "Ptr", QuickSwitch._eventCallback, "UInt", pid, "UInt", tid, "UInt", 0, "Ptr")   ; WINEVENT_OUTOFCONTEXT
    }

    ; Keep the callback short: the work runs from a timer, after the event is handled
    static _OnEvent(hook, event, hwnd, idObject, idChild, thread, time) {
        if (event = 0x3)                                                    ; EVENT_SYSTEM_FOREGROUND
            return SetTimer(QuickSwitch._checkTimer, -1)
        if (event = 0xA || event = 0xB) {                                   ; EVENT_SYSTEM_MOVESIZESTART / END: the dialog is dragged or resized
            if (hwnd && hwnd = QuickSwitchPanel._panelFor)
                QuickSwitch._Drag(event = 0xA)
            return
        }
        if (idObject != 0 || idChild != 0 || !hwnd)                         ; windows only (OBJID_WINDOW), not the caret, the cursor or parts of controls
            return
        if (event = 0x800B)                                                 ; EVENT_OBJECT_LOCATIONCHANGE while the dialog is dragged
            return (hwnd = QuickSwitchPanel._panelFor) ? SetTimer(QuickSwitch._placeTimer, -1) : 0
        SetTimer(QuickSwitch._checkTimer, -QuickSwitch.CheckDelayMs)
    }

    ; The UI thread of this window (one window at a time), with the events of its phase (see Phases above)
    static _WatchWindow(hwnd, phase) {
        if (hwnd = QuickSwitch._watched && phase = QuickSwitch._phase)
            return
        QuickSwitch._UnwatchWindow()
        QuickSwitch._watched := hwnd, QuickSwitch._phase := phase
        ranges := QuickSwitch.Phases[phase]
        if (tid := DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "UInt*", &pid := 0, "UInt"))
            for range in ranges
                if (hook := QuickSwitch._HookEvent(range[1], range[2], pid, tid))
                    QuickSwitch._threadHooks.Push(hook)
        if (QuickSwitch._threadHooks.Length < ranges.Length) {
            QuickSwitch._threadPoll := true
            Logger.Warn("QuickSwitch: cannot watch the events of a window (" phase "), checking every " QuickSwitch.PollMs " ms meanwhile")
        }
    }

    static _UnwatchWindow() {
        for hook in QuickSwitch._threadHooks
            DllCall("UnhookWinEvent", "Ptr", hook)
        QuickSwitch._threadHooks := [], QuickSwitch._watched := 0, QuickSwitch._phase := "", QuickSwitch._threadPoll := false
        QuickSwitch._Drag(false, false)
    }

    ; The dialog is dragged or resized (on = MOVESIZESTART): the location changes of its thread move the panel along, and it is placed once
    ; more at the end. Every DragMs instead if that hook can't be set
    static _Drag(on, place := true) {
        if QuickSwitch._dragHook
            DllCall("UnhookWinEvent", "Ptr", QuickSwitch._dragHook), QuickSwitch._dragHook := 0
        if (QuickSwitch._dragTimer != "")                                   ; not before _StartWatching
            SetTimer(QuickSwitch._dragTimer, 0)
        if (on && (tid := DllCall("GetWindowThreadProcessId", "Ptr", QuickSwitchPanel._panelFor, "UInt*", &pid := 0, "UInt"))
            && !(QuickSwitch._dragHook := QuickSwitch._HookEvent(0x800B, 0x800B, pid, tid)))   ; EVENT_OBJECT_LOCATIONCHANGE
            SetTimer(QuickSwitch._dragTimer, QuickSwitch.DragMs)
        if (place && QuickSwitch._placeTimer != "")
            SetTimer(QuickSwitch._placeTimer, -1)
    }

    static _Poll(on) {
        if (on = QuickSwitch._polling || QuickSwitch._pollTimer = "")
            return
        QuickSwitch._polling := on
        SetTimer(QuickSwitch._pollTimer, on ? QuickSwitch.PollMs : 0)
    }

    ; The window (default: the one in front) matches the dialog lists: it may be a dialog, or become one once it is shown and its controls or title are there
    static _MayBeDialog(hwnd := 0) {
        if !(hwnd := hwnd || DllCall("GetForegroundWindow", "Ptr"))
            return false
        browse := QuickSwitch.Options.Has("BrowseDialogs") && QuickSwitch.Options["BrowseDialogs"]
        if !(QuickSwitch._Matches("ahk_group ALTRunDialogs", hwnd) || browse && QuickSwitch._Matches("ahk_class #32770", hwnd))
            return false
        return !(QuickSwitch._hasExclude && QuickSwitch._Matches("ahk_group ALTRunDialogExclude", hwnd))
    }

    ; The window matches the window conditions, hidden or not
    static _Matches(criteria, hwnd) {
        hidden := A_DetectHiddenWindows
        DetectHiddenWindows(true)
        try return WinExist(criteria " ahk_id " hwnd) != 0
        finally DetectHiddenWindows(hidden)
    }

    ; Windows that are in front only for a moment while switching (taskbar, Alt+Tab): they don't count as "the window before" for AutoSwitch.
    ; Also any window already hidden or gone when the event is handled: Windows 11's Alt+Tab window hides right after the switch, and
    ; read with hidden windows not detected its class was "", so it counted as the window before the dialog and AutoSwitch didn't jump
    static IsPassingWindow(hwnd, className) => QuickSwitch.IsTransientWindow(className) || !DllCall("IsWindowVisible", "Ptr", hwnd)

    ; A window's class, also when it is hidden (WinGetClass doesn't find hidden windows unless DetectHiddenWindows is on); "" if it is gone
    static ClassOf(hwnd) {
        name := Buffer(512)
        return DllCall("GetClassNameW", "Ptr", hwnd, "Ptr", name, "Int", 256) ? StrGet(name) : ""
    }

    static IsTransientWindow(className) => RegExMatch(className, "i)^(Shell_TrayWnd|Shell_SecondaryTrayWnd|MultitaskingViewFrame|XamlExplorerHostIslandWindow|TaskSwitcherWnd|ForegroundStaging)$") > 0

    ; WM_KEYDOWN rather than a hotkey: while a Chinese IME is composing (not yet committed) keys are VK_PROCESSKEY and aren't taken for Enter / Esc
    ; compare only the window handles noted: after the panel closes its control objects are destroyed and their Hwnd can't be read
    ; Alt+↓ or F4 opens the ☆ menu, as they open a drop-down list (a combo box) in Windows
    static _OnKeyDown(wParam, lParam, msg, hwnd) {
        if !(hwnd && (hwnd = QuickSwitchPanel._searchHwnd || hwnd = QuickSwitchPanel._listHwnd))
            return
        if (msg = 0x104) {                                                  ; WM_SYSKEYDOWN: only Alt+↓ is ours, other Alt keys go on (Alt+F4...)
            if (wParam != 0x28)
                return
            QuickSwitchFavorites.ShowFavoritesMenu()
            return 0
        }
        switch wParam {
            case 0x26: QuickSwitchPanel.PanelKey("Up")
            case 0x28: QuickSwitchPanel.PanelKey("Down")
            case 0x73:                                                      ; F4 (alone)
                if (GetKeyState("Ctrl") || GetKeyState("Shift"))
                    return
                QuickSwitchFavorites.ShowFavoritesMenu()
            case 0x0D: QuickSwitchPanel.PanelKey("Enter")
            case 0x1B: QuickSwitchPanel.PanelKey("Escape")
            case 0x44:                                                      ; Ctrl+D: the dialog's current folder becomes a favorite (as bookmarking in a browser)
                if !(GetKeyState("Ctrl") && !GetKeyState("Alt") && !GetKeyState("Shift"))
                    return
                QuickSwitchFavorites.AddDialogFolder()
            default:   return
        }
        return 0
    }

    ; On window events (see above): with AutoSwitch, jump when switching from TC to the dialog; show / hide / position the folder panel;
    ; watch the events of the window in front while it may be a dialog
    static _Check() {
        if !(active := DllCall("GetForegroundWindow", "Ptr"))               ; read once: every step below is about the same window
            return                                                          ; none for a moment while switching: the next event comes
        if (IsObject(QuickSwitchPanel._panel) && active = QuickSwitchPanel._panel.Hwnd) {   ; typing in the panel's search box: the dialog and panel stay as they are
            if !DllCall("IsWindow", "Ptr", QuickSwitchPanel._panelFor)          ; the dialog was closed (by its program): close the panel too
                QuickSwitchPanel.HidePanel(), QuickSwitch._UnwatchWindow()
            else {
                QuickSwitchPanel._PlacePanel(QuickSwitchPanel._panelFor)             ; a panel shown as only its search box opens its list now (see PanelLayout)
                QuickSwitch._WatchWindow(QuickSwitchPanel._panelFor, "panel")    ; until the caret leaves the panel (a foreground event)
            }
            return QuickSwitch._Poll(QuickSwitch._alwaysPoll || QuickSwitch._threadPoll)
        }
        className := QuickSwitch.ClassOf(active)
        switched := false, previous := QuickSwitch._foreground
        if (active != previous && !QuickSwitch.IsPassingWindow(active, className)) {   ; another window came to the front: was the one before a file manager?
            owner := DllCall("GetWindow", "Ptr", active, "UInt", 4, "Ptr")  ; GW_OWNER
            if !(owner && owner = previous && A_TickCount - QuickSwitch._foregroundAt < QuickSwitch.OwnerPassMs)   ; see OwnerPassMs
                QuickSwitch._fromManager := FileManagers.IsFileManager(previous)
            QuickSwitch._foreground := active, QuickSwitch._foregroundAt := A_TickCount, QuickSwitch._settledSince := 0, switched := true
        }
        isDialog := FileDialogs._IsFileDialogWindow(active)
        if (Logger.Enabled && active != QuickSwitch._loggedForeground)
            QuickSwitch._LogForeground(active, className, isDialog)
        if (isDialog && QuickSwitch.Options["AutoSwitch"] && QuickSwitch._fromManager && !WinExist("ahk_group ALTRunAutoSwitchExclude ahk_id " active))
            QuickSwitch._AutoSwitch(active, switched ? previous : 0)

        if (QuickSwitch._PanelEnabled() || IsObject(QuickSwitchPanel._panel))    ; with automatic showing off, a panel opened by the hotkey still follows the dialog
            QuickSwitchPanel._UpdatePanel(isDialog ? active : 0)
        if isDialog
            QuickSwitch._WatchWindow(active, "dialog")
        else if QuickSwitch._MayBeDialog(active)
            QuickSwitch._WatchWindow(active, "candidate")
        else if (active = QuickSwitch._foreground && (left := QuickSwitch.SettleMs - (A_TickCount - QuickSwitch._foregroundAt)) > 0) {
            QuickSwitch._WatchWindow(active, "settle")
            SetTimer(QuickSwitch._settleTimer, -left)                       ; then the title no longer counts: unwatched
        } else
            QuickSwitch._UnwatchWindow()
        QuickSwitch._Poll(QuickSwitch._alwaysPoll || QuickSwitch._threadPoll)
    }

    ; For the debug log: each window that comes to the front, and whether AutoSwitch counts the window before it as a file manager
    static _loggedForeground := 0
    static _LogForeground(active, className, isDialog) {
        QuickSwitch._loggedForeground := active
        exe := "", ownerClass := ""
        if DllCall("GetWindowThreadProcessId", "Ptr", active, "UInt*", &pid := 0, "UInt")   ; WinGetProcessName doesn't see hidden windows
            try exe := ProcessGetName(pid)
        if (owner := DllCall("GetWindow", "Ptr", active, "UInt", 4, "Ptr"))
            ownerClass := QuickSwitch.ClassOf(owner)
        Logger.Debug(Format("QuickSwitch: in front 0x{:X} {} of {}, owner 0x{:X} {}, passing {}, file dialog {}, from a file manager {}", active, className, exe
            , owner, ownerClass, QuickSwitch.IsPassingWindow(active, className), isDialog, QuickSwitch._fromManager))
    }

    ; AutoSwitch: jump to the file manager's folder once per switch. previous: the window before, on the first check after the switch (debug log)
    static _AutoSwitch(dialog, previous) {
        method := FileDialogs.DialogJumpMethod(dialog)
        if (previous && Logger.Enabled) {
            Logger.Debug("QuickSwitch: AutoSwitch from " QuickSwitch.ClassOf(previous) ", jump by " method ", " (A_TickCount - QuickSwitch._foregroundAt) " ms after the switch")
        }
        if (method = "message" || method = "browse" || QuickSwitch._InputSettled(dialog)) {
            QuickSwitch._fromManager := false                               ; once per switch
            return QuickSwitch.SyncFileManagerPath(true)
        }
        if (A_TickCount - QuickSwitch._foregroundAt < QuickSwitch.AutoSwitchWaitMs)
            return SetTimer(QuickSwitch._checkTimer, -QuickSwitch.CheckDelayMs)   ; the switch is still going on: check again a moment later
        QuickSwitch._fromManager := false
        Logger.Debug("QuickSwitch: AutoSwitch gave up, keys or the mouse still held")
    }

    ; A jump that types (see SetDialogPath) needs the switch to be over. The foreground event comes as soon as the dialog is activated, while
    ; Alt (Alt+Tab) or the mouse button (a click on the title bar) may still be held: keys sent then would be read with Alt still down, or
    ; taken by the title bar's move loop. Settled = no mouse button or modifier held (App.IsKeyDown: also a remote desktop's) and the dialog's thread not in a move / size
    ; or menu loop (GetGUIThreadInfo), then InputSettleMs more: posted keys are read before input still queued (the Alt release)
    static InputSettleMs := 100, _settledSince := 0

    static _InputSettled(dialog) {
        for key in ["LButton", "RButton", "MButton", "Alt", "Ctrl", "Shift", "LWin", "RWin"]
            if App.IsKeyDown(key)
                return (QuickSwitch._settledSince := 0, false)
        info := Buffer(A_PtrSize = 8 ? 72 : 48, 0)                         ; GUITHREADINFO
        NumPut("UInt", info.Size, info)
        tid := DllCall("GetWindowThreadProcessId", "Ptr", dialog, "Ptr", 0, "UInt")
        if (tid && DllCall("GetGUIThreadInfo", "UInt", tid, "Ptr", info) && NumGet(info, 4, "UInt") & 0x1E)   ; GUI_INMOVESIZE | GUI_INMENUMODE | GUI_SYSTEMMENUMODE | GUI_POPUPMENUMODE
            return (QuickSwitch._settledSince := 0, false)
        if !QuickSwitch._settledSince
            QuickSwitch._settledSince := A_TickCount
        return A_TickCount - QuickSwitch._settledSince >= QuickSwitch.InputSettleMs
    }

    static _SearchesFiles() => QuickSwitch.Options.Has("PanelSearch") && QuickSwitch.Options["PanelSearch"] = "all"

    static _PanelEnabled() => QuickSwitch.Options.Has("ShowPanel") && QuickSwitch.Options["ShowPanel"]

    ; Hotkey (MenuHotkey, default Ctrl+Shift+G): the caret jumps to the panel's search box to choose with the keyboard; shows the panel for the moment when automatic showing is off
    static FocusPanel() {
        dialog := WinExist("A")
        if (!IsObject(QuickSwitchPanel._panel) || QuickSwitchPanel._panelFor != dialog) {
            QuickSwitchPanel._BuildPanel(dialog)
            QuickSwitchPanel._PlacePanel(dialog)
        }
        if !IsObject(QuickSwitchPanel._panel)
            return
        try {
            WinActivate("ahk_id " QuickSwitchPanel._panel.Hwnd)
            ControlFocus(QuickSwitchPanel._panelSearch)
        }
    }

    ; The jump hotkey: the folder of the file manager used last (TC, Opus or Explorer, see FileManagerFolder)
    static SyncFileManagerPath(silent := false) {
        manager := FileManagers.FileManagerFolder()
        if !(IsObject(manager) && manager.Path != "")
            return silent ? "" : MsgBox(I18n.T("QuickSwitch.NoFileManager"), App.Name, 48)
        if Logger.Enabled
            Logger.Debug("QuickSwitch: " manager.Tag " is in [" manager.Path "]")
        FileDialogs.SetDialogPath(RTrim(manager.Path, "\") "\")             ; with a trailing backslash: AutoCAD and others need it
    }
}
