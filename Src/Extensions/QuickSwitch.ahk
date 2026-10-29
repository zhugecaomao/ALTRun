;===============================================================================
; QuickSwitch.ahk - quick folder jumps in open / save dialogs (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like Listary's "Quick Switch": in an Open / Save dialog, the hotkey (Ctrl+G by default) jumps the dialog to the
; folder open in the file manager used last: Total Commander, Directory Opus or File Explorer (see FileManagerFolder).
; AutoSwitch = 1: switching from one of them to a dialog jumps automatically.
; The folder panel (ShowPanel, on by default, like Listary's Quick Switch window): as soon as a dialog appears, a search box +
; folder list (both panels of every TC window, Explorer windows, recently used folders) sits below it; one click jumps there, no hotkey to remember;
; the search box also searches folders (with PanelSearch = all, files too: choosing a file jumps to its folder). Hidden when the dialog isn't in front,
; refreshed when back in the dialog (a folder change in TC shows right away). The colors follow ALTRun's theme.
; Hotkey (MenuHotkey, default Ctrl+Shift+G): the caret jumps to the panel's search box to choose with the keyboard; shows the panel for the moment when automatic showing is off.
;
; Settings (ALTRun.json -> Extensions.QuickSwitch):
;   Enabled / Hotkey (the jump hotkey, for all three) / MenuHotkey / RecentFolders / ShowPanel / PanelSearch / AutoSwitch /
;   StandardDialogs (ahk_class #32770) / BrowseDialogs / DialogWindows / ExcludeWindows / AutoSwitchExclude
;   The window lists are arrays of window conditions (ahk_class / ahk_exe / part of the title)
;
; Usage:
;   QuickSwitch.Init(AppSettings.Extension("QuickSwitch"))    at start-up
;   QuickSwitch.FolderOfWindow(hwnd)     current folder of a TC / Opus / Explorer window (read only, for "terminal here")
;   QuickSwitch.MenuFolders()            the folders in the panel [{Path, Group, Tag}]
;   QuickSwitch.RecentFolders(n)         the n most recently used folders (opened files count by their folder)
;===============================================================================

class QuickSwitch {
    static Options := Map(), _hasExclude := false
    static _panel := "", _panelFor := 0, _panelFolders := [], _panelPos := ""

    static Init(options) {
        QuickSwitch.Options := options
        if !options["Enabled"]
            return
        if options["StandardDialogs"]
            GroupAdd("ALTRunDialogs", "ahk_class #32770")
        for window in options["DialogWindows"]
            GroupAdd("ALTRunDialogs", window)
        QuickSwitch._hasExclude := false
        for window in options["ExcludeWindows"]
            GroupAdd("ALTRunDialogExclude", window), QuickSwitch._hasExclude := true
        for window in options["AutoSwitchExclude"]
            GroupAdd("ALTRunAutoSwitchExclude", window)

        HotIf((*) => QuickSwitch.IsFileDialog())
        try {
            if (options["Hotkey"] != "")
                Hotkey(options["Hotkey"], (*) => QuickSwitch.SyncFileManagerPath())
            if (options["MenuHotkey"] != "")
                Hotkey(options["MenuHotkey"], (*) => QuickSwitch.FocusPanel())
        } catch as e {
            Logger.Error("QuickSwitch: cannot register hotkeys - " e.Message)
        }
        OnMessage(0x100, (p*) => QuickSwitch._OnKeyDown(p*))                ; WM_KEYDOWN: ↑ ↓ Enter Esc in the panel's search box

        QuickSwitch._StartWatching()
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
        QuickSwitch._placeTimer := () => QuickSwitch._panelFor && QuickSwitch._PlacePanel(QuickSwitch._panelFor)
        QuickSwitch._dragTimer := () => QuickSwitch._panelFor && QuickSwitch._PlacePanel(QuickSwitch._panelFor)   ; only if the drag hook can't be set
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
            if (hwnd && hwnd = QuickSwitch._panelFor)
                QuickSwitch._Drag(event = 0xA)
            return
        }
        if (idObject != 0 || idChild != 0 || !hwnd)                         ; windows only (OBJID_WINDOW), not the caret, the cursor or parts of controls
            return
        if (event = 0x800B)                                                 ; EVENT_OBJECT_LOCATIONCHANGE while the dialog is dragged
            return (hwnd = QuickSwitch._panelFor) ? SetTimer(QuickSwitch._placeTimer, -1) : 0
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
        if (on && (tid := DllCall("GetWindowThreadProcessId", "Ptr", QuickSwitch._panelFor, "UInt*", &pid := 0, "UInt"))
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
    static _OnKeyDown(wParam, lParam, msg, hwnd) {
        if !(hwnd && (hwnd = QuickSwitch._searchHwnd || hwnd = QuickSwitch._listHwnd))
            return
        switch wParam {
            case 0x26: QuickSwitch.PanelKey("Up")
            case 0x28: QuickSwitch.PanelKey("Down")
            case 0x0D: QuickSwitch.PanelKey("Enter")
            case 0x1B: QuickSwitch.PanelKey("Escape")
            default:   return
        }
        return 0
    }

    ; On window events (see above): with AutoSwitch, jump when switching from TC to the dialog; show / hide / position the folder panel;
    ; watch the events of the window in front while it may be a dialog
    static _Check() {
        if !(active := DllCall("GetForegroundWindow", "Ptr"))               ; read once: every step below is about the same window
            return                                                          ; none for a moment while switching: the next event comes
        if (IsObject(QuickSwitch._panel) && active = QuickSwitch._panel.Hwnd) {   ; typing in the panel's search box: the dialog and panel stay as they are
            if !DllCall("IsWindow", "Ptr", QuickSwitch._panelFor)          ; the dialog was closed (by its program): close the panel too
                QuickSwitch.HidePanel(), QuickSwitch._UnwatchWindow()
            else
                QuickSwitch._WatchWindow(QuickSwitch._panelFor, "panel")    ; until the caret leaves the panel (a foreground event)
            return QuickSwitch._Poll(QuickSwitch._alwaysPoll || QuickSwitch._threadPoll)
        }
        className := QuickSwitch.ClassOf(active)
        switched := false, previous := QuickSwitch._foreground
        if (active != previous && !QuickSwitch.IsPassingWindow(active, className)) {   ; another window came to the front: was the one before a file manager?
            owner := DllCall("GetWindow", "Ptr", active, "UInt", 4, "Ptr")  ; GW_OWNER
            if !(owner && owner = previous && A_TickCount - QuickSwitch._foregroundAt < QuickSwitch.OwnerPassMs)   ; see OwnerPassMs
                QuickSwitch._fromManager := QuickSwitch.IsFileManager(previous)
            QuickSwitch._foreground := active, QuickSwitch._foregroundAt := A_TickCount, QuickSwitch._settledSince := 0, switched := true
        }
        isDialog := QuickSwitch._IsFileDialogWindow(active)
        if (Logger.Enabled && active != QuickSwitch._loggedForeground)
            QuickSwitch._LogForeground(active, className, isDialog)
        if (isDialog && QuickSwitch.Options["AutoSwitch"] && QuickSwitch._fromManager && !WinExist("ahk_group ALTRunAutoSwitchExclude ahk_id " active))
            QuickSwitch._AutoSwitch(active, switched ? previous : 0)

        if (QuickSwitch._PanelEnabled() || IsObject(QuickSwitch._panel))    ; with automatic showing off, a panel opened by the hotkey still follows the dialog
            QuickSwitch._UpdatePanel(isDialog ? active : 0)
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
        method := QuickSwitch.DialogJumpMethod(dialog)
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
    ; taken by the title bar's move loop. Settled = no mouse button or modifier physically down and the dialog's thread not in a move / size
    ; or menu loop (GetGUIThreadInfo), then InputSettleMs more: posted keys are read before input still queued (the Alt release)
    static InputSettleMs := 100, _settledSince := 0

    static _InputSettled(dialog) {
        for key in ["LButton", "RButton", "MButton", "Alt", "Ctrl", "Shift", "LWin", "RWin"]
            if GetKeyState(key, "P")
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
        if (!IsObject(QuickSwitch._panel) || QuickSwitch._panelFor != dialog) {
            QuickSwitch._BuildPanel(dialog)
            QuickSwitch._PlacePanel(dialog)
        }
        if !IsObject(QuickSwitch._panel)
            return
        try {
            WinActivate("ahk_id " QuickSwitch._panel.Hwnd)
            ControlFocus(QuickSwitch._panelSearch)
        }
    }

    ;---------------------------------------------------------------------------
    ; The folder panel (like Listary's Quick Switch window): attached below the dialog, as wide as it;
    ; the search box on top (filters the list first, then finds folders and files with Everything / the built-in index), the list below.
    ; Clicking an item or Enter in the search box jumps there (files: see PickFile); Esc goes back to the dialog
    ;---------------------------------------------------------------------------
    static _UpdatePanel(dialog) {
        if !dialog
            return QuickSwitch.HidePanel()
        if (dialog != QuickSwitch._panelFor) {                              ; a new dialog, or back from another window: list the folders again
            if !QuickSwitch._PanelEnabled()
                return QuickSwitch.HidePanel()
            QuickSwitch._BuildPanel(dialog)
        }
        QuickSwitch._PlacePanel(dialog)
    }

    static HidePanel() {
        if IsObject(QuickSwitch._panel)
            try QuickSwitch._panel.Destroy()
        QuickSwitch._panel := "", QuickSwitch._panelFor := 0, QuickSwitch._panelPos := ""
        QuickSwitch._panelSearch := "", QuickSwitch._panelList := "", QuickSwitch._panelLine := ""   ; the controls are destroyed: no more references
        QuickSwitch._searchHwnd := 0, QuickSwitch._listHwnd := 0
    }

    static PanelRows := 8

    static _BuildPanel(dialog) {
        QuickSwitch.HidePanel()
        QuickSwitch._panelFor := dialog
        QuickSwitch._panelBase := QuickSwitch.MenuFolders(false)
        rect := QuickSwitch._FrameRect(dialog)
        width := QuickSwitch.PanelWidthFor(IsObject(rect) ? rect.W : 600)
        colors := QuickSwitch.PanelColors()                                 ; follows ALTRun's theme (light / dark...)
        ; -DPIScale: the panel is as wide as the dialog, in the dialog's real pixels; the row height is measured in pixels too. Letting AHK scale by DPI again
        ; would make the panel half again as wide as the dialog at 150%, with an empty area below. Fixed margins are converted with Win.Scale
        panel := Gui("-Caption +ToolWindow +AlwaysOnTop +Border -DPIScale", "ALTRun Quick Switch")
        panel.MarginX := 0, panel.MarginY := 0, panel.BackColor := colors.Background
        panel.SetFont("s10 c" colors.Text, QuickSwitch._FontName())
        search := panel.AddEdit("x" Win.Scale(10) " y" Win.Scale(8) " w" (width - Win.Scale(20)) " r1 -Multi -E0x200 Background" colors.Background)
        hint := (QuickSwitch.Options.Has("MenuHotkey") && QuickSwitch.Options["MenuHotkey"] != "") ? "  (" Win.HotkeyLabel(QuickSwitch.Options["MenuHotkey"]) ")" : ""
        DllCall("SendMessage", "Ptr", search.Hwnd, "UInt", 0x1501, "Ptr", 1, "WStr", " " I18n.T(QuickSwitch._SearchesFiles() ? "QuickSwitch.SearchCue" : "QuickSwitch.SearchCueFolders") hint)   ; EM_SETCUEBANNER: gray cue text
        search.OnEvent("Change", (*) => SetTimer(QuickSwitch._searchTimer, -150))
        line := panel.AddText("x0 y+" Win.Scale(6) " w" width " h1 Background" colors.Separator)   ; separator between the search box and the list
        panel.SetFont("s9 c" colors.Text)
        list := panel.AddListView("x0 y+" Win.Scale(2) " w" width " h100 -Hdr -Multi -E0x200 +LV0x10400 Background" colors.Background, ["Name", "Folder", "From"])   ; 0x400 = the full path on hover when it is too long
        icons := IL_Create(2)
        IL_Add(icons, "shell32.dll", 4)                                     ; 1 = folder
        IL_Add(icons, "shell32.dll", 1)                                     ; 2 = file
        list.SetImageList(icons)
        list.OnEvent("Click", (ctrl, row) => QuickSwitch._PanelJump(row))
        list.OnNotify(-12, (ctrl, lParam) => QuickSwitch._OnCustomDraw(lParam))   ; NM_CUSTOMDRAW: the selected row and gray text in the theme's colors
        list.Add("Icon1", "X")                                              ; the row height can only be measured with one row in (LVM_APPROXIMATEVIEWRECT)
        QuickSwitch._rowTop := SendMessage(0x1040, 1, 0, list) >> 16
        QuickSwitch._rowHeight := (SendMessage(0x1040, 2, 0, list) >> 16) - QuickSwitch._rowTop
        list.Delete()
        QuickSwitch._panel := panel, QuickSwitch._panelList := list, QuickSwitch._panelSearch := search, QuickSwitch._panelWidth := width
        QuickSwitch._panelLine := line, QuickSwitch._panelDialogW := IsObject(rect) ? rect.W : 0, QuickSwitch._panelColors := colors
        QuickSwitch._searchHwnd := search.Hwnd, QuickSwitch._listHwnd := list.Hwnd
        if colors.Dark
            try DllCall("uxtheme\SetWindowTheme", "Ptr", list.Hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)   ; dark theme: dark scroll bars and selected row too
        try DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", panel.Hwnd, "UInt", 33, "Int*", 3, "UInt", 4)   ; Win11: small rounded corners
        QuickSwitch._ShowFolders(QuickSwitch._panelBase)
        if (QuickSwitch._placeTimer != "")
            SetTimer(QuickSwitch._placeTimer, -500)                         ; once more: some dialogs move or resize themselves right after they are shown
        if WinExist("ahk_class dopus.lister")                               ; reading Opus takes a moment: show the panel first
            SetTimer(QuickSwitch._AddDOpusFolders.Bind(QuickSwitch, dialog), -1)
    }

    static _AddDOpusFolders(dialog) {
        if (QuickSwitch._panelFor != dialog || !IsObject(QuickSwitch._panelList))
            return
        QuickSwitch._panelBase := QuickSwitch.MenuFolders()
        if (QuickSwitch._panelFor = dialog && IsObject(QuickSwitch._panelSearch) && QuickSwitch._panelSearch.Value = "")   ; not while searching
            QuickSwitch._ShowFolders(QuickSwitch._panelBase)
    }

    static _searchHwnd := 0, _listHwnd := 0, _panelColors := "", _panelBase := [], _panelList := "", _panelSearch := "", _panelLine := "", _panelWidth := 600, _panelDialogW := 0, _rowTop := 20, _rowHeight := 18

    ; Panel width = dialog width (420 ~ 1000, converted by DPI; all real pixels)
    static PanelWidthFor(dialogWidth) => Max(Win.Scale(420), Min(dialogWidth, Win.Scale(1000)))

    ; The panel's colors: background / text / separator colors of ALTRun's current theme; light while no theme is loaded yet
    static PanelColors() {
        get(key, fallback) => ((value := ThemeManager.Get(key)) != "") ? value : fallback
        background := get("Background", "FFFFFF")
        return {Background: background, Text: get("Title", "1F1F1F"), Subtitle: get("Subtitle", "808080"), Separator: get("Separator", "E4E4E4")
              , Selected: get("SelectedBackground", "DDE7F6"), SelectedText: get("SelectedTitle", "000000"), SelectedSubtitle: get("SelectedSubtitle", "4A5568")
              , Dark: QuickSwitch.IsDarkColor(background)}
    }

    ; "RRGGBB" -> a Windows COLORREF (0x00BBGGRR)
    static ColorRef(hex) {
        value := Integer("0x" hex)
        return ((value & 0xFF) << 16) | (value & 0xFF00) | ((value >> 16) & 0xFF)
    }

    ; Each cell sets its own colors: the selected row in the theme's selection color, paths and sources in gray (NMLVCUSTOMDRAW)
    static _OnCustomDraw(lParam) {
        static CDDS_PREPAINT := 0x1, CDDS_ITEMPREPAINT := 0x10001, CDDS_SUBITEMPREPAINT := 0x30001
        static CDRF_NEWFONT := 0x2, CDRF_NOTIFYITEMDRAW := 0x20, CDRF_NOTIFYSUBITEMDRAW := 0x20
        x64 := (A_PtrSize = 8)
        stage := NumGet(lParam, A_PtrSize * 3, "UInt")
        if (stage = CDDS_PREPAINT)
            return CDRF_NOTIFYITEMDRAW
        if (stage = CDDS_ITEMPREPAINT)
            return CDRF_NOTIFYSUBITEMDRAW
        if (stage != CDDS_SUBITEMPREPAINT)
            return 0
        colors := QuickSwitch._panelColors
        row := NumGet(lParam, x64 ? 56 : 36, "UPtr")
        column := NumGet(lParam, x64 ? 88 : 56, "Int")
        selected := SendMessage(0x102C, row, 0x2, QuickSwitch._panelList) & 0x2   ; LVM_GETITEMSTATE / LVIS_SELECTED
        stateOffset := x64 ? 64 : 40
        NumPut("UInt", NumGet(lParam, stateOffset, "UInt") & ~0x11, lParam, stateOffset)   ; remove CDIS_SELECTED / CDIS_FOCUS so the system's blue highlight isn't drawn
        text := selected ? (column ? colors.SelectedSubtitle : colors.SelectedText) : (column ? colors.Subtitle : colors.Text)
        NumPut("UInt", QuickSwitch.ColorRef(text), lParam, x64 ? 80 : 48)
        NumPut("UInt", QuickSwitch.ColorRef(selected ? colors.Selected : colors.Background), lParam, x64 ? 84 : 52)
        return CDRF_NEWFONT
    }

    static IsDarkColor(hex) => ThemeManager.IsDarkColor(hex)

    static _FontName() {
        try return ThemeManager.FontName()
        return "Segoe UI"
    }

    ; the dialog's width changed: the panel follows
    static _ResizePanel(dialogWidth) {
        width := QuickSwitch.PanelWidthFor(dialogWidth)
        QuickSwitch._panelDialogW := dialogWidth
        if (width = QuickSwitch._panelWidth)
            return
        QuickSwitch._panelWidth := width
        QuickSwitch._panelSearch.Move(, , width - Win.Scale(20))
        QuickSwitch._panelLine.Move(, , width)
        QuickSwitch._panelList.Move(, , width)
        QuickSwitch._ShowFolders(QuickSwitch._panelFolders)                ; Recomputes the column widths and adjusts the panel's size and position
    }
    static _searchTimer := () => QuickSwitch._SearchPanel()

    static ColumnsWidth(list) {
        total := 0
        loop 3
            total += SendMessage(0x101D, A_Index - 1, 0, list)            ; LVM_GETCOLUMNWIDTH
        return total
    }

    static ClientWidth(list) {
        rect := Buffer(16)
        DllCall("GetClientRect", "Ptr", list.Hwnd, "Ptr", rect)
        return NumGet(rect, 8, "Int")
    }

    ; List: folder name | full path | source (TC / Explorer / recent / search)
    static _ShowFolders(folders) {
        list := QuickSwitch._panelList
        if !IsObject(list)
            return
        QuickSwitch._panelFolders := folders
        list.Opt("-Redraw")
        list.Delete()
        for entry in folders
            list.Add(entry.HasOwnProp("IsFile") && entry.IsFile ? "Icon2" : "Icon1", QuickSwitch.FolderName(entry.Path), entry.Path, entry.Tag)
        list.ModifyCol(3, "Auto"), list.ModifyCol(1, "Auto")
        nameW := Max(Min(SendMessage(0x101D, 0, 0, list), QuickSwitch._panelWidth // 3), Win.Scale(120))   ; LVM_GETCOLUMNWIDTH
        list.ModifyCol(1, nameW)
        list.ModifyCol(2, QuickSwitch._panelWidth - nameW - SendMessage(0x101D, 2, 0, list) - Win.Scale(folders.Length > QuickSwitch.PanelRows ? 22 : 4))
        if folders.Length
            list.Modify(1, "Select Focus")
        list.Opt("+Redraw")
        rows := Max(3, Min(folders.Length, QuickSwitch.PanelRows))           ; the height follows the content (3 ~ PanelRows rows)
        list.Move(, , , QuickSwitch._rowTop + (rows - 1) * QuickSwitch._rowHeight + Win.Scale(4))
        list.ModifyCol(3, "AutoHdr")                                        ; the last column fills to the right so the selected row's color spans the row; computed after the height is set (with a scroll bar where needed), so no horizontal scroll bar appears
        ; The last column is never narrower than its (hidden) header text: with no rows, or only short tags, it can stick out
        ; past the right edge and bring up a horizontal scrollbar. Take the difference off the path column
        if (over := QuickSwitch.ColumnsWidth(list) - QuickSwitch.ClientWidth(list)) > 0
            list.ModifyCol(2, Max(0, SendMessage(0x101D, 1, 0, list) - over))
        if DllCall("IsWindowVisible", "Ptr", QuickSwitch._panel.Hwnd) {     ; The search results changed: adjust the size and position again
            QuickSwitch._panelPos := ""
            QuickSwitch._PlacePanel(QuickSwitch._panelFor)
        }
    }

    ; "D:\Projects\Tower" -> "Tower"; drive roots show as "D:\"
    static FolderName(folder) {
        trimmed := RTrim(folder, "\")
        if (trimmed ~= "^[A-Za-z]:$")
            return trimmed "\"
        SplitPath(trimmed, &name)
        return (name != "") ? name : folder
    }

    ; The text in the search box: filter the list first (paths containing every word), then add the folders and files found (Everything or the built-in index, at most 30)
    static _SearchPanel() {
        if !IsObject(QuickSwitch._panelSearch)
            return
        term := Trim(QuickSwitch._panelSearch.Value)
        QuickSwitch._ShowFolders(QuickSwitch.FilterFolders(QuickSwitch._panelBase, term, term = "" ? [] : QuickSwitch._FindItems(term)))
    }

    ; [{Path, IsFolder}], folders first. Folders are searched separately (searched together with files, a folder with exactly the name might not make it in for not being recently modified);
    ; PanelSearch = "folders" searches folders only
    static _FindItems(term) {
        found := [], seen := Map()
        try {
            for item in FileSearchProvider.Query(term, 15, false, true)
                found.Push({Path: item.Path, IsFolder: true}), seen[StrLower(item.Path)] := true
            if !QuickSwitch._SearchesFiles()
                return found
            for item in FileSearchProvider.Query(term, 30, false)
                if (!item.IsFolder && !seen.Has(StrLower(item.Path)))
                    found.Push({Path: item.Path, IsFolder: false})
        }
        return found
    }

    ; Filters base (case-insensitive, every word must appear in the path), then appends the items of extra that aren't duplicates:
    ; each item of extra is a path (folder) or {Path, IsFolder}; folders get the source "search", files "file"
    static FilterFolders(base, term, extra) {
        result := [], seen := Map(), words := StrSplit(Trim(term), " ")
        for entry in base {
            matched := true
            for word in words
                if (word != "" && !InStr(entry.Path, word))
                    matched := false
            if matched {
                result.Push(entry)
                seen[StrLower(RTrim(entry.Path, "\"))] := true
            }
        }
        for item in extra {
            itemPath := IsObject(item) ? item.Path : item
            isFile := IsObject(item) && !item.IsFolder
            if seen.Has(StrLower(RTrim(itemPath, "\")))
                continue
            seen[StrLower(RTrim(itemPath, "\"))] := true
            result.Push({Path: itemPath, Group: "search", IsFile: isFile, Tag: I18n.T(isFile ? "QuickSwitch.TagFile" : "QuickSwitch.TagSearch")})
        }
        return result
    }

    static _PanelJump(row) {
        if (row < 1 || row > QuickSwitch._panelFolders.Length || !QuickSwitch._panelFor)
            return
        entry := QuickSwitch._panelFolders[row], dialog := QuickSwitch._panelFor
        QuickSwitch._BackToDialog()
        if (entry.HasOwnProp("IsFile") && entry.IsFile)
            QuickSwitch.PickFile(entry.Path, dialog)
        else
            QuickSwitch.SetDialogPath(RTrim(entry.Path, "\") "\", dialog)
        if (IsObject(QuickSwitch._panelSearch) && QuickSwitch._panelSearch.Value != "") {   ; after jumping, clear the search so the list is back as it was
            QuickSwitch._panelSearch.Value := ""
            QuickSwitch._ShowFolders(QuickSwitch._panelBase)
        }
    }

    static _BackToDialog() {
        dialog := QuickSwitch._panelFor
        if (!dialog || WinActive("ahk_id " dialog))
            return
        try {
            WinActivate("ahk_id " dialog)
            WinWaitActive("ahk_id " dialog, , 1)
        }
    }

    ; Keys in the search box: ↑ ↓ choose, Enter jumps, Esc goes back to the dialog
    static PanelKey(key) {
        list := QuickSwitch._panelList
        if !IsObject(list)
            return
        count := list.GetCount(), row := list.GetNext()
        switch key {
            case "Up", "Down":
                if !count
                    return
                row := (key = "Down") ? Min(row + 1, count) : Max(row - 1, 1)
                list.Modify(0, "-Select"), list.Modify(row, "Select Focus Vis")
            case "Enter":
                QuickSwitch._PanelJump(row ? row : 1)
            case "Escape":
                QuickSwitch._BackToDialog()
        }
    }

    static _PlacePanel(dialog) {
        if !IsObject(QuickSwitch._panel)
            return
        rect := QuickSwitch._FrameRect(dialog)
        if !IsObject(rect)
            return QuickSwitch.HidePanel()
        if (QuickSwitch._panelDialogW && rect.W != QuickSwitch._panelDialogW)
            return QuickSwitch._ResizePanel(rect.W)
        if !QuickSwitch._panelPos                                           ; first showing, or the size changed: compute the size from the content first
            QuickSwitch._panel.Show(DllCall("IsWindowVisible", "Ptr", QuickSwitch._panel.Hwnd) ? "NA AutoSize" : "NA Hide AutoSize")
        QuickSwitch._panel.GetPos(, , &panelW, &panelH)
        area := Win.WorkAreaAt(rect.X + rect.W // 2, rect.Y + rect.H // 2)
        pos := QuickSwitch.PanelPosition(rect.X, rect.Y, rect.W, rect.H, panelW, panelH, area)
        key := pos.X "," pos.Y
        if (key != QuickSwitch._panelPos) {                                 ; reposition only when the dialog moved
            QuickSwitch._panel.Show("NA x" pos.X " y" pos.Y)
            QuickSwitch._panelPos := key
        }
    }

    ; the window's visible frame (Win10/11 windows have an invisible resize border around them, which WinGetPos includes)
    static _FrameRect(hwnd) {
        rect := Buffer(16, 0)
        if !DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", hwnd, "UInt", 9, "Ptr", rect, "UInt", 16) {   ; DWMWA_EXTENDED_FRAME_BOUNDS
            left := NumGet(rect, 0, "Int"), top := NumGet(rect, 4, "Int")
            if (NumGet(rect, 8, "Int") > left)
                return {X: left, Y: top, W: NumGet(rect, 8, "Int") - left, H: NumGet(rect, 12, "Int") - top}
        }
        try {
            WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
            return {X: x, Y: y, W: w, H: h}
        }
        return ""
    }

    ; Panel position (as in Listary): right below the dialog, left-aligned; above the dialog when there is no room below;
    ; inside the dialog at its bottom when there is room neither above nor below
    static PanelPosition(x, y, w, h, panelW, panelH, area) {
        left := Max(area.Left, Min(x, area.Right - panelW))
        if (y + h + panelH <= area.Bottom)
            return {X: left, Y: y + h}
        if (y - panelH >= area.Top)
            return {X: left, Y: y - panelH}
        return {X: left, Y: Max(area.Top, area.Bottom - panelH)}
    }

    ; Only real open / save file dialogs count (a file name box + a file list + buttons or an address bar)
    static IsFileDialog(*) => QuickSwitch._IsFileDialogWindow(WinExist("A"))   ; also the hotkeys' condition (HotIf)

    static _IsFileDialogWindow(hwnd) {
        if !(hwnd && DllCall("IsWindowVisible", "Ptr", hwnd) && QuickSwitch._MayBeDialog(hwnd))
            return false
        target := "ahk_id " hwnd, controls := [], title := "", winClass := ""  ; read once for both kinds
        try controls := WinGetControls(target)
        try title := WinGetTitle(target)
        try winClass := WinGetClass(target)
        browse := QuickSwitch.Options.Has("BrowseDialogs") && QuickSwitch.Options["BrowseDialogs"]
        if (browse && QuickSwitch.LooksLikeBrowseDialog(winClass, controls, QuickSwitch._TreeId(controls, target)))
            return true
        return QuickSwitch._Matches("ahk_group ALTRunDialogs", hwnd) && QuickSwitch.LooksLikeFileDialog(winClass, controls, title)
    }

    ; Decided by the window's controls (independent of the system language). #32770 is a system dialog: it needs Explorer's file view to count;
    ; one with only a plain list (e.g. the main window of "System Information" msinfo32, with a find box, a list and buttons) doesn't.
    ; Other programs ("other dialogs" added in Preferences) count with a plain list too; those whose controls can't be read (Qt etc.) are decided by title
    static LooksLikeFileDialog(winClass, controls, title := "") {
        if !controls.Length
            return RegExMatch(title, "i)(open|save|select|import|export|browse|打开|另存|选择|导入|导出)") > 0
        hasNameEdit := QuickSwitch._HasControl(controls, "^Edit\d+$")
        hasFileView := (winClass = "#32770")
            ? QuickSwitch._HasControl(controls, "^DirectUIHWND\d+$", "^SHELLDLL_DefView\d*$")
            : QuickSwitch._HasControl(controls, "^DirectUIHWND\d+$", "^SHELLDLL_DefView\d*$", "^SysListView32\d*$")
        hasButtons  := QuickSwitch._HasControl(controls, "^Button\d+$", "^ToolbarWindow32\d+$", "^ComboBoxEx32\d+$", "^ComboBox\d+$")
        return hasNameEdit && hasFileView && hasButtons
    }

    ; The old "Browse for Folder" dialog (SHBrowseForFolder): only a folder tree, no file list, so it is not one
    ; of the dialogs above. Told apart by fixed names, whatever the system language: the new style puts the tree
    ; in a "SHBrowseForFolder ShellNameSpace Control", the classic style has the tree itself with control ID 0x3741.
    ; Turned on by Works in → "Browse for Folder dialogs" (BrowseDialogs); ExcludeWindows applies too
    static IsBrowseWindow(hwnd) {
        controls := [], winClass := ""
        try controls := WinGetControls("ahk_id " hwnd), winClass := WinGetClass("ahk_id " hwnd)
        return QuickSwitch.LooksLikeBrowseDialog(winClass, controls, QuickSwitch._TreeId(controls, "ahk_id " hwnd))
    }

    static LooksLikeBrowseDialog(winClass, controls, treeId := 0) {
        return winClass = "#32770" && (treeId = 0x3741 || QuickSwitch._HasControl(controls, "^SHBrowseForFolder ShellNameSpace Control\d+$"))
    }

    ; Control ID of the dialog's first tree view (0 = none)
    static _TreeId(controls, win) {
        for name in controls
            if (name = "SysTreeView321")
                try return DllCall("GetDlgCtrlID", "Ptr", ControlGetHwnd(name, win), "Int")
        return 0
    }

    ; Select a folder in a "Browse for Folder" dialog's tree, like Listary: BFFM_SETSELECTIONW expands the tree to it.
    ; The message carries a pointer, so the path is first written into the dialog's own process. Doesn't press OK
    static SelectInBrowseDialog(folder, dialog) {
        folder := RTrim(folder, "\")
        if RegExMatch(folder, "^[A-Za-z]:$")
            folder .= "\"                                                   ; a drive needs its backslash
        pid := 0
        try pid := WinGetPID("ahk_id " dialog)
        if !pid || !(process := DllCall("OpenProcess", "UInt", 0x38, "Int", false, "UInt", pid, "Ptr"))   ; PROCESS_VM_OPERATION | VM_READ | VM_WRITE
            return false
        size := (StrLen(folder) + 1) * 2, ok := false
        remote := DllCall("VirtualAllocEx", "Ptr", process, "Ptr", 0, "UPtr", size, "UInt", 0x3000, "UInt", 0x04, "Ptr")   ; MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE
        if remote {
            DllCall("IsWow64Process", "Ptr", process, "Int*", &wow64 := 0)
            text := Buffer(size), StrPut(folder, text, "UTF-16")
            if (!(wow64 && remote > 0xFFFFFFFF)                             ; a 32-bit program only sees the low 4 GB
                && DllCall("WriteProcessMemory", "Ptr", process, "Ptr", remote, "Ptr", text, "UPtr", size, "Ptr", 0)) {
                try {
                    SendMessage(0x467, 1, remote, , "ahk_id " dialog, , , , 2000)   ; BFFM_SETSELECTIONW (WM_USER + 103); wParam TRUE = lParam is a path
                    ok := true
                }
            }
            DllCall("VirtualFreeEx", "Ptr", process, "Ptr", remote, "UPtr", 0, "UInt", 0x8000)   ; MEM_RELEASE
        }
        DllCall("CloseHandle", "Ptr", process)
        if !ok
            Logger.Error("QuickSwitch: cannot select " folder " in the Browse for Folder dialog")
        return ok
    }

    static _HasControl(controls, patterns*) {
        for name in controls
            for pattern in patterns
                if RegExMatch(name, "i)" pattern)
                    return true
        return false
    }

    ; The jump hotkey: the folder of the file manager used last (TC, Opus or Explorer, see FileManagerFolder)
    static SyncFileManagerPath(silent := false) {
        manager := QuickSwitch.FileManagerFolder()
        if !(IsObject(manager) && manager.Path != "")
            return silent ? "" : MsgBox(I18n.T("QuickSwitch.NoFileManager"), App.Name, 48)
        if Logger.Enabled
            Logger.Debug("QuickSwitch: " manager.Tag " is in [" manager.Path "]")
        QuickSwitch.SetDialogPath(RTrim(manager.Path, "\") "\")             ; with a trailing backslash: AutoCAD and others need it
    }

    ; The folder of a Total Commander panel: asked directly through TC's WM_COPYDATA interface (since TC 8.0), without touching the clipboard.
    ; hwnd = 0 takes the frontmost TC window; otherPanel = true takes the other panel
    static TotalCmdFolder(hwnd := 0, otherPanel := false) {
        if !(hwnd := hwnd ? hwnd : WinExist("ahk_class TTOTAL_CMD"))
            return ""
        folder := RTrim(QuickSwitch.TotalCmdAsk(hwnd, otherPanel ? "TP" : "SP"), "\")
        return (folder ~= "^([A-Za-z]:|\\\\)") ? folder : ""               ; only real paths (no FTP, inside archives etc.)
    }

    ; Sends WM_COPYDATA (dwData = "GW", an ANSI command) to TC, which replies at once with WM_COPYDATA (dwData = "RW", UTF-16)
    ; to the window in wParam. Commands: A = which side is active (L / R); two letters: first L left / R right / S active / T the other,
    ; second P path / C file count / I caret position / N file name under the caret. "" if TC doesn't reply
    static _tcWaiting := false, _tcReply := "", _tcListening := false
    static TotalCmdAsk(hwnd, command, timeoutMs := 500) {
        if !QuickSwitch._tcListening {
            OnMessage(0x4A, (p*) => QuickSwitch._OnTotalCmdReply(p*))         ; WM_COPYDATA
            QuickSwitch._tcListening := true
        }
        request := Buffer(StrPut(command, "CP0"))
        StrPut(command, request, "CP0")
        copyData := Buffer(A_PtrSize * 3, 0)                                    ; COPYDATASTRUCT
        NumPut("UPtr", Ord("G") + 256 * Ord("W"), copyData, 0)
        NumPut("UInt", request.Size, copyData, A_PtrSize)
        NumPut("Ptr", request.Ptr, copyData, A_PtrSize * 2)
        QuickSwitch._tcReply := "", QuickSwitch._tcWaiting := true, accepted := 0
        DllCall("SendMessageTimeoutW", "Ptr", hwnd, "UInt", 0x4A, "Ptr", A_ScriptHwnd, "Ptr", copyData.Ptr
              , "UInt", 2, "UInt", timeoutMs, "UPtr*", &accepted)                 ; 2 = SMTO_ABORTIFHUNG: don't wait when TC hangs
        QuickSwitch._tcWaiting := false                                         ; TC replies while handling this message, so no extra wait
        return QuickSwitch._tcReply
    }

    static _OnTotalCmdReply(wParam, lParam, msg, hwnd) {
        if (!QuickSwitch._tcWaiting || NumGet(lParam, 0, "UPtr") != Ord("R") + 256 * Ord("W"))
            return                                                              ; not the reply waited for: leave it to other handlers (Everything)
        size := NumGet(lParam, A_PtrSize, "UInt"), data := NumGet(lParam, A_PtrSize * 2, "Ptr")
        reply := (data && size >= 2) ? StrGet(data, size // 2, "UTF-16") : ""  ; up to the terminating 0
        QuickSwitch._tcReply := reply, QuickSwitch._tcWaiting := false
        return true
    }

    ;---------------------------------------------------------------------------
    ; The folders in the panel
    ;---------------------------------------------------------------------------
    ; [{Path, Group: "open" / "recent", Tag}]: the active and other panel of each TC window, each Explorer window,
    ; recent folders (RecentFolders of them), without duplicates
    ; opus = false: without Directory Opus (the panel adds those a moment later, see _AddDOpusFolders)
    static MenuFolders(opus := true) {
        list := [], seen := Map()
        add(folder, group, tag) {
            if (folder = "" || seen.Has(StrLower(RTrim(folder, "\"))))
                return
            seen[StrLower(RTrim(folder, "\"))] := true
            list.Push({Path: folder, Group: group, Tag: tag})
        }
        for hwnd in WinGetList("ahk_class TTOTAL_CMD") {
            add(QuickSwitch.TotalCmdFolder(hwnd), "open", I18n.T("QuickSwitch.TagTC"))
            add(QuickSwitch.TotalCmdFolder(hwnd, true), "open", I18n.T("QuickSwitch.TagTCOther"))
        }
        if opus                                                             ; the source tab of each lister, then the other side
            for state in [1, 2]
                for tab in QuickSwitch.DOpusTabs()
                    if (tab.State = state)
                        add(tab.Path, "open", I18n.T(state = 1 ? "QuickSwitch.TagOpus" : "QuickSwitch.TagOpusOther"))
        try {
            for window in ComObject("Shell.Application").Windows
                try add(window.Document.Folder.Self.Path, "open", I18n.T("QuickSwitch.TagExplorer"))
        }
        count := QuickSwitch.Options.Has("RecentFolders") ? QuickSwitch.Options["RecentFolders"] : 10
        for folder in QuickSwitch.RecentFolders(count)
            add(folder, "recent", I18n.T("QuickSwitch.TagRecent"))
        return list
    }

    ; The newest limit folders of Windows "Recent items" (WindowsRecent); a file gives its folder.
    ; Only the newest limit * 6 shortcuts are looked at, not all of the hundreds there may be
    static RecentFolders(limit, recentDir := "") {
        folders := [], seen := Map()
        if (limit <= 0)
            return folders
        for item in WindowsRecent.Items(limit * 6, recentDir) {
            folder := RTrim(item.IsFolder ? item.Path : QuickSwitch._Parent(item.Path), "\")
            if (folder ~= "^[A-Za-z]:$")
                folder .= "\"
            if (folder = "" || seen.Has(StrLower(folder)))
                continue
            seen[StrLower(folder)] := true
            folders.Push(folder)
            if (folders.Length >= limit)
                break
        }
        return folders
    }

    static _Parent(target) {
        SplitPath(RTrim(target, "\"), , &dir)
        return dir
    }

    ; The current folder of an Explorer window; hwnd = 0 takes the frontmost Explorer window
    static ExplorerFolder(hwnd := 0) {
        if !hwnd
            hwnd := WinExist("ahk_class CabinetWClass")
        if !hwnd
            return ""
        try {
            for window in ComObject("Shell.Application").Windows
                if (window.HWND = hwnd)
                    return window.Document.Folder.Self.Path
        }
        return ""
    }

    ; Current folder of the TC / Opus / Explorer window nearest the top: {Path, Tag, Hwnd}, "" when none.
    ; Windows are listed front to back, so with a dialog in front this is the file manager used last
    ; (the jump hotkey, auto jump, "Copy to ..." in the actions)
    static FileManagerFolder() {
        for hwnd in WinGetList() {
            className := ""
            try className := WinGetClass("ahk_id " hwnd)
            switch className {
                case "TTOTAL_CMD"   : return {Path: QuickSwitch.TotalCmdFolder(hwnd), Tag: "Total Commander", Hwnd: hwnd}
                case "dopus.lister" : return {Path: QuickSwitch.DOpusFolder(hwnd), Tag: "Directory Opus", Hwnd: hwnd}
                case "CabinetWClass": return {Path: QuickSwitch.ExplorerFolder(hwnd), Tag: I18n.T("QuickSwitch.TagExplorer"), Hwnd: hwnd}
            }
        }
        return ""
    }

    static IsFileManager(hwnd) {
        className := ""
        try className := WinGetClass("ahk_id " hwnd)
        return className = "TTOTAL_CMD" || className = "dopus.lister" || className = "CabinetWClass"
    }

    static FolderOfWindow(hwnd) {
        if (!hwnd || !WinExist("ahk_id " hwnd))
            return ""
        switch WinGetClass("ahk_id " hwnd) {
            case "TTOTAL_CMD"   : return QuickSwitch.TotalCmdFolder(hwnd)
            case "dopus.lister" : return QuickSwitch.DOpusFolder(hwnd)
            case "CabinetWClass": return QuickSwitch.ExplorerFolder(hwnd)
        }
        return ""
    }

    ;---------------------------------------------------------------------------
    ; Directory Opus: its own "dopusrt.exe /info <file>,paths" writes every lister's tabs to a small XML file
    ; (read only, no clipboard). It runs only when an Opus window is open and a folder is actually needed,
    ; and a result is reused for DOpusCacheMs
    ;---------------------------------------------------------------------------
    static DOpusCacheMs := 1500, _dopusTabs := [], _dopusAt := 0

    ; Source (focused) tab of a lister; hwnd 0 = the active lister
    static DOpusFolder(hwnd := 0) => QuickSwitch.PickDOpusFolder(QuickSwitch.DOpusTabs(), hwnd)

    static PickDOpusFolder(tabs, hwnd := 0) {
        for tab in tabs
            if (tab.State = 1 && (hwnd ? tab.Lister = hwnd : tab.ActiveLister))
                return tab.Path
        for tab in tabs                                                     ; lister not found (e.g. Opus just opened it): the active one
            if (tab.State = 1 && tab.ActiveLister)
                return tab.Path
        return ""
    }

    static DOpusTabs() {
        if !(lister := WinExist("ahk_class dopus.lister"))
            return []
        if (QuickSwitch._dopusAt && A_TickCount - QuickSwitch._dopusAt < QuickSwitch.DOpusCacheMs)
            return QuickSwitch._dopusTabs
        started := Logger.Ms(), tabs := []
        try {
            SplitPath(WinGetProcessPath("ahk_id " lister), , &dir)
            xmlFile := A_Temp "\ALTRun.dopus-paths.xml"
            try FileDelete(xmlFile)
            Run('"' dir '\dopusrt.exe" /info "' xmlFile '",paths', , "Hide", &pid)
            if ProcessWaitClose(pid, 3)                                     ; still running after 3 s (returns its PID): never hang the dialog on it
                try ProcessClose(pid)
            tabs := QuickSwitch.ParseDOpusPaths(FileRead(xmlFile, "UTF-8"))
            try FileDelete(xmlFile)
        } catch as e {
            Logger.Error("QuickSwitch: cannot read the Directory Opus folders - " e.Message)
        }
        QuickSwitch._dopusTabs := tabs, QuickSwitch._dopusAt := A_TickCount
        if Logger.Enabled
            Logger.Debug(Format("Perf: Directory Opus paths {} ms, {} tabs", Round(Logger.Ms() - started), tabs.Length))
        return tabs
    }

    ; <path active_lister="1" active_tab="1" tab_state="1" side="1" lister="0x70e2a" tab="0x50e1a" ...>C:\Folder</path>
    ; tab_state: 1 = the source (focused) tab of its lister, 2 = the tab shown on the other side of a dual lister
    static ParseDOpusPaths(xml) {
        tabs := [], pos := 1, attrs := ""
        value(name) => RegExMatch(attrs, "\b" name '="([^"]*)"', &a) ? a[1] : ""
        while (pos := RegExMatch(xml, "s)<path\b([^>]*)>(.*?)</path>", &m, pos)) {
            pos += m.Len
            attrs := m[1], folder := m[2]
            for entity in [["&lt;", "<"], ["&gt;", ">"], ["&quot;", '"'], ["&apos;", "'"], ["&amp;", "&"]]   ; &amp; last
                folder := StrReplace(folder, entity[1], entity[2])
            lister := value("lister"), state := value("tab_state")
            tabs.Push({Path: Trim(folder), Lister: IsInteger(lister) ? Integer(lister) : 0, State: IsInteger(state) ? Integer(state) : 0
                , ActiveLister: value("active_lister") = "1"})
        }
        return tabs
    }

    ; Choosing a file in a dialog: for open and save dialogs alike, jump to its folder (without opening or saving for the user)
    static PickFile(filePath, dialog := 0) {
        SplitPath(filePath, , &dir)
        if (dir != "")
            QuickSwitch.SetDialogPath(RTrim(dir, "\") "\", dialog)
    }

    ; fill the old name back only when the file name box held a normal file name (not when empty, a path or a filter like *.txt)
    static ShouldRestoreName(saved) {
        saved := Trim(saved)
        return saved != "" && !RegExMatch(saved, "[\\/:*?]")
    }

    ; dialog: the dialog to jump in (the current window by default). Before simulating keyboard input, make sure it is the foreground window,
    ; otherwise the text would go to another window (e.g. switching back to the dialog failed and the panel is still in front)
    static SetDialogPath(folder, dialog := 0) {
        if (folder = "" || !FileExist(folder))
            return
        if !dialog
            dialog := WinExist("A")
        if !(dialog && WinExist("ahk_id " dialog))
            return
        Usage.Count("QuickSwitch")
        method := QuickSwitch.DialogJumpMethod(dialog)
        if (method = "browse")                                              ; folder tree only: select it there, pressing Enter would mean OK
            return QuickSwitch.SelectInBrowseDialog(folder, dialog)
        if (method = "message") {
            if QuickSwitch._JumpByMessage(folder, dialog)
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
                if QuickSwitch.ShouldRestoreName(saved) {                   ; put it back after the jump, so the suggested name isn't lost
                    Sleep(150)
                    try ControlSetText(saved, "Edit1", "ahk_id " dialog)
                }
                return
            }
            method := (className = "#32770") ? "address" : "type"           ; the box refused: as if there were none
        }
        if !WinActive("ahk_id " dialog)                                     ; typing goes to the foreground window: only into the dialog
            return
        if (method = "address") {                                           ; a system dialog without the box: its address bar
            SendInput("!d")
            Sleep(60)
        }
        SendText(folder)                                                    ; dialogs a program draws itself (WPS / Qt, GTK, Java...): the
        SendInput("{Enter}")                                                ; file name box has the focus; a folder path + Enter goes there
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
            Logger.Debug("QuickSwitch: jump to [" folder "] by message" QuickSwitch._DialogState(dialog))
        try {
            box := ControlGetHwnd("Edit1", target)
            try saved := ControlGetText(box)
            if !QuickSwitch._SetBoxText(box, folder) {
                Logger.Debug("QuickSwitch: the file name box didn't take the folder, or the dialog didn't take the change notice")
                return false
            }
        } catch as e {
            Logger.Debug("QuickSwitch: no file name box - " e.Message)
            return false
        }
        started := A_TickCount
        if !(answered := QuickSwitch._SendCommand(dialog, DllCall("GetDlgItem", "Ptr", dialog, "Int", 1, "Ptr"), 0, 1))   ; BN_CLICKED from IDOK
            Logger.Warn("QuickSwitch: the dialog didn't answer the OK command in 2 s")   ; it may still go there: no second try with keys
        else if Logger.Enabled
            Logger.Debug("QuickSwitch: OK command answered in " (A_TickCount - started) " ms" QuickSwitch._DialogState(dialog))
        if (answered && QuickSwitch.ShouldRestoreName(saved)) {             ; put back a name Save As suggested, so it isn't lost
            Loop 6 {                                                        ; the dialog may put it back itself after going there: leave that
                Sleep(50)
                text := ""
                try text := ControlGetText(box)
                if (text != folder && text != "")
                    return true
            }
            try QuickSwitch._SetBoxText(box, saved)
        }
        return true
    }

    ; The folder a system dialog shows, from its address bar ("Address: C:\Docs", the label in the UI language); "" if it can't be read.
    ; The tests check a jump with it
    static DialogFolder(dialog) {
        text := ""
        try text := WinGetText("ahk_id " dialog)
        Loop Parse, text, "`n", "`r"
            if RegExMatch(A_LoopField, "^[^:：\\]{1,40}[:：]\s*((?:[A-Za-z]:\\|\\\\)[^\r\n]*)$", &m)
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
            return QuickSwitch._SendCommand(DllCall("GetParent", "Ptr", parent, "Ptr"), parent, 5)   ; CBN_EDITCHANGE
        return QuickSwitch._SendCommand(parent, box, 0x300)                                       ; EN_CHANGE
    }

    ; WM_COMMAND (code, the control's id) from control to window, sent with a 2 s limit; false if it didn't get through
    static _SendCommand(window, control, code, id := "") {
        id := (id != "") ? id : DllCall("GetDlgCtrlID", "Ptr", control, "Int")
        return window && DllCall("SendMessageTimeoutW", "Ptr", window, "UInt", 0x111, "Ptr", (code << 16) | (id & 0xFFFF), "Ptr", control
            , "UInt", 0x2, "UInt", 2000, "Ptr", 0, "Ptr") != 0              ; SMTO_ABORTIFHUNG
    }

    ; How the folder goes into this dialog, see JumpMethod; "browse" = a folder tree dialog (SelectInBrowseDialog)
    static DialogJumpMethod(dialog) {
        if QuickSwitch.IsBrowseWindow(dialog)
            return "browse"
        className := "", hasEdit := false
        try className := WinGetClass("ahk_id " dialog)
        try hasEdit := ControlGetHwnd("Edit1", "ahk_id " dialog) != 0
        return QuickSwitch.JumpMethod(className, hasEdit, DllCall("GetDlgItem", "Ptr", dialog, "Int", 1, "Ptr") != 0)
    }

    ; How to put the folder into a dialog:
    ;   message = a system dialog (#32770) with a file name box and an OK button (IDOK): the folder and OK as messages (see _JumpByMessage)
    ;   edit    = into its file name box (Edit1) + Enter: programs with real edit controls that aren't system dialogs
    ;   address = a system dialog (#32770) without that box: Alt+D (address bar), type, Enter
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
