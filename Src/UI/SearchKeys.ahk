;===============================================================================
; SearchKeys.ahk - the search window's keyboard handling (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Every key pressed in the search box or the result list goes through OnKeyDown (WM_KEYDOWN / WM_SYSKEYDOWN via
; OnMessage, or the subclassed window procedure while a system dialog's modal loop runs); the keys are listed at
; the top of SearchWindow.ahk. Also the caret helpers of the search box. Part of SearchWindow, whose state and
; actions it uses.
;
; Usage:
;   OnMessage(0x100, (p*) => SearchKeys.OnKeyMessage(p*))   in SearchWindow.Create
;   SearchKeys.HookIfModal() / Unhook()                      when the window shows / hides
;===============================================================================

class SearchKeys {
    static IsOwnWindow(hwnd) {
        return IsObject(SearchWindow.Gui) && (hwnd = SearchWindow.Gui.Hwnd || hwnd = SearchWindow.Input.Hwnd || hwnd = SearchWindow.List.Hwnd)
    }

    static OnKeyMessage(wParam, lParam, msg, hwnd) {
        result := SearchKeys.OnKeyDown(wParam, lParam, msg, hwnd)
        SearchKeys._seen := (result = "" && SearchKeys.IsOwnWindow(hwnd))           ; dispatched next: the subclass (if any) lets it through
        return result
    }

    ; While ALTRun shows a system dialog (MsgBox, FileSelect, DirSelect...), that dialog's modal message loop
    ; dispatches posted messages without calling OnMessage, so the keys never reach OnKeyDown: typing still
    ; searches, but Enter, Esc, arrows... do nothing. Shown at such a time, the window subclasses the search box
    ; and the list and handles the keys in their window procedure instead (until it hides)
    static _callback := 0, _hooked := false, _seen := false, _swallowChar := false

    ; A visible dialog (#32770) of this process: MsgBox, FileSelect, DirSelect... (about 0.04 ms: only dialogs are looked at)
    static InSystemDialog() {
        pid := DllCall("GetCurrentProcessId"), hwnd := 0
        while (hwnd := DllCall("FindWindowEx", "Ptr", 0, "Ptr", hwnd, "Str", "#32770", "Ptr", 0, "Ptr")) {
            DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "UInt*", &owner := 0)
            if (owner = pid && DllCall("IsWindowVisible", "Ptr", hwnd))
                return hwnd
        }
        return 0
    }

    static HookIfModal() {
        if SearchKeys.InSystemDialog()
            SearchKeys.Hook()
    }

    static Hook() {
        if SearchKeys._hooked
            return
        if !SearchKeys._callback                                            ; created once and kept: it may be running when the window hides
            SearchKeys._callback := CallbackCreate((hwnd, msg, wParam, lParam, id, data) => SearchKeys._Proc(hwnd, msg, wParam, lParam), , 6)
        SearchKeys._seen := false, SearchKeys._swallowChar := false
        for hwnd in [SearchWindow.Input.Hwnd, SearchWindow.List.Hwnd]
            DllCall("comctl32\SetWindowSubclass", "Ptr", hwnd, "Ptr", SearchKeys._callback, "UPtr", 1, "UPtr", 0)
        SearchKeys._hooked := true
        Logger.Debug("SearchWindow: keys handled in the window procedure (a system dialog is open)")
    }

    static Unhook() {
        if !SearchKeys._hooked
            return
        for hwnd in [SearchWindow.Input.Hwnd, SearchWindow.List.Hwnd]
            DllCall("comctl32\RemoveWindowSubclass", "Ptr", hwnd, "Ptr", SearchKeys._callback, "UPtr", 1)
        SearchKeys._hooked := false
    }

    static _Proc(hwnd, msg, wParam, lParam) {
        if (msg = 0x100 || msg = 0x104) {                                   ; WM_KEYDOWN / WM_SYSKEYDOWN
            SearchKeys._swallowChar := false
            if SearchKeys._seen                                             ; OnMessage already saw it (no modal loop)
                SearchKeys._seen := false
            else if (SearchKeys.OnKeyDown(wParam, lParam, msg, hwnd) != "") {
                SearchKeys._swallowChar := true                             ; the WM_CHAR made from it (Enter, Esc, Space...) must not reach the box either
                return 0
            }
        } else if ((msg = 0x102 || msg = 0x106) && SearchKeys._swallowChar) {     ; WM_CHAR / WM_SYSCHAR
            SearchKeys._swallowChar := false
            return 0
        }
        return DllCall("comctl32\DefSubclassProc", "Ptr", hwnd, "UInt", msg, "UPtr", wParam, "Ptr", lParam, "Ptr")
    }

    ; 0 = the key is handled and not passed to the input box; empty = handled the normal way
    static OnKeyDown(wParam, lParam, msg, hwnd) {
        if !SearchKeys.IsOwnWindow(hwnd)
            return
        vk := wParam
        ctrl  := GetKeyState("Ctrl")
        shift := GetKeyState("Shift")
        alt   := (msg = 0x104) || GetKeyState("Alt")
        actions := SearchWindow.Mode = "actions"

        switch vk {
            case 0x0D:                                                      ; Enter
                SearchWindow._Execute(ctrl ? "ctrl" : alt ? "alt" : "")
                return 0
            case 0x1B:                                                      ; Esc
                if (actions && SearchWindow._actionsOnly)
                    SearchWindow.Hide("Esc")
                else if actions
                    SearchWindow._CloseActions()
                else
                    SearchWindow.Hide("Esc")
                return 0
            case 0x20:                                                      ; Space: in the empty search box (or with the caret at the start), enter file search mode
                if (ctrl || alt)
                    return
                ; with all text selected (the restored last search) a space would replace it, so treat it as an empty search box
                if (!shift && !actions && !SearchWindow.FileMode && SearchKeys.AllTextSelected()) {
                    SearchWindow._SetInput("")
                    if SearchWindow._CanEnterFileMode() {
                        SearchWindow._SetFileMode(true)
                        return 0
                    }
                    return
                }
                if (!shift && SearchWindow._CanEnterFileMode()) {
                    SearchWindow._SetFileMode(true)
                    return 0
                }
                ; SpaceToRun (the 2.x option of the same name): with text typed, Space runs the selected result; Shift+Space types a space as usual
                if (!shift && AppSettings.General["SpaceToRun"] && SearchWindow.Input.Value != "" && IsObject(SearchWindow.SelectedItem())) {
                    SearchWindow._Execute()
                    return 0
                }
                return
            case 0x26:                                                      ; ↑ / Ctrl+↑ previous search in the history
                if ctrl
                    SearchWindow.RecallHistory(1)
                else
                    SearchWindow.MoveSelection(-1)
                return 0
            case 0x28:                                                      ; ↓ / Ctrl+↓ next search in the history
                if ctrl
                    SearchWindow.RecallHistory(-1)
                else
                    SearchWindow.MoveSelection(1)
                return 0
            case 0x21:                                                      ; PgUp
                SearchWindow.MoveSelection(-SearchWindow.VisibleRows)
                return 0
            case 0x22:                                                      ; PgDn
                SearchWindow.MoveSelection(SearchWindow.VisibleRows)
                return 0
            case 0x2D:                                                      ; Insert: mark / unmark (as in Total Commander) and move to the next row
                if (!actions && !ctrl && !shift && !alt) {
                    SearchWindow.ToggleMark()
                    return 0
                }
                return
            case 0x09:                                                      ; Tab: complete, or (TabKey = "Next") next / previous result; Ctrl+Tab always completes
                if (!ctrl && AppSettings.General["TabKey"] = "Next")
                    SearchWindow.MoveSelection(shift ? -1 : 1, true)
                else
                    SearchWindow._AutoComplete()
                return 0
            case 0x27:                                                      ; →
                ; Only a fresh →: holding → to move the caret doesn't open the actions panel by auto-repeat when the caret reaches the end (lParam bit 30 = already down)
                ; With text selected (e.g. the restored last search selected all), like a normal input box it only deselects and moves the caret to the end, so editing an expression doesn't open the panel
                if (!(lParam & 0x40000000) && SearchKeys.CaretAtEnd() && !SearchKeys.HasTextSelection() && SearchWindow._OpenActions())
                    return 0
                return
            case 0x25:                                                      ; ←
                if (actions && SearchWindow.Input.Value = "") {
                    if SearchWindow._actionsOnly
                        SearchWindow.Hide()
                    else
                        SearchWindow._CloseActions()
                    return 0
                }
                return
            case 0x70:                                                      ; F1
                SearchWindow.Hide()
                App.About()
                return 0
            case 0x72:                                                      ; F3
                SearchWindow._EditSelected()
                return 0
            case 0x71:                                                      ; F2
                SearchWindow.Hide()
                App.OpenPreferences()
                return 0
            case 0x73:                                                      ; F4
                SearchWindow.Hide()
                App.EditSettingsFile()
                return 0
            case 0x2E:                                                      ; Ctrl+Del
                if (ctrl && SearchKeys.CaretAtEnd() && IsObject(item := SearchWindow._TargetItem()) && ResultActions.CanDelete(item)) {
                    SearchWindow._Later(() => SearchWindow.DeleteItem(item))
                    return 0
                }
                return
            case 0x08:                                                      ; Backspace
                if (actions && SearchWindow.Input.Value = "") {
                    if SearchWindow._actionsOnly
                        SearchWindow.Hide()
                    else
                        SearchWindow._CloseActions()
                    return 0
                }
                ; pressed again after emptying the box in file search mode, or with the caret at the start: back to normal search
                if (SearchWindow.FileMode && (SearchWindow.Input.Value = "" || !actions && SearchKeys.CaretAtStart())) {
                    SearchWindow._SetFileMode(false)
                    return 0
                }
                if ctrl {                                                   ; Ctrl+Backspace: delete the previous word (a normal Edit control doesn't support it)
                    SearchKeys.DeletePreviousWord()
                    return 0
                }
                return
        }

        if !ctrl
            return
        switch vk {
            case 0x4E:                                                      ; Ctrl+N
                SearchWindow.MoveSelection(1)
                return 0
            case 0x50:                                                      ; Ctrl+P
                SearchWindow.MoveSelection(-1)
                return 0
            case 0x43:                                                      ; Ctrl+C
                if (shift || SearchKeys.HasTextSelection())
                    return
                SearchWindow._CopySelected()
                return 0
            case 0x4C:                                                      ; Ctrl+L
                if LargeType.Enabled()
                    SearchWindow._ShowLargeType()
                return 0
            case 0x45:                                                      ; Ctrl+E
                SearchWindow._EditFileSelected()
                return 0
            case 0xBC:                                                      ; Ctrl+,
                SearchWindow.Hide()
                App.OpenPreferences()
                return 0
        }
        if (vk >= 0x31 && vk <= 0x39) {                                     ; Ctrl+1..9
            SearchWindow._RunVisibleRow(vk - 0x30)
            return 0
        }
    }

    ; Caret at the start, no text selected
    static CaretAtStart() {
        return SendMessage(0xB0, 0, 0, SearchWindow.Input.Hwnd) = 0             ; EM_GETSEL: start and end both 0
    }

    static CaretAtEnd() {
        selection := SendMessage(0xB0, 0, 0, SearchWindow.Input.Hwnd)        ; EM_GETSEL
        return ((selection >> 16) & 0xFFFF) >= StrLen(SearchWindow.Input.Value)
    }

    ; All text in the box is selected (after restoring the last search): a key press now replaces it
    static AllTextSelected() {
        len := StrLen(SearchWindow.Input.Value)
        selection := SendMessage(0xB0, 0, 0, SearchWindow.Input.Hwnd)        ; EM_GETSEL
        return len && (selection & 0xFFFF) = 0 && ((selection >> 16) & 0xFFFF) >= len
    }

    static HasTextSelection() {
        selection := SendMessage(0xB0, 0, 0, SearchWindow.Input.Hwnd)
        return (selection & 0xFFFF) != ((selection >> 16) & 0xFFFF)
    }

    static DeletePreviousWord() {
        text := SearchWindow.Input.Value
        caret := (SendMessage(0xB0, 0, 0, SearchWindow.Input.Hwnd) >> 16) & 0xFFFF
        before := RegExReplace(SubStr(text, 1, caret), "\S*\s*$")
        SearchWindow.Input.Value := before SubStr(text, caret + 1)
        len := StrLen(before)
        SendMessage(0xB1, len, len, SearchWindow.Input.Hwnd)
        SearchWindow._OnInputChange()
    }
}
