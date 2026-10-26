;===============================================================================
; SearchInputLanguage.ahk - switches the search box to English input while it is shown, and back afterwards (AutoHotkey v2)
;-------------------------------------------------------------------------------
; General.SwitchToEnglishInput: 1 = the English (US) keyboard, 2 = the English mode of the current input method;
; General.RestoreInput switches back when the window hides.
;
; Usage (SearchWindow):
;   SearchInputLanguage.SwitchOnShow(wasVisible)   after the search box got focus
;   SearchInputLanguage.Restore()                  when the window hides
;===============================================================================

class SearchInputLanguage {
    static _layoutBefore := 0                                               ; input language of the foreground window before showing (switched back after hiding when SwitchToEnglishInput = 1)
    static _imeBefore := ""                                                 ; {Open, Conv} of the IME when shown (restored after hiding when SwitchToEnglishInput = 2)

    ; Called by SearchWindow._Show once the search box has focus; the steps are logged like the other steps of showing
    static SwitchOnShow(wasVisible) {
        switch AppSettings.General["SwitchToEnglishInput"] {
            case 1:                                                         ; switch to the English (US) keyboard
                SearchWindow._Step("input language: read")                  ; logged in two steps: right after boot this once took 17 s; see which step
                step := Logger.Ms()
                if (!wasVisible || !SearchInputLanguage._layoutBefore)          ; switch back to the previous input language on hide
                    SearchInputLanguage._layoutBefore := App.PreviousWindow ? Win.KeyboardLayout(App.PreviousWindow) : 0
                SearchWindow._LogSlowStep("input language: read", step)
                SearchWindow._Step("input language: switch")
                step := Logger.Ms()
                Win.SwitchToEnglishIME()
                SearchWindow._LogSlowStep("input language: switch", step)
            case 2:                                                         ; keep the input method, switch it to English mode
                SearchWindow._Step("input mode")
                step := Logger.Ms()
                SearchInputLanguage._SwitchModeToEnglish(wasVisible)
                SearchWindow._LogSlowStep("input mode", step)
        }
    }

    ; Switches the search box's IME to English mode (setting the open status and conversion mode like the CapsLock Chinese / English toggle), keeping the old state to restore on hide.
    ; No extra English (US) keyboard needed: the input method stays the same, only in English (suggested in #154, 2026-10)
    static _SwitchModeToEnglish(wasVisible) {
        if !(imeWnd := DllCall("imm32\ImmGetDefaultIMEWnd", "Ptr", SearchWindow.Input.Hwnd, "Ptr"))
            return
        before := Win.ImeState(imeWnd)
        if (!wasVisible || !IsObject(SearchInputLanguage._imeBefore))
            SearchInputLanguage._imeBefore := before
        target := Win.ImeEnglish(before.Conv)
        Win.SetImeState(imeWnd, target.Open, target.Conv)
        if Logger.Enabled {
            after := Win.ImeState(imeWnd)
            Logger.Debug(Format("IME to English: open {} conv 0x{:X} -> now open {} conv 0x{:X}", before.Open, before.Conv, after.Open, after.Conv))
        }
    }

    ; Showing switched to an English input language: switch back on hide. Only needed when all programs share one input language (the Windows default);
    ; with "use a different input method for each app window" only ALTRun itself switched and other programs are unaffected
    static Restore() {
        SearchInputLanguage._RestoreMode()
        layout := SearchInputLanguage._layoutBefore, SearchInputLanguage._layoutBefore := 0
        if (!layout || AppSettings.General["SwitchToEnglishInput"] != 1 || !AppSettings.General["RestoreInput"]
            || Win.PerWindowInputMethod() || layout = Win.KeyboardLayout())
            return
        if SearchWindow.IsActive() {                                        ; still in the foreground: switch back ourselves, which applies to all programs
            try DllCall("ActivateKeyboardLayout", "Ptr", layout, "UInt", 0)
        } else if (hwnd := WinExist("A")) {                                 ; focus already lost: ask the foreground window to switch
            try PostMessage(0x50, 0, layout, , "ahk_id " hwnd)              ; WM_INPUTLANGCHANGEREQUEST
        }
    }

    ; Showing switched the IME to English mode: restore the old open status and conversion mode on hide. Most Chinese IMEs share one Chinese / English state across programs,
    ; so without restoring, other programs would be in English too. If focus is already lost, set it on the foreground window (like switching the input language back above)
    static _RestoreMode() {
        before := SearchInputLanguage._imeBefore, SearchInputLanguage._imeBefore := ""
        if (!IsObject(before) || AppSettings.General["SwitchToEnglishInput"] != 2 || !AppSettings.General["RestoreInput"])
            return
        if !(hwnd := WinExist("A")) || !(imeWnd := Win.ImeWindow(hwnd))
            return
        Win.SetImeState(imeWnd, before.Open, before.Conv)
        if Logger.Enabled {
            after := Win.ImeState(imeWnd)
            Logger.Debug(Format("IME restore: open {} conv 0x{:X}{} -> now open {} conv 0x{:X}", before.Open, before.Conv
                , SearchWindow.IsActive() ? "" : " (other window)", after.Open, after.Conv))
        }
    }
}
