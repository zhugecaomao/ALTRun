;===============================================================================
; WindowProvider.ahk - switches to windows that are already open (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like PowerToys Run's Window Walker and Raycast's Switch Windows: type a window title or program name,
; Enter switches to it (restoring it first if minimized).
;   "w word" (Keyword)     searches windows only; only "w " lists all windows (in z-order)
;   typing directly        with InDefaultResults = 1, windows whose title matches well show too (after applications and commands, at most 3)
; Actions panel (→): close the window.
; Only windows visible on the taskbar are listed: visible, with a title, not tool windows / windows owned by another window,
; not on another virtual desktop (DWM cloaked), and none of ALTRun's own windows.
; The window list is reused within a search; typing within 1 s doesn't enumerate again.
;
; Settings (ALTRun.json -> Features.Windows): Enabled / Keyword / InDefaultResults
;===============================================================================

class WindowProvider {
    static Id := "Windows"
    static CacheMs := 1000
    static _list := [], _listTime := 0

    static Init() {
    }

    static Search(query) {
        options := AppSettings.Feature("Windows")
        exclusive := false, term := query.Text
        if query.MatchKeyword([options["Keyword"]], &rest) {
            term := rest, exclusive := true
        } else if (!options["InDefaultResults"] || StrLen(query.Text) < 2) {
            return []
        }
        windows := WindowProvider.List()
        needle := StrLower(Trim(term))
        scores := Map()
        for index, window in windows {
            score := (needle = "") ? 100 - index * 0.01 : FuzzyMatcher.BestKey(needle, window.Keys)
            if (score > 0 && (exclusive || score >= 60))                   ; only good matches in the default results, to avoid a pile of browser tabs
                scores[index] := score
        }
        results := []
        for index in FuzzyMatcher.TopIndexes(scores, exclusive ? ProviderRegistry.MaxResults : 3) {
            window := windows[index]
            results.Push(WindowProvider.ItemFor(window, scores[index] - (exclusive ? 0 : 14), exclusive))
        }
        return results
    }

    static ItemFor(window, score, exclusive := false) {
        hwnd := window.Hwnd
        return ResultItem(window.Title, I18n.T("Win.Subtitle", window.Process), {
            Icon: (window.Path != "") ? window.Path : "res:imageres.dll,-102", Score: score, Exclusive: exclusive,
            Uid: "window:" StrLower(window.Process), Arg: window.Title,
            OnRun: (*) => WindowProvider.Activate(hwnd),
            Actions: [ResultItem(I18n.T("Action.CloseWindow"), window.Title, {Icon: "res:imageres.dll,-98", OnRun: (*) => WindowProvider.Close(hwnd)})]
        })
    }

    ; The windows on the taskbar, in z-order (frontmost first): [{Hwnd, Title, Process, Path, Keys}]
    static List() {
        if (A_TickCount - WindowProvider._listTime < WindowProvider.CacheMs)
            return WindowProvider._list
        list := []
        for hwnd in WinGetList() {
            if !WindowProvider.IsSwitchable(hwnd)
                continue
            try {
                title := WinGetTitle("ahk_id " hwnd), process := WinGetProcessName("ahk_id " hwnd)
            } catch {
                continue
            }
            exePath := ""
            try exePath := WinGetProcessPath("ahk_id " hwnd)
            name := RegExReplace(process, "i)\.exe$")
            pinyinText := Pinyin.Initials(title)
            list.Push({Hwnd: hwnd, Title: title, Process: process, Path: exePath
                , Keys: [FuzzyMatcher.Key(title), FuzzyMatcher.Key(name), (pinyinText != title) ? FuzzyMatcher.Key(pinyinText) : ""]})
        }
        WindowProvider._list := list, WindowProvider._listTime := A_TickCount
        return list
    }

    ; The same test as the taskbar: visible, with a title, not a tool window, no owner, not hidden by DWM (another virtual desktop / a suspended Store app)
    static IsSwitchable(hwnd) {
        try {
            if !DllCall("IsWindowVisible", "Ptr", hwnd) || WinGetTitle("ahk_id " hwnd) = ""
                return false
            exStyle := WinGetExStyle("ahk_id " hwnd)
            if (exStyle & 0x80) && !(exStyle & 0x40000)                     ; WS_EX_TOOLWINDOW, unless WS_EX_APPWINDOW
                return false
            if DllCall("GetWindow", "Ptr", hwnd, "UInt", 4, "Ptr") && !(exStyle & 0x40000)   ; GW_OWNER: dialogs etc.
                return false
            cloaked := 0
            DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", hwnd, "UInt", 14, "UInt*", &cloaked, "UInt", 4)   ; DWMWA_CLOAKED
            if cloaked
                return false
            if (WinGetPID("ahk_id " hwnd) = DllCall("GetCurrentProcessId"))  ; ALTRun's own windows
                return false
            return WinGetClass("ahk_id " hwnd) != "Progman"
        }
        return false
    }

    static Activate(hwnd) {
        if !WinExist("ahk_id " hwnd)
            return App.Notify(I18n.T("Win.Gone"))
        if (WinGetMinMax("ahk_id " hwnd) = -1)
            WinRestore("ahk_id " hwnd)
        WinActivate("ahk_id " hwnd)
    }

    static Close(hwnd) {
        if WinExist("ahk_id " hwnd)
            WinClose("ahk_id " hwnd)
        WindowProvider._listTime := 0
    }
}
