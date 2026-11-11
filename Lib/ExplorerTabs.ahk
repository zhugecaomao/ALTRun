;===============================================================================
; ExplorerTabs.ahk - open a folder in a new tab of a File Explorer window that is already open (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Explorer has tabs since Windows 11 22H2 (build 22621) but no switch or API to open one. What works (probed on Windows CI,
; Windows Server 2025 24H2, see Tests\Integration\ExplorerTabsCheck.ahk):
;   1. WM_COMMAND 0xA21B to the window's tab view (ShellTabWindowClass1) opens a new tab, also when the window isn't in front
;      (the command behind Ctrl+T; not documented, ExplorerTabUtility and others use it)
;   2. the new tab is the item of Shell.Application.Windows whose tab view handle (IShellBrowser::GetWindow) wasn't there before
;      (counting items would be fooled by other windows opening and closing meanwhile)
;   3. Navigate2 with the folder's PIDL as a byte array reaches any folder: no URL, so spaces, # % and Chinese need no escaping
;   4. Document.SelectItem selects a file in it (Reveal)
; Tried and not working: explorer.exe has no switch for it; Navigate2's navOpenInNewTab replaces the current tab;
; BrowseObject(SBSP_NEWBROWSER) opens no tab; the "opennewtab" verb through ShellExecute opens a new window; keys (Ctrl+T, Ctrl+L,
; typing the path) depend on the focus, the input method and tab bar add-ons. Thanks to the ideas in #440 and AHK forum t=141197 #47.
; Anything that fails returns false (LastError says why) and the caller opens a new window as before; an empty tab this
; opened is closed again.
;
; Usage:
;   ExplorerTabs.Supported()                          Windows 11 22H2 or later
;   ExplorerTabs.Open(folder [, file, preferred])     true = shown in a new tab of an open window (file selected in it);
;                                                     preferred: an Explorer window to use first (e.g. the one in front before)
;   ExplorerTabs.LastError                            why the last Open returned false
;===============================================================================

class ExplorerTabs {
    static MinVersion := "10.0.22621"
    static TabTimeout := 3000, NavigateTimeout := 3000                    ; ms; on CI a tab came in 100-800 ms, a navigation in 300-550 ms
    static SelectTimeout := 1500
    static LastError := ""

    static Supported(version := A_OSVersion) => VerCompare(version, ExplorerTabs.MinVersion) >= 0

    static Open(folder, file := "", preferred := 0) {
        ExplorerTabs.LastError := ""
        if !ExplorerTabs.Supported()
            return ExplorerTabs._Fail("no tabs before Windows 11 22H2")
        if !(hwnd := ExplorerTabs.Target(preferred))
            return ExplorerTabs._Fail("no Explorer window open")
        try {
            before := ExplorerTabs._TabSet(hwnd)
            SendMessage(0x0111, 0xA21B, 0, "ShellTabWindowClass1", "ahk_id " hwnd, , , , 2000)   ; WM_COMMAND: new tab
            if !(tab := ExplorerTabs._WaitNewTab(hwnd, before))
                return ExplorerTabs._Fail("no new tab came")
            if !(ExplorerTabs._Navigate(hwnd, tab, folder) && ExplorerTabs._WaitPath(hwnd, tab, folder)) {
                ExplorerTabs._CloseTab(hwnd, tab)                           ; don't leave an empty tab behind
                return ExplorerTabs._Fail("the new tab didn't go to the folder")
            }
            if (file != "")
                ExplorerTabs._Select(hwnd, tab, file)
            ExplorerTabs._Show(hwnd)
            return true
        } catch as err {
            return ExplorerTabs._Fail(err.Message)
        }
    }

    ; The window to use: preferred when it is a visible Explorer window, else the frontmost visible one on the current virtual desktop
    static Target(preferred := 0) {
        if ExplorerTabs._Usable(preferred)
            return preferred
        for hwnd in WinGetList("ahk_class CabinetWClass")                  ; in Z-order, frontmost first
            if ExplorerTabs._Usable(hwnd)
                return hwnd
        return 0
    }

    static _Usable(hwnd) {
        if !(hwnd && WinExist("ahk_id " hwnd) && DllCall("IsWindowVisible", "Ptr", hwnd))
            return false
        try {
            if (WinGetClass("ahk_id " hwnd) != "CabinetWClass")
                return false
        } catch
            return false
        return ExplorerTabs._OnCurrentDesktop(hwnd)
    }

    ; IVirtualDesktopManager::IsWindowOnCurrentVirtualDesktop; true when it can't be told
    static _OnCurrentDesktop(hwnd) {
        static manager := ""
        try {
            if !IsObject(manager)
                manager := ComObject("{AA509086-5CA9-4C25-8F95-589D3C07B48A}", "{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}")
            onCurrent := 1
            ComCall(3, manager, "Ptr", hwnd, "Int*", &onCurrent)
            return onCurrent != 0
        }
        return true
    }

    ; [{Tab: tab view handle, Item}] of the window's tabs (Shell.Application.Windows has an item for every tab)
    static _Tabs(hwnd) {
        tabs := []
        for item in ComObject("Shell.Application").Windows {
            try {
                if (item.HWND != hwnd)
                    continue
            } catch
                continue
            tabs.Push({Tab: ExplorerTabs._TabOf(item), Item: item})
        }
        return tabs
    }

    static _TabSet(hwnd) {
        set := Map()
        for each in ExplorerTabs._Tabs(hwnd)
            if each.Tab
                set[each.Tab] := true
        return set
    }

    static _Item(hwnd, tab) {
        for each in ExplorerTabs._Tabs(hwnd)
            if (each.Tab = tab)
                return each.Item
        return ""
    }

    ; The tab view handle of an item: IShellBrowser's GetWindow (IOleWindow, the fourth method)
    static _TabOf(item) {
        static IID_IShellBrowser := "{000214E2-0000-0000-C000-000000000046}"
        try {
            browser := ComObjQuery(item, IID_IShellBrowser, IID_IShellBrowser)
            tab := 0
            ComCall(3, browser, "Ptr*", &tab)
            return tab
        }
        return 0
    }

    static _WaitNewTab(hwnd, before) {
        deadline := A_TickCount + ExplorerTabs.TabTimeout
        while (A_TickCount < deadline) {
            for each in ExplorerTabs._Tabs(hwnd)
                if (each.Tab && !before.Has(each.Tab))
                    return each.Tab
            Sleep(50)
        }
        return 0
    }

    ; Navigate2 with the folder's PIDL (VT_ARRAY | VT_UI1), which only moves that tab
    static _Navigate(hwnd, tab, folder) {
        if !IsObject(item := ExplorerTabs._Item(hwnd, tab))
            return false
        if !IsObject(bytes := ExplorerTabs.PidlBytes(folder))
            return false
        item.Navigate2(bytes)
        return true
    }

    ; The PIDL of a path as a COM byte array, "" when the path can't be parsed
    static PidlBytes(path) {
        pidl := 0
        if (DllCall("shell32\SHParseDisplayName", "WStr", path, "Ptr", 0, "Ptr*", &pidl, "UInt", 0, "Ptr", 0, "Int") < 0 || !pidl)
            return ""
        try {
            size := DllCall("shell32\ILGetSize", "Ptr", pidl, "UInt")
            bytes := ComObjArray(0x11, size)                                ; VT_UI1
            Loop size
                bytes[A_Index - 1] := NumGet(pidl, A_Index - 1, "UChar")
            return bytes
        } finally
            DllCall("ole32\CoTaskMemFree", "Ptr", pidl)
    }

    static _WaitPath(hwnd, tab, folder) {
        deadline := A_TickCount + ExplorerTabs.NavigateTimeout
        while (A_TickCount < deadline) {
            if ExplorerTabs.SameFolder(ExplorerTabs._Path(hwnd, tab), folder)
                return true
            Sleep(50)
        }
        return false
    }

    static _Path(hwnd, tab) {
        try return ExplorerTabs._Item(hwnd, tab).Document.Folder.Self.Path
        return ""
    }

    static SameFolder(a, b) => (a != "" && StrLower(RTrim(a, "\ ")) = StrLower(RTrim(b, "\ ")))

    ; Select a file of the folder in that tab (select, deselect the others, make it visible, focus it). A tab that has only just
    ; reached the folder may not list its files yet and ignore it (seen on CI with the first tab), so it is checked and tried
    ; again for a while; true = selected
    static _Select(hwnd, tab, file) {
        deadline := A_TickCount + ExplorerTabs.SelectTimeout
        loop {
            try {
                doc := ExplorerTabs._Item(hwnd, tab).Document
                if IsObject(fileItem := doc.Folder.ParseName(file)) {
                    doc.SelectItem(fileItem, 1 | 4 | 8 | 16)
                    chosen := doc.SelectedItems()
                    if (chosen.Count = 1 && chosen.Item(0).Path = fileItem.Path)
                        return true
                }
            }
            if (A_TickCount >= deadline)
                return false
            Sleep(100)
        }
    }

    static _CloseTab(hwnd, tab) {
        try ExplorerTabs._Item(hwnd, tab).Quit()                          ; IWebBrowser2.Quit on a tab's item closes that tab only
    }

    static _Show(hwnd) {
        try {
            if (WinGetMinMax("ahk_id " hwnd) = -1)
                WinRestore("ahk_id " hwnd)
            WinActivate("ahk_id " hwnd)
        }
    }

    static _Fail(reason) {
        ExplorerTabs.LastError := reason
        return false
    }
}
