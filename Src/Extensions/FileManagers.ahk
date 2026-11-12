;===============================================================================
; FileManagers.ahk - the current folder of file managers (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Reads, without touching the clipboard, the folder open in Total Commander (its WM_COPYDATA interface), Directory Opus
; (dopusrt.exe /info) and File Explorer windows, which of them was used last, and Windows' recently used folders.
; Used by Quick Switch, "terminal here" and the result actions.
;
; Usage:
;   FileManagers.FileManagerFolder()        {Path, Tag} of the file manager used last
;   FileManagers.FolderOfWindow(hwnd)       current folder of a TC / Opus / Explorer window (read only, for "terminal here")
;   FileManagers.RecentFolders(n)           the n most recently used folders (opened files count by their folder)
;===============================================================================

class FileManagers {
    ; The folder of a Total Commander panel: asked directly through TC's WM_COPYDATA interface (since TC 8.0), without touching the clipboard.
    ; hwnd = 0 takes the frontmost TC window; otherPanel = true takes the other panel
    static TotalCmdFolder(hwnd := 0, otherPanel := false) {
        if !(hwnd := hwnd ? hwnd : WinExist("ahk_class TTOTAL_CMD"))
            return ""
        folder := RTrim(FileManagers.TotalCmdAsk(hwnd, otherPanel ? "TP" : "SP"), "\")
        return (folder ~= "^([A-Za-z]:|\\\\)") ? folder : ""               ; only real paths (no FTP, inside archives etc.)
    }

    ; Sends WM_COPYDATA (dwData = "GW", an ANSI command) to TC, which replies at once with WM_COPYDATA (dwData = "RW", UTF-16)
    ; to the window in wParam. Commands: A = which side is active (L / R); two letters: first L left / R right / S active / T the other,
    ; second P path / C file count / I caret position / N file name under the caret. "" if TC doesn't reply
    static _tcWaiting := false, _tcReply := "", _tcListening := false
    static TotalCmdAsk(hwnd, command, timeoutMs := 500) {
        if !FileManagers._tcListening {
            OnMessage(0x4A, (p*) => FileManagers._OnTotalCmdReply(p*))         ; WM_COPYDATA
            FileManagers._tcListening := true
        }
        request := Buffer(StrPut(command, "CP0"))
        StrPut(command, request, "CP0")
        copyData := Buffer(A_PtrSize * 3, 0)                                    ; COPYDATASTRUCT
        NumPut("UPtr", Ord("G") + 256 * Ord("W"), copyData, 0)
        NumPut("UInt", request.Size, copyData, A_PtrSize)
        NumPut("Ptr", request.Ptr, copyData, A_PtrSize * 2)
        FileManagers._tcReply := "", FileManagers._tcWaiting := true, accepted := 0
        DllCall("SendMessageTimeoutW", "Ptr", hwnd, "UInt", 0x4A, "Ptr", A_ScriptHwnd, "Ptr", copyData.Ptr
              , "UInt", 2, "UInt", timeoutMs, "UPtr*", &accepted)                 ; 2 = SMTO_ABORTIFHUNG: don't wait when TC hangs
        FileManagers._tcWaiting := false                                         ; TC replies while handling this message, so no extra wait
        return FileManagers._tcReply
    }

    static _OnTotalCmdReply(wParam, lParam, msg, hwnd) {
        if (!FileManagers._tcWaiting || NumGet(lParam, 0, "UPtr") != Ord("R") + 256 * Ord("W"))
            return                                                              ; not the reply waited for: leave it to other handlers (Everything)
        FileManagers._tcReply := Win.CopyDataText(lParam), FileManagers._tcWaiting := false
        return true
    }

    ; The newest limit folders of Windows "Recent items" (WindowsRecent); a file gives its folder.
    ; Only the newest limit * 6 shortcuts are looked at, not all of the hundreds there may be
    static RecentFolders(limit, recentDir := "") {
        folders := [], seen := Map()
        if (limit <= 0)
            return folders
        for item in WindowsRecent.Items(limit * 6, recentDir) {
            folder := RTrim(item.IsFolder ? item.Path : FileManagers._Parent(item.Path), "\")
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
                case "TTOTAL_CMD"   : return {Path: FileManagers.TotalCmdFolder(hwnd), Tag: "Total Commander", Hwnd: hwnd}
                case "dopus.lister" : return {Path: FileManagers.DOpusFolder(hwnd), Tag: "Directory Opus", Hwnd: hwnd}
                case "CabinetWClass": return {Path: FileManagers.ExplorerFolder(hwnd), Tag: I18n.T("QuickSwitch.TagExplorer"), Hwnd: hwnd}
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
            case "TTOTAL_CMD"   : return FileManagers.TotalCmdFolder(hwnd)
            case "dopus.lister" : return FileManagers.DOpusFolder(hwnd)
            case "CabinetWClass": return FileManagers.ExplorerFolder(hwnd)
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
    static DOpusFolder(hwnd := 0) => FileManagers.PickDOpusFolder(FileManagers.DOpusTabs(), hwnd)

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
        if (FileManagers._dopusAt && A_TickCount - FileManagers._dopusAt < FileManagers.DOpusCacheMs)
            return FileManagers._dopusTabs
        started := Logger.Ms(), tabs := []
        try {
            SplitPath(WinGetProcessPath("ahk_id " lister), , &dir)
            xmlFile := A_Temp "\ALTRun.dopus-paths.xml"
            try FileDelete(xmlFile)
            Run('"' dir '\dopusrt.exe" /info "' xmlFile '",paths', , "Hide", &pid)
            if ProcessWaitClose(pid, 3)                                     ; still running after 3 s (returns its PID): never hang the dialog on it
                try ProcessClose(pid)
            tabs := FileManagers.ParseDOpusPaths(FileRead(xmlFile, "UTF-8"))
            try FileDelete(xmlFile)
        } catch as e {
            Logger.Error("QuickSwitch: cannot read the Directory Opus folders - " e.Message)
        }
        FileManagers._dopusTabs := tabs, FileManagers._dopusAt := A_TickCount
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
}
