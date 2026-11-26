;===============================================================================
; Arke.ahk - asks the Arke file manager (0.8.0 or later; 0.9.1+ leaves out the selection) for its folders and to open one (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Arke's integration interface (docs/integration.md in github.com/Koimsurai/Arke): a request is WM_COPYDATA with
; dwData = "ARKE" and UTF-8 JSON; Arke replies while handling it, as Total Commander does, with a WM_COPYDATA (same dwData,
; UTF-8 JSON, no terminating 0) to the window in wParam, so no process is started and the clipboard isn't touched.
; Its main windows have the class ArkeWindow (menus and other popups have ArkeWindow.Popup).
; Format version 1 only grows (new fields, location kinds, error codes), so unknown fields are ignored.
;
; Usage:
;   Arke.Windows()                    [{Hwnd, Folder, Other}] front to back; "" for what isn't a folder
;   Arke.Folder(hwnd := 0)            the folder of that Arke window (0: the one used last)
;   Arke.Open(path, newTab, previous) shows a folder (a file: its folder, the file selected); false if Arke didn't take it
;===============================================================================

class Arke {
    static Id := 0x454B5241                                                    ; "ARKE"
    static _waiting := false, _reply := "", _listening := false
    static _withoutSelection := true                                            ; state asked with "selection":false (Arke 0.9.1+)

    ; The reply to a JSON request, "" if Arke doesn't answer (not running, older than 0.8.0, hung, or running as administrator:
    ; Windows doesn't pass WM_COPYDATA up to a program with higher rights)
    static Ask(hwnd, request, timeoutMs := 500) {
        if !Arke._listening {
            OnMessage(0x4A, (p*) => Arke._OnReply(p*))                         ; WM_COPYDATA
            Arke._listening := true
        }
        data := Buffer(StrPut(request, "UTF-8"))
        size := StrPut(request, data, "UTF-8") - 1                              ; bytes, without the terminating 0 (optional for Arke)
        copyData := Buffer(A_PtrSize * 3, 0)                                    ; COPYDATASTRUCT
        NumPut("UPtr", Arke.Id, copyData, 0)
        NumPut("UInt", size, copyData, A_PtrSize)
        NumPut("Ptr", data.Ptr, copyData, A_PtrSize * 2)
        Arke._reply := "", Arke._waiting := true, accepted := 0
        DllCall("SendMessageTimeoutW", "Ptr", hwnd, "UInt", 0x4A, "Ptr", A_ScriptHwnd, "Ptr", copyData.Ptr
              , "UInt", 2, "UInt", timeoutMs, "UPtr*", &accepted)                 ; 2 = SMTO_ABORTIFHUNG: don't wait when Arke hangs
        Arke._waiting := false                                                   ; the reply comes while Arke handles the message
        return Arke._reply
    }

    static _OnReply(wParam, lParam, msg, hwnd) {
        if (!Arke._waiting || NumGet(lParam, 0, "UPtr") != Arke.Id)
            return                                                              ; not the reply waited for: leave it to other handlers
        size := NumGet(lParam, A_PtrSize, "UInt"), data := NumGet(lParam, A_PtrSize * 2, "Ptr")
        Arke._reply := (data && size) ? StrGet(data, size, "UTF-8") : ""        ; size in bytes
        Arke._waiting := false
        return true
    }

    ; The Arke windows front to back: [{Hwnd, Folder, Other}]; Folder = the active pane of the window's active tab, Other = the
    ; other pane in split view. [] when no Arke window is open or it doesn't answer. Any Arke window answers for all of them
    static Windows() {
        if !(hwnd := WinExist("ahk_class ArkeWindow"))
            return []
        started := Logger.Ms()
        ; Without the selection of every pane (Arke 0.9.1+): it isn't used, and thousands of selected files make the reply big and slow
        ; to read. An Arke that answers it with an error is asked as before, and from then on in this run (unless that fails too)
        reply := Arke.Ask(hwnd, Arke._withoutSelection ? '{"v":1,"cmd":"state","selection":false}' : '{"v":1,"cmd":"state"}')
        if (Arke._withoutSelection && Arke._IsError(reply)) {
            again := Arke.Ask(hwnd, '{"v":1,"cmd":"state"}')
            if (again != "" && !Arke._IsError(again))
                Arke._withoutSelection := false, reply := again
        }
        windows := Arke.ParseState(reply)
        if Logger.Enabled
            Logger.Debug(Format("Perf: Arke state {} ms, {} windows", Round(Logger.Ms() - started), windows.Length))
        return windows
    }

    ; {"v":1,"windows":[{"hwnd":"0x1A2B3C","active":true,"tabs":[{"active":true,"panes":[{"active":true,"location":{"kind":"folder",
    ; "path":"D:\\Work"},"selection":[...]}, ...]}]}]}: windows in Z-order, one active tab per window, one active pane per tab.
    ; Only location kind "folder" (a folder, drive or share) counts; home, search results, the Recycle Bin, SFTP / WebDAV and
    ; other pages are no folder ("")
    static ParseState(text) {
        windows := []
        try state := JSON.Parse(text)
        catch
            return windows
        for window in Arke._List(state, "windows") {
            folder := "", other := ""
            for tab in Arke._List(window, "tabs") {
                if !Arke._Active(tab)
                    continue
                for pane in Arke._List(tab, "panes") {
                    if Arke._Active(pane)
                        folder := Arke._PaneFolder(pane)
                    else if (other = "")
                        other := Arke._PaneFolder(pane)
                }
                break
            }
            hwnd := (window is Map && window.Has("hwnd")) ? window["hwnd"] : ""
            windows.Push({Hwnd: IsInteger(hwnd) ? Integer(hwnd) : 0, Folder: folder, Other: other})
        }
        return windows
    }

    static _IsError(text) {
        try reply := JSON.Parse(text)
        catch
            return false
        return reply is Map && reply.Has("error")
    }

    static _List(obj, key) => (obj is Map && obj.Has(key) && obj[key] is Array) ? obj[key] : []
    static _Active(obj) => obj is Map && obj.Has("active") && obj["active"] = 1

    static _PaneFolder(pane) {
        if !(pane is Map && pane.Has("location") && pane["location"] is Map)
            return ""
        location := pane["location"]
        if !(location.Has("kind") && location["kind"] = "folder" && location.Has("path"))
            return ""
        folder := RTrim(location["path"], "\")
        if TextTools.Match(folder, "^[A-Za-z]:$")
            return folder "\"                                                   ; a drive: D:\ (D: alone means the current folder on D)
        return TextTools.Match(folder, "^([A-Za-z]:\\|\\\\)") ? folder : ""
    }

    ; The folder of a given Arke window (hwnd), or of the one used last (hwnd = 0, or a window the reply doesn't list)
    static Folder(hwnd := 0, otherPane := false) {
        windows := Arke.Windows()
        for window in windows
            if (hwnd && window.Hwnd = hwnd)
                return otherPane ? window.Other : window.Folder
        return windows.Length ? (otherPane ? windows[1].Other : windows[1].Folder) : ""
    }

    ; Shows path in an Arke window: previous when it is one (the window in front before ALTRun was shown), else the one used last.
    ; A folder, or a file's folder with the file selected; newTab = in a new tab. Arke brings its window to the front itself, so it
    ; is allowed to first. false when no Arke window is open or Arke didn't take the request (older than 0.8.0, a path that
    ; doesn't exist...): open it another way
    static Open(path, newTab := false, previous := 0) {
        className := ""
        try className := previous ? WinGetClass("ahk_id " previous) : ""
        if !(hwnd := (className = "ArkeWindow") ? previous : WinExist("ahk_class ArkeWindow"))
            return false
        try DllCall("AllowSetForegroundWindow", "UInt", WinGetPID("ahk_id " hwnd))
        request := '{"v":1,"cmd":"open","path":"' JSON.Escape(path) '"' (newTab ? ',"new_tab":true' : "") "}"
        reply := Arke.Ask(hwnd, request, 2000)                                  ; Arke checks that the path exists (a share may take a moment)
        try {
            answer := JSON.Parse(reply)
            if (answer is Map && answer.Has("ok") && answer["ok"] = 1)
                return true
        }
        Logger.Debug("Arke: didn't open " path " - " (reply = "" ? "no reply" : reply))
        return false
    }
}
