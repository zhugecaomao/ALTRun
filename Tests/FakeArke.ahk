;===============================================================================
; FakeArke.ahk - a stand-in for an Arke window in the tests (AutoHotkey v2)
;-------------------------------------------------------------------------------
; A tiny visible top-level window with Arke's class (ArkeWindow) that answers WM_COPYDATA requests (dwData = "ARKE", UTF-8 JSON)
; the way Arke 0.8.0 does (docs/integration.md in the Arke repository): while handling the message, a WM_COPYDATA with the same
; dwData and the UTF-8 JSON reply, without a terminating 0, to the window in wParam. Lib\Arke.ahk finds it with WinExist.
;
; Usage:
;   fake := FakeArke('{"v":1,"windows":[...]}')    ; the reply to "state"
;   fake.OpenReply := '{"v":1,"ok":true}'          ; the reply to "open" ("" = don't answer, as Arke before 0.8.0)
;   fake.SelectionReply := '{"v":1,"error":...}'    ; the reply to "state" with "selection":false ("" = as to "state")
;   ... Arke.Folder() ...   fake.Requests[1] = '{"v":1,"cmd":"state","selection":false}'
;   fake.Close()                                    ; always (try / finally): the class can be registered only once
;===============================================================================

class FakeArke {
    State := "", OpenReply := '{"v":1,"ok":true}', SelectionReply := "", Requests := []
    Hwnd := 0, _class := "", _proc := 0

    __New(state := "") {
        this.State := state
        className := "ArkeWindow"
        this._class := Buffer(StrPut(className, "UTF-16") * 2)
        StrPut(className, this._class, "UTF-16")
        this._proc := CallbackCreate(this._WndProc.Bind(this), "F", 4)
        p := A_PtrSize, instance := DllCall("GetModuleHandle", "Ptr", 0, "Ptr")
        wc := Buffer(16 + 8 * p, 0)                                         ; WNDCLASSEXW
        NumPut("UInt", wc.Size, wc, 0)
        NumPut("Ptr", this._proc, wc, 8)
        NumPut("Ptr", instance, wc, 16 + p)
        NumPut("Ptr", this._class.Ptr, wc, 16 + 6 * p)
        if !DllCall("RegisterClassExW", "Ptr", wc, "UShort")
            throw Error("FakeArke: RegisterClassExW failed (" A_LastError ")")
        ; WS_POPUP | WS_VISIBLE (WinExist finds only visible windows); WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE: no taskbar button, no focus
        this.Hwnd := DllCall("CreateWindowExW", "UInt", 0x08000080, "Ptr", this._class.Ptr, "Str", "ALTRun Fake Arke", "UInt", 0x90000000
            , "Int", 0, "Int", 0, "Int", 1, "Int", 1, "Ptr", 0, "Ptr", 0, "Ptr", instance, "Ptr", 0, "Ptr")
        if !this.Hwnd
            throw Error("FakeArke: CreateWindowExW failed (" A_LastError ")")
    }

    Close() {
        if this.Hwnd
            DllCall("DestroyWindow", "Ptr", this.Hwnd), this.Hwnd := 0
        if this._class
            DllCall("UnregisterClassW", "Ptr", this._class.Ptr, "Ptr", DllCall("GetModuleHandle", "Ptr", 0, "Ptr")), this._class := ""
        if this._proc
            CallbackFree(this._proc), this._proc := 0
    }

    _WndProc(hwnd, msg, wParam, lParam) {
        if (msg != 0x4A || NumGet(lParam, 0, "UPtr") != 0x454B5241)        ; WM_COPYDATA, "ARKE"
            return DllCall("DefWindowProcW", "Ptr", hwnd, "UInt", msg, "UPtr", wParam, "Ptr", lParam, "Ptr")
        request := StrGet(NumGet(lParam, A_PtrSize * 2, "Ptr"), NumGet(lParam, A_PtrSize, "UInt"), "UTF-8")
        this.Requests.Push(request)
        reply := InStr(request, '"cmd":"state"') ? ((this.SelectionReply != "" && InStr(request, '"selection":false')) ? this.SelectionReply : this.State)
            : InStr(request, '"cmd":"open"') ? this.OpenReply : ""
        if (reply = "")
            return 0
        data := Buffer(StrPut(reply, "UTF-8"))
        size := StrPut(reply, data, "UTF-8") - 1                            ; no terminating 0, as Arke
        copyData := Buffer(A_PtrSize * 3, 0)
        NumPut("UPtr", 0x454B5241, copyData, 0)
        NumPut("UInt", size, copyData, A_PtrSize)
        NumPut("Ptr", data.Ptr, copyData, A_PtrSize * 2)
        DllCall("SendMessageW", "Ptr", wParam, "UInt", 0x4A, "Ptr", hwnd, "Ptr", copyData.Ptr, "Ptr")
        return 1
    }
}
