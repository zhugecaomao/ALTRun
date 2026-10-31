;===============================================================================
; FakeEverything.ahk - a stand-in for Everything's IPC window in the tests (AutoHotkey v2)
;-------------------------------------------------------------------------------
; A hidden top-level window with Everything's class that answers EVERYTHING_IPC_COPYDATA_QUERY2W queries the way
; Everything does: an EVERYTHING_IPC_LIST2 sent back with WM_COPYDATA to the window that asked, under the query's reply
; id. Lib\Everything.ahk finds it with FindWindow, as it finds Everything. Checked against the real Everything 1.4 and
; 1.5 by .github/workflows/everything.yml.
;
; Usage:
;   fake := FakeEverything([{Path: "C:\Work\Report.docx", IsFolder: false}])
;   fake.Delay := 50                ; reply after 50 ms (0 = at once, while the query is being sent; < 0 = never)
;   fake.StaleFirst := true         ; first a reply with an older id (a search the user has typed past)
;   ... Everything.Query("report") ...   fake.Queries[1].Search = "report"
;   fake.Close()                    ; always (try / finally): the class can be registered only once
;===============================================================================

class FakeEverything {
    Results := [], Delay := 0, StaleFirst := false, Queries := []
    Hwnd := 0, _class := "", _proc := 0

    __New(results := [], className := "EVERYTHING_TASKBAR_NOTIFICATION") {
        this.Results := results
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
            throw Error("FakeEverything: RegisterClassExW failed (" A_LastError ")")
        this.Hwnd := DllCall("CreateWindowExW", "UInt", 0, "Ptr", this._class.Ptr, "Str", "Everything", "UInt", 0
            , "Int", 0, "Int", 0, "Int", 0, "Int", 0, "Ptr", 0, "Ptr", 0, "Ptr", instance, "Ptr", 0, "Ptr")
        if !this.Hwnd
            throw Error("FakeEverything: CreateWindowExW failed (" A_LastError ")")
    }

    Close() {
        if this.Hwnd
            DllCall("DestroyWindow", "Ptr", this.Hwnd), this.Hwnd := 0
        if this._class
            DllCall("UnregisterClassW", "Ptr", this._class.Ptr, "Ptr", DllCall("GetModuleHandle", "Ptr", 0, "Ptr")), this._class := ""
        if this._proc
            CallbackFree(this._proc), this._proc := 0
    }

    ; EVERYTHING_IPC_QUERY2: reply_hwnd, reply_copydata_message, search_flags, offset, max_results, request_flags,
    ; sort_type (7 DWORDs), then the search text
    _WndProc(hwnd, msg, wParam, lParam) {
        if (msg != 0x4A || NumGet(lParam, 0, "UPtr") != 18)                ; WM_COPYDATA, EVERYTHING_IPC_COPYDATA_QUERY2W
            return DllCall("DefWindowProcW", "Ptr", hwnd, "UInt", msg, "UPtr", wParam, "Ptr", lParam, "Ptr")
        data := NumGet(lParam, A_PtrSize * 2, "Ptr")
        query := {ReplyHwnd: NumGet(data, 0, "UInt"), ReplyId: NumGet(data, 4, "UInt"), Max: NumGet(data, 16, "UInt")
                , RequestFlags: NumGet(data, 20, "UInt"), Sort: NumGet(data, 24, "UInt"), Search: StrGet(data + 28, "UTF-16")}
        this.Queries.Push(query)
        if this.StaleFirst
            this._Reply(query.ReplyHwnd, query.ReplyId - 1, [{Path: "C:\Stale\old result.txt", IsFolder: false}])
        if (this.Delay = 0)
            this._Reply(query.ReplyHwnd, query.ReplyId, this.Results)
        else if (this.Delay > 0)
            SetTimer(() => this._Reply(query.ReplyHwnd, query.ReplyId, this.Results), -this.Delay)
        return 1
    }

    ; EVERYTHING_IPC_LIST2: totitems, numitems, offset, request_flags, sort_type (5 DWORDs), numitems {flags, data_offset},
    ; then at each data_offset a DWORD length and the UTF-16 path (flags 1 = folder)
    _Reply(replyHwnd, replyId, results) {
        size := 20 + results.Length * 8
        for item in results
            size += 4 + (StrLen(item.Path) + 1) * 2
        list := Buffer(size, 0)
        NumPut("UInt", results.Length, "UInt", results.Length, list, 0)
        offset := 20 + results.Length * 8
        for index, item in results {
            NumPut("UInt", item.IsFolder ? 1 : 0, "UInt", offset, list, 20 + (index - 1) * 8)
            NumPut("UInt", StrLen(item.Path), list, offset)
            StrPut(item.Path, list.Ptr + offset + 4, "UTF-16")
            offset += 4 + (StrLen(item.Path) + 1) * 2
        }
        copyData := Buffer(A_PtrSize * 3, 0)
        NumPut("UPtr", replyId, copyData, 0)
        NumPut("UInt", size, copyData, A_PtrSize)
        NumPut("Ptr", list.Ptr, copyData, A_PtrSize * 2)
        DllCall("SendMessageW", "Ptr", replyHwnd, "UInt", 0x4A, "Ptr", this.Hwnd, "Ptr", copyData.Ptr, "Ptr")
    }
}
