;===============================================================================
; Everything.ahk - queries Everything (voidtools) directly over IPC (AutoHotkey v2)
;-------------------------------------------------------------------------------
; No Everything64.dll / es.exe needed: as long as Everything runs, it is queried through its public IPC
; interface (WM_COPYDATA, EVERYTHING_IPC_COPYDATA_QUERY2W), the same thing the official SDK does.
; Everything sends the results back to our window (A_ScriptHwnd) with WM_COPYDATA.
;
; Usage:
;   if Everything.IsRunning()
;       for item in Everything.Query("report", 30)      ; -> [{Path, IsFolder}]
;
; Sort constants (sort parameter): 1 = name ascending, 14 = date modified descending (most recent first)
;===============================================================================

class Everything {
    static SORT_NAME := 1, SORT_DATE_MODIFIED_DESC := 14
    static _replyId := 0x414C5400, _pending := 0, _results := "", _listening := false
    static LastProblem := "", _foreign := 0                               ; Why the last query came back empty (not received / timed out), for the caller's log

    static IsRunning() {
        return DllCall("FindWindowW", "Str", "EVERYTHING_TASKBAR_NOTIFICATION", "Ptr", 0, "Ptr") != 0
    }

    ; Returns [{Path, IsFolder}]; an empty array when Everything isn't running or times out (the reason is in LastProblem)
    static Query(search, maxResults := 30, sort := 14, timeoutMs := 1000) {
        Everything.LastProblem := ""
        everythingHwnd := DllCall("FindWindowW", "Str", "EVERYTHING_TASKBAR_NOTIFICATION", "Ptr", 0, "Ptr")
        if !everythingHwnd
            return []
        if !Everything._listening {
            OnMessage(0x4A, (p*) => Everything._OnCopyData(p*))                ; WM_COPYDATA
            Everything._listening := true
        }

        Everything._pending := ++Everything._replyId
        Everything._results := ""

        ; EVERYTHING_IPC_QUERY2: 7 DWORDs + the UTF-16 search text (with the terminating 0)
        size := 28 + (StrLen(search) + 1) * 2
        query := Buffer(size, 0)
        NumPut("UInt", A_ScriptHwnd, "UInt", Everything._pending, "UInt", 0, "UInt", 0
             , "UInt", maxResults, "UInt", 0x4, "UInt", sort, query)            ; 0x4 = REQUEST_FULL_PATH_AND_NAME
        StrPut(search, query.Ptr + 28, "UTF-16")

        copyData := Buffer(A_PtrSize * 3, 0)                                    ; COPYDATASTRUCT
        NumPut("UPtr", 18, copyData, 0)                                         ; EVERYTHING_IPC_COPYDATA_QUERY2W
        NumPut("UInt", size, copyData, A_PtrSize)
        NumPut("Ptr", query.Ptr, copyData, A_PtrSize * 2)
        ; Everything's IPC window is hidden and AHK's SendMessage("ahk_id ...") doesn't find hidden windows by default,
        ; so SendMessageTimeout is called directly (2 = SMTO_ABORTIFHUNG)
        accepted := 0, started := A_TickCount
        if !DllCall("SendMessageTimeoutW", "Ptr", everythingHwnd, "UInt", 0x4A, "Ptr", A_ScriptHwnd, "Ptr", copyData.Ptr
                  , "UInt", 2, "UInt", timeoutMs, "UPtr*", &accepted) || !accepted {
            Everything.LastProblem := "query not accepted after " (A_TickCount - started) " ms"
            return []
        }

        Everything._foreign := 0
        deadline := A_TickCount + timeoutMs
        while (!IsObject(Everything._results) && A_TickCount < deadline)
            Sleep(5)
        if !IsObject(Everything._results)                                   ; no reply: Everything is busy, or the reply collided with another WM_COPYDATA
            Everything.LastProblem := "no reply in " (A_TickCount - started) " ms (other WM_COPYDATA while waiting: " Everything._foreign ")"
        results := IsObject(Everything._results) ? Everything._results : []
        Everything._pending := 0
        return results
    }

    ; The EVERYTHING_IPC_LIST2 Everything sends back:
    ;   totitems, numitems, offset, request_flags, sort_type (5 DWORDs)
    ;   then numitems {flags, data_offset}; at data_offset a DWORD length + the UTF-16 path
    static _OnCopyData(wParam, lParam, msg, hwnd) {
        replyId := NumGet(lParam, 0, "UPtr")
        if (!Everything._pending || replyId != Everything._pending) {
            if Everything._pending
                Everything._foreign++
            return
        }
        list := NumGet(lParam, A_PtrSize * 2, "Ptr")
        count := NumGet(list, 4, "UInt")
        results := []
        Loop count {
            itemPtr := list + 20 + (A_Index - 1) * 8
            flags := NumGet(itemPtr, 0, "UInt")
            dataPtr := list + NumGet(itemPtr, 4, "UInt")
            length := NumGet(dataPtr, 0, "UInt")
            results.Push({Path: StrGet(dataPtr + 4, length, "UTF-16"), IsFolder: (flags & 1) ? true : false})
        }
        Everything._results := results
        return true
    }
}
