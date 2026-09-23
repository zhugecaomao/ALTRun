;===============================================================================
; Everything.ahk - 通过 IPC 直接查询 Everything (voidtools) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 不需要 Everything64.dll / es.exe: 只要 Everything 在运行, 就用它公开的 IPC
; 接口 (WM_COPYDATA, EVERYTHING_IPC_COPYDATA_QUERY2W) 查询, 和官方 SDK 做的事
; 一样。Everything 把结果用 WM_COPYDATA 发回我们的窗口 (A_ScriptHwnd)。
;
; 用法:
;   if Everything.IsRunning()
;       for item in Everything.Query("report", 30)      ; -> [{Path, IsFolder}]
;
; 排序常量 (sort 参数): 1 = 名称升序, 14 = 修改日期降序 (最近修改的在前)
;===============================================================================

class Everything {
    static SORT_NAME := 1, SORT_DATE_MODIFIED_DESC := 14
    static _replyId := 0x414C5400, _pending := 0, _results := "", _listening := false

    static IsRunning() {
        return DllCall("FindWindowW", "Str", "EVERYTHING_TASKBAR_NOTIFICATION", "Ptr", 0, "Ptr") != 0
    }

    ; 返回 [{Path, IsFolder}], Everything 没运行或超时返回空数组
    static Query(search, maxResults := 30, sort := 14, timeoutMs := 1000) {
        everythingHwnd := DllCall("FindWindowW", "Str", "EVERYTHING_TASKBAR_NOTIFICATION", "Ptr", 0, "Ptr")
        if !everythingHwnd
            return []
        if !Everything._listening {
            OnMessage(0x4A, (p*) => Everything._OnCopyData(p*))                ; WM_COPYDATA
            Everything._listening := true
        }

        Everything._pending := ++Everything._replyId
        Everything._results := ""

        ; EVERYTHING_IPC_QUERY2: 7 个 DWORD + UTF-16 搜索文字 (带结尾 0)
        size := 28 + (StrLen(search) + 1) * 2
        query := Buffer(size, 0)
        NumPut("UInt", A_ScriptHwnd, "UInt", Everything._pending, "UInt", 0, "UInt", 0
             , "UInt", maxResults, "UInt", 0x4, "UInt", sort, query)            ; 0x4 = REQUEST_FULL_PATH_AND_NAME
        StrPut(search, query.Ptr + 28, "UTF-16")

        copyData := Buffer(A_PtrSize * 3, 0)                                    ; COPYDATASTRUCT
        NumPut("UPtr", 18, copyData, 0)                                         ; EVERYTHING_IPC_COPYDATA_QUERY2W
        NumPut("UInt", size, copyData, A_PtrSize)
        NumPut("Ptr", query.Ptr, copyData, A_PtrSize * 2)
        ; Everything 的 IPC 窗口是隐藏的, AHK 的 SendMessage("ahk_id ...") 默认找不到隐藏窗口,
        ; 所以直接调用 SendMessageTimeout (2 = SMTO_ABORTIFHUNG)
        accepted := 0
        if !DllCall("SendMessageTimeoutW", "Ptr", everythingHwnd, "UInt", 0x4A, "Ptr", A_ScriptHwnd, "Ptr", copyData.Ptr
                  , "UInt", 2, "UInt", timeoutMs, "UPtr*", &accepted) || !accepted
            return []

        deadline := A_TickCount + timeoutMs
        while (!IsObject(Everything._results) && A_TickCount < deadline)
            Sleep(5)
        results := IsObject(Everything._results) ? Everything._results : []
        Everything._pending := 0
        return results
    }

    ; Everything 发回的 EVERYTHING_IPC_LIST2:
    ;   totitems, numitems, offset, request_flags, sort_type (5 个 DWORD)
    ;   然后 numitems 个 {flags, data_offset}; data_offset 处是 DWORD 长度 + UTF-16 路径
    static _OnCopyData(wParam, lParam, msg, hwnd) {
        replyId := NumGet(lParam, 0, "UPtr")
        if (!Everything._pending || replyId != Everything._pending)
            return
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
