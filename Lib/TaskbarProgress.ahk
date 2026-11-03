;===============================================================================
; TaskbarProgress.ahk - progress on a window's taskbar button (ITaskbarList3) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Shows how far a long task is on the window's taskbar button, the way Explorer does while copying: green for normal, red for
; an error, yellow for paused (stopped), moving back and forth when the total isn't known. The window needs a taskbar button
; (a window without an owner). Windows ignores these calls until the button exists, and forgets them when Explorer restarts:
; both times it sends the window "TaskbarButtonCreated", and the last state set is set again then. Where ITaskbarList3 isn't
; there (Wine, Windows without Explorer), nothing happens.
;
; Usage:
;   TaskbarProgress.Set(hwnd, "Normal", 500)      ; 500 of 1000; states: None, Indeterminate, Normal, Error, Paused
;   TaskbarProgress.Set(hwnd, "None")             ; no progress (also forgets the window)
;===============================================================================

class TaskbarProgress {
    static States := Map("None", 0, "Indeterminate", 1, "Normal", 2, "Error", 4, "Paused", 8)   ; TBPF_*
    static Last := Map()                                                    ; hwnd -> [state, value, total] set last
    static _taskbarList := ""                                               ; the ITaskbarList3, "" until first used, 0 when not there
    static _createdFn := (wParam, lParam, msg, hwnd) => TaskbarProgress._Created(hwnd)

    ; value: 0-total (ignored for None and Indeterminate)
    static Set(hwnd, state, value := 0, total := 1000) {
        if !(list := TaskbarProgress._List())
            return false
        if (state != "None")
            TaskbarProgress.Last[hwnd] := [state, value, total]
        else if TaskbarProgress.Last.Has(hwnd)
            TaskbarProgress.Last.Delete(hwnd)
        return TaskbarProgress._Apply(list, hwnd, state, value, total)
    }

    static _Apply(list, hwnd, state, value, total) {
        try {
            ComCall(10, list, "Ptr", hwnd, "Int", TaskbarProgress.States[state])            ; SetProgressState
            if (state != "None" && state != "Indeterminate")
                ComCall(9, list, "Ptr", hwnd, "UInt64", Max(0, Min(value, total)), "UInt64", total)   ; SetProgressValue
            return true
        }
        return false
    }

    ; The button was (re)made: what was set before it existed
    static _Created(hwnd) {
        if !TaskbarProgress.Last.Has(hwnd)
            return
        if !WinExist("ahk_id " hwnd)
            return TaskbarProgress.Last.Delete(hwnd)
        last := TaskbarProgress.Last[hwnd]
        TaskbarProgress._Apply(TaskbarProgress._taskbarList, hwnd, last*)
    }

    static _List() {
        if (TaskbarProgress._taskbarList = "") {
            TaskbarProgress._taskbarList := 0
            try {
                list := ComObject("{56FDF344-FD6D-11d0-958A-006097C9A090}", "{ea1afb91-9e28-4b86-90e9-9e9f8a5eefaf}")
                ComCall(3, list)                                            ; HrInit
                TaskbarProgress._taskbarList := list
                OnMessage(DllCall("RegisterWindowMessage", "Str", "TaskbarButtonCreated", "UInt"), TaskbarProgress._createdFn)
            }
        }
        return TaskbarProgress._taskbarList
    }
}
