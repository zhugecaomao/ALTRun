;===============================================================================
; BackgroundWork.ahk - when background work may run (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Index scans, cache writes and icon prefetching wait while the search window is
; open or opening, so typing never stalls on them. The search window sets IsBusy
; when it is created; Core and the providers only ask CanRun() and don't need to
; know about the window.
;
; Usage:
;   BackgroundWork.IsBusy := () => ...     true while background work must wait
;   if !BackgroundWork.CanRun()
;       return                             try again on the next timer tick
;===============================================================================

class BackgroundWork {
    static IsBusy := ""                                                     ; not set (tests, before start-up): never busy

    static CanRun() => !(IsObject(BackgroundWork.IsBusy) && BackgroundWork.IsBusy.Call())
}
