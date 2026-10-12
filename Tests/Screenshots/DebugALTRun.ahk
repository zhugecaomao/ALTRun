;===============================================================================
; DebugALTRun.ahk - 临时诊断: 结果行空白 (截图脚本复制到 ALTRun 文件夹里运行, 不改 ALTRun 本身)
;-------------------------------------------------------------------------------
; 先挂上钩子再载入 ALTRun.ahk, 把下面这些写进 %Temp%\ALTRunDebug.log (每行前面是启动后的毫秒数):
;   - 每次 NM_CUSTOMDRAW 回调 (阶段、行号、当时的结果数、是否正在读图标)
;   - 每次读图标的开始、结束和耗时
;   - 画行时出的错
;   - 每秒一次心跳 (定时器还能不能运行) 和回调次数
;===============================================================================
#Requires AutoHotkey v2.0

Dbg.Install()

#Include %A_ScriptDir%\ALTRun.ahk

class Dbg {
    static LogFile := A_Temp "\ALTRunDebug.log"
    static Started := A_TickCount
    static Loading := ""                    ; 正在读的图标 (空 = 没有在读)
    static Prepaints := 0, ItemDraws := 0, Logged := 0

    static Install() {
        try FileDelete(Dbg.LogFile)
        Dbg.Log("debug wrapper started, pid " DllCall("GetCurrentProcessId"))
        Dbg.Hook(SearchWindow, "_OnCustomDraw", (original, cls, params*) => Dbg.OnDraw(original, cls, params*))
        Dbg.Hook(SearchWindow, "_PaintRow", (original, cls, params*) => Dbg.OnPaintRow(original, cls, params*))
        Dbg.Hook(IconCache, "_Load", (original, cls, params*) => Dbg.OnLoad(original, cls, params*))
        SetTimer(() => Dbg.Beat(), 1000)
    }

    static Log(text) {
        try FileAppend(Format("{:7} ", A_TickCount - Dbg.Started) text "`n", Dbg.LogFile, "UTF-8")
    }

    ; 把类的静态方法换成 hook(原来的方法, 类, 参数...)
    static Hook(cls, name, hook) {
        original := cls.GetOwnPropDesc(name).Call
        cls.DefineProp(name, {Call: (target, params*) => hook(original, target, params*)})
    }

    static OnDraw(original, cls, lParam) {
        stage := NumGet(lParam, A_PtrSize * 3, "UInt")
        row := NumGet(lParam, A_PtrSize = 8 ? 56 : 36, "UPtr")
        if (stage = 0x1)
            Dbg.Prepaints++
        else if (stage = 0x10001)
            Dbg.ItemDraws++
        if (Dbg.Logged < 60 && (stage = 0x1 || stage = 0x10001)) {
            Dbg.Logged++
            Dbg.Log("draw stage " Format("0x{:X}", stage) " row " row " results " SearchWindow.Results.Length " offset " SearchWindow.Offset
                . " list rows " SearchWindow.List.GetCount() (Dbg.Loading != "" ? " (while loading " Dbg.Loading ")" : ""))
        }
        return original(cls, lParam)
    }

    static OnPaintRow(original, cls, params*) {
        try {
            return original(cls, params*)
        } catch as e {
            Dbg.Log("paint row error: " e.Message " (" e.What ", line " e.Line ")")
            throw e
        }
    }

    static OnLoad(original, cls, spec) {
        Dbg.Log("icon load start: " spec)
        Dbg.Loading := spec
        start := A_TickCount
        try {
            result := original(cls, spec)
        } finally {
            Dbg.Loading := ""
            Dbg.Log("icon load end (" (A_TickCount - start) " ms): " spec)
        }
        return result
    }

    static Beat() {
        Dbg.Log("beat: prepaint " Dbg.Prepaints " item " Dbg.ItemDraws " results " SearchWindow.Results.Length
            . " visible " SearchWindow.IsVisible() " icon queue " IconCache._queue.Count (Dbg.Loading != "" ? " loading " Dbg.Loading : ""))
    }
}
