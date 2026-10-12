;===============================================================================
; DebugALTRun.ahk - 临时诊断: 结果行空白 (截图脚本复制到 ALTRun 文件夹里运行, 不改 ALTRun 本身)
;-------------------------------------------------------------------------------
; 先挂上钩子再载入 ALTRun.ahk, 把下面这些写进 %Temp%\ALTRunDebug.log (每行前面是启动后的毫秒数):
;   - 每次 NM_CUSTOMDRAW 回调 (阶段、行号、当时的结果数、是否正在读图标)
;   - 每次读图标的开始、结束和耗时
;   - 画行时出的错
;   - 每秒一次心跳 (定时器还能不能运行) 和回调次数
;   - ALTRun 各个类的静态方法和 InputHook 的 Start / Stop 的进入 / 退出 (">" / "<", "C" = Critical; QuickSwitch._Watch 太频繁, 不记)
;===============================================================================
#Requires AutoHotkey v2.0

Dbg.Install()

#Include %A_ScriptDir%\ALTRun.ahk

class Dbg {
    static LogFile := A_Temp "\ALTRunDebug.log"
    static Started := A_TickCount
    static Loading := ""                    ; 正在读的图标 (空 = 没有在读)
    static Prepaints := 0, ItemDraws := 0, Logged := 0
    static Tracing := true, TraceLines := 0, Depth := 0, Muted := 0
    static Mute := Map("QuickSwitch._Watch", true)
    static TraceClasses := ["ActionCatalog", "App", "AppSettings", "ApplicationProvider", "AutoDate", "BookmarkProvider", "CalculatorProvider"
        , "CapsLockSwitch", "ClipboardData", "ClipboardProvider", "CurrencyRates", "CustomCommandProvider", "Everything", "FileIndex"
        , "FileSearchProvider", "HelpProvider", "Hud", "IconCache", "Knowledge", "NavIcons", "Path", "ProviderRegistry", "QuickSwitch"
        , "RecentProvider", "ScriptProvider", "SearchWindow", "SelectionActions", "SnippetExpander", "SnippetProvider", "SystemProvider"
        , "TerminalProvider", "ThemeManager", "UpdateChecker", "Usage", "WebSearchProvider", "Win", "WindowProvider"]

    static Install() {
        try FileDelete(Dbg.LogFile)
        Dbg.Log("debug wrapper started, pid " DllCall("GetCurrentProcessId"))
        for name in Dbg.TraceClasses
            Dbg.TraceClass(name)
        Dbg.Hook(SearchWindow, "_OnCustomDraw", (original, cls, params*) => Dbg.OnDraw(original, cls, params*))
        Dbg.Hook(SearchWindow, "_PaintRow", (original, cls, params*) => Dbg.OnPaintRow(original, cls, params*))
        Dbg.Hook(IconCache, "_Load", (original, cls, params*) => Dbg.OnLoad(original, cls, params*))
        for method in ["Start", "Stop"]
            Dbg.HookTrace(InputHook.Prototype, method, "InputHook." method, InputHook.Prototype.GetOwnPropDesc(method).Call)
        SetTimer(() => Dbg.Beat(), 1000)
    }

    ; 一直开着日志文件 (每行都打开 / 关闭文件太慢, 会改变时序); 读 Handle 把缓冲写进文件, 进程被结束也不丢
    static _file := ""
    static Log(text) {
        try {
            if (Dbg._file = "")
                Dbg._file := FileOpen(Dbg.LogFile, "a", "UTF-8")
            Dbg._file.Write(Format("{:7} ", A_TickCount - Dbg.Started) text "`n")
            flushed := Dbg._file.Handle                                     ; 读 Handle 时 AHK 把缓冲写进文件
        }
    }

    ; 把类的静态方法换成 hook(原来的方法, 类, 参数...)
    static Hook(cls, name, hook) {
        original := cls.GetOwnPropDesc(name).Call
        cls.DefineProp(name, {Call: (target, params*) => hook(original, target, params*)})
    }

    ; 给类的每个静态方法套一层: 开始跟踪以后记下进入和退出
    static TraceClass(name) {
        cls := %name%
        for prop in cls.OwnProps() {
            if (prop = "__Init" || prop = "__New" || prop = "Call" || prop = "Prototype")
                continue
            desc := cls.GetOwnPropDesc(prop)
            if !desc.HasProp("Call")
                continue
            Dbg.HookTrace(cls, prop, name "." prop, desc.Call)
        }
    }

    static HookTrace(cls, prop, label, original) {
        cls.DefineProp(prop, {Call: (target, params*) => Dbg.Traced(label, original, target, params*)})
    }

    static Traced(label, original, target, params*) {
        if (!Dbg.Tracing || Dbg.Muted)
            return original(target, params*)
        if Dbg.Mute.Has(label) {
            Dbg.Muted++
            try return original(target, params*)
            finally Dbg.Muted--
        }
        Dbg.Trace("> " label (A_IsCritical ? " C" : ""))
        Dbg.Depth++
        try {
            result := original(target, params*)
        } catch as e {
            Dbg.Depth--
            Dbg.Trace("! " label ": " e.Message)
            throw e
        }
        Dbg.Depth--
        Dbg.Trace("< " label)
        return result
    }

    static Trace(text) {
        if (++Dbg.TraceLines > 20000)
            return
        Dbg.Log(Format("{:" Max(1, Dbg.Depth * 2) "}", "") text)
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
