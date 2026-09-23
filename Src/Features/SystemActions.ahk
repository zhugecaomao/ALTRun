;===============================================================================
; SystemActions.ahk - 系统级/杂项内置命令 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 收的是那些跟"当前搜索状态/主窗口"没什么关系, 单纯"做一件系统级或对外部服务
; 的事情"的内置 Func 命令: 显示器/音量/电源、进程与服务列表、搜索引擎跳转、
; IP 地址、URL 编码这些杂项工具。
;
; 用法 (ALTRun.ahk 里, 全部通过同名的裸全局函数外壳调用):
;   TurnMonitorOff() { SystemActions.TurnMonitorOff() }   等等
;
; 这些外壳必须是裸的全局函数, 不能直接把 SystemActions.Xxx 登记成内置命令:
; "Func | Xxx | ..." 命令和 FuncList 自定义热键都是靠 CommandRunner.Execute() 里的
; %cmdPath%() 按名字动态调用, 只认裸的全局函数名, 不认 Class.Method - 和
; NewClip()/PTTools() 等已有的外壳是同一个原因, 详见那边的注释。
;===============================================================================

Class SystemActions {

    ; --- 显示器 / 音量 / 电源 --------------------------------------------------

    static TurnMonitorOff() {
        SendMessage(0x112, 0xF170, 2, , "Program Manager")   ; WM_SYSCOMMAND / SC_MONITORPOWER
    }

    static MuteVolume() {
        SoundSetMute(true)
    }

    static IncreaseVolume() {
        SoundSetVolume("+5")
    }

    static DecreaseVolume() {
        SoundSetVolume("-5")
    }

    static EmptyRecycle() {
        if (MsgBox("Do you really want to empty the Recycle Bin?", , "YesNo") = "Yes")
            FileRecycleEmpty
    }

    static Logoff() {
        if (MsgBox(g_LNG[850], g_TITLE, "YesNo") = "Yes")
            Shutdown(0)
    }

    static ShutdownMachine() {
        if (MsgBox(g_LNG[851], g_TITLE, "YesNo") = "Yes")
            Shutdown(1)
    }

    static RestartMachine() {
        if (MsgBox(g_LNG[852], g_TITLE, "YesNo") = "Yes")
            Shutdown(2)
    }

    static HibernateMachine() {
        if (MsgBox(g_LNG[853], g_TITLE, "YesNo") = "Yes")
            DllCall("PowrProf\SetSuspendState", "Int", 1, "Int", 0, "Int", 0)
    }

    ; --- 进程 / 服务 -------------------------------------------------------------

    static ListProcess() {
        SystemActions._ShowCmdOutputInNotepad("tasklist", "ALTRun.Processes.txt")
    }

    static ListService() {
        SystemActions._ShowCmdOutputInNotepad("net start", "ALTRun.Services.txt")
    }

    ; 运行一条控制台命令, 把结果写到命名临时文件后用记事本打开 - 一次性的长列表
    ; 不值得专门做一个可滚动/可过滤的窗口。
    static _ShowCmdOutputInNotepad(command, tempName) {
        tempFile := A_Temp "\" tempName
        try FileDelete(tempFile)
        FileAppend(SystemActions._GetCmdOutput(command), tempFile, "UTF-8")
        Run("Notepad.exe `"" tempFile "`"")
    }

    static _GetCmdOutput(command) {
        tempFile := A_Temp "\ALTRun.stdout"
        RunWait(A_ComSpec " /C " command " > " tempFile, A_Temp, "Hide")
        result := FileRead(tempFile)
        try FileDelete(tempFile)
        return RTrim(result, "`r`n")                                   ; 去掉结果末尾的换行
    }

    ; --- 搜索引擎跳转 -------------------------------------------------------------

    static Google() {
        Run("https://www.google.com/search?q=" SystemActions._Keyword() "&newwindow=1")
    }

    static Bing() {
        Run("https://cn.bing.com/search?q=" SystemActions._Keyword())
    }

    static Baidu() {
        Run("https://www.baidu.com/s?wd=" SystemActions._Keyword())
    }

    static Taobao() {
        Run("https://s.taobao.com/search?q=" SystemActions._Keyword())
    }

    static JD() {
        Run("http://search.jd.com/Search?keyword=" SystemActions._Keyword() "&enc=utf-8")
    }

    static Everything() {
        try {
            Run(g_CONFIG["Everything"] . ' -s `"' g_RUNTIME["Arg"] '`"')
        } catch as e {
            MsgBox("Everything software not found.`n`nPlease check ALTRun setting and Everything program file.`n`nError message=" . e.Message)
        }
    }

    ; 搜索类动作共用: 优先用输入框里的参数, 没有就用剪贴板。
    static _Keyword() {
        return g_RUNTIME["Arg"] = "" ? A_Clipboard : g_RUNTIME["Arg"]
    }

    ; --- 其它杂项 -------------------------------------------------------------

    static AhkRun() {
        try {
            Run(g_RUNTIME["Arg"])
        } catch as e {
            g_LOG.Debug("AhkRun: Error occur=" . e.Message)
        }
    }

    static ShowIP() {
        ips := []
        for _, ip in [A_IPAddress1, A_IPAddress2, A_IPAddress3, A_IPAddress4]
            if (ip != "0.0.0.0")
                ips.Push(ip)
        if (!ips.Length)
            return MsgBox(g_LNG[843], g_TITLE, 48)

        text := ""
        for _, ip in ips
            text .= (text = "" ? "" : "`n") . ip
        A_Clipboard := ips[1]
        MsgBox(text, g_LNG[844], 64)
    }

    ; Percent-encodes the arg/clipboard and writes the result back to the clipboard
    ; (byte-by-byte over its UTF-8 encoding, so non-ASCII text encodes correctly too).
    static UrlEncode() {
        text := SystemActions._Keyword()
        if (text = "")
            return

        out := ""
        for ch in StrSplit(text) {
            if RegExMatch(ch, "^[0-9A-Za-z\-_.~]$") {
                out .= ch
                continue
            }
            buf := Buffer(8, 0)
            len := StrPut(ch, buf, "UTF-8") - 1
            Loop len
                out .= Format("%{:02X}", NumGet(buf, A_Index - 1, "UChar"))
        }
        A_Clipboard := out
        ToolTip(g_LNG[845])
        SetTimer(() => ToolTip(""), -1500)
    }
}
