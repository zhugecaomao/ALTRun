;===============================================================================
; WindowProvider.ahk - 切换到已经打开的窗口 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 PowerToys Run 的 Window Walker、Raycast 的 Switch Windows 一样: 输入窗口标题或程序名,
; Enter 切换过去 (最小化的先还原)。
;   "w 关键词" (Keyword)   只搜窗口; 只输入 "w " 列出所有窗口 (按前后顺序)
;   直接输入               InDefaultResults = 1 时, 标题匹配得好的窗口也显示 (排在应用和命令后面, 最多 3 个)
; 操作面板 (→): 关闭窗口。
; 只列出任务栏上能看到的窗口: 可见、有标题、不是工具窗口 / 被别的窗口拥有的窗口、
; 不在别的虚拟桌面上 (DWM cloaked), 也不包括 ALTRun 自己的窗口。
; 窗口列表在一次搜索中复用, 1 秒内的连续输入不重新枚举。
;
; 设置 (ALTRun.json -> Features.Windows): Enabled / Keyword / InDefaultResults
;===============================================================================

class WindowProvider {
    static Id := "Windows"
    static CacheMs := 1000
    static _list := [], _listTime := 0

    static Init() {
    }

    static Search(query) {
        options := AppSettings.Feature("Windows")
        exclusive := false, term := query.Text
        if query.MatchKeyword([options["Keyword"]], &rest) {
            term := rest, exclusive := true
        } else if (!options["InDefaultResults"] || StrLen(query.Text) < 2) {
            return []
        }
        windows := WindowProvider.List()
        needle := StrLower(Trim(term))
        scores := Map()
        for index, window in windows {
            score := (needle = "") ? 100 - index * 0.01 : FuzzyMatcher.BestKey(needle, window.Keys)
            if (score > 0 && (exclusive || score >= 60))                   ; 默认结果里只要匹配得好的, 免得一堆浏览器标签页
                scores[index] := score
        }
        results := []
        for index in FuzzyMatcher.TopIndexes(scores, exclusive ? ProviderRegistry.MaxResults : 3) {
            window := windows[index]
            results.Push(WindowProvider.ItemFor(window, scores[index] - (exclusive ? 0 : 14), exclusive))
        }
        return results
    }

    static ItemFor(window, score, exclusive := false) {
        hwnd := window.Hwnd
        return ResultItem(window.Title, I18n.T("Win.Subtitle", window.Process), {
            Icon: (window.Path != "") ? window.Path : "res:imageres.dll,-102", Score: score, Exclusive: exclusive,
            Uid: "window:" StrLower(window.Process), Arg: window.Title,
            OnRun: (*) => WindowProvider.Activate(hwnd),
            Actions: [ResultItem(I18n.T("Action.CloseWindow"), window.Title, {Icon: "res:imageres.dll,-98", OnRun: (*) => WindowProvider.Close(hwnd)})]
        })
    }

    ; 任务栏上的窗口, 按前后顺序 (最前面的在前): [{Hwnd, Title, Process, Path, Keys}]
    static List() {
        if (A_TickCount - WindowProvider._listTime < WindowProvider.CacheMs)
            return WindowProvider._list
        list := []
        for hwnd in WinGetList() {
            if !WindowProvider.IsSwitchable(hwnd)
                continue
            try {
                title := WinGetTitle("ahk_id " hwnd), process := WinGetProcessName("ahk_id " hwnd)
            } catch {
                continue
            }
            exePath := ""
            try exePath := WinGetProcessPath("ahk_id " hwnd)
            name := RegExReplace(process, "i)\.exe$")
            pinyinText := Pinyin.Initials(title)
            list.Push({Hwnd: hwnd, Title: title, Process: process, Path: exePath
                , Keys: [FuzzyMatcher.Key(title), FuzzyMatcher.Key(name), (pinyinText != title) ? FuzzyMatcher.Key(pinyinText) : ""]})
        }
        WindowProvider._list := list, WindowProvider._listTime := A_TickCount
        return list
    }

    ; 和任务栏一样的判断: 可见、有标题、不是工具窗口、没有所有者、没有被 DWM 隐藏 (别的虚拟桌面 / 挂起的商店应用)
    static IsSwitchable(hwnd) {
        try {
            if !DllCall("IsWindowVisible", "Ptr", hwnd) || WinGetTitle("ahk_id " hwnd) = ""
                return false
            exStyle := WinGetExStyle("ahk_id " hwnd)
            if (exStyle & 0x80) && !(exStyle & 0x40000)                     ; WS_EX_TOOLWINDOW, 除非 WS_EX_APPWINDOW
                return false
            if DllCall("GetWindow", "Ptr", hwnd, "UInt", 4, "Ptr") && !(exStyle & 0x40000)   ; GW_OWNER: 对话框等
                return false
            cloaked := 0
            DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", hwnd, "UInt", 14, "UInt*", &cloaked, "UInt", 4)   ; DWMWA_CLOAKED
            if cloaked
                return false
            if (WinGetPID("ahk_id " hwnd) = DllCall("GetCurrentProcessId"))  ; ALTRun 自己的窗口
                return false
            return WinGetClass("ahk_id " hwnd) != "Progman"
        }
        return false
    }

    static Activate(hwnd) {
        if !WinExist("ahk_id " hwnd)
            return App.Notify(I18n.T("Win.Gone"))
        if (WinGetMinMax("ahk_id " hwnd) = -1)
            WinRestore("ahk_id " hwnd)
        WinActivate("ahk_id " hwnd)
    }

    static Close(hwnd) {
        if WinExist("ahk_id " hwnd)
            WinClose("ahk_id " hwnd)
        WindowProvider._listTime := 0
    }
}
