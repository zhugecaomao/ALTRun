;===============================================================================
; HttpGet.ahk - 在后台下载网页内容, 下载期间程序照常响应 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; AHK 的 Download() 下载完之前整个程序都停着: 网络慢、代理要等超时时, 热键、搜索窗口都没反应。
; 这里用 Msxml2.XMLHTTP.6.0 异步请求 (和 Download() 一样走 WinINet, 用系统的代理设置),
; 用计时器看是否完成 (只在有请求时运行), 完成后再调用回调。
;
; 用法:
;   HttpGet.Start(url, done, fail?, timeoutMs?)
;     done(text)      HTTP 2xx 时调用, text 是回复的内容; done 出错 (例如内容不对) 时改为调用 fail
;     fail(message)   网络错误、HTTP 错误、超时 (默认 30 秒) 时调用; 没给 fail 时写进日志
;   HttpGet.Pending   还没完成的请求数
;===============================================================================

class HttpGet {
    static PollMs := 100, TimeoutMs := 30000
    static NewRequest := () => ComObject("Msxml2.XMLHTTP.6.0")              ; 测试时换成假的请求
    static _active := [], _timer := ""

    static Pending => HttpGet._active.Length

    static Start(url, done, fail := "", timeoutMs := "") {
        if (timeoutMs = "")
            timeoutMs := HttpGet.TimeoutMs
        try {
            xhr := HttpGet.NewRequest.Call()
            xhr.open("GET", url, true)
            xhr.setRequestHeader("If-Modified-Since", "Sat, 01 Jan 2000 00:00:00 GMT")   ; 不用 WinINet 缓存里的旧内容 (和 Download() 一样每次都问服务器)
            xhr.send()
        } catch as e {
            return HttpGet._Later(fail, "", e.Message)
        }
        HttpGet._active.Push({Xhr: xhr, Done: done, Fail: fail, Deadline: A_TickCount + timeoutMs, TimeoutMs: timeoutMs})
        if (HttpGet._timer = "")
            HttpGet._timer := () => HttpGet._Poll()
        SetTimer(HttpGet._timer, HttpGet.PollMs)
    }

    static _Poll() {
        i := HttpGet._active.Length
        while (i >= 1) {
            request := HttpGet._active[i]
            state := 0
            try state := request.Xhr.readyState
            if (state = 4) {
                HttpGet._active.RemoveAt(i)
                HttpGet._Finish(request)
            } else if (A_TickCount >= request.Deadline) {
                HttpGet._active.RemoveAt(i)
                try request.Xhr.abort()
                HttpGet._Later(request.Fail, "", "Timed out after " Round(request.TimeoutMs / 1000) " s")
            }
            i--
        }
        if !HttpGet._active.Length
            SetTimer(HttpGet._timer, 0)
    }

    static _Finish(request) {
        try {
            status := request.Xhr.status
            if (status >= 12000)                                            ; WinINet 的错误号 (连不上、找不到主机...)
                throw Error("Network error " status)
            if (status < 200 || status >= 300)
                throw Error("HTTP " status " " request.Xhr.statusText)
            text := request.Xhr.responseText
        } catch as e {
            return HttpGet._Later(request.Fail, "", e.Message)
        }
        HttpGet._Later(request.Fail, request.Done, text)
    }

    ; 回调放到单独的线程里: 回调弹出对话框时, 不耽误其它请求的检查
    static _Later(fail, done, value) {
        SetTimer(() => HttpGet._Call(fail, done, value), -1)
    }

    static _Call(fail, done, value) {
        if IsObject(done) {
            try {
                done.Call(value)
                return
            } catch as e {
                value := e.Message
            }
        }
        if IsObject(fail)
            fail.Call(value)
        else
            Logger.Error("HttpGet: " value)
    }
}
