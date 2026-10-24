;===============================================================================
; HttpGet.ahk - downloads web content in the background; the program keeps responding meanwhile (AutoHotkey v2)
;-------------------------------------------------------------------------------
; AHK's Download() stops the whole program until it finishes: with a slow network or a proxy timing out, hotkeys and the search window don't respond.
; This uses asynchronous Msxml2.XMLHTTP.6.0 requests (WinINet like Download(), with the system proxy settings),
; checks for completion on a timer (running only while there are requests) and then calls the callback.
;
; Usage:
;   HttpGet.Start(url, done, fail?, timeoutMs?)
;     done(text)      called for HTTP 2xx with the response text; if done throws (e.g. bad content), fail is called instead
;     fail(message)   called for network errors, HTTP errors and timeouts (30 s by default); without fail it goes to the log
;   HttpGet.Pending   number of requests not finished yet
;===============================================================================

class HttpGet {
    static PollMs := 100, TimeoutMs := 30000
    static NewRequest := () => ComObject("Msxml2.XMLHTTP.6.0")              ; replaced by a fake request in tests
    static _active := [], _timer := ""

    static Pending => HttpGet._active.Length

    static Start(url, done, fail := "", timeoutMs := "") {
        if (timeoutMs = "")
            timeoutMs := HttpGet.TimeoutMs
        try {
            xhr := HttpGet.NewRequest.Call()
            xhr.open("GET", url, true)
            xhr.setRequestHeader("If-Modified-Since", "Sat, 01 Jan 2000 00:00:00 GMT")   ; don't use old content from the WinINet cache (ask the server every time, like Download())
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
            if (status >= 12000)                                            ; WinINet error number (can't connect, host not found...)
                throw Error("Network error " status)
            if (status < 200 || status >= 300)
                throw Error("HTTP " status " " request.Xhr.statusText)
            text := request.Xhr.responseText
        } catch as e {
            return HttpGet._Later(request.Fail, "", e.Message)
        }
        HttpGet._Later(request.Fail, request.Done, text)
    }

    ; The callback runs in its own thread: if it shows a dialog, the other requests are still checked
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
