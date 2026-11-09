;===============================================================================
; HttpGet.ahk - downloads web content in the background; the program keeps responding meanwhile (AutoHotkey v2)
;-------------------------------------------------------------------------------
; AHK's Download() stops the whole program until it finishes: with a slow network or a proxy timing out, hotkeys and the search window don't respond.
; This uses asynchronous Msxml2.XMLHTTP.6.0 requests (WinINet like Download(), with the system proxy settings),
; checks for completion on a timer (running only while there are requests) and then calls the callback.
;
; Usage:
;   request := HttpGet.Start(url, done, fail?, timeoutMs?, file?)
;     done(text)      called for HTTP 2xx with the response text; if done throws (e.g. bad content), fail is called instead
;                     with file: the response is saved there as it came (binary) and done gets the file name
;     fail(message)   called for network errors, HTTP errors and timeouts (30 s by default); without fail it goes to the log
;   HttpGet.Cancel(request)   stops a request that hasn't finished: no callback is called
;   HttpGet.Pending   number of requests not finished yet
;===============================================================================

class HttpGet {
    static PollMs := 100, TimeoutMs := 30000
    static NewRequest := () => ComObject("Msxml2.XMLHTTP.6.0")              ; replaced by a fake request in tests
    static _active := [], _timer := ""

    static Pending => HttpGet._active.Length

    static Start(url, done, fail := "", timeoutMs := "", file := "") {
        if (timeoutMs = "")
            timeoutMs := HttpGet.TimeoutMs
        try {
            xhr := HttpGet.NewRequest.Call()
            xhr.open("GET", url, true)
            xhr.setRequestHeader("If-Modified-Since", "Sat, 01 Jan 2000 00:00:00 GMT")   ; don't use old content from the WinINet cache (ask the server every time, like Download())
            xhr.send()
        } catch as e {
            HttpGet._Later(fail, "", e.Message)
            return ""
        }
        request := {Xhr: xhr, Done: done, Fail: fail, Deadline: A_TickCount + timeoutMs, TimeoutMs: timeoutMs, File: file}
        HttpGet._active.Push(request)
        if (HttpGet._timer = "")
            HttpGet._timer := () => HttpGet._Poll()
        SetTimer(HttpGet._timer, HttpGet.PollMs)
        return request
    }

    ; Stops request (from Start) if it is still running; its callbacks aren't called. Returns whether it was running
    static Cancel(request) {
        for i, active in HttpGet._active
            if (active == request) {
                HttpGet._active.RemoveAt(i)
                try request.Xhr.abort()
                if !HttpGet._active.Length
                    SetTimer(HttpGet._timer, 0)
                return true
            }
        return false
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
            text := (request.File = "") ? request.Xhr.responseText : HttpGet.SaveBody(request.Xhr.responseBody, request.File)
        } catch as e {
            return HttpGet._Later(request.Fail, "", e.Message)
        }
        HttpGet._Later(request.Fail, request.Done, text)
    }

    ; responseBody (a SAFEARRAY of bytes) -> file, as it came; returns the file name
    static SaveBody(body, file) {
        size := body.MaxIndex() - body.MinIndex() + 1
        handle := FileOpen(file, "w")
        try {
            if (size > 0) {
                DllCall("oleaut32\SafeArrayAccessData", "Ptr", ComObjValue(body), "Ptr*", &data := 0, "HRESULT")
                try handle.RawWrite(data, size)
                finally DllCall("oleaut32\SafeArrayUnaccessData", "Ptr", ComObjValue(body))
            }
        } finally
            handle.Close()
        return file
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
