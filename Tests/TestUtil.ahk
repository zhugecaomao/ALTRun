;===============================================================================
; TestUtil.ahk - helpers shared by the test suites, and test doubles
;===============================================================================
class TestUtil {
    static WaitFor(fn, ms := 2000) {
        start := A_TickCount
        while !fn() {
            if (A_TickCount - start > ms)
                return false
            Sleep(10)
        }
        return true
    }
}

class FakeHttpRequest {                                                    ; a fake request for the HttpGet tests (the parts of Msxml2.XMLHTTP that are used)
    readyState := 1, status := 0, statusText := "", responseText := ""
    Method := "", Url := "", Async := "", Header := "", Aborted := false
    open(method, url, async) => (this.Method := method, this.Url := url, this.Async := async)
    setRequestHeader(name, value) => (this.Header := name)
    send() => ""
    abort() => (this.Aborted := true)
}

JsonReadWrite_Parses(text) {
    try {
        JSON.Parse(text)
        return true
    }
    return false
}

StrRepeat(text, count) => StrReplace(Format("{:" count "}", ""), " ", text)
