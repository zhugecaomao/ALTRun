;===============================================================================
; StartupApproved.ahk - startup apps turned on or off as in Task Manager → Startup apps (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Task Manager doesn't delete a startup app it turns off: it writes a value of the same name under
; HKCU|HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\<Run|Run32|StartupFolder>.
; ALTRun reads and writes the one of its own start-up shortcut (see App.LaunchOffInTaskManager).
;
; Usage:
;   StartupApproved.Key                     "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\" (after HKCU / HKLM)
;   StartupApproved.Data(off [, fileTime])  the value to write (hex, REG_BINARY)
;   StartupApproved.IsOff(data)             true when a value read means turned off
;===============================================================================

class StartupApproved {
    static Key := "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\"

    ; 12 bytes: the first 02 = on, 03 = off, then (when off) the time it was turned off (FILETIME)
    static Data(off, fileTime := "") {
        if !off
            return "020000000000000000000000"
        if (fileTime = "")
            DllCall("GetSystemTimeAsFileTime", "Int64*", &fileTime := 0)
        hex := "03000000"
        Loop 8
            hex .= Format("{:02X}", (fileTime >> (8 * (A_Index - 1))) & 0xFF)
        return hex
    }

    ; An odd first byte (01, 03) = turned off
    static IsOff(data) => (StrLen(data) >= 2 && IsXDigit(SubStr(data, 1, 2)) && Mod(Integer("0x" SubStr(data, 1, 2)), 2) = 1)
}
