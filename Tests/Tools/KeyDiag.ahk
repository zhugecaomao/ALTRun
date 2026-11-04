;===============================================================================
; KeyDiag.ahk - records which keys reach which window, to find out why a program opens a window by itself (AutoHotkey v2)
;-------------------------------------------------------------------------------
; For key problems around ALTRun that only happen on someone's own PC (2026-11: Total Commander's Synchronize directories
; window opened a properties dialog after Alt+Space and Enter on Quit; a lone Enter key-up was taken as Alt+Enter).
; Runs on its own in the tray, so it keeps recording while ALTRun exits and restarts. Not shipped with ALTRun.
;
; Only while a program in KeyDiagSettings.Watch (or ALTRun) is in front, it records:
;   - keys going down / up: which key, whether it came from the keyboard or was sent by a program, auto-repeat.
;     Letters, digits, symbols and typed text are NOT recorded (they show as "key"), so passwords never end up in a file.
;   - which window comes to the front, menu bars / menus opening, windows those programs show, ALTRun quitting / starting
; When the window that is the problem shows (KeyDiagSettings.Trigger*; by default a properties dialog of Total Commander):
; a beep, and KeyDiag-hit-<time>.txt with everything needed: the 15 s before it and 2 s after it, the keyboard settings
; and the programs running then (names and versions only). Nothing to do by hand.
;   Ctrl+F9   in the test window (TC's Synchronize directories window, a file selected): automatic test that presses
;             ALTRun's hotkey, types QuitText, presses Enter, checks for that window and restarts ALTRun
;   Ctrl+F10  KeyDiag-report-<time>.txt now (the same, for the last 60 s)
; Tray menu: start with Windows, open the folder. Files go to %LOCALAPPDATA%\ALTRun\KeyDiag (day logs older than 14 days are
; deleted); nothing is sent anywhere. Window titles are not recorded.
;
; Usage: run it with AutoHotkey v2 (double-click) with the same rights as the program in question (as administrator if it
; runs so), keep working as usual, and send the KeyDiag-hit file once it beeps.
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

class KeyDiagSettings {
    static AltRunExe := "ALTRun.exe"
    static Watch := ["TOTALCMD64.EXE", "TOTALCMD.EXE"]          ; programs recorded while in front (ALTRun always)
    static TriggerPrograms := ["TOTALCMD64.EXE", "TOTALCMD.EXE"]   ; the window that is the problem: shown by one of these programs,
    static TriggerClass := "#32770"                             ; of this window class,
    static TriggerTitle := "(属性|Properties|內容|プロパティ)$"   ; with a title matching this (read only to tell it; never recorded)
    static TestWindow := "ahk_class TCmpForm"                   ; automatic test: the window it starts in
    static QuitText := "quit"                                   ; typed in ALTRun's search window to find its Quit command
    static HotkeyVks := [0xA4, 0x20]                            ; ALTRun's hotkey: LAlt + Space
    static Rounds := 3                                          ; automatic test: rounds per variant
    static HitBefore := 15000, HitAfter := 2000, ReportSpan := 60000   ; ms
}

KeyDiag.Start()

^F9::KeyDiag.Test()
^F10::KeyDiag.Saved(KeyDiag.Report())

KbdHookProc(nCode, wParam, lParam) {
    if (nCode = 0)
        try KeyDiag.OnKey(wParam, lParam)
    return DllCall("CallNextHookEx", "Ptr", 0, "Int", nCode, "Ptr", wParam, "Ptr", lParam, "Ptr")
}

WinEventProc(hook, event, hwnd, idObject, idChild, thread, tick) {
    try KeyDiag.OnWin(event, hwnd, idObject, idChild, tick)
}

class KeyDiag {
    static Events := []                                         ; [A_TickCount, line] of the last ReportSpan ms
    static Pending := []                                        ; lines not in the day's log yet
    static Names := Map()                                       ; pid -> process name
    static Down := Map()                                        ; vk -> true while down (to tell auto-repeat)
    static LastTick := 0
    static FrontWatched := false
    static AltRunPid := 0
    static Testing := false
    static KbdHook := 0
    static KbdFn := 0
    static SentMark := 0xD1A6                                   ; dwExtraInfo of the keys this script sends
    static KeyNames := Map(0x08, "Backspace", 0x09, "Tab", 0x10, "Shift", 0x11, "Ctrl", 0x12, "Alt", 0x13, "Pause"
        , 0x14, "CapsLock", 0x15, "IME-Kana", 0x19, "IME-Kanji", 0x1B, "Esc", 0x1C, "IME-Convert", 0x1D, "IME-NonConvert"
        , 0x20, "Space", 0x21, "PgUp", 0x22, "PgDn", 0x23, "End", 0x24, "Home", 0x25, "Left", 0x26, "Up", 0x27, "Right"
        , 0x28, "Down", 0x2C, "PrintScreen", 0x2D, "Insert", 0x2E, "Delete", 0x5B, "LWin", 0x5C, "RWin", 0x5D, "Apps"
        , 0xA0, "LShift", 0xA1, "RShift", 0xA2, "LCtrl", 0xA3, "RCtrl", 0xA4, "LAlt", 0xA5, "RAlt"
        , 0xE5, "IME-Process", 0xE7, "char(text)", 0xE8, "vkE8(mask)", 0xFF, "vkFF")

    static Start() {
        try DirCreate(KeyDiag.Dir)
        A_IconTip := "KeyDiag: recording to " KeyDiag.Dir
        tray := A_TrayMenu
        tray.Delete()
        tray.Add("Save a report now (Ctrl+F10)", (*) => KeyDiag.Saved(KeyDiag.Report()))
        tray.Add("Help / automatic test (Ctrl+F9)", (*) => KeyDiag.Help())
        tray.Add("Open the folder", (*) => Run(KeyDiag.Dir))
        tray.Add("Start with Windows", (name, *) => KeyDiag.ToggleStartup(name))
        if FileExist(KeyDiag.StartupLink)
            tray.Check("Start with Windows")
        tray.Add()
        tray.Add("Exit", (*) => ExitApp())
        KeyDiag.KbdFn := CallbackCreate(KbdHookProc, "F", 3)
        if !KeyDiag.Hook()
            return MsgBox("Could not install the keyboard hook (SetWindowsHookEx error " A_LastError ").", "KeyDiag", "Icon!")
        winFn := CallbackCreate(WinEventProc, "F", 7)
        for span in [[0x3, 0x7], [0x8002, 0x8002]]               ; foreground, menus; window shown
            DllCall("SetWinEventHook", "UInt", span[1], "UInt", span[2], "Ptr", 0, "Ptr", winFn, "UInt", 0, "UInt", 0, "UInt", 0x2, "Ptr")   ; WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS
        OnExit((*) => KeyDiag.Stop())
        try {
            if (found := KeyDiag.FindAltRun())
                KeyDiag.AltRunPid := found.Pid
        }
        KeyDiag.CleanLogs()
        KeyDiag.Pending.Push("", "==== KeyDiag started " FormatTime(, "yyyy-MM-dd HH:mm:ss"), KeyDiag.Environment()*)
        KeyDiag.Flush()
        SetTimer(() => KeyDiag.Flush(), 2000)
        SetTimer(() => KeyDiag.WatchAltRun(), 250)
        SetTimer(() => KeyDiag.Rehook(), 60000)
        TrayTip("Recording. When the window in question shows, it beeps and saves the file to send.", "KeyDiag")
    }

    static Dir := EnvGet("LOCALAPPDATA") "\ALTRun\KeyDiag"
    static StartupLink := A_Startup "\KeyDiag.lnk"

    static ToggleStartup(name) {
        if FileExist(KeyDiag.StartupLink) {
            try FileDelete(KeyDiag.StartupLink)
        } else
            try FileCreateShortcut(A_AhkPath, KeyDiag.StartupLink, A_ScriptDir, '"' A_ScriptFullPath '"', "KeyDiag")
        if FileExist(KeyDiag.StartupLink)
            A_TrayMenu.Check(name)
        else
            A_TrayMenu.Uncheck(name)
    }

    static CleanLogs() {
        loop files KeyDiag.Dir "\KeyDiag-*.log"
            if (DateDiff(A_Now, A_LoopFileTimeModified, "Days") > 14)
                try FileDelete(A_LoopFileFullPath)
    }

    static Hook() {
        KeyDiag.KbdHook := DllCall("SetWindowsHookEx", "Int", 13, "Ptr", KeyDiag.KbdFn, "Ptr", DllCall("GetModuleHandle", "Ptr", 0, "Ptr"), "UInt", 0, "Ptr")
        return KeyDiag.KbdHook
    }

    ; Windows drops a keyboard hook that answered too slowly once, without telling: set it again now and then, while nobody types
    static Rehook() {
        if (A_TimeIdle < 3000 || KeyDiag.Testing)
            return
        DllCall("UnhookWindowsHookEx", "Ptr", KeyDiag.KbdHook)
        KeyDiag.Hook()
    }

    static Stop() {
        DllCall("UnhookWindowsHookEx", "Ptr", KeyDiag.KbdHook)
        KeyDiag.Pending.Push("==== KeyDiag stopped " FormatTime(, "yyyy-MM-dd HH:mm:ss"))
        KeyDiag.Flush()
    }

    ;---------------------------------------------------------------- recording
    static OnKey(wParam, lParam) {
        vk := NumGet(lParam, 0, "UInt"), flags := NumGet(lParam, 8, "UInt"), tick := NumGet(lParam, 12, "UInt"), extra := NumGet(lParam, 16, "UPtr")
        up := wParam = 0x101 || wParam = 0x105                  ; WM_KEYUP, WM_SYSKEYUP
        repeated := !up && KeyDiag.Down.Has(vk)
        if up {
            if KeyDiag.Down.Has(vk)
                KeyDiag.Down.Delete(vk)
        } else
            KeyDiag.Down[vk] := true
        front := KeyDiag.Front()
        if !front.Watched
            return
        KeyDiag.Add(tick, Format("KEY   {1:-12} {2:-6} {3}{4}   {5}", KeyDiag.KeyName(vk, flags), up ? "up" : repeated ? "repeat" : "down"
            , KeyDiag.Origin(flags, extra), (wParam = 0x104 || wParam = 0x105) ? ", system key (Alt held)" : "", front.Text))
    }

    static OnWin(event, hwnd, idObject, idChild, tick) {
        if !hwnd
            return
        if (event = 0x8002 && (idObject != 0 || idChild != 0 || DllCall("GetAncestor", "Ptr", hwnd, "UInt", 2, "Ptr") != hwnd))
            return                                              ; EVENT_OBJECT_SHOW: top-level windows only
        pid := KeyDiag.Pid(hwnd), name := KeyDiag.Name(pid), watched := KeyDiag.Watched(name, pid)
        if (event = 0x3) {                                      ; EVENT_SYSTEM_FOREGROUND
            if (watched || KeyDiag.FrontWatched)
                KeyDiag.Add(tick, "FRONT " (watched ? name " " KeyDiag.Class(hwnd) : "another program"))
            KeyDiag.FrontWatched := watched
            return
        }
        if !watched
            return
        cls := KeyDiag.Class(hwnd)
        if (event = 0x8002) {
            hit := KeyDiag.IsTrigger(name, cls, hwnd)
            KeyDiag.Add(tick, "SHOW  " name " " cls (hit ? "   >>> THE WINDOW IN QUESTION <<<" : ""))
            if hit
                KeyDiag.Hit()
            return
        }
        static menus := Map(4, "menu bar active (it took a lone Alt or F10)", 5, "menu bar inactive", 6, "menu shown", 7, "menu closed")
        KeyDiag.Add(tick, "MENU  " menus[event] "   " name " " cls)
    }

    static WatchAltRun() {
        if (KeyDiag.AltRunPid && ProcessExist(KeyDiag.AltRunPid))
            return
        if KeyDiag.AltRunPid {
            KeyDiag.Add(A_TickCount, "ALTRUN quit (pid " KeyDiag.AltRunPid ")")
            if KeyDiag.Names.Has(KeyDiag.AltRunPid)
                KeyDiag.Names.Delete(KeyDiag.AltRunPid)
            KeyDiag.AltRunPid := 0
        }
        if (pid := ProcessExist(KeyDiagSettings.AltRunExe)) {
            KeyDiag.Add(A_TickCount, "ALTRUN running (pid " pid ")")
            KeyDiag.AltRunPid := pid
        }
    }

    static Add(tick, text) {
        delta := KeyDiag.LastTick ? tick - KeyDiag.LastTick : 0
        KeyDiag.LastTick := tick
        line := Format("{1}.{2}  {3:7}  {4}", FormatTime(, "HH:mm:ss"), A_MSec, (delta >= 0 ? "+" : "") delta, text)
        KeyDiag.Events.Push([A_TickCount, line])
        KeyDiag.Pending.Push(line)
        while (KeyDiag.Events.Length > 5000 || (KeyDiag.Events.Length && A_TickCount - KeyDiag.Events[1][1] > KeyDiagSettings.ReportSpan))
            KeyDiag.Events.RemoveAt(1)
    }

    static Hit() {
        at := A_TickCount, stamp := FormatTime(, "yyyyMMdd-HHmmss")
        SetTimer(() => KeyDiag.SaveHit(stamp, at), -KeyDiagSettings.HitAfter)
    }

    static SaveHit(stamp, at) {
        lines := ["KeyDiag: the window in question showed at " stamp, "", KeyDiag.Environment()*]
        lines.Push("", Format("---- from {1} s before to {2} s after it (>>> marks the window in question)", KeyDiagSettings.HitBefore // 1000, KeyDiagSettings.HitAfter // 1000))
        lines.Push(KeyDiag.Recent(A_TickCount - at + KeyDiagSettings.HitBefore)*)
        ; the programs running: reading their versions takes a moment, so the keyboard hook is off meanwhile (typing isn't held up)
        DllCall("UnhookWindowsHookEx", "Ptr", KeyDiag.KbdHook)
        try lines.Push("", "---- programs running", KeyDiag.Programs()*)
        KeyDiag.Hook()
        file := KeyDiag.Dir "\KeyDiag-hit-" stamp ".txt"
        KeyDiag.Write(file, lines)
        SoundBeep(1000, 150)
        TrayTip("Caught it. Send this file:`n" file, "KeyDiag")
    }

    static Recent(span) {
        Critical
        lines := []
        for e in KeyDiag.Events
            if (A_TickCount - e[1] <= span)
                lines.Push(e[2])
        return lines
    }

    static Flush() {
        Critical
        lines := KeyDiag.Pending, KeyDiag.Pending := []
        Critical "Off"
        if lines.Length
            KeyDiag.Write(KeyDiag.Dir "\KeyDiag-" FormatTime(, "yyyyMMdd") ".log", lines)
    }

    static Write(file, lines) {
        text := ""
        for line in lines
            text .= line "`r`n"
        try FileAppend(text, file, "UTF-8")
    }

    ;---------------------------------------------------------------- what the lines say
    static KeyName(vk, flags) {
        if (vk = 0x0D)
            return (flags & 0x1) ? "NumpadEnter" : "Enter"
        if (vk >= 0x70 && vk <= 0x87)
            return "F" (vk - 0x6F)
        return KeyDiag.KeyNames.Has(vk) ? KeyDiag.KeyNames[vk] : "key"          ; letters, digits, symbols: not recorded
    }

    static Origin(flags, extra) {
        if !(flags & 0x12)                                      ; LLKHF_INJECTED, LLKHF_LOWER_IL_INJECTED
            return "keyboard"
        if (extra = KeyDiag.SentMark)
            return "sent by KeyDiag"
        if ((extra & ~0xFF) = 0xFFC3D400)
            return "sent by AutoHotkey"                         ; AutoHotkey marks the keys it sends this way
        return "sent by a program" (extra ? Format(" (0x{:X})", extra) : "") ((flags & 0x2) ? ", lower integrity" : "")
    }

    static Front() {
        hwnd := DllCall("GetForegroundWindow", "Ptr"), pid := KeyDiag.Pid(hwnd), name := KeyDiag.Name(pid)
        watched := KeyDiag.Watched(name, pid)
        return {Watched: watched, Text: watched ? "[" name " " KeyDiag.Class(hwnd) "]" : "[another program]"}
    }

    static Watched(name, pid) {
        return pid && pid != DllCall("GetCurrentProcessId") && (KeyDiag.InList(KeyDiagSettings.Watch, name) || name = KeyDiagSettings.AltRunExe || pid = KeyDiag.AltRunPid)
    }

    static InList(list, name) {
        for exe in list
            if (name = exe)
                return true
        return false
    }

    static IsTrigger(name, cls, hwnd) => KeyDiag.InList(KeyDiagSettings.TriggerPrograms, name) && cls = KeyDiagSettings.TriggerClass
        && RegExMatch(KeyDiag.Title(hwnd), KeyDiagSettings.TriggerTitle)

    static Pid(hwnd) {
        pid := 0
        DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "UInt*", &pid)
        return pid
    }

    static Name(pid) {
        if !pid
            return ""
        if !KeyDiag.Names.Has(pid) {
            try
                KeyDiag.Names[pid] := ProcessGetName(pid)
            catch
                KeyDiag.Names[pid] := "?"
        }
        return KeyDiag.Names[pid]
    }

    static Class(hwnd) {
        buf := Buffer(512, 0)
        DllCall("GetClassName", "Ptr", hwnd, "Ptr", buf, "Int", 256)
        return StrGet(buf)
    }

    static Title(hwnd) {                                        ; read only to tell a properties window; never recorded
        buf := Buffer(1024, 0)
        DllCall("GetWindowText", "Ptr", hwnd, "Ptr", buf, "Int", 512)
        return StrGet(buf)
    }

    ;---------------------------------------------------------------- environment
    static Environment() {
        lines := [Format("Windows {1} {2}, KeyDiag {3}, AutoHotkey {4}", A_OSVersion, A_Is64bitOS ? "64-bit" : "32-bit"
            , A_IsAdmin ? "as administrator" : "not as administrator", A_AhkVersion)]
        for p in KeyDiag.Processes()
            if (KeyDiag.InList(KeyDiagSettings.Watch, p.Name) || p.Name = KeyDiagSettings.AltRunExe || (KeyDiag.AltRunPid && p.Pid = KeyDiag.AltRunPid))
                lines.Push(Format("{1} {2} (pid {3}), as administrator: {4}", p.Name, KeyDiag.Version(KeyDiag.ExePath(p.Pid)), p.Pid, KeyDiag.Elevated(p.Pid)))
        lines.Push("Keyboard layouts: " KeyDiag.Layouts(), KeyDiag.KeyboardSettings())
        return lines
    }

    static Processes() {                                        ; [{Pid, Name}] of every process running
        list := []
        snap := DllCall("CreateToolhelp32Snapshot", "UInt", 0x2, "UInt", 0, "Ptr")
        if (snap = -1)
            return list
        size := A_PtrSize = 8 ? 568 : 556                       ; PROCESSENTRY32W
        entry := Buffer(size, 0)
        NumPut("UInt", size, entry)
        ok := DllCall("Process32FirstW", "Ptr", snap, "Ptr", entry)
        while ok {
            list.Push({Pid: NumGet(entry, 8, "UInt"), Name: StrGet(entry.Ptr + 28 + 2 * A_PtrSize, 260)})
            ok := DllCall("Process32NextW", "Ptr", snap, "Ptr", entry)
        }
        DllCall("CloseHandle", "Ptr", snap)
        return list
    }

    static ExePath(pid) {
        if !(h := DllCall("OpenProcess", "UInt", 0x1000, "Int", 0, "UInt", pid, "Ptr"))   ; PROCESS_QUERY_LIMITED_INFORMATION
            return ""
        buf := Buffer(2048, 0), len := 1024
        ok := DllCall("QueryFullProcessImageNameW", "Ptr", h, "UInt", 0, "Ptr", buf, "UInt*", &len)
        DllCall("CloseHandle", "Ptr", h)
        return ok ? StrGet(buf, len) : ""
    }

    static Version(exe) {
        if (exe = "")
            return ""
        try return FileGetVersion(exe)
        return ""
    }

    static Elevated(pid) {
        if !(h := DllCall("OpenProcess", "UInt", 0x1000, "Int", 0, "UInt", pid, "Ptr"))
            return "unknown"
        result := "unknown (no access, maybe yes)", token := 0
        if DllCall("advapi32\OpenProcessToken", "Ptr", h, "UInt", 0x8, "Ptr*", &token) {   ; TOKEN_QUERY
            info := Buffer(4, 0), len := 0
            if DllCall("advapi32\GetTokenInformation", "Ptr", token, "Int", 20, "Ptr", info, "UInt", 4, "UInt*", &len)   ; TokenElevation
                result := NumGet(info, "UInt") ? "yes" : "no"
            DllCall("CloseHandle", "Ptr", token)
        }
        DllCall("CloseHandle", "Ptr", h)
        return result
    }

    static Layouts() {
        count := DllCall("GetKeyboardLayoutList", "Int", 0, "Ptr", 0)
        buf := Buffer(count * A_PtrSize, 0)
        DllCall("GetKeyboardLayoutList", "Int", count, "Ptr", buf)
        text := ""
        loop count
            text .= (A_Index > 1 ? ", " : "") Format("{:08X}", NumGet(buf, (A_Index - 1) * A_PtrSize, "UPtr") & 0xFFFFFFFF)
        return text
    }

    static KeyboardSettings() {
        delay := 0, speed := 0
        DllCall("SystemParametersInfo", "UInt", 0x16, "UInt", 0, "UInt*", &delay, "UInt", 0)   ; SPI_GETKEYBOARDDELAY
        DllCall("SystemParametersInfo", "UInt", 0x0A, "UInt", 0, "UInt*", &speed, "UInt", 0)   ; SPI_GETKEYBOARDSPEED
        fk := KeyDiag.Spi(0x32, 24), sk := KeyDiag.Spi(0x3A, 8), tk := KeyDiag.Spi(0x34, 8)             ; FILTERKEYS, STICKYKEYS, TOGGLEKEYS
        return Format("Keyboard: repeat delay {1} ms, repeat rate {2}/31; Filter Keys {3} (wait {4}, delay {5}, repeat {6}, bounce {7} ms); Sticky Keys {8}; Toggle Keys {9}"
            , (delay + 1) * 250, speed, (NumGet(fk, 4, "UInt") & 1) ? "on" : "off", NumGet(fk, 8, "UInt"), NumGet(fk, 12, "UInt")
            , NumGet(fk, 16, "UInt"), NumGet(fk, 20, "UInt"), (NumGet(sk, 4, "UInt") & 1) ? "on" : "off", (NumGet(tk, 4, "UInt") & 1) ? "on" : "off")
    }

    static Spi(action, size) {
        buf := Buffer(size, 0)
        NumPut("UInt", size, buf)
        DllCall("SystemParametersInfo", "UInt", action, "UInt", size, "Ptr", buf, "UInt", 0)
        return buf
    }

    static Programs() {                                         ; names and versions; Windows' own only counted
        static hints := ["autohotkey", "ahk", "powertoys", "keyboard", "kbd", "hotkey", "macro", "logi", "razer", "corsair", "icue"
            , "steelseries", "synapse", "armoury", "synaptics", "elan", "x-mouse", "strokes", "gesture", "quicker", "utools", "wox"
            , "listary", "everything", "ditto", "snipaste", "pixpin", "sogou", "baidu", "qqpinyin", "wetype", "weasel", "rime"
            , "barrier", "synergy", "inputdirector", "teamviewer", "anydesk", "todesk", "sunlogin", "vmware", "virtualbox"
            , "translat", "youdao", "eudic", "deepl", "clip", "capslock", "sharpkeys", "kanata", "kmonad", "launcher", "keepass"
            , "1password", "bitwarden"]
        seen := Map(), seen.CaseSense := "Off", text := "", windows := 0
        for p in KeyDiag.Processes() {
            if seen.Has(p.Name)
                continue
            seen[p.Name] := true
            exe := KeyDiag.ExePath(p.Pid)
            if (exe != "" && InStr(exe, A_WinDir "\") = 1) {
                windows++
                continue
            }
            hint := ""
            for word in hints
                if InStr(p.Name, word) {
                    hint := "   <- may watch or send keys"
                    break
                }
            text .= p.Name "  " (exe = "" ? "(no access)" : KeyDiag.Version(exe)) hint "`n"
        }
        lines := StrSplit(Trim(Sort(text), "`n"), "`n")
        lines.Push("(and " windows " of Windows' own)")
        return lines
    }

    static Report(results := "") {
        stamp := FormatTime(, "yyyyMMdd-HHmmss")
        lines := ["KeyDiag report " stamp, "", KeyDiag.Environment()*]
        if IsObject(results)
            lines.Push("", "---- automatic test", results*)
        lines.Push("", "---- programs running", KeyDiag.Programs()*)
        lines.Push("", "---- the last " KeyDiagSettings.ReportSpan // 1000 " s", KeyDiag.Recent(KeyDiagSettings.ReportSpan)*)
        file := KeyDiag.Dir "\KeyDiag-report-" stamp ".txt"
        KeyDiag.Write(file, lines)
        return file
    }

    static Saved(file) => TrayTip("Report saved:`n" file, "KeyDiag")

    static Help() {
        MsgBox("
        (
        KeyDiag records in the background (the tray icon). It records only while ALTRun or a program in its Watch list is in front;
        letters, digits, symbols and typed text are not recorded, nor window titles.

        Normally: keep working as usual. When the window in question shows, it beeps and saves KeyDiag-hit-<time>.txt
        (with the programs running and the keyboard settings) in the folder of the tray menu. Send that one file.
        Run KeyDiag with the same rights as that program (as administrator if it runs so). "Start with Windows" is in the tray menu.

        Automatic test (Ctrl+F9), for Total Commander by default:
        1. ALTRun is running
        2. Open TC's Synchronize directories window, compare, select a file
        3. Press Ctrl+F9, confirm, and don't touch the keyboard or mouse for about a minute
        It presses ALTRun's hotkey, types the quit text and presses Enter, checks whether the window in question opened,
        then restarts ALTRun; twice 3 rounds (Enter pressed briefly / held with auto-repeat), then saves a report.

        Ctrl+F10: save a report now (KeyDiag-report-<time>.txt)
        )", "KeyDiag", "Iconi")
    }

    ;---------------------------------------------------------------- automatic test
    static Test() {
        if KeyDiag.Testing
            return
        if !(sync := WinActive(KeyDiagSettings.TestWindow))
            return KeyDiag.Help()
        if !(altrun := KeyDiag.FindAltRun())
            return MsgBox("ALTRun isn't running.", "KeyDiag", "Icon!")
        KeyWait("Ctrl"), KeyWait("F9")
        if (MsgBox("Start the automatic test? It takes about a minute; don't touch the keyboard or mouse meanwhile.`n(ALTRun quits and restarts " KeyDiagSettings.Rounds * 2 " times.)", "KeyDiag", "OKCancel Iconi") != "OK")
            return
        KeyDiag.Testing := true
        testPid := WinGetPID(sync), results := []
        variants := [["A", "Enter pressed 60 ms", 60, false], ["B", "Enter held 700 ms with auto-repeat, like a real keyboard", 700, true]]
        for v in variants {
            loop KeyDiagSettings.Rounds {
                KeyDiag.Add(A_TickCount, Format("TEST  {1} round {2}: {3}", v[1], A_Index, v[2]))
                result := KeyDiag.Round(sync, testPid, altrun, v[3], v[4])
                KeyDiag.Add(A_TickCount, "TEST  result: " result)
                results.Push(Format("{1} round {2} ({3}): {4}", v[1], A_Index, v[2], result))
                if !altrun.Pid
                    break 2
            }
        }
        KeyDiag.Testing := false
        file := KeyDiag.Report(results)
        text := ""
        for line in results
            text .= line "`n"
        MsgBox("Automatic test done:`n`n" text "`nReport: " file, "KeyDiag", "Iconi")
    }

    static Round(sync, testPid, altrun, holdMs, autoRepeat) {
        try WinActivate(sync)
        if !WinWaitActive(sync, , 3)
            return "the test window didn't come to the front"
        Sleep(500)
        modVk := KeyDiagSettings.HotkeyVks[1], keyVk := KeyDiagSettings.HotkeyVks[2]
        KeyDiag.Send([[modVk, false]]), Sleep(40)
        KeyDiag.Send([[keyVk, false]]), Sleep(30), KeyDiag.Send([[keyVk, true]])
        shown := WinWaitActive("ahk_pid " altrun.Pid, , 3)
        Sleep(60)
        KeyDiag.Send([[modVk, true]])                              ; Alt goes up while ALTRun is in front, as when typing by hand
        if !shown
            return "ALTRun's window didn't show"
        Sleep(200)
        KeyDiag.SendText(KeyDiagSettings.QuitText)
        Sleep(600)
        KeyDiag.Send([[0x0D, false]])
        if autoRepeat {
            Sleep(500)
            loop Max(0, (holdMs - 500) // 33)
                KeyDiag.Send([[0x0D, false]]), Sleep(33)
        } else
            Sleep(holdMs)
        KeyDiag.Send([[0x0D, true]])
        quit := !ProcessWaitClose(altrun.Pid, 5)
        Sleep(1500)
        found := KeyDiag.FindTrigger(testPid)
        if found
            try WinClose(found)
        result := found ? "THE WINDOW IN QUESTION OPENED" : quit ? "it didn't open" : "ALTRun didn't quit (change QuitText at the top?)"
        if quit
            altrun.Pid := KeyDiag.Restart(altrun)
        return result
    }

    static FindTrigger(pid) {
        for hwnd in WinGetList("ahk_class " KeyDiagSettings.TriggerClass " ahk_pid " pid)
            if KeyDiag.IsTrigger(KeyDiag.Name(pid), KeyDiagSettings.TriggerClass, hwnd)
                return hwnd
        return 0
    }

    static FindAltRun() {                                       ; {Pid, Command, Dir} of the ALTRun running (exe or script), or ""
        own := DllCall("GetCurrentProcessId")
        query := "SELECT ProcessId, CommandLine FROM Win32_Process WHERE Name = '" KeyDiagSettings.AltRunExe "' OR CommandLine LIKE '%ALTRun.ahk%'"
        for p in ComObjGet("winmgmts:").ExecQuery(query) {
            if (p.ProcessId = own)
                continue
            exe := KeyDiag.ExePath(p.ProcessId), command := ""
            try command := String(p.CommandLine)
            SplitPath(exe, , &dir)
            return {Pid: p.ProcessId, Command: command != "" ? command : '"' exe '"', Dir: dir}
        }
        return ""
    }

    static Restart(altrun) {
        try Run(altrun.Command, altrun.Dir)
        catch
            return 0
        loop 40 {
            Sleep(250)
            if (found := KeyDiag.FindAltRun()) {
                KeyDiag.AltRunPid := found.Pid
                Sleep(2500)                                     ; until its hotkey works
                return found.Pid
            }
        }
        return 0
    }

    ; SendInput of [[vk, up]...] in one batch; a third item true sends vk as a character
    static Send(events) {
        size := 8 + 3 * A_PtrSize + 8                           ; INPUT (40 bytes on 64-bit)
        buf := Buffer(size * events.Length, 0)
        for index, ev in events {
            offset := (index - 1) * size, vk := ev[1], upFlag := ev[2] ? 0x2 : 0
            NumPut("UInt", 1, buf, offset)                      ; INPUT_KEYBOARD
            if (ev.Length > 2 && ev[3])
                NumPut("UShort", 0, "UShort", vk, "UInt", 0x4 | upFlag, buf, offset + A_PtrSize)   ; KEYEVENTF_UNICODE
            else
                NumPut("UShort", vk, "UShort", DllCall("MapVirtualKey", "UInt", vk, "UInt", 0, "UInt")
                    , "UInt", upFlag | ((vk = 0xA5 || vk = 0xA3) ? 0x1 : 0), buf, offset + A_PtrSize)
            NumPut("UPtr", KeyDiag.SentMark, buf, offset + A_PtrSize + 8 + A_PtrSize)   ; dwExtraInfo
        }
        return DllCall("SendInput", "UInt", events.Length, "Ptr", buf, "Int", size, "UInt")
    }

    static SendText(text) {
        loop parse text {
            KeyDiag.Send([[Ord(A_LoopField), false, true], [Ord(A_LoopField), true, true]])
            Sleep(40)
        }
    }
}
