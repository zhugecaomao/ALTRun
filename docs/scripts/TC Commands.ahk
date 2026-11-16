; @altrun.title   Total Commander Commands
; @altrun.keyword tc
; @altrun.mode    filter
; @altrun.filter  altrun
;
; A command palette for Total Commander in ALTRun: type "tc " and part of a command's name (in English or in the language
; Total Commander uses, pinyin initials too), then Enter runs it in Total Commander.
; Lists Total Commander's internal commands (cm_..., from TOTALCMD.INC; named by its language file when it has one) and your
; own commands (em_..., from usercmd.ini). Works on the Total Commander window you came from, or else the one used last.
; Save it in ALTRun's Scripts folder (Preferences -> Scripts -> Open Scripts Folder). See the Scripts page of ALTRun's wiki.
;
; How it works: ALTRun runs this script once when you type "tc " (ALTRUN_ACTION=filter) and filters the list it prints.
; Enter runs it again with the command's name (ALTRUN_ACTION=run, A_Args[1]): cm_ commands are sent as their number
; (WM_USER+51), em_ commands by name (WM_COPYDATA, "EM"). Total Commander running as administrator doesn't accept them.
#Requires AutoHotkey v2.0
#NoTrayIcon

tc := FindTotalCmd()
if (EnvGet("ALTRUN_ACTION") = "run")
    RunCommand(tc, A_Args.Length ? A_Args[1] : "")
else
    ListCommands(tc)
ExitApp

; The Total Commander window to work on: the one ALTRun was opened from, else the one used last -> {Hwnd, Dir, Exe}, or "" when none
FindTotalCmd() {
    hwnd := EnvGet("ALTRUN_HWND")
    if (hwnd != "" && RegExMatch(EnvGet("ALTRUN_PROCESS"), "i)^TOTALCMD(64)?\.EXE$") && WinExist("ahk_id " hwnd)) {
        if (WinGetClass("ahk_id " hwnd) != "TTOTAL_CMD")                    ; one of its dialogs: its main window
            hwnd := WinExist("ahk_class TTOTAL_CMD ahk_pid " EnvGet("ALTRUN_PID")) || hwnd
    } else
        hwnd := WinExist("ahk_class TTOTAL_CMD")
    if !hwnd
        return ""
    exe := ""
    try exe := WinGetProcessPath("ahk_id " hwnd)
    dir := EnvGet("COMMANDER_PATH")                                         ; set by Total Commander for what it starts
    if (dir = "")
        SplitPath(exe, , &dir)
    return {Hwnd: hwnd, Dir: RTrim(dir, "\"), Exe: exe}
}

ListCommands(tc) {
    if !IsObject(tc) {
        FileAppend('{"items": [{"title": "Total Commander is not running", "subtitle": "Start it, then type tc again", "valid": false}]}', "*", "UTF-8-RAW")
        return
    }
    inc := tc.Dir "\TOTALCMD.INC"
    if !FileExist(inc)
        Fail("Can't find TOTALCMD.INC in " tc.Dir)
    ini := WincmdIni(tc.Dir)
    SplitPath(ini, , &iniDir)
    names := LanguageNames(ini, tc.Dir)
    items := []
    if FileExist(iniDir "\usercmd.ini")                                     ; your own commands first
        for section in ReadIni(ReadText(iniDir "\usercmd.ini"))
            if (StrLower(SubStr(section.Name, 1, 3)) = "em_") {
                title := (section.Keys.Has("menu") && section.Keys["menu"] != "") ? section.Keys["menu"] : SubStr(section.Name, 4)
                items.Push(Item(title, "usercmd.ini" Dot() section.Name, section.Name, tc.Exe, title " " section.Name))
            }
    group := ""
    Loop Parse, ReadText(inc), "`n", "`r" {
        if RegExMatch(A_LoopField, "^\[_*(.*?)_*\]", &m) {                 ; [_____File Operations_____]=0
            key := StrLower(StrReplace(m[1], " "))
            group := names.Has(key) ? names[key] : m[1]
            continue
        }
        if !RegExMatch(A_LoopField, "i)^(cm_\w+)=(\d+);(.*)$", &m)          ; cm_SrcComments=300;Source: Show comments
            continue                                                        ; also "CD <directory>=-2;...": not a command
        name := m[1], english := Usable(m[3]) ? Trim(m[3]) : SubStr(name, 4)
        title := (names.Has(m[2]) && Usable(names[m[2]])) ? names[m[2]] : english
        items.Push(Item(title, group Dot() name, name, tc.Exe, title (title != english ? " " english : "") " " name))
    }
    out := ""
    for each in items
        out .= (A_Index > 1 ? ",`n" : "") each
    FileAppend('{"items": [' out ']}', "*", "UTF-8-RAW")
}

Dot() => " " Chr(0xB7) " "                                                   ; "Group - cm_Name" with a middle dot (the file stays ASCII)

; A description that can be a title: not "(etc, define your own if)", "etc." or "..." as some entries have
Usable(text) => (text := Trim(text)) != "" && SubStr(text, 1, 1) != "(" && !RegExMatch(text, "i)^(etc\.?|\.\.\.)$")

Item(title, subtitle, arg, icon, match) => '{"title": ' Json(title) ', "subtitle": ' Json(subtitle) ', "arg": ' Json(arg)
    . ', "type": "script", "uid": ' Json(arg) ', "match": ' Json(match) (icon != "" ? ', "icon": ' Json(icon) : "") "}"

Json(text) => '"' StrReplace(StrReplace(StrReplace(StrReplace(StrReplace(text, "\", "\\"), '"', '\"'), "`t", " "), "`r", ""), "`n", " ") '"'

RunCommand(tc, name) {
    if !IsObject(tc)
        Fail("Total Commander isn't running")
    if !RegExMatch(name, "i)^(cm|em)_\w+$")
        Fail("Not a Total Commander command: " name)
    name := StrLower(SubStr(name, 1, 3)) SubStr(name, 4)                    ; TC wants the prefix in lower case
    if !WinActive("ahk_id " tc.Hwnd) {                                      ; ALTRun was opened from somewhere else
        WinActivate("ahk_id " tc.Hwnd)
        try WinWaitActive("ahk_id " tc.Hwnd, , 1)
    }
    number := (SubStr(name, 1, 3) = "cm_") ? CommandNumber(tc.Dir "\TOTALCMD.INC", name) : ""
    if (number != "")
        return PostMessage(0x433, Integer(number), 0, , "ahk_id " tc.Hwnd)           ; WM_USER+51; posted: a command opening a dialog doesn't hold this script
    data := Buffer(StrPut(name, "CP0")), StrPut(name, data, "CP0")          ; em_ (or a cm_ missing from TOTALCMD.INC) by name
    copyData := Buffer(A_PtrSize * 3, 0)                                    ; COPYDATASTRUCT
    NumPut("UPtr", Ord("E") + 256 * Ord("M"), copyData, 0)
    NumPut("UInt", data.Size, copyData, A_PtrSize)
    NumPut("Ptr", data.Ptr, copyData, A_PtrSize * 2)
    answer := 0
    DllCall("SendMessageTimeoutW", "Ptr", tc.Hwnd, "UInt", 0x4A, "Ptr", A_ScriptHwnd, "Ptr", copyData.Ptr
        , "UInt", 0x2, "UInt", 2000, "Ptr*", &answer)                       ; SMTO_ABORTIFHUNG; TC may still be busy with it after 2 s
    if !WinExist("ahk_id " tc.Hwnd)
        Fail("Total Commander was closed")
}

CommandNumber(inc, name) {
    if FileExist(inc)
        Loop Parse, ReadText(inc), "`n", "`r"
            if RegExMatch(A_LoopField, "i)^" name "=(\d+);", &m)
                return m[1]
    return ""
}

; wincmd.ini: COMMANDER_INI, else the registry (IniFileName, as the installer sets it), else the program folder, else %APPDATA%\GHISLER.
; (Started with /i=<file>, Total Commander uses that file instead: not seen here.) usercmd.ini is in the same folder
WincmdIni(dir) {
    ini := EnvGet("COMMANDER_INI")
    for root in ["HKCU", "HKLM"]
        if (ini = "")
            try ini := RegRead(root "\Software\Ghisler\Total Commander", "IniFileName")
    if (ini != "") {
        ini := StrReplace(ini, "%COMMANDER_PATH%", dir)
        size := DllCall("ExpandEnvironmentStringsW", "Str", ini, "Ptr", 0, "UInt", 0, "UInt")
        expanded := Buffer(size * 2), DllCall("ExpandEnvironmentStringsW", "Str", ini, "Ptr", expanded, "UInt", size, "UInt")
        ini := StrGet(expanded)
        if !RegExMatch(ini, "^([A-Za-z]:\\|\\\\)")                          ; relative: to the program folder
            ini := dir "\" RegExReplace(ini, "^\.\\")
        if FileExist(ini)
            return ini
    }
    return FileExist(dir "\wincmd.ini") ? dir "\wincmd.ini" : A_AppData "\GHISLER\wincmd.ini"
}

; The names in Total Commander's language: wincmd.ini [Configuration] LanguageIni=WCMD_CHN.LNG -> Language\WCMD_CHN.INC, with
; 300="..." (a command's number) and FileOperations="..." (a group) -> Map: number or group name (lower case, no spaces) -> text
LanguageNames(ini, dir) {
    names := Map()
    if !FileExist(ini)
        return names
    language := ""
    for section in ReadIni(ReadText(ini))
        if (section.Name = "Configuration" && section.Keys.Has("LanguageIni"))
            language := StrReplace(section.Keys["LanguageIni"], "%COMMANDER_PATH%", dir)
    if (language = "")
        return names
    language := RegExReplace(language, "\.[^.\\]*$") ".INC"
    if !RegExMatch(language, "^([A-Za-z]:\\|\\\\)")
        language := FileExist(dir "\Language\" language) ? dir "\Language\" language : dir "\" language
    if !FileExist(language)
        return names
    Loop Parse, ReadText(language), "`n", "`r"
        if RegExMatch(A_LoopField, '^\s*([^=;]+?)\s*=\s*"?(.*?)"?\s*$', &m) && m[2] != ""
            names[IsDigit(m[1]) ? m[1] : StrLower(StrReplace(m[1], " "))] := m[2]
    return names
}

; [section] -> [{Name, Keys: Map(key -> value), keys in any case}], in the file's order
ReadIni(text) {
    sections := [], keys := ""
    Loop Parse, text, "`n", "`r" {
        line := Trim(A_LoopField)
        if RegExMatch(line, "^\[(.+)\]$", &m) {
            keys := Map(), keys.CaseSense := "Off"
            sections.Push({Name: m[1], Keys: keys})
        } else if (IsObject(keys) && SubStr(line, 1, 1) != ";" && (split := InStr(line, "=")) > 1)
            keys[Trim(SubStr(line, 1, split - 1))] := Trim(SubStr(line, split + 1))
    }
    return sections
}

; A text file as Total Commander writes them: UTF-16 or UTF-8 with a byte order mark, UTF-8, or else the ANSI code page
; (language files such as WCMD_CHN.INC are in their language's code page, which is the system's for that language)
ReadText(path) {
    data := FileRead(path, "RAW"), size := data.Size
    if (size >= 2 && NumGet(data, 0, "UShort") = 0xFEFF)
        return StrGet(data.Ptr + 2, (size - 2) // 2, "UTF-16")
    if (size >= 3 && NumGet(data, 0, "UShort") = 0xBBEF && NumGet(data, 2, "UChar") = 0xBF)
        return StrGet(data.Ptr + 3, size - 3, "UTF-8")
    if !size
        return ""
    utf8 := DllCall("MultiByteToWideChar", "UInt", 65001, "UInt", 8, "Ptr", data, "Int", size, "Ptr", 0, "Int", 0)   ; MB_ERR_INVALID_CHARS
    return StrGet(data, size, utf8 ? "UTF-8" : "CP0")
}

Fail(message) {                                                             ; ALTRun shows the first line of standard error
    FileAppend(message "`n", "**", "UTF-8-RAW")
    ExitApp(1)
}
