; @altrun.title          Save Screenshot
; @altrun.keyword        shot
; @altrun.mode           silent
; @altrun.author         sfcboy
; @altrun.description    Save the most recent clipboard screenshot as a PNG file; Save As opens in the current Explorer folder
; @altrun.description_zh 把剪贴板里最近的一张截图另存为 PNG 文件; "另存为"对话框默认打开当前资源管理器所在文件夹
; @altrun.network        none
;
; Why a script: ALTRun already stores every screenshot it sees in the clipboard history as
; %LOCALAPPDATA%\ALTRun\Clipboard\img-*.png. This script just takes the newest one and copies it
; to a location you pick in a Save As dialog - no changes to ALTRun's core code needed.
;
; Trigger options (pick one or both):
;   1) Keyword: type "shot" in ALTRun and press Enter.
;   2) Hotkey:  Preferences -> Hotkeys -> add one, Action = "script:Save Screenshot.ahk"
;              (e.g. Ctrl+Alt+S). Press it right after taking a screenshot.
;
; Optional settings via environment variables (user level):
;   ALTRUN_SHOT_DIR   default folder the Save As dialog opens in (e.g. D:\Pictures\Screenshots)
;   ALTRUN_SHOT_PREFIX default filename prefix (default: screenshot)
; If ALTRUN_SHOT_DIR is set, the file name is auto-generated as <prefix>_yyyyMMdd_HHmmss.png
; and the dialog opens there; otherwise it opens in Pictures with the same suggested name.
#Requires AutoHotkey v2.0
#NoTrayIcon

Chinese := RegExMatch(A_Language, "i)^(0804|0404|0C04|1004|1404)$") > 0
; Output is best-effort: if ALTRun still has this script cached with the default "window" mode
; (a plain Reload of the script process does NOT refresh ALTRun's metadata cache), stdout has no
; handle and FileAppend throws error 6. Never let a notification break a save that already worked.
; Restart ALTRun once to pick up @altrun.mode silent and enable the notification.
Say(msg) {
    try FileAppend(msg "`n", "*", "UTF-8-RAW")
}
Fail(msg) {
    Say(msg)
    ExitApp(1)
}

; The folder that was in front when the user triggered ALTRun (ALTRUN_HWND) - i.e. the Explorer
; window whose folder the user is looking at. Falls back to the frontmost Explorer window.
ActiveFolder() {
    hwndStr := EnvGet("ALTRUN_HWND")
    hwnd := (hwndStr != "" && IsInteger(hwndStr)) ? Integer(hwndStr) : 0
    shell := ComObject("Shell.Application")
    ; Map each Explorer window handle to its current folder path
    paths := Map()
    for window in shell.Windows() {
        try {
            p := window.Document.Folder.Self.Path
            if (p != "")
                paths[window.HWND] := p
        }
    }
    ; Only VISIBLE (not minimized) Explorer windows count. A minimized folder window is not the
    ; "current folder" - when all folders are minimized, the user is effectively on the desktop.
    ; 1) Exact window that was in front when ALTRun was triggered, if still visible
    if (hwnd && paths.Has(hwnd)) {
        mm := 0
        try mm := WinGetMinMax("ahk_id " hwnd)
        if (mm != -1)
            return paths[hwnd]
    }
    ; 2) Frontmost VISIBLE Explorer window by Z-order (WinGetList is front-to-back)
    for h in WinGetList() {
        cls := ""
        try cls := WinGetClass("ahk_id " h)
        if (cls = "CabinetWClass" || cls = "ExploreWClass") {
            mm := 0
            try mm := WinGetMinMax("ahk_id " h)
            if (mm = -1)                                ; skip minimized windows
                continue
            if (paths.Has(h))
                return paths[h]
        }
    }
    return ""                                           ; no visible folder -> caller falls back to Desktop
}

; --- Locate ALTRun's screenshot folder (where the running app stores clipboard images) ---
dir := EnvGet("LOCALAPPDATA")
clipFolder := (dir != "") ? dir "\ALTRun\Clipboard" : ""
if (clipFolder = "" || !DirExist(clipFolder))
    Fail(Chinese ? "找不到 ALTRun 的剪贴板截图目录" : "ALTRun clipboard image folder not found")

; --- Find the newest img-*.png ---
newest := "", newestTime := 0
loop Files clipFolder "\img-*.png" {
    t := FileGetTime(A_LoopFileFullPath, "M")
    if (t > newestTime) {
        newestTime := t
        newest := A_LoopFileFullPath
    }
}
if (newest = "")
    Fail(Chinese ? "剪贴板里还没有截图" : "No screenshot in the clipboard history yet")

; --- Decide starting folder and default file name ---
startDir := EnvGet("ALTRUN_SHOT_DIR")
prefix := EnvGet("ALTRUN_SHOT_PREFIX")
if (prefix = "")
    prefix := "screenshot"
if (startDir = "")
    startDir := ActiveFolder()          ; folder the user was in when ALTRun was invoked
if (startDir = "" || !DirExist(startDir)) {
    ; No visible folder window -> the user is on the desktop, so default there.
    if DirExist(A_Desktop)
        startDir := A_Desktop
    else
        startDir := A_MyDocuments
}
defaultName := prefix "_" FormatTime(newestTime, "yyyyMMdd_HHmmss") ".png"

; --- Temp diagnostic: record what folder logic resolved (remove once confirmed) ---
try {
    diag := "ALTRUN_HWND=[" EnvGet("ALTRUN_HWND") "]`nstartDir=[" startDir "]`n"
    for window in ComObject("Shell.Application").Windows() {
        mm := 0
        try mm := WinGetMinMax("ahk_id " window.HWND)
        try diag .= "  win " window.HWND " [minmax=" mm "] = " window.Document.Folder.Self.Path "`n"
    }
    FileAppend(diag, A_Temp "\altrun_shot_diag.txt", "UTF-8")
}

; --- Ask where to save ("S" = Save As, 24 = prompt to create / overwrite) ---
title := Chinese ? "保存截图" : "Save Screenshot"
filter := Chinese ? "PNG 图片 (*.png)" : "PNG Images (*.png)"
dest := FileSelect("S24", startDir "\" defaultName, title, filter)
if (dest = "")                                      ; user cancelled
    ExitApp(0)

; Add .png when the user typed a name without an extension
if !RegExMatch(dest, "i)\.png$")
    dest .= ".png"

; --- Copy the already-stored PNG to the chosen path (Windows copy, with conflict prompt) ---
try FileCopy(newest, dest, true)
catch as e {
    Fail((Chinese ? "保存失败: " : "Failed to save: ") e.Message)
}

; --- Notify via stdout. Keep it ASCII: ALTRun's silent mode can mis-decode non-ASCII stdout ---
; (the file itself is already on disk, so a short confirmation is enough)
SplitPath(dest, &name)
Say("Saved: " name)
