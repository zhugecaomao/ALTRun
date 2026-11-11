;===============================================================================
; ExplorerTabsCheck.ahk - opening folders in a new tab of File Explorer on a real Windows desktop (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Run by .github/workflows/explorer-tabs.yml (windows-latest: Windows Server 2025, whose Explorer has tabs):
;   - with no Explorer window open, ExplorerTabs.Open returns false (ALTRun opens a new window instead)
;   - with one open, a plain folder, one with spaces / # / % and a Chinese one each open in a new tab of it, with the file selected,
;     also when the window isn't in front
;   - Ctrl still held from Ctrl+Enter (Reveal) when it starts: it waits for the key, so the tab still goes to the folder and
;     no other window opens (with Ctrl down, Explorer took the navigation for Ctrl+click and the tab stayed empty)
;   - a folder that can't be reached returns false and leaves no empty tab behind
; Prints PASS / FAIL lines; the exit code is the number of failures. On a system without Explorer tabs it prints SKIP and exits 0
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Off
#Warn All, StdOut
#Include %A_ScriptDir%\..\..\Lib\ExplorerTabs.ahk

Main()

Main() {
    failures := 0
    out(text) => FileAppend(text "`n", "*", "UTF-8")
    check(name, ok, detail := "") {
        out((ok ? "PASS " : "FAIL ") name (detail != "" ? " (" detail ")" : ""))
        failures += ok ? 0 : 1
    }
    out("Windows " A_OSVersion)
    if !ExplorerTabs.Supported() {
        out("SKIP: no Explorer tabs on this Windows")
        ExitApp(0)
    }
    root := "C:\ExplorerTabsCheck"
    try DirDelete(root, true)
    folders := [root "\Plain", root "\A b #1 100% x", root "\中文 文件夹"                    ; two of each kind: more tries of the selection
             , root "\Plain 2", root "\A b #2 100% y", root "\中文 文件夹 二"]
    for folder in folders {
        DirCreate(folder)
        FileAppend("x", folder "\note.txt", "UTF-8")
    }
    for hwnd in WinGetList("ahk_class CabinetWClass")
        try WinClose("ahk_id " hwnd)
    Sleep(500)

    check("no window: not opened", !ExplorerTabs.Open(folders[1]), ExplorerTabs.LastError)

    Run('explorer.exe "' root '"')                                         ; the window starts at the root: each folder below is only in its new tab
    hwnd := WinWait("ahk_class CabinetWClass", , 15)
    check("explorer window", hwnd != 0)
    if !hwnd
        ExitApp(failures)
    Sleep(1500)
    front := Gui("+AlwaysOnTop", "ExplorerTabsCheck front")                ; Explorer isn't the window in front
    front.Show("w300 h80")
    Sleep(300)

    for folder in folders {
        before := ExplorerTabs._Tabs(hwnd).Length
        front.Show()
        t0 := A_TickCount
        opened := ExplorerTabs.Open(folder, "note.txt")
        check("opened " folder, opened, (A_TickCount - t0) " ms " ExplorerTabs.LastError)
        check("one more tab " folder, ExplorerTabs._Tabs(hwnd).Length = before + 1, ExplorerTabs._Tabs(hwnd).Length " tabs")
        check("in that window " folder, HasTab(hwnd, folder))
        check("file selected " folder, Selected(hwnd, folder) = "note.txt", Selected(hwnd, folder) "; " ExplorerTabs.SelectInfo)
    }

    ctrlFolder := root "\Ctrl held"                                        ; Reveal with Ctrl+Enter, Ctrl let go 0.7 s later
    DirCreate(ctrlFolder), FileAppend("x", ctrlFolder "\note.txt", "UTF-8")
    before := ExplorerTabs._Tabs(hwnd).Length
    windows := WinGetList("ahk_class CabinetWClass").Length
    front.Show()
    Send("{Ctrl down}")
    SetTimer(() => Send("{Ctrl up}"), -700)
    try opened := ExplorerTabs.Open(ctrlFolder, "note.txt")
    finally {
        Sleep(800)
        if GetKeyState("Ctrl")
            Send("{Ctrl up}")
    }
    check("Ctrl held: opened", opened, ExplorerTabs.LastError)
    check("Ctrl held: one more tab", ExplorerTabs._Tabs(hwnd).Length = before + 1, ExplorerTabs._Tabs(hwnd).Length " tabs, " before " before")
    check("Ctrl held: in that window", HasTab(hwnd, ctrlFolder))
    check("Ctrl held: file selected", Selected(hwnd, ctrlFolder) = "note.txt", Selected(hwnd, ctrlFolder))
    check("Ctrl held: no other window", WinGetList("ahk_class CabinetWClass").Length = windows, WinGetList("ahk_class CabinetWClass").Length " windows")
    for other in WinGetList("ahk_class CabinetWClass")                     ; a window Explorer opened for the folder instead
        if (other != hwnd)
            try WinClose("ahk_id " other)

    before := ExplorerTabs._Tabs(hwnd).Length
    check("unreachable folder: not opened", !ExplorerTabs.Open(root "\Missing\Nowhere"), ExplorerTabs.LastError)
    check("unreachable folder: a normal case (not logged as Explorer failing)", ExplorerTabs.LastErrorExpected = true, ExplorerTabs.LastError)
    Sleep(500)
    check("unreachable folder: no empty tab left", ExplorerTabs._Tabs(hwnd).Length = before, ExplorerTabs._Tabs(hwnd).Length " tabs, " before " before")
    check("window still open", WinExist("ahk_id " hwnd) != 0)

    front.Destroy()
    try WinClose("ahk_id " hwnd)
    try DirDelete(root, true)
    ExitApp(failures)
}

HasTab(hwnd, folder) {
    for each in ExplorerTabs._Tabs(hwnd)
        if ExplorerTabs.SameFolder(ExplorerTabs._Path(hwnd, each.Tab), folder)
            return true
    return false
}

; the selected file's name in the tab showing folder (from its path: Name is the display name, without the extension when
; Explorer hides extensions)
Selected(hwnd, folder) {
    for each in ExplorerTabs._Tabs(hwnd) {
        try {
            doc := each.Item.Document
            if !ExplorerTabs.SameFolder(doc.Folder.Self.Path, folder)
                continue
            items := doc.SelectedItems()
            if (items.Count != 1)
                return items.Count " selected"
            SplitPath(items.Item(0).Path, &name)
            return name
        }
    }
    return "no tab"
}
