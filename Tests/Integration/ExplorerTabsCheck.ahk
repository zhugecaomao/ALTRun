;===============================================================================
; ExplorerTabsCheck.ahk - opening folders in a new tab of File Explorer on a real Windows desktop (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Run by .github/workflows/explorer-tabs.yml (windows-latest: Windows Server 2025, whose Explorer has tabs):
;   - with no Explorer window open, ExplorerTabs.Open returns false (ALTRun opens a new window instead)
;   - with one open, a plain folder, one with spaces / # / % and a Chinese one each open in a new tab of it, with the file selected,
;     also when the window isn't in front
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
    folders := [root "\Plain", root "\A b #1 100% x", root "\中文 文件夹"]
    for folder in folders {
        DirCreate(folder)
        FileAppend("x", folder "\note.txt", "UTF-8")
    }
    for hwnd in WinGetList("ahk_class CabinetWClass")
        try WinClose("ahk_id " hwnd)
    Sleep(500)

    check("no window: not opened", !ExplorerTabs.Open(folders[1]), ExplorerTabs.LastError)

    Run('explorer.exe "' folders[1] '"')
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
        check("file selected " folder, Selected(hwnd, folder) = "note.txt", Selected(hwnd, folder))
    }

    before := ExplorerTabs._Tabs(hwnd).Length
    check("unreachable folder: not opened", !ExplorerTabs.Open(root "\Missing\Nowhere"), ExplorerTabs.LastError)
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

; the selected file's name in the tab showing folder
Selected(hwnd, folder) {
    for each in ExplorerTabs._Tabs(hwnd) {
        try {
            doc := each.Item.Document
            if !ExplorerTabs.SameFolder(doc.Folder.Self.Path, folder)
                continue
            items := doc.SelectedItems()
            return (items.Count = 1) ? items.Item(0).Name : items.Count " selected"
        }
    }
    return "no tab"
}
