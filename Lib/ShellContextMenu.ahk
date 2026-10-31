;===============================================================================
; ShellContextMenu.ahk - the Windows (Explorer) context menu of files and folders (AutoHotkey v2)
;-------------------------------------------------------------------------------
; The same menu Explorer shows on right-click, with the items other programs add (7-Zip, Git, antivirus...):
; the parent folder's IShellFolder gives the items' IContextMenu, which fills a popup menu. Submenus that are filled
; when opened (Send to, Open with, most third-party submenus) need the menu messages passed on to IContextMenu2/3
; while the menu shows, so the owner window's WM_INITMENUPOPUP / WM_DRAWITEM / WM_MEASUREITEM / WM_MENUCHAR are
; handed to it (OnMessage). Several items must be in the same folder (as in Explorer, one menu for one selection).
;
; Showing and running are separate steps, so the caller can close its own window in between:
;   menu := ShellContextMenu(["C:\Reports\a.docx"])     throws when the menu can't be made (missing file...)
;   id := menu.Track(ownerHwnd, x, y, extended)         0 = closed without choosing; extended = Shift's extra items
;   if id
;       menu.Invoke(id)                                 runs the chosen item (Open, Send to, Properties...)
;   ShellContextMenu.CanShow(paths)                     -> whether paths can share one menu (they exist, same folder)
; The COM objects are released when the object goes away.
;===============================================================================

class ShellContextMenu {
    static IID_IShellFolder  := "{000214E6-0000-0000-C000-000000000046}"
    static IID_IContextMenu  := "{000214E4-0000-0000-C000-000000000046}"
    static IID_IContextMenu2 := "{000214F4-0000-0000-C000-000000000046}"
    static IID_IContextMenu3 := "{BCFCE0A0-EC17-11D0-8D10-00A0C90F2719}"
    static FirstId := 1, LastId := 0x7FFF                                    ; command ids given to the menu items
    static MenuMessages := [0x117, 0x2B, 0x2C, 0x120]                       ; WM_INITMENUPOPUP, WM_DRAWITEM, WM_MEASUREITEM, WM_MENUCHAR
    static _active := ""                                                    ; the menu being shown (gets the menu messages)
    Paths := [], Menu := 0, Owner := 0, Handler := 0, HandlerLevel := 0, _pidls := [], _folder := 0, _contextMenu := 0

    ; Whether the paths can share one context menu: at least one, all exist, all in the same folder (drive roots each on their own)
    static CanShow(paths) {
        if !(paths is Array) || !paths.Length
            return false
        parent := ""
        for item in paths {
            if !FileExist(item)
                return false
            folder := ShellContextMenu._Parent(item)
            if (A_Index = 1)
                parent := folder
            else if (folder = "" || folder != parent)
                return false
        }
        return true
    }

    __New(paths) {
        if !ShellContextMenu.CanShow(paths)
            throw ValueError("No context menu for these paths", -1)
        this.Paths := paths
        for item in paths {
            pidl := 0
            if (DllCall("shell32\SHParseDisplayName", "WStr", item, "Ptr", 0, "Ptr*", &pidl, "UInt", 0, "Ptr", 0) != 0 || !pidl)
                throw OSError("SHParseDisplayName failed: " item, -1)
            this._pidls.Push(pidl)
        }
        folder := 0, child := 0                                             ; child points into the first pidl (not freed on its own)
        iid := ShellContextMenu._Guid(ShellContextMenu.IID_IShellFolder)
        if (DllCall("shell32\SHBindToParent", "Ptr", this._pidls[1], "Ptr", iid, "Ptr*", &folder, "Ptr*", &child) != 0 || !folder)
            throw OSError("SHBindToParent failed", -1)
        this._folder := folder
        children := Buffer(A_PtrSize * this._pidls.Length)
        for pidl in this._pidls
            NumPut("Ptr", DllCall("shell32\ILFindLastID", "Ptr", pidl, "Ptr"), children, (A_Index - 1) * A_PtrSize)
        contextMenu := 0
        ComCall(10, folder, "Ptr", 0, "UInt", this._pidls.Length, "Ptr", children    ; IShellFolder::GetUIObjectOf
            , "Ptr", ShellContextMenu._Guid(ShellContextMenu.IID_IContextMenu), "Ptr", 0, "Ptr*", &contextMenu)
        this._contextMenu := contextMenu
        for pair in [[3, ShellContextMenu.IID_IContextMenu3], [2, ShellContextMenu.IID_IContextMenu2]] {
            handler := 0                                                     ; the submenus' messages go to the newest interface it has
            if (ComCall(0, contextMenu, "Ptr", ShellContextMenu._Guid(pair[2]), "Ptr*", &handler, "Int") = 0 && handler) {   ; QueryInterface
                this.Handler := handler, this.HandlerLevel := pair[1]
                break
            }
        }
    }

    ; Shows the menu at x, y (screen) and returns the chosen command id, or 0. extended: the extra items Explorer shows with Shift
    Track(ownerHwnd, x, y, extended := false) {
        this.Fill(extended)
        this.Owner := ownerHwnd
        ShellContextMenu._Listen(this)
        try {
            DllCall("SetForegroundWindow", "Ptr", ownerHwnd)                ; otherwise the menu doesn't close when clicking elsewhere
            id := DllCall("TrackPopupMenuEx", "Ptr", this.Menu, "UInt", 0x0100 | 0x0002, "Int", x, "Int", y, "Ptr", ownerHwnd, "Ptr", 0, "UInt")   ; TPM_RETURNCMD | TPM_RIGHTBUTTON
            DllCall("PostMessage", "Ptr", ownerHwnd, "UInt", 0, "Ptr", 0, "Ptr", 0)   ; WM_NULL, as the TrackPopupMenu documentation asks
        } finally
            ShellContextMenu._Listen("")
        return id
    }

    ; Runs the command chosen in Track (verb = its offset from FirstId). ownerHwnd: the window for any dialog it shows (0 = none).
    ; Returns the HRESULT (negative = failed; 0x800704C7 = canceled by the user, e.g. in a confirmation)
    Invoke(id, ownerHwnd := 0) {
        p := A_PtrSize, info := Buffer(24 + 10 * p, 0)                      ; CMINVOKECOMMANDINFOEX
        NumPut("UInt", info.Size, info, 0)
        mask := 0x4000                                                      ; CMIC_MASK_UNICODE
        mask |= GetKeyState("Shift") ? 0x10000000 : 0                       ; CMIC_MASK_SHIFT_DOWN
        mask |= GetKeyState("Ctrl") ? 0x40000000 : 0                        ; CMIC_MASK_CONTROL_DOWN
        NumPut("UInt", mask, info, 4)
        NumPut("Ptr", ownerHwnd, info, 8)
        verb := id - ShellContextMenu.FirstId                               ; MAKEINTRESOURCE(offset), in both the ANSI and the Unicode field
        NumPut("Ptr", verb, info, 8 + p)
        directory := ShellContextMenu._Parent(this.Paths[1])
        NumPut("Int", 1, info, 8 + 4 * p)                                   ; SW_SHOWNORMAL
        NumPut("Ptr", verb, info, 16 + 6 * p)
        if (directory != "")
            NumPut("Ptr", StrPtr(directory), info, 16 + 8 * p)
        return ComCall(4, this._contextMenu, "Ptr", info, "Int")            ; InvokeCommand
    }

    ; The text of every item of the filled menu, with submenus indented (for tests and the log): Track fills the menu, this reads it
    static MenuText(hMenu, indent := "") {
        text := ""
        Loop DllCall("GetMenuItemCount", "Ptr", hMenu, "Int") {
            index := A_Index - 1
            length := DllCall("GetMenuStringW", "Ptr", hMenu, "UInt", index, "Ptr", 0, "Int", 0, "UInt", 0x400, "Int")   ; MF_BYPOSITION
            label := ""
            if (length > 0) {
                buf := Buffer((length + 1) * 2, 0)
                DllCall("GetMenuStringW", "Ptr", hMenu, "UInt", index, "Ptr", buf, "Int", length + 1, "UInt", 0x400, "Int")
                label := StrGet(buf)
            }
            text .= indent (label = "" ? "-" : label) "`n"
            if (sub := DllCall("GetSubMenu", "Ptr", hMenu, "Int", index, "Ptr"))
                text .= ShellContextMenu.MenuText(sub, indent "  ")
        }
        return text
    }

    ; Fills the popup menu (QueryContextMenu) without showing it; Track does this first. Returns the menu handle (destroyed with the object).
    ; extended: CMF_EXTENDEDVERBS; no CMF_CANRENAME, since renaming needs Explorer's own view
    Fill(extended := false) {
        if this.Menu
            DllCall("DestroyMenu", "Ptr", this.Menu)
        this.Menu := DllCall("CreatePopupMenu", "Ptr")
        ComCall(3, this._contextMenu, "Ptr", this.Menu, "UInt", 0, "UInt", ShellContextMenu.FirstId, "UInt", ShellContextMenu.LastId, "UInt", extended ? 0x100 : 0, "Int")
        return this.Menu
    }

    __Delete() {
        if this.Menu
            DllCall("DestroyMenu", "Ptr", this.Menu)
        for pointer in [this.Handler, this._contextMenu, this._folder]
            if pointer
                ObjRelease(pointer)
        for pidl in this._pidls
            DllCall("ole32\CoTaskMemFree", "Ptr", pidl)
    }

    ;---------------------------------------------------------------------------
    ; The menu messages while the menu shows: owner-drawn items (icons) and submenus filled when opened
    ;---------------------------------------------------------------------------
    static _Listen(shown) {
        static handler := (wParam, lParam, msg, hwnd) => ShellContextMenu._OnMenuMessage(wParam, lParam, msg, hwnd)
        ShellContextMenu._active := shown
        for msg in ShellContextMenu.MenuMessages
            OnMessage(msg, handler, IsObject(shown) ? 1 : 0)
    }

    static _OnMenuMessage(wParam, lParam, msg, hwnd) {
        shown := ShellContextMenu._active
        if (!IsObject(shown) || hwnd != shown.Owner || !shown.Handler)
            return
        if ((msg = 0x2B || msg = 0x2C) && NumGet(lParam, 0, "UInt") != 1)   ; WM_DRAWITEM / WM_MEASUREITEM of a control, not of the menu (ODT_MENU)
            return
        if (shown.HandlerLevel = 3) {
            result := 0
            if (ComCall(7, shown.Handler, "UInt", msg, "Ptr", wParam, "Ptr", lParam, "Ptr*", &result, "Int") = 0)   ; IContextMenu3::HandleMenuMsg2
                return result
            return
        }
        ComCall(6, shown.Handler, "UInt", msg, "Ptr", wParam, "Ptr", lParam, "Int")   ; IContextMenu2::HandleMenuMsg
        return (msg = 0x117) ? 0 : 1
    }

    ; The parent folder of a path ("" for a drive root: it has no folder to share with another item)
    static _Parent(fullPath) {
        fullPath := RTrim(fullPath, "\")
        if RegExMatch(fullPath, "^[A-Za-z]:$")
            return ""
        SplitPath(fullPath, , &dir)
        return dir
    }

    static _Guid(text) {
        guid := Buffer(16)
        DllCall("ole32\CLSIDFromString", "WStr", text, "Ptr", guid)
        return guid
    }
}
