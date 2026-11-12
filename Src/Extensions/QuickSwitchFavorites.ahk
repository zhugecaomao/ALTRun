;===============================================================================
; QuickSwitchFavorites.ahk - Quick Switch's folders, favorite folders and ☆ menu (AutoHotkey v2)
;-------------------------------------------------------------------------------
; The folders listed in the panel (TC, Opus and Explorer windows, recent folders), the favorite folders
; (Extensions.QuickSwitch.Favorites in ALTRun.json), the ☆ menu with one level of subfolders and the rows' right-click menu.
; Part of QuickSwitch (Src\Extensions\QuickSwitch.ahk); reads and refreshes QuickSwitchPanel's state.
;
; Usage:
;   QuickSwitchFavorites.MenuFolders()            the folders in the panel [{Path, Group, Tag}]
;   QuickSwitchFavorites.AddFavorite(folder) / RemoveFavorite(folder) / IsFavorite(folder)   the favorite folders (saved in ALTRun.json)
;===============================================================================

class QuickSwitchFavorites {
    ;---------------------------------------------------------------------------
    ; The folders in the panel
    ;---------------------------------------------------------------------------
    ; [{Path, Group: "open" / "recent", Tag}]: the active and other panel of each TC window, each Explorer window,
    ; recent folders (RecentFolders of them), without duplicates. The favorite folders are in the ☆ menu (and in searches, SearchBase)
    ; opus = false: without Directory Opus (the panel adds those a moment later, see _AddDOpusFolders)
    static MenuFolders(opus := true) {
        list := [], seen := Map()
        add(folder, group, tag) {
            if (folder = "" || seen.Has(StrLower(RTrim(folder, "\"))))
                return
            seen[StrLower(RTrim(folder, "\"))] := true
            list.Push({Path: folder, Group: group, Tag: tag})
        }
        for hwnd in WinGetList("ahk_class TTOTAL_CMD") {
            add(FileManagers.TotalCmdFolder(hwnd), "open", I18n.T("QuickSwitch.TagTC"))
            add(FileManagers.TotalCmdFolder(hwnd, true), "open", I18n.T("QuickSwitch.TagTCOther"))
        }
        if opus                                                             ; the source tab of each lister, then the other side
            for state in [1, 2]
                for tab in FileManagers.DOpusTabs()
                    if (tab.State = state)
                        add(tab.Path, "open", I18n.T(state = 1 ? "QuickSwitch.TagOpus" : "QuickSwitch.TagOpusOther"))
        try {
            for window in ComObject("Shell.Application").Windows
                try add(window.Document.Folder.Self.Path, "open", I18n.T("QuickSwitch.TagExplorer"))
        }
        count := QuickSwitch.Options.Has("RecentFolders") ? QuickSwitch.Options["RecentFolders"] : 10
        for folder in FileManagers.RecentFolders(count)
            add(folder, "recent", I18n.T("QuickSwitch.TagRecent"))
        return list
    }

    ;---------------------------------------------------------------------------
    ; Favorite folders (Options["Favorites"], saved in ALTRun.json): listed first in the panel, in the order they were added.
    ; Environment variables in the setting (e.g. %USERPROFILE%\Documents, typed in Preferences) are expanded when listed;
    ; a folder that doesn't exist (an unplugged drive) is still listed: checking network paths could hang
    ;---------------------------------------------------------------------------
    static Favorites() {
        folders := []
        for folder in QuickSwitchFavorites._FavoriteSetting()
            folders.Push(QuickSwitchFavorites._FavoritePath(folder))
        return folders
    }

    static _FavoriteSetting() => (QuickSwitch.Options.Has("Favorites") && QuickSwitch.Options["Favorites"] is Array) ? QuickSwitch.Options["Favorites"] : []

    ; One entry of the setting as a folder: environment variables expanded, no trailing backslash ("D:\" for a drive root); "" if empty
    static _FavoritePath(folder) {
        folder := Trim(Path.ExpandEnv(folder))
        return RegExMatch(folder, "^[A-Za-z]:\\?$") ? SubStr(folder, 1, 2) "\" : RTrim(folder, "\")
    }

    ; Position of folder in the setting (however it is written there), 0 if it isn't a favorite
    static _FavoriteIndex(folder) {
        key := StrLower(RTrim(folder, "\"))
        if (key = "")
            return 0
        for index, entry in QuickSwitchFavorites._FavoriteSetting()
            if (StrLower(RTrim(QuickSwitchFavorites._FavoritePath(entry), "\")) = key)
                return index
        return 0
    }

    static IsFavorite(folder) => QuickSwitchFavorites._FavoriteIndex(folder) > 0

    ; Adds folder at the end and saves the settings; false if it is empty or already a favorite
    static AddFavorite(folder) {
        folder := QuickSwitchFavorites._FavoritePath(folder)
        if (folder = "" || QuickSwitchFavorites.IsFavorite(folder))
            return false
        if !(QuickSwitch.Options.Has("Favorites") && QuickSwitch.Options["Favorites"] is Array)
            QuickSwitch.Options["Favorites"] := []
        QuickSwitch.Options["Favorites"].Push(folder)
        AppSettings.Save()
        return true
    }

    ; Removes folder and saves the settings; false if it isn't a favorite
    static RemoveFavorite(folder) {
        if !(index := QuickSwitchFavorites._FavoriteIndex(folder))
            return false
        QuickSwitch.Options["Favorites"].RemoveAt(index)
        AppSettings.Save()
        return true
    }

    ; Ctrl+D in the panel (the caret is in the panel: after clicking it, or after the MenuHotkey): the folder the dialog shows becomes a favorite.
    ; In the dialog itself Ctrl+D is the date stamp's hotkey (AutoDate), so it is only taken here
    static AddDialogFolder() {
        folder := QuickSwitchPanel._panelFor ? FileDialogs.DialogFolder(QuickSwitchPanel._panelFor) : ""
        if (folder = "")
            return App.Notify(I18n.T("QuickSwitch.NoDialogFolder"), 2500)
        QuickSwitchFavorites._ToggleFavorite(folder, true)
    }

    ; Right-click on the panel's list: add / remove the row's folder; add the dialog's current folder
    static _ShowRowMenu(row) {
        if !QuickSwitchPanel._panelFor
            return
        rowMenu := Menu(), items := 0
        if (row >= 1 && row <= QuickSwitchPanel._panelFolders.Length) {
            entry := QuickSwitchPanel._panelFolders[row]
            if !(entry.HasOwnProp("IsFile") && entry.IsFile) && !entry.HasOwnProp("Hint") {
                folder := entry.Path, favorite := QuickSwitchFavorites.IsFavorite(folder)
                rowMenu.Add(I18n.T(favorite ? "QuickSwitch.RemoveFavorite" : "QuickSwitch.AddFavorite"), (*) => QuickSwitchFavorites._ToggleFavorite(folder, !favorite))
                items++
            }
        }
        dialogFolder := FileDialogs.DialogFolder(QuickSwitchPanel._panelFor)
        if (dialogFolder != "" && !QuickSwitchFavorites.IsFavorite(dialogFolder)) {
            rowMenu.Add(QuickSwitchFavorites.MenuLabel(I18n.T("QuickSwitch.AddDialogFolder", QuickSwitchPanel.FolderName(dialogFolder))) "`tCtrl+D", (*) => QuickSwitchFavorites._ToggleFavorite(dialogFolder, true))
            items++
        }
        if items {
            rowMenu.Show()
            QuickSwitchFavorites._DialogAfterMenu()
        }
    }

    ; After a panel menu: Menu.Show and ShowMenuBelow make ALTRun's hidden window the foreground window for the menu (or it wouldn't
    ; close on a click elsewhere). A chosen item jumps back to the dialog by itself, but a menu closed with Esc or a click elsewhere
    ; left the hidden window in front, and typing went nowhere (checked on Windows CI: Tests\Integration\QuickSwitchCheck.ahk).
    ; So the dialog comes back to the front when nothing else has taken it
    static _DialogAfterMenu() {
        if (DllCall("GetForegroundWindow", "Ptr") = A_ScriptHwnd)            ; not when another program was clicked to close the menu
            QuickSwitchPanel._BackToDialog()
    }

    ; Adds or removes a favorite, then lists the panel's folders again (without Opus first, as when the panel opens), keeping the search.
    ; The notice is ALTRun's own (Hud), shown at once: Windows notifications may be turned off
    static _ToggleFavorite(folder, add) {
        changed := add ? QuickSwitchFavorites.AddFavorite(folder) : QuickSwitchFavorites.RemoveFavorite(folder)
        App.Notify(I18n.T(!add ? "QuickSwitch.FavoriteRemoved" : changed ? "QuickSwitch.FavoriteAdded" : "QuickSwitch.FavoriteExists", QuickSwitchPanel.FolderName(folder)))
        if !(changed && QuickSwitchPanel._panelFor && IsObject(QuickSwitchPanel._panelList))
            return
        QuickSwitchPanel._panelBase := QuickSwitchFavorites.MenuFolders(false)
        QuickSwitchPanel._SearchPanel()
        if WinExist("ahk_class dopus.lister")
            SetTimer(QuickSwitchPanel._AddDOpusFolders.Bind(QuickSwitch, QuickSwitchPanel._panelFor), -1)
    }

    ; The favorite folders as rows of a search ([{Path, Group: "favorite", Tag}])
    static FavoriteEntries() {
        entries := []
        for folder in QuickSwitchFavorites.Favorites()
            if (folder != "")
                entries.Push({Path: folder, Group: "favorite", Tag: I18n.T("QuickSwitch.TagFavorite")})
        return entries
    }

    ;---------------------------------------------------------------------------
    ; The ☆ menu: the favorite folders, each with one level of its subfolders (read when the menu opens). Only folders on this PC's own
    ; drives are read: a network path (\\server\share, a mapped drive) or a CD would wait for the network or the drive, and a menu can't
    ; show "loading". A favorite with subfolders is a submenu whose first item is the favorite itself (in bold): clicking an item always
    ; jumps, never only opens a submenu (in Listary's favorites menu it was unclear when a click opens the folder and when it expands it).
    ; The menu is read on the UI thread each time it opens, so at most MaxReadSubfolders subfolders are read ("470+ more folders"):
    ; a favorite with thousands of them, or a disk that has to spin up, would hold the click. They are sorted by name as Explorer
    ; sorts them (StrCmpLogicalW: "2" before "10"), since FAT / exFAT drives (USB sticks) list folders unsorted; then the first
    ; MaxSubfolders are shown. With more than MaxReadSubfolders on such a drive, those are the first of the ones read
    ;---------------------------------------------------------------------------
    static MaxSubfolders := 30, MaxReadSubfolders := 500

    ; [{Path, Name, Subfolders: [paths], More: how many more than MaxSubfolders, MoreAtLeast: true when reading stopped at maxRead}];
    ; hidden and system folders are left out
    static FavoritesMenuItems(maxSubfolders := QuickSwitchFavorites.MaxSubfolders, maxRead := QuickSwitchFavorites.MaxReadSubfolders) {
        items := []
        for folder in QuickSwitchFavorites.Favorites() {
            if (folder = "")
                continue
            item := {Path: folder, Name: QuickSwitchPanel.FolderName(folder), Subfolders: [], More: 0, MoreAtLeast: false}
            if QuickSwitchFavorites.ReadsSubfolders(folder) {
                found := "", count := 0
                try Loop Files RTrim(folder, "\") "\*", "D" {
                    if InStr(A_LoopFileAttrib, "H") || InStr(A_LoopFileAttrib, "S")
                        continue
                    if (++count > maxRead) {
                        count := maxRead, item.MoreAtLeast := true
                        break
                    }
                    found .= (found = "" ? "" : "`n") A_LoopFileFullPath
                }
                if (found != "")
                    for subfolder in StrSplit(Sort(found, "CLogical"), "`n")   ; same parent: sorting the paths sorts the names
                        if (A_Index <= maxSubfolders)
                            item.Subfolders.Push(subfolder)
                item.More := Max(0, count - maxSubfolders)
            }
            items.Push(item)
        }
        return items
    }

    ; Whether the menu reads the subfolders of folder: on a fixed, removable or RAM drive of this PC (GetDriveType doesn't touch the drive)
    static ReadsSubfolders(folder) {
        if !RegExMatch(folder, "^[A-Za-z]:\\", &root)
            return false
        driveType := DllCall("GetDriveType", "Str", root[0], "UInt")
        return (driveType = 2 || driveType = 3 || driveType = 6)        ; DRIVE_REMOVABLE, DRIVE_FIXED, DRIVE_RAMDISK
    }

    ; A menu item's text: & shows as & (not as an underlined access key)
    static MenuLabel(text) => StrReplace(text, "&", "&&")

    ; The ☆ button: shows the menu below it
    static ShowFavoritesMenu() {
        if !QuickSwitchPanel._panelFor
            return
        favMenu := QuickSwitchFavorites.BuildFavoritesMenu(FileDialogs.DialogFolder(QuickSwitchPanel._panelFor))
        QuickSwitchPanel._panelStar.GetPos(&x, &y, &w, &h)
        rect := Buffer(16), NumPut("Int", x, "Int", y, "Int", x + w, "Int", y + h, rect)
        DllCall("MapWindowPoints", "Ptr", QuickSwitchPanel._panel.Hwnd, "Ptr", 0, "Ptr", rect, "UInt", 2)   ; client -> screen
        QuickSwitchFavorites.ShowMenuBelow(favMenu, rect)
    }

    ; The ☆ menu: the favorites (their path on the right, as two may have the same name), each with subfolders as a submenu whose
    ; first, default item is the favorite itself; "Add <dialog folder>" when it isn't a favorite yet; "Edit favorites"
    static BuildFavoritesMenu(dialogFolder) {
        favMenu := Menu()
        items := QuickSwitchFavorites.FavoritesMenuItems()
        for item in items {
            label := QuickSwitchFavorites.MenuLabel(item.Name) "`t" QuickSwitchFavorites.MenuLabel(Path.Compact(item.Path))
            jump := QuickSwitchPanel.JumpTo.Bind(QuickSwitch, item.Path)
            if item.Subfolders.Length {
                subMenu := Menu()
                subMenu.Add(label, jump), subMenu.Default := label
                subMenu.SetIcon(label, "shell32.dll", 44)
                subMenu.Add()
                for folder in item.Subfolders {
                    name := QuickSwitchFavorites.MenuLabel(QuickSwitchPanel.FolderName(folder))
                    subMenu.Add(name, QuickSwitchPanel.JumpTo.Bind(QuickSwitch, folder))
                    subMenu.SetIcon(name, "shell32.dll", 4)
                }
                if item.More {
                    more := I18n.T("QuickSwitch.MoreFolders", item.More (item.MoreAtLeast ? "+" : ""))
                    subMenu.Add(more, (*) => 0), subMenu.Disable(more)
                }
                favMenu.Add(label, subMenu)
            } else
                favMenu.Add(label, jump)
            favMenu.SetIcon(label, "shell32.dll", 44)
        }
        if !items.Length {
            none := I18n.T("QuickSwitch.NoFavorites")
            favMenu.Add(none, (*) => 0), favMenu.Disable(none)
        }
        favMenu.Add()
        if (dialogFolder != "" && !QuickSwitchFavorites.IsFavorite(dialogFolder))
            favMenu.Add(QuickSwitchFavorites.MenuLabel(I18n.T("QuickSwitch.AddDialogFolder", QuickSwitchPanel.FolderName(dialogFolder))) "`tCtrl+D"
                , (*) => QuickSwitchFavorites._ToggleFavorite(dialogFolder, true))
        favMenu.Add(I18n.T("QuickSwitch.EditFavorites"), (*) => App.OpenPreferences("Prefs.Page.QSPanel"))
        return favMenu
    }

    ; Shows m as a drop-down of a button at the right (button = screen RECT): its right edge at the button's, below it, or above it
    ; when there is no room below (TPMPARAMS keeps the button uncovered). Menu.Show can only align the left edge, so the menu would
    ; run past the panel. The items run as with Menu.Show: TrackPopupMenuEx with the script's window sends it the item's WM_COMMAND
    static ShowMenuBelow(m, button) {
        params := Buffer(20, 0), NumPut("UInt", 20, params)                 ; TPMPARAMS: cbSize, rcExclude
        DllCall("RtlMoveMemory", "Ptr", params.Ptr + 4, "Ptr", button, "Ptr", 16)
        DllCall("SetForegroundWindow", "Ptr", A_ScriptHwnd)                 ; else the menu doesn't close when clicking elsewhere
        DllCall("TrackPopupMenuEx", "Ptr", m.Handle, "UInt", 0x0008 | 0x0040   ; TPM_RIGHTALIGN | TPM_VERTICAL
            , "Int", NumGet(button, 8, "Int"), "Int", NumGet(button, 12, "Int"), "Ptr", A_ScriptHwnd, "Ptr", params)
        DllCall("PostMessage", "Ptr", A_ScriptHwnd, "UInt", 0, "Ptr", 0, "Ptr", 0)   ; WM_NULL, as the TrackPopupMenu documentation asks
        QuickSwitchFavorites._DialogAfterMenu()
    }
}
