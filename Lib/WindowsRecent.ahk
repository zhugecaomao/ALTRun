;===============================================================================
; WindowsRecent.ahk - Windows "Recent items" (AutoHotkey v2)
;-------------------------------------------------------------------------------
; %AppData%\Microsoft\Windows\Recent holds a shortcut to every file and folder opened in any program.
; (%UserProfile%\Recent is only a legacy junction that programs may not list.)
;
; Used by:
;   FileSearchProvider.RecentResults   "recent report": recent files and folders by name
;   QuickSwitch.RecentFolders          the recent folders in the dialog panel
;
; Usage:
;   WindowsRecent.Items(300)            -> [{Path, Name, IsFolder}], newest first
;===============================================================================

class WindowsRecent {
    ; The targets of the newest maxLinks shortcuts in the Recent folder, newest first, each once:
    ; [{Path, Name, IsFolder}]. Local targets that are gone are left out; network targets are not checked (the network can
    ; be slow): no extension = folder. Reading a shortcut is a COM call, so each one is read once and kept until it
    ; changes; the whole list is kept for CacheMs (typing "recent rep" searches once per character)
    static Dir := A_AppData "\Microsoft\Windows\Recent"
    static CacheMs := 2000, MaxLinks := 1200, _links := Map(), _lists := Map()
    static Items(maxLinks, recentDir := "") {
        recentDir := (recentDir != "") ? recentDir : WindowsRecent.Dir
        listKey := StrLower(recentDir) "|" maxLinks
        lists := WindowsRecent._lists
        if (lists.Has(listKey) && A_TickCount - lists[listKey].Time < WindowsRecent.CacheMs)
            return lists[listKey].Items
        lines := ""
        Loop Files, recentDir "\*.lnk"
            lines .= A_LoopFileTimeModified "`t" A_LoopFileFullPath "`n"
        links := WindowsRecent._links
        if (links.Count > WindowsRecent.MaxLinks)                           ; old shortcuts pile up: start over
            links := WindowsRecent._links := Map()
        items := [], seen := Map()
        for line in StrSplit(Sort(RTrim(lines, "`n"), "R"), "`n") {
            if (line = "" || A_Index > maxLinks)
                break
            parts := StrSplit(line, "`t", , 2)
            key := StrLower(parts[2])
            if !(links.Has(key) && links[key].Time = parts[1]) {
                try FileGetShortcut(parts[2], &target)
                catch                                                       ; a broken shortcut: target is left unset
                    target := ""
                links[key] := {Time: parts[1], Target: target}
            }
            target := links[key].Target
            if !(target ~= "^([A-Za-z]:\\|\\\\)")                          ; shell items without a path (Control Panel...)
                continue
            target := RTrim(target, "\")
            if (target ~= "^[A-Za-z]:$")
                target .= "\"
            if seen.Has(StrLower(target))
                continue
            if Path.IsRemote(target)
                isFolder := (Path.Extension(target) = "")
            else if (attributes := FileExist(target))
                isFolder := InStr(attributes, "D") ? true : false
            else
                continue                                                    ; deleted or moved
            seen[StrLower(target)] := true
            SplitPath(target, &name)
            items.Push({Path: target, Name: (name != "") ? name : target, IsFolder: isFolder})
        }
        if (lists.Count > 10)
            lists.Clear()
        lists[listKey] := {Time: A_TickCount, Items: items}
        return items
    }
}
