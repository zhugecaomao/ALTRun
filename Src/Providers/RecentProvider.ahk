;===============================================================================
; RecentProvider.ahk - pinned and recently used items in the empty search box (AutoHotkey v2)
;-------------------------------------------------------------------------------
; As in Raycast / Alfred, they are listed when the search window is shown and nothing is typed yet:
;   1. pinned items (Window.ShowPinned): "Pin to empty search box" in the actions panel (→), stored in ALTRun.json (Features.Recent.Pinned),
;      so they go with the settings; reordered / removed in Preferences → Search Window → Edit... (PrefsPages.EditPinned)
;   2. recently opened items (Window.ShowRecent, at most RecentCount), stored in Data\Knowledge.json (Knowledge.Recent); Clear... in Preferences
; Ctrl+Del: unpin / remove from the recently used. Files / folders dropped on the search window are pinned too (PinPaths).
; Only results that can be opened again are kept: files, folders, programs, URLs (path saved) and system commands (Id saved, shown with the name in the current language);
; clipboard history, calculation results, snippets, web searches, windows and others that depend on the input or the moment aren't.
;
; Settings (ALTRun.json): Features.Recent: Enabled (always on) / Pinned; Window: ShowPinned / ShowRecent / RecentCount
;===============================================================================

class RecentProvider {
    static Id := "Recent"
    static Keep := 30                                                       ; at most this many entries kept in Knowledge.json
    static Skip := Map("WebSearch", 1, "Windows", 1, "Clipboard", 1, "Calculator", 1, "Snippets", 1, "Help", 1, "Terminal", 1)

    static Init() {
    }

    static Search(query) {
        return []
    }

    ; Empty search box: pinned first, then the recently used (without the pinned ones shown above)
    static EmptyResults() {
        options := AppSettings.Feature("Recent"), window := AppSettings.Window
        results := [], seen := Map()
        if window["ShowPinned"]
            for entry in options["Pinned"] {
                if IsObject(item := RecentProvider.ItemFor(entry, true))
                    results.Push(item), seen[entry["Uid"]] := true
            }
        shown := 0, limit := window["ShowRecent"] ? window["RecentCount"] : 0
        for entry in Knowledge.Recent {
            if (shown >= limit)
                break
            if (!(entry is Map) || !entry.Has("Uid") || seen.Has(entry["Uid"]))
                continue
            if IsObject(item := RecentProvider.ItemFor(entry, false)) {
                results.Push(item), seen[entry["Uid"]] := true, shown += 1
            }
        }
        return results
    }

    ; Result -> the Map to save; "" for results that can't be opened again
    static Snapshot(item) {
        if (!IsObject(item) || item.Provider = "" || RecentProvider.Skip.Has(item.Provider) || !item.Valid)
            return ""
        if RegExMatch(item.Uid, "^system:(.+)$", &m)
            return Map("System", m[1], "Title", item.Title, "Uid", item.Uid)
        if (RegExMatch(item.Kind, "^(file|folder|url)$") && item.Arg != "" && !IsObject(item.OnRun)) {
            entry := Map("Title", item.Title, "Subtitle", item.Subtitle, "Kind", item.Kind, "Arg", item.Arg, "Arguments", item.Arguments
                       , "Icon", IsObject(item.Icon) ? "" : item.Icon, "Uid", (item.Uid != "") ? item.Uid : item.Kind ":" StrLower(item.Arg))
            if item.WindowArgs                                              ; {hwnd} / {pid} stay as they are, filled in each time it runs
                entry["WindowArgs"] := 1
            return entry
        }
        return ""
    }

    ; An item was opened (Enter / an action / the context menu in the search window)
    static Remember(item) {
        if !IsObject(entry := RecentProvider.Snapshot(item))
            return
        recent := Knowledge.Recent
        for index, previous in recent {
            if (previous is Map && previous.Has("Uid") && previous["Uid"] = entry["Uid"]) {
                recent.RemoveAt(index)
                break
            }
        }
        recent.InsertAt(1, entry)
        while (recent.Length > RecentProvider.Keep)
            recent.Pop()
        Knowledge._SaveLater()
    }

    static IsPinned(item) {
        uid := IsObject(entry := RecentProvider.Snapshot(item)) ? entry["Uid"] : ""
        for pinned in AppSettings.Feature("Recent")["Pinned"]
            if (pinned is Map && pinned.Has("Uid") && pinned["Uid"] = uid)
                return true
        return false
    }

    static CanPin(item) => IsObject(RecentProvider.Snapshot(item))

    static Pin(item) {
        if (!IsObject(entry := RecentProvider.Snapshot(item)) || RecentProvider.IsPinned(item))
            return false
        for pinned in RecentProvider._PinnedLists()
            pinned.Push(A_Index = 1 ? entry : entry.Clone())
        App.Notify(I18n.T("Recent.PinnedHint", item.Title))
        return AppSettings.Save()
    }

    ; Files / folders dropped on the search window: pins the first limit of them that exist and aren't pinned yet, saving once;
    ; returns how many were pinned
    static PinPaths(paths, limit := 20) {
        count := 0, last := ""
        for filePath in paths {
            if (count >= limit)
                break
            if ((attributes := FileExist(filePath)) = "")
                continue
            isFolder := InStr(attributes, "D") > 0
            name := Path.Leaf(RTrim(filePath, "\"))
            item := ResultItem((name != "") ? name : filePath, filePath, {Kind: isFolder ? "folder" : "file", Arg: filePath
                , Icon: isFolder ? IconCache.FolderIcon(filePath) : filePath, Provider: "Drop"})
            if (!IsObject(entry := RecentProvider.Snapshot(item)) || RecentProvider.IsPinned(item))
                continue
            for pinned in RecentProvider._PinnedLists()
                pinned.Push(A_Index = 1 ? entry : entry.Clone())
            count++, last := item.Title
        }
        if !count
            return 0
        App.Notify(count = 1 ? I18n.T("Recent.PinnedHint", last) : I18n.T("Recent.PinnedMany", count))
        AppSettings.Save()
        return count
    }

    static Unpin(uid) {
        found := false
        for pinned in RecentProvider._PinnedLists() {
            for index, entry in pinned {
                if (entry is Map && entry.Has("Uid") && entry["Uid"] = uid) {
                    pinned.RemoveAt(index), found := true
                    break
                }
            }
        }
        return found ? AppSettings.Save() : false
    }

    ; The pinned lists a pin / unpin goes to: the settings, and while Preferences is open its copy too
    ; (OK there replaces all settings with that copy, which would otherwise undo the change)
    static _PinnedLists() {
        lists := [AppSettings.Feature("Recent")["Pinned"]]
        if (IsObject(PreferencesWindow.Gui) && IsObject(PreferencesWindow.Working))
            try {
                working := PreferencesWindow.Working["Features"]["Recent"]["Pinned"]
                if (working is Array && working != lists[1])
                    lists.Push(working)
            }
        return lists
    }

    ; Saved Map -> result; system commands are looked up by Id ("" if the command no longer exists)
    static ItemFor(entry, pinned) {
        if !(entry is Map) || !entry.Has("Uid")
            return ""
        extra := {Source: entry, Uid: entry["Uid"], Pinned: pinned}
        if entry.Has("System") {
            id := SystemProvider.Aliases.Get(entry["System"], entry["System"])   ; a removed command saved by an older version (ListProcesses -> TaskManager)
            for command in SystemProvider.Commands() {
                if (command["Id"] = id) {
                    extra.Icon := command["Icon"], extra.OnRun := SystemProvider._Runner(command["Id"])
                    return RecentProvider._Item(command["Title"], command["Subtitle"], extra, pinned)
                }
            }
            return ""
        }
        if !(entry.Has("Kind") && entry.Has("Arg"))
            return ""
        kind := entry["Kind"], icon := entry.Has("Icon") ? entry["Icon"] : ""
        if (icon = "")
            icon := (kind = "folder") ? IconCache.FolderIcon(entry["Arg"]) : (kind = "url") ? "url:" : Path.Resolve(entry["Arg"])
        extra.Kind := kind, extra.Arg := entry["Arg"], extra.Arguments := entry.Has("Arguments") ? entry["Arguments"] : "", extra.Icon := icon
        extra.WindowArgs := entry.Has("WindowArgs") && entry["WindowArgs"] ? true : false
        return RecentProvider._Item(entry["Title"], entry.Has("Subtitle") ? entry["Subtitle"] : "", extra, pinned)
    }

    static _Item(title, subtitle, extra, pinned) {
        item := ResultItem(title, subtitle)
        for name, value in extra.OwnProps()
            item.%name% := value
        item.Provider := RecentProvider.Id
        return item
    }

    static DeleteItem(item) {
        if item.Pinned
            return RecentProvider.Unpin(item.Uid)
        for index, entry in Knowledge.Recent {
            if (entry is Map && entry.Has("Uid") && entry["Uid"] = item.Uid) {
                Knowledge.Recent.RemoveAt(index)
                Knowledge.Save()
                return true
            }
        }
        return false
    }

    ; Clear... in Preferences: forgets all the recently used items (the pinned ones stay); returns how many there were
    static ClearRecent() {
        count := Knowledge.Recent.Length
        if count {
            Knowledge.Recent.Length := 0
            Knowledge.Save()
        }
        return count
    }

    static DeletePrompt(item) {
        return I18n.T(item.Pinned ? "Recent.ConfirmUnpin" : "Recent.ConfirmForget", item.Title)
    }
}
