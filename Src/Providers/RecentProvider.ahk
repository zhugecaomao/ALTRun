;===============================================================================
; RecentProvider.ahk - pinned and recently used items in the empty search box (AutoHotkey v2)
;-------------------------------------------------------------------------------
; As in Raycast / Alfred, they are listed when the search window is shown and nothing is typed yet:
;   1. pinned items: "Pin to empty search box" in the actions panel (→), stored in ALTRun.json (Features.Recent.Pinned), so they go with the settings
;   2. recently opened items (RecentCount, 0 = none), stored in Data\Knowledge.json (Knowledge.Recent)
; Ctrl+Del: unpin / remove from the recently used.
; Only results that can be opened again are kept: files, folders, programs, URLs (path saved) and system commands (Id saved, shown with the name in the current language);
; clipboard history, calculation results, snippets, web searches, windows and others that depend on the input or the moment aren't.
;
; Settings (ALTRun.json -> Features.Recent): Enabled / RecentCount / Pinned
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

    ; Empty search box: pinned first, then the recently used (without the pinned ones)
    static EmptyResults() {
        options := AppSettings.Feature("Recent")
        results := [], seen := Map()
        for entry in options["Pinned"] {
            if IsObject(item := RecentProvider.ItemFor(entry, true)) {
                results.Push(item), seen[entry["Uid"]] := true
            }
        }
        shown := 0
        for entry in Knowledge.Recent {
            if (shown >= options["RecentCount"])
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
        if (RegExMatch(item.Kind, "^(file|folder|url)$") && item.Arg != "" && !IsObject(item.OnRun))
            return Map("Title", item.Title, "Subtitle", item.Subtitle, "Kind", item.Kind, "Arg", item.Arg, "Arguments", item.Arguments
                     , "Icon", IsObject(item.Icon) ? "" : item.Icon, "Uid", (item.Uid != "") ? item.Uid : item.Kind ":" StrLower(item.Arg))
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
        AppSettings.Feature("Recent")["Pinned"].Push(entry)
        App.Notify(I18n.T("Recent.PinnedHint", item.Title))
        return AppSettings.Save()
    }

    static Unpin(uid) {
        pinned := AppSettings.Feature("Recent")["Pinned"]
        for index, entry in pinned {
            if (entry is Map && entry.Has("Uid") && entry["Uid"] = uid) {
                pinned.RemoveAt(index)
                return AppSettings.Save()
            }
        }
        return false
    }

    ; Saved Map -> result; system commands are looked up by Id ("" if the command no longer exists)
    static ItemFor(entry, pinned) {
        if !(entry is Map) || !entry.Has("Uid")
            return ""
        extra := {Source: entry, Uid: entry["Uid"], Pinned: pinned}
        if entry.Has("System") {
            for command in SystemProvider.Commands() {
                if (command["Id"] = entry["System"]) {
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

    static DeletePrompt(item) {
        return I18n.T(item.Pinned ? "Recent.ConfirmUnpin" : "Recent.ConfirmForget", item.Title)
    }
}
