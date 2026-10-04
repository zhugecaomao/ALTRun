;===============================================================================
; RecentProvider.ahk - 空搜索框里的置顶项目和最近使用 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 Raycast / Alfred 一样, 呼出搜索窗口还没输入时直接列出:
;   1. 置顶的项目: 操作面板 (→) 里 "置顶到空搜索框", 存在 ALTRun.json (Features.Recent.Pinned), 跟设置走
;   2. 最近打开的项目 (RecentCount 个, 0 = 不显示), 存在 Data\Knowledge.json (Knowledge.Recent)
; Ctrl+Del: 取消置顶 / 从最近使用里去掉。
; 只记能重新打开的结果: 文件、文件夹、程序、网址 (保存路径) 和系统命令 (保存 Id, 显示时用当前语言的名称);
; 剪贴板历史、计算结果、片段、网页搜索、窗口等跟输入或当时状态有关的不记。
;
; 设置 (ALTRun.json -> Features.Recent): Enabled / RecentCount / Pinned
;===============================================================================

class RecentProvider {
    static Id := "Recent"
    static Keep := 30                                                       ; Knowledge.json 里最多记多少条
    static Skip := Map("WebSearch", 1, "Windows", 1, "Clipboard", 1, "Calculator", 1, "Snippets", 1, "Help", 1, "Terminal", 1)

    static Init() {
    }

    static Search(query) {
        return []
    }

    ; 空搜索框: 置顶的在前, 然后是最近使用 (去掉已经置顶的)
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

    ; 结果 -> 保存用的 Map; 不能重新打开的结果返回 ""
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

    ; 打开了某一项 (搜索窗口里 Enter / 操作 / 右键菜单)
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

    ; 保存的 Map -> 结果; 系统命令按 Id 找当前的命令 (已经没有这个命令时 "")
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
