;===============================================================================
; HelpProvider.ahk - Cheat sheet (type ?) and usage tips in the status bar (AutoHotkey v2)
;-------------------------------------------------------------------------------
; One table (Entries) serves two purposes:
;   ?          lists all input syntax and shortcuts; text after ? filters them (? file), Enter opens the matching Wiki page (by UI language, see WikiPage)
;   NextTip()  each time the search window is shown, the status bar shows the next tip in turn (General.ShowTips)
; Each entry: [Id, feature it belongs to (hidden when that feature is off, "" = always shown), Wiki page]
; The order is the order in the cheat sheet and in the tips: the most useful entries come first.
; The text is in I18n: Help.<Id>.Key (input / keys) and Help.<Id>.Text (what it does);
; {1} {2} in Key are replaced with the keywords / hotkeys from the current settings (see _Args).
;===============================================================================

class HelpProvider {
    static Id      := "Help"
    static Prefix  := "?"
    static WikiUrl := "https://github.com/zhugecaomao/ALTRun/wiki/"
    static _tipIndex := 0

    static Init() {
    }

    ; Wiki page URL: the Chinese UI opens the Chinese page (Usage-zh), other languages the English page (Usage)
    static WikiPage(page, lang := I18n.Lang) {
        return HelpProvider.WikiUrl page (SubStr(lang, 1, 2) = "zh" ? "-zh" : "")
    }

    static Entries() {
        return [
            ["Files"      , "FileSearch"    , "File-Search"],
            ["Folders"    , "FileSearch"    , "File-Search"],
            ["Recent"     , "FileSearch"    , "File-Search"],
            ["FileKeyword", "FileSearch"    , "File-Search"],
            ["FileTypes"  , "FileSearch"    , "File-Search"],
            ["Reveal"     , ""              , "Usage"],
            ["Actions"    , ""              , "Usage"],
            ["QuickSwitch", "QuickSwitch"   , "Quick-Switch"],
            ["About"      , ""              , "Installation"],
            ["Prefs"      , ""              , "Configuration"],
            ["Edit"       , ""              , "Custom-Commands"],
            ["EditJson"   , ""              , "Configuration"],
            ["Selection"  , ""              , "Usage"],
            ["Delete"     , ""              , "Usage"],
            ["Copy"       , ""              , "Usage"],
            ["Number"     , ""              , "Usage"],
            ["History"    , ""              , "Usage"],
            ["Tab"        , ""              , "Usage"],
            ["Clipboard"  , "Clipboard"     , "Clipboard"],
            ["Snippets"   , "Snippets"      , "Snippets"],
            ["Expand"     , "Snippets"      , "Snippets"],
            ["Terminal"   , "Terminal"      , "System-Commands"],
            ["Calculator" , "Calculator"    , "Calculator"],
            ["Convert"    , "Calculator"    , "Calculator"],
            ["WebSearch"  , "WebSearch"     , "Web-Search"],
            ["Windows"    , "Windows"       , "Window-Switching"],
            ["Scripts"    , "Scripts"       , "Scripts"],
            ["ScriptKeyword", "Scripts"     , "Scripts"],
            ["Query"      , "CustomCommands", "Custom-Commands"],
            ["EditFile"   , "FileSearch"    , "Usage"],
            ["Pinyin"     , ""              , "File-Search"],
            ["Browse"     , "FileSearch"    , "File-Search"],
            ["Mark"       , "FileSearch"    , "File-Search"],
            ["Bookmarks"  , "Bookmarks"     , "Bookmarks"],
            ["LargeType"  , "LargeType"     , "Usage"],
            ["CheckPaths" , "CustomCommands", "Custom-Commands"],
            ["Help"       , ""              , "Usage"]
        ]
    }

    ; Entries to show with the current settings: [{Id, Key, Text, Url}...]
    static Items() {
        items := []
        for entry in HelpProvider.Entries() {
            if (entry[2] != "" && !HelpProvider._FeatureEnabled(entry[2]))
                continue
            if (entry[1] = "Expand" && !AppSettings.Feature("Snippets")["AutoExpand"])
                continue
            if (entry[1] = "Selection" && AppSettings.General["SelectionHotkey"] = "")
                continue
            if (entry[1] = "FileTypes" && !FileSearchProvider.TypeFilters().Length)
                continue
            if (entry[1] ~= "^(Folders|FileKeyword|QuickSwitch)$" && HelpProvider._Args(entry[1])[1] = "")   ; no keyword / hotkey left
                continue
            if (entry[1] = "Pinyin" && I18n.Lang != "zh-CN")                ; pinyin initials only cover Simplified Chinese, so only that UI shows it
                continue
            textId := (entry[1] = "Tab" && AppSettings.General["TabKey"] = "Next") ? "TabNext" : entry[1]   ; the Tab tip follows the Tab key setting
            items.Push({Id: entry[1], Key: I18n.T("Help." textId ".Key", HelpProvider._Args(entry[1])*)
                      , Text: I18n.T("Help." textId ".Text"), Url: HelpProvider.WikiPage(entry[3])})
        }
        return items
    }

    static Search(query) {
        if !query.MatchPrefix(HelpProvider.Prefix, &term)
            return []
        results := []
        for item in HelpProvider.Items() {
            if (term != "" && FuzzyMatcher.Best(term, [item.Key, item.Text]) <= 0)
                continue
            wikiPage := item.Url                                            ; not "url": that would shadow the Url class
            results.Push(ResultItem(item.Key, item.Text, {
                Icon: IconCache.Own("Help", "res:shell32.dll,-24"), Score: 200 - results.Length * 0.01, Exclusive: true,
                OnRun: (*) => ActionCatalog.OpenUrl(wikiPage)
            }))
        }
        return results
    }

    ; The next tip in turn (each start begins at a random one)
    static NextTip() {
        items := HelpProvider.Items()
        if !items.Length
            return ""
        if !HelpProvider._tipIndex
            HelpProvider._tipIndex := Random(1, items.Length)
        else
            HelpProvider._tipIndex := Mod(HelpProvider._tipIndex, items.Length) + 1
        item := items[Min(HelpProvider._tipIndex, items.Length)]
        return I18n.T("Help.TipFormat", item.Key, item.Text)
    }

    static _FeatureEnabled(id) {
        try return AppSettings.Feature(id)["Enabled"] ? true : false
        try return AppSettings.Extension(id)["Enabled"] ? true : false          ; e.g. QuickSwitch
        return false
    }

    ; Placeholders in Key: keywords and hotkeys the user changed are shown as changed
    static _Args(id) {
        switch id {
            case "Folders":     return [HelpProvider._Keywords(AppSettings.Feature("FileSearch")["FolderKeywords"], ["folder"], 1)]
            case "Recent":      return [HelpProvider._Keywords(AppSettings.Feature("FileSearch")["RecentKeywords"], ["recent"], 1)]
            case "FileKeyword":                                             ; e.g. "file / open / '"
                keywords := HelpProvider._Keywords(AppSettings.Feature("FileSearch")["Keywords"], ["file", "open"], 2)
                if AppSettings.Feature("FileSearch")["QuotePrefix"]
                    keywords .= (keywords = "" ? "" : " / ") "'"
                return [keywords]
            case "FileTypes":                                               ; e.g. "doc / zip"
                keywords := []
                for filter in FileSearchProvider.TypeFilters()
                    keywords.Push(filter.Keyword)
                return [HelpProvider._Keywords(keywords, ["doc", "zip"], 2)]
            case "QuickSwitch":
                options := AppSettings.Extension("QuickSwitch")                ; one hotkey for TC / Opus / Explorer
                return [(options["Hotkey"] != "") ? Win.HotkeyLabel(options["Hotkey"]) : ""]
            case "Clipboard":   return [AppSettings.Feature("Clipboard")["Keyword"], Win.HotkeyLabel(AppSettings.Feature("Clipboard")["Hotkey"])]
            case "Snippets":    return [AppSettings.Feature("Snippets")["Keyword"]]
            case "ScriptKeyword": return [AppSettings.Feature("Scripts")["Keyword"]]
            case "Expand":      return [AppSettings.Feature("Snippets")["ExpandPrefix"]]
            case "Terminal":    return [AppSettings.Feature("Terminal")["Prefix"]]
            case "Bookmarks":   return [AppSettings.Feature("Bookmarks")["Keyword"]]
            case "Windows":     return [AppSettings.Feature("Windows")["Keyword"]]
            case "Selection":   return [Win.HotkeyLabel(AppSettings.General["SelectionHotkey"])]
            case "WebSearch":
                engines := AppSettings.Feature("WebSearch")["Engines"]
                return [(engines.Length && WebSearchProvider.FirstKeyword(engines[1]) != "") ? WebSearchProvider.FirstKeyword(engines[1]) : "g"]
        }
        return []
    }

    ; Up to max keywords for a short tip: the preferred ones first (when the user still has them), then the rest in order
    static _Keywords(keywords, preferred, max) {
        picked := []
        for word in preferred
            for keyword in keywords
                if (keyword = word && picked.Length < max) {
                    picked.Push(keyword)
                    break
                }
        for keyword in keywords {
            if (picked.Length >= max)
                break
            known := false
            for p in picked
                known := known || (p = keyword)
            if !known
                picked.Push(keyword)
        }
        text := ""
        for keyword in picked
            text .= (text = "" ? "" : " / ") keyword
        return text
    }
}
