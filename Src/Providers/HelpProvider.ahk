;===============================================================================
; HelpProvider.ahk - 速查表 (输入 ?) 和空搜索框里的使用提示 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 一张表 (Entries) 两种用途:
;   ?          列出所有输入语法和快捷键, ? 后面的文字用来过滤 (? 文件), Enter 打开 Wiki 对应的页面
;   NextTip()  搜索窗口每次显示时, 空搜索框里的灰色提示文字按顺序换一条 (General.ShowTips)
; 每一条: [Id, 所属功能 (关闭时不显示, "" = 总是显示), Wiki 页面]
; 显示的文字在 I18n: Help.<Id>.Key (输入 / 按键) 和 Help.<Id>.Text (作用);
; Key 里的 {1} {2} 换成当前设置里的关键字 / 热键 (见 _Args)。
;===============================================================================

class HelpProvider {
    static Id      := "Help"
    static Prefix  := "?"
    static WikiUrl := "https://github.com/zhugecaomao/ALTRun/wiki/"
    static _tipIndex := 0

    static Init() {
    }

    static Entries() {
        return [
            ["Files"     , "FileSearch"    , "File-Search"],
            ["Folders"   , "FileSearch"    , "File-Search"],
            ["FileKeyword", "FileSearch"   , "File-Search"],
            ["Actions"   , ""              , "Usage"],
            ["Edit"      , ""              , "Commands-and-Snippets"],
            ["Delete"    , ""              , "Usage"],
            ["Reveal"    , ""              , "Usage"],
            ["Copy"      , ""              , "Usage"],
            ["Number"    , ""              , "Usage"],
            ["History"   , ""              , "Usage"],
            ["Tab"       , ""              , "Usage"],
            ["Clipboard" , "Clipboard"     , "Commands-and-Snippets"],
            ["Snippets"  , "Snippets"      , "Commands-and-Snippets"],
            ["Expand"    , "Snippets"      , "Commands-and-Snippets"],
            ["Terminal"  , "Terminal"      , "Usage"],
            ["Calculator", "Calculator"    , "Usage"],
            ["WebSearch" , "WebSearch"     , "Usage"],
            ["CheckPaths", "CustomCommands", "Commands-and-Snippets"],
            ["LargeType" , ""              , "Usage"],
            ["Prefs"     , ""              , "Configuration"],
            ["Help"      , ""              , "Usage"]
        ]
    }

    ; 当前设置下要显示的条目: [{Id, Key, Text, Url}...]
    static Items() {
        items := []
        for entry in HelpProvider.Entries() {
            if (entry[2] != "" && !HelpProvider._FeatureEnabled(entry[2]))
                continue
            if (entry[1] = "Expand" && !AppSettings.Feature("Snippets")["AutoExpand"])
                continue
            items.Push({Id: entry[1], Key: I18n.T("Help." entry[1] ".Key", HelpProvider._Args(entry[1])*)
                      , Text: I18n.T("Help." entry[1] ".Text"), Url: HelpProvider.WikiUrl entry[3]})
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
            wikiPage := item.Url                                            ; 不能叫 url: 和 Url 类同名
            results.Push(ResultItem(item.Key, item.Text, {
                Icon: "res:shell32.dll,-24", Score: 200 - results.Length * 0.01, Exclusive: true,
                OnRun: (*) => ActionCatalog.OpenUrl(wikiPage)
            }))
        }
        return results
    }

    ; 按顺序轮换的下一条提示 (每次启动从随机的一条开始)
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
        return false
    }

    ; Key 里的占位符: 用户改过的关键字、热键也显示成改过之后的
    static _Args(id) {
        switch id {
            case "Folders":     return [AppSettings.Feature("FileSearch")["FolderKeywords"].Length ? AppSettings.Feature("FileSearch")["FolderKeywords"][1] : "folder"]
            case "FileKeyword": return [AppSettings.Feature("FileSearch")["Keywords"].Length ? AppSettings.Feature("FileSearch")["Keywords"][1] : "open"]
            case "Clipboard":   return [AppSettings.Feature("Clipboard")["Keyword"], Win.HotkeyLabel(AppSettings.Feature("Clipboard")["Hotkey"])]
            case "Snippets":    return [AppSettings.Feature("Snippets")["Keyword"]]
            case "Expand":      return [AppSettings.Feature("Snippets")["ExpandPrefix"]]
            case "Terminal":    return [AppSettings.Feature("Terminal")["Prefix"]]
            case "WebSearch":
                engines := AppSettings.Feature("WebSearch")["Engines"]
                return [engines.Length ? engines[1]["Keyword"] : "g"]
        }
        return []
    }
}
