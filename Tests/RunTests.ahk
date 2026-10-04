;===============================================================================
; RunTests.ahk - ALTRun 单元测试 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 只测试不依赖界面的纯逻辑: 匹配打分、输入解析、设置升级、计算器、加日期、
; 文字工具、排序等。运行:
;   AutoHotkey64.exe /ErrorStdOut Tests\RunTests.ahk
; 结果输出到标准输出, 退出码 = 失败的数量 (0 = 全部通过)。
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut

#Include %A_ScriptDir%\..\Lib\JSON.ahk
#Include %A_ScriptDir%\..\Lib\Logger.ahk
#Include %A_ScriptDir%\..\Lib\Util.ahk
#Include %A_ScriptDir%\..\Lib\TextTools.ahk
#Include %A_ScriptDir%\..\Lib\Kanji.ahk
#Include %A_ScriptDir%\..\Lib\Everything.ahk
#Include %A_ScriptDir%\..\Lib\Units.ahk
#Include %A_ScriptDir%\..\Lib\ClipboardData.ahk
#Include %A_ScriptDir%\..\Lib\Dialogs.ahk
#Include %A_ScriptDir%\..\Src\Core\App.ahk
#Include %A_ScriptDir%\..\Src\Core\I18n.ahk
#Include %A_ScriptDir%\..\Src\Core\AppSettings.ahk
#Include %A_ScriptDir%\..\Src\Core\SchemaMigration.ahk
#Include %A_ScriptDir%\..\Src\Core\SearchQuery.ahk
#Include %A_ScriptDir%\..\Src\Core\ResultItem.ahk
#Include %A_ScriptDir%\..\Src\Core\FuzzyMatcher.ahk
#Include %A_ScriptDir%\..\Src\Core\Knowledge.ahk
#Include %A_ScriptDir%\..\Src\Core\Usage.ahk
#Include %A_ScriptDir%\..\Src\Core\ActionCatalog.ahk
#Include %A_ScriptDir%\..\Src\Core\ProviderRegistry.ahk
#Include %A_ScriptDir%\..\Src\Core\FileIndex.ahk
#Include %A_ScriptDir%\..\Src\UI\ThemeManager.ahk
#Include %A_ScriptDir%\..\Src\UI\ThemePreview.ahk
#Include %A_ScriptDir%\..\Src\UI\IconCache.ahk
#Include %A_ScriptDir%\..\Src\UI\SearchWindow.ahk
#Include %A_ScriptDir%\..\Src\UI\LargeType.ahk
#Include %A_ScriptDir%\..\Src\UI\Hud.ahk
#Include %A_ScriptDir%\..\Src\UI\ItemEditor.ahk
#Include %A_ScriptDir%\..\Src\UI\HotkeyBox.ahk
#Include %A_ScriptDir%\..\Src\UI\PreferencesWindow.ahk
#Include %A_ScriptDir%\..\Src\Providers\ClipboardProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\ApplicationProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\CustomCommandProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\SnippetProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\SystemProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\CalculatorProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\WebSearchProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\BookmarkProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\FileSearchProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\TerminalProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\WindowProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\RecentProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\ScriptProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\HelpProvider.ahk
#Include %A_ScriptDir%\..\Src\Extensions\SnippetExpander.ahk
#Include %A_ScriptDir%\..\Src\Extensions\QuickSwitch.ahk
#Include %A_ScriptDir%\..\Src\Extensions\AutoDate.ahk
#Include %A_ScriptDir%\..\Src\Extensions\TendonProfile.ahk
#Include %A_ScriptDir%\..\Src\Extensions\PTToolsWindow.ahk
#Include %A_ScriptDir%\..\Src\Extensions\UpdateChecker.ahk
#Include %A_ScriptDir%\..\Src\Extensions\CurrencyRates.ahk
#Include %A_ScriptDir%\..\Src\Extensions\SelectionActions.ahk

OnError((err, mode) => TestRunner.OnUncaught(err, mode))                                             ; 运行错误时输出并退出, 不弹对话框卡住
Logger.Enabled := false
I18n.LangDir := A_ScriptDir "\..\Resources\Lang"
I18n.Init("en")
AppSettings.Data := AppSettings.Defaults()                                  ; 内存里的默认设置, 不读写文件
AppSettings.Feature("Clipboard")["Persist"] := 0                           ; 剪贴板历史测试不写盘

TestRunner.Run()

class TestRunner {
    static Passed := 0, Failed := 0

    static Run() {
        for name in ["FuzzyMatcher", "SearchQuery", "SchemaMigration", "Calculator", "WebSearch"
                    , "AutoDate", "TextTools", "Sorting", "Knowledge", "Clipboard", "ClipboardKinds", "ClipboardPin", "SnippetExpander", "Preferences", "FileIndex", "TopIndexes", "EditActions", "Themes", "ThemeGallery", "IconScaling", "BuiltinIcons", "MatchHighlight", "HighlightSpacing", "CommandTargets", "CommandSearchScale", "CheckTargets", "EditRows", "HiddenApps", "HiddenSystemCommands", "SettingsPages", "WindowSwitch", "RecentItems", "Scripts", "ListFilter", "DefaultFolders", "FileSearchModes", "FolderSearch", "HelpAndTips", "PreferencesButtons", "PreferencesFit", "I18nLanguages", "I18nUnused", "DefaultExamples", "WindowPosition", "PreferenceDescriptions", "SendTo", "HistoryKeys", "TendonProfileVsSpf2m", "TendonProfileInputs", "LegacyIni", "SettingsLocation", "DataLocation", "ReleaseVersion", "ChangelogLinks", "SelfUpdate", "UpdateNotice", "HotkeyText", "JsonReadWrite", "UnitConversion", "CalcBasesDates", "SnippetPlaceholders", "SnippetTextSearch", "SnippetEditor", "Bookmarks", "SelectionItems", "FileTypes", "FolderMenu", "FileActions", "FolderBrowse", "DoubleTap", "BrowseKind", "UsageStats", "HudPlacement", "Misc"] {
            try {
                Tests.%name%()
            } catch as e {
                SplitPath(e.File, &fileName)                                ; 写出文件名和调用栈, 偶发的失败也能看出在哪里
                stack := RegExReplace(Trim(e.Stack, "`r`n"), "\R", " <- ")
                TestRunner.Fail(name, "exception: " e.Message " (" fileName ":" e.Line ") " SubStr(stack, 1, 600))
            }
        }
        FileAppend("`n" TestRunner.Passed " passed, " TestRunner.Failed " failed`n", "*")
        ExitApp(TestRunner.Failed)
    }

    static OnUncaught(err, mode) {
        FileAppend("UNCAUGHT " err.Message " (" err.File ":" err.Line ")`n", "*")
        ExitApp(99)
    }

    static Equal(name, actual, expected) {
        if (actual == expected)
            TestRunner.Passed++
        else
            TestRunner.Fail(name, "expected [" expected "] got [" actual "]")
    }

    static True(name, condition) {
        if condition
            TestRunner.Passed++
        else
            TestRunner.Fail(name, "condition is false")
        return condition ? true : false
    }

    static Fail(name, message) {
        TestRunner.Failed++
        FileAppend("FAIL " name ": " message "`n", "*")
    }
}

class Tests {
    static FuzzyMatcher() {
        eq := (n, a, e) => TestRunner.Equal("FuzzyMatcher." n, a, e)
        ok := (n, c) => TestRunner.True("FuzzyMatcher." n, c)
        eq("exact", FuzzyMatcher.Score("notepad", "Notepad"), 100)
        ok("prefix beats contains", FuzzyMatcher.Score("note", "Notepad") > FuzzyMatcher.Score("pad", "Notepad"))
        ok("word start", FuzzyMatcher.Score("code", "Visual Studio Code") >= 70)
        ok("initials", FuzzyMatcher.Score("vsc", "Visual Studio Code") >= 70)
        ok("camel initials", FuzzyMatcher.Score("ah", "AutoHotkey") >= 70)
        ok("multi token", FuzzyMatcher.Score("st co", "Visual Studio Code") > 0)
        ok("subsequence", FuzzyMatcher.Score("ntpd", "Notepad") > 0)
        eq("no match", FuzzyMatcher.Score("xyz", "Notepad"), 0)
        eq("no scattered match", FuzzyMatcher.Score("snip", "List Running Processes"), 0)
        ok("shorter wins", FuzzyMatcher.Score("word", "Word") > FuzzyMatcher.Score("word", "WordPad Document Editor"))
        ok("regex chars are literal", FuzzyMatcher.Score("c++", "Dev C++ Compiler") > 0)
        ok("\E is safe", FuzzyMatcher.Score("a\Eb", "xx a\Eb") >= 79)
        eq("Initials", FuzzyMatcher.Initials("Visual Studio Code"), "vsc")
        eq("Best", FuzzyMatcher.Best("wx", ["微信", "WX"]), 100)
    }

    static SearchQuery() {
        q := SearchQuery("  g hello  world ")
        TestRunner.Equal("SearchQuery.keyword", q.Keyword, "g")
        TestRunner.Equal("SearchQuery.rest", q.Rest, "hello  world")
        TestRunner.True("SearchQuery.match", q.MatchKeyword(["google", "G"], &term) && term = "hello  world")
        TestRunner.True("SearchQuery.nomatch", !q.MatchKeyword(["bd"], &term))
        q := SearchQuery(">ipconfig /all")
        TestRunner.True("SearchQuery.prefix", q.MatchPrefix(">", &term) && term = "ipconfig /all")
        q := SearchQuery("g")
        TestRunner.True("SearchQuery.keyword only", q.MatchKeyword(["g"], &term) && term = "" && !q.HasRest)
        q := SearchQuery("clip ")
        TestRunner.True("SearchQuery.keyword + space", q.MatchKeyword(["clip"], &term) && term = "" && q.HasRest)
    }

    static SchemaMigration() {
        old := JSON.Parse('
        (
        {
          "Config": {"AutoStartup": 0, "HideOnLostFocus": 1, "FileMgr": "D:\\TC\\TOTALCMD64.EXE", "StruCalc": 1,
                     "IndexDir": "A_ProgramsCommon,A_StartMenu,C:\\Path\\IndexLocation", "IndexDepth": 2,
                     "ClipSendMode": 2, "Chinese": 1, "AutoSwitchDir": 1},
          "Hotkey": {"GlobalHotkey1": "~!Space", "GlobalHotkey2": "!r", "TotalCMDDir": "^g", "ExplorerDir": "^e",
                     "CondTitle": "ahk_exe RAPTW.exe", "CondHotkey": "~Mbutton", "CondAction": "PTTools",
                     "AutoDateBEHKey": "^d"},
          "UserCommand": {
            "File | C:\\Windows\\Notepad.exe": 9,
            "Dir | A_Desktop | Desktop": 99,
            "CMD | cmd.exe /k ipconfig | Check IP Address": 9,
            "URL | www.google.com | Google": 9,
            "Clip | Dear Sir,\\n\\nThanks.\\nLiming | sig": 9,
            "Func | NewCommand | New Command": 99
          },
          "PTTools": {"RebarSpanWidth": 4000},
          "DefaultCommand": {"Func | About | Help": 99}
        }
        )')
        eq := (n, a, e) => TestRunner.Equal("Migration." n, a, e)
        eq("detect", SchemaMigration.DetectVersion(old), 2)
        data := SchemaMigration.Upgrade(old, 2)
        eq("version", data["SchemaVersion"], AppSettings.CurrentVersion)
        eq("hotkey", data["General"]["Hotkey"], "!Space")
        eq("hotkey2", data["General"]["SecondaryHotkey"], "!r")
        eq("startup", data["General"]["LaunchAtLogin"], 0)
        eq("language", data["General"]["Language"], "zh-CN")
        eq("filemanager", data["General"]["FileManager"], "D:\TC\TOTALCMD64.EXE")
        eq("strucalc", data["Features"]["Calculator"]["StructuralCalc"], 1)
        eq("folders", data["Features"]["Applications"]["Folders"].Length, 2)
        eq("depth", data["Features"]["Applications"]["Depth"], 2)
        eq("pastemode", data["Features"]["Snippets"]["PasteMode"], "Type")
        eq("autoswitch", data["Extensions"]["QuickSwitch"]["AutoSwitch"], 1)
        eq("pttools", data["Extensions"]["PTTools"]["RebarSpanWidth"], 4000)
        eq("hotkeys", data["Hotkeys"].Length, 1)
        eq("hotkey action", data["Hotkeys"][1]["Action"], "PTTools")
        eq("commands", data["CustomCommands"].Length, 4)                    ; Func 不迁移
        commands := Map()
        for command in data["CustomCommands"]
            commands[command["Title"]] := command
        eq("cmd split target", commands["Check IP Address"]["Target"], "cmd.exe")
        eq("cmd split args", commands["Check IP Address"]["Arguments"], "/k ipconfig")
        eq("dir type", commands["Desktop"]["Type"], "Folder")
        eq("url type", commands["Google"]["Type"], "Url")
        eq("file title", commands["Notepad"]["Type"], "File")
        eq("snippets", data["Snippets"].Length, 1)
        eq("snippet text", data["Snippets"][1]["Text"], "Dear Sir,`r`n`r`nThanks.`r`nLiming")
        eq("snippet keyword", data["Snippets"][1]["Keyword"], "sig")
        eq("current detect", SchemaMigration.DetectVersion(data), AppSettings.CurrentVersion)
        eq("files not in default results", data["Features"]["FileSearch"]["InDefaultResults"], 0)
        v3 := Map("SchemaVersion", 3, "Features", Map("FileSearch", Map("InDefaultResults", 1, "MaxResults", 50)))
        v4 := SchemaMigration.Upgrade(v3, 3)
        eq("3->4 default results off", v4["Features"]["FileSearch"]["InDefaultResults"], 0)
        eq("3->4 keeps other settings", v4["Features"]["FileSearch"]["MaxResults"], 50)
        eq("3->4 version", v4["SchemaVersion"], 4)
        eq("3->4 without section", SchemaMigration.Upgrade(Map("SchemaVersion", 3), 3)["SchemaVersion"], 4)
        TestRunner.True("Migration.defaults filled", !AppSettings._MergeDefaults(data, AppSettings.Defaults()))
        TestRunner.True("Migration.no downgrade", SchemaMigration.Upgrade(Map("SchemaVersion", 4), 4)["SchemaVersion"] = 4)
    }

    static Calculator() {
        eq := (n, a, e) => TestRunner.Equal("Calculator." n, a, e)
        eq("basic", CalculatorProvider.Search(SearchQuery("12*(3+4)"))[1].Title, "84")
        eq("prefix =", CalculatorProvider.Search(SearchQuery("=2^10"))[1].Title, "1,024")
        eq("decimal", CalculatorProvider.Search(SearchQuery("10/4"))[1].Title, "2.5")
        eq("not math", CalculatorProvider.Search(SearchQuery("notepad")).Length, 0)
        title := (text) => (r := CalculatorProvider.Search(SearchQuery(text))).Length ? r[1].Title : "(none)"
        eq("two decimals", title("113.6+296.6+87.7+143.6+132.9"), "774.4")
        eq("0.1+0.2", title("0.1+0.2"), "0.3")
        eq("round 10/3", title("10/3"), "3.33")
        eq("round 2/3", title("2/3"), "0.67")
        eq("small", title("0.001*4"), "0.004")
        eq("negative", title("-1234.5*2"), "-2,469")
        eq("big", title("100000^4"), "100,000,000,000,000,000,000")
        eq("same operands", title("1*2+11*2"), "24")
        eq("same operands 2", title("2*3+12*3"), "42")
        eq("precedence", title("2+3*4-6/2"), "11")
        eq("left assoc", title("10-2-3"), "5")
        eq("nested parens", title("2*(3+(4-1))"), "12")
        eq("power right assoc", title("2^3^2"), "512")
        eq("power **", title("2**10"), "1,024")
        eq("negative power", title("2^-2"), "0.25")
        eq("sign before power", title("-2^2"), "-4")
        eq("divide by zero", title("5/0"), "(none)")
        eq("incomplete", title("1+"), "(none)")
        eq("unclosed", title("(1+2"), "(none)")
        eq("extra )", title("(1+2))"), "(none)")
        eq("two dots", title("1..2+1"), "(none)")
        TestRunner.True("Calculator.eval full precision", Abs(Calc.Eval("10/3") - 3.3333333333333335) < 1e-12)
        AppSettings.Feature("Calculator")["StructuralCalc"] := 1
        eq("structural rows", CalculatorProvider.Search(SearchQuery("300*2")).Length, 3)
        AppSettings.Feature("Calculator")["StructuralCalc"] := 0
    }

    static WebSearch() {
        items := WebSearchProvider.Search(SearchQuery("g hello world"))
        TestRunner.Equal("WebSearch.count", items.Length, 1)
        TestRunner.Equal("WebSearch.url", items[1].Arg, "https://www.google.com/search?q=hello%20world")
        items := WebSearchProvider.Search(SearchQuery("bd 中文"))
        TestRunner.Equal("WebSearch.utf8", items[1].Arg, "https://www.baidu.com/s?wd=%E4%B8%AD%E6%96%87")
        items := WebSearchProvider.Search(SearchQuery("g"))
        TestRunner.True("WebSearch.keyword only is not runnable", items.Length = 1 && !items[1].Valid)
        fallbacks := WebSearchProvider.Fallbacks(SearchQuery("xyz"))
        TestRunner.Equal("WebSearch.fallbacks", fallbacks.Length, 3)
    }

    static AutoDate() {
        eq := (n, a, e) => TestRunner.Equal("AutoDate." n, a, e)
        today := "23.09.2026"
        eq("with ext", AutoDate.AddDateToName("Report.docx", today), "Report - 23.09.2026.docx")
        eq("update date", AutoDate.AddDateToName("Report - 01.01.2020.docx", today), "Report - 23.09.2026.docx")
        eq("update dash date", AutoDate.AddDateToName("Report-01.01.2020.docx", today), "Report - 23.09.2026.docx")
        eq("folder", AutoDate.AddDateToName("Project", today), "Project - 23.09.2026")
        eq("folder update", AutoDate.AddDateToName("Project - 01.01.2020", today), "Project - 23.09.2026")
        eq("numeric ext", AutoDate.AddDateToName("Version 1.2", today), "Version 1.2 - 23.09.2026")
        eq("dot space", AutoDate.AddDateToName("1. DWG", today), "1. DWG - 23.09.2026")
        ; 设置里换了日期格式: 末尾那种格式的日期照样更新, 不会再加一个
        eq("iso update", AutoDate.AddDateToName("Report - 2020-01-01.docx", "2026-09-23", "yyyy-MM-dd"), "Report - 2026-09-23.docx")
        eq("iso folder", AutoDate.AddDateToName("Project - 2020-01-01", "2026-09-23", "yyyy-MM-dd"), "Project - 2026-09-23")
        eq("dotted iso folder", AutoDate.AddDateToName("Project - 2020.01.01", "2026.09.23", "yyyy.MM.dd"), "Project - 2026.09.23")
        eq("compact", AutoDate.AddDateToName("Plan_20200101.pdf", "20260923", "yyyyMMdd"), "Plan_20200101 - 20260923.pdf")
        eq("compact update", AutoDate.AddDateToName("Plan - 20200101.pdf", "20260923", "yyyyMMdd"), "Plan - 20260923.pdf")
        eq("month name", AutoDate.AddDateToName("Memo - 1 Jan 2020", "23 Sep 2026", "d MMM yyyy"), "Memo - 23 Sep 2026")
        eq("other format kept", AutoDate.AddDateToName("Report - 01.01.2020.docx", "2026-09-23", "yyyy-MM-dd"), "Report - 01.01.2020 - 2026-09-23.docx")
        eq("pattern", AutoDate.DatePattern("dd.MM.yyyy"), "\d{2}\.\d{2}\.\d{4}")
    }

    static TextTools() {
        eq := (n, a, e) => TestRunner.Equal("TextTools." n, a, e)
        eq("sort", TextTools.SortLines("b`r`na`r`nc"), "a`r`nb`r`nc")
        eq("sort desc", TextTools.SortLines("b`na`nc", true), "c`r`nb`r`na")
        eq("dedupe", TextTools.DedupeLines("a`nb`na`nc"), "a`r`nb`r`nc")
        eq("blank", TextTools.RemoveBlankLines("a`n`n  `nb"), "a`r`nb")
        eq("trim", TextTools.TrimLines("  a `n`tb"), "a`r`nb")
        eq("reverse", TextTools.Reverse("abc"), "cba")
        eq("url", Url.Encode("a b/中"), "a%20b%2F%E4%B8%AD")
    }

    static TopIndexes() {
        top := FuzzyMatcher.TopIndexes(Map(1, 10, 2, 90, 3, 50, 4, 90), 3)
        TestRunner.Equal("TopIndexes.order", top[1] " " top[2] " " top[3], "2 4 3")
        TestRunner.Equal("TopIndexes.few", FuzzyMatcher.TopIndexes(Map(5, 1, 7, 9), 10)[1], 7)
    }

    static Sorting() {
        items := [ResultItem("low", "", {Score: 10}), ResultItem("high", "", {Score: 90}), ResultItem("mid1", "", {Score: 50}), ResultItem("mid2", "", {Score: 50})]
        sorted := ProviderRegistry.SortByScore(items)
        order := ""
        for item in sorted
            order .= item.Title " "
        TestRunner.Equal("Sorting.order", Trim(order), "high mid1 mid2 low")
    }

    static Knowledge() {
        Knowledge.Picks := Map(), Knowledge.QueryPicks := Map(), Knowledge.History := []
        Knowledge._saveTimer := () => 0                                     ; 测试时不写盘
        Knowledge.Record("no", "app:notepad")
        Knowledge.Record("no", "app:notepad")
        TestRunner.True("Knowledge.boost exact", Knowledge.Boost("no", "app:notepad") > Knowledge.Boost("no", "app:other"))
        TestRunner.True("Knowledge.boost prefix", Knowledge.Boost("not", "app:notepad") > 20)
        TestRunner.Equal("Knowledge.history", Knowledge.History[1], "no")
        TestRunner.Equal("Knowledge.history dedupe", Knowledge.History.Length, 1)
    }

    static Clipboard() {
        eq := (n, a, e) => TestRunner.Equal("Clipboard." n, a, e)
        ; 保存: 很长的条目单独存成文件, 读回来和原来一样; 不在历史里的文件会删掉
        saved := [ClipboardProvider.File, ClipboardProvider.Folder, ClipboardProvider.Entries, AppSettings.Feature("Clipboard")["Persist"]]
        root := A_Temp "\ALTRun-test-clip-store"
        try DirDelete(root, true)
        DirCreate(root)
        ClipboardProvider.File := root "\ClipboardHistory.json", ClipboardProvider.Folder := root "\Clipboard"
        AppSettings.Feature("Clipboard")["Persist"] := 1
        long := ""
        Loop 500
            long .= "line " A_Index " `"quoted`" \ tab`t end`r`n"
        ClipboardProvider.Entries := [Map("Text", "short one", "Time", "20260929010101", "App", "a.exe"), Map("Text", long, "Time", "20260929010102", "App", "b.exe")]
        ClipboardProvider.Save()
        files := 0
        Loop Files, root "\Clipboard\*.txt"
            files++
        eq("store: one file", files, 1)
        eq("store: json has no long text", InStr(FileRead(ClipboardProvider.File, "UTF-8"), "line 250") ? 1 : 0, 0)
        ClipboardProvider.Entries := []
        ClipboardProvider._Load()
        eq("store: reload count", ClipboardProvider.Entries.Length, 2)
        eq("store: long text exact", ClipboardProvider.Entries[2]["Text"] == long ? 1 : 0, 1)
        eq("store: short text", ClipboardProvider.Entries[1]["Text"], "short one")
        ClipboardProvider.Entries.RemoveAt(2)
        ClipboardProvider.Save()
        files := 0
        Loop Files, root "\Clipboard\*.txt"
            files++
        eq("store: orphan deleted", files, 0)
        ClipboardProvider.File := saved[1], ClipboardProvider.Folder := saved[2], ClipboardProvider.Entries := saved[3]
        AppSettings.Feature("Clipboard")["Persist"] := saved[4]
        try DirDelete(root, true)
        ClipboardProvider.Entries := []
        eq("empty", ClipboardProvider.Search(SearchQuery("clip")).Length, 1)
        eq("empty invalid", ClipboardProvider.Search(SearchQuery("clip"))[1].Valid, false)
        ClipboardProvider.Add("first entry", "notepad.exe")
        ClipboardProvider.Add("second entry", "code.exe")
        ClipboardProvider.Add("first entry", "notepad.exe")                 ; 重复内容移到最前面
        eq("dedupe", ClipboardProvider.Entries.Length, 2)
        eq("newest first", ClipboardProvider.Entries[1]["Text"], "first entry")
        eq("blank ignored", ClipboardProvider.Add("  `r`n ", ""), false)
        items := ClipboardProvider.Search(SearchQuery("clip"))
        eq("list + clear item", items.Length, 3)
        eq("clear item last", items[3].Title, I18n.T("Clipboard.Clear"))
        eq("filter", ClipboardProvider.Search(SearchQuery("clip second")).Length, 1)
        eq("multi token filter", ClipboardProvider.Search(SearchQuery("clip ent sec")).Length, 1)
        eq("keyword only", ClipboardProvider.Search(SearchQuery("second")).Length, 0)
        TestRunner.True("Clipboard.exclusive after space", ClipboardProvider.Search(SearchQuery("clip "))[1].Exclusive)
        TestRunner.True("Clipboard.not exclusive without space", !ClipboardProvider.Search(SearchQuery("clip"))[1].Exclusive)
        mixed := [ResultItem("a", "", {Exclusive: true}), ResultItem("b")]
        eq("registry keeps exclusive", ProviderRegistry._KeepExclusive(mixed).Length, 1)
        AppSettings.Feature("Clipboard")["MaxItems"] := 2
        ClipboardProvider.Add("third", "")
        eq("max items", ClipboardProvider.Entries.Length, 2)
        AppSettings.Feature("Clipboard")["MaxItems"] := 200
        ClipboardProvider.Remove("third")
        eq("remove", ClipboardProvider.Entries[1]["Text"], "first entry")
        ClipboardProvider.PauseRecording(10000)
        ClipboardProvider._OnChange(1)
        eq("paused", ClipboardProvider.Entries.Length, 1)
        ClipboardProvider.Entries := []
    }

    ; 剪贴板历史里的文件和图片, 连按两次 Ctrl+C 合并
    static ClipboardKinds() {
        eq := (n, a, e) => TestRunner.Equal("ClipboardKinds." n, a, e)
        ok := (n, v) => TestRunner.True("ClipboardKinds." n, v)
        options := AppSettings.Feature("Clipboard")
        saved := [ClipboardProvider.File, ClipboardProvider.Folder, ClipboardProvider.Entries, options["Persist"], options["MaxImages"], ProviderRegistry.Providers]
        root := A_Temp "\ALTRun-test-clip-kinds"
        try DirDelete(root, true)
        DirCreate(root "\Clipboard")
        ClipboardProvider.File := root "\ClipboardHistory.json", ClipboardProvider.Folder := root "\Clipboard"
        options["Persist"] := 1
        ProviderRegistry.Providers := [ClipboardProvider]
        ClipboardProvider.Entries := []

        ; 文件: 多个文件一条, 同样的文件再复制一次只移到最前面
        ClipboardProvider.Add("some text", "notepad.exe")
        ClipboardProvider.AddFiles(["C:\Reports\a.xlsx", "C:\Reports\b.docx"], "explorer.exe")
        ClipboardProvider.Add("C:\Reports\a.xlsx`r`nC:\Reports\b.docx", "")  ; 同样的文字是另一条
        ClipboardProvider.AddFiles(["C:\Reports\a.xlsx", "C:\Reports\b.docx"], "explorer.exe")
        eq("files dedupe", ClipboardProvider.Entries.Length, 3)
        eq("files newest", ClipboardProvider.TypeOf(ClipboardProvider.Entries[1]), "files")
        items := ClipboardProvider.Search(SearchQuery("clip b.docx"))
        eq("files search", items.Length, 2)
        item := items[1], item.Provider := "Clipboard"
        eq("files title", item.Title, "a.xlsx, b.docx")
        eq("files subtitle", item.Subtitle, I18n.T("Clipboard.FilesSubtitle", ClipboardProvider._FormatTime(item.Source["Time"]), "explorer.exe", 2))
        eq("files kind", item.Kind, "text")
        actions := ActionCatalog.ListFor(item)
        eq("files first action", actions[1].Title, I18n.T("Clipboard.PasteFiles"))
        eq("files first hint", actions[1].Subtitle, "Enter")
        eq("files copy no enter", actions[2].Subtitle, "")
        ok("files add all", ActionCatalog.CanEdit(item))
        ClipboardProvider.AddFiles([A_WinDir], "explorer.exe")               ; 一个文件夹: 和文件搜索的结果一样
        single := ClipboardProvider.Search(SearchQuery("clip"))[1]
        eq("single kind", single.Kind, "folder")
        eq("single arg", single.Arg, A_WinDir)
        eq("text at 1", ClipboardProvider.TextAt(1), "C:\Reports\a.xlsx`r`nC:\Reports\b.docx")
        eq("text at 2", ClipboardProvider.TextAt(2), "some text")
        eq("text at 3", ClipboardProvider.TextAt(3), "")

        ; 图片: 一模一样的图片只保留一张; 超过 MaxImages 删掉最早的
        FileAppend("image-one", root "\Clipboard\one.png")
        FileAppend("image-two", root "\Clipboard\two.png")
        FileAppend("image-one", root "\Clipboard\dup.png")
        ClipboardProvider.AddImageFile("one.png", 800, 600, "mspaint.exe")
        ClipboardProvider.AddImageFile("two.png", 800, 600, "mspaint.exe")
        ClipboardProvider.AddImageFile("dup.png", 800, 600, "mspaint.exe")
        ok("image dup deleted", !FileExist(root "\Clipboard\dup.png"))
        eq("image dup moved", ClipboardProvider.Entries[1]["Image"], "one.png")
        imageItem := ClipboardProvider.Search(SearchQuery("clip 800"))[1], imageItem.Provider := "Clipboard"
        eq("image title", imageItem.Title, I18n.T("Clipboard.Image", 800, 600))
        eq("image kind", imageItem.Kind, "file")
        eq("image arg", imageItem.Arg, root "\Clipboard\one.png")
        eq("image thumbnail", imageItem.Icon, "thumb:" root "\Clipboard\one.png")
        eq("image search word", ClipboardProvider.Search(SearchQuery("clip image")).Length, 2)
        ok("image not editable", !ActionCatalog.CanEdit(imageItem))
        ok("image deletable", ActionCatalog.CanDelete(imageItem))
        eq("image first action", ActionCatalog.ListFor(imageItem)[1].Title, I18n.T("Clipboard.PasteImage"))
        options["MaxImages"] := 1
        ClipboardProvider._Trim()
        imageCount := 0
        for entry in ClipboardProvider.Entries
            imageCount += ClipboardProvider.TypeOf(entry) = "image"
        eq("max images", imageCount, 1)

        ; 保存和读回; 不在历史里的图片文件删掉
        ClipboardProvider.Save()
        ok("orphan png deleted", !FileExist(root "\Clipboard\two.png"))
        ok("kept png", FileExist(root "\Clipboard\one.png"))
        before := ClipboardProvider.Entries.Length
        ClipboardProvider.Entries := []
        ClipboardProvider._Load()
        eq("reload count", ClipboardProvider.Entries.Length, before)
        eq("reload image", ClipboardProvider.Entries[1]["Image"], "one.png")
        eq("reload image width", ClipboardProvider.Entries[1]["Width"], 800)
        eq("reload folder", ClipboardProvider.Entries[2]["Files"][1], A_WinDir)
        eq("reload files text", ClipboardProvider.Entries[3]["Text"], "C:\Reports\a.xlsx`r`nC:\Reports\b.docx")
        ClipboardProvider.DeleteItem(imageItem)                              ; 旧的对象已经不在了: 什么也不删
        eq("delete stale", ClipboardProvider.Entries.Length, before)
        ClipboardProvider.RemoveEntry(ClipboardProvider.Entries[1])
        eq("delete image", ClipboardProvider.TypeOf(ClipboardProvider.Entries[1]), "files")
        FileDelete(root "\Clipboard\one.png")                               ; 图片文件没了: 读回时跳过
        ClipboardProvider.Entries.InsertAt(1, Map("Type", "image", "Image", "one.png", "Width", 1, "Height", 1, "Text", "", "Time", A_Now, "App", ""))
        ClipboardProvider.Save(), ClipboardProvider._Load()
        eq("missing image skipped", ClipboardProvider.TypeOf(ClipboardProvider.Entries[1]), "files")

        ; 连按两次 Ctrl+C: 第二次复制同样的文字时接到前一条后面
        ClipboardProvider.Entries := []
        ClipboardProvider.Add("first part", "")
        ClipboardProvider.Add("second part", "")
        eq("merge different text", ClipboardProvider._Merge("other", ""), false)
        ok("merge", ClipboardProvider._Merge("second part", ""))
        eq("merge count", ClipboardProvider.Entries.Length, 1)
        eq("merge text", ClipboardProvider.Entries[1]["Text"], "first part`r`nsecond part")
        eq("merge clipboard", A_Clipboard, "first part`r`nsecond part")
        ClipboardProvider.AddFiles(["C:\x.txt"], "")
        eq("merge needs text", ClipboardProvider._Merge("C:\x.txt", ""), false)

        ; 剪贴板里的文件 (CF_HDROP) 和图片 (位图) 放进去再读出来
        ClipboardProvider.PauseRecording(3000)
        ok("set files", ClipboardData.SetFiles([A_WinDir "\win.ini", A_WinDir "\System32"]))
        ok("has files", ClipboardData.HasFiles())
        clipFiles := ClipboardData.Files()
        eq("files count", clipFiles.Length, 2)
        eq("files path", clipFiles.Length ? clipFiles[1] : "", A_WinDir "\win.ini")
        sample := A_ScriptDir "\..\docs\images\screenshots\clipboard.png"
        if FileExist(sample) {
            ok("set image", ClipboardData.SetImage(sample))
            ok("has image", ClipboardData.HasImage())
            size := ClipboardData.SaveImage(root "\out.png")
            ok("save image", IsObject(size) && FileExist(root "\out.png"))
            ok("thumbnail icon", IconCache._FromImage(root "\out.png"))
            if IsObject(size) {
                width := 0, height := 0, bitmap := 0
                DllCall("gdiplus\GdipCreateBitmapFromFile", "WStr", sample, "Ptr*", &bitmap)
                DllCall("gdiplus\GdipGetImageWidth", "Ptr", bitmap, "UInt*", &width)
                DllCall("gdiplus\GdipGetImageHeight", "Ptr", bitmap, "UInt*", &height)
                DllCall("gdiplus\GdipDisposeImage", "Ptr", bitmap)
                eq("image size", size.Width "x" size.Height, width "x" height)
            }
        }
        A_Clipboard := ""

        ClipboardProvider.File := saved[1], ClipboardProvider.Folder := saved[2], ClipboardProvider.Entries := saved[3]
        options["Persist"] := saved[4], options["MaxImages"] := saved[5], ProviderRegistry.Providers := saved[6]
        try DirDelete(root, true)
    }

    static SnippetExpander() {
        eq := (n, a, e) => TestRunner.Equal("SnippetExpander." n, a, e)
        eq("prefix", SnippetExpander.Abbreviation(Map("Keyword", "sig", "Text", "x"), ";"), ";sig")
        eq("no keyword", SnippetExpander.Abbreviation(Map("Keyword", "", "Text", "x"), ";"), "")
        eq("disabled", SnippetExpander.Abbreviation(Map("Keyword", "sig", "Text", "x", "AutoExpand", 0), ";"), "")
        eq("space", SnippetExpander.Abbreviation(Map("Keyword", "a b", "Text", "x"), ";"), "")
        eq("no prefix", SnippetExpander.Abbreviation(Map("Keyword", "sig", "Text", "x"), ""), "sig")
        options := AppSettings.Feature("Snippets"), saved := options["ExpandExclude"]
        options["ExpandExclude"] := " ahk_exe mstsc.exe ,, ahk_class KeePass "
        SnippetExpander.Refresh()
        eq("exclude list", SnippetExpander._exclude.Length "|" SnippetExpander._exclude[2], "2|ahk_class KeePass")
        eq("default exclude", AppSettings.Defaults()["Features"]["Snippets"]["ExpandExclude"] != "", true)
        options["ExpandExclude"] := saved
        SnippetExpander.Refresh()
    }

    static Preferences() {
        eq := (n, a, e) => TestRunner.Equal("Preferences." n, a, e)
        data := PreferencesWindow.DeepCopy(AppSettings.Defaults())
        eq("get", PreferencesWindow.GetPath(data, "General.Hotkey"), "!Space")
        eq("get missing", PreferencesWindow.GetPath(data, "General.NoSuchKey"), "")
        PreferencesWindow.SetPath(data, "Features.Calculator.StructuralCalc", 1)
        eq("set", data["Features"]["Calculator"]["StructuralCalc"], 1)
        PreferencesWindow.SetPath(data, "New.Section.Value", "x")
        eq("set creates", data["New"]["Section"]["Value"], "x")
        original := AppSettings.Defaults()
        copy := PreferencesWindow.DeepCopy(original)
        copy["General"]["Hotkey"] := "^Space"
        eq("deep copy independent", original["General"]["Hotkey"], "!Space")
        eq("lines", PreferencesWindow.SplitLines(" a `r`n`r`nb ").Length, 2)
        eq("join lines", PreferencesWindow.JoinLines(["a", "b"]), "a`r`nb")
        eq("csv", PreferencesWindow.SplitCsv("open, find ,, x")[3], "x")
        eq("join csv", PreferencesWindow.JoinCsv(["open", "find"]), "open, find")
    }

    static FileIndex() {
        eq := (n, a, e) => TestRunner.Equal("FileIndex." n, a, e)
        FileIndex.Paths := ["C:\Docs\Project Omega", "C:\Docs\Project Omega\Omega Plan.dwg", "C:\Docs\notes.txt", "C:\Docs\Tender Report.pdf"]
        FileIndex.Names := ["project omega", "omega plan.dwg", "notes.txt", "tender report.pdf"]
        FileIndex.Folders := [1, 0, 0, 0]
        FileIndex._lastNeedle := "", FileIndex._lastMatches := ""
        found := FileIndex.Search("omega", 10)
        eq("count", found.Length, 2)
        eq("prefix first", found[1].Path, "C:\Docs\Project Omega\Omega Plan.dwg")
        eq("narrowed", FileIndex.Search("omega p", 10).Length, 1)
        eq("new search", FileIndex.Search("rep", 10)[1].Path, "C:\Docs\Tender Report.pdf")
        TestRunner.True("FileIndex.word start >= 70", FileIndex.ScoreName("omega", "project omega") >= 70)
        eq("no match", FileIndex.ScoreName("xyz", "notes.txt"), 0)
        FileIndex.Paths := [], FileIndex.Names := [], FileIndex.Folders := []
    }

    ; 文件类型筛选: doc 报告 / cad 平面图 (和 Listary 一样)
    static FileTypes() {
        eq := (n, a, e) => TestRunner.Equal("FileTypes." n, a, e)
        options := AppSettings.Feature("FileSearch")
        savedEverything := options["UseEverything"]
        options["UseEverything"] := 0                                       ; 测试用内置索引
        parsed := FileSearchProvider.ParseTypeFilter(" Doc = doc, .DOCX;pdf  *.doc ")
        eq("parse keyword", parsed.Keyword, "doc")
        eq("parse everything", parsed.Everything, "ext:doc;docx;pdf")
        eq("parse invalid", FileSearchProvider.ParseTypeFilter("no equals sign"), "")
        eq("parse empty list", FileSearchProvider.ParseTypeFilter("x = , `;"), "")
        keywords := ""
        for filter in FileSearchProvider.TypeFilters()
            keywords .= filter.Keyword " "
        eq("default keywords", keywords, "doc pic video audio zip exe cad ")
        FileIndex.Paths := ["C:\P\Tower Plan.dwg", "C:\P\Tower Plan.pdf", "C:\P\Tower photo.jpg", "C:\P\Tower"]
        FileIndex.Names := ["tower plan.dwg", "tower plan.pdf", "tower photo.jpg", "tower"]
        FileIndex.Folders := [0, 0, 0, 1]
        FileIndex._lastNeedle := "", FileIndex._lastMatches := ""
        titles(results) {
            list := ""
            for item in results
                if (item.Kind = "file" || item.Kind = "folder")
                    list .= item.Title "|"
            return RTrim(list, "|")
        }
        eq("cad", titles(FileSearchProvider.Search(SearchQuery("cad tower"))), "Tower Plan.dwg")
        eq("doc", titles(FileSearchProvider.Search(SearchQuery("doc tower"))), "Tower Plan.pdf")
        eq("pic", titles(FileSearchProvider.Search(SearchQuery("pic tower"))), "Tower photo.jpg")
        eq("exclusive", FileSearchProvider.Search(SearchQuery("cad tower"))[1].Exclusive, true)
        eq("file mode", titles(ProviderRegistry.SearchFiles("cad tower")), "Tower Plan.dwg")
        eq("file mode plain", titles(ProviderRegistry.SearchFiles("tower")), "Tower|Tower Plan.dwg|Tower Plan.pdf|Tower photo.jpg")
        hint := FileSearchProvider.Search(SearchQuery("cad"))
        eq("keyword alone hint", hint.Length, 1)
        eq("keyword alone low score", hint[1].Score < 10 && !hint[1].Exclusive, true)
        eq("hint lists extensions", InStr(hint[1].Title, "dwg dxf") ? 1 : 0, 1)
        eq("not a keyword", FileSearchProvider.Search(SearchQuery("cadence")).Length, 0)
        eq("everything prefix", FileSearchProvider._ScopePrefix(parsed), "ext:doc;docx;pdf ")
        eq("everything folder prefix", FileSearchProvider._ScopePrefix(true), "folder:")
        options["TypeFilters"].Push("dwgonly = dwg")                        ; 改了设置马上生效
        eq("settings change", titles(FileSearchProvider.Search(SearchQuery("dwgonly tower"))), "Tower Plan.dwg")
        options["TypeFilters"].Pop()
        options["UseEverything"] := savedEverything
        FileIndex.Paths := [], FileIndex.Names := [], FileIndex.Folders := []
    }

    ; 对话框的文件夹菜单: 最近的文件夹 (Windows 的 "最近使用的项目")、菜单文字
    static FolderMenu() {
        eq := (n, a, e) => TestRunner.Equal("FolderMenu." n, a, e)
        root := A_Temp "\ALTRun-test-recent"
        try DirDelete(root, true)
        DirCreate(root "\Recent"), DirCreate(root "\Projects\Tower"), DirCreate(root "\Docs")
        longPath := Buffer(2048)                                            ; A_Temp 可能是 8.3 短路径 (C:\Users\RUNNER~1), 快捷方式读回来的是长路径
        if DllCall("GetLongPathNameW", "Str", root, "Ptr", longPath, "UInt", 1024)
            root := StrGet(longPath)
        FileAppend("x", root "\Docs\report.pdf")
        FileCreateShortcut(root "\Projects\Tower", root "\Recent\Tower.lnk")
        FileSetTime("20260101000000", root "\Recent\Tower.lnk")
        FileCreateShortcut(root "\Docs\report.pdf", root "\Recent\report.pdf.lnk")    ; 文件: 取所在的文件夹
        FileSetTime("20260201000000", root "\Recent\report.pdf.lnk")
        FileCreateShortcut(root "\Docs", root "\Recent\Docs.lnk")                     ; 和上一条同一个文件夹: 只出现一次
        FileSetTime("20260115000000", root "\Recent\Docs.lnk")
        FileCreateShortcut(root "\Gone", root "\Recent\Gone.lnk")                     ; 已经不存在
        FileSetTime("20260301000000", root "\Recent\Gone.lnk")
        folders := QuickSwitch.RecentFolders(10, root "\Recent")
        eq("count", folders.Length, 2)
        eq("newest first", folders.Length ? folders[1] : "", root "\Docs")
        eq("second", folders.Length > 1 ? folders[2] : "", root "\Projects\Tower")
        eq("limit", QuickSwitch.RecentFolders(1, root "\Recent").Length, 1)
        eq("zero", QuickSwitch.RecentFolders(0, root "\Recent").Length, 0)
        eq("default hotkey", AppSettings.Defaults()["Extensions"]["QuickSwitch"]["MenuHotkey"], "^+g")
        eq("panel on by default", AppSettings.Defaults()["Extensions"]["QuickSwitch"]["ShowPanel"], 1)
        area := {Left: 0, Top: 0, Right: 1920, Bottom: 1040}
        pos := QuickSwitch.PanelPosition(400, 200, 800, 500, 800, 300, area)
        eq("panel below", pos.X "," pos.Y, "400,700")
        pos := QuickSwitch.PanelPosition(100, 600, 800, 500, 800, 300, area)
        eq("panel above", pos.X "," pos.Y, "100,300")
        pos := QuickSwitch.PanelPosition(1500, 100, 800, 900, 800, 300, area)
        eq("panel kept on screen", pos.X "," pos.Y, "1120,740")
        eq("folder name", QuickSwitch.FolderName("D:\Projects\Tower A\"), "Tower A")
        eq("drive name", QuickSwitch.FolderName("D:\"), "D:\")
        base := [{Path: "D:\Projects\Tower A", Tag: "TC"}, {Path: "C:\Docs", Tag: "Recent"}]
        filtered := QuickSwitch.FilterFolders(base, "tower", ["D:\Projects\Tower A", "E:\Archive\Tower B"])
        eq("filter count", filtered.Length, 2)
        eq("filter keeps listed first", filtered[1].Tag, "TC")
        eq("filter adds found", filtered[2].Path "|" filtered[2].Tag, "E:\Archive\Tower B|" I18n.T("QuickSwitch.TagSearch"))
        eq("filter words", QuickSwitch.FilterFolders(base, "proj tow", []).Length, 1)
        eq("filter empty", QuickSwitch.FilterFolders(base, "", []).Length, 2)
        withFile := QuickSwitch.FilterFolders(base, "tower", [{Path: "E:\Tower", IsFolder: true}, {Path: "E:\Tower\plan.dwg", IsFolder: false}])
        eq("file entry", withFile[3].IsFile "|" withFile[3].Tag, "1|" I18n.T("QuickSwitch.TagFile"))
        eq("folder entry", withFile[2].IsFile, false)
        eq("restore file name", QuickSwitch.ShouldRestoreName("Report 2026.docx"), true)
        eq("no restore empty", QuickSwitch.ShouldRestoreName("  "), false)
        eq("no restore filter", QuickSwitch.ShouldRestoreName("*.txt"), false)
        eq("no restore path", QuickSwitch.ShouldRestoreName("C:\Docs\a.txt"), false)
        ; TC 的路径用 WM_COPYDATA 问 (另一个进程模拟 TC 的回复), 不碰剪贴板
        A_Clipboard := "clipboard before"
        Run('"' A_AhkPath '" "' A_ScriptDir '\Fixtures\FakeTotalCmd.ahk"', , , &fakePid)
        fakeHwnd := WinWait("ALTRun Fake Total Commander ahk_pid " fakePid, , 5)
        eq("tc source panel", QuickSwitch.TotalCmdFolder(fakeHwnd), "C:\Work\Current")
        eq("tc other panel", QuickSwitch.TotalCmdFolder(fakeHwnd, true), "\\server\share\Other")
        eq("tc active side", QuickSwitch.TotalCmdAsk(fakeHwnd, "A"), "L")
        eq("tc not a folder", QuickSwitch.TotalCmdAsk(fakeHwnd, "LP"), "ftp://example.com/pub/")
        eq("tc clipboard untouched", A_Clipboard, "clipboard before")
        ProcessClose(fakePid), ProcessWaitClose(fakePid, 5)
        eq("tc closed", QuickSwitch.TotalCmdFolder(fakeHwnd), "")
        closed := Gui(), closedHwnd := closed.Hwnd
        closed.Destroy()
        threw := false
        try QuickSwitch.SetDialogPath(A_Temp "\", closedHwnd)               ; 对话框已经关掉: 什么也不做 (不往别的窗口里打字)
        catch
            threw := true
        eq("closed dialog ignored", threw, false)
        QuickSwitch._BuildPanel(A_ScriptHwnd)                              ; 真的建一次面板 (Gui 选项写错时这里就会报错)
        TestRunner.True("FolderMenu.panel built", IsObject(QuickSwitch._panel) && IsObject(QuickSwitch._panelList))
        QuickSwitch._ShowFolders(base)
        eq("panel rows", QuickSwitch._panelList.GetCount(), 2)
        eq("panel name column", QuickSwitch._panelList.GetText(1, 1), "Tower A")
        QuickSwitch._panelSearch.Value := "docs"
        QuickSwitch._panelBase := base
        QuickSwitch._SearchPanel()
        eq("panel search filters", QuickSwitch._panelList.GetText(1, 2), "C:\Docs")
        eq("resize before", QuickSwitch._panelWidth, QuickSwitch.PanelWidthFor(QuickSwitch._panelDialogW ? QuickSwitch._panelDialogW : 600))
        QuickSwitch._ResizePanel(900)                                      ; 对话框变宽: 面板跟着变宽
        QuickSwitch._panelList.GetPos(, , &listW)
        eq("resize width", QuickSwitch._panelWidth "|" listW, "900|900")
        eq("key other window", QuickSwitch._OnKeyDown(0x28, 0, 0x100, A_ScriptHwnd), "")
        eq("key down handled", QuickSwitch._OnKeyDown(0x28, 0, 0x100, QuickSwitch._panelSearch.Hwnd), 0)
        eq("ime key passes", QuickSwitch._OnKeyDown(0xE5, 0, 0x100, QuickSwitch._panelSearch.Hwnd), "")   ; 输入法正在输入: 不拦截
        searchHwnd := QuickSwitch._searchHwnd
        QuickSwitch.HidePanel()
        eq("panel hidden", QuickSwitch._panel, "")
        eq("controls released", QuickSwitch._panelSearch "|" QuickSwitch._searchHwnd, "|0")
        eq("key after hide", QuickSwitch._OnKeyDown(0x0D, 0, 0x100, searchHwnd), "")   ; 面板关掉之后: 不报错, 也不处理
        eq("key no panel", QuickSwitch._OnKeyDown(0x0D, 0, 0x100, 0), "")
        eq("width min", QuickSwitch.PanelWidthFor(300), 420)
        eq("width max", QuickSwitch.PanelWidthFor(1600), 1000)
        eq("dark color", QuickSwitch.IsDarkColor("1E1E1E"), true)
        eq("light color", QuickSwitch.IsDarkColor("FAFAFA"), false)
        eq("bad color", QuickSwitch.IsDarkColor("abc"), false)
        savedTheme := ThemeManager.Current
        ThemeManager.Current := Map("Background", "202020", "Title", "EEEEEE", "Separator", "333333")
        colors := QuickSwitch.PanelColors()
        eq("theme colors", colors.Background "|" colors.Text "|" colors.Dark, "202020|EEEEEE|1")
        ThemeManager.Current := Map()
        eq("no theme", QuickSwitch.PanelColors().Background, "FFFFFF")
        ThemeManager.Current := savedTheme
        eq("panel search default", AppSettings.Defaults()["Extensions"]["QuickSwitch"]["PanelSearch"], "folders")
        try DirDelete(root, true)
    }

    ; 文件的操作: 复制 / 剪切文件、打开方式、移到回收站
    static FileActions() {
        eq := (n, a, e) => TestRunner.Equal("FileActions." n, a, e)
        root := A_Temp "\ALTRun-test-fileactions"
        try DirDelete(root, true)
        DirCreate(root)
        FileAppend("x", root "\plan.dwg")
        item := FileSearchProvider._ToItem({Path: root "\plan.dwg", IsFolder: false}, 10)
        titles := "|"
        for action in ActionCatalog.ListFor(item)
            titles .= action.Title "|"
        for key in ["Action.OpenWith", "Action.CopyFile", "Action.CutFile", "Action.Recycle"]
            eq("has " key, InStr(titles, "|" I18n.T(key) "|") ? 1 : 0, 1)
        missing := FileSearchProvider._ToItem({Path: root "\gone.dwg", IsFolder: false}, 10)
        titles := "|"
        for action in ActionCatalog.ListFor(missing)
            titles .= action.Title "|"
        eq("missing file: no copy", InStr(titles, "|" I18n.T("Action.CopyFile") "|") ? 1 : 0, 0)
        drive := FileSearchProvider._ToItem({Path: "C:\", IsFolder: true}, 10)
        titles := "|"
        for action in ActionCatalog.ListFor(drive)
            titles .= action.Title "|"
        eq("drive: no recycle", InStr(titles, "|" I18n.T("Action.Recycle") "|") ? 1 : 0, 0)
        ClipboardProvider.PauseRecording(3000)
        TestRunner.True("FileActions.copy file", ActionCatalog.CopyFiles([root "\plan.dwg"]))
        eq("copied path", ClipboardData.Files().Length ? ClipboardData.Files()[1] : "", root "\plan.dwg")
        eq("copy is not cut", ClipboardData.IsCut(), false)
        ActionCatalog.CopyFiles([root "\plan.dwg"], true)
        eq("cut", ClipboardData.IsCut(), true)
        A_Clipboard := ""
        try DirDelete(root, true)
    }

    ; 自定义命令编辑框: 类型选 "文件夹" 时浏览按钮选择文件夹
    static BrowseKind() {
        eq := (n, a, e) => TestRunner.Equal("BrowseKind." n, a, e)
        fields := CustomCommandProvider.EditorFields()
        eq("target has rule", fields[3].FolderWhen[2], "Folder")
        controls := Map("Type", {Value: 2})                                  ; 第 2 项 = 文件夹
        eq("folder type", ItemEditor.KindFor(fields[3].FolderWhen, fields, controls), "folder")
        controls["Type"].Value := 1
        eq("file type", ItemEditor.KindFor(fields[3].FolderWhen, fields, controls), "file")
        controls["Type"].Value := 4
        eq("url type", ItemEditor.KindFor(fields[3].FolderWhen, fields, controls), "file")
        eq("no such field", ItemEditor.KindFor(["Nope", "Folder"], fields, controls), "file")
    }

    ; 双击 Ctrl / Shift: 只认两次单独的短按
    static DoubleTap() {
        eq := (n, a, e) => TestRunner.Equal("DoubleTap." n, a, e)
        eq("first tap", App.TapDecision("LControl", "LCtrl", 80, 99999), "first")
        eq("second tap", App.TapDecision("LControl", "LCtrl", 80, 250), "show")
        eq("right ctrl", App.TapDecision("RControl", "RCtrl", 80, 250), "show")
        eq("shift", App.TapDecision("LShift", "LShift", 80, 250), "show")
        eq("too slow", App.TapDecision("LControl", "LCtrl", 80, 900), "first")
        eq("ctrl+c", App.TapDecision("c", "LCtrl", 80, 250), "reset")
        eq("held too long", App.TapDecision("LControl", "LCtrl", 800, 250), "reset")
        eq("default off", AppSettings.Defaults()["General"]["DoubleTap"], "")
    }

    static EditActions() {
        eq := (n, a, e) => TestRunner.Equal("EditActions." n, a, e)
        saved := ProviderRegistry.Providers
        ProviderRegistry.Providers := [CustomCommandProvider, FileSearchProvider, ClipboardProvider]
        command := Map("Title", "Notes", "Type", "File", "Target", "C:\notes.txt", "Arguments", "", "Keyword", "")
        custom := CustomCommandProvider._ToItem(command, 50), custom.Provider := "CustomCommands"
        eq("custom edit", ActionCatalog.CanEdit(custom), true)
        eq("custom delete", ActionCatalog.CanDelete(custom), true)
        found := FileSearchProvider._ToItem({Path: "C:\Docs\a.pdf", IsFolder: false}, 10), found.Provider := "FileSearch"
        eq("file edit (add command)", ActionCatalog.CanEdit(found), true)
        eq("file delete", ActionCatalog.CanDelete(found), false)
        system := ResultItem("Lock", "", {OnRun: (*) => 0}), system.Provider := "System"
        eq("system edit", ActionCatalog.CanEdit(system), false)
        titles := ""
        for action in ActionCatalog.ListFor(custom)
            titles .= action.Title "|"
        TestRunner.True("EditActions.list has edit", InStr(titles, I18n.T("Action.Edit")) && InStr(titles, I18n.T("Action.Delete")))
        ProviderRegistry.Providers := saved
    }

    ; 搜索结果标题里高亮匹配的字
    static MatchHighlight() {
        eq := (n, a, e) => TestRunner.Equal("Highlight." n, a, e)
        str(ranges) {
            out := ""
            for range in ranges
                out .= range[1] "+" range[2] " "
            return RTrim(out)
        }
        eq("substring", str(FuzzyMatcher.MatchRanges("stud", "Visual Studio Code")), "8+4")
        eq("prefix", str(FuzzyMatcher.MatchRanges("VIS", "Visual Studio Code")), "1+3")
        eq("word start preferred", str(FuzzyMatcher.MatchRanges("note", "Denote Notepad")), "8+4")
        eq("initials", str(FuzzyMatcher.MatchRanges("vsc", "Visual Studio Code")), "1+1 8+1 15+1")
        eq("camel initials", str(FuzzyMatcher.MatchRanges("ah", "AutoHotkey")), "1+1 5+1")
        eq("tokens", str(FuzzyMatcher.MatchRanges("code vis", "Visual Studio Code")), "1+3 15+4")
        eq("keyword prefix", str(FuzzyMatcher.MatchRanges("snip quote", "Quote reply")), "1+5")
        eq("subsequence", str(FuzzyMatcher.MatchRanges("ntpd", "Notepad")), "1+1 3+1 5+1 7+1")
        eq("adjacent merged", str(FuzzyMatcher.MatchRanges("vis ual", "Visual")), "1+6")
        eq("pinyin", str(FuzzyMatcher.MatchRanges("jsb", "记事本")), "1+3")
        eq("pinyin partial", str(FuzzyMatcher.MatchRanges("sb", "记事本 Notepad")), "2+2")
        eq("pinyin merged", str(FuzzyMatcher._MergeRanges([[2, 1], [1, 1], [3, 1]])), "1+3")
        eq("no match", str(FuzzyMatcher.MatchRanges("xyz", "Notepad")), "")
        eq("empty needle", str(FuzzyMatcher.MatchRanges("  ", "Notepad")), "")
        eq("word starts", str(FuzzyMatcher._MergeRanges([[1, 1]])), "1+1")
        starts := ""
        for start in FuzzyMatcher.WordStarts("my_file-Name AutoHotkey 记事本")
            starts .= start " "
        eq("word start list", RTrim(starts), "1 4 9 14 18 25")
        segs(text, ranges) {
            out := ""
            for segment in FuzzyMatcher.Segments(text, ranges)
                out .= (segment[2] ? "[" segment[1] "]" : segment[1])
            return out
        }
        eq("segments", segs("Visual Studio Code", [[1, 1], [8, 1], [15, 1]]), "[V]isual [S]tudio [C]ode")
        eq("segments whole", segs("Code", [[1, 4]]), "[Code]")
        eq("segments clipped", segs("Visual St…", [[8, 6]]), "Visual [St…]")
        eq("segments beyond", segs("Visual", [[8, 2]]), "Visual")
        item := ResultItem("Visual Studio Code")
        savedText := SearchWindow.HighlightText
        SearchWindow.HighlightText := "vsc"
        eq("item ranges", str(SearchWindow.HighlightRanges(item)), "1+1 8+1 15+1")
        SearchWindow.HighlightText := "code"
        eq("item ranges updated", str(SearchWindow.HighlightRanges(item)), "15+4")
        SearchWindow.HighlightText := ""
        eq("no query", str(SearchWindow.HighlightRanges(item)), "")
        SearchWindow.HighlightText := savedText
    }

    ; 高亮的字和其它字分段画, 字的位置不能随高亮边界移动 (输入时字距跳动)
    static HighlightSpacing() {
        title := "Weima Wave Station", width := 320, height := 40
        hdc0 := DllCall("GetDC", "Ptr", 0, "Ptr")
        hdc := DllCall("CreateCompatibleDC", "Ptr", hdc0, "Ptr")
        hbm := DllCall("CreateCompatibleBitmap", "Ptr", hdc0, "Int", width, "Int", height, "Ptr")
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdc0)
        oldBitmap := DllCall("SelectObject", "Ptr", hdc, "Ptr", hbm, "Ptr")
        font := SearchWindow._CreateFont("Segoe UI", 13, 400)
        oldFont := DllCall("SelectObject", "Ptr", hdc, "Ptr", font, "Ptr")
        DllCall("SetBkMode", "Ptr", hdc, "Int", 1)
        gdi := SearchWindow._gdi, saved := Map()
        for key in ["TitleColor", "HighlightColor", "SelectedTitleColor", "SelectedHighlightColor"]
            saved[key] := gdi.Has(key) ? gdi[key] : ""
        gdi["TitleColor"] := 0, gdi["HighlightColor"] := 0, gdi["SelectedTitleColor"] := 0, gdi["SelectedHighlightColor"] := 0   ; 同一种颜色: 只比较位置
        ink(query) {                                                        ; 每一列有没有字 (不看颜色)
            rect := Buffer(16), NumPut("Int", 0, "Int", 0, "Int", width, "Int", height, rect)
            DllCall("FillRect", "Ptr", hdc, "Ptr", rect, "Ptr", DllCall("GetStockObject", "Int", 0, "Ptr"))
            ranges := (query = "") ? [] : FuzzyMatcher.MatchRanges(query, title)
            SearchWindow._DrawTitle(hdc, title, ranges, 4, 0, width - 4, height, 0x24 | 0x800, false)
            header := Buffer(40, 0), NumPut("UInt", 40, "Int", width, "Int", -height, "UShort", 1, "UShort", 32, header)
            pixels := Buffer(width * height * 4)
            DllCall("GetDIBits", "Ptr", hdc, "Ptr", hbm, "UInt", 0, "UInt", height, "Ptr", pixels, "Ptr", header, "UInt", 0)
            columns := ""
            Loop width {
                x := A_Index - 1, dark := false
                Loop height {
                    color := NumGet(pixels, ((A_Index - 1) * width + x) * 4, "UInt")
                    if (Min(color & 0xFF, color >> 8 & 0xFF, color >> 16 & 0xFF) < 64) {   ; 只看笔画中心 (抗锯齿的边缘在分段交界处会叠加)
                        dark := true
                        break
                    }
                }
                columns .= dark ? "1" : "0"
            }
            return columns
        }
        plainHighlight := ink("weima wave station")                         ; 整个标题都高亮: 和逐段画的位置比较
        for query in ["", "w", "we", "wei", "weim", "wave", "ws", "zzz"]                ; "" / "ws" / "zzz": 没有高亮的字
            TestRunner.Equal("HighlightSpacing." query, ink(query), plainHighlight)
        for key, value in saved
            if (value != "")
                gdi[key] := value
        DllCall("SelectObject", "Ptr", hdc, "Ptr", oldFont)
        DllCall("SelectObject", "Ptr", hdc, "Ptr", oldBitmap)
        DllCall("DeleteObject", "Ptr", font)
        DllCall("DeleteObject", "Ptr", hbm)
        DllCall("DeleteDC", "Ptr", hdc)
    }

    ; 不是系统尺寸的图标 (紧凑主题 28 px 等) 平滑缩放成正好的尺寸
    ; 每条内置命令都要有图标 (Wine 里没有控制面板等 Shell 项目, 只在 Windows 上检查)
    static BuiltinIcons() {
        if DllCall("GetProcAddress", "Ptr", DllCall("GetModuleHandle", "Str", "ntdll", "Ptr"), "AStr", "wine_get_version", "Ptr")
            return
        missing := ""
        for command in SystemProvider.Commands() {
            hIcon := IconCache._Load(command["Icon"])
            if hIcon
                DllCall("DestroyIcon", "Ptr", hIcon)
            else
                missing .= command["Id"] " "
        }
        TestRunner.Equal("BuiltinIcons.missing", Trim(missing), "")
        hIcon := IconCache._Load("::{A8A91A66-3A7D-4424-8D24-04E180695C7A}")      ; 只写 CLSID 的控制面板项目
        TestRunner.True("BuiltinIcons.control panel item", hIcon)
        if hIcon
            DllCall("DestroyIcon", "Ptr", hIcon)
    }

    static IconScaling() {
        eq := (n, a, e) => TestRunner.Equal("IconScaling." n, a, e)
        eq("source 16", IconCache.SourceSize(16), 16)
        eq("source 20", IconCache.SourceSize(20), 24)
        eq("source 28", IconCache.SourceSize(28), 32)
        eq("source 32", IconCache.SourceSize(32), 32)
        eq("source 40", IconCache.SourceSize(40), 48)
        eq("source 56", IconCache.SourceSize(56), 256)
        saved := IconCache.Size
        for size in [28, 32, 40] {
            IconCache.Size := size
            for spec in ["folder:", "ext:.txt", "res:shell32.dll,-4"] {
                hIcon := IconCache._Load(spec)
                eq(spec " at " size, IconCache.IconWidth(hIcon), size)
                if hIcon
                    DllCall("DestroyIcon", "Ptr", hIcon)
            }
        }
        IconCache.Size := 28                                                ; 没有 Alpha 通道的旧式图标 (按掩码透明)
        hIcon := IconCache.FitSize(DllCall("CopyIcon", "Ptr", DllCall("LoadIcon", "Ptr", 0, "Ptr", 32512, "Ptr"), "Ptr"))
        eq("legacy icon", IconCache.IconWidth(hIcon), 28)
        DllCall("DestroyIcon", "Ptr", hIcon)
        eq("no icon", IconCache.FitSize(0), 0)
        IconCache.Size := saved
    }

    ; 主题缩略图: 每个主题都画得出来, 背景、选中行用的是主题自己的颜色; 外观页点选缩略图就是选主题
    static ThemeGallery() {
        eq := (n, a, e) => TestRunner.Equal("ThemeGallery." n, a, e)
        pixel(hbm, x, y) {
            dc := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
            old := DllCall("SelectObject", "Ptr", dc, "Ptr", hbm, "Ptr")
            value := DllCall("GetPixel", "Ptr", dc, "Int", x, "Int", y, "UInt")
            DllCall("SelectObject", "Ptr", dc, "Ptr", old), DllCall("DeleteDC", "Ptr", dc)
            return value
        }
        for themeName in ThemeManager.Names() {
            hbm := ThemePreview.ForName(themeName, 112, 76)
            TestRunner.True("ThemeGallery.bitmap " themeName, hbm != 0)
            theme := ThemeManager.Resolve(themeName = "System" ? "Light" : themeName)
            if (hbm && themeName != "System") {
                eq("background " themeName, pixel(hbm, 56, 6), Win.ColorToBgr(theme["Background"]))
                eq("selected row " themeName, pixel(hbm, 104, 26), Win.ColorToBgr(theme["SelectedBackground"]))
            }
            if hbm
                DllCall("DeleteObject", "Ptr", hbm)
        }
        hbm := ThemePreview.ForName("System", 112, 76)                      ; 左半边 Light, 右半边 Dark
        eq("system left", pixel(hbm, 20, 6), Win.ColorToBgr(ThemeManager.Defaults()["Background"]))
        eq("system right", pixel(hbm, 100, 6), Win.ColorToBgr(ThemeManager.Resolve("Dark")["Background"]))
        DllCall("DeleteObject", "Ptr", hbm)

        PreferencesWindow.Show(1, -3000, -3000)
        Sleep(400)                                                          ; 打开 300 ms 之后的修改才算 (见 PreferencesWindow.Show)
        PreferencesWindow.SelectPage(PreferencesWindow._PageIndex("Prefs.Page.Appearance"))
        page := PreferencesWindow.Pages[PreferencesWindow._PageIndex("Prefs.Page.Appearance")]
        gallery := ""
        for ctrl in page.Controls
            if (ctrl.Type = "ListView")
                gallery := ctrl
        themes := ThemeManager.Names()
        eq("one item per theme", gallery.GetCount(), themes.Length)
        current := AppSettings.Appearance["Theme"]
        eq("current theme selected", themes[gallery.GetNext()], current)
        target := (current = "Dark") ? "Monokai" : "Dark"
        for index, themeName in themes
            if (themeName = target)
                gallery.Modify(index, "Select Focus")
        Sleep(50)
        eq("dirty after picking", (PreferencesWindow._ready ? "ready " : "not ready ") (PreferencesWindow._dirty ? 1 : 0), "ready 1")
        picked := ""
        for bind in PreferencesWindow.Binds
            if (bind.Path = "Appearance.Theme")
                picked := bind.Read.Call()
        eq("picked theme", picked, target)
        PreferencesWindow._dirty := false
        PreferencesWindow.Close()
    }

    static Themes() {
        eq := (n, a, e) => TestRunner.Equal("Themes." n, a, e)
        ThemeManager.BuiltinDir := A_ScriptDir "\..\Resources\Themes"
        ThemeManager.UserDir := A_Temp "\ALTRunThemeTest"
        try DirDelete(ThemeManager.UserDir, true)
        fullKeys := ThemeManager.Defaults()
        for themeName in ThemeManager.Names() {
            if (themeName = "System")
                continue
            ThemeManager.Load(themeName)
            eq(themeName " loaded", ThemeManager.Resolved, themeName)
            missing := ""
            for key in fullKeys
                if (ThemeManager.Get(key) = "")
                    missing .= key " "
            eq(themeName " complete", missing, "")
            for key in ["Background", "Title", "SelectedBackground", "SelectedTitle"]
                TestRunner.True("Themes." themeName "." key " is RRGGBB", RegExMatch(ThemeManager.Get(key), "^[0-9A-Fa-f]{6}$"))
        }
        eq("builtin count", ThemeManager.Names().Length, 21)
        TestRunner.True("Themes.Ocean is builtin", ThemeManager.IsBuiltin("Ocean"))

        ; 用户主题: 同名覆盖内置主题 ("Base" 写自己 = 在内置那一套上改), 以及 Base 链
        DirCreate(ThemeManager.UserDir)
        FileAppend('{ "Base": "Ocean", "SelectedRadius": 12 }', ThemeManager.UserDir "\Mine.json", "UTF-8")
        FileAppend('{ "Base": "Dark", "Title": "FF0000" }', ThemeManager.UserDir "\Dark.json", "UTF-8")
        ThemeManager.Load("Mine")
        eq("user base color", ThemeManager.Get("Background"), "2E3440")
        eq("user override", ThemeManager.Get("SelectedRadius"), 12)
        TestRunner.True("Themes.user listed", ThemeManager.Names().Length = 22 && !ThemeManager.IsBuiltin("Mine"))
        ThemeManager.Load("Dark")
        eq("user overrides builtin", ThemeManager.Get("Title"), "FF0000")
        eq("user override keeps builtin", ThemeManager.Get("Background"), "1E1F22")
        DirDelete(ThemeManager.UserDir, true)

        ThemeManager.Load("System")
        TestRunner.True("Themes.System resolves", ThemeManager.Resolved = "Light" || ThemeManager.Resolved = "Dark")
        ThemeManager.Load("No Such Theme")
        eq("missing falls back", ThemeManager.Resolved, "Light")
        ThemeManager.Load("Light")
    }

    ; 命令很多时: 只为前 MaxResults 条生成结果 (算上学习加分), 继续输入时只在上次的结果里找
    static CommandSearchScale() {
        eq := (n, a, e) => TestRunner.Equal("CommandSearchScale." n, a, e)
        saved := AppSettings.Data["CustomCommands"]
        savedPicks := Knowledge.Picks, savedQueryPicks := Knowledge.QueryPicks, savedHistory := Knowledge.History
        Knowledge.Picks := Map(), Knowledge.QueryPicks := Map(), Knowledge.History := []
        commands := []
        Loop 120
            commands.Push(Map("Title", "Note " A_Index, "Type", "Folder", "Target", "C:\Work\Folder " A_Index, "Arguments", "", "Keyword", ""))
        rare := Map("Title", "Annual Report", "Type", "File", "Target", "C:\Work\annual.pdf", "Arguments", "", "Keyword", "")
        commands.Push(rare)
        AppSettings.Data["CustomCommands"] := commands
        titles(text) {
            list := ""
            for item in ProviderRegistry.SortByScore(CustomCommandProvider.Search(SearchQuery(text)))
                list .= item.Title "|"
            return RTrim(list, "|")
        }
        results := CustomCommandProvider.Search(SearchQuery("n"))
        eq("capped at MaxResults", results.Length, ProviderRegistry.MaxResults)
        found := false
        for item in results
            found := found || (item.Title = "Annual Report")
        eq("weak match not in top without learning", found, false)
        ; 常选的命令: 学习加分让它进入前 MaxResults 条
        Knowledge.Record("n", CustomCommandProvider._Uid(rare))
        Knowledge.Record("n", CustomCommandProvider._Uid(rare))
        found := false
        for item in CustomCommandProvider.Search(SearchQuery("n"))
            found := found || (item.Title = "Annual Report")
        eq("learned pick kept", found, true)

        ; 继续输入时只在上一次的结果里找, 结果和从头找一样
        CustomCommandProvider._ResetNarrowing()
        full := titles("note 11")
        titles("not"), titles("note"), titles("note ")
        eq("narrowed equals full search", titles("note 11"), full)
        eq("narrowed result", full, "Note 11|Note 110|Note 111|Note 112|Note 113|Note 114|Note 115|Note 116|Note 117|Note 118|Note 119")
        ; 命令被修改后重新从全部命令里找
        titles("rep")
        commands[5]["Title"] := "Report Folder"
        CustomCommandProvider._ResetNarrowing()
        eq("edited command found", InStr(titles("repo"), "Report Folder") > 0, true)
        ; 增删命令 (数量变化) 也会重新找
        titles("zzz")
        commands.Push(Map("Title", "zzz top", "Type", "Folder", "Target", "C:\z", "Arguments", "", "Keyword", ""))
        eq("added command found", titles("zzz t"), "zzz top")

        AppSettings.Data["CustomCommands"] := saved
        Knowledge.Picks := savedPicks, Knowledge.QueryPicks := savedQueryPicks, Knowledge.History := savedHistory
        if (Knowledge._saveTimer != "")
            SetTimer(Knowledge._saveTimer, 0)                               ; Record() 排好的写盘不要执行
        CustomCommandProvider._ResetNarrowing()

        ; 网络位置的图标不读磁盘: UNC 路径用文件夹 / 扩展名的通用图标
        ; 缓存上限: 只留最近用到的 (句柄为 0 的假图标, 不会调用 DestroyIcon)
        savedIcons := IconCache._icons, savedUsed := IconCache._used
        IconCache._icons := Map("a", 0, "b", 0, "c", 0, "d", 0), IconCache._used := Map("a", 5, "b", 1, "c", 9, "d", 3)
        IconCache.Trim(2)
        keys := ""
        for key in IconCache._icons
            keys .= key
        eq("trim keeps newest", keys, "ac")
        eq("trim used map", IconCache._used.Count, 2)
        IconCache._icons := savedIcons, IconCache._used := savedUsed
        eq("remote unc", IconCache.IsRemote("\\server\share\PT1931"), true)
        eq("remote local", IconCache.IsRemote("C:\Windows"), false)
        eq("remote folder key", IconCache._CacheKey("\\server\share\PT1931 - 24 NIR"), "folder:")
        eq("remote folder slash", IconCache._CacheKey("\\server\share\PT1931\"), "folder:")
        eq("remote file key", IconCache._CacheKey("\\server\share\Report.PDF"), "ext:.pdf")
        eq("remote exe key", IconCache._CacheKey("\\server\share\tool.exe"), "ext:.exe")
        ; 名字里带点的文件夹不是 "扩展名" (以前变成空白图标)
        eq("remote dotted folder", IconCache._CacheKey("\\server\TENDER PROPOSAL\2025\26. 18 New Industrial Road (EA)"), "folder:")
        eq("remote dotted folder slash", IconCache._CacheKey("\\server\share\10.PT2310-29NIR\"), "folder:")
        eq("local dotted folder", IconCache._CacheKey("C:\Projects\26. 18 New Industrial Road (EA)"), "c:\projects\26. 18 new industrial road (ea)")
        eq("extension", IconCache._Extension("C:\a\Report.PDF"), "PDF")
        eq("extension with spaces", IconCache._Extension("C:\a\26. 18 New Road"), "")
        eq("extension appref-ms", IconCache._Extension("C:\a\App.appref-ms"), "appref-ms")
        eq("extension 7z", IconCache._Extension("C:\a\backup.7z"), "7z")
        eq("extension sldprt", IconCache._Extension("C:\a\Beam.SLDPRT"), "SLDPRT")
        eq("extension digits only", IconCache._Extension("C:\a\Design.2019"), "")
        eq("folder command icon remote", CustomCommandProvider._FolderIcon("\\server\share\26. 18 New Industrial Road (EA)"), "folder:")
        ; 普通文件夹用通用图标, 带 desktop.ini (自定义图标) 的和磁盘根目录用自己的
        iconRoot := A_Temp "\ALTRun-folder-icon-test"
        try DirDelete(iconRoot, true)

        ; 启动后预先算搜索 Key: 坏掉的条目跳过, 不影响后面的
        saved := AppSettings.CustomCommands
        AppSettings.Data["CustomCommands"] := ["not a map", Map("Title", "No target"), Map("Title", "Warm Test Folder", "Type", "Folder", "Target", "C:\"), Map("Title", "记事本", "Target", "notepad.exe")]
        warmed := true
        try CustomCommandProvider.Warm()
        catch
            warmed := false
        eq("warm skips bad entries", warmed, true)
        eq("warm then search", CustomCommandProvider.Search(SearchQuery("jsb")).Length, 1)
        AppSettings.Data["CustomCommands"] := saved
        CustomCommandProvider._ResetNarrowing()
        DirCreate(iconRoot "\Plain 26. 18 Road")
        DirCreate(iconRoot "\Custom")
        FileAppend("[.ShellClassInfo]`n", iconRoot "\Custom\desktop.ini")
        eq("folder command icon local", CustomCommandProvider._FolderIcon(iconRoot "\Plain 26. 18 Road"), "folder:")
        eq("folder icon drive root", CustomCommandProvider._FolderIcon("C:\"), "C:\")
        ; 搜索时不访问磁盘: 先是通用图标, 后台看过 desktop.ini 后才是它自己的
        eq("folder icon with desktop.ini, not probed yet", CustomCommandProvider._FolderIcon(iconRoot "\Custom"), "folder:")
        IconCache._ProbeQueued()
        eq("folder icon with desktop.ini, after probe", CustomCommandProvider._FolderIcon(iconRoot "\Custom"), iconRoot "\Custom")
        ; 处理队列时被另一轮 (后台定时器) 打断、剩下的已经处理掉: 跳过, 不报 "Item has no value"
        savedProbe := IconCache.GetOwnPropDesc("_ProbeFolder"), savedProbes := IconCache._folderProbes
        IconCache._folderProbes := Map("X:\probe-a", true, "X:\probe-b", true, "X:\probe-c", true)
        IconCache.DefineProp("_ProbeFolder", {Call: (this, folder) => (IconCache._folderProbes.Clear(), "folder:")})
        interrupted := true
        try IconCache._ProbeQueued()
        catch
            interrupted := false
        IconCache.DefineProp("_ProbeFolder", savedProbe), IconCache._folderProbes := savedProbes
        eq("probe queue emptied by another run", interrupted, true)
        eq("folder icon probe now (warm-up)", IconCache.FolderIcon(iconRoot "\Plain 26. 18 Road", true), "folder:")
        try DirDelete(iconRoot, true)
        eq("remote load spec", IconCache._LoadSpec("\\server\share\Report.pdf", "ext:.pdf"), "ext:.pdf")
        eq("local exe key", IconCache._CacheKey("C:\Tools\app.exe"), "c:\tools\app.exe")
        eq("local load spec", IconCache._LoadSpec("C:\Tools\app.exe", "c:\tools\app.exe"), "C:\Tools\app.exe")
    }

    static CheckTargets() {
        eq := (n, a, e) => TestRunner.Equal("CheckTargets." n, a, e)
        dir := A_Temp "\ALTRunTest_CheckTargets"
        try DirDelete(dir, true)
        DirCreate(dir "\PT1931 - 24 NIR")
        FileAppend("", dir "\summary.pdf")
        check(type, target, roots := "") => CustomCommandProvider.CheckTarget(Map("Title", "t", "Type", type, "Target", target), roots)
        eq("folder ok", check("Folder", dir "\PT1931 - 24 NIR"), "OK")
        eq("folder renamed", check("Folder", dir "\PT1931 - 24 NIR (old)"), "Missing")
        eq("folder trailing slash", check("Folder", dir "\PT1931 - 24 NIR\"), "OK")
        eq("file ok", check("File", dir "\summary.pdf"), "OK")
        eq("quoted file", check("File", '"' dir '\summary.pdf"'), "OK")
        eq("file moved", check("File", dir "\summary-old.pdf"), "Missing")
        eq("file type pointing at folder", check("File", dir), "OK")
        eq("folder type pointing at file", check("Folder", dir "\summary.pdf"), "Missing")
        eq("runas prefix", check("File", "*RunAs " dir "\summary.pdf"), "OK")
        eq("builtin variable", check("Folder", "A_WinDir"), "OK")
        eq("program in PATH", check("Command", "cmd.exe"), "OK")
        eq("program without extension", check("Command", "cmd"), "OK")
        eq("unknown program", check("Command", "altrun-no-such-program-12345.exe"), "Missing")
        eq("empty target", check("File", ""), "Missing")
        eq("url skipped", check("Url", "https://github.com"), "Skipped")
        eq("shell location skipped", check("Folder", "shell:Downloads"), "Skipped")
        eq("clsid skipped", check("Folder", "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"), "Skipped")
        ; 断开的网络盘: 驱动器能否访问记在 roots 里, 不再逐条等待
        offline := Map("Q:", false)
        eq("drive offline", check("Folder", "Q:\DESIGN PROJECTS\PT1931", offline), "Unavailable")
        eq("unc offline", check("File", "\\server\share\a.pdf", Map("\\server\share", false)), "Unavailable")
        roots := Map()
        check("Folder", dir, roots)
        SplitPath(dir, , , , , &drive)
        eq("root remembered", roots.Has(drive) && roots[drive], true)
        try DirDelete(dir, true)
    }

    static CommandTargets() {
        eq := (n, a, e) => TestRunner.Equal("CommandTargets." n, a, e)
        saved := AppSettings.Data["CustomCommands"]
        AppSettings.Data["CustomCommands"] := [
            Map("Title", "CKR, EA, JIB", "Type", "Folder", "Target", "Q:\DESIGN PROJECTS\Design-2019\PT1931 - 24 NIR", "Arguments", "", "Keyword", ""),
            Map("Title", "NIR Report", "Type", "File", "Target", "Q:\Docs\summary.pdf", "Arguments", "", "Keyword", ""),
            Map("Title", "Check IP", "Type", "Command", "Target", "cmd.exe", "Arguments", "/k ipconfig", "Keyword", ""),
            Map("Title", "Drive Q", "Type", "Folder", "Target", "Q:\", "Arguments", "", "Keyword", "")]
        titles(text) {
            list := ""
            for item in ProviderRegistry.SortByScore(CustomCommandProvider.Search(SearchQuery(text)))
                list .= item.Title "|"
            return RTrim(list, "|")
        }
        eq("folder name", titles("1931"), "CKR, EA, JIB")
        eq("title first", titles("nir"), "NIR Report|CKR, EA, JIB")
        eq("file name without extension", titles("summary"), "NIR Report")
        eq("extension not searched", titles("pdf"), "")
        eq("command target not searched", titles("cmd"), "")
        eq("drive root not a name", titles("q:"), "")

        ; {query}: "关键字 文字" 换掉目标 / 参数里的 {query} (网址里编码), 只显示这条
        AppSettings.Data["CustomCommands"].Push(
            Map("Title", "Jira", "Type", "Url", "Target", "https://jira.example.com/browse/{query}", "Arguments", "", "Keyword", "jira"),
            Map("Title", "Ping Host", "Type", "Command", "Target", "cmd.exe", "Arguments", "/k ping {query}", "Keyword", "ping"),
            Map("Title", "Notes Folder", "Type", "Folder", "Target", "D:\Notes\{query}", "Arguments", "", "Keyword", ""))
        CustomCommandProvider._ResetNarrowing()
        items := CustomCommandProvider.Search(SearchQuery("jira ABC 12"))
        eq("query: only that command", items.Length, 1)
        eq("query: url encoded", items[1].Title "|" items[1].Arg "|" items[1].Exclusive, "Jira: ABC 12|https://jira.example.com/browse/ABC%2012|1")
        items := CustomCommandProvider.Search(SearchQuery("ping 10.0.0.1"))
        eq("query: arguments", items[1].Arg "|" items[1].Arguments, "cmd.exe|/k ping 10.0.0.1")
        items := CustomCommandProvider.Search(SearchQuery("ping 10.0.0.12"))   ; 继续输入: 不用上一次的结果缩小范围
        eq("query: keep typing", items.Length "|" items[1].Arguments, "1|/k ping 10.0.0.12")
        item := ""
        for found in CustomCommandProvider.Search(SearchQuery("jir"))
            if (found.Title = "Jira")
                item := found
        eq("query: not typed yet -> complete keyword", IsObject(item) ? item.Valid "|" item.AutoComplete : "none", "0|jira ")
        for found in CustomCommandProvider.Search(SearchQuery("notes"))
            item := found
        eq("query: no keyword -> empty", item.Valid "|" item.Arg, "1|D:\Notes\")
        eq("query: path not checked", CustomCommandProvider.CheckTarget(AppSettings.Data["CustomCommands"][7]), "Skipped")
        AppSettings.Data["CustomCommands"] := saved
        CustomCommandProvider._ResetNarrowing()
    }

    ; 所有 Edit 控件都要写明行数 (r1 / r8 ...): 不写时长文字会让 AHK 自动变成多行并加高, 盖住下面的控件
    static EditRows() {
        missing := ""
        for folder in ["Src", "Lib"] {
            Loop Files, A_ScriptDir "\..\" folder "\*.ahk", "R" {
                Loop Parse, FileRead(A_LoopFileFullPath, "UTF-8"), "`n", "`r" {
                    if (RegExMatch(A_LoopField, 'AddEdit\(|_Add\("Edit"') && !RegExMatch(A_LoopField, '[" ]r(\d|"\s)'))
                        missing .= A_LoopFileName ":" A_Index " "
                }
            }
        }
        TestRunner.Equal("EditRows.all edits set rows", missing, "")
    }

    static HiddenApps() {
        eq := (n, a, e) => TestRunner.Equal("HiddenApps." n, a, e)
        saved := ProviderRegistry.Providers
        ProviderRegistry.Providers := [ApplicationProvider]
        options := AppSettings.Feature("Applications")
        options["Hidden"] := []
        savedFile := AppSettings.File
        AppSettings.File := A_Temp "\ALTRunTest.json"                     ; 删除时会保存设置, 写到临时文件
        ApplicationProvider.Apps := [
            Map("Title", "Notepad", "Target", "C:\Start Menu\Notepad.lnk", "Detail", "", "Search", ""),
            Map("Title", "Notepad++", "Target", "C:\Start Menu\Notepad++.lnk", "Detail", "", "Search", "")]
        titles() {
            list := ""
            for item in ApplicationProvider.Search(SearchQuery("notepad"))
                list .= item.Title "|"
            return RTrim(list, "|")
        }
        eq("before", titles(), "Notepad|Notepad++")
        item := ApplicationProvider.Search(SearchQuery("notepad++"))[1]
        item.Provider := "Applications"
        TestRunner.True("HiddenApps.can delete", ActionCatalog.CanDelete(item))
        TestRunner.True("HiddenApps.prompt", InStr(ActionCatalog.DeletePrompt(item), "Notepad++") && InStr(ActionCatalog.DeletePrompt(item), "uninstalled"))
        ActionCatalog.DeleteItem(item)
        eq("hidden", titles(), "Notepad")
        eq("saved in settings", options["Hidden"][1], "C:\Start Menu\Notepad++.lnk")
        ApplicationProvider.Apps := [
            Map("Title", "Notepad", "Target", "C:\Start Menu\Notepad.lnk", "Detail", "", "Search", ""),
            Map("Title", "Notepad++", "Target", "c:\start menu\NOTEPAD++.lnk", "Detail", "", "Search", "")]
        eq("stays hidden after rebuild", titles(), "Notepad")
        options["Hidden"] := []
        eq("restored", titles(), "Notepad|Notepad++")
        ApplicationProvider.Apps := []
        ProviderRegistry.Providers := saved
        AppSettings.File := savedFile
        try FileDelete(A_Temp "\ALTRunTest.json")
    }

    static HiddenSystemCommands() {
        eq := (n, a, e) => TestRunner.Equal("HiddenSystemCommands." n, a, e)
        saved := ProviderRegistry.Providers
        ProviderRegistry.Providers := [SystemProvider]
        options := AppSettings.Feature("System")
        savedHidden := options["Hidden"], options["Hidden"] := []
        savedFile := AppSettings.File
        AppSettings.File := A_Temp "\ALTRunTest.json"                     ; 删除时会保存设置, 写到临时文件
        titles() {
            list := ""
            for item in SystemProvider.Search(SearchQuery("devices and printers"))
                list .= item.Title "|"
            return list
        }
        TestRunner.True("HiddenSystemCommands.found", InStr(titles(), I18n.T("Tool.Printers")))
        item := SystemProvider.Search(SearchQuery("devices and printers"))[1]
        item.Provider := "System"
        TestRunner.True("HiddenSystemCommands.can delete", ActionCatalog.CanDelete(item))
        TestRunner.True("HiddenSystemCommands.prompt", InStr(ActionCatalog.DeletePrompt(item), I18n.T("Tool.Printers")))
        ActionCatalog.DeleteItem(item)
        eq("saved in settings", options["Hidden"].Length ? options["Hidden"][1] : "", "Printers")
        TestRunner.True("HiddenSystemCommands.hidden", !InStr(titles(), I18n.T("Tool.Printers")))
        SystemProvider.Hide("Printers")
        eq("added once", options["Hidden"].Length, 1)
        options["Hidden"] := [" printers "]                                 ; 设置里手写的: 不分大小写, 去掉空格
        TestRunner.True("HiddenSystemCommands.case-insensitive", !InStr(titles(), I18n.T("Tool.Printers")))
        options["Hidden"] := []
        TestRunner.True("HiddenSystemCommands.restored", InStr(titles(), I18n.T("Tool.Printers")))
        eq("default", AppSettings.Defaults()["Features"]["System"]["Hidden"].Length, 0)
        options["Hidden"] := savedHidden
        ProviderRegistry.Providers := saved
        AppSettings.File := savedFile
        try FileDelete(A_Temp "\ALTRunTest.json")
    }

    static ListFilter() {
        eq := (n, a, e) => TestRunner.Equal("ListFilter." n, a, e)
        cells := ["Devices and Printers", "File", "C:\Windows\System32\control.exe printers", "dp"]
        eq("empty", PreferencesWindow.FilterMatch("", cells), true)
        eq("spaces only", PreferencesWindow.FilterMatch("   ", cells), true)
        eq("title", PreferencesWindow.FilterMatch("printer", cells), true)
        eq("case-insensitive", PreferencesWindow.FilterMatch("CONTROL", cells), true)
        eq("all words, any column", PreferencesWindow.FilterMatch("devices  system32", cells), true)
        eq("missing word", PreferencesWindow.FilterMatch("devices notepad", cells), false)
        eq("not across columns", PreferencesWindow.FilterMatch("printersfile", cells), false)

        ; 自定义命令页: 筛选框里输入后只显示匹配的行, 筛选不算修改设置
        saved := AppSettings.Data["CustomCommands"]
        AppSettings.Data["CustomCommands"] := [
            Map("Title", "Devices and Printers", "Type", "Command", "Target", "control", "Arguments", "printers", "Keyword", ""),
            Map("Title", "Notepad", "Type", "File", "Target", "C:\Windows\notepad.exe", "Arguments", "", "Keyword", "np"),
            Map("Title", "Printer queue", "Type", "Command", "Target", "control", "Arguments", "printers", "Keyword", "")]
        PreferencesWindow.Show(1, -3000, -3000)
        page := PreferencesWindow.Pages[PreferencesWindow._PageIndex("Prefs.Page.Commands")]
        listView := "", filterBox := ""
        for ctrl in page.Controls {
            if (ctrl.Type = "ListView")
                listView := ctrl
            else if (ctrl.Type = "Edit")
                filterBox := ctrl
        }
        eq("all rows", listView.GetCount(), 3)
        ControlSetText("print", filterBox)
        Sleep(50)
        eq("filtered rows", listView.GetCount(), 2)
        eq("filtered second row", listView.GetText(2, 1), "Printer queue")
        ControlSetText("NP", filterBox)
        Sleep(50)
        eq("keyword column", listView.GetCount() "|" listView.GetText(1, 1), "1|Notepad")
        eq("not dirty", PreferencesWindow._dirty ? 1 : 0, 0)
        ControlSetText("", filterBox)
        Sleep(50)
        eq("cleared", listView.GetCount(), 3)
        PreferencesWindow.Close()
        AppSettings.Data["CustomCommands"] := saved
    }

    ; 默认设置里用 A_ 变量写的文件夹都要能解析成真实路径 (否则那个文件夹从来不会被索引)
    static DefaultFolders() {
        defaults := AppSettings.Defaults()
        unresolved := ""
        for folder in defaults["Features"]["Applications"]["Folders"]
            if (InStr(Path.Resolve(folder), "A_") = 1)
                unresolved .= folder " "
        for folder in defaults["Features"]["FileSearch"]["ScopeFolders"]
            if (InStr(Path.Resolve(folder), "A_") = 1 || InStr(Path.Resolve(folder), "%"))
                unresolved .= folder " "
        TestRunner.Equal("DefaultFolders.all resolve", unresolved, "")
    }

    static FileSearchModes() {
        eq := (n, a, e) => TestRunner.Equal("FileSearchModes." n, a, e)
        options := AppSettings.Feature("FileSearch")
        saved := [options["InDefaultResults"], options["UseEverything"]]
        options["UseEverything"] := 0                                       ; 测试用内置索引
        FileIndex.Paths := ["C:\Windows\Fonts\NIRMALA.TTF", "C:\Docs\Nir Report.pdf"]
        FileIndex.Names := ["nirmala.ttf", "nir report.pdf"]
        FileIndex.Folders := [0, 0]
        FileIndex._lastNeedle := "", FileIndex._lastMatches := ""
        eq("default off", AppSettings.Defaults()["Features"]["FileSearch"]["InDefaultResults"], 0)
        eq("space prefix on", AppSettings.Defaults()["Features"]["FileSearch"]["SpacePrefix"], 1)
        options["InDefaultResults"] := 0
        eq("normal search has no files", FileSearchProvider.Search(SearchQuery("nir")).Length, 0)
        files := ProviderRegistry.SearchFiles("nir")
        TestRunner.True("FileSearchModes.file mode finds files", files.Length >= 2 && files[1].Provider = "FileSearch")
        eq("file mode empty", ProviderRegistry.SearchFiles("  ").Length, 0)
        TestRunner.True("FileSearchModes.quote still works", FileSearchProvider.Search(SearchQuery("'nir")).Length >= 2)
        options["InDefaultResults"] := saved[1], options["UseEverything"] := saved[2]
        FileIndex.Paths := [], FileIndex.Names := [], FileIndex.Folders := []
    }

    ; folder bk 只搜文件夹; 关键字后面还没输入空格时不挡住其它结果; 结果按名称匹配程度排序
    static FolderSearch() {
        eq := (n, a, e) => TestRunner.Equal("FolderSearch." n, a, e)
        options := AppSettings.Feature("FileSearch")
        savedEverything := options["UseEverything"]
        options["UseEverything"] := 0                                       ; 测试用内置索引
        FileIndex.Paths := ["C:\Docs\notebk.pdf", "C:\Projects\BK Tower", "C:\Docs\BK drawing.dwg", "C:\Projects\Old BK"]
        FileIndex.Names := ["notebk.pdf", "bk tower", "bk drawing.dwg", "old bk"]
        FileIndex.Folders := [0, 1, 0, 1]
        FileIndex._lastNeedle := "", FileIndex._lastMatches := ""
        titles(results) {
            list := ""
            for item in results
                if (item.Kind = "file" || item.Kind = "folder")             ; 不算最后的 "用 Everything / Windows 搜索"
                    list .= item.Title "|"
            return RTrim(list, "|")
        }
        eq("default keyword", AppSettings.Defaults()["Features"]["FileSearch"]["FolderKeywords"][1], "folder")
        eq("folders only", titles(FileSearchProvider.Search(SearchQuery("folder bk"))), "BK Tower|Old BK")
        eq("folders exclusive", FileSearchProvider.Search(SearchQuery("folder bk"))[1].Exclusive, true)
        eq("files keyword", titles(FileSearchProvider.Search(SearchQuery("open bk"))), "BK Tower|BK drawing.dwg|Old BK|notebk.pdf")
        eq("file mode folder keyword", titles(ProviderRegistry.SearchFiles("folder bk")), "BK Tower|Old BK")
        eq("file mode plain", titles(ProviderRegistry.SearchFiles("bk")), "BK Tower|BK drawing.dwg|Old BK|notebk.pdf")
        ; 只输入关键字 (还没有空格): 提示排在后面, 不挡住名字带 folder 的命令
        hint := FileSearchProvider.Search(SearchQuery("folder"))
        eq("keyword alone hint", hint.Length, 1)
        eq("keyword alone not exclusive", hint[1].Exclusive, false)
        eq("keyword alone low score", hint[1].Score < 10, true)
        eq("open alone not exclusive", FileSearchProvider.Search(SearchQuery("open"))[1].Exclusive, false)
        eq("keyword with space exclusive", FileSearchProvider.Search(SearchQuery("folder "))[1].Exclusive, true)
        eq("quote alone exclusive", FileSearchProvider.Search(SearchQuery("'"))[1].Exclusive, true)
        ; Everything 的结果按修改时间来: 名称开头 / 单词开头的排前面, 同分时文件夹在前
        found := [{Path: "C:\a\notebk.pdf", IsFolder: 0, Score: FileIndex.ScoreName("bk", "notebk.pdf")}
                , {Path: "C:\a\BK x.dwg", IsFolder: 0, Score: FileIndex.ScoreName("bk", "bk x.dwg")}
                , {Path: "C:\a\BK x.pdf", IsFolder: 1, Score: FileIndex.ScoreName("bk", "bk x.pdf")}]
        best := FileSearchProvider._Best(found, 3)
        eq("best order", best[1].Path "|" best[2].Path "|" best[3].Path, "C:\a\BK x.pdf|C:\a\BK x.dwg|C:\a\notebk.pdf")
        eq("everything fetches more", FileSearchProvider.EverythingFetch >= 300, true)
        eq("score needle folder:", FileSearchProvider.ScoreNeedle("folder:bk"), "bk")
        eq("score needle ext:", FileSearchProvider.ScoreNeedle("ext:pdf report"), "pdf report")
        eq("score needle plain", FileSearchProvider.ScoreNeedle("bk tower"), "bk tower")
        options["UseEverything"] := savedEverything
        FileIndex.Paths := [], FileIndex.Names := [], FileIndex.Folders := []
        FileIndex._lastNeedle := "", FileIndex._lastMatches := ""
    }

    ; 输入 ? 显示速查表; 空搜索框里轮换的使用提示
    static HelpAndTips() {
        eq := (n, a, e) => TestRunner.Equal("HelpAndTips." n, a, e)
        eq("tips default on", AppSettings.Defaults()["General"]["ShowTips"], 1)
        eq("help default on", AppSettings.Defaults()["Features"]["Help"]["Enabled"], 1)
        items := HelpProvider.Items()
        all := HelpProvider.Search(SearchQuery("?"))
        eq("all entries", all.Length, items.Length)
        eq("exclusive", all[1].Exclusive, true)
        eq("no help without ?", HelpProvider.Search(SearchQuery("folder")).Length, 0)
        ids := ""
        for item in items
            ids .= item.Id " "
        eq("folder entry listed", InStr(ids, "Folders ") > 0, true)
        ; 过滤: ? 后面的文字
        keys := ""
        for item in HelpProvider.Search(SearchQuery("? f3"))
            keys .= item.Title "|"
        eq("filter", keys, "F3|")
        ; 占位符换成当前设置里的关键字
        folderItem := ""
        for item in items
            if (item.Id = "Folders")
                folderItem := item
        eq("configured keyword", folderItem.Key, AppSettings.Feature("FileSearch")["FolderKeywords"][1] " bk")
        eq("wiki page", folderItem.Url, "https://github.com/zhugecaomao/ALTRun/wiki/File-Search")
        ; 关掉的功能不显示
        clip := AppSettings.Feature("Clipboard"), savedClip := clip["Enabled"]
        clip["Enabled"] := 0
        ids := ""
        for item in HelpProvider.Items()
            ids .= item.Id " "
        eq("disabled feature hidden", InStr(ids, "Clipboard "), 0)
        clip["Enabled"] := savedClip
        ; 提示按顺序轮换, 每条都会轮到
        seen := Map()
        Loop items.Length
            seen[HelpProvider.NextTip()] := true
        eq("tips rotate", seen.Count, items.Length)
        tip := HelpProvider.NextTip(), matched := false
        for item in items
            matched := matched || (tip = I18n.T("Help.TipFormat", item.Key, item.Text))
        eq("tip text", matched, true)
    }

    ; 偏好设置: 应用 之后用 "-Preferences 页码 x y" 重新打开; 帮助 打开每页对应的 Wiki
    static PreferencesButtons() {
        eq := (n, a, e) => TestRunner.Equal("PreferencesButtons." n, a, e)
        args := PreferencesWindow.ParseArgs(["-Preferences", "5", "120", "-8"])
        eq("page", args.Page, 5)
        eq("x", args.X, 120)
        eq("y (negative, second monitor)", args.Y, -8)
        args := PreferencesWindow.ParseArgs(["-Preferences", "3"])
        eq("page only", args.Page "|" args.X "|" args.Y, "3||")
        args := PreferencesWindow.ParseArgs(["-Preferences"])
        eq("default page", args.Page, 1)
        eq("wiki file search", PreferencesWindow.WikiPage("Prefs.Page.FileSearch"), "File-Search")
        eq("wiki commands", PreferencesWindow.WikiPage("Prefs.Page.Commands"), "Commands-and-Snippets")
        eq("wiki appearance", PreferencesWindow.WikiPage("Prefs.Page.Appearance"), "Themes")
        eq("wiki default", PreferencesWindow.WikiPage("Prefs.Page.General"), "Configuration")
        for key in ["Prefs.OK", "Prefs.Cancel", "Prefs.Apply", "Prefs.Help", "Prefs.DiscardChanges"]
            eq("text " key, I18n.T(key) != key, true)
    }

    ; 搜索窗口的位置: 默认居中、离顶部 20%; 记住的位置按屏幕里的千分比换算, 换一块屏幕也放在对应的地方
    ; 每一页的控件都在底部按钮上面 (中英文都检查, 文字长短不同)
    static PreferencesFit() {
        savedLang := I18n.Lang
        for language in I18n.Languages() {                                   ; 按代码排序, ja 在 zh-TW 前 (Wine 没有这两种字体, 先用过 JhengHei 后量出的日文宽度偏小)
            lang := language[1]
            I18n.Init(lang)
            PreferencesWindow.Show(1, -3000, -3000)
            limit := PreferencesWindow.ButtonY - 6
            rightLimit := PreferencesWindow.ContentX + PreferencesWindow.ContentW + 3   ; 分节下面的细线 (SS_ETCHEDHORZ) 会宽 2 像素
            for index, page in PreferencesWindow.Pages {
                bottom := 0, right := 0
                for ctrl in page.Controls {
                    ctrl.GetPos(&x, &y, &w, &h)
                    bottom := Max(bottom, y + h), right := Max(right, x + w)
                }
                TestRunner.True("PreferencesFit." lang " page " index " (bottom " bottom ", limit " limit ")", bottom <= limit)
                TestRunner.True("PreferencesFit." lang " page " index " (right " right ", limit " rightLimit ")", right <= rightLimit)
            }
            PreferencesWindow.Close()
            if (lang = "en" || lang = "zh-CN") {                            ; 中英文的左列宽度是调好的, 不需要加宽
                widened := ""
                for key in PreferencesWindow._labelWidths
                    if (SubStr(key, 1, StrLen(lang) + 1) = lang " ")
                        widened .= key "; "
                TestRunner.Equal("PreferencesFit." lang " no widened pages", widened, "")
            }
        }
        I18n.Init(savedLang)
    }

    ; 新用户的默认示例: 自定义命令每种类型都有, 片段只用支持的占位符; 默认热键; 窗口列表
    static DefaultExamples() {
        eq := (n, a, e) => TestRunner.Equal("DefaultExamples." n, a, e)
        defaults := AppSettings.Defaults()
        types := Map()
        for command in defaults["CustomCommands"]
            types[command["Type"]] := true
        for commandType in ["File", "Folder", "Command", "Url"]
            eq("command type " commandType, types.Has(commandType), true)
        keywords := Map()
        for snippet in defaults["Snippets"] {
            TestRunner.True("DefaultExamples.snippet keyword " snippet["Name"], snippet["Keyword"] != "" && !keywords.Has(snippet["Keyword"]))
            keywords[snippet["Keyword"]] := true
            unknown := RegExReplace(snippet["Text"], "\{(date|time|datetime|clipboard|cursor)\}")
            eq("snippet placeholders " snippet["Name"], RegExMatch(unknown, "\{\w+\}"), 0)
        }
        TestRunner.True("DefaultExamples.snippets for email", keywords.Has("sig") && keywords.Has("thx"))
        eq("second hotkey", defaults["General"]["SecondaryHotkey"], "!r")
        for entry in defaults["Hotkeys"]                                    ; F1 ~ F4 内置在搜索窗口里, 不在自定义热键中
            TestRunner.True("DefaultExamples.no F-key hotkey " entry["Key"], !RegExMatch(entry["Key"], "i)^F[1-4]$"))
        help := ""
        for item in HelpProvider.Items()
            help .= item.Id " "
        eq("help lists F1 and F4", (InStr(help, "About ") && InStr(help, "EditJson ")) ? 1 : 0, 1)
        eq("command target shows arguments", PreferencesWindow._Cell(Map("Target", "cmd.exe", "Arguments", "/k ipconfig /all"), "Target"), "cmd.exe /k ipconfig /all")
        eq("url type name", I18n.Strings["Prefs.Type.Url"], "Web address / link")
        ids := Map()
        for command in SystemProvider.Commands()
            ids[command["Id"]] := true
        eq("about command exists", ids.Has("About"), true)
        eq("auto switch exclude default", defaults["Extensions"]["QuickSwitch"]["AutoSwitchExclude"], "")

        eq("split windows", PreferencesWindow.JoinLines(PreferencesWindow.SplitWindows("ahk_class A,  ahk_exe b.exe ,,")), "ahk_class A`r`nahk_exe b.exe")
        eq("join windows", PreferencesWindow.JoinWindows(["ahk_class #32770"], ["ahk_class Qt5QWindowIcon", " ", "ahk_class #32770"]), "ahk_class #32770, ahk_class Qt5QWindowIcon")
    }

    ; 英文原文都不为空; 每个语言文件 (Resources\Lang\*.json) 每个键都有翻译、没有多余的键、占位符 {1} {2} ... 和英文一样;
    ; 读语言文件: 代码的写法、_fallback、_font、文件不在或缺键时退回
    static I18nLanguages() {
        eq := (n, a, e) => TestRunner.Equal("I18nLanguages." n, a, e)
        for key, text in I18n.Strings
            TestRunner.True("I18nLanguages.english " key, Type(text) = "String" && Trim(text) != "")
        languages := I18n.Languages()
        codes := ""
        for language in languages
            codes .= language[1] "=" language[2] " "
        eq("list", codes, "en=English ja=日本語 zh-CN=简体中文 zh-TW=繁體中文 ")
        values := ""
        for value in PreferencesWindow._LanguageValues()
            values .= value " "
        eq("preference values", values, "auto en ja zh-CN zh-TW ")
        for language in languages {
            code := language[1]
            if (code = "en")
                continue
            data := I18n.ReadFile(code)
            missing := "", extra := ""
            for key, english in I18n.Strings {
                if !data.Has(key) {
                    missing .= key " "
                    continue
                }
                TestRunner.True("I18nLanguages." code " not empty " key, Trim(data[key]) != "")
                RegExReplace(english, "\{\d\}", , &countEn)
                RegExReplace(data[key], "\{\d\}", , &count)
                TestRunner.True("I18nLanguages." code " placeholders " key, count = countEn)
            }
            for key in data
                if (SubStr(key, 1, 1) != "_" && !I18n.Strings.Has(key))
                    extra .= key " "
            eq(code " missing keys", missing, "")
            eq(code " extra keys", extra, "")
            eq(code " font", data.Get("_font", "") != "", true)
        }

        savedLang := I18n.Lang
        eq("normalize old zh", I18n.Normalize("zh"), "zh-CN")
        eq("normalize case", I18n.Normalize(" ZH_tw "), "zh-TW")
        eq("normalize region", I18n.Normalize("zh-HK"), "zh-TW")
        eq("normalize empty", I18n.Normalize(""), "auto")
        I18n.Init("zh")
        eq("zh-CN", I18n.Lang "|" I18n.T("Prefs.Page.General") "|" ThemeManager.FontName(), "zh-CN|通用|Microsoft YaHei UI")
        I18n.Init("ja")
        eq("ja", I18n.Lang "|" I18n.T("Prefs.Page.General") "|" ThemeManager.FontName(), "ja|全般|Yu Gothic UI")
        I18n.Init("zh-TW")
        eq("zh-TW", I18n.Lang "|" I18n.T("Prefs.Page.General") "|" ThemeManager.FontName(), "zh-TW|一般|Microsoft JhengHei UI")
        eq("zh-TW args", I18n.T("Clipboard.ClearHint", 3), "3 筆")
        I18n.Init("en")
        eq("en", I18n.Lang "|" I18n.T("Prefs.Page.General") "|" ThemeManager.FontName(), "en|General|Segoe UI")
        I18n.Init("ko-KR")                                                  ; 没有的语言: 英文
        eq("unknown", I18n.Lang "|" I18n.T("Prefs.Page.General"), "en|General")

        savedDir := I18n.LangDir
        dir := A_Temp "\ALTRun-lang-test"
        try DirDelete(dir, true)
        DirCreate(dir)
        I18n.LangDir := dir
        FileCopy(savedDir "\zh-CN.json", dir "\zh-CN.json")
        I18n.Init("zh-TW")                                                  ; 繁体的文件不在: 简体中文
        eq("no zh-TW file", I18n.Lang "|" I18n.T("Prefs.Page.General"), "zh-CN|通用")
        FileAppend('{"_name": "繁體中文", "_font": "Test Font", "_fallback": "zh-CN", "Prefs.Page.General": "一般"}', dir "\zh-TW.json", "UTF-8")
        I18n.Init("zh-TW")                                                  ; 缺的键: 先找 _fallback, 再用英文
        eq("fallback", I18n.Lang "|" I18n.T("Prefs.Page.General") "|" I18n.T("Prefs.Page.Appearance") "|" ThemeManager.FontName(), "zh-TW|一般|外观|Test Font")
        FileAppend('{"_name": "Test", "_fallback": "xx", "Prefs.Page.General": "Allgemein"}', dir "\de.json", "UTF-8")
        I18n.Init("de-DE")                                                  ; 只有语言部分的文件; _fallback 的文件不在: 英文
        eq("base language", I18n.Lang "|" I18n.T("Prefs.Page.General") "|" I18n.T("Prefs.Page.Appearance") "|" ThemeManager.FontName(), "de|Allgemein|Appearance|Segoe UI")
        eq("list from folder", I18n.Languages().Length, 4)
        FileDelete(dir "\zh-CN.json")
        FileAppend("{broken", dir "\zh-CN.json", "UTF-8")
        I18n.Init("zh-CN")                                                  ; 读不了: 英文
        eq("broken file", I18n.Lang "|" I18n.T("Prefs.Page.General"), "en|General")
        I18n.LangDir := savedDir
        DirDelete(dir, true)
        I18n.Init(savedLang)
    }

    ; 每条界面文字都有代码用到: 键名原样出现在源码里, 或者属于按名称拼出来的那几组 (Theme.<主题名> 等)
    static I18nUnused() {
        source := ""
        Loop Files, A_ScriptDir "\..\Src\*.ahk", "R"
            if (A_LoopFileName != "I18n.ahk")
                source .= FileRead(A_LoopFileFullPath, "UTF-8")
        dynamic := ["Help.", "Theme.", "Setting.", "Usage.F.", "Usage.Col.", "Prefs.Feature.", "Prefs.Status.", "Prefs.TypeShort.", "Cmd.Field."]
        unused := ""
        for key in I18n.Strings {
            if (InStr(source, '"' key '"', true) || SubStr(key, -5) = ".Desc")
                continue
            for prefix in dynamic
                if (SubStr(key, 1, StrLen(prefix)) == prefix)
                    continue 2
            unused .= key " "
        }
        TestRunner.Equal("I18nUnused", unused, "")
    }

    ; CHANGELOG.md 的每个版本标题都有链接 (文件末尾的 [版本]: 网址), "未发布" 比较的是最新发布的版本
    static ChangelogLinks() {
        text := FileRead(A_ScriptDir "\..\CHANGELOG.md", "UTF-8")
        missing := "", newest := "", pos := 1
        while (pos := RegExMatch(text, "m)^## \[([^\]]+)\]", &m, pos)) {
            pos += m.Len
            if (newest = "" && m[1] != "未发布")
                newest := m[1]
            if !RegExMatch(text, "m)^\[\Q" m[1] "\E\]: https://github\.com/\S+$")
                missing .= m[1] " "
        }
        TestRunner.Equal("ChangelogLinks.missing", missing, "")
        TestRunner.True("ChangelogLinks.unreleased compares " newest, InStr(text, "[未发布]: https://github.com/zhugecaomao/ALTRun/compare/" newest "...HEAD"))
        TestRunner.True("ChangelogLinks.newest is App.Version", newest = App.Version)
    }

    ; 数据文件夹: DataLocation.txt 指定的文件夹 (环境变量、相对路径), 空文件 = 默认位置; 写 / 恢复默认
    static DataLocation() {
        eq := (n, a, e) => TestRunner.Equal("DataLocation." n, a, e)
        eq("full: absolute", Path.Full("C:\Sync\ALTRun\"), "C:\Sync\ALTRun")
        eq("full: env", Path.Full("%SystemRoot%\Temp"), A_WinDir "\Temp")
        eq("full: relative", Path.Full("..\Shared\Data", "C:\Apps\ALTRun"), "C:\Apps\Shared\Data")
        eq("full: builtin", Path.Full("A_AppData\ALTRun"), A_AppData "\ALTRun")
        root := A_Temp "\ALTRun-location-test"
        try DirDelete(root, true)
        DirCreate(root "\app"), DirCreate(root "\user")
        dirs := [root "\app", root "\user"]
        eq("no file", AppSettings.CustomDataDir(dirs), "")
        FileAppend("", root "\app\DataLocation.txt")
        eq("empty file", AppSettings.CustomDataDir(dirs), "")
        FileAppend("  D:\OneDrive\ALTRun\  `r`n", root "\user\DataLocation.txt", "UTF-8")
        eq("user folder", AppSettings.CustomDataDir(dirs), "D:\OneDrive\ALTRun")
        FileOpen(root "\app\DataLocation.txt", "w").Write("..\Shared")
        eq("program folder first, relative", AppSettings.CustomDataDir(dirs), Path.Full(A_ScriptDir "\..\Shared"))

        saved := {Portable: AppSettings.Portable, UserDir: AppSettings.UserDir}
        AppSettings.Portable := false, AppSettings.UserDir := root "\user"   ; 不能写程序目录时写在 %APPDATA%\ALTRun
        AppSettings.SetDataLocation("E:\Sync\ALTRun")
        eq("set", FileRead(root "\user\DataLocation.txt", "UTF-8"), "E:\Sync\ALTRun")
        AppSettings.SetDataLocation("")
        eq("reset keeps an empty file", FileExist(root "\user\DataLocation.txt") ? FileRead(root "\user\DataLocation.txt") : "missing", "")
        AppSettings.SetDataLocation(AppSettings.DefaultDataDir())
        eq("default location = reset", FileRead(root "\user\DataLocation.txt"), "")
        eq("default when not portable", AppSettings.DefaultDataDir(), root "\user\Data")
        AppSettings.Portable := saved.Portable, AppSettings.UserDir := saved.UserDir
        eq("default when portable", AppSettings.DefaultDataDir(), A_ScriptDir "\Data")
        DirDelete(root, true)
    }

    ; Windows 设置的页面: 按名称 (和英文名称) 搜到, 可以整组关掉
    static SettingsPages() {
        eq := (n, a, e) => TestRunner.Equal("SettingsPages." n, a, e)
        find(text) {
            for item in SystemProvider.Search(SearchQuery(text))
                if (item.Source["Id"] = "SetBluetooth")
                    return item
            return ""
        }
        item := find("bluetooth")
        eq("found", IsObject(item) ? item.Title "|" item.Subtitle : "", "Bluetooth & devices|Windows Settings")
        pages := 0
        for command in SystemProvider.Commands()
            if command.Has("IsSetting")
                pages += 1
        TestRunner.True("SettingsPages.count " pages, pages >= 40)
        options := AppSettings.Feature("System")
        options["SettingsPages"] := 0
        eq("turned off", IsObject(find("bluetooth")), false)
        options["SettingsPages"] := 1
        eq("default on", AppSettings.Defaults()["Features"]["System"]["SettingsPages"], 1)
    }

    ; 切换窗口: 另开一个进程显示一个窗口, "w 标题" 找到它, 关闭; 自己的窗口不列出
    static WindowSwitch() {
        eq := (n, a, e) => TestRunner.Equal("WindowSwitch." n, a, e)
        script := A_Temp "\ALTRun-window-test.ahk", title := "ALTRun Window Test " A_TickCount
        try FileDelete(script)
        FileAppend('#NoTrayIcon`ng := Gui(, "' title '")`ng.OnEvent("Close", (*) => ExitApp())`ng.Show("w300 h120")`nSetTimer(() => ExitApp(), -20000)', script, "UTF-8")
        Run('"' A_AhkPath '" "' script '"', , , &pid)
        found := WinWait(title, , 10)
        eq("test window shown", found != 0, true)
        WindowProvider._listTime := 0
        items := WindowProvider.Search(SearchQuery("w " SubStr(title, 1, 18)))
        item := ""
        for candidate in items
            if (candidate.Title = title)
                item := candidate
        eq("found by keyword", IsObject(item) ? item.Exclusive "|" InStr(item.Subtitle, "AutoHotkey") : "missing", "1|1")
        all := WindowProvider.Search(SearchQuery("w "))
        own := false
        for candidate in all
            if (candidate.Title = title)
                own := true
        eq("listed with empty keyword", own, true)
        test := Gui(, "ALTRun own window " A_TickCount)
        test.Show("w200 h80")
        WindowProvider._listTime := 0
        mine := false
        for window in WindowProvider.List()
            if (window.Hwnd = test.Hwnd)
                mine := true
        eq("own windows not listed", mine, false)
        test.Destroy()
        if IsObject(item) {
            item.Actions[1].OnRun.Call()                                     ; 操作面板里的 "关闭窗口"
            eq("closed", WinWaitClose(title, , 5), 1)
        }
        try ProcessClose(pid)
        try FileDelete(script)
        AppSettings.Feature("Windows")["InDefaultResults"] := 0
        eq("default results off", WindowProvider.Search(SearchQuery("ALTRun Window")).Length, 0)
        AppSettings.Feature("Windows")["InDefaultResults"] := 1
    }

    ; 空搜索框: 置顶在前, 然后最近打开的 (去重、限制个数); 只记能重新打开的结果
    static RecentItems() {
        eq := (n, a, e) => TestRunner.Equal("RecentItems." n, a, e)
        saved := {Recent: Knowledge.Recent, File: Knowledge.File, SettingsFile: AppSettings.File}
        options := AppSettings.Feature("Recent"), savedPinned := options["Pinned"], savedCount := options["RecentCount"]
        Knowledge.Recent := [], Knowledge.File := A_Temp "\ALTRun-recent-test.json", AppSettings.File := A_Temp "\ALTRun-recent-settings.json"
        options["Pinned"] := [], options["RecentCount"] := 2
        fileItem := (name) => ResultItem(name, "C:\Docs\" name, {Kind: "file", Arg: "C:\Docs\" name, Uid: "file:" name, Provider: "FileSearch", Icon: "C:\Docs\" name})
        RecentProvider.Remember(fileItem("a.txt"))
        RecentProvider.Remember(fileItem("b.txt"))
        RecentProvider.Remember(fileItem("c.txt"))
        RecentProvider.Remember(fileItem("a.txt"))                               ; 再打开一次: 移到最前面, 不重复
        clip := ResultItem("secret", "", {Kind: "text", Arg: "secret", Uid: "clip:1", Provider: "Clipboard"})
        RecentProvider.Remember(clip)                                        ; 剪贴板历史不记
        sys := ResultItem("Lock", "", {Uid: "system:Lock", Provider: "System", OnRun: (*) => 0})
        RecentProvider.Remember(sys)
        eq("order", Knowledge.Recent.Length "|" Knowledge.Recent[1]["Uid"] "|" Knowledge.Recent[2]["Uid"] "|" Knowledge.Recent[3]["Uid"], "4|system:Lock|file:a.txt|file:c.txt")
        titles(items) {
            list := ""
            for item in items
                list .= item.Title "|"
            return RTrim(list, "|")
        }
        eq("empty box: recent count", titles(RecentProvider.EmptyResults()), I18n.T("Sys.Lock") "|a.txt")
        options["Pinned"].Push(RecentProvider.Snapshot(fileItem("c.txt")))
        items := RecentProvider.EmptyResults()
        eq("pinned first", titles(items), "c.txt|" I18n.T("Sys.Lock") "|a.txt")
        eq("pinned flag", items[1].Pinned "|" items[2].Pinned "|" items[1].Provider, "1|0|Recent")
        eq("is pinned", RecentProvider.IsPinned(fileItem("c.txt")) "|" RecentProvider.IsPinned(fileItem("b.txt")), "1|0")
        eq("can pin", RecentProvider.CanPin(fileItem("b.txt")) "|" RecentProvider.CanPin(clip), "1|0")
        actions := ""
        for action in ActionCatalog.ListFor(items[2])
            actions .= action.Title "|"
        eq("pin action", InStr(actions, I18n.T("Action.Pin")) > 0, true)
        RecentProvider.DeleteItem(items[3])                                  ; Ctrl+Del: 从最近使用里去掉
        eq("forget", titles(RecentProvider.EmptyResults()), "c.txt|" I18n.T("Sys.Lock") "|b.txt")
        RecentProvider.DeleteItem(items[1])                                  ; 置顶的: 取消置顶
        eq("unpin", options["Pinned"].Length, 0)
        options["RecentCount"] := 0
        eq("count 0", RecentProvider.EmptyResults().Length, 0)
        Knowledge.Recent := saved.Recent, Knowledge.File := saved.File, AppSettings.File := saved.SettingsFile
        options["Pinned"] := savedPinned, options["RecentCount"] := savedCount
        try FileDelete(A_Temp "\ALTRun-recent-test.json")
        try FileDelete(A_Temp "\ALTRun-recent-settings.json")
    }

    ; 剪贴板置顶: 排在最前面, 超过条数不删, 再复制仍然置顶, 清空时保留, 保存后还在
    static ClipboardPin() {
        eq := (n, a, e) => TestRunner.Equal("ClipboardPin." n, a, e)
        root := A_Temp "\ALTRun-clip-pin-test"
        try DirDelete(root, true)
        DirCreate(root)
        saved := [ClipboardProvider.File, ClipboardProvider.Folder, ClipboardProvider.Entries]
        options := AppSettings.Feature("Clipboard"), savedMax := options["MaxItems"], savedPersist := options["Persist"]
        ClipboardProvider.File := root "\ClipboardHistory.json", ClipboardProvider.Folder := root "\Clipboard", ClipboardProvider.Entries := []
        options["MaxItems"] := 3, options["Persist"] := 1
        ClipboardProvider.Add("one"), ClipboardProvider.Add("two")
        ClipboardProvider.Entries[2]["Pinned"] := 1                          ; "one" 置顶
        ClipboardProvider.Add("three"), ClipboardProvider.Add("four"), ClipboardProvider.Add("five")
        texts() {
            list := ""
            for item in ClipboardProvider.Search(SearchQuery("clip "))
                if (item.Kind = "text")
                    list .= item.Arg "|"
            return RTrim(list, "|")
        }
        eq("pinned first, kept when trimming", texts(), "one|five|four")
        items := ClipboardProvider.Search(SearchQuery("clip one"))
        eq("pinned tag", InStr(items[1].Subtitle, I18n.T("Clipboard.PinnedTag")) = 1, true)
        actions := ""
        for action in items[1].Actions
            actions .= action.Title "|"
        eq("unpin action", InStr(actions, I18n.T("Clipboard.Unpin")) > 0, true)
        ClipboardProvider.Add("one")                                         ; 再复制一次: 还是置顶
        eq("copy again keeps pin", ClipboardProvider.IsPinned(ClipboardProvider.Entries[1]), true)
        ClipboardProvider.Save()
        ClipboardProvider.Entries := [], ClipboardProvider._Load()
        eq("saved", texts(), "one|five|four")
        ClipboardProvider.Clear()
        eq("clear keeps pinned", texts(), "one")
        ClipboardProvider.Entries[1].Delete("Pinned"), ClipboardProvider.Clear()
        eq("unpinned cleared", ClipboardProvider.Entries.Length, 0)
        ClipboardProvider.File := saved[1], ClipboardProvider.Folder := saved[2], ClipboardProvider.Entries := saved[3]
        options["MaxItems"] := savedMax, options["Persist"] := savedPersist
        try DirDelete(root, true)
    }

    ; 输入路径浏览文件夹: 文件夹在前, 最后一段过滤, 隐藏文件不列出, Tab 进入下一级; Insert 标记多个一起操作
    static FolderBrowse() {
        eq := (n, a, e) => TestRunner.Equal("FolderBrowse." n, a, e)
        root := A_Temp "\ALTRun-browse-test"
        try DirDelete(root, true)
        DirCreate(root "\Reports"), DirCreate(root "\Archive")
        FileAppend("x", root "\readme.txt"), FileAppend("x", root "\report 2026.docx"), FileAppend("x", root "\secret.txt")
        FileSetAttrib("+H", root "\secret.txt")
        browse := FileSearchProvider.BrowsePath(root "\rep")
        eq("parse", IsObject(browse) ? browse.Dir "|" browse.Filter : "", root "\|rep")
        eq("not a path", FileSearchProvider.BrowsePath("report"), "")
        eq("drive only", FileSearchProvider.BrowsePath("C:"), "")
        eq("missing folder", FileSearchProvider.BrowsePath(root "\Nope\x"), "")
        home := FileSearchProvider.BrowsePath("~")
        eq("home", IsObject(home) ? home.Dir : "", EnvGet("UserProfile") "\")
        titles(text) {
            list := ""
            for item in FileSearchProvider.Search(SearchQuery(text))
                list .= item.Title "|"
            return RTrim(list, "|")
        }
        eq("all, folders first, no hidden", titles(root "\"), "Archive|Reports|readme.txt|report 2026.docx")
        eq("filter", titles(root "\rep"), "Reports|report 2026.docx")
        items := FileSearchProvider.Search(SearchQuery(root "\Rep"))
        eq("folder tab completes", items[1].AutoComplete "|" items[1].Exclusive "|" items[1].Kind, root "\Reports\|1|folder")
        eq("file", items[2].Kind "|" items[2].Arg, "file|" root "\report 2026.docx")
        eq("empty", titles(root "\Reports\"), I18n.T("Files.EmptyFolder"))
        eq("file mode", FileSearchProvider.SearchFiles(root "\arc")[1].Title, "Archive")

        SearchWindow.Marked := Map(StrLower(items[2].Arg), items[2].Arg)
        eq("marked", SearchWindow.IsMarked(items[2]) "|" SearchWindow.IsMarked(items[1]), "1|0")
        SearchWindow.Marked := Map()
        actions := ""
        for action in ActionCatalog.ListForMany([root "\readme.txt", root "\Archive"])
            actions .= action.Title "|"
        eq("many: open all", InStr(actions, I18n.T("Action.OpenAll", 2)) = 1, true)
        eq("many: copy, recycle", InStr(actions, I18n.T("Action.CopyFiles", 2)) && InStr(actions, I18n.T("Action.RecycleAll", 2)) ? 1 : 0, 1)
        FileSetAttrib("-H", root "\secret.txt")
        DirDelete(root, true)
    }

    ; 计算器: 进制换算和日期加减
    static CalcBasesDates() {
        eq := (n, a, e) => TestRunner.Equal("CalcBasesDates." n, a, e)
        first(text) {
            items := CalculatorProvider.Search(SearchQuery(text))
            return items.Length ? items[1].Title : ""
        }
        eq("hex in", first("255 in hex"), "FF")
        eq("to dec", first("0xff to dec"), "255")
        eq("bin", first("10 in bin"), "1010")
        eq("oct", first("0o17 in dec"), "15")
        items := CalculatorProvider.Search(SearchQuery("0b1010"))
        eq("prefixed alone: three forms", items.Length "|" items[1].Title "|" items[2].Title "|" items[3].Title, "3|10|0xA|0b1010")
        eq("plain number not a base", first("255"), "")
        eq("zero", CalculatorProvider.ToBase(0, 16), "0")
        eq("big", CalculatorProvider.ParseInteger("0xFFFFFFFF"), 4294967295)
        eq("too big", CalculatorProvider.ParseInteger("0x" "FFFFFFFFFFFFFFFFFF"), "")

        date := (text) => (items := CalculatorProvider._Dates(text, "20261004")).Length ? SubStr(items[1].Title, 1, 10) : ""
        eq("today + days", date("today + 30 days"), "2026-11-03")
        eq("today - weeks", date("today - 2w"), "2026-09-20")
        eq("chinese units", date("今天 + 1 年"), "2027-10-04")
        eq("months clamp", date("2026-01-31 + 1 month"), "2026-02-28")
        eq("leap year", date("2028-01-31 + 1m"), "2028-02-29")
        eq("months back across year", date("2026-02-15 - 3 months"), "2025-11-15")
        eq("date difference", CalculatorProvider._Dates("2026-12-25 - today", "20261004")[1].Title, I18n.T("Calc.Days", 82))
        eq("invalid date", date("2026-02-30 + 1d"), "")
        eq("normal math still works", first("12*(3+4)"), "84")
    }

    ; 脚本扩展: 读开头注释里的 @altrun.xxx, 按名称 / 关键字搜到, 带参数, 命令行, 后台运行的输出
    static Scripts() {
        eq := (n, a, e) => TestRunner.Equal("Scripts." n, a, e)
        root := A_Temp "\ALTRun-scripts-test"
        try DirDelete(root, true)
        DirCreate(root)
        FileAppend("; @altrun.title Restart Explorer`n; @altrun.keyword rex`n; @altrun.mode silent`nProcessClose(`"explorer.exe`")`n", root "\restart.ahk", "UTF-8")
        FileAppend("# @altrun.title  Ping Host`n# @altrun.keyword ping`n# @altrun.argument Host name or IP`n# @altrun.mode output`nping $args[0]`n", root "\ping.ps1", "UTF-8")
        FileAppend("@echo off`nREM @altrun.title Clean Temp`necho done`n", root "\clean.bat", "UTF-8")
        FileAppend("not a script", root "\notes.txt")
        saved := [ScriptProvider.Folder, ScriptProvider.Scripts]
        ScriptProvider.Folder := root
        ScriptProvider.Load()
        eq("loaded", ScriptProvider.Scripts.Length, 3)
        byTitle := Map()
        for script in ScriptProvider.Scripts
            byTitle[script.Title] := script
        eq("meta", byTitle["Ping Host"].Keyword "|" byTitle["Ping Host"].Argument "|" byTitle["Ping Host"].Mode, "ping|Host name or IP|output")
        eq("bat comment, default mode", byTitle["Clean Temp"].Mode, "window")
        titles(text) {
            list := ""
            for item in ProviderRegistry.SortByScore(ScriptProvider.Search(SearchQuery(text)))
                list .= item.Title "|"
            return RTrim(list, "|")
        }
        eq("by title", titles("restart"), "Restart Explorer")
        eq("by keyword", titles("rex"), "Restart Explorer")
        items := ScriptProvider.Search(SearchQuery("ping 10.0.0.1"))
        eq("argument", items.Length "|" items[1].Title "|" items[1].Exclusive, "1|Ping Host: 10.0.0.1|1")
        item := ""
        for found in ScriptProvider.Search(SearchQuery("pin"))
            if (found.Source.Title = "Ping Host")
                item := found
        eq("needs argument: complete keyword", item.Valid "|" item.AutoComplete, "0|ping ")
        eq("ps1 command", ScriptProvider.CommandLine(byTitle["Ping Host"], "a b"), 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' root '\ping.ps1" "a b"')
        eq("bat command", ScriptProvider.CommandLine(byTitle["Clean Temp"]), A_ComSpec ' /c ""' root '\clean.bat""')
        FileAppend("line one`r`n`r`nlast line  `r`n", root "\out.txt")
        eq("silent: last line", ScriptProvider._Finished(byTitle["Restart Explorer"], root "\out.txt"), "last line")
        FileAppend("", root "\empty.txt")
        eq("silent: no output", ScriptProvider._Finished(byTitle["Restart Explorer"], root "\empty.txt"), I18n.T("Script.Done", "Restart Explorer"))
        FileAppend("; @altrun.title New One`n", root "\new.ahk", "UTF-8")
        ScriptProvider._checked := 0                                         ; 文件夹变了: 下一次搜索时重新读
        eq("refresh", titles("new one"), "New One")
        ScriptProvider.Folder := saved[1], ScriptProvider.Scripts := saved[2]
        DirDelete(root, true)
    }

    static WindowPosition() {
        eq := (n, a, e) => TestRunner.Equal("WindowPosition." n, a, e)
        area := {Left: 0, Top: 0, Right: 1920, Bottom: 1040}
        pos := SearchWindow.Place(area, 700, 500)
        eq("default", pos.X "," pos.Y, "610,208")
        pos := SearchWindow.Place(area, 700, 500, Map("X", 500, "Y", 200))
        eq("default map", pos.X "," pos.Y, "610,208")
        pos := SearchWindow.Place(area, 700, 500, Map("X", 0, "Y", 0))
        eq("top left", pos.X "," pos.Y, "0,0")
        pos := SearchWindow.Place(area, 700, 500, Map("X", 1000, "Y", 1000))
        eq("bottom right kept on screen", pos.X "," pos.Y, "1220,540")
        pos := SearchWindow.Place(area, 700, 500, Map("X", 1500, "Y", -30))
        eq("out of range clamped", pos.X "," pos.Y, "1220,0")
        second := {Left: 1920, Top: 0, Right: 3200, Bottom: 1024}
        pos := SearchWindow.Place(second, 700, 500)
        eq("second screen", pos.X "," pos.Y, "2210,205")
        left := {Left: -1280, Top: -200, Right: 0, Bottom: 824}
        pos := SearchWindow.Place(left, 700, 500, Map("X", 250, "Y", 100))
        eq("screen with negative coordinates", pos.X "," pos.Y, "-1135,-98")
        for bad in ["", Map(), Map("X", "a", "Y", 1), 5]
            eq("invalid position " A_Index, SearchWindow.Place(area, 700, 500, bad).X, 610)
        relative := SearchWindow.RelativePosition(area, 700, 305, 312)
        eq("relative", relative["X"] "," relative["Y"], "250,300")
        pos := SearchWindow.Place(area, 700, 500, relative)
        eq("round trip", pos.X "," pos.Y, "305,312")
        pos := SearchWindow.Place(second, 700, 500, relative)
        eq("same relative place on another screen", pos.X "," pos.Y, "2065,307")
        relative := SearchWindow.RelativePosition(area, 700, -50, 5000)
        eq("relative clamped", relative["X"] "," relative["Y"], "0,1000")
        appearance := AppSettings.Defaults()["Appearance"]
        eq("default show on", appearance["ShowOn"], "Mouse")
        eq("default remember", appearance["RememberPosition"], 0)
        eq("default position", appearance["Position"]["X"] "," appearance["Position"]["Y"], "500,200")
        eq("wiki window page", PreferencesWindow.WikiPage("Prefs.Page.Window"), "Usage")
    }

    ; 偏好设置里的灰色说明: "<标签>.Desc" 要有对应的标签, 中英文都不能空
    static PreferenceDescriptions() {
        count := 0
        for key, english in I18n.Strings {
            if !(SubStr(key, -5) = ".Desc")
                continue
            count += 1
            base := SubStr(key, 1, -5)
            TestRunner.True("PreferenceDescriptions.label " base, I18n.Strings.Has(base))
            TestRunner.True("PreferenceDescriptions.text " key, Trim(english) != "")
        }
        TestRunner.True("PreferenceDescriptions.count " count, count >= 50)
        for key in ["Prefs.RememberPosition.Desc", "Prefs.ShowOn.Desc", "Prefs.HideOnDeactivate.Desc", "Prefs.Hotkey.Desc"]
            TestRunner.True("PreferenceDescriptions.has " key, I18n.Strings.Has(key))
    }

    ; 资源管理器 "发送到": 多个文件直接添加, 已经有的不重复添加 (1 个时弹对话框, 这里不测)
    static SendTo() {
        eq := (n, a, e) => TestRunner.Equal("SendTo." n, a, e)
        ; 关掉开机启动等设置时只删 ALTRun 自己建的快捷方式, 用户自己建的同名快捷方式保留 (#113)
        eq("own by description", App.IsOwnShortcut("", "ALTRun - 高效的 Windows 启动器", ""), true)
        eq("own by startup flag", App.IsOwnShortcut('"D:\ALTRun\ALTRun.ahk" -Startup', "", "-Startup"), true)
        eq("user shortcut", App.IsOwnShortcut("", "", "-Startup"), false)
        eq("user start menu shortcut", App.IsOwnShortcut("", "", ""), false)
        eq("other flag", App.IsOwnShortcut("-SendTo", "", "-Startup"), false)
        lnkDir := A_Temp "\ALTRunShortcutTest"
        try DirDelete(lnkDir, true)
        DirCreate(lnkDir)
        FileCreateShortcut(A_AhkPath, lnkDir "\own.lnk", A_ScriptDir, "-Startup", "ALTRun - An effective launcher for Windows")
        FileCreateShortcut(A_AhkPath, lnkDir "\user.lnk", A_ScriptDir)
        eq("remove own", App.RemoveOwnShortcut(lnkDir "\own.lnk", "-Startup") && !FileExist(lnkDir "\own.lnk"), true)
        eq("keep user", !App.RemoveOwnShortcut(lnkDir "\user.lnk", "-Startup") && FileExist(lnkDir "\user.lnk") != "", true)
        eq("missing", App.RemoveOwnShortcut(lnkDir "\none.lnk", "-Startup"), false)
        try DirDelete(lnkDir, true)
        folder := A_Temp "\ALTRunSendToTest"
        try DirDelete(folder, true)
        DirCreate(folder "\PT2415 - Riverside")
        FileAppend("", folder "\Design Report.docx")
        command := CustomCommandProvider.FromPath(folder "\PT2415 - Riverside\")
        eq("folder title", command["Title"], "PT2415 - Riverside")
        eq("folder type", command["Type"], "Folder")
        eq("folder target (no trailing slash)", command["Target"], folder "\PT2415 - Riverside")
        command := CustomCommandProvider.FromPath(folder "\Design Report.docx")
        eq("file title", command["Title"], "Design Report")
        eq("file type", command["Type"], "File")
        eq("drive root", CustomCommandProvider.FromPath("C:")["Target"], "C:\")

        savedFile := AppSettings.File, savedCommands := AppSettings.Data["CustomCommands"]
        AppSettings.File := A_Temp "\ALTRunTest.json"
        AppSettings.Data["CustomCommands"] := [Map("Title", "Old", "Type", "Folder", "Target", folder "\pt2415 - riverside\", "Arguments", "", "Keyword", "")]
        eq("find existing (case, trailing slash)", CustomCommandProvider.FindByTarget(folder "\PT2415 - Riverside")["Title"], "Old")
        eq("find missing", CustomCommandProvider.FindByTarget(folder "\Design Report.docx"), "")
        CustomCommandProvider.AddFromPaths([folder "\PT2415 - Riverside", folder "\Design Report.docx", "C:\Windows"])
        commands := AppSettings.CustomCommands
        eq("several added, existing skipped", commands.Length, 3)
        eq("added in order", commands[2]["Title"] "|" commands[3]["Title"], "Design Report|Windows")
        ToolTip(, , , 20)
        AppSettings.Data["CustomCommands"] := savedCommands, AppSettings.File := savedFile
        try FileDelete(A_Temp "\ALTRunTest.json")
        try DirDelete(folder, true)
    }

    ; Ctrl+↑ / Ctrl+↓ 翻搜索记录 (↑ ↓ 只移动选择); 用隐藏的输入框代替搜索窗口, 不真的搜索
    static HistoryKeys() {
        eq := (n, a, e) => TestRunner.Equal("HistoryKeys." n, a, e)
        savedInput := SearchWindow.Input, savedHistory := Knowledge.History, savedSearch := SearchWindow.GetOwnPropDesc("_RunSearch")
        SearchWindow.DefineProp("_RunSearch", {Call: (*) => 0})
        g := Gui()
        SearchWindow.Input := g.AddEdit("w200")
        SearchWindow.Mode := "results", SearchWindow.FileMode := false, SearchWindow.HistoryIndex := 0
        Knowledge.History := ["nir", "wah", "pt tools"]
        try {
            SearchWindow.Input.Value := "typed"
            SearchWindow.RecallHistory(1)
            eq("ctrl+up newest", SearchWindow.Input.Value, "nir")
            SearchWindow.RecallHistory(1)
            eq("ctrl+up older", SearchWindow.Input.Value, "wah")
            SearchWindow.RecallHistory(1), SearchWindow.RecallHistory(1)
            eq("stops at oldest", SearchWindow.Input.Value "|" SearchWindow.HistoryIndex, "pt tools|3")
            SearchWindow.RecallHistory(-1)
            eq("ctrl+down newer", SearchWindow.Input.Value, "wah")
            SearchWindow.RecallHistory(-1), SearchWindow.RecallHistory(-1)
            eq("back to typed text", SearchWindow.Input.Value "|" SearchWindow.HistoryIndex, "typed|0")
            SearchWindow.RecallHistory(-1)
            eq("ctrl+down at the start does nothing", SearchWindow.Input.Value, "typed")
            SearchWindow.Input.Value := "nir", SearchWindow.HistoryIndex := 0     ; 例如保留的上一次搜索
            SearchWindow.RecallHistory(1)
            eq("same text skipped", SearchWindow.Input.Value, "wah")
            SearchWindow.HistoryIndex := 0, SearchWindow.FileMode := true
            SearchWindow.RecallHistory(1)
            eq("not in file mode", SearchWindow.Input.Value, "wah")
        } finally {
            SearchWindow.FileMode := false, SearchWindow.HistoryIndex := 0
            SearchWindow.DefineProp("_RunSearch", savedSearch)
            SearchWindow.Input := savedInput, Knowledge.History := savedHistory
            g.Destroy()
        }
    }

    ; 束线型计算和原来的 SPF2M.EXE 逐个比对: Tests\Fixtures\SPF2M-Reference.json 是在 DOSBox 里运行
    ; SPF2M 得到的结果 (4 种线型 x 6 种钢绞线, 上升 / 下降, 非整米跨度, 自定义半径 / 反弯点 /
    ; 支架间距, 以及 SPF2M 报错或崩溃的情况); 生成方法见 Tests\Tools\SPF2M
    static TendonProfileVsSpf2m() {
        data := JSON.Parse(FileRead(A_ScriptDir "\Fixtures\SPF2M-Reference.json", "UTF-8"))
        checked := 0, failed := 0
        for ref in data["Cases"] {
            input := Map()
            for key in ["Profile", "Tendon", "Start", "End", "Distance", "Radius", "Contraflexure", "Intervals"]
                if ref.Has(key)
                    input[key] := ref[key]
            result := TendonProfile.Calc(input)
            name := "TendonProfile." ref["Profile"] "/" ref["Tendon"] " " ref["Start"] "->" ref["End"] " L" ref["Distance"]
                . (ref.Has("Radius") ? " R" ref["Radius"] : "") . (ref.Has("Contraflexure") ? " C" ref["Contraflexure"] : "")
            if ref.Has("Expect") {
                TestRunner.True(name " error", result.Error != "")
                continue
            }
            if (result.Error != "") {
                TestRunner.Fail(name, "unexpected error: " result.Error)
                continue
            }
            got := "", want := ""
            for row in result.Rows
                got .= row[1] "," row[2] "," row[3] "," row[4] "," row[5] ";"
            for row in ref["Rows"]
                want .= row[1] "," row[2] "," row[3] "," row[4] "," row[5] ";"
            if (got != want) {
                failed += 1
                TestRunner.Equal(name " rows", got, want)
            }
            checked += ref["Rows"].Length
            TestRunner.Equal(name " contraflexure", Integer(result.Contraflexure), ref["ShownContraflexure"])
            if ref.Has("ShownRadius")
                TestRunner.Equal(name " radius", Integer(result.Radius), ref["ShownRadius"])
        }
        TestRunner.True("TendonProfile.rows checked (" checked ")", checked > 1000 && !failed)
    }

    ; 输入检查、支架间距、SPF2M 没有处理好的情况 (崩溃 / 不显示任何东西)
    static TendonProfileInputs() {
        eq := (n, a, e) => TestRunner.Equal("TendonProfileInputs." n, a, e)
        profileOf := (profile, start, finish, l, extra := "") => TendonProfile.Calc(TendonProfileInputsMap(profile, start, finish, l, extra))
        TendonProfileInputsMap(profile, start, finish, l, extra) {
            input := Map("Profile", profile, "Tendon", 3, "Start", start, "End", finish, "Distance", l)
            if IsObject(extra)
                for key, value in extra
                    input[key] := value
            return input
        }
        intervals := (l) => TendonProfileJoin(TendonProfile.DefaultIntervals(l))
        TendonProfileJoin(list) {
            text := ""
            for v in list
                text .= (text = "" ? "" : ",") v
            return text
        }
        eq("intervals whole metres", intervals(8000), "1000,1000,1000,1000,1000,1000,1000,1000")
        eq("intervals single split", intervals(9750), "750,1000,1000,1000,1000,1000,1000,1000,1000,1000")
        eq("intervals double split", intervals(7350), "650,700,1000,1000,1000,1000,1000,1000")
        eq("intervals double split 2", intervals(8420), "720,700,1000,1000,1000,1000,1000,1000,1000")
        eq("intervals short span", intervals(450), "450")
        eq("intervals 1300", intervals(1300), "600,700")
        eq("parse intervals", TendonProfileJoin(TendonProfile.ParseIntervals("500, 1500 800;550")), "500,1500,800,550")
        eq("parse empty = auto", TendonProfile.ParseIntervals("  "), "")
        eq("intervals must add up", InStr(profileOf(1, 600, 100, 8000, Map("Intervals", [1000, 1000])).Error, "add up") > 0, true)
        eq("interval not a number", profileOf(1, 600, 100, 8000, Map("Intervals", TendonProfile.ParseIntervals("1000 abc"))).Error != "", true)
        eq("missing distance", profileOf(1, 600, 100, "").Error != "", true)
        eq("missing level", profileOf(1, "", 100, 8000).Error != "", true)
        ; SPF2M 崩溃的情况: 这里给出错误说明
        eq("psp not achievable", InStr(profileOf(2, 900, 100, 3000, Map("Tendon", 6)).Error, "NOT ACHIEVABLE") > 0, true)
        ; SPF2M 什么都不显示的情况: 反弯点超过半跨
        eq("contraflexure beyond half span", InStr(profileOf(1, 600, 100, 9000, Map("Contraflexure", 5000)).Error, "half") > 0, true)
        eq("contraflexure with equal levels", profileOf(1, 300, 300, 6000, Map("Contraflexure", 500)).Error != "", true)
        result := profileOf(1, 600, 100, 9000, Map("Contraflexure", 1200))
        eq("contraflexure -> radius", Integer(result.Radius) "|" result.RadiusChanged, "10800|1")
        result := profileOf(1, 600, 100, 8000, Map("Contraflexure", 525))
        eq("contraflexure equal to default keeps radius", result.Radius "|" result.RadiusChanged, "4200|0")
        eq("default radius 22s", TendonProfile.DefaultRadius(5), 5700)
        eq("default duct dia", TendonProfile.DefaultDuctDiameter(1) "|" TendonProfile.DefaultDuctDiameter(3) "|" TendonProfile.DefaultDuctDiameter(6), "25|90|130")
        eq("default tendon 12S", PTToolsWindow.Defaults["TendonType"], 3)
    }

    ; 已发布的 v2026.08.12 用 ALTRun.ini: 第一次启动新版本时整体导入
    static LegacyIni() {
        eq := (n, a, e) => TestRunner.Equal("LegacyIni." n, a, e)
        folder := A_Temp "\ALTRunLegacyTest"
        try DirDelete(folder, true)
        DirCreate(folder)
        iniFile := folder "\ALTRun.ini"
        ; 和 2.x 写出的文件一样: UTF-16, 注释行, 键里的 = 和 ; 被转义
        FileAppend("
        (
[Config]
AutoStartup=0
Chinese=1
FileMgr=C:\Apps\TotalCMD64.exe /O /T /S
IndexDir=A_ProgramsCommon,A_StartMenu,C:\Path\IndexLocation
IndexType=*.lnk,*.exe
IndexDepth=2
StruCalc=1
AutoSwitchDir=1
SpaceToRun=1
KeepInput=0
AutoEngIME=1
[Hotkey]
GlobalHotkey1=~!Space
GlobalHotkey2=!r
TotalCMDDir=^g
CondTitle=ahk_exe RAPTW.exe
CondHotkey=~Mbutton
CondAction=PTTools
[UserCommand]
; This section is User-Defined commands, modify as desired
; Format: Command Type | Command | Description=Rank
File | C:\Windows\Notepad.exe=9
Dir | A_Desktop | Desktop=99
CMD | cmd.exe /k ipconfig | Check IP Address=9
URL | https://www.google.com/search?q_Equal_x_Semicolon_y | Google Query=5
Dir | Q:\DESIGN PROJECTS\Design-2019\PT1931 - 24 NIR | CKR, EA, JIB=3
Func | PTTools | PT Tools (AHK)=99
[History]
1=Dir | A_Desktop | Desktop
        )", iniFile, "UTF-16")

        data := SchemaMigration.ReadLegacyIni(iniFile)
        eq("config number", data["Config"]["IndexDepth"], 2)
        eq("config value with spaces", data["Config"]["FileMgr"], "C:\Apps\TotalCMD64.exe /O /T /S")
        eq("comments skipped", data["UserCommand"].Count, 6)
        TestRunner.True("LegacyIni.escaped key restored", data["UserCommand"].Has("URL | https://www.google.com/search?q=x;y | Google Query"))
        eq("detected as 2.x", SchemaMigration.DetectVersion(data), 2)

        ; 完整的启动流程: 没有 ALTRun.json, 只有程序目录下的 ALTRun.ini
        savedFile := AppSettings.File, savedLegacy := AppSettings.LegacyFile, savedData := AppSettings.Data
        AppSettings.File := folder "\Data\ALTRun.json", AppSettings.LegacyFile := folder "\ALTRun.json"
        AppSettings.ImportedFrom := "", AppSettings.MigratedFrom := 0, AppSettings.MovedFrom := ""
        AppSettings.Load()
        settings := AppSettings.Data
        eq("imported from", AppSettings.ImportedFrom, iniFile)
        eq("version", settings["SchemaVersion"], AppSettings.CurrentVersion)
        eq("json written to Data", FileExist(folder "\Data\ALTRun.json") != "", true)
        eq("nothing moved", AppSettings.MovedFrom, "")
        eq("ini kept", FileExist(iniFile) != "", true)
        eq("hotkey", settings["General"]["Hotkey"], "!Space")
        eq("second hotkey", settings["General"]["SecondaryHotkey"], "!r")
        eq("language", settings["General"]["Language"], "zh-CN")
        eq("startup", settings["General"]["LaunchAtLogin"], 0)
        eq("file manager", settings["General"]["FileManager"], "C:\Apps\TotalCMD64.exe /O /T /S")
        eq("index depth", settings["Features"]["Applications"]["Depth"], 2)
        eq("index placeholder dropped", settings["Features"]["Applications"]["Folders"].Length, 2)
        eq("structural calc", settings["Features"]["Calculator"]["StructuralCalc"], 1)
        eq("quick switch", settings["Extensions"]["QuickSwitch"]["AutoSwitch"], 1)
        eq("space to run", settings["General"]["SpaceToRun"], 1)
        eq("keep input off", settings["General"]["KeepLastQuery"], 0)
        eq("english input", settings["General"]["SwitchToEnglishInput"], 1)
        eq("conditional hotkey", settings["Hotkeys"][1]["Action"], "PTTools")
        eq("commands (Func skipped)", settings["CustomCommands"].Length, 5)
        commands := Map()
        for command in settings["CustomCommands"]
            commands[command["Title"]] := command
        eq("folder command", commands["CKR, EA, JIB"]["Target"], "Q:\DESIGN PROJECTS\Design-2019\PT1931 - 24 NIR")
        eq("command args", commands["Check IP Address"]["Arguments"], "/k ipconfig")
        eq("url unescaped", commands["Google Query"]["Target"], "https://www.google.com/search?q=x;y")
        eq("file title", commands["Notepad"]["Type"], "File")

        ; 真实的 v2026.08.12 生成的 ALTRun.ini (UTF-16, 带默认命令) + 用户加的一条命令和设置
        fixture := SchemaMigration.Upgrade(SchemaMigration.ReadLegacyIni(A_ScriptDir "\Fixtures\ALTRun.v2026.08.12.ini"), 2)
        eq("fixture commands", fixture["CustomCommands"].Length, 10)
        eq("fixture language", fixture["General"]["Language"], "zh-CN")
        eq("fixture keep input (2.x default on)", fixture["General"]["KeepLastQuery"], 1)
        eq("keep last query default off", AppSettings.Defaults()["General"]["KeepLastQuery"], 0)
        eq("fixture hotkey", fixture["Hotkeys"][1]["Key"], "~Mbutton")

        ; 第二次启动: 已经有 ALTRun.json, 不再导入
        AppSettings.ImportedFrom := ""
        AppSettings.Load()
        eq("no second import", AppSettings.ImportedFrom, "")

        AppSettings.File := savedFile, AppSettings.LegacyFile := savedLegacy, AppSettings.Data := savedData
        AppSettings.ImportedFrom := "", AppSettings.MigratedFrom := 0
        try DirDelete(folder, true)
    }

    ; 旧版本的 ALTRun.json 在程序目录, 启动时移到 Data\
    static SettingsLocation() {
        eq := (n, a, e) => TestRunner.Equal("SettingsLocation." n, a, e)
        read := (f) => FileExist(f) ? FileRead(f, "UTF-8") : "<missing>"
        write := (f, text) => (FileExist(f) && FileDelete(f), FileAppend(text, f, "UTF-8"))
        eq("default file in Data", AppSettings.File, AppSettings.DataDir "\ALTRun.json")

        root := A_Temp "\ALTRunSettingsTest"
        try DirDelete(root, true)
        DirCreate(root)
        savedFile := AppSettings.File, savedLegacy := AppSettings.LegacyFile, savedData := AppSettings.Data
        AppSettings.File := root "\Data\ALTRun.json", AppSettings.LegacyFile := root "\ALTRun.json"

        ; 3.x 的布局: ALTRun.json、升级备份和 .bad 都在程序目录, 还没有 Data\
        old := AppSettings.Defaults()
        old["SchemaVersion"] := 3, old["General"]["Hotkey"] := "#Space"
        write(root "\ALTRun.json", JSON.Stringify(old))
        write(root "\ALTRun.v2.backup.json", "backup")
        write(root "\ALTRun.json.bad", "bad")
        write(root "\ALTRun.ini", "[Config]")
        AppSettings.MovedFrom := "", AppSettings.MigratedFrom := 0, AppSettings.ImportedFrom := ""
        AppSettings.Load()
        eq("moved from", AppSettings.MovedFrom, root "\ALTRun.json")
        eq("old file gone", FileExist(root "\ALTRun.json"), "")
        eq("settings kept", AppSettings.General["Hotkey"], "#Space")
        eq("not imported again", AppSettings.ImportedFrom, "")
        eq("schema upgraded", AppSettings.MigratedFrom, 3)
        eq("upgrade backup in Data", FileExist(root "\Data\ALTRun.v3.backup.json") != "", true)
        eq("old backups moved", read(root "\Data\ALTRun.v2.backup.json") "|" read(root "\Data\ALTRun.json.bad"), "backup|bad")
        eq("ini left alone", read(root "\ALTRun.ini"), "[Config]")

        ; 下一次启动: 已经在 Data\ 里, 不再移动
        AppSettings.MovedFrom := "", AppSettings.MigratedFrom := 0
        AppSettings.Load()
        eq("second start", AppSettings.MovedFrom "|" AppSettings.MigratedFrom, "|0")

        ; 两边都有 (例如退回旧版本又生成了一份): 用 Data\ 里的, 旧文件不动
        write(root "\ALTRun.json", "{}")
        AppSettings.Load()
        eq("both exist: Data wins", AppSettings.General["Hotkey"] "|" AppSettings.MovedFrom, "#Space|")
        eq("both exist: old file kept", read(root "\ALTRun.json"), "{}")

        ; 移不动 (文件被占用): 这次继续用旧位置
        FileDelete(root "\Data\ALTRun.json")
        lock := FileOpen(root "\ALTRun.json", "r -rwd")
        AppSettings.MoveLegacyFile()
        lock.Close()
        eq("locked: keep using old file", AppSettings.File "|" AppSettings.MovedFrom, root "\ALTRun.json|")

        AppSettings.File := savedFile, AppSettings.LegacyFile := savedLegacy, AppSettings.Data := savedData
        AppSettings.MovedFrom := "", AppSettings.MigratedFrom := 0, AppSettings.ImportedFrom := ""
        try DirDelete(root, true)
    }

    ; 编译信息里的版本号 (ALTRun.ahk 的 ;@Ahk2Exe-SetVersion) 要和 App.Version 一致
    static ReleaseVersion() {
        main := FileRead(A_ScriptDir "\..\ALTRun.ahk", "UTF-8")
        RegExMatch(main, "m);@Ahk2Exe-SetVersion\s+(\S+)", &m)
        TestRunner.Equal("ReleaseVersion.exe version = App.Version", IsObject(m) ? m[1] : "", App.Version)
        RegExMatch(main, "m);@Ahk2Exe-SetDescription\s+(.+?)\s*$", &m)                ; Windows 用它当显示名 (通知、任务管理器), 只写软件名
        TestRunner.Equal("ReleaseVersion.exe description = App.Name", IsObject(m) ? m[1] : "", App.Name)
        TestRunner.True("ReleaseVersion.date format", RegExMatch(App.Version, "^\d{4}\.\d{2}\.\d{2}(\.\d+)?$"))   ; 同一天再发布: 2026.09.30.1
        TestRunner.True("ReleaseVersion.same-day release is newer", UpdateChecker.Compare("2026.09.30.1", "2026.09.30") > 0 && UpdateChecker.Compare("2026.10.01", "2026.09.30.2") > 0)
    }

    ; 一键更新: 解析 GitHub 的 Release、SHA256、替换文件 (用临时文件夹里的假程序, 不下载)
    static SelfUpdate() {
        eq := (n, a, e) => TestRunner.Equal("SelfUpdate." n, a, e)
        ok := (n, c) => TestRunner.True("SelfUpdate." n, c)
        digest := "3779b0aedbd58a7bf235eab4c0cc52d3b4c9f3188cb98be38319837605cef017"
        response := '{"tag_name": "v2026.10.01", "html_url": "https://github.com/zhugecaomao/ALTRun/releases/tag/2026.10.01", "assets": ['
            . '{"name": "notes.txt", "browser_download_url": "https://x/notes.txt", "digest": null},'
            . '{"name": "ALTRun_v2026.10.01.zip", "browser_download_url": "https://x/ALTRun_v2026.10.01.zip", "digest": "sha256:' StrUpper(digest) '"}]}'
        release := UpdateChecker.ParseRelease(response)
        eq("version", release.Version, "2026.10.01")
        eq("page", release.Page, "https://github.com/zhugecaomao/ALTRun/releases/tag/2026.10.01")
        eq("zip", release.ZipUrl, "https://x/ALTRun_v2026.10.01.zip")
        eq("sha256", release.Sha256, digest)
        release := UpdateChecker.ParseRelease('{"tag_name": "2026.10.01", "assets": [{"name": "ALTRun_v2026.10.01.zip", "browser_download_url": "https://x/a.zip", "digest": null}]}')
        eq("no digest", release.Sha256 "|" release.Page, "|" UpdateChecker.ReleasePage)
        ok("no digest -> no install", !UpdateChecker.CanInstall(release))
        failed := false
        try UpdateChecker.ParseRelease('{"message": "API rate limit exceeded"}')
        catch
            failed := true
        ok("no tag throws", failed)

        root := A_Temp "\ALTRun-selfupdate-test"
        try DirDelete(root, true)
        src := root "\new", dest := root "\app"
        write := (path, text) => (DirCreate(RegExReplace(path, "\\[^\\]+$")), FileAppend(text, path, "UTF-8-RAW"))
        read := (path) => FileExist(path) ? FileRead(path, "UTF-8") : "<missing>"
        write(root "\abc.txt", "abc")
        eq("sha256 file", UpdateChecker.Sha256File(root "\abc.txt"), "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad")
        ok("writable", UpdateChecker.IsWritable(root))
        tar := A_WinDir "\System32\tar.exe"                                ; 解压 (Windows 10 起自带 tar.exe; Wine 里没有就跳过)
        if FileExist(tar) {
            write(root "\pkg\ALTRun.exe", "exe"), write(root "\pkg\Resources\Lang\zh-CN.json", "{}")
            RunWait('"' tar '" -a -cf "' root '\pkg.zip" -C "' root '\pkg" ALTRun.exe Resources', , "Hide")
            UpdateChecker.Extract(root "\pkg.zip", root "\out")
            eq("extract", read(root "\out\ALTRun.exe") "|" read(root "\out\Resources\Lang\zh-CN.json"), "exe|{}")
        }

        write(src "\ALTRun.exe", "new exe")
        write(src "\Resources\Kanji.txt", "new kanji")
        write(src "\Resources\Themes\Dark.json", "{}")
        write(src "\Resources\SDL.dll", "still shipped")
        write(src "\README.md", "readme")
        write(dest "\ALTRun.exe", "old exe")
        write(dest "\ALTRun.exe.old", "stale")
        write(dest "\Data\ALTRun.json", "settings")
        write(dest "\Data\Knowledge.json", "learned")
        write(dest "\Resources\Kanji.txt", "old kanji")
        write(dest "\Resources\DOSBox.exe", "obsolete")
        write(dest "\Resources\Mine.txt", "user file")
        UpdateChecker.Apply(src, dest, "ALTRun.exe")
        eq("exe replaced", read(dest "\ALTRun.exe"), "new exe")
        eq("old exe kept as .old", read(dest "\ALTRun.exe.old"), "old exe")
        eq("resources updated", read(dest "\Resources\Kanji.txt") "|" read(dest "\Resources\Themes\Dark.json"), "new kanji|{}")
        eq("settings untouched", read(dest "\Data\ALTRun.json") "|" read(dest "\Data\Knowledge.json"), "settings|learned")
        eq("obsolete removed, user file kept", read(dest "\Resources\DOSBox.exe") "|" read(dest "\Resources\Mine.txt"), "<missing>|user file")
        eq("obsolete but still in the package: kept", read(dest "\Resources\SDL.dll"), "still shipped")
        eq("top-level file copied", read(dest "\README.md"), "readme")
        extra := ""
        Loop Files dest "\*", "D"
            if !(A_LoopFileName = "Resources" || A_LoopFileName = "Data")
                extra .= A_LoopFileName " "
        eq("no stray folders", extra, "")

        write(dest "\Launcher.exe", "renamed old")                          ; 用户把 ALTRun.exe 改过名
        FileDelete(dest "\ALTRun.exe")
        UpdateChecker.Apply(src, dest, "Launcher.exe")
        eq("renamed exe replaced", read(dest "\Launcher.exe") "|" read(dest "\Launcher.exe.old"), "new exe|renamed old")
        eq("no extra ALTRun.exe", read(dest "\ALTRun.exe"), "<missing>")

        FileDelete(src "\ALTRun.exe")
        FileDelete(dest "\Launcher.exe.old")
        failed := false
        try UpdateChecker.Apply(src, dest, "Launcher.exe")
        catch
            failed := true
        ok("package without exe throws", failed)
        eq("package without exe changes nothing", read(dest "\Launcher.exe") "|" read(dest "\Launcher.exe.old"), "new exe|<missing>")

        ; 复制到一半失败 (目标文件被别的程序独占打开): 已经覆盖的改回去, 新增的文件和文件夹删掉, exe 不动
        try DirDelete(root, true)
        write(src "\ALTRun.exe", "new exe")
        write(src "\README.md", "new readme")
        write(src "\Resources\Kanji.txt", "new kanji")
        write(src "\Resources\New.txt", "new file")
        write(src "\Resources\Fonts\a.ttf", "font")
        write(dest "\ALTRun.exe", "old exe")
        write(dest "\README.md", "old readme")
        write(dest "\Resources\Kanji.txt", "old kanji")
        locked := FileOpen(dest "\Resources\Kanji.txt", "rw -rwd")
        failed := false
        try UpdateChecker.Apply(src, dest, "ALTRun.exe")
        catch
            failed := true
        locked.Close()
        ok("copy failure throws", failed)
        eq("copy failure: files restored", read(dest "\README.md") "|" read(dest "\Resources\Kanji.txt") "|" read(dest "\ALTRun.exe"), "old readme|old kanji|old exe")
        eq("copy failure: new files removed", read(dest "\Resources\New.txt") "|" read(dest "\Resources\Fonts\a.ttf") "|" (InStr(FileExist(dest "\Resources\Fonts"), "D") ? "dir" : "no dir"), "<missing>|<missing>|no dir")
        eq("copy failure: no .old, no backup left", read(dest "\ALTRun.exe.old") "|" (FileExist(src ".backup") ? "backup" : "none"), "<missing>|none")
        try DirDelete(root, true)
    }

    ; 使用统计: 按天、按功能计数, 合计 / 每天 / 保存和载入 (用临时文件, 不碰真正的 Data\Usage.json)
    static UsageStats() {
        eq := (n, a, e) => TestRunner.Equal("Usage." n, a, e)
        saved := {File: Usage.File, Days: Usage.Days, Since: Usage.Since}
        Usage.File := A_Temp "\ALTRun-usage-test.json"
        try FileDelete(Usage.File)
        Usage.Days := Map(), Usage.Since := ""
        try {
            Usage.Count("Show", "2026-09-20")
            Usage.Count("Applications", "2026-09-20")
            Usage.Count("Applications", "2026-09-26")
            Usage.Count("Applications", "2026-09-26")
            Usage.Count("Show", "2026-09-26")
            Usage.Count("Calculator", "2026-09-25")
            Usage.Count("Calculator", "2026-08-01")
            eq("since", Usage.Since, "2026-09-20")
            today := Usage.Summary(1, "2026-09-26")
            eq("today", today["Applications"] "|" today["Show"] "|" Usage.Total(today), "2|1|2")
            eq("7 days (from 09-20)", Usage.Total(Usage.Summary(7, "2026-09-26")), 4)
            eq("6 days (not 09-20)", Usage.Total(Usage.Summary(6, "2026-09-26")), 3)
            eq("all", Usage.Total(Usage.Summary(0)), 5)
            daily := Usage.Daily(3, "2026-09-26")
            eq("daily", daily.Length "|" daily[1][1] "|" daily[1][2] "|" daily[2][2] "|" daily[3][1] "|" daily[3][2], "3|2026-09-24|0|1|2026-09-26|2")
            eq("daily across month end", Usage.Daily(2, "2026-10-01")[1][1], "2026-09-30")

            Usage.Days := Map()
            Usage.CountItem({Provider: "CustomCommands", Kind: "file"})
            Usage.CountItem({Provider: "", Kind: "url"})                    ; 兜底的网页搜索
            Usage.CountItem({Provider: "", Kind: ""})                       ; 兜底的文件搜索
            Usage.CountItem({Provider: "Help", Kind: "url"})                ; 速查表不算
            all := Usage.Summary(0)
            eq("count item", all["CustomCommands"] "|" all["WebSearch"] "|" all["FileSearch"] "|" all.Has("Help"), "1|1|1|0")

            Usage.Days := Map("2020-01-01", Map("Calculator", 9))           ; 太旧的保存时删掉
            Usage.Count("Snippets")
            Usage.Save()
            Usage.Days := Map(), Usage.Since := ""
            Usage.Load()
            eq("save / load", Usage.Total(Usage.Summary(0)) "|" Usage.Days.Has("2020-01-01"), "1|0")
            Usage.Clear()
            eq("clear", Usage.Days.Count "|" Usage.Since, "0|")
        } finally {
            try FileDelete(Usage.File)
            Usage.File := saved.File, Usage.Days := saved.Days, Usage.Since := saved.Since
        }
    }

    ; 操作后的提示: 搜索窗口开着时在它下方居中, 否则在屏幕中间偏下; 不出屏幕
    static HudPlacement() {
        eq := (n, a, e) => TestRunner.Equal("HudPlacement." n, a, e)
        gap := Win.Scale(12)
        area := {Left: 0, Top: 0, Right: 1920, Bottom: 1040}
        pos := Hud.Position(200, 40, {Window: "", Area: area})
        eq("no window: centered", pos.X, 860)
        eq("no window: lower middle", pos.Y, 673)
        pos := Hud.Position(200, 40, {Window: {X: 610, Y: 200, W: 700, H: 400}, Area: area})
        eq("below window", pos.X "," pos.Y, "860," (600 + gap))
        pos := Hud.Position(200, 40, {Window: {X: 610, Y: 700, W: 700, H: 330}, Area: area})
        eq("above window when no room below", pos.Y, 700 - gap - 40)
        pos := Hud.Position(200, 40, {Window: {X: 1800, Y: 100, W: 700, H: 300}, Area: area})
        eq("kept on screen", pos.X, 1720)
        second := {Left: 1920, Top: 0, Right: 3840, Bottom: 1080}
        eq("second monitor", Hud.Position(200, 40, {Window: "", Area: second}).X, 2780)

        long := ""
        Loop 40
            long .= "long text "
        Hud.Show(long, 300)                                                 ; 超过最大宽度: 自动换行, 不报错
        Hud.Gui.GetPos(, , &longW)
        eq("long text wraps", longW <= Win.Scale(Hud.MaxWidth) + Win.Scale(60), true)
        LargeType.Show(long long long)
        eq("large type long text", IsObject(LargeType.Gui), true)
        LargeType.Close()
        Hud.Show("Copied", 300)
        eq("shown", IsObject(Hud.Gui), true)
        Hud.Show("Second", 300)
        eq("replaced, one window", WinExist("ALTRun HUD ahk_class AutoHotkeyGUI") = Hud.Gui.Hwnd, true)
        Hud.Hide()
        eq("hidden", Hud.Gui, "")
    }

    ; 后台检查更新: 每天一次、跳过的版本、搜索窗口里的更新提示 (不联网: 直接设置 Pending)
    static UpdateNotice() {
        eq := (n, a, e) => TestRunner.Equal("UpdateNotice." n, a, e)
        ok := (n, c) => TestRunner.True("UpdateNotice." n, c)
        ok("due: never checked", UpdateChecker.IsDue("", "20260928120000"))
        ok("not due: 5 h", !UpdateChecker.IsDue("20260928070000", "20260928120000"))
        ok("due: 6 h", UpdateChecker.IsDue("20260928060000", "20260928120000"))
        ok("startup: due after 1 h", UpdateChecker.IsDue("20260928110000", "20260928120000", UpdateChecker.StartupHours))
        ok("startup: not due within 1 h", !UpdateChecker.IsDue("20260928113000", "20260928120000", UpdateChecker.StartupHours))
        ok("due: clock moved back", UpdateChecker.IsDue("20260929120000", "20260928120000"))
        ok("due: bad value", UpdateChecker.IsDue("garbage", "20260928120000"))
        ok("wanted: newer", UpdateChecker.IsWanted("2026.10.01", "2026.09.28", ""))
        ok("not wanted: same", !UpdateChecker.IsWanted("2026.09.28", "2026.09.28", ""))
        ok("not wanted: skipped", !UpdateChecker.IsWanted("2026.10.01", "2026.09.28", "2026.10.01"))
        ok("wanted: newer than skipped", UpdateChecker.IsWanted("2026.10.02", "2026.09.28", "2026.10.01"))

        ; 状态文件
        savedFile := UpdateChecker.StateFile
        root := A_Temp "\ALTRun-test-update-state"
        try DirDelete(root, true)
        UpdateChecker.StateFile := root "\Data\Update.json"
        savedProviders := ProviderRegistry.Providers
        ProviderRegistry.Providers := [SystemProvider]
        eq("state: empty", UpdateChecker.LoadState()["LastCheck"] "|" UpdateChecker.LoadState()["Skip"], "|")
        UpdateChecker.SaveState(Map("LastCheck", "20260928120000", "Skip", "2026.10.01"))
        state := UpdateChecker.LoadState()
        eq("state: saved", state["LastCheck"] "|" state["Skip"], "20260928120000|2026.10.01")

        ; 没有新版本: 空搜索框没有结果
        UpdateChecker.Pending := ""
        eq("no update: empty query", ProviderRegistry.Search("").Length, 0)
        eq("no update: item", UpdateChecker.PendingItem(), "")

        ; 有新版本: 空搜索框和搜索 "update" 都显示, 排在 "检查更新" 前面
        release := {Version: "2026.10.02", Page: "https://x/notes", ZipUrl: "https://x/a.zip", Sha256: "ab"}
        UpdateChecker.SetPending(release)
        eq("mode (tests run the source)", release.Mode, "source")
        results := ProviderRegistry.Search("")
        eq("empty query: one item", results.Length, 1)
        eq("empty query: title", results[1].Title, "Update Available: ALTRun 2026.10.02")
        eq("empty query: provider", results[1].Provider, "System")
        ok("source hint", InStr(results[1].Subtitle, "git pull"))
        eq("actions: notes + skip", results[1].Actions.Length, 2)
        titles := ""
        for action in ActionCatalog.ListFor(results[1])
            titles .= action.Title "|"
        ok("action panel", InStr(titles, "Release Notes|Skip This Version|"))
        results := ProviderRegistry.Search("update")
        ok("search update: first", results.Length >= 2 && results[1].Title = "Update Available: ALTRun 2026.10.02")
        ; 操作面板右侧的说明: 放得下时完整显示, 放不下时名称优先 (最多一半), 说明用剩下的
        eq("hint fits", SearchWindow.HintWidth(600, 120, 300, 12), 300)
        eq("hint shares the row", SearchWindow.HintWidth(400, 120, 600, 12), 268)
        eq("long title keeps half", SearchWindow.HintWidth(400, 500, 600, 12), 188)
        eq("no room", SearchWindow.HintWidth(20, 100, 50, 12), 0)
        for mode, text in Map("install", "Install Update", "page", "Download Page", "scoop", "scoop update altrun", "winget", "winget upgrade")
            release.Mode := mode, ok("subtitle " mode, InStr(UpdateChecker.PendingItem().Subtitle, text))

        ; 跳过这个版本: 不再显示, 记在状态文件里
        UpdateChecker.SkipVersion("2026.10.02")
        eq("skipped: no item", ProviderRegistry.Search("").Length, 0)
        eq("skipped: saved", UpdateChecker.LoadState()["Skip"], "2026.10.02")
        Hud.Hide()
        UpdateChecker.StateFile := savedFile
        ProviderRegistry.Providers := savedProviders
        try DirDelete(root, true)
    }

    ; 热键的显示 (Alt+Space) 和录制框的规则; 偏好设置保存前检查重复的全局热键
    static HotkeyText() {
        eq := (n, a, e) => TestRunner.Equal("HotkeyText." n, a, e)
        ok := (n, c) => TestRunner.True("HotkeyText." n, c)
        for pair in [["!Space", "Alt+Space"], ["^!c", "Ctrl+Alt+C"], ["!^c", "Ctrl+Alt+C"], ["#e", "Win+E"], ["^+F12", "Ctrl+Shift+F12"]
                   , ["~MButton", "MButton"], ["$*F5", "F5"], ["CapsLock & j", "CapsLock+J"], ["^+", "Ctrl++"], ["!r", "Alt+R"]
                   , ["<^>!a", "Ctrl+Alt+A"], ["F1 up", "F1 Up"], ["", ""], ["None", ""]]
            eq("label " pair[1], Win.HotkeyLabel(pair[1]), pair[2])
        eq("names", Win.HotkeyLabel("~^MButton", Map("MButton", "Middle")), "Ctrl+Middle")
        eq("box label", HotkeyBox.Label("~MButton"), "Middle mouse button")
        eq("box label XButton1", HotkeyBox.Label("^XButton1"), "Ctrl+Mouse back button")

        eq("compose order", HotkeyBox.Compose("#+!^", "C"), "^!+#c")
        eq("compose name", HotkeyBox.Compose("!", "Space"), "!Space")
        eq("compose none", HotkeyBox.Compose("", "F5"), "F5")
        ; 录制时的修饰键以 InputHook 收到的按下 / 松开为准 (远程桌面模拟的按键不算物理按下)
        savedActive := HotkeyBox._active
        HotkeyBox._active := {Hwnd: 0}
        HotkeyBox._held := Map("RControl", true, "LAlt", true)
        eq("held mods", HotkeyBox._Mods(), "^!")
        HotkeyBox._held := Map("Control", true, "LWin", true, "RShift", true)
        eq("held generic Control", HotkeyBox._Mods(), "^+#")
        HotkeyBox._held := Map()
        eq("nothing held", HotkeyBox._Mods(), "")
        HotkeyBox._active := savedActive
        ok("modifier names", HotkeyBox.IsModifier("LControl") && HotkeyBox.IsModifier("RWin") && HotkeyBox.IsModifier("Shift") && !HotkeyBox.IsModifier("k"))
        for hk in ["^!c", "!Space", "#e", "F5", "^Numpad1", "+F3", "Pause", "MButton", "!Enter"]
            ok("allowed " hk, HotkeyBox.IsAllowed(hk))
        for hk in ["a", "+a", "Space", "+Space", "Enter", "Tab", "Numpad5", "1", "Delete"]
            ok("not allowed " hk, !HotkeyBox.IsAllowed(hk))
        ; 关掉的窗口里的框要去掉: 句柄会被重新使用, 留着的话在别的控件上点中键会被当成录制
        g := Gui()
        box := HotkeyBox.Add(g, "w120", "^!k", true)
        boxHwnd := box.Hwnd, state := HotkeyBox.Boxes[boxHwnd]
        ok("box alive", HotkeyBox.IsAlive(state))
        g.Destroy()
        ok("box destroyed", !HotkeyBox.IsAlive(state))
        eq("mouse on destroyed box", HotkeyBox._OnMouse(0, 0x207, boxHwnd), "")
        ok("destroyed box pruned", !HotkeyBox.Boxes.Has(boxHwnd))
        g := Gui()
        HotkeyBox.Add(g, "w120", "", false)
        HotkeyBox.Boxes[boxHwnd] := state                                   ; 下一次 Add 时也会清理
        HotkeyBox.Add(g, "w120", "", false)
        ok("pruned on add", !HotkeyBox.Boxes.Has(boxHwnd) || HotkeyBox.Boxes[boxHwnd] != state)
        g.Destroy()
        HotkeyBox.Prune()

        ; 偏好设置里重复的全局热键
        list := [{Key: "!Space", Label: "A"}, {Key: "!r", Label: "B"}, {Key: "~!R", Label: "C"}]
        clash := PreferencesWindow.DuplicateHotkey(list)
        eq("duplicate", IsObject(clash) ? clash.First "|" clash.Second : "", "B|C")
        eq("duplicate order", PreferencesWindow.DuplicateHotkey([{Key: "^!c", Label: "A"}, {Key: "!^c", Label: "B"}]).Second, "B")
        eq("no duplicate", PreferencesWindow.DuplicateHotkey([{Key: "!Space", Label: "A"}, {Key: "!r", Label: "B"}]), "")
        data := AppSettings.Defaults()
        data["Hotkeys"] := [Map("Key", "^!c", "Action", "Lock", "WinTitle", ""), Map("Key", "!Space", "Action", "Lock", "WinTitle", "ahk_exe notepad.exe")]
        keys := ""
        for entry in PreferencesWindow._GlobalHotkeys(data)
            keys .= entry.Key "|"
        eq("global hotkeys (window-limited left out)", keys, "!Space|!r|^!\|^!c|^!c|")
        ok("default clipboard hotkey clashes with a global ^!c", IsObject(PreferencesWindow.DuplicateHotkey(PreferencesWindow._GlobalHotkeys(data))))
        ok("defaults have no clash", !IsObject(PreferencesWindow.DuplicateHotkey(PreferencesWindow._GlobalHotkeys(AppSettings.Defaults()))))
        eq("list cell", PreferencesWindow._Cell(Map("Key", "~MButton"), "Key"), "Middle mouse button")
    }

    ; JSON: 转义、数字、嵌套、往返; 大文件读取是线性的 (以前按值传整段文本, 越大越慢)
    static JsonReadWrite() {
        eq := (n, a, e) => TestRunner.Equal("Json." n, a, e)
        ok := (n, c) => TestRunner.True("Json." n, c)
        data := JSON.Parse('{"a": "x\"y", "b": [1, -2.5, 3e2, true, false, null], "c": {"d": "\u4e2d\n\t\/\\"}, "e": "", "f": "C:\\Temp\\"}')
        eq("escaped quote", data["a"], 'x"y')
        eq("numbers", data["b"][1] "|" data["b"][2] "|" data["b"][3], "1|-2.5|300.0")
        eq("true false null", data["b"][4] "|" data["b"][5] "|" data["b"][6], "1|0|")
        eq("escapes", data["c"]["d"], "中`n`t/\")
        eq("empty string", data["e"], "")
        eq("backslash before quote", data["f"], "C:\Temp\")
        eq("empty containers", JSON.Stringify(JSON.Parse('{"x": [], "y": {}}')), '{`r`n  "x": [],`r`n  "y": {}`r`n}')
        original := Map("Text", 'line1`r`nline2 "quoted" \ back`ttab', "List", [1, "two", Map("k", "v")])
        again := JSON.Parse(JSON.Stringify(original))
        eq("round trip text", again["Text"], original["Text"])
        eq("round trip nested", again["List"][3]["k"], "v")
        for bad in ['{"a": 1', '{"a" 1}', '[1, 2', '"open', '{a: 1}', '[1 2]', '@']
            ok("error: " bad, !JsonReadWrite_Parses(bad))

        ; 控制字符 (例如从终端复制的 ESC) 写成 \u001B, 文件里不出现原样的控制字符
        control := "a" Chr(27) "[0m" Chr(8) Chr(12) Chr(1) "z"
        written := JSON.Stringify(Map("Text", control))
        ok("control chars escaped", !RegExMatch(written, "[\x00-\x08\x0B\x0C\x0E-\x1F]") && InStr(written, "\u001B[0m\u0008\u000C\u0001z"))
        eq("control chars round trip", JSON.Parse(written)["Text"], control)
        ; 写文件: 先写临时文件再替换, 不带 BOM, 不留下 .tmp
        dir := A_Temp "\ALTRun-json-test"
        try DirDelete(dir, true)
        jsonFile := dir "\sub\data.json"
        JSON.WriteFile(jsonFile, Map("k", "中文"))
        JSON.WriteFile(jsonFile, Map("k", "v2"))
        firstByte := FileOpen(jsonFile, "r").RawRead(bytes := Buffer(3), 3) ? NumGet(bytes, 0, "UChar") : 0
        eq("write file", JSON.Parse(FileRead(jsonFile, "UTF-8"))["k"] "|" (firstByte = 0xEF ? "bom" : "no bom") "|" (FileExist(jsonFile ".tmp") ? "tmp" : ""), "v2|no bom|")
        DirDelete(dir, true)

        paths := []
        Loop 20000
            paths.Push("C:\Users\someone\Documents\Folder" A_Index "\Report " A_Index ".docx")
        text := JSON.Stringify(Map("Paths", paths))
        start := A_TickCount
        parsed := JSON.Parse(text)
        elapsed := A_TickCount - start
        eq("big file", parsed["Paths"].Length "|" parsed["Paths"][20000], "20000|" paths[20000])
        ok("big file is fast (" elapsed " ms)", elapsed < 3000)             ; 以前要好几分钟
    }

    ; 单位 / 货币换算 (计算器)
    static UnitConversion() {
        eq := (n, a, e) => TestRunner.Equal("Units." n, a, e)
        ok := (n, c) => TestRunner.True("Units." n, c)
        r := Units.Parse("10 km in mi")
        eq("parse", r.Value "|" r.From "|" r.To, "10|km|mi")
        r := Units.Parse("5ft to cm")
        eq("parse no space", r.From "|" r.To, "ft|cm")
        r := Units.Parse("1,500 kg->t")
        eq("parse arrow + thousands", r.Value "|" r.From "|" r.To, "1500|kg|t")
        r := Units.Parse("3亩转平方米")
        eq("parse chinese", r.From "|" r.To, "亩|平方米")
        eq("not a conversion", Units.Parse("notepad"), "")
        eq("not a conversion 2", Units.Parse("12*3"), "")
        fmt := (v) => CalculatorProvider.Format(v)
        eq("km mi", fmt(Units.Convert(10, "km", "mi")), "6.21")
        eq("f c", fmt(Units.Convert(212, "f", "c")), "100")
        eq("c f", fmt(Units.Convert(-40, "°c", "f")), "-40")
        eq("c k", fmt(Units.Convert(0, "C", "K")), "273.15")
        eq("mpa psi", fmt(Units.Convert(1, "MPa", "psi")), "145.04")
        eq("n/mm2 mpa", fmt(Units.Convert(30, "n/mm2", "mpa")), "30")
        eq("kn kip", fmt(Units.Convert(100, "kN", "kip")), "22.48")
        eq("mu m2", fmt(Units.Convert(1, "亩", "m²")), "666.67")
        eq("gb mb", fmt(Units.Convert(1, "GB", "MB")), "1,024")
        eq("mph km/h", fmt(Units.Convert(60, "mph", "km/h")), "96.56")
        eq("lb kg", fmt(Units.Convert(1, "lb", "kg")), "0.45")
        eq("different kinds", Units.Convert(1, "km", "kg"), "")
        eq("unknown unit", Units.Convert(1, "km", "xyz"), "")
        eq("label", Units.Label("m2"), "m²")

        ; 计算器里的结果
        results := CalculatorProvider.Search(SearchQuery("10 km in mi"))
        eq("calc result", results.Length ? results[1].Title "|" results[1].Arg : "", "6.21 mi|6.21")
        ok("calc subtitle", results.Length && InStr(results[1].Subtitle, "10 km = 6.21 mi"))
        saved := Units.Rates
        Units.Rates := Map()
        results := CalculatorProvider.Search(SearchQuery("100 usd to sgd"))
        ok("currency off hint", results.Length = 1 && !results[1].Valid && InStr(results[1].Subtitle, "Preferences"))
        CurrencyRates.Apply(JSON.Parse('{"amount": 1.0, "base": "EUR", "date": "2026-09-26", "rates": {"USD": 1.10, "SGD": 1.43, "CNY": 7.8}}'))
        eq("rates date", CurrencyRates.Date, "2026-09-26")
        eq("usd sgd", fmt(Units.Convert(100, "usd", "sgd")), "130")
        eq("eur base", fmt(Units.Convert(10, "EUR", "CNY")), "78")
        eq("chinese currency name", fmt(Units.Convert(110, "美元", "人民币")), "780")
        results := CalculatorProvider.Search(SearchQuery("100 usd to sgd"))
        ok("currency result", results.Length && results[1].Title = "130 SGD" && InStr(results[1].Subtitle, "2026-09-26"))
        Units.Rates := saved
        ok("stale: never", CurrencyRates.IsStale("", "20260929120000"))
        ok("stale: 13 h", CurrencyRates.IsStale("20260928230000", "20260929120000"))
        ok("fresh: 2 h", !CurrencyRates.IsStale("20260929100000", "20260929120000"))
    }

    ; 片段的占位符
    static SnippetPlaceholders() {
        eq := (n, a, e) => TestRunner.Equal("Placeholders." n, a, e)
        ok := (n, c) => TestRunner.True("Placeholders." n, c)
        eq("date format", SnippetProvider.Expand("{date:yyyy}"), FormatTime(, "yyyy"))
        eq("date plus", SnippetProvider.Expand("{date+7:yyyyMMdd}"), FormatTime(DateAdd(A_Now, 7, "Days"), "yyyyMMdd"))
        eq("date minus", SnippetProvider.Expand("x {date-1:yyyyMMdd} y"), "x " FormatTime(DateAdd(A_Now, -1, "Days"), "yyyyMMdd") " y")
        eq("default date", SnippetProvider.Expand("{date}"), FormatTime(, AppSettings.Extension("AutoDate")["DateFormat"]))
        eq("time format", StrLen(SnippetProvider.Expand("{time:HH:mm:ss}")), 8)
        eq("cursor untouched", SnippetProvider.Expand("a{cursor}b"), "a{cursor}b")
        uuids := SnippetProvider.Expand("{uuid} {uuid}")
        ok("uuid format", RegExMatch(uuids, "^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12} [0-9a-f-]{36}$"))
        ok("uuid different", StrSplit(uuids, " ")[1] != StrSplit(uuids, " ")[2])
        saved := ClipboardProvider.Entries
        ClipboardProvider.Entries := [Map("Text", "newest"), Map("Text", "previous"), Map("Text", "older")]
        eq("clipboard:1", SnippetProvider.Expand("[{clipboard:1}] [{clipboard:2}] [{clipboard:9}]"), "[previous] [older] []")
        ClipboardProvider.Entries := saved
    }

    ; 片段: 也搜索正文
    static SnippetTextSearch() {
        eq := (n, a, e) => TestRunner.Equal("SnippetText." n, a, e)
        ok := (n, c) => TestRunner.True("SnippetText." n, c)
        options := AppSettings.Feature("Snippets")
        savedSnippets := AppSettings.Snippets, savedSearch := options["SearchText"]
        long := "Dear Sir, " StrRepeat("please see the attached drawings. ", 4) "We offer a goodwill discount on the revised quotation."
        AppSettings.Data["Snippets"] := [Map("Name", "Quote reply", "Keyword", "qr", "Text", long)
                               , Map("Name", "Signature", "Keyword", "sig", "Text", "Best regards")
                               , Map("Name", "Discount note", "Keyword", "dn", "Text", "No text match here")]
        titles(text) {
            list := ""
            for item in SnippetProvider.Search(SearchQuery(text))
                list .= item.Title "=" Round(item.Score) "|"
            return list
        }
        options["SearchText"] := 1
        list := titles("goodwill")
        eq("text match", list, "Quote reply=" SnippetProvider.TextScore "|")
        ok("case-insensitive", InStr(titles("GOODWILL"), "Quote reply="))
        ok("every word", InStr(titles("revised goodwill"), "Quote reply="))
        ok("missing word", !InStr(titles("revised tender"), "Quote reply="))
        ok("too short", !InStr(titles("se"), "Quote reply="))
        scores := Map()
        for item in SnippetProvider.Search(SearchQuery("discount"))
            scores[item.Title] := item.Score
        ok("name match ranks higher", scores.Has("Discount note") && scores.Has("Quote reply") && scores["Discount note"] > scores["Quote reply"])
        eq("text-only score", scores.Has("Quote reply") ? scores["Quote reply"] : "", SnippetProvider.TextScore)
        ok("no fuzzy on text", !InStr(titles("dsgd"), "Quote reply="))
        item := SnippetProvider.Search(SearchQuery("goodwill"))[1]
        ok("preview shows match", InStr(item.Subtitle, "...") && InStr(item.Subtitle, "goodwill"))
        item := SnippetProvider.Search(SearchQuery("Best reg"))[1]
        eq("preview from start", item.Subtitle, I18n.T("Snippet.Subtitle", "Best regards"))
        ok("snip keyword boosts", InStr(titles("snip goodwill"), "Quote reply=" (SnippetProvider.TextScore + 30) "|"))
        options["SearchText"] := 0
        eq("switched off", titles("goodwill"), "")
        options["SearchText"] := savedSearch
        AppSettings.Data["Snippets"] := savedSnippets
        eq("match position", SnippetProvider.TextMatch("abc def", ["DEF", "abc"]), 5)
        eq("no match", SnippetProvider.TextMatch("abc def", ["abc", "xyz"]), 0)
    }

    ; 片段编辑框: 每个字段都有灰色说明; 关键字太长 / 有空格时提示
    static SnippetEditor() {
        eq := (n, a, e) => TestRunner.Equal("SnippetEditor." n, a, e)
        ok := (n, c) => TestRunner.True("SnippetEditor." n, c)
        for field in SnippetProvider.EditorFields()
            ok("hint: " field.Key, field.HasOwnProp("Hint") && field.Hint != "")
        keyword := SnippetProvider.EditorFields()[2]
        ok("keyword has check", keyword.HasOwnProp("Check"))
        ok("check as ItemEditor calls it", keyword.Check.Call(StrRepeat("a", 40), Map("AutoExpand", 1)) != "")
        eq("check short keyword", keyword.Check.Call("sig", Map("AutoExpand", 1)), "")
        ok("hint shows limit", InStr(keyword.Hint, SnippetExpander.MaxAbbreviation - StrLen(AppSettings.Feature("Snippets")["ExpandPrefix"])))
        longKeyword := StrRepeat("a", 40)
        eq("short ok", SnippetProvider.CheckKeyword("sig", Map("AutoExpand", 1), ";"), "")
        eq("empty ok", SnippetProvider.CheckKeyword("", Map("AutoExpand", 1), ";"), "")
        ok("too long", SnippetProvider.CheckKeyword(longKeyword, Map("AutoExpand", 1), ";") != "")
        eq("39 + prefix ok", SnippetProvider.CheckKeyword(StrRepeat("a", 39), Map("AutoExpand", 1), ";"), "")
        ok("space", SnippetProvider.CheckKeyword("a b", Map("AutoExpand", 1), ";") != "")
        eq("no auto-expand", SnippetProvider.CheckKeyword(longKeyword, Map("AutoExpand", 0), ";"), "")
        eq("too long not registered", SnippetExpander.Abbreviation(Map("Keyword", longKeyword, "Text", "x"), ";"), "")
        eq("longest registered", SnippetExpander.Abbreviation(Map("Keyword", StrRepeat("a", 39), "Text", "x"), ";"), ";" StrRepeat("a", 39))
        store := ResultItem("Notepad", "", {Kind: "file", Arg: "shell:AppsFolder\Microsoft.WindowsNotepad_8wekyb3d8bbwe!App"})
        list := ""
        for action in ActionCatalog.ListFor(store)
            list .= action.Title "|"
        ok("store app: no properties", !InStr(list, I18n.T("Action.Properties") "|"))
        list := ""
        for action in ActionCatalog.ListFor(ResultItem("Windows", "", {Kind: "folder", Arg: A_WinDir}))
            list .= action.Title "|"
        ok("folder: properties", InStr(list, I18n.T("Action.Properties") "|"))
    }

    ; 浏览器书签
    static Bookmarks() {
        eq := (n, a, e) => TestRunner.Equal("Bookmarks." n, a, e)
        ok := (n, c) => TestRunner.True("Bookmarks." n, c)
        sample := '{"roots": {"bookmark_bar": {"type": "folder", "children": ['
            . '{"type": "url", "name": "GitHub", "url": "https://github.com/"},'
            . '{"type": "folder", "name": "Work", "children": [{"type": "url", "name": "Tender Portal", "url": "https://www.tenders.example.com/login"}]},'
            . '{"type": "url", "name": "Bookmarklet", "url": "javascript:alert(1)"}]},'
            . '"other": {"type": "folder", "children": [{"type": "url", "name": "", "url": "http://intranet.local/"}]}}}'
        items := BookmarkProvider.ParseFile(sample)
        eq("count (http only)", items.Length, 3)
        eq("nested", items[2].Title, "Tender Portal")
        eq("empty name uses url", items[3].Title, "http://intranet.local/")
        savedItems := BookmarkProvider.Items, savedChecked := BookmarkProvider._checked
        BookmarkProvider.Items := items, BookmarkProvider._checked := A_TickCount   ; 不去读本机的书签文件
        results := BookmarkProvider.Search(SearchQuery("tender"))
        eq("default results", results.Length ? results[1].Arg : "", "https://www.tenders.example.com/login")
        ok("default results rank below commands", results.Length && results[1].Score < 100 && !results[1].Exclusive)
        results := BookmarkProvider.Search(SearchQuery("bm tenders"))
        ok("keyword + host", results.Length = 1 && results[1].Exclusive)
        eq("too short", BookmarkProvider.Search(SearchQuery("g")).Length, 0)
        ; 书签文件读不了 (浏览器正在写) 时不记下修改时间, 下次搜索时再读
        root := A_Temp "\ALTRun-test-bookmarks"
        try DirDelete(root, true)
        DirCreate(root "\Google\Chrome\User Data\Default")
        bookmarkPath := root "\Google\Chrome\User Data\Default\Bookmarks"
        FileAppend('{"roots": {"bookmark_bar": {"type": "folder", "children": [', bookmarkPath, "UTF-8")
        savedLocal := EnvGet("LOCALAPPDATA"), savedStamp := BookmarkProvider._stamp
        EnvSet("LOCALAPPDATA", root)
        BookmarkProvider.Load()
        eq("broken file not stamped", BookmarkProvider._stamp, "")
        eq("broken file no items", BookmarkProvider.Items.Length, 0)
        FileDelete(bookmarkPath)
        FileAppend(sample, bookmarkPath, "UTF-8")
        BookmarkProvider.Load()
        ok("good file stamped", InStr(BookmarkProvider._stamp, "\Default\Bookmarks"))   ; A_Temp 可能是 8.3 短路径, 不比较整个路径
        eq("good file items", BookmarkProvider.Items.Length, 3)
        EnvSet("LOCALAPPDATA", savedLocal)
        try DirDelete(root, true)
        BookmarkProvider.Items := savedItems, BookmarkProvider._checked := savedChecked, BookmarkProvider._stamp := savedStamp
        ids := ""
        for command in SystemProvider.Commands()
            ids .= command["Id"] "|"
        ok("media commands", InStr(ids, "MediaPlayPause|MediaNext|MediaPrev|MediaStop|"))
    }

    ; 选中内容的操作: 按内容类型生成操作面板的来源
    static SelectionItems() {
        eq := (n, a, e) => TestRunner.Equal("Selection." n, a, e)
        ok := (n, c) => TestRunner.True("Selection." n, c)
        titles(item) {
            text := ""
            for action in ActionCatalog.ListFor(item)
                text .= action.Title "|"
            return text
        }
        item := SelectionActions.ItemFor({Text: "hello world"})
        eq("text kind", item.Kind "|" item.Arg "|" item.Provider, "text|hello world|Selection")
        list := titles(item)
        ok("text: copy first", InStr(list, "Copy to Clipboard|") = 1)
        ok("text: search engines", InStr(list, "Search Google|") && InStr(list, "Search GitHub|"))
        ok("text: save as snippet", InStr(list, "Save as Snippet...|"))
        ok("text: replace with uppercase", InStr(list, "Replace with: UPPERCASE|"))
        ok("text: no calculation", !InStr(list, "= "))
        list := titles(SelectionActions.ItemFor({Text: "12*(3+4)"}))
        ok("math: result first extra", InStr(list, "= 84|"))
        item := SelectionActions.ItemFor({Text: "  https://github.com/zhugecaomao/ALTRun  "})
        eq("link", item.Kind "|" item.Arg, "url|https://github.com/zhugecaomao/ALTRun")
        item := SelectionActions.ItemFor({Text: A_WinDir})
        eq("path text -> folder", item.Kind "|" item.Arg, "folder|" A_WinDir)
        item := SelectionActions.ItemFor({Files: [A_WinDir "\notepad.exe"]})
        eq("one file", item.Kind "|" item.Title, "file|notepad.exe")
        item := SelectionActions.ItemFor({Files: ["C:\a\one.txt", "C:\b\two.txt"]})
        eq("files", item.Title "|" item.Subtitle, "2 files|one.txt, two.txt")
        ok("files: add all", InStr(titles(item), "Add All to Custom Commands|"))
        eq("transform", SelectionActions.Transforms()[1][2].Call("abc"), "ABC")
    }

    static Misc() {
        TestRunner.True("Input.layout of this thread", Win.KeyboardLayout() != 0)
        TestRunner.True("Input.per-window setting is a flag", Win.PerWindowInputMethod() = true || Win.PerWindowInputMethod() = false)
        saved := AppSettings.General["SwitchToEnglishInput"]
        AppSettings.General["SwitchToEnglishInput"] := 0
        SearchWindow._layoutBefore := 0x08040804
        SearchWindow._RestoreInputLanguage()                                 ; 没开这个选项: 什么也不做, 只清掉记录
        TestRunner.Equal("Input.restore clears the saved layout", SearchWindow._layoutBefore, 0)
        AppSettings.General["SwitchToEnglishInput"] := saved
        TestRunner.True("App.Website", RegExMatch(App.Website, "^https://\S+/$") > 0)
        TestRunner.True("UpdateChecker.newer", UpdateChecker.Compare("2026.10.01", "2026.09.23") > 0)
        TestRunner.True("UpdateChecker.same", UpdateChecker.Compare("2026.09.23", "2026.09.23") = 0)
        TestRunner.Equal("UpdateChecker.scoop", UpdateChecker.InstalledBy("C:\Users\me\scoop\apps\altrun\current"), "scoop")
        TestRunner.Equal("UpdateChecker.scoop version dir", UpdateChecker.InstalledBy("D:\Scoop\apps\ALTRun\2026.09.25"), "scoop")
        TestRunner.Equal("UpdateChecker.winget", UpdateChecker.InstalledBy("C:\Users\me\AppData\Local\Microsoft\WinGet\Packages\zhugecaomao.ALTRun_Microsoft.Winget.Source_8wekyb3d8bbwe"), "winget")
        TestRunner.Equal("UpdateChecker.by hand", UpdateChecker.InstalledBy("C:\Tools\ALTRun"), "")
        TestRunner.Equal("UpdateChecker.other altrun folder", UpdateChecker.InstalledBy("C:\apps\altrun\sub\deeper"), "")
        TestRunner.Equal("I18n.args", I18n.T("Web.SearchFor", "Google", "x"), "Search Google for 'x'")
        TestRunner.Equal("I18n.missing", I18n.T("No.Such.Key"), "No.Such.Key")
        TestRunner.True("System commands", SystemProvider.Commands().Length > 40)
        TestRunner.True("System search", SystemProvider.Search(SearchQuery("lock")).Length >= 1)
        TestRunner.Equal("Color", Win.ColorToBgr("#112233"), 0x332211)
    }
}

JsonReadWrite_Parses(text) {
    try {
        JSON.Parse(text)
        return true
    }
    return false
}

StrRepeat(text, count) => StrReplace(Format("{:" count "}", ""), " ", text)
