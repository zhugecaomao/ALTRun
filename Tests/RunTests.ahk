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
#Include %A_ScriptDir%\..\Lib\Dialogs.ahk
#Include %A_ScriptDir%\..\Src\Core\App.ahk
#Include %A_ScriptDir%\..\Src\Core\I18n.ahk
#Include %A_ScriptDir%\..\Src\Core\AppSettings.ahk
#Include %A_ScriptDir%\..\Src\Core\SchemaMigration.ahk
#Include %A_ScriptDir%\..\Src\Core\SearchQuery.ahk
#Include %A_ScriptDir%\..\Src\Core\ResultItem.ahk
#Include %A_ScriptDir%\..\Src\Core\FuzzyMatcher.ahk
#Include %A_ScriptDir%\..\Src\Core\Knowledge.ahk
#Include %A_ScriptDir%\..\Src\Core\ActionCatalog.ahk
#Include %A_ScriptDir%\..\Src\Core\ProviderRegistry.ahk
#Include %A_ScriptDir%\..\Src\UI\ThemeManager.ahk
#Include %A_ScriptDir%\..\Src\UI\IconCache.ahk
#Include %A_ScriptDir%\..\Src\UI\SearchWindow.ahk
#Include %A_ScriptDir%\..\Src\UI\LargeType.ahk
#Include %A_ScriptDir%\..\Src\Providers\ApplicationProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\CustomCommandProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\SnippetProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\SystemProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\CalculatorProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\WebSearchProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\FileSearchProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\TerminalProvider.ahk
#Include %A_ScriptDir%\..\Src\Extensions\QuickSwitch.ahk
#Include %A_ScriptDir%\..\Src\Extensions\AutoDate.ahk
#Include %A_ScriptDir%\..\Src\Extensions\PTToolsWindow.ahk
#Include %A_ScriptDir%\..\Src\Extensions\UpdateChecker.ahk

Logger.Enabled := false
I18n.Init("en")
AppSettings.Data := AppSettings.Defaults()                                  ; 内存里的默认设置, 不读写文件

TestRunner.Run()

class TestRunner {
    static Passed := 0, Failed := 0

    static Run() {
        for name in ["FuzzyMatcher", "SearchQuery", "SchemaMigration", "Calculator", "WebSearch"
                    , "AutoDate", "TextTools", "Sorting", "Knowledge", "Misc"] {
            try {
                Tests.%name%()
            } catch as e {
                TestRunner.Fail(name, "exception: " e.Message " (line " e.Line ")")
            }
        }
        FileAppend("`n" TestRunner.Passed " passed, " TestRunner.Failed " failed`n", "*")
        ExitApp(TestRunner.Failed)
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
        eq("version", data["SchemaVersion"], 3)
        eq("hotkey", data["General"]["Hotkey"], "!Space")
        eq("hotkey2", data["General"]["SecondaryHotkey"], "!r")
        eq("startup", data["General"]["LaunchAtLogin"], 0)
        eq("language", data["General"]["Language"], "zh")
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
        eq("v3 detect", SchemaMigration.DetectVersion(data), 3)
        TestRunner.True("Migration.defaults filled", !AppSettings._MergeDefaults(data, AppSettings.Defaults()))
        TestRunner.True("Migration.no downgrade", SchemaMigration.Upgrade(Map("SchemaVersion", 3), 3)["SchemaVersion"] = 3)
    }

    static Calculator() {
        eq := (n, a, e) => TestRunner.Equal("Calculator." n, a, e)
        eq("basic", CalculatorProvider.Search(SearchQuery("12*(3+4)"))[1].Title, "84")
        eq("prefix =", CalculatorProvider.Search(SearchQuery("=2^10"))[1].Title, "1,024")
        eq("decimal", CalculatorProvider.Search(SearchQuery("10/4"))[1].Title, "2.5")
        eq("not math", CalculatorProvider.Search(SearchQuery("notepad")).Length, 0)
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

    static Misc() {
        TestRunner.True("UpdateChecker.newer", UpdateChecker.Compare("2026.10.01", "2026.09.23") > 0)
        TestRunner.True("UpdateChecker.same", UpdateChecker.Compare("2026.09.23", "2026.09.23") = 0)
        TestRunner.Equal("I18n.args", I18n.T("Web.SearchFor", "Google", "x"), "Search Google for 'x'")
        TestRunner.Equal("I18n.missing", I18n.T("No.Such.Key"), "No.Such.Key")
        TestRunner.True("System commands", SystemProvider.Commands().Length > 40)
        TestRunner.True("System search", SystemProvider.Search(SearchQuery("lock")).Length >= 1)
        TestRunner.Equal("Color", Win.ColorToBgr("#112233"), 0x332211)
    }
}
