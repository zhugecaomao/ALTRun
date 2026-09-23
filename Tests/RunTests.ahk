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
#Include %A_ScriptDir%\..\Src\Core\FileIndex.ahk
#Include %A_ScriptDir%\..\Src\UI\ThemeManager.ahk
#Include %A_ScriptDir%\..\Src\UI\IconCache.ahk
#Include %A_ScriptDir%\..\Src\UI\SearchWindow.ahk
#Include %A_ScriptDir%\..\Src\UI\LargeType.ahk
#Include %A_ScriptDir%\..\Src\UI\ItemEditor.ahk
#Include %A_ScriptDir%\..\Src\UI\PreferencesWindow.ahk
#Include %A_ScriptDir%\..\Src\Providers\ClipboardProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\ApplicationProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\CustomCommandProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\SnippetProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\SystemProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\CalculatorProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\WebSearchProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\FileSearchProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\TerminalProvider.ahk
#Include %A_ScriptDir%\..\Src\Extensions\SnippetExpander.ahk
#Include %A_ScriptDir%\..\Src\Extensions\QuickSwitch.ahk
#Include %A_ScriptDir%\..\Src\Extensions\AutoDate.ahk
#Include %A_ScriptDir%\..\Src\Extensions\PTToolsWindow.ahk
#Include %A_ScriptDir%\..\Src\Extensions\UpdateChecker.ahk

OnError((err, mode) => TestRunner.OnUncaught(err, mode))                                             ; 运行错误时输出并退出, 不弹对话框卡住
Logger.Enabled := false
I18n.Init("en")
AppSettings.Data := AppSettings.Defaults()                                  ; 内存里的默认设置, 不读写文件
AppSettings.Feature("Clipboard")["Persist"] := 0                           ; 剪贴板历史测试不写盘

TestRunner.Run()

class TestRunner {
    static Passed := 0, Failed := 0

    static Run() {
        for name in ["FuzzyMatcher", "SearchQuery", "SchemaMigration", "Calculator", "WebSearch"
                    , "AutoDate", "TextTools", "Sorting", "Knowledge", "Clipboard", "SnippetExpander", "Preferences", "FileIndex", "TopIndexes", "EditActions", "Themes", "CommandTargets", "CheckTargets", "EditRows", "HiddenApps", "DefaultFolders", "FileSearchModes", "LegacyIni", "ReleaseVersion", "Misc"] {
            try {
                Tests.%name%()
            } catch as e {
                TestRunner.Fail(name, "exception: " e.Message " (line " e.Line ")")
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

    static SnippetExpander() {
        eq := (n, a, e) => TestRunner.Equal("SnippetExpander." n, a, e)
        eq("prefix", SnippetExpander.Abbreviation(Map("Keyword", "sig", "Text", "x"), ";"), ";sig")
        eq("no keyword", SnippetExpander.Abbreviation(Map("Keyword", "", "Text", "x"), ";"), "")
        eq("disabled", SnippetExpander.Abbreviation(Map("Keyword", "sig", "Text", "x", "AutoExpand", 0), ";"), "")
        eq("space", SnippetExpander.Abbreviation(Map("Keyword", "a b", "Text", "x"), ";"), "")
        eq("no prefix", SnippetExpander.Abbreviation(Map("Keyword", "sig", "Text", "x"), ""), "sig")
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
        eq("builtin count", ThemeManager.Names().Length, 9)
        TestRunner.True("Themes.Ocean is builtin", ThemeManager.IsBuiltin("Ocean"))

        ; 用户主题: 同名覆盖内置主题 ("Base" 写自己 = 在内置那一套上改), 以及 Base 链
        DirCreate(ThemeManager.UserDir)
        FileAppend('{ "Base": "Ocean", "SelectedRadius": 12 }', ThemeManager.UserDir "\Mine.json", "UTF-8")
        FileAppend('{ "Base": "Dark", "Title": "FF0000" }', ThemeManager.UserDir "\Dark.json", "UTF-8")
        ThemeManager.Load("Mine")
        eq("user base color", ThemeManager.Get("Background"), "2E3440")
        eq("user override", ThemeManager.Get("SelectedRadius"), 12)
        TestRunner.True("Themes.user listed", ThemeManager.Names().Length = 10 && !ThemeManager.IsBuiltin("Mine"))
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
        AppSettings.Data["CustomCommands"] := saved
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

        ; 完整的启动流程: 没有 ALTRun.json, 只有 ALTRun.ini
        savedFile := AppSettings.File, savedData := AppSettings.Data
        AppSettings.File := folder "\ALTRun.json", AppSettings.ImportedFrom := "", AppSettings.MigratedFrom := 0
        AppSettings.Load()
        settings := AppSettings.Data
        eq("imported from", AppSettings.ImportedFrom, iniFile)
        eq("version", settings["SchemaVersion"], AppSettings.CurrentVersion)
        eq("json written", FileExist(folder "\ALTRun.json") != "", true)
        eq("ini kept", FileExist(iniFile) != "", true)
        eq("hotkey", settings["General"]["Hotkey"], "!Space")
        eq("second hotkey", settings["General"]["SecondaryHotkey"], "!r")
        eq("language", settings["General"]["Language"], "zh")
        eq("startup", settings["General"]["LaunchAtLogin"], 0)
        eq("file manager", settings["General"]["FileManager"], "C:\Apps\TotalCMD64.exe /O /T /S")
        eq("index depth", settings["Features"]["Applications"]["Depth"], 2)
        eq("index placeholder dropped", settings["Features"]["Applications"]["Folders"].Length, 2)
        eq("structural calc", settings["Features"]["Calculator"]["StructuralCalc"], 1)
        eq("quick switch", settings["Extensions"]["QuickSwitch"]["AutoSwitch"], 1)
        eq("space to run", settings["General"]["SpaceToRun"], 1)
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
        eq("fixture language", fixture["General"]["Language"], "zh")
        eq("fixture hotkey", fixture["Hotkeys"][1]["Key"], "~Mbutton")

        ; 第二次启动: 已经有 ALTRun.json, 不再导入
        AppSettings.ImportedFrom := ""
        AppSettings.Load()
        eq("no second import", AppSettings.ImportedFrom, "")

        AppSettings.File := savedFile, AppSettings.Data := savedData
        AppSettings.ImportedFrom := "", AppSettings.MigratedFrom := 0
        try DirDelete(folder, true)
    }

    ; 编译信息里的版本号 (ALTRun.ahk 的 ;@Ahk2Exe-SetVersion) 要和 App.Version 一致
    static ReleaseVersion() {
        main := FileRead(A_ScriptDir "\..\ALTRun.ahk", "UTF-8")
        RegExMatch(main, "m);@Ahk2Exe-SetVersion\s+(\S+)", &m)
        TestRunner.Equal("ReleaseVersion.exe version = App.Version", IsObject(m) ? m[1] : "", App.Version)
        TestRunner.True("ReleaseVersion.date format", RegExMatch(App.Version, "^\d{4}\.\d{2}\.\d{2}$"))
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
