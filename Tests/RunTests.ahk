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
#Include %A_ScriptDir%\..\Lib\DelimitedText.ahk
#Include %A_ScriptDir%\..\Lib\Kanji.ahk
#Include %A_ScriptDir%\..\Lib\Everything.ahk
#Include %A_ScriptDir%\..\Lib\Units.ahk
#Include %A_ScriptDir%\..\Lib\ClipboardData.ahk
#Include %A_ScriptDir%\..\Lib\Dialogs.ahk
#Include %A_ScriptDir%\..\Lib\HttpGet.ahk
#Include %A_ScriptDir%\..\Lib\FolderWatcher.ahk
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
#Include %A_ScriptDir%\..\Src\UI\NavIcons.ahk
#Include %A_ScriptDir%\..\Src\UI\IconCache.ahk
#Include %A_ScriptDir%\..\Src\UI\SearchWindow.ahk
#Include %A_ScriptDir%\..\Src\UI\WindowFade.ahk
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
#Include %A_ScriptDir%\..\Src\Extensions\CapsLockSwitch.ahk
#Include %A_ScriptDir%\..\Src\Extensions\CapsLockHelper.ahk
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
IconCache.DiskDir := ""                                                     ; 图标不写磁盘缓存 (IconDiskCache 测试里临时打开)

TestRunner.Run()

class TestRunner {
    static Passed := 0, Failed := 0

    ; 8.3 短路径 (C:\Users\RUNNER~1\...) -> 长路径, 和 Loop Files 列出的路径一样
    static LongPath(shortPath) {
        buf := Buffer(32767 * 2)
        return DllCall("GetLongPathNameW", "WStr", shortPath, "Ptr", buf, "UInt", 32767, "UInt") ? StrGet(buf) : shortPath
    }

    static Run() {
        for name in ["FuzzyMatcher", "SearchQuery", "SchemaMigration", "Calculator", "WebSearch"
                    , "AutoDate", "TextTools", "Sorting", "Knowledge", "Clipboard", "ClipboardKinds", "ClipboardPin", "ClipboardLocal", "SnippetExpander", "SnippetPasteMode", "Preferences", "FileIndex", "FileIndexWatch", "TopIndexes", "EditActions", "Themes", "Densities", "MenuMode", "ThemeGallery", "IconScaling", "OwnIcons", "RoundedFill", "BuiltinIcons", "MatchHighlight", "HighlightSpacing", "CommandTargets", "CommandSearchScale", "CheckTargets", "EditRows", "HiddenApps", "BrokenShortcut", "StaleIndexes", "AppIndexRefresh", "WatcherRecovery", "HiddenSystemCommands", "SettingsPages", "PreferencePages", "WindowSwitch", "RecentItems", "Scripts", "ListFilter", "DefaultFolders", "FileSearchModes", "FileSearchPinyin", "FolderSearch", "HelpAndTips", "PreferencesButtons", "PreferencesFit", "PreferencesSpacing", "ListColumnsFit", "I18nLanguages", "I18nUnused", "DefaultExamples", "WindowPosition", "PreferenceDescriptions", "SendTo", "SingleInstance", "AdvancedPage", "HistoryKeys", "FileModeKeys", "TabKey", "SystemDialogKeys", "StatusBar", "TypeAhead", "ShowPrepared", "RestoreQuick", "SiteIcon", "CapsLockSwitch", "ImeToggle", "WindowFade", "IconPrefetch", "IconDiskCache", "WakeHooks", "IconLoadingOnce", "HttpGet", "SlowTaskLog", "LoggerFile", "PreferencesLazyPages", "SkippedPaint", "FontWarmUp", "CustomDrawCritical", "TendonProfileVsSpf2m", "TendonProfileInputs", "PTToolsWindowUi", "LegacyIni", "SettingsLocation", "DataLocation", "ReleaseVersion", "ChangelogLinks", "WikiPages", "SelfUpdate", "UpdateNotice", "HotkeyText", "HotkeyBoxFocus", "JsonReadWrite", "UnitConversion", "CalcBasesDates", "SnippetPlaceholders", "SnippetTextSearch", "SnippetEditor", "MultipleKeywords", "LargeTypeSwitch", "DelimitedText", "SnippetImport", "Bookmarks", "SelectionItems", "FileTypes", "FolderMenu", "FileActions", "EditAction", "FolderBrowse", "DoubleTap", "BrowseKind", "UsageStats", "HudPlacement", "StuckShow", "Misc"] {
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
        eq("3->4 upgraded to current", v4["SchemaVersion"], AppSettings.CurrentVersion)
        eq("3->4 without section", SchemaMigration.Upgrade(Map("SchemaVersion", 3), 3)["SchemaVersion"], AppSettings.CurrentVersion)
        ; 4 -> 5: 片段的粘贴方式 Clipboard (旧的默认值) -> Auto; Type 不变
        v5 := SchemaMigration.Upgrade(Map("SchemaVersion", 4, "Features", Map("Snippets", Map("PasteMode", "Clipboard", "PasteDelay", 500))), 4)
        eq("4->5 clipboard -> auto", v5["Features"]["Snippets"]["PasteMode"], "Auto")
        eq("4->5 keeps other settings", v5["Features"]["Snippets"]["PasteDelay"], 500)
        eq("4->5 version", v5["SchemaVersion"], 5)
        eq("4->5 type stays", SchemaMigration.Upgrade(Map("SchemaVersion", 4, "Features", Map("Snippets", Map("PasteMode", "Type"))), 4)["Features"]["Snippets"]["PasteMode"], "Type")
        eq("4->5 without section", InStr(JSON.Stringify(SchemaMigration.Upgrade(Map("SchemaVersion", 4), 4)), "PasteMode"), 0)
        ; 4 -> 5: 后来补进默认值的窗口 / 关键字也补给老用户 (缺什么补什么, 用户自己加的保留, 清空了的不动); 小清理
        defaults := AppSettings.Defaults()
        old := Map("SchemaVersion", 4
            , "Features", Map("FileSearch", Map("Keywords", ["open", "find"]), "Calculator", Map("BarSizes", ["13", "16", " 20 ", 25, "abc"]))
            , "Extensions", Map("AutoDate", Map("RenameWindows", "ahk_class CabinetWClass,ahk_class Progman,ahk_class WorkerW,ahk_class #32770,ahk_class TTOTAL_CMD"
                                              , "AppendWindows", "ahk_class TCmtEditForm,ahk_exe Notepad4.exe")
                              , "QuickSwitch", Map("ExcludeWindows", "ahk_class SysListView32, ahk_exe Explorer.exe"))
            , "Hotkeys", [Map("Key", "~MButton", "Action", "PTTools", "WinTitle", "ahk_exe RAPTW.exe"), Map("Key", "F1", "Action", "About", "WinTitle", "ALTRun")])
        v5 := SchemaMigration.Upgrade(old, 4)
        eq("4->5 keywords: file added", Tests._Join(v5["Features"]["FileSearch"]["Keywords"]), "open,find,file")
        eq("4->5 keywords = new default", Tests._Join(v5["Features"]["FileSearch"]["Keywords"]), Tests._Join(defaults["Features"]["FileSearch"]["Keywords"]))
        eq("4->5 rename windows: TC dialogs added", v5["Extensions"]["AutoDate"]["RenameWindows"]
            , "ahk_class CabinetWClass,ahk_class Progman,ahk_class WorkerW,ahk_class #32770,ahk_class TTOTAL_CMD, ahk_class TSTDTREEDLG, ahk_class TCOMBOINPUT")
        eq("4->5 append windows = new default", v5["Extensions"]["AutoDate"]["AppendWindows"], defaults["Extensions"]["AutoDate"]["AppendWindows"])
        eq("4->5 exclude windows = new default", v5["Extensions"]["QuickSwitch"]["ExcludeWindows"], defaults["Extensions"]["QuickSwitch"]["ExcludeWindows"])
        eq("4->5 old F1 hotkey removed, others kept", v5["Hotkeys"].Length "|" v5["Hotkeys"][1]["Key"], "1|~MButton")
        sizes := v5["Features"]["Calculator"]["BarSizes"]
        eq("4->5 bar sizes as numbers", (sizes[1] is Integer) "|" sizes[1] "|" sizes[3] "|" (sizes[3] is Integer) "|" sizes[4] "|" sizes[5], "1|13|20|1|25|abc")
        custom := Map("SchemaVersion", 4
            , "Features", Map("FileSearch", Map("Keywords", ["open", "FILE"])), "Extensions", Map("AutoDate", Map("RenameWindows", "ahk_exe foo.exe, AHK_CLASS tstdtreedlg"
                , "AppendWindows", "ahk_exe bar.exe"), "QuickSwitch", Map("ExcludeWindows", "")))
        v5 := SchemaMigration.Upgrade(custom, 4)
        eq("4->5 keywords: no duplicate file", Tests._Join(v5["Features"]["FileSearch"]["Keywords"]), "open,FILE")
        eq("4->5 custom windows kept, only missing added", v5["Extensions"]["AutoDate"]["RenameWindows"], "ahk_exe foo.exe, AHK_CLASS tstdtreedlg, ahk_class TCOMBOINPUT")
        eq("4->5 custom append windows untouched", v5["Extensions"]["AutoDate"]["AppendWindows"], "ahk_exe bar.exe")
        eq("4->5 emptied list stays empty", v5["Extensions"]["QuickSwitch"]["ExcludeWindows"], "")
        eq("4->5 missing sections not created", SchemaMigration.Upgrade(Map("SchemaVersion", 4), 4).Count, 1)
        TestRunner.True("Migration.defaults filled", !AppSettings._MergeDefaults(data, AppSettings.Defaults()))
        TestRunner.True("Migration.no downgrade", SchemaMigration.Upgrade(Map("SchemaVersion", AppSettings.CurrentVersion), AppSettings.CurrentVersion)["SchemaVersion"] = AppSettings.CurrentVersion)
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
        options := AppSettings.Feature("Calculator"), saved := Map()
        for key in ["StructuralCalc", "RebarCover", "MaxBarSpacing", "BarSizes", "BarPrefix"]
            saved[key] := options[key]
        options["StructuralCalc"] := 1
        rows := CalculatorProvider.Search(SearchQuery("300*2"))
        eq("structural rows", rows.Length, 3)
        eq("structural: beam (defaults)", rows[2].Title, I18n.T("Calc.BeamWidth", "600", 3, 260))
        eq("structural: As (defaults)", rows[3].Title, I18n.T("Calc.RebarArea", "600", "5H13  3H16  2H20  2H25  1H32"))
        options["RebarCover"] := 50, options["MaxBarSpacing"] := 150, options["BarSizes"] := ["10", "12"], options["BarPrefix"] := "T"
        rows := CalculatorProvider.Search(SearchQuery("300*2"))
        eq("structural: beam (settings)", rows[2].Title, I18n.T("Calc.BeamWidth", "600", 5, 125))
        eq("structural: As (settings)", rows[3].Title, I18n.T("Calc.RebarArea", "600", "8T10  6T12"))
        eq("structural: below 2 x edge", CalculatorProvider.Search(SearchQuery("50*2")).Length, 1)
        options["RebarCover"] := "abc", options["MaxBarSpacing"] := 0, options["BarSizes"] := ["x", 16]
        rows := CalculatorProvider.Search(SearchQuery("300*2"))
        eq("structural: invalid values use defaults", rows[2].Title "|" rows[3].Title, I18n.T("Calc.BeamWidth", "600", 3, 260) "|" I18n.T("Calc.RebarArea", "600", "3T16"))
        for key, value in saved
            options[key] := value
    }

    static WebSearch() {
        items := WebSearchProvider.Search(SearchQuery("g hello world"))
        TestRunner.Equal("WebSearch.count", items.Length, 1)
        TestRunner.Equal("WebSearch.url", items[1].Arg, "https://www.google.com/search?q=hello%20world")
        items := WebSearchProvider.Search(SearchQuery("bd 中文"))
        TestRunner.Equal("WebSearch.utf8", items[1].Arg, "https://www.baidu.com/s?wd=%E4%B8%AD%E6%96%87")
        saved := ClipboardAll()
        A_Clipboard := ""
        items := WebSearchProvider.Search(SearchQuery("g "))
        TestRunner.True("WebSearch.keyword and space with empty clipboard is not runnable", items.Length = 1 && !items[1].Valid && items[1].Subtitle = I18n.T("Web.ClipboardEmpty"))
        A_Clipboard := "from clipboard"
        items := WebSearchProvider.Search(SearchQuery("g"))
        TestRunner.True("WebSearch.keyword only: hint, clipboard not used", items.Length = 1 && !items[1].Valid && items[1].Title = I18n.T("Web.SearchEmpty", "Google") && items[1].AutoComplete = "g ")
        items := WebSearchProvider.Search(SearchQuery("g "))
        TestRunner.True("WebSearch.keyword and space: clipboard", items.Length = 1 && items[1].Valid && items[1].Arg = "https://www.google.com/search?q=from%20clipboard")
        items := WebSearchProvider.Search(SearchQuery("g typed"))
        TestRunner.Equal("WebSearch.keyword and text: typed text", items[1].Arg, "https://www.google.com/search?q=typed")
        A_Clipboard := ""
        fallbacks := WebSearchProvider.Fallbacks(SearchQuery("xyz"))
        TestRunner.Equal("WebSearch.fallbacks", fallbacks.Length, 3)

        ; {query|默认值}: 只输入关键字时用默认值, 可以直接打开
        eq := (n, a, e) => TestRunner.Equal("WebSearch." n, a, e)
        engine := Map("Id", "engine", "Keyword", "lh", "Title", "Localhost", "Url", "http://localhost:{query|8000}/")
        eq("default value", WebSearchProvider.Fill(engine["Url"], ""), "http://localhost:8000/")
        eq("typed value", WebSearchProvider.Fill(engine["Url"], "3000"), "http://localhost:3000/")
        eq("plain {query}", WebSearchProvider.Fill("https://x.com/?q={query}&r={query}", "a b"), "https://x.com/?q=a%20b&r=a%20b")
        eq("no default", WebSearchProvider.DefaultTerm("https://x.com/?q={query}"), "")
        item := WebSearchProvider.ItemFor(engine, "")
        eq("keyword only with default", (item.Valid ? 1 : 0) "|" item.Arg "|" item.AutoComplete, "1|http://localhost:8000/|lh ")
        eq("keyword only with default: title", item.Title, I18n.T("Web.SearchFor", "Localhost", "8000"))
        item := WebSearchProvider.ItemFor(engine, "", "lh", true)
        eq("keyword without space: hint, default not used", (item.Valid ? 1 : 0) "|" item.Title "|" item.AutoComplete, "0|" I18n.T("Web.SearchEmpty", "Localhost") "|lh ")
        ; {query|clipboard}: 只输入关键字时搜索剪贴板里的文字
        google := Map("Id", "google", "Keyword", "g", "Title", "Google", "Url", "https://www.google.com/search?q={query|clipboard}")
        eq("clipboard default", WebSearchProvider.UsesClipboard(google["Url"]) && WebSearchProvider.UsesClipboard("https://x/?q={query|Clipboard}") && !WebSearchProvider.UsesClipboard(engine["Url"]), true)
        A_Clipboard := "  hello`r`n  world  "
        item := WebSearchProvider.ItemFor(google, "")
        eq("clipboard: item", (item.Valid ? 1 : 0) "|" item.Arg "|" item.AutoComplete, "1|https://www.google.com/search?q=hello%20world|g ")
        eq("clipboard: title", item.Title, I18n.T("Web.SearchClipboardFor", "Google", "hello world"))
        item := WebSearchProvider.ItemFor(google, "typed")
        eq("clipboard: typed text wins", item.Arg "|" item.Title, "https://www.google.com/search?q=typed|" I18n.T("Web.SearchFor", "Google", "typed"))
        item := WebSearchProvider.ItemFor(google, "", "", true)
        eq("clipboard: name hint does not read the clipboard", (item.Valid ? 1 : 0) "|" item.Title, "0|" I18n.T("Web.SearchEmpty", "Google"))
        A_Clipboard := ""
        item := WebSearchProvider.ItemFor(google, "")
        eq("clipboard: empty", (item.Valid ? 1 : 0) "|" item.Title "|" item.Subtitle, "0|" I18n.T("Web.SearchClipboard", "Google") "|" I18n.T("Web.ClipboardEmpty"))
        A_Clipboard := "line1`nline2"
        eq("clipboard: one line", WebSearchProvider.ClipboardText(), "line1 line2")
        A_Clipboard := ""
        Loop 300
            A_Clipboard .= "a"
        eq("clipboard: truncated", StrLen(WebSearchProvider.ClipboardText()), WebSearchProvider.MaxTermChars)
        eq("term: truncated", StrLen(WebSearchProvider.Term(A_Clipboard " more")), WebSearchProvider.MaxTermChars)
        ; 文字和图片一起 (Excel 复制单元格): 用文字; 密码管理器标了 "不要记录": 不读; 复制的是文件: 不读
        putText("cell text", ClipboardData.CF_BITMAP)
        eq("clipboard: text with a bitmap", WebSearchProvider.ClipboardText(), "cell text")
        putText("secret", DllCall("RegisterClipboardFormat", "Str", "ExcludeClipboardContentFromMonitorProcessing", "UInt"))
        eq("clipboard: private", WebSearchProvider.ClipboardText(), "")
        ClipboardData.SetFiles([A_ScriptDir])
        eq("clipboard: files", WebSearchProvider.ClipboardText(), "")
        A_Clipboard := saved
        ; 已有设置文件里自带引擎的旧默认网址换成 {query|clipboard}; 改过的、自己加的不动
        data := Map("Features", Map("WebSearch", Map("Engines", [
            Map("Id", "google", "Url", "https://www.google.com/search?q={query}"),
            Map("Id", "bing", "Url", "https://www.bing.com/search?q={query}&setlang=en"),
            Map("Id", "amz", "Url", "https://www.amazon.com/s?k={query}")])))
        engines := data["Features"]["WebSearch"]["Engines"]
        eq("upgrade engines", AppSettings._UpgradeEngines(data) "|" engines[1]["Url"] "|" engines[2]["Url"] "|" engines[3]["Url"]
            , "1|https://www.google.com/search?q={query|clipboard}|https://www.bing.com/search?q={query}&setlang=en|https://www.amazon.com/s?k={query}")
        eq("upgrade engines: only once", AppSettings._UpgradeEngines(data), false)
        eq("upgrade engines: no web search", AppSettings._UpgradeEngines(Map("Features", Map())), false)
        for engine2 in AppSettings.Feature("WebSearch")["Engines"]
            eq("default engine uses clipboard: " engine2["Id"], WebSearchProvider.UsesClipboard(engine2["Url"]), true)
        ; Icon: 没写 / 文件不存在时用浏览器的图标; 图片用缩略图; 相对路径从 Data 文件夹算起
        eq("icon: none", WebSearchProvider.IconFor(engine), "url:")

        ; 剪贴板里放文字和另一种格式 (format 的内容: 位图格式放一个 1x1 位图, 其它放一个 0)
        putText(text, format) {
            if !DllCall("OpenClipboard", "Ptr", A_ScriptHwnd)
                return
            DllCall("EmptyClipboard")
            size := (StrLen(text) + 1) * 2
            hText := DllCall("GlobalAlloc", "UInt", 0x42, "UPtr", size, "Ptr")
            StrPut(text, DllCall("GlobalLock", "Ptr", hText, "Ptr"), "UTF-16"), DllCall("GlobalUnlock", "Ptr", hText)
            DllCall("SetClipboardData", "UInt", 13, "Ptr", hText)            ; CF_UNICODETEXT
            if (format = ClipboardData.CF_BITMAP) {
                DllCall("SetClipboardData", "UInt", format, "Ptr", DllCall("CreateBitmap", "Int", 1, "Int", 1, "UInt", 1, "UInt", 32, "Ptr", 0, "Ptr"))
            } else {
                hFlag := DllCall("GlobalAlloc", "UInt", 0x42, "UPtr", 4, "Ptr")
                DllCall("SetClipboardData", "UInt", format, "Ptr", hFlag)
            }
            DllCall("CloseClipboard")
        }
        engine["Icon"] := "res:imageres.dll,-1"
        eq("icon: resource", WebSearchProvider.IconFor(engine), "res:imageres.dll,-1")
        engine["Icon"] := A_Temp "\no-such-icon.png"
        eq("icon: missing file", WebSearchProvider.IconFor(engine), "url:")
        png := A_Temp "\ALTRunEngineIcon.png"
        FileOpen(png, "w").Close()
        engine["Icon"] := png
        eq("icon: image", WebSearchProvider.IconFor(engine), "thumb:" png)
        engine["Icon"] := A_ScriptDir "\..\Resources\ALTRun.ico"
        eq("icon: ico file", WebSearchProvider.IconFor(engine), engine["Icon"])
        FileDelete(png)
        engine["Icon"] := "Icons\site.png"
        eq("icon: relative to Data", WebSearchProvider.IconFor(engine), FileExist(AppSettings.DataDir "\Icons\site.png") ? "thumb:" AppSettings.DataDir "\Icons\site.png" : "url:")
        eq("item uses the icon", WebSearchProvider.ItemFor(Map("Keyword", "x", "Title", "X", "Url", "https://x/{query}", "Icon", "res:imageres.dll,-1"), "a").Icon, "res:imageres.dll,-1")

        ; 自带的引擎有自己的图标 (Resources\Icons\Web\<Id>.png)
        savedDir := WebSearchProvider.BuiltinIconDir
        WebSearchProvider.BuiltinIconDir := A_ScriptDir "\..\Resources\Icons\Web", WebSearchProvider._builtinIcons := Map()
        missing := ""
        for engine in AppSettings._DefaultEngines()
            if !FileExist(WebSearchProvider.BuiltinIconDir "\" engine["Id"] ".png")
                missing .= engine["Id"] " "
        eq("every default engine has an icon", missing, "")
        eq("builtin icon", WebSearchProvider.IconFor(Map("Id", "google", "Url", "x")), "thumb:" WebSearchProvider.BuiltinIconDir "\google.png")
        eq("builtin icon when the file is missing", WebSearchProvider.IconFor(Map("Id", "google", "Url", "x", "Icon", "C:\no\such.png")), "thumb:" WebSearchProvider.BuiltinIconDir "\google.png")
        eq("custom engine: browser icon", WebSearchProvider.IconFor(Map("Id", "mysite", "Url", "x")), "url:")
        eq("odd id: browser icon", WebSearchProvider.IconFor(Map("Id", "..\x", "Url", "x")), "url:")
        WebSearchProvider.BuiltinIconDir := savedDir, WebSearchProvider._builtinIcons := Map()
    }

    ; 按 CapsLock 切换输入法: 设置的取值、下一个输入法的选择
    static CapsLockSwitch() {
        eq := (n, a, e) => TestRunner.Equal("CapsLockSwitch." n, a, e)
        eq("off by default", AppSettings.Defaults()["General"]["CapsLock"], "")
        eq("modes", (CapsLockSwitch.IsMode("Layout") ? 1 : 0) (CapsLockSwitch.IsMode("Mode") ? 1 : 0) (CapsLockSwitch.IsMode("") ? 1 : 0) (CapsLockSwitch.IsMode("x") ? 1 : 0), "1100")
        eq("hold time", CapsLockSwitch.HoldMs, 300)
        eq("notification text", CapsLockSwitch.StateText(true) "|" CapsLockSwitch.StateText(false), "Caps Lock On|Caps Lock Off")
        eq("uses the keyboard hook", InStr(CapsLockSwitch.Key, "$") ? 1 : 0, 1)    ; 不用钩子时拦不住大写切换, 按住时也会不停切换
        eq("with any modifier", InStr(CapsLockSwitch.Key, "*") ? 1 : 0, 1)        ; 远程桌面补发的 Alt+CapsLock 等也要拦下
        registered := [], savedInterval := A_HotkeyInterval
        CapsLockSwitch.Register((key, fn) => registered.Push(key))
        eq("down and up hotkeys", registered.Length = 2 ? registered[1] "|" registered[2] : "", "*$CapsLock|*$CapsLock up")
        eq("no too-many-hotkeys warning", A_HotkeyInterval, 0)
        A_HotkeyInterval := savedInterval

        ; 大写锁定的提示可以关掉 (General.CapsLockNotify)
        eq("notice on by default", AppSettings.Defaults()["General"]["CapsLockNotify"], 1)
        notices := [], savedNotify := App.GetOwnPropDesc("Notify")
        App.DefineProp("Notify", {Call: (this, text, *) => notices.Push(text)})
        savedNotice := AppSettings.General.Has("CapsLockNotify") ? AppSettings.General["CapsLockNotify"] : 1
        AppSettings.General["CapsLockNotify"] := 1, App.CapsLockNotice(true)
        AppSettings.General["CapsLockNotify"] := 0, App.CapsLockNotice(false)
        AppSettings.General["CapsLockNotify"] := savedNotice
        App.DefineProp("Notify", savedNotify)
        eq("notice shown only when on", notices.Length = 1 ? notices[1] : notices.Length, "Caps Lock On")

        ; 管理员窗口里的小助手 (CapsLockHelper, #158)
        eq("helper off by default", AppSettings.Defaults()["General"]["CapsLockElevated"], 0)
        eq("helper wanted", (CapsLockHelper.Wanted(Map("CapsLock", "Mode", "CapsLockElevated", 1)) ? 1 : 0)
            . (CapsLockHelper.Wanted(Map("CapsLock", "", "CapsLockElevated", 1)) ? 1 : 0)
            . (CapsLockHelper.Wanted(Map("CapsLock", "Layout", "CapsLockElevated", 0)) ? 1 : 0)
            . (CapsLockHelper.Wanted(Map("CapsLock", "Layout")) ? 1 : 0), "1000")
        settingsFile := A_Temp "\ALTRunHelperTest.json"
        try FileDelete(settingsFile)
        FileAppend('{"General": {"CapsLock": "Layout", "CapsLockElevated": 1}}', settingsFile, "UTF-8")
        eq("helper reads the mode", CapsLockHelper.ModeFromSettings(settingsFile), "Layout")
        FileDelete(settingsFile)
        FileAppend('{"General": {"CapsLock": "Layout", "CapsLockElevated": 0}}', settingsFile, "UTF-8")
        eq("helper: option off", CapsLockHelper.ModeFromSettings(settingsFile), "")
        FileDelete(settingsFile)
        eq("helper: no settings file", CapsLockHelper.ModeFromSettings(settingsFile), "")
        xml := CapsLockHelper.TaskXml("C:\A & B\ALTRun.exe", "-CapsLockHelper", "C:\A & B", "PC\me")
        eq("task: highest privileges", InStr(xml, "<RunLevel>HighestAvailable</RunLevel>") ? 1 : 0, 1)
        eq("task: interactive user", InStr(xml, "<UserId>PC\me</UserId><LogonType>InteractiveToken</LogonType>") ? 1 : 0, 1)
        eq("task: no time limit, runs on battery", (InStr(xml, "<ExecutionTimeLimit>PT0S</ExecutionTimeLimit>") && InStr(xml, "<DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>")) ? 1 : 0, 1)
        eq("task: escaped path", InStr(xml, "<Command>C:\A &amp; B\ALTRun.exe</Command><Arguments>-CapsLockHelper</Arguments>") ? 1 : 0, 1)
        eq("task: no trigger (started by ALTRun)", InStr(xml, "<Triggers>") ? 1 : 0, 0)
        shown := []                                                         ; 大写锁定的提示交给 Notify (小助手里换成告诉 ALTRun)
        savedNotify := CapsLockSwitch.Notify, CapsLockSwitch.Notify := (on) => shown.Push(on)
        before := GetKeyState("CapsLock", "T")
        CapsLockSwitch.ToggleCapsLock()
        eq("notify on toggle", shown.Length = 1 && shown[1] = !before ? 1 : 0, 1)
        SetCapsLockState(before ? "On" : "Off"), CapsLockSwitch.Notify := savedNotify

        ; 按下 / 松开的顺序 (假的时钟, 定时器手动触发)
        saved := Map()
        for name in ["Clock", "ToggleCapsLock", "Switch", "ShiftDown"]
            saved[name] := CapsLockSwitch.GetOwnPropDesc(name)
        clock := {Now: 0}, calls := {Text: ""}
        CapsLockSwitch.Clock := () => clock.Now
        CapsLockSwitch.DefineProp("ToggleCapsLock", {Call: (*) => calls.Text .= "caps;"})
        CapsLockSwitch.DefineProp("Switch", {Call: (*) => calls.Text .= "switch;"})
        shift := {Down: false}
        CapsLockSwitch.DefineProp("ShiftDown", {Call: (*) => shift.Down})
        at := (ms) => clock.Now := ms
        try {
            CapsLockSwitch._down := false
            at(1000), CapsLockSwitch.Down(), at(1100), CapsLockSwitch.Up()
            eq("tap: switch once", calls.Text, "switch;")
            calls.Text := ""
            at(2000), CapsLockSwitch.Down()
            Loop 20                                                         ; 按住时的自动重复, 每 30 ms 一下
                at(2000 + A_Index * 30), CapsLockSwitch.Down()
            CapsLockSwitch.Hold()                                           ; 定时器: 按住满 0.3 秒
            Loop 10
                at(2600 + A_Index * 30), CapsLockSwitch.Down()
            at(3000), CapsLockSwitch.Up()
            eq("hold: caps lock once, never switches", calls.Text, "caps;")
            calls.Text := ""
            at(4000), CapsLockSwitch.Down(), at(4400), CapsLockSwitch.Up()  ; 定时器没来得及运行
            eq("late timer: still a hold", calls.Text, "caps;")
            calls.Text := ""
            at(5000), CapsLockSwitch.Down(), at(7000), CapsLockSwitch.Down(), at(7050), CapsLockSwitch.Up()   ; 上一次的松开丢了
            eq("lost key-up: new press", calls.Text, "switch;")
            calls.Text := ""
            CapsLockSwitch.Up()
            eq("extra key-up ignored", calls.Text, "")
            at(8000), CapsLockSwitch.Down(), at(8000.4), CapsLockSwitch.Up()  ; Chrome 远程桌面补发来对齐大写锁定的
            eq("remote sync tap ignored", calls.Text, "")
            at(9000), CapsLockSwitch.Down(), at(9040), CapsLockSwitch.Up()
            eq("quick real tap still switches", calls.Text, "switch;")
            calls.Text := ""
            shift.Down := true
            at(10000), CapsLockSwitch.Down(), at(10080), CapsLockSwitch.Up()
            eq("shift+caps lock: caps lock as usual", calls.Text, "caps;")
            calls.Text := ""
            at(11000), CapsLockSwitch.Down(), at(11000.3), CapsLockSwitch.Up()   ; 打 Shift+字母时远程桌面补发的
            eq("remote sync tap with shift ignored", calls.Text, "")
            shift.Down := false
            CapsLockSwitch.Clock := saved["Clock"].Value
            ; Sleep 按 15.6 ms 一档的 A_TickCount 判断, 实际可能少睡将近一档: 睡 100 ms, 只要求量到 50 ms 以上
            first := CapsLockSwitch.Clock.Call(), Sleep(100), second := CapsLockSwitch.Clock.Call()
            TestRunner.True("CapsLockSwitch.precise clock", second - first >= 50 && second - first < 5000)
        } finally {
            SetTimer(CapsLockSwitch._holdTimer, 0)
            CapsLockSwitch._down := false
            for name, desc in saved
                CapsLockSwitch.DefineProp(name, desc)
        }
        eq("next layout", Win.NextLayout([0x4090409, 0x8040804], 0x4090409), 0x8040804)
        eq("wraps around", Win.NextLayout([0x4090409, 0x8040804], 0x8040804), 0x4090409)
        eq("unknown current: first", Win.NextLayout([0x4090409, 0x8040804], 0x4110411), 0x4090409)
        eq("no layouts", Win.NextLayout([], 0x4090409), 0)
    }

    ; 下载网站图标: 找网页里写的图标、检查文件头、存进 Data\Icons (联网用假的 fetch 代替)
    static SiteIcon() {
        eq := (n, a, e) => TestRunner.Equal("SiteIcon." n, a, e)
        html := '<html><head><link rel="icon" href="/favicon-16.png" sizes="16x16"><link rel="shortcut icon" href="/favicon.ico">'
            . '<LINK REL="icon" TYPE="image/svg+xml" HREF="/icon.svg"><link href="//cdn.example.com/touch.png?v=2&amp;x=1" rel="apple-touch-icon">'
            . "<link rel='icon' href='img/icon-64.png' sizes='64x64'><link rel=stylesheet href=/a.css><link rel=icon href=data:image/png;base64,AAAA></head>"
        urls := WebSearchProvider.IconCandidates(html, "https://www.example.com/sub/page.html?q=1")
        eq("candidates", urls.Length, 4)
        eq("apple-touch-icon first", urls[1], "https://cdn.example.com/touch.png?v=2&x=1")
        eq("then the biggest size", urls[2], "https://www.example.com/sub/img/icon-64.png")
        eq("unknown size before 16 px", urls[3], "https://www.example.com/favicon.ico")
        eq("16 px last", urls[4], "https://www.example.com/favicon-16.png")
        eq("relative to the site root", WebSearchProvider._Absolute("a.png", "https://x.com"), "https://x.com/a.png")

        png := Buffer(16, 0), bytes := [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]
        for b in bytes
            NumPut("UChar", b, png, A_Index - 1)
        ico := Buffer(16, 0), NumPut("UChar", 1, ico, 2)
        text := Buffer(16, 0), StrPut("<html>", text, "UTF-8")
        eq("png", WebSearchProvider.ImageType(png, 16), "png")
        eq("ico", WebSearchProvider.ImageType(ico, 16), "ico")
        eq("html is not an image", WebSearchProvider.ImageType(text, 16), "")

        page := Buffer(StrPut(html, "UTF-8")), StrPut(html, page, "UTF-8")
        responses := Map("https://www.example.com/", {Status: 200, Body: page, Size: page.Size - 1, Url: "https://www.example.com/sub/page.html"}
            , "https://cdn.example.com/touch.png?v=2&x=1", {Status: 404, Body: Buffer(0), Size: 0, Url: ""}
            , "https://www.example.com/sub/img/icon-64.png", {Status: 200, Body: text, Size: 16, Url: ""}          ; 不是图片
            , "https://www.example.com/favicon.ico", {Status: 200, Body: ico, Size: 16, Url: ""})
        asked := []
        fetch := (target) => (asked.Push(target), responses.Has(target) ? responses[target] : {Status: 404, Body: Buffer(0), Size: 0, Url: ""})
        savedData := AppSettings.DataDir
        AppSettings.DataDir := A_Temp "\ALTRunIconTest"
        try DirDelete(AppSettings.DataDir, true)
        DirCreate(AppSettings.DataDir "\Icons")
        FileAppend("old", AppSettings.DataDir "\Icons\ex.png")
        try {
            result := WebSearchProvider.DownloadIcon("https://www.example.com/search?q={query}", "ex", fetch)
            eq("saved the first real image", result.HasOwnProp("File") ? result.File : result.Error, "Icons\ex.ico")
            eq("file written", FileGetSize(AppSettings.DataDir "\Icons\ex.ico"), 16)
            eq("old icon with the same name removed", FileExist(AppSettings.DataDir "\Icons\ex.png") ? 1 : 0, 0)
            eq("visited only that site", asked[1], "https://www.example.com/")
            result := WebSearchProvider.DownloadIcon("http://localhost:{query|8000}/x", "", (t) => (asked.Push(t), {Status: 404, Body: Buffer(0), Size: 0, Url: ""}))
            eq("default value in the host", asked[asked.Length], "http://localhost:8000/favicon.ico")
            eq("nothing found", result.HasOwnProp("Error") && InStr(result.Error, "localhost:8000") ? 1 : 0, 1)
            eq("not a web address", WebSearchProvider.DownloadIcon("notes.txt", "x", fetch).HasOwnProp("Error"), true)
            big := WebSearchProvider.DownloadIcon("https://big.example.com/", "big", (t) => {Status: 200, Body: png, Size: WebSearchProvider.MaxIconBytes + 1, Url: ""})
            eq("too big", big.HasOwnProp("Error"), true)
        } finally {
            try DirDelete(AppSettings.DataDir, true)
            AppSettings.DataDir := savedData
        }
    }

    static AutoDate() {
        eq := (n, a, e) => TestRunner.Equal("AutoDate." n, a, e)
        today := "23.09.2026"
        ; Ctrl+D 只在改名的编辑框里接管 (其它时候按键原样交给程序, 不补发)
        eq("edit box", AutoDate.IsEditControl("Edit1"), true)
        eq("TC rename box", AutoDate.IsEditControl("TInEdit1"), true)
        eq("scintilla", AutoDate.IsEditControl("Scintilla1"), true)
        eq("file list", AutoDate.IsEditControl("DirectUIHWND3"), false)
        eq("no focus", AutoDate.IsEditControl(""), false)
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
        ; 剪贴板变了之后等 ReadDelay 再读 (不占着剪贴板); 等的时候又变了: 先读上一次的
        saved := ClipboardAll(), ClipboardProvider._pausedUntil := 0
        A_Clipboard := "first copy"
        ClipboardProvider.Entries := []
        ClipboardProvider._OnChange(1)
        eq("read later", ClipboardProvider.Entries.Length, 0)
        A_Clipboard := "second copy"
        ClipboardProvider._OnChange(1)
        eq("earlier change read first", ClipboardProvider.Entries.Length ? ClipboardProvider.Entries[1]["Text"] : "", "second copy")
        Sleep(ClipboardProvider.ReadDelay + 150)
        eq("read after the delay", ClipboardProvider.Entries.Length " " ClipboardProvider.Entries[1]["Text"], "1 second copy")
        eq("nothing pending", IsObject(ClipboardProvider._pending), false)
        A_Clipboard := "copied before ALTRun borrows the clipboard"
        ClipboardProvider.Entries := []
        ClipboardProvider._OnChange(1)
        ClipboardProvider.PauseRecording(1000)                              ; 例如片段粘贴: 剪贴板里马上换成临时的内容
        Sleep(ClipboardProvider.ReadDelay + 150)
        eq("not read while ALTRun borrows it", ClipboardProvider.Entries.Length, 0)
        ClipboardProvider._pausedUntil := 0
        A_Clipboard := saved
        Sleep(ClipboardProvider.ReadDelay + 150)
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

    ; 粘贴方式 Auto: 短的单行文字直接输入 (不动剪贴板), 多行 / Tab / 很长 / emoji 用剪贴板
    static SnippetPasteMode() {
        eq := (n, a, e) => TestRunner.Equal("SnippetPasteMode." n, a, e)
        types := (text, mode) => ActionCatalog.TypesText(text, mode)
        long := StrReplace(Format("{:" ActionCatalog.AutoTypeMax "}", ""), " ", "x")
        eq("default", AppSettings.Defaults()["Features"]["Snippets"]["PasteMode"], "Auto")
        eq("auto: short line typed", types("Best regards, Bench", "Auto"), true)
        eq("auto: CJK typed", types("你好, 谢谢", "Auto"), true)
        eq("auto: at the limit typed", types(long, "Auto"), true)
        eq("auto: longer pasted", types(long "x", "Auto"), false)
        eq("auto: several lines pasted", types("Dear Sir,`nThanks", "Auto"), false)
        eq("auto: CR pasted", types("a`rb", "Auto"), false)
        eq("auto: tab pasted", types("a`tb", "Auto"), false)
        eq("auto: emoji pasted", types("ok " Chr(0x1F600), "Auto"), false)
        eq("clipboard: never typed", types("hi", "Clipboard"), false)
        eq("type: always typed", types("Dear Sir,`nThanks", "Type"), true)
        eq("unknown mode: clipboard", types("hi", "Other"), false)

        ; 真的输入到一个编辑框: 文字出来了, 剪贴板没动
        snippets := AppSettings.Feature("Snippets")
        savedMode := snippets["PasteMode"], savedClip := ClipboardAll()
        g := Gui("+AlwaysOnTop", "ALTRun paste test")
        box := g.Add("Edit", "w300 h40")
        g.Show()
        try {
            WinActivate("ahk_id " g.Hwnd)
            if WinWaitActive("ahk_id " g.Hwnd, , 2) {
                ControlFocus(box)
                A_Clipboard := "keep me"
                snippets["PasteMode"] := "Auto"
                eq("auto: pasted", ActionCatalog.PasteText("typed text", false), true)
                Sleep(100)
                eq("auto: text typed", box.Value, "typed text")
                eq("auto: clipboard untouched", A_Clipboard, "keep me")
            } else {
                FileAppend("NOTE SnippetPasteMode: test window not active, typing test skipped`n", "*")
            }
        } finally {
            snippets["PasteMode"] := savedMode
            g.Destroy()
            A_Clipboard := savedClip
        }
    }

    static SnippetExpander() {
        eq := (n, a, e) => TestRunner.Equal("SnippetExpander." n, a, e)
        abbr := (snippet, prefix) => Tests._Join(SnippetExpander.Abbreviations(snippet, prefix))
        eq("prefix", abbr(Map("Keyword", "sig", "Text", "x"), ";"), ";sig")
        eq("no keyword", abbr(Map("Keyword", "", "Text", "x"), ";"), "")
        eq("disabled", abbr(Map("Keyword", "sig", "Text", "x", "AutoExpand", 0), ";"), "")
        eq("space", abbr(Map("Keyword", "a b", "Text", "x"), ";"), "")
        eq("no prefix", abbr(Map("Keyword", "sig", "Text", "x"), ""), "sig")
        eq("several keywords", abbr(Map("Keyword", " sig, ㄴ ,，sg ", "Text", "x"), ";"), ";sig,;ㄴ,;sg")
        eq("several: bad one skipped", abbr(Map("Keyword", "a b, ab", "Text", "x"), ";"), ";ab")
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
        sizes := PreferencesWindow.SplitCsv("10, 12,16", true)             ; 数字列表 (钢筋直径) 存成数字
        eq("numeric csv", (sizes[1] is Number) && sizes[3] = 16 ? 1 : 0, 1)
        eq("text csv stays text", PreferencesWindow.SplitCsv("10")[1] is String ? 1 : 0, 1)
    }

    static FileIndex() {
        eq := (n, a, e) => TestRunner.Equal("FileIndex." n, a, e)
        FileIndex.Paths := ["C:\Docs\Project Atlas", "C:\Docs\Project Atlas\Atlas Plan.dwg", "C:\Docs\notes.txt", "C:\Docs\Status Report.pdf"]
        FileIndex.Names := ["project atlas", "atlas plan.dwg", "notes.txt", "status report.pdf"]
        FileIndex.Folders := [1, 0, 0, 0]
        FileIndex._lastNeedle := "", FileIndex._lastMatches := ""
        found := FileIndex.Search("atlas", 10)
        eq("count", found.Length, 2)
        eq("prefix first", found[1].Path, "C:\Docs\Project Atlas\Atlas Plan.dwg")
        eq("narrowed", FileIndex.Search("atlas p", 10).Length, 1)
        eq("new search", FileIndex.Search("rep", 10)[1].Path, "C:\Docs\Status Report.pdf")
        TestRunner.True("FileIndex.word start >= 70", FileIndex.ScoreName("atlas", "project atlas") >= 70)
        eq("no match", FileIndex.ScoreName("xyz", "notes.txt"), 0)
        FileIndex.Paths := [], FileIndex.Names := [], FileIndex.Folders := []
    }

    ; 监视索引的文件夹: 哪些变化要重新扫描、变化后重新扫描、扫完后等空闲再写缓存
    static FileIndexWatch() {
        eq := (n, a, e) => TestRunner.Equal("FileIndexWatch." n, a, e)
        roots := ["C:\Users\me\Desktop", "D:\Docs"]
        changed := (fullPath, exclude := "") => FileIndex.IsWatchedChange(fullPath, roots, 2, exclude, "D:\ALTRun")
        eq("file in root", changed("C:\Users\me\Desktop\Plan.dwg"), true)
        eq("second level", changed("D:\Docs\Tower\Plan.pdf"), true)
        eq("deeper than the scan", changed("D:\Docs\Tower\Old\Plan.pdf"), false)
        eq("outside the roots", changed("C:\Users\me\Downloads\a.zip"), false)
        eq("similar root name", changed("D:\Docs2\a.txt"), false)
        eq("download in progress", changed("C:\Users\me\Desktop\setup.exe.crdownload"), false)
        eq("office lock file", changed("D:\Docs\~$Report.docx"), false)
        eq("temp file", changed("D:\Docs\~WRL0001.tmp"), false)
        eq("libreoffice lock", changed("D:\Docs\.~lock.Report.odt#"), false)
        eq("excluded", changed("D:\Docs\node_modules", "i)\\node_modules(\\|$)"), false)
        ; 默认的排除: 代码仓库、虚拟环境、开发工具的缓存 (整个文件夹, 不管大小写); 名字只是相似的照常扫描
        exclude := AppSettings.Defaults()["Features"]["FileSearch"]["ScopeExclude"]
        for folder in ["node_modules", ".git", ".SVN", ".hg", "__pycache__", ".venv", "venv", ".tox", ".cache", ".idea", ".vs", "$RECYCLE.BIN"]
            eq("default exclude " folder, RegExMatch("D:\Docs\Tool\" folder "\x.txt", exclude) > 0, true)
        for folder in ["venvs", "my.venv", "cache", "Visual Studio", ".vscode", "github", "idea"]
            eq("default keeps " folder, RegExMatch("D:\Docs\" folder "\x.txt", exclude), 0)
        eq("own folder", FileIndex.IsWatchedChange("D:\Docs\ALTRun\Data\FileIndex.json", roots, 3, "", "D:\Docs\ALTRun"), false)

        root := A_Temp "\ALTRunFileIndexWatch"
        try DirDelete(root, true)
        DirCreate(root)
        FileAppend("", root "\Old Plan.dwg")
        options := AppSettings.Feature("FileSearch")
        saved := {Folders: options["ScopeFolders"], Everything: options["UseEverything"], File: FileIndex.File
            , Paths: FileIndex.Paths, Names: FileIndex.Names, Folders2: FileIndex.Folders
            , Quiet: FileIndex.WatchQuietMs, Gap: FileIndex.WatchMinGapMs, Idle: FileIndex.SaveIdleMs}
        options["ScopeFolders"] := [root], options["UseEverything"] := 0
        FileIndex.File := A_Temp "\ALTRunFileIndexWatch.json"
        try FileDelete(FileIndex.File)
        FileIndex.WatchQuietMs := 0, FileIndex.WatchMinGapMs := 0, FileIndex.SaveIdleMs := 999999999   ; 不自动写, 测试里自己写
        waitScan() {
            Loop 100 {
                if !FileIndex.Scanning
                    break
                Sleep(20)
            }
        }
        try {
            FileIndex.StartWatching()
            eq("watching", FileIndex._watcher.Count, 1)
            FileIndex.Rebuild(), waitScan()
            eq("scanned", FileIndex.Paths.Length, 1)
            eq("save waits for idle", FileIndex._savePending && !FileExist(FileIndex.File), true)
            FileIndex.SavePending()
            eq("saved on demand", !FileIndex._savePending && FileExist(FileIndex.File) != "", true)
            FileIndex._PollWatches()
            eq("no change, no scan", FileIndex.Scanning, false)
            FileAppend("x", root "\Old Plan.dwg")                         ; 修改文件: 索引里只有路径, 不用重新扫描
            FileAppend("", root "\Report.docx.crdownload")                ; 临时文件也不算
            Sleep(300)
            FileIndex._PollWatches()
            eq("edit and temp file ignored", FileIndex.Scanning || FileIndex._changedAt != 0, false)
            FileAppend("", root "\New Plan.dwg")
            Sleep(300)
            FileIndex._PollWatches()
            eq("new file rescans", FileIndex.Scanning || FileIndex.Paths.Length = 3, true)
            waitScan()
            eq("new file indexed", FileIndex.Paths.Length, 3)               ; 包括 .crdownload (扫描不跳过, 只是它不触发)
            FileIndex.WatchMinGapMs := 600000
            FileDelete(root "\New Plan.dwg")
            Sleep(300)
            FileIndex._PollWatches()
            eq("minimum gap", FileIndex.Scanning = false && FileIndex._changedAt != 0, true)
            FileIndex.StopWatching()
            eq("stopped", FileIndex._watcher, "")
        } finally {
            FileIndex.StopWatching()
            waitScan()
            if (FileIndex._saveTimer != "")
                SetTimer(FileIndex._saveTimer, 0)
            FileIndex._savePending := false, FileIndex._changedAt := 0, FileIndex._lastScan := 0
            options["ScopeFolders"] := saved.Folders, options["UseEverything"] := saved.Everything
            FileIndex.File := saved.File
            FileIndex.Paths := saved.Paths, FileIndex.Names := saved.Names, FileIndex.Folders := saved.Folders2
            FileIndex.WatchQuietMs := saved.Quiet, FileIndex.WatchMinGapMs := saved.Gap, FileIndex.SaveIdleMs := saved.Idle
            try FileDelete(A_Temp "\ALTRunFileIndexWatch.json")
            try DirDelete(root, true)
        }
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
        FileAppend("not a shortcut", root "\Recent\Broken.lnk")                     ; 损坏的快捷方式: 跳过, 不报错
        FileSetTime("20260401000000", root "\Recent\Broken.lnk")
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
        ; 文件对话框: #32770 要有资源管理器的文件视图; "系统信息" (msinfo32) 的主窗口只有普通列表, 不算
        fileDialog := ["DUIViewWndClassName1", "DirectUIHWND1", "SHELLDLL_DefView1", "DirectUIHWND2", "ComboBoxEx321", "ComboBox1", "Edit1", "Button1", "Button2", "ToolbarWindow323"]
        oldDialog := ["ToolbarWindow321", "ComboBox1", "SHELLDLL_DefView1", "SysListView321", "Edit1", "Button1", "Button2"]
        msinfo := ["SysTreeView321", "SysListView321", "Static1", "Edit1", "Button1", "Button2", "Button3", "Button4"]
        eq("file dialog", QuickSwitch.LooksLikeFileDialog("#32770", fileDialog), true)
        eq("old-style file dialog", QuickSwitch.LooksLikeFileDialog("#32770", oldDialog), true)
        eq("msinfo32 main window", QuickSwitch.LooksLikeFileDialog("#32770", msinfo), false)
        eq("other app with a plain list", QuickSwitch.LooksLikeFileDialog("TMyOpenDialog", msinfo), true)
        eq("no controls: title", QuickSwitch.LooksLikeFileDialog("Qt5QWindowIcon", [], "Open File") " " QuickSwitch.LooksLikeFileDialog("Qt5QWindowIcon", [], "Writer"), "1 0")
        ; "Browse for Folder" (folder tree only): its own kind, never a file dialog (pressing Enter there means OK)
        newStyle := ["Static1", "SHBrowseForFolder ShellNameSpace Control1", "SysTreeView321", "Static2", "Edit1", "Button1", "Button2", "Button3"]
        classic := ["SysTreeView321", "Static1", "Static2", "Static3", "Edit1", "Button1", "Button2", "Button3"]
        eq("browse dialog: new style", QuickSwitch.LooksLikeBrowseDialog("#32770", newStyle), true)
        eq("browse dialog: classic tree ID", QuickSwitch.LooksLikeBrowseDialog("#32770", classic, 0x3741), true)
        eq("other dialog with a tree", QuickSwitch.LooksLikeBrowseDialog("#32770", classic, 0x3F2), false)
        eq("browse dialog: other class", QuickSwitch.LooksLikeBrowseDialog("TMyDialog", newStyle), false)
        eq("file dialog is not a browse dialog", QuickSwitch.LooksLikeBrowseDialog("#32770", fileDialog), false)
        eq("browse dialog is not a file dialog", QuickSwitch.LooksLikeFileDialog("#32770", newStyle) QuickSwitch.LooksLikeFileDialog("#32770", classic), "00")
        eq("browse dialogs on by default", AppSettings.Defaults()["Extensions"]["QuickSwitch"]["BrowseDialogs"], 1)
        ; A real one: AutoHotkey's DirSelect is SHBrowseForFolder (option 4 = classic style). Detected, the folder is
        ; selected in its tree, and OK returns it
        target := A_WinDir "\System32"
        for options in [0, 4] {
            seen := {Detected: "", Selected: ""}, tries := 0
            probe() {
                if !(dlg := SearchWindow.InSystemDialog()) {
                    if (++tries < 50)
                        SetTimer(probe, -100)
                    return
                }
                seen.Detected := QuickSwitch.IsBrowseWindow(dlg)
                seen.Selected := QuickSwitch.SelectInBrowseDialog(target "\", dlg)
                Sleep(300)
                PostMessage(0x111, 1, 0, , "ahk_id " dlg)                   ; WM_COMMAND IDOK
                SetTimer(() => WinExist("ahk_id " dlg) && WinClose("ahk_id " dlg), -3000)   ; never left open
            }
            SetTimer(probe, -300)
            picked := DirSelect("C:\", options)
            eq("DirSelect " options ": detected", seen.Detected, true)
            eq("DirSelect " options ": selected", seen.Selected, true)
            eq("DirSelect " options ": OK returns it", StrLower(RTrim(picked, "\")), StrLower(target))
        }
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
    ; 文件类型登记了 "编辑" 时, 操作面板里有 "编辑" (右边是编辑程序的名称), Ctrl+E 也能用; 没登记的类型没有
    static EditAction() {
        eq := (n, a, e) => TestRunner.Equal("EditAction." n, a, e)
        root := A_Temp "\ALTRun-test-edit"
        try DirDelete(root, true)
        DirCreate(root)
        FileAppend("x", root "\script.altruntest")
        FileAppend("x", root "\plain.altrunnone")
        classes := "HKCU\Software\Classes"
        try {
            RegWrite("ALTRunTest.File", "REG_SZ", classes "\.altruntest")
            RegWrite('"' A_WinDir '\notepad.exe" "%1"', "REG_SZ", classes "\ALTRunTest.File\shell\edit\command")
            ; Wine 的 AssocQueryStringW 查不到任何类型 (连 .txt 都没有), 登记了的情况只在 Windows 上检查
            onWine := DllCall("GetProcAddress", "Ptr", DllCall("GetModuleHandle", "Str", "ntdll", "Ptr"), "AStr", "wine_get_version", "Ptr")
            if !onWine {
                editor := ActionCatalog.EditorFor(root "\script.altruntest")
                eq("registered", IsObject(editor), true)
                eq("editor icon", IsObject(editor) ? editor.Icon : "", A_WinDir "\notepad.exe")
            }
            eq("not registered", ActionCatalog.EditorFor(root "\plain.altrunnone"), "")
            eq("folder", ActionCatalog.EditorFor(root), "")
            eq("missing file", ActionCatalog.EditorFor(root "\gone.altruntest"), "")
            eq("unknown type", ActionCatalog.AssocString(1, ".altrunnone", "edit"), "")
            titles(path) {
                list := "|"
                for action in ActionCatalog.ListFor(FileSearchProvider._ToItem({Path: path, IsFolder: false}, 10))
                    list .= action.Title "|"
                return list
            }
            if !onWine
                eq("action shown", InStr(titles(root "\script.altruntest"), "|" I18n.T("Action.EditFile") "|") ? 1 : 0, 1)
            eq("action hidden", InStr(titles(root "\plain.altrunnone"), "|" I18n.T("Action.EditFile") "|") ? 1 : 0, 0)
            openWith := ""
            for action in ActionCatalog.ListFor(FileSearchProvider._ToItem({Path: root "\plain.altrunnone", IsFolder: false}, 10))
                if (action.Title = I18n.T("Action.OpenWith"))
                    openWith := action.Icon
            eq("open with icon", openWith, "res:shell32.dll,-3")
            if !onWine                                                      ; 原来的 shell32.dll,-16702 在 Windows 上没有, 菜单里没有图标
                eq("open with icon loads", IconCache.GetNow(openWith) != 0, true)
        } finally {
            try RegDeleteKey(classes "\ALTRunTest.File")
            try RegDeleteKey(classes "\.altruntest")
            try DirDelete(root, true)
        }
    }

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

    ; Highlighted and other characters are drawn in segments; character positions must not move with the highlight boundary (jitter while typing)
    static HighlightSpacing() {
        title := "Western Wave Station", width := 320, height := 40
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
        plainHighlight := ink("western wave station")                       ; the whole title highlighted: compare with the segmented drawing
        for query in ["", "w", "we", "wes", "west", "wave", "ws", "zzz"]                ; "" / "ws" / "zzz": no highlighted characters
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

    ; Resources\Icons 里 ALTRun 自己的图标: 每个都在, 各种尺寸都能读出正好的大小; 文件不在时用系统图标
    static OwnIcons() {
        eq := (n, a, e) => TestRunner.Equal("OwnIcons." n, a, e)
        saved := IconCache.Size
        for name in ["Clipboard", "Snippet", "Help", "Calculator", "Pinned"] {
            iconFile := A_ScriptDir "\..\Resources\Icons\" name ".ico"
            if !TestRunner.True("OwnIcons." name " exists", FileExist(iconFile))
                continue
            for size in [16, 32, 40, 64] {
                IconCache.Size := size
                hIcon := IconCache._Load(iconFile)
                eq(name " at " size, IconCache.IconWidth(hIcon), size)
                if hIcon
                    DllCall("DestroyIcon", "Ptr", hIcon)
            }
        }
        IconCache.Size := saved
        eq("fallback", IconCache.Own("NoSuchIcon", "res:shell32.dll,-24"), "res:shell32.dll,-24")
    }

    ; 选中行和主题缩略图的圆角矩形: 里面和直边是实心的, 圆角外面不画; 圆角边缘抗锯齿 (Wine 的 GDI+ 不做抗锯齿, 跳过这一项)
    static RoundedFill() {
        eq := (n, a, e) => TestRunner.Equal("RoundedFill." n, a, e)
        screen := DllCall("GetDC", "Ptr", 0, "Ptr")
        dc := DllCall("CreateCompatibleDC", "Ptr", screen, "Ptr"), hbm := DllCall("CreateCompatibleBitmap", "Ptr", screen, "Int", 40, "Int", 24, "Ptr")
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", screen)
        old := DllCall("SelectObject", "Ptr", dc, "Ptr", hbm, "Ptr")
        rect := Buffer(16), NumPut("Int", 0, "Int", 0, "Int", 40, "Int", 24, rect)
        white := DllCall("GetStockObject", "Int", 0, "Ptr")                   ; WHITE_BRUSH
        DllCall("FillRect", "Ptr", dc, "Ptr", rect, "Ptr", white)
        Win.FillRoundRect(dc, 0, 0, 40, 24, 8, 0x000000)
        px := (x, y) => DllCall("GetPixel", "Ptr", dc, "Int", x, "Int", y, "UInt") & 0xFFFFFF
        eq("center", px(20, 12), 0)
        eq("top edge", px(20, 0), 0)
        eq("right edge", px(39, 12), 0)
        eq("bottom edge", px(20, 23), 0)
        eq("outside corner", px(0, 0), 0xFFFFFF)
        eq("outside right of last column", px(39, 0), 0xFFFFFF)
        if !DllCall("GetProcAddress", "Ptr", DllCall("GetModuleHandle", "Str", "ntdll", "Ptr"), "AStr", "wine_get_version", "Ptr") {
            blended := 0
            Loop 8 {
                c := px(A_Index - 1, 8 - A_Index) & 0xFF
                blended += (c > 0 && c < 0xFF)
            }
            TestRunner.True("RoundedFill.anti-aliased corner", blended > 0)
        }
        DllCall("FillRect", "Ptr", dc, "Ptr", rect, "Ptr", white)
        Win.FillRoundRect(dc, 0, 0, 40, 24, 0, 0x000000)                       ; 半径 0: 直角矩形
        eq("square corner", px(0, 0), 0)
        DllCall("SelectObject", "Ptr", dc, "Ptr", old)
        DllCall("DeleteObject", "Ptr", hbm), DllCall("DeleteDC", "Ptr", dc)
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
        ; 读得慢的图标写进调试日志 (阈值设为 0 时每个都算慢)
        savedLog := [Logger.Enabled, Logger.File, Logger.Immediate, IconCache.SlowLoadMs]
        Logger.Enabled := true, Logger.Immediate := true, Logger.File := A_Temp "\ALTRunIconLogTest.log"
        try FileDelete(Logger.File)
        IconCache.SlowLoadMs := 0
        IconCache.Get("ext:.altrunslowtest"), IconCache._LoadQueued()
        IconCache.SlowLoadMs := 100000
        IconCache.Get("ext:.altrunfasttest"), IconCache._LoadQueued()
        logText := FileExist(Logger.File) ? FileRead(Logger.File, "UTF-8") : ""
        TestRunner.True("IconScaling.slow icon logged", RegExMatch(logText, "Perf: icon \d+ ms: ext:\.altrunslowtest") > 0)
        TestRunner.True("IconScaling.fast icon not logged", !InStr(logText, "altrunfasttest"))
        try FileDelete(Logger.File)
        Logger.Enabled := savedLog[1], Logger.File := savedLog[2], Logger.Immediate := savedLog[3], IconCache.SlowLoadMs := savedLog[4]
    }

    ; 主题缩略图: 每个主题都画得出来, 背景、选中行用的是主题自己的颜色; 外观页点选缩略图就是选主题
    ; 弹出菜单跟随主题的深浅: 按背景色判断深色; Windows 10 1903 (18362) 起才设置
    static MenuMode() {
        eq := (n, a, e) => TestRunner.Equal("MenuMode." n, a, e)
        saved := [ThemeManager.Name, ThemeManager.Current, ThemeManager.Resolved, ThemeManager.MenuMode]
        try {
            for themeName, dark in Map("Light", false, "Paper", false, "SolarizedLight", false, "Dark", true, "Midnight", true, "TokyoNight", true) {
                ThemeManager.Load(themeName)
                eq(themeName, ThemeManager.IsDark(), dark)
            }
            eq("win11", ThemeManager.SupportsMenuMode("10.0.22631"), true)
            eq("win10 1903", ThemeManager.SupportsMenuMode("10.0.18362"), true)
            eq("win10 1809", ThemeManager.SupportsMenuMode("10.0.17763"), false)
            eq("win7", ThemeManager.SupportsMenuMode("6.1.7601"), false)
            ThemeManager.MenuMode := 2
            eq("same mode: nothing to do", ThemeManager.ApplyMenuMode(true), false)
        } finally {
            ThemeManager.Name := saved[1], ThemeManager.Current := saved[2], ThemeManager.Resolved := saved[3], ThemeManager.MenuMode := saved[4]
        }
    }

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
        il := ThemePreview.ImageList(["Dark"], 112, 76)                       ; 经过图像列表画出来: 背景不透明 (GDI+ 的圆角不能让背景变透明)
        screen := DllCall("GetDC", "Ptr", 0, "Ptr")
        dc := DllCall("CreateCompatibleDC", "Ptr", screen, "Ptr"), canvas := DllCall("CreateCompatibleBitmap", "Ptr", screen, "Int", 112, "Int", 76, "Ptr")
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", screen)
        oldCanvas := DllCall("SelectObject", "Ptr", dc, "Ptr", canvas, "Ptr")
        rect := Buffer(16), NumPut("Int", 0, "Int", 0, "Int", 112, "Int", 76, rect)
        DllCall("FillRect", "Ptr", dc, "Ptr", rect, "Ptr", DllCall("GetStockObject", "Int", 0, "Ptr"))   ; 白色
        DllCall("comctl32\ImageList_Draw", "Ptr", il, "Int", 0, "Ptr", dc, "Int", 0, "Int", 0, "UInt", 0)   ; ILD_NORMAL
        eq("image list background opaque", DllCall("GetPixel", "Ptr", dc, "Int", 56, "Int", 6, "UInt"), Win.ColorToBgr(ThemeManager.Resolve("Dark")["Background"]))
        DllCall("SelectObject", "Ptr", dc, "Ptr", oldCanvas), DllCall("DeleteObject", "Ptr", canvas), DllCall("DeleteDC", "Ptr", dc)
        DllCall("comctl32\ImageList_Destroy", "Ptr", il)
        hbm := ThemePreview.ForName("System", 112, 76)                      ; 左半边 Light, 右半边 Dark
        eq("system left", pixel(hbm, 20, 6), Win.ColorToBgr(ThemeManager.Defaults()["Background"]))
        eq("system right", pixel(hbm, 100, 6), Win.ColorToBgr(ThemeManager.Resolve("Dark")["Background"]))
        DllCall("DeleteObject", "Ptr", hbm)

        PreferencesWindow.Show(1, -3000, -3000)
        PreferencesWindow.BuildAll()                                         ; 页面用到时才建: 检查每一页之前先全部建好
        Sleep(400)                                                          ; 打开 300 ms 之后的修改才算 (见 PreferencesWindow.Show)
        focused := DllCall("GetFocus", "Ptr")                               ; 打开时焦点在页面列表 (↑ ↓ 切换页面, 不改第一页的设置)
        eq("focus on page list", focused = PreferencesWindow.PageList.Hwnd, true)
        nextTab := DllCall("GetNextDlgTabItem", "Ptr", PreferencesWindow.Gui.Hwnd, "Ptr", PreferencesWindow.PageList.Hwnd, "Int", 0, "Ptr")
        onPage := false                                                     ; Tab: 列表之后是当前页面里的控件, 不是底部按钮
        for ctrl in PreferencesWindow.Pages[1].Controls
            onPage := onPage || ctrl.Hwnd = nextTab
        eq("tab from page list goes into the page", onPage, true)
        keys := ""                                                         ; PageKeys 和实际建的页面一致 (搜索里的子页面靠它)
        for page in PreferencesWindow.Pages
            keys .= page.Key " "
        expected := ""
        for pageKey in PreferencesWindow.PageKeys
            expected .= pageKey " "
        eq("page keys match the pages", keys, expected)
        SystemProvider._PageOpener("Prefs.Page.Calculator")()               ; 窗口开着时跳到那一页
        eq("open a page from search", PreferencesWindow._page, PreferencesWindow._PageIndex("Prefs.Page.Calculator"))
        eq("page list follows", PreferencesWindow.PageList.GetNext(), PreferencesWindow._page)
        eq("page list rows", PreferencesWindow.PageList.GetCount(), PreferencesWindow.Pages.Length)
        appearancePage := PreferencesWindow._PageIndex("Prefs.Page.Appearance")
        PreferencesWindow.Show(appearancePage)                              ; opening by number while already open ("-Preferences 4" handed over)
        eq("open by number while open", PreferencesWindow._page "|" PreferencesWindow.PageList.GetNext(), appearancePage "|" appearancePage)
        gallery := ""                                                       ; the theme list on the Appearance page: exactly 4 thumbnails per row, no wide gap on the right
        for ctrl in PreferencesWindow.Pages[appearancePage].Controls
            if (ctrl.Type = "ListView" && !IsObject(gallery))
                gallery := ctrl
        itemPos := (index) => (point := Buffer(8, 0), SendMessage(0x1010, index, point.Ptr, gallery), [NumGet(point, 0, "Int"), NumGet(point, 4, "Int")])   ; LVM_GETITEMPOSITION
        spacing := SendMessage(0x1033, 0, 0, gallery) & 0xFFFF              ; LVM_GETITEMSPACING (大图标)
        galleryRect := Buffer(16)
        DllCall("GetClientRect", "Ptr", gallery.Hwnd, "Ptr", galleryRect)
        first := itemPos(0), fourth := itemPos(3), fifth := itemPos(4)
        eq("gallery: 4 per row", (fourth[2] = first[2] && fifth[2] > first[2]) ? 1 : 0, 1)
        TestRunner.True("ThemeGallery.gallery fills the row (left " (NumGet(galleryRect, 8, "Int") - 4 * spacing) ")"
            , NumGet(galleryRect, 8, "Int") - 4 * spacing < 8)
        PreferencesWindow.PageList.Modify(appearancePage, "-Select")        ; a click on the blank area below the list: the current page is selected again on release
        click := Buffer(3 * A_PtrSize + 40, 0)                              ; NMITEMACTIVATE, iItem = -1 (没有点到行)
        NumPut("Int", -1, click, 3 * A_PtrSize)
        PreferencesWindow._OnPageClick(click.Ptr)
        eq("blank click keeps the page selected", PreferencesWindow.PageList.GetNext(), appearancePage)
        draw := Buffer(6 * A_PtrSize + 40, 0)                               ; NMCUSTOMDRAW: 选中行不画虚线焦点框
        NumPut("UInt", 0x1, draw, 3 * A_PtrSize)
        eq("custom draw asks for each row", PreferencesWindow._OnPageListDraw(draw.Ptr), 0x20)
        NumPut("UInt", 0x10001, draw, 3 * A_PtrSize), NumPut("UInt", 0x11, draw, 6 * A_PtrSize + 16)   ; CDIS_SELECTED | CDIS_FOCUS
        PreferencesWindow._OnPageListDraw(draw.Ptr)
        eq("row drawn without the focus rectangle", NumGet(draw, 6 * A_PtrSize + 16, "UInt"), 0x1)
        listRect := Buffer(16)                                              ; 唯一的一列占满列表宽度 (不留列分隔线)
        DllCall("GetClientRect", "Ptr", PreferencesWindow.PageList.Hwnd, "Ptr", listRect)
        eq("column fills the list", SendMessage(0x101D, 0, 0, PreferencesWindow.PageList), NumGet(listRect, 8, "Int"))   ; LVM_GETCOLUMNWIDTH
        PreferencesWindow.PageList.Modify(5, "Select Focus")                ; 用户在列表里选了一页 (点击或 ↑ ↓)
        Sleep(50)
        eq("selecting a row shows the page", PreferencesWindow._page, 5)
        wrong := 0                                                          ; 换页只显示这一页的控件, 别的页都隐藏
        for index, page in PreferencesWindow.Pages
            for ctrl in page.Controls
                wrong += (ctrl.Visible != (index = 5))
        eq("only the current page is visible", wrong, 0)
        eq("every page has an icon", NavIcons.MissingPages(PreferencesWindow.PageKeys), "")
        if (NavIcons.FontName() != "") {                                    ; Windows 上有图标字体 (Wine 没有)
            il := NavIcons.ImageList(PreferencesWindow.PageKeys)
            eq("one icon per page", DllCall("comctl32\ImageList_GetImageCount", "Ptr", il), PreferencesWindow.PageKeys.Length)
            DllCall("comctl32\ImageList_Destroy", "Ptr", il)
            eq("glyphs exist in " NavIcons.FontName(), NavIcons.MissingGlyphs(), "")
        }
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
        eq("builtin count", ThemeManager.Names().Length, 23)
        TestRunner.True("Themes.Ocean is builtin", ThemeManager.IsBuiltin("Ocean"))

        ; 用户主题: 同名覆盖内置主题 ("Base" 写自己 = 在内置那一套上改), 以及 Base 链
        DirCreate(ThemeManager.UserDir)
        FileAppend('{ "Base": "Ocean", "SelectedRadius": 12 }', ThemeManager.UserDir "\Mine.json", "UTF-8")
        FileAppend('{ "Base": "Dark", "Title": "FF0000" }', ThemeManager.UserDir "\Dark.json", "UTF-8")
        ThemeManager.Load("Mine")
        eq("user base color", ThemeManager.Get("Background"), "2E3440")
        eq("user override", ThemeManager.Get("SelectedRadius"), 12)
        TestRunner.True("Themes.user listed", ThemeManager.Names().Length = 24 && !ThemeManager.IsBuiltin("Mine"))
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

    ; 显示密度 (Appearance.Density): 任何主题的尺寸按比例缩放, 颜色不变; 以前的 DarkCompact 等当作 Dark + 常规
    static Densities() {
        eq := (n, a, e) => TestRunner.Equal("Densities." n, a, e)
        ThemeManager.BuiltinDir := A_ScriptDir "\..\Resources\Themes"
        ThemeManager.UserDir := A_Temp "\ALTRunDensityTest"
        try DirDelete(ThemeManager.UserDir, true)
        sizes := () => ThemeManager.Get("InputFontSize") "/" ThemeManager.Get("TitleFontSize") "/" ThemeManager.Get("SubtitleFontSize") "/"
            . ThemeManager.Get("ShortcutFontSize") "/" ThemeManager.Get("Padding") "/" ThemeManager.Get("RowHeight") "/" ThemeManager.Get("IconSize")
        ThemeManager.Load("Dark", "Comfortable")
        eq("comfortable = theme size", sizes() "|" ThemeManager.Density, "20/13/9.5/9/14/54/32|Comfortable")
        ThemeManager.Load("Dark", "Normal")
        eq("normal = old Dark Compact", sizes(), "17/11.5/9/8.5/12/48/28")
        eq("colors kept", ThemeManager.Get("Background") "|" ThemeManager.Resolved "|" ThemeManager.Density, "1E1F22|Dark|Normal")
        ThemeManager.Load("Dark", "Compact")
        eq("compact", sizes(), "15/10.5/8.5/8/10/42/24")
        ThemeManager.Load("Dark")
        eq("default comfortable", ThemeManager.Get("RowHeight") "|" ThemeManager.Density, "54|Comfortable")
        ThemeManager.Load("Dark", "Huge")
        eq("unknown = comfortable", ThemeManager.Get("RowHeight") "|" ThemeManager.Density, "54|Comfortable")
        eq("scales theme sizes", ThemeManager.Scaled(Map("RowHeight", 60, "IconSize", 40), "Normal")["RowHeight"] "|" ThemeManager.Scaled(Map("RowHeight", 60, "IconSize", 40), "Normal")["IconSize"], "53|36")   ; 自定义主题自己的尺寸也按比例缩小
        ; 内置主题只有颜色不同, 尺寸都和浅色一样: 同一档密度下切换主题, 窗口不会忽大忽小
        differ := ""
        for themeName in ThemeManager.Names() {
            if (themeName = "System")
                continue
            ThemeManager.Load(themeName, "Normal")
            if (sizes() != "17/11.5/9/8.5/12/48/28")
                differ .= themeName " " sizes() "; "
        }
        eq("builtin themes same size", differ, "")
        ThemeManager.Load("System", "Compact")
        eq("system compact", ThemeManager.Get("RowHeight"), 42)
        for oldName, base in Map("DarkCompact", "Dark", "LightCompact", "Light", "MidnightCompact", "Midnight") {
            ThemeManager.Load(oldName)
            eq("old " oldName, ThemeManager.Name "|" ThemeManager.Resolved "|" ThemeManager.Density "|" ThemeManager.Get("RowHeight"), base "|" base "|Normal|48")
            eq("old " oldName " not listed", ThemeManager._Contains(ThemeManager.Names(), oldName), false)
        }
        ThemeManager.Load("DarkCompact", "Normal")                          ; 不会缩小两次
        eq("old name + normal once", ThemeManager.Get("RowHeight"), 48)
        ThemeManager.Load("DarkCompact", "Compact")
        eq("old name + compact", ThemeManager.Get("RowHeight"), 42)

        ; 用户主题的 "Base" 还写着以前的紧凑主题: 照常是 Dark 的颜色、常规的尺寸; 用户自己有同名主题时用用户的
        DirCreate(ThemeManager.UserDir)
        FileAppend('{ "Base": "DarkCompact", "Title": "FF0000" }', ThemeManager.UserDir "\Mine.json", "UTF-8")
        ThemeManager.Load("Mine")
        eq("user base old name", ThemeManager.Get("Background") "|" ThemeManager.Get("Title") "|" ThemeManager.Get("RowHeight"), "1E1F22|FF0000|48")
        FileAppend('{ "Base": "Dark", "RowHeight": 40 }', ThemeManager.UserDir "\MidnightCompact.json", "UTF-8")
        eq("user file wins", ThemeManager.LegacyCompact("MidnightCompact"), "")
        ThemeManager.Load("MidnightCompact")
        eq("user file loaded", ThemeManager.Resolved "|" ThemeManager.Get("RowHeight") "|" ThemeManager.Density, "MidnightCompact|40|Comfortable")
        DirDelete(ThemeManager.UserDir, true)
        ThemeManager.Load("Light")

        ; 新安装默认常规; 已有的设置文件没有这一项时保持原来的大小 (宽松), 用以前的紧凑主题的是常规
        eq("new install default", AppSettings.Defaults()["Appearance"]["Density"], "Normal")
        keep := (theme) => (data := Map("Appearance", Map("Theme", theme)), AppSettings._KeepDensity(data), data["Appearance"]["Density"])
        eq("existing file keeps size", keep("Dark"), "Comfortable")
        eq("existing old compact theme", keep("darkcompact"), "Normal")
        set := Map("Appearance", Map("Theme", "Dark", "Density", "Compact"))
        eq("already set", AppSettings._KeepDensity(set) "|" set["Appearance"]["Density"], "0|Compact")
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
        eq("remote unc", IconCache.IsRemote("\\server\share\PT1001"), true)
        eq("remote local", IconCache.IsRemote("C:\Windows"), false)
        eq("remote folder key", IconCache._CacheKey("\\server\share\PT1001 - 24 Harbor"), "folder:")
        eq("remote folder slash", IconCache._CacheKey("\\server\share\PT1001\"), "folder:")
        eq("remote file key", IconCache._CacheKey("\\server\share\Report.PDF"), "ext:.pdf")
        eq("remote exe key", IconCache._CacheKey("\\server\share\tool.exe"), "ext:.exe")
        ; 名字里带点的文件夹不是 "扩展名" (以前变成空白图标)
        eq("remote dotted folder", IconCache._CacheKey("\\server\Proposals\2025\26. 18 Sample Road (A)"), "folder:")
        eq("remote dotted folder slash", IconCache._CacheKey("\\server\share\10.PT1002-29Harbor\"), "folder:")
        eq("local dotted folder", IconCache._CacheKey("C:\Projects\26. 18 Sample Road (A)"), "c:\projects\26. 18 sample road (a)")
        eq("extension", IconCache._Extension("C:\a\Report.PDF"), "PDF")
        eq("extension with spaces", IconCache._Extension("C:\a\26. 18 New Road"), "")
        eq("extension appref-ms", IconCache._Extension("C:\a\App.appref-ms"), "appref-ms")
        eq("extension 7z", IconCache._Extension("C:\a\backup.7z"), "7z")
        eq("extension sldprt", IconCache._Extension("C:\a\Beam.SLDPRT"), "SLDPRT")
        eq("extension digits only", IconCache._Extension("C:\a\Design.2019"), "")
        eq("folder command icon remote", CustomCommandProvider._FolderIcon("\\server\share\26. 18 Sample Road (A)"), "folder:")
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
        DirCreate(dir "\PT1001 - 24 Harbor")
        FileAppend("", dir "\summary.pdf")
        check(type, target, roots := "") => CustomCommandProvider.CheckTarget(Map("Title", "t", "Type", type, "Target", target), roots)
        eq("folder ok", check("Folder", dir "\PT1001 - 24 Harbor"), "OK")
        eq("folder renamed", check("Folder", dir "\PT1001 - 24 Harbor (old)"), "Missing")
        eq("folder trailing slash", check("Folder", dir "\PT1001 - 24 Harbor\"), "OK")
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
        eq("drive offline", check("Folder", "Q:\Projects\PT1001", offline), "Unavailable")
        eq("unc offline", check("File", "\\server\share\a.pdf", Map("\\server\share", false)), "Unavailable")
        roots := Map()
        check("Folder", dir, roots)
        SplitPath(dir, , , , , &drive)
        eq("root remembered", roots.Has(drive) && roots[drive], true)
        try DirDelete(dir, true)
    }

    ; 大字显示可以在 偏好设置 → 功能 里关掉: 新安装默认关, 以前的设置文件保持打开; 关掉时操作面板和速查表里都没有
    static LargeTypeSwitch() {
        eq := (n, a, e) => TestRunner.Equal("LargeTypeSwitch." n, a, e)
        eq("new install off", AppSettings.Defaults()["Features"]["LargeType"]["Enabled"], 0)
        old := Map("Features", Map("Help", Map("Enabled", 1)))
        eq("existing file keeps it on", AppSettings._KeepPrevious(old) "|" old["Features"]["LargeType"]["Enabled"], "1|1")
        set := Map("Features", Map("LargeType", Map("Enabled", 0)))
        eq("already set", AppSettings._KeepLargeType(set) "|" set["Features"]["LargeType"]["Enabled"], "0|0")
        options := AppSettings.Feature("LargeType"), saved := options["Enabled"]
        hasAction() {
            for action in ActionCatalog.ListFor(ResultItem("42", "", {Kind: "text", Arg: "42"}))
                if (action.Title = I18n.T("Action.LargeType"))
                    return 1
            return 0
        }
        helpHas() {
            for entry in HelpProvider.Items()
                if (entry.Id = "LargeType")
                    return 1
            return 0
        }
        options["Enabled"] := 0
        eq("off: no action, no help", hasAction() "|" helpHas() "|" LargeType.Enabled(), "0|0|0")
        options["Enabled"] := 1
        eq("on: action and help", hasAction() "|" helpHas() "|" LargeType.Enabled(), "1|1|1")
        options["Enabled"] := saved
    }

    ; 关键字一栏写几个 (逗号分隔): 网页搜索、自定义命令、片段输入哪个都行, 提示和补全用输入的那个 (没有时用第一个)
    static MultipleKeywords() {
        eq := (n, a, e) => TestRunner.Equal("MultipleKeywords." n, a, e)
        eq("split", Tests._Join(SearchQuery.Keywords(" g , ㅎ,,google，구글 ")), "g,ㅎ,google,구글")
        eq("empty", SearchQuery.Keywords("").Length "|" SearchQuery.Keywords(" , ").Length, "0|0")
        eq("matched keyword", SearchQuery("ㅎ hello").MatchKeyword(["g", "ㅎ"], &term, &matched) ? matched "|" term : "no", "ㅎ|hello")

        options := AppSettings.Feature("WebSearch"), savedEngines := options["Engines"]
        options["Engines"] := [Map("Id", "multi", "Keyword", "gg, ㅎㅎ", "Title", "Multi", "Url", "https://x.test/?q={query}")]
        items := WebSearchProvider.Search(SearchQuery("ㅎㅎ hello"))
        eq("web: second keyword", items.Length "|" items[1].Arg "|" items[1].AutoComplete, "1|https://x.test/?q=hello|ㅎㅎ hello")
        items := WebSearchProvider.Search(SearchQuery("GG"))
        eq("web: first keyword, any case", items.Length "|" items[1].Valid, "1|0")
        eq("web: waiting hint", items[1].Subtitle, I18n.T("Search.TypeAfter", "gg"))
        eq("web: whole text is not a keyword", WebSearchProvider.Search(SearchQuery("gg, x")).Length, 0)
        eq("web: fallback uses first", WebSearchProvider.ItemFor(options["Engines"][1], "a").AutoComplete, "gg a")
        eq("web: help example", HelpProvider._Args("WebSearch")[1], "gg")
        options["Engines"] := savedEngines

        saved := AppSettings.Data["CustomCommands"]
        AppSettings.Data["CustomCommands"] := [
            Map("Title", "Notepad", "Type", "File", "Target", "notepad.exe", "Arguments", "", "Keyword", "np, ㅜㅔ"),
            Map("Title", "Jira", "Type", "Url", "Target", "https://jira.test/{query}", "Arguments", "", "Keyword", "jira, ㅓㅑㄱㅁ")]
        CustomCommandProvider._ResetNarrowing()
        top(text) {
            items := ProviderRegistry.SortByScore(CustomCommandProvider.Search(SearchQuery(text)))
            return items.Length ? items[1].Title "|" items[1].Score : ""
        }
        eq("command: first keyword", top("np"), "Notepad|115")
        eq("command: second keyword", top("ㅜㅔ"), "Notepad|115")
        items := CustomCommandProvider.Search(SearchQuery("ㅓㅑㄱㅁ ABC-1"))
        eq("command: query after second keyword", items.Length "|" items[1].Arg "|" items[1].AutoComplete, "1|https://jira.test/ABC-1|ㅓㅑㄱㅁ ABC-1")
        item := CustomCommandProvider._ToItem(AppSettings.Data["CustomCommands"][2], 50)
        eq("command: hint uses first", item.AutoComplete "|" (InStr(item.Subtitle, I18n.T("Search.TypeAfter", "jira")) = 1), "jira |1")
        AppSettings.Data["CustomCommands"] := saved
        CustomCommandProvider._ResetNarrowing()

        savedSnippets := AppSettings.Data["Snippets"]
        AppSettings.Data["Snippets"] := [Map("Name", "Signature", "Keyword", "sig, ㄴㅑㅎ", "Text", "Best regards", "AutoExpand", 1)]
        found := (text) => SnippetProvider.Search(SearchQuery(text)).Length
        eq("snippet: each keyword", found("sig") "|" found("ㄴㅑㅎ") "|" found("signature"), "1|1|1")
        AppSettings.Data["Snippets"] := savedSnippets
    }

    static CommandTargets() {
        eq := (n, a, e) => TestRunner.Equal("CommandTargets." n, a, e)
        saved := AppSettings.Data["CustomCommands"]
        AppSettings.Data["CustomCommands"] := [
            Map("Title", "Tower A, B, C", "Type", "Folder", "Target", "Q:\Projects\Design-2019\PT1001 - 24 Harbor", "Arguments", "", "Keyword", ""),
            Map("Title", "Harbor Report", "Type", "File", "Target", "Q:\Docs\summary.pdf", "Arguments", "", "Keyword", ""),
            Map("Title", "Check IP", "Type", "Command", "Target", "cmd.exe", "Arguments", "/k ipconfig", "Keyword", ""),
            Map("Title", "Drive Q", "Type", "Folder", "Target", "Q:\", "Arguments", "", "Keyword", "")]
        titles(text) {
            list := ""
            for item in ProviderRegistry.SortByScore(CustomCommandProvider.Search(SearchQuery(text)))
                list .= item.Title "|"
            return RTrim(list, "|")
        }
        eq("folder name", titles("1001"), "Tower A, B, C")
        eq("title first", titles("harbor"), "Harbor Report|Tower A, B, C")
        eq("file name without extension", titles("summary"), "Harbor Report")
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

    ; 索引的文件夹里有损坏的快捷方式: 照样收进索引 (没有目标), 扫描不中断 (曾经报 "linkTarget 未赋值")
    static BrokenShortcut() {
        eq := (n, a, e) => TestRunner.Equal("BrokenShortcut." n, a, e)
        root := A_Temp "\ALTRunBrokenShortcut"
        try DirDelete(root, true)
        DirCreate(root)
        FileAppend("not a shortcut", root "\Broken.lnk")
        FileAppend("", root "\Editor.exe")
        options := AppSettings.Feature("Applications")
        saved := Map("Folders", options["Folders"], "StoreApps", options["StoreApps"])
        savedFile := ApplicationProvider.IndexFile, savedApps := ApplicationProvider.Apps
        options["Folders"] := [root], options["StoreApps"] := 0
        ApplicationProvider.IndexFile := A_Temp "\ALTRunBrokenShortcut.json"
        try {
            eq("count", ApplicationProvider.Rebuild(), 2)
            titles := ""
            for entry in ApplicationProvider.Apps
                titles .= entry["Title"] "|"
            eq("titles", titles, "Broken|Editor|")
        } finally {
            options["Folders"] := saved["Folders"], options["StoreApps"] := saved["StoreApps"]
            ApplicationProvider.IndexFile := savedFile, ApplicationProvider.Apps := savedApps
            try FileDelete(A_Temp "\ALTRunBrokenShortcut.json")
            try DirDelete(root, true)
        }
    }

    ; 偏好设置里改了索引的扫描范围: 保存后让缓存过期, 重新启动后在后台重新扫描
    static StaleIndexes() {
        eq := (n, a, e) => TestRunner.Equal("StaleIndexes." n, a, e)
        before := AppSettings.Defaults()
        stale(change) {
            after := PreferencesWindow.DeepCopy(before)
            change(after)
            result := PreferencesWindow.StaleIndexes(before, after)
            return result.Apps "/" result.Files
        }
        eq("nothing changed", stale((data) => 0), "0/0")
        eq("app depth", stale((data) => data["Features"]["Applications"]["Depth"] := 5), "1/0")
        eq("app folders", stale((data) => data["Features"]["Applications"]["Folders"].Push("D:\Tools")), "1/0")
        eq("app hidden list is not the scan range", stale((data) => data["Features"]["Applications"]["Hidden"].Push("x")), "0/0")
        eq("file exclude", stale((data) => data["Features"]["FileSearch"]["ScopeExclude"] := "i)temp"), "0/1")
        eq("file max entries", stale((data) => data["Features"]["FileSearch"]["MaxEntries"] := 1000), "0/1")
        eq("refresh interval is not the scan range", stale((data) => data["Features"]["FileSearch"]["RefreshMinutes"] := 30), "0/0")
        savedApp := ApplicationProvider.IndexFile, savedFile := FileIndex.File
        ApplicationProvider.IndexFile := A_Temp "\ALTRunStaleApp.json", FileIndex.File := A_Temp "\ALTRunStaleFile.json"
        try {
            for path in [ApplicationProvider.IndexFile, FileIndex.File] {
                try FileDelete(path)
                FileAppend("{}", path)
            }
            eq("fresh caches", ApplicationProvider._CacheExpired() || FileIndex._CacheExpired(), false)
            FileIndex._savePending := true
            ApplicationProvider.ExpireCache(), FileIndex.ExpireCache()
            eq("expired", ApplicationProvider._CacheExpired() && FileIndex._CacheExpired(), true)
            eq("pending save dropped", FileIndex._savePending, false)
            ApplicationProvider.ExpireCache()                               ; 没有缓存文件也不报错
        } finally {
            FileIndex._savePending := false
            try FileDelete(ApplicationProvider.IndexFile)
            try FileDelete(FileIndex.File)
            ApplicationProvider.IndexFile := savedApp, FileIndex.File := savedFile
        }
    }

    static AppIndexRefresh() {
        eq := (n, a, e) => TestRunner.Equal("AppIndexRefresh." n, a, e)
        root := A_Temp "\ALTRunAppIndexTest"
        try DirDelete(root, true)
        for dir in ["", "\Office", "\Office\Tools", "\Office\Tools\More", "\Office\Tools\More\TooDeep", "\Games"]
            DirCreate(root dir)                                             ; 一层一层建 (Wine 里一次建多层会卡住)
        for relative in ["\Editor.exe", "\Office\Writer.exe", "\Office\Uninstall Writer.exe", "\Office\Tools\Sheet.url"
                , "\Office\Tools\More\Draw.exe", "\Office\Tools\More\TooDeep\Hidden.exe", "\Games\Chess.exe", "\Games\notes.txt"]
            FileAppend("", root relative)
        options := AppSettings.Feature("Applications")
        saved := Map("Folders", options["Folders"], "StoreApps", options["StoreApps"])
        savedFile := ApplicationProvider.IndexFile, savedApps := ApplicationProvider.Apps
        options["Folders"] := [root], options["StoreApps"] := 0
        ApplicationProvider.IndexFile := A_Temp "\ALTRunAppIndexTest.json"
        try FileDelete(ApplicationProvider.IndexFile)
        titles() {
            list := ""
            for entry in ApplicationProvider.Apps
                list .= entry["Title"] "|"
            return RTrim(list, "|")
        }
        expected := "Chess|Draw|Editor|Sheet|Writer"                       ; 深度 4 的 Hidden、Exclude 里的 Uninstall、txt 不要

        eq("no cache = expired", ApplicationProvider._CacheExpired(), true)
        eq("manual count", ApplicationProvider.Rebuild(), 5)
        eq("manual", titles(), expected)
        eq("fresh cache", ApplicationProvider._CacheExpired(), false)
        FileSetTime(DateAdd(A_Now, -(ApplicationProvider.FallbackMinutes + 1), "Minutes"), ApplicationProvider.IndexFile, "M")
        eq("older than the fallback interval", ApplicationProvider._CacheExpired(), true)

        ApplicationProvider.Apps := []
        job := ApplicationProvider._NewJob(), steps := 1
        while !ApplicationProvider._Step(job, -1)                          ; -1: 每扫一个文件夹就停一次, 下次接着扫
            steps++
        eq("background in steps", steps > 5, true)
        eq("not replaced while scanning", titles(), "")
        ApplicationProvider._Finish(job, false)
        eq("background", titles(), expected)

        ; 监视文件夹: 有新文件时 (安静一会儿后) 在后台重建; 后台重建的缓存等空闲再写
        savedQuiet := ApplicationProvider.WatchQuietMs, savedGap := ApplicationProvider.WatchMinGapMs
        ApplicationProvider.WatchQuietMs := 0, ApplicationProvider.WatchMinGapMs := 0
        ApplicationProvider.StartWatching()
        eq("watching the folder", ApplicationProvider._watcher.Count, 1)
        eq("watch read pending", ApplicationProvider._watcher.Watches[1].Pending, true)
        ApplicationProvider._PollWatches()
        eq("no change, no rebuild", IsObject(ApplicationProvider._job), false)
        FileAppend("", root "\notes2.txt")                                  ; 不是应用的文件: 不重建
        Sleep(300)
        ApplicationProvider._PollWatches()
        eq("other file, no rebuild", IsObject(ApplicationProvider._job), false)
        FileAppend("", root "\Solitaire.exe")
        Sleep(300)
        ApplicationProvider._PollWatches()
        eq("new file starts a rebuild", IsObject(ApplicationProvider._job), true)
        job := ApplicationProvider._job
        ApplicationProvider._StopJob()
        if !DllCall("GetProcAddress", "Ptr", DllCall("GetModuleHandle", "Str", "ntdll", "Ptr"), "AStr", "wine_get_version", "Ptr") {
            ApplicationProvider._lastBackground := 0                        ; 子文件夹里的变化也收得到 (开始菜单的快捷方式大多在子文件夹里; Wine 不支持)
            FileAppend("", root "\Games\Chess2.exe")
            Sleep(300)
            ApplicationProvider._PollWatches()
            eq("change in a subfolder starts a rebuild", IsObject(ApplicationProvider._job), true)
            ApplicationProvider._StopJob()
            job := ApplicationProvider._NewJob()
        }
        while !ApplicationProvider._Step(job, 0)
            continue
        modified := FileGetTime(ApplicationProvider.IndexFile, "M")
        ApplicationProvider._Finish(job, false)
        eq("new app found", InStr(titles(), "Solitaire") > 0, true)
        eq("background save waits", ApplicationProvider._savePending && FileGetTime(ApplicationProvider.IndexFile, "M") = modified, true)
        ApplicationProvider.SavePending()
        eq("saved", ApplicationProvider._savePending, false)
        ApplicationProvider.StopWatching()
        types := ApplicationProvider._TypesPattern(options)
        dotFolder := root "\Vendor.Office"
        DirCreate(dotFolder)
        eq("watch: folder with extension", ApplicationProvider.IsWatchedChange(dotFolder, types, "D:\ALTRun"), true)
        DirDelete(dotFolder)
        eq("watch: shortcut", ApplicationProvider.IsWatchedChange("C:\Start\Tool.lnk", types, "D:\ALTRun"), true)
        eq("watch: folder (no extension)", ApplicationProvider.IsWatchedChange("C:\Start\Vendor", types, "D:\ALTRun"), true)
        eq("watch: other file", ApplicationProvider.IsWatchedChange("C:\Desktop\Report.docx", types, "D:\ALTRun"), false)
        eq("watch: ALTRun's own folder", ApplicationProvider.IsWatchedChange("C:\Desktop\ALTRun\Data\AppIndex.json", types, "C:\Desktop\ALTRun"), false)
        eq("watch: ALTRun's own exe", ApplicationProvider.IsWatchedChange("C:\Desktop\ALTRun\ALTRun.exe", types, "C:\Desktop\ALTRun"), false)
        ApplicationProvider.WatchQuietMs := savedQuiet, ApplicationProvider.WatchMinGapMs := savedGap
        if (ApplicationProvider._saveTimer != "")
            SetTimer(ApplicationProvider._saveTimer, 0)

        for key, value in saved
            options[key] := value
        ApplicationProvider.IndexFile := savedFile, ApplicationProvider.Apps := savedApps
        try FileDelete(A_Temp "\ALTRunAppIndexTest.json")
        try DirDelete(root, true)
    }

    ; 故障后的目录监视器应能重新挂载，并在关闭时安全结束待处理的异步 I/O
    static WatcherRecovery() {
        eq := (n, a, e) => TestRunner.Equal("WatcherRecovery." n, a, e)
        root := A_Temp "\ALTRunWatcherRecoveryTest"
        try DirDelete(root, true)
        DirCreate(root)
        watcher := ""
        try {
            watcher := FolderWatcher([root, root "\missing"])
            eq("missing folder skipped", watcher.Count, 1)
            if !watcher.Count
                return
            watch := watcher.Watches[1]
            eq("initial read pending", watch.Pending, true)
            eq("nothing yet", watcher.Poll().Length, 0)

            ; 模拟再开始等失败后停掉的文件夹: 到时间重新打开, 报告 "*" (停掉期间可能漏了变化), 之后照常监视
            watch.Disabled := true, watch.RetryAt := A_TickCount + 60000
            eq("waits for retry", watcher.Poll().Length, 0)
            watch.RetryAt := 0
            changes := watcher.Poll()
            recovered := watcher.Watches[1]
            eq("disabled watch replaced", recovered != watch, true)
            eq("recovery reports unknown changes", changes.Length = 1 && changes[1] = "*", true)
            eq("recovered read pending", recovered.Pending && !recovered.Disabled, true)
            FileAppend("", root "\Recovered.exe")
            Sleep(300)
            changes := watcher.Poll()
            eq("recovered watch reports the file", changes.Length >= 1 && changes[1] = root "\Recovered.exe", true)

            ; 关闭还在等的读取: 等它结束后才放掉缓冲区, 多次打开 / 关闭都正常
            for i in [1, 2, 3] {
                pendingWatch := FolderWatcher.Create(root)
                eq("close cycle " i " created", IsObject(pendingWatch) && pendingWatch.Pending, true)
                if !IsObject(pendingWatch)
                    continue
                FolderWatcher.Close(pendingWatch)
                eq("close cycle " i " completed", pendingWatch.Pending || pendingWatch.Handle || pendingWatch.Event, false)
                FolderWatcher.Close(pendingWatch)                           ; 再关一次也没事
            }
            eq("create missing", FolderWatcher.Create(root "\missing"), "")
        } finally {
            if IsObject(watcher)
                watcher.Stop()
            try DirDelete(root, true)
        }
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
        PreferencesWindow.BuildAll()                                         ; 页面用到时才建: 检查每一页之前先全部建好
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

    ; Everything 打开拼音后找到的中文文件名: 按拼音首字母排序 (比同一档的直接匹配低一点), 高亮对应的字
    static FileSearchPinyin() {
        eq := (n, a, e) => TestRunner.Equal("FileSearchPinyin." n, a, e)
        score := (needle, name) => FileSearchProvider.NameScore(needle, name)
        eq("pinyin exact", score("cs.bat", "测试.bat"), 95)
        eq("pinyin prefix", Round(score("cs", "测试.bat"), 1), 84.6)
        eq("pinyin word start", score("bg", "项目-报告.docx") > 70 && score("bg", "项目-报告.docx") < 80, true)
        eq("name still first", score("cs.bat", "cs.bat"), 100)
        order := [score("cs.bat", "cs.bat"), score("cs.bat", "测试.bat"), score("cs", "cs_report.bat"), score("cs", "测试.bat")]
        eq("order: exact > pinyin exact > prefix > pinyin prefix", order[1] > order[2] && order[2] > order[3] && order[3] > order[4], true)
        eq("chinese input", Round(score("测试", "测试.bat"), 1), 89.6)
        eq("no match", score("report", "测试.bat"), 50)
        eq("no chinese in the name", score("cs", "report.bat"), 50)
        eq("chinese needle: no pinyin", score("测s", "测试.bat"), 50)
        eq("several words: no pinyin", score("cs bat", "测 试.bat"), 50)
        eq("initials cached", Pinyin.Initials("测试") Pinyin.Initials("试测") Pinyin._cache.Has("测"), "CSSC1")
        eq("unknown character", Pinyin.Initials("A龘B"), "AB")

        ; 内置索引 (没有 Everything 时) 也按拼音首字母找
        paths := ["C:\D\测试.bat", "C:\D\cs_report.bat", "C:\D\证件扫描文件.zip", "C:\D\项目测试报告.docx", "C:\D\notes.txt"]
        FileIndex.Paths := paths, FileIndex.Names := [], FileIndex.Folders := [0, 0, 0, 0, 0]
        for fullPath in paths {
            SplitPath(fullPath, &fileName)
            FileIndex.Names.Push(StrLower(fileName))
        }
        eq("index: before initials are ready", names("zjsm"), "")
        FileIndex.BuildInitials()                                             ; 后台分批算
        eq("index: initials built in the background", Tests._WaitFor(() => FileIndex._initialsFor == FileIndex.Names), true)
        eq("index initials", FileIndex.Initials.Count "|" FileIndex.Initials[1] "|" FileIndex.Initials.Has(2) "|" FileIndex.Initials[3], "3|cs.bat|0|zjsmwj.zip")
        names(needle, fresh := true) {
            if fresh
                FileIndex._lastNeedle := "", FileIndex._lastMatches := ""
            list := ""
            for item in FileIndex.Search(needle, 10) {
                SplitPath(item.Path, &fileName)
                list .= fileName "|"
            }
            return RTrim(list, "|")
        }
        eq("index: direct first, then pinyin", names("cs"), "cs_report.bat|测试.bat|项目测试报告.docx")
        eq("index: initials", names("zjsm"), "证件扫描文件.zip")
        eq("index: narrowing keeps pinyin", (names("c"), names("cs", false)), "cs_report.bat|测试.bat|项目测试报告.docx")
        eq("index: chinese input", names("测试"), "测试.bat|项目测试报告.docx")
        eq("index: several words", names("zj sm"), "")
        FileIndex.Names := FileIndex.Names.Clone()                            ; 换了 Names (例如测试直接设的): 首字母对不上, 不按拼音
        eq("index: initials for other names", names("zjsm"), "")
        savedFile := FileIndex.File, FileIndex.File := A_Temp "\ALTRunPinyinIndex.json"
        try {
            JSON.WriteFile(FileIndex.File, Map("Paths", ["C:\D\证件扫描文件.zip", "C:\D\notes.txt"], "Folders", [0, 0]))
            FileIndex._LoadCache()
            Tests._WaitFor(() => FileIndex._initialsFor == FileIndex.Names)
            eq("cache load builds initials", FileIndex.Initials.Count "|" FileIndex.Initials[1], "1|zjsmwj.zip")
        } finally {
            try FileDelete(FileIndex.File)
            FileIndex.File := savedFile
        }
        FileIndex.Paths := [], FileIndex.Names := [], FileIndex.Folders := [], FileIndex.Initials := Map(), FileIndex._initialsFor := ""
        FileIndex._lastNeedle := "", FileIndex._lastMatches := ""
        ranges := FuzzyMatcher.MatchRanges("cs", "测试.bat")
        eq("highlight", ranges.Length = 1 ? ranges[1][1] "," ranges[1][2] : ranges.Length, "1,2")
    }

    static FileSearchModes() {
        eq := (n, a, e) => TestRunner.Equal("FileSearchModes." n, a, e)
        options := AppSettings.Feature("FileSearch")
        saved := [options["InDefaultResults"], options["UseEverything"]]
        options["UseEverything"] := 0                                       ; the tests use the built-in index
        FileIndex.Paths := ["C:\Windows\Fonts\MARLETT.TTF", "C:\Docs\Market Notes.pdf"]
        FileIndex.Names := ["marlett.ttf", "market notes.pdf"]
        FileIndex.Folders := [0, 0]
        FileIndex._lastNeedle := "", FileIndex._lastMatches := ""
        eq("default off", AppSettings.Defaults()["Features"]["FileSearch"]["InDefaultResults"], 0)
        eq("space prefix on", AppSettings.Defaults()["Features"]["FileSearch"]["SpacePrefix"], 1)
        options["InDefaultResults"] := 0
        eq("normal search has no files", FileSearchProvider.Search(SearchQuery("mar")).Length, 0)
        files := ProviderRegistry.SearchFiles("mar")
        TestRunner.True("FileSearchModes.file mode finds files", files.Length >= 2 && files[1].Provider = "FileSearch")
        eq("file mode empty", ProviderRegistry.SearchFiles("  ").Length, 0)
        TestRunner.True("FileSearchModes.quote still works", FileSearchProvider.Search(SearchQuery("'mar")).Length >= 2)
        options["InDefaultResults"] := saved[1], options["UseEverything"] := saved[2]
        FileIndex.Paths := [], FileIndex.Names := [], FileIndex.Folders := []
    }

    ; "folder atlas" searches folders only; a keyword without a following space does not hide other results; results are sorted by name match
    static FolderSearch() {
        eq := (n, a, e) => TestRunner.Equal("FolderSearch." n, a, e)
        options := AppSettings.Feature("FileSearch")
        savedEverything := options["UseEverything"]
        options["UseEverything"] := 0                                       ; the tests use the built-in index
        FileIndex.Paths := ["C:\Docs\noteatlas.pdf", "C:\Projects\Atlas Tower", "C:\Docs\Atlas drawing.dwg", "C:\Projects\Old Atlas"]
        FileIndex.Names := ["noteatlas.pdf", "atlas tower", "atlas drawing.dwg", "old atlas"]
        FileIndex.Folders := [0, 1, 0, 1]
        FileIndex._lastNeedle := "", FileIndex._lastMatches := ""
        titles(results) {
            list := ""
            for item in results
                if (item.Kind = "file" || item.Kind = "folder")             ; not the trailing "Search with Everything / Windows"
                    list .= item.Title "|"
            return RTrim(list, "|")
        }
        eq("default keyword", AppSettings.Defaults()["Features"]["FileSearch"]["FolderKeywords"][1], "folder")
        eq("folders only", titles(FileSearchProvider.Search(SearchQuery("folder atlas"))), "Atlas Tower|Old Atlas")
        eq("folders exclusive", FileSearchProvider.Search(SearchQuery("folder atlas"))[1].Exclusive, true)
        eq("files keyword", titles(FileSearchProvider.Search(SearchQuery("open atlas"))), "Atlas Tower|Atlas drawing.dwg|Old Atlas|noteatlas.pdf")
        eq("file mode folder keyword", titles(ProviderRegistry.SearchFiles("folder atlas")), "Atlas Tower|Old Atlas")
        eq("file mode plain", titles(ProviderRegistry.SearchFiles("atlas")), "Atlas Tower|Atlas drawing.dwg|Old Atlas|noteatlas.pdf")
        ; Only the keyword (no space yet): the hint ranks low and does not hide commands with folder in their name
        hint := FileSearchProvider.Search(SearchQuery("folder"))
        eq("keyword alone hint", hint.Length, 1)
        eq("keyword alone not exclusive", hint[1].Exclusive, false)
        eq("keyword alone low score", hint[1].Score < 10, true)
        eq("open alone not exclusive", FileSearchProvider.Search(SearchQuery("open"))[1].Exclusive, false)
        eq("keyword with space exclusive", FileSearchProvider.Search(SearchQuery("folder "))[1].Exclusive, true)
        eq("quote alone exclusive", FileSearchProvider.Search(SearchQuery("'"))[1].Exclusive, true)
        ; Everything results come by modification time: name-start / word-start matches first, folders first on ties
        found := [{Path: "C:\a\noteatlas.pdf", IsFolder: 0, Score: FileIndex.ScoreName("atlas", "noteatlas.pdf")}
                , {Path: "C:\a\Atlas x.dwg", IsFolder: 0, Score: FileIndex.ScoreName("atlas", "atlas x.dwg")}
                , {Path: "C:\a\Atlas x.pdf", IsFolder: 1, Score: FileIndex.ScoreName("atlas", "atlas x.pdf")}]
        best := FileSearchProvider._Best(found, 3)
        eq("best order", best[1].Path "|" best[2].Path "|" best[3].Path, "C:\a\Atlas x.pdf|C:\a\Atlas x.dwg|C:\a\noteatlas.pdf")
        eq("everything fetches more", FileSearchProvider.EverythingFetch >= 300, true)
        eq("score needle folder:", FileSearchProvider.ScoreNeedle("folder:atlas"), "atlas")
        eq("score needle ext:", FileSearchProvider.ScoreNeedle("ext:pdf report"), "pdf report")
        eq("score needle plain", FileSearchProvider.ScoreNeedle("atlas tower"), "atlas tower")
        eq("only Everything syntax (d:)", FileSearchProvider.NameScore(FileSearchProvider.ScoreNeedle("d:"), "AhkScripts"), 50)
        eq("only Everything syntax (folder:)", FileSearchProvider.NameScore(FileSearchProvider.ScoreNeedle("folder:"), "AhkScripts"), 50)
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
        ; The most useful entries come first
        eq("order", InStr(ids, "Files Folders FileKeyword FileTypes Reveal Actions QuickSwitch About Prefs Edit EditJson ") = 1, true)
        ; Short keyword lists for the tips: preferred ones first, then the rest in order
        eq("keywords preferred", HelpProvider._Keywords(["open", "find", "file"], ["file", "open"], 2), "file / open")
        eq("keywords changed by user", HelpProvider._Keywords(["find", "search"], ["file", "open"], 2), "find / search")
        eq("keywords type filters", HelpProvider._Keywords(["doc", "pic", "video", "zip"], ["doc", "zip"], 2), "doc / zip")
        eq("keywords none", HelpProvider._Keywords([], ["doc"], 2), "")
        savedLang := I18n.Lang
        I18n.Init("en")
        keyOf(id) {
            for item in HelpProvider.Items()
                if (item.Id = id)
                    return item.Key
            return ""
        }
        options := AppSettings.Feature("FileSearch"), savedKeywords := options["Keywords"]
        options["Keywords"] := ["open", "find", "file"]
        eq("file keyword key", keyOf("FileKeyword"), "file / open / ' + name")
        options["Keywords"] := savedKeywords
        quick := AppSettings.Extension("QuickSwitch"), savedQuick := [quick["Enabled"], quick["TotalCmdHotkey"], quick["ExplorerHotkey"]]
        quick["Enabled"] := 1, quick["TotalCmdHotkey"] := "^g", quick["ExplorerHotkey"] := "^e"
        eq("quick switch key", keyOf("QuickSwitch"), "Ctrl+G / Ctrl+E (in an Open / Save dialog)")
        quick["TotalCmdHotkey"] := "", quick["ExplorerHotkey"] := ""
        eq("quick switch without hotkeys", keyOf("QuickSwitch"), "")
        quick["Enabled"] := savedQuick[1], quick["TotalCmdHotkey"] := savedQuick[2], quick["ExplorerHotkey"] := savedQuick[3]
        I18n.Init(savedLang)
        savedLang := I18n.Lang                                              ; 拼音提示只给简体中文界面
        hasPinyin(lang) {
            I18n.Init(lang)
            for item in HelpProvider.Items()
                if (item.Id = "Pinyin")
                    return item.Key " " item.Url
            return ""
        }
        eq("pinyin tip in zh-CN", hasPinyin("zh-CN"), "cs " HelpProvider.WikiUrl "File-Search-zh")
        eq("no pinyin tip in zh-TW / en", hasPinyin("zh-TW") hasPinyin("en"), "")
        I18n.Init(savedLang)
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
        eq("configured keyword", InStr(folderItem.Key, AppSettings.Feature("FileSearch")["FolderKeywords"][1] " + ") = 1, true)
        eq("wiki page", folderItem.Url, HelpProvider.WikiPage("File-Search"))
        eq("wiki page in English", HelpProvider.WikiPage("File-Search", "en"), "https://github.com/zhugecaomao/ALTRun/wiki/File-Search")
        eq("wiki page in Japanese UI", HelpProvider.WikiPage("Usage", "ja"), "https://github.com/zhugecaomao/ALTRun/wiki/Usage")
        eq("wiki page in Chinese", HelpProvider.WikiPage("File-Search", "zh-CN"), "https://github.com/zhugecaomao/ALTRun/wiki/File-Search-zh")
        eq("wiki page in Traditional Chinese", HelpProvider.WikiPage("Usage", "zh-TW"), "https://github.com/zhugecaomao/ALTRun/wiki/Usage-zh")
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
        eq("wiki general", PreferencesWindow.WikiPage("Prefs.Page.General"), "Configuration")
        eq("wiki window", PreferencesWindow.WikiPage("Prefs.Page.Window"), "Usage")
        eq("wiki keyboard", PreferencesWindow.WikiPage("Prefs.Page.Keyboard"), "Hotkeys")
        eq("wiki appearance", PreferencesWindow.WikiPage("Prefs.Page.Appearance"), "Themes")
        eq("wiki features", PreferencesWindow.WikiPage("Prefs.Page.Features"), "Usage")
        eq("wiki applications", PreferencesWindow.WikiPage("Prefs.Page.Applications"), "Configuration")
        eq("wiki file search", PreferencesWindow.WikiPage("Prefs.Page.FileSearch"), "File-Search")
        eq("wiki file index", PreferencesWindow.WikiPage("Prefs.Page.FileIndex"), "File-Search")
        eq("wiki commands", PreferencesWindow.WikiPage("Prefs.Page.Commands"), "Custom-Commands")
        eq("wiki snippets", PreferencesWindow.WikiPage("Prefs.Page.Snippets"), "Snippets")
        eq("wiki clipboard", PreferencesWindow.WikiPage("Prefs.Page.Clipboard"), "Clipboard")
        eq("wiki web search", PreferencesWindow.WikiPage("Prefs.Page.WebSearch"), "Web-Search")
        eq("wiki calculator", PreferencesWindow.WikiPage("Prefs.Page.Calculator"), "Calculator")
        eq("wiki scripts", PreferencesWindow.WikiPage("Prefs.Page.Scripts"), "Scripts")
        eq("wiki hotkeys", PreferencesWindow.WikiPage("Prefs.Page.Hotkeys"), "Hotkeys")
        eq("wiki quick switch", PreferencesWindow.WikiPage("Prefs.Page.QuickSwitch"), "Quick-Switch")
        eq("wiki dialog panel", PreferencesWindow.WikiPage("Prefs.Page.QSPanel"), "Quick-Switch-Panel")
        eq("wiki date stamp", PreferencesWindow.WikiPage("Prefs.Page.DateStamp"), "Date-Stamp")
        eq("wiki usage", PreferencesWindow.WikiPage("Prefs.Page.Usage"), "Usage")
        eq("wiki advanced default", PreferencesWindow.WikiPage("Prefs.Page.Advanced"), "Configuration")
        for key in ["Prefs.OK", "Prefs.Cancel", "Prefs.Apply", "Prefs.Help", "Prefs.DiscardChanges"]
            eq("text " key, I18n.T(key) != key, true)
    }

    ; 搜索窗口的位置: 默认居中、离顶部 20%; 记住的位置按屏幕里的千分比换算, 换一块屏幕也放在对应的地方
    ; 每一页的控件都在底部按钮上面 (中英文都检查, 文字长短不同)
    ; 偏好设置每一页: 上下相邻的控件 (左右有重叠) 至少隔 4 像素, 不会挤在一起。
    ; 灰色说明 (小一号的字) 本来就紧贴在上面的设置下面, 不算
    ; 偏好设置里的表格: 最后一列补满, 各列刚好铺满列表 (有竖滚动条时减掉它), 没有横向滚动条
    static ListColumnsFit() {
        eq := (n, a, e) => TestRunner.Equal("ListColumnsFit." n, a, e)
        PreferencesWindow.Show(1, -3000, -3000)
        PreferencesWindow.BuildAll()
        lists := []
        for page in PreferencesWindow.Pages
            for ctrl in page.Controls
                if (ctrl.Type = "ListView" && SendMessage(0x108F, 0, 0, ctrl.Hwnd) = 1)   ; LVM_GETVIEW = LV_VIEW_DETAILS (主题缩略图不算)
                    lists.Push([page.Key, ctrl])
        TestRunner.True("ListColumnsFit.found the tables (" lists.Length ")", lists.Length >= 6)
        for entry in lists {
            listView := entry[2]
            columns := SendMessage(0x1200, 0, 0, SendMessage(0x101F, 0, 0, listView.Hwnd))
            total := 0
            Loop columns
                total += SendMessage(0x101D, A_Index - 1, 0, listView.Hwnd)
            eq(entry[1] " fills the width", total, PreferencesWindow.ListWidth(listView))
            style := DllCall("GetWindowLong", "Ptr", listView.Hwnd, "Int", -16, "Int")
            eq(entry[1] " no horizontal scroll bar", style & 0x100000 ? 1 : 0, 0)   ; WS_HSCROLL
        }
        listView := lists[1][2]                                             ; 拖窄了前面的列: 最后一列补上
        SendMessage(0x101E, 0, 33, listView.Hwnd)
        PreferencesWindow.FitColumns(listView)
        columns := SendMessage(0x1200, 0, 0, SendMessage(0x101F, 0, 0, listView.Hwnd)), total := 0
        Loop columns
            total += SendMessage(0x101D, A_Index - 1, 0, listView.Hwnd)
        eq("user width kept", SendMessage(0x101D, 0, 0, listView.Hwnd) "|" total, "33|" PreferencesWindow.ListWidth(listView))
        PreferencesWindow.Close()
    }

    static PreferencesSpacing() {
        smallFont(ctrl) {
            if !(ctrl.Type = "Text" || ctrl.Type = "Link") || !(font := SendMessage(0x31, 0, 0, ctrl.Hwnd))   ; WM_GETFONT
                return false
            logFont := Buffer(92, 0)
            DllCall("GetObject", "Ptr", font, "Int", 92, "Ptr", logFont)
            return Abs(NumGet(logFont, 0, "Int")) <= 11 * A_ScreenDPI / 96      ; s8 约 11 像素, s9 约 12 像素
        }
        savedLang := I18n.Lang
        for lang in ["en", "zh-CN", "ja"] {
            I18n.Init(lang)
            PreferencesWindow.Show(1, -3000, -3000)
            PreferencesWindow.BuildAll()
            for index, page in PreferencesWindow.Pages {
                boxes := []
                for ctrl in page.Controls {
                    ctrl.GetPos(&x, &y, &w, &h)
                    if (h > 2)                                               ; 分节下面的细线不算
                        boxes.Push({Ctrl: ctrl, X: x, Y: y, W: w, H: h, Small: smallFont(ctrl)})
                }
                tight := ""
                for upper in boxes {
                    below := "", gap := 9999
                    for other in boxes {
                        if (other = upper || other.Y < upper.Y + upper.H - 2 || other.X >= upper.X + upper.W || upper.X >= other.X + other.W)
                            continue
                        if (other.Y - upper.Y - upper.H < gap)
                            below := other, gap := other.Y - upper.Y - upper.H
                    }
                    if (IsObject(below) && gap < 4 && !below.Small && !upper.Small)
                        tight .= upper.Ctrl.Type " '" SubStr(upper.Ctrl.Text, 1, 20) "' -> " below.Ctrl.Type " '" SubStr(below.Ctrl.Text, 1, 20) "' (" gap "); "
                }
                TestRunner.Equal("PreferencesSpacing." lang " page " index " " page.Key, tight, "")
            }
            PreferencesWindow._dirty := false
            PreferencesWindow.Close()
        }
        I18n.Init(savedLang)
    }

    static PreferencesFit() {
        savedLang := I18n.Lang
        for language in I18n.Languages() {                                   ; 按代码排序, ja 在 zh-TW 前 (Wine 没有这两种字体, 先用过 JhengHei 后量出的日文宽度偏小)
            lang := language[1]
            I18n.Init(lang)
            PreferencesWindow.Show(1, -3000, -3000)
            PreferencesWindow.BuildAll()                                         ; 页面用到时才建: 检查每一页之前先全部建好
            limit := PreferencesWindow.ButtonY - 6
            rightLimit := PreferencesWindow.ContentX + PreferencesWindow.ContentW + 3   ; 分节下面的细线 (SS_ETCHEDHORZ) 会宽 2 像素
            for index, page in PreferencesWindow.Pages {
                if (page.Key = "Prefs.Page.Keyboard" && !AppSettings.General["CapsLock"]) {  ; Caps Lock as usual: "also in administrator windows" is grayed out, with the reason
                    hint := 0
                    for ctrl in page.Controls
                        hint := hint || (ctrl.Type = "Text" && ctrl.Text = I18n.T("Prefs.CapsLockElevated.NeedsMode"))
                    TestRunner.True("PreferencesFit." lang " caps lock helper hint", hint)
                    TestRunner.Equal("PreferencesFit." lang " shield space", StrLen(PreferencesWindow.ShieldPad(4)) " " StrLen(PreferencesWindow.ShieldPad(3)), "5 7")
                }
                if (page.Key = "Prefs.Page.Window" || page.Key = "Prefs.Page.Keyboard") {   ; the input boxes / drop-downs of these pages share one left edge
                    lefts := "", count := 0
                    for ctrl in page.Controls
                        if (ctrl.Type = "Edit" || ctrl.Type = "DDL")
                            ctrl.GetPos(&x), lefts .= (lefts = "" ? "" : " ") x, count += 1
                    first := StrSplit(lefts, " ")[1], expected := ""
                    Loop count
                        expected .= (A_Index = 1 ? "" : " ") first
                    TestRunner.Equal("PreferencesFit." lang " " page.Key " controls aligned", lefts, expected)
                    TestRunner.Equal("PreferencesFit." lang " " page.Key " control count", count, page.Key = "Prefs.Page.Window" ? 2 : 4)
                }
                suffix := I18n.T("Prefs.DefaultChoice", "")
                for ctrl in page.Controls                                   ; every drop-down marks exactly one option "(default)" and shows its longest option in full
                    if (ctrl.Type = "DDL") {
                        marked := 0
                        for text in ControlGetItems(ctrl)
                            marked += (SubStr(text, -StrLen(suffix)) = suffix)
                        ctrl.GetPos(, , &w)
                        TestRunner.Equal("PreferencesFit." lang " " page.Key " default marked once (" ControlGetItems(ctrl)[1] ")", marked, 1)
                        TestRunner.True("PreferencesFit." lang " " page.Key " drop-down fits (" ControlGetItems(ctrl)[1] ", " w " < " PreferencesWindow._ChoiceWidth(ctrl) ")", w >= PreferencesWindow._ChoiceWidth(ctrl))
                    }
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
        commandIds := Map()
        for command in SystemProvider.Commands()
            commandIds[command["Id"]] := true
        scopes := ""
        for entry in defaults["Hotkeys"] {                                  ; F1 ~ F4 are built into the search window, not custom hotkeys
            TestRunner.True("DefaultExamples.no F-key hotkey " entry["Key"], !RegExMatch(entry["Key"], "i)^F[1-4]$"))
            TestRunner.True("DefaultExamples.hotkey action " entry["Action"], commandIds.Has(entry["Action"]))
            eq("hotkey key order " entry["Key"], HotkeyBox.Compose(RegExReplace(entry["Key"], "i)\w+$"), RegExReplace(entry["Key"], "^[\^!+#]+")), entry["Key"])   ; as Preferences would record it
            scopes .= (entry["WinTitle"] = "" ? "global" : entry["WinTitle"] = "ALTRun" ? "altrun" : "window") " "
        }
        eq("hotkey examples: global, one program, ALTRun", scopes, "global window altrun ")
        eq("no mouse hotkeys by default (no mouse hook)", App.UsesMouseHotkeys(defaults["Hotkeys"]), false)
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

    ; Wiki 中英两套页面 (英文 Usage / 中文 Usage-zh): 每页都有中文页, 页面里的 Wiki 链接都指向存在的页面,
    ; 程序里打开的帮助页 (速查表、偏好设置的 F1) 两种语言都有; 不再有旧的 en- 页面 (过渡页由 wiki.yml 发布时生成)
    static WikiPages() {
        dir := A_ScriptDir "\..\docs\wiki\"
        pages := Map()
        Loop Files, dir "*.md"
            pages[SubStr(A_LoopFileName, 1, -3)] := true
        missing := "", old := ""
        for page in pages {
            if (SubStr(page, 1, 3) = "en-")
                old .= page " "
            else if (SubStr(page, 1, 1) != "_" && !RegExMatch(page, "-zh$") && !pages.Has(page "-zh"))
                missing .= page " "
            else if (RegExMatch(page, "^(.+)-zh$", &zh) && !pages.Has(zh[1]))
                missing .= page " "
        }
        TestRunner.Equal("WikiPages.chinese page for every page", missing, "")
        TestRunner.Equal("WikiPages.no old en- pages", old, "")
        broken := ""
        for page in pages {
            text := FileRead(dir page ".md", "UTF-8"), pos := 1
            while (pos := RegExMatch(text, "\]\(([A-Za-z][\w-]*)(#[^)]*)?\)", &m, pos)) {   ; [文字](Page) 或 [文字](Page#标题)
                pos += m.Len
                if !pages.Has(m[1])
                    broken .= page " -> " m[1] "; "
            }
        }
        TestRunner.Equal("WikiPages.links point to existing pages", broken, "")
        used := Map()
        for entry in HelpProvider.Entries()
            used[entry[3]] := true
        for pageKey in PreferencesWindow.PageKeys
            used[PreferencesWindow.WikiPage(pageKey)] := true
        missing := ""
        for page in used
            if !pages.Has(page) || !pages.Has(page "-zh")
                missing .= page " "
        TestRunner.Equal("WikiPages.help pages exist in both languages", missing, "")
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

    ; 数据文件夹: 默认位置的 ALTRun.json 里的 "DataLocation" (环境变量、相对路径); 写 / 恢复默认
    static DataLocation() {
        eq := (n, a, e) => TestRunner.Equal("DataLocation." n, a, e)
        eq("full: absolute", Path.Full("C:\Sync\ALTRun\"), "C:\Sync\ALTRun")
        eq("full: env", Path.Full("%SystemRoot%\Temp"), A_WinDir "\Temp")
        eq("full: relative", Path.Full("..\Shared\Data", "C:\Apps\ALTRun"), "C:\Apps\Shared\Data")
        eq("full: builtin", Path.Full("A_AppData\ALTRun"), A_AppData "\ALTRun")
        root := A_Temp "\ALTRun-location-test"
        try DirDelete(root, true)
        DirCreate(root "\user\Data")
        saved := {Portable: AppSettings.Portable, UserDir: AppSettings.UserDir, DataDir: AppSettings.DataDir}
        AppSettings.Portable := false, AppSettings.UserDir := root "\user"   ; 默认位置 = root\user\Data
        settingsFile := root "\user\Data\ALTRun.json"
        eq("default when not portable", AppSettings.DefaultDataDir(), root "\user\Data")
        eq("no settings settingsFile", AppSettings.ResolveDataDir(), root "\user\Data")
        FileAppend('{"General": {"Hotkey": "!Space"}, "Snippets": [{"Text": "\"DataLocation\": \"X:\\\\fake\""}]}', settingsFile, "UTF-8")
        eq("no key (text in a snippet does not count)", AppSettings.ResolveDataDir(), root "\user\Data")
        FileDelete(settingsFile)
        FileAppend('{"General": {}, "DataLocation": "D:\\OneDrive\\ALTRun\\"}', settingsFile, "UTF-8")
        eq("key", AppSettings.ResolveDataDir(), "D:\OneDrive\ALTRun")
        FileDelete(settingsFile)
        FileAppend('{"DataLocation": "%SystemRoot%\\Temp"}', settingsFile, "UTF-8")
        eq("environment variable", AppSettings.ResolveDataDir(), A_WinDir "\Temp")
        FileDelete(settingsFile)
        FileAppend('{"DataLocation": ""}', settingsFile, "UTF-8")
        eq("empty = default", AppSettings.ResolveDataDir(), root "\user\Data")
        FileDelete(settingsFile)

        AppSettings.DataDir := "E:\Elsewhere"                                ; 正在用别处的设置: 只改默认位置的文件
        FileAppend('{"General": {"Hotkey": "!r"}}', settingsFile, "UTF-8")
        AppSettings.SetDataLocation("E:\Sync\ALTRun")
        data := JSON.Parse(FileRead(settingsFile, "UTF-8"))
        eq("set: key added, other settings kept", data["DataLocation"] "|" data["General"]["Hotkey"], "E:\Sync\ALTRun|!r")
        eq("set: resolves", AppSettings.ResolveDataDir(), "E:\Sync\ALTRun")
        AppSettings.SetDataLocation("")
        data := JSON.Parse(FileRead(settingsFile, "UTF-8"))
        eq("reset: key removed", data.Has("DataLocation") "|" data["General"]["Hotkey"], "0|!r")
        AppSettings.SetDataLocation("E:\Sync\ALTRun"), AppSettings.SetDataLocation(AppSettings.DefaultDataDir())
        eq("default folder = reset", JSON.Parse(FileRead(settingsFile, "UTF-8")).Has("DataLocation"), 0)
        FileDelete(settingsFile)
        AppSettings.SetDataLocation("")
        eq("reset without a settingsFile: nothing written", FileExist(settingsFile) ? "exists" : "missing", "missing")
        AppSettings.SetDataLocation("E:\Sync\ALTRun")
        eq("set without a settingsFile: creates it", AppSettings.ResolveDataDir(), "E:\Sync\ALTRun")

        AppSettings.DataDir := saved.DataDir, AppSettings.Portable := saved.Portable, AppSettings.UserDir := saved.UserDir
        eq("default when portable", AppSettings.DefaultDataDir(), A_ScriptDir "\Data")
        DirDelete(root, true)
    }

    ; Windows 设置的页面: 按名称 (和英文名称) 搜到, 可以整组关掉
    ; 搜索里的偏好设置子页面: 每页一条 ("ALTRun Preferences: Appearance"), 按页名找得到, 主项 "ALTRun Preferences" 排在子项前面
    static PreferencePages() {
        eq := (n, a, e) => TestRunner.Equal("PreferencePages." n, a, e)
        best(text) {
            top := ""
            for item in SystemProvider.Search(SearchQuery(text))
                if (!IsObject(top) || item.Score > top.Score)
                    top := item
            return IsObject(top) ? top.Source["Id"] : ""
        }
        count := 0
        for command in SystemProvider.Commands()
            count += (SubStr(command["Id"], 1, 12) = "Preferences.")
        eq("one per page", count, PreferencesWindow.PageKeys.Length)
        eq("appearance", best("appearance"), "Preferences.Appearance")
        eq("hotkeys page", best("preferences hotkeys"), "Preferences.Hotkeys")
        eq("main item first", best("altrun pref"), "Preferences")
        for command in SystemProvider.Commands()
            if (command["Id"] = "Preferences.Calculator")
                eq("title", command["Title"], "ALTRun Preferences: Calculator")
    }

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
        eq("default count", AppSettings.Defaults()["Features"]["Recent"]["RecentCount"], 0)   ; 默认只显示置顶的项目
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

    ; 剪贴板存在哪里 (LocalFiles): 1 = Clipboard 文件夹搬到本机, ClipboardHistory.json 留在 Data (照常同步);
    ; 2 = json 也搬到本机; 0 = 都搬回 Data; 两边都有时不覆盖; 没有本机文件夹时都在 Data
    static ClipboardLocal() {
        eq := (n, a, e) => TestRunner.Equal("ClipboardLocal." n, a, e)
        root := A_Temp "\ALTRun-clip-local-test"
        data := root "\Data", localDir := root "\Local"
        try DirDelete(root, true)
        DirCreate(data "\Clipboard")
        FileAppend('{"Entries": []}', data "\ClipboardHistory.json", "UTF-8")
        FileAppend("long text", data "\Clipboard\t-1.txt", "UTF-8")
        saved := [ClipboardProvider.File, ClipboardProvider.Folder]
        where := () => (FileExist(data "\ClipboardHistory.json") ? "D" : "-") (FileExist(localDir "\ClipboardHistory.json") ? "L" : "-")
            . " " (FileExist(data "\Clipboard\t-1.txt") ? "D" : "-") (FileExist(localDir "\Clipboard\t-1.txt") ? "L" : "-")
        ClipboardProvider.PlaceStorage(1, data, localDir)
        eq("1 json stays in Data", ClipboardProvider.File, data "\ClipboardHistory.json")
        eq("1 folder", ClipboardProvider.Folder, localDir "\Clipboard")
        eq("1 files", where(), "D- -L")
        ClipboardProvider.PlaceStorage(1, data, localDir)                      ; 已经在本机: 什么都不做
        eq("1 again", where(), "D- -L")
        ClipboardProvider.PlaceStorage(2, data, localDir)
        eq("2 json", ClipboardProvider.File, localDir "\ClipboardHistory.json")
        eq("2 files", where(), "-L -L")
        ClipboardProvider.PlaceStorage(0, data, localDir)
        eq("0 json", ClipboardProvider.File, data "\ClipboardHistory.json")
        eq("0 folder", ClipboardProvider.Folder, data "\Clipboard")
        eq("0 files", where(), "D- D-")
        FileAppend('{"Entries": []}', localDir "\ClipboardHistory.json", "UTF-8")   ; 两边都有: 不覆盖, 用新位置的
        ClipboardProvider.PlaceStorage(0, data, localDir)
        eq("both kept", where(), "DL D-")
        ClipboardProvider.PlaceStorage(2, data, "")                            ; 没有 %LOCALAPPDATA%: 都在 Data
        eq("no local", ClipboardProvider.File "|" ClipboardProvider.Folder, data "\ClipboardHistory.json|" data "\Clipboard")
        eq("default", AppSettings.Defaults()["Features"]["Clipboard"]["LocalFiles"], 1)
        ClipboardProvider.File := saved[1], ClipboardProvider.Folder := saved[2]
        try DirDelete(root, true)
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
        eq("pinned badge", items[1].Pinned, true)
        eq("unpinned: no badge", ClipboardProvider.Search(SearchQuery("clip five"))[1].Pinned, false)
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
        root := TestRunner.LongPath(root)                                   ; %Temp% 可能是 8.3 短路径 (C:\Users\RUNNER~1), 列出的文件是长路径
        FileAppend("x", root "\readme.txt"), FileAppend("x", root "\report 2026.docx"), FileAppend("x", root "\secret.txt")
        FileSetAttrib("+H", root "\secret.txt")
        browse := FileSearchProvider.BrowsePath(root "\rep")
        eq("parse", IsObject(browse) ? browse.Dir "|" browse.Filter : "", root "\|rep")
        eq("not a path", FileSearchProvider.BrowsePath("report"), "")
        eq("drive only", FileSearchProvider.BrowsePath("C:"), "")
        missing := FileSearchProvider.BrowsePath(root "\Nope\x")
        eq("missing folder", IsObject(missing) && missing.HasOwnProp("Missing") ? missing.Dir : "", root "\Nope\")
        eq("missing folder: hint", titles(root "\Nope\x"), I18n.T("Files.FolderNotFound"))
        wide := FileSearchProvider.BrowsePath(SubStr(root, 1, 1) "：、" StrReplace(SubStr(root, 4), "\", "、") "、rep")   ; 中文输入法的全角冒号和顿号
        eq("chinese punctuation", IsObject(wide) ? wide.Dir "|" wide.Filter : "", root "\|rep")
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
        root := TestRunner.LongPath(root)
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
        ; 第一次打开 Scripts 文件夹: 建三个示例 (各一种写法), 都读得出来; 文件夹已经有了就不动
        examples := A_Temp "\ALTRunScriptExamples"
        try DirDelete(examples, true)
        savedOpen := ActionCatalog.GetOwnPropDesc("OpenFolder"), opened := ""
        ActionCatalog.DefineProp("OpenFolder", {Call: (this, folder) => opened := folder})
        ScriptProvider.Folder := examples
        try {
            ScriptProvider.OpenFolder()
            eq("examples: folder opened", opened, examples)
            ScriptProvider.Load()
            found := Map()
            for script in ScriptProvider.Scripts
                found[script.Name] := script.Title "|" script.Mode "|" script.Keyword "|" script.Argument
            eq("examples: count", found.Count, 3)
            eq("example ahk", found.Get("Today.ahk", ""), "Example: Today's Date|silent||")
            eq("example bat", found.Get("Port.bat", ""), "Example: Who Uses a Port|window|port|a port number")
            eq("example ps1", found.Get("IP Addresses.ps1", ""), "Example: IP Addresses|output||")
            bom := FileOpen(examples "\Port.bat", "r", "CP0"), first := bom.RawRead(head := Buffer(3), 3) ? NumGet(head, 0, "UChar") : 0, bom.Close()
            eq("example bat has no BOM", first, Ord("@"))
            for name, text in ScriptProvider.ExampleScripts()
                TestRunner.True("Scripts.example ASCII only " name, !RegExMatch(text, "[^\x00-\x7F]"))
            ; 等待参数时副标题说明要接着输入什么 (脚本、自定义命令、网页搜索)
            port := ScriptProvider.Search(SearchQuery("port"))
            eq("waiting hint: script", port.Length ? port[1].Subtitle : "", I18n.T("Search.TypeArgAfter", "port", "a port number") " · " I18n.T("Script.Subtitle", "Port.bat"))
            eq("waiting: script autocompletes", port.Length ? port[1].Valid "|" port[1].AutoComplete : "", "0|port ")
            ping := Map("Title", "Ping", "Type", "Command", "Target", "cmd.exe", "Arguments", "/k ping {query}", "Keyword", "ping")
            TestRunner.True("Scripts.waiting hint: custom command", InStr(CustomCommandProvider._ToItem(ping, 100).Subtitle, I18n.T("Search.TypeAfter", "ping") " · ") = 1)
            TestRunner.True("Scripts.with argument: no hint", !InStr(CustomCommandProvider._ToItem(ping, 100, "8.8.8.8").Subtitle, I18n.T("Search.TypeAfter", "ping")))
            eq("waiting hint: web search", WebSearchProvider.ItemFor(Map("Title", "Google", "Keyword", "g", "Url", "https://example.com/?q={query}"), "").Subtitle, I18n.T("Search.TypeAfter", "g"))
            FileDelete(examples "\Today.ahk")
            ScriptProvider.OpenFolder()
            eq("examples: existing folder untouched", FileExist(examples "\Today.ahk"), "")
        } finally {
            ActionCatalog.DefineProp("OpenFolder", savedOpen)
            ScriptProvider.Folder := saved[1], ScriptProvider.Scripts := saved[2]
            try DirDelete(examples, true)
        }
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
        ; 带 Wiki 链接的说明: 每种语言都写了 <a href="{1}">, 网址按界面语言打开中文 / 英文页; 其它说明不能有链接 (会原样显示)
        for language in I18n.Languages() {
            strings := (language[1] = "en") ? I18n.Strings : I18n.ReadFile(language[1])
            for key, text in strings
                if (SubStr(key, -5) = ".Desc")
                    TestRunner.True("PreferenceDescriptions.link " language[1] " " key
                        , PreferencesWindow.DescLinks.Has(SubStr(key, 1, -5)) ? InStr(text, '<a href="{1}">') : !InStr(text, "<a "))
        }
        TestRunner.True("PreferenceDescriptions.link url", InStr(PreferencesWindow._DescLink("Prefs.SelectionHotkey"), "/wiki/") && PreferencesWindow._DescLink("Prefs.Hotkey") = "")
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
        ; 启动时快捷方式没变就不重写
        description := "ALTRun - An effective launcher for Windows", iconFile := Path.Full("..\Resources\ALTRun.ico")
        FileCreateShortcut(A_AhkPath, lnkDir "\same.lnk", A_ScriptDir, "-SendTo", description, iconFile)
        eq("matches", App.ShortcutMatches(lnkDir "\same.lnk", A_AhkPath, A_ScriptDir, "-SendTo", description, iconFile), true)
        eq("other arguments", App.ShortcutMatches(lnkDir "\same.lnk", A_AhkPath, A_ScriptDir, "-Startup", description, iconFile), false)
        eq("other description", App.ShortcutMatches(lnkDir "\same.lnk", A_AhkPath, A_ScriptDir, "-SendTo", "ALTRun - 高效的 Windows 启动器", iconFile), false)
        eq("other target", App.ShortcutMatches(lnkDir "\same.lnk", A_ScriptFullPath, A_ScriptDir, "-SendTo", description, iconFile), false)
        eq("other icon", App.ShortcutMatches(lnkDir "\same.lnk", A_AhkPath, A_ScriptDir, "-SendTo", description, ""), false)
        FileCreateShortcut(A_AhkPath, lnkDir "\noicon.lnk", A_ScriptDir, "", description)
        eq("matches without icon", App.ShortcutMatches(lnkDir "\noicon.lnk", A_AhkPath, A_ScriptDir, "", description), true)
        eq("missing shortcut", App.ShortcutMatches(lnkDir "\none.lnk", A_AhkPath, A_ScriptDir, "", description), false)
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

        ; 偏好设置开着时: 加进 "命令" 页的列表 (算作修改, 按 确定 / 应用 才保存), 不直接写进设置
        AppSettings.Data["CustomCommands"] := [Map("Title", "Old", "Type", "Folder", "Target", folder "\pt2415 - riverside\", "Arguments", "", "Keyword", "")]
        PreferencesWindow.Show(1, -3000, -3000)                             ; 刚打开 (还没开始记录修改) 时加进来的也算修改
        CustomCommandProvider.AddFromPaths([folder "\PT2415 - Riverside", folder "\Design Report.docx", folder "\design report.docx", "C:\Windows"])
        working := PreferencesWindow.Working["CustomCommands"]
        eq("prefs: added to the working list", working.Length "|" working[2]["Title"] "|" working[3]["Title"], "3|Design Report|Windows")
        eq("prefs: settings not saved yet", AppSettings.CustomCommands.Length, 1)
        eq("prefs: commands page shown, dirty", PreferencesWindow.Pages[PreferencesWindow._page].Key "|" PreferencesWindow._dirty, "Prefs.Page.Commands|1")
        listView := ""
        for ctrl in PreferencesWindow.Pages[PreferencesWindow._page].Controls
            if (ctrl.Type = "ListView")
                listView := ctrl
        eq("prefs: list shows them, last selected", listView.GetCount() "|" listView.GetNext(), "3|3")
        PreferencesWindow.Close()
        eq("prefs: closed clears list hooks", PreferencesWindow._lists.Count, 0)
        ToolTip(, , , 20)
        AppSettings.Data["CustomCommands"] := savedCommands, AppSettings.File := savedFile
        try FileDelete(A_Temp "\ALTRunTest.json")
        try DirDelete(folder, true)
    }

    ; 偏好设置 → 高级: 链接显示名称 (主页 · GitHub · 更新说明 · 报告问题), 版权, 数据文件夹一行 (长路径用省略号, 鼠标停留显示完整路径)
    static AdvancedPage() {
        eq := (n, a, e) => TestRunner.Equal("AdvancedPage." n, a, e)
        PreferencesWindow.Show("Prefs.Page.Advanced", -3000, -3000)
        try {
            texts := "", link := "", folder := ""
            for ctrl in PreferencesWindow.Pages[PreferencesWindow._page].Controls {
                if (ctrl.Type = "Link")
                    link := ctrl
                else if (ctrl.Type = "Text" && ctrl.Text = AppSettings.DataDir)
                    folder := ctrl
                if (ctrl.Type = "Text")
                    texts .= ctrl.Text "`n"
            }
            eq("four links", StrSplit(link.Text, "<a ").Length - 1, 4)
            eq("no bare website address", InStr(link.Text, ">zhugecaomao.github.io"), 0)
            eq("issues link", InStr(link.Text, App.RepoUrl "/issues/new/choose") > 0, true)
            eq("copyright", InStr(texts, "© 2013–" SubStr(App.Version, 1, 4) " zhugecaomao") > 0, true)
            eq("settings file row removed", InStr(texts, AppSettings.File), 0)
            eq("data folder: one line, path ellipsis", IsObject(folder) && (ControlGetStyle(folder) & 0x8100) = 0x8100, true)
        } finally {
            PreferencesWindow.Close()
        }
        g := Gui()
        label := g.AddText("w100 0x100", "C:\A very long path\to a folder")
        tip := Win.AddTooltip(label.Hwnd, "C:\A very long path\to a folder")
        eq("tooltip added", DllCall("SendMessageW", "Ptr", tip, "UInt", 0x40D, "Ptr", 0, "Ptr", 0), 1)   ; TTM_GETTOOLCOUNT
        g.Destroy()
    }

    ; 只运行一个 ALTRun: 再启动时把命令行参数交给正在运行的 (WM_COPYDATA), 重新启动 (/restart) 时替换它
    static SingleInstance() {
        eq := (n, a, e) => TestRunner.Equal("SingleInstance." n, a, e)
        eq("restart (compiled)", App.IsRestart('"C:\ALTRun\ALTRun.exe" /restart -Reloaded'), true)
        eq("restart (script)", App.IsRestart('"C:\AHK\AutoHotkey64.exe" /restart "C:\ALTRun\ALTRun.ahk" -Preferences 3'), true)
        eq("restart (no quotes)", App.IsRestart('ALTRun.exe /RESTART'), true)
        eq("double-click", App.IsRestart('"C:\ALTRun\ALTRun.exe"'), false)
        eq("send to", App.IsRestart('"C:\ALTRun\ALTRun.exe" -SendTo "D:\a /restart b"'), false)
        eq("no other instance", App._RunningInstance(), 0)
        eq("title: compiled exe", App.IsInstanceTitle("C:\ALTRun\ALTRun.exe", "C:\ALTRun\ALTRun.exe"), true)
        eq("title: script", App.IsInstanceTitle("C:\ALTRun\ALTRun.ahk - AutoHotkey v2.0.19", "C:\ALTRun\ALTRun.ahk"), true)
        eq("title: other case", App.IsInstanceTitle("c:\altrun\ALTRun.exe", "C:\ALTRun\ALTRun.exe"), true)
        eq("title: other script", App.IsInstanceTitle("C:\ALTRun\ALTRun.ahk.bak - AutoHotkey v2.0.19", "C:\ALTRun\ALTRun.ahk"), false)
        eq("title: other folder", App.IsInstanceTitle("D:\Test\ALTRun.exe", "C:\ALTRun\ALTRun.exe"), false)

        received := []
        saved := App.GetOwnPropDesc("_HandleCommandLine")
        App.DefineProp("_HandleCommandLine", {Call: (self, args, forwarded := false) => received.Push([args, forwarded])})
        OnMessage(0x4A, onCopyData := (p*) => App._OnCopyData(p*))
        try {
            eq("sent", App.SendArgs(A_ScriptHwnd, ["-SendTo", "D:\项目 A\Report.docx", "C:\Windows"]), true)
            eq("no arguments sent", App.SendArgs(A_ScriptHwnd, []), true)
            Sleep(50)
            eq("handled after replying", received.Length, 2)
            if (received.Length = 2) {
                args := received[1][1]
                eq("arguments", args.Length "|" args[1] "|" args[2] "|" args[3], "3|-SendTo|D:\项目 A\Report.docx|C:\Windows")
                eq("forwarded", received[1][2], true)
                eq("double-click: no arguments", received[2][1].Length, 0)
            }
            copyData := Buffer(A_PtrSize * 3, 0)                            ; 别的程序 (例如 Everything) 的 WM_COPYDATA 不处理
            NumPut("UPtr", 0x1234, copyData, 0)
            eq("other copydata ignored", App._OnCopyData(0, copyData.Ptr, 0x4A, A_ScriptHwnd), "")
        } finally {
            OnMessage(0x4A, onCopyData, 0)
            App.DefineProp("_HandleCommandLine", saved)
        }
    }

    ; Ctrl+↑ / Ctrl+↓ 翻搜索记录 (↑ ↓ 只移动选择); 用隐藏的输入框代替搜索窗口, 不真的搜索
    static HistoryKeys() {
        eq := (n, a, e) => TestRunner.Equal("HistoryKeys." n, a, e)
        savedInput := SearchWindow.Input, savedHistory := Knowledge.History, savedSearch := SearchWindow.GetOwnPropDesc("_RunSearch")
        SearchWindow.DefineProp("_RunSearch", {Call: (*) => 0})
        g := Gui()
        SearchWindow.Input := g.AddEdit("w200")
        SearchWindow.Mode := "results", SearchWindow.FileMode := false, SearchWindow.HistoryIndex := 0
        Knowledge.History := ["report", "notes", "pt tools"]
        try {
            SearchWindow.Input.Value := "typed"
            SearchWindow.RecallHistory(1)
            eq("ctrl+up newest", SearchWindow.Input.Value, "report")
            SearchWindow.RecallHistory(1)
            eq("ctrl+up older", SearchWindow.Input.Value, "notes")
            SearchWindow.RecallHistory(1), SearchWindow.RecallHistory(1)
            eq("stops at oldest", SearchWindow.Input.Value "|" SearchWindow.HistoryIndex, "pt tools|3")
            SearchWindow.RecallHistory(-1)
            eq("ctrl+down newer", SearchWindow.Input.Value, "notes")
            SearchWindow.RecallHistory(-1), SearchWindow.RecallHistory(-1)
            eq("back to typed text", SearchWindow.Input.Value "|" SearchWindow.HistoryIndex, "typed|0")
            SearchWindow.RecallHistory(-1)
            eq("ctrl+down at the start does nothing", SearchWindow.Input.Value, "typed")
            SearchWindow.Input.Value := "report", SearchWindow.HistoryIndex := 0     ; e.g. the restored last search
            SearchWindow.RecallHistory(1)
            eq("same text skipped", SearchWindow.Input.Value, "notes")
            SearchWindow.HistoryIndex := 0, SearchWindow.FileMode := true
            SearchWindow.RecallHistory(1)
            eq("not in file mode", SearchWindow.Input.Value, "notes")
        } finally {
            SearchWindow.FileMode := false, SearchWindow.HistoryIndex := 0
            SearchWindow.DefineProp("_RunSearch", savedSearch)
            SearchWindow.Input := savedInput, Knowledge.History := savedHistory
            g.Destroy()
        }
    }

    ; 光标在最前面按空格: 已经输入的文字改用文件搜索; 文件搜索模式下光标在最前面按 Backspace: 回到普通搜索, 文字都保留
    static FileModeKeys() {
        eq := (n, a, e) => TestRunner.Equal("FileModeKeys." n, a, e)
        savedInput := SearchWindow.Input, savedGui := SearchWindow.Gui, savedSearch := SearchWindow.GetOwnPropDesc("_RunSearch")
        options := AppSettings.Feature("FileSearch"), savedOptions := [options["SpacePrefix"], options["Enabled"]]
        SearchWindow.DefineProp("_RunSearch", {Call: (*) => 0})
        g := Gui()
        SearchWindow.Gui := g, SearchWindow.Input := g.AddEdit("w200")
        hwnd := SearchWindow.Input.Hwnd
        caret := (start, end := "") => SendMessage(0xB1, start, end = "" ? start : end, hwnd)   ; EM_SETSEL
        key := (vk) => SearchWindow._OnKeyDown(vk, 0, 0x100, hwnd)
        SearchWindow.Mode := "results", SearchWindow.FileMode := false
        options["SpacePrefix"] := 1, options["Enabled"] := 1
        try {
            SearchWindow.Input.Value := "seismic", caret(7)
            eq("space at the end types a space", key(0x20) "|" SearchWindow.FileMode, "|0")
            caret(0, 3)
            eq("space with a selection types a space", key(0x20) "|" SearchWindow.FileMode, "|0")
            caret(0)
            eq("space at the start enters file mode", key(0x20) "|" SearchWindow.FileMode "|" SearchWindow.Input.Value, "0|1|seismic")
            caret(3)
            eq("backspace inside the text deletes", key(0x08) "|" SearchWindow.FileMode, "|1")
            caret(0)
            eq("backspace at the start leaves file mode", key(0x08) "|" SearchWindow.FileMode "|" SearchWindow.Input.Value, "0|0|seismic")
            caret(0)
            eq("backspace at the start in normal search", key(0x08) "|" SearchWindow.FileMode, "|0")
            options["SpacePrefix"] := 0
            caret(0)
            eq("space prefix off", key(0x20) "|" SearchWindow.FileMode, "|0")
            options["SpacePrefix"] := 1, SearchWindow.Input.Value := ""
            eq("empty box still enters file mode", key(0x20) "|" SearchWindow.FileMode, "0|1")
            eq("backspace in the empty box leaves it", key(0x08) "|" SearchWindow.FileMode, "0|0")
        } finally {
            SearchWindow.FileMode := false
            options["SpacePrefix"] := savedOptions[1], options["Enabled"] := savedOptions[2]
            SearchWindow.DefineProp("_RunSearch", savedSearch)
            SearchWindow.Input := savedInput, SearchWindow.Gui := savedGui
            g.Destroy()
        }
    }

    ; Tab: completes by default; with TabKey = "Next" it moves to the next result (wrapping), Shift+Tab to the previous
    static TabKey() {
        eq := (n, a, e) => TestRunner.Equal("TabKey." n, a, e)
        eq("default complete", AppSettings.Defaults()["General"]["TabKey"], "Complete")
        saved := {Input: SearchWindow.Input, Gui: SearchWindow.Gui, Results: SearchWindow.Results, Selected: SearchWindow.Selected
            , Offset: SearchWindow.Offset, Mode: SearchWindow.Mode, TabKey: AppSettings.General["TabKey"]}
        savedSearch := SearchWindow.GetOwnPropDesc("_RunSearch"), savedRepaint := SearchWindow.GetOwnPropDesc("_Repaint")
        SearchWindow.DefineProp("_RunSearch", {Call: (*) => 0}), SearchWindow.DefineProp("_Repaint", {Call: (*) => 0})
        g := Gui()
        SearchWindow.Gui := g, SearchWindow.Input := g.AddEdit("w200")
        hwnd := SearchWindow.Input.Hwnd
        tab := () => SearchWindow._OnKeyDown(0x09, 0, 0x100, hwnd)
        try {
            SearchWindow.Mode := "results", SearchWindow.Offset := 0
            SearchWindow.Results := [ResultItem("Notepad"), ResultItem("Notes"), ResultItem("Notion")], SearchWindow.Selected := 1
            SearchWindow.Input.Value := "no"
            AppSettings.General["TabKey"] := "Complete"
            eq("complete: handled", tab(), 0)
            eq("complete: fills in the name", SearchWindow.Input.Value "|" SearchWindow.Selected, "Notepad|1")
            AppSettings.General["TabKey"] := "Next", SearchWindow.Input.Value := "no"
            tab()
            eq("next: second result", SearchWindow.Selected "|" SearchWindow.Input.Value, "2|no")
            tab()
            eq("next: third result", SearchWindow.Selected, 3)
            tab()
            eq("next: wraps to the first", SearchWindow.Selected, 1)
            SearchWindow.MoveSelection(-1, true)                            ; what Shift+Tab does
            eq("previous: wraps to the last", SearchWindow.Selected, 3)
            SearchWindow.MoveSelection(1)
            eq("down arrow still stops at the last", SearchWindow.Selected, 3)
            SearchWindow.Results := []
            eq("next: no results", tab() "|" SearchWindow.Selected, "0|3")
            ; The cheat sheet / tip for Tab follows the setting
            savedLang := I18n.Lang
            I18n.Init("en")
            keyOf(id) {
                for item in HelpProvider.Items()
                    if (item.Id = id)
                        return item.Key "|" item.Text
                return ""
            }
            eq("tip in next mode", keyOf("Tab"), "Tab / Shift+Tab|next / previous result; Ctrl+Tab autocompletes")
            AppSettings.General["TabKey"] := "Complete"
            eq("tip in complete mode", keyOf("Tab"), "Tab|" I18n.T("Help.Tab.Text"))
            I18n.Init(savedLang)
        } finally {
            AppSettings.General["TabKey"] := saved.TabKey
            SearchWindow.DefineProp("_RunSearch", savedSearch), SearchWindow.DefineProp("_Repaint", savedRepaint)
            SearchWindow.Input := saved.Input, SearchWindow.Gui := saved.Gui, SearchWindow.Results := saved.Results
            SearchWindow.Selected := saved.Selected, SearchWindow.Offset := saved.Offset, SearchWindow.Mode := saved.Mode
            g.Destroy()
        }
    }

    ; While ALTRun shows a system dialog, keys reach the search box through its window procedure (subclass),
    ; not OnMessage. A sent WM_KEYDOWN is not monitored either, so it stands in for the modal loop here
    static SystemDialogKeys() {
        eq := (n, a, e) => TestRunner.Equal("SystemDialogKeys." n, a, e)
        saved := {Input: SearchWindow.Input, List: SearchWindow.List, Gui: SearchWindow.Gui, Results: SearchWindow.Results
            , Selected: SearchWindow.Selected, Offset: SearchWindow.Offset, Mode: SearchWindow.Mode, TabKey: AppSettings.General["TabKey"]}
        savedSearch := SearchWindow.GetOwnPropDesc("_RunSearch"), savedRepaint := SearchWindow.GetOwnPropDesc("_Repaint")
        SearchWindow.DefineProp("_RunSearch", {Call: (*) => 0}), SearchWindow.DefineProp("_Repaint", {Call: (*) => 0})
        g := Gui()
        SearchWindow.Gui := g, SearchWindow.Input := g.AddEdit("w200"), SearchWindow.List := g.AddListView("w200 r3", ["Title"])
        hwnd := SearchWindow.Input.Hwnd
        tab := () => SendMessage(0x100, 0x09, 0, hwnd)                       ; WM_KEYDOWN Tab
        try {
            eq("no system dialog open", SearchWindow.InSystemDialog(), 0)
            found := 0
            SetTimer(() => (found := SearchWindow.InSystemDialog(), found && WinClose("ahk_id " found)), -300)
            MsgBox("SystemDialogKeys", , "T3")                              ; closed by the timer above (or after 3 s)
            eq("finds an open MsgBox", found != 0, true)
            SearchWindow.Mode := "results", SearchWindow.Offset := 0, SearchWindow.Input.Value := "no"
            SearchWindow.Results := [ResultItem("Notepad"), ResultItem("Notes"), ResultItem("Notion")], SearchWindow.Selected := 1
            AppSettings.General["TabKey"] := "Next"
            tab()
            eq("not hooked: the key is lost", SearchWindow.Selected, 1)
            SearchWindow._HookKeys()
            eq("hooked: handled", tab(), 0)
            eq("hooked: next result", SearchWindow.Selected, 2)
            eq("its WM_CHAR is swallowed", SendMessage(0x102, 0x09, 0, hwnd) "|" SearchWindow._swallowChar, "0|0")
            SendMessage(0x100, 0x26, 0, SearchWindow.List.Hwnd)              ; ↑ with the list focused
            eq("list: previous result", SearchWindow.Selected, 1)
            SearchWindow._keySeen := true                                   ; OnMessage already handled this one
            tab()
            eq("seen by OnMessage: not handled twice", SearchWindow.Selected "|" SearchWindow._keySeen, "1|0")
            SendMessage(0x100, 0x41, 0, hwnd)                               ; A: not a search window key
            eq("other keys go on to the box", SearchWindow._swallowChar, 0)
            SearchWindow._UnhookKeys()
            tab()
            eq("unhooked", SearchWindow.Selected "|" SearchWindow._keyHooked, "1|0")
        } finally {
            SearchWindow._UnhookKeys()
            AppSettings.General["TabKey"] := saved.TabKey
            SearchWindow.DefineProp("_RunSearch", savedSearch), SearchWindow.DefineProp("_Repaint", savedRepaint)
            SearchWindow.Input := saved.Input, SearchWindow.List := saved.List, SearchWindow.Gui := saved.Gui, SearchWindow.Results := saved.Results
            SearchWindow.Selected := saved.Selected, SearchWindow.Offset := saved.Offset, SearchWindow.Mode := saved.Mode
            g.Destroy()
        }
    }

    ; 底部状态栏: 左边的提示 / 状态, 右边选中项的 Enter 操作; 窗口高度算上状态栏; → 不认自动重复
    static StatusBar() {
        eq := (n, a, e) => TestRunner.Equal("StatusBar." n, a, e)
        eq("on by default", AppSettings.Defaults()["Appearance"]["StatusBar"], 1)
        fileItem := ResultItem("Notepad", "", {Kind: "file", Arg: "C:\Windows\notepad.exe"})
        eq("file: open", SearchWindow.EnterTitle(fileItem), "Open")
        eq("folder: open", SearchWindow.EnterTitle(ResultItem("Docs", "", {Kind: "folder"})), "Open")
        eq("url: open", SearchWindow.EnterTitle(ResultItem("Web", "", {Kind: "url"})), "Open")
        eq("text: copy", SearchWindow.EnterTitle(ResultItem("42", "", {Kind: "text"})), "Copy")
        eq("text with its own action: paste", SearchWindow.EnterTitle(ResultItem("snippet", "", {Kind: "text", OnRun: (*) => 0})), "Paste")
        eq("command: run", SearchWindow.EnterTitle(ResultItem("Lock", "", {OnRun: (*) => 0})), "Run")
        eq("own name", SearchWindow.EnterTitle(ResultItem("img", "", {Kind: "text", OnRun: (*) => 0, RunTitle: "Paste Image"})), "Paste Image")
        eq("hint row: complete", SearchWindow.EnterTitle(ResultItem("g", "", {Valid: false, AutoComplete: "g "})), "Complete")
        eq("hint row without completion", SearchWindow.EnterTitle(ResultItem("x", "", {Valid: false})), "")

        saved := {Input: SearchWindow.Input, Gui: SearchWindow.Gui, Results: SearchWindow.Results, Selected: SearchWindow.Selected
            , Mode: SearchWindow.Mode, FileMode: SearchWindow.FileMode, Tip: SearchWindow._tip, Status: SearchWindow.StatusHeight}
        savedOpen := SearchWindow.GetOwnPropDesc("_OpenActions")
        g := Gui()
        SearchWindow.Gui := g, SearchWindow.Input := g.AddEdit("w200")
        try {
            SearchWindow.Mode := "results", SearchWindow.FileMode := false, SearchWindow._tip := "Tip: ? cheat sheet"
            SearchWindow.Input.Value := "", SearchWindow.Results := [], SearchWindow.Selected := 0
            texts := SearchWindow.StatusTexts()
            eq("empty box: tip", texts.Left "|" texts.Enter "|" texts.Actions, "Tip: ? cheat sheet||0")
            SearchWindow.Input.Value := "note", SearchWindow.Results := [fileItem, ResultItem("b")], SearchWindow.Selected := 1
            texts := SearchWindow.StatusTexts()
            eq("results: tip, enter and actions", texts.Left "|" texts.Enter "|" texts.Actions, "Tip: ? cheat sheet|Open|1")
            parts := SearchWindow.StatusParts(texts)
            eq("parts like Raycast", parts.Length = 2 ? parts[1].Label " " parts[1].Keys[1] " | " parts[2].Label " " parts[2].Keys[1] : "", "Open ↵ | Actions →")
            items := SearchWindow.MainMenuItems(), labels := ""
            for entry in items
                labels .= (IsObject(entry) ? StrSplit(entry[1], "`t")[1] : "-") ";"
            eq("menu", labels, "Preferences...;Shortcuts and Syntax;-;Rebuild Index;Check for Updates;-;Reload;Quit;")
            SearchWindow.FileMode := true
            eq("fileItem mode", SearchWindow.StatusTexts().Left, I18n.T("Status.FileMode"))
            SearchWindow.FileMode := false, SearchWindow.Mode := "actions"
            texts := SearchWindow.StatusTexts()
            eq("action panel", texts.Left "|" texts.Enter "|" texts.Actions, "Esc  Back|Run|0")
            SearchWindow.Mode := "results"
            eq("mix colors", Format("{:06X}", SearchWindow.MixColor("FFFFFF", "000000", 0.5)), "808080")

            SearchWindow.StatusHeight := 0
            plain := SearchWindow._WindowHeight(3)
            SearchWindow.StatusHeight := 26
            eq("window height includes the status bar", SearchWindow._WindowHeight(3) - plain, 26)
            eq("no status bar without result rows", SearchWindow._WindowHeight(0) - SearchWindow._ContentHeight(0), 0)

            opened := {Count: 0}
            SearchWindow.DefineProp("_OpenActions", {Call: (*) => (opened.Count++, true)})
            hwnd := SearchWindow.Input.Hwnd
            SearchWindow.Input.Value := "note", SendMessage(0xB1, 4, 4, hwnd)   ; 光标在最后
            SearchWindow._OnKeyDown(0x27, 0x40000000 | 1, 0x100, hwnd)       ; 按住 → 的自动重复
            eq("held right arrow does not open actions", opened.Count, 0)
            SearchWindow._OnKeyDown(0x27, 1, 0x100, hwnd)                    ; 新按下的 →
            eq("fresh right arrow opens actions", opened.Count, 1)
            SearchWindow.Input.Value := "1250+500", SendMessage(0xB1, 0, -1, hwnd)   ; 恢复上次的搜索: 全选
            eq("right arrow on selected text: no actions", (SearchWindow._OnKeyDown(0x27, 1, 0x100, hwnd), opened.Count), 1)
            SendMessage(0xB1, 8, 8, hwnd)                                     ; Edit 自己处理 →: 取消选择, 光标到末尾
            eq("right arrow again opens actions", (SearchWindow._OnKeyDown(0x27, 1, 0x100, hwnd), opened.Count), 2)
        } finally {
            SearchWindow.DefineProp("_OpenActions", savedOpen)
            SearchWindow.Input := saved.Input, SearchWindow.Gui := saved.Gui, SearchWindow.Results := saved.Results, SearchWindow.Selected := saved.Selected
            SearchWindow.Mode := saved.Mode, SearchWindow.FileMode := saved.FileMode, SearchWindow._tip := saved.Tip, SearchWindow.StatusHeight := saved.Status
            g.Destroy()
        }
    }

    ; 呼出窗口期间打的字: 显示完补进搜索框, Enter 搜完后执行, Esc 关掉
    ; 隐藏后准备好空的搜索框: 呼出时不用先搜索; 输入框改过或用过一次就不算; 显示后的刷新只在结果变了时换上
    static ShowPrepared() {
        eq := (n, a, e) => TestRunner.Equal("ShowPrepared." n, a, e)
        savedInput := SearchWindow.Input, savedGui := SearchWindow.Gui, savedSearch := SearchWindow.GetOwnPropDesc("_RunSearch")
        saved := [SearchWindow.Mode, SearchWindow.FileMode, SearchWindow._prepared]
        searches := {Count: 0}
        SearchWindow.DefineProp("_RunSearch", {Call: (*) => searches.Count++})
        g := Gui()
        SearchWindow.Gui := g, SearchWindow.Input := g.AddEdit("w200")
        try {
            SearchWindow.Input.Value := "abc", SearchWindow.Mode := "actions", SearchWindow.FileMode := true
            SearchWindow.PrepareEmpty()                                     ; 窗口隐藏着: 清空输入框, 搜一次
            eq("prepared", SearchWindow.IsPrepared() "|" SearchWindow.Input.Value "|" SearchWindow.Mode "|" SearchWindow.FileMode "|" searches.Count, "1||results|0|1")
            SearchWindow.Input.Value := "x"
            eq("input changed", SearchWindow.IsPrepared(), false)
            SearchWindow.Input.Value := "", SearchWindow.FileMode := true
            eq("file mode", SearchWindow.IsPrepared(), false)
            SearchWindow.FileMode := false, SearchWindow._prepared := false
            eq("used", SearchWindow.IsPrepared(), false)
            g.Show("Hide"), g.Show("NoActivate")
            searches.Count := 0, SearchWindow.PrepareEmpty()
            eq("visible: not prepared", SearchWindow.IsPrepared() "|" searches.Count, "0|0")
            a := [ResultItem("Notepad", "C:\Windows\notepad.exe"), ResultItem("Paint", "mspaint.exe")]
            b := [ResultItem("Notepad", "C:\Windows\notepad.exe"), ResultItem("Paint", "mspaint.exe")]
            eq("same results", SearchWindow.SameResults(a, b), true)
            b[2] := ResultItem("Paint 3D", "mspaint.exe")
            eq("changed title", SearchWindow.SameResults(a, b), false)
            eq("changed count", SearchWindow.SameResults(a, [a[1]]), false)
        } finally {
            SearchWindow.Mode := saved[1], SearchWindow.FileMode := saved[2], SearchWindow._prepared := saved[3]
            SearchWindow.DefineProp("_RunSearch", savedSearch)
            SearchWindow.Input := savedInput, SearchWindow.Gui := savedGui
            g.Destroy()
        }
    }

    ; 呼出时恢复上次的搜索: 先不搜文件 (马上显示窗口), 显示后再完整搜一次并选中上次的那一行; 已经开始打字就不用了
    static RestoreQuick() {
        eq := (n, a, e) => TestRunner.Equal("RestoreQuick." n, a, e)
        saved := Map()
        for name in ["SetResults", "MoveSelection"]
            saved[name] := SearchWindow.GetOwnPropDesc(name)
        savedRegistry := ProviderRegistry.GetOwnPropDesc("Search")
        savedInput := SearchWindow.Input, savedGui := SearchWindow.Gui, savedTimer := SearchWindow._searchTimer
        savedState := [SearchWindow.Mode, SearchWindow.FileMode, SearchWindow.Results, SearchWindow.Selected]
        calls := []
        ProviderRegistry.DefineProp("Search", {Call: (this, text, skip := "") => (calls.Push((skip = FileSearchProvider ? "quick " : "full ") text)
            , [ResultItem("A", "a"), ResultItem("B", "b"), ResultItem("C", "c")])})
        SearchWindow.DefineProp("SetResults", {Call: (this, results) => (SearchWindow.Results := results, SearchWindow.Selected := results.Length ? 1 : 0)})
        SearchWindow.DefineProp("MoveSelection", {Call: (this, delta) => SearchWindow.Selected += delta})
        searches := {Count: 0}
        g := Gui()
        SearchWindow.Gui := g, SearchWindow.Input := g.AddEdit("w200")
        try {
            SearchWindow.Mode := "results", SearchWindow.FileMode := false
            SearchWindow._ShowQuickResults("rep")
            eq("quick search, no files", calls.Length = 1 && calls[1] = "quick rep", true)
            eq("text restored", SearchWindow.Input.Value, "rep")
            SearchWindow._searchTimer := () => searches.Count++
            SearchWindow._OnInputChange()                                   ; 自己填的文字引起的 Change: 不另外搜
            Sleep(30)
            eq("own text change ignored", searches.Count, 0)
            g.Show("NoActivate")
            gen := SearchWindow._searchGen
            SearchWindow._FinishRestore(gen, {Text: "rep", Selected: 3})
            eq("full search after showing", calls.Length = 2 && calls[2] = "full rep", true)
            eq("last row selected", SearchWindow.Selected, 3)
            eq("quiet text cleared", SearchWindow._quietText, "")
            SearchWindow._ShowQuickResults("rep")
            gen := SearchWindow._searchGen
            SearchWindow._quietText := "", SearchWindow.Input.Value := "repo"   ; 显示后马上打了字
            SearchWindow._FinishRestore(gen, {Text: "rep", Selected: 3})
            eq("typed: no full search for the old text", calls.Length, 3)
            SearchWindow.FileMode := true
            SearchWindow._ShowQuickResults("plan")
            eq("file mode: nothing until the full search", SearchWindow.Results.Length "|" calls.Length, "0|3")
        } finally {
            for name, desc in saved
                SearchWindow.DefineProp(name, desc)
            ProviderRegistry.DefineProp("Search", savedRegistry)
            SearchWindow.Mode := savedState[1], SearchWindow.FileMode := savedState[2], SearchWindow.Results := savedState[3], SearchWindow.Selected := savedState[4]
            SearchWindow.Input := savedInput, SearchWindow.Gui := savedGui, SearchWindow._searchTimer := savedTimer, SearchWindow._quietText := ""
            g.Destroy()
        }
    }

    static TypeAhead() {
        eq := (n, a, e) => TestRunner.Equal("TypeAhead." n, a, e)
        saved := Map()
        for name in ["_RunSearch", "_Execute", "Hide"]
            saved[name] := SearchWindow.GetOwnPropDesc(name)
        savedInput := SearchWindow.Input, savedGui := SearchWindow.Gui
        options := AppSettings.Feature("FileSearch"), savedOptions := [options["SpacePrefix"], options["Enabled"]]
        calls := {Text: ""}
        SearchWindow.DefineProp("_RunSearch", {Call: (*) => calls.Text .= "search:" SearchWindow.Input.Value ";"})
        SearchWindow.DefineProp("_Execute", {Call: (this, modifier := "") => calls.Text .= "run:" modifier ";"})
        SearchWindow.DefineProp("Hide", {Call: (*) => calls.Text .= "hide;"})
        g := Gui()
        SearchWindow.Gui := g, SearchWindow.Input := g.AddEdit("w200")
        hwnd := SearchWindow.Input.Hwnd
        apply := (text, key := "", modifier := "") => (calls.Text := "", SearchWindow._ApplyTypeAhead({Text: text, Key: key, Modifier: modifier}), calls.Text)
        SearchWindow.Mode := "results", SearchWindow.FileMode := false
        options["SpacePrefix"] := 1, options["Enabled"] := 1
        try {
            SearchWindow.Input.Value := ""
            eq("nothing typed", apply(""), "")
            eq("text typed", apply("note") "|" SearchWindow.Input.Value, "search:note;|note")
            SendMessage(0xB1, 0, -1, hwnd)                                  ; 恢复的上次搜索是全选的
            eq("replaces selected text", apply("calc") "|" SearchWindow.Input.Value, "search:calc;|calc")
            SendMessage(0xB1, 4, 4, hwnd)
            eq("appends at the caret", apply("s") "|" SearchWindow.Input.Value, "search:calcs;|calcs")
            eq("enter runs the first result", apply("", "Enter"), "run:;")
            SearchWindow.Input.Value := ""
            eq("text then ctrl+enter", apply("vsc", "NumpadEnter", "ctrl"), "search:vsc;run:ctrl;")
            eq("escape hides", apply("x", "Escape"), "hide;")
            SearchWindow.Input.Value := ""
            eq("leading space enters file mode", apply(" seismic") "|" SearchWindow.FileMode "|" SearchWindow.Input.Value, "search:seismic;|1|seismic")
            SearchWindow.FileMode := false, options["SpacePrefix"] := 0, SearchWindow.Input.Value := ""
            eq("space prefix off", apply(" a") "|" SearchWindow.FileMode, "search: a;|0")
            eq("no hook: nothing", SearchWindow._TakeTypeAhead().Text, "")
        } finally {
            SearchWindow.FileMode := false
            options["SpacePrefix"] := savedOptions[1], options["Enabled"] := savedOptions[2]
            for name, desc in saved
                SearchWindow.DefineProp(name, desc)
            SearchWindow.Input := savedInput, SearchWindow.Gui := savedGui
            g.Destroy()
        }
        ; 真正的 InputHook: 模拟打字被拦下记住, 停止后不再拦
        SearchWindow._typeAhead := ""
        ih := InputHook("T3")
        ih.KeyOpt("{Enter}{NumpadEnter}{Esc}", "ES")
        ih.Start(), SearchWindow._typeAhead := ih
        SendEvent("{Text}ab")
        SendEvent("{Enter}")
        typed := SearchWindow._TakeTypeAhead()
        eq("hook collects text", typed.Text "|" typed.Key, "ab|Enter")
        eq("hook stopped", ih.InProgress "|" (SearchWindow._typeAhead = "" ? 1 : 0), "0|1")
    }

    ; 列表画完了但自绘回调没有运行 (空白行): 写日志并再画一次; 回调运行过、还没画时不算
    ; 后备字体预热: 在内存 DC 上用结果行的字体量一次中日韩文字 (见 SearchWindow.WarmUpFonts)
    static FontWarmUp() {
        eq := (n, a, e) => TestRunner.Equal("FontWarmUp." n, a, e)
        saved := [SearchWindow._gdi, Logger.Enabled, Logger.File, Logger.Immediate]
        Logger.Enabled := true, Logger.Immediate := true, Logger.File := A_Temp "\ALTRunFontTest.log"
        try FileDelete(Logger.File)
        created := []
        try {
            SearchWindow._gdi := Map()
            SearchWindow.WarmUpFonts()                                      ; 字体还没建好: 什么都不做
            eq("no fonts yet: nothing done", FileExist(Logger.File) ? 1 : 0, 0)
            for key in ["TitleFont", "SubtitleFont"]
                created.Push(SearchWindow._gdi[key] := SearchWindow._CreateFont("Segoe UI", 12, 400))
            SearchWindow.WarmUpFonts()
            logText := FileExist(Logger.File) ? FileRead(Logger.File, "UTF-8") : ""
            TestRunner.True("FontWarmUp.logged", InStr(logText, "Perf: font warm-up") > 0)
            text := SearchWindow.FontWarmUpText
            eq("Chinese, Japanese and Korean", (RegExMatch(text, "[\x{4E00}-\x{9FFF}]") && RegExMatch(text, "[\x{3040}-\x{30FF}]")
                && RegExMatch(text, "[\x{AC00}-\x{D7AF}]")) ? 1 : 0, 1)
        } finally {
            for font in created
                DllCall("DeleteObject", "Ptr", font)
            SearchWindow._gdi := saved[1], Logger.Enabled := saved[2], Logger.File := saved[3], Logger.Immediate := saved[4]
            try FileDelete(A_Temp "\ALTRunFontTest.log")
        }
    }

    ; 自绘回调不可打断 (见 SearchWindow._OnCustomDraw): Critical 在这个线程里一直有效到线程结束, 调用后检查再关掉
    static CustomDrawCritical() {
        nmcd := Buffer(A_PtrSize * 4 + 64, 0)
        NumPut("UInt", 0x1, nmcd, A_PtrSize * 3)                            ; CDDS_PREPAINT
        saved := SearchWindow._paints
        Critical("Off")
        result := SearchWindow._OnCustomDraw(nmcd.Ptr)
        wasCritical := A_IsCritical
        Critical("Off")
        SearchWindow._paints := saved
        TestRunner.Equal("CustomDrawCritical.prepaint result", result, 0x20)   ; CDRF_NOTIFYITEMDRAW
        TestRunner.True("CustomDrawCritical.critical", wasCritical != 0)
    }

    ; CapsLock 切换中 / 英模式要做什么 (Win.ImeTarget), 包括读数不跟着变的输入法 (#154)
    ; 启动后预读常用图标: 选过的按次数排在前面; 去重、跳过已缓存的、最多 PrefetchMax 个; 窗口显示时不读
    static IconPrefetch() {
        eq := (n, a, e) => TestRunner.Equal("IconPrefetch." n, a, e)
        savedPicks := Knowledge.Picks
        Knowledge.Picks := Map("app:b", Map("Count", 2), "app:c", Map("Count", 5), "app:d", Map("Count", 2), "app:e", Map("Count", 0))
        order := ""
        for item in ProviderRegistry.ByPicks([{Uid: "app:a"}, {Uid: "app:b"}, {Uid: "app:c"}, {Uid: "app:d"}, {Uid: "app:e"}, {Uid: "app:f"}])
            order .= SubStr(item.Uid, 5)
        eq("picked first, by count, ties keep order", order, "cbdaef")
        Knowledge.Picks := Map("x", "not a map")
        eq("bad pick data is ignored", ProviderRegistry.ByPicks([{Uid: "x"}]).Length, 1)
        Knowledge.Picks := savedPicks

        savedMax := IconCache.PrefetchMax, savedWhen := IconCache.PrefetchWhen
        cached := "ext:.prefetchcached"
        IconCache.GetNow(cached)
        IconCache.PrefetchWhen := () => false                               ; 窗口显示中
        IconCache.PrefetchMax := 3
        IconCache.Prefetch(["ext:.pf1", "", "ext:.PF1", cached, "ext:.pf2", "ext:.pf3", "ext:.pf4"])
        eq("dedupe, skip cached, limit", IconCache.PrefetchPending, 3)
        IconCache._PrefetchStep()
        eq("waits while the window is shown", IconCache.PrefetchPending, 3)
        IconCache.PrefetchWhen := () => true
        start := A_TickCount
        while (IconCache.PrefetchPending && A_TickCount - start < 3000)
            IconCache._PrefetchStep()
        eq("all loaded", IconCache.PrefetchPending, 0)
        eq("in the cache", IconCache._icons.Has(IconCache._CacheKey("ext:.pf3")), true)
        eq("over the limit not loaded", IconCache._icons.Has(IconCache._CacheKey("ext:.pf4")), false)
        IconCache.Prefetch([])
        eq("empty list", IconCache.PrefetchPending, 0)
        IconCache.PrefetchMax := savedMax, IconCache.PrefetchWhen := savedWhen

        savedShowing := SearchWindow._showing                               ; 显示到一半 (窗口还没出来) 时也不预读
        SearchWindow._showing := true
        eq("not while showing", SearchWindow._CanPrefetch(), false)
        SearchWindow._showing := false
        eq("hidden and not showing", SearchWindow._CanPrefetch(), !SearchWindow.IsVisible())
        SearchWindow._showing := savedShowing
    }

    ; 图标的磁盘缓存: 像素来回转换、文件格式 (按最近用到的保留、损坏的文件不用)、从缓存直接显示、过期判断
    static IconDiskCache() {
        eq := (n, a, e) => TestRunner.Equal("IconDiskCache." n, a, e)
        same := (a, b) => (a.Size = b.Size && DllCall("ntdll\RtlCompareMemory", "Ptr", a, "Ptr", b, "UPtr", a.Size, "UPtr") = a.Size)
        savedSize := IconCache.Size, savedDir := IconCache.DiskDir, savedDisk := IconCache._disk
        dir := A_Temp "\ALTRunIconDiskTest"
        try DirDelete(dir, true)
        IconCache.Size := 32, IconCache.DiskDir := dir, IconCache._disk := Map()
        try {
            hIcon := IconCache._Load("res:shell32.dll,-24")
            pixels := IconCache.IconPixels(hIcon)
            eq("pixels from icon", IsObject(pixels) ? pixels.Size : 0, 32 * 32 * 4)
            remade := IconCache.MakeIcon(pixels)
            eq("icon from pixels", IconCache.IconWidth(remade), 32)
            again := IconCache.IconPixels(remade)
            eq("round trip keeps pixels", IsObject(again) && same(pixels, again), true)
            for handle in [hIcon, remade]
                DllCall("DestroyIcon", "Ptr", handle)
            eq("short data", IconCache.MakeIcon(Buffer(16)), 0)

            entries := Map()
            entries["c:\a.exe"] := {Stamp: "20260101|10", Saved: 20260101, Used: 20260105, Data: pixels, Offset: 0}
            entries["ext:.pdf"] := {Stamp: "", Saved: 20260102, Used: 20260110, Data: pixels, Offset: 0}
            entries["res:x.dll,-1"] := {Stamp: "", Saved: 20260103, Used: 20250101, Data: pixels, Offset: 0}
            parsed := IconCache._ParseDisk(IconCache._SerializeDisk(entries, 32, 2), 32)
            eq("keeps the most recently used", parsed.Count "|" parsed.Has("ext:.pdf") "|" parsed.Has("c:\a.exe") "|" parsed.Has("res:x.dll,-1"), "2|1|1|0")
            e := parsed["c:\a.exe"]
            eq("fields", e.Stamp "|" e.Saved "|" e.Used, "20260101|10|20260101|20260105")
            stored := Buffer(32 * 32 * 4)
            DllCall("RtlMoveMemory", "Ptr", stored, "Ptr", e.Data.Ptr + e.Offset, "UPtr", stored.Size)
            eq("pixels kept", same(stored, pixels), true)
            eq("other size ignored", IconCache._ParseDisk(IconCache._SerializeDisk(entries, 32, 9), 24).Count, 0)
            eq("bad header ignored", IconCache._ParseDisk(Buffer(40, 0), 32).Count, 0)
            eq("empty cache", IconCache._ParseDisk(IconCache._SerializeDisk(Map(), 32, 9), 32).Count, 0)
            old := IconCache._SerializeDisk(entries, 32, 9)
            NumPut("UShort", 1, old, 4)
            eq("older format ignored", IconCache._ParseDisk(old, 32).Count, 0)
            cut := IconCache._SerializeDisk(entries, 32, 9), cut.Size -= 100
            rejected := false
            try IconCache._ParseDisk(cut, 32)
            catch
                rejected := true
            eq("truncated file rejected", rejected, true)

            cacheFile := dir "\Icons-32.bin"
            IconCache._disk := entries, IconCache._diskDirty := true
            eq("save", IconCache.SaveDisk(cacheFile), true)
            eq("nothing to save", IconCache.SaveDisk(cacheFile), false)
            eq("load", IconCache.LoadDisk(cacheFile), 3)
            ; 分块写 (每块很小时写很多块), 内容一样; 没有留下临时文件
            savedChunk := IconCache.DiskChunk, IconCache.DiskChunk := 1000
            try {
                IconCache._disk := entries, IconCache._diskDirty := true
                eq("save in chunks", IconCache.SaveDisk(cacheFile), true)
            } finally {
                IconCache.DiskChunk := savedChunk
            }
            eq("chunks: same size", FileGetSize(cacheFile), IconCache._SerializeDisk(entries, 32, IconCache.DiskMax).Size)
            eq("chunks: no temp file", FileExist(cacheFile ".tmp") ? 1 : 0, 0)
            eq("chunks: load", IconCache.LoadDisk(cacheFile), 3)
            IconCache._disk := entries, IconCache._diskDirty := true, IconCache._saving := true
            eq("not while saving", IconCache.SaveDisk(cacheFile), false)
            IconCache._saving := false
            eq("failed save stays dirty", IconCache.SaveDisk(dir "\bad<name>\Icons-32.bin") "|" IconCache._diskDirty "|" IconCache._saving, "0|1|0")
            eq("save delay batches new icons", IconCache.DiskSaveDelay >= 30000, true)
            f := FileOpen(cacheFile, "w"), f.RawWrite(Buffer(30, 0x41), 30), f.Close()
            eq("broken file", IconCache.LoadDisk(cacheFile), 0)
            eq("missing file", IconCache.LoadDisk(dir "\none.bin"), 0)

            key := "ext:.diskcachetest"
            IconCache._disk := Map(key, {Stamp: "", Saved: 20200101, Used: 20200101, Data: pixels, Offset: 0})
            IconCache._verified := Map(), IconCache._verify := []
            hIcon := IconCache.Get(key)
            eq("shown from the disk cache at once", hIcon != 0 && !IconCache._queue.Has(key), true)
            eq("queued for checking", IconCache._verify.Length, 1)
            eq("used today", IconCache._disk[key].Used, IconCache._Today())
            eq("old non-file entry is stale", IconCache.IsStale(key, IconCache._disk[key]), true)
            eq("recent non-file entry is fresh", IconCache.IsStale(key, {Saved: IconCache._Today()}), false)
            IconCache._verify := []
            try IconCache._icons.Delete(key)

            target := dir "\stamp.ico"
            FileAppend("one", target)
            stamp := IconCache.Stamp(target)
            eq("file key", IconCache.IsFileKey(StrLower(target)) "|" IconCache.IsFileKey("shell:appsfolder\x") "|" IconCache.IsFileKey("::{abc}"), "1|0|0")
            eq("same file is fresh", IconCache.IsStale(StrLower(target), {Stamp: stamp}), false)
            FileAppend("two", target)
            eq("changed file is stale", IconCache.IsStale(StrLower(target), {Stamp: stamp}), true)
            eq("missing file stamp", IconCache.Stamp(dir "\none.exe"), "")

            IconCache._disk := Map()
            hIcon := IconCache._Load("folder:")
            IconCache._Remember("thumb:c:\x.png", hIcon)
            eq("thumbnails not kept", IconCache._disk.Count, 0)
            IconCache._Remember("folder:", hIcon)
            eq("new icon kept", IconCache._disk.Has("folder:") && IconCache._disk["folder:"].Saved = IconCache._Today(), true)
            DllCall("DestroyIcon", "Ptr", hIcon)
            IconCache._disk := Map()
            Loop 5
                IconCache._disk["k" A_Index] := {Stamp: "", Saved: 1, Used: 20260100 + A_Index, Data: pixels, Offset: 0}
            IconCache._TrimDisk(3)
            eq("trim keeps the most recently used", IconCache._disk.Count "|" IconCache._disk.Has("k5") "|" IconCache._disk.Has("k1"), "3|1|0")
            IconCache._disk := Map("ext:.cleartest", {Stamp: "", Saved: 1, Used: 1, Data: pixels, Offset: 0})
            IconCache.Clear()
            eq("clear empties the disk cache", IconCache._disk.Count, 0)
        } finally {
            if (IconCache._saveTimer != "")
                SetTimer(IconCache._saveTimer, 0)
            if (IconCache._verifyTimer != "")
                SetTimer(IconCache._verifyTimer, 0)
            IconCache._diskDirty := false, IconCache._diskTouched := false
            IconCache.Size := savedSize, IconCache.DiskDir := savedDir, IconCache._disk := savedDisk
            try DirDelete(dir, true)
        }
    }

    ; 偏好设置的页面用到时才建: 打开时只有要显示的那一页; 换页时建, Tab 键顺序 列表 -> 页面 -> 底部按钮;
    ; 建页面时程序自己填的值不算修改, 之后的修改照常记录
    static PreferencesLazyPages() {
        eq := (n, a, e) => TestRunner.Equal("PreferencesLazyPages." n, a, e)
        PreferencesWindow.Show("Prefs.Page.Appearance", -3000, -3000)
        try {
            built := ""
            for page in PreferencesWindow.Pages
                built .= page.Built ? page.Key " " : ""
            eq("only the shown page", Trim(built), "Prefs.Page.Appearance")
            eq("page list has every page", PreferencesWindow.PageList.GetCount(), PreferencesWindow.PageKeys.Length)
            Sleep(400)
            commands := PreferencesWindow._PageIndex("Prefs.Page.Commands")
            PreferencesWindow._SelectRow(commands)
            eq("built when selected", PreferencesWindow.Pages[commands].Built, true)
            eq("building is not a change", PreferencesWindow._dirty, false)
            eq("ready again", PreferencesWindow._ready, true)
            first := ""
            for ctrl in PreferencesWindow.Pages[commands].Controls
                if (!first && ctrl.Visible && (DllCall("GetWindowLong", "Ptr", ctrl.Hwnd, "Int", -16, "Int") & 0x10000))   ; WS_TABSTOP
                    first := ctrl.Hwnd
            next := DllCall("GetNextDlgTabItem", "Ptr", PreferencesWindow.Gui.Hwnd, "Ptr", PreferencesWindow.PageList.Hwnd, "Int", 0, "Ptr")
            eq("tab from the list goes into the page", next = first && next != PreferencesWindow._helpButton.Hwnd, true)
            general := PreferencesWindow._PageIndex("Prefs.Page.General")
            PreferencesWindow._SelectRow(general)
            for ctrl in PreferencesWindow.Pages[general].Controls
                if (ctrl.Type = "CheckBox") {
                    ctrl.Value := !ctrl.Value
                    ControlClick(ctrl)                                      ; 用户点一下 (Click 事件)
                    break
                }
            Sleep(100)
            eq("user change after building is recorded", PreferencesWindow._dirty, true)
            built := 0
            for page in PreferencesWindow.Pages
                built += page.Built
            eq("unvisited pages still not built", built, 3)
            ; 建页面的中途窗口被关掉 (别的线程打断了建页面): 不再往已经没有的窗口上加控件, 不报错
            usagePage := PreferencesWindow._PageIndex("Prefs.Page.Usage")
            PreferencesWindow.Pages[usagePage].Builder := "_TestCloseWhileBuilding"
            PreferencesWindow.DefineProp("_TestCloseWhileBuilding", {Call: (this) => (PreferencesWindow.Close(), PreferencesWindow._Add("Text", "x200 w100", "x"))})
            closedOk := true
            try PreferencesWindow._SelectRow(usagePage)
            catch
                closedOk := false
            PreferencesWindow.DeleteProp("_TestCloseWhileBuilding")
            eq("closed while building: no error", closedOk, true)
            eq("closed while building: window gone", IsObject(PreferencesWindow.Gui), false)
        } finally {
            PreferencesWindow._dirty := false
            PreferencesWindow.Close()
        }
    }

    ; 睡眠唤醒 / 解锁 / 显示器变化后重新装键盘钩子: 连着几个事件只装一次; 有鼠标热键时也装鼠标钩子
    static WakeHooks() {
        eq := (n, a, e) => TestRunner.Equal("WakeHooks." n, a, e)
        eq("mouse hotkeys", App.UsesMouseHotkeys([Map("Key", "~MButton", "Action", "PTTools")]), true)
        eq("wheel", App.UsesMouseHotkeys([Map("Key", "^WheelUp", "Action", "x")]), true)
        eq("keyboard only", App.UsesMouseHotkeys([Map("Key", "^!t", "Action", "x"), "bad"]), false)
        eq("none", App.UsesMouseHotkeys([]), false)
        savedDelay := App.WakeDelay
        App.WakeDelay := 50
        App._AfterWake("resume"), App._AfterWake("unlock"), App._AfterWake("display change")
        eq("last reason wins", App._wakeReason, "display change")
        Sleep(150)                                                          ; 定时器运行 ReinstallHooks (不出错)
        App.WakeDelay := savedDelay
        App.ReinstallHooks("test")
        eq("hook still works", A_IsSuspended, 0)
    }

    ; 正在读的图标: 界面又来要时不再排一次队 (以前慢的图标会读两遍)
    ; 后台下载: 异步请求, 完成 / 出错 / 超时后在单独的线程里调用回调 (用假的请求, 不连网)
    static HttpGet() {
        eq := (n, a, e) => TestRunner.Equal("HttpGet." n, a, e)
        saved := [HttpGet.NewRequest, HttpGet.PollMs]
        fakes := []
        HttpGet.NewRequest := () => (fakes.Push(r := FakeHttpRequest()), r)
        HttpGet.PollMs := 10
        got := []
        try {
            HttpGet.Start("https://example.test/ok", (t) => got.Push("done " t), (m) => got.Push("fail " m))
            eq("async GET", fakes[1].Method " " fakes[1].Url " " fakes[1].Async, "GET https://example.test/ok 1")
            eq("asks the server, not the cache", fakes[1].Header, "If-Modified-Since")
            eq("pending", HttpGet.Pending, 1)
            Sleep(60)
            eq("waits until complete", got.Length, 0)
            fakes[1].status := 200, fakes[1].responseText := "hello", fakes[1].readyState := 4
            eq("done", Tests._WaitFor(() => got.Length) ? got[1] : "", "done hello")
            eq("nothing pending", HttpGet.Pending, 0)

            got.Length := 0
            HttpGet.Start("https://example.test/missing", (t) => got.Push("done " t), (m) => got.Push("fail " m))
            fakes[2].status := 404, fakes[2].statusText := "Not Found", fakes[2].readyState := 4
            eq("http error", Tests._WaitFor(() => got.Length) ? got[1] : "", "fail HTTP 404 Not Found")

            got.Length := 0
            HttpGet.Start("https://example.test/offline", (t) => got.Push("done " t), (m) => got.Push("fail " m))
            fakes[3].status := 12007, fakes[3].readyState := 4
            eq("network error", Tests._WaitFor(() => got.Length) ? got[1] : "", "fail Network error 12007")

            got.Length := 0
            HttpGet.Start("https://example.test/bad", (t) => (got.Push("parsing"), JSON.Parse(t)), (m) => got.Push("fail"))
            fakes[4].status := 200, fakes[4].responseText := "{not json", fakes[4].readyState := 4
            Tests._WaitFor(() => got.Length >= 2)
            eq("done throws -> fail", got.Length = 2 ? got[1] " " got[2] : got.Length, "parsing fail")

            got.Length := 0
            HttpGet.Start("https://example.test/slow", (t) => got.Push("done"), (m) => got.Push("fail " m), 50)
            eq("timeout", Tests._WaitFor(() => got.Length) ? got[1] : "", "fail Timed out after 0 s")
            eq("timeout aborts", fakes[5].Aborted, true)
            eq("nothing pending after timeout", HttpGet.Pending, 0)

            ; 一个请求的回调还在运行 (例如弹出了对话框) 时, 另一个照样完成
            got.Length := 0, release := false
            HttpGet.Start("https://example.test/b", (t) => (got.Push("b"), release := true))
            HttpGet.Start("https://example.test/a", (t) => (got.Push("a start"), Tests._WaitFor(() => release, 2000), got.Push("a end")))
            fakes[6].status := 200, fakes[6].readyState := 4
            fakes[7].status := 200, fakes[7].readyState := 4
            Tests._WaitFor(() => got.Length >= 3, 4000)
            eq("callbacks do not block each other", got.Length = 3 ? got[3] : Tests._Join(got), "a end")

            HttpGet.NewRequest := () => ComObject("ALTRun.NoSuchClass")
            got.Length := 0
            HttpGet.Start("https://example.test/x", (t) => got.Push("done"), (m) => got.Push("fail " m))
            eq("cannot start -> fail", Tests._WaitFor(() => got.Length) ? SubStr(got[1], 1, 5) : "", "fail ")
        } finally {
            HttpGet.NewRequest := saved[1], HttpGet.PollMs := saved[2]
        }

        ; 检查更新 / 汇率不再用 Download() (下载期间整个程序都停着); 一键更新下载 zip 时有进度窗口, 仍用 Download()
        for name in ["UpdateChecker", "CurrencyRates"] {
            text := FileRead(A_ScriptDir "\..\Src\Extensions\" name ".ahk", "UTF-8")
            pos := 1, count := 0
            while (pos := RegExMatch(text, "\bDownload\(", , pos)) {
                count++, pos++
            }
            eq(name ": synchronous downloads", count, name = "UpdateChecker" ? 1 : 0)
        }

        ; 检查更新: 版本信息下载完再处理 (假的回复)
        savedState := [UpdateChecker.StateFile, UpdateChecker.Pending]
        root := A_Temp "\ALTRun-test-update-async"
        try DirDelete(root, true)
        UpdateChecker.StateFile := root "\Update.json"
        fakes := []
        HttpGet.NewRequest := () => (fakes.Push(r := FakeHttpRequest()), r)
        HttpGet.PollMs := 10
        try {
            UpdateChecker.Pending := ""
            UpdateChecker.CheckInBackground()
            eq("update: asks GitHub", fakes[1].Url, UpdateChecker.ApiUrl)
            eq("update: not yet", UpdateChecker.Pending, "")
            fakes[1].status := 200, fakes[1].responseText := '{"tag_name": "v2999.01.01", "html_url": "https://x/notes"}', fakes[1].readyState := 4
            Tests._WaitFor(() => IsObject(UpdateChecker.Pending))
            eq("update: found in background", IsObject(UpdateChecker.Pending) ? UpdateChecker.Pending.Version : "", "2999.01.01")
            eq("update: check time saved", UpdateChecker.LoadState()["LastCheck"] != "", true)
            eq("manual check: busy flag starts clear", UpdateChecker._checking, false)
        } finally {
            HttpGet.NewRequest := saved[1], HttpGet.PollMs := saved[2]
            UpdateChecker.StateFile := savedState[1], UpdateChecker.Pending := savedState[2]
            try DirDelete(root, true)
        }

        ; 真的请求 (Msxml2.XMLHTTP.6.0, 只在 GitHub Actions 的 Windows 上): 等待时程序照常运行, 内容正确。
        ; 连不上网 / 超时不算失败 (和程序无关), 只写一行说明
        if (EnvGet("GITHUB_ACTIONS") = "true") {
            got := [], ticks := 0
            HttpGet.Start("https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/LICENSE", (t) => got.Push("done " SubStr(Trim(t, " `t`r`n"), 1, 26)), (m) => got.Push("fail " m))
            start := A_TickCount
            while (!got.Length && A_TickCount - start < 35000)
                Sleep(10), ticks++
            result := got.Length ? got[1] : "fail no answer"
            if RegExMatch(result, "^fail (Network error|Timed out|no answer)")
                FileAppend("NOTE HttpGet.real request skipped: " result "`n", "*")
            else {
                eq("real request", result, "done GNU GENERAL PUBLIC LICENSE")
                eq("real request: kept running while waiting", ticks > 1, true)
            }
        }
    }

    ; 日志文件一直开着 (不是每条都打开、关闭一次): 写完马上能读到; 换文件、文件被删、截断后重新打开
    static LoggerFile() {
        eq := (n, a, e) => TestRunner.Equal("LoggerFile." n, a, e)
        read := (file) => FileExist(file) ? FileRead(file, "UTF-8") : ""
        lines := (file) => (text := read(file), StrSplit(Trim(text, "`r`n"), "`r`n").Length * (text != ""))
        first := A_Temp "\ALTRunLoggerTest1.log", second := A_Temp "\ALTRunLoggerTest2.log"
        saved := [Logger.Enabled, Logger.File, Logger.Immediate]
        Logger.Close()
        for logPath in [first, second, second ".old"]
            try FileDelete(logPath)
        Logger.Enabled := true, Logger.Immediate := true, Logger.File := first
        try {
            Logger.Debug("one 中文")
            eq("kept open", Logger._file != "", true)
            eq("written right away", InStr(read(first), "[DBG] one 中文") > 0, true)
            Logger.Warn("two")
            eq("appended", lines(first), 2)
            Logger.File := second
            Logger.Debug("three")
            eq("new file", lines(second), 1)
            eq("old file unchanged", lines(first), 2)
            Logger.Close()
            eq("closed", Logger._file, "")
            FileDelete(second)
            Logger.Debug("four")
            eq("reopened after close", InStr(read(second), "four") > 0 && lines(second) = 1, true)
            if !DllCall("GetProcAddress", "Ptr", DllCall("GetModuleHandle", "Str", "ntdll", "Ptr"), "AStr", "wine_get_version", "Ptr") {
                FileDelete(second)                                          ; 开着的时候被删掉 (Wine 里删不掉开着的文件, 跳过)
                Logger.Debug("five")
                eq("recreated after delete", read(second) ~= "^\x{FEFF}?[^\r\n]*five\r\n$" ? 1 : 0, 1)
            }
            Logger.Rotate(1)
            eq("rotated", FileExist(second ".old") != "", true)
            Logger.Debug("six")
            eq("new file after rotate", lines(second), 1)
            Logger.Immediate := false
            Logger.Debug("seven")
            eq("buffered", lines(second), 1)
            Logger.Flush()
            eq("flushed", lines(second), 2)
        } finally {
            Logger.Close()
            for logPath in [first, second, second ".old"]
                try FileDelete(logPath)
            Logger.Enabled := saved[1], Logger.File := saved[2], Logger.Immediate := saved[3]
        }
    }

    ; 后台任务一次超过 SlowMs 时记一条 "Perf: slow"; 快的不记; 返回值、参数照常
    static SlowTaskLog() {
        eq := (n, a, e) => TestRunner.Equal("SlowTaskLog." n, a, e)
        saved := [Logger.Enabled, Logger.File, Logger.Immediate, Logger.SlowMs]
        Logger.Enabled := true, Logger.Immediate := true, Logger.File := A_Temp "\ALTRunSlowTaskTest.log", Logger.SlowMs := 30
        try FileDelete(Logger.File)
        try {
            fast := Logger.Watch("fast task", (a, b) => a + b)
            eq("returns the result", fast(2, 3), 5)
            slow := Logger.Watch("slow task", () => (Sleep(60), "ok"))
            eq("slow returns", slow(), "ok")
            failing := Logger.Watch("failing task", () => (Sleep(60), JSON.Parse("{bad")))
            try failing()
            catch
                threw := true
            eq("errors still thrown", IsSet(threw), true)
            Logger.Flush()
            logText := FileExist(Logger.File) ? FileRead(Logger.File, "UTF-8") : ""
            eq("fast not logged", InStr(logText, "fast task") > 0, false)
            eq("slow logged", RegExMatch(logText, "Perf: slow slow task \d+ ms") > 0, true)
            eq("failing logged too", InStr(logText, "Perf: slow failing task") > 0, true)
        } finally {
            try FileDelete(Logger.File)
            Logger.Enabled := saved[1], Logger.File := saved[2], Logger.Immediate := saved[3], Logger.SlowMs := saved[4]
        }
        eq("default threshold", Logger.SlowMs, 50)
        for source, label in Map("Core\Usage", "save usage", "Core\Knowledge", "save knowledge", "Providers\ClipboardProvider", "save clipboard"
                             , "Extensions\QuickSwitch", "quick switch watch", "Core\FileIndex", "file index scan", "UI\IconCache", "save icon cache")
            eq("watched: " label, InStr(FileRead(A_ScriptDir "\..\Src\" source ".ahk", "UTF-8"), 'Logger.Watch("' label '"') > 0, true)
    }

    static _WaitFor(fn, ms := 2000) {
        start := A_TickCount
        while !fn() {
            if (A_TickCount - start > ms)
                return false
            Sleep(10)
        }
        return true
    }

    static _Join(list) {
        text := ""
        for item in list
            text .= (A_Index > 1 ? "," : "") item
        return text
    }

    static IconLoadingOnce() {
        eq := (n, a, e) => TestRunner.Equal("IconLoadingOnce." n, a, e)
        key := IconCache._CacheKey("ext:.loadingonce")
        IconCache._loading[key] := true
        try {
            eq("not ready", IconCache.Get("ext:.loadingonce"), 0)
            eq("not queued again", IconCache._queue.Has(key), false)
        } finally {
            IconCache._loading.Delete(key)
        }
        IconCache.Get("ext:.loadingonce")
        eq("queued when not loading", IconCache._queue.Has(key), true)
        try IconCache._queue.Delete(key)
    }

    ; 淡入淡出: 透明度曲线; 真的窗口淡入时先透明、结束后是主题的透明度; 淡出时留一个影子, 结束后消失
    static WindowFade() {
        eq := (n, a, e) => TestRunner.Equal("WindowFade." n, a, e)
        eq("alpha start", WindowFade.Alpha(0, 120, 0, 255), 0)
        eq("alpha end", WindowFade.Alpha(120, 120, 0, 255), 255)
        eq("alpha after end", WindowFade.Alpha(500, 120, 0, 230), 230)
        eq("alpha ease-out (half way is past half)", WindowFade.Alpha(60, 120, 0, 255) > 200, true)
        eq("alpha fade out", WindowFade.Alpha(50, 100, 255, 0) < 64, true)
        eq("alpha no duration", WindowFade.Alpha(0, 0, 255, 0), 0)
        eq("alpha before start", WindowFade.Alpha(-5, 100, 0, 255), 0)
        eq("setting off by default", AppSettings.Defaults()["Appearance"]["Fade"], 0)

        g := Gui("-Caption +ToolWindow")
        g.Add("Text", "w200 h60", "fade")
        DetectHiddenWindows(true)
        WindowFade.In(g.Hwnd, 230)
        eq("in: transparent before showing", WinGetTransparent("ahk_id " g.Hwnd), 0)
        eq("in: running", WindowFade.Running, true)
        g.Show("NA x10 y10 w200 h60")
        WindowFade.Stop()
        eq("in: theme opacity after", WinGetTransparent("ahk_id " g.Hwnd), 230)
        eq("in: stopped", WindowFade.Running, false)

        WindowFade.In(g.Hwnd, 255)
        WindowFade.Out(g.Hwnd)                                              ; 淡入还没完就隐藏: 窗口恢复不透明, 换成影子
        eq("out: window opaque again", WinGetTransparent("ahk_id " g.Hwnd), 255)
        eq("out: ghost shown", WindowFade.GhostVisible, true)
        g.Hide()
        start := A_TickCount
        while (WindowFade.Running && A_TickCount - start < 2000)
            Sleep(10)
        eq("out: finished by itself", WindowFade.Running, false)
        eq("out: ghost gone", WindowFade.GhostVisible, false)

        WindowFade.Out(g.Hwnd)                                              ; 窗口不在屏幕上: 没有影子
        eq("out: hidden window has no ghost", WindowFade.Running || WindowFade.GhostVisible, false)
        g.Destroy()
    }

    static ImeToggle() {
        show := (p) => (p.Chinese ? "zh" : "en") " open " p.Open " conv " p.Conv (p.Remembered ? " remembered" : "")
        eq := (n, p, e) => TestRunner.Equal("ImeToggle." n, show(p), e)
        eq("english -> chinese", Win.ImeTarget(1, 0), "zh open 1 conv 1")
        eq("chinese -> english (also closes the IME)", Win.ImeTarget(1, 1025), "en open 0 conv 1024")
        eq("closed -> open chinese", Win.ImeTarget(0, 0), "zh open 1 conv 1")
        eq("closed with native bit -> open chinese", Win.ImeTarget(0, 1025), "zh open 1 conv 1025")
        ; 读数可信: 切换后的读数和这次不同, 按读数
        eq("reliable readings", Win.ImeTarget(1, 1, {Open: 0, Conv: 0, Chinese: false}), "en open 0 conv 0")
        ; 微信输入法 (#154 的日志): 读数跟着设置变, 一直是打开; 上次切到了中文 -> 这次英文 (关闭输入法)
        eq("wechat ime: to english", Win.ImeTarget(1, 0x401, {Open: 1, Conv: 0x401, Chinese: true}), "en open 0 conv 1024 remembered")
        ; 读数不跟着变: 一直读到 "没打开 / 0", 上次切到了中文 -> 这次切回英文, 再下次又切到中文
        eq("stuck readings: back to english", Win.ImeTarget(0, 0, {Open: 0, Conv: 0, Chinese: true}), "en open 0 conv 0 remembered")
        eq("stuck readings: chinese again", Win.ImeTarget(0, 0, {Open: 0, Conv: 0, Chinese: false}), "zh open 1 conv 1 remembered")
        ; 设置的顺序: 打开时先打开再设转换模式, 关闭时先设转换模式再关闭 (6 = IMC_SETOPENSTATUS, 2 = IMC_SETCONVERSIONMODE)
        steps(open, conv) {
            list := ""
            for step in Win.ImeSteps(open, conv)
                list .= " " step[1] ":" step[2]
            return Trim(list)
        }
        TestRunner.Equal("ImeToggle.steps: open", steps(1, 0x401), "6:1 2:1025")
        TestRunner.Equal("ImeToggle.steps: close", steps(0, 0x400), "2:1024 6:0")
        ; 呼出时切到英文模式 (SwitchToEnglishInput = 2): 去掉中文位, 关闭输入法, 其它位不变
        english := Win.ImeEnglish(0x409)
        TestRunner.Equal("ImeToggle.english mode", english.Open " " english.Conv, "0 1032")
        english := Win.ImeEnglish(0)
        TestRunner.Equal("ImeToggle.english mode (already english)", english.Open " " english.Conv, "0 0")
    }

    static SkippedPaint() {
        eq := (n, a, e) => TestRunner.Equal("SkippedPaint." n, a, e)
        saved := [SearchWindow.Gui, SearchWindow.List, SearchWindow.Input, SearchWindow.Results, Logger.Enabled, Logger.File, Logger.Immediate]
        g := Gui("-Caption")
        input := g.AddEdit("w200")
        list := g.AddListView("w200 h100 -Hdr", [""])                       ; 没有接自绘回调: 每次画都相当于回调被跳过
        list.Add("", "")
        g.Show("NA x0 y0")
        SearchWindow.Gui := g, SearchWindow.List := list, SearchWindow.Input := input, SearchWindow.Results := []
        Logger.Enabled := true, Logger.Immediate := true, Logger.File := A_Temp "\ALTRunPaintTest.log"
        try FileDelete(Logger.File)
        pending := () => DllCall("GetUpdateRect", "Ptr", list.Hwnd, "Ptr", 0, "Int", 0)
        paint := () => (DllCall("InvalidateRect", "Ptr", list.Hwnd, "Ptr", 0, "Int", 0), DllCall("UpdateWindow", "Ptr", list.Hwnd))
        try {
            SearchWindow._skipLogged := false
            SearchWindow._WatchPaint(), paint(), SearchWindow._paints++     ; 回调运行过
            SearchWindow._CheckPaint()
            eq("drawn: no repaint", pending() ? 1 : 0, 0)
            DllCall("InvalidateRect", "Ptr", list.Hwnd, "Ptr", 0, "Int", 0)
            SearchWindow._WatchPaint(), SearchWindow._CheckPaint()          ; 还没画
            eq("not painted yet: wait", SearchWindow._paintTries "|" SearchWindow._paintRetries, "1|0")
            DllCall("UpdateWindow", "Ptr", list.Hwnd)
            Critical(1000)                                                  ; 执行中不处理消息 (写日志慢时, 补画的请求会先被处理掉)
            SearchWindow._WatchPaint(), paint(), SearchWindow._CheckPaint()  ; 画了, 回调没运行
            eq("skipped: repaint", (pending() ? 1 : 0) "|" SearchWindow._paintRetries, "1|1")
            Critical("Off")
            logText := FileExist(Logger.File) ? FileRead(Logger.File, "UTF-8") : ""
            TestRunner.True("SkippedPaint.logged", InStr(logText, "[WRN] SearchWindow: custom draw skipped") > 0)
            DllCall("UpdateWindow", "Ptr", list.Hwnd)
            SearchWindow._paintRetries := 3, SearchWindow._CheckPaint()     ; 最多补画 3 次
            eq("gives up after 3 repaints", pending() ? 1 : 0, 0)
            SearchWindow._WatchPaint(), paint(), SearchWindow._CheckPaint()
            logText := FileExist(Logger.File) ? FileRead(Logger.File, "UTF-8") : ""
            StrReplace(logText, "custom draw skipped", , , &count)
            eq("logged once per show", count, 1)
        } finally {
            Critical("Off")
            SetTimer(SearchWindow._paintCheck, 0)
            try FileDelete(Logger.File)
            SearchWindow.Gui := saved[1], SearchWindow.List := saved[2], SearchWindow.Input := saved[3], SearchWindow.Results := saved[4]
            Logger.Enabled := saved[5], Logger.File := saved[6], Logger.Immediate := saved[7]
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
    ; PT Tools 窗口: 结果框不停留 Tab 焦点 (Area Expression 之后直接到 Rebar Diameter), 单位写成 mm², 点进输入框时全选
    static PTToolsWindowUi() {
        eq := (n, a, e) => TestRunner.Equal("PTToolsWindowUi." n, a, e)
        saved := [PTToolsWindow.Settings["WinLeft"], PTToolsWindow.Settings["WinTop"]]
        PTToolsWindow.Settings["WinLeft"] := -3000, PTToolsWindow.Settings["WinTop"] := -3000
        PTToolsWindow.Show()
        g := PTToolsWindow.G
        tabbed := "", oldUnit := ""
        for ctrl in g {
            style := ControlGetStyle(ctrl)
            if (ctrl.Type = "Edit" && (style & 0x800) && (style & 0x10000))     ; ES_READONLY 还有 WS_TABSTOP
                tabbed .= ctrl.Name " "
            if InStr(ctrl.Text, "mm2")
                oldUnit .= ctrl.Text " | "
        }
        eq("read-only fields skip tab", tabbed, "")
        eq("units", oldUnit, "")
        next := DllCall("GetNextDlgTabItem", "Ptr", g.Hwnd, "Ptr", g["RequiredAreaExpression"].Hwnd, "Int", 0, "Ptr")
        eq("expression -> diameter", next = g["RequiredBarDiameter"].Hwnd, true)
        span := g["RebarSpanWidth"]
        SendMessage(0xB1, 1, 1, span)                                       ; 光标放在中间
        span.Focus()
        Sleep(100)                                                          ; 处理 PostMessage 的全选
        start := 0, finish := 0
        DllCall("SendMessage", "Ptr", span.Hwnd, "UInt", 0xB0, "UInt*", &start, "UInt*", &finish)   ; EM_GETSEL
        eq("select all on focus", start "-" finish, "0-" StrLen(span.Value))
        eq("right margin", (SendMessage(0xD4, 0, 0, span) >> 16) >= 4, true)   ; EM_GETMARGINS: 光标不紧贴数字
        g["RebarSpanWidth"].Value := 5000                                   ; 第一组的 Span Width -> BRC 的 Span Width -> 钢筋网面积一起重算
        PTToolsWindow.OnRebarFieldChanged()
        eq("brc span follows", g["BrcSpanWidth"].Value, "5000")
        eq("top mesh area recalculated", g["BrcTopArea"].Value, PTToolsWindow.Fmt2(5 * PTToolsWindow.CalculateMeshArea(g["BrcTopMeshMark"].Value)))
        eq("bottom mesh area recalculated", g["BrcBotArea"].Value, PTToolsWindow.Fmt2(5 * PTToolsWindow.CalculateMeshArea(g["BrcBotMeshMark"].Value)))
        safety := g["SafetyFactorEnabled"]
        eq("safety factor checkbox skips tab", (ControlGetStyle(safety) & 0x10000) != 0, false)
        ; 钢筋直径下拉框里的数字右对齐: 第一个字的位置 (EM_POSFROMCHAR) 在输入框的右半边 (Wine 不支持建好后改对齐, 只在 Windows 上看位置)
        diaEdit := DllCall("FindWindowEx", "Ptr", g["RebarBarDiameter"].Hwnd, "Ptr", 0, "Str", "Edit", "Ptr", 0, "Ptr")
        eq("diameter diaEdit has ES_RIGHT", (ControlGetStyle(diaEdit) & 0x2) != 0, true)
        if !DllCall("GetProcAddress", "Ptr", DllCall("GetModuleHandle", "Str", "ntdll", "Ptr"), "AStr", "wine_get_version", "Ptr") {
            WinGetClientPos(, , &ew, , diaEdit)
            firstX := SendMessage(0xD6, 0, 0, diaEdit) & 0xFFFF                     ; EM_POSFROMCHAR: 第 0 个字的 x
            TestRunner.True("PTToolsWindowUi.diameter text on the right (first char at " firstX " of " ew ")", firstX > ew // 2)
        }
        g.Destroy(), PTToolsWindow.G := ""
        PTToolsWindow.Settings["WinLeft"] := saved[1], PTToolsWindow.Settings["WinTop"] := saved[2]
    }

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
Dir | Q:\Projects\Design-2019\PT1001 - 24 Harbor | Tower A, B, C=3
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
        eq("folder command", commands["Tower A, B, C"]["Target"], "Q:\Projects\Design-2019\PT1001 - 24 Harbor")
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
    ; 显示搜索窗口时不重入, 但上一次卡住超过 StuckShowMs 就不再挡住新的显示 (关掉一台显示器后卡了 2 分半)
    static StuckShow() {
        eq := (n, a, e) => TestRunner.Equal("StuckShow." n, a, e)
        saved := [SearchWindow._showing, SearchWindow._showStarted, SearchWindow._showStep]
        SearchWindow._showing := false
        eq("not showing", SearchWindow._ShouldSkipShow(), false)
        SearchWindow._showing := true, SearchWindow._showStarted := A_TickCount, SearchWindow._showStep := "redraw"
        eq("still showing", SearchWindow._ShouldSkipShow(), true)
        SearchWindow._showStarted := A_TickCount - SearchWindow.StuckShowMs - 100
        eq("stuck", SearchWindow._ShouldSkipShow(), false)
        SearchWindow._showing := saved[1], SearchWindow._showStarted := saved[2], SearchWindow._showStep := saved[3]
    }

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
        savedArea := SearchWindow.LastArea, savedAt := SearchWindow.HiddenAt, savedShowOn := AppSettings.Appearance["ShowOn"]
        SearchWindow.LastArea := second, SearchWindow.HiddenAt := A_TickCount                          ; 搜索窗口刚在第二块屏幕上隐藏
        eq("just hidden: same screen", Hud._Anchor().Area.Left, 1920)
        SearchWindow.HiddenAt := A_TickCount - Hud.RecentWindowMs - 1
        AppSettings.Appearance["ShowOn"] := "Primary"
        primary := Win.WorkArea(MonitorGetPrimary())
        eq("later: ShowOn primary", Hud._Anchor().Area.Left "," Hud._Anchor().Area.Top, primary.Left "," primary.Top)
        SearchWindow.LastArea := savedArea, SearchWindow.HiddenAt := savedAt, AppSettings.Appearance["ShowOn"] := savedShowOn

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

    ; 录完热键: 焦点移到下一个控件但不全选里面的文字 (下面是窗口列表时, 接着打字会整个换掉);
    ; 录制时框下面的提示跟着窗口移动
    static HotkeyBoxFocus() {
        eq := (n, a, e) => TestRunner.Equal("HotkeyBoxFocus." n, a, e)
        g := Gui(, "ALTRun hotkey box test")
        box := HotkeyBox.Add(g, "w160", "^!k")
        list := g.AddEdit("w300 r3 +Multi", "ahk_exe a.exe`r`nahk_exe b.exe")
        g.Show("x100 y100")
        try {
            WinActivate("ahk_id " g.Hwnd)
            if !WinWaitActive("ahk_id " g.Hwnd, , 2) {
                FileAppend("NOTE HotkeyBoxFocus: test window not active, skipped`n", "*")
                return
            }
            state := HotkeyBox.Boxes[box.Hwnd]
            box.Focus()                                                     ; 点进框里: 开始录制
            Sleep(50)
            eq("recording", HotkeyBox._active == state, true)
            eq("tip shown", state.Tip != "", true)
            before := state.TipAt
            WinMove(300, 250, , , "ahk_id " g.Hwnd)
            HotkeyBox._Watch()
            eq("tip follows the window", state.TipAt != before, true)
            HotkeyBox._Finish(state, "^d")
            Sleep(50)
            eq("value", HotkeyBox.Value(box), "^d")
            eq("hotkeys back on", A_IsSuspended, 0)
            eq("focus on the next control", DllCall("GetFocus", "Ptr") = list.Hwnd, true)
            start := 0, end := 0
            DllCall("SendMessage", "Ptr", list.Hwnd, "UInt", 0xB0, "UInt*", &start, "UInt*", &end)   ; EM_GETSEL
            eq("next control's text not selected", end - start, 0)
            eq("tip cleared", state.Tip, "")
        } finally {
            HotkeyBox.CancelActive()
            g.Destroy()
            HotkeyBox.Prune()
        }
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
        locked := FileOpen(jsonFile, "r -rwd")                              ; 同步盘 / 杀毒软件暂时占着文件: 等它放开后写入
        SetTimer(() => locked.Close(), -300)
        JSON.WriteFile(jsonFile, Map("k", "v3"))
        eq("write file: retries while locked", JSON.Parse(FileRead(jsonFile, "UTF-8"))["k"], "v3")
        locked := FileOpen(jsonFile, "r -rwd"), saved := JSON.MoveTries, JSON.MoveTries := 2
        failed := false
        try JSON.WriteFile(jsonFile, Map("k", "v4"))
        catch
            failed := true
        locked.Close(), JSON.MoveTries := saved
        if failed                                                           ; Wine 不一定支持独占打开, 这时直接写成功
            eq("write file: gives up, original kept, no tmp left", JSON.Parse(FileRead(jsonFile, "UTF-8"))["k"] "|" (FileExist(jsonFile ".tmp") ? "tmp" : ""), "v3|")
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
        ok("currency off hint", results.Length = 1 && results[1].Valid && IsObject(results[1].OnRun) && InStr(results[1].Subtitle, "Preferences"))
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
    ; 表格文字: 从 Excel 复制的 (Tab 分隔, 单元格里有换行时加引号)、CSV、欧洲 Excel 的 ; 分隔
    static DelimitedText() {
        eq := (n, a, e) => TestRunner.Equal("DelimitedText." n, a, e)
        joined(rows) {
            text := ""
            for row in rows {
                for index, cell in row
                    text .= (index > 1 ? "|" : "") cell
                text .= "/"
            }
            return text
        }
        eq("excel paste", joined(DelimitedText.Parse("Name`tKey`tText`r`nMail`tem`t" '"a@b.com' "`r`n" 'Thanks"' "`r`n")), "Name|Key|Text/Mail|em|a@b.com`nThanks/")
        eq("csv quotes", joined(DelimitedText.Parse('a,"b,c","say ""hi"""' "`n" "d,,e")), 'a|b,c|say "hi"/d||e/')
        eq("semicolon", joined(DelimitedText.Parse("a;b,c;d`n1;2;3")), "a|b,c|d/1|2|3/")
        eq("blank lines skipped", joined(DelimitedText.Parse("`n`na`tb`n`t`n")), "a|b/")
        eq("quote inside a field", joined(DelimitedText.Parse('5" pipe,ok')), '5" pipe|ok/')
        eq("delimiter in quotes ignored", DelimitedText.Delimiter('"a`tb",c,d'), ",")
        eq("tab wins", DelimitedText.Delimiter("a`tb,c`td"), "`t")
        rows := [["Name", "Text"], ["Mail", "line1`r`nline2"], ["Quote", 'say "hi", ok'], [" pad", "1"]]
        csv := DelimitedText.ToCsv(rows)
        eq("to csv", csv, "Name,Text`r`nMail," '"line1' "`r`n" 'line2"' "`r`nQuote," '"say ""hi"", ok"' "`r`n" '" pad",1' "`r`n")
        eq("round trip", joined(DelimitedText.Parse(csv)), "Name|Text/Mail|line1`nline2/Quote|say " '"hi", ok/ pad|1/')
        path := A_Temp "\ALTRunDelimited.csv"
        for encoding in ["UTF-8", "UTF-8-RAW"] {                            ; 带 BOM / 不带 BOM
            try FileDelete(path)
            FileAppend("名称,正文`r`n邮箱,谢谢", path, encoding)
            eq("read " encoding, DelimitedText.ReadFile(path), "名称,正文`r`n邮箱,谢谢")
        }
        try FileDelete(path)
        FileAppend("名称,正文", path, "CP936")                              ; GBK (中文 Excel 默认保存的 CSV): 按系统编码读 (CI 不是中文系统, 只比较两种读法一样)
        eq("read system code page", DelimitedText.ReadFile(path), FileRead(path, "CP0"))
        try FileDelete(path)
    }

    ; 文字片段的导入 / 导出: 合并规则、表头、自动展开
    static SnippetImport() {
        eq := (n, a, e) => TestRunner.Equal("SnippetImport." n, a, e)
        snippets := [Map("Name", "Mail", "Keyword", "em", "Text", "old", "AutoExpand", 1)
            , Map("Name", "Address", "Keyword", "", "Text", "old address", "AutoExpand", 1)]
        rows := [["名称", "关键字", "正文", "自动展开"]
            , ["Mail (work)", "EM", "new@b.com", ""]                       ; 关键字相同 (不分大小写): 更新
            , ["Address", "", "new address"]                               ; 没有关键字, 名称相同: 更新
            , ["", "sig", "Best regards`nCaomao", "0"]                     ; 没有名称: 用关键字; 0 = 不自动展开
            , ["", "", "  Line one  `nLine two"]                           ; 没有名称和关键字: 用正文第一行
            , ["Spaced", "two words", "x"]                                 ; 关键字有空格 (自动展开用不了): 自动展开关掉
            , ["Empty", "e", "   "]]                                       ; 没有正文: 跳过
        result := SnippetProvider.ImportRows(rows, snippets, ";")
        eq("counts", result.Added "/" result.Updated "/" result.Skipped, "3/2/1")
        eq("updated by keyword", snippets[1]["Name"] "|" snippets[1]["Text"] "|" snippets[1]["Keyword"], "Mail (work)|new@b.com|EM")
        eq("updated by name", snippets[2]["Text"], "new address")
        eq("name from keyword", snippets[3]["Name"] "|" snippets[3]["AutoExpand"], "sig|0")
        eq("name from text", snippets[4]["Name"], "Line one")
        eq("keyword with a space", snippets[5]["AutoExpand"], 0)
        eq("total", snippets.Length, 5)
        eq("header only on the first row", SnippetProvider.ImportRows([["Mail", "m", "x"], ["Name", "k", "Text"]], [], ";").Added, 2)
        eq("header english", SnippetProvider.IsHeaderRow(["Name", "Keyword", "Text"]), true)
        eq("header japanese", SnippetProvider.IsHeaderRow(["名前", "キーワード", "本文", "自動"]), true)
        eq("not header", SnippetProvider.IsHeaderRow(["Mail", "em", "a@b.com"]), false)
        exported := SnippetProvider.ExportRows(snippets)
        eq("export header", SnippetProvider.IsHeaderRow(exported[1]), true)
        eq("export rows", exported.Length, 6)
        eq("export auto expand", exported[4][4], 0)
        again := []                                                         ; 导出再导入: 内容不变
        SnippetProvider.ImportRows(DelimitedText.Parse(DelimitedText.ToCsv(exported)), again, ";")
        same := again.Length = snippets.Length
        for index, snippet in snippets
            for key in ["Name", "Keyword", "Text", "AutoExpand"]
                same := same && again[index][key] == snippet[key]
        eq("round trip", same, true)
    }

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
        eq("too long not registered", SnippetExpander.Abbreviations(Map("Keyword", longKeyword, "Text", "x"), ";").Length, 0)
        eq("longest registered", Tests._Join(SnippetExpander.Abbreviations(Map("Keyword", StrRepeat("a", 39), "Text", "x"), ";")), ";" StrRepeat("a", 39))
        eq("several keywords ok", SnippetProvider.CheckKeyword("sig, ㄴ", Map("AutoExpand", 1), ";"), "")
        ok("several: one too long", SnippetProvider.CheckKeyword("sig, " longKeyword, Map("AutoExpand", 1), ";") != "")
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
        ; 剪贴板历史里的文字条目: → 菜单里也有搜索操作 (图片、文件没有)
        clipItem := ResultItem("hello", "", {Kind: "text", Arg: "hello world", Provider: ClipboardProvider.Id, Source: Map("Text", "hello world", "Time", A_Now)})
        ok("clipboard history: search engines", InStr(titles(clipItem), "Search Google|"))
        clipItem.Source := Map("Type", "image", "Time", A_Now)
        ok("clipboard history: image has no search", !InStr(titles(clipItem), "Search Google|"))
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
        ; 日志和 .txt 输出用默认程序打开, 打不开 (没有关联) 时用记事本
        calls := []
        Path.OpenText("C:\Temp\a.log", (cmd) => calls.Push(cmd))
        TestRunner.Equal("Misc.open text: default program", calls.Length ? calls[1] : "", '"C:\Temp\a.log"')
        calls := []
        failFirst(cmd) {
            calls.Push(cmd)
            if (calls.Length = 1)
                throw Error("no association")
        }
        Path.OpenText("C:\Temp\a.log", failFirst)
        TestRunner.Equal("Misc.open text: falls back to Notepad", calls.Length = 2 ? calls[2] : "", 'notepad.exe "C:\Temp\a.log"')
        TestRunner.True("Input.layout of this thread", Win.KeyboardLayout() != 0)
        TestRunner.True("Input.per-window setting is a flag", Win.PerWindowInputMethod() = true || Win.PerWindowInputMethod() = false)
        saved := AppSettings.General["SwitchToEnglishInput"]
        AppSettings.General["SwitchToEnglishInput"] := 0
        SearchWindow._layoutBefore := 0x08040804
        SearchWindow._RestoreInputLanguage()                                 ; 没开这个选项: 什么也不做, 只清掉记录
        TestRunner.Equal("Input.restore clears the saved layout", SearchWindow._layoutBefore, 0)
        SearchWindow._imeBefore := {Open: 1, Conv: 0x401}
        SearchWindow._RestoreInputLanguage()                                 ; 英文模式也一样: 没开时不恢复, 只清掉记录
        TestRunner.Equal("Input.restore clears the saved IME mode", SearchWindow._imeBefore, "")
        AppSettings.General["SwitchToEnglishInput"] := 1
        SearchWindow._imeBefore := {Open: 1, Conv: 0x401}
        SearchWindow._RestoreInputLanguage()                                 ; 选的是英语键盘: 不动中 / 英模式
        TestRunner.Equal("Input.keyboard choice leaves the IME mode alone", SearchWindow._imeBefore, "")
        savedPlaceholder := AppSettings.Appearance["Placeholder"]
        AppSettings.Appearance["Placeholder"] := ""
        TestRunner.Equal("Placeholder.default", SearchWindow.Placeholder(), "ALTRun Search")
        AppSettings.Appearance["Placeholder"] := " "
        TestRunner.Equal("Placeholder.only spaces: nothing", SearchWindow.Placeholder(), "")
        AppSettings.Appearance["Placeholder"] := " Type to search "
        TestRunner.Equal("Placeholder.custom (trimmed)", SearchWindow.Placeholder(), "Type to search")
        AppSettings.Appearance["Placeholder"] := savedPlaceholder
        ; 输入框按大写字母的高度居中 (#158): 150% 时实测基线在 42、大写字母高 33, 往下挪 7 像素
        TestRunner.Equal("InputTop.cap centered", SearchWindow.InputTop(64, 64, 21, 42, 33), 7)
        TestRunner.Equal("InputTop.at most half the padding", SearchWindow.InputTop(64, 64, 8, 42, 33), 4)
        TestRunner.Equal("InputTop.no metrics: centered", SearchWindow.InputTop(60, 40, 14, 0, 0), 10)
        TestRunner.Equal("InputTop.never above centered", SearchWindow.InputTop(60, 40, 14, 42, 10), 10)
        g := Gui()
        g.SetFont("s20", "Segoe UI")
        metrics := Win.CapMetrics(g.AddEdit("w200 r1").Hwnd)
        ; 搜索结果列表的实际行高 (最后一行不被切掉): 和 SearchWindow 一样用 48 像素高的图标列表撑开行高
        list := g.AddListView("w300 h0 -Hdr +0x2000", ["Result"])
        images := DllCall("comctl32\ImageList_Create", "Int", 1, "Int", 48, "UInt", 0x21, "Int", 1, "Int", 1, "Ptr")
        list.SetImageList(images, 1)
        rowH := Win.ListRowHeight(list, 0)
        TestRunner.True("ListRowHeight.measured (" rowH ")", rowH >= 48 && rowH <= 50)
        TestRunner.Equal("ListRowHeight.temporary row removed", list.GetCount(), 0)
        list.Add(), list.Add()
        TestRunner.Equal("ListRowHeight.existing rows kept", (Win.ListRowHeight(list, 0) = rowH) " " list.GetCount(), "1 2")
        g.Destroy()
        TestRunner.True("InputTop.font metrics (ascent " metrics.Ascent ", cap " metrics.Cap ")", metrics.Ascent > metrics.Cap && metrics.Cap > 10)
        TestRunner.Equal("Input.english choices", AppSettings.Defaults()["General"]["SwitchToEnglishInput"] " " I18n.T("Prefs.EnglishInput.Mode"), "0 English mode of the current input method")
        AppSettings.General["SwitchToEnglishInput"] := saved
        ; 恢复默认设置: 选项换成默认值, 自己加的内容 (命令、片段、热键、网页搜索、置顶、隐藏的系统命令) 保留
        current := AppSettings.Defaults()
        current["General"]["Hotkey"] := "^Space", current["Appearance"]["Theme"] := "Dark"
        current["CustomCommands"] := [Map("Title", "Mine")], current["Snippets"] := [Map("Keyword", "sig")]
        current["Features"]["WebSearch"]["Engines"] := [Map("Id", "mine")], current["Features"]["Recent"]["Pinned"] := ["x"]
        current["Features"]["System"]["Hidden"] := ["Lock"], current["Features"]["WebSearch"]["Fallbacks"] := []
        reset := AppSettings.ResetData(current)
        TestRunner.Equal("RestoreDefaults.options reset", reset["General"]["Hotkey"] " " reset["Appearance"]["Theme"] " " reset["Features"]["WebSearch"]["Fallbacks"].Length
            , AppSettings.Defaults()["General"]["Hotkey"] " " AppSettings.Defaults()["Appearance"]["Theme"] " " AppSettings.Defaults()["Features"]["WebSearch"]["Fallbacks"].Length)
        TestRunner.Equal("RestoreDefaults.content kept", reset["CustomCommands"][1]["Title"] " " reset["Snippets"][1]["Keyword"] " " reset["Features"]["WebSearch"]["Engines"][1]["Id"]
            " " reset["Features"]["Recent"]["Pinned"][1] " " reset["Features"]["System"]["Hidden"][1], "Mine sig mine x Lock")
        TestRunner.Equal("RestoreDefaults.no data location", AppSettings.ResetData(Map()).Has("DataLocation") ? 1 : 0, 0)
        TestRunner.Equal("RestoreDefaults.data location kept", AppSettings.ResetData(Map("DataLocation", "D:\Sync"))["DataLocation"], "D:\Sync")
        ; 系统信息: 有版本号, 不含数据文件夹的路径
        info := App.SystemInfo()
        TestRunner.True("SystemInfo.version", InStr(info, App.Name " " App.Version " (") = 1)
        TestRunner.True("SystemInfo.no paths", !InStr(info, AppSettings.DataDir) && !InStr(info, A_ScriptDir))
        TestRunner.Equal("SystemInfo.lines", StrSplit(Trim(info, "`n"), "`n").Length, 9)
        ; 以管理员身份运行 (#158): 新启动的是管理员、正在运行的不是时替换它, 其它情况照常把参数交过去
        TestRunner.Equal("Admin.takes over", (App.TakesOver(true, false) ? 1 : 0) (App.TakesOver(true, true) ? 1 : 0) (App.TakesOver(false, false) ? 1 : 0) (App.TakesOver(false, true) ? 1 : 0), "1000")
        TestRunner.Equal("Admin.this process", Win.IsElevated(ProcessExist()) ? 1 : 0, A_IsAdmin ? 1 : 0)
        TestRunner.Equal("Admin.no such process", Win.IsElevated(0) ? 1 : 0, 0)
        TestRunner.Equal("Admin.suffix", App.AdminSuffix(), A_IsAdmin ? " (Administrator)" : "")
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

class FakeHttpRequest {                                                    ; HttpGet 测试用的假请求 (Msxml2.XMLHTTP 用到的部分)
    readyState := 1, status := 0, statusText := "", responseText := ""
    Method := "", Url := "", Async := "", Header := "", Aborted := false
    open(method, url, async) => (this.Method := method, this.Url := url, this.Async := async)
    setRequestHeader(name, value) => (this.Header := name)
    send() => ""
    abort() => (this.Aborted := true)
}

JsonReadWrite_Parses(text) {
    try {
        JSON.Parse(text)
        return true
    }
    return false
}

StrRepeat(text, count) => StrReplace(Format("{:" count "}", ""), " ", text)
