;===============================================================================
; CoreTests.ahk - Search core, settings and shared libraries: matching, queries, settings and
;   their migration, calculator, text and list tools, JSON, logging, data locations and app
;   startup logic
;-------------------------------------------------------------------------------
; Included by Tests\RunTests.ahk. Every static method whose name doesn't start with "_" is a test;
; TestRunner runs them in name order, and each one must put back the shared state it changes.
;===============================================================================
class CoreTests {
    ; WinTimer: Windows timers instead of SetTimer for timers that stay set while idle (AutoHotkey checks its own every 10 ms while any is set)
    static WinTimers() {
        eq := (n, a, e) => TestRunner.Equal("WinTimers." n, a, e)
        counts := {Every: 0, Once: 0}
        every := () => counts.Every++, once := () => counts.Once++
        before := WinTimer.Count
        WinTimer.Set(every, 50), WinTimer.Set(once, -50)
        eq("set", WinTimer.Count - before, 2)
        Sleep(400)
        WinTimer.Set(every, 0)
        ran := counts.Every
        TestRunner.True("WinTimers.repeats (" ran ")", ran >= 3)
        eq("once", counts.Once, 1)
        eq("stopped and done ones removed", WinTimer.Count, before)
        Sleep(200)
        eq("stopped", counts.Every, ran)
        WinTimer.Set(once, -50), WinTimer.Set(once, -60000), Sleep(200)  ; setting it again restarts it with the new time
        eq("restarted", counts.Once, 1)
        WinTimer.Set(once, 0), WinTimer.Set(() => 0, 0)
        eq("stop", WinTimer.Count, before)
        ; Not run again while it is still running (like SetTimer): a slow function with a short period runs one at a time
        state := {Now: 0, Most: 0, Runs: 0}
        slow := () => (state.Now++, state.Most := Max(state.Most, state.Now), state.Runs++, Sleep(200), state.Now--)
        WinTimer.Set(slow, 30), Sleep(700), WinTimer.Set(slow, 0), Sleep(300)
        eq("no re-entry", state.Most, 1)
        TestRunner.True("WinTimers.slow one still repeats (" state.Runs ")", state.Runs >= 2)
        eq("tolerance: a tenth, at most 1 s", WinTimer.Tolerance(2000) " " WinTimer.Tolerance(-60000) " " WinTimer.Tolerance(50), "200 1000 5")
        hidden := A_DetectHiddenWindows
        DetectHiddenWindows(true)
        eq("message-only window", WinExist("ALTRun timers ahk_pid " ProcessExist()), 0)
        DetectHiddenWindows(hidden)
    }

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
        ; the start of a keyword not typed in full yet: at least 3 letters, no space, a longer keyword (case-insensitive)
        keyword := ""
        TestRunner.True("SearchQuery.starts keyword", SearchQuery("Scrip").StartsKeyword(["open", "script"], &keyword) && keyword = "script")
        TestRunner.True("SearchQuery.starts keyword: 3 letters", SearchQuery("scr").StartsKeyword(["script"], &keyword))
        TestRunner.True("SearchQuery.starts keyword: too short", !SearchQuery("sc").StartsKeyword(["script"], &keyword))
        TestRunner.True("SearchQuery.starts keyword: typed in full", !SearchQuery("script").StartsKeyword(["script"], &keyword))
        TestRunner.True("SearchQuery.starts keyword: after a space", !SearchQuery("scr ").StartsKeyword(["script"], &keyword))
        TestRunner.True("SearchQuery.starts keyword: not the start", !SearchQuery("ript").StartsKeyword(["script"], &keyword))
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
        eq("2.x jump hotkey -> Hotkey", data["Extensions"]["QuickSwitch"]["Hotkey"] "|" data["Extensions"]["AutoDate"]["Hotkey"], "^g|^d")
        eq("2.x dialogs: standard ones, as arrays", data["Extensions"]["QuickSwitch"]["StandardDialogs"] "|" data["Extensions"]["QuickSwitch"]["DialogWindows"].Length
            "|" (data["Extensions"]["AutoDate"]["RenameWindows"] is Array), "1|0|1")
        eq("2.x old hotkeys not kept", data["Extensions"]["QuickSwitch"].Has("ExplorerHotkey") || data["Extensions"]["QuickSwitch"].Has("TotalCmdHotkey"), false)
        eq("pttools", data["Extensions"]["PTTools"]["RebarSpanWidth"], 4000)
        eq("hotkeys", data["Hotkeys"].Length, 1)
        eq("hotkey action", data["Hotkeys"][1]["Action"], "PTTools")
        eq("commands", data["CustomCommands"].Length, 4)                    ; Func isn't migrated
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
        ; 4 -> 5: snippet paste mode Clipboard (the old default) -> Auto; Type stays
        v5 := SchemaMigration.Upgrade(Map("SchemaVersion", 4, "Features", Map("Snippets", Map("PasteMode", "Clipboard", "PasteDelay", 500))), 4)
        eq("4->5 clipboard -> auto", v5["Features"]["Snippets"]["PasteMode"], "Auto")
        eq("4->5 keeps other settings", v5["Features"]["Snippets"]["PasteDelay"], 500)
        eq("4->5 on to the current version", v5["SchemaVersion"], AppSettings.CurrentVersion)
        eq("4->5 type stays", SchemaMigration.Upgrade(Map("SchemaVersion", 4, "Features", Map("Snippets", Map("PasteMode", "Type"))), 4)["Features"]["Snippets"]["PasteMode"], "Type")
        eq("4->5 without section", InStr(JSON.Stringify(SchemaMigration.Upgrade(Map("SchemaVersion", 4), 4)), "PasteMode"), 0)
        ; 4 -> 5: windows / keywords added to the defaults later are added for old users too (only what is missing; the user's own stay; emptied lists are left alone); small cleanups
        defaults := AppSettings.Defaults()
        old := Map("SchemaVersion", 4
            , "Features", Map("FileSearch", Map("Keywords", ["open", "find"]), "Calculator", Map("BarSizes", ["13", "16", " 20 ", 25, "abc"]))
            , "Extensions", Map("AutoDate", Map("RenameWindows", "ahk_class CabinetWClass,ahk_class Progman,ahk_class WorkerW,ahk_class #32770,ahk_class TTOTAL_CMD"
                                              , "AppendWindows", "ahk_class TCmtEditForm,ahk_exe Notepad4.exe")
                              , "QuickSwitch", Map("ExcludeWindows", "ahk_class SysListView32, ahk_exe Explorer.exe"))
            , "Hotkeys", [Map("Key", "~MButton", "Action", "PTTools", "WinTitle", "ahk_exe RAPTW.exe"), Map("Key", "F1", "Action", "About", "WinTitle", "ALTRun")])
        v5 := SchemaMigration.Upgrade(old, 4)
        eq("4->5 keywords: file added", TextTools.JoinList(v5["Features"]["FileSearch"]["Keywords"], ","), "open,find,file")
        eq("4->5 keywords = new default", TextTools.JoinList(v5["Features"]["FileSearch"]["Keywords"], ","), TextTools.JoinList(defaults["Features"]["FileSearch"]["Keywords"], ","))
        eq("4->5 rename windows: TC dialogs added", TextTools.JoinList(v5["Extensions"]["AutoDate"]["RenameWindows"], ",")
            , "ahk_class CabinetWClass,ahk_class Progman,ahk_class WorkerW,ahk_class #32770,ahk_class TTOTAL_CMD,ahk_class TSTDTREEDLG,ahk_class TCOMBOINPUT")
        eq("4->5 rename windows = new default", TextTools.JoinList(v5["Extensions"]["AutoDate"]["RenameWindows"], ","), TextTools.JoinList(defaults["Extensions"]["AutoDate"]["RenameWindows"], ","))
        eq("4->5 append windows = new default", TextTools.JoinList(v5["Extensions"]["AutoDate"]["AppendWindows"], ","), TextTools.JoinList(defaults["Extensions"]["AutoDate"]["AppendWindows"], ","))
        eq("4->5 exclude windows = new default", TextTools.JoinList(v5["Extensions"]["QuickSwitch"]["ExcludeWindows"], ","), TextTools.JoinList(defaults["Extensions"]["QuickSwitch"]["ExcludeWindows"], ","))
        eq("4->5 old F1 hotkey removed, others kept", v5["Hotkeys"].Length "|" v5["Hotkeys"][1]["Key"], "1|~MButton")
        sizes := v5["Features"]["Calculator"]["BarSizes"]
        eq("4->5 bar sizes as numbers", (sizes[1] is Integer) "|" sizes[1] "|" sizes[3] "|" (sizes[3] is Integer) "|" sizes[4] "|" sizes[5], "1|13|20|1|25|abc")
        custom := Map("SchemaVersion", 4
            , "Features", Map("FileSearch", Map("Keywords", ["open", "FILE"])), "Extensions", Map("AutoDate", Map("RenameWindows", "ahk_exe foo.exe, AHK_CLASS tstdtreedlg"
                , "AppendWindows", "ahk_exe bar.exe"), "QuickSwitch", Map("ExcludeWindows", "")))
        v5 := SchemaMigration.Upgrade(custom, 4)
        eq("4->5 keywords: no duplicate file", TextTools.JoinList(v5["Features"]["FileSearch"]["Keywords"], ","), "open,FILE")
        eq("4->5 custom windows kept, only missing added", TextTools.JoinList(v5["Extensions"]["AutoDate"]["RenameWindows"], ","), "ahk_exe foo.exe,AHK_CLASS tstdtreedlg,ahk_class TCOMBOINPUT")
        eq("4->5 custom append windows untouched", TextTools.JoinList(v5["Extensions"]["AutoDate"]["AppendWindows"], ","), "ahk_exe bar.exe")
        eq("4->5 emptied list stays empty", TextTools.JoinList(v5["Extensions"]["QuickSwitch"]["ExcludeWindows"], ","), "")
        eq("4->5 missing sections not created", SchemaMigration.Upgrade(Map("SchemaVersion", 4), 4).Count, 1)
        ; 5 -> 6: one format for lists (arrays), StandardDialogs, hotkeys renamed, unused keys removed (#176)
        v5 := Map("SchemaVersion", 5
            , "Appearance", Map("Theme", "DarkCompact", "Density", "Comfortable")
            , "Features", Map("Snippets", Map("ExpandExclude", "ahk_exe mstsc.exe,  ahk_exe KeePass.exe ,")
                , "Applications", Map("RefreshMinutes", 120, "Depth", 3)
                , "FileSearch", Map("ScopeExclude", SchemaMigration.OldScopeExclude))
            , "Extensions", Map("QuickSwitch", Map("TotalCmdHotkey", "^!g", "ExplorerHotkey", "^e", "DialogWindows", "ahk_class #32770, ahk_class Qt5QWindowIcon"
                    , "ExcludeWindows", "ahk_exe 7zG.exe", "AutoSwitchExclude", "")
                , "AutoDate", Map("RenameHotkey", "^!d", "AppendHotkey", "^d", "RenameWindows", "ahk_class CabinetWClass", "AppendWindows", "ahk_exe a.exe, ahk_exe b.exe")))
        v6 := SchemaMigration.Upgrade(v5, 5)
        quick := v6["Extensions"]["QuickSwitch"], dates := v6["Extensions"]["AutoDate"]
        eq("5->6 version", v6["SchemaVersion"], 6)
        eq("5->6 jump hotkey renamed", quick["Hotkey"] "|" quick.Has("TotalCmdHotkey") "|" quick.Has("ExplorerHotkey"), "^!g|0|0")
        eq("5->6 date hotkey renamed", dates["Hotkey"] "|" dates.Has("RenameHotkey") "|" dates.Has("AppendHotkey"), "^!d|0|0")
        eq("5->6 standard dialogs split out", quick["StandardDialogs"] "|" TextTools.JoinList(quick["DialogWindows"], ","), "1|ahk_class Qt5QWindowIcon")
        eq("5->6 window lists as arrays", TextTools.JoinList(quick["ExcludeWindows"], ",") "|" quick["AutoSwitchExclude"].Length "|" TextTools.JoinList(dates["RenameWindows"], ",") "|" TextTools.JoinList(dates["AppendWindows"], ",")
            , "ahk_exe 7zG.exe|0|ahk_class CabinetWClass|ahk_exe a.exe,ahk_exe b.exe")
        eq("5->6 snippet exclusions as an array", TextTools.JoinList(v6["Features"]["Snippets"]["ExpandExclude"], ","), "ahk_exe mstsc.exe,ahk_exe KeePass.exe")
        eq("5->6 unused refresh removed", v6["Features"]["Applications"].Has("RefreshMinutes") "|" v6["Features"]["Applications"]["Depth"], "0|3")
        eq("5->6 old scope exclude -> new default", v6["Features"]["FileSearch"]["ScopeExclude"], defaults["Features"]["FileSearch"]["ScopeExclude"])
        eq("5->6 compact theme -> theme at normal density", v6["Appearance"]["Theme"] "|" v6["Appearance"]["Density"], "Dark|Normal")
        custom := Map("SchemaVersion", 5, "Appearance", Map("Theme", "LightCompact"), "Features", Map("FileSearch", Map("ScopeExclude", "i)\(build)(\|$)"))
            , "Extensions", Map("QuickSwitch", Map("Hotkey", "^h", "TotalCmdHotkey", "^g", "DialogWindows", "ahk_class Qt5QWindowIcon", "ExcludeWindows", ["ahk_exe x.exe"])))
        v6 := SchemaMigration.Upgrade(custom, 5)
        quick := v6["Extensions"]["QuickSwitch"]
        eq("5->6 new name already there wins", quick["Hotkey"] "|" quick.Has("TotalCmdHotkey"), "^h|0")
        eq("5->6 no standard dialogs", quick["StandardDialogs"] "|" TextTools.JoinList(quick["DialogWindows"], ","), "0|ahk_class Qt5QWindowIcon")
        eq("5->6 arrays left alone", TextTools.JoinList(quick["ExcludeWindows"], ","), "ahk_exe x.exe")
        eq("5->6 custom scope exclude kept", v6["Features"]["FileSearch"]["ScopeExclude"], "i)\(build)(\|$)")
        eq("5->6 compact theme without density", v6["Appearance"]["Theme"] "|" v6["Appearance"]["Density"], "Light|Normal")
        eq("5->6 other theme left alone", SchemaMigration.Upgrade(Map("SchemaVersion", 5, "Appearance", Map("Theme", "Midnight", "Density", "Comfortable")), 5)["Appearance"]["Density"], "Comfortable")
        eq("5->6 missing sections not created", SchemaMigration.Upgrade(Map("SchemaVersion", 5), 5).Count, 1)
        for setting in ["Extensions.QuickSwitch.DialogWindows", "Extensions.QuickSwitch.ExcludeWindows", "Extensions.QuickSwitch.AutoSwitchExclude"
                   , "Extensions.AutoDate.RenameWindows", "Extensions.AutoDate.AppendWindows", "Features.Snippets.ExpandExclude"]
            eq("defaults: " setting " is an array", PreferencesWindow.GetPath(defaults, setting) is Array, true)
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
        ; a few of many (one pass keeping a short list): the same as sorting them all, ties by lower index
        many := Map(), expected := []
        Loop 500
            many[A_Index * 2] := Mod(A_Index * 37, 11) + (Mod(A_Index, 3) = 0 ? 0.5 : 0)
        for pair in StrSplit(Sort(TopListAll(many), "R"), "`n")
            expected.Push(9999999 - Integer(StrSplit(pair, "`t")[2]))
        got := FuzzyMatcher.TopIndexes(many, 10), same := ""
        Loop 10
            same .= (got[A_Index] = expected[A_Index]) ? "" : A_Index " "
        TestRunner.Equal("TopIndexes.few of many like sorting all", got.Length "|" same, "10|")
        TopListAll(scores) {
            lines := ""
            for index, score in scores
                lines .= Format("{:09.4f}", score) "`t" Format("{:07}", 9999999 - index) "`n"
            return RTrim(lines, "`n")
        }
    }

    static Sorting() {
        items := [ResultItem("low", "", {Score: 10}), ResultItem("high", "", {Score: 90}), ResultItem("mid1", "", {Score: 50}), ResultItem("mid2", "", {Score: 50})]
        sorted := ProviderRegistry.SortByScore(items)
        order := ""
        for item in sorted
            order .= item.Title " "
        TestRunner.Equal("Sorting.order", Trim(order), "high mid1 mid2 low")
    }

    ; Background work waits while the search window says it is busy; with nothing set (tests) it never waits
    static BackgroundGate() {
        eq := (n, a, e) => TestRunner.Equal("BackgroundGate." n, a, e)
        saved := BackgroundWork.IsBusy
        BackgroundWork.IsBusy := ""
        eq("nothing set", BackgroundWork.CanRun(), true)
        BackgroundWork.IsBusy := () => true
        eq("busy", BackgroundWork.CanRun(), false)
        BackgroundWork.IsBusy := () => false
        eq("idle", BackgroundWork.CanRun(), true)
        BackgroundWork.IsBusy := saved
    }

    static Knowledge() {
        Knowledge.Picks := Map(), Knowledge.QueryPicks := Map(), Knowledge.History := []
        Knowledge._saveTimer := () => 0                                     ; no disk writes in tests
        Knowledge.Record("no", "app:notepad")
        Knowledge.Record("no", "app:notepad")
        TestRunner.True("Knowledge.boost exact", Knowledge.Boost("no", "app:notepad") > Knowledge.Boost("no", "app:other"))
        TestRunner.True("Knowledge.boost prefix", Knowledge.Boost("not", "app:notepad") > 20)
        TestRunner.Equal("Knowledge.history", Knowledge.History[1], "no")
        TestRunner.Equal("Knowledge.history dedupe", Knowledge.History.Length, 1)
    }

    ; Every setting the code reads exists in AppSettings.Defaults() (a typo in a key name would only show at run time).
    ; Reads guarded by .Has("Key") are optional and may be missing. SchemaMigration reads old names on purpose
    static SettingsKeys() {
        defaults := AppSettings.Defaults(), missing := "", count := 0
        Loop Files, A_ScriptDir "\..\Src\*.ahk", "R" {
            if (A_LoopFileName = "SchemaMigration.ahk" || A_LoopFileName = "AppSettings.ahk")
                continue
            text := RegExReplace(FileRead(A_LoopFileFullPath, "UTF-8"), "m)(^|\s);[^\n]*")      ; comments out
            for setting in CoreTests._SettingPaths(text) {
                count++
                parts := StrSplit(setting, ".")
                if !CoreTests._HasSetting(defaults, parts) && !InStr(text, '.Has("' parts[parts.Length] '")')
                    missing .= setting " (" A_LoopFileName ") "
            }
        }
        TestRunner.True("SettingsKeys.found reads (" count ")", count > 100)
        TestRunner.Equal("SettingsKeys.every read key has a default", missing, "")
        TestRunner.True("SettingsKeys.a typo is caught", !CoreTests._HasSetting(defaults, ["Features", "FileSearch", "Keyword"]))
    }

    ; "Features.FileSearch.Keywords"-style paths read in one source file: AppSettings.Feature("X")["Y"], .General["Y"],
    ; a variable holding AppSettings.Feature("X") then var["Y"] in the same method, base := "Extensions.X." + base "Y",
    ; and quoted paths (Preferences binds those)
    static _SettingPaths(text) {
        paths := Map()
        for m in CoreTests._Matches(text, 'AppSettings\.(Feature|Extension)\("(\w+)"\)\["(\w+)"\]')
            paths[CoreTests._Section(m[1]) "." m[2] "." m[3]] := true
        for m in CoreTests._Matches(text, 'AppSettings\.(General|Appearance)\["(\w+)"\]')
            paths[m[1] "." m[2]] := true
        for m in CoreTests._Matches(text, '"((?:Features|Extensions)\.\w+\.\w+|(?:General|Appearance)\.\w+)"')
            paths[m[1]] := true
        for method in StrSplit(text, "static ") {
            for m in CoreTests._Matches(method, '(\w+) := AppSettings\.(Feature|Extension)\("(\w+)"\)')
                for key in CoreTests._Matches(method, '\b' m[1] '\["(\w+)"\]')
                    paths[CoreTests._Section(m[2]) "." m[3] "." key[1]] := true
            if RegExMatch(method, 'base := "((?:Features|Extensions)\.\w+)\."', &b)
                for key in CoreTests._Matches(method, 'base "(\w+)"')
                    paths[b[1] "." key[1]] := true
        }
        list := []
        for settingPath in paths
            list.Push(settingPath)
        return list
    }

    static _Section(kind) => (kind = "Feature") ? "Features" : "Extensions"

    static _Matches(haystack, pattern) {
        found := [], pos := 1
        while (pos := RegExMatch(haystack, pattern, &m, pos))
            found.Push(m), pos += m.Len
        return found
    }

    static _HasSetting(data, parts) {
        for part in parts {
            if !(data is Map) || !data.Has(part)
                return false
            data := data[part]
        }
        return true
    }

    ; Double-tap Ctrl / Shift: only two separate short presses count
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

    ; Several keywords in the keyword field (comma-separated): web searches, custom commands, snippets work with any of them; hints and completion use the one typed (the first if none)
    static MultipleKeywords() {
        eq := (n, a, e) => TestRunner.Equal("MultipleKeywords." n, a, e)
        eq("split", TextTools.JoinList(SearchQuery.Keywords(" g , ㅎ,,google，구글 "), ","), "g,ㅎ,google,구글")
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
        ; "sni": the snip keyword's hint at the end of the list; "snip" alone lists the snippets (no hint)
        hint := SnippetProvider.Search(SearchQuery("sni"))
        eq("snippet keyword start: hint", hint.Length "|" hint[1].Valid "|" hint[1].AutoComplete "|" hint[1].Score, "1|0|snip |" ResultItem.KeywordPrefixHintScore)
        listed := SnippetProvider.Search(SearchQuery("snip"))
        eq("snippet keyword alone: the snippets", listed.Length "|" listed[1].Title, "1|Signature")
        AppSettings.Data["Snippets"] := savedSnippets
    }

    ; The hint row for a built-in keyword: can't run, completes "keyword ", typed in full / only the start typed decide the score
    static KeywordHint() {
        eq := (n, a, e) => TestRunner.Equal("KeywordHint." n, a, e)
        hint := ResultItem.KeywordHint("Find scripts by name", "script", "res:imageres.dll,-5323")
        eq("typed in full", hint.Valid "|" hint.AutoComplete "|" hint.Subtitle "|" hint.Score, "0|script |script ...|" ResultItem.KeywordHintScore)
        eq("only the start typed", ResultItem.KeywordHint("t", "script", "", true).Score, ResultItem.KeywordPrefixHintScore)
        eq("own subtitle", ResultItem.KeywordHint("t", "open", "", false, "'... / open ...").Subtitle, "'... / open ...")
        eq("no keyword: nothing to complete", ResultItem.KeywordHint("t", "", "").AutoComplete, "")
        eq("the start after every match", ResultItem.KeywordPrefixHintScore < ResultItem.KeywordHintScore, true)
    }

    ; Data folder: "DataLocation" in ALTRun.json at the default location (environment variables, relative paths); setting / back to default
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
        AppSettings.Portable := false, AppSettings.UserDir := root "\user"   ; default location = root\user\Data
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

        AppSettings.DataDir := "E:\Elsewhere"                                ; settings from elsewhere in use: only the file at the default location changes
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

    ; Calculator: number bases and date arithmetic
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

    ; Only one ALTRun: a second start hands the command-line arguments to the running one (WM_COPYDATA); a restart (/restart) replaces it
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
            copyData := Buffer(A_PtrSize * 3, 0)                            ; WM_COPYDATA of other programs (e.g. Everything) isn't handled
            NumPut("UPtr", 0x1234, copyData, 0)
            eq("other copydata ignored", App._OnCopyData(0, copyData.Ptr, 0x4A, A_ScriptHwnd), "")
        } finally {
            OnMessage(0x4A, onCopyData, 0)
            App.DefineProp("_HandleCommandLine", saved)
        }
    }

    ; an icon being read: when the UI asks again it isn't queued twice (slow icons used to be read twice)
    ; Background downloads: asynchronous requests; on completion / error / timeout the callback runs in its own thread (fake requests, no network)
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
            eq("done", TestUtil.WaitFor(() => got.Length) ? got[1] : "", "done hello")
            eq("nothing pending", HttpGet.Pending, 0)

            got.Length := 0
            HttpGet.Start("https://example.test/missing", (t) => got.Push("done " t), (m) => got.Push("fail " m))
            fakes[2].status := 404, fakes[2].statusText := "Not Found", fakes[2].readyState := 4
            eq("http error", TestUtil.WaitFor(() => got.Length) ? got[1] : "", "fail HTTP 404 Not Found")

            got.Length := 0
            HttpGet.Start("https://example.test/offline", (t) => got.Push("done " t), (m) => got.Push("fail " m))
            fakes[3].status := 12007, fakes[3].readyState := 4
            eq("network error", TestUtil.WaitFor(() => got.Length) ? got[1] : "", "fail Network error 12007")

            got.Length := 0
            HttpGet.Start("https://example.test/bad", (t) => (got.Push("parsing"), JSON.Parse(t)), (m) => got.Push("fail"))
            fakes[4].status := 200, fakes[4].responseText := "{not json", fakes[4].readyState := 4
            TestUtil.WaitFor(() => got.Length >= 2)
            eq("done throws -> fail", got.Length = 2 ? got[1] " " got[2] : got.Length, "parsing fail")

            got.Length := 0
            HttpGet.Start("https://example.test/slow", (t) => got.Push("done"), (m) => got.Push("fail " m), 50)
            eq("timeout", TestUtil.WaitFor(() => got.Length) ? got[1] : "", "fail Timed out after 0 s")
            eq("timeout aborts", fakes[5].Aborted, true)
            eq("nothing pending after timeout", HttpGet.Pending, 0)

            ; while one request's callback is still running (e.g. it showed a dialog), another one still completes
            got.Length := 0, release := false
            HttpGet.Start("https://example.test/b", (t) => (got.Push("b"), release := true))
            HttpGet.Start("https://example.test/a", (t) => (got.Push("a start"), TestUtil.WaitFor(() => release, 2000), got.Push("a end")))
            fakes[6].status := 200, fakes[6].readyState := 4
            fakes[7].status := 200, fakes[7].readyState := 4
            TestUtil.WaitFor(() => got.Length >= 3, 4000)
            eq("callbacks do not block each other", got.Length = 3 ? got[3] : TextTools.JoinList(got, ","), "a end")

            HttpGet.NewRequest := () => ComObject("ALTRun.NoSuchClass")
            got.Length := 0
            HttpGet.Start("https://example.test/x", (t) => got.Push("done"), (m) => got.Push("fail " m))
            eq("cannot start -> fail", TestUtil.WaitFor(() => got.Length) ? SubStr(got[1], 1, 5) : "", "fail ")
        } finally {
            HttpGet.NewRequest := saved[1], HttpGet.PollMs := saved[2]
        }

        ; update and exchange-rate checks no longer use Download() (the whole program stops while it downloads); the one-click update's zip download has a progress window and still uses Download()
        for name in ["UpdateChecker", "CurrencyRates"] {
            text := FileRead(A_ScriptDir "\..\Src\Extensions\" name ".ahk", "UTF-8")
            pos := 1, count := 0
            while (pos := RegExMatch(text, "\bDownload\(", , pos)) {
                count++, pos++
            }
            eq(name ": synchronous downloads", count, name = "UpdateChecker" ? 1 : 0)
        }

        ; Check for updates: the release info is handled after it has downloaded (a fake reply)
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
            TestUtil.WaitFor(() => IsObject(UpdateChecker.Pending))
            eq("update: found in background", IsObject(UpdateChecker.Pending) ? UpdateChecker.Pending.Version : "", "2999.01.01")
            eq("update: check time saved", UpdateChecker.LoadState()["LastCheck"] != "", true)
            eq("manual check: busy flag starts clear", UpdateChecker._checking, false)
        } finally {
            HttpGet.NewRequest := saved[1], HttpGet.PollMs := saved[2]
            UpdateChecker.StateFile := savedState[1], UpdateChecker.Pending := savedState[2]
            try DirDelete(root, true)
        }

        ; A real request (Msxml2.XMLHTTP.6.0, only on Windows in GitHub Actions): the program keeps running while waiting, and the content is right.
        ; no network / a timeout isn't a failure (it has nothing to do with the program); it only prints a note
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

    ; The log file stays open (not opened and closed per line): readable right after writing; reopened after a file change, deletion or truncation
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
                FileDelete(second)                                          ; deleted while open (Wine can't delete an open file, so skipped there)
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

    ; A background task run over SlowMs logs "Perf: slow"; fast ones don't; return values and parameters as usual
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

    ; The released v2026.08.12 used ALTRun.ini: imported as a whole at the first start of the new version
    static LegacyIni() {
        eq := (n, a, e) => TestRunner.Equal("LegacyIni." n, a, e)
        folder := A_Temp "\ALTRunLegacyTest"
        try DirDelete(folder, true)
        DirCreate(folder)
        iniFile := folder "\ALTRun.ini"
        ; like the files 2.x writes: UTF-16, comment lines, = and ; escaped in keys
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

        ; The full start-up: no ALTRun.json, only ALTRun.ini in the program folder
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

        ; an ALTRun.ini written by the real v2026.08.12 (UTF-16, with default commands) + a command and settings the user added
        fixture := SchemaMigration.Upgrade(SchemaMigration.ReadLegacyIni(A_ScriptDir "\Fixtures\ALTRun.v2026.08.12.ini"), 2)
        eq("fixture commands", fixture["CustomCommands"].Length, 10)
        eq("fixture language", fixture["General"]["Language"], "zh-CN")
        eq("fixture keep input (2.x default on)", fixture["General"]["KeepLastQuery"], 1)
        eq("keep last query default off", AppSettings.Defaults()["General"]["KeepLastQuery"], 0)
        eq("fixture hotkey", fixture["Hotkeys"][1]["Key"], "~Mbutton")

        ; second start: ALTRun.json exists now, no import again
        AppSettings.ImportedFrom := ""
        AppSettings.Load()
        eq("no second import", AppSettings.ImportedFrom, "")

        AppSettings.File := savedFile, AppSettings.LegacyFile := savedLegacy, AppSettings.Data := savedData
        AppSettings.ImportedFrom := "", AppSettings.MigratedFrom := 0
        try DirDelete(folder, true)
    }

    ; An old ALTRun.json in the program folder moves to Data\ at start-up
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

        ; the 3.x layout: ALTRun.json, upgrade backups and .bad all in the program folder, no Data\ yet
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

        ; the next start: already in Data\, no move
        AppSettings.MovedFrom := "", AppSettings.MigratedFrom := 0
        AppSettings.Load()
        eq("second start", AppSettings.MovedFrom "|" AppSettings.MigratedFrom, "|0")

        ; both exist (e.g. an old version run again made another): the one in Data\ is used, the old file left alone
        write(root "\ALTRun.json", "{}")
        AppSettings.Load()
        eq("both exist: Data wins", AppSettings.General["Hotkey"] "|" AppSettings.MovedFrom, "#Space|")
        eq("both exist: old file kept", read(root "\ALTRun.json"), "{}")

        ; can't move (the file is in use): keep the old location this time
        FileDelete(root "\Data\ALTRun.json")
        lock := FileOpen(root "\ALTRun.json", "r -rwd")
        AppSettings.MoveLegacyFile()
        lock.Close()
        eq("locked: keep using old file", AppSettings.File "|" AppSettings.MovedFrom, root "\ALTRun.json|")

        AppSettings.File := savedFile, AppSettings.LegacyFile := savedLegacy, AppSettings.Data := savedData
        AppSettings.MovedFrom := "", AppSettings.MigratedFrom := 0, AppSettings.ImportedFrom := ""
        try DirDelete(root, true)
    }

    ; Usage statistics: counted per day and feature; totals / per day / save and load (with temporary files, not the real Data\Usage.json)
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
            Usage.CountItem({Provider: "", Kind: "url"})                    ; web search fallback
            Usage.CountItem({Provider: "", Kind: ""})                       ; file search fallback
            Usage.CountItem({Provider: "Help", Kind: "url"})                ; the cheat sheet doesn't count
            all := Usage.Summary(0)
            eq("count item", all["CustomCommands"] "|" all["WebSearch"] "|" all["FileSearch"] "|" all.Has("Help"), "1|1|1|0")

            Usage.Days := Map("2020-01-01", Map("Calculator", 9))           ; too old: removed when saving
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

    ; JSON: escaping, numbers, nesting, round trips; reading large files is linear (passing the whole text by value used to get slower with size)
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

        ; control characters (e.g. ESC copied from a terminal) are written as \u001B; no raw control characters in the file
        control := "a" Chr(27) "[0m" Chr(8) Chr(12) Chr(1) "z"
        written := JSON.Stringify(Map("Text", control))
        ok("control chars escaped", !RegExMatch(written, "[\x00-\x08\x0B\x0C\x0E-\x1F]") && InStr(written, "\u001B[0m\u0008\u000C\u0001z"))
        eq("control chars round trip", JSON.Parse(written)["Text"], control)
        ; Writing files (SafeFile.WriteText, used by JSON.WriteFile): a temporary file first, then replaced; no BOM, no .tmp left behind
        dir := A_Temp "\ALTRun-json-test"
        try DirDelete(dir, true)
        jsonFile := dir "\sub\data.json"
        JSON.WriteFile(jsonFile, Map("k", "中文"))
        JSON.WriteFile(jsonFile, Map("k", "v2"))
        firstByte := FileOpen(jsonFile, "r").RawRead(bytes := Buffer(3), 3) ? NumGet(bytes, 0, "UChar") : 0
        eq("write file", JSON.Parse(FileRead(jsonFile, "UTF-8"))["k"] "|" (firstByte = 0xEF ? "bom" : "no bom") "|" (FileExist(jsonFile ".tmp") ? "tmp" : ""), "v2|no bom|")
        locked := FileOpen(jsonFile, "r -rwd")                              ; a sync folder / antivirus holds the file for a moment: written once it is released
        SetTimer(() => locked.Close(), -300)
        JSON.WriteFile(jsonFile, Map("k", "v3"))
        eq("write file: retries while locked", JSON.Parse(FileRead(jsonFile, "UTF-8"))["k"], "v3")
        locked := FileOpen(jsonFile, "r -rwd"), saved := SafeFile.MoveTries, SafeFile.MoveTries := 2
        failed := false
        try JSON.WriteFile(jsonFile, Map("k", "v4"))
        catch
            failed := true
        locked.Close(), SafeFile.MoveTries := saved
        if failed                                                           ; Wine may not support exclusive opening; then the write just succeeds
            eq("write file: gives up, original kept, no tmp left", JSON.Parse(FileRead(jsonFile, "UTF-8"))["k"] "|" (FileExist(jsonFile ".tmp") ? "tmp" : ""), "v3|")
        binFile := dir "\data.bin", bytes := Buffer(4), NumPut("UInt", 0x01020304, bytes)    ; binary content through SafeFile.Write
        SafeFile.Write(binFile, (out) => (out.Write("AB"), out.RawWrite(bytes, 4)))
        back := FileRead(binFile, "RAW")
        eq("write binary", back.Size "|" StrGet(back, 2, "CP0") "|" Format("{:X}", NumGet(back, 2, "UInt")), "6|AB|1020304")
        failed := false
        try SafeFile.Write(binFile, (out) => (out.Write("x"), Integer("not a number")))
        catch
            failed := true
        eq("write: a writer error keeps the old file, no tmp left", failed "|" FileGetSize(binFile) "|" (FileExist(binFile ".tmp") ? "tmp" : ""), "1|6|")
        DirDelete(dir, true)

        paths := []
        Loop 20000
            paths.Push("C:\Users\someone\Documents\Folder" A_Index "\Report " A_Index ".docx")
        text := JSON.Stringify(Map("Paths", paths))
        start := A_TickCount
        parsed := JSON.Parse(text)
        elapsed := A_TickCount - start
        eq("big file", parsed["Paths"].Length "|" parsed["Paths"][20000], "20000|" paths[20000])
        ok("big file is fast (" elapsed " ms)", elapsed < 3000)             ; used to take minutes
    }

    ; Unit / currency conversion (calculator)
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

        ; results in the calculator
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

    ; Snippet editor: every field has a gray description; warns when a keyword is too long / has spaces
    ; Table text: copied from Excel (Tab-separated, quoted when a cell has line breaks), CSV, European Excel's ; separator
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
        csvFile := A_Temp "\ALTRunDelimited.csv"
        for encoding in ["UTF-8", "UTF-8-RAW"] {                            ; with BOM / without BOM
            try FileDelete(csvFile)
            FileAppend("名称,正文`r`n邮箱,谢谢", csvFile, encoding)
            eq("read " encoding, DelimitedText.ReadFile(csvFile), "名称,正文`r`n邮箱,谢谢")
        }
        try FileDelete(csvFile)
        FileAppend("名称,正文", csvFile, "CP936")                              ; GBK (the CSV Chinese Excel saves by default): read with the system code page (CI isn't a Chinese system, so it only checks both ways read the same)
        eq("read system code page", DelimitedText.ReadFile(csvFile), FileRead(csvFile, "CP0"))
        try FileDelete(csvFile)
    }

    static Misc() {
        ; The log and .txt output open with the default program, or Notepad when that fails (no association)
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
        SearchInputLanguage._layoutBefore := 0x08040804
        SearchInputLanguage.Restore()                                        ; the option is off: do nothing, only clear the record
        TestRunner.Equal("Input.restore clears the saved layout", SearchInputLanguage._layoutBefore, 0)
        SearchInputLanguage._imeBefore := {Open: 1, Conv: 0x401}
        SearchInputLanguage.Restore()                                        ; English mode too: not restored when off, only the record cleared
        TestRunner.Equal("Input.restore clears the saved IME mode", SearchInputLanguage._imeBefore, "")
        AppSettings.General["SwitchToEnglishInput"] := 1
        SearchInputLanguage._imeBefore := {Open: 1, Conv: 0x401}
        SearchInputLanguage.Restore()                                        ; the English keyboard was chosen: the Chinese / English mode is left alone
        TestRunner.Equal("Input.keyboard choice leaves the IME mode alone", SearchInputLanguage._imeBefore, "")
        savedPlaceholder := AppSettings.Appearance["Placeholder"]
        AppSettings.Appearance["Placeholder"] := ""
        TestRunner.Equal("Placeholder.default", SearchWindow.Placeholder(), "ALTRun Search")
        AppSettings.Appearance["Placeholder"] := " "
        TestRunner.Equal("Placeholder.only spaces: nothing", SearchWindow.Placeholder(), "")
        AppSettings.Appearance["Placeholder"] := " Type to search "
        TestRunner.Equal("Placeholder.custom (trimmed)", SearchWindow.Placeholder(), "Type to search")
        AppSettings.Appearance["Placeholder"] := savedPlaceholder
        ; The input box is centered on the capital-letter height (#158): at 150% the baseline measured 42 and the capital height 33, so it moves down 7 pixels
        TestRunner.Equal("InputTop.cap centered", SearchWindow.InputTop(64, 64, 21, 42, 33), 7)
        TestRunner.Equal("InputTop.at most half the padding", SearchWindow.InputTop(64, 64, 8, 42, 33), 4)
        TestRunner.Equal("InputTop.no metrics: centered", SearchWindow.InputTop(60, 40, 14, 0, 0), 10)
        TestRunner.Equal("InputTop.never above centered", SearchWindow.InputTop(60, 40, 14, 42, 10), 10)
        g := Gui()
        g.SetFont("s20", "Segoe UI")
        metrics := Win.CapMetrics(g.AddEdit("w200 r1").Hwnd)
        ; The real row height of the result list (the last row isn't cut off): rows set by a 48-pixel image list, as in SearchWindow
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
        ; Reset to defaults: options become the defaults, what the user added (commands, snippets, hotkeys, web searches, pinned items, hidden system commands) stays
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
        ; System info: has the version, no data folder path
        info := App.SystemInfo()
        TestRunner.True("SystemInfo.version", InStr(info, App.Name " " App.Version " (") = 1)
        TestRunner.True("SystemInfo.no paths", !InStr(info, AppSettings.DataDir) && !InStr(info, A_ScriptDir))
        TestRunner.Equal("SystemInfo.lines", StrSplit(Trim(info, "`n"), "`n").Length, 9)
        ; Running as administrator (#158): a new elevated instance replaces a running non-elevated one; otherwise the arguments are handed over as usual
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
