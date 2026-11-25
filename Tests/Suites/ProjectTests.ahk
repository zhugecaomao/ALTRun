;===============================================================================
; ProjectTests.ahk - Repository checks: translations, Wiki pages, changelog links, version numbers
;   and source conventions
;-------------------------------------------------------------------------------
; Included by Tests\RunTests.ahk. Every static method whose name doesn't start with "_" is a test;
; TestRunner runs them in name order, and each one must put back the shared state it changes.
;===============================================================================
class ProjectTests {
    ; Layers: Lib uses nothing from Src; Src\Core uses nothing from the UI, the providers or the extensions (only
    ; App.ahk, which starts everything, does). Lower layers offer callbacks instead (e.g. BackgroundWork.IsBusy)
    static Layers() {
        ProjectTests._CheckLayer("Lib", ["Core", "UI", "Providers", "Extensions"])
        ProjectTests._CheckLayer("Src\Core", ["UI", "Providers", "Extensions"], ["App.ahk"])
    }

    ; AHK names are case-insensitive: a parameter or local variable named like a class hides the class in its function,
    ; so a call such as TypeAhead.Start() there fails at run time ("Integer has no method named Start"). Flags a function
    ; that has such a name and also uses that class (a parameter named path in a function that never uses Path is fine)
    static ShadowedClasses() {
        files := [], classes := Map()
        for folder in ["Src", "Lib"]
            Loop Files, A_ScriptDir "\..\" folder "\*.ahk", "R"
                files.Push(A_LoopFileFullPath)
        for sourceFile in files
            Loop Parse, FileRead(sourceFile, "UTF-8"), "`n", "`r"
                if RegExMatch(A_LoopField, "^class (\w+)", &m)
                    classes[StrLower(m[1])] := m[1]
        found := ""
        for sourceFile in files {
            SplitPath(sourceFile, &fileName)
            lines := [], startLine := 0
            Loop Parse, FileRead(sourceFile, "UTF-8") "`n}", "`n", "`r" {
                code := RegExReplace(A_LoopField, "`"[^`"]*`"|'[^']*'", "''")    ; strings and comments don't count
                code := RegExReplace(code, "(^|\s);.*$")
                if (RegExMatch(code, "^    static \w+\(") || RegExMatch(code, "^\}")) {
                    found .= ProjectTests._Shadowed(lines, classes, fileName ":" startLine)
                    lines := [], startLine := A_Index
                }
                lines.Push(code)
            }
        }
        TestRunner.Equal("ShadowedClasses", found, "")
    }

    ; The names a function's lines declare (parameters, assigned locals, for variables) that hide a class it also uses
    static _Shadowed(lines, classes, where) {
        if !lines.Length
            return ""
        names := Map()
        if RegExMatch(lines[1], "^    static \w+\((.*?)\)\s*(\{|=>)", &m)
            for param in StrSplit(m[1], ",")
                if RegExMatch(param, "^\s*&?(\w+)", &p)
                    names[StrLower(p[1])] := true
        for line in lines {
            pos := 1
            while (pos := RegExMatch(line, "(?<![\w.%])(\w+)\s*:=", &m, pos))
                names[StrLower(m[1])] := true, pos += StrLen(m[0])
            if RegExMatch(line, "^\s*for\s+(\w+)(?:\s*,\s*(\w+))?\s+in\b", &m)
                Loop m.Count
                    if (m[A_Index] != "")
                        names[StrLower(m[A_Index])] := true
        }
        found := ""
        for name in names {
            if !classes.Has(name)
                continue
            for line in lines
                if RegExMatch(line, "i)(?<![\w.])" classes[name] "\.\w") {
                    found .= where " " classes[name] ", "
                    break
                }
        }
        return found
    }

    static _CheckLayer(folder, upperFolders, skip := []) {
        names := ""
        for upper in upperFolders
            Loop Files, A_ScriptDir "\..\Src\" upper "\*.ahk"
                Loop Parse, FileRead(A_LoopFileFullPath, "UTF-8"), "`n", "`r"
                    if RegExMatch(A_LoopField, "^class (\w+)", &m)
                        names .= (names = "" ? "" : "|") m[1]
        TestRunner.True("Layers." folder " has upper classes to check", names != "")
        found := ""
        Loop Files, A_ScriptDir "\..\" folder "\*.ahk" {
            fileName := A_LoopFileName
            for skipped in skip
                if (fileName = skipped)
                    continue 2
            Loop Parse, FileRead(A_LoopFileFullPath, "UTF-8"), "`n", "`r" {
                code := RegExReplace(A_LoopField, "`"[^`"]*`"|'[^']*'", "''")    ; strings (e.g. "QuickSwitch.Title" keys) don't count
                code := RegExReplace(code, "(^|\s);.*$")                    ; nor comments
                if RegExMatch(code, "\b(" names ")\.", &m)
                    found .= (found = "" ? "" : ", ") fileName ":" A_Index " " m[1]
            }
        }
        TestRunner.Equal("Layers." folder " uses no upper layer", found, "")
    }

    ; Every Edit control must give a row count (r1 / r8 ...): without it long text makes AHK turn it multi-line and taller, covering the controls below
    static EditRows() {
        missing := ""
        for folder in ["Src", "Lib"] {
            Loop Files, A_ScriptDir "\..\" folder "\*.ahk", "R" {
                Loop Parse, FileRead(A_LoopFileFullPath, "UTF-8"), "`n", "`r" {
                    if (RegExMatch(A_LoopField, 'AddEdit\(|PrefsForm\.Add\("Edit"') && !RegExMatch(A_LoopField, '[" ]r(\d|"\s)'))
                        missing .= A_LoopFileName ":" A_Index " "
                }
            }
        }
        TestRunner.Equal("EditRows.all edits set rows", missing, "")
    }

    ; No empty English text; every key in every language file (Resources\Lang\*.json) is translated, no extra keys, placeholders {1} {2} ... as in English;
    ; reading language files: code forms, _fallback, _font, falling back when a file or key is missing
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
        for value in PrefsPages._LanguageValues()
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
        I18n.Init("ko-KR")                                                  ; a language that doesn't exist: English
        eq("unknown", I18n.Lang "|" I18n.T("Prefs.Page.General"), "en|General")

        savedDir := I18n.LangDir
        dir := A_Temp "\ALTRun-lang-test"
        try DirDelete(dir, true)
        DirCreate(dir)
        I18n.LangDir := dir
        FileCopy(savedDir "\zh-CN.json", dir "\zh-CN.json")
        I18n.Init("zh-TW")                                                  ; the Traditional Chinese file is missing: Simplified Chinese
        eq("no zh-TW file", I18n.Lang "|" I18n.T("Prefs.Page.General"), "zh-CN|通用")
        FileAppend('{"_name": "繁體中文", "_font": "Test Font", "_fallback": "zh-CN", "Prefs.Page.General": "一般"}', dir "\zh-TW.json", "UTF-8")
        I18n.Init("zh-TW")                                                  ; missing keys: _fallback first, then English
        eq("fallback", I18n.Lang "|" I18n.T("Prefs.Page.General") "|" I18n.T("Prefs.Page.Appearance") "|" ThemeManager.FontName(), "zh-TW|一般|外观|Test Font")
        FileAppend('{"_name": "Test", "_fallback": "xx", "Prefs.Page.General": "Allgemein"}', dir "\de.json", "UTF-8")
        I18n.Init("de-DE")                                                  ; a file for the language part only; the _fallback file missing: English
        eq("base language", I18n.Lang "|" I18n.T("Prefs.Page.General") "|" I18n.T("Prefs.Page.Appearance") "|" ThemeManager.FontName(), "de|Allgemein|Appearance|Segoe UI")
        eq("list from folder", I18n.Languages().Length, 4)
        FileDelete(dir "\zh-CN.json")
        FileAppend("{broken", dir "\zh-CN.json", "UTF-8")
        I18n.Init("zh-CN")                                                  ; can't be read: English
        eq("broken file", I18n.Lang "|" I18n.T("Prefs.Page.General"), "en|General")
        I18n.LangDir := savedDir
        DirDelete(dir, true)
        I18n.Init(savedLang)
    }

    ; Every UI text is used by the code: the key appears as is in the source, or belongs to one of the groups built from names (Theme.<theme name> etc.)
    static I18nUnused() {
        source := ""
        Loop Files, A_ScriptDir "\..\Src\*.ahk", "R"
            if (A_LoopFileName != "I18n.ahk")
                source .= FileRead(A_LoopFileFullPath, "UTF-8")
        dynamic := ["Help.", "Theme.", "Setting.", "Usage.F.", "Usage.Col.", "Prefs.Feature.", "Prefs.Status.", "Prefs.TypeShort.", "Cmd.Field.", "Sync.Problem.", "Sync.Mode.", "Sync.ModeHint.", "Sync.Preview.Copy", "Sync.Preview.Delete", "Sync.Preview.Why.", "Sync.Warn."]
        unused := ""
        for key in I18n.Strings {
            if (InStr(source, '"' key '"', true) || SubStr(key, -5) = ".Desc" || SubStr(key, -5) = ".Help")
                continue
            for prefix in dynamic
                if (SubStr(key, 1, StrLen(prefix)) == prefix)
                    continue 2
            unused .= key " "
        }
        TestRunner.Equal("I18nUnused", unused, "")
    }

    ; Wiki pages in English and Chinese (English Usage / Chinese Usage-zh): every page has a Chinese page, Wiki links in pages point to pages that exist,
    ; and the help pages the program opens (cheat sheet, F1 in Preferences) exist in both languages; no old en- pages any more
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
            while (pos := RegExMatch(text, "\]\(([A-Za-z][\w-]*)(#[^)]*)?\)", &m, pos)) {   ; [text](Page) or [text](Page#heading)
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

    ; Markdown tables in the Wiki and the READMEs have as many cells in every row as the delimiter row has columns. GitHub splits cells
    ; at every "|" not written "\|", inside `inline code` too, so a cell showing `{query|x}` shifts the rest of its row
    static MarkdownTables() {
        root := A_ScriptDir "\..\", files := ["README.md", "README.zh-CN.md", "CONTRIBUTING.md", "docs\scripts\README.md"]
        Loop Files, root "docs\wiki\*.md"
            files.Push("docs\wiki\" A_LoopFileName)
        problems := "", tables := 0
        for relative in files
            problems .= ProjectTests._TableProblems(StrReplace(relative, "\", "/"), FileRead(root relative, "UTF-8"), &tables)
        TestRunner.True("MarkdownTables.has tables (" tables ")", tables >= 100)
        TestRunner.Equal("MarkdownTables.problems", problems, "")
        ; the check itself, on made-up text
        table := "| A | B |`n|---|:--:|`n"
        TestRunner.Equal("MarkdownTables.check: good", ProjectTests._TableProblems("a", table "| ``x`` | y \| z |`n``a\|b`` | c`n"), "")
        TestRunner.Equal("MarkdownTables.check: pipe in inline code", ProjectTests._TableProblems("a", table "| ``{query|x}`` | y |`n"), "a:3 expected 2 columns, got 3; ")
        TestRunner.Equal("MarkdownTables.check: header", ProjectTests._TableProblems("a", "| A | B | C |`n|---|---|`n| x | y |`n"), "a:1 expected 2 columns, got 3; ")
        TestRunner.Equal("MarkdownTables.check: short row", ProjectTests._TableProblems("a", table "| x |`n`n| not | a | table |`n"), "a:3 expected 2 columns, got 1; ")
        TestRunner.Equal("MarkdownTables.check: code fence", ProjectTests._TableProblems("a", "~~~~md`n" table "| x | y | z |`n~~~`n~~~~`n``````"), "")
        TestRunner.Equal("MarkdownTables.check: heading ends a table", ProjectTests._TableProblems("a", table "| x | y |`n## Next`n"), "")
    }

    ; Rows of the Markdown tables in text (file name) whose cell count isn't the delimiter row's, as "name:line expected N columns, got M; "
    ; for each; adds the tables found to tables. Fenced code is skipped; a table ends at a blank line, a heading or a quote
    static _TableProblems(name, text, &tables := 0) {
        problems := "", fence := "", columns := 0, previous := ""
        for index, line in StrSplit(text, "`n", "`r") {
            if (fence != "") {                                             ; a fence closes with the same character, at least as many times
                if RegExMatch(line, "^\s*" SubStr(fence, 1, 1) "{" StrLen(fence) ",}\s*$")
                    fence := ""
                continue
            }
            if RegExMatch(line, "^\s*(``{3,}|~{3,})", &m) {
                fence := m[1], columns := 0, previous := ""
                continue
            }
            row := Trim(line)
            if (columns > 0) {
                if (row = "" || RegExMatch(row, "^[#>]"))
                    columns := 0
                else if ((cells := ProjectTests._TableCells(row)) != columns)
                    problems .= name ":" index " expected " columns " columns, got " cells "; "
            } else if (InStr(previous, "|") && InStr(row, "|") && RegExMatch(row, "^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?$")) {
                columns := ProjectTests._TableCells(row), tables += 1
                if ((cells := ProjectTests._TableCells(previous)) != columns)
                    problems .= name ":" (index - 1) " expected " columns " columns, got " cells "; "
            }
            previous := row
        }
        return problems
    }

    ; Cells in a table row: "\|" is a pipe in a cell, and a pipe at the start or end of the row is optional
    static _TableCells(row) {
        row := RegExReplace(Trim(StrReplace(row, "\|")), "^\||\|$")
        StrReplace(row, "|", "", , &count)
        return count + 1
    }

    ; The script library (docs\scripts, listed on the website by site\build.py): each script has in its first 30 lines (where ALTRun reads
    ; @altrun.xxx) a title, a description in English and Chinese and @altrun.network (the hosts it connects to, or "none"); a script that
    ; makes web requests names its hosts, each host named appears in the script, and no API key or token is written in it
    static ScriptLibrary() {
        dir := A_ScriptDir "\..\docs\scripts\"
        count := 0, problems := ""
        Loop Files, dir "*.*", "F" {
            if !RegExMatch(A_LoopFileExt, "i)^(ahk|ps1|bat|cmd)$") {
                if (A_LoopFileName != "README.md")
                    problems .= A_LoopFileName ": not .ahk, .ps1, .bat or .cmd; "
                continue
            }
            count += 1, text := FileRead(A_LoopFileFullPath, "UTF-8")
            problems .= ProjectTests._ScriptProblems(A_LoopFileName, text)
            if (A_LoopFileExt = "bat" || A_LoopFileExt = "cmd") {         ; cmd takes a byte order mark as part of the first command, and reads the file in the console's code page
                handle := FileOpen(A_LoopFileFullPath, "r", "CP0"), head := Buffer(3, 0), handle.RawRead(head, 3), handle.Close()
                if (NumGet(head, 0, "UChar") = 0xEF)
                    problems .= A_LoopFileName ": has a byte order mark; "
                for line in StrSplit(text, "`n", "`r")
                    if RegExMatch(line, "[^\x00-\x7F]") && !RegExMatch(line, "i)^\s*(rem\b|::)")
                        problems .= A_LoopFileName ": non-ASCII text outside a rem line; "
            }
        }
        TestRunner.True("ScriptLibrary.has scripts (" count ")", count >= 6)
        TestRunner.Equal("ScriptLibrary.problems", problems, "")
        ; the checks themselves, on made-up scripts
        head := "; @altrun.title T`n; @altrun.description D`n; @altrun.description_zh 说明`n"
        TestRunner.Equal("ScriptLibrary.check: good", ProjectTests._ScriptProblems("a", head "; @altrun.network example.com`nurl := `"https://example.com`"`nr := ComObject(`"WinHttp.WinHttpRequest.5.1`")"), "")
        TestRunner.Equal("ScriptLibrary.check: web, network none", ProjectTests._ScriptProblems("a", head "; @altrun.network none`nr := ComObject(`"WinHttp.WinHttpRequest.5.1`")"), "a: makes web requests but @altrun.network is none; ")
        TestRunner.Equal("ScriptLibrary.check: host not used", ProjectTests._ScriptProblems("a", head "; @altrun.network example.org`n"), "a: network host example.org isn't used in the script; ")
        TestRunner.Equal("ScriptLibrary.check: key written", ProjectTests._ScriptProblems("a", head "; @altrun.network none`nkey := `"sk-abcdefghijklmnopqrstuvwxyz`""), "a: looks like a key or token is written in it (sk-abc...); ")
        TestRunner.Equal("ScriptLibrary.check: COM ID in braces", ProjectTests._ScriptProblems("a", head "; @altrun.network none`nclsid := `"{557CF406-1A04-11D3-9A73-0000F81EF32E}`""), "")
        TestRunner.Equal("ScriptLibrary.check: key that looks like a GUID", ProjectTests._ScriptProblems("a", head "; @altrun.network none`nkey := `"557cf406-1a04-11d3-9a73-0000f81ef32e:fx`""), "a: looks like a key or token is written in it (557cf4...); ")
        TestRunner.Equal("ScriptLibrary.check: author", ProjectTests._ScriptProblems("a", head "; @altrun.network none`n; @altrun.author some-user1`n"), "")
        TestRunner.Equal("ScriptLibrary.check: author not a user name", ProjectTests._ScriptProblems("a", head "; @altrun.network none`n; @altrun.author Some One`n"), "a: @altrun.author isn't a GitHub user name; ")
        TestRunner.Equal("ScriptLibrary.check: no description", ProjectTests._ScriptProblems("a", "# @altrun.title T`n# @altrun.network none`n"), "a: no @altrun.description; a: no @altrun.description_zh; ")
        TestRunner.Equal("ScriptLibrary.check: description after line 30", ProjectTests._ScriptProblems("a", "; @altrun.title T`n; @altrun.network none`n" StrReplace(Format("{:30}", ""), " ", "`n") "; @altrun.description D`n; @altrun.description_zh 说明`n"), "a: no @altrun.description; a: no @altrun.description_zh; ")
    }

    ; What a library script (name, text) lacks or does wrong, as "name: problem; " for each, or ""
    static _ScriptProblems(name, text) {
        meta := ProjectTests._ScriptMeta(text), problems := ""
        for key in ["title", "description", "description_zh", "network"]
            if !meta.Has(key) || meta[key] = ""
                problems .= name ": no @altrun." key "; "
        if (meta.Has("description_zh") && !RegExMatch(meta["description_zh"], "[\x{4E00}-\x{9FFF}]"))
            problems .= name ": description_zh isn't Chinese; "
        if (meta.Has("requires") != meta.Has("requires_zh"))
            problems .= name ": requires and requires_zh go together; "
        if (meta.Has("author") && !RegExMatch(meta["author"], "^[A-Za-z0-9][A-Za-z0-9-]{0,38}$"))   ; the website links it to github.com/<author>
            problems .= name ": @altrun.author isn't a GitHub user name; "
        network := meta.Has("network") ? Trim(meta["network"]) : ""
        web := RegExMatch(text, "i)WinHttp|XMLHTTP|\bDownload\(|Invoke-WebRequest|Invoke-RestMethod|Net\.WebClient|HttpClient|\bcurl(\.exe)?\s")
        if (web > 0 && network = "none")
            problems .= name ": makes web requests but @altrun.network is none; "
        if (network != "" && network != "none") {
            rest := RegExReplace(text, "im)^.*@altrun\.network\b.*$", "")           ; the script without its network line
            for host in StrSplit(network, ",", " `t")
                if !RegExMatch(host, "^[a-z0-9-]+(\.[a-z0-9-]+)+$")
                    problems .= name ": network host " host " isn't a host name; "
                else if !InStr(rest, host)
                    problems .= name ": network host " host " isn't used in the script; "
        }
        keyText := RegExReplace(text, "\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}", "{GUID}")   ; a COM class or interface ID in braces isn't a key
        if RegExMatch(keyText, "i)\b(?:[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}(?::fx)?|sk-[A-Za-z0-9_-]{20,}|gh[pousr]_[A-Za-z0-9]{30,}|AIza[0-9A-Za-z_-]{30,})", &key)
            problems .= name ": looks like a key or token is written in it (" SubStr(key[0], 1, 6) "...); "
        return problems
    }

    ; @altrun.xxx values in the first 30 lines, as ScriptProvider.ReadScript reads them (keys in lower case)
    static _ScriptMeta(text) {
        meta := Map()
        for line in StrSplit(text, "`n", "`r") {
            if (A_Index > 30)
                break
            if RegExMatch(line, "i)^\s*(?:;|#|//|::|rem\b)\s*@altrun\.(\w+)\s+(.*?)\s*$", &m)
                meta[StrLower(m[1])] := m[2]
        }
        return meta
    }

    ; Every version heading in CHANGELOG.md has a link ([version]: URL at the end of the file), and "未发布" compares against the latest release
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

    ; Timers that stay set while idle use WinTimer, not SetTimer: while any SetTimer timer is set, AutoHotkey checks its timers every 10 ms
    ; (about 330 million CPU cycles a minute). Flags SetTimer with a literal period of 2 s or more, or a one-shot delay of a minute or more.
    ; Periods held in variables aren't seen: the benchmark's idle CPU budget (Tests\Bench\bench_stats.py) checks the running program
    static LongTimers() {
        found := ""
        for folder in ["Src", "Lib"]
            Loop Files, A_ScriptDir "\..\" folder "\*.ahk", "R"
                Loop Parse, FileRead(A_LoopFileFullPath, "UTF-8"), "`n", "`r"
                    if (RegExMatch(A_LoopField, "^\s*[^;]*\bSetTimer\(.+,\s*(-?\d+)\)", &m) && (m[1] >= 2000 || m[1] <= -60000))
                        found .= A_LoopFileName ":" A_Index " "
        TestRunner.Equal("LongTimers.SetTimer kept set while idle (use WinTimer)", found, "")
    }

    ; A timer's function is called without arguments, so a static method passed on its own (WinTimer.Set(Class.Method, ...)) fails:
    ; it lacks its "this" ("Missing a required parameter"), and WinTimer only logs it. Pass a function such as () => Class.Method(),
    ; kept in a static property so the same function stops it again. Flags Class.Name passed to WinTimer.Set / SetTimer when Class has a
    ; method of that name (DelayedStart's start-up check never ran this way)
    static TimerCallbacks() {
        source := "", found := ""
        for folder in ["Src", "Lib"]
            Loop Files, A_ScriptDir "\..\" folder "\*.ahk", "R"
                source .= FileRead(A_LoopFileFullPath, "UTF-8") "`n"
        pos := 1
        while (pos := RegExMatch(source, "(?:WinTimer\.Set|SetTimer)\((\w+)\.(\w+)\s*,", &m, pos)) {
            pos += m.Len
            if RegExMatch(source, "im)^\s*static\s+\Q" m[2] "\E\(") && RegExMatch(source, "im)^class\s+\Q" m[1] "\E\b")
                found .= m[1] "." m[2] " "
        }
        TestRunner.Equal("TimerCallbacks.static method passed as a timer function", found, "")
    }

    ; The Preferences list (PrefsForm.List) hands out its functions in an object ({Items, Insert, Append, Edit, Reload}). They are plain
    ; functions, not methods: list.Insert(item) would pass the object as well ("Too many parameters"), so they are called with .Call
    ; (Startup Apps... on the Delayed Start page failed this way)
    static ListFunctions() {
        found := ""
        Loop Files, A_ScriptDir "\..\Src\*.ahk", "R"
            Loop Parse, FileRead(A_LoopFileFullPath, "UTF-8"), "`n", "`r"
                if RegExMatch(A_LoopField, "^\s*[^;]*\blist\.(Insert|Append|Edit|Reload)\(")
                    found .= A_LoopFileName ":" A_Index " "
        TestRunner.Equal("ListFunctions.called without .Call", found, "")
    }

    ; AutoHotkey's RegExMatch / RegExReplace / ~= throw (PCRE error -10) on text starting with ，（Ａ or ending with 﨑 (#653, #675), and
    ; almost any text can be like that: file names, paths, window titles, settings, clipboard. TextTools.Match / Replace handle it; flags a
    ; direct call whose text isn't a string literal. Not checked: TextTools.ahk itself, and Calc.Eval in Util.ahk, which stays the same
    ; as PT Tools' copy (PTCalcSameAsCalc) and gets only text Calc.Looks / CalculatorProvider let through
    static RawRegex() {
        found := ""
        for folder in ["Src", "Lib"]
            Loop Files, A_ScriptDir "\..\" folder "\*.ahk", "R" {
                if (A_LoopFileName = "TextTools.ahk")
                    continue
                inCalc := false
                Loop Parse, FileRead(A_LoopFileFullPath, "UTF-8"), "`n", "`r" {
                    if (A_LoopFileName = "Util.ahk" && InStr(A_LoopField, "`; Returns the result (a float)") || InStr(A_LoopField, "`; Thousands separators"))
                        inCalc := InStr(A_LoopField, "Returns the result") > 0
                    code := TextTools.Replace(A_LoopField, "(^|\s);.*$")
                    if (!inCalc && TextTools.Match(code, "(?<![\w.])RegEx(Match|Replace)\(\s*[^\s`"']|\s~=\s"))
                        found .= A_LoopFileName ":" A_Index " "
                }
            }
        TestRunner.Equal("RawRegex.RegExMatch / RegExReplace / ~= on text (use TextTools.Match / Replace)", found, "")
    }

    ; The log never holds what was typed in the search box, only its length (Privacy: the debug log is attached to bug reports).
    ; Flags Input.Value in a Logger call (its line or the ". " lines that continue it) unless it is inside StrLen(
    static NoInputTextInLog() {
        found := ""
        Loop Files, A_ScriptDir "\..\Src\*.ahk", "R" {
            inLog := false
            Loop Parse, FileRead(A_LoopFileFullPath, "UTF-8"), "`n", "`r" {
                code := RegExReplace(A_LoopField, "(^|\s);.*$")
                inLog := RegExMatch(code, "\bLogger\.\w+\(") || (inLog && RegExMatch(code, "^\s*\."))
                if (inLog && RegExMatch(RegExReplace(code, "StrLen\([\w.]*Input\.Value\)"), "Input\.Value"))
                    found .= A_LoopFileName ":" A_Index " "
            }
        }
        TestRunner.Equal("NoInputTextInLog.search box text in a log line", found, "")
    }

    ; The data files ALTRun reads back whole (the .json files in Data\ and FileIndex.bin, held in a File / IndexFile / StateFile property)
    ; are written with SafeFile (or JSON.WriteFile, which uses it): a temporary file flushed to the disk, then swapped in, so a crash or
    ; power cut halfway can't leave them empty or cut short. Flags FileAppend / FileOpen for writing on such a file or on a ".json" name
    static DataWrites() {
        found := ""
        Loop Files, A_ScriptDir "\..\Src\*.ahk", "R" {
            fileName := A_LoopFileName
            Loop Parse, FileRead(A_LoopFileFullPath, "UTF-8"), "`n", "`r" {
                code := RegExReplace(A_LoopField, "(^|\s);.*$")
                target := ""
                if RegExMatch(code, "\bFileAppend\((?:[^,()]|(\((?:[^()]|(?1))*\)))*,\s*([^,)]+)", &m)   ; the first argument may hold nested calls
                    target := m[2]
                else if RegExMatch(code, '\bFileOpen\(([^,]+),\s*"([^"]*)"', &m) && RegExMatch(m[2], "i)^[^-]*[wa]")
                    target := m[1]
                if (target != "" && RegExMatch(target, 'i)\.(File|IndexFile|StateFile)\b(?!\s*")|\.json"'))
                    found .= fileName ":" A_Index " "
            }
        }
        TestRunner.Equal("DataWrites.data file written without SafeFile", found, "")
    }

    ; Code that borrows the clipboard (saved := ClipboardAll()) puts it back in a finally block, so an error half way doesn't
    ; leave the user's clipboard empty or overwritten. Flags a ClipboardAll() with no "finally" before the next method
    static ClipboardRestore() {
        found := ""
        Loop Files, A_ScriptDir "\..\Src\*.ahk", "R" {
            fileName := A_LoopFileName, open := 0
            Loop Parse, FileRead(A_LoopFileFullPath, "UTF-8") "`n    static End(", "`n", "`r" {
                if (open && RegExMatch(A_LoopField, "^    static \w+\(")) {
                    found .= fileName ":" open " ", open := 0
                } else if (open && RegExMatch(A_LoopField, "^\s*(\}\s*)?finally\b")) {
                    open := 0
                }
                if RegExMatch(A_LoopField, "^[^;]*:=\s*ClipboardAll\(\)")
                    open := A_Index
            }
        }
        TestRunner.Equal("ClipboardRestore.borrowed clipboard without finally", found, "")
    }

    ; The version in the compile info (;@Ahk2Exe-SetVersion in ALTRun.ahk) must match App.Version
    static ReleaseVersion() {
        main := FileRead(A_ScriptDir "\..\ALTRun.ahk", "UTF-8")
        RegExMatch(main, "m);@Ahk2Exe-SetVersion\s+(\S+)", &m)
        TestRunner.Equal("ReleaseVersion.exe version = App.Version", IsObject(m) ? m[1] : "", App.Version)
        RegExMatch(main, "m);@Ahk2Exe-SetDescription\s+(.+?)\s*$", &m)                ; Windows uses it as the display name (notifications, Task Manager): only the software's name
        TestRunner.Equal("ReleaseVersion.exe description = App.Name", IsObject(m) ? m[1] : "", App.Name)
        TestRunner.True("ReleaseVersion.date format", RegExMatch(App.Version, "^\d{4}\.\d{2}\.\d{2}(\.\d+)?$"))   ; a second release on the same day: 2026.09.30.1
        TestRunner.True("ReleaseVersion.same-day release is newer", UpdateChecker.Compare("2026.09.30.1", "2026.09.30") > 0 && UpdateChecker.Compare("2026.10.01", "2026.09.30.2") > 0)
    }
}
