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
                    if (RegExMatch(A_LoopField, 'AddEdit\(|_Add\("Edit"') && !RegExMatch(A_LoopField, '[" ]r(\d|"\s)'))
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
        dynamic := ["Help.", "Theme.", "Setting.", "Usage.F.", "Usage.Col.", "Prefs.Feature.", "Prefs.Status.", "Prefs.TypeShort.", "Cmd.Field."]
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
    ; and the help pages the program opens (cheat sheet, F1 in Preferences) exist in both languages; no old en- pages any more (wiki.yml generates the redirect pages when publishing)
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
