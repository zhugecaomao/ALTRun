;===============================================================================
; ProjectTests.ahk - Repository checks: translations, Wiki pages, changelog links, version numbers
;   and source conventions
;-------------------------------------------------------------------------------
; Included by Tests\RunTests.ahk. Every static method whose name doesn't start with "_" is a test;
; TestRunner runs them in name order, and each one must put back the shared state it changes.
;===============================================================================
class ProjectTests {
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
            if (InStr(source, '"' key '"', true) || SubStr(key, -5) = ".Desc" || SubStr(key, -5) = ".Help")
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
}
