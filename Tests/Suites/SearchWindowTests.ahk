;===============================================================================
; SearchWindowTests.ahk - The search window: keys, actions, result list drawing, icons, themes,
;   status bar, showing / hiding and the small windows around it (Large Type, HUD)
;-------------------------------------------------------------------------------
; Included by Tests\RunTests.ahk. Every static method whose name doesn't start with "_" is a test;
; TestRunner runs them in name order, and each one must put back the shared state it changes.
;===============================================================================
class SearchWindowTests {
    ; File actions: copy / cut the file, Open with, move to the Recycle Bin
    ; When the file type has an "Edit" verb the actions panel has "Edit" (with the editor's name on the right) and Ctrl+E works; unregistered types don't
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
            ; Wine's AssocQueryStringW finds no type at all (not even .txt), so the registered case is only checked on Windows
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
                for action in ResultActions.ListFor(FileSearchProvider._ToItem({Path: path, IsFolder: false}, 10))
                    list .= action.Title "|"
                return list
            }
            if !onWine
                eq("action shown", InStr(titles(root "\script.altruntest"), "|" I18n.T("Action.EditFile") "|") ? 1 : 0, 1)
            eq("action hidden", InStr(titles(root "\plain.altrunnone"), "|" I18n.T("Action.EditFile") "|") ? 1 : 0, 0)
            openWith := ""
            for action in ResultActions.ListFor(FileSearchProvider._ToItem({Path: root "\plain.altrunnone", IsFolder: false}, 10))
                if (action.Title = I18n.T("Action.OpenWith"))
                    openWith := action.Icon
            eq("open with icon", openWith, "res:shell32.dll,-3")
            if !onWine                                                      ; the former shell32.dll,-16702 doesn't exist on Windows, leaving the menu without an icon
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
        for action in ResultActions.ListFor(item)
            titles .= action.Title "|"
        for key in ["Action.OpenWith", "Action.CopyFile", "Action.CutFile", "Action.Recycle"]
            eq("has " key, InStr(titles, "|" I18n.T(key) "|") ? 1 : 0, 1)
        missing := FileSearchProvider._ToItem({Path: root "\gone.dwg", IsFolder: false}, 10)
        titles := "|"
        for action in ResultActions.ListFor(missing)
            titles .= action.Title "|"
        eq("missing file: no copy", InStr(titles, "|" I18n.T("Action.CopyFile") "|") ? 1 : 0, 0)
        drive := FileSearchProvider._ToItem({Path: "C:\", IsFolder: true}, 10)
        titles := "|"
        for action in ResultActions.ListFor(drive)
            titles .= action.Title "|"
        eq("drive: no recycle", InStr(titles, "|" I18n.T("Action.Recycle") "|") ? 1 : 0, 0)
        ClipboardData.Borrow(3000)
        TestRunner.True("FileActions.copy file", ActionCatalog.CopyFiles([root "\plan.dwg"]))
        eq("copied path", ClipboardData.Files().Length ? ClipboardData.Files()[1] : "", root "\plan.dwg")
        eq("copy is not cut", ClipboardData.IsCut(), false)
        ActionCatalog.CopyFiles([root "\plan.dwg"], true)
        eq("cut", ClipboardData.IsCut(), true)
        A_Clipboard := ""
        try DirDelete(root, true)
    }

    static EditActions() {
        eq := (n, a, e) => TestRunner.Equal("EditActions." n, a, e)
        saved := ProviderRegistry.Providers
        ProviderRegistry.Providers := [CustomCommandProvider, FileSearchProvider, ClipboardProvider]
        command := Map("Title", "Notes", "Type", "File", "Target", "C:\notes.txt", "Arguments", "", "Keyword", "")
        custom := CustomCommandProvider._ToItem(command, 50), custom.Provider := "CustomCommands"
        eq("custom edit", ResultActions.CanEdit(custom), true)
        eq("custom delete", ResultActions.CanDelete(custom), true)
        found := FileSearchProvider._ToItem({Path: "C:\Docs\a.pdf", IsFolder: false}, 10), found.Provider := "FileSearch"
        eq("file edit (add command)", ResultActions.CanEdit(found), true)
        eq("file delete", ResultActions.CanDelete(found), false)
        system := ResultItem("Lock", "", {OnRun: (*) => 0}), system.Provider := "System"
        eq("system edit", ResultActions.CanEdit(system), false)
        titles := ""
        for action in ResultActions.ListFor(custom)
            titles .= action.Title "|"
        TestRunner.True("EditActions.list has edit", InStr(titles, I18n.T("Action.Edit")) && InStr(titles, I18n.T("Action.Delete")))
        ProviderRegistry.Providers := saved
    }

    ; Matching characters highlighted in result titles
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
        eq("item ranges", str(ResultRows.HighlightRanges(item)), "1+1 8+1 15+1")
        SearchWindow.HighlightText := "code"
        eq("item ranges updated", str(ResultRows.HighlightRanges(item)), "15+4")
        SearchWindow.HighlightText := ""
        eq("no query", str(ResultRows.HighlightRanges(item)), "")
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
        font := ResultRows._CreateFont("Segoe UI", 13, 400)
        oldFont := DllCall("SelectObject", "Ptr", hdc, "Ptr", font, "Ptr")
        DllCall("SetBkMode", "Ptr", hdc, "Int", 1)
        gdi := ResultRows.Gdi, saved := Map()
        for key in ["TitleColor", "HighlightColor", "SelectedTitleColor", "SelectedHighlightColor"]
            saved[key] := gdi.Has(key) ? gdi[key] : ""
        gdi["TitleColor"] := 0, gdi["HighlightColor"] := 0, gdi["SelectedTitleColor"] := 0, gdi["SelectedHighlightColor"] := 0   ; one color: compare positions only
        ink(query) {                                                        ; whether each column has ink (ignoring color)
            rect := Buffer(16), NumPut("Int", 0, "Int", 0, "Int", width, "Int", height, rect)
            DllCall("FillRect", "Ptr", hdc, "Ptr", rect, "Ptr", DllCall("GetStockObject", "Int", 0, "Ptr"))
            ranges := (query = "") ? [] : FuzzyMatcher.MatchRanges(query, title)
            ResultRows.DrawTitle(hdc, title, ranges, 4, 0, width - 4, height, 0x24 | 0x800, false)
            header := Buffer(40, 0), NumPut("UInt", 40, "Int", width, "Int", -height, "UShort", 1, "UShort", 32, header)
            pixels := Buffer(width * height * 4)
            DllCall("GetDIBits", "Ptr", hdc, "Ptr", hbm, "UInt", 0, "UInt", height, "Ptr", pixels, "Ptr", header, "UInt", 0)
            columns := ""
            Loop width {
                x := A_Index - 1, dark := false
                Loop height {
                    color := NumGet(pixels, ((A_Index - 1) * width + x) * 4, "UInt")
                    if (Min(color & 0xFF, color >> 8 & 0xFF, color >> 16 & 0xFF) < 64) {   ; only stroke centers (antialiased edges overlap where parts meet)
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

    ; icons that aren't a system size (28 px for the compact density etc.) are scaled smoothly to the exact size
    ; Every built-in command has an icon (Wine lacks Control Panel and other shell items, so only checked on Windows)
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
        hIcon := IconCache._Load("::{A8A91A66-3A7D-4424-8D24-04E180695C7A}")      ; a Control Panel item given only as a CLSID
        TestRunner.True("BuiltinIcons.control panel item", hIcon)
        if hIcon
            DllCall("DestroyIcon", "Ptr", hIcon)
    }

    ; ALTRun's own icons in Resources\Icons: all present, each size reads at exactly that size; system icons when a file is missing
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

    ; Rounded rectangles of the selected row and theme thumbnails: solid inside and along straight edges, nothing outside the corners; antialiased corner edges (Wine's GDI+ doesn't antialias, so that part is skipped)
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
        Win.FillRoundRect(dc, 0, 0, 40, 24, 0, 0x000000)                       ; radius 0: a square rectangle
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
        IconCache.Size := 28                                                ; old-style icons without an alpha channel (transparent by mask)
        hIcon := IconCache.FitSize(DllCall("CopyIcon", "Ptr", DllCall("LoadIcon", "Ptr", 0, "Ptr", 32512, "Ptr"), "Ptr"))
        eq("legacy icon", IconCache.IconWidth(hIcon), 28)
        DllCall("DestroyIcon", "Ptr", hIcon)
        eq("no icon", IconCache.FitSize(0), 0)
        IconCache.Size := saved
        ; slow icon reads go into the debug log (with the threshold at 0 every one counts as slow)
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

    ; Theme thumbnails: every theme can be drawn, the background and selected row use the theme's own colors; clicking a thumbnail on the Appearance page chooses the theme
    ; Popup menus follow the theme's darkness: dark is decided by the background color; only set from Windows 10 1903 (18362)
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

        ; User themes: one with a built-in name overrides it ("Base" naming itself = changing the built-in one), and Base chains
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

    ; Density (Appearance.Density): any theme's sizes scale proportionally, colors unchanged; the former DarkCompact etc. count as Dark + Normal
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
        eq("scales theme sizes", ThemeManager.Scaled(Map("RowHeight", 60, "IconSize", 40), "Normal")["RowHeight"] "|" ThemeManager.Scaled(Map("RowHeight", 60, "IconSize", 40), "Normal")["IconSize"], "53|36")   ; a custom theme's own sizes scale down proportionally too
        ; built-in themes only differ in color, all sizes as in Light: switching themes at one density doesn't make the window jump in size
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
        ThemeManager.Load("DarkCompact", "Normal")                          ; never scaled twice
        eq("old name + normal once", ThemeManager.Get("RowHeight"), 48)
        ThemeManager.Load("DarkCompact", "Compact")
        eq("old name + compact", ThemeManager.Get("RowHeight"), 42)

        ; a user theme whose "Base" still names a former compact theme: Dark's colors and Normal's sizes as usual; the user's own theme of that name wins
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

        ; new installs default to Normal; an existing settings file without the entry keeps the old size (Comfortable), Normal with a former compact theme
        eq("new install default", AppSettings.Defaults()["Appearance"]["Density"], "Normal")
        keep := (theme) => (data := Map("Appearance", Map("Theme", theme)), AppSettings._KeepDensity(data), data["Appearance"]["Density"])
        eq("existing file keeps size", keep("Dark"), "Comfortable")
        eq("existing old compact theme", keep("darkcompact"), "Normal")
        set := Map("Appearance", Map("Theme", "Dark", "Density", "Compact"))
        eq("already set", AppSettings._KeepDensity(set) "|" set["Appearance"]["Density"], "0|Compact")
    }

    ; Large Type can be turned off in Preferences → Features: off for new installs, on for older settings files; when off neither the actions panel nor the cheat sheet lists it
    static LargeTypeSwitch() {
        eq := (n, a, e) => TestRunner.Equal("LargeTypeSwitch." n, a, e)
        eq("new install off", AppSettings.Defaults()["Features"]["LargeType"]["Enabled"], 0)
        old := Map("Features", Map("Help", Map("Enabled", 1)))
        eq("existing file keeps it on", AppSettings._KeepPrevious(old) "|" old["Features"]["LargeType"]["Enabled"], "1|1")
        set := Map("Features", Map("LargeType", Map("Enabled", 0)))
        eq("already set", AppSettings._KeepLargeType(set) "|" set["Features"]["LargeType"]["Enabled"], "0|0")
        options := AppSettings.Feature("LargeType"), saved := options["Enabled"]
        hasAction() {
            for action in ResultActions.ListFor(ResultItem("42", "", {Kind: "text", Arg: "42"}))
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

        ; Custom commands page: typing in the filter box shows only matching rows; filtering isn't a settings change
        saved := AppSettings.Data["CustomCommands"]
        AppSettings.Data["CustomCommands"] := [
            Map("Title", "Devices and Printers", "Type", "Command", "Target", "control", "Arguments", "printers", "Keyword", ""),
            Map("Title", "Notepad", "Type", "File", "Target", "C:\Windows\notepad.exe", "Arguments", "", "Keyword", "np"),
            Map("Title", "Printer queue", "Type", "Command", "Target", "control", "Arguments", "printers", "Keyword", "")]
        PreferencesWindow.Show(1, -3000, -3000)
        PreferencesWindow.BuildAll()                                         ; pages are built when needed: build them all before checking each page
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

    ; Browsing folders by typing a path: folders first, the last part filters, hidden files aren't listed, Tab goes one level down; Insert marks several to act on together
    static FolderBrowse() {
        eq := (n, a, e) => TestRunner.Equal("FolderBrowse." n, a, e)
        root := A_Temp "\ALTRun-browse-test"
        try DirDelete(root, true)
        DirCreate(root "\Reports"), DirCreate(root "\Archive")
        root := TestRunner.LongPath(root)                                   ; %Temp% may be an 8.3 short path (C:\Users\RUNNER~1), while files are listed with long paths
        FileAppend("x", root "\readme.txt"), FileAppend("x", root "\report 2026.docx"), FileAppend("x", root "\secret.txt")
        FileSetAttrib("+H", root "\secret.txt")
        browse := FileSearchProvider.BrowsePath(root "\rep")
        eq("parse", IsObject(browse) ? browse.Dir "|" browse.Filter : "", root "\|rep")
        eq("not a path", FileSearchProvider.BrowsePath("report"), "")
        eq("drive only", FileSearchProvider.BrowsePath("C:"), "")
        missing := FileSearchProvider.BrowsePath(root "\Nope\x")
        eq("missing folder", IsObject(missing) && missing.HasOwnProp("Missing") ? missing.Dir : "", root "\Nope\")
        eq("missing folder: hint", titles(root "\Nope\x"), I18n.T("Files.FolderNotFound"))
        wide := FileSearchProvider.BrowsePath(SubStr(root, 1, 1) "：、" StrReplace(SubStr(root, 4), "\", "、") "、rep")   ; full-width colon and ideographic comma from a Chinese IME
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
        eq("file mode", FileSearchProvider.SearchOnly(root "\arc")[1].Title, "Archive")

        SearchWindow.Marked := Map(StrLower(items[2].Arg), items[2].Arg)
        eq("marked", SearchWindow.IsMarked(items[2]) "|" SearchWindow.IsMarked(items[1]), "1|0")
        SearchWindow.Marked := Map()
        actions := ""
        for action in ResultActions.ListForMany([root "\readme.txt", root "\Archive"])
            actions .= action.Title "|"
        eq("many: open all", InStr(actions, I18n.T("Action.OpenAll", 2)) = 1, true)
        eq("many: copy, recycle", InStr(actions, I18n.T("Action.CopyFiles", 2)) && InStr(actions, I18n.T("Action.RecycleAll", 2)) ? 1 : 0, 1)
        FileSetAttrib("-H", root "\secret.txt")
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

    ; Explorer's "Send to": several files are added directly, existing ones not added again (one opens a dialog, not tested here)
    static SendTo() {
        eq := (n, a, e) => TestRunner.Equal("SendTo." n, a, e)
        ; Turning off start with Windows etc. deletes only ALTRun's own shortcuts; a shortcut of the same name the user made stays (#113)
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
        ; at start-up an unchanged shortcut isn't rewritten
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

        ; With Preferences open: added to the list on the Commands page (counts as a change, saved by OK / Apply), not written to the settings directly
        AppSettings.Data["CustomCommands"] := [Map("Title", "Old", "Type", "Folder", "Target", folder "\pt2415 - riverside\", "Arguments", "", "Keyword", "")]
        PreferencesWindow.Show(1, -3000, -3000)                             ; added right after opening (before changes are recorded) still counts as a change
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

    ; Ctrl+↑ / Ctrl+↓ browse the search history (↑ ↓ only move the selection); a hidden input box stands in for the search window, no real search
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

    ; Space with the caret at the start: the typed text switches to file search; Backspace at the start in file search mode: back to normal search; the text stays either way
    static FileModeKeys() {
        eq := (n, a, e) => TestRunner.Equal("FileModeKeys." n, a, e)
        savedInput := SearchWindow.Input, savedGui := SearchWindow.Gui, savedSearch := SearchWindow.GetOwnPropDesc("_RunSearch")
        options := AppSettings.Feature("FileSearch"), savedOptions := [options["SpacePrefix"], options["Enabled"]]
        SearchWindow.DefineProp("_RunSearch", {Call: (*) => 0})
        g := Gui()
        SearchWindow.Gui := g, SearchWindow.Input := g.AddEdit("w200")
        hwnd := SearchWindow.Input.Hwnd
        caret := (start, end := "") => SendMessage(0xB1, start, end = "" ? start : end, hwnd)   ; EM_SETSEL
        key := (vk) => SearchKeys.OnKeyDown(vk, 0, 0x100, hwnd)
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
        tab := () => SearchKeys.OnKeyDown(0x09, 0, 0x100, hwnd)
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
            eq("no system dialog open", SearchKeys.InSystemDialog(), 0)
            found := 0
            SetTimer(() => (found := SearchKeys.InSystemDialog(), found && WinClose("ahk_id " found)), -300)
            MsgBox("SystemDialogKeys", , "T3")                              ; closed by the timer above (or after 3 s)
            eq("finds an open MsgBox", found != 0, true)
            SearchWindow.Mode := "results", SearchWindow.Offset := 0, SearchWindow.Input.Value := "no"
            SearchWindow.Results := [ResultItem("Notepad"), ResultItem("Notes"), ResultItem("Notion")], SearchWindow.Selected := 1
            AppSettings.General["TabKey"] := "Next"
            tab()
            eq("not hooked: the key is lost", SearchWindow.Selected, 1)
            SearchKeys.Hook()
            eq("hooked: handled", tab(), 0)
            eq("hooked: next result", SearchWindow.Selected, 2)
            eq("its WM_CHAR is swallowed", SendMessage(0x102, 0x09, 0, hwnd) "|" SearchKeys._swallowChar, "0|0")
            SendMessage(0x100, 0x26, 0, SearchWindow.List.Hwnd)              ; ↑ with the list focused
            eq("list: previous result", SearchWindow.Selected, 1)
            SearchKeys._seen := true                                        ; OnMessage already handled this one
            tab()
            eq("seen by OnMessage: not handled twice", SearchWindow.Selected "|" SearchKeys._seen, "1|0")
            SendMessage(0x100, 0x41, 0, hwnd)                               ; A: not a search window key
            eq("other keys go on to the box", SearchKeys._swallowChar, 0)
            SearchKeys.Unhook()
            tab()
            eq("unhooked", SearchWindow.Selected "|" SearchKeys._hooked, "1|0")
        } finally {
            SearchKeys.Unhook()
            AppSettings.General["TabKey"] := saved.TabKey
            SearchWindow.DefineProp("_RunSearch", savedSearch), SearchWindow.DefineProp("_Repaint", savedRepaint)
            SearchWindow.Input := saved.Input, SearchWindow.List := saved.List, SearchWindow.Gui := saved.Gui, SearchWindow.Results := saved.Results
            SearchWindow.Selected := saved.Selected, SearchWindow.Offset := saved.Offset, SearchWindow.Mode := saved.Mode
            g.Destroy()
        }
    }

    ; Bottom status bar: the tip / state on the left, the selected result's Enter action on the right; the window height includes it; → ignores auto-repeat
    static StatusBar() {
        eq := (n, a, e) => TestRunner.Equal("StatusBar." n, a, e)
        eq("on by default", AppSettings.Defaults()["Appearance"]["StatusBar"], 1)
        fileItem := ResultItem("Notepad", "", {Kind: "file", Arg: "C:\Windows\notepad.exe"})
        eq("file: open", SearchStatusBar.EnterTitle(fileItem), "Open")
        eq("folder: open", SearchStatusBar.EnterTitle(ResultItem("Docs", "", {Kind: "folder"})), "Open")
        eq("url: open", SearchStatusBar.EnterTitle(ResultItem("Web", "", {Kind: "url"})), "Open")
        eq("text: copy", SearchStatusBar.EnterTitle(ResultItem("42", "", {Kind: "text"})), "Copy")
        eq("text with its own action: paste", SearchStatusBar.EnterTitle(ResultItem("snippet", "", {Kind: "text", OnRun: (*) => 0})), "Paste")
        eq("command: run", SearchStatusBar.EnterTitle(ResultItem("Lock", "", {OnRun: (*) => 0})), "Run")
        eq("own name", SearchStatusBar.EnterTitle(ResultItem("img", "", {Kind: "text", OnRun: (*) => 0, RunTitle: "Paste Image"})), "Paste Image")
        eq("hint row: complete", SearchStatusBar.EnterTitle(ResultItem("g", "", {Valid: false, AutoComplete: "g "})), "Complete")
        eq("hint row without completion", SearchStatusBar.EnterTitle(ResultItem("x", "", {Valid: false})), "")

        saved := {Input: SearchWindow.Input, Gui: SearchWindow.Gui, Results: SearchWindow.Results, Selected: SearchWindow.Selected
            , Mode: SearchWindow.Mode, FileMode: SearchWindow.FileMode, Tip: SearchWindow._tip, Status: SearchStatusBar.Height}
        savedOpen := SearchWindow.GetOwnPropDesc("_OpenActions")
        g := Gui()
        SearchWindow.Gui := g, SearchWindow.Input := g.AddEdit("w200")
        try {
            SearchWindow.Mode := "results", SearchWindow.FileMode := false, SearchWindow._tip := "Tip: ? cheat sheet"
            SearchWindow.Input.Value := "", SearchWindow.Results := [], SearchWindow.Selected := 0
            texts := SearchStatusBar.Texts()
            eq("empty box: the tip is in the box, not the status bar", texts.Left "|" texts.Enter "|" texts.Actions, "||0")
            SearchWindow.Input.Value := "note", SearchWindow.Results := [fileItem, ResultItem("b")], SearchWindow.Selected := 1
            texts := SearchStatusBar.Texts()
            eq("results: enter, ctrl+enter, edit, actions", texts.Left "|" texts.Enter "|" texts.Ctrl "|" texts.Edit "|" texts.Actions, "|Open|Reveal|1|1")
            parts := SearchStatusBar.Parts(texts), shown := ""
            for part in parts {
                keys := ""
                for key in part.Keys
                    keys .= (keys = "" ? "" : " ") key
                shown .= (shown = "" ? "" : " | ") part.Label " " keys
            }
            eq("four parts", shown, "Open ↵ | Reveal Ctrl ↵ | Edit F3 | Actions →")
            ; A tip too long for the box ends with "…" (one character per unit here)
            eq("fit: short tip unchanged", SearchWindow.FitWidth("Tip: ? cheat sheet", 30, (t) => StrLen(t)), "Tip: ? cheat sheet")
            eq("fit: long tip shortened", SearchWindow.FitWidth("Tip: folder + name  search folders only", 20, (t) => StrLen(t)), "Tip: folder + name…")
            eq("ctrl+enter: text pastes", SearchStatusBar.CtrlEnterTitle(ResultItem("42", "", {Kind: "text"}), "Copy"), "Paste")
            eq("ctrl+enter: same as enter is not shown", SearchStatusBar.CtrlEnterTitle(ResultItem("s", "", {Kind: "text", OnRun: (*) => 0}), "Paste"), "")
            eq("ctrl+enter: url copies", SearchStatusBar.CtrlEnterTitle(ResultItem("w", "", {Kind: "url"}), "Open"), "Copy")
            eq("ctrl+enter: commands have none", SearchStatusBar.CtrlEnterTitle(ResultItem("Lock", "", {OnRun: (*) => 0}), "Run"), "")
            items := SearchStatusBar.MenuItems(), labels := ""
            for entry in items
                labels .= (IsObject(entry) ? StrSplit(entry[1], "`t")[1] : "-") ";"
            eq("menu", labels, "Preferences...;Shortcuts and Syntax;-;Rebuild Index;Check for Updates;-;Reload;Quit;")
            SearchWindow.FileMode := true, SearchWindow.Input.Value := ""
            eq("fileItem mode", SearchStatusBar.Texts().Left, I18n.T("Status.FileMode"))
            SearchWindow.FileMode := false, SearchWindow.Mode := "actions"
            texts := SearchStatusBar.Texts()
            eq("action panel", texts.Left "|" texts.Enter "|" texts.Actions, "Esc  Back|Run|0")
            SearchWindow.Mode := "results"
            eq("mix colors", Format("{:06X}", SearchStatusBar.MixColor("FFFFFF", "000000", 0.5)), "808080")

            SearchStatusBar.Height := 0
            plain := SearchWindow._WindowHeight(3)
            SearchStatusBar.Height := 26
            eq("window height includes the status bar", SearchWindow._WindowHeight(3) - plain, 26)
            eq("no status bar without result rows", SearchWindow._WindowHeight(0) - SearchWindow._ContentHeight(0), 0)

            opened := {Count: 0}
            SearchWindow.DefineProp("_OpenActions", {Call: (*) => (opened.Count++, true)})
            hwnd := SearchWindow.Input.Hwnd
            SearchWindow.Input.Value := "note", SendMessage(0xB1, 4, 4, hwnd)   ; caret at the end
            SearchKeys.OnKeyDown(0x27, 0x40000000 | 1, 0x100, hwnd)          ; auto-repeat from holding →
            eq("held right arrow does not open actions", opened.Count, 0)
            SearchKeys.OnKeyDown(0x27, 1, 0x100, hwnd)                       ; a fresh →
            eq("fresh right arrow opens actions", opened.Count, 1)
            SearchWindow.Input.Value := "1250+500", SendMessage(0xB1, 0, -1, hwnd)   ; the restored last search: all selected
            eq("right arrow on selected text: no actions", (SearchKeys.OnKeyDown(0x27, 1, 0x100, hwnd), opened.Count), 1)
            SendMessage(0xB1, 8, 8, hwnd)                                     ; the Edit handles → itself: deselects and moves the caret to the end
            eq("right arrow again opens actions", (SearchKeys.OnKeyDown(0x27, 1, 0x100, hwnd), opened.Count), 2)
        } finally {
            SearchWindow.DefineProp("_OpenActions", savedOpen)
            SearchWindow.Input := saved.Input, SearchWindow.Gui := saved.Gui, SearchWindow.Results := saved.Results, SearchWindow.Selected := saved.Selected
            SearchWindow.Mode := saved.Mode, SearchWindow.FileMode := saved.FileMode, SearchWindow._tip := saved.Tip, SearchStatusBar.Height := saved.Status
            g.Destroy()
        }
    }

    ; At the narrowest window (450 px) the four parts on the right still leave room for the ☰ button, in every language
    static StatusBarFit() {
        savedLang := I18n.Lang, savedGdi := ResultRows.Gdi.Clone()
        ResultRows.CreateGdi()
        pad := Win.Scale(ThemeManager.Get("Padding"))
        try {
            for language in I18n.Languages() {
                I18n.Init(language[1])
                parts := SearchStatusBar.Parts({Enter: I18n.T("Status.Open"), Ctrl: I18n.T("Status.Reveal"), Edit: 1, Actions: 1})
                used := SearchStatusBar.PartsWidth(parts) + pad + SearchStatusBar.MenuWidthFor(pad)         ; right padding + the ☰ button
                TestRunner.True("StatusBarFit." language[1] " (" used " of " Win.Scale(PreferencesWindow.MinWindowWidth) ")", used <= Win.Scale(PreferencesWindow.MinWindowWidth))
            }
        } finally {
            I18n.Init(savedLang)
            ResultRows.Gdi := savedGdi
        }
    }

    ; Characters typed while the window is showing: added to the search box afterwards; Enter runs after the search, Esc closes
    ; The empty search box prepared after hiding: no search needed when showing; not used once the box changed or it was used once; the refresh after showing swaps in only when the results changed
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
            SearchWindow.PrepareEmpty()                                     ; the window is hidden: clear the box and search once
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

    ; Restoring the last search when shown: files aren't searched first (the window shows at once), then a full search after showing selects last time's row; not once typing has started
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
            SearchWindow._OnInputChange()                                   ; a Change caused by the text filled in by the program: no extra search
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
            SearchWindow._quietText := "", SearchWindow.Input.Value := "repo"   ; typed right after showing
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
            SendMessage(0xB1, 0, -1, hwnd)                                  ; the restored last search is all selected
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
            eq("no hook: nothing", TypeAhead.Take().Text, "")
            eq("log: count only", TypeAhead.Summary({Text: "secret", Key: "", Modifier: ""}), "6 chars")
            eq("log: leading space and key", TypeAhead.Summary({Text: " ab", Key: "Enter", Modifier: ""}), "3 chars (leading space) + Enter")
        } finally {
            SearchWindow.FileMode := false
            options["SpacePrefix"] := savedOptions[1], options["Enabled"] := savedOptions[2]
            for name, desc in saved
                SearchWindow.DefineProp(name, desc)
            SearchWindow.Input := savedInput, SearchWindow.Gui := savedGui
            g.Destroy()
        }
        ; a real InputHook: simulated typing is caught and kept, and no longer caught after stopping
        TypeAhead._hook := ""
        ih := InputHook("T3")
        ih.KeyOpt("{Enter}{NumpadEnter}{Esc}", "ES")
        ih.Start(), TypeAhead._hook := ih
        SendEvent("{Text}ab")
        SendEvent("{Enter}")
        typed := TypeAhead.Take()
        eq("hook collects text", typed.Text "|" typed.Key, "ab|Enter")
        eq("hook stopped", ih.InProgress "|" (TypeAhead._hook = "" ? 1 : 0), "0|1")
    }

    ; Shown again while still open (e.g. ALTRun.exe started again by another program, 2026-10 appinn): the text, mode
    ; and caret stay, nothing is searched again
    static ShowWhileOpen() {
        eq := (n, a, e) => TestRunner.Equal("ShowWhileOpen." n, a, e)
        savedSet := SearchWindow.GetOwnPropDesc("_SetInput")
        savedInput := SearchWindow.Input, savedGui := SearchWindow.Gui, savedMode := SearchWindow.Mode
        calls := {Text: ""}
        SearchWindow.DefineProp("_SetInput", {Call: (this, text) => calls.Text .= "set:" text ";"})
        g := Gui("+ToolWindow")
        SearchWindow.Gui := g, SearchWindow.Input := g.AddEdit("w200")
        try {
            g.Show("NoActivate w220 h40")
            SearchWindow.Input.Value := "half typed", SearchWindow.Mode := "actions"
            SendMessage(0xB1, 4, 4, SearchWindow.Input.Hwnd)
            SearchWindow._Show("")
            selection := SendMessage(0xB0, 0, 0, SearchWindow.Input.Hwnd)
            eq("text kept", SearchWindow.Input.Value, "half typed")
            eq("mode kept", SearchWindow.Mode, "actions")
            eq("caret kept", (selection & 0xFFFF) "-" ((selection >> 16) & 0xFFFF), "4-4")
            eq("not searched again", calls.Text, "")
            eq("still visible", SearchWindow.IsVisible() ? 1 : 0, 1)
        } finally {
            SearchWindow.DefineProp("_SetInput", savedSet)
            SearchWindow.Input := savedInput, SearchWindow.Gui := savedGui, SearchWindow.Mode := savedMode
            g.Destroy()
        }
    }

    ; Debug log for a search box that seems to empty itself (2026-10, appinn): counts and program names, never the text
    static EmptiedLog() {
        eq := (n, a, e) => TestRunner.Equal("EmptiedLog." n, a, e)
        eq("show: hidden", SearchWindow.ShowNote(false, false, ""), "")
        eq("show: visible, active", SearchWindow.ShowNote(true, true, "x.exe X"), " (already visible, active)")
        eq("show: visible, not active", SearchWindow.ShowNote(true, false, "ime.exe Cand"), " (already visible, not active; foreground: ime.exe Cand)")
        eq("foreground text", RegExMatch(SearchWindow.ForegroundText(), "^(none|\S+\.exe .+)$") > 0, true)
        saved := Map()
        for name in ["_RunSearch", "IsVisible"]
            saved[name] := SearchWindow.GetOwnPropDesc(name)
        savedInput := SearchWindow.Input, savedGui := SearchWindow.Gui, savedMode := SearchWindow.Mode
        savedList := SearchWindow.List, savedRecent := SearchWindow._recent.Clone()
        savedLog := [Logger.Enabled, Logger.File, Logger.Immediate]
        SearchWindow.DefineProp("_RunSearch", {Call: (*) => ""})
        SearchWindow.DefineProp("IsVisible", {Call: (*) => true})
        g := Gui()
        SearchWindow.Gui := g, SearchWindow.Input := g.AddEdit("w200"), SearchWindow.List := g.AddListView("w200", ["Title"])
        SearchWindow.Mode := "results"
        Logger.Enabled := true, Logger.Immediate := true, Logger.File := A_Temp "\ALTRunEmptiedTest.log"
        try FileDelete(Logger.File)
        try {
            SearchWindow._SetInput("secret words")
            SearchWindow._SetInput("")
            SearchWindow._SetInput("a")
            SearchWindow._SetInput("")                                      ; one character: not worth a line
            SearchWindow._SetInput("hello")
            SearchWindow.Input.Value := ""
            SearchWindow._OnInputChange()
            SetTimer(SearchWindow._searchTimer, 0)                          ; its search must not run in a later test
            Logger.Flush()
            logText := FileExist(Logger.File) ? FileRead(Logger.File, "UTF-8") : ""
            eq("by ALTRun", InStr(logText, "SearchWindow: input emptied (12 chars, by ALTRun, mode results)") > 0, true)
            eq("by an edit", InStr(logText, "SearchWindow: input emptied (5 chars, by an edit)") > 0, true)
            eq("one character skipped", InStr(logText, "(1 chars"), 0)
            eq("text never logged", InStr(logText, "secret") || InStr(logText, "hello"), 0)
            SearchWindow.Input.Value := "secret words", SearchWindow._NoteResults(3)
            eq("recent results: length only", RegExMatch(SearchWindow._recent[-1], "input 12 chars, mode results results 3 ") > 0
                && !InStr(SearchWindow._recent[-1], "secret"), true)
        } finally {
            try FileDelete(Logger.File)
            Logger.Enabled := savedLog[1], Logger.File := savedLog[2], Logger.Immediate := savedLog[3]
            for name, desc in saved
                SearchWindow.DefineProp(name, desc)
            SearchWindow.Input := savedInput, SearchWindow.Gui := savedGui, SearchWindow.Mode := savedMode, SearchWindow._inputLen := 0
            SearchWindow.List := savedList, SearchWindow._recent := savedRecent
            g.Destroy()
        }
    }

    ; The list finished painting but the custom-draw callback didn't run (blank rows): log and paint again; a callback that ran, or nothing painted yet, doesn't count
    ; Fallback font warm-up: CJK text measured once in each result font on a memory DC (see ResultRows.WarmUpFonts)
    static FontWarmUp() {
        eq := (n, a, e) => TestRunner.Equal("FontWarmUp." n, a, e)
        saved := [ResultRows.Gdi, Logger.Enabled, Logger.File, Logger.Immediate]
        Logger.Enabled := true, Logger.Immediate := true, Logger.File := A_Temp "\ALTRunFontTest.log"
        try FileDelete(Logger.File)
        created := []
        try {
            ResultRows.Gdi := Map()
            ResultRows.WarmUpFonts()                                        ; fonts not created yet: do nothing
            eq("no fonts yet: nothing done", FileExist(Logger.File) ? 1 : 0, 0)
            for key in ["TitleFont", "SubtitleFont"]
                created.Push(ResultRows.Gdi[key] := ResultRows._CreateFont("Segoe UI", 12, 400))
            ResultRows.WarmUpFonts()
            logText := FileExist(Logger.File) ? FileRead(Logger.File, "UTF-8") : ""
            TestRunner.True("FontWarmUp.logged", InStr(logText, "Perf: font warm-up") > 0)
            text := ResultRows.FontWarmUpText
            eq("Chinese, Japanese and Korean", (RegExMatch(text, "[\x{4E00}-\x{9FFF}]") && RegExMatch(text, "[\x{3040}-\x{30FF}]")
                && RegExMatch(text, "[\x{AC00}-\x{D7AF}]")) ? 1 : 0, 1)
        } finally {
            for font in created
                DllCall("DeleteObject", "Ptr", font)
            ResultRows.Gdi := saved[1], Logger.Enabled := saved[2], Logger.File := saved[3], Logger.Immediate := saved[4]
            try FileDelete(A_Temp "\ALTRunFontTest.log")
        }
    }

    ; The custom-draw callback isn't interruptible (see ResultRows.OnCustomDraw): Critical stays on in this thread until it ends, so check after the call and then turn it off
    static CustomDrawCritical() {
        nmcd := Buffer(A_PtrSize * 4 + 64, 0)
        NumPut("UInt", 0x1, nmcd, A_PtrSize * 3)                            ; CDDS_PREPAINT
        saved := SearchWindow._paints
        Critical("Off")
        result := ResultRows.OnCustomDraw(nmcd.Ptr)
        wasCritical := A_IsCritical
        Critical("Off")
        SearchWindow._paints := saved
        TestRunner.Equal("CustomDrawCritical.prepaint result", result, 0x20)   ; CDRF_NOTIFYITEMDRAW
        TestRunner.True("CustomDrawCritical.critical", wasCritical != 0)
    }

    ; What the CapsLock Chinese / English toggle does (Win.ImeTarget), including IMEs whose reading doesn't follow (#154)
    ; Prefetching common icons after start-up: picked ones first by count; duplicates removed, cached ones skipped, at most PrefetchMax; nothing read while the window shows
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

        savedMax := IconCache.PrefetchMax, savedBusy := BackgroundWork.IsBusy
        cached := "ext:.prefetchcached"
        IconCache.GetNow(cached)
        BackgroundWork.IsBusy := () => true                                 ; the window is shown
        IconCache.PrefetchMax := 3
        IconCache.Prefetch(["ext:.pf1", "", "ext:.PF1", cached, "ext:.pf2", "ext:.pf3", "ext:.pf4"])
        eq("dedupe, skip cached, limit", IconCache.PrefetchPending, 3)
        IconCache._PrefetchStep()
        eq("waits while the window is shown", IconCache.PrefetchPending, 3)
        BackgroundWork.IsBusy := () => false
        start := A_TickCount
        while (IconCache.PrefetchPending && A_TickCount - start < 3000)
            IconCache._PrefetchStep()
        eq("all loaded", IconCache.PrefetchPending, 0)
        eq("in the cache", IconCache._icons.Has(IconCache._CacheKey("ext:.pf3")), true)
        eq("over the limit not loaded", IconCache._icons.Has(IconCache._CacheKey("ext:.pf4")), false)
        IconCache.Prefetch([])
        eq("empty list", IconCache.PrefetchPending, 0)
        IconCache.PrefetchMax := savedMax, BackgroundWork.IsBusy := savedBusy

        savedShowing := SearchWindow._showing                               ; not while half shown either (the window isn't up yet)
        SearchWindow._showing := true
        eq("not while showing", SearchWindow._AllowsBackgroundWork(), false)
        SearchWindow._showing := false
        eq("hidden and not showing", SearchWindow._AllowsBackgroundWork(), !SearchWindow.IsVisible())
        SearchWindow._showing := savedShowing
    }

    ; Icon disk cache: pixel round trips, file format (keeping the most recently used, ignoring a damaged file), showing straight from the cache, expiry
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
            ; written in chunks (many chunks when each is tiny), same content; no temporary file left behind
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

    ; Reinstalling the keyboard hook after waking, unlocking or a monitor change: once for several events in a row; the mouse hook too when there are mouse hotkeys
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
        Sleep(150)                                                          ; the timer runs ReinstallHooks (without errors)
        App.WakeDelay := savedDelay
        App.ReinstallHooks("test")
        eq("hook still works", A_IsSuspended, 0)
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

    ; Fading: the opacity curve; a real window is transparent first when fading in and ends at the theme's opacity; fading out leaves a shadow that disappears at the end
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
        WindowFade.Out(g.Hwnd)                                              ; hidden before the fade in finished: the window gets its opacity back and the shadow takes over
        eq("out: window opaque again", WinGetTransparent("ahk_id " g.Hwnd), 255)
        eq("out: ghost shown", WindowFade.GhostVisible, true)
        g.Hide()
        start := A_TickCount
        while (WindowFade.Running && A_TickCount - start < 2000)
            Sleep(10)
        eq("out: finished by itself", WindowFade.Running, false)
        eq("out: ghost gone", WindowFade.GhostVisible, false)

        WindowFade.Out(g.Hwnd)                                              ; the window isn't on screen: no shadow
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
        ; a trustworthy reading: the reading after the toggle differs from this one, so go by the reading
        eq("reliable readings", Win.ImeTarget(1, 1, {Open: 0, Conv: 0, Chinese: false}), "en open 0 conv 0")
        ; WeChat IME (log of #154): the reading follows the setting and is always open; last switched to Chinese -> English this time (close the IME)
        eq("wechat ime: to english", Win.ImeTarget(1, 0x401, {Open: 1, Conv: 0x401, Chinese: true}), "en open 0 conv 1024 remembered")
        ; a reading that doesn't follow: always "not open / 0"; last switched to Chinese -> back to English this time, then Chinese again next time
        eq("stuck readings: back to english", Win.ImeTarget(0, 0, {Open: 0, Conv: 0, Chinese: true}), "en open 0 conv 0 remembered")
        eq("stuck readings: chinese again", Win.ImeTarget(0, 0, {Open: 0, Conv: 0, Chinese: false}), "zh open 1 conv 1 remembered")
        ; Order of the settings: when opening, open first then set the conversion mode; when closing, set the mode first then close (6 = IMC_SETOPENSTATUS, 2 = IMC_SETCONVERSIONMODE)
        steps(open, conv) {
            list := ""
            for step in Win.ImeSteps(open, conv)
                list .= " " step[1] ":" step[2]
            return Trim(list)
        }
        TestRunner.Equal("ImeToggle.steps: open", steps(1, 0x401), "6:1 2:1025")
        TestRunner.Equal("ImeToggle.steps: close", steps(0, 0x400), "2:1024 6:0")
        ; Switching to English mode when shown (SwitchToEnglishInput = 2): remove the native bit, close the IME, other bits unchanged
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
        list := g.AddListView("w200 h100 -Hdr", [""])                       ; no custom-draw callback connected: every paint is as if the callback was skipped
        list.Add("", "")
        g.Show("NA x0 y0")
        SearchWindow.Gui := g, SearchWindow.List := list, SearchWindow.Input := input, SearchWindow.Results := []
        Logger.Enabled := true, Logger.Immediate := true, Logger.File := A_Temp "\ALTRunPaintTest.log"
        try FileDelete(Logger.File)
        pending := () => DllCall("GetUpdateRect", "Ptr", list.Hwnd, "Ptr", 0, "Int", 0)
        paint := () => (DllCall("InvalidateRect", "Ptr", list.Hwnd, "Ptr", 0, "Int", 0), DllCall("UpdateWindow", "Ptr", list.Hwnd))
        try {
            SearchWindow._skipLogged := false
            SearchWindow._WatchPaint(), paint(), SearchWindow._paints++     ; the callback ran
            SearchWindow._CheckPaint()
            eq("drawn: no repaint", pending() ? 1 : 0, 0)
            DllCall("InvalidateRect", "Ptr", list.Hwnd, "Ptr", 0, "Int", 0)
            SearchWindow._WatchPaint(), SearchWindow._CheckPaint()          ; nothing painted yet
            eq("not painted yet: wait", SearchWindow._paintTries "|" SearchWindow._paintRetries, "1|0")
            DllCall("UpdateWindow", "Ptr", list.Hwnd)
            Critical(1000)                                                  ; no messages handled while running (with a slow log write, the repaint request would be handled first)
            SearchWindow._WatchPaint(), paint(), SearchWindow._CheckPaint()  ; painted, callback didn't run
            eq("skipped: repaint", (pending() ? 1 : 0) "|" SearchWindow._paintRetries, "1|1")
            Critical("Off")
            logText := FileExist(Logger.File) ? FileRead(Logger.File, "UTF-8") : ""
            TestRunner.True("SkippedPaint.logged", InStr(logText, "[WRN] SearchWindow: custom draw skipped") > 0)
            DllCall("UpdateWindow", "Ptr", list.Hwnd)
            SearchWindow._paintRetries := 3, SearchWindow._CheckPaint()     ; at most 3 extra paints
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

    ; Notices after an action: below the search window, centered, while it is open, otherwise low in the middle of the screen; never off screen
    ; Showing the search window doesn't re-enter, but a previous showing stuck for over StuckShowMs no longer blocks a new one (stuck for 2.5 min after a monitor was turned off)
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
        SearchWindow.LastArea := second, SearchWindow.HiddenAt := A_TickCount                          ; the search window just hid on the second screen
        eq("just hidden: same screen", Hud._Anchor().Area.Left, 1920)
        SearchWindow.HiddenAt := A_TickCount - Hud.RecentWindowMs - 1
        AppSettings.Appearance["ShowOn"] := "Primary"
        primary := Win.WorkArea(MonitorGetPrimary())
        eq("later: ShowOn primary", Hud._Anchor().Area.Left "," Hud._Anchor().Area.Top, primary.Left "," primary.Top)
        SearchWindow.LastArea := savedArea, SearchWindow.HiddenAt := savedAt, AppSettings.Appearance["ShowOn"] := savedShowOn

        long := ""
        Loop 40
            long .= "long text "
        Hud.Show(long, 300)                                                 ; wider than the maximum: wraps without an error
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
}
