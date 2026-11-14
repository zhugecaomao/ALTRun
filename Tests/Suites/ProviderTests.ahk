;===============================================================================
; ProviderTests.ahk - Result providers: web search, clipboard, file index and file search,
;   applications, commands, windows, recent items, scripts, bookmarks, browser history and snippets search
;-------------------------------------------------------------------------------
; Included by Tests\RunTests.ahk. Every static method whose name doesn't start with "_" is a test;
; TestRunner runs them in name order, and each one must put back the shared state it changes.
;===============================================================================
class ProviderTests {
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

        ; {query|default}: with only the keyword typed the default is used, so it opens directly
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
        ; {query|clipboard}: with only the keyword typed the text on the clipboard is searched
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
        ; text with an image (cells copied from Excel): the text is used; marked "don't record" by a password manager: not read; copied files: not read
        putText("cell text", ClipboardData.CF_BITMAP)
        eq("clipboard: text with a bitmap", WebSearchProvider.ClipboardText(), "cell text")
        putText("secret", DllCall("RegisterClipboardFormat", "Str", "ExcludeClipboardContentFromMonitorProcessing", "UInt"))
        eq("clipboard: private", WebSearchProvider.ClipboardText(), "")
        ClipboardData.SetFiles([A_ScriptDir])
        eq("clipboard: files", WebSearchProvider.ClipboardText(), "")
        A_Clipboard := saved
        ; Built-in engines' old default URLs in existing settings files become {query|clipboard}; changed and user-added ones are left alone
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
        ; Icon: the browser's icon when not given / the file doesn't exist; images as thumbnails; relative paths start from the Data folder
        eq("icon: none", WebSearchProvider.IconFor(engine), "url:")

        ; puts text and another format on the clipboard (the format's content: a 1x1 bitmap for bitmap formats, a 0 otherwise)
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

        ; built-in engines have their own icons (Resources\Icons\Web\<Id>.png)
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

    ; Downloading site icons: the icons a page names, file header checks, stored in Data\Icons (a fake fetch stands in for the network)
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
            , "https://www.example.com/sub/img/icon-64.png", {Status: 200, Body: text, Size: 16, Url: ""}          ; not an image
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

    ; A saved history that can't be read at start-up is not an empty one: nothing is saved over it and none of its files are deleted;
    ; once it can be read, its entries come back after this run's. A file that doesn't parse is kept as .bad
    static _CountFiles(pattern) {
        count := 0
        Loop Files, pattern
            count++
        return count
    }

    static ClipboardUnread() {
        eq := (n, a, e) => TestRunner.Equal("ClipboardUnread." n, a, e)
        saved := {File: ClipboardProvider.File, Folder: ClipboardProvider.Folder, Entries: ClipboardProvider.Entries, Unread: ClipboardProvider._unread
            , Persist: AppSettings.Feature("Clipboard")["Persist"], Read: SafeFile.readFn, Wait: SafeFile.ReadWaitMs}
        root := A_Temp "\ALTRun-test-clip-unread"
        try DirDelete(root, true)
        DirCreate(root)
        fail := {On: false}
        try {
            ClipboardProvider.File := root "\ClipboardHistory.json", ClipboardProvider.Folder := root "\Clipboard"
            AppSettings.Feature("Clipboard")["Persist"] := 1
            long := ""
            Loop 500
                long .= "line " A_Index "`r`n"
            ClipboardProvider._unread := false
            ClipboardProvider.Entries := [Map("Text", "old pinned", "Time", "20260929010101", "App", "a.exe", "Pinned", 1), Map("Text", long, "Time", "20260929010102", "App", "b.exe")
                , Map("Text", "both", "Time", "20260929010103", "App", "c.exe", "Pinned", 1)]
            ClipboardProvider.Save()
            onDisk := FileRead(ClipboardProvider.File, "UTF-8")
            longFiles := () => ProviderTests._CountFiles(root "\Clipboard\*.txt")

            SafeFile.ReadWaitMs := 0
            SafeFile.readFn := (file, encoding) => fail.On ? CoreTests._ReadFails() : FileRead(file, encoding)
            fail.On := true
            ClipboardProvider.Entries := [], ClipboardProvider._Load()
            eq("unreadable: noted, nothing loaded", ClipboardProvider._unread "|" ClipboardProvider.Entries.Length, "1|0")
            ClipboardProvider.Entries := [Map("Text", "new one", "Time", A_Now, "App", "d.exe"), Map("Text", "both", "Time", A_Now, "App", "d.exe")]
            ClipboardProvider.Save()
            eq("still unreadable: file and long entry untouched", (FileRead(ClipboardProvider.File, "UTF-8") == onDisk) "|" longFiles(), "1|1")

            fail.On := false                                                ; readable again: merged before the next save
            ClipboardProvider.Save()
            texts := ""
            for entry in ClipboardProvider.Entries
                texts .= (StrLen(entry["Text"]) > 50 ? "long" : entry["Text"]) (ClipboardProvider.IsPinned(entry) ? "*" : "") "|"
            eq("readable again: this run's first, then the saved ones, the same text once", texts, "new one|both*|old pinned*|long|")
            eq("readable again: saved, long entry kept", ClipboardProvider._unread "|" longFiles(), "0|1")
            ClipboardProvider.Entries := [], ClipboardProvider._Load()
            eq("reloaded", ClipboardProvider.Entries.Length, 4)

            FileDelete(ClipboardProvider.File), FileAppend("{broken", ClipboardProvider.File, "UTF-8-RAW")   ; doesn't parse: kept as .bad
            ClipboardProvider.Entries := [], ClipboardProvider._Load()
            eq("invalid: kept as .bad, empty history", ClipboardProvider.Entries.Length "|" ClipboardProvider._unread "|" FileRead(ClipboardProvider.File ".bad", "UTF-8"), "0|0|{broken")
        } finally {
            SafeFile.readFn := saved.Read, SafeFile.ReadWaitMs := saved.Wait
            ClipboardProvider.File := saved.File, ClipboardProvider.Folder := saved.Folder, ClipboardProvider.Entries := saved.Entries, ClipboardProvider._unread := saved.Unread
            AppSettings.Feature("Clipboard")["Persist"] := saved.Persist
            try DirDelete(root, true)
        }
    }

    static Clipboard() {
        eq := (n, a, e) => TestRunner.Equal("Clipboard." n, a, e)
        ; saving: very long entries are stored as separate files and read back the same; files not in the history are deleted
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
        ClipboardProvider.Add("first entry", "notepad.exe")                 ; duplicate content moves to the top
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
        hint := ClipboardProvider.Search(SearchQuery("cli"))                ; only the start of the keyword: its hint at the end of the list
        eq("keyword start: hint", hint.Length "|" hint[1].Valid "|" hint[1].AutoComplete "|" hint[1].Score, "1|0|clip |" ResultItem.KeywordPrefixHintScore)
        eq("two letters: nothing", ClipboardProvider.Search(SearchQuery("cl")).Length, 0)
        mixed := [ResultItem("a", "", {Exclusive: true}), ResultItem("b")]
        eq("registry keeps exclusive", ProviderRegistry._KeepExclusive(mixed).Length, 1)
        AppSettings.Feature("Clipboard")["MaxItems"] := 2
        ClipboardProvider.Add("third", "")
        eq("max items", ClipboardProvider.Entries.Length, 2)
        AppSettings.Feature("Clipboard")["MaxItems"] := 200
        ClipboardProvider.Remove("third")
        eq("remove", ClipboardProvider.Entries[1]["Text"], "first entry")
        ClipboardData.Borrow(10000)
        ClipboardProvider._OnChange(1)
        eq("paused", ClipboardProvider.Entries.Length, 1)
        ClipboardProvider.Entries := []
        ; After the clipboard changes, read after ReadDelay (without holding the clipboard); if it changes again while waiting, read the previous one first
        saved := ClipboardAll(), ClipboardData.BorrowedUntil := 0
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
        ClipboardData.Borrow(1000)                              ; e.g. pasting a snippet: the clipboard gets temporary content right away
        Sleep(ClipboardProvider.ReadDelay + 150)
        eq("not read while ALTRun borrows it", ClipboardProvider.Entries.Length, 0)
        ClipboardData.BorrowedUntil := 0
        A_Clipboard := saved
        Sleep(ClipboardProvider.ReadDelay + 150)
        ClipboardProvider.Entries := []
    }

    ; Files and images in the clipboard history, merging a double Ctrl+C
    ; Thumbnails in the clipboard history: images and one copied image file show a thumbnail, filling the square (FillThumbnails) or the whole image;
    ; a wide image filled covers the corners, fitted it leaves them transparent; big files go through the shell's thumbnails (falling back to GDI+)
    static ClipboardThumbnails() {
        eq := (n, a, e) => TestRunner.Equal("ClipboardThumbnails." n, a, e)
        ok := (n, v) => TestRunner.True("ClipboardThumbnails." n, v)
        options := AppSettings.Feature("Clipboard")
        saved := [options["FillThumbnails"], IconCache.ShellThumbBytes]
        root := A_Temp "\ALTRun-test-clip-thumbs"
        try DirDelete(root, true)
        DirCreate(root)
        try {
            eq("default fill", options["FillThumbnails"], 1)
            eq("image fill", ClipboardProvider.Thumbnail("C:\x.png"), "thumbfill:C:\x.png")
            eq("image file", ClipboardProvider.FileIcon("C:\Cache\a.JPG"), "thumbfill:C:\Cache\a.JPG")
            eq("webp file", ClipboardProvider.FileIcon("C:\Cache\a.webp"), "thumbfill:C:\Cache\a.webp")
            eq("other file", ClipboardProvider.FileIcon("C:\x.txt"), "C:\x.txt")
            eq("network image file", ClipboardProvider.FileIcon("\\server\share\a.png"), "\\server\share\a.png")
            options["FillThumbnails"] := 0
            eq("image fit", ClipboardProvider.Thumbnail("C:\x.png"), "thumb:C:\x.png")
            eq("image file fit", ClipboardProvider.FileIcon("C:\a.png"), "thumb:C:\a.png")
            eq("cache key", IconCache._CacheKey("thumbfill:C:\A.png"), "thumbfill:c:\a.png")
            ok("not a file key", !IconCache.IsFileKey("thumbfill:c:\a.png"))

            ; a 64 x 16 opaque image: fitted, the top-left corner of the square stays transparent; filled, it is covered.
            ; Wine's GDI+ can't write or read image files (ClipboardKinds.thumbnail icon fails there too): Windows only
            isWine := DllCall("GetProcAddress", "Ptr", DllCall("GetModuleHandle", "Str", "ntdll", "Ptr"), "AStr", "wine_get_version", "Ptr")
            if !isWine {
                ok("gdiplus", ClipboardData.StartGdiplus())
                wide := root "\wide.png", bitmap := 0, graphics := 0
                DllCall("gdiplus\GdipCreateBitmapFromScan0", "Int", 64, "Int", 16, "Int", 0, "Int", 0x26200A, "Ptr", 0, "Ptr*", &bitmap)
                DllCall("gdiplus\GdipGetImageGraphicsContext", "Ptr", bitmap, "Ptr*", &graphics)
                DllCall("gdiplus\GdipGraphicsClear", "Ptr", graphics, "UInt", 0xFFCC2020)
                DllCall("gdiplus\GdipDeleteGraphics", "Ptr", graphics)
                clsid := Buffer(16)
                DllCall("ole32\CLSIDFromString", "Str", "{557CF406-1A04-11D3-9A73-0000F81EF32E}", "Ptr", clsid)   ; PNG encoder
                DllCall("gdiplus\GdipSaveImageToFile", "Ptr", bitmap, "WStr", wide, "Ptr", clsid, "Ptr", 0)
                DllCall("gdiplus\GdipDisposeImage", "Ptr", bitmap)
                ok("wide image saved", FileExist(wide))
                corner := (hIcon) => (hIcon && (pixels := IconCache.IconPixels(hIcon))) ? NumGet(pixels, 3, "UChar") : -1   ; alpha of the top-left pixel
                middle := (hIcon) => (hIcon && (pixels := IconCache.IconPixels(hIcon))) ? NumGet(pixels, ((IconCache.Size // 2) * IconCache.Size + IconCache.Size // 2) * 4 + 3, "UChar") : -1
                for name, bytes in Map("gdiplus", 1 << 30, "shell", 0) {    ; "shell": every file goes through the shell's thumbnails first
                    IconCache.ShellThumbBytes := bytes
                    fit := IconCache._FromImage(wide, false), fill := IconCache._FromImage(wide, true)
                    ok(name " fit icon", fit)
                    ok(name " fill icon", fill)
                    eq(name " fit corner transparent", corner(fit), 0)
                    eq(name " fit middle opaque", middle(fit), 255)
                    eq(name " fill corner covered", corner(fill), 255)   ; solid up to the edge: no half-transparent frame from scaling
                    for hIcon in [fit, fill]
                        if hIcon
                            DllCall("DestroyIcon", "Ptr", hIcon)
                }
                IconCache.ShellThumbBytes := saved[2]
                hbm := 0                                                    ; Windows has a thumbnail handler for PNG: the shell path itself gives an image
                image := IconCache._ShellThumbnail(wide, 128, &hbm)
                ok("shell thumbnail", image && hbm)
                if image
                    DllCall("gdiplus\GdipDisposeImage", "Ptr", image)
                if hbm
                    DllCall("DeleteObject", "Ptr", hbm)
            }
            ; a file that can't be read as an image: the file's own icon instead of none
            FileAppend("not an image", root "\broken.png")
            eq("broken image", IconCache._FromImage(root "\broken.png", true), 0)
            fallback := IconCache._LoadRaw("thumbfill:" root "\broken.png")
            ok("broken image falls back to the file icon", fallback)
            if fallback
                DllCall("DestroyIcon", "Ptr", fallback)
        } finally {
            options["FillThumbnails"] := saved[1], IconCache.ShellThumbBytes := saved[2]
            try DirDelete(root, true)
        }
    }

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

        ; files: several files in one entry; copying the same files again only moves it to the top
        ClipboardProvider.Add("some text", "notepad.exe")
        ClipboardProvider.AddFiles(["C:\Reports\a.xlsx", "C:\Reports\b.docx"], "explorer.exe")
        ClipboardProvider.Add("C:\Reports\a.xlsx`r`nC:\Reports\b.docx", "")  ; the same text is another entry
        ClipboardProvider.AddFiles(["C:\Reports\a.xlsx", "C:\Reports\b.docx"], "explorer.exe")
        eq("files dedupe", ClipboardProvider.Entries.Length, 3)
        eq("files newest", ClipboardProvider.TypeOf(ClipboardProvider.Entries[1]), "files")
        items := ClipboardProvider.Search(SearchQuery("clip b.docx"))
        eq("files search", items.Length, 2)
        item := items[1], item.Provider := "Clipboard"
        eq("files title", item.Title, "a.xlsx, b.docx")
        eq("files subtitle", item.Subtitle, I18n.T("Clipboard.FilesSubtitle", ClipboardProvider._FormatTime(item.Source["Time"]), "explorer.exe", 2))
        eq("files kind", item.Kind, "text")
        actions := ResultActions.ListFor(item)
        eq("files first action", actions[1].Title, I18n.T("Clipboard.PasteFiles"))
        eq("files first hint", actions[1].Subtitle, "Enter")
        eq("files copy no enter", actions[2].Subtitle, "")
        ok("files add all", ResultActions.CanEdit(item))
        ClipboardProvider.AddFiles([A_WinDir], "explorer.exe")               ; one folder: like a file search result
        single := ClipboardProvider.Search(SearchQuery("clip"))[1]
        eq("single kind", single.Kind, "folder")
        eq("single arg", single.Arg, A_WinDir)
        eq("text at 1", ClipboardProvider.TextAt(1), "C:\Reports\a.xlsx`r`nC:\Reports\b.docx")
        eq("text at 2", ClipboardProvider.TextAt(2), "some text")
        eq("text at 3", ClipboardProvider.TextAt(3), "")

        ; images: identical images are kept once; over MaxImages the oldest is removed
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
        eq("image thumbnail", imageItem.Icon, "thumbfill:" root "\Clipboard\one.png")
        eq("image search word", ClipboardProvider.Search(SearchQuery("clip image")).Length, 2)
        ok("image not editable", !ResultActions.CanEdit(imageItem))
        ok("image deletable", ResultActions.CanDelete(imageItem))
        eq("image first action", ResultActions.ListFor(imageItem)[1].Title, I18n.T("Clipboard.PasteImage"))
        options["MaxImages"] := 1
        ClipboardProvider._Trim()
        imageCount := 0
        for entry in ClipboardProvider.Entries
            imageCount += ClipboardProvider.TypeOf(entry) = "image"
        eq("max images", imageCount, 1)

        ; save and read back; image files not in the history are deleted
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
        ClipboardProvider.DeleteItem(imageItem)                              ; the old object is gone: nothing deleted
        eq("delete stale", ClipboardProvider.Entries.Length, before)
        ClipboardProvider.RemoveEntry(ClipboardProvider.Entries[1])
        eq("delete image", ClipboardProvider.TypeOf(ClipboardProvider.Entries[1]), "files")
        FileDelete(root "\Clipboard\one.png")                               ; the image file is gone: skipped when read back
        ClipboardProvider.Entries.InsertAt(1, Map("Type", "image", "Image", "one.png", "Width", 1, "Height", 1, "Text", "", "Time", A_Now, "App", ""))
        ClipboardProvider.Save(), ClipboardProvider._Load()
        eq("missing image skipped", ClipboardProvider.TypeOf(ClipboardProvider.Entries[1]), "files")

        ; Double Ctrl+C: the same text copied a second time is appended to the previous entry
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

        ; Files (CF_HDROP) and images (bitmaps) put on the clipboard and read back
        ClipboardData.Borrow(3000)
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

    static FileIndex() {
        eq := (n, a, e) => TestRunner.Equal("FileIndex." n, a, e)
        FileIndex.SetEntries(["C:\Docs\Project Atlas", "C:\Docs\Project Atlas\Atlas Plan.dwg", "C:\Docs\notes.txt", "C:\Docs\Status Report.pdf"], [1, 0, 0, 0])
        found := FileIndex.Search("atlas", 10)
        eq("count", found.Length, 2)
        eq("prefix first", found[1].Path, "C:\Docs\Project Atlas\Atlas Plan.dwg")
        eq("narrowed", FileIndex.Search("atlas p", 10).Length, 1)
        eq("new search", FileIndex.Search("rep", 10)[1].Path, "C:\Docs\Status Report.pdf")
        TestRunner.True("FileIndex.word start >= 70", FileIndex.ScoreName("atlas", "project atlas") >= 70)
        eq("no match", FileIndex.ScoreName("xyz", "notes.txt"), 0)
        FileIndex.SetEntries([])
    }

    ; The packed index (FileIndex.Entries): names, folders and paths come back as added; Find scores like ScoreName; the cache file round-trips
    static FileIndexPacked() {
        eq := (n, a, e) => TestRunner.Equal("FileIndexPacked." n, a, e)
        ; Find == ScoreName on every name, for matches close together, far apart (the stride doubles) and at the very end
        data := FileIndex.Entries()
        Loop 3000
            data.Add("C:\Big", (Mod(A_Index, 1000) = 0 ? "far " : Mod(A_Index, 3) = 0 ? "near " : "x ") A_Index ".txt", false)
        data.Seal()
        for needle in ["near", "far", "3000.txt", "x 1", "zzz"] {
            scores := Map(), matches := [], expected := "", got := ""
            data.Find(needle, scores, matches)
            Loop data.Count
                if (score := FileIndex.ScoreName(needle, StrLower(data.NameAt(A_Index))))
                    expected .= A_Index "=" score " "
            for index in matches
                got .= index "=" scores[index] " "
            eq("find all, like ScoreName: " needle, got, expected)
        }
        paths := ["C:\Docs\Project Atlas", "C:\Docs\Project Atlas\Atlas Plan.DWG", "D:\notes.txt", "C:\Docs\Status Report (atlas).pdf"
            , "C:\Docs\atlas", "C:\Docs\测试 Atlas.docx", "C:\Docs\atlasatlas.txt"]
        FileIndex.SetEntries(paths, [1, 0, 0, 0, 1, 0, 0])
        data := FileIndex.Data
        eq("count", FileIndex.Count, 7)
        back := ""
        Loop data.Count
            back .= data.PathAt(A_Index) "|" data.IsFolderAt(A_Index) ";"
        eq("paths, case and folders kept", back, "C:\Docs\Project Atlas|1;C:\Docs\Project Atlas\Atlas Plan.DWG|0;D:\notes.txt|0;C:\Docs\Status Report (atlas).pdf|0;"
            . "C:\Docs\atlas|1;C:\Docs\测试 Atlas.docx|0;C:\Docs\atlasatlas.txt|0;")
        eq("each folder once", data.Dirs.Length, 3)
        for needle in ["atlas", "a", "plan.dwg", "notes.txt", "(", "测试", "xyz", "atlasatlas.txt"] {
            scores := Map(), matches := []
            data.Find(needle, scores, matches)
            expected := "", found := ""
            Loop data.Count
                if (score := FileIndex.ScoreName(needle, StrLower(data.NameAt(A_Index))))
                    expected .= A_Index "=" score " "
            for index in matches
                found .= index "=" scores[index] " "
            eq("Find scores like ScoreName: " needle, found, expected)
        }
        eq("an entry once (two matches in the name)", FileIndex.Search("atlas", 10).Length, 6)
        eq("extension filter", FileIndex.Search("atlas", 10, false, Map("dwg", true))[1].Path, "C:\Docs\Project Atlas\Atlas Plan.DWG")
        eq("folders only", FileIndex.Search("atlas", 10, true).Length, 2)
        savedFile := FileIndex.File, savedOld := FileIndex.OldFiles
        FileIndex.File := A_Temp "\ALTRunPackedIndex.bin", FileIndex.OldFiles := [A_Temp "\ALTRunPackedIndex.txt", A_Temp "\ALTRunPackedIndex.json"]
        try {
            FileIndex._SaveCache()
            FileIndex.SetEntries([])
            for oldFile in FileIndex.OldFiles
                FileAppend("ALTRun file index 2`n", oldFile)
            eq("cache read", FileIndex._LoadCache(), true)
            again := ""
            Loop FileIndex.Count
                again .= FileIndex.Data.PathAt(A_Index) "|" FileIndex.Data.IsFolderAt(A_Index) ";"
            eq("cache round trip", again, back)
            eq("cache round trip: each folder once", FileIndex.Data.Dirs.Length, 3)
            eq("cache round trip: search", FileIndex.Search("atlas", 10).Length "|" FileIndex.Search("plan.dwg", 10)[1].Path, "6|C:\Docs\Project Atlas\Atlas Plan.DWG")
            eq("cache round trip: pinyin initials", FileIndex.Initials.Count "|" FileIndex.Initials[6] "|" (FileIndex._initialsFor == FileIndex.Data), "1|cs atlas.docx|1")
            eq("old caches deleted", FileExist(FileIndex.OldFiles[1]) FileExist(FileIndex.OldFiles[2]), "")
            whole := FileRead(FileIndex.File, "RAW")
            cut := FileOpen(FileIndex.File, "w"), cut.RawWrite(whole, whole.Size - 2), cut.Close()
            eq("a cache cut short isn't read", FileIndex._LoadCache(), false)
            FileDelete(FileIndex.File), FileAppend("ALTRun file index 2`nD`tC:\Docs`n1`tatlas.txt`n", FileIndex.File)
            eq("a cache in another format isn't read", FileIndex._LoadCache(), false)
            FileIndex.SetEntries([])
            FileIndex._SaveCache()
            eq("an empty index round-trips", FileIndex._LoadCache() "|" FileIndex.Count "|" FileIndex.Data.Dirs.Length, "1|0|0")
        } finally {
            try FileDelete(FileIndex.File)
            for oldFile in FileIndex.OldFiles
                try FileDelete(oldFile)
            FileIndex.File := savedFile, FileIndex.OldFiles := savedOld
        }
        FileIndex.SetEntries([])
        eq("empty index", FileIndex.Search("atlas", 10).Length, 0)
    }

    ; Watching the indexed folders: which changes need a rescan, rescanning after changes, writing the cache after the scan when idle
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
        ; the default exclusions: code repositories, virtual environments, development tool caches (whole folders, case-insensitive); names that are only similar are scanned as usual
        exclude := AppSettings.Defaults()["Features"]["FileSearch"]["ScopeExclude"]
        for folder in ["node_modules", ".git", ".SVN", ".hg", "__pycache__", ".venv", "venv", ".tox", ".cache", ".idea", ".vs", "$RECYCLE.BIN"]
            eq("default exclude " folder, RegExMatch("D:\Docs\Tool\" folder "\x.txt", exclude) > 0, true)
        for folder in ["venvs", "my.venv", "cache", "Visual Studio", ".vscode", "github", "idea"]
            eq("default keeps " folder, RegExMatch("D:\Docs\" folder "\x.txt", exclude), 0)
        eq("own folder", FileIndex.IsWatchedChange("D:\Docs\ALTRun\Data\FileIndex.txt", roots, 3, "", "D:\Docs\ALTRun"), false)
        ; a ScopeExclude that isn't a valid regex (edited by hand) is ignored instead of failing every scan step
        eq("valid exclude used", FileIndex._Exclude(Map("ScopeExclude", exclude)), exclude)
        eq("bad exclude ignored", FileIndex._Exclude(Map("ScopeExclude", "i)\\(node_modules")), "")

        root := A_Temp "\ALTRunFileIndexWatch"
        try DirDelete(root, true)
        DirCreate(root)
        FileAppend("", root "\Old Plan.dwg")
        options := AppSettings.Feature("FileSearch")
        saved := {Folders: options["ScopeFolders"], Everything: options["UseEverything"], File: FileIndex.File
            , Data: FileIndex.Data, Initials: FileIndex.Initials, InitialsFor: FileIndex._initialsFor
            , Quiet: FileIndex.WatchQuietMs, Gap: FileIndex.WatchMinGapMs, Idle: FileIndex.SaveIdleMs}
        options["ScopeFolders"] := [root], options["UseEverything"] := 0
        FileIndex.File := A_Temp "\ALTRunFileIndexWatch.txt"
        try FileDelete(FileIndex.File)
        FileIndex.WatchQuietMs := 0, FileIndex.WatchMinGapMs := 0, FileIndex.SaveIdleMs := 999999999   ; not written automatically; the test writes it itself
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
            eq("scanned", FileIndex.Count, 1)
            eq("save waits for idle", FileIndex._savePending && !FileExist(FileIndex.File), true)
            FileIndex.SavePending()
            eq("saved on demand", !FileIndex._savePending && FileExist(FileIndex.File) != "", true)
            FileIndex._PollWatches()
            eq("no change, no scan", FileIndex.Scanning, false)
            FileAppend("x", root "\Old Plan.dwg")                         ; modified files: the index only has paths, no rescan needed
            FileAppend("", root "\Report.docx.crdownload")                ; temporary files don't count either
            Sleep(300)
            FileIndex._PollWatches()
            eq("edit and temp file ignored", FileIndex.Scanning || FileIndex._changedAt != 0, false)
            FileAppend("", root "\New Plan.dwg")
            Sleep(300)
            FileIndex._PollWatches()
            eq("new file rescans", FileIndex.Scanning || FileIndex.Count = 3, true)
            waitScan()
            eq("new file indexed", FileIndex.Count, 3)               ; including .crdownload (scanning doesn't skip it, it just doesn't trigger)
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
            FileIndex.Data := saved.Data, FileIndex.Initials := saved.Initials, FileIndex._initialsFor := saved.InitialsFor
            FileIndex.WatchQuietMs := saved.Quiet, FileIndex.WatchMinGapMs := saved.Gap, FileIndex.SaveIdleMs := saved.Idle
            try FileDelete(A_Temp "\ALTRunFileIndexWatch.txt")
            try DirDelete(root, true)
        }
    }

    ; File type filters: doc report / cad plan (as in Listary)
    static FileTypes() {
        eq := (n, a, e) => TestRunner.Equal("FileTypes." n, a, e)
        options := AppSettings.Feature("FileSearch")
        savedEverything := options["UseEverything"]
        options["UseEverything"] := 0                                       ; the test uses the built-in index
        parsed := FileSearchProvider.ParseTypeFilter(" Doc = doc, .DOCX;pdf  *.doc ")
        eq("parse keyword", parsed.Keyword, "doc")
        eq("parse everything", parsed.Everything, "ext:doc;docx;pdf")
        eq("parse invalid", FileSearchProvider.ParseTypeFilter("no equals sign"), "")
        eq("parse empty list", FileSearchProvider.ParseTypeFilter("x = , `;"), "")
        keywords := ""
        for filter in FileSearchProvider.TypeFilters()
            keywords .= filter.Keyword " "
        eq("default keywords", keywords, "doc pic video audio zip exe cad ")
        FileIndex.SetEntries(["C:\P\Tower Plan.dwg", "C:\P\Tower Plan.pdf", "C:\P\Tower photo.jpg", "C:\P\Tower"], [0, 0, 0, 1])
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
        eq("file mode", titles(ProviderRegistry.SearchOnly(FileSearchProvider, "cad tower")), "Tower Plan.dwg")
        eq("file mode plain", titles(ProviderRegistry.SearchOnly(FileSearchProvider, "tower")), "Tower|Tower Plan.dwg|Tower Plan.pdf|Tower photo.jpg")
        hint := FileSearchProvider.Search(SearchQuery("cad"))
        eq("keyword alone hint", hint.Length, 1)
        eq("keyword alone hint score", hint[1].Score = ResultItem.KeywordHintScore && !hint[1].Exclusive, true)
        eq("hint lists extensions", InStr(hint[1].Title, "dwg dxf") ? 1 : 0, 1)
        eq("not a keyword", FileSearchProvider.Search(SearchQuery("cadence")).Length, 0)
        eq("everything prefix", FileSearchProvider._ScopePrefix(parsed), "ext:doc;docx;pdf ")
        eq("everything folder prefix", FileSearchProvider._ScopePrefix(true), "folder:")
        options["TypeFilters"].Push("dwgonly = dwg")                        ; a changed setting applies at once
        eq("settings change", titles(FileSearchProvider.Search(SearchQuery("dwgonly tower"))), "Tower Plan.dwg")
        options["TypeFilters"].Pop()
        options["UseEverything"] := savedEverything
        FileIndex.SetEntries([])
    }

    ; "recent report": Windows' recent files and folders (shortcuts in the Recent folder), newest first, without Everything
    static RecentFiles() {
        eq := (n, a, e) => TestRunner.Equal("RecentFiles." n, a, e)
        root := A_Temp "\ALTRun-test-recentfiles"
        try DirDelete(root, true)
        DirCreate(root "\Recent"), DirCreate(root "\Projects\Tower"), DirCreate(root "\Docs")
        longPath := Buffer(2048)                                            ; A_Temp may be an 8.3 path; shortcuts give the long one
        if DllCall("GetLongPathNameW", "Str", root, "Ptr", longPath, "UInt", 1024)
            root := StrGet(longPath)
        recent := root "\Recent"
        link := (target, name, time) => (FileCreateShortcut(target, recent "\" name ".lnk"), FileSetTime(time, recent "\" name ".lnk"))
        FileAppend("x", root "\Docs\Report 2026.pdf"), FileAppend("x", root "\Docs\notes.txt")
        link(root "\Docs\Report 2026.pdf", "Report 2026.pdf", "20260301000000")
        link(root "\Docs\notes.txt", "notes.txt", "20260201000000")
        link(root "\Projects\Tower", "Tower", "20260101000000")
        link(root "\Docs\Report 2026.pdf", "Report copy", "20251201000000")   ; the same target twice: listed once
        link(root "\Docs\Gone.docx", "Gone.docx", "20260401000000")           ; deleted: left out
        FileAppend("not a shortcut", recent "\Broken.lnk"), FileSetTime("20260501000000", recent "\Broken.lnk")
        FileAppend("x", recent "\desktop.ini")                             ; not a shortcut
        saved := [WindowsRecent.Dir, WindowsRecent.CacheMs]
        WindowsRecent.Dir := recent, WindowsRecent.CacheMs := 0
        options := AppSettings.Feature("FileSearch")
        try {
            items := WindowsRecent.Items(300)
            names := ""
            for item in items
                names .= item.Name (item.IsFolder ? "/" : "") ";"
            eq("newest first, once, existing only", names, "Report 2026.pdf;notes.txt;Tower/;")
            eq("target path", items.Length ? items[1].Path : "", root "\Docs\Report 2026.pdf")
            eq("newest shortcuts only", WindowsRecent.Items(2).Length, 0)   ; Broken and Gone are the two newest
            listed := FileSearchProvider.RecentResults("", options, true)
            eq("recent + space: newest", listed.Length "|" listed[1].Title "|" listed[1].Subtitle "|" listed[1].Arg, "3|Report 2026.pdf|" root "\Docs|" root "\Docs\Report 2026.pdf")
            eq("folder kind and Tab", listed[3].Kind "|" listed[3].AutoComplete, "folder|" root "\Projects\Tower\")
            eq("exclusive", listed[1].Exclusive, true)
            found := FileSearchProvider.RecentResults("rep", options, true)
            eq("filtered", found.Length "|" found[1].Title, "1|Report 2026.pdf")
            none := FileSearchProvider.RecentResults("zzz", options, true)
            eq("nothing found", none.Length "|" none[1].Valid "|" none[1].Title, "1|0|" I18n.T("Files.RecentNone"))
            hint := FileSearchProvider.RecentResults("", options, false)
            eq("keyword only: a hint", hint.Length "|" hint[1].Valid "|" hint[1].Score "|" hint[1].Subtitle, "1|0|" ResultItem.KeywordHintScore "|recent ...")
            results := FileSearchProvider.Search(SearchQuery("recent note"))
            eq("through Search", results.Length ? results[1].Title : "", "notes.txt")
            eq("file search mode", FileSearchProvider.SearchOnly("recent tower")[1].Title, "Tower")
            eq("default keyword", AppSettings.Defaults()["Features"]["FileSearch"]["RecentKeywords"][1], "recent")
            FileSetTime("20260601000000", recent "\notes.txt.lnk")             ; opened again: moves to the top
            eq("reopened moves up", WindowsRecent.Items(300)[1].Name, "notes.txt")
            eq("quick switch folders", FileManagers.RecentFolders(10, recent).Length, 2)
        } finally {
            WindowsRecent.Dir := saved[1], WindowsRecent.CacheMs := saved[2]
            try DirDelete(root, true)
        }
    }

    ; File-only search (space / ' first): up to Features.FileSearch.MaxResults + the fallback row, not the general 50 (#515)
    static ResultLimit() {
        eq := (n, a, e) => TestRunner.Equal("ResultLimit." n, a, e)
        options := AppSettings.Feature("FileSearch"), saved := options["MaxResults"]
        try {
            files := [], mixed := []
            Loop 3
                files.Push(ResultItem("f" A_Index, "", {Exclusive: true, Provider: FileSearchProvider.Id}))
            mixed.Push(ResultItem("f", "", {Exclusive: true, Provider: FileSearchProvider.Id}), ResultItem("c", "", {Exclusive: true, Provider: "Clipboard"}))
            options["MaxResults"] := 200
            eq("file search", ProviderRegistry.ResultLimit(files), 201)
            eq("other exclusive", ProviderRegistry.ResultLimit(mixed), ProviderRegistry.MaxResults)
            eq("not exclusive", ProviderRegistry.ResultLimit([ResultItem("a", "", {Provider: FileSearchProvider.Id})]), ProviderRegistry.MaxResults)
            eq("empty", ProviderRegistry.ResultLimit([]), ProviderRegistry.MaxResults)
            options["MaxResults"] := 30
            eq("smaller setting keeps 50", ProviderRegistry.ResultLimit(files), ProviderRegistry.MaxResults)
        } finally
            options["MaxResults"] := saved
    }

    ; With many commands: results only for the top MaxResults (counting the learned bonus); typing on only searches the previous results
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
        ; an often chosen command: the learned bonus brings it into the top MaxResults
        Knowledge.Record("n", CustomCommandProvider._Uid(rare))
        Knowledge.Record("n", CustomCommandProvider._Uid(rare))
        found := false
        for item in CustomCommandProvider.Search(SearchQuery("n"))
            found := found || (item.Title = "Annual Report")
        eq("learned pick kept", found, true)

        ; typing on only searches the previous results, with the same result as searching from scratch
        CustomCommandProvider._ResetNarrowing()
        full := titles("note 11")
        titles("not"), titles("note"), titles("note ")
        eq("narrowed equals full search", titles("note 11"), full)
        eq("narrowed result", full, "Note 11|Note 110|Note 111|Note 112|Note 113|Note 114|Note 115|Note 116|Note 117|Note 118|Note 119")
        ; after a command changes, all commands are searched again
        titles("rep")
        commands[5]["Title"] := "Report Folder"
        CustomCommandProvider._ResetNarrowing()
        eq("edited command found", InStr(titles("repo"), "Report Folder") > 0, true)
        ; adding / deleting commands (a different count) searches again too
        titles("zzz")
        commands.Push(Map("Title", "zzz top", "Type", "Folder", "Target", "C:\z", "Arguments", "", "Keyword", ""))
        eq("added command found", titles("zzz t"), "zzz top")

        AppSettings.Data["CustomCommands"] := saved
        Knowledge.Picks := savedPicks, Knowledge.QueryPicks := savedQueryPicks, Knowledge.History := savedHistory
        if (Knowledge._saveTimer != "")
            SetTimer(Knowledge._saveTimer, 0)                               ; don't run the disk write Record() scheduled
        CustomCommandProvider._ResetNarrowing()

        ; Icons on network locations don't read the disk: UNC paths use the generic folder / extension icons
        ; Cache limit: only the most recently used stay (fake icons with handle 0, so DestroyIcon isn't called)
        savedIcons := IconCache._icons, savedUsed := IconCache._used
        IconCache._icons := Map("a", 0, "b", 0, "c", 0, "d", 0), IconCache._used := Map("a", 5, "b", 1, "c", 9, "d", 3)
        IconCache.Trim(2)
        keys := ""
        for key in IconCache._icons
            keys .= key
        eq("trim keeps newest", keys, "ac")
        eq("trim used map", IconCache._used.Count, 2)
        IconCache._icons := savedIcons, IconCache._used := savedUsed
        eq("remote unc", Path.IsRemote("\\server\share\P1001"), true)
        eq("remote local", Path.IsRemote("C:\Windows"), false)
        eq("remote folder key", IconCache._CacheKey("\\server\share\P1001 - Riverside Park"), "folder:")
        eq("remote folder slash", IconCache._CacheKey("\\server\share\P1001\"), "folder:")
        eq("remote file key", IconCache._CacheKey("\\server\share\Report.PDF"), "ext:.pdf")
        eq("remote exe key", IconCache._CacheKey("\\server\share\tool.exe"), "ext:.exe")
        ; a folder with a dot in its name has no "extension" (it used to get a blank icon)
        eq("remote dotted folder", IconCache._CacheKey("\\server\Proposals\2025\26. 18 Sample Road (A)"), "folder:")
        eq("remote dotted folder slash", IconCache._CacheKey("\\server\share\10.P1002-Riverside\"), "folder:")
        eq("local dotted folder", IconCache._CacheKey("C:\Projects\26. 18 Sample Road (A)"), "c:\projects\26. 18 sample road (a)")
        eq("extension", Path.Extension("C:\a\Report.PDF"), "PDF")
        eq("extension with spaces", Path.Extension("C:\a\26. 18 New Road"), "")
        eq("extension appref-ms", Path.Extension("C:\a\App.appref-ms"), "appref-ms")
        eq("extension 7z", Path.Extension("C:\a\backup.7z"), "7z")
        eq("extension sldprt", Path.Extension("C:\a\Beam.SLDPRT"), "SLDPRT")
        eq("extension digits only", Path.Extension("C:\a\Design.2019"), "")
        eq("folder command icon remote", CustomCommandProvider._FolderIcon("\\server\share\26. 18 Sample Road (A)"), "folder:")
        ; normal folders use the generic icon; those with desktop.ini (a custom icon) and drive roots use their own
        iconRoot := A_Temp "\ALTRun-folder-icon-test"
        try DirDelete(iconRoot, true)

        ; Computing search keys after start-up: broken entries are skipped without affecting the rest
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
        ; no disk access while searching: the generic icon first, its own only after desktop.ini was checked in the background
        eq("folder icon with desktop.ini, not probed yet", CustomCommandProvider._FolderIcon(iconRoot "\Custom"), "folder:")
        IconCache._ProbeQueued()
        eq("folder icon with desktop.ini, after probe", CustomCommandProvider._FolderIcon(iconRoot "\Custom"), iconRoot "\Custom")
        ; handling the queue is interrupted by another round (the background timer) that already handled the rest: skipped, no "Item has no value"
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

    ; Placeholders that won't work (the Commands page's Status column): {selection}, {query} that gets no text, other {word}s
    static PlaceholderProblems() {
        eq := (n, a, e) => TestRunner.Equal("PlaceholderProblems." n, a, e)
        kinds(command, hotkeys := "") {
            text := ""
            for each in CustomCommandProvider.PlaceholderProblems(command, IsObject(hotkeys) ? hotkeys : [])
                text .= each.Kind (each.HasOwnProp("Name") ? " " each.Name : "") "|"
            return text
        }
        cmd(target, arguments := "", keyword := "", extra*) {
            command := Map("Id", "pq7k2m4x", "Title", "t", "Type", "Command", "Target", target, "Arguments", arguments, "Keyword", keyword)
            Loop extra.Length // 2
                command[extra[A_Index * 2 - 1]] := extra[A_Index * 2]
            return command
        }
        hk(key) => Map("Key", key, "Action", "custom:pq7k2m4x", "SearchText", "", "WinTitle", "")
        eq("plain command", kinds(cmd("notepad.exe")), "")
        eq("query with a keyword", kinds(cmd("cmd.exe", "/k ping {query}", "ping")), "")
        eq("query with a content type", kinds(cmd("cmd.exe", "/k ping {query}", "", "Selection", "IP")), "")
        eq("query with a hotkey", kinds(cmd("cmd.exe", "/k ping {query}"), [hk("^!p")]), "")
        eq("query with nothing", kinds(cmd("cmd.exe", "/k ping {query}")), "Query|")
        eq("query: any case, in the target", kinds(cmd("https://x.example.com/{Query}")), "Query|")
        eq("query: blank keyword", kinds(cmd("cmd.exe", "{query}", " , ")), "Query|")
        eq("query: hotkey without keys", kinds(cmd("cmd.exe", "{query}"), [hk(""), hk("~")]), "Query|")
        other := hk("^!o"), other["Action"] := "custom:zz9k2m4x"
        eq("query: another command's hotkey", kinds(cmd("cmd.exe", "{query}"), [other]), "Query|")
        eq("window placeholders", kinds(cmd("cmd.exe", "/c echo {hwnd} {PID}")), "")
        eq("selection", kinds(cmd("cmd.exe", "/c echo {selection}", "e")), "Selection|")
        eq("selection first", kinds(cmd("cmd.exe", "{date} {query} {Selection}")), "Selection|Query|Unknown {date}|")
        eq("unknown, each once", kinds(cmd("x.ahk", "{clipboard} {Enter} {clipboard}")), "Unknown {clipboard}|Unknown {Enter}|")
        eq("GUIDs and code aren't placeholders", kinds(cmd("shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "-Command & { $_ } {} {1x}")), "")

        savedLang := I18n.Lang
        I18n.Init("en")
        try {
            status := (command, hotkeys := "") => CustomCommandProvider.PlaceholderStatus(command, IsObject(hotkeys) ? hotkeys : [])
            eq("status: none", status(cmd("notepad.exe")), "")
            eq("status: query", status(cmd("cmd.exe", "{query}")), "{query} empty")
            eq("status: selection first", status(cmd("cmd.exe", "{query} {selection}")), "Use {query}")
            eq("status: unknown", status(cmd("cmd.exe", "{date}", "d")), "Unknown {date}")
            warning := (command) => CustomCommandProvider.PlaceholderWarning(command, [])
            eq("warning: none", warning(cmd("notepad.exe")), "")
            eq("warning: each problem, the unknown ones together", warning(cmd("cmd.exe", "{query} {selection} {date} {Enter}"))
                , I18n.T("Cmd.Placeholder.Selection") "`n`n" I18n.T("Cmd.Placeholder.Query") "`n`n"
                . "Not a placeholder in commands: {date}, {Enter}. Only {query}, {hwnd} and {pid} are replaced; others are passed to the program as written.")
        } finally I18n.Init(savedLang)
    }

    static CheckTargets() {
        eq := (n, a, e) => TestRunner.Equal("CheckTargets." n, a, e)
        dir := A_Temp "\ALTRunTest_CheckTargets"
        try DirDelete(dir, true)
        DirCreate(dir "\P1001 - Riverside Park")
        FileAppend("", dir "\summary.pdf")
        check(type, target, roots := "") => CustomCommandProvider.CheckTarget(Map("Title", "t", "Type", type, "Target", target), roots)
        eq("folder ok", check("Folder", dir "\P1001 - Riverside Park"), "OK")
        eq("folder renamed", check("Folder", dir "\P1001 - Riverside Park (old)"), "Missing")
        eq("folder trailing slash", check("Folder", dir "\P1001 - Riverside Park\"), "OK")
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
        ; disconnected network drives: whether a drive is reachable is kept in roots, so it isn't waited for item by item
        offline := Map("Q:", false)
        eq("drive offline", check("Folder", "Q:\Demo\P1001", offline), "Unavailable")
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
            Map("Title", "Tower A, B, C", "Type", "Folder", "Target", "Q:\Demo\P1001 - Riverside Park", "Arguments", "", "Keyword", ""),
            Map("Title", "Riverside Report", "Type", "File", "Target", "Q:\Docs\summary.pdf", "Arguments", "", "Keyword", ""),
            Map("Title", "Check IP", "Type", "Command", "Target", "cmd.exe", "Arguments", "/k ipconfig", "Keyword", ""),
            Map("Title", "Drive Q", "Type", "Folder", "Target", "Q:\", "Arguments", "", "Keyword", "")]
        titles(text) {
            list := ""
            for item in ProviderRegistry.SortByScore(CustomCommandProvider.Search(SearchQuery(text)))
                list .= item.Title "|"
            return RTrim(list, "|")
        }
        eq("folder name", titles("1001"), "Tower A, B, C")
        eq("title first", titles("riverside"), "Riverside Report|Tower A, B, C")
        eq("file name without extension", titles("summary"), "Riverside Report")
        eq("extension not searched", titles("pdf"), "")
        eq("command target not searched", titles("cmd"), "")
        eq("drive root not a name", titles("q:"), "")

        ; {query}: "keyword text" replaces {query} in the target / arguments (encoded in URLs), only this command shown
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
        items := CustomCommandProvider.Search(SearchQuery("ping 10.0.0.12"))   ; typing on: don't narrow down with the previous results
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

    static HiddenApps() {
        eq := (n, a, e) => TestRunner.Equal("HiddenApps." n, a, e)
        saved := ProviderRegistry.Providers
        ProviderRegistry.Providers := [ApplicationProvider]
        options := AppSettings.Feature("Applications")
        options["Hidden"] := []
        savedFile := AppSettings.File
        AppSettings.File := A_Temp "\ALTRunTest.json"                     ; deleting saves the settings: write to a temporary file
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
        TestRunner.True("HiddenApps.can delete", ResultActions.CanDelete(item))
        TestRunner.True("HiddenApps.prompt", InStr(ResultActions.DeletePrompt(item), "Notepad++") && InStr(ResultActions.DeletePrompt(item), "uninstalled"))
        ResultActions.DeleteItem(item)
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

    ; Broken shortcuts in an indexed folder: still indexed (without a target), the scan isn't interrupted (it once threw "linkTarget not assigned")
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

    ; The index scan scope changed in Preferences: the cache expires after saving and is rescanned in the background after restarting
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
            for cacheFile in [ApplicationProvider.IndexFile, FileIndex.File] {
                try FileDelete(cacheFile)
                FileAppend("{}", cacheFile)
            }
            eq("fresh caches", ApplicationProvider._CacheExpired() || FileIndex._CacheExpired(), false)
            FileIndex._savePending := true
            ApplicationProvider.ExpireCache(), FileIndex.ExpireCache()
            eq("expired", ApplicationProvider._CacheExpired() && FileIndex._CacheExpired(), true)
            eq("pending save dropped", FileIndex._savePending, false)
            ApplicationProvider.ExpireCache()                               ; no error without a cache file either
        } finally {
            FileIndex._savePending := false
            try FileDelete(ApplicationProvider.IndexFile)
            try FileDelete(FileIndex.File)
            ApplicationProvider.IndexFile := savedApp, FileIndex.File := savedFile
        }
    }

    ; Store apps are listed through COM (shell:AppsFolder): they replace the previous ones, other apps are kept, Exclude applies.
    ; Wine has no Store apps: only checked that nothing else is lost
    static StoreAppsScan() {
        eq := (n, a, e) => TestRunner.Equal("StoreAppsScan." n, a, e)
        options := AppSettings.Feature("Applications"), savedExclude := options["Exclude"]
        savedFile := ApplicationProvider.IndexFile, savedApps := ApplicationProvider.Apps, savedScanned := ApplicationProvider._storeScanned
        ApplicationProvider.IndexFile := A_Temp "\ALTRunStoreAppsScan.json"
        normal := Map("Title", "Editor", "Target", "C:\Tools\Editor.exe", "Detail", "", "Search", "")
        old := Map("Title", "Old", "Target", "shell:AppsFolder\Old_1!App", "Detail", "", "Search", "")
        store() {
            list := []
            for entry in ApplicationProvider.Apps
                if (InStr(entry["Target"], "shell:AppsFolder\") = 1)
                    list.Push(entry)
            return list
        }
        try {
            ApplicationProvider.Apps := [normal, old]
            start := A_TickCount
            ApplicationProvider._ScanStoreApps()
            ms := A_TickCount - start
            eq("other apps kept", ApplicationProvider.Apps.Length && ApplicationProvider.Apps[1] == normal, true)
            if DllCall("GetProcAddress", "Ptr", DllCall("GetModuleHandle", "Str", "ntdll", "Ptr"), "AStr", "wine_get_version", "Ptr")
                return
            apps := store()
            FileAppend("INFO StoreAppsScan: " apps.Length " store apps in " ms " ms`n", "*")
            eq("listed", apps.Length > 0, true)
            eq("previous replaced", InStr(JSON.Stringify(apps), "Old_1!App"), 0)
            good := true
            for entry in apps
                good := good && InStr(entry["Target"], "!") > 0 && entry["Title"] != ""
            eq("package!app targets with names", good, true)
            if !apps.Length
                return
            name := apps[1]["Title"]
            options["Exclude"] := "^\Q" name "\E$"
            ApplicationProvider._ScanStoreApps()
            excluded := true
            for entry in store()
                excluded := excluded && entry["Title"] != name
            eq("Exclude applies", excluded, true)
        } finally {
            options["Exclude"] := savedExclude
            ApplicationProvider.IndexFile := savedFile, ApplicationProvider.Apps := savedApps, ApplicationProvider._storeScanned := savedScanned
            try FileDelete(A_Temp "\ALTRunStoreAppsScan.json")
        }
    }

    static AppIndexRefresh() {
        eq := (n, a, e) => TestRunner.Equal("AppIndexRefresh." n, a, e)
        root := A_Temp "\ALTRunAppIndexTest"
        try DirDelete(root, true)
        for dir in ["", "\Office", "\Office\Tools", "\Office\Tools\More", "\Office\Tools\More\TooDeep", "\Games"]
            DirCreate(root dir)                                             ; created level by level (creating several levels at once hangs in Wine)
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
        expected := "Chess|Draw|Editor|Sheet|Writer"                       ; Hidden at depth 4, Uninstall in Exclude and txt files are left out

        eq("no cache = expired", ApplicationProvider._CacheExpired(), true)
        eq("manual count", ApplicationProvider.Rebuild(), 5)
        eq("manual", titles(), expected)
        eq("fresh cache", ApplicationProvider._CacheExpired(), false)
        FileSetTime(DateAdd(A_Now, -(ApplicationProvider.FallbackMinutes + 1), "Minutes"), ApplicationProvider.IndexFile, "M")
        eq("older than the fallback interval", ApplicationProvider._CacheExpired(), true)

        ApplicationProvider.Apps := []
        job := ApplicationProvider._NewJob(), steps := 1
        while !ApplicationProvider._Step(job, -1)                          ; -1: stops after each folder scanned and goes on next time
            steps++
        eq("background in steps", steps > 5, true)
        eq("not replaced while scanning", titles(), "")
        ApplicationProvider._Finish(job, false)
        eq("background", titles(), expected)
        eq("bad Exclude ignored", ApplicationProvider._Exclude(Map("Exclude", "(uninstall")), "")
        eq("valid Exclude used", ApplicationProvider._Exclude(Map("Exclude", "i)uninstall")), "i)uninstall")

        ; Watching folders: a new file rebuilds in the background (after a short quiet); the cache of a background rebuild is written when idle
        savedQuiet := ApplicationProvider.WatchQuietMs, savedGap := ApplicationProvider.WatchMinGapMs
        ApplicationProvider.WatchQuietMs := 0, ApplicationProvider.WatchMinGapMs := 0
        ApplicationProvider.StartWatching()
        eq("watching the folder", ApplicationProvider._watcher.Count, 1)
        eq("watch read pending", ApplicationProvider._watcher.Watches[1].Pending, true)
        ApplicationProvider._PollWatches()
        eq("no change, no rebuild", IsObject(ApplicationProvider._job), false)
        FileAppend("", root "\notes2.txt")                                  ; files that aren't applications: no rebuild
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
            ApplicationProvider._lastBackground := 0                        ; changes in subfolders arrive too (most Start menu shortcuts are in subfolders; Wine doesn't support it)
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

    ; A folder watcher should reattach after a failure and safely end pending asynchronous I/O when closed
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

            ; a folder stopped after it failed to start waiting again: reopened when it is time, reports "*" (changes may have been missed meanwhile), then watched as usual
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

            ; closing a pending read: the buffer is freed only after it ended; opening / closing several times works
            for i in [1, 2, 3] {
                pendingWatch := FolderWatcher.Create(root)
                eq("close cycle " i " created", IsObject(pendingWatch) && pendingWatch.Pending, true)
                if !IsObject(pendingWatch)
                    continue
                FolderWatcher.Close(pendingWatch)
                eq("close cycle " i " completed", pendingWatch.Pending || pendingWatch.Handle || pendingWatch.Event, false)
                FolderWatcher.Close(pendingWatch)                           ; closing again is fine
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
        AppSettings.File := A_Temp "\ALTRunTest.json"                     ; deleting saves the settings: write to a temporary file
        titles() {
            list := ""
            for item in SystemProvider.Search(SearchQuery("devices and printers"))
                list .= item.Title "|"
            return list
        }
        TestRunner.True("HiddenSystemCommands.found", InStr(titles(), I18n.T("Tool.Printers")))
        item := SystemProvider.Search(SearchQuery("devices and printers"))[1]
        item.Provider := "System"
        TestRunner.True("HiddenSystemCommands.can delete", ResultActions.CanDelete(item))
        TestRunner.True("HiddenSystemCommands.prompt", InStr(ResultActions.DeletePrompt(item), I18n.T("Tool.Printers")))
        ResultActions.DeleteItem(item)
        eq("saved in settings", options["Hidden"].Length ? options["Hidden"][1] : "", "Printers")
        TestRunner.True("HiddenSystemCommands.hidden", !InStr(titles(), I18n.T("Tool.Printers")))
        SystemProvider.Hide("Printers")
        eq("added once", options["Hidden"].Length, 1)
        options["Hidden"] := [" printers "]                                 ; written by hand in the settings: case-insensitive, spaces trimmed
        TestRunner.True("HiddenSystemCommands.case-insensitive", !InStr(titles(), I18n.T("Tool.Printers")))
        options["Hidden"] := []
        TestRunner.True("HiddenSystemCommands.restored", InStr(titles(), I18n.T("Tool.Printers")))
        eq("default", AppSettings.Defaults()["Features"]["System"]["Hidden"].Length, 0)
        options["Hidden"] := savedHidden
        ProviderRegistry.Providers := saved
        AppSettings.File := savedFile
        try FileDelete(A_Temp "\ALTRunTest.json")
    }

    ; Every folder written with A_ variables in the default settings must resolve to a real path (otherwise that folder would never be indexed)
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

    ; Chinese file names Everything finds with pinyin on: sorted by pinyin initials (a little below a direct match of the same tier), with the matching characters highlighted
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
        eq("unknown character", Pinyin.Initials("A" Chr(0x9FFF) "B"), "AB")
        ; The full table (#257): GB2312 level-2, GBK-only and traditional characters; level-1 ones keep their letter
        eq("table loaded", Pinyin._letters.Size, 0x9FFF - 0x3400 + 1)
        eq("level-2 characters", Pinyin.Initials("麋鹿昊鑫琪璐"), "MLHXQL")
        eq("GBK-only and traditional", Pinyin.Initials("喆記事本"), "ZJSB")
        eq("polyphones keep their old letter", Pinyin.Initials("长城長"), "CCC")
        eq("CJK punctuation dropped as before", Pinyin.Initials("测试。"), "CS")
        letters := Pinyin._letters, Pinyin._letters := Buffer(0), Pinyin._cache := Map()     ; without the file: the ranges
        eq("no table: level-1 only", Pinyin.Initials("记事本麋"), "JSB")
        Pinyin._letters := letters, Pinyin._cache := Map()

        ; the built-in index (without Everything) also finds by pinyin initials
        paths := ["C:\D\测试.bat", "C:\D\cs_report.bat", "C:\D\证件扫描文件.zip", "C:\D\项目测试报告.docx", "C:\D\notes.txt"]
        FileIndex.SetEntries(paths)                                           ; with the initials (a scan computes them as it goes)
        eq("index initials", FileIndex.Initials.Count "|" FileIndex.Initials[1] "|" FileIndex.Initials.Has(2) "|" FileIndex.Initials[3], "3|cs.bat|0|zjsmwj.zip")
        names(needle) {
            list := ""
            for item in FileIndex.Search(needle, 10) {
                SplitPath(item.Path, &fileName)
                list .= fileName "|"
            }
            return RTrim(list, "|")
        }
        eq("index: direct first, then pinyin", names("cs"), "cs_report.bat|测试.bat|项目测试报告.docx")
        eq("index: initials", names("zjsm"), "证件扫描文件.zip")
        eq("index: typing on keeps pinyin", (names("c"), names("cs")), "cs_report.bat|测试.bat|项目测试报告.docx")
        eq("index: chinese input", names("测试"), "测试.bat|项目测试报告.docx")
        eq("index: several words", names("zj sm"), "")
        ; many Chinese names, matches far apart: the packed initials find the same as checking each one
        many := []
        Loop 400
            many.Push("C:\D\" (Mod(A_Index, 97) = 0 ? "证件扫描" : Mod(A_Index, 2) ? "测试" : "项目") A_Index ".txt")
        FileIndex.SetEntries(many)
        expected := 0
        for index, text in FileIndex.Initials
            expected += FileIndex.PinyinScore("zjsm", text) ? 1 : 0
        scores := Map(), matches := []
        FileIndex._FindInitials("zjsm", scores, matches)
        eq("index: packed initials, far apart", matches.Length "|" expected "|" FileIndex.Search("zjsm", 10)[1].Path, "4|4|C:\D\证件扫描97.txt")
        FileIndex.SetEntries(paths)
        FileIndex._initialsFor := ""                                          ; initials of another index: no pinyin
        eq("index: initials for another index", names("zjsm"), "")
        savedFile := FileIndex.File, FileIndex.File := A_Temp "\ALTRunPinyinIndex.bin"
        try {
            FileIndex.SetEntries(["C:\D\证件扫描文件.zip", "C:\D\notes.txt"])
            FileIndex._SaveCache(), FileIndex.SetEntries([])
            FileIndex._LoadCache()
            eq("cache keeps the initials", (FileIndex._initialsFor == FileIndex.Data) "|" FileIndex.Initials.Count "|" FileIndex.Initials[1], "1|1|zjsmwj.zip")
            eq("index: pinyin after reading the cache", names("zjsm"), "证件扫描文件.zip")
        } finally {
            try FileDelete(FileIndex.File)
            FileIndex.File := savedFile
        }
        FileIndex.SetEntries([])
        ranges := FuzzyMatcher.MatchRanges("cs", "测试.bat")
        eq("highlight", ranges.Length = 1 ? ranges[1][1] "," ranges[1][2] : ranges.Length, "1,2")
    }

    static FileSearchModes() {
        eq := (n, a, e) => TestRunner.Equal("FileSearchModes." n, a, e)
        options := AppSettings.Feature("FileSearch")
        saved := [options["InDefaultResults"], options["UseEverything"]]
        options["UseEverything"] := 0                                       ; the tests use the built-in index
        FileIndex.SetEntries(["C:\Windows\Fonts\MARLETT.TTF", "C:\Docs\Market Notes.pdf"], [0, 0])
        eq("default off", AppSettings.Defaults()["Features"]["FileSearch"]["InDefaultResults"], 0)
        eq("space prefix on", AppSettings.Defaults()["Features"]["FileSearch"]["SpacePrefix"], 1)
        options["InDefaultResults"] := 0
        eq("normal search has no files", FileSearchProvider.Search(SearchQuery("mar")).Length, 0)
        files := ProviderRegistry.SearchOnly(FileSearchProvider, "mar")
        TestRunner.True("FileSearchModes.file mode finds files", files.Length >= 2 && files[1].Provider = "FileSearch")
        eq("file mode empty", ProviderRegistry.SearchOnly(FileSearchProvider, "  ").Length, 0)
        TestRunner.True("FileSearchModes.quote still works", FileSearchProvider.Search(SearchQuery("'mar")).Length >= 2)
        ; files in normal results: by default only the start of the name or of a word; Contains = anywhere in the name (discussion #308)
        savedMatch := options["DefaultResultsMatch"]
        eq("match default", AppSettings.Defaults()["Features"]["FileSearch"]["DefaultResultsMatch"], "Start")
        options["InDefaultResults"] := 1, options["DefaultResultsMatch"] := "Start"
        eq("start: word start found", FileSearchProvider.Search(SearchQuery("notes")).Length, 1)
        eq("start: middle of a word not found", FileSearchProvider.Search(SearchQuery("ket")).Length, 0)
        options["DefaultResultsMatch"] := "Contains"
        found := FileSearchProvider.Search(SearchQuery("ket"))
        eq("contains: middle of a word found", found.Length = 1 ? found[1].Title : found.Length, "Market Notes.pdf")
        FileIndex.SetEntries(["C:\Docs\Annotated.txt", "C:\Docs\Market Notes.pdf"], [0, 0])
        found := FileSearchProvider.Search(SearchQuery("not"))
        eq("contains: a word start still ranks first", found.Length = 2 ? found[1].Title : found.Length, "Market Notes.pdf")
        options["DefaultResultsMatch"] := "Start"
        eq("start: only the word start", FileSearchProvider.Search(SearchQuery("not")).Length, 1)
        options["DefaultResultsMatch"] := savedMatch
        options["InDefaultResults"] := saved[1], options["UseEverything"] := saved[2]
        FileIndex.SetEntries([])
    }

    ; "folder atlas" searches folders only; a keyword without a following space does not hide other results; results are sorted by name match
    static FolderSearch() {
        eq := (n, a, e) => TestRunner.Equal("FolderSearch." n, a, e)
        options := AppSettings.Feature("FileSearch")
        savedEverything := options["UseEverything"]
        options["UseEverything"] := 0                                       ; the tests use the built-in index
        FileIndex.SetEntries(["C:\Docs\noteatlas.pdf", "C:\Projects\Atlas Tower", "C:\Docs\Atlas drawing.dwg", "C:\Projects\Old Atlas"], [0, 1, 0, 1])
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
        eq("file mode folder keyword", titles(ProviderRegistry.SearchOnly(FileSearchProvider, "folder atlas")), "Atlas Tower|Old Atlas")
        eq("file mode plain", titles(ProviderRegistry.SearchOnly(FileSearchProvider, "atlas")), "Atlas Tower|Atlas drawing.dwg|Old Atlas|noteatlas.pdf")
        ; Only the keyword (no space yet): the hint ranks low and does not hide commands with folder in their name
        hint := FileSearchProvider.Search(SearchQuery("folder"))
        eq("keyword alone hint", hint.Length, 1)
        eq("keyword alone not exclusive", hint[1].Exclusive, false)
        eq("keyword alone hint score", hint[1].Score, ResultItem.KeywordHintScore)
        ; where the hint ranks: after names starting with the word or matching by initials, before "contains" (apps get +10, see ApplicationProvider)
        hintScore := ResultItem.KeywordHintScore
        eq("hint after app starting with it", FuzzyMatcher.Score("doc", "Docker Desktop") + 10 > hintScore, true)
        eq("hint after command starting with it", FuzzyMatcher.Score("doc", "Docker") > hintScore, true)
        eq("hint after command word start", FuzzyMatcher.Score("doc", "Shared Documents Folder") > hintScore, true)
        eq("hint after command initials", FuzzyMatcher.Score("doc", "Daily Order Check") > hintScore, true)
        eq("hint before app containing it", FuzzyMatcher.Score("doc", "MyDocs") + 10 < hintScore, true)
        eq("hint before command containing it", FuzzyMatcher.Score("doc", "MyDocs") < hintScore, true)
        eq("open alone not exclusive", FileSearchProvider.Search(SearchQuery("open"))[1].Exclusive, false)
        eq("keyword with space exclusive", FileSearchProvider.Search(SearchQuery("folder "))[1].Exclusive, true)
        ; the hint completes to the keyword typed and a space (Enter / Tab / click); nothing to complete once the space is there
        complete := (text) => FileSearchProvider.Search(SearchQuery(text))[1].AutoComplete
        eq("hint completes: folder", complete("folder"), "folder ")
        eq("hint completes: the keyword typed", complete("Find"), "find ")
        eq("hint completes: file type", complete("doc"), "doc ")
        eq("hint completes: recent", complete("recent"), "recent ")
        eq("keyword with space: no completion", complete("folder "), "")
        ; a keyword not typed in full yet: the same hint at the end of the list
        prefixHint := (text) => FileSearchProvider.Search(SearchQuery(text))
        eq("prefix hint: folder", prefixHint("fol").Length "|" prefixHint("fol")[1].AutoComplete "|" prefixHint("fol")[1].Score, "1|folder |" ResultItem.KeywordPrefixHintScore)
        eq("prefix hint: recent", prefixHint("rec")[1].AutoComplete, "recent ")
        eq("prefix hint: file type", prefixHint("vid")[1].AutoComplete, "video ")
        eq("prefix hint: the keyword it starts", prefixHint("fin")[1].AutoComplete, "find ")
        eq("prefix hint: not exclusive, can't run", prefixHint("fol")[1].Exclusive "|" prefixHint("fol")[1].Valid, "0|0")
        eq("prefix hint: two letters are not enough", prefixHint("fo").Length, 0)
        eq("prefix hint: after every match", ResultItem.KeywordPrefixHintScore < 30 - 4, true)      ; the weakest match (letters in order, long text) scores 26
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
        FileIndex.SetEntries([])
    }

    ; Typing ? shows the cheat sheet; usage tips rotate in the empty search box
    static HelpAndTips() {
        eq := (n, a, e) => TestRunner.Equal("HelpAndTips." n, a, e)
        eq("tips default on", AppSettings.Defaults()["Window"]["ShowTips"], 1)
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
        eq("recent entry listed", InStr(ids, "Recent ") > 0, true)
        ; The most useful entries come first
        eq("order", InStr(ids, "Files Folders Recent FileKeyword FileTypes Reveal Actions QuickSwitch About Prefs Edit EditJson ") = 1, true)
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
        quick := AppSettings.Extension("QuickSwitch"), savedQuick := [quick["Enabled"], quick["Hotkey"]]
        quick["Enabled"] := 1, quick["Hotkey"] := "^g"
        eq("quick switch key", keyOf("QuickSwitch"), "Ctrl+G (in an Open / Save dialog)")
        quick["Hotkey"] := ""
        eq("quick switch without hotkeys", keyOf("QuickSwitch"), "")
        quick["Enabled"] := savedQuick[1], quick["Hotkey"] := savedQuick[2]
        I18n.Init(savedLang)
        savedLang := I18n.Lang                                              ; the pinyin tip is only for the Simplified Chinese UI
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
        ; filtering: the text after ?
        keys := ""
        for item in HelpProvider.Search(SearchQuery("? f3"))
            keys .= item.Title "|"
        eq("filter", keys, "F3|")
        ; placeholders become the keywords from the current settings
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
        ; features that are off aren't shown
        clip := AppSettings.Feature("Clipboard"), savedClip := clip["Enabled"]
        clip["Enabled"] := 0
        ids := ""
        for item in HelpProvider.Items()
            ids .= item.Id " "
        eq("disabled feature hidden", InStr(ids, "Clipboard "), 0)
        clip["Enabled"] := savedClip
        ; tips rotate in order, each gets its turn
        seen := Map()
        Loop items.Length
            seen[HelpProvider.NextTip()] := true
        eq("tips rotate", seen.Count, items.Length)
        tip := HelpProvider.NextTip(), matched := false
        for item in items
            matched := matched || (tip = I18n.T("Help.TipFormat", item.Key, item.Text))
        eq("tip text", matched, true)
    }

    ; Default examples for new users: custom commands of every type, snippets only with supported placeholders; default hotkeys; window lists
    ; Search hotkeys: {selection} becomes the selected text (one line, at most 100 characters), nothing / files selected: left out
    static HotkeySearchText() {
        eq := (n, a, e) => TestRunner.Equal("HotkeySearchText." n, a, e)
        fill := (template, selection) => SystemProvider.FillSearchText(template, selection)
        eq("selected text", fill("folder {selection}", {Text: "  Project X  "}), "folder Project X")
        eq("any case", fill("doc {Selection}", {Text: "budget"}), "doc budget")
        eq("several lines on one", fill("folder {selection}", {Text: "Tower`r`n  B`tlevel 3`n"}), "folder Tower B level 3")
        eq("nothing selected", fill("folder {selection}", ""), "folder ")
        eq("files selected", fill("folder {selection}", {Files: ["C:\a.txt"]}), "folder ")
        eq("dollar kept", fill("g {selection}", {Text: "$1 and $$"}), "g $1 and $$")
        long := ""
        Loop 30
            long .= "word "
        eq("at most 100 characters", StrLen(fill("{selection}", {Text: long})), 99)   ; 100, the space at the end trimmed
        eq("without {selection}: as it is, nothing copied", fill("recent ", "not looked at"), "recent ")
        eq("twice", fill("{selection} {selection}", {Text: "a"}), "a a")
    }

    ; A hotkey with Action "custom:<Id>": runs the custom command; the selection goes to commands that take it, with the panel's checks
    static HotkeyCustomCommand() {
        eq := (n, a, e) => TestRunner.Equal("HotkeyCustomCommand." n, a, e)
        ran := [], notices := [], shown := [], captured := 0
        savedRun := ActionCatalog.GetOwnPropDesc("RunDefault"), savedNotify := App.GetOwnPropDesc("Notify"), savedShow := SearchWindow.GetOwnPropDesc("Show")
        savedCommands := AppSettings.Data["CustomCommands"], savedUsage := Usage.Days.Clone()
        ActionCatalog.DefineProp("RunDefault", {Call: (this, item) => ran.Push(item)})
        App.DefineProp("Notify", {Call: (this, text, *) => notices.Push(text)})
        SearchWindow.DefineProp("Show", {Call: (this, text := "", *) => shown.Push(text)})
        capture(selection) => () => (captured++, selection)
        try {
            AppSettings.Data["CustomCommands"] := [
                Map("Id", "plain001", "Title", "Notepad", "Type", "File", "Target", "notepad.exe", "Arguments", "", "Keyword", "np")
              , Map("Id", "query001", "Title", "Google", "Type", "Url", "Target", "https://www.google.com/search?q={query}", "Arguments", "", "Keyword", "gg")
              , Map("Id", "query002", "Title", "Ping", "Type", "Command", "Target", "cmd.exe", "Arguments", "/k ping {query}", "Keyword", "")
              , Map("Id", "types001", "Title", "Open in Notepad", "Type", "File", "Target", "notepad.exe", "Arguments", "", "Keyword", "", "Selection", "Text files")]
            runHotkey := (id, selection := "") => CustomCommandProvider.RunFromHotkey(id, capture(selection))

            runHotkey("plain001", {Text: "ignored"})
            eq("plain: runs, the selection isn't copied", ran.Length "|" ran[1].Arg "|" ran[1].Arguments "|" captured, "1|notepad.exe||0")
            runHotkey("query001", {Text: "red apple"})
            eq("{query}: the selected text, encoded", ran.Length "|" ran[2].Arg "|" captured, "2|https://www.google.com/search?q=red%20apple|1")
            runHotkey("query001")
            eq("nothing selected, with a keyword: the search window with it", shown.Length "|" shown[1] "|" ran.Length, "1|gg |2")
            runHotkey("query002")
            eq("nothing selected, no keyword: says so", notices.Length "|" notices[1] "|" ran.Length, "1|" I18n.T("Selection.None") "|2")
            runHotkey("query002", {Text: "8.8.8.8 & del x"})
            eq("shell characters for cmd: refused", notices.Length "|" notices[2] "|" ran.Length, "2|" I18n.T("Hotkey.SelectionRefused", "Ping") "|2")
            runHotkey("query002", {Text: "8.8.8.8"})
            eq("safe text: passed", ran.Length "|" ran[3].Arguments, "3|/k ping 8.8.8.8")
            runHotkey("types001", {Files: ["C:\Docs\report.pdf"]})
            eq("content types: the path as the last argument, type not checked", ran.Length "|" ran[4].Arguments, '4|"C:\Docs\report.pdf"')
            runHotkey("types001")
            eq("content types, nothing selected: runs as it is", ran.Length "|" ran[5].Arguments, "5|")
            runHotkey("gone0001", {Text: "x"})
            eq("deleted command: says so, nothing runs", notices.Length "|" notices[3] "|" ran.Length, "3|" I18n.T("Hotkey.CommandMissing") "|5")
            eq("summary names the command", InStr(HotkeyEditor.Summary(Map("Key", "^!n", "Action", "custom:plain001", "WinTitle", "")), "Notepad") > 0, true)
            eq("summary of a deleted command: its Id", InStr(HotkeyEditor.Summary(Map("Key", "^!n", "Action", "custom:gone0001", "WinTitle", "")), "custom:gone0001") > 0, true)
            eq("editor keeps the action", HotkeyEditor.Normalize(Map("Key", "^!n", "Action", "custom:plain001"))["Action"], "custom:plain001")
        } finally {
            ActionCatalog.DefineProp("RunDefault", savedRun), App.DefineProp("Notify", savedNotify), SearchWindow.DefineProp("Show", savedShow)
            AppSettings.Data["CustomCommands"] := savedCommands, Usage.Days := savedUsage
        }
    }

    ; The hotkey edit window: old forms shown as Search, the summary line, and every control inside the window in every language
    static HotkeyEditorWindow() {
        eq := (n, a, e) => TestRunner.Equal("HotkeyEditorWindow." n, a, e)
        item := HotkeyEditor.Normalize(Map("Key", "^!a", "Action", "Activate", "WinTitle", ""))   ; the 2.x name of Show
        eq("old Show: an empty search", item["Action"] "|" item["SearchText"], "Search|")
        item := HotkeyEditor.Normalize(Map("Key", "^!r", "Action", "Reindex"))
        eq("old command name", item["Action"] "|" item["WinTitle"], "RebuildIndex|")
        eq("unknown action kept as it is", HotkeyEditor.Normalize(Map("Key", "^!x", "Action", "FutureCommand"))["Action"] "|" HotkeyEditor.ActionName("FutureCommand", []), "FutureCommand|FutureCommand")
        savedLang := I18n.Lang
        I18n.Init("en")
        eq("summary search", HotkeyEditor.Summary(Map("Key", "^!f", "Action", "Search", "SearchText", "folder {selection}", "WinTitle", ""))
            , 'Ctrl+Alt+F opens ALTRun and searches "folder <selected text>" in all programs.')
        eq("summary empty search, ALTRun", HotkeyEditor.Summary(Map("Key", "~^r", "Action", "Search", "SearchText", "", "WinTitle", "ALTRun"))
            , "Ctrl+R opens ALTRun only in the ALTRun search window.")
        lockName := ""
        for command in SystemProvider.Commands()                            ; the names as the command list has them (made once, in the language of that time)
            if (command["Id"] = "Lock")
                lockName := command["Title"]
        eq("summary command in a window", HotkeyEditor.Summary(Map("Key", "^!t", "Action", "Lock", "WinTitle", "ahk_class CabinetWClass"))
            , 'Ctrl+Alt+T runs "' lockName '" only in windows matching ahk_class CabinetWClass.')
        eq("summary no key", HotkeyEditor.Summary(Map("Key", "", "Action", "Search")), I18n.T("HotkeyEditor.Summary.NoKey"))
        owner := Gui()
        for language in I18n.Languages() {
            I18n.Init(language[1])
            ui := HotkeyEditor.Build(owner, Map("Key", "^!f", "Action", "Search", "SearchText", "folder {selection}", "WinTitle", "ahk_exe WINWORD.EXE"), Map())
            g := ui.Gui
            g.Show("Hide")
            WinGetClientPos(, , &clientW, &clientH, g.Hwnd)
            right := 0, bottom := 0
            for ctrl in g {
                ctrl.GetPos(&x, &y, &w, &h)
                right := Max(right, x + w), bottom := Max(bottom, y + h)
            }
            TestRunner.True("HotkeyEditorWindow." language[1] " fits (" right " x " bottom " in " clientW " x " clientH ")", right <= clientW && bottom <= clientH)
            eq(language[1] " title", g.Title, I18n.T("HotkeyEditor.TitleEdit"))
            eq(language[1] " read back", HotkeyEditor.Read(ui)["Action"] "|" HotkeyEditor.Read(ui)["SearchText"] "|" HotkeyEditor.Read(ui)["WinTitle"] "|" ui.Only.Value
                , "Search|folder {selection}|ahk_exe WINWORD.EXE|1")
            eq(language[1] " summary shown", ui.SummaryText.Value, HotkeyEditor.Summary(HotkeyEditor.Read(ui)))
            eq(language[1] " insert is a drop-down", ControlGetStyle(ui.Insert) & 0xF, 12)   ; BS_SPLITBUTTON
            for ctrl in g                                                   ; the hotkey hint on one line, the check box beside the box
                if (ctrl.Type = "Text" && ctrl.Text = I18n.T("Prefs.HotkeyHint")) {
                    ctrl.GetPos(, , &hintW)
                    TestRunner.True("HotkeyEditorWindow." language[1] " hotkey hint on one line (" Win.TextExtent(ctrl.Hwnd, ctrl.Text) " <= " hintW ")", Win.TextExtent(ctrl.Hwnd, ctrl.Text) <= hintW)
                }
            ui.Key.GetPos(&keyX, &keyY, &keyW), ui.Pass.GetPos(&passX, &passY)
            TestRunner.True("HotkeyEditorWindow." language[1] " keep-function box beside the hotkey box", passX > keyX + keyW && Abs(passY - keyY) <= 4)
            eq(language[1] " action name", ui.Action.Value, I18n.T("HotkeyEditor.Action.Search"))
            ui.Action.GetPos(&actionX, &actionY, &actionW), ui.Choose.GetPos(&chooseX, &chooseY)
            TestRunner.True("HotkeyEditorWindow." language[1] " Choose beside the action box", chooseX > actionX + actionW && Abs(chooseY - actionY) <= 4)
            ui.SetAction.Call("ToggleWindow")                               ; Show / hide: no search text
            eq(language[1] " read without search", HotkeyEditor.Read(ui)["Action"] "|" HotkeyEditor.Read(ui)["SearchText"] "|" ui.SearchText.Enabled, "ToggleWindow||0")
            eq(language[1] " detail", ui.Detail.Value, I18n.T("HotkeyEditor.Detail.Toggle"))
            eq(language[1] " no clash: how to record a hotkey", ui.KeyHint.Value, I18n.T("Prefs.HotkeyHint"))
            eq(language[1] " the line under the box: one line, ends in ...", ControlGetStyle(ui.KeyHint) & 0x400F, 0x400C)   ; SS_ENDELLIPSIS | SS_LEFTNOWORDWRAP
            g.Destroy()
            ; a clash: said on the line under the hotkey box, in place of how to record one; the hotkey being edited isn't a clash with itself
            editing := Map("Key", "^!f", "Action", "Lock", "WinTitle", "")
            inUse := [{Key: "!Space", Label: I18n.T("Prefs.Hotkey"), WinTitle: "", Entry: ""}, {Key: "^!f", Label: "x", WinTitle: "", Entry: editing}
                , {Key: "~^!F", Label: I18n.T("Prefs.Page.Hotkeys") " (" I18n.T("Prefs.Hotkey.Search", "folder {selection} with a rather long search text") ")", WinTitle: "", Entry: ""}]
            ui := HotkeyEditor.Build(owner, editing, editing, "", [], inUse)
            g := ui.Gui
            g.Show("Hide")
            eq(language[1] " clash shown", ui.KeyHint.Value, I18n.T("HotkeyEditor.Clash", inUse[3].Label))
            ui.Key.GetPos(, &keyY, , &keyH), ui.KeyHint.GetPos(, &hintY)
            TestRunner.True("HotkeyEditorWindow." language[1] " clash right under the hotkey box", hintY >= keyY + keyH && hintY <= keyY + keyH + 10)
            ui.Only.Value := 1, ui.WinTitle.Value := "ahk_exe notepad.exe"
            ui.SetAction.Call("Lock")                                       ; the window checks again (as on any change)
            eq(language[1] " only in a window: no clash with global ones", ui.KeyHint.Value, I18n.T("Prefs.HotkeyHint"))
            g.Destroy()
        }
        owner.Destroy()
        I18n.Init(savedLang)
    }

    ; A hotkey that the same key press of another one also triggers in the same windows (the red line of the edit window)
    static HotkeyEditorClash() {
        eq := (n, a, e) => TestRunner.Equal("HotkeyEditorClash." n, a, e)
        label(clash) => IsObject(clash) ? clash.Label : ""
        inUse := [{Key: "!Space", Label: "show", WinTitle: "", Entry: ""}
            , {Key: "^!c", Label: "clipboard", WinTitle: "", Entry: ""}
            , {Key: "^!t", Label: "terminal", WinTitle: "ahk_class CabinetWClass", Entry: ""}]
        hk(key, winTitle := "") => Map("Key", key, "Action", "Lock", "WinTitle", winTitle)
        eq("same key", label(HotkeyEditor.Clash(hk("!Space"), inUse)), "show")
        eq("~ and order of modifiers don't matter", label(HotkeyEditor.Clash(hk("~!^C"), inUse)), "clipboard")
        eq("left Alt is one of either Alt", label(HotkeyEditor.Clash(hk("<!Space"), inUse)), "show")
        eq("other key", label(HotkeyEditor.Clash(hk("!r"), inUse)), "")
        eq("no key", label(HotkeyEditor.Clash(hk(""), inUse)), "")
        eq("limited to a window over a global one: fine", label(HotkeyEditor.Clash(hk("!Space", "ahk_exe notepad.exe"), inUse)), "")
        eq("the same window", label(HotkeyEditor.Clash(hk("^!T", " ahk_class CabinetWClass"), inUse)), "terminal")
        eq("a global one over one limited to a window: fine", label(HotkeyEditor.Clash(hk("^!t"), inUse)), "")

        ; the hotkeys in use: the fixed ones as settings, custom ones named as in the list (window-limited ones too); the one being edited left out
        savedLang := I18n.Lang
        I18n.Init("en")
        data := AppSettings.Defaults()
        data["CustomCommands"] := [Map("Id", "plain001", "Title", "Notepad", "Type", "File", "Target", "notepad.exe")]
        own := Map("Key", "^!n", "Action", "custom:plain001", "WinTitle", "")
        data["Hotkeys"] := [own, Map("Key", "^!t", "Action", "custom:plain001", "WinTitle", "ahk_exe code.exe")]
        list := PreferencesWindow.AllHotkeys(data)
        keys := ""
        for each in list
            keys .= each.Key "/" each.WinTitle "|"
        eq("all hotkeys", keys, "!Space/|!r/|^!\/|^!c/|^!n/|^!t/ahk_exe code.exe|")
        eq("custom one named as in the list", list[5].Label, "Hotkeys (Run: Notepad)")
        others := HotkeyEditor._Others(list, own)
        eq("the one being edited left out", others.Length, list.Length - 1)
        eq("not a clash with itself", label(HotkeyEditor.Clash(own, others)), "")
        eq("a new one with its keys", label(HotkeyEditor.Clash(hk("^!n"), HotkeyEditor._Others(list, ""))), "Hotkeys (Run: Notepad)")
        eq("the boxes not applied yet", PreferencesWindow.AllHotkeys(data, (path) => (path = "General.Hotkey") ? "#Space" : PreferencesWindow.GetPath(data, path))[1].Key, "#Space")
        data["Features"]["Clipboard"]["Enabled"] := false
        keys := ""
        for each in PreferencesWindow._GlobalHotkeys(data)
            keys .= each.Key "|"
        eq("global ones; clipboard history off", keys, "!Space|!r|^!\|^!n|")

        ; the hotkeys set on other pages (the Hotkeys page's "Other hotkeys"): only those in effect, with the page they are set on
        list := ""
        row := (each) => each.Key "/" StrReplace(each.Page, "Prefs.Page.") "/" (each.Global ? "all" : "some") "|"
        data["Features"]["Clipboard"]["Enabled"] := true
        data["Extensions"]["QuickSwitch"]["Enabled"] := true, data["Extensions"]["QuickSwitch"]["Hotkey"] := "^g", data["Extensions"]["QuickSwitch"]["MenuHotkey"] := "^+g"
        data["Extensions"]["AutoDate"]["Enabled"] := true, data["Extensions"]["AutoDate"]["Hotkey"] := "^d"
        for each in PreferencesWindow.OtherHotkeys(data)
            list .= row(each)
        eq("other hotkeys, all on", list, "!Space/General/all|!r/General/all|^!\/Selection/all|^!c/Clipboard/all|^g/QuickSwitch/some|^+g/QSPanel/some|^d/DateStamp/some|")
        eq("other hotkeys: what each is for", PreferencesWindow.OtherHotkeys(data)[4].Purpose "|" PreferencesWindow.OtherHotkeys(data)[7].Purpose, "Clipboard history|Add the date")
        data["Features"]["Clipboard"]["Enabled"] := false, data["Extensions"]["QuickSwitch"]["Enabled"] := false, data["Extensions"]["AutoDate"]["Enabled"] := false
        data["General"]["SecondaryHotkey"] := ""
        list := ""
        for each in PreferencesWindow.OtherHotkeys(data)
            list .= row(each)
        eq("other hotkeys: features off and an empty box not listed", list, "!Space/General/all|^!\/Selection/all|")
        list := ""
        for each in PreferencesWindow.OtherHotkeys(data, (path) => (path = "Extensions.AutoDate.Enabled") ? true : PreferencesWindow.GetPath(data, path))
            list .= row(each)
        eq("other hotkeys: as the boxes are now (not applied yet)", list, "!Space/General/all|^!\/Selection/all|^d/DateStamp/some|")
        I18n.Init(savedLang)
    }

    ; Choosing a hotkey's command: the names and the line below the name, the list (order, filtering), and the window in every language
    static HotkeyCommandPicker() {
        eq := (n, a, e) => TestRunner.Equal("HotkeyCommandPicker." n, a, e)
        savedLang := I18n.Lang
        I18n.Init("en")
        commands := [
            Map("Id", "plain001", "Title", "Notepad", "Type", "File", "Target", "notepad.exe", "Arguments", "", "Keyword", "np")
          , Map("Id", "query001", "Title", "Google", "Type", "Url", "Target", "https://www.google.com/search?q={query}", "Arguments", "", "Keyword", "gg")
          , Map("Id", "types001", "Title", "Open in Notepad", "Type", "File", "Target", "notepad.exe", "Arguments", "", "Keyword", "", "Selection", "Text files")
          , Map("Title", "Just added", "Type", "File", "Target", "calc.exe", "Arguments", "", "Keyword", "ja")]   ; added in Preferences, no Id yet

        eq("name: custom", HotkeyEditor.ActionName("custom:plain001", commands), "Notepad")
        eq("name: deleted", HotkeyEditor.ActionName("custom:gone0001", commands), "(deleted command) custom:gone0001")
        eq("name: own", HotkeyEditor.ActionName("Search", commands), I18n.T("HotkeyEditor.Action.Search"))
        eq("summary from the commands given", InStr(HotkeyEditor.Summary(Map("Key", "^!g", "Action", "custom:query001", "WinTitle", ""), commands), '"Google"') > 0, true)
        detail := HotkeyEditor.ActionDetail("custom:plain001", commands)
        eq("detail: plain", detail.Text "|" detail.Missing, "Custom command · notepad.exe · doesn't use the selection|0")
        eq("detail: {query}", HotkeyEditor.ActionDetail("custom:query001", commands).Text, "Custom command · https://www.google.com/search?q={query} · " I18n.T("HotkeyEditor.Detail.Query"))
        eq("detail: content types", HotkeyEditor.ActionDetail("custom:types001", commands).Text, "Custom command · notepad.exe · " I18n.T("HotkeyEditor.Detail.Types"))
        detail := HotkeyEditor.ActionDetail("custom:gone0001", commands)
        eq("detail: deleted", detail.Text "|" detail.Missing, I18n.T("HotkeyEditor.Detail.Missing") "|1")

        choices := CommandPicker.Choices(commands, [])
        eq("own first, then custom (with an Id), then system", choices[1].Id "|" choices[2].Id "|" choices[3].Id "|" choices[5].Id "|" choices[6].Kind
            , "Search|ToggleWindow|custom:plain001|custom:types001|" I18n.T("CommandPicker.Kind.System"))
        ids(list) => list.Length ? list[1].Id : ""
        firstOf(list, kind) {
            for choice in list
                if (choice.Kind = kind)
                    return choice.Id
            return ""
        }
        eq("filter all", CommandPicker.Filter(choices, "  ").Length, choices.Length)
        eq("filter by keyword", ids(CommandPicker.Filter(choices, "gg")), "custom:query001")
        filtered := CommandPicker.Filter(choices, "notepad")
        eq("filter: best first within a kind", firstOf(filtered, I18n.T("CommandPicker.Kind.Custom")), "custom:plain001")
        eq("filter: system commands by English name", firstOf(CommandPicker.Filter(choices, "lock"), I18n.T("CommandPicker.Kind.System")), "Lock")
        eq("filter: nothing", CommandPicker.Filter(choices, "zzqqxx").Length, 0)

        AppSettings.EnsureIds(Map("CustomCommands", commands))              ; what Choose does first
        eq("a new command gets an Id to be chosen", AppSettings.IsId(commands[4]["Id"]) "|" (CommandPicker.Choices(commands, []).Length = choices.Length + 1), "1|1")

        owner := Gui()
        for language in I18n.Languages() {
            I18n.Init(language[1])
            ui := CommandPicker.Build(owner, "custom:types001", commands)
            g := ui.Gui
            g.Show("Hide")
            WinGetClientPos(, , &clientW, &clientH, g.Hwnd)
            right := 0, bottom := 0
            for ctrl in g {
                ctrl.GetPos(&x, &y, &w, &h)
                right := Max(right, x + w), bottom := Max(bottom, y + h)
            }
            TestRunner.True("HotkeyCommandPicker." language[1] " fits (" right " x " bottom " in " clientW " x " clientH ")", right <= clientW && bottom <= clientH)
            eq(language[1] " current selected", CommandPicker.Selected(ui), "custom:types001")
            total := 0
            Loop 3
                total += SendMessage(0x101D, A_Index - 1, 0, ui.List)        ; LVM_GETCOLUMNWIDTH
            ui.List.GetPos(, , &listW)
            TestRunner.True("HotkeyCommandPicker." language[1] " columns fit without a horizontal scroll bar (" total " < " listW ")", total < listW - SysGet(2))   ; SM_CXVSCROLL
            row := ui.List.GetNext()
            eq(language[1] " down key moves in the list", CommandPicker._OnKeyDown(ui, 0x28, 0, 0x100, ui.Search.Hwnd) "|" ui.List.GetNext(), "0|" (row + 1))
            eq(language[1] " other keys stay in the box", CommandPicker._OnKeyDown(ui, 0x41, 0, 0x100, ui.Search.Hwnd) "|" ui.List.GetNext(), "|" (row + 1))
            ui.Release.Call()
            g.Destroy()
        }
        owner.Destroy()
        I18n.Init(savedLang)
    }

    ; Preferences → Hotkeys shows what each hotkey runs by name, a deleted custom command in red; deleting a command deletes its hotkeys
    static HotkeyCommandList() {
        eq := (n, a, e) => TestRunner.Equal("HotkeyCommandList." n, a, e)
        savedLang := I18n.Lang
        I18n.Init("en")
        commands := [Map("Id", "plain001", "Title", "Notepad", "Type", "File", "Target", "notepad.exe", "Arguments", "", "Keyword", "np")
                   , Map("Id", "other001", "Title", "Calculator", "Type", "File", "Target", "calc.exe", "Arguments", "", "Keyword", "")]
        runKey := Map("Key", "^!n", "Action", "custom:plain001", "WinTitle", "")
        gone := Map("Key", "^!q", "Action", "custom:gone0001", "WinTitle", "")
        eq("custom command", HotkeyEditor.ListText(runKey, commands), "Run: Notepad")
        eq("deleted command", HotkeyEditor.ListText(gone, commands), "(deleted command) custom:gone0001")
        eq("system command by name", HotkeyEditor.ListText(Map("Key", "^!t", "Action", "TerminalHere"), commands), HotkeyEditor.ActionName("TerminalHere"))
        eq("show / hide", HotkeyEditor.ListText(Map("Key", "^!h", "Action", "ToggleWindow"), commands), "Show / hide ALTRun")
        eq("old Show: an empty search", HotkeyEditor.ListText(Map("Key", "^!s", "Action", "Show"), commands), "Search in ALTRun")
        eq("missing", HotkeyEditor.IsMissing(runKey, commands) "|" HotkeyEditor.IsMissing(gone, commands) "|" HotkeyEditor.IsMissing(Map("Action", "Lock"), commands), "0|1|0")

        savedWorking := PreferencesWindow.Working
        PreferencesWindow.Working := Map("CustomCommands", commands, "Hotkeys", [runKey, gone])
        try {
            eq("list cell: the commands being edited", PrefsForm.Cell(runKey, "Action"), "Run: Notepad")
            eq("row color: red only when deleted", PrefsPages.HotkeyRowColor(runKey) "|" PrefsPages.HotkeyRowColor(gone), "|" 0x1C2BC4)
        } finally
            PreferencesWindow.Working := savedWorking

        hotkeys := [runKey, gone, Map("Key", "^!m", "Action", "custom:plain001", "WinTitle", "ahk_exe WINWORD.EXE"), Map("Key", "^!c", "Action", "custom:other001")]
        eq("hotkeys of a command", CustomCommandProvider.HotkeysFor(commands[1], hotkeys).Length, 2)
        eq("question names them", CustomCommandProvider.DeleteQuestion(commands[1], hotkeys), "Delete 'Notepad'?`n`nThe hotkeys that run it are deleted too: Ctrl+Alt+N, Ctrl+Alt+M")
        eq("no hotkey: the usual question", CustomCommandProvider.DeleteQuestion(Map("Id", "none0001", "Title", "X"), hotkeys), "")
        removed := CustomCommandProvider.RemoveHotkeys(commands[1], hotkeys)
        eq("removed in place", removed.Length "|" hotkeys.Length "|" hotkeys[1]["Action"] "|" hotkeys[2]["Action"], "2|2|custom:gone0001|custom:other001")

        ; Ctrl+Del in the results: the command and its hotkeys go, the hotkeys are turned off, the settings saved
        savedCommands := AppSettings.Data["CustomCommands"], savedHotkeys := AppSettings.Data["Hotkeys"], savedFile := AppSettings.File
        savedUnregister := App.GetOwnPropDesc("UnregisterHotkey")
        turnedOff := []
        App.DefineProp("UnregisterHotkey", {Call: (this, entry) => turnedOff.Push(entry["Key"])})
        AppSettings.File := A_Temp "\ALTRunTest.json"                     ; deleting saves the settings: write to a temporary file
        try {
            AppSettings.Data["CustomCommands"] := [commands[1].Clone(), commands[2].Clone()]
            AppSettings.Data["Hotkeys"] := [runKey.Clone(), Map("Key", "^!c", "Action", "custom:other001", "WinTitle", "")]
            item := {Source: AppSettings.Data["CustomCommands"][1], Title: "Notepad"}
            eq("search prompt names the hotkey", CustomCommandProvider.DeletePrompt(item), "Delete 'Notepad'?`n`nThe hotkeys that run it are deleted too: Ctrl+Alt+N")
            eq("search prompt without hotkeys", CustomCommandProvider.DeletePrompt({Source: Map("Id", "none0001", "Title", "Paint"), Title: "Paint"}), I18n.T("Search.ConfirmDelete", "Paint"))
            CustomCommandProvider.DeleteItem(item)
            eq("deleted with its hotkey, turned off", AppSettings.CustomCommands.Length "|" AppSettings.Hotkeys.Length "|" AppSettings.Hotkeys[1]["Key"] "|" turnedOff.Length "|" turnedOff[1], "1|1|^!c|1|^!n")
        } finally {
            App.DefineProp("UnregisterHotkey", savedUnregister)
            AppSettings.Data["CustomCommands"] := savedCommands, AppSettings.Data["Hotkeys"] := savedHotkeys, AppSettings.File := savedFile
            try FileDelete(A_Temp "\ALTRunTest.json")
            CustomCommandProvider._ResetNarrowing()
        }
        App.UnregisterHotkey(Map("Key", "^!F24", "WinTitle", "ALTRun"))      ; never registered: no error
        I18n.Init(savedLang)
    }

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
        commandIds["Search"] := true                                        ; the Ids only for hotkeys
        scopes := ""
        prefixes := ""
        for entry in defaults["Hotkeys"] {                                  ; F1 ~ F4 are built into the search window, not custom hotkeys
            TestRunner.True("DefaultExamples.no F-key hotkey " entry["Key"], !RegExMatch(entry["Key"], "i)^F[1-4]$"))
            TestRunner.True("DefaultExamples.hotkey action " entry["Action"], commandIds.Has(entry["Action"]))
            eq("hotkey key order " entry["Key"], HotkeyBox.Compose(RegExReplace(entry["Key"], "i)\w+$"), RegExReplace(entry["Key"], "^[\^!+#]+")), entry["Key"])   ; as Preferences would record it
            scopes .= (entry["WinTitle"] = "" ? "global" : entry["WinTitle"] = "ALTRun" ? "altrun" : "window") " "
            if (entry["Action"] = "Search")
                prefixes .= entry["SearchText"] "|"
        }
        eq("hotkey examples: global, one program, ALTRun", scopes, "global global global global window altrun ")
        eq("hotkey examples: Search with the selected text", prefixes, "folder {selection}|recent |doc {selection}|")
        eq("hotkey search text: Search types it", App.HotkeySearchText(Map("Key", "^!f", "Action", "Search", "SearchText", "folder {selection}", "WinTitle", "")), "folder {selection}")
        eq("hotkey search text: only for Search", App.HotkeySearchText(Map("Key", "^!f", "Action", "Lock", "SearchText", "folder ")), "")
        eq("hotkey search text: none set", App.HotkeySearchText(Map("Key", "^!f", "Action", "Search")), "")
        eq("hotkey plain Show stays", App.HotkeyActionId(Map("Key", "^!f", "Action", "Show")), "Show")
        eq("no mouse hotkeys by default (no mouse hook)", App.UsesMouseHotkeys(defaults["Hotkeys"]), false)
        help := ""
        for item in HelpProvider.Items()
            help .= item.Id " "
        eq("help lists F1 and F4", (InStr(help, "About ") && InStr(help, "EditJson ")) ? 1 : 0, 1)
        eq("hotkey action shows its search text", PrefsForm.Cell(Map("Key", "^!f", "Action", "Search", "SearchText", "folder {selection}"), "Action"), I18n.T("Prefs.Hotkey.Search", "folder {selection}"))
        eq("hotkey action without text", PrefsForm.Cell(Map("Key", "^!f", "Action", "Search", "SearchText", " "), "Action"), I18n.T("HotkeyEditor.Action.Search"))
        eq("hotkey search text only for Search; a system command by name", PrefsForm.Cell(Map("Key", "^!f", "Action", "Lock", "SearchText", "folder "), "Action"), HotkeyEditor.ActionName("Lock"))
        eq("command target shows arguments", PrefsForm.Cell(Map("Target", "cmd.exe", "Arguments", "/k ipconfig /all"), "Target"), "cmd.exe /k ipconfig /all")
        eq("url type name", I18n.Strings["Prefs.Type.Url"], "Web address / link")
        ids := Map()
        for command in SystemProvider.Commands()
            ids[command["Id"]] := true
        eq("about command exists", ids.Has("About"), true)
        eq("auto switch exclude default", defaults["Extensions"]["QuickSwitch"]["AutoSwitchExclude"].Length, 0)
        eq("standard dialogs default", defaults["Extensions"]["QuickSwitch"]["StandardDialogs"] "|" defaults["Extensions"]["QuickSwitch"]["DialogWindows"].Length, "1|0")
    }

    ; Switching windows: another process shows a window, "w title" finds it, then it is closed; our own windows aren't listed
    static WindowSwitch() {
        eq := (n, a, e) => TestRunner.Equal("WindowSwitch." n, a, e)
        script := A_Temp "\ALTRun-window-test.ahk", title := "ALTRun Window Test " A_TickCount
        try FileDelete(script)
        FileAppend('#NoTrayIcon`ng := Gui(, "' title '")`ng.OnEvent("Close", (*) => ExitApp())`ng.Show("w300 h120")'
            . '`ng2 := Gui(, "' title ' second")`ng2.Show("x10 y10 w300 h120")`nSetTimer(() => ExitApp(), -20000)', script, "UTF-8")
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
        WinWait(title " second", , 5)
        WindowProvider._listTime := 0
        both := 0
        for candidate in ProviderRegistry._Dedupe(WindowProvider.Search(SearchQuery("w " SubStr(title, 1, 18))))   ; as the results are merged
            both += InStr(candidate.Title, title) ? 1 : 0
        eq("two windows of one program: both", both, 2)
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
            item.Actions[1].OnRun.Call()                                     ; "Close window" in the actions panel
            eq("closed", WinWaitClose(title, , 5), 1)
        }
        try ProcessClose(pid)
        try FileDelete(script)
        AppSettings.Feature("Windows")["InDefaultResults"] := 0
        eq("default results off", WindowProvider.Search(SearchQuery("ALTRun Window")).Length, 0)
        AppSettings.Feature("Windows")["InDefaultResults"] := 1
    }

    ; Empty search box: pinned first, then recently opened (deduplicated, limited count); only results that can be opened again are kept
    ; {hwnd} / {pid} in a custom command (#445): left in the result, filled in with the window in front before ALTRun when it runs,
    ; also for pinned / recently used entries; empty when that window is gone
    static WindowPlaceholders() {
        eq := (n, a, e) => TestRunner.Equal("WindowPlaceholders." n, a, e)
        g := Gui("+ToolWindow", "ALTRun window placeholder test")
        g.Show("x-3000 y-3000 w120 h40 NoActivate")
        hwnd := g.Hwnd, pid := DllCall("GetCurrentProcessId")
        eq("fill", ActionCatalog.FillWindow("/c {hwnd} {PID} {Hwnd}", hwnd), "/c " hwnd " " pid " " hwnd)
        eq("nothing to fill", ActionCatalog.FillWindow("/k ping {query}", hwnd), "/k ping {query}")
        eq("no window", ActionCatalog.FillWindow("{hwnd}|{pid}", 0), "|")
        eq("uses window", ActionCatalog.UsesWindow("cmd.exe", "/c {pid}") "|" ActionCatalog.UsesWindow("cmd.exe", "/c {query}"), "1|0")

        command := Map("Title", "Window Info", "Type", "Command", "Target", "cmd.exe", "Arguments", "/c echo {hwnd} {pid}", "Keyword", "wi")
        item := CustomCommandProvider._ToItem(command, 0)
        eq("result keeps placeholders", item.WindowArgs "|" item.Arguments, "1|/c echo {hwnd} {pid}")
        eq("plain command", CustomCommandProvider._ToItem(Map("Title", "t", "Type", "File", "Target", "notepad.exe"), 0).WindowArgs, false)
        eq("path not checked", CustomCommandProvider.CheckTarget(Map("Title", "t", "Type", "File", "Target", "C:\Tools\{pid}.txt")), "Skipped")

        savedPrevious := App.PreviousWindow
        App.PreviousWindow := hwnd
        filled := ActionCatalog.ForWindow(item)
        eq("filled when run", filled.Kind "|" filled.Arg "|" filled.Arguments "|" filled.WindowArgs, "file|cmd.exe|/c echo " hwnd " " pid "|0")
        eq("others unchanged", ActionCatalog.ForWindow(item := ResultItem("a", "", {Kind: "file", Arg: "C:\a{pid}.txt"})) == item, true)

        entry := RecentProvider.Snapshot(CustomCommandProvider._ToItem(command, 0).DefineProp("Provider", {Value: "CustomCommands"}))
        eq("saved with placeholders", entry["Arguments"] "|" entry["WindowArgs"], "/c echo {hwnd} {pid}|1")
        again := RecentProvider.ItemFor(entry, true)
        eq("recent entry filled when run", ActionCatalog.ForWindow(again).Arguments, "/c echo " hwnd " " pid)
        g.Destroy()
        eq("window gone", ActionCatalog.ForWindow(again).Arguments, "/c echo  ")
        App.PreviousWindow := savedPrevious
    }

    static RecentItems() {
        eq := (n, a, e) => TestRunner.Equal("RecentItems." n, a, e)
        eq("default count", AppSettings.Defaults()["Window"]["RecentCount"], 0)   ; by default only pinned items show
        saved := {Recent: Knowledge.Recent, File: Knowledge.File, SettingsFile: AppSettings.File}
        options := AppSettings.Feature("Recent"), savedPinned := options["Pinned"], savedCount := AppSettings.Window["RecentCount"]
        Knowledge.Recent := [], Knowledge.File := A_Temp "\ALTRun-recent-test.json", AppSettings.File := A_Temp "\ALTRun-recent-settings.json"
        options["Pinned"] := [], AppSettings.Window["RecentCount"] := 2
        fileItem := (name) => ResultItem(name, "C:\Docs\" name, {Kind: "file", Arg: "C:\Docs\" name, Uid: "file:" name, Provider: "FileSearch", Icon: "C:\Docs\" name})
        RecentProvider.Remember(fileItem("a.txt"))
        RecentProvider.Remember(fileItem("b.txt"))
        RecentProvider.Remember(fileItem("c.txt"))
        RecentProvider.Remember(fileItem("a.txt"))                               ; opened again: moves to the top, no duplicate
        clip := ResultItem("secret", "", {Kind: "text", Arg: "secret", Uid: "clip:1", Provider: "Clipboard"})
        RecentProvider.Remember(clip)                                        ; the clipboard history isn't kept
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
        for action in ResultActions.ListFor(items[2])
            actions .= action.Title "|"
        eq("pin action", InStr(actions, I18n.T("Action.Pin")) > 0, true)
        RecentProvider.DeleteItem(items[3])                                  ; Ctrl+Del: removed from the recently used
        eq("forget", titles(RecentProvider.EmptyResults()), "c.txt|" I18n.T("Sys.Lock") "|b.txt")
        RecentProvider.DeleteItem(items[1])                                  ; pinned: unpinned
        eq("unpin", options["Pinned"].Length, 0)
        AppSettings.Window["RecentCount"] := 0
        eq("count 0", RecentProvider.EmptyResults().Length, 0)
        Knowledge.Recent := saved.Recent, Knowledge.File := saved.File, AppSettings.File := saved.SettingsFile
        options["Pinned"] := savedPinned, AppSettings.Window["RecentCount"] := savedCount
        try FileDelete(A_Temp "\ALTRun-recent-test.json")
        try FileDelete(A_Temp "\ALTRun-recent-settings.json")
    }

    ; Where the clipboard is stored (LocalFiles): 1 = the Clipboard folder moves to this PC, ClipboardHistory.json stays in Data (syncs as usual);
    ; 2 = the json moves to this PC too; 0 = everything back to Data; nothing overwritten when both exist; all in Data without a local folder
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
        ClipboardProvider.PlaceStorage(1, data, localDir)                      ; already on this PC: nothing to do
        eq("1 again", where(), "D- -L")
        ClipboardProvider.PlaceStorage(2, data, localDir)
        eq("2 json", ClipboardProvider.File, localDir "\ClipboardHistory.json")
        eq("2 files", where(), "-L -L")
        ClipboardProvider.PlaceStorage(0, data, localDir)
        eq("0 json", ClipboardProvider.File, data "\ClipboardHistory.json")
        eq("0 folder", ClipboardProvider.Folder, data "\Clipboard")
        eq("0 files", where(), "D- D-")
        FileAppend('{"Entries": []}', localDir "\ClipboardHistory.json", "UTF-8")   ; both exist: nothing overwritten, the new location's is used
        ClipboardProvider.PlaceStorage(0, data, localDir)
        eq("both kept", where(), "DL D-")
        ClipboardProvider.PlaceStorage(2, data, "")                            ; no %LOCALAPPDATA%: all in Data
        eq("no local", ClipboardProvider.File "|" ClipboardProvider.Folder, data "\ClipboardHistory.json|" data "\Clipboard")
        eq("default", AppSettings.Defaults()["Features"]["Clipboard"]["LocalFiles"], 2)
        old := Map("Features", Map("Clipboard", Map("Persist", 1)))          ; a settings file from before the default changed: stays 1
        eq("existing file keeps 1", AppSettings._KeepPrevious(old) "|" old["Features"]["Clipboard"]["LocalFiles"], "1|1")
        old := Map("Features", Map("Help", Map("Enabled", 1)))
        eq("existing file without Clipboard", AppSettings._KeepClipboardStorage(old) "|" old["Features"]["Clipboard"]["LocalFiles"], "1|1")
        set := Map("Features", Map("Clipboard", Map("LocalFiles", 0)))
        eq("already set", AppSettings._KeepClipboardStorage(set) "|" set["Features"]["Clipboard"]["LocalFiles"], "0|0")
        eq("new file", AppSettings._KeepClipboardStorage(Map()), false)
        savedImport := AppSettings.ImportedFrom, AppSettings.ImportedFrom := "ALTRun.ini"   ; imported from 2.x (no clipboard history there): the new default
        eq("imported from 2.x", AppSettings._KeepClipboardStorage(Map("Features", Map("Help", Map("Enabled", 1)))), false)
        AppSettings.ImportedFrom := savedImport
        ; turning off saving the history may delete what was saved (the list, its .bad copy, long entries and images)
        ClipboardProvider.PlaceStorage(0, data, localDir)
        FileAppend("x", ClipboardProvider.File ".bad")
        eq("delete saved: all gone", ClipboardProvider.DeleteSaved() "|" FileExist(ClipboardProvider.File ".bad") "|" (FileExist(data "\ClipboardHistory.json") || DirExist(data "\Clipboard")), "1||")
        ClipboardProvider.File := saved[1], ClipboardProvider.Folder := saved[2]
        try DirDelete(root, true)
    }

    ; Pinned clipboard entries: listed first, not removed over the limit, still pinned when copied again, kept by Clear, still there after saving
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
        ClipboardProvider.Entries[2]["Pinned"] := 1                          ; "one" pinned
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
        ClipboardProvider.Add("one")                                         ; copied again: still pinned
        eq("copy again keeps pin", ClipboardProvider.IsPinned(ClipboardProvider.Entries[1]), true)
        ClipboardProvider.Save()
        ClipboardProvider.Entries := [], ClipboardProvider._Load()
        eq("saved", texts(), "one|five|four")
        counts := ClipboardProvider.Counts()                                 ; for the question before "Clear History Now"
        eq("counts", counts.Total " " counts.Pinned, "3 1")
        ClipboardProvider.Clear()
        eq("clear keeps pinned", texts(), "one")
        ClipboardProvider.Entries[1].Delete("Pinned"), ClipboardProvider.Clear()
        eq("unpinned cleared", ClipboardProvider.Entries.Length, 0)
        ClipboardProvider.File := saved[1], ClipboardProvider.Folder := saved[2], ClipboardProvider.Entries := saved[3]
        options["MaxItems"] := savedMax, options["Persist"] := savedPersist
        try DirDelete(root, true)
    }

    ; Scripts extension: @altrun.xxx read from the leading comments, found by name / keyword, arguments, command lines, output of background runs
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
        eq("script keyword: all scripts", titles("script "), "Clean Temp|Ping Host|Restart Explorer")
        eq("script keyword: by name", titles("script expl"), "Restart Explorer")
        onlyScripts := ScriptProvider.Search(SearchQuery("script clean"))
        eq("script keyword: exclusive", onlyScripts.Length "|" onlyScripts[1].Exclusive, "1|1")
        hint := ScriptProvider.Search(SearchQuery("script"))                ; without the space: only a hint
        eq("keyword without space: a hint", hint.Length "|" hint[1].Valid "|" hint[1].Exclusive "|" hint[1].Subtitle, "1|0|0|script ...")
        eq("keyword without space: hint score", hint[1].Score, ResultItem.KeywordHintScore)
        eq("keyword without space: Enter / Tab complete", hint[1].AutoComplete, "script ")
        prefix := ScriptProvider.Search(SearchQuery("scrip"))               ; not typed in full yet: the hint at the end of the list
        eq("keyword start: a hint", prefix.Length "|" prefix[1].Valid "|" prefix[1].AutoComplete "|" prefix[1].Score, "1|0|script |" ResultItem.KeywordPrefixHintScore)
        eq("keyword start: two letters are not enough", ScriptProvider.Search(SearchQuery("sc")).Length, 0)
        eq("keyword with space: no hint", titles("script "), "Clean Temp|Ping Host|Restart Explorer")
        eq("default script keyword", AppSettings.Defaults()["Features"]["Scripts"]["Keyword"], "script")
        eq("ps1 command", ScriptProvider.CommandLine(byTitle["Ping Host"], "a b"), 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' root '\ping.ps1" "a b"')
        eq("bat command", ScriptProvider.CommandLine(byTitle["Clean Temp"]), A_ComSpec ' /c ""' root '\clean.bat""')
        ; arguments as programs read them; through cmd.exe (batch files, captured output) a quote isn't safe and is refused
        eq("argument: quote and trailing backslash", ScriptProvider.CommandLine(byTitle["Ping Host"], 'say "hi" C:\'), 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' root '\ping.ps1" "say \"hi\" C:\\"')
        eq("through cmd", ScriptProvider.ThroughCmd({Ext: "bat", Mode: "window"}) "|" ScriptProvider.ThroughCmd({Ext: "ps1", Mode: "output"}) "|" ScriptProvider.ThroughCmd({Ext: "ps1", Mode: "window"}), "1|1|0")
        ; output files of earlier runs are deleted after a day
        FileAppend("old", root "\ALTRun-script-1.txt"), FileAppend("new", root "\ALTRun-script-2.txt")
        FileSetTime(DateAdd(A_Now, -25, "Hours"), root "\ALTRun-script-1.txt", "M")
        ScriptProvider.CleanOutputs(root)
        eq("old output deleted, new kept", (FileExist(root "\ALTRun-script-1.txt") ? 1 : 0) "|" (FileExist(root "\ALTRun-script-2.txt") ? 1 : 0), "0|1")
        FileAppend("line one`r`n`r`nlast line  `r`n", root "\out.txt")
        eq("silent: last line", ScriptProvider._Finished(byTitle["Restart Explorer"], root "\out.txt"), "last line")
        FileAppend("", root "\empty.txt")
        eq("silent: no output", ScriptProvider._Finished(byTitle["Restart Explorer"], root "\empty.txt"), I18n.T("Script.Done", "Restart Explorer"))
        FileAppend("; @altrun.title New One`n", root "\new.ahk", "UTF-8")
        ScriptProvider._checked := 0                                         ; the folder changed: read again at the next search
        eq("refresh", titles("new one"), "New One")
        ScriptProvider.Folder := saved[1], ScriptProvider.Scripts := saved[2]
        ; Opening the Scripts folder the first time: three examples (one form each) are created and all can be read; an existing folder is left alone
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
            eq("examples: count", found.Count, 4)
            eq("example filter", found.Get("Environment Variables.ahk", ""), "Example: Environment Variables|filter|envvar|part of a name")
            eq("example ahk", found.Get("Today.ahk", ""), "Example: Today's Date|silent||")
            eq("example bat", found.Get("Port.bat", ""), "Example: Who Uses a Port|window|port|a port number")
            eq("example ps1", found.Get("IP Addresses.ps1", ""), "Example: IP Addresses|output||")
            bom := FileOpen(examples "\Port.bat", "r", "CP0"), first := bom.RawRead(head := Buffer(3), 3) ? NumGet(head, 0, "UChar") : 0, bom.Close()
            eq("example bat has no BOM", first, Ord("@"))
            for name, text in ScriptProvider.ExampleScripts()
                TestRunner.True("Scripts.example ASCII only " name, !RegExMatch(text, "[^\x00-\x7F]"))
            ; While waiting for an argument the subtitle says what to type next (scripts, custom commands, web searches)
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

    ; A filter script as ScriptProvider.ReadScript makes it, in C:\s (ScriptFilter tests)
    static _FilterScript(name, keyword) {
        SplitPath(name, , , &ext, &baseName)
        return {Path: "C:\s\" name, Name: name, Ext: ext, Title: StrTitle(baseName), Keyword: keyword, Argument: "", Mode: "filter", Selection: [], FilterBy: "script", Keys: []}
    }

    ; Live results of a filter script (ScriptFilter): what it prints -> results (JSON or one per line), the type guessed from the
    ; target, web icons dropped, at most MaxItems; a run's result -> items, the error row or "No results"
    static ScriptFilterParse() {
        eq := (n, a, e) => TestRunner.Equal("ScriptFilterParse." n, a, e)
        show(items) {
            list := ""
            for each in items
                list .= each.Title "|" each.Type "|" each.Arg "|" each.Icon "|" each.Uid "|" each.Valid "|" each.AutoComplete "|" each.LargeText "/"
            return list
        }
        items := ScriptFilter.Parse('`n {"items": [{"title": "A", "subtitle": "s", "arg": "C:\\x\\", "uid": "a", "icon": {"path": "C:\\i.ico"}, "text": {"largetype": "L"}},'
            . ' {"title": ""}, {"title": "Web", "arg": "https://e.test", "icon": "https://e.test/i.png", "valid": false, "autocomplete": "web"}, {"title": "T", "type": "default"}, 5, {"title": 7, "extra": [1]}]}')
        eq("JSON", show(items), "A|folder|C:\x\|C:\i.ico|a|1||L/Web|url|https://e.test||" "|0|web|/T|text|T|||1||/7|text|7|||1||/")
        eq("subtitle", items[1].Subtitle, "s")
        items := ScriptFilter.Parse("One`tsub`tC:\f.txt`r`n`r`nTwo`t`t`tfolder`t res:imageres.dll,-3 `nmailto:x@example.test`n`tno title")
        eq("lines", show(items), "One|file|C:\f.txt||" "|1||/Two|folder|Two|res:imageres.dll,-3||1||/mailto:x@example.test|url|mailto:x@example.test|||1||/")
        eq("type guessed from the form only", ScriptFilter.KindOf("", "\\server\share\a") ScriptFilter.KindOf("", "C:\") ScriptFilter.KindOf("", "obsidian://open")
            . ScriptFilter.KindOf("FILE", "x") ScriptFilter.KindOf("", "hello") ScriptFilter.KindOf("", "\\"), "filefolderurlfiletexttext")
        many := ""
        Loop 250
            many .= "item " A_Index "`n"
        eq("at most MaxItems", ScriptFilter.Parse(many).Length, ScriptFilter.MaxItems)
        problem := ""
        try ScriptFilter.Parse('{"items": [')
        catch as e
            problem := e.Message
        TestRunner.True("ScriptFilterParse.bad JSON throws: " problem, problem != "")
        problem := ""
        try ScriptFilter.Parse('{"x": 1}')
        catch as e
            problem := e.Message
        eq("no items list", problem, I18n.T("ScriptFilter.NoItems"))

        readRun := (out, err := "", code := 0, timedOut := false, failed := "") => ScriptFilter.Read({Stdout: out, Stderr: err, ExitCode: code, TimedOut: timedOut, Error: failed})
        entry := readRun("A`nB", "a warning", 1)
        eq("items win over the exit code", entry.Items.Length "|" entry.Error, "2|")
        entry := readRun("", "`n boom `nmore", 2)
        eq("stderr's first line", entry.Error "|" entry.Detail, "boom|boom `nmore")
        eq("no items, exit 0", readRun("").Error "|" readRun("").Items.Length, "|0")
        eq("exit code", readRun("", "", 3).Error, I18n.T("ScriptFilter.ExitCode", 3))
        TestRunner.True("ScriptFilterParse.bad JSON row", InStr(readRun('{"items": [').Error, I18n.T("ScriptFilter.BadJson", "")) = 1)
        eq("timed out", readRun("", "", "", true).Error, I18n.T("ScriptFilter.TimedOut", 10))
        eq("can't start", readRun("", "", "", false, "(2) not found").Error, I18n.T("ScriptFilter.CantStart", "(2) not found"))

        script := ProviderTests._FilterScript("find.ahk", "fd")
        rows := ScriptFilter.Rows(script, readRun('{"items": [{"title": "F", "arg": "C:\\f.txt", "uid": "f"}, {"title": "U", "arg": "https://e.test", "autocomplete": "u"}, {"title": "T"}]}'))
        list := ""
        for row in rows
            list .= row.Title "|" row.Kind "|" row.Icon "|" row.Uid "|" row.Exclusive "|" row.Score "|" row.AutoComplete "|" (row.Source == script) "/"
        eq("rows", list, "F|file|C:\f.txt|script:find.ahk:f|1|999||1/U|url|url:||1|998|fd u|1/T|text|" ScriptProvider.IconOf(script) "||1|997||1/")
        rows := ScriptFilter.Rows(script, readRun("", "bad thing", 1))
        eq("error row", rows.Length "|" rows[1].Title "|" rows[1].Subtitle "|" rows[1].Valid "|" rows[1].LargeText "|" rows[1].Exclusive, "1|" I18n.T("ScriptFilter.Error", "Find") "|bad thing|1|bad thing|1")
        rows := ScriptFilter.Rows(script, readRun(""))
        eq("no results row", rows.Length "|" rows[1].Title "|" rows[1].Valid, "1|" I18n.T("ScriptFilter.NoResults") "|0")

        eq("ahk command", ScriptFilter.CommandLine(script, 'a "b"'), '"' A_AhkPath '" /ErrorStdOut "C:\s\find.ahk" "a \"b\""')
        eq("py command", ScriptFilter.CommandLine(ProviderTests._FilterScript("f.py", "fp"), ""), 'py "C:\s\f.py" ""')
        ps := ProviderTests._FilterScript("it's.ps1", "fs")
        eq("ps1: the text only from the variable", ScriptFilter.CommandLine(ps, "x; calc")
            , '"' A_WinDir "\System32\WindowsPowerShell\v1.0\powershell.exe" '" -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "'
            . "[Console]::OutputEncoding=[Text.Encoding]::UTF8; & 'C:\s\it''s.ps1' $env:ALTRUN_QUERY; exit $LASTEXITCODE" '"')
        eq("problems", ScriptFilter.Problem(script) "|" ScriptFilter.Problem(ProviderTests._FilterScript("b.bat", "fb")) "|" ScriptFilter.Problem(ProviderTests._FilterScript("n.ahk", ""))
            . "|" ScriptFilter.Problem({Mode: "window", Keyword: "", Ext: "bat"}), "|NotForType|NeedsKeyword|")
    }

    ; When a filter script runs (fake runs): at once on entering the keyword, then once typing paused; a new run ends the one before;
    ; the last results stay while the next one runs; "Running..." only after WaitRowMs; results kept per text; late results of an
    ; ended run dropped; leaving the keyword / hiding ends it
    static ScriptFilterRules() {
        eq := (n, a, e) => TestRunner.Equal("ScriptFilterRules." n, a, e)
        saved := [ScriptFilter.StartFn, ScriptFilter.KillFn, ScriptFilter.OnResults]
        starts := [], kills := [], refreshed := {Count: 0}
        ScriptFilter.StartFn := (line, done, options) => (starts.Push({Line: line, Done: done, Options: options}), {Id: starts.Length})
        ScriptFilter.KillFn := (child) => kills.Push(child.Id)
        ScriptFilter.OnResults := () => refreshed.Count++
        ScriptFilter.Stop()
        script := ProviderTests._FilterScript("find.ahk", "fd")
        result := (out, code := 0) => {Stdout: out, Stderr: "", ExitCode: code, TimedOut: false, Error: "", Ms: 5}
        titles(rows) {
            list := ""
            for row in rows
                list .= row.Title "|"
            return RTrim(list, "|")
        }
        try {
            rows := ScriptFilter.Results(script, "")
            eq("first: the script's own row", rows[1].Title "|" rows[1].Valid "|" rows[1].Exclusive "|" rows[1].Subtitle, "Find|0|1|" I18n.T("Script.FilterHint"))
            Sleep(40)
            eq("first: starts at once", starts.Length, 1)
            eq("first: text, variables, folder, time limit", starts[1].Options.Env["ALTRUN_QUERY"] "|" starts[1].Options.Env["ALTRUN_KEYWORD"] "|"
                . starts[1].Options.Dir "|" starts[1].Options.TimeoutMs "|" starts[1].Line, "|fd|C:\s|10000|" ScriptFilter.CommandLine(script, ""))
            ScriptFilter.Results(script, "a"), Sleep(40), ScriptFilter.Results(script, "ab"), Sleep(40), ScriptFilter.Results(script, "abc"), Sleep(40)
            eq("typing: waits for a pause", starts.Length, 1)
            Sleep(250)
            eq("typing: one run for the last text", starts.Length "|" starts[2].Options.Env["ALTRUN_QUERY"], "2|abc")
            eq("the run before was ended", kills.Length ? kills[1] : 0, 1)
            eq("nothing yet after WaitRowMs: Running...", titles(ScriptFilter.Results(script, "abc")), I18n.T("ScriptFilter.Running", "Find"))
            TestRunner.True("ScriptFilterRules.the window searched again for the Running row", refreshed.Count >= 1)
            before := refreshed.Count
            starts[2].Done.Call(result("Hit`tsub"))
            eq("results arrived: searched again", refreshed.Count, before + 1)
            eq("results shown", titles(ScriptFilter.Results(script, "abc")), "Hit")
            ScriptFilter.Results(script, "abcd")
            eq("while the next runs: the last results stay", titles(ScriptFilter.Results(script, "abcd")), "Hit")
            Sleep(250)
            eq("next run", starts.Length, 3)
            eq("back to a text with results: no run, the running one ended", titles(ScriptFilter.Results(script, "abc")) "|" starts.Length "|" kills[kills.Length], "Hit|3|3")
            before := refreshed.Count
            starts[3].Done.Call(result("Late"))
            eq("late results of an ended run dropped", refreshed.Count, before)
            ScriptFilter.Results(script, "x")
            ScriptFilter.Leave()
            Sleep(250)
            eq("left the keyword: no run", starts.Length, 3)
            ScriptFilter.Results(script, "abc")
            eq("kept until hidden", starts.Length, 3)
            ScriptFilter.Stop()
            eq("hidden: forgotten", titles(ScriptFilter.Results(script, "abc")), "Find")
            Sleep(40)
            eq("hidden: runs again at once", starts.Length, 4)
            starts[4].Done.Call(result("", 1))
            eq("error", titles(ScriptFilter.Results(script, "abc")), I18n.T("ScriptFilter.Error", "Find"))
        } finally {
            ScriptFilter.Stop()
            ScriptFilter.StartFn := saved[1], ScriptFilter.KillFn := saved[2], ScriptFilter.OnResults := saved[3]
        }
    }

    ; "@altrun.filter altrun" (fake runs): one run with an empty text on entering the keyword, none while typing; ALTRun filters the
    ; items on "match" or the title, best first; nothing matching: "No results"; hidden: forgotten
    static ScriptFilterLocal() {
        eq := (n, a, e) => TestRunner.Equal("ScriptFilterLocal." n, a, e)
        saved := [ScriptFilter.StartFn, ScriptFilter.KillFn, ScriptFilter.OnResults]
        starts := []
        ScriptFilter.StartFn := (line, done, options) => (starts.Push({Done: done, Options: options}), {Id: starts.Length})
        ScriptFilter.KillFn := (child) => 0
        ScriptFilter.OnResults := () => 0
        ScriptFilter.Stop()
        script := ProviderTests._FilterScript("hosts.ps1", "host"), script.FilterBy := "altrun"
        titles(rows) {
            list := ""
            for row in ProviderRegistry.SortByScore(rows)
                list .= row.Title "|"
            return RTrim(list, "|")
        }
        try {
            ScriptFilter.Results(script, "al")
            Sleep(40)
            eq("runs at once, with an empty text", starts.Length "|" starts[1].Options.Env["ALTRUN_QUERY"], "1|")
            ScriptFilter.Results(script, "alt"), Sleep(250)
            eq("no run while typing", starts.Length, 1)
            starts[1].Done.Call({Stdout: '{"items": [{"title": "Coral Bay"}, {"title": "Altair Tower"}, {"title": "Zed", "match": "alder"}, {"title": "Harbor"}]}'
                , Stderr: "", ExitCode: 0, TimedOut: false, Error: "", Ms: 5})
            eq("filtered, best first (a shorter match ranks higher)", titles(ScriptFilter.Results(script, "al")), "Zed|Altair Tower|Coral Bay")
            eq("match instead of the title", titles(ScriptFilter.Results(script, "alder")), "Zed")
            eq("empty: all, in the script's order", titles(ScriptFilter.Results(script, "")), "Coral Bay|Altair Tower|Zed|Harbor")
            eq("nothing matching", titles(ScriptFilter.Results(script, "qqq")), I18n.T("ScriptFilter.NoResults"))
            Sleep(250)
            eq("still one run", starts.Length, 1)
            ScriptFilter.Stop()
            ScriptFilter.Results(script, "al"), Sleep(40)
            eq("hidden: runs again", starts.Length, 2)
        } finally {
            ScriptFilter.Stop()
            ScriptFilter.StartFn := saved[1], ScriptFilter.KillFn := saved[2], ScriptFilter.OnResults := saved[3]
        }
    }

    ; A real filter script found by the search: "keyword text" runs it (an .ahk; .ps1 and .py on Windows when there, with their
    ; times in the log), "keyword" alone completes "keyword ", a batch file and a script without keyword don't run, not in the
    ; selection panel, a hotkey opens the search window with "keyword " and the selected line
    static ScriptFilterSearch() {
        eq := (n, a, e) => TestRunner.Equal("ScriptFilterSearch." n, a, e)
        root := A_Temp "\ALTRun-script-filter-test"
        try DirDelete(root, true)
        DirCreate(root)
        root := TestRunner.LongPath(root)
        FileAppend("; @altrun.keyword fd`n; @altrun.mode filter`n#NoTrayIcon`n"
            . 'FileAppend("got " A_Args[1] "``t" EnvGet("ALTRUN_KEYWORD") "``t" A_Args[1] " 中文``ttext``n", "*", "UTF-8-RAW")`n', root "\find.ahk", "UTF-8")
        FileAppend("# @altrun.keyword fl`n# @altrun.mode filter`n# @altrun.filter  ALTRun`n"
            . "Write-Output ('Alpha host [' + $env:ALTRUN_QUERY + ']')`nWrite-Output 'Beta host'`n", root "\list.ps1", "UTF-8")
        FileAppend("# @altrun.keyword fs`n# @altrun.mode filter`n"
            . "Write-Output ('{`"items`": [{`"title`": `"ps ' + $env:ALTRUN_QUERY + '`"}]}')`n", root "\ps.ps1", "UTF-8")
        FileAppend("# @altrun.keyword fp`n# @altrun.mode filter`nimport sys`nprint('py ' + sys.argv[1])`n", root "\py.py", "UTF-8")
        FileAppend("rem @altrun.keyword bt`nrem @altrun.mode filter`necho x`n", root "\bad.bat", "UTF-8")
        FileAppend("; @altrun.mode filter`n; @altrun.selection Text`n", root "\nokw.ahk", "UTF-8")
        saved := [ScriptProvider.Folder, ScriptProvider.Scripts, ScriptProvider._checked, ScriptFilter.OnResults]
        savedShow := SearchWindow.GetOwnPropDesc("Show"), savedNotify := App.GetOwnPropDesc("Notify")
        arrived := {Count: 0}, shown := [], notices := []
        ScriptFilter.OnResults := () => arrived.Count++
        try {
            ScriptProvider.Folder := root
            ScriptProvider.Load()
            byName := Map()
            for script in ScriptProvider.Scripts
                byName[script.Name] := script
            eq("mode read", byName["find.ahk"].Mode "|" byName["bad.bat"].Mode "|" byName["find.ahk"].FilterBy "|" byName["list.ps1"].FilterBy, "filter|filter|script|altrun")
            hint := ScriptProvider.Search(SearchQuery("fd"))
            found := ""
            for item in hint
                if (item.Source == byName["find.ahk"])
                    found := item.Valid "|" item.AutoComplete "|" item.Subtitle
            eq("keyword alone: completes", found, "0|fd |" I18n.T("Search.TypeArgAfter", "fd", I18n.T("Script.FilterHint")) " · " I18n.T("Script.Subtitle", "find.ahk"))

            timed(keyword, text, expected) {
                started := A_TickCount, arrived.Count := 0
                ScriptProvider.Search(SearchQuery(keyword " " text))
                TestUtil.WaitFor(() => arrived.Count && !ScriptFilter.Running && !IsObject(ScriptFilter._wanted), 20000)
                rows := ScriptProvider.Search(SearchQuery(keyword " " text))
                eq(keyword ": results", rows.Length ? rows[1].Title "|" rows[1].Exclusive : "none", expected "|1")
                FileAppend("NOTE ScriptFilter: " keyword " " (A_TickCount - started) " ms from typing to results`n", "*")
                return rows
            }
            rows := timed("fd", 'a & "b"', 'got a & "b"')
            eq("fd: the rest of the line", rows.Length ? rows[1].Subtitle "|" rows[1].Arg "|" rows[1].Kind : "", 'fd|a & "b" 中文|text')
            onWine := DllCall("GetProcAddress", "Ptr", DllCall("GetModuleHandle", "Str", "ntdll", "Ptr"), "AStr", "wine_get_version", "Ptr")
            if !onWine {
                timed("fs", "x; y", "ps x; y")
                timed("fl", "beta", "Beta host")                            ; @altrun.filter altrun: one run, ALTRun filters
                rows := ScriptProvider.Search(SearchQuery("fl "))
                eq("fl: run once with an empty text", rows.Length ? rows[1].Title : "", "Alpha host []")
                if FileExist(A_WinDir "\py.exe")
                    timed("fp", "q", "py q")
                else
                    FileAppend("NOTE ScriptFilter: no py.exe, the .py script wasn't run`n", "*")
            }
            plain := ScriptProvider.Search(SearchQuery("bt x"))
            exclusive := false
            for item in plain
                exclusive := exclusive || item.Exclusive
            eq("batch file: not run", exclusive "|" ScriptFilter.Running, "0|0")

            SearchWindow.DefineProp("Show", {Call: (this, text := "", *) => shown.Push(text)})
            App.DefineProp("Notify", {Call: (this, text, *) => notices.Push(text)})
            ScriptProvider.RunFromHotkey("find.ahk", () => {Text: "hello there"})
            ScriptProvider.RunFromHotkey("find.ahk", () => {Text: "two`nlines"})
            ScriptProvider.RunFromHotkey("find.ahk", () => "")
            eq("hotkey: search window with the keyword and the selected line", shown.Length "|" shown[1] "|" shown[2] "|" shown[3], "3|fd hello there|fd |fd ")
            ScriptProvider.RunFromHotkey("nokw.ahk", () => {Text: "hello"})
            eq("hotkey: no keyword", notices.Length ? notices[1] : "", I18n.T("Script.Filter.NeedsKeyword", "nokw"))
            panel := ScriptProvider.ActionsFor({Kind: "text", Types: ["Text"], Term: "hello", Quoted: '"hello"', Text: "hello"})
            eq("not in the selection panel", panel.Length, 0)
            ScriptProvider.Hidden()
            eq("hidden: nothing running", ScriptFilter.Running "|" ChildProcess.Pending, "0|0")
        } finally {
            ScriptFilter.Stop()
            SearchWindow.DefineProp("Show", savedShow), App.DefineProp("Notify", savedNotify)
            ScriptProvider.Folder := saved[1], ScriptProvider.Scripts := saved[2], ScriptProvider._checked := saved[3], ScriptFilter.OnResults := saved[4]
            try DirDelete(root, true)
        }
    }

    ; Scripts for a selection ("@altrun.selection") and hotkeys that run a script ("script:<file name>"): the panel rows with the
    ; command-line checks of the script's real command line, several files as one argument each, the hotkey rules, the names in Preferences
    static ScriptSelection() {
        eq := (n, a, e) => TestRunner.Equal("ScriptSelection." n, a, e)
        root := A_Temp "\ALTRun-script-selection-test"
        try DirDelete(root, true)
        DirCreate(root)
        root := TestRunner.LongPath(root)
        FileAppend("; @altrun.title PDF Tool`n; @altrun.selection PDF, Files`n", root "\pdf.ahk", "UTF-8")
        FileAppend("@echo off`nREM @altrun.title Ping IP`nREM @altrun.selection IP`n", root "\ip.bat", "UTF-8")
        FileAppend("; @altrun.title Quiet`n; @altrun.mode silent`n; @altrun.selection Text`n", root "\quiet.ahk", "UTF-8")
        FileAppend("# @altrun.title Lookup`n# @altrun.keyword lk`n# @altrun.argument a word`n", root "\lookup.ps1", "UTF-8")
        FileAppend("# @altrun.title No Keyword`n# @altrun.argument some text`n", root "\noword.py", "UTF-8")
        FileAppend("; @altrun.title Plain`n", root "\plain.ahk", "UTF-8")
        saved := [ScriptProvider.Folder, ScriptProvider.Scripts, ScriptProvider._checked]
        savedRun := ScriptProvider.GetOwnPropDesc("Run"), savedNotify := App.GetOwnPropDesc("Notify"), savedShow := SearchWindow.GetOwnPropDesc("Show")
        savedUsage := Usage.Days.Clone(), scripts := AppSettings.Feature("Scripts"), savedEnabled := scripts["Enabled"]
        ran := [], notices := [], shown := [], captured := 0
        try {
            ScriptProvider.Folder := root
            ScriptProvider.Load()
            byName := Map()
            for script in ScriptProvider.Scripts
                byName[script.Name] := script
            names := ""
            for script in ScriptProvider.Scripts
                names .= script.Name "|"
            eq("by file name", names, "ip.bat|lookup.ps1|noword.py|pdf.ahk|plain.ahk|quiet.ahk|")
            eq("types read", PrefsPages.JoinNames(byName["pdf.ahk"].Selection) "/" byName["plain.ahk"].Selection.Length, "PDF, Files/0")
            eq("types any case", ScriptProvider.ForTypes(byName["pdf.ahk"], ["File", "pdf"]) "|" ScriptProvider.ForTypes(byName["pdf.ahk"], ["Text"]), "1|0")
            eq("shell of the real command line", SelectionActions.ShellIn(ScriptProvider.ShellCommand(byName["pdf.ahk"])["Target"]) "|"
                SelectionActions.ShellIn(ScriptProvider.ShellCommand(byName["quiet.ahk"])["Target"]) "|" SelectionActions.ShellIn(ScriptProvider.ShellCommand(byName["ip.bat"])["Target"]) "|"
                SelectionActions.ShellIn(ScriptProvider.ShellCommand(byName["lookup.ps1"])["Target"]), "|cmd|cmd|posix")

            scripts["Enabled"] := 1
            titles(info) {
                list := ""
                for row in ScriptProvider.ActionsFor(info)
                    list .= row.Title "|"
                return RTrim(list, "|")
            }
            text(term, types) => {Kind: "text", Types: types, Term: term, Quoted: SelectionActions.Quote(term), Text: term}
            eq("panel: a file of the type", titles({Kind: "file", Types: ["File", "PDF"], Term: "C:\a.pdf", Quoted: '"C:\a.pdf"', Path: "C:\a.pdf"}), "PDF Tool")
            eq("panel: safe text for cmd", titles(text("10.0.0.1", ["Text", "IP"])), "Ping IP|Quiet")
            eq("panel: cmd characters refused (batch file, and an ahk whose output goes through cmd)", titles(text("10.0.0.1 & del x", ["Text", "IP"])), "")
            eq("panel: no type, no script", titles(text("hello", ["Url"])), "")
            files := ["C:\a b\1.pdf", "C:\2.pdf"]
            argument := ScriptProvider.SelectionArgument({Kind: "files", Files: files, Term: '"C:\a b\1.pdf" "C:\2.pdf"'})
            eq("several files: one argument each", (argument is Array) "|" ScriptProvider.CommandLine(byName["pdf.ahk"], argument)
                , "1|" SubStr(ScriptProvider.CommandLine(byName["pdf.ahk"]), 1) ' "C:\a b\1.pdf" "C:\2.pdf"')
            scripts["Enabled"] := 0
            eq("panel: scripts turned off", titles(text("10.0.0.1", ["Text", "IP"])), "")
            scripts["Enabled"] := 1
            eq("titles for Preferences", PrefsPages.JoinNames(SelectionActions.CommandTitles(["pdf"], [Map("Title", "Open", "Selection", "PDF")], ScriptProvider.Scripts))
                , "Open, " I18n.T("Selection.ScriptTitle", "PDF Tool"))

            ScriptProvider.DefineProp("Run", {Call: (this, script, argument := "") => ran.Push({Name: script.Name, Argument: argument})})
            App.DefineProp("Notify", {Call: (this, message, *) => notices.Push(message)})
            SearchWindow.DefineProp("Show", {Call: (this, value := "", *) => shown.Push(value)})
            capture(selection) => () => (captured++, selection)
            runHotkey := (name, selection := "") => ScriptProvider.RunFromHotkey(name, capture(selection))
            runHotkey("plain.ahk", {Text: "ignored"})
            eq("plain: runs, the selection isn't copied", ran.Length "|" ran[1].Argument "|" captured, "1||0")
            runHotkey("PDF.AHK", {Files: ["C:\Docs\report.docx"]})
            eq("any case; the type isn't checked", ran.Length "|" ran[2].Name "|" ran[2].Argument "|" captured, "2|pdf.ahk|C:\Docs\report.docx|1")
            runHotkey("pdf.ahk")
            eq("content types, nothing selected: runs without one", ran.Length "|" ran[3].Argument, "3|")
            runHotkey("lookup.ps1")
            eq("argument and keyword, nothing selected: the search window", shown.Length "|" shown[1] "|" ran.Length, "1|lk |3")
            runHotkey("noword.py")
            eq("argument without keyword, nothing selected: says so", notices.Length "|" notices[1] "|" ran.Length, "1|" I18n.T("Selection.None") "|3")
            runHotkey("lookup.ps1", {Text: "red apple"})
            eq("argument: the selected text", ran.Length "|" ran[4].Argument, "4|red apple")
            runHotkey("ip.bat", {Text: "8.8.8.8 & del x"})
            eq("cmd characters: refused", notices.Length "|" notices[2] "|" ran.Length, "2|" I18n.T("Hotkey.SelectionRefused", "Ping IP") "|4")
            runHotkey("pdf.ahk", {Files: files})
            eq("several files", ran.Length "|" (ran[5].Argument is Array) "|" ran[5].Argument.Length, "5|1|2")
            runHotkey("gone.ahk", {Text: "x"})
            eq("missing script: says so", notices.Length "|" notices[3] "|" ran.Length, "3|" I18n.T("Hotkey.ScriptMissing", "gone.ahk") "|5")

            eq("name", HotkeyEditor.ActionName("script:pdf.ahk") "|" HotkeyEditor.ActionName("script:gone.ahk"), "PDF Tool|" I18n.T("HotkeyEditor.ScriptMissing", "script:gone.ahk"))
            eq("list text", HotkeyEditor.ListText(Map("Key", "^!p", "Action", "script:pdf.ahk")), I18n.T("Prefs.Hotkey.Run", "PDF Tool"))
            eq("missing", HotkeyEditor.IsMissing(Map("Action", "script:gone.ahk")) "|" HotkeyEditor.IsMissing(Map("Action", "script:pdf.ahk")), "1|0")
            detail := HotkeyEditor.ActionDetail("script:pdf.ahk")
            eq("detail: takes the selection", detail.Text "|" detail.Missing, I18n.T("CommandPicker.Kind.Script") " · pdf.ahk · " I18n.T("HotkeyEditor.Detail.ScriptArgument") "|0")
            eq("detail: plain", HotkeyEditor.ActionDetail("script:plain.ahk").Text, I18n.T("CommandPicker.Kind.Script") " · plain.ahk · " I18n.T("HotkeyEditor.Detail.Plain"))
            detail := HotkeyEditor.ActionDetail("script:gone.ahk")
            eq("detail: missing", detail.Text "|" detail.Missing, I18n.T("HotkeyEditor.Detail.ScriptMissing") "|1")
            eq("summary", InStr(HotkeyEditor.Summary(Map("Key", "^!p", "Action", "script:pdf.ahk", "WinTitle", "")), '"PDF Tool"') > 0, true)

            choices := CommandPicker.Choices([], ScriptProvider.Scripts)
            kinds := Map()
            for choice in choices
                kinds[choice.Kind] := (kinds.Has(choice.Kind) ? kinds[choice.Kind] : 0) + 1
            eq("picker: the scripts as their own kind, before the system commands", kinds[I18n.T("CommandPicker.Kind.Script")] "|" choices[3].Id "|" choices[3].Kind "|" choices[9].Kind
                , "6|script:ip.bat|" I18n.T("CommandPicker.Kind.Script") "|" I18n.T("CommandPicker.Kind.System"))
            found := CommandPicker.Filter(choices, "noword")
            eq("picker: found by file name", found.Length ? found[1].Id : "", "script:noword.py")
            found := CommandPicker.Filter(choices, "lk")
            eq("picker: found by keyword", found.Length ? found[1].Id : "", "script:lookup.ps1")
        } finally {
            ScriptProvider.DefineProp("Run", savedRun), App.DefineProp("Notify", savedNotify), SearchWindow.DefineProp("Show", savedShow)
            Usage.Days := savedUsage, scripts["Enabled"] := savedEnabled
            ScriptProvider.Folder := saved[1], ScriptProvider.Scripts := saved[2], ScriptProvider._checked := saved[3]
            try DirDelete(root, true)
        }
    }

    ; Snippets: the text is searched too
    static SnippetTextSearch() {
        eq := (n, a, e) => TestRunner.Equal("SnippetText." n, a, e)
        ok := (n, c) => TestRunner.True("SnippetText." n, c)
        options := AppSettings.Feature("Snippets")
        savedSnippets := AppSettings.Snippets, savedSearch := options["SearchText"], savedExpand := [options["AutoExpand"], options["ExpandPrefix"]]
        options["AutoExpand"] := 0                                          ; the subtitles below without the abbreviations (SnippetTriggers)
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
        ok("large text worked out only when shown", IsObject(item.LargeText) && item.DisplayText() = "Best regards")
        AppSettings.Data["Snippets"].Push(Map("Name", "Today", "Keyword", "td", "Text", "On {date:yyyy}"))
        eq("large text expands placeholders", SnippetProvider.Search(SearchQuery("snip td"))[1].DisplayText(), "On " A_YYYY)
        ok("snip keyword boosts", InStr(titles("snip goodwill"), "Quote reply=" (SnippetProvider.TextScore + 30) "|"))
        options["SearchText"] := 0
        eq("switched off", titles("goodwill"), "")
        options["SearchText"] := savedSearch, options["AutoExpand"] := savedExpand[1], options["ExpandPrefix"] := savedExpand[2]
        AppSettings.Data["Snippets"] := savedSnippets
        eq("match position", SnippetProvider.TextMatch("abc def", ["DEF", "abc"]), 5)
        eq("no match", SnippetProvider.TextMatch("abc def", ["abc", "xyz"]), 0)
    }

    ; A snippet's subtitle starts with what expands it while typing (AHK forum: "a view of all the hotstrings"): ";sig, ;signature · ..."
    static SnippetTriggers() {
        eq := (n, a, e) => TestRunner.Equal("SnippetTriggers." n, a, e)
        options := AppSettings.Feature("Snippets")
        saved := [options["Enabled"], options["AutoExpand"], options["ExpandPrefix"]]
        paste := I18n.T("Snippet.Subtitle", "Best regards")
        try {
            options["Enabled"] := 1, options["AutoExpand"] := 1, options["ExpandPrefix"] := ";"
            eq("one keyword", SnippetProvider.Subtitle(Map("Keyword", "sig", "Text", "x"), "Best regards", options), ";sig · " paste)
            eq("two keywords", SnippetProvider.Subtitle(Map("Keyword", "sig, signature", "Text", "x"), "Best regards", options), ";sig, `;signature · " paste)
            eq("its own expansion off", SnippetProvider.Subtitle(Map("Keyword", "sig", "Text", "x", "AutoExpand", 0), "Best regards", options), paste)
            eq("no keyword", SnippetProvider.Subtitle(Map("Keyword", "", "Text", "x"), "Best regards", options), paste)
            options["ExpandPrefix"] := "//"
            eq("other prefix", SnippetProvider.Subtitle(Map("Keyword", "sig", "Text", "x"), "Best regards", options), "//sig · " paste)
            options["AutoExpand"] := 0
            eq("auto expansion off", SnippetProvider.Subtitle(Map("Keyword", "sig", "Text", "x"), "Best regards", options), paste)
        } finally {
            options["Enabled"] := saved[1], options["AutoExpand"] := saved[2], options["ExpandPrefix"] := saved[3]
        }
    }

    ; Browser bookmarks
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
        BookmarkProvider.Items := items, BookmarkProvider._checked := A_TickCount   ; don't read this PC's bookmark files
        results := BookmarkProvider.Search(SearchQuery("tender"))
        eq("default results", results.Length ? results[1].Arg : "", "https://www.tenders.example.com/login")
        ok("default results rank below commands", results.Length && results[1].Score < 100 && !results[1].Exclusive)
        results := BookmarkProvider.Search(SearchQuery("bm tenders"))
        ok("keyword + host", results.Length = 1 && results[1].Exclusive)
        eq("too short", BookmarkProvider.Search(SearchQuery("g")).Length, 0)
        hint := BookmarkProvider.Search(SearchQuery("bm"))                  ; the keyword alone: its hint, Enter completes "bm "
        eq("keyword alone: hint", hint.Length "|" hint[1].Valid "|" hint[1].AutoComplete "|" hint[1].Score "|" hint[1].Exclusive, "1|0|bm |" ResultItem.KeywordHintScore "|0")
        ; when a bookmark file can't be read (the browser is writing it) its time isn't noted, so it is read again at the next search
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
        ok("good file stamped", InStr(BookmarkProvider._stamp, "\Default\Bookmarks"))   ; A_Temp may be an 8.3 short path, so whole paths aren't compared
        eq("good file items", BookmarkProvider.Items.Length, 3)
        EnvSet("LOCALAPPDATA", savedLocal)
        try DirDelete(root, true)
        BookmarkProvider.Items := savedItems, BookmarkProvider._checked := savedChecked, BookmarkProvider._stamp := savedStamp
        ids := ""
        for command in SystemProvider.Commands()
            ids .= command["Id"] "|"
        ok("media commands", InStr(ids, "MediaPlayPause|MediaNext|MediaPrev|MediaStop|"))
    }

    ; Other browsers (#492): Features.Bookmarks.BrowserFolders adds the profiles of portable browsers to bookmarks and history,
    ; a User Data folder (profiles inside) or one profile folder, relative to the program folder or with variables; each file once
    static BrowserFolders() {
        eq := (n, a, e) => TestRunner.Equal("BrowserFolders." n, a, e)
        ok := (n, c) => TestRunner.True("BrowserFolders." n, c)
        root := A_Temp "\ALTRun-test-browser-folders"
        try DirDelete(root, true)
        portable := root "\Portable\Data\profile"                         ; a User Data folder: Default and Profile 1 inside
        DirCreate(portable "\Default"), DirCreate(portable "\Profile 1"), DirCreate(portable "\Crashpad")
        FileAppend("{}", portable "\Default\Bookmarks"), FileAppend("{}", portable "\Default\Preferences")
        FileAppend("{}", portable "\Profile 1\Preferences"), FileAppend("", portable "\Profile 1\History")
        single := root "\Single"                                            ; one profile folder
        DirCreate(single), FileAppend("{}", single "\Bookmarks"), FileAppend("{}", single "\Preferences"), FileAppend("", single "\History")
        DirCreate(root "\Local")                                            ; %LOCALAPPDATA% without any browser
        options := AppSettings.Feature("Bookmarks"), saved := options["BrowserFolders"], savedLocal := EnvGet("LOCALAPPDATA")
        try {
            EnvSet("LOCALAPPDATA", root "\Local"), EnvSet("ALTRUN_TEST_BROWSERS", root)
            options["BrowserFolders"] := []
            eq("none", BookmarkProvider.Files().Length, 0)
            options["BrowserFolders"] := ["%ALTRUN_TEST_BROWSERS%\Portable\Data\profile", single, "", root "\Missing", single "\"]
            short := (path) => RegExReplace(path, "i)^.*\\ALTRun-test-browser-folders")   ; A_Temp may be an 8.3 short path, expanded or not
            names := ""
            for each in BookmarkProvider.Files()
                names .= each[1] ":" short(each[2]) "|"
            eq("bookmarks: user data + profile folder, once each", names, "Other:\Portable\Data\profile\Default\Bookmarks|Other:\Single\Bookmarks|")
            names := ""
            for historyFile in BrowserHistoryProvider.Files()
                names .= short(historyFile) "|"
            eq("history: the same folders", names, "\Portable\Data\profile\Profile 1\History|\Single\History|")
            eq("profiles in user data", BookmarkProvider.ProfileCount(portable), 2)
            eq("profile folder", BookmarkProvider.ProfileCount(single), 1)
            eq("missing folder", BookmarkProvider.ProfileCount(root "\Missing"), 0)
            summary := PrefsPages.BrowserFolderSummary([portable, root "\Missing"])
            ok("summary: profiles found / not found", InStr(summary, I18n.T("Prefs.BrowserFolders.Profiles", 2)) > 0 && InStr(summary, I18n.T("Prefs.BrowserFolders.NotFound")) > 0)
            eq("empty lines skipped", BookmarkProvider.ExtraFolders().Length, 4)
            options["BrowserFolders"] := ["..\x"]                           ; a portable browser next to ALTRun: relative to the program folder
            eq("relative path", BookmarkProvider.ExtraFolders()[1], Path.Full(A_ScriptDir "\..\x"))
            ; the standard Chrome folder listed again isn't read twice
            DirCreate(root "\Local\Google\Chrome\User Data\Default"), FileAppend("{}", root "\Local\Google\Chrome\User Data\Default\Bookmarks")
            options["BrowserFolders"] := ["%LOCALAPPDATA%\Google\Chrome\User Data"]
            files := BookmarkProvider.Files()
            eq("standard folder once", files.Length "|" files[1][1], "1|Chrome")
        } finally {
            options["BrowserFolders"] := saved
            EnvSet("LOCALAPPDATA", savedLocal), EnvSet("ALTRUN_TEST_BROWSERS")
            try DirDelete(root, true)
        }
    }

    ; Browser history: "h word" (keyword only), most recent first, every word in the title or URL; the History files read through winsqlite3.dll
    static BrowserHistory() {
        eq := (n, a, e) => TestRunner.Equal("BrowserHistory." n, a, e)
        ok := (n, c) => TestRunner.True("BrowserHistory." n, c)
        savedItems := BrowserHistoryProvider.Items, savedLoaded := BrowserHistoryProvider._loaded, savedActive := BrowserHistoryProvider._active
        page := (title, url, time) => {Title: title, Url: url, Time: time, Text: StrLower(title " " url), Cjk: RegExMatch(title, "[\x{2E80}-\x{9FFF}]") > 0, Initials: ""}
        BrowserHistoryProvider.Items := [page("Tender Portal - Login", "https://tenders.example.com/login", 3)       ; most recent first
            , page("中文 标题 测试", "https://example.cn/zh", 2), page("GitHub", "https://github.com/", 1)]
        BrowserHistoryProvider._loaded := A_TickCount, BrowserHistoryProvider._active := true   ; don't read this PC's history files
        eq("not without the keyword", BrowserHistoryProvider.Search(SearchQuery("tender")).Length, 0)
        hint := BrowserHistoryProvider.Search(SearchQuery("h"))           ; the keyword alone: its hint, Enter completes "h "
        eq("keyword alone: hint", hint.Length "|" hint[1].Valid "|" hint[1].AutoComplete "|" hint[1].Exclusive, "1|0|h |0")
        eq("every word, title or URL", BrowserHistoryProvider.Matches("tender LOGIN", 10).Length, 1)
        eq("a word missing", BrowserHistoryProvider.Matches("tender github", 10).Length, 0)
        eq("Chinese title", BrowserHistoryProvider.Matches("标题", 10).Length, 1)
        eq("pinyin initials", BrowserHistoryProvider.Matches("zwbt", 10).Length "|" BrowserHistoryProvider.Matches("zwbt", 10)[1].Url, "1|https://example.cn/zh")
        eq("empty: most recent first", BrowserHistoryProvider.Matches("", 2).Length "|" BrowserHistoryProvider.Matches("", 2)[1].Title, "2|Tender Portal - Login")
        eq("most recent first", BrowserHistoryProvider.Matches("https", 10)[3].Title, "GitHub")
        if Sqlite.Available() {
            results := BrowserHistoryProvider.Search(SearchQuery("h com"))
            eq("keyword: urls in order", results.Length "|" results[1].Arg "|" results[2].Arg, "2|https://tenders.example.com/login|https://github.com/")
            ok("keyword: exclusive url results", results[1].Exclusive && results[1].Kind = "url" && results[1].Uid = "history:https://tenders.example.com/login" && results[1].Score = results[2].Score)
            none := BrowserHistoryProvider.Search(SearchQuery("h nothing-like-this"))
            eq("no match: says so", none.Length "|" none[1].Valid "|" none[1].Exclusive "|" none[1].Title, "1|0|1|" I18n.T("BrowserHistory.NoMatch"))
        } else {
            none := BrowserHistoryProvider.Search(SearchQuery("h com"))
            eq("no winsqlite3.dll: says so", none.Length "|" none[1].Valid "|" none[1].Title, "1|0|" I18n.T("BrowserHistory.Unavailable"))
        }
        BrowserHistoryProvider.Search(SearchQuery("tender"))                ; leaving history search: the next one reads again when old
        eq("leaving history search", BrowserHistoryProvider._active, false)
        SetTimer(BrowserHistoryProvider._releaseTimer, 0)                   ; let go ReloadMs later; called here right away
        kept := BrowserHistoryProvider.Items
        BrowserHistoryProvider._releaseTimer.Call()
        eq("pages let go after leaving", BrowserHistoryProvider.Items.Length "|" BrowserHistoryProvider._loaded, "0|0")
        BrowserHistoryProvider.Items := kept, BrowserHistoryProvider._loaded := A_TickCount, BrowserHistoryProvider._active := true
        BrowserHistoryProvider._releaseTimer.Call()
        eq("kept while searching history again", BrowserHistoryProvider.Items.Length > 0, true)
        BrowserHistoryProvider._active := true                               ; the search window hidden in history search ("keep the last search": no other search follows)
        savedProviders := ProviderRegistry.Providers
        ProviderRegistry.Providers := [ApplicationProvider, BrowserHistoryProvider]   ; one without Hidden() is skipped
        ProviderRegistry.WindowHidden()
        ProviderRegistry.Providers := savedProviders
        eq("hidden: leaves history search", BrowserHistoryProvider._active, false)
        SetTimer(BrowserHistoryProvider._releaseTimer, 0)
        kept := BrowserHistoryProvider.Items
        BrowserHistoryProvider._releaseTimer.Call()
        eq("hidden: pages let go", BrowserHistoryProvider.Items.Length "|" BrowserHistoryProvider._loaded, "0|0")
        BrowserHistoryProvider.Items := kept
        BrowserHistoryProvider._active := false
        ; Chromium's time: microseconds since 1601-01-01 UTC
        offset := DateDiff(A_Now, A_NowUTC, "Seconds")
        chrome := (localStamp) => (DateDiff(localStamp, "16010101", "Seconds") - offset) * 1000000
        eq("time today", BrowserHistoryProvider.TimeText(chrome("20261111140500"), "20261111200000"), "14:05")
        eq("time this year", BrowserHistoryProvider.TimeText(chrome("20260309081500"), "20261111200000"), "03-09 08:15")
        eq("time earlier", BrowserHistoryProvider.TimeText(chrome("20251231235900"), "20261111200000"), "2025-12-31")
        eq("no time", BrowserHistoryProvider.TimeText(0), "")
        eq("file uri", Sqlite.FileUri("C:\A b#%\中文\History"), "file:///C:/A%20b%23%25/中文/History")
        if !Sqlite.Available() {
            TestRunner.Equal("BrowserHistory.files (skipped: no winsqlite3.dll, e.g. Wine)", true, true)
        } else {
            ; real History files: two browsers / profiles, a page in both (its last visit counts), hidden and non-http pages left out
            root := A_Temp "\ALTRun-test-history"
            try DirDelete(root, true)
            schema := "CREATE TABLE urls(id INTEGER PRIMARY KEY, url LONGVARCHAR, title LONGVARCHAR, visit_count INTEGER DEFAULT 0 NOT NULL,"
                . " typed_count INTEGER DEFAULT 0 NOT NULL, last_visit_time INTEGER NOT NULL, hidden INTEGER DEFAULT 0 NOT NULL);"
            DirCreate(root "\Google\Chrome\User Data\Default")
            Sqlite.Exec(root "\Google\Chrome\User Data\Default\History", schema
                . "INSERT INTO urls(url, title, last_visit_time) VALUES ('https://a.example.com/', 'Page A', 13400000000000000),"
                . " ('https://b.example.com/', '中文 页面', 13400000000000300), ('javascript:alert(1)', 'Script', 13400000000000400),"
                . " ('chrome://settings/', 'Settings', 13400000000000500), ('https://shared.example.com/', 'Shared', 13400000000000100);"
                . "INSERT INTO urls(url, title, last_visit_time, hidden) VALUES ('https://hidden.example.com/', 'Hidden', 13400000000000600, 1);")
            DirCreate(root "\Microsoft\Edge\User Data\Profile 1")
            Sqlite.Exec(root "\Microsoft\Edge\User Data\Profile 1\History", schema
                . "INSERT INTO urls(url, title, last_visit_time) VALUES ('https://shared.example.com/', 'Shared', 13400000000000700),"
                . " ('http://c.example.com/', '', 13400000000000200);")
            savedLocal := EnvGet("LOCALAPPDATA")
            EnvSet("LOCALAPPDATA", root)
            BrowserHistoryProvider.Load()
            EnvSet("LOCALAPPDATA", savedLocal)
            list := ""
            for item in BrowserHistoryProvider.Items
                list .= item.Title "|"
            eq("files: merged, most recent first", list, "Shared|中文 页面|http://c.example.com/|Page A|")
            eq("files: Chinese title read", BrowserHistoryProvider.Matches("zwym", 5).Length, 1)
            rows := Sqlite.Query(root "\Google\Chrome\User Data\Default\History", "SELECT count(*) FROM urls", 0, "immutable")
            eq("immutable read", rows[1][1], 6)
            ok("nothing copied to TEMP", !DirExist(A_Temp "\ALTRun-History-" DllCall("GetCurrentProcessId")))
            threw := false
            try Sqlite.Query(root "\nothing\History", "SELECT 1", 0, "read")
            catch
                threw := true
            ok("missing file throws", threw)
            ; time: 10,000 pages (MaxItems of one profile), half with Chinese titles; the first pinyin search works out every initial
            try DirDelete(root, true)
            DirCreate(root "\Google\Chrome\User Data\Default")
            Sqlite.Exec(root "\Google\Chrome\User Data\Default\History", schema
                . "INSERT INTO urls(url, title, last_visit_time) WITH RECURSIVE n(i) AS (SELECT 1 UNION ALL SELECT i + 1 FROM n WHERE i < 10000)"
                . " SELECT 'https://site' || i || '.example.com/docs/page', CASE WHEN i % 2 THEN 'Project page ' || i ELSE '项目 文档 第' || i || '页' END,"
                . " 13400000000000000 + i FROM n;")
            EnvSet("LOCALAPPDATA", root)
            started := A_TickCount
            BrowserHistoryProvider.Load()
            loadMs := A_TickCount - started
            EnvSet("LOCALAPPDATA", savedLocal)
            started := A_TickCount
            words := BrowserHistoryProvider.Matches("site9 docs", 30).Length
            wordsMs := A_TickCount - started
            started := A_TickCount
            pinyin := BrowserHistoryProvider.Matches("xmwd", 30).Length
            pinyinMs := A_TickCount - started
            started := A_TickCount
            none := BrowserHistoryProvider.Matches("nothing here", 30).Length
            noneMs := A_TickCount - started
            FileAppend("NOTE BrowserHistory: 10000 pages read in " loadMs " ms; search 2 words " wordsMs " ms, pinyin (first) " pinyinMs " ms, no match " noneMs " ms`n", "*")
            eq("time: pages", BrowserHistoryProvider.Items.Length "|" words "|" pinyin "|" none, "10000|30|30|0")
            ok("time: read under 3 s", loadMs < 3000)
            try DirDelete(root, true)
        }
        BrowserHistoryProvider.Items := savedItems, BrowserHistoryProvider._loaded := savedLoaded, BrowserHistoryProvider._active := savedActive
    }
}
