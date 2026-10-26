;===============================================================================
; ProviderTests.ahk - Result providers: web search, clipboard, file index and file search,
;   applications, commands, windows, recent items, scripts, bookmarks and snippets search
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
        eq("image thumbnail", imageItem.Icon, "thumb:" root "\Clipboard\one.png")
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
            eq("scanned", FileIndex.Paths.Length, 1)
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
            eq("new file rescans", FileIndex.Scanning || FileIndex.Paths.Length = 3, true)
            waitScan()
            eq("new file indexed", FileIndex.Paths.Length, 3)               ; including .crdownload (scanning doesn't skip it, it just doesn't trigger)
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
        eq("file mode", titles(ProviderRegistry.SearchOnly(FileSearchProvider, "cad tower")), "Tower Plan.dwg")
        eq("file mode plain", titles(ProviderRegistry.SearchOnly(FileSearchProvider, "tower")), "Tower|Tower Plan.dwg|Tower Plan.pdf|Tower photo.jpg")
        hint := FileSearchProvider.Search(SearchQuery("cad"))
        eq("keyword alone hint", hint.Length, 1)
        eq("keyword alone low score", hint[1].Score < 10 && !hint[1].Exclusive, true)
        eq("hint lists extensions", InStr(hint[1].Title, "dwg dxf") ? 1 : 0, 1)
        eq("not a keyword", FileSearchProvider.Search(SearchQuery("cadence")).Length, 0)
        eq("everything prefix", FileSearchProvider._ScopePrefix(parsed), "ext:doc;docx;pdf ")
        eq("everything folder prefix", FileSearchProvider._ScopePrefix(true), "folder:")
        options["TypeFilters"].Push("dwgonly = dwg")                        ; a changed setting applies at once
        eq("settings change", titles(FileSearchProvider.Search(SearchQuery("dwgonly tower"))), "Tower Plan.dwg")
        options["TypeFilters"].Pop()
        options["UseEverything"] := savedEverything
        FileIndex.Paths := [], FileIndex.Names := [], FileIndex.Folders := []
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
            eq("keyword only: a hint", hint.Length "|" hint[1].Valid "|" hint[1].Score "|" hint[1].Subtitle, "1|0|5|recent ...")
            results := FileSearchProvider.Search(SearchQuery("recent note"))
            eq("through Search", results.Length ? results[1].Title : "", "notes.txt")
            eq("file search mode", FileSearchProvider.SearchOnly("recent tower")[1].Title, "Tower")
            eq("default keyword", AppSettings.Defaults()["Features"]["FileSearch"]["RecentKeywords"][1], "recent")
            FileSetTime("20260601000000", recent "\notes.txt.lnk")             ; opened again: moves to the top
            eq("reopened moves up", WindowsRecent.Items(300)[1].Name, "notes.txt")
            eq("quick switch folders", QuickSwitch.RecentFolders(10, recent).Length, 2)
        } finally {
            WindowsRecent.Dir := saved[1], WindowsRecent.CacheMs := saved[2]
            try DirDelete(root, true)
        }
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
        eq("remote unc", Path.IsRemote("\\server\share\PT1001"), true)
        eq("remote local", Path.IsRemote("C:\Windows"), false)
        eq("remote folder key", IconCache._CacheKey("\\server\share\PT1001 - 24 Harbor"), "folder:")
        eq("remote folder slash", IconCache._CacheKey("\\server\share\PT1001\"), "folder:")
        eq("remote file key", IconCache._CacheKey("\\server\share\Report.PDF"), "ext:.pdf")
        eq("remote exe key", IconCache._CacheKey("\\server\share\tool.exe"), "ext:.exe")
        ; a folder with a dot in its name has no "extension" (it used to get a blank icon)
        eq("remote dotted folder", IconCache._CacheKey("\\server\Proposals\2025\26. 18 Sample Road (A)"), "folder:")
        eq("remote dotted folder slash", IconCache._CacheKey("\\server\share\10.PT1002-29Harbor\"), "folder:")
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
        ; disconnected network drives: whether a drive is reachable is kept in roots, so it isn't waited for item by item
        offline := Map("Q:", false)
        eq("drive offline", check("Folder", "Q:\Projects\PT1001", offline), "Unavailable")
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
        FileIndex.Paths := paths, FileIndex.Names := [], FileIndex.Folders := [0, 0, 0, 0, 0]
        for fullPath in paths {
            SplitPath(fullPath, &fileName)
            FileIndex.Names.Push(StrLower(fileName))
        }
        eq("index: before initials are ready", names("zjsm"), "")
        FileIndex.BuildInitials()                                             ; computed in the background in batches
        eq("index: initials built in the background", TestUtil.WaitFor(() => FileIndex._initialsFor == FileIndex.Names), true)
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
        FileIndex.Names := FileIndex.Names.Clone()                            ; Names replaced (e.g. set directly by a test): the initials don't match, so no pinyin
        eq("index: initials for other names", names("zjsm"), "")
        savedFile := FileIndex.File, FileIndex.File := A_Temp "\ALTRunPinyinIndex.json"
        try {
            JSON.WriteFile(FileIndex.File, Map("Paths", ["C:\D\证件扫描文件.zip", "C:\D\notes.txt"], "Folders", [0, 0]))
            FileIndex._LoadCache()
            TestUtil.WaitFor(() => FileIndex._initialsFor == FileIndex.Names)
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
        files := ProviderRegistry.SearchOnly(FileSearchProvider, "mar")
        TestRunner.True("FileSearchModes.file mode finds files", files.Length >= 2 && files[1].Provider = "FileSearch")
        eq("file mode empty", ProviderRegistry.SearchOnly(FileSearchProvider, "  ").Length, 0)
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
        eq("file mode folder keyword", titles(ProviderRegistry.SearchOnly(FileSearchProvider, "folder atlas")), "Atlas Tower|Old Atlas")
        eq("file mode plain", titles(ProviderRegistry.SearchOnly(FileSearchProvider, "atlas")), "Atlas Tower|Atlas drawing.dwg|Old Atlas|noteatlas.pdf")
        ; Only the keyword (no space yet): the hint ranks low and does not hide commands with folder in their name
        hint := FileSearchProvider.Search(SearchQuery("folder"))
        eq("keyword alone hint", hint.Length, 1)
        eq("keyword alone not exclusive", hint[1].Exclusive, false)
        eq("keyword alone low score", hint[1].Score < 10, true)
        eq("open alone not exclusive", FileSearchProvider.Search(SearchQuery("open"))[1].Exclusive, false)
        eq("keyword with space exclusive", FileSearchProvider.Search(SearchQuery("folder "))[1].Exclusive, true)
        ; the hint completes to the keyword typed and a space (Enter / Tab / click); nothing to complete once the space is there
        complete := (text) => FileSearchProvider.Search(SearchQuery(text))[1].AutoComplete
        eq("hint completes: folder", complete("folder"), "folder ")
        eq("hint completes: the keyword typed", complete("Find"), "find ")
        eq("hint completes: file type", complete("doc"), "doc ")
        eq("hint completes: recent", complete("recent"), "recent ")
        eq("keyword with space: no completion", complete("folder "), "")
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

    ; Typing ? shows the cheat sheet; usage tips rotate in the empty search box
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
    static RecentItems() {
        eq := (n, a, e) => TestRunner.Equal("RecentItems." n, a, e)
        eq("default count", AppSettings.Defaults()["Features"]["Recent"]["RecentCount"], 0)   ; by default only pinned items show
        saved := {Recent: Knowledge.Recent, File: Knowledge.File, SettingsFile: AppSettings.File}
        options := AppSettings.Feature("Recent"), savedPinned := options["Pinned"], savedCount := options["RecentCount"]
        Knowledge.Recent := [], Knowledge.File := A_Temp "\ALTRun-recent-test.json", AppSettings.File := A_Temp "\ALTRun-recent-settings.json"
        options["Pinned"] := [], options["RecentCount"] := 2
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
        options["RecentCount"] := 0
        eq("count 0", RecentProvider.EmptyResults().Length, 0)
        Knowledge.Recent := saved.Recent, Knowledge.File := saved.File, AppSettings.File := saved.SettingsFile
        options["Pinned"] := savedPinned, options["RecentCount"] := savedCount
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
        eq("default", AppSettings.Defaults()["Features"]["Clipboard"]["LocalFiles"], 1)
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
        hint := ScriptProvider.Search(SearchQuery("script"))                ; without the space: only a hint below the other results
        eq("keyword without space: a hint", hint.Length "|" hint[1].Valid "|" hint[1].Exclusive "|" hint[1].Subtitle, "1|0|0|script ...")
        eq("keyword without space: hint comes last", hint[1].Score, 5)
        eq("keyword without space: Enter / Tab complete", hint[1].AutoComplete, "script ")
        eq("keyword with space: no hint", titles("script "), "Clean Temp|Ping Host|Restart Explorer")
        eq("default script keyword", AppSettings.Defaults()["Features"]["Scripts"]["Keyword"], "script")
        eq("ps1 command", ScriptProvider.CommandLine(byTitle["Ping Host"], "a b"), 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' root '\ping.ps1" "a b"')
        eq("bat command", ScriptProvider.CommandLine(byTitle["Clean Temp"]), A_ComSpec ' /c ""' root '\clean.bat""')
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
            eq("examples: count", found.Count, 3)
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

    ; Snippets: the text is searched too
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
}
