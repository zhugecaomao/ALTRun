;===============================================================================
; WebSearchProvider.ahk - web search + fallback results (AutoHotkey v2)
;-------------------------------------------------------------------------------
; ALTRun.json -> Features.WebSearch.Engines, each engine:
;   { "Id": "google", "Keyword": "g", "Title": "Google", "Url": "https://...?q={query}", "Icon": "" }
; Url can hold several addresses, one per line (#565): one search opens them all, in that order (at most MaxUrls).
;   The result's Arg holds the filled addresses one per line; ActionCatalog.OpenUrl opens each line.
; "g words" searches with that engine; typing only the keyword "g" offers completion. Keyword can hold several, separated by commas ("g, ㅎ").
; With {query|default} in the Url it opens even without search words after the keyword, using the default, e.g. "http://localhost:{query|8000}".
; A default of clipboard ({query|clipboard}) searches the text on the clipboard (text only, joined into one line, at most MaxTermChars characters).
; {query|} (an empty default) opens the URL with nothing in place of {query}, e.g. "https://www.baidu.com/s?wd={query|}" opens the site.
; The default is used only with a space after the keyword and no search words ("g "); only the keyword ("g") offers completion.
; The rules are shared with custom commands (SearchQuery.PlaceholderDefault / FillPlaceholder).
; Icon: the result's icon, a .ico / .png / .exe etc. file (relative paths start from the Data folder) or an IconCache form
; (res:, ext: etc.); if not given, built-in engines (by Id) use the icons in Resources\Icons\Web, others the default browser's icon.
; "Download site icon" in the edit dialog: only when the user clicks does it visit the URL's site (no third-party icon service); the icon found goes into Data\Icons.
;
; Fallbacks: shown when no other feature has results, e.g. "Search Google for 'xxx'".
; The list holds engine Ids, or "files" (handed to FileSearchProvider to search files).
;===============================================================================

class WebSearchProvider {
    static Id := "WebSearch"

    static Init() {
    }

    static Engines() {
        return AppSettings.Feature("WebSearch")["Engines"]
    }

    static Search(query) {
        results := []
        for engine in WebSearchProvider.Engines() {
            if !(engine is Map) || !engine.Has("Keyword") || !(keywords := SearchQuery.Keywords(engine["Keyword"])).Length
                continue
            if query.MatchKeyword(keywords, &term, &keyword) {
                item := WebSearchProvider.ItemFor(engine, term, keyword, !query.HasRest)   ; only "g" -> offer completion; "g " -> default / clipboard
                item.Score := 150
                item.Exclusive := query.HasRest
                results.Push(item)
            } else if (!query.HasRest && StrLen(query.Text) >= 2 && FuzzyMatcher.Score(query.Text, engine["Title"]) >= 80) {
                item := WebSearchProvider.ItemFor(engine, "", "", true)      ; typing an engine's name -> offer "g " completion (without reading the clipboard)
                item.Score := 20
                results.Push(item)
            }
        }
        return results
    }

    ; keyword: the keyword typed (kept when completing); the first if none
    ; hintOnly: only offer completion (just the keyword typed, or the engine's name): no default and no clipboard reading
    static ItemFor(engine, term, keyword := "", hintOnly := false) {
        icon := WebSearchProvider.IconFor(engine)
        if (keyword = "")
            keyword := WebSearchProvider.FirstKeyword(engine)
        autoComplete := keyword " " term
        fromClipboard := false, empty := false
        if (term = "" && !hintOnly && SearchQuery.PlaceholderDefault(engine["Url"], &defaultText)) {
            if SearchQuery.IsClipboardDefault(defaultText) {
                term := WebSearchProvider.ClipboardText(), fromClipboard := true
                if (term = "")
                    return ResultItem(I18n.T("Web.SearchClipboard", engine["Title"]), I18n.T("Web.ClipboardEmpty"), {
                        Kind: "url", Icon: icon, Valid: false, AutoComplete: autoComplete, Source: engine
                    })
            } else
                term := defaultText, empty := (defaultText = "")          ; {query|}: open it with nothing in place of {query}
        }
        if empty {
            target := WebSearchProvider.FillAll(engine["Url"], "")
            return ResultItem(I18n.T("Web.Open", engine["Title"]), WebSearchProvider.Subtitle(target), {
                Kind: "url", Arg: target, Icon: icon, AutoComplete: autoComplete, Source: engine
            })
        }
        if (term = "") {
            return ResultItem(I18n.T("Web.SearchEmpty", engine["Title"]), I18n.T("Search.TypeAfter", keyword), {
                Kind: "url", Icon: icon, Valid: false, AutoComplete: autoComplete, Source: engine
            })
        }
        target := WebSearchProvider.FillAll(engine["Url"], term)
        title := fromClipboard ? I18n.T("Web.SearchClipboardFor", engine["Title"], WebSearchProvider.Preview(term)) : I18n.T("Web.SearchFor", engine["Title"], term)
        return ResultItem(title, WebSearchProvider.Subtitle(target), {
            Kind: "url", Arg: target, Icon: icon, AutoComplete: autoComplete, Source: engine
        })
    }

    static FirstKeyword(engine) {
        keywords := SearchQuery.Keywords((engine is Map && engine.Has("Keyword")) ? engine["Keyword"] : "")
        return keywords.Length ? keywords[1] : ""
    }

    ; {query} / {query|default} in the URL -> the search words (encoded); with empty search words the default (clipboard = the text on the clipboard)
    ; readClipboard: false = don't read the clipboard even when the default is clipboard (when only the rest of the URL is needed)
    static Fill(template, term, readClipboard := true) {
        if (term = "") {
            term := WebSearchProvider.DefaultTerm(template)
            if SearchQuery.IsClipboardDefault(term)
                term := readClipboard ? WebSearchProvider.ClipboardText() : ""
        }
        return SearchQuery.FillPlaceholder(template, Url.Encode(term))
    }

    ; The addresses in an engine's Url, one per line (blank lines skipped), at most MaxUrls
    static MaxUrls := 10
    static Urls(template) {
        urls := []
        for line in StrSplit(template, "`n", "`r `t")
            if (line != "" && urls.Length < WebSearchProvider.MaxUrls)
                urls.Push(line)
        return urls
    }

    static FirstUrl(template) => (urls := WebSearchProvider.Urls(template)).Length ? urls[1] : ""

    ; Fill for each address in template; the results one per line (the default of {query|default} is read from each line)
    static FillAll(template, term, readClipboard := true) {
        filled := ""
        for line in WebSearchProvider.Urls(template)
            filled .= (filled = "" ? "" : "`n") WebSearchProvider.Fill(line, term, readClipboard)
        return filled
    }

    ; A result's second line: the address, or the first one and how many more
    static Subtitle(target) {
        lines := StrSplit(target, "`n")
        return (lines.Length > 1) ? I18n.T("Web.MoreUrls", lines[1], lines.Length - 1) : target
    }

    ; The default in {query|default}, "" if none (or {query|})
    static DefaultTerm(template) {
        SearchQuery.PlaceholderDefault(template, &value)
        return value
    }

    ; the default is clipboard (case-insensitive)
    static UsesClipboard(template) => SearchQuery.IsClipboardDefault(WebSearchProvider.DefaultTerm(template))

    static MaxTermChars := 200

    ; The text on the clipboard (processed by Term); "" when files were copied, there is no text, or a password manager marked it "don't record".
    ; With both text and an image (e.g. cells copied from Excel) the text is used, as in the clipboard history
    static ClipboardText() {
        if (ClipboardData.HasFiles() || ClipboardProvider.IsPrivate())
            return ""
        text := ""
        try text := A_Clipboard
        return WebSearchProvider.Term(text)
    }

    ; Search words for the URL: white space collapsed to one space, at most MaxTermChars characters (search sites reject URLs that are too long)
    static Term(text) {
        text := Trim(TextTools.Replace(text, "\s+", " "))
        return StrLen(text) > WebSearchProvider.MaxTermChars ? Trim(SubStr(text, 1, WebSearchProvider.MaxTermChars)) : text
    }

    ; Search words shown in the result title: shortened when too long
    static Preview(text) => StrLen(text) > 60 ? TextTools.Head(text, 60) "..." : text

    ; The engine's Icon: image files as thumbnails, .ico / .exe etc. with their own icon; when not given or the file doesn't exist,
    ; built-in engines use their own icon (BuiltinIcon), others the default browser's icon
    static IconFor(engine) {
        spec := engine.Has("Icon") ? Trim(engine["Icon"], " `t`"") : ""
        if (spec = "")
            return WebSearchProvider.BuiltinIcon(engine)
        if RegExMatch(spec, "i)^(res|ext|url|folder|thumb|shell):")
            return spec
        if !RegExMatch(spec, "i)^([a-z]:\\|\\\\|%|A_)")                    ; relative paths start from the Data folder
            spec := AppSettings.DataDir "\" spec
        iconFile := Path.Resolve(spec)
        if !FileExist(iconFile)
            return WebSearchProvider.BuiltinIcon(engine)
        return RegExMatch(iconFile, "i)\.(png|jpe?g|bmp|gif)$") ? "thumb:" iconFile : iconFile
    }

    ; Icons of the built-in engines (Google, Bing, GitHub...): Resources\Icons\Web\<Id>.png (each site's own icon), the default browser's icon if missing
    static BuiltinIconDir := A_ScriptDir "\Resources\Icons\Web", _builtinIcons := Map()
    static BuiltinIcon(engine) {
        id := (engine.Has("Id") && engine["Id"] != "") ? engine["Id"] : ""
        if (id = "" || !RegExMatch(id, "^[\w-]+$"))
            return "url:"
        found := WebSearchProvider._builtinIcons
        if !found.Has(id) {
            iconFile := WebSearchProvider.BuiltinIconDir "\" id ".png"
            found[id] := FileExist(iconFile) ? "thumb:" iconFile : "url:"
        }
        return found[id]
    }

    ;---------------------------------------------------------------------------
    ; Download site icon (the button in the edit dialog)
    ;---------------------------------------------------------------------------
    static MaxIconBytes := 1024 * 1024, MaxPageBytes := 2 * 1024 * 1024

    ; Downloads the icon of url's site and stores it as Data\Icons\<name>.<extension>.
    ; Returns {File: "Icons\x.png"} (relative to the Data folder, so it works on another PC) or {Error: message}.
    ; First the icons the home page names (apple-touch-icon is usually a 180 px PNG, the sharpest), then /apple-touch-icon.png and /favicon.ico.
    ; Only PNG / ICO / JPG / GIF / BMP are accepted (by file header), at most 1 MB. fetch is passed only by tests (instead of the network)
    static DownloadIcon(address, name, fetch := "") {
        fetch := IsObject(fetch) ? fetch : (target) => WebSearchProvider._Fetch(target)
        page := WebSearchProvider.Fill(WebSearchProvider.FirstUrl(address), "", false)   ; the first address's site, no clipboard reading
        if !RegExMatch(page, "i)^(https?://[^/?#{}]+)", &m)
            return {Error: I18n.T("Web.IconBadUrl")}
        origin := RTrim(m[1], ":")
        candidates := []
        try {
            response := fetch(origin "/")
            if (response.Status = 200 && response.Size && response.Size <= WebSearchProvider.MaxPageBytes)
                candidates := WebSearchProvider.IconCandidates(StrGet(response.Body, Min(response.Size, 512 * 1024), "UTF-8")
                    , response.Url != "" ? response.Url : origin "/")
        }
        for fallback in [origin "/apple-touch-icon.png", origin "/favicon.ico"]
            if !WebSearchProvider._Contains(candidates, fallback)
                candidates.Push(fallback)
        name := RegExReplace(name, "[^\w.-]", "_")
        if (name = "")
            name := RegExReplace(RegExReplace(origin, "i)^https?://"), "[^\w.-]", "_")
        for candidate in candidates {
            try response := fetch(candidate)
            catch
                continue
            if (response.Status != 200 || !response.Size || response.Size > WebSearchProvider.MaxIconBytes)
                continue
            if ((ext := WebSearchProvider.ImageType(response.Body, response.Size)) = "")
                continue
            folder := AppSettings.DataDir "\Icons"
            DirCreate(folder)
            for old in ["png", "ico", "jpg", "gif", "bmp"]                     ; an icon of the same name downloaded earlier (maybe another format)
                try FileDelete(folder "\" name "." old)
            f := FileOpen(folder "\" name "." ext, "w")
            f.RawWrite(response.Body, response.Size)
            f.Close()
            IconCache.Clear()                                               ; the old icon of the same file name is still in the cache
            Logger.Debug("WebSearch: downloaded icon " candidate " -> Icons\" name "." ext)
            return {File: "Icons\" name "." ext}
        }
        Logger.Debug("WebSearch: no icon found for " origin)
        return {Error: I18n.T("Web.IconNotFound", origin)}
    }

    ; The icons a page names in <link rel="apple-touch-icon"> / <link rel="icon">, sharpest first (apple-touch-icon, then by sizes);
    ; no SVG (can't be drawn) and no data: URLs. pageUrl: the page's URL, to turn relative paths into full URLs
    static IconCandidates(html, pageUrl) {
        found := []
        pos := 1
        while (pos := TextTools.Match(html, "is)<link\b[^>]*>", &m, pos)) {
            tag := m[0], pos += StrLen(tag)
            rel := WebSearchProvider._Attribute(tag, "rel"), href := WebSearchProvider._Attribute(tag, "href")
            if (href = "" || !RegExMatch(rel, "i)(^|\s)(apple-touch-icon(-precomposed)?|icon)(\s|$)"))
                continue
            href := StrReplace(href, "&amp;", "&")
            if (RegExMatch(href, "i)^data:|\.svgz?([?#]|$)") || InStr(WebSearchProvider._Attribute(tag, "type"), "svg"))
                continue
            apple := InStr(rel, "apple-touch-icon") > 0
            size := RegExMatch(WebSearchProvider._Attribute(tag, "sizes"), "(\d+)x\d+", &sz) ? Integer(sz[1]) : (apple ? 180 : 32)
            found.Push({Url: WebSearchProvider._Absolute(href, pageUrl), Score: (apple ? 10000 : 0) + size})
        }
        ; by Score, highest first (insertion sort, usually only a few)
        Loop found.Length - 1 {
            i := A_Index + 1, current := found[i], j := i - 1
            while (j >= 1 && found[j].Score < current.Score)
                found[j + 1] := found[j], j--
            found[j + 1] := current
        }
        urls := []
        for item in found
            if !WebSearchProvider._Contains(urls, item.Url)
                urls.Push(item.Url)
        return urls
    }

    ; Which image type the file header is: "png" / "ico" / "jpg" / "gif" / "bmp", "" if not an image
    static ImageType(buf, size) {
        byte := (i) => NumGet(buf, i, "UChar")
        if (size >= 8 && byte(0) = 0x89 && byte(1) = 0x50 && byte(2) = 0x4E && byte(3) = 0x47)
            return "png"
        if (size >= 6 && byte(0) = 0 && byte(1) = 0 && byte(2) = 1 && byte(3) = 0)
            return "ico"
        if (size >= 3 && byte(0) = 0xFF && byte(1) = 0xD8 && byte(2) = 0xFF)
            return "jpg"
        if (size >= 6 && byte(0) = 0x47 && byte(1) = 0x49 && byte(2) = 0x46 && byte(3) = 0x38)
            return "gif"
        if (size >= 26 && byte(0) = 0x42 && byte(1) = 0x4D)
            return "bmp"
        return ""
    }

    static _Attribute(tag, name) {
        return RegExMatch(tag, 'is)\s' name '\s*=\s*(?:"([^"]*)"|`'([^`']*)`'|([^\s>]+))', &m) ? Trim(m[1] m[2] m[3]) : ""
    }

    static _Absolute(href, base) {
        if RegExMatch(href, "i)^https?://")
            return href
        if !RegExMatch(base, "i)^(https?:)//([^/?#]+)", &b)
            return href
        if (SubStr(href, 1, 2) = "//")
            return b[1] href
        if (SubStr(href, 1, 1) = "/")
            return b[1] "//" b[2] href
        dir := RegExReplace(base, "[?#].*")
        dir := RegExMatch(dir, "i)^https?://[^/]+$") ? dir "/" : RegExReplace(dir, "[^/]*$")
        return dir href
    }

    static _Contains(list, value) {
        for item in list
            if (item = value)
                return true
        return false
    }

    ; GETs a URL: {Status, Body (Buffer), Size, Url (after redirects)}; throws if it can't connect
    static _Fetch(target) {
        request := ComObject("WinHttp.WinHttpRequest.5.1")
        request.SetTimeouts(5000, 5000, 8000, 8000)
        request.Open("GET", target, false)
        request.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) ALTRun/" App.Version)
        request.Send()
        size := 0, body := Buffer(0)
        try {
            data := request.ResponseBody                                    ; byte array (SAFEARRAY)
            size := data.MaxIndex() + 1
            if (size > 0) {
                body := Buffer(size)
                DllCall("RtlMoveMemory", "Ptr", body, "Ptr", NumGet(ComObjValue(data) + 8 + A_PtrSize, "Ptr"), "Ptr", size)
            }
        }
        finalUrl := ""
        try finalUrl := request.Option(1)                                   ; WinHttpRequestOption_URL
        return {Status: request.Status, Body: body, Size: size, Url: finalUrl}
    }

    static EditorFields() {
        return [ItemEditor.Field("Title", "Prefs.Col.Title", "text", true)
              , ItemEditor.Field("Keyword", "Prefs.Col.Keyword", "text", true, "", I18n.T("Web.Field.Keyword"))
              , WebSearchProvider._UrlField()
              , ItemEditor.Field("Id", "Prefs.Col.Id", "text", true)
              , WebSearchProvider._IconField()]
    }

    ; The Url field: a few lines, one address per line; long addresses don't wrap (each line is one address) and scroll
    ; sideways with the caret, without a horizontal scroll bar (always shown, even when not needed)
    static _UrlField() {
        field := ItemEditor.Field("Url", "Prefs.Col.Url", "multiline", true, "", I18n.T("Web.Field.Url"))
        field.EditOptions := "r3 -Wrap"
        return field
    }

    ; The Icon field: text box + Browse + "Download site icon" (finds the site's icon from the URL, stores it in Data\Icons and fills it in)
    static _IconField() {
        field := ItemEditor.Field("Icon", "Prefs.Col.Icon", "file", false, "", I18n.T("Web.Field.Icon"))
        field.Button := {Label: I18n.T("Web.IconDownload"), OnClick: (controls, g) => WebSearchProvider._DownloadClicked(controls, g)}
        return field
    }

    static _DownloadClicked(controls, g) {
        g.Opt("+OwnDialogs")
        address := WebSearchProvider.FirstUrl(controls["Url"].Value)
        if !TextTools.Match(address, "i)^https?://")
            return MsgBox(I18n.T("Web.IconBadUrl"), App.Name, 48)
        name := controls.Has("Id") && Trim(controls["Id"].Value) != "" ? Trim(controls["Id"].Value) : (controls.Has("Keyword") ? Trim(controls["Keyword"].Value) : "")
        cursor := DllCall("SetCursor", "Ptr", DllCall("LoadCursor", "Ptr", 0, "Ptr", 32514, "Ptr"), "Ptr")   ; IDC_WAIT
        try result := WebSearchProvider.DownloadIcon(address, name)
        catch as e
            result := {Error: e.Message}
        DllCall("SetCursor", "Ptr", cursor)
        if result.HasOwnProp("File")
            controls["Icon"].Value := result.File
        else
            MsgBox(result.Error, App.Name, 48)
    }

    static NewEngine() {
        return Map("Id", "", "Keyword", "", "Title", "", "Url", "https://", "Icon", "")
    }

    ; F3 / right-click "Edit" in the results: edit that search engine
    static EditItem(item) {
        engine := item.Source
        if !IsObject(engine)
            return false
        edited := ItemEditor.Edit(ItemEditor.Owner(), I18n.T("Prefs.Page.WebSearch"), WebSearchProvider.EditorFields(), engine)
        if !IsObject(edited)
            return false
        for key, value in edited
            engine[key] := value
        return AppSettings.Save()
    }

    static DeleteItem(item) {
        engines := WebSearchProvider.Engines()
        for index, engine in engines {
            if (ObjPtr(engine) = ObjPtr(item.Source)) {
                engines.RemoveAt(index)
                return AppSettings.Save()
            }
        }
        return false
    }

    static Fallbacks(query) {
        results := []
        for id in AppSettings.Feature("WebSearch")["Fallbacks"] {
            if (id = "files") {
                if ProviderRegistry.IsEnabled(FileSearchProvider)
                    results.Push(FileSearchProvider.FallbackItem(query.Text))
                continue
            }
            for engine in WebSearchProvider.Engines() {
                if (engine is Map && engine.Has("Id") && engine["Id"] = id) {
                    results.Push(WebSearchProvider.ItemFor(engine, query.Text))
                    break
                }
            }
        }
        return results
    }
}
