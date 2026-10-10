;===============================================================================
; WebSearchProvider.ahk - 网页搜索 + 兜底结果 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; ALTRun.json -> Features.WebSearch.Engines, 每个引擎:
;   { "Id": "google", "Keyword": "g", "Title": "Google", "Url": "https://...?q={query}", "Icon": "" }
; 输入 "g 关键词" 用对应引擎搜索; 只输入关键字 "g" 时提示补全。
; Url 里写 {query|默认值} 时, 只输入关键字也能直接打开: 用默认值, 例如 "http://localhost:{query|8000}"。
; Icon: 结果的图标, 可以是 .ico / .png / .exe 等文件 (相对路径从 Data 文件夹算起), 或 IconCache 的写法
; (res:、ext: 等); 不写时自带的引擎 (按 Id) 用 Resources\Icons\Web 里的图标, 其它用默认浏览器的图标。
; 编辑框里的 "下载网站图标": 用户点了才访问网址所在的网站 (不经过第三方图标服务), 找到的图标存进 Data\Icons。
;
; Fallbacks: 其它功能都没有结果时显示的兜底项, 例如 "Search Google for 'xxx'"。
; 列表里写引擎的 Id, 或 "files" (交给 FileSearchProvider 搜索文件)。
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
            if !(engine is Map) || !engine.Has("Keyword") || engine["Keyword"] = ""
                continue
            if query.MatchKeyword([engine["Keyword"]], &term) {
                item := WebSearchProvider.ItemFor(engine, term)
                item.Score := 150
                item.Exclusive := query.HasRest
                results.Push(item)
            } else if (!query.HasRest && StrLen(query.Text) >= 2 && FuzzyMatcher.Score(query.Text, engine["Title"]) >= 80) {
                item := WebSearchProvider.ItemFor(engine, "")                ; 输入引擎名称 -> 提示 "g " 补全
                item.Score := 20
                results.Push(item)
            }
        }
        return results
    }

    static ItemFor(engine, term) {
        icon := WebSearchProvider.IconFor(engine)
        autoComplete := engine["Keyword"] " " term
        if (term = "")
            term := WebSearchProvider.DefaultTerm(engine["Url"])
        if (term = "") {
            return ResultItem(I18n.T("Web.SearchEmpty", engine["Title"]), I18n.T("Search.TypeAfter", engine["Keyword"]), {
                Kind: "url", Icon: icon, Valid: false, AutoComplete: autoComplete, Source: engine
            })
        }
        target := WebSearchProvider.Fill(engine["Url"], term)
        return ResultItem(I18n.T("Web.SearchFor", engine["Title"], term), target, {
            Kind: "url", Arg: target, Icon: icon, AutoComplete: autoComplete, Source: engine
        })
    }

    ; 网址里的 {query} / {query|默认值} 换成搜索词 (编码); 搜索词是空的时用默认值
    static Fill(template, term) {
        if (term = "")
            term := WebSearchProvider.DefaultTerm(template)
        return RegExReplace(template, "\{query(?:\|[^}]*)?\}", Url.Encode(term))   ; 编码后只有字母、数字和 %-_.~, 没有 $
    }

    ; {query|默认值} 里的默认值, 没有时 ""
    static DefaultTerm(template) {
        return RegExMatch(template, "\{query\|([^}]*)\}", &m) ? m[1] : ""
    }

    ; 引擎的 Icon: 图片文件用缩略图, .ico / .exe 等用文件自己的图标; 没写或文件不存在时
    ; 自带的引擎用自带的图标 (BuiltinIcon), 其它用默认浏览器的图标
    static IconFor(engine) {
        spec := engine.Has("Icon") ? Trim(engine["Icon"], " `t`"") : ""
        if (spec = "")
            return WebSearchProvider.BuiltinIcon(engine)
        if RegExMatch(spec, "i)^(res|ext|url|folder|thumb|shell):")
            return spec
        if !RegExMatch(spec, "i)^([a-z]:\\|\\\\|%|A_)")                    ; 相对路径: 从 Data 文件夹算起
            spec := AppSettings.DataDir "\" spec
        iconFile := Path.Resolve(spec)
        if !FileExist(iconFile)
            return WebSearchProvider.BuiltinIcon(engine)
        return RegExMatch(iconFile, "i)\.(png|jpe?g|bmp|gif)$") ? "thumb:" iconFile : iconFile
    }

    ; 自带的引擎 (Google、Bing、GitHub...) 的图标: Resources\Icons\Web\<Id>.png, 没有时用默认浏览器的图标
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
    ; 下载网站图标 (编辑框里的按钮)
    ;---------------------------------------------------------------------------
    static MaxIconBytes := 1024 * 1024, MaxPageBytes := 2 * 1024 * 1024

    ; 下载 url 所在网站的图标, 存成 Data\Icons\<name>.<扩展名>。
    ; 返回 {File: "Icons\x.png"} (相对 Data 文件夹, 换电脑也能用) 或 {Error: 提示文字}。
    ; 先看首页里写的图标 (apple-touch-icon 一般是 180 px 的 PNG, 最清晰), 再试 /apple-touch-icon.png 和 /favicon.ico。
    ; 只接受 PNG / ICO / JPG / GIF / BMP (看文件头), 最大 1 MB。fetch 只在测试时传入 (代替联网)
    static DownloadIcon(address, name, fetch := "") {
        fetch := IsObject(fetch) ? fetch : (target) => WebSearchProvider._Fetch(target)
        page := WebSearchProvider.Fill(address, "")
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
            for old in ["png", "ico", "jpg", "gif", "bmp"]                     ; 以前下载的同名图标 (可能是别的格式)
                try FileDelete(folder "\" name "." old)
            f := FileOpen(folder "\" name "." ext, "w")
            f.RawWrite(response.Body, response.Size)
            f.Close()
            IconCache.Clear()                                               ; 同一个文件名的旧图标还在缓存里
            Logger.Debug("WebSearch: downloaded icon " candidate " -> Icons\" name "." ext)
            return {File: "Icons\" name "." ext}
        }
        Logger.Debug("WebSearch: no icon found for " origin)
        return {Error: I18n.T("Web.IconNotFound", origin)}
    }

    ; 网页里 <link rel="apple-touch-icon"> / <link rel="icon"> 写的图标, 清晰的在前 (apple-touch-icon, 再按 sizes);
    ; 不要 SVG (画不了) 和 data: 网址。pageUrl: 网页的网址, 用来把相对路径换成完整网址
    static IconCandidates(html, pageUrl) {
        found := []
        pos := 1
        while (pos := RegExMatch(html, "is)<link\b[^>]*>", &m, pos)) {
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
        ; 按 Score 从大到小 (插入排序, 一般只有几条)
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

    ; 文件头是哪种图片: "png" / "ico" / "jpg" / "gif" / "bmp", 不是图片时 ""
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

    ; GET 一个网址: {Status, Body (Buffer), Size, Url (跳转后的网址)}; 连不上时抛出异常
    static _Fetch(target) {
        request := ComObject("WinHttp.WinHttpRequest.5.1")
        request.SetTimeouts(5000, 5000, 8000, 8000)
        request.Open("GET", target, false)
        request.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) ALTRun/" App.Version)
        request.Send()
        size := 0, body := Buffer(0)
        try {
            data := request.ResponseBody                                    ; 字节数组 (SAFEARRAY)
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
              , ItemEditor.Field("Keyword", "Prefs.Col.Keyword", "text", true)
              , ItemEditor.Field("Url", "Prefs.Col.Url", "text", true, "", I18n.T("Web.Field.Url"))
              , ItemEditor.Field("Id", "Prefs.Col.Id", "text", true)
              , WebSearchProvider._IconField()]
    }

    ; Icon 字段: 文本框 + 浏览 + "下载网站图标" (按网址找网站的图标, 存进 Data\Icons 并填进来)
    static _IconField() {
        field := ItemEditor.Field("Icon", "Prefs.Col.Icon", "file", false, "", I18n.T("Web.Field.Icon"))
        field.Button := {Label: I18n.T("Web.IconDownload"), OnClick: (controls, g) => WebSearchProvider._DownloadClicked(controls, g)}
        return field
    }

    static _DownloadClicked(controls, g) {
        g.Opt("+OwnDialogs")
        address := Trim(controls["Url"].Value)
        if !RegExMatch(address, "i)^https?://")
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

    ; 搜索结果里 F3 / 右键 "编辑": 修改对应的搜索引擎
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
