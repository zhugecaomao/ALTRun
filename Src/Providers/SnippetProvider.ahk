;===============================================================================
; SnippetProvider.ahk - 文字片段 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; ALTRun.json -> Snippets, 每一条: { "Name": "名称", "Keyword": "关键字", "Text": "正文" }
; 搜索名称或关键字, Enter 把正文粘贴到呼出 ALTRun 之前的窗口。输入 "snip xxx"
; (Features.Snippets.Keyword) 只在片段里搜索, "snip" 单独输入列出全部片段。
;
; 正文里的占位符在粘贴时展开:
;   {date} {time} {datetime} {clipboard} {cursor} (粘贴后光标停在这里)
;   {date:yyyy-MM-dd} {time:HH:mm:ss} 自己指定格式 (FormatTime 的写法); {date+7} {date-1:dddd} 前后几天
;   {uuid} 随机的 UUID; {clipboard:1} {clipboard:2} 剪贴板历史里前一条、前两条 (和 Alfred 一样)
; 自动展开 (输入 ";关键字") 见 Src\Extensions\SnippetExpander.ahk。
;===============================================================================

class SnippetProvider {
    static Id := "Snippets"

    static Init() {
    }

    static Search(query) {
        options := AppSettings.Feature("Snippets")
        results := []
        onlySnippets := query.MatchKeyword([options["Keyword"]], &term)
        needle := onlySnippets ? term : query.Text
        for snippet in AppSettings.Snippets {
            if !(snippet is Map) || !snippet.Has("Text")
                continue
            name    := snippet.Has("Name") ? snippet["Name"] : ""
            keyword := snippet.Has("Keyword") ? snippet["Keyword"] : ""
            score := (onlySnippets && needle = "") ? 50 : FuzzyMatcher.Best(needle, [keyword, name])
            if (score <= 0)
                continue
            preview := SnippetProvider._Preview(snippet["Text"])
            item := ResultItem((name != "") ? name : preview, I18n.T("Snippet.Subtitle", preview), {
                Kind: "text", Arg: snippet["Text"], Icon: "res:imageres.dll,-102",
                Uid: "snippet:" StrLower(keyword "|" name), Score: score + (onlySnippets ? 30 : 0), Source: snippet,
                Exclusive: onlySnippets && query.HasRest,
                LargeText: SnippetProvider.Expand(snippet["Text"])
            })
            item.OnRun := (resultItem) => SnippetProvider.Paste(resultItem.Arg)
            results.Push(item)
        }
        return results
    }

    static EditorFields() {
        return [ItemEditor.Field("Name", "Prefs.Col.Name", "text", true)
              , ItemEditor.Field("Keyword", "Prefs.Col.Keyword")
              , ItemEditor.Field("Text", "Prefs.Col.Text", "multiline", true, "", "{date} {date:yyyy-MM-dd} {date+7} {time} {time:HH:mm:ss} {datetime} {clipboard} {clipboard:1} {uuid} {cursor}")
              , ItemEditor.Field("AutoExpand", "Prefs.Col.AutoExpand", "check")]
    }

    static NewSnippet() {
        return Map("Name", "", "Keyword", "", "Text", "", "AutoExpand", 1)
    }

    ; snippet: AppSettings.Snippets 里的一条 (就地修改), 或 "" 新建 (prefill 预先填好的字段)
    static Edit(snippet := "", prefill := "") {
        isNew := !IsObject(snippet)
        base := isNew ? SnippetProvider.NewSnippet() : ItemEditor.WithDefaults(snippet, SnippetProvider.NewSnippet())
        if IsObject(prefill)
            for key, value in prefill
                base[key] := value
        edited := ItemEditor.Edit(ItemEditor.Owner(), I18n.T("Prefs.Page.Snippets"), SnippetProvider.EditorFields(), base)
        if !IsObject(edited)
            return false
        if isNew
            AppSettings.Snippets.Push(edited)
        else
            for key, value in edited
                snippet[key] := value
        saved := AppSettings.Save()
        SnippetExpander.Refresh()                                           ; 关键字可能变了, 重新注册自动展开
        return saved
    }

    static EditItem(item) {
        return SnippetProvider.Edit(item.Source)
    }

    static DeleteItem(item) {
        for index, snippet in AppSettings.Snippets {
            if (ObjPtr(snippet) = ObjPtr(item.Source)) {
                AppSettings.Snippets.RemoveAt(index)
                saved := AppSettings.Save()
                SnippetExpander.Refresh()
                return saved
            }
        }
        return false
    }

    ; focusPrevious = false: 自动展开时直接粘贴到当前窗口
    static Paste(text, focusPrevious := true) {
        text := SnippetProvider.Expand(text)
        caretBack := 0
        if (cursorPos := InStr(text, "{cursor}")) {
            text := StrReplace(text, "{cursor}")
            caretBack := StrLen(text) - cursorPos + 1
        }
        if !ActionCatalog.PasteText(text, focusPrevious)
            return
        if (caretBack > 0)
            SendInput("{Left " caretBack "}")
    }

    static Expand(text) {
        if !InStr(text, "{")
            return text
        dateFormat := AppSettings.Extension("AutoDate")["DateFormat"]
        ; {date} {time} {datetime}, 可以带 +N / -N 天和 :格式
        pos := 1
        while (pos := RegExMatch(text, "\{(datetime|date|time)([+-]\d+)?(?::([^{}]*))?\}", &m, pos)) {
            stamp := (m[2] != "") ? DateAdd(A_Now, Integer(m[2]), "Days") : A_Now
            fmt := (m[3] != "") ? m[3] : (m[1] = "date") ? dateFormat : (m[1] = "time") ? "HH:mm" : dateFormat " HH:mm"
            value := FormatTime(stamp, fmt)
            text := SubStr(text, 1, pos - 1) value SubStr(text, pos + m.Len)
            pos += StrLen(value)
        }
        while InStr(text, "{uuid}")                                        ; 每个 {uuid} 各不相同
            text := StrReplace(text, "{uuid}", SnippetProvider.NewUuid(), , , 1)
        pos := 1                                                            ; {clipboard:N}: 剪贴板历史里往前第 N 条
        while (pos := RegExMatch(text, "\{clipboard:(\d+)\}", &m, pos)) {
            index := Integer(m[1]) + 1, entries := ClipboardProvider.Entries
            value := (index = 1) ? A_Clipboard : (index <= entries.Length) ? entries[index]["Text"] : ""
            text := SubStr(text, 1, pos - 1) value SubStr(text, pos + m.Len)
            pos += StrLen(value)
        }
        text := StrReplace(text, "{clipboard}", A_Clipboard)
        return text
    }

    ; 小写、不带大括号的 UUID, 例如 3f2b8c1e-...
    static NewUuid() {
        guid := Buffer(16)
        DllCall("ole32\CoCreateGuid", "Ptr", guid)
        text := Buffer(78)
        DllCall("ole32\StringFromGUID2", "Ptr", guid, "Ptr", text, "Int", 39)
        return StrLower(Trim(StrGet(text), "{}"))
    }

    static _Preview(text) {
        text := RegExReplace(text, "\s+", " ")
        return (StrLen(text) > 80) ? SubStr(text, 1, 80) "..." : text
    }
}
