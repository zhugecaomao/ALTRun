;===============================================================================
; SnippetProvider.ahk - snippets (AutoHotkey v2)
;-------------------------------------------------------------------------------
; ALTRun.json -> Snippets, each entry: { "Name": "name", "Keyword": "keyword", "Text": "text" }
; Keyword can hold several, separated by commas ("sig, signature"); all work for search and auto expansion.
; Searches name, keyword and text (Features.Snippets.SearchText; the text needs at least 3 characters typed, every word must appear,
; and text-only matches come later); Enter pastes the text into the window active before ALTRun was shown. "snip xxx"
; (Features.Snippets.Keyword) searches snippets only; "snip" alone lists all snippets.
;
; Placeholders in the text are expanded when pasting:
;   {date} {time} {datetime} {clipboard} {cursor} (the caret stays here after pasting)
;   {date:yyyy-MM-dd} {time:HH:mm:ss} with your own format (FormatTime's syntax); {date+7} {date-1:dddd} days before / after
;   {uuid} a random UUID; {clipboard:1} {clipboard:2} the previous and the one before in the clipboard history (as in Alfred)
; Auto expansion (typing ";keyword"): see Src\Extensions\SnippetExpander.ahk.
;===============================================================================

class SnippetProvider {
    static Id := "Snippets"
    static TextMinChars := 3                                                ; the text is searched only from this many characters (shorter would hit a heap)
    static TextScore := 35                                                  ; score of text-only matches: after name / keyword matches and applications, commands

    static Init() {
    }

    static Icon => IconCache.Own("Snippet", "res:imageres.dll,-102")

    static Search(query) {
        options := AppSettings.Feature("Snippets")
        results := []
        onlySnippets := query.MatchKeyword([options["Keyword"]], &term), full := ""
        if (!onlySnippets && query.StartsKeyword([options["Keyword"]], &full))   ; "sni": the keyword's hint at the end of the list ("snip" alone lists the snippets)
            results.Push(ResultItem.KeywordHint(I18n.T("Snippets.Keyword"), full, SnippetProvider.Icon, true))
        needle := onlySnippets ? term : query.Text
        tokens := StrSplit(Trim(needle), " ")
        searchText := options["SearchText"] && StrLen(Trim(needle)) >= SnippetProvider.TextMinChars
        for snippet in AppSettings.Snippets {
            if !(snippet is Map) || !snippet.Has("Text")
                continue
            name    := snippet.Has("Name") ? snippet["Name"] : ""
            keyword := snippet.Has("Keyword") ? snippet["Keyword"] : ""
            score := (onlySnippets && needle = "") ? 50 : FuzzyMatcher.Best(needle, SnippetProvider._MatchTexts(keyword, name))
            found := searchText ? SnippetProvider.TextMatch(snippet["Text"], tokens) : 0
            if (found && score < SnippetProvider.TextScore)
                score := SnippetProvider.TextScore
            else
                found := 0
            if (score <= 0)
                continue
            preview := SnippetProvider._Preview(snippet["Text"], found)
            item := ResultItem((name != "") ? name : preview, I18n.T("Snippet.Subtitle", preview), {
                Kind: "text", Arg: snippet["Text"], Icon: SnippetProvider.Icon,
                Uid: "snippet:" StrLower(keyword "|" name), Score: score + (onlySnippets ? 30 : 0), Source: snippet,
                Exclusive: onlySnippets && query.HasRest,
                LargeText: SnippetProvider.Expand(snippet["Text"])
            })
            item.OnRun := (resultItem) => SnippetProvider.Paste(resultItem.Arg)
            results.Push(item)
        }
        return results
    }

    ; The text compared when searching: each keyword and the name
    static _MatchTexts(keyword, name) {
        texts := SearchQuery.Keywords(keyword).Clone()
        texts.Push(name)
        return texts
    }

    static EditorFields() {
        prefix := AppSettings.Feature("Snippets")["ExpandPrefix"]
        maxKeyword := SnippetExpander.MaxAbbreviation - StrLen(prefix)
        keyword := ItemEditor.Field("Keyword", "Prefs.Col.Keyword", , , , I18n.T("Snippet.Hint.Keyword", prefix, maxKeyword))
        keyword.Check := (value, edited) => SnippetProvider.CheckKeyword(value, edited, prefix)
        return [ItemEditor.Field("Name", "Prefs.Col.Name", "text", true, "", I18n.T("Snippet.Hint.Name"))
              , keyword
              , ItemEditor.Field("Text", "Prefs.Col.Text", "multiline", true, "", I18n.T("Snippet.Hint.Text") "`n{date} {date:yyyy-MM-dd} {date+7} {time} {datetime} {clipboard} {clipboard:1} {uuid} {cursor}")
              , ItemEditor.Field("AutoExpand", "Prefs.Col.AutoExpand", "check", , , I18n.T("Snippet.Hint.AutoExpand", prefix))]
    }

    ; On OK in the edit dialog: auto expansion with a keyword that is too long / has spaces -> a warning ("" = fine)
    static CheckKeyword(keyword, edited, prefix) {
        if (edited.Has("AutoExpand") && !edited["AutoExpand"])
            return ""
        for each in SearchQuery.Keywords(keyword)
            if (StrLen(prefix each) > SnippetExpander.MaxAbbreviation || RegExMatch(each, "\s"))
                return I18n.T("Snippet.KeywordTooLong", prefix, SnippetExpander.MaxAbbreviation)
        return ""
    }

    static NewSnippet() {
        return Map("Name", "", "Keyword", "", "Text", "", "AutoExpand", 1)
    }

    ;---------------------------------------------------------------------------
    ; Import / export (Preferences → Snippets → Import / Export): table columns Name | Keyword | Text | Auto expand (optional)
    ;---------------------------------------------------------------------------
    ; Pure function: merges rows (the result of DelimitedText.Parse) into snippets (in place). Rows with the same keyword (case-insensitive; without keywords,
    ; the same name) update, the others are appended; rows without text are skipped; a header first row is skipped. Without a name the keyword or the text's first line is used.
    ; Auto expand: on when blank, off for 0 / no / 否 etc.; also off when the keyword is too long or has spaces (auto expansion can't use it).
    ; Returns {Added, Updated, Skipped}
    static ImportRows(rows, snippets, prefix := "") {
        static Off := Map("0", 1, "no", 1, "n", 1, "false", 1, "off", 1, "否", 1, "いいえ", 1)
        result := {Added: 0, Updated: 0, Skipped: 0}
        for index, row in rows {
            if (index = 1 && SnippetProvider.IsHeaderRow(row))
                continue
            name := (row.Length >= 1) ? Trim(row[1]) : "", keyword := (row.Length >= 2) ? Trim(row[2]) : ""
            text := (row.Length >= 3) ? row[3] : "", expand := (row.Length >= 4) ? StrLower(Trim(row[4])) : ""
            if (Trim(text, " `t`r`n") = "") {
                result.Skipped++
                continue
            }
            if (name = "")
                name := (keyword != "") ? keyword : SubStr(Trim(StrSplit(Trim(text, " `t`r`n"), "`n", "`r")[1]), 1, 40)
            snippet := Map("Name", name, "Keyword", keyword, "Text", text, "AutoExpand", Off.Has(expand) ? 0 : 1)
            if (snippet["AutoExpand"] && SnippetProvider.CheckKeyword(keyword, snippet, prefix) != "")
                snippet["AutoExpand"] := 0
            existing := SnippetProvider._Find(snippets, keyword, name)
            if IsObject(existing) {
                for key, value in snippet
                    existing[key] := value
                result.Updated++
            } else {
                snippets.Push(snippet)
                result.Added++
            }
        }
        return result
    }

    ; The exported table: the first row is the header (column names in the current UI language)
    static ExportRows(snippets) {
        rows := [[I18n.T("Prefs.Col.Name"), I18n.T("Prefs.Col.Keyword"), I18n.T("Prefs.Col.Text"), I18n.T("Prefs.Col.AutoExpand")]]
        for snippet in snippets
            if (snippet is Map)
                rows.Push([snippet.Has("Name") ? snippet["Name"] : "", snippet.Has("Keyword") ? snippet["Keyword"] : ""
                    , snippet.Has("Text") ? snippet["Text"] : "", (!snippet.Has("AutoExpand") || snippet["AutoExpand"]) ? 1 : 0])
        return rows
    }

    ; a first column "Name" and a third column "Text" (in any UI language) make it a header
    static IsHeaderRow(row) {
        static Names := "|name|title|名称|名稱|名前|标题|標題|", Texts := "|text|content|正文|本文|内容|內容|"
        if (row.Length < 3)
            return false
        first := StrLower(Trim(row[1])), third := StrLower(Trim(row[3]))
        return ((InStr(Names, "|" first "|") || first = StrLower(I18n.T("Prefs.Col.Name")))
            && (InStr(Texts, "|" third "|") || third = StrLower(I18n.T("Prefs.Col.Text")))) ? true : false
    }

    static _Find(snippets, keyword, name) {
        for snippet in snippets {
            if !(snippet is Map)
                continue
            other := snippet.Has("Keyword") ? Trim(snippet["Keyword"]) : ""
            if (keyword != "" ? (other = keyword) : (other = "" && snippet.Has("Name") && Trim(snippet["Name"]) = name))
                return snippet
        }
        return ""
    }

    ; snippet: an entry of AppSettings.Snippets (changed in place), or "" for a new one (prefill = the prefilled fields)
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
        SnippetExpander.Refresh()                                           ; the keyword may have changed: register the auto expansion again
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

    ; focusPrevious = false: auto expansion pastes into the current window directly
    static Paste(text, focusPrevious := true) {
        text := SnippetProvider.Expand(text)
        caretBack := SnippetProvider.CaretBack(&text)
        if !ActionCatalog.PasteText(text, focusPrevious)
            return
        if (caretBack > 0)
            SendInput("{Left " caretBack "}")
    }

    ; Takes {cursor} out of the text; returns how many times Left moves the caret back there after pasting (0 = no {cursor}).
    ; A line break is one step for Left in Notepad, Word, browsers and edit controls alike, so `r`n counts once
    static CaretBack(&text) {
        if !(cursorPos := InStr(text, "{cursor}"))
            return 0
        text := StrReplace(text, "{cursor}")                                ; only the first counts; any others just go
        after := SubStr(text, cursorPos)
        return StrLen(StrReplace(after, "`r`n", "`n"))
    }

    static Expand(text) {
        if !InStr(text, "{")
            return text
        dateFormat := AppSettings.Extension("AutoDate")["DateFormat"]
        ; {date} {time} {datetime}, optionally with +N / -N days and :format
        pos := 1
        while (pos := RegExMatch(text, "\{(datetime|date|time)([+-]\d+)?(?::([^{}]*))?\}", &m, pos)) {
            stamp := (m[2] != "") ? DateAdd(A_Now, Integer(m[2]), "Days") : A_Now
            fmt := (m[3] != "") ? m[3] : (m[1] = "date") ? dateFormat : (m[1] = "time") ? "HH:mm" : dateFormat " HH:mm"
            value := FormatTime(stamp, fmt)
            text := SubStr(text, 1, pos - 1) value SubStr(text, pos + m.Len)
            pos += StrLen(value)
        }
        while InStr(text, "{uuid}")                                        ; each {uuid} differs
            text := StrReplace(text, "{uuid}", SnippetProvider.NewUuid(), , , 1)
        pos := 1                                                            ; {clipboard:N}: the Nth entry back in the clipboard history
        while (pos := RegExMatch(text, "\{clipboard:(\d+)\}", &m, pos)) {
            index := Integer(m[1]) + 1                                      ; only text counts (copied files and images are skipped)
            value := (index = 1) ? A_Clipboard : ClipboardProvider.TextAt(index)
            text := SubStr(text, 1, pos - 1) value SubStr(text, pos + m.Len)
            pos += StrLen(value)
        }
        text := StrReplace(text, "{clipboard}", A_Clipboard)
        return text
    }

    ; A lower-case UUID without braces, e.g. 3f2b8c1e-...
    static NewUuid() {
        guid := Buffer(16)
        DllCall("ole32\CoCreateGuid", "Ptr", guid)
        text := Buffer(78)
        DllCall("ole32\StringFromGUID2", "Ptr", guid, "Ptr", text, "Int", 39)
        return StrLower(Trim(StrGet(text), "{}"))
    }

    ; Whether every word appears in the text (case-insensitive, no fuzzy matching: fuzzy matching on long text would hit almost anything).
    ; Returns the position of the first word in the text, 0 unless all appear
    static TextMatch(text, tokens) {
        first := 0
        for token in tokens {
            if (token = "")
                continue
            if !(pos := InStr(text, token))
                return 0
            if !first
                first := pos
        }
        return first
    }

    ; found: where the text matched; when far in (past 60 characters) the preview starts a little before it, otherwise the match wouldn't be visible
    static _Preview(text, found := 0) {
        if (found > 60)
            text := "..." SubStr(text, found - 20)
        text := RegExReplace(text, "\s+", " ")
        return (StrLen(text) > 80) ? SubStr(text, 1, 80) "..." : text
    }
}
