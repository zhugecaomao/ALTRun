;===============================================================================
; SnippetExpander.ahk - 片段关键字自动展开 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 Alfred 的 Snippet 自动展开一样: 在任何程序里输入 "前缀 + 关键字" (默认前缀
; ";", 例如 ";sig"), 输入的文字自动删掉并换成片段正文 (占位符在展开时替换)。
; 前缀是为了避免平时正常打字时误触发; 在 ALTRun 自己的搜索窗口里不展开。
;
; 设置:
;   Features.Snippets.AutoExpand      1 = 开启
;   Features.Snippets.ExpandPrefix    关键字前面要加的前缀, 默认 ";"
;   Snippets[n].AutoExpand            单个片段设为 0 可以不自动展开 (仍可搜索)
; 修改片段后调用 Refresh(), 新的关键字随即生效。
;
; 用法:
;   SnippetExpander.Init()            启动时, 在 SearchWindow.Create() 之后调用
;===============================================================================

class SnippetExpander {
    static Count := 0
    static _texts := Map()                                                  ; 已注册的缩写 -> 正文

    static Init() {
        SnippetExpander.Refresh()
    }

    ; 按当前的 AppSettings.Snippets 重新注册: 删掉的缩写关闭, 新的 / 改过的缩写注册。
    ; 在搜索结果里编辑 / 删除片段后调用, 不需要重新载入 ALTRun。
    static Refresh() {
        options := AppSettings.Feature("Snippets")
        wanted := Map()
        wanted.CaseSense := "Off"                                           ; 和热字串一样不区分大小写
        if (options["Enabled"] && options["AutoExpand"]) {
            for snippet in AppSettings.Snippets {
                abbreviation := SnippetExpander.Abbreviation(snippet, options["ExpandPrefix"])
                if (abbreviation != "")
                    wanted[abbreviation] := snippet["Text"]
            }
        }
        HotIfWinNotActive("ahk_id " SearchWindow.Gui.Hwnd)
        for abbreviation in SnippetExpander._texts
            if !wanted.Has(abbreviation)
                try Hotstring(":*?:" abbreviation, , "Off")
        registered := Map()
        registered.CaseSense := "Off"
        for abbreviation, text in wanted {
            try {
                ; * = 不需要结束符, ? = 前面紧挨着其它字符也触发; 输入的缩写由 AHK 自动删除
                Hotstring(":*?:" abbreviation, SnippetExpander._Expander(abbreviation), "On")
                registered[abbreviation] := text
            } catch as e {
                Logger.Error("SnippetExpander: cannot register " abbreviation " - " e.Message)
            }
        }
        HotIfWinNotActive()
        SnippetExpander._texts := registered
        SnippetExpander.Count := registered.Count
        Logger.Debug("SnippetExpander: " SnippetExpander.Count " snippets")
    }

    ; 片段 -> 触发的缩写 ("" = 不自动展开)
    static Abbreviation(snippet, prefix) {
        if !(snippet is Map) || !snippet.Has("Keyword") || !snippet.Has("Text")
            return ""
        if (snippet.Has("AutoExpand") && !snippet["AutoExpand"])
            return ""
        keyword := Trim(snippet["Keyword"])
        if (keyword = "" || RegExMatch(keyword, "[\s``]"))                  ; 空白和反引号不能用在缩写里
            return ""
        return prefix keyword
    }

    ; 单独一个方法生成闭包; 正文在触发时再查, 这样修改片段后不必重新生成
    static _Expander(abbreviation) {
        return (*) => (Usage.Count("SnippetExpand"), SnippetProvider.Paste(SnippetExpander._texts[abbreviation], false))
    }
}
