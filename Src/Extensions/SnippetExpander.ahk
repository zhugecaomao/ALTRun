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
; 修改片段后 ALTRun 自动重新载入, 新的关键字随即生效。
;
; 用法:
;   SnippetExpander.Init()            启动时, 在 SearchWindow.Create() 之后调用
;===============================================================================

class SnippetExpander {
    static Count := 0

    static Init() {
        options := AppSettings.Feature("Snippets")
        if !(options["Enabled"] && options["AutoExpand"])
            return
        HotIfWinNotActive("ahk_id " SearchWindow.Gui.Hwnd)
        for snippet in AppSettings.Snippets {
            abbreviation := SnippetExpander.Abbreviation(snippet, options["ExpandPrefix"])
            if (abbreviation = "")
                continue
            try {
                ; * = 不需要结束符, ? = 前面紧挨着其它字符也触发; 输入的缩写由 AHK 自动删除
                Hotstring(":*?:" abbreviation, SnippetExpander._Expander(snippet["Text"]))
                SnippetExpander.Count++
            } catch as e {
                Logger.Error("SnippetExpander: cannot register " abbreviation " - " e.Message)
            }
        }
        HotIfWinNotActive()
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

    ; 单独一个方法生成闭包, 每个缩写各自记住自己的正文
    static _Expander(text) {
        return (*) => SnippetProvider.Paste(text, false)
    }
}
