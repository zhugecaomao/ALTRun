;===============================================================================
; Clip.ahk - "Clip" 剪贴文本片段命令 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Command format:  Clip | <snippet body> | <short name to type>=<rank>
; The body is stored on ONE INI/JSON line, so real newlines/tabs are encoded:
;     \n = new line      \t = tab      \\ = a literal backslash
; Placeholders expanded at paste time:
;     {date} {time} {datetime} {clipboard} {arg} {cursor}
;
; 用法:
;   Clip.PasteClipText(rawText)      ; RunCommand() 执行 Clip 类型命令时调用
;   Clip.ClipPreview(text)           ; 主列表/状态栏里显示的单行预览
;   Clip.EditClipText()              ; 命令管理器里 "..." 按钮打开的多行编辑器
;
; 注意: 全局函数 NewClip() (F4 菜单 / "Func | NewClip | ..." 内置命令用它的函数名
; 做动态调用 %cmdPath%(), 必须是裸的全局函数) 仍留在 ALTRun.ahk 里, 只是内部
; 改成调用 Clip.EscapeClipText() - 详见那边的注释。
;===============================================================================

Class Clip {

    static EscapeClipText(text) {                                          ; Real text  ->  single INI line
        if (text = "")
            return ""
        text := StrReplace(text, "\", "\\")                                 ; backslash first
        text := StrReplace(text, "`r`n", "\n")
        text := StrReplace(text, "`n", "\n")
        text := StrReplace(text, "`r", "\n")
        text := StrReplace(text, "`t", "\t")
        return text
    }

    static UnescapeClipText(text) {                                         ; Single INI line  ->  real text
        if (text = "")
            return ""
        ; Placeholder trick keeps an escaped backslash (\\) from eating the next token.
        text := StrReplace(text, "\\", Chr(1))
        text := StrReplace(text, "\n", "`r`n")
        text := StrReplace(text, "\t", "`t")
        text := StrReplace(text, Chr(1), "\")
        return text
    }

    static ClipPreview(text, maxLen := 70) {                                ; One-line preview for ListView / StatusBar
        preview := StrReplace(StrReplace(text, "\n", " "), "\t", " ")
        preview := RegExReplace(preview, "\s{2,}", " ")
        return (StrLen(preview) > maxLen) ? SubStr(preview, 1, maxLen) " ..." : preview
    }

    static ExpandClipPlaceholders(text) {                                   ; {date} {time} {datetime} {clipboard} {arg}
        if (InStr(text, "{") = 0)
            return text
        text := StrReplace(text, "{date}", FormatTime(, "dd.MM.yyyy"))
        text := StrReplace(text, "{time}", FormatTime(, "HH:mm"))
        text := StrReplace(text, "{datetime}", FormatTime(, "dd.MM.yyyy HH:mm"))
        text := StrReplace(text, "{clipboard}", A_Clipboard)
        text := StrReplace(text, "{arg}", g_RUNTIME["Arg"])
        return text
    }

    static FocusLastWindow() {                                              ; Give focus back to the app the user came from
        target := g_RUNTIME["LastWin"]
        if (!target || !WinExist("ahk_id " target))
            return false
        if WinActive("ahk_id " target)
            return true
        try {
            WinActivate("ahk_id " target)
            WinWaitActive("ahk_id " target, , 1)
        } catch as e {
            g_LOG.Debug("FocusLastWindow: Failed to activate hwnd=" target ", " e.Message)
            return false
        }
        return WinActive("ahk_id " target) ? true : false
    }

    static PasteClipText(rawText) {                                         ; Main entry, called by RunCommand for type Clip
        text := Clip.ExpandClipPlaceholders(Clip.UnescapeClipText(rawText))
        if (text = "") {
            g_LOG.Debug("PasteClipText: Empty clip text, nothing to paste")
            return false
        }

        ; {cursor} marks where the caret should end up after pasting.
        caretBack := 0
        if (cursorPos := InStr(text, "{cursor}")) {
            text := StrReplace(text, "{cursor}", "")
            caretBack := StrLen(text) - cursorPos + 1                       ; how many chars sit after the caret
        }

        Sleep 80                                                            ; let the hidden GUI release focus
        Clip.FocusLastWindow()
        Win.WaitModifiersUp()

        if (g_CONFIG["ClipSendMode"] = 2) {                                 ; Mode 2: type it out, for apps that block clipboard paste
            SendInput("{Text}" text)
        } else {                                                            ; Mode 1 (default): clipboard + Ctrl+V, fast and safe for long text
            oldClip := ClipboardAll()
            A_Clipboard := ""
            A_Clipboard := text
            if !ClipWait(1) {
                A_Clipboard := oldClip
                g_LOG.Debug("PasteClipText: ClipWait timeout, paste aborted")
                return false
            }
            SendInput("^v")
            Sleep g_CONFIG["ClipPasteDelay"]
            A_Clipboard := oldClip                                          ; restore whatever the user had before
            oldClip := ""
        }

        if (caretBack > 0)
            SendInput("{Left " caretBack "}")

        g_LOG.Debug("PasteClipText: Pasted " StrLen(text) " chars, mode=" g_CONFIG["ClipSendMode"])
        return true
    }

    static EditClipText(*) {                                                ; Multi-line editor for the Clip body, opened by the "..." button
        Global g_CmdMgrGui, g_ClipEditGui

        g_ClipEditGui := Gui("+Owner" g_CmdMgrGui.Hwnd, "Clip Text  -  line breaks are stored as \n")
        g_ClipEditGui.SetFont("S10 Norm", "Consolas")
        clipEdit := g_ClipEditGui.AddEdit("w620 r18 +Multi +WantReturn +WantTab vClipBody", Clip.UnescapeClipText(g_CmdMgrGui["Path"].Text))
        g_ClipEditGui.SetFont("S9 Norm", "Microsoft Yahei")
        g_ClipEditGui.AddText("xm w440 cGray", "Placeholders: {date} {time} {datetime} {clipboard} {arg} {cursor}")
        g_ClipEditGui.AddButton("Default x+10 yp-6 w80", "OK").OnEvent("Click", SaveClipText)
        g_ClipEditGui.AddButton("x+8 yp w80", "Cancel").OnEvent("Click", (p*) => Clip.CloseClipEditor(p*))
        g_ClipEditGui.OnEvent("Close", (p*) => Clip.CloseClipEditor(p*))
        g_ClipEditGui.OnEvent("Escape", (p*) => Clip.CloseClipEditor(p*))

        g_CmdMgrGui.Opt("+Disabled")
        g_ClipEditGui.Show("Center")
        clipEdit.Focus()

        SaveClipText(*) {
            g_CmdMgrGui["Path"].Value := Clip.EscapeClipText(clipEdit.Value)
            Clip.CloseClipEditor()
        }
    }

    static CloseClipEditor(*) {
        Global g_CmdMgrGui, g_ClipEditGui
        try g_CmdMgrGui.Opt("-Disabled")
        try g_ClipEditGui.Destroy()
        try WinActivate("ahk_id " g_CmdMgrGui.Hwnd)
    }

    ;===========================================================================
    ; 一键文本转换: 直接在当前剪贴板内容上转换并写回, 配合 Ctrl+V 使用。
    ; 每个转换都是 Func | ClipXxx | ... 内置命令, 见 ALTRun.ahk 里同名的裸全局
    ; 函数外壳 (RunCommand 的 FUNC 类型只认裸函数名, 不认 Class.Method)。
    ;===========================================================================

    ; 转换前的公共检查 + 转换后的公共反馈, 每个具体转换只需要传一个处理函数。
    static _Transform(fn, label) {
        text := A_Clipboard
        if (text = "") {
            ToolTip(g_LNG[830])
            SetTimer(() => ToolTip(""), -1500)
            return
        }
        A_Clipboard := fn(text)
        ToolTip(label)
        SetTimer(() => ToolTip(""), -1500)
    }

    static ToUpper()     => Clip._Transform((s) => StrUpper(s), g_LNG[831])
    static ToLower()     => Clip._Transform((s) => StrLower(s), g_LNG[832])
    static ToTitleCase() => Clip._Transform((s) => StrTitle(s), g_LNG[833])
    static Reverse()     => Clip._Transform((s) => Clip._ReverseChars(s), g_LNG[834])
    static SortAsc()     => Clip._Transform((s) => Clip._Join(Clip._Lines(Sort(s))), g_LNG[835])
    static SortDesc()    => Clip._Transform((s) => Clip._Join(Clip._Lines(Sort(s, "R"))), g_LNG[836])
    static TrimLines()   => Clip._Transform((s) => Clip._TrimLines(s), g_LNG[837])
    static RemoveBlankLines() => Clip._Transform((s) => Clip._DropBlankLines(s), g_LNG[838])
    static DedupeLines() => Clip._Transform((s) => Clip._DedupeLines(s), g_LNG[839])

    static _ReverseChars(text) {
        out := ""
        Loop Parse text
            out := A_LoopField . out
        return out
    }

    ; 拆行/合并行的公共小工具 - 统一按 `n`/`r` 拆, 统一用 `r`n` 合并回去,
    ; 不管剪贴板原本是 LF 还是 CRLF, 结果都是规规矩矩的 Windows 换行。
    static _Lines(text) => StrSplit(text, "`n", "`r")

    static _Join(lines) {
        out := ""
        for _, line in lines
            out .= (out = "" ? "" : "`r`n") . line
        return out
    }

    ; 逐行去掉首尾空格/Tab, 保留空行本身(要连空行一起清掉用 RemoveBlankLines)。
    static _TrimLines(text) {
        trimmed := []
        for _, line in Clip._Lines(text)
            trimmed.Push(Trim(line, " `t"))
        return Clip._Join(trimmed)
    }

    static _DropBlankLines(text) {
        kept := []
        for _, line in Clip._Lines(text)
            if (Trim(line) != "")
                kept.Push(line)
        return Clip._Join(kept)
    }

    ; 按原顺序去重(不像 Sort 的 U 选项那样会顺带重新排序)。
    static _DedupeLines(text) {
        seen := Map(), kept := []
        for _, line in Clip._Lines(text) {
            if seen.Has(line)
                continue
            seen[line] := true
            kept.Push(line)
        }
        return Clip._Join(kept)
    }
}
