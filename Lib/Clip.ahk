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
        g_ClipEditGui.AddButton("x+8 yp w80", "Cancel").OnEvent("Click", Clip.CloseClipEditor)
        g_ClipEditGui.OnEvent("Close", Clip.CloseClipEditor)
        g_ClipEditGui.OnEvent("Escape", Clip.CloseClipEditor)

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
}
