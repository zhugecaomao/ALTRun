;===============================================================================
; ScriptProvider.ahk - 脚本扩展: Scripts\ 文件夹里的脚本变成命令 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 Raycast 的 Script Commands、RunZ 的插件一样: 把脚本放进 Scripts\ (程序目录, 程序目录不能写入时
; 在 %APPDATA%\ALTRun\Scripts), 就能按名称或关键字搜到并运行。支持 .ahk .ps1 .bat .cmd .py。
; 脚本开头的注释里可以写 (注释符 ; # REM :: // 都行, 都可以不写):
;   @altrun.title     名称              (默认用文件名)
;   @altrun.keyword   关键字            输入完全相同时排在最前面
;   @altrun.argument  提示文字          需要参数: 输入 "关键字 文字", 文字作为第一个参数传给脚本
;   @altrun.mode      window | silent | output
;                     window (默认) 正常运行, 有窗口; silent 在后台运行, 结束后把输出的最后一行显示成通知;
;                     output 在后台运行, 结束后用记事本打开全部输出
; 文件夹有变化时 (修改时间) 重新读取, 不用重新载入 ALTRun。F3 用记事本编辑脚本。
;
; 设置 (ALTRun.json -> Features.Scripts): Enabled
;===============================================================================

class ScriptProvider {
    static Id := "Scripts"
    static Extensions := Map("ahk", 1, "ps1", 1, "bat", 1, "cmd", 1, "py", 1)
    static Scripts := []
    static _lastStamp := "", _checked := 0
    static TimeoutMs := 120000                                              ; 后台脚本最多等多久

    static Folder := ""                                                     ; 不为空时用这个文件夹 (测试用)
    static Dir => (ScriptProvider.Folder != "") ? ScriptProvider.Folder : AppSettings.Portable ? A_ScriptDir "\Scripts" : AppSettings.UserDir "\Scripts"

    static Init() {
        ScriptProvider.Load()
    }

    static Search(query) {
        if (A_TickCount - ScriptProvider._checked > 2000)                   ; 最多 2 秒看一次文件夹有没有变
            ScriptProvider._RefreshIfChanged()
        results := []
        needle := StrLower(query.Text)
        for script in ScriptProvider.Scripts {
            if (script.Keyword != "" && query.Keyword = script.Keyword && script.Argument != "" && query.HasRest) {
                item := ScriptProvider.ItemFor(script, 150, query.Rest)
                item.Exclusive := true
                return [item]
            }
            score := FuzzyMatcher.BestKey(needle, script.Keys)
            if (script.Keyword != "" && query.Keyword = script.Keyword)
                score := Max(score, 100)
            if (score > 0)
                results.Push(ScriptProvider.ItemFor(script, score + 10))
        }
        return results
    }

    static ItemFor(script, score, argument := "") {
        props := {Kind: "file", Arg: script.Path, Icon: ScriptProvider._Icon(script), Score: score, Source: script
            , Uid: "script:" StrLower(script.Path), OnRun: (*) => ScriptProvider.Run(script, argument)}
        title := script.Title, subtitle := I18n.T("Script.Subtitle", script.Name)
        if (script.Argument != "") {
            if (argument = "") {
                props.Valid := (script.Keyword = "")                        ; 有关键字: Enter 补全 "关键字 ", 接着输入参数
                props.AutoComplete := (script.Keyword != "") ? script.Keyword " " : ""
                subtitle := ((script.Keyword != "") ? I18n.T("Search.TypeArgAfter", script.Keyword, script.Argument) : script.Argument) " · " subtitle
            } else
                title .= ": " argument
        }
        return ResultItem(title, subtitle, props)
    }

    static _Icon(script) {
        switch script.Ext {
            case "ahk": return A_IsCompiled ? A_ScriptFullPath : A_AhkPath
            case "ps1": return A_WinDir "\System32\WindowsPowerShell\v1.0\powershell.exe"
            default:    return A_ComSpec
        }
    }

    ; 运行脚本的命令行
    static CommandLine(script, argument := "") {
        quoted := '"' script.Path '"' ((argument != "") ? ' "' StrReplace(argument, '"', '\"') '"' : "")
        switch script.Ext {
            case "ahk": return '"' A_AhkPath '"' (A_IsCompiled ? " /script " : " ") quoted
            case "ps1": return 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File ' quoted
            case "py":  return 'py ' quoted
            default:    return A_ComSpec ' /c "' quoted '"'
        }
    }

    static Run(script, argument := "") {
        SplitPath(script.Path, , &dir)
        command := ScriptProvider.CommandLine(script, argument)
        if (script.Mode = "window")
            return Run(command, dir)
        output := A_Temp "\ALTRun-script-" A_TickCount ".txt"              ; 后台运行, 输出写到临时文件
        Run(A_ComSpec ' /c "' command ' > "' output '" 2>&1"', dir, "Hide", &pid)
        started := A_TickCount
        check() {
            if (ProcessExist(pid) && A_TickCount - started < ScriptProvider.TimeoutMs)
                return
            SetTimer(check, 0)
            ScriptProvider._Finished(script, output)
        }
        SetTimer(check, 300)
    }

    static _Finished(script, output) {
        text := ""
        try text := FileRead(output)
        if (script.Mode = "output") {
            Run('notepad.exe "' output '"')                                 ; 记事本打开后由用户关闭, 临时文件留在 %Temp%
            return
        }
        try FileDelete(output)
        last := ""
        Loop Parse, text, "`n", "`r"
            if (Trim(A_LoopField) != "")
                last := Trim(A_LoopField)
        message := (last != "") ? last : I18n.T("Script.Done", script.Title)
        App.Notify(message, 3000)
        return message
    }

    ; F3: 用记事本编辑脚本
    static EditItem(item) {
        Run('notepad.exe "' item.Source.Path '"')
        return false
    }

    static Load() {
        scripts := []
        if DirExist(ScriptProvider.Dir) {
            Loop Files, ScriptProvider.Dir "\*.*", "F" {
                ext := StrLower(A_LoopFileExt)
                if !ScriptProvider.Extensions.Has(ext)
                    continue
                if IsObject(script := ScriptProvider.ReadScript(A_LoopFileFullPath))
                    scripts.Push(script)
            }
        }
        ScriptProvider.Scripts := scripts
        ScriptProvider._lastStamp := ScriptProvider._Stamp()
        ScriptProvider._checked := A_TickCount
    }

    ; 脚本文件 -> {Path, Name, Ext, Title, Keyword, Argument, Mode, Keys}; 读开头 30 行里的 @altrun.xxx
    static ReadScript(path) {
        SplitPath(path, &name, , &ext, &baseName)
        meta := Map("title", baseName, "keyword", "", "argument", "", "mode", "window")
        try {
            Loop Read, path {
                if (A_Index > 30)
                    break
                if RegExMatch(A_LoopReadLine, "i)^\s*(?:;|#|//|::|rem\b)\s*@altrun\.(\w+)\s+(.*?)\s*$", &m)
                    meta[StrLower(m[1])] := m[2]
            }
        } catch
            return ""
        mode := StrLower(meta["mode"])
        if !(mode = "silent" || mode = "output")
            mode := "window"
        title := meta["title"], keyword := StrLower(Trim(meta["keyword"]))
        pinyinText := Pinyin.Initials(title)
        return {Path: path, Name: name, Ext: StrLower(ext), Title: title, Keyword: keyword, Argument: Trim(meta["argument"]), Mode: mode
              , Keys: [FuzzyMatcher.Key(title), (keyword != "") ? FuzzyMatcher.Key(keyword) : "", (pinyinText != title) ? FuzzyMatcher.Key(pinyinText) : ""]}
    }

    ; 文件夹和里面每个文件的修改时间拼起来, 有变化就重新读
    static _Stamp() {
        stamp := ""
        if DirExist(ScriptProvider.Dir)
            Loop Files, ScriptProvider.Dir "\*.*", "F"
                stamp .= A_LoopFileName A_LoopFileTimeModified "|"
        return stamp
    }

    static _RefreshIfChanged() {
        ScriptProvider._checked := A_TickCount
        if (ScriptProvider._Stamp() != ScriptProvider._lastStamp)
            ScriptProvider.Load()
    }

    ; 示例脚本, 各演示一种写法: 后台运行显示通知 (ahk, silent)、关键字后面带参数 (bat, window; 关键字不要和默认命令重复)、
    ; 用记事本看全部输出 (ps1, output)。只用 ASCII 字符: Windows PowerShell 5 按 ANSI 读不带 BOM 的文件
    static ExampleScripts() {
        return Map(
            "Today.ahk",
                "; @altrun.title  Example: Today's Date`r`n"
              . "; @altrun.mode   silent`r`n"
              . "; Runs in the background; the last line it prints is shown as a notification.`r`n"
              . "; Change the mode to output to read everything it prints in Notepad, or remove it to run normally.`r`n"
              . 'FileAppend(FormatTime(, "dddd, d MMMM yyyy") ", week " SubStr(FormatTime(, "YWeek"), 5), "*")' "`r`n",
            "Port.bat",
                "@echo off`r`n"
              . "rem @altrun.title     Example: Who Uses a Port`r`n"
              . "rem @altrun.keyword   port`r`n"
              . "rem @altrun.argument  a port number`r`n"
              . 'rem Type "port 8080" in ALTRun: the text after the keyword is passed to the script as %1.' "`r`n"
              . "rem No mode line: it runs normally in its own window.`r`n"
              . 'netstat -ano | findstr /c:":%~1 "' "`r`n"
              . "pause`r`n",
            "IP Addresses.ps1",
                "# @altrun.title  Example: IP Addresses`r`n"
              . "# @altrun.mode   output`r`n"
              . "# Runs in the background, then opens everything it prints in Notepad.`r`n"
              . "Get-NetIPAddress -AddressFamily IPv4 | Where-Object IPAddress -ne '127.0.0.1' | Format-Table InterfaceAlias, IPAddress -AutoSize`r`n")
    }

    ; 打开 Scripts 文件夹; 还没有时先建几个示例脚本 (ExampleScripts)
    static OpenFolder() {
        dir := ScriptProvider.Dir
        if !DirExist(dir) {
            DirCreate(dir)
            for name, text in ScriptProvider.ExampleScripts()
                FileAppend(text, dir "\" name, "UTF-8-RAW")                 ; 不带 BOM: cmd 读到 BOM 会把第一行当成命令报错
        }
        ActionCatalog.OpenFolder(dir)                                       ; 设置的文件管理器 (例如 Total Commander)
    }
}
