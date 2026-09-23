;===============================================================================
; CustomCommandProvider.ahk - 用户自定义命令 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; ALTRun.json -> CustomCommands, 每一条:
;   { "Title": "显示名称", "Type": "File|Folder|Command|Url",
;     "Target": "路径 / 程序 / 网址", "Arguments": "命令行参数", "Keyword": "可选关键字" }
; Target 可以用 A_Desktop / A_ScriptDir 等内置变量开头, 或 %AppData% 等环境变量。
; Keyword 完全相同时排在最前面。
;
; 用法:
;   CustomCommandProvider.AddFromPath(path [, title])   "发送到" 菜单 / 操作面板添加
;===============================================================================

class CustomCommandProvider {
    static Id := "CustomCommands"

    static Init() {
    }

    static Search(query) {
        results := []
        for command in AppSettings.CustomCommands {
            if !(command is Map) || !command.Has("Title") || !command.Has("Target")
                continue
            keyword := command.Has("Keyword") ? command["Keyword"] : ""
            score := FuzzyMatcher.Best(query.Text, [command["Title"], keyword, Pinyin.Initials(command["Title"])])
            if (keyword != "" && query.Keyword = keyword)
                score := Max(score, 100)
            if (score <= 0)
                continue
            results.Push(CustomCommandProvider._ToItem(command, score + 15))  ; 用户自己加的命令比索引出来的应用优先
        }
        return results
    }

    static _ToItem(command, score) {
        target := command["Target"]
        arguments := command.Has("Arguments") ? command["Arguments"] : ""
        commandType := command.Has("Type") ? command["Type"] : "File"
        switch commandType, false {
            case "Folder": kind := "folder", icon := Path.Resolve(target)
            case "Url"   : kind := "url",    icon := "url:"
            default      : kind := "file",   icon := Path.Resolve(target)
        }
        displayTarget := (kind = "url") ? target : Path.Resolve(target)
        return ResultItem(command["Title"], Trim(displayTarget " " arguments), {
            Kind: kind, Arg: target, Arguments: arguments, Icon: icon, Score: score,
            Uid: "custom:" StrLower(commandType "|" target "|" arguments)
        })
    }

    static AddFromPath(target, title := "") {
        if (title = "") {
            SplitPath(target, &fileName, , , &nameNoExt)
            title := DirExist(target) ? fileName : nameNoExt
        }
        AppSettings.CustomCommands.Push(Map(
            "Title", title,
            "Type", DirExist(target) ? "Folder" : "File",
            "Target", target,
            "Arguments", "",
            "Keyword", ""
        ))
        AppSettings.Save()
        App.Notify(I18n.T("Custom.Added", title), 2500)
    }
}
