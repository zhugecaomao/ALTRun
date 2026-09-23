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
;   CustomCommandProvider.AddFromPath(path [, title])   "发送到" 菜单添加 (不弹对话框)
;   CustomCommandProvider.Edit(command [, prefill])      打开编辑对话框; command 为 "" 时新建
;   搜索结果里 F3 / 右键 "编辑" 调用 EditItem(), Ctrl+Del 调用 DeleteItem()
;===============================================================================

class CustomCommandProvider {
    static Id := "CustomCommands"

    static Init() {
    }

    static Search(query) {
        results := []
        needle := StrLower(query.Text)
        for command in AppSettings.CustomCommands {
            if !(command is Map) || !command.Has("Title") || !command.Has("Target")
                continue
            keyword := command.Has("Keyword") ? command["Keyword"] : ""
            score := FuzzyMatcher.BestKey(needle, CustomCommandProvider._KeysFor(command["Title"], keyword))
            if (keyword != "" && query.Keyword = keyword)
                score := Max(score, 100)
            if (score <= 0)
                continue
            results.Push(CustomCommandProvider._ToItem(command, score + 15))  ; 用户自己加的命令比索引出来的应用优先
        }
        return results
    }

    ; 名称 / 关键字 / 拼音首字母的搜索 Key, 按 "名称|关键字" 缓存 (修改命令后自然换成新的键)
    static _KeysFor(title, keyword) {
        static cache := Map()
        cacheKey := title "|" keyword
        if !cache.Has(cacheKey) {
            pinyinText := Pinyin.Initials(title)
            cache[cacheKey] := [FuzzyMatcher.Key(title), (keyword != "") ? FuzzyMatcher.Key(keyword) : "", (pinyinText != title) ? FuzzyMatcher.Key(pinyinText) : ""]
        }
        return cache[cacheKey]
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
            Kind: kind, Arg: target, Arguments: arguments, Icon: icon, Score: score, Source: command,
            Uid: "custom:" StrLower(commandType "|" target "|" arguments)
        })
    }

    static EditorFields() {
        types := [["File", I18n.T("Prefs.Type.File")], ["Folder", I18n.T("Prefs.Type.Folder")]
                , ["Command", I18n.T("Prefs.Type.Command")], ["Url", I18n.T("Prefs.Type.Url")]]
        return [ItemEditor.Field("Title", "Prefs.Col.Title", "text", true)
              , ItemEditor.Field("Type", "Prefs.Col.Type", "choice", false, types)
              , ItemEditor.Field("Target", "Prefs.Col.Target", "file", true)
              , ItemEditor.Field("Arguments", "Prefs.Col.Arguments")
              , ItemEditor.Field("Keyword", "Prefs.Col.Keyword")]
    }

    static NewCommand() {
        return Map("Title", "", "Type", "File", "Target", "", "Arguments", "", "Keyword", "")
    }

    ; command: AppSettings.CustomCommands 里的一条 (就地修改), 或 "" 新建 (prefill 预先填好的字段)
    static Edit(command := "", prefill := "") {
        isNew := !IsObject(command)
        base := isNew ? CustomCommandProvider.NewCommand() : ItemEditor.WithDefaults(command, CustomCommandProvider.NewCommand())
        if IsObject(prefill)
            for key, value in prefill
                base[key] := value
        edited := ItemEditor.Edit(ItemEditor.Owner(), I18n.T("Prefs.Page.Commands"), CustomCommandProvider.EditorFields(), base)
        if !IsObject(edited)
            return false
        if isNew
            AppSettings.CustomCommands.Push(edited)
        else
            for key, value in edited
                command[key] := value
        return AppSettings.Save()
    }

    static EditItem(item) {
        return CustomCommandProvider.Edit(item.Source)
    }

    static DeleteItem(item) {
        for index, command in AppSettings.CustomCommands {
            if (ObjPtr(command) = ObjPtr(item.Source)) {
                AppSettings.CustomCommands.RemoveAt(index)
                return AppSettings.Save()
            }
        }
        return false
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
