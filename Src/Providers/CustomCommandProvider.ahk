;===============================================================================
; CustomCommandProvider.ahk - 用户自定义命令 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; ALTRun.json -> CustomCommands, 每一条:
;   { "Title": "显示名称", "Type": "File|Folder|Command|Url",
;     "Target": "路径 / 程序 / 网址", "Arguments": "命令行参数", "Keyword": "可选关键字" }
; Target 可以用 A_Desktop / A_ScriptDir 等内置变量开头, 或 %AppData% 等环境变量。
; Keyword 完全相同时排在最前面。
; 搜索范围: 名称、关键字、名称的拼音首字母; File / Folder 类型还包括目标的文件名
; (不含扩展名) 或文件夹名, 例如 Target "Q:\Projects\PT1931 - 24 NIR" 输入 "nir" 也能找到。
; 同样的匹配程度, 名称匹配排在文件名 / 文件夹名匹配前面。
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
            commandType := command.Has("Type") ? command["Type"] : "File"
            score := FuzzyMatcher.BestKey(needle, CustomCommandProvider._KeysFor(command["Title"], keyword))
            targetScore := FuzzyMatcher.BestKey(needle, CustomCommandProvider._TargetKeysFor(commandType, command["Target"]))
            if (targetScore > 0)
                score := Max(score, targetScore - 5)                        ; 名称优先
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

    ; File / Folder 的目标文件名 (不含扩展名) 或文件夹名及其拼音首字母, 按 "类型|目标" 缓存
    static _TargetKeysFor(commandType, target) {
        static cache := Map()
        cacheKey := commandType "|" target
        if !cache.Has(cacheKey)
            cache[cacheKey] := CustomCommandProvider._TargetKeys(commandType, target)
        return cache[cacheKey]
    }

    static _TargetKeys(commandType, target) {
        if !(commandType = "Folder" || commandType = "File")
            return []
        SplitPath(RTrim(Path.Resolve(target), "\/"), &fileName, , , &nameNoExt)
        name := (commandType = "File") ? nameNoExt : fileName
        if (name = "" || InStr(name, ":"))                                  ; 驱动器根目录 "Q:" 不算名称
            return []
        pinyinText := Pinyin.Initials(name)
        return [FuzzyMatcher.Key(name), (pinyinText != name) ? FuzzyMatcher.Key(pinyinText) : ""]
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

    ; 检查命令的目标是否还在 (文件夹改名、文件移走后命令会失效), 返回:
    ;   "OK"          存在
    ;   "Missing"     找不到
    ;   "Unavailable" 所在的驱动器或网络位置现在无法访问 (例如网络盘没连上), 不能判断
    ;   "Skipped"     网址、shell: 等不是文件的目标, 不检查
    ; roots: 同一次检查共用的 Map, 记住每个驱动器 / 网络共享能否访问, 断开的网络盘只等一次
    static CheckTarget(command, roots := "") {
        commandType := command.Has("Type") ? command["Type"] : "File"
        if (commandType = "Url")
            return "Skipped"
        raw := Trim(command.Has("Target") ? command["Target"] : "", "`" `t")
        if (raw = "")
            return "Missing"
        target := Path.Resolve(raw)
        if RegExMatch(target, "i)^([a-z][a-z0-9+.-]+:|::\{)")               ; shell:、ms-settings:、http: 等 (不是 "C:")
            return "Skipped"
        if !InStr(target, "\")                                              ; 裸文件名, PATH 里没有
            return CustomCommandProvider._InAppPaths(target) ? "OK" : "Missing"
        SplitPath(target, , , , , &drive)
        if (drive != "") {
            if !IsObject(roots)
                roots := Map()
            if !roots.Has(drive)
                roots[drive] := DirExist(drive "\") != ""
            if !roots[drive]
                return "Unavailable"
        }
        found := (commandType = "Folder") ? DirExist(target) : FileExist(target)
        return (found != "") ? "OK" : "Missing"
    }

    ; "App Paths" 里注册的程序 (winword、chrome...), Run 也能直接打开
    static _InAppPaths(name) {
        SplitPath(name, , , &ext)
        if (ext = "")
            name .= ".exe"
        for root in ["HKCU", "HKLM"]
            try if (RegRead(root "\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" name, "") != "")
                return true
        return false
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
