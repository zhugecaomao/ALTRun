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
;   CustomCommandProvider.AddFromPaths(paths)            "发送到" 菜单: 1 个弹出编辑对话框, 多个直接添加
;   CustomCommandProvider.Edit(command [, prefill])      打开编辑对话框; command 为 "" 时新建
;   搜索结果里 F3 / 右键 "编辑" 调用 EditItem(), Ctrl+Del 调用 DeleteItem()
;===============================================================================

class CustomCommandProvider {
    static Id := "CustomCommands"
    static _lastNeedle := "", _lastMatches := [], _lastList := ""

    static Init() {
    }

    ; 先给所有命令打分, 只为排在前面的 MaxResults 条生成结果 (命令多时, 输入一个字母就能
    ; 匹配上几百条, 每条都生成结果会让输入卡顿)。取前几条时算上学习加分, 常选的命令不会被挤掉。
    ; 继续输入 (新输入以上一次的输入开头, 至少 3 个字母) 时只在上一次匹配到的命令里找,
    ; 和 ApplicationProvider 一样; 命令有增删改时重新从全部命令里找
    static Search(query) {
        needle := StrLower(query.Text)
        commands := AppSettings.CustomCommands
        list := ObjPtr(commands) "|" commands.Length
        candidates := commands
        if (list = CustomCommandProvider._lastList && StrLen(CustomCommandProvider._lastNeedle) >= 3
            && InStr(needle, CustomCommandProvider._lastNeedle) = 1) {
            candidates := Map()
            for index in CustomCommandProvider._lastMatches
                candidates[index] := commands[index]
        }
        scores := Map(), ranks := Map(), matches := []
        for index, command in candidates {
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
            scores[index] := score + 15                                     ; 用户自己加的命令比索引出来的应用优先
            ranks[index] := scores[index] + Knowledge.Boost(query.Text, CustomCommandProvider._Uid(command))
            matches.Push(index)
        }
        CustomCommandProvider._lastNeedle := needle, CustomCommandProvider._lastMatches := matches, CustomCommandProvider._lastList := list
        results := []
        for index in FuzzyMatcher.TopIndexes(ranks, ProviderRegistry.MaxResults)
            results.Push(CustomCommandProvider._ToItem(commands[index], scores[index]))
        return results
    }

    ; 启动后空闲时预先算好每条命令的搜索 Key (名称 / 关键字 / 拼音 / 目标名) 和目标路径,
    ; 第一次输入时就不用等 (几百条命令第一次算要几百毫秒)。每次算 WarmBatch 条, 中间可以处理输入
    static WarmBatch := 40
    static Warm(start := 1) {
        commands := AppSettings.CustomCommands
        last := Min(commands.Length, start + CustomCommandProvider.WarmBatch - 1)
        Loop Max(0, last - start + 1) {
            command := commands[start + A_Index - 1]
            if !(command is Map) || !command.Has("Title") || !command.Has("Target")
                continue
            commandType := command.Has("Type") ? command["Type"] : "File"
            CustomCommandProvider._KeysFor(command["Title"], command.Has("Keyword") ? command["Keyword"] : "")
            CustomCommandProvider._TargetKeysFor(commandType, command["Target"])
            if (commandType = "Folder")
                IconCache.FolderIcon(CustomCommandProvider._Resolve(command["Target"]), true)
        }
        if (last < commands.Length)
            SetTimer(() => CustomCommandProvider.Warm(last + 1), -10)
    }

    ; 命令被修改后, 下一次搜索从全部命令里找
    static _ResetNarrowing() {
        CustomCommandProvider._lastNeedle := ""
    }

    static _Uid(command) {
        arguments := command.Has("Arguments") ? command["Arguments"] : ""
        commandType := command.Has("Type") ? command["Type"] : "File"
        return "custom:" StrLower(commandType "|" command["Target"] "|" arguments)
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
        SplitPath(RTrim(CustomCommandProvider._Resolve(target), "\/"), &fileName, , , &nameNoExt)
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
            case "Folder": kind := "folder", icon := CustomCommandProvider._FolderIcon(CustomCommandProvider._Resolve(target))
            case "Url"   : kind := "url",    icon := "url:"
            default      : kind := "file",   icon := CustomCommandProvider._Resolve(target)
        }
        displayTarget := (kind = "url") ? target : CustomCommandProvider._Resolve(target)
        return ResultItem(command["Title"], Trim(displayTarget " " arguments), {
            Kind: kind, Arg: target, Arguments: arguments, Icon: icon, Score: score, Source: command,
            Uid: CustomCommandProvider._Uid(command)
        })
    }

    ; 网络位置上的文件夹直接用通用的文件夹图标 (不读网络, 也不会因为名字里带点被当成文件)
    static _FolderIcon(folder) {
        return IconCache.FolderIcon(folder)
    }

    ; Path.Resolve 对只写程序名的目标 ("cmd.exe") 要查磁盘和 PATH, 结果缓存起来
    static _Resolve(target) {
        static cache := Map()
        if !cache.Has(target)
            cache[target] := Path.Resolve(target)
        return cache[target]
    }

    static EditorFields() {
        types := [["File", I18n.T("Prefs.Type.File")], ["Folder", I18n.T("Prefs.Type.Folder")]
                , ["Command", I18n.T("Prefs.Type.Command")], ["Url", I18n.T("Prefs.Type.Url")]]
        hint := (name) => I18n.T("Cmd.Field." name)                        ; 每个字段下面的灰色说明
        return [ItemEditor.Field("Title", "Prefs.Col.Title", "text", true, "", hint("Title"))
              , ItemEditor.Field("Type", "Prefs.Col.Type", "choice", false, types, hint("Type"))
              , ItemEditor.Field("Target", "Prefs.Col.Target", "file", true, "", hint("Target"))
              , ItemEditor.Field("Arguments", "Prefs.Col.Arguments", "text", false, "", hint("Arguments"))
              , ItemEditor.Field("Keyword", "Prefs.Col.Keyword", "text", false, "", hint("Keyword"))]
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
        CustomCommandProvider._ResetNarrowing()
        return AppSettings.Save()
    }

    static EditItem(item) {
        return CustomCommandProvider.Edit(item.Source)
    }

    static DeleteItem(item) {
        for index, command in AppSettings.CustomCommands {
            if (ObjPtr(command) = ObjPtr(item.Source)) {
                AppSettings.CustomCommands.RemoveAt(index)
                CustomCommandProvider._ResetNarrowing()
                return AppSettings.Save()
            }
        }
        return false
    }

    ; 资源管理器 "发送到 → ALTRun" (选中几个就一次传进来几个路径):
    ;   1 个: 弹出编辑对话框, 名称 / 类型 / 目标已填好, 确定后才添加; 已经有这条命令时打开它修改
    ;   多个: 全部直接添加 (一个个弹对话框太多), 已经有的跳过
    static AddFromPaths(paths) {
        if !paths.Length
            return
        if (paths.Length = 1) {
            existing := CustomCommandProvider.FindByTarget(paths[1])
            if IsObject(existing) {
                App.Notify(I18n.T("Custom.Exists", existing["Title"]), 2500)
                CustomCommandProvider.Edit(existing)
            } else if CustomCommandProvider.Edit("", CustomCommandProvider.FromPath(paths[1])) {
                App.Notify(I18n.T("Custom.Added", AppSettings.CustomCommands[-1]["Title"]), 2500)
            }
            return
        }
        added := 0, skipped := 0
        for target in paths {
            if IsObject(CustomCommandProvider.FindByTarget(target)) {
                skipped += 1
                continue
            }
            AppSettings.CustomCommands.Push(CustomCommandProvider.FromPath(target))
            added += 1
        }
        if added {
            CustomCommandProvider._ResetNarrowing()
            AppSettings.Save()
        }
        App.Notify(I18n.T("Custom.AddedMany", added) (skipped ? " " I18n.T("Custom.SkippedExisting", skipped) : ""), 3000)
    }

    ; 文件 / 文件夹路径 -> 一条新命令 (名称 = 文件名去掉扩展名, 或文件夹名)
    static FromPath(target) {
        target := RTrim(target, "\/")
        if RegExMatch(target, "^[A-Za-z]:$")                                ; 整个驱动器 "D:"
            target .= "\"
        isFolder := DirExist(target) != ""
        SplitPath(RTrim(target, "\"), &fileName, , , &nameNoExt)
        title := isFolder ? fileName : nameNoExt
        if (title = "")
            title := target
        return Map("Title", title, "Type", isFolder ? "Folder" : "File", "Target", target, "Arguments", "", "Keyword", "")
    }

    ; 目标指向同一个文件 / 文件夹的命令 (比较展开变量后的路径, 不分大小写), 没有返回 ""
    static FindByTarget(target) {
        wanted := StrLower(RTrim(Path.Resolve(target), "\/"))
        for command in AppSettings.CustomCommands {
            if !(command is Map) || !command.Has("Target")
                continue
            if (command.Has("Type") && command["Type"] = "Url")
                continue
            if (StrLower(RTrim(CustomCommandProvider._Resolve(command["Target"]), "\/")) = wanted)
                return command
        }
        return ""
    }
}
