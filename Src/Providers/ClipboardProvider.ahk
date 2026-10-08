;===============================================================================
; ClipboardProvider.ahk - 剪贴板历史 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 和 Alfred 的 Clipboard History 一样: 记住复制过的文字、文件和图片, 输入 "clip" (Keyword)
; 或按热键 (默认 Ctrl+Alt+C) 列出, "clip 关键词" 过滤, Enter 粘贴到前台窗口
; (文件粘贴回去还是文件, 图片还是图片)。
;
; 隐私:
;   - 密码管理器等程序复制的内容 (带 ExcludeClipboardContentFromMonitorProcessing
;     或 CanIncludeInClipboardHistory = 0 标记) 不会被记录
;   - IgnoreApps 里的程序 (进程名) 复制的内容不会被记录
;   - Persist = 0 时只保存在内存里, 退出即清空 (这时不记录图片)
;
; 保存: ClipboardHistory.json; 超过 LargeText 个字的条目单独存成 Clipboard\*.txt,
; 图片存成 Clipboard\img-*.png, JSON 里只记文件名。每次复制都会保存, 这样不用每次都
; 把很长的文字重新写一遍 (最多 200 条 x 10 万字)。
;
; 置顶 (操作面板里 "置顶"): 条目带 "Pinned": 1, 列在最前面, 超过条数时不删, "清空" 时也保留。
;
; 条目: Map("Text", "Time", "App") 是文字; 另外两种带 "Type":
;   "files"  复制的文件: "Files" [路径...], "Text" 是每行一个路径 (用来搜索)
;   "image"  图片: "Image" 文件名, "Width" / "Height"
;
; 设置 (ALTRun.json -> Features.Clipboard):
;   Enabled / Keyword / Hotkey / MaxItems / MaxItemLength / Persist / IgnoreApps
;   Images / MaxImages     记录图片、最多保存几张
;   MergeDoubleCopy        快速按两次 Ctrl+C: 把这次复制的文字接到上一条后面 (默认关闭)
;
; 用法:
;   ClipboardProvider.PauseRecording(ms)   ALTRun 自己临时借用剪贴板时调用, 这段时间不记录
;   ClipboardProvider.Add(text, source)    手动加入一条 (source = 来源程序的进程名)
;   ClipboardProvider.AddFiles(paths, source) / AddImage(source)
;   ClipboardProvider.TextAt(n)            第 n 条文字 (1 = 最新的), 片段的 {clipboard:N} 用
;===============================================================================

class ClipboardProvider {
    static Id      := "Clipboard"
    static File    := AppSettings.DataDir "\ClipboardHistory.json"
    static Folder  := AppSettings.DataDir "\Clipboard"                         ; 很长的条目和图片
    static LocalDir := EnvGet("LOCALAPPDATA") != "" ? EnvGet("LOCALAPPDATA") "\ALTRun" : ""   ; LocalFiles = 1 时 (默认) 图片和长条目存在这里
    static LargeText := 4000
    static Icon => IconCache.Own("Clipboard", "res:imageres.dll,-5314")   ; 文字条目的图标
    static MergeWindow := 400                                               ; 两次 Ctrl+C 最多隔多少毫秒算 "连按"
    static Entries := []              ; 最新的在前, 见文件开头的说明
    static _pausedUntil := 0, _saveTimer := "", _mergeUntil := 0

    static Init() {
        options := AppSettings.Feature("Clipboard")
        if (options["LocalFiles"] && ClipboardProvider.LocalDir != "")
            ClipboardProvider.UseLocalStorage(ClipboardProvider.LocalDir)
        if options["Persist"]
            ClipboardProvider._Load()
        OnClipboardChange((dataType) => ClipboardProvider._OnChange(dataType))
        if options["MergeDoubleCopy"] {
            try Hotkey("~^c", (*) => ClipboardProvider._OnCopyKey())         ; ~ : Ctrl+C 照常复制
            catch as e
                Logger.Error("ClipboardProvider: cannot register Ctrl+C - " e.Message)
        }
    }

    ; 图片和很长的条目 (Clipboard 文件夹) 放在本机, 不跟着 Data 文件夹进 OneDrive 等同步盘: 图片大, 放在同步盘里
    ; 每复制一张都要上传。ClipboardHistory.json 仍在 Data 里 (文字历史照常同步); 另一台电脑上没有对应文件的
    ; 图片 / 长条目读取时跳过。第一次切换时把 Data 里原来的文件夹搬过来; 搬不动时这次仍用原来的位置, 下次再试
    static UseLocalStorage(dir) {
        newFolder := dir "\Clipboard"
        if (newFolder = ClipboardProvider.Folder)
            return true
        try {
            if (InStr(FileExist(ClipboardProvider.Folder), "D") && !InStr(FileExist(newFolder), "D")) {
                DirCreate(dir)
                DirCopy(ClipboardProvider.Folder, newFolder)
                DirDelete(ClipboardProvider.Folder, true)
                Logger.Debug("ClipboardProvider: images and long entries moved to " newFolder)
            }
        } catch as e {
            Logger.Error("ClipboardProvider: cannot move " ClipboardProvider.Folder " to " newFolder " - " e.Message)
            return false
        }
        ClipboardProvider.Folder := newFolder
        return true
    }

    static PauseRecording(milliseconds) {
        ClipboardProvider._pausedUntil := A_TickCount + milliseconds
    }

    static Search(query) {
        options := AppSettings.Feature("Clipboard")
        if !query.MatchKeyword([options["Keyword"]], &term)
            return []
        icon := ClipboardProvider.Icon
        exclusive := query.HasRest                                          ; "clip " 之后只显示剪贴板历史
        if !ClipboardProvider.Entries.Length
            return [ResultItem(I18n.T("Clipboard.Empty"), I18n.T("Clipboard.EmptyHint", Win.HotkeyLabel(options["Hotkey"])), {Icon: icon, Valid: false, Score: 300, Exclusive: exclusive})]

        results := []
        tokens := StrSplit(Trim(term), " ")
        for pass in [true, false] {                                         ; 置顶的在前, 其它按时间
            for index, entry in ClipboardProvider.Entries {
                if (ClipboardProvider.IsPinned(entry) != pass || !ClipboardProvider._Matches(ClipboardProvider._SearchText(entry), tokens))
                    continue
                item := ClipboardProvider._ToItem(entry, icon, 300 - results.Length * 0.001)
                item.Exclusive := exclusive
                results.Push(item)
                if (results.Length >= ProviderRegistry.MaxResults - 1)
                    break 2
            }
        }
        if (term = "")
            results.Push(ResultItem(I18n.T("Clipboard.Clear"), I18n.T("Clipboard.ClearHint", ClipboardProvider.Entries.Length), {
                Icon: "res:shell32.dll,-32", Score: 0, Exclusive: exclusive, OnRun: (*) => ClipboardProvider.Clear()
            }))
        return results
    }

    static _Matches(text, tokens) {
        for token in tokens
            if (token != "" && !InStr(text, token))
                return false
        return true
    }

    static TypeOf(entry) => entry.Has("Type") ? entry["Type"] : "text"

    static IsPinned(entry) => (entry.Has("Pinned") && entry["Pinned"]) ? true : false

    static SetPinned(entry, pinned) {
        if pinned
            entry["Pinned"] := 1
        else if entry.Has("Pinned")
            entry.Delete("Pinned")
        ClipboardProvider._SaveLater()
        App.Notify(I18n.T(pinned ? "Clipboard.PinnedDone" : "Clipboard.UnpinnedDone"))
    }

    static _SearchText(entry) {
        if (ClipboardProvider.TypeOf(entry) = "image")                      ; 图片: 按 "图片" / "image" 和尺寸找
            return ClipboardProvider._ImageTitle(entry) " image png"
        return entry["Text"]
    }

    static _ToItem(entry, icon, score) {
        when := ClipboardProvider._FormatTime(entry["Time"])
        source := (entry.Has("App") && entry["App"] != "") ? entry["App"] : "?"
        pinned := ClipboardProvider.IsPinned(entry)
        props := {Score: score, Source: entry, OnRun: (item) => ClipboardProvider.PasteEntry(item.Source), Actions: []}
        switch ClipboardProvider.TypeOf(entry) {
            case "files":
                files := entry["Files"]
                if (files.Length = 1) {                                     ; 一个文件: 和文件搜索的结果一样, 有打开 / 显示位置等操作
                    filePath := files[1], isFolder := InStr(FileExist(filePath), "D") ? true : false
                    props.Kind := isFolder ? "folder" : "file", props.Arg := filePath
                    props.Icon := isFolder ? IconCache.FolderIcon(filePath) : filePath
                    title := Path.Leaf(RTrim(filePath, "\"))
                    subtitle := I18n.T("Clipboard.FileSubtitle", when, source, filePath)
                } else {
                    props.Kind := "text", props.Arg := entry["Text"], props.Icon := "folder:"
                    props.Actions.Push(ResultItem(I18n.T("Action.AddAllCommands"), "", {Icon: "res:imageres.dll,-2", OnRun: (*) => CustomCommandProvider.AddFromPaths(files)}))
                    title := ClipboardProvider._Names(files)
                    subtitle := I18n.T("Clipboard.FilesSubtitle", when, source, files.Length)
                }
                props.RunTitle := I18n.T("Clipboard.PasteFiles")
            case "image":
                imageFile := ClipboardProvider.Folder "\" entry["Image"]
                props.Kind := "file", props.Arg := imageFile, props.Icon := "thumb:" imageFile
                props.RunTitle := I18n.T("Clipboard.PasteImage")
                title := ClipboardProvider._ImageTitle(entry)
                subtitle := I18n.T("Clipboard.ImageSubtitle", when, source)
            default:
                text := entry["Text"]
                props.Kind := "text", props.Arg := text, props.Icon := icon, props.LargeText := text
                title := ClipboardProvider._Preview(text)
                subtitle := I18n.T("Clipboard.Subtitle", when, source, StrLen(text))
        }
        props.Actions.Push(ResultItem(I18n.T(pinned ? "Clipboard.Unpin" : "Clipboard.Pin"), "", {Icon: IconCache.Own("Pinned", "res:imageres.dll,-5303")
            , OnRun: (*) => ClipboardProvider.SetPinned(entry, !pinned)}))
        props.Pinned := pinned                                              ; 置顶的: 图标上画一个图钉 (SearchWindow)
        return ResultItem(title, subtitle, props)
    }

    static _ImageTitle(entry) => I18n.T("Clipboard.Image", entry.Has("Width") ? entry["Width"] : "?", entry.Has("Height") ? entry["Height"] : "?")

    static _Names(files) {
        names := ""
        for filePath in files {
            names .= (names = "" ? "" : ", ") Path.Leaf(RTrim(filePath, "\"))
            if (StrLen(names) > 100)
                return SubStr(names, 1, 100) "..."
        }
        return names
    }

    ; F3 / 右键 "编辑": 文字保存为片段, 文件添加到自定义命令; 图片不能编辑
    static CanEditItem(item) => ClipboardProvider.TypeOf(item.Source) != "image"

    static EditItem(item) {
        switch ClipboardProvider.TypeOf(item.Source) {
            case "files": return CustomCommandProvider.AddFromPaths(item.Source["Files"])
            case "image": return false
        }
        return SnippetProvider.Edit("", Map("Name", SubStr(ClipboardProvider._Preview(item.Arg), 1, 40), "Text", item.Arg))
    }

    static DeleteItem(item) {
        ClipboardProvider.RemoveEntry(item.Source)
        return true
    }

    ; 粘贴一条历史到前台窗口, 并把它移到最前面
    static Paste(text) {
        ClipboardProvider.Add(text, ClipboardProvider._EntryApp(text))
        ActionCatalog.PasteText(text)
    }

    static PasteEntry(entry) {
        switch ClipboardProvider.TypeOf(entry) {
            case "files":
                ClipboardProvider._MoveToTop(entry)
                files := entry["Files"]
                return ClipboardProvider._PasteWith(() => ClipboardData.SetFiles(files))
            case "image":
                ClipboardProvider._MoveToTop(entry)
                imageFile := ClipboardProvider.Folder "\" entry["Image"]
                return ClipboardProvider._PasteWith(() => ClipboardData.SetImage(imageFile))
        }
        return ClipboardProvider.Paste(entry["Text"])
    }

    ; 临时把文件 / 图片放进剪贴板, Ctrl+V, 再还原 (和 ActionCatalog.PasteText 一样, 这段时间不记录)
    static _PasteWith(setter) {
        delay := Max(AppSettings.Feature("Snippets")["PasteDelay"], 500)    ; 资源管理器粘贴文件时读剪贴板比较慢
        App.FocusPreviousWindow()
        Win.WaitModifiersUp()
        ClipboardProvider.PauseRecording(delay + 1000)
        saved := ClipboardAll()
        if !setter() {
            A_Clipboard := saved
            return false
        }
        SendInput("^v")
        Sleep(delay)
        A_Clipboard := saved
        return true
    }

    ;---------------------------------------------------------------------------
    ; History
    ;---------------------------------------------------------------------------
    static Add(text, source := "") {
        options := AppSettings.Feature("Clipboard")
        if (Trim(text, " `t`r`n") = "" || StrLen(text) > options["MaxItemLength"])
            return false
        entry := Map("Text", text, "Time", A_Now, "App", source)
        if IsObject(previous := ClipboardProvider._RemoveSame("text", text)) && ClipboardProvider.IsPinned(previous)
            entry["Pinned"] := 1                                            ; 再次复制 / 粘贴置顶的条目: 仍然置顶
        ClipboardProvider.Entries.InsertAt(1, entry)
        ClipboardProvider._Trim()
        ClipboardProvider._SaveLater()
        return true
    }

    static AddFiles(files, source := "") {
        if !files.Length
            return false
        joined := ""
        for filePath in files
            joined .= (joined = "" ? "" : "`r`n") filePath
        entry := Map("Type", "files", "Files", files, "Text", joined, "Time", A_Now, "App", source)
        if IsObject(previous := ClipboardProvider._RemoveSame("files", joined)) && ClipboardProvider.IsPinned(previous)
            entry["Pinned"] := 1
        ClipboardProvider.Entries.InsertAt(1, entry)
        ClipboardProvider._Trim()
        ClipboardProvider._SaveLater()
        return true
    }

    ; 把剪贴板里的图片存成 PNG 加入历史; 和已有的某张一模一样时只把那张移到最前面
    static AddImage(source := "") {
        try DirCreate(ClipboardProvider.Folder)
        name := "img-" A_Now "-" Random(100000, 999999) ".png"
        imageFile := ClipboardProvider.Folder "\" name
        size := ClipboardData.SaveImage(imageFile)
        if !IsObject(size)
            return false
        return ClipboardProvider.AddImageFile(name, size.Width, size.Height, source)
    }

    ; Data\Clipboard 里已经存好的 PNG 加入历史 (AddImage 和测试用)
    static AddImageFile(name, width, height, source := "") {
        imageFile := ClipboardProvider.Folder "\" name
        for entry in ClipboardProvider.Entries {
            if (ClipboardProvider.TypeOf(entry) = "image" && entry["Image"] != name && entry["Width"] = width && entry["Height"] = height
                && ClipboardProvider._SameFile(ClipboardProvider.Folder "\" entry["Image"], imageFile)) {
                try FileDelete(imageFile)
                ClipboardProvider._MoveToTop(entry)
                ClipboardProvider._SaveLater()
                return true
            }
        }
        ClipboardProvider.Entries.InsertAt(1, Map("Type", "image", "Image", name, "Width", width, "Height", height, "Text", "", "Time", A_Now, "App", source))
        ClipboardProvider._Trim()
        ClipboardProvider._SaveLater()
        return true
    }

    static Remove(text) {
        ClipboardProvider._RemoveSame("text", text)
        ClipboardProvider._SaveLater()
    }

    static RemoveEntry(target) {
        for index, entry in ClipboardProvider.Entries {
            if (entry = target) {
                ClipboardProvider.Entries.RemoveAt(index)
                break
            }
        }
        ClipboardProvider._SaveLater()
    }

    static Clear() {                                                        ; 置顶的条目保留
        kept := []
        for entry in ClipboardProvider.Entries
            if ClipboardProvider.IsPinned(entry)
                kept.Push(entry)
        ClipboardProvider.Entries := kept
        ClipboardProvider._SaveLater()
        App.Notify(I18n.T("Clipboard.Cleared"))
    }

    ; 第 n 条文字 (1 = 最新的), 跳过文件和图片; 没有时返回 ""
    static TextAt(n) {
        for entry in ClipboardProvider.Entries
            if (ClipboardProvider.TypeOf(entry) = "text" && --n = 0)
                return entry["Text"]
        return ""
    }

    ; 删掉同样内容的条目, 返回删掉的条目 (没有时 "")
    static _RemoveSame(type, text) {
        for index, entry in ClipboardProvider.Entries {
            if (ClipboardProvider.TypeOf(entry) = type && entry["Text"] == text)
                return ClipboardProvider.Entries.RemoveAt(index)
        }
        return ""
    }

    static _MoveToTop(target) {
        for index, entry in ClipboardProvider.Entries {
            if (entry = target) {
                ClipboardProvider.Entries.RemoveAt(index)
                break
            }
        }
        target["Time"] := A_Now
        ClipboardProvider.Entries.InsertAt(1, target)
        ClipboardProvider._SaveLater()
    }

    ; 超过条数时删掉最早的 (置顶的不删); 图片另有上限 (MaxImages); 删掉的图片文件在保存时清理
    static _Trim() {
        options := AppSettings.Feature("Clipboard"), entries := ClipboardProvider.Entries
        index := entries.Length
        while (entries.Length > options["MaxItems"] && index >= 1) {
            if !ClipboardProvider.IsPinned(entries[index])
                entries.RemoveAt(index)
            index--
        }
        images := 0, index := 1
        while (index <= entries.Length) {
            if (ClipboardProvider.TypeOf(entries[index]) = "image" && ++images > options["MaxImages"] && !ClipboardProvider.IsPinned(entries[index])) {
                entries.RemoveAt(index)
                continue
            }
            index++
        }
    }

    static _SameFile(a, b) {
        try {
            if (FileGetSize(a) != FileGetSize(b))
                return false
            first := FileRead(a, "RAW"), second := FileRead(b, "RAW")
            return DllCall("ntdll\RtlCompareMemory", "Ptr", first, "Ptr", second, "UPtr", first.Size, "UPtr") = first.Size
        }
        return false
    }

    static _EntryApp(text) {
        for entry in ClipboardProvider.Entries
            if (ClipboardProvider.TypeOf(entry) = "text" && entry["Text"] == text)
                return entry["App"]
        return ""
    }

    ;---------------------------------------------------------------------------
    ; Recording
    ;---------------------------------------------------------------------------
    static _OnChange(dataType) {
        if (dataType = 0 || A_TickCount < ClipboardProvider._pausedUntil)  ; 0 = 剪贴板被清空
            return
        if ClipboardProvider._IsPrivate()
            return
        processName := ""
        try processName := WinGetProcessName("A")
        options := AppSettings.Feature("Clipboard")
        for ignored in options["IgnoreApps"]
            if (processName = ignored)
                return
        if (dataType = 1) {                                                 ; 1 = 文字 (复制文件时也是, 内容是路径)
            if ClipboardData.HasFiles()
                return ClipboardProvider.AddFiles(ClipboardData.Files(), processName)
            text := ""
            try text := A_Clipboard
            if (A_TickCount < ClipboardProvider._mergeUntil && ClipboardProvider._Merge(text, processName))
                return
            return ClipboardProvider.Add(text, processName)
        }
        if (options["Images"] && options["Persist"] && ClipboardData.HasImage())   ; 2 = 其他格式; 图片只在保存历史时记录
            ClipboardProvider.AddImage(processName)
    }

    ; 连按两次 Ctrl+C: 第一次复制的文字已经是最新的一条, 第二次复制到同样的文字时把它接到前一条后面,
    ; 剪贴板里也换成合并后的文字
    static _OnCopyKey() {
        if (A_PriorHotkey = A_ThisHotkey && A_TimeSincePriorHotkey < ClipboardProvider.MergeWindow)
            ClipboardProvider._mergeUntil := A_TickCount + 1500
    }

    static _Merge(text, source := "") {
        entries := ClipboardProvider.Entries
        if (entries.Length < 2 || ClipboardProvider.TypeOf(entries[1]) != "text" || !(entries[1]["Text"] == text)
            || ClipboardProvider.TypeOf(entries[2]) != "text")
            return false
        merged := RTrim(entries[2]["Text"], "`r`n") "`r`n" text
        if (StrLen(merged) > AppSettings.Feature("Clipboard")["MaxItemLength"])
            return false
        ClipboardProvider._mergeUntil := 0
        entries.RemoveAt(1, 2)
        ClipboardProvider.Add(merged, source)
        ClipboardProvider.PauseRecording(500)
        A_Clipboard := merged
        App.Notify(I18n.T("Clipboard.Merged"))
        return true
    }

    ; 密码管理器等会给剪贴板内容加上 "不要记录" 的标记 (Windows 剪贴板历史也遵守)
    static _IsPrivate() {
        static excludeFormat := DllCall("RegisterClipboardFormat", "Str", "ExcludeClipboardContentFromMonitorProcessing", "UInt")
        static historyFormat := DllCall("RegisterClipboardFormat", "Str", "CanIncludeInClipboardHistory", "UInt")
        if DllCall("IsClipboardFormatAvailable", "UInt", excludeFormat)
            return true
        if !DllCall("IsClipboardFormatAvailable", "UInt", historyFormat)
            return false
        allowed := 1
        if DllCall("OpenClipboard", "Ptr", A_ScriptHwnd) {
            if (handle := DllCall("GetClipboardData", "UInt", historyFormat, "Ptr")) {
                if (pointer := DllCall("GlobalLock", "Ptr", handle, "Ptr")) {
                    allowed := NumGet(pointer, "UInt")
                    DllCall("GlobalUnlock", "Ptr", handle)
                }
            }
            DllCall("CloseClipboard")
        }
        return !allowed
    }

    ;---------------------------------------------------------------------------
    ; Formatting
    ;---------------------------------------------------------------------------
    static _Preview(text) {
        text := Trim(RegExReplace(text, "\s+", " "))
        return (StrLen(text) > 120) ? SubStr(text, 1, 120) "..." : text
    }

    static _FormatTime(timestamp) {
        if (SubStr(timestamp, 1, 8) = FormatTime(, "yyyyMMdd"))
            return FormatTime(timestamp, "HH:mm")
        return FormatTime(timestamp, "yyyy-MM-dd HH:mm")
    }

    ;---------------------------------------------------------------------------
    ; Persistence
    ;---------------------------------------------------------------------------
    static _Load() {
        if !FileExist(ClipboardProvider.File)
            return
        try {
            data := JSON.Parse(FileRead(ClipboardProvider.File, "UTF-8"))
            if !(data is Map && data.Has("Entries") && data["Entries"] is Array)
                return
            entries := []
            for entry in data["Entries"] {
                if !(entry is Map)
                    continue
                switch ClipboardProvider.TypeOf(entry) {
                    case "files":
                        if !(entry.Has("Files") && entry["Files"] is Array && entry["Files"].Length)
                            continue
                        joined := ""
                        for filePath in entry["Files"]
                            joined .= (joined = "" ? "" : "`r`n") filePath
                        entry["Text"] := joined
                    case "image":
                        if !(entry.Has("Image") && FileExist(ClipboardProvider.Folder "\" entry["Image"]))
                            continue
                        entry["Text"] := ""
                    default:
                        if entry.Has("File") {                              ; 很长的条目: 正文在单独的文件里
                            try entry["Text"] := FileRead(ClipboardProvider.Folder "\" entry["File"], "UTF-8")
                            catch
                                continue
                        }
                        if !entry.Has("Text")
                            continue
                }
                if !entry.Has("Time")
                    entry["Time"] := A_Now
                entries.Push(entry)
            }
            ClipboardProvider.Entries := entries
        } catch as e {
            Logger.Error("ClipboardProvider: cannot read history - " e.Message)
        }
    }

    static Save() {
        if !AppSettings.Feature("Clipboard")["Persist"]
            return
        try {
            DirCreate(AppSettings.DataDir)
            list := [], keep := Map()
            for entry in ClipboardProvider.Entries {
                stamp := entry.Has("Time") ? entry["Time"] : A_Now, source := entry.Has("App") ? entry["App"] : ""
                switch ClipboardProvider.TypeOf(entry) {
                    case "files":
                        list.Push(ClipboardProvider._WithPin(entry, Map("Type", "files", "Files", entry["Files"], "Time", stamp, "App", source)))
                        continue
                    case "image":
                        keep[StrLower(entry["Image"])] := true
                        list.Push(ClipboardProvider._WithPin(entry, Map("Type", "image", "Image", entry["Image"], "Width", entry["Width"], "Height", entry["Height"], "Time", stamp, "App", source)))
                        continue
                }
                text := entry["Text"]
                if (StrLen(text) <= ClipboardProvider.LargeText) {
                    list.Push(ClipboardProvider._WithPin(entry, Map("Text", text, "Time", stamp, "App", source)))
                    continue
                }
                if (!entry.Has("File") || !FileExist(ClipboardProvider.Folder "\" entry["File"])) {   ; 新的长条目: 只写这一次
                    DirCreate(ClipboardProvider.Folder)
                    entry["File"] := stamp "-" Random(100000, 999999) ".txt"
                    FileAppend(text, ClipboardProvider.Folder "\" entry["File"], "UTF-8")
                }
                keep[StrLower(entry["File"])] := true
                list.Push(ClipboardProvider._WithPin(entry, Map("File", entry["File"], "Time", stamp, "App", source)))
            }
            JSON.WriteFile(ClipboardProvider.File, Map("Entries", list))
            Loop Files, ClipboardProvider.Folder "\*.*" {                  ; 删掉已经不在历史里的长条目和图片
                if (RegExMatch(A_LoopFileName, "i)\.(txt|png)$") && !keep.Has(StrLower(A_LoopFileName)))
                    try FileDelete(A_LoopFileFullPath)
            }
        } catch as e {
            Logger.Error("ClipboardProvider: cannot write history - " e.Message)
        }
    }

    static _WithPin(entry, saved) {
        if ClipboardProvider.IsPinned(entry)
            saved["Pinned"] := 1
        return saved
    }

    static _SaveLater() {
        if (ClipboardProvider._saveTimer = "")
            ClipboardProvider._saveTimer := () => ClipboardProvider.Save()
        SetTimer(ClipboardProvider._saveTimer, -2000)
    }
}
