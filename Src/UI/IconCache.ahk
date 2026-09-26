;===============================================================================
; IconCache.ahk - 结果图标 (HICON) 的读取和缓存 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; ResultItem.Icon 的写法:
;   "C:\Windows\notepad.exe"     文件 / 文件夹 / 快捷方式自己的图标
;   "shell:AppsFolder\<AppID>"   shell 路径 (应用商店应用等)
;   "res:imageres.dll,-102"      dll / exe 里的图标资源 (负数 = 资源 Id)
;   "ext:.pdf"                   某种扩展名的默认图标
;   "url:"                       默认浏览器 (网址) 图标
;   "folder:"                    通用文件夹图标
; exe/lnk/ico 等每个文件图标不同, 按完整路径缓存; 其它文件按扩展名缓存。
; 网络位置 (\\server\share、映射的网络盘) 上的文件和文件夹不读磁盘, 用扩展名 / 文件夹的
; 通用图标: 读一次网络上的图标可能要几百毫秒, 而加载图标和打字在同一个线程里。
; 普通文件夹也用通用的文件夹图标 (FolderIcon): 只有带 desktop.ini (自定义图标, 如 OneDrive、
; 桌面、下载) 的文件夹和磁盘根目录才单独读取。
;
; 读取图标 (特别是 exe/lnk、网络路径) 可能要几毫秒到几十毫秒, 所以 Get() 遇到
; 还没加载的图标先返回 0 并放进队列, 由定时器在后台加载, 加载完调用 OnLoaded
; (SearchWindow 在那里重画列表)。这样打字时不会被图标卡住。
;
; 用法:
;   IconCache.Size := 32                 (由 SearchWindow 按主题和 DPI 设置)
;   IconCache.OnLoaded := () => ...      后台加载完一批图标后调用
;   hIcon := IconCache.Get(item.Icon)    0 = 没有图标 / 还在加载
;   item.Icon := IconCache.FolderIcon(path)   文件夹: 通用图标 "folder:" 或它自己的路径
;===============================================================================

class IconCache {
    static Size     := 32
    static OnLoaded := ""
    static _icons   := Map()
    static _queue   := Map()              ; key -> spec, 等待后台加载
    static _timer   := ""

    static Get(spec) {
        if (spec = "")
            return 0
        key := IconCache._CacheKey(spec)
        if IconCache._icons.Has(key)
            return IconCache._icons[key]
        if !IconCache._queue.Has(key) {
            IconCache._queue[key] := IconCache._LoadSpec(spec, key)
            if (IconCache._timer = "")
                IconCache._timer := () => IconCache._LoadQueued()
            SetTimer(IconCache._timer, -1)
        }
        return 0
    }

    ; 立即加载 (不经过队列), 测试或需要马上拿到图标时用
    static GetNow(spec) {
        if (spec = "")
            return 0
        key := IconCache._CacheKey(spec)
        if !IconCache._icons.Has(key) {
            hIcon := 0
            try hIcon := IconCache._Load(IconCache._LoadSpec(spec, key))
            IconCache._icons[key] := hIcon
        }
        return IconCache._icons[key]
    }

    ; 文件夹结果的图标: 普通文件夹都一样, 用通用图标 (不用逐个读, 几百个文件夹时明显更快);
    ; 有自定义图标 (desktop.ini) 的文件夹和磁盘根目录用它自己的。按路径缓存, 每个文件夹只看一次。
    ; 看 desktop.ini 要访问磁盘, 搜索时不做: 先返回通用图标, 放进队列在后台看 (probeNow: 启动后
    ; 预热时直接看)
    static _folderIcons := Map(), _folderProbes := Map(), _probeTimer := ""
    static FolderIcon(folder, probeNow := false) {
        if IconCache._folderIcons.Has(folder)
            return IconCache._folderIcons[folder]
        trimmed := RTrim(folder, "\/")
        if (RegExMatch(trimmed, "^[A-Za-z]:$") || RegExMatch(folder, "i)^(shell:|::\{)"))
            return IconCache._folderIcons[folder] := folder
        if IconCache.IsRemote(folder)
            return IconCache._folderIcons[folder] := "folder:"
        if probeNow
            return IconCache._ProbeFolder(folder)
        IconCache._folderProbes[folder] := true
        if (IconCache._probeTimer = "")
            IconCache._probeTimer := () => IconCache._ProbeQueued()
        SetTimer(IconCache._probeTimer, -1)
        return "folder:"
    }

    static _ProbeFolder(folder) {
        return IconCache._folderIcons[folder] := FileExist(RTrim(folder, "\/") "\desktop.ini") ? folder : "folder:"
    }

    ; 和加载图标一样, 每轮最多 10 ms
    static _ProbeQueued() {
        start := IconCache._Ms()
        for folder in IconCache._folderProbes.Clone() {
            IconCache._folderProbes.Delete(folder)
            IconCache._ProbeFolder(folder)
            if (IconCache._Ms() - start > 10)
                break
        }
        if IconCache._folderProbes.Count
            SetTimer(IconCache._probeTimer, -1)
    }

    ; 每次最多加载约 10 ms, 剩下的留到下一轮, 中间可以处理键盘输入
    static _LoadQueued() {
        start := IconCache._Ms()
        loaded := false
        for key, spec in IconCache._queue.Clone() {
            IconCache._queue.Delete(key)
            hIcon := 0
            try hIcon := IconCache._Load(spec)
            IconCache._icons[key] := hIcon
            loaded := true
            if (IconCache._Ms() - start > 10)
                break
        }
        if IconCache._queue.Count
            SetTimer(IconCache._timer, -1)
        if (loaded && IsObject(IconCache.OnLoaded))
            IconCache.OnLoaded.Call()
    }

    ; 毫秒 (QueryPerformanceCounter; A_TickCount 的精度只有 10 ~ 16 ms)
    static _Ms() {
        static freq := 0
        if !freq
            DllCall("QueryPerformanceFrequency", "Int64*", &freq)
        DllCall("QueryPerformanceCounter", "Int64*", &now := 0)
        return now * 1000 / freq
    }

    static Clear() {
        for key, hIcon in IconCache._icons
            if hIcon
                DllCall("DestroyIcon", "Ptr", hIcon)
        IconCache._icons := Map()
    }

    static _CacheKey(spec) {
        if RegExMatch(spec, "i)^(res|ext|url|folder):")
            return StrLower(spec)
        ext := IconCache._Extension(spec)
        if IconCache.IsRemote(spec)
            return (ext = "") ? "folder:" : "ext:." StrLower(ext)
        if (ext = "" || RegExMatch(ext, "i)^(exe|lnk|ico|url|appref-ms|msc|cpl|scr)$") || InStr(spec, "shell:") = 1)
            return StrLower(spec)                                           ; 每个文件自己的图标
        return "ext:." StrLower(ext)
    }

    ; 像扩展名的才算扩展名 (1~6 个字母数字, 至少一个字母, 如 pdf / docx / dwg / sldprt):
    ; "26. 18 New Industrial Road (EA)"、"10.PT2310-29NIR"、"Design.2019" 这种名字里带点的文件夹没有扩展名
    static _Extension(path) {
        SplitPath(RTrim(path, "\/"), , , &ext)
        return (RegExMatch(ext, "^(?=.*[A-Za-z])[A-Za-z0-9]{1,6}$") || ext = "appref-ms") ? ext : ""
    }

    ; 网络位置按通用图标加载 (缓存键就是 "folder:" / "ext:.pdf"), 其它按原来的路径
    static _LoadSpec(spec, key) {
        return (IconCache.IsRemote(spec) && RegExMatch(key, "^(folder|ext):")) ? key : spec
    }

    ; UNC 路径, 或者 Windows 认为是网络驱动器的盘符 (GetDriveType 不访问网络, 每个盘符只查一次)
    static IsRemote(target) {
        static drives := Map()
        if (SubStr(target, 1, 2) = "\\")
            return true
        if !RegExMatch(target, "^([A-Za-z]):", &match)
            return false
        letter := StrUpper(match[1])
        if !drives.Has(letter)
            drives[letter] := DllCall("GetDriveTypeW", "WStr", letter ":\", "UInt") = 4      ; DRIVE_REMOTE
        return drives[letter]
    }

    static _Load(spec) {
        static FILE_ATTRIBUTE_DIRECTORY := 0x10, FILE_ATTRIBUTE_NORMAL := 0x80
        if (SubStr(spec, 1, 4) = "res:") {
            parts := StrSplit(SubStr(spec, 5), ",")
            return IconCache._FromResource(Trim(parts[1]), parts.Length >= 2 ? Integer(parts[2]) : 0)
        }
        if (SubStr(spec, 1, 4) = "ext:")
            return IconCache._FromShell(SubStr(spec, 5), FILE_ATTRIBUTE_NORMAL, true)
        if (spec = "url:")
            return IconCache._FromShell(".html", FILE_ATTRIBUTE_NORMAL, true)
        if (spec = "folder:")
            return IconCache._FromShell("folder", FILE_ATTRIBUTE_DIRECTORY, true)

        target := Path.Resolve(spec)
        if RegExMatch(target, "i)^(shell:|::\{)")
            return IconCache._FromPidl(target)
        if FileExist(target)
            return IconCache._FromShell(target, 0, false)
        SplitPath(target, , , &ext)
        return IconCache._FromShell(ext != "" ? "." ext : ".exe", FILE_ATTRIBUTE_NORMAL, true)
    }

    ; SHGetFileInfo 取系统图标列表里的序号, 再从合适尺寸的图标列表里取图标
    static _FromShell(target, attributes, useAttributes) {
        static SHGFI_SYSICONINDEX := 0x4000, SHGFI_USEFILEATTRIBUTES := 0x10
        info := Buffer(A_PtrSize + 688, 0)
        flags := SHGFI_SYSICONINDEX | (useAttributes ? SHGFI_USEFILEATTRIBUTES : 0)
        if !DllCall("shell32\SHGetFileInfoW", "WStr", target, "UInt", attributes, "Ptr", info, "UInt", info.Size, "UInt", flags, "UPtr")
            return 0
        return IconCache._FromSystemList(NumGet(info, A_PtrSize, "Int"))
    }

    static _FromPidl(target) {
        static SHGFI_SYSICONINDEX := 0x4000, SHGFI_PIDL := 0x8
        pidl := 0
        if DllCall("shell32\SHParseDisplayName", "WStr", target, "Ptr", 0, "Ptr*", &pidl, "UInt", 0, "Ptr", 0) != 0
            return 0
        info := Buffer(A_PtrSize + 688, 0)
        ok := DllCall("shell32\SHGetFileInfoW", "Ptr", pidl, "UInt", 0, "Ptr", info, "UInt", info.Size, "UInt", SHGFI_SYSICONINDEX | SHGFI_PIDL, "UPtr")
        DllCall("ole32\CoTaskMemFree", "Ptr", pidl)
        return ok ? IconCache._FromSystemList(NumGet(info, A_PtrSize, "Int")) : 0
    }

    ; 系统图标列表: SHIL_LARGE (32px) / SHIL_EXTRALARGE (48px), 按需要的尺寸选
    static _FromSystemList(index) {
        static IID_IImageList := "", lists := Map()
        if (IID_IImageList = "") {
            IID_IImageList := Buffer(16)
            DllCall("ole32\CLSIDFromString", "WStr", "{46EB5926-582E-4017-9FDF-E8998DAA0950}", "Ptr", IID_IImageList)
        }
        listType := (IconCache.Size > 32) ? 2 : 0                           ; 2 = SHIL_EXTRALARGE, 0 = SHIL_LARGE
        if !lists.Has(listType) {
            imageList := 0
            DllCall("shell32\SHGetImageList", "Int", listType, "Ptr", IID_IImageList, "Ptr*", &imageList)
            lists[listType] := imageList
        }
        if !lists[listType]
            return 0
        return DllCall("comctl32\ImageList_GetIcon", "Ptr", lists[listType], "Int", index, "UInt", 1, "Ptr")   ; 1 = ILD_TRANSPARENT
    }

    static _FromResource(file, index) {
        hIcon := 0
        size := IconCache.Size
        ; PrivateExtractIcons 可以直接取指定尺寸; 负数序号表示资源 Id
        DllCall("user32\PrivateExtractIconsW", "WStr", file, "Int", index, "Int", size, "Int", size, "Ptr*", &hIcon, "Ptr", 0, "UInt", 1, "UInt", 0)
        return hIcon
    }
}
