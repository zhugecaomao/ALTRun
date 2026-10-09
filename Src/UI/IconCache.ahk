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
;   "thumb:C:\x.png"             图片本身的缩略图 (剪贴板历史里的图片)
; exe/lnk/ico 等每个文件图标不同, 按完整路径缓存; 其它文件按扩展名缓存。
; 网络位置 (\\server\share、映射的网络盘) 上的文件和文件夹不读磁盘, 用扩展名 / 文件夹的
; 通用图标: 读一次网络上的图标可能要几百毫秒, 而加载图标和打字在同一个线程里。
; 普通文件夹也用通用的文件夹图标 (FolderIcon): 只有带 desktop.ini (自定义图标, 如 OneDrive、
; 桌面、下载) 的文件夹和磁盘根目录才单独读取。
;
;
; 图标按 Size 读取: 取不小于 Size 的系统尺寸 (16 / 32 / 48...), 不是正好的尺寸 (例如紧凑主题的 28 px、
; 125% 缩放时的 40 px) 用 GDI+ 高质量双三次插值缩放成 Size x Size 再缓存 (FitSize), 边缘平滑没有锯齿。
;
; 缓存最多 MaxIcons 个: ALTRun 常常开机后连续运行几周, 图标句柄 (每个程序最多 1 万个) 不能只增不减。
; 超过时去掉最久没用到的 (DestroyIcon), 剩 80%; 再用到时重新加载。
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
    static MaxIcons := 1500
    static SlowLoadMs := 200              ; 读一个图标超过这么多毫秒时写进调试日志 (见 _LoadQueued)
    static _used    := Map(), _tick := 0  ; key -> 最后一次用到的序号 (越大越新)

    static Get(spec) {
        if (spec = "")
            return 0
        key := IconCache._CacheKey(spec)
        if IconCache._icons.Has(key) {
            IconCache._used[key] := ++IconCache._tick
            return IconCache._icons[key]
        }
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
            IconCache._Store(key, hIcon)
        }
        IconCache._used[key] := ++IconCache._tick
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

    ; 和加载图标一样, 每轮最多 10 ms。遍历的是副本: 中途另一轮 (定时器打断了这一轮) 可能已经处理掉
    ; 后面的文件夹, 已经不在队列里的跳过。用 try 删除, 不先 Has 再 Delete: 定时器也可能正好在两行之间
    ; 打断, Delete 一个已经没有的项会报 "Item has no value"
    static _ProbeQueued() {
        start := IconCache._Ms()
        for folder in IconCache._folderProbes.Clone() {
            try IconCache._folderProbes.Delete(folder)
            catch
                continue
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
            try IconCache._queue.Delete(key)                                ; 另一轮已经加载了 (同 _ProbeQueued)
            catch
                continue
            hIcon := 0
            Logger.Trace("icon: " spec)
            loadStart := IconCache._Ms()
            try hIcon := IconCache._Load(spec)
            ; 开机时资源管理器 / OneDrive 很忙, 第一次读图标 (外壳初始化图标列表) 可能卡很久,
            ; 这期间打字、失去焦点隐藏都在排队。慢的记下来, 看得出卡在哪个图标上
            if ((took := IconCache._Ms() - loadStart) >= IconCache.SlowLoadMs)
                Logger.Debug("Perf: icon " Round(took) " ms: " spec)
            IconCache._Store(key, hIcon)
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
        IconCache._icons := Map(), IconCache._used := Map()
    }

    static _Store(key, hIcon) {
        IconCache._icons[key] := hIcon
        IconCache._used[key] := ++IconCache._tick
        if (IconCache._icons.Count > IconCache.MaxIcons)
            IconCache.Trim(Round(IconCache.MaxIcons * 0.8))
    }

    ; 只留 keep 个最近用到的图标
    static Trim(keep) {
        drop := IconCache._icons.Count - keep
        if (drop <= 0)
            return
        lines := ""
        for key in IconCache._icons
            lines .= (IconCache._used.Has(key) ? IconCache._used[key] : 0) "`t" key "`n"
        for line in StrSplit(RTrim(Sort(lines, "N"), "`n"), "`n") {
            if (drop-- <= 0)
                break
            key := SubStr(line, InStr(line, "`t") + 1)
            if (hIcon := IconCache._icons[key])
                DllCall("DestroyIcon", "Ptr", hIcon)
            IconCache._icons.Delete(key)
            if IconCache._used.Has(key)
                IconCache._used.Delete(key)
        }
    }

    static _CacheKey(spec) {
        if RegExMatch(spec, "i)^(res|ext|url|folder|thumb):")
            return StrLower(spec)
        ext := IconCache._Extension(spec)
        if IconCache.IsRemote(spec)
            return (ext = "") ? "folder:" : "ext:." StrLower(ext)
        if (ext = "" || RegExMatch(ext, "i)^(exe|lnk|ico|url|appref-ms|msc|cpl|scr)$") || InStr(spec, "shell:") = 1)
            return StrLower(spec)                                           ; 每个文件自己的图标
        return "ext:." StrLower(ext)
    }

    ; 像扩展名的才算扩展名 (1~6 个字母数字, 至少一个字母, 如 pdf / docx / dwg / sldprt):
    ; "26. Main Street (A)"、"10.Project-2310"、"Design.2019" 这种名字里带点的文件夹没有扩展名
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

    ; 读出图标后缩放到正好 Size x Size (见 FitSize)
    static _Load(spec) {
        return IconCache.FitSize(IconCache._LoadRaw(spec))
    }

    static _LoadRaw(spec) {
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
        if (SubStr(spec, 1, 6) = "thumb:")
            return IconCache._FromImage(SubStr(spec, 7))

        target := Path.Resolve(spec)
        if RegExMatch(target, "i)^(shell:|::\{)") {
            ; 控制面板里的项目 (设备和打印机等) 只写 ::{CLSID} 时在桌面下找不到, 再到控制面板下面找
            if (!(hIcon := IconCache._FromPidl(target)) && RegExMatch(target, "^::\{[^\\]+\}$"))
                hIcon := IconCache._FromPidl("::{21EC2020-3AEA-1069-A2DD-08002B30309D}\" target)
            return hIcon
        }
        SplitPath(target, , , &ext)
        if FileExist(target) {
            if (ext = "ico") {                                              ; .ico 直接读里面的图片 (Resources\Icons 的图标), 不依赖资源管理器的缩略图
                size := IconCache.SourceSize(IconCache.Size)
                if (hIcon := DllCall("LoadImageW", "Ptr", 0, "WStr", target, "UInt", 1, "Int", size, "Int", size, "UInt", 0x10, "Ptr"))   ; IMAGE_ICON, LR_LOADFROMFILE
                    return hIcon
            }
            return IconCache._FromShell(target, 0, false)
        }
        return IconCache._FromShell(ext != "" ? "." ext : ".exe", FILE_ATTRIBUTE_NORMAL, true)
    }

    ; 图片缩小后放在 Size x Size 的透明方块中间 (保持比例), 做成图标
    static _FromImage(file) {
        if (!FileExist(file) || !ClipboardData.StartGdiplus())
            return 0
        image := 0
        if (DllCall("gdiplus\GdipLoadImageFromFile", "WStr", file, "Ptr*", &image) || !image)
            return 0
        width := 0, height := 0, canvas := 0, graphics := 0, hIcon := 0
        DllCall("gdiplus\GdipGetImageWidth", "Ptr", image, "UInt*", &width)
        DllCall("gdiplus\GdipGetImageHeight", "Ptr", image, "UInt*", &height)
        size := IconCache.Size, scale := (width && height) ? Min(size / width, size / height) : 1
        w := Max(1, Round(width * scale)), h := Max(1, Round(height * scale))
        DllCall("gdiplus\GdipCreateBitmapFromScan0", "Int", size, "Int", size, "Int", 0, "Int", 0x26200A, "Ptr", 0, "Ptr*", &canvas)   ; 32bppARGB
        DllCall("gdiplus\GdipGetImageGraphicsContext", "Ptr", canvas, "Ptr*", &graphics)
        DllCall("gdiplus\GdipSetInterpolationMode", "Ptr", graphics, "Int", 7)                  ; HighQualityBicubic
        DllCall("gdiplus\GdipDrawImageRectI", "Ptr", graphics, "Ptr", image, "Int", (size - w) // 2, "Int", (size - h) // 2, "Int", w, "Int", h)
        DllCall("gdiplus\GdipDeleteGraphics", "Ptr", graphics)
        DllCall("gdiplus\GdipDisposeImage", "Ptr", image)
        DllCall("gdiplus\GdipCreateHICONFromBitmap", "Ptr", canvas, "Ptr*", &hIcon)
        DllCall("gdiplus\GdipDisposeImage", "Ptr", canvas)
        return hIcon
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

    ; 系统图标列表: SHIL_SMALL / SHIL_LARGE / SHIL_EXTRALARGE (96 DPI 下 16 / 32 / 48 px, 随系统 DPI 变大),
    ; 取不小于 Size 的最小的一个 (都比 Size 小时取最大的), 不是正好的尺寸由 FitSize 平滑缩放
    static _FromSystemList(index) {
        static IID_IImageList := "", lists := Map(), sizes := Map()
        if (IID_IImageList = "") {
            IID_IImageList := Buffer(16)
            DllCall("ole32\CLSIDFromString", "WStr", "{46EB5926-582E-4017-9FDF-E8998DAA0950}", "Ptr", IID_IImageList)
        }
        chosen := ""
        for listType in [1, 0, 2] {                                         ; 1 = SHIL_SMALL, 0 = SHIL_LARGE, 2 = SHIL_EXTRALARGE
            if !lists.Has(listType) {
                imageList := 0, width := 0, height := 0
                DllCall("shell32\SHGetImageList", "Int", listType, "Ptr", IID_IImageList, "Ptr*", &imageList)
                if imageList
                    DllCall("comctl32\ImageList_GetIconSize", "Ptr", imageList, "Int*", &width, "Int*", &height)
                lists[listType] := imageList, sizes[listType] := width
            }
            if !lists[listType]
                continue
            chosen := listType
            if (sizes[listType] >= IconCache.Size)
                break
        }
        if (chosen = "")
            return 0
        return DllCall("comctl32\ImageList_GetIcon", "Ptr", lists[chosen], "Int", index, "UInt", 1, "Ptr")   ; 1 = ILD_TRANSPARENT
    }

    static _FromResource(file, index) {
        hIcon := 0
        size := IconCache.SourceSize(IconCache.Size)
        ; PrivateExtractIcons 可以直接取指定尺寸; 负数序号表示资源 Id。取图标里一般都有的尺寸,
        ; 让 Windows 不必自己缩放 (它的缩放有锯齿), 不是正好的尺寸由 FitSize 平滑缩放
        DllCall("user32\PrivateExtractIconsW", "WStr", file, "Int", index, "Int", size, "Int", size, "Ptr*", &hIcon, "Ptr", 0, "UInt", 1, "UInt", 0)
        return hIcon
    }

    ; ALTRun 自己的图标 Resources\Icons\<name>.ico (片段、计算器、帮助、置顶...), 和 Clipboard.ico 同一套画法;
    ; 文件不在时 (例如只复制了 exe) 用 fallback 的系统图标。每个名字只看一次文件在不在
    static Own(name, fallback) {
        static found := Map()
        if !found.Has(name) {
            iconFile := A_ScriptDir "\Resources\Icons\" name ".ico"
            found[name] := FileExist(iconFile) ? iconFile : fallback
        }
        return found[name]
    }

    ; 图标文件里一般都有的尺寸里, 不小于 size 的最小的一个
    static SourceSize(size) {
        for native in [16, 24, 32, 48, 256]
            if (native >= size)
                return native
        return 256
    }

    ; 图标的宽度 (像素), 读不到返回 0
    static IconWidth(hIcon) {
        info := Buffer(8 + 3 * A_PtrSize, 0)                                ; ICONINFO
        if (!hIcon || !DllCall("GetIconInfo", "Ptr", hIcon, "Ptr", info))
            return 0
        maskOffset := (A_PtrSize = 8) ? 16 : 12
        hbmMask := NumGet(info, maskOffset, "Ptr"), hbmColor := NumGet(info, maskOffset + A_PtrSize, "Ptr")
        bitmap := Buffer(16 + 2 * A_PtrSize, 0)                             ; BITMAP
        width := DllCall("GetObject", "Ptr", hbmColor ? hbmColor : hbmMask, "Int", bitmap.Size, "Ptr", bitmap) ? NumGet(bitmap, 4, "Int") : 0
        for handle in [hbmMask, hbmColor]
            if handle
                DllCall("DeleteObject", "Ptr", handle)
        return width
    }

    ; 尺寸不是正好 Size 的图标 (例如紧凑主题的 28 px、125% 缩放时的 40 px) 用 GDI+ 高质量双三次插值
    ; 缩放成 Size x Size, 换掉原来的图标。绘制时 DrawIconEx 不用再缩放: 它的缩放直接丢掉 / 重复像素, 边缘有锯齿
    static FitSize(hIcon) {
        size := IconCache.Size
        if (!hIcon || size <= 0)
            return hIcon
        info := Buffer(8 + 3 * A_PtrSize, 0)                                ; ICONINFO: fIcon, xHotspot, yHotspot, hbmMask, hbmColor
        if !DllCall("GetIconInfo", "Ptr", hIcon, "Ptr", info)
            return hIcon
        maskOffset := (A_PtrSize = 8) ? 16 : 12
        hbmMask := NumGet(info, maskOffset, "Ptr"), hbmColor := NumGet(info, maskOffset + A_PtrSize, "Ptr")
        bitmap := Buffer(16 + 2 * A_PtrSize, 0)                             ; BITMAP: bmType, bmWidth, bmHeight, ...
        width := 0, height := 0
        if (hbmColor && DllCall("GetObject", "Ptr", hbmColor, "Int", bitmap.Size, "Ptr", bitmap))
            width := NumGet(bitmap, 4, "Int"), height := Abs(NumGet(bitmap, 8, "Int"))
        resized := 0
        if (width > 0 && height > 0 && (width != size || height != size) && ClipboardData.StartGdiplus())
            resized := IconCache._Resample(hbmColor, hbmMask, width, height, size)
        if hbmMask
            DllCall("DeleteObject", "Ptr", hbmMask)
        if hbmColor
            DllCall("DeleteObject", "Ptr", hbmColor)
        if !resized
            return hIcon
        DllCall("DestroyIcon", "Ptr", hIcon)
        return resized
    }

    static _Resample(hbmColor, hbmMask, width, height, size) {
        pixels := IconCache._Bits(hbmColor, width, height)
        if !pixels
            return 0
        hasAlpha := false
        Loop width * height
            if (NumGet(pixels, A_Index * 4 - 1, "UChar")) {
                hasAlpha := true
                break
            }
        if !hasAlpha {                                                      ; 旧式图标没有 Alpha 通道: 按掩码 (白 = 透明) 补上
            mask := hbmMask ? IconCache._Bits(hbmMask, width, height) : 0
            Loop width * height {
                offset := A_Index * 4 - 4
                NumPut("UChar", (mask && NumGet(mask, offset, "UInt") & 0xFFFFFF) ? 0 : 255, pixels, offset + 3)
            }
        }
        source := 0, canvas := 0, graphics := 0, attributes := 0, hIcon := 0
        DllCall("gdiplus\GdipCreateBitmapFromScan0", "Int", width, "Int", height, "Int", width * 4, "Int", 0x26200A, "Ptr", pixels, "Ptr*", &source)   ; 32bppARGB
        DllCall("gdiplus\GdipCreateBitmapFromScan0", "Int", size, "Int", size, "Int", 0, "Int", 0x26200A, "Ptr", 0, "Ptr*", &canvas)
        if (source && canvas) {
            DllCall("gdiplus\GdipGetImageGraphicsContext", "Ptr", canvas, "Ptr*", &graphics)
            DllCall("gdiplus\GdipSetInterpolationMode", "Ptr", graphics, "Int", 7)              ; HighQualityBicubic
            DllCall("gdiplus\GdipSetPixelOffsetMode", "Ptr", graphics, "Int", 4)                ; HighQuality: 不偏半个像素
            DllCall("gdiplus\GdipCreateImageAttributes", "Ptr*", &attributes)
            DllCall("gdiplus\GdipSetImageAttributesWrapMode", "Ptr", attributes, "Int", 3, "UInt", 0, "Int", 0)   ; TileFlipXY: 边缘不混进透明的黑边
            DllCall("gdiplus\GdipDrawImageRectRectI", "Ptr", graphics, "Ptr", source, "Int", 0, "Int", 0, "Int", size, "Int", size
                , "Int", 0, "Int", 0, "Int", width, "Int", height, "Int", 2, "Ptr", attributes, "Ptr", 0, "Ptr", 0)   ; 2 = UnitPixel
            DllCall("gdiplus\GdipDisposeImageAttributes", "Ptr", attributes)
            DllCall("gdiplus\GdipDeleteGraphics", "Ptr", graphics)
            DllCall("gdiplus\GdipCreateHICONFromBitmap", "Ptr", canvas, "Ptr*", &hIcon)
        }
        if source
            DllCall("gdiplus\GdipDisposeImage", "Ptr", source)
        if canvas
            DllCall("gdiplus\GdipDisposeImage", "Ptr", canvas)
        return hIcon
    }

    ; 位图 -> 自上而下的 32 位像素 (BGRA), 失败返回 0
    static _Bits(hBitmap, width, height) {
        header := Buffer(40, 0)                                             ; BITMAPINFOHEADER
        NumPut("UInt", 40, "Int", width, "Int", -height, "UShort", 1, "UShort", 32, header)
        pixels := Buffer(width * height * 4, 0)
        hdc := DllCall("GetDC", "Ptr", 0, "Ptr")
        lines := DllCall("GetDIBits", "Ptr", hdc, "Ptr", hBitmap, "UInt", 0, "UInt", height, "Ptr", pixels, "Ptr", header, "UInt", 0)
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdc)
        return lines ? pixels : 0
    }
}
