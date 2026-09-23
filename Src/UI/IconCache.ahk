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
;
; 用法:
;   IconCache.Size := 32                 (由 SearchWindow 按主题和 DPI 设置)
;   hIcon := IconCache.Get(item.Icon)    0 = 没有图标
;===============================================================================

class IconCache {
    static Size   := 32
    static _icons := Map()

    static Get(spec) {
        if (spec = "")
            return 0
        key := IconCache._CacheKey(spec)
        if IconCache._icons.Has(key)
            return IconCache._icons[key]
        hIcon := 0
        try hIcon := IconCache._Load(spec)
        IconCache._icons[key] := hIcon
        return hIcon
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
        SplitPath(spec, , , &ext)
        if (ext = "" || RegExMatch(ext, "i)^(exe|lnk|ico|url|appref-ms|msc|cpl|scr)$") || InStr(spec, "shell:") = 1)
            return StrLower(spec)                                           ; 每个文件自己的图标
        return "ext:." StrLower(ext)
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
