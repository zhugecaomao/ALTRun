;===============================================================================
; IconCache.ahk - reads and caches result icons (HICON) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Forms of ResultItem.Icon:
;   "C:\Windows\notepad.exe"     the icon of the file / folder / shortcut itself
;   "shell:AppsFolder\<AppID>"   a shell path (Store apps etc.)
;   "res:imageres.dll,-102"      an icon resource in a dll / exe (negative = resource Id)
;   "ext:.pdf"                   the default icon of an extension
;   "url:"                       the default browser (URL) icon
;   "folder:"                    the generic folder icon
;   "thumb:C:\x.png"             a thumbnail of the image itself (images in the clipboard history)
; exe / lnk / ico etc. differ per file and are cached by full path; other files are cached by extension.
; Files and folders on network locations (\\server\share, mapped drives) aren't read from disk; the generic icons of the extension / folder are used:
; reading one icon over the network can take hundreds of milliseconds, and loading icons runs on the same thread as typing.
; Normal folders also use the generic folder icon (FolderIcon): only folders with a desktop.ini (custom icons, e.g. OneDrive,
; Desktop, Downloads) and drive roots are read individually.
;
;
; Icons are read at Size: the smallest system size not below Size (16 / 32 / 48...); sizes that don't fit exactly (e.g. 28 px for the compact density,
; 40 px at 125% scaling) are scaled to Size x Size with high-quality bicubic interpolation in GDI+ before caching (FitSize), so edges are smooth.
;
; At most MaxIcons are cached: ALTRun often runs for weeks after boot, and icon handles (at most 10,000 per program) mustn't only grow.
; Beyond that the least recently used are destroyed (DestroyIcon) down to 80%; they reload when used again.
;
; Disk cache: icons read (scaled to Size x Size pixels) are stored in %LOCALAPPDATA%\ALTRun\IconCache\Icons-<Size>.bin,
; one per PC (the same path may be another program on another PC, and it isn't synced with Data). Loaded after start-up; when Get finds an icon
; missing in memory but in the disk cache, it makes the HICON directly (under 1 ms) without waiting for a background read; after a restart icons show at once too.
; The cached icon is used first and checked in the background (while the window is hidden): when the file's time or size changed (the program updated its icon) it is read again;
; icons that aren't files (extensions, resources, Store apps...) are read again after DiskRecheckDays days.
; At most DiskMax, keeping the most recently used; new icons are collected for a while (DiskSaveDelay) and saved while the window is hidden, and on exit. Deleting the folder does no harm.
;
; Reading an icon (especially exe / lnk or network paths) can take milliseconds to tens of milliseconds, so Get() returns 0 for an
; icon not loaded yet and queues it; a timer loads it in the background and then calls OnLoaded
; (SearchWindow repaints the list there). So icons never hold up typing.
;
; Usage:
;   IconCache.Size := 32                 (set by SearchWindow from the theme and DPI)
;   IconCache.OnLoaded := () => ...      called after a batch of icons was loaded in the background
;   hIcon := IconCache.Get(item.Icon)    0 = no icon / still loading
;   IconCache.Prefetch(specs)            reads common icons in the background after start-up (see Prefetch)
;   IconCache.LoadDisk() / SaveDisk()    loads the disk cache after start-up / saves it on exit (see "disk cache" below)
;   item.Icon := IconCache.FolderIcon(path)   folders: the generic icon "folder:" or the folder's own path
;===============================================================================

class IconCache {
    static Size     := 32
    static OnLoaded := ""
    static _icons   := Map()
    static _queue   := Map()              ; key -> spec, waiting to be loaded in the background
    static _loading := Map()              ; being read (while reading, the UI may ask for the same icon again; don't queue it twice)
    static _timer   := ""
    static MaxIcons := 1500
    static SlowLoadMs := 200              ; reading one icon slower than this many milliseconds goes into the debug log (see _LoadQueued)
    static _used    := Map(), _tick := 0  ; key -> sequence number of its last use (higher = more recent)

    static Get(spec) {
        if (spec = "")
            return 0
        key := IconCache._CacheKey(spec)
        if IconCache._icons.Has(key) {
            IconCache._used[key] := ++IconCache._tick
            return IconCache._icons[key]
        }
        if (hIcon := IconCache._FromDisk(key, spec)) {                      ; in the disk cache: show it now, check it in the background later
            IconCache._Store(key, hIcon)
            return hIcon
        }
        if !(IconCache._queue.Has(key) || IconCache._loading.Has(key)) {
            IconCache._queue[key] := IconCache._LoadSpec(spec, key)
            if (IconCache._timer = "")
                IconCache._timer := () => IconCache._LoadQueued()
            SetTimer(IconCache._timer, -1)
        }
        return 0
    }

    ; Loads immediately (bypassing the queue), for tests or when the icon is needed right away
    static GetNow(spec) {
        if (spec = "")
            return 0
        key := IconCache._CacheKey(spec)
        if !IconCache._icons.Has(key) {
            hIcon := 0
            try hIcon := IconCache._Load(IconCache._LoadSpec(spec, key))
            IconCache._StoreLoaded(key, hIcon)
        }
        IconCache._used[key] := ++IconCache._tick
        return IconCache._icons[key]
    }

    ; After start-up, read the common items' icons in the background (specs in order of priority, at most PrefetchMax):
    ; the first time an item is found its icon appears with the text instead of a moment later (reading a program's
    ; icon takes tens of milliseconds on an old PC). Like the queue of Get, at most about 10 ms per round, and only
    ; when BackgroundWork.CanRun() (the search window is hidden, nobody is typing), so it doesn't compete with the
    ; results on screen; icons already cached are skipped. Prefetched icons aren't on screen: no repaint after reading
    static PrefetchMax := 300, PrefetchGapMs := 30
    static _prefetch := [], _prefetchTimer := "", _prefetchStats := ""
    static Prefetch(specs) {
        IconCache._prefetch := [], seen := Map()
        for spec in specs {
            if (spec = "" || IconCache._prefetch.Length >= IconCache.PrefetchMax)
                continue
            key := IconCache._CacheKey(spec)
            if (seen.Has(key) || IconCache._icons.Has(key) || IconCache._disk.Has(key))
                continue
            seen[key] := true
            IconCache._prefetch.Push(spec)
        }
        IconCache._prefetchStats := {Count: 0, Ms: 0, Total: IconCache._prefetch.Length}
        if (IconCache._prefetchTimer = "")
            IconCache._prefetchTimer := Logger.Watch("icon prefetch", () => IconCache._PrefetchStep())
        SetTimer(IconCache._prefetchTimer, IconCache._prefetch.Length ? IconCache.PrefetchGapMs : 0)
    }

    static PrefetchPending => IconCache._prefetch.Length

    static _PrefetchStep() {
        if (IconCache._queue.Count || !BackgroundWork.CanRun())
            return                                                          ; icons on screen first / the window is shown: next round
        start := IconCache._Ms(), stats := IconCache._prefetchStats
        while IconCache._prefetch.Length {
            spec := IconCache._prefetch.RemoveAt(1)
            key := IconCache._CacheKey(spec)
            if (IconCache._icons.Has(key) || IconCache._disk.Has(key) || IconCache._loading.Has(key))   ; icons in the disk cache are made directly when used, no prefetch needed
                continue
            hIcon := 0
            try hIcon := IconCache._Load(IconCache._LoadSpec(spec, key))
            IconCache._StoreLoaded(key, hIcon)
            stats.Count += 1
            if (IconCache._Ms() - start > 10)
                break
        }
        stats.Ms += IconCache._Ms() - start
        if !IconCache._prefetch.Length {
            SetTimer(IconCache._prefetchTimer, 0)
            Logger.Debug(Format("Perf: icons prefetched {} of {} in {} ms", stats.Count, stats.Total, Round(stats.Ms)))
        }
    }

    ;---------------------------------------------------------------------------
    ; The disk cache (see the top of the file). File format: a 16-byte header (marker "ARIC", version, reserved, icon size, count), then per icon:
    ; key / stamp (UTF-16, preceded by the length), date saved, date last used (yyyyMMdd), Size x Size BGRA pixels.
    ; Not compressed: measured on Windows, 830 icons were 1.2 MB compressed / 3.4 MB uncompressed, but loading took 12 ms / 5 ms and saving 24 ms / 11 ms;
    ; speed matters more than the two or three MB saved
    ;---------------------------------------------------------------------------
    static DiskDir := EnvGet("LOCALAPPDATA") != "" ? EnvGet("LOCALAPPDATA") "\ALTRun\IconCache" : ""
    static DiskMax := 1000, DiskRecheckDays := 7, DiskSaveDelay := 30000, DiskVersion := 2, DiskChunk := 262144
    static _disk := Map()                 ; key -> {Stamp, Saved, Used, Data, Offset}: the pixels are at Offset in Data
    static _diskDirty := false, _diskTouched := false, _saveTimer := "", _saving := false
    static _verify := [], _verified := Map(), _verifyTimer := ""

    static DiskFile() => (IconCache.DiskDir = "") ? "" : IconCache.DiskDir "\Icons-" IconCache.Size ".bin"

    static LoadDisk(file := "") {
        file := (file = "") ? IconCache.DiskFile() : file
        IconCache._disk := Map()
        if (file = "" || !FileExist(file))
            return 0
        start := IconCache._Ms()
        try {
            IconCache._disk := IconCache._ParseDisk(FileRead(file, "RAW"), IconCache.Size)
        } catch as e {
            Logger.Debug("IconCache: disk cache ignored - " e.Message)
            IconCache._disk := Map()
        }
        Logger.Debug(Format("Perf: icon cache loaded {} icons in {} ms", IconCache._disk.Count, Round(IconCache._Ms() - start)))
        return IconCache._disk.Count
    }

    ; Only preparing the data keeps other threads out (a few milliseconds); the file is written in chunks (it can take hundreds of milliseconds with a busy disk or antivirus scan),
    ; so a hotkey pressed meanwhile still shows the window at once. Written to a temporary file and then replaced, so exiting halfway never leaves a half-written cache
    static SaveDisk(file := "") {
        file := (file = "") ? IconCache.DiskFile() : file
        if (file = "" || !(IconCache._diskDirty || IconCache._diskTouched) || IconCache._saving)
            return false
        temp := file ".tmp", f := "", IconCache._saving := true
        try {
            start := IconCache._Ms()
            Critical("On")
            try data := IconCache._SerializeDisk(IconCache._disk, IconCache.Size, IconCache.DiskMax)
            finally Critical("Off")
            IconCache._diskDirty := false, IconCache._diskTouched := false   ; new icons arrived while writing: save again later
            built := IconCache._Ms()
            SplitPath(file, , &folder)
            DirCreate(folder)
            f := FileOpen(temp, "w")
            pos := 0
            while (pos < data.Size) {
                n := Min(IconCache.DiskChunk, data.Size - pos)
                f.RawWrite(data.Ptr + pos, n)
                pos += n
            }
            f.Close(), f := ""
            written := IconCache._Ms()
            FileMove(temp, file, true)
            Logger.Debug(Format("Perf: icon cache saved {} KB: build {} ms, write {} ms, replace {} ms", data.Size // 1024
                , Round(built - start), Round(written - built), Round(IconCache._Ms() - written)))
            return true
        } catch as e {
            if IsObject(f)
                try f.Close()
            try FileDelete(temp)
            IconCache._diskDirty := true
            Logger.Debug("IconCache: saving the disk cache failed - " e.Message)
            return false
        } finally {
            IconCache._saving := false
        }
    }

    static _ScheduleSave() {
        IconCache._diskDirty := true
        if (IconCache.DiskDir = "")
            return
        if (IconCache._saveTimer = "")
            IconCache._saveTimer := Logger.Watch("save icon cache", () => IconCache._SaveWhenIdle())
        SetTimer(IconCache._saveTimer, -IconCache.DiskSaveDelay)
    }

    static _SaveWhenIdle() {
        if !BackgroundWork.CanRun()
            return SetTimer(IconCache._saveTimer, -IconCache.DiskSaveDelay)  ; the window is shown: save after it hides
        IconCache.SaveDisk()
    }

    ; Makes an HICON from the disk cache if it is there, and schedules a background check (each icon once per run)
    static _FromDisk(key, spec) {
        if !IconCache._disk.Has(key)
            return 0
        entry := IconCache._disk[key]
        if !(hIcon := IconCache.MakeIcon(entry.Data, entry.Offset)) {
            IconCache._disk.Delete(key)
            return 0
        }
        today := IconCache._Today()
        if (entry.Used != today)
            entry.Used := today, IconCache._diskTouched := true
        if !IconCache._verified.Has(key) {
            IconCache._verified[key] := true
            IconCache._verify.Push([key, spec])
            if (IconCache._verifyTimer = "")
                IconCache._verifyTimer := Logger.Watch("icon verify", () => IconCache._VerifyStep())
            SetTimer(IconCache._verifyTimer, IconCache.PrefetchGapMs)
        }
        return hIcon
    }

    ; like the prefetch, only while the window is hidden, at most about 10 ms per round
    static _VerifyStep() {
        if (IconCache._queue.Count || !BackgroundWork.CanRun())
            return
        start := IconCache._Ms(), changed := false
        while IconCache._verify.Length {
            item := IconCache._verify.RemoveAt(1), key := item[1], spec := item[2]
            if (!IconCache._disk.Has(key) || !IconCache.IsStale(key, IconCache._disk[key]))
                continue
            hIcon := 0
            try hIcon := IconCache._Load(IconCache._LoadSpec(spec, key))
            if hIcon {
                IconCache._Remember(key, hIcon)
                if IconCache._icons.Has(key) {                              ; swap in the new icon
                    old := IconCache._icons[key], IconCache._icons[key] := hIcon, changed := true
                    if old
                        DllCall("DestroyIcon", "Ptr", old)
                } else {
                    DllCall("DestroyIcon", "Ptr", hIcon)
                }
            }
            if (IconCache._Ms() - start > 10)
                break
        }
        if !IconCache._verify.Length
            SetTimer(IconCache._verifyTimer, 0)
        if (changed && IsObject(IconCache.OnLoaded))
            IconCache.OnLoaded.Call()
    }

    ; files: the time or size changed; others: stored DiskRecheckDays days or more ago
    static IsStale(key, entry) {
        if IconCache.IsFileKey(key)
            return IconCache.Stamp(key) != entry.Stamp
        return IconCache.DaysSince(entry.Saved) >= IconCache.DiskRecheckDays
    }

    static IsFileKey(key) => !RegExMatch(key, "i)^(res|ext|url|folder|thumb|shell):|^::\{")

    ; A file's "modified time|size" (folders have size 0); "" if it doesn't exist
    static Stamp(key) {
        try {
            target := Path.Resolve(key)
            if FileExist(target)
                return FileGetTime(target, "M") "|" FileGetSize(target)
        }
        return ""
    }

    static DaysSince(day) {
        try return DateDiff(A_Now, day "000000", "Days")
        return 9999
    }

    static _Today() => Integer(A_YYYY A_MM A_DD)

    ; A newly read icon goes into the disk cache (thumbnails of clipboard images don't: each image is only used for a while)
    static _Remember(key, hIcon) {
        if (!hIcon || IconCache.DiskDir = "" || SubStr(key, 1, 6) = "thumb:")
            return
        if !(pixels := IconCache.IconPixels(hIcon))
            return
        today := IconCache._Today()
        IconCache._disk[key] := {Stamp: IconCache.IsFileKey(key) ? IconCache.Stamp(key) : "", Saved: today, Used: today, Data: pixels, Offset: 0}
        IconCache._verified[key] := true
        if (IconCache._disk.Count > IconCache.DiskMax + 200)               ; ALTRun often runs for weeks: what is in memory mustn't only grow either
            IconCache._TrimDisk(IconCache.DiskMax)
        IconCache._ScheduleSave()
    }

    ; keep only the keep most recently used
    static _TrimDisk(keep) {
        lines := ""
        for key, entry in IconCache._disk
            lines .= entry.Used "`t" key "`n"
        for line in StrSplit(RTrim(Sort(lines, "N R"), "`n"), "`n")
            if (A_Index > keep)
                IconCache._disk.Delete(SubStr(line, InStr(line, "`t") + 1))
    }

    ; Icon -> Size x Size BGRA pixels (with alpha); 0 for a wrong size or if it can't be read
    static IconPixels(hIcon) {
        size := IconCache.Size
        info := Buffer(8 + 3 * A_PtrSize, 0)                                ; ICONINFO
        if (!hIcon || !DllCall("GetIconInfo", "Ptr", hIcon, "Ptr", info))
            return 0
        maskOffset := (A_PtrSize = 8) ? 16 : 12
        hbmMask := NumGet(info, maskOffset, "Ptr"), hbmColor := NumGet(info, maskOffset + A_PtrSize, "Ptr")
        pixels := 0, bitmap := Buffer(16 + 2 * A_PtrSize, 0)
        if (hbmColor && DllCall("GetObject", "Ptr", hbmColor, "Int", bitmap.Size, "Ptr", bitmap)
            && NumGet(bitmap, 4, "Int") = size && Abs(NumGet(bitmap, 8, "Int")) = size
            && (pixels := IconCache._Bits(hbmColor, size, size)))
            IconCache._FixAlpha(pixels, hbmMask, size, size)
        for handle in [hbmMask, hbmColor]
            if handle
                DllCall("DeleteObject", "Ptr", handle)
        return pixels
    }

    ; Size x Size BGRA pixels at offset in data -> an icon. The pixels go unchanged into a 32-bit DIB (CreateIconIndirect), exactly as they were read;
    ; not through GDI+ (it multiplies the colors by alpha when making icons, so semi-transparent edges get darker with each round trip).
    ; 32-bit icons with alpha are drawn using the alpha, so an all-zero mask will do
    static MakeIcon(data, offset := 0) {
        size := IconCache.Size
        if (size <= 0 || data.Size < offset + size * size * 4)
            return 0
        header := Buffer(40, 0)                                             ; BITMAPINFOHEADER: top-down, 32-bit
        NumPut("UInt", 40, "Int", size, "Int", -size, "UShort", 1, "UShort", 32, header)
        bits := 0
        if !(color := DllCall("CreateDIBSection", "Ptr", 0, "Ptr", header, "UInt", 0, "Ptr*", &bits, "Ptr", 0, "UInt", 0, "Ptr"))
            return 0
        DllCall("RtlMoveMemory", "Ptr", bits, "Ptr", data.Ptr + offset, "UPtr", size * size * 4)
        mask := DllCall("CreateBitmap", "Int", size, "Int", size, "UInt", 1, "UInt", 1, "Ptr", Buffer((size + 15) // 16 * 2 * size, 0), "Ptr")
        info := Buffer(8 + 3 * A_PtrSize, 0)                                ; ICONINFO: fIcon, xHotspot, yHotspot, hbmMask, hbmColor
        maskOffset := (A_PtrSize = 8) ? 16 : 12
        NumPut("Int", 1, info, 0)
        NumPut("Ptr", mask, info, maskOffset)
        NumPut("Ptr", color, info, maskOffset + A_PtrSize)
        hIcon := DllCall("CreateIconIndirect", "Ptr", info, "Ptr")
        DllCall("DeleteObject", "Ptr", mask)
        DllCall("DeleteObject", "Ptr", color)
        return hIcon
    }

    ; most recently used first, at most limit
    static _SerializeDisk(entries, size, limit) {
        lines := ""
        for key, entry in entries
            lines .= entry.Used "`t" key "`n"
        keys := []
        for line in StrSplit(RTrim(Sort(lines, "N R"), "`n"), "`n") {
            if (line = "" || keys.Length >= limit)
                continue
            keys.Push(SubStr(line, InStr(line, "`t") + 1))
        }
        pixelBytes := size * size * 4, total := 16
        for key in keys
            total += 2 + StrLen(key) * 2 + 2 + StrLen(entries[key].Stamp) * 2 + 8 + pixelBytes
        out := Buffer(total)
        NumPut("UInt", 0x43495241, "UShort", IconCache.DiskVersion, "UShort", 0, "UInt", size, "UInt", keys.Length, out)   ; "ARIC"
        pos := 16
        for key in keys {
            entry := entries[key]
            for text in [key, entry.Stamp] {
                NumPut("UShort", StrLen(text), out, pos)
                if (text != "")
                    StrPut(text, out.Ptr + pos + 2, StrLen(text), "UTF-16")
                pos += 2 + StrLen(text) * 2
            }
            NumPut("UInt", entry.Saved, "UInt", entry.Used, out, pos)
            DllCall("RtlMoveMemory", "Ptr", out.Ptr + pos + 8, "Ptr", entry.Data.Ptr + entry.Offset, "UPtr", pixelBytes)
            pos += 8 + pixelBytes
        }
        return out
    }

    ; the pixels aren't copied: the data read is used directly (Offset points into data)
    static _ParseDisk(data, size) {
        entries := Map()
        if (data.Size < 16 || NumGet(data, 0, "UInt") != 0x43495241 || NumGet(data, 4, "UShort") != IconCache.DiskVersion
            || NumGet(data, 8, "UInt") != size)
            return entries
        count := NumGet(data, 12, "UInt"), end := data.Size
        pixelBytes := size * size * 4, pos := 16
        Loop count {
            texts := []
            Loop 2 {
                if (pos + 2 > end)
                    throw Error("bad entry")
                length := NumGet(data, pos, "UShort")
                if (pos + 2 + length * 2 > end)
                    throw Error("bad entry")
                texts.Push(length ? StrGet(data.Ptr + pos + 2, length, "UTF-16") : "")
                pos += 2 + length * 2
            }
            if (pos + 8 + pixelBytes > end)
                throw Error("bad entry")
            entries[texts[1]] := {Stamp: texts[2], Saved: NumGet(data, pos, "UInt"), Used: NumGet(data, pos + 4, "UInt"), Data: data, Offset: pos + 8}
            pos += 8 + pixelBytes
        }
        return entries
    }

    ; Icons of folder results: normal folders all look the same, so the generic icon is used (not read one by one, clearly faster with hundreds of folders);
    ; folders with a custom icon (desktop.ini) and drive roots use their own. Cached by path; each folder is looked at once.
    ; Checking desktop.ini touches the disk, so it isn't done while searching: the generic icon is returned first and the folder queued to check in the background (probeNow:
    ; check right away, when warming up after start-up)
    static _folderIcons := Map(), _folderProbes := Map(), _probeTimer := ""
    static FolderIcon(folder, probeNow := false) {
        if IconCache._folderIcons.Has(folder)
            return IconCache._folderIcons[folder]
        trimmed := RTrim(folder, "\/")
        if (RegExMatch(trimmed, "^[A-Za-z]:$") || RegExMatch(folder, "i)^(shell:|::\{)"))
            return IconCache._folderIcons[folder] := folder
        if Path.IsRemote(folder)
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

    ; Like loading icons, at most 10 ms per round. It goes through a copy: another round (a timer interrupting this one) may already have handled
    ; later folders; those no longer queued are skipped. Deleted with try instead of Has then Delete: a timer may also interrupt between the two lines,
    ; and deleting a missing item throws "Item has no value"
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

    ; load for at most about 10 ms each time and leave the rest for the next round, so keyboard input can be handled in between
    static _LoadQueued() {
        start := IconCache._Ms()
        loaded := false
        for key, spec in IconCache._queue.Clone() {
            try IconCache._queue.Delete(key)                                ; another round already loaded it (as in _ProbeQueued)
            catch
                continue
            hIcon := 0
            Logger.Trace("icon: " spec)
            loadStart := IconCache._Ms()
            IconCache._loading[key] := true                                 ; reading shell icons may handle messages and repaint the list, which asks for this icon again
            try hIcon := IconCache._Load(spec)
            IconCache._loading.Delete(key)
            ; At boot Explorer / OneDrive are busy, and the first icon read (the shell initializing its image list) can stall for a long time,
            ; with typing and hiding on focus loss queued meanwhile. Slow ones are logged to show which icon it got stuck on
            if ((took := IconCache._Ms() - loadStart) >= IconCache.SlowLoadMs)
                Logger.Debug("Perf: icon " Round(took) " ms: " spec)
            IconCache._StoreLoaded(key, hIcon)
            loaded := true
            if (IconCache._Ms() - start > 10)
                break
        }
        if IconCache._queue.Count
            SetTimer(IconCache._timer, -1)
        if (loaded && IsObject(IconCache.OnLoaded))
            IconCache.OnLoaded.Call()
    }

    ; milliseconds (QueryPerformanceCounter; A_TickCount is only accurate to 10 ~ 16 ms)
    static _Ms() {
        static freq := 0
        if !freq
            DllCall("QueryPerformanceFrequency", "Int64*", &freq)
        DllCall("QueryPerformanceCounter", "Int64*", &now := 0)
        return now * 1000 / freq
    }

    ; Removes all icons and clears the disk cache too (the size changed, or a file got new content, e.g. a site icon downloaded again)
    static Clear() {
        IconCache._prefetch := [], IconCache._verify := [], IconCache._verified := Map()
        for key, hIcon in IconCache._icons
            if hIcon
                DllCall("DestroyIcon", "Ptr", hIcon)
        IconCache._icons := Map(), IconCache._used := Map()
        if IconCache._disk.Count
            IconCache._disk := Map(), IconCache._ScheduleSave()
    }

    ; A newly read icon: into the cache, and into the disk cache
    static _StoreLoaded(key, hIcon) {
        IconCache._Store(key, hIcon)
        IconCache._Remember(key, hIcon)
    }

    static _Store(key, hIcon) {
        IconCache._icons[key] := hIcon
        IconCache._used[key] := ++IconCache._tick
        if (IconCache._icons.Count > IconCache.MaxIcons)
            IconCache.Trim(Round(IconCache.MaxIcons * 0.8))
    }

    ; keep only the keep most recently used icons
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
        ext := Path.Extension(spec)
        if Path.IsRemote(spec)
            return (ext = "") ? "folder:" : "ext:." StrLower(ext)
        if (ext = "" || RegExMatch(ext, "i)^(exe|lnk|ico|url|appref-ms|msc|cpl|scr)$") || InStr(spec, "shell:") = 1)
            return StrLower(spec)                                           ; each file has its own icon
        return "ext:." StrLower(ext)
    }

    ; network locations load the generic icons (the cache key is "folder:" / "ext:.pdf"), others the original path
    static _LoadSpec(spec, key) {
        return (Path.IsRemote(spec) && RegExMatch(key, "^(folder|ext):")) ? key : spec
    }

    ; after reading, scale the icon to exactly Size x Size (see FitSize)
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
            ; Control Panel items (Devices and Printers etc.) given only as ::{CLSID} aren't found under the desktop; look under Control Panel too
            if (!(hIcon := IconCache._FromPidl(target)) && RegExMatch(target, "^::\{[^\\]+\}$"))
                hIcon := IconCache._FromPidl("::{21EC2020-3AEA-1069-A2DD-08002B30309D}\" target)
            return hIcon
        }
        SplitPath(target, , , &ext)
        if FileExist(target) {
            if (ext = "ico") {                                              ; .ico files are read directly (the icons in Resources\Icons), without relying on Explorer's thumbnails
                size := IconCache.SourceSize(IconCache.Size)
                if (hIcon := DllCall("LoadImageW", "Ptr", 0, "WStr", target, "UInt", 1, "Int", size, "Int", size, "UInt", 0x10, "Ptr"))   ; IMAGE_ICON, LR_LOADFROMFILE
                    return hIcon
            }
            return IconCache._FromShell(target, 0, false)
        }
        return IconCache._FromShell(ext != "" ? "." ext : ".exe", FILE_ATTRIBUTE_NORMAL, true)
    }

    ; The image is shrunk and centered in a transparent Size x Size square (keeping its proportions) and made into an icon
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

    ; SHGetFileInfo gives the index in the system image list, then the icon is taken from the image list of a suitable size
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

    ; System image lists: SHIL_SMALL / SHIL_LARGE / SHIL_EXTRALARGE (16 / 32 / 48 px at 96 DPI, larger with the system DPI);
    ; take the smallest not below Size (the largest if all are smaller); sizes that don't fit exactly are scaled smoothly by FitSize
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
        ; PrivateExtractIcons can take a given size directly; a negative index is a resource Id. Take a size icons usually have,
        ; so Windows doesn't scale it itself (its scaling is jagged); sizes that don't fit exactly are scaled smoothly by FitSize
        DllCall("user32\PrivateExtractIconsW", "WStr", file, "Int", index, "Int", size, "Int", size, "Ptr*", &hIcon, "Ptr", 0, "UInt", 1, "UInt", 0)
        return hIcon
    }

    ; ALTRun's own icons Resources\Icons\<name>.ico (snippets, calculator, help, pin...), drawn in the same style as Clipboard.ico;
    ; without the file (e.g. only the exe was copied) the fallback system icon is used. Each name checks for the file once
    static Own(name, fallback) {
        static found := Map()
        if !found.Has(name) {
            iconFile := A_ScriptDir "\Resources\Icons\" name ".ico"
            found[name] := FileExist(iconFile) ? iconFile : fallback
        }
        return found[name]
    }

    ; Of the sizes icon files usually have, the smallest not below size
    static SourceSize(size) {
        for native in [16, 24, 32, 48, 256]
            if (native >= size)
                return native
        return 256
    }

    ; Width of the icon (pixels), 0 if it can't be read
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

    ; Icons not exactly Size (e.g. 28 px for the compact density, 40 px at 125% scaling) are scaled with high-quality bicubic interpolation in GDI+
    ; to Size x Size, replacing the original. Drawing with DrawIconEx then needs no scaling: its scaling just drops / repeats pixels, so edges are jagged
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
        IconCache._FixAlpha(pixels, hbmMask, width, height)
        source := 0, canvas := 0, graphics := 0, attributes := 0, hIcon := 0
        DllCall("gdiplus\GdipCreateBitmapFromScan0", "Int", width, "Int", height, "Int", width * 4, "Int", 0x26200A, "Ptr", pixels, "Ptr*", &source)   ; 32bppARGB
        DllCall("gdiplus\GdipCreateBitmapFromScan0", "Int", size, "Int", size, "Int", 0, "Int", 0x26200A, "Ptr", 0, "Ptr*", &canvas)
        if (source && canvas) {
            DllCall("gdiplus\GdipGetImageGraphicsContext", "Ptr", canvas, "Ptr*", &graphics)
            DllCall("gdiplus\GdipSetInterpolationMode", "Ptr", graphics, "Int", 7)              ; HighQualityBicubic
            DllCall("gdiplus\GdipSetPixelOffsetMode", "Ptr", graphics, "Int", 4)                ; HighQuality: no half-pixel offset
            DllCall("gdiplus\GdipCreateImageAttributes", "Ptr*", &attributes)
            DllCall("gdiplus\GdipSetImageAttributesWrapMode", "Ptr", attributes, "Int", 3, "UInt", 0, "Int", 0)   ; TileFlipXY: no transparent black fringe at the edges
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

    ; old-style icons have no alpha channel (all 0): fill it from the mask (white = transparent)
    static _FixAlpha(pixels, hbmMask, width, height) {
        Loop width * height
            if (NumGet(pixels, A_Index * 4 - 1, "UChar"))
                return
        mask := hbmMask ? IconCache._Bits(hbmMask, width, height) : 0
        Loop width * height {
            offset := A_Index * 4 - 4
            NumPut("UChar", (mask && NumGet(mask, offset, "UInt") & 0xFFFFFF) ? 0 : 255, pixels, offset + 3)
        }
    }

    ; Bitmap -> top-down 32-bit pixels (BGRA), 0 on failure
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
