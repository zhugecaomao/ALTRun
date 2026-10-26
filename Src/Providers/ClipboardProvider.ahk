;===============================================================================
; ClipboardProvider.ahk - clipboard history (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like Alfred's Clipboard History: remembers copied text, files and images; "clip" (Keyword)
; or the hotkey (default Ctrl+Alt+C) lists them, "clip word" filters, Enter pastes into the foreground window
; (files paste back as files, images as images).
;
; Privacy:
;   - content copied by password managers and the like (marked with ExcludeClipboardContentFromMonitorProcessing
;     or CanIncludeInClipboardHistory = 0) is never recorded
;   - content copied by programs in IgnoreApps (process names) is never recorded
;   - with Persist = 0 it is kept in memory only and cleared on exit (images aren't recorded then)
;
; Storage: ClipboardHistory.json; entries longer than LargeText characters are stored separately as Clipboard\*.txt,
; images as Clipboard\img-*.png, with only the file name in the JSON. It is saved on every copy, so very long text
; isn't rewritten each time (up to 200 entries x 100,000 characters).
;
; Pinning ("Pin" in the actions panel): the entry gets "Pinned": 1, is listed first, isn't removed when over the limit and survives "Clear".
;
; Entries: Map("Text", "Time", "App") is text; the two other kinds have "Type":
;   "files"  copied files: "Files" [paths...], "Text" has one path per line (for searching)
;   "image"  an image: "Image" file name, "Width" / "Height"
;
; Settings (ALTRun.json -> Features.Clipboard):
;   Enabled / Keyword / Hotkey / MaxItems / MaxItemLength / Persist / IgnoreApps
;   Images / MaxImages     record images, how many to keep at most
;   MergeDoubleCopy        press Ctrl+C twice quickly: append the copied text to the previous entry (off by default)
;   LocalFiles             where it is stored: 0 = all in the Data folder; 1 = images and long entries (Clipboard folder) only on this PC (default);
;                          2 = the history (ClipboardHistory.json) too. This PC = %LOCALAPPDATA%\ALTRun, see PlaceStorage
;
; Usage:
;   ClipboardProvider.Add(text, source)    adds an entry by hand (source = process name of the program it came from)
;   ClipboardProvider.AddFiles(paths, source) / AddImage(source)
;   ClipboardProvider.TextAt(n)            text of entry n (1 = newest), for the {clipboard:N} of snippets
;===============================================================================

class ClipboardProvider {
    static Id      := "Clipboard"
    static File    := AppSettings.DataDir "\ClipboardHistory.json"
    static Folder  := AppSettings.DataDir "\Clipboard"                         ; very long entries and images
    static LocalDir := EnvGet("LOCALAPPDATA") != "" ? EnvGet("LOCALAPPDATA") "\ALTRun" : ""   ; stored here with LocalFiles = 1 (default) / 2
    static LargeText := 4000
    static Icon => IconCache.Own("Clipboard", "res:imageres.dll,-5314")   ; icon of text entries
    static MergeWindow := 400                                               ; at most this many milliseconds between two Ctrl+C count as a double press
    static Entries := []              ; newest first, see the top of the file
    static ReadDelay := 150                                                 ; how long to wait after the clipboard changed before reading it (milliseconds)
    static _saveTimer := "", _mergeUntil := 0, _pending := "", _readTimer := ""

    static Init() {
        options := AppSettings.Feature("Clipboard")
        ClipboardProvider.PlaceStorage(options["LocalFiles"])
        if options["Persist"]
            ClipboardProvider._Load()
        OnClipboardChange(Logger.Watch("clipboard change", (dataType) => ClipboardProvider._OnChange(dataType)))
        if options["MergeDoubleCopy"] {
            try Hotkey("~^c", (*) => ClipboardProvider._OnCopyKey())         ; ~ : Ctrl+C copies as usual
            catch as e
                Logger.Error("ClipboardProvider: cannot register Ctrl+C - " e.Message)
        }
    }

    ; Where the clipboard history is stored (LocalFiles): on this PC it doesn't travel with the Data folder into OneDrive and other sync folders or a USB drive.
    ;   1 (default): images and very long entries (Clipboard folder) on this PC: images are big, and in a sync folder every copy would upload one.
    ;      ClipboardHistory.json stays in Data (the text history syncs as usual); on another PC images / long entries without their files are skipped when read
    ;   2: ClipboardHistory.json on this PC too: the clipboard history never leaves this PC
    ;   0: all in Data (e.g. the whole folder carried on a USB drive)
    ; After the setting changes, the files are moved from the other location at the first start-up; if they can't be moved, the old location is used this time and it tries again next time.
    ; When both exist (e.g. another PC's history in a sync folder) nothing is overwritten and the new location's is used
    static PlaceStorage(level, dataDir := AppSettings.DataDir, localDir := ClipboardProvider.LocalDir) {
        if (localDir = "" || localDir = dataDir)
            level := 0
        folderDir := (level >= 1) ? localDir : dataDir
        fileDir := (level >= 2) ? localDir : dataDir
        otherDir := (folderDir = dataDir) ? localDir : dataDir
        ClipboardProvider.Folder := (otherDir = "") ? folderDir "\Clipboard" : ClipboardProvider._Relocate(otherDir "\Clipboard", folderDir "\Clipboard", true)
        otherDir := (fileDir = dataDir) ? localDir : dataDir
        ClipboardProvider.File := (otherDir = "") ? fileDir "\ClipboardHistory.json" : ClipboardProvider._Relocate(otherDir "\ClipboardHistory.json", fileDir "\ClipboardHistory.json", false)
    }

    ; Moves source (folder or file) to target and returns the location to use: target when already there / source doesn't exist / both exist, source if it can't be moved
    static _Relocate(source, target, isFolder) {
        exists := (where) => isFolder ? InStr(FileExist(where), "D") : (FileExist(where) != "" && !InStr(FileExist(where), "D"))
        if (source = target || !exists(source))
            return target
        if exists(target) {
            Logger.Debug("ClipboardProvider: both " source " and " target " exist, using " target)
            return target
        }
        try {
            SplitPath(target, , &dir)
            DirCreate(dir)
            if isFolder {
                DirCopy(source, target)
                DirDelete(source, true)
            } else {
                FileMove(source, target)
            }
            Logger.Debug("ClipboardProvider: moved " source " to " target)
            return target
        } catch as e {
            Logger.Error("ClipboardProvider: cannot move " source " to " target " - " e.Message)
            return source
        }
    }

    static Search(query) {
        options := AppSettings.Feature("Clipboard")
        if !query.MatchKeyword([options["Keyword"]], &term) {
            full := ""
            if query.StartsKeyword([options["Keyword"]], &full)            ; "cli": the keyword's hint at the end of the list ("clip" alone lists the history)
                return [ResultItem.KeywordHint(I18n.T("Clipboard.Keyword"), full, ClipboardProvider.Icon, true)]
            return []
        }
        icon := ClipboardProvider.Icon
        exclusive := query.HasRest                                          ; after "clip " only the clipboard history is shown
        if !ClipboardProvider.Entries.Length
            return [ResultItem(I18n.T("Clipboard.Empty"), I18n.T("Clipboard.EmptyHint", Win.HotkeyLabel(options["Hotkey"])), {Icon: icon, Valid: false, Score: 300, Exclusive: exclusive})]

        results := []
        tokens := StrSplit(Trim(term), " ")
        for pass in [true, false] {                                         ; pinned first, the rest by time
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
        if (ClipboardProvider.TypeOf(entry) = "image")                      ; images: found by "image" and their size
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
                if (files.Length = 1) {                                     ; one file: like a file search result, with open / reveal and the other actions
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
        props.Pinned := pinned                                              ; pinned: a pin is drawn on the icon (SearchWindow)
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

    ; F3 / right-click "Edit": text is saved as a snippet, files are added as custom commands; images can't be edited
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

    ; Pastes a history entry into the foreground window and moves it to the top
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

    ; Puts files / an image on the clipboard for a moment, Ctrl+V, then restores it (like ActionCatalog.PasteText, not recorded meanwhile)
    static _PasteWith(setter) {
        delay := Max(AppSettings.Feature("Snippets")["PasteDelay"], 500)    ; Explorer reads the clipboard slowly when pasting files
        App.FocusPreviousWindow()
        Win.WaitModifiersUp()
        ClipboardData.Borrow(delay + 1000)
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
            entry["Pinned"] := 1                                            ; a pinned entry copied / pasted again stays pinned
        ClipboardProvider.Entries.InsertAt(1, entry)
        ClipboardProvider._Trim()
        ClipboardProvider._SaveLater()
        return true
    }

    static AddFiles(files, source := "") {
        if !files.Length
            return false
        joined := TextTools.Join(files)
        entry := Map("Type", "files", "Files", files, "Text", joined, "Time", A_Now, "App", source)
        if IsObject(previous := ClipboardProvider._RemoveSame("files", joined)) && ClipboardProvider.IsPinned(previous)
            entry["Pinned"] := 1
        ClipboardProvider.Entries.InsertAt(1, entry)
        ClipboardProvider._Trim()
        ClipboardProvider._SaveLater()
        return true
    }

    ; Saves the image on the clipboard as a PNG and adds it to the history; if it is identical to an existing one, that one just moves to the top
    static AddImage(source := "") {
        try DirCreate(ClipboardProvider.Folder)
        name := "img-" A_Now "-" Random(100000, 999999) ".png"
        imageFile := ClipboardProvider.Folder "\" name
        size := ClipboardData.SaveImage(imageFile)
        if !IsObject(size)
            return false
        return ClipboardProvider.AddImageFile(name, size.Width, size.Height, source)
    }

    ; Adds a PNG already saved in Data\Clipboard to the history (for AddImage and tests)
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

    static Clear() {                                                        ; pinned entries are kept
        kept := []
        for entry in ClipboardProvider.Entries
            if ClipboardProvider.IsPinned(entry)
                kept.Push(entry)
        ClipboardProvider.Entries := kept
        ClipboardProvider._SaveLater()
        App.Notify(I18n.T("Clipboard.Cleared"))
    }

    ; Text of entry n (1 = newest), skipping files and images; "" if none
    static TextAt(n) {
        for entry in ClipboardProvider.Entries
            if (ClipboardProvider.TypeOf(entry) = "text" && --n = 0)
                return entry["Text"]
        return ""
    }

    ; Removes entries with the same content and returns the removed one ("" if none)
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

    ; Over the limit, removes the oldest (not pinned ones); images have their own limit (MaxImages); removed image files are cleaned up when saving
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
    ; The clipboard changed: wait ReadDelay before reading. Many programs (screenshot tools etc.) open the clipboard again right after writing it,
    ; and reading at once would hold it and make their OpenClipboard fail. If it changes again while waiting: read the previous change first, then wait for this one
    ; (when merging a double Ctrl+C, the first copy has to be recorded first)
    static _OnChange(dataType) {
        if (dataType = 0 || ClipboardData.Borrowed())                      ; 0 = the clipboard was emptied
            return
        processName := ""
        try processName := WinGetProcessName("A")                          ; the foreground program at the time of copying
        if IsObject(ClipboardProvider._pending)
            ClipboardProvider._ReadPending()
        ClipboardProvider._pending := {Type: dataType, App: processName}
        if !IsObject(ClipboardProvider._readTimer)
            ClipboardProvider._readTimer := ObjBindMethod(ClipboardProvider, "_ReadPending")
        SetTimer(ClipboardProvider._readTimer, -ClipboardProvider.ReadDelay)
    }

    static _ReadPending() {
        pending := ClipboardProvider._pending
        ClipboardProvider._pending := ""
        if IsObject(ClipboardProvider._readTimer)
            SetTimer(ClipboardProvider._readTimer, 0)
        if (IsObject(pending) && !ClipboardData.Borrowed())                 ; ALTRun borrowed the clipboard meanwhile: it now holds temporary content, not recorded
            ClipboardProvider._Record(pending.Type, pending.App)
    }

    static _Record(dataType, processName) {
        if ClipboardProvider.IsPrivate()
            return
        options := AppSettings.Feature("Clipboard")
        for ignored in options["IgnoreApps"]
            if (processName = ignored)
                return
        if (dataType = 1) {                                                 ; 1 = text (copied files too; the content is their paths)
            if ClipboardData.HasFiles()
                return ClipboardProvider.AddFiles(ClipboardData.Files(), processName)
            text := ""
            try text := A_Clipboard
            if (A_TickCount < ClipboardProvider._mergeUntil && ClipboardProvider._Merge(text, processName))
                return
            return ClipboardProvider.Add(text, processName)
        }
        if (options["Images"] && options["Persist"] && ClipboardData.HasImage())   ; 2 = other formats; images are only recorded when the history is saved
            ClipboardProvider.AddImage(processName)
    }

    ; Double Ctrl+C: the text of the first copy is already the newest entry; when the second copy has the same text it is appended to the previous entry,
    ; and the clipboard gets the merged text too
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
        ClipboardData.Borrow(500)
        A_Clipboard := merged
        App.Notify(I18n.T("Clipboard.Merged"))
        return true
    }

    ; Password managers and the like mark clipboard content "don't record" (Windows' clipboard history respects it too; {query|clipboard} in web searches doesn't read it either)
    static IsPrivate() {
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
                        entry["Text"] := TextTools.Join(entry["Files"])
                    case "image":
                        if !(entry.Has("Image") && FileExist(ClipboardProvider.Folder "\" entry["Image"]))
                            continue
                        entry["Text"] := ""
                    default:
                        if entry.Has("File") {                              ; a very long entry: the text is in a separate file
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
                if (!entry.Has("File") || !FileExist(ClipboardProvider.Folder "\" entry["File"])) {   ; a new long entry: written only this once
                    DirCreate(ClipboardProvider.Folder)
                    entry["File"] := stamp "-" Random(100000, 999999) ".txt"
                    FileAppend(text, ClipboardProvider.Folder "\" entry["File"], "UTF-8")
                }
                keep[StrLower(entry["File"])] := true
                list.Push(ClipboardProvider._WithPin(entry, Map("File", entry["File"], "Time", stamp, "App", source)))
            }
            JSON.WriteFile(ClipboardProvider.File, Map("Entries", list))
            Loop Files, ClipboardProvider.Folder "\*.*" {                  ; Deletes long entries and images that are no longer in the history
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
            ClipboardProvider._saveTimer := Logger.Watch("save clipboard", () => ClipboardProvider.Save())
        SetTimer(ClipboardProvider._saveTimer, -2000)
    }
}
