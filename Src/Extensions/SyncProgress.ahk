;===============================================================================
; SyncProgress.ahk - what robocopy has done so far, read from its log while it is written (AutoHotkey v2)
;-------------------------------------------------------------------------------
; robocopy writes its log as it goes, also with /MT (probed on Windows: the log can be read while robocopy writes it, and it
; grows line by line). A SyncProgress reads the part written since the last Read, from Start on: the log's size when the task
; started, since a run appends to the day's log. It keeps:
;   Changes   the first Limit files and folders copied / deleted, [{Kind: "Copy" | "Delete", Folder, Path, Label, Size}...]
;   Total     how many there are in all; CopyFiles / CopyBytes, NewDirs, DeleteFiles / DeleteDirs
;   Errors    robocopy's ERROR lines (a file that couldn't be copied after the retries), LastError the last one
; Log lines are tab-separated: a folder line ends with its full path ("\t  New Dir  2\tC:\src\sub\"; an unchanged folder only
; has the count) and the file lines after it give the name ("\t    New File  \t\t  1024\ta.txt"). With /MT (several threads, the
; default) there are no folder lines (so no New Dir lines either) and each file line has the full path. robocopy writes the labels (New File, Newer,
; *EXTRA File...) in Windows' language; the ones it deletes start with "*" (*EXTRA). Sizes are bytes ("966803") or a number with
; a unit of 1024 ("400.0 m", "1.5 g"): ParseSize; FormatSize writes them for people ("966.8 KB").
;
; Usage:
;   progress := SyncProgress(logFile, start)    progress.Read()  -> the changes kept since the last Read
;   progress := SyncProgress(), progress.Feed(text, true)        -> text from elsewhere (a whole log; true: nothing more follows)
;   SyncProgress.ParseSize("1.5 m") -> 1572864     SyncProgress.FormatSize(1572864) -> "1.5 MB"
;===============================================================================

class SyncProgress {
    __New(logFile := "", start := 0, limit := 2000) {
        this.LogFile := logFile, this.Pos := start, this.Limit := limit
        this.Changes := [], this.Total := 0
        this.CopyFiles := 0, this.CopyBytes := 0, this.NewDirs := 0, this.DeleteFiles := 0, this.DeleteDirs := 0
        this.Errors := 0, this.LastError := ""
        this._rest := "", this._folder := ""
    }

    ; The changes written since the last call (only the ones kept: at most Limit in all). The log is UTF-16 (/UNILOG)
    Read() {
        if (this.LogFile = "" || !FileExist(this.LogFile))
            return []
        try {
            handle := FileOpen(this.LogFile, "r", "UTF-16")                ; opened for reading only while robocopy writes it (it allows that)
            if (this.Pos > handle.Pos)                                      ; after the byte order mark when there is one
                handle.Pos := this.Pos
            text := handle.Read()
            this.Pos := handle.Pos
            handle.Close()
        } catch
            return []
        return this.Feed(LTrim(text, Chr(0xFEFF)))
    }

    ; Text of the log in order; the last line waits for the rest of it unless final
    Feed(text, final := false) {
        lines := StrSplit(this._rest StrReplace(text, "`r"), "`n")
        if !lines.Length                                                    ; nothing new (StrSplit of "" has no items)
            return []
        this._rest := final ? "" : lines.Pop()
        kept := []
        for line in lines {
            if !IsObject(change := this._Line(line))
                continue
            this.Total++
            if (change.Folder && change.Kind = "Delete")
                this.DeleteDirs++
            else if (!change.Folder && change.Kind = "Delete")
                this.DeleteFiles++
            else if change.Folder
                this.NewDirs++
            else if !change.Folder
                this.CopyFiles++, this.CopyBytes += change.Size
            if (this.Changes.Length < this.Limit)
                this.Changes.Push(change), kept.Push(change)
        }
        return kept
    }

    ; One log line -> a change, or "" (a header, an unchanged folder, a summary row...)
    _Line(line) {
        if RegExMatch(line, "\sERROR \d+ \(0x[0-9A-Fa-f]+\)") {               ; "2026/11/03 10:00:00 ERROR 5 (0x00000005) Copying File ..."
            this.Errors++, this.LastError := Trim(line)
            return ""
        }
        if !InStr(line, "`t")
            return ""
        fields := StrSplit(line, "`t")
        name := fields[fields.Length]
        if (Trim(name) = "")
            return ""
        if (RegExMatch(name, "^(?:[A-Za-z]:\\|\\\\)") && SubStr(name, -1) = "\") {   ; a folder: C:\x\ or \\server\share\x\
            this._folder := name
            label := Trim(RegExReplace(fields.Length >= 2 ? fields[fields.Length - 1] : "", "\s*-?\d+\s*$"))
            if (label = "")
                return ""                                                   ; unchanged: only the folder the next files are in
            return {Kind: SubStr(label, 1, 1) = "*" ? "Delete" : "Copy", Folder: true, Path: name, Label: label, Size: ""}
        }
        label := "", size := 0
        Loop fields.Length - 1 {
            field := fields[A_Index]
            if (label = "" && RegExMatch(field, "[^\d\s.,]{2,}"))           ; the label, not a size such as "1.5 m"
                label := Trim(field)
            else if (Trim(field) != "" && (bytes := SyncProgress.ParseSize(field)) != "")
                size := bytes
        }
        full := RegExMatch(name, "^(?:[A-Za-z]:\\|\\\\)")                     ; /MT: the file's full path
        if (label = "" || (!full && this._folder = ""))
            return ""
        return {Kind: SubStr(label, 1, 1) = "*" ? "Delete" : "Copy", Folder: false, Path: full ? name : this._folder name, Label: label, Size: size}
    }

    ; robocopy's size: bytes, or a number with k / m / g / t (units of 1024) -> bytes; "" when it isn't a size
    static ParseSize(text) {
        if !RegExMatch(Trim(text), "i)^(\d+(?:\.\d+)?)(?:\s*([kmgt]))?$", &m)
            return ""
        return Round(m[1] * (1024 ** (m[2] = "" ? 0 : InStr("kmgt", m[2]))))
    }

    ; Bytes for people: "512 B", "966.8 KB", "1.5 MB", "38.6 GB"
    static FormatSize(bytes) {
        if (bytes < 1024)
            return Round(bytes) " B"
        for unit in ["KB", "MB", "GB", "TB"] {
            bytes /= 1024
            if (bytes < 1024 || unit = "TB")
                return Format("{:.1f} {}", bytes, unit)
        }
    }
}
