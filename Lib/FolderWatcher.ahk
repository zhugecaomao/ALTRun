;===============================================================================
; FolderWatcher.ahk - watches folders (with subfolders) for file changes (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Reads with ReadDirectoryChangesW asynchronously; the caller calls Poll() on a timer (no waiting) to get the full paths changed since
; the last time (files / folders added, deleted, renamed; files modified). "*" when it isn't known which: the buffer overflowed (too many at once),
; a read failed, or a folder couldn't be watched for a while and is back (e.g. deleted and created again; retried every RetryMs).
; Uses almost no resources while idle; Wine doesn't support subfolders (only the first level is reported).
;
; Usage:
;   watcher := FolderWatcher(["C:\Users\me\Desktop", "D:\Docs"])   folders that can't be opened are skipped
;   FolderWatcher(folders, FolderWatcher.NameChanges)              only additions, deletions and renames, no modifications
;   watcher.Count                                                  how many folders are watched
;   for fullPath in watcher.Poll()                                 "C:\Users\me\Desktop\a.lnk" or "*"
;   watcher.Stop()                                                 (also stops when the object is released)
;===============================================================================

class FolderWatcher {
    static RetryMs := 30000
    static NameChanges := 0x1 | 0x2, AllChanges := 0x1 | 0x2 | 0x10    ; file names, folder names (| modified time)

    __New(folders, filter := FolderWatcher.AllChanges) {
        this.Watches := []
        for folder in folders
            if IsObject(watch := FolderWatcher.Create(folder, filter))
                this.Watches.Push(watch)
    }

    Count => this.Watches.Length

    ; The changes that have arrived (no waiting); each folder starts waiting for the next batch right after it is read
    Poll() {
        changes := []
        for index, watch in this.Watches {
            if watch.Disabled {                                             ; couldn't start waiting again last time: reopen when it is time
                if (A_TickCount < watch.RetryAt)
                    continue
                if IsObject(replacement := FolderWatcher.Create(watch.Folder, watch.Filter)) {
                    FolderWatcher.Close(watch)
                    this.Watches[index] := replacement
                    changes.Push("*")                                       ; changes may have been missed while it was stopped
                } else {
                    watch.RetryAt := A_TickCount + FolderWatcher.RetryMs
                }
                continue
            }
            if (DllCall("WaitForSingleObject", "Ptr", watch.Event, "UInt", 0, "UInt") != 0)
                continue
            watch.Pending := false                                          ; signaled = this read has completed
            bytes := 0
            if !DllCall("GetOverlappedResult", "Ptr", watch.Handle, "Ptr", watch.Overlapped, "UInt*", &bytes, "Int", false)
                changes.Push("*")                                           ; the read failed: changes may have been missed
            else if !bytes                                                  ; the buffer overflowed: unknown which, treat it as "everything changed"
                changes.Push("*")
            offset := 0
            while (bytes && offset < bytes) {                               ; FILE_NOTIFY_INFORMATION list
                next := NumGet(watch.Buffer, offset, "UInt"), length := NumGet(watch.Buffer, offset + 8, "UInt")
                changes.Push(watch.Folder "\" StrGet(watch.Buffer.Ptr + offset + 12, length // 2, "UTF-16"))
                if !next
                    break
                offset += next
            }
            if !FolderWatcher._Read(watch) {
                watch.Disabled := true, watch.RetryAt := A_TickCount + FolderWatcher.RetryMs
                changes.Push("*")
            }
        }
        return changes
    }

    Stop() {
        for watch in this.Watches
            FolderWatcher.Close(watch)
        this.Watches := []
    }

    __Delete() => this.Stop()

    ; Opens a folder and starts waiting for changes; "" if it can't be opened
    static Create(folder, filter := FolderWatcher.AllChanges) {
        static FILE_LIST_DIRECTORY := 0x1, SHARE_ALL := 0x7, OPEN_EXISTING := 3, BACKUP_SEMANTICS := 0x02000000, OVERLAPPED := 0x40000000
        folder := RTrim(folder, "\")
        if (folder = "" || !DirExist(folder))
            return ""
        handle := DllCall("CreateFileW", "Str", folder, "UInt", FILE_LIST_DIRECTORY, "UInt", SHARE_ALL, "Ptr", 0, "UInt", OPEN_EXISTING
            , "UInt", BACKUP_SEMANTICS | OVERLAPPED, "Ptr", 0, "Ptr")
        if (handle = -1 || !handle)
            return ""
        watch := {Folder: folder, Handle: handle, Event: DllCall("CreateEventW", "Ptr", 0, "Int", true, "Int", false, "Ptr", 0, "Ptr")
            , Overlapped: Buffer(A_PtrSize * 2 + 8 + A_PtrSize, 0), Buffer: Buffer(32768), Filter: filter, Pending: false, Disabled: false, RetryAt: 0}
        if !watch.Event {
            DllCall("CloseHandle", "Ptr", handle)
            return ""
        }
        NumPut("Ptr", watch.Event, watch.Overlapped, A_PtrSize * 2 + 8)    ; OVERLAPPED.hEvent
        if !FolderWatcher._Read(watch) {
            FolderWatcher.Close(watch)
            return ""
        }
        return watch
    }

    ; Cancels a pending read and waits until it has really ended (CancelIoEx only requests it); only then can the OVERLAPPED and buffer be freed
    static Close(watch) {
        if watch.Pending {
            DllCall("CancelIoEx", "Ptr", watch.Handle, "Ptr", watch.Overlapped)
            bytes := 0
            DllCall("GetOverlappedResult", "Ptr", watch.Handle, "Ptr", watch.Overlapped, "UInt*", &bytes, "Int", true)
            watch.Pending := false
        }
        if watch.Handle
            DllCall("CloseHandle", "Ptr", watch.Handle), watch.Handle := 0
        if watch.Event
            DllCall("CloseHandle", "Ptr", watch.Event), watch.Event := 0
    }

    ; Starts (again) waiting asynchronously for the next batch of changes in this folder (with subfolders)
    static _Read(watch) {
        DllCall("ResetEvent", "Ptr", watch.Event)
        watch.Pending := DllCall("ReadDirectoryChangesW", "Ptr", watch.Handle, "Ptr", watch.Buffer, "UInt", watch.Buffer.Size, "Int", true
            , "UInt", watch.Filter, "Ptr", 0, "Ptr", watch.Overlapped, "Ptr", 0) ? true : false
        return watch.Pending
    }
}
