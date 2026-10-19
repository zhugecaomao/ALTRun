;===============================================================================
; FolderWatcher.ahk - 监视文件夹 (包括子文件夹) 里的文件变化 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; ReadDirectoryChangesW 异步读取, 调用方定时 Poll() 看一眼 (不等待): 返回上次以来变化了的完整路径
; (文件 / 文件夹的新增、删除、改名, 文件的修改)。不知道具体是哪些时返回 "*": 缓冲区放不下 (一下子变化太多)、
; 读取出错、某个文件夹暂时监视不了又恢复了 (例如文件夹被删掉又建回来; 每 RetryMs 重试一次)。
; 平时几乎不占资源; Wine 不支持子文件夹 (只报告第一层)。
;
; 用法:
;   watcher := FolderWatcher(["C:\Users\me\Desktop", "D:\Docs"])   打不开的文件夹跳过
;   FolderWatcher(folders, FolderWatcher.NameChanges)              只要新增、删除、改名, 不要修改
;   watcher.Count                                                  在监视几个文件夹
;   for fullPath in watcher.Poll()                                 "C:\Users\me\Desktop\a.lnk" 或 "*"
;   watcher.Stop()                                                 (对象释放时也会停)
;===============================================================================

class FolderWatcher {
    static RetryMs := 30000
    static NameChanges := 0x1 | 0x2, AllChanges := 0x1 | 0x2 | 0x10    ; 文件名、文件夹名 (| 修改时间)

    __New(folders, filter := FolderWatcher.AllChanges) {
        this.Watches := []
        for folder in folders
            if IsObject(watch := FolderWatcher.Create(folder, filter))
                this.Watches.Push(watch)
    }

    Count => this.Watches.Length

    ; 已经到了的变化 (不等待); 每个文件夹读完后马上开始等下一批
    Poll() {
        changes := []
        for index, watch in this.Watches {
            if watch.Disabled {                                             ; 上次没能再开始等: 到时间就重新打开
                if (A_TickCount < watch.RetryAt)
                    continue
                if IsObject(replacement := FolderWatcher.Create(watch.Folder, watch.Filter)) {
                    FolderWatcher.Close(watch)
                    this.Watches[index] := replacement
                    changes.Push("*")                                       ; 停掉的这段时间可能漏了变化
                } else {
                    watch.RetryAt := A_TickCount + FolderWatcher.RetryMs
                }
                continue
            }
            if (DllCall("WaitForSingleObject", "Ptr", watch.Event, "UInt", 0, "UInt") != 0)
                continue
            watch.Pending := false                                          ; 有信号 = 这次读取已经完成
            bytes := 0
            if !DllCall("GetOverlappedResult", "Ptr", watch.Handle, "Ptr", watch.Overlapped, "UInt*", &bytes, "Int", false)
                changes.Push("*")                                           ; 读取出错: 可能漏了变化
            else if !bytes                                                  ; 缓冲区放不下: 不知道是哪些, 当作 "都变了"
                changes.Push("*")
            offset := 0
            while (bytes && offset < bytes) {                               ; FILE_NOTIFY_INFORMATION 链表
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

    ; 打开一个文件夹并开始等变化; 打不开返回 ""
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

    ; 还在等的读取先取消并等它真的结束 (CancelIoEx 只是请求取消), 之后才能放掉 OVERLAPPED 和缓冲区
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

    ; 开始 (再次) 异步等这个文件夹 (包括子文件夹) 的下一批变化
    static _Read(watch) {
        DllCall("ResetEvent", "Ptr", watch.Event)
        watch.Pending := DllCall("ReadDirectoryChangesW", "Ptr", watch.Handle, "Ptr", watch.Buffer, "UInt", watch.Buffer.Size, "Int", true
            , "UInt", watch.Filter, "Ptr", 0, "Ptr", watch.Overlapped, "Ptr", 0) ? true : false
        return watch.Pending
    }
}
