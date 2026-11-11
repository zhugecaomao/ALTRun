;===============================================================================
; SafeFile.ahk - writing a file so that a crash or power cut halfway can't damage it, and reading one that is busy (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Writes a temporary file in the same folder, flushes it to the disk, then replaces the original (MoveFileEx with
; MOVEFILE_WRITE_THROUGH: the rename is on the disk too before it returns). Without the flush, a power cut or blue screen right
; after saving could leave the new file zero-filled once the old one is gone. Creates the folder if needed; UTF-8 without BOM.
; Sync folders such as OneDrive and antivirus software hold a just-changed file for a moment: on failure retry every 100 ms,
; up to MoveTries times, then throw (the original is kept, the temporary file deleted).
;
; Writes of the same file can overlap: while one waits to retry (Sleep), a timer or hotkey thread may write the file again. So
; each write has a temporary file of its own (<file>.<process id>.<count>.tmp; one shared name let the second write delete the
; first one's, which then failed after 2 s of retries), and a write that finds a later write of the same file already done
; keeps that newer content and drops its own instead of putting older content over it. A temporary file that has gone
; (ERROR_FILE_NOT_FOUND / ERROR_PATH_NOT_FOUND) isn't retried: it can't come back
;
; Usage:
;   SafeFile.WriteText("x.txt", text)
;   SafeFile.Write("x.bin", (out) => out.RawWrite(buf, buf.Size))   writer gets the File object (opened as UTF-8 without BOM)
;   text := SafeFile.ReadText("x.txt")       tries again while the file is busy, then throws; never "" for a file that couldn't be read
;===============================================================================
class SafeFile {
    static MoveTries := 20
    static _count := 0, _saved := Map()                                    ; writes started; file -> number of the last write that replaced it

    static WriteText(file, text) => SafeFile.Write(file, (out) => out.Write(text))

    ; Reads a data file. The same sync folders and antivirus software, or a program writing it, can hold the file for a moment: on failure
    ; it is read again every ReadWaitMs, up to tries times, then the error is thrown (at once when the file isn't there). A caller must
    ; not take a file that couldn't be read for an empty or damaged one: it would write new content over a good file
    static ReadTries := 10, ReadWaitMs := 200
    static readFn := (file, encoding) => FileRead(file, encoding)           ; replaced in the tests

    static ReadText(file, encoding := "UTF-8", tries := SafeFile.ReadTries) {
        Loop tries {
            try return SafeFile.readFn.Call(file, encoding)
            catch OSError as e {
                if (A_Index >= tries || e.Number = 2 || e.Number = 3)       ; ERROR_FILE_NOT_FOUND / ERROR_PATH_NOT_FOUND: waiting doesn't help
                    throw e
            }
            Sleep(SafeFile.ReadWaitMs)
        }
    }

    ; writer(out) writes the content into out, a File object open for writing (UTF-8 without BOM for text; RawWrite for binary data)
    static Write(file, writer) {
        SplitPath(file, , &dir)
        if (dir != "")
            DirCreate(dir)
        seq := ++SafeFile._count, key := StrLower(file)
        tmpFile := file "." DllCall("GetCurrentProcessId", "UInt") "." seq ".tmp"
        if !(out := FileOpen(tmpFile, "w", "UTF-8-RAW"))
            throw OSError(A_LastError, -1, tmpFile)
        try {
            writer(out)
            DllCall("FlushFileBuffers", "Ptr", out.Handle)                  ; reading Handle writes AutoHotkey's buffer to the file; then Windows' cache to the disk
        } catch as e {
            out.Close()
            try FileDelete(tmpFile)
            throw e
        }
        out.Close()
        Loop SafeFile.MoveTries {
            if (SafeFile._saved.Get(key, 0) > seq) {                       ; a later write of this file is already on the disk
                try FileDelete(tmpFile)
                return
            }
            if DllCall("MoveFileEx", "Str", tmpFile, "Str", file, "UInt", 0x9) {  ; MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH
                SafeFile._saved[key] := seq
                return
            }
            code := A_LastError
            if (A_Index = SafeFile.MoveTries || code = 2 || code = 3) {
                try FileDelete(tmpFile)
                throw OSError(code, -1, file)
            }
            Sleep(100)
        }
    }
}
