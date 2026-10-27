;===============================================================================
; SafeFile.ahk - writing a file so that a crash or power cut halfway can't damage it (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Writes a temporary file in the same folder, then replaces the original. Creates the folder if needed; UTF-8 without BOM.
; Sync folders such as OneDrive and antivirus software hold a just-changed file for a moment: on failure retry every 100 ms,
; up to MoveTries times, then throw (the original is kept, the temporary file deleted)
;
; Usage:
;   SafeFile.WriteText("x.txt", text)
;===============================================================================
class SafeFile {
    static MoveTries := 20

    static WriteText(file, text) {
        SplitPath(file, , &dir)
        if (dir != "")
            DirCreate(dir)
        tmpFile := file ".tmp"
        try FileDelete(tmpFile)
        if FileExist(tmpFile)                                               ; a temporary file left over from last time can't be deleted (in use): use another name
            tmpFile := file "." A_TickCount ".tmp"
        FileAppend(text, tmpFile, "UTF-8-RAW")
        Loop SafeFile.MoveTries {
            try {
                FileMove(tmpFile, file, true)
                return
            } catch as e {
                if (A_Index = SafeFile.MoveTries) {
                    try FileDelete(tmpFile)
                    throw e
                }
                Sleep(100)
            }
        }
    }
}
