;===============================================================================
; FileIndexBench.ahk - times the built-in file index on 64,000 made-up files (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Run by Tests\RunTests.ahk FileIndexBench (on Windows CI: the File index benchmark workflow, also on pull requests that change the
; index). Creates 64,000 empty files in 400 folders under %TEMP% (one in ten with a Chinese name), then times a full rescan as ALTRun
; does it (in the background, in time slices), writing and reading the cache, and searches with few and many matches. Prints one
; line per measurement and deletes the files again
;===============================================================================
class FileIndexBench {
    static Run() {
        out(text) => FileAppend(text "`n", "*")
        root := A_Temp "\ALTRunIndexBench"
        try DirDelete(root, true)
        made := FileIndexBench.MakeFiles(root)
        options := AppSettings.Feature("FileSearch")
        saved := [options["ScopeFolders"], options["ScopeDepth"], options["MaxEntries"], options["UseEverything"], FileIndex.File]
        options["ScopeFolders"] := [root], options["ScopeDepth"] := 6, options["MaxEntries"] := 200000, options["UseEverything"] := 0
        FileIndex.File := A_Temp "\ALTRunIndexBench.bin"
        try {
            state := {Done: 0}, start := Logger.Ms()
            FileIndex.Rebuild((count) => (state.Done := count))
            while !state.Done
                Sleep(5)
            out(Format("rescan (in the background, as in ALTRun): {:.0f} ms, {} entries ({} files made)", Logger.Ms() - start, state.Done, made))
            start := Logger.Ms()
            FileIndex._SaveCache()
            out(Format("cache write: {:.0f} ms, {} KB", Logger.Ms() - start, FileGetSize(FileIndex.File) // 1024))
            start := Logger.Ms()
            loaded := FileIndex._LoadCache()
            out(Format("cache read: {:.0f} ms (read {}, {} entries, {} with pinyin initials)", Logger.Ms() - start, loaded, FileIndex.Count, FileIndex.Initials.Count))
            for needle in ["re", "report", "udge", "1234", "plan 12", "xyz", "bg"] {
                count := 0, start := Logger.Ms()
                Loop 5
                    count := FileIndex.Search(needle, 30).Length
                scores := Map(), matches := []
                FileIndex.Data.Find(needle, scores, matches)
                out(Format("search [{}]: {:.1f} ms ({} names match, {} shown)", needle, (Logger.Ms() - start) / 5, matches.Length, count))
            }
        } finally {
            options["ScopeFolders"] := saved[1], options["ScopeDepth"] := saved[2], options["MaxEntries"] := saved[3], options["UseEverything"] := saved[4]
            try FileDelete(FileIndex.File)
            FileIndex.File := saved[5]
            FileIndex.SetEntries([])
            try DirDelete(root, true)
        }
    }

    ; 20 x 20 folders of 160 empty files each, made-up names (budget_report 10203.pdf, 测试报告 10210.docx...)
    static MakeFiles(root) {
        words := ["budget", "report", "drawing", "plan", "notes", "invoice", "tender", "design", "section", "level", "schedule", "minutes"
            , "photo", "contract", "summary", "meeting", "letter", "calc", "model", "review"]
        chinese := ["测试", "报告", "图纸", "会议", "合同", "预算", "设计", "计算"]
        exts := [".pdf", ".docx", ".xlsx", ".dwg", ".jpg", ".txt", ".pptx", ".zip"]
        made := 0
        Loop 20 {
            a := A_Index
            Loop 20 {
                b := A_Index
                dir := root "\Project " Format("{:02}", a) " " words[a] "\Sub " Format("{:02}", b) " " words[Mod(a + b, 20) + 1]
                DirCreate(dir)
                Loop 160 {
                    k := a * 10000 + b * 200 + A_Index
                    name := Mod(k, 10) = 0 ? chinese[Mod(k, 8) + 1] chinese[Mod(k * 3, 8) + 1] " " k : words[Mod(k, 20) + 1] "_" words[Mod(k * 7, 20) + 1] " " k
                    handle := DllCall("CreateFileW", "Str", dir "\" name exts[Mod(k, 8) + 1], "UInt", 0x40000000, "UInt", 0, "Ptr", 0, "UInt", 1, "UInt", 0x80, "Ptr", 0, "Ptr")   ; GENERIC_WRITE, CREATE_NEW
                    if (handle != -1)
                        DllCall("CloseHandle", "Ptr", handle), made++
                }
            }
        }
        return made
    }
}
