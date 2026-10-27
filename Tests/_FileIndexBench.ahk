;===============================================================================
; RunTests.ahk - ALTRun unit tests (AutoHotkey v2)
;-------------------------------------------------------------------------------
; The tests are in Tests\Suites (one class per area); this file loads the code,
; sets up in-memory settings and runs every suite. Run:
;   AutoHotkey64.exe /ErrorStdOut Tests\RunTests.ahk
; Results go to standard output; the exit code is the number of failures (0 = all passed).
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut

#Include %A_ScriptDir%\..\Lib\SafeFile.ahk
#Include %A_ScriptDir%\..\Lib\JSON.ahk
#Include %A_ScriptDir%\..\Lib\Logger.ahk
#Include %A_ScriptDir%\..\Lib\Util.ahk
#Include %A_ScriptDir%\..\Lib\TextTools.ahk
#Include %A_ScriptDir%\..\Lib\DelimitedText.ahk
#Include %A_ScriptDir%\..\Lib\Kanji.ahk
#Include %A_ScriptDir%\..\Lib\Everything.ahk
#Include %A_ScriptDir%\..\Lib\Units.ahk
#Include %A_ScriptDir%\..\Lib\ClipboardData.ahk
#Include %A_ScriptDir%\..\Lib\Dialogs.ahk
#Include %A_ScriptDir%\..\Lib\HttpGet.ahk
#Include %A_ScriptDir%\..\Lib\FolderWatcher.ahk
#Include %A_ScriptDir%\..\Lib\WinTimer.ahk
#Include %A_ScriptDir%\..\Lib\WindowsRecent.ahk
#Include %A_ScriptDir%\..\Src\Core\App.ahk
#Include %A_ScriptDir%\..\Src\Core\I18n.ahk
#Include %A_ScriptDir%\..\Src\Core\AppSettings.ahk
#Include %A_ScriptDir%\..\Src\Core\SchemaMigration.ahk
#Include %A_ScriptDir%\..\Src\Core\SearchQuery.ahk
#Include %A_ScriptDir%\..\Src\Core\ResultItem.ahk
#Include %A_ScriptDir%\..\Src\Core\FuzzyMatcher.ahk
#Include %A_ScriptDir%\..\Src\Core\BackgroundWork.ahk
#Include %A_ScriptDir%\..\Src\Core\Knowledge.ahk
#Include %A_ScriptDir%\..\Src\Core\Usage.ahk
#Include %A_ScriptDir%\..\Src\Core\ActionCatalog.ahk
#Include %A_ScriptDir%\..\Src\Core\ProviderRegistry.ahk
#Include %A_ScriptDir%\..\Src\Core\FileIndex.ahk
#Include %A_ScriptDir%\..\Src\UI\ThemeManager.ahk
#Include %A_ScriptDir%\..\Src\UI\ThemePreview.ahk
#Include %A_ScriptDir%\..\Src\UI\NavIcons.ahk
#Include %A_ScriptDir%\..\Src\UI\IconCache.ahk
#Include %A_ScriptDir%\..\Src\UI\SearchWindow.ahk
#Include %A_ScriptDir%\..\Src\UI\ResultRows.ahk
#Include %A_ScriptDir%\..\Src\UI\SearchStatusBar.ahk
#Include %A_ScriptDir%\..\Src\UI\TypeAhead.ahk
#Include %A_ScriptDir%\..\Src\UI\SearchInputLanguage.ahk
#Include %A_ScriptDir%\..\Src\UI\SearchKeys.ahk
#Include %A_ScriptDir%\..\Src\UI\ResultActions.ahk
#Include %A_ScriptDir%\..\Src\UI\WindowFade.ahk
#Include %A_ScriptDir%\..\Src\UI\LargeType.ahk
#Include %A_ScriptDir%\..\Src\UI\Hud.ahk
#Include %A_ScriptDir%\..\Src\UI\ItemEditor.ahk
#Include %A_ScriptDir%\..\Src\UI\HotkeyBox.ahk
#Include %A_ScriptDir%\..\Src\UI\PreferencesWindow.ahk
#Include %A_ScriptDir%\..\Src\UI\PrefsForm.ahk
#Include %A_ScriptDir%\..\Src\UI\PrefsPages.ahk
#Include %A_ScriptDir%\..\Src\Providers\ClipboardProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\ApplicationProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\CustomCommandProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\SnippetProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\SystemProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\CalculatorProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\WebSearchProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\BookmarkProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\FileSearchProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\TerminalProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\WindowProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\RecentProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\ScriptProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\HelpProvider.ahk
#Include %A_ScriptDir%\..\Src\Extensions\SnippetExpander.ahk
#Include %A_ScriptDir%\..\Src\Extensions\QuickSwitch.ahk
#Include %A_ScriptDir%\..\Src\Extensions\AutoDate.ahk
#Include %A_ScriptDir%\..\Src\Extensions\CapsLockSwitch.ahk
#Include %A_ScriptDir%\..\Src\Extensions\CapsLockHelper.ahk
#Include %A_ScriptDir%\..\Src\Extensions\TendonProfile.ahk
#Include %A_ScriptDir%\..\Src\Extensions\PTToolsWindow.ahk
#Include %A_ScriptDir%\..\Src\Extensions\UpdateChecker.ahk
#Include %A_ScriptDir%\..\Src\Extensions\CurrencyRates.ahk
#Include %A_ScriptDir%\..\Src\Extensions\SelectionActions.ahk

#Include %A_ScriptDir%\TestUtil.ahk
#Include %A_ScriptDir%\Suites\CoreTests.ahk
#Include %A_ScriptDir%\Suites\ProviderTests.ahk
#Include %A_ScriptDir%\Suites\SearchWindowTests.ahk
#Include %A_ScriptDir%\Suites\PreferencesTests.ahk
#Include %A_ScriptDir%\Suites\ExtensionTests.ahk
#Include %A_ScriptDir%\Suites\ProjectTests.ahk

OnError((err, mode) => TestRunner.OnUncaught(err, mode))                                             ; on a runtime error, print it and exit instead of hanging on a dialog
Logger.Enabled := false
I18n.LangDir := A_ScriptDir "\..\Resources\Lang"
I18n.Init("en")
AppSettings.Data := AppSettings.Defaults()                                  ; default settings in memory; no file is read or written
AppSettings.Feature("Clipboard")["Persist"] := 0                           ; the clipboard history tests don't write to disk
IconCache.DiskDir := ""                                                     ; icons don't use the disk cache (switched on for the IconDiskCache test only)
Pinyin.DataFile := A_ScriptDir "\..\Resources\Pinyin.txt"                   ; the scripts run from Tests\

Bench()
ExitApp(0)

Bench() {
    out(s) => FileAppend(s "`n", "*")
    root := A_Temp "\ALTRunIndexBench"
    words := ["budget", "report", "drawing", "plan", "notes", "invoice", "tender", "design", "section", "level", "schedule", "minutes", "photo", "contract", "summary", "meeting", "letter", "calc", "model", "review"]
    cn := ["测试", "报告", "图纸", "会议", "合同", "预算", "设计", "计算"]
    exts := [".pdf", ".docx", ".xlsx", ".dwg", ".jpg", ".txt", ".pptx", ".zip"]
    t := Logger.Ms(), made := 0
    DirCreate(root)
    Loop 20 {
        a := A_Index
        Loop 20 {
            b := A_Index
            dir := root "\Project " Format("{:02}", a) " " words[a] "\Sub " Format("{:02}", b) " " words[Mod(a + b, 20) + 1]
            DirCreate(dir)
            Loop 160 {
                i := A_Index, k := a * 10000 + b * 200 + i
                name := Mod(k, 10) = 0 ? cn[Mod(k, 8) + 1] cn[Mod(k * 3, 8) + 1] " " k : words[Mod(k, 20) + 1] "_" words[Mod(k * 7, 20) + 1] " " k
                h := DllCall("CreateFileW", "Str", dir "\" name exts[Mod(k, 8) + 1], "UInt", 0x40000000, "UInt", 0, "Ptr", 0, "UInt", 1, "UInt", 0x80, "Ptr", 0, "Ptr")
                if (h != -1)
                    DllCall("CloseHandle", "Ptr", h), made++
            }
        }
    }
    out(Format("created {} files + 420 folders in {:.0f} ms", made, Logger.Ms() - t))
    options := AppSettings.Feature("FileSearch")
    options["ScopeFolders"] := [root], options["ScopeDepth"] := 6, options["MaxEntries"] := 200000, options["UseEverything"] := 0

    ; A. the real Rebuild (batched in the background, as in ALTRun)
    state := {Done: 0}
    t := Logger.Ms()
    FileIndex.Rebuild((count) => (state.Done := count))
    while !state.Done
        Sleep(5)
    out(Format("A  Rebuild (real, batched 15 ms per tick): {:.0f} ms, {} entries", Logger.Ms() - t, state.Done))

    ; B. Loop Files, same work as _ScanSome, in one go (no time slices)
    exclude := options["ScopeExclude"]
    t := Logger.Ms(), stack := [[root, 0]], data := FileIndex.Entries(), initials := Map()
    while stack.Length {
        entry := stack.Pop()
        Loop Files, entry[1] "\*", "FD" {
            if (A_LoopFileAttrib ~= "[HS]")
                continue
            if (exclude != "" && RegExMatch(A_LoopFileFullPath, exclude))
                continue
            isFolder := InStr(A_LoopFileAttrib, "D") ? true : false
            data.Add(A_LoopFileDir, A_LoopFileName, isFolder)
            if ((text := FileIndex.InitialsOf(A_LoopFileName)) != "")
                initials[data.Count] := text
            if (isFolder && entry[2] + 1 < 6)
                stack.Push([A_LoopFileFullPath, entry[2] + 1])
        }
    }
    data.Seal()
    out(Format("B  Loop Files, same work, one go: {:.0f} ms ({} entries)", Logger.Ms() - t, data.Count))

    ; B2. Loop Files, enumeration only
    t := Logger.Ms(), stack := [root], n := 0
    while stack.Length {
        folder := stack.Pop()
        Loop Files, folder "\*", "FD" {
            n++
            if InStr(A_LoopFileAttrib, "D")
                stack.Push(A_LoopFileFullPath)
        }
    }
    out(Format("B2 Loop Files, enumeration only: {:.0f} ms ({} entries)", Logger.Ms() - t, n))

    ; C. FindFirstFileExW (FindExInfoBasic, FIND_FIRST_EX_LARGE_FETCH), enumeration only
    t := Logger.Ms(), stack := [root], n := 0, fd := Buffer(592)
    while stack.Length {
        folder := stack.Pop()
        h := DllCall("FindFirstFileExW", "Str", folder "\*", "Int", 1, "Ptr", fd, "Int", 0, "Ptr", 0, "UInt", 2, "Ptr")
        if (h = -1)
            continue
        loop {
            name := StrGet(fd.Ptr + 44)
            if (name != "." && name != "..") {
                n++
                if (NumGet(fd, 0, "UInt") & 0x10)
                    stack.Push(folder "\" name)
            }
        } until !DllCall("FindNextFileW", "Ptr", h, "Ptr", fd)
        DllCall("FindClose", "Ptr", h)
    }
    out(Format("C  FindFirstFileExW basic+large fetch, enumeration only: {:.0f} ms ({} entries)", Logger.Ms() - t, n))

    ; C2. FindFirstFileExW with the same work as _ScanSome but attributes by bit test, exclude only on folders, initials only for non-ASCII names
    t := Logger.Ms(), stack := [[root, 0]], data2 := FileIndex.Entries(), initials2 := Map()
    while stack.Length {
        entry := stack.Pop(), folder := entry[1]
        h := DllCall("FindFirstFileExW", "Str", folder "\*", "Int", 1, "Ptr", fd, "Int", 0, "Ptr", 0, "UInt", 2, "Ptr")
        if (h = -1)
            continue
        loop {
            attr := NumGet(fd, 0, "UInt"), name := StrGet(fd.Ptr + 44)
            if (name = "." || name = ".." || (attr & 0x6))
                continue
            isFolder := attr & 0x10
            if (isFolder && exclude != "" && RegExMatch(folder "\" name, exclude))
                continue
            data2.Add(folder, name, isFolder)
            if ((text := FileIndex.InitialsOf(name)) != "")
                initials2[data2.Count] := text
            if (isFolder && entry[2] + 1 < 6)
                stack.Push([folder "\" name, entry[2] + 1])
        } until !DllCall("FindNextFileW", "Ptr", h, "Ptr", fd)
        DllCall("FindClose", "Ptr", h)
    }
    data2.Seal()
    out(Format("C2 FindFirstFileExW, same work (bit tests, exclude on folders): {:.0f} ms ({} entries)", Logger.Ms() - t, data2.Count))

    ; D. cache: the text file as now
    t := Logger.Ms()
    FileIndex._SaveCache()
    out(Format("D  text cache write: {:.0f} ms, {} KB", Logger.Ms() - t, FileGetSize(FileIndex.File) // 1024))
    t := Logger.Ms()
    FileIndex._LoadCache()
    out(Format("D  text cache read (parse + Add + Seal): {:.0f} ms", Logger.Ms() - t))
    SetTimer(FileIndex._initialsTimer, 0)
    t := Logger.Ms(), cnt := 0
    Loop FileIndex.Data.Count
        if (FileIndex.InitialsOf(FileIndex.Data.NameAt(A_Index)) != "")
            cnt++
    out(Format("F  pinyin initials of every name (now done in the background after reading): {:.0f} ms ({} Chinese names)", Logger.Ms() - t, cnt))

    ; D2. binary cache prototype: header, Info as raw Int32, Shown as UTF-16, folders joined by `n; Lower = StrLower(Shown) on load
    d := FileIndex.Data, binFile := A_Temp "\ALTRunIndexBench.bin"
    t := Logger.Ms()
    dirsText := ""
    for dir in d.Dirs
        dirsText .= dir "`n"
    f := FileOpen(binFile, "w")
    f.RawWrite("AIX1", 4), f.WriteUInt(d.Count), f.WriteUInt(StrLen(d.Shown)), f.WriteUInt(StrLen(dirsText))
    f.RawWrite(d.Info, d.Info.Size), f.RawWrite(StrPtr(d.Shown), StrLen(d.Shown) * 2), f.RawWrite(StrPtr(dirsText), StrLen(dirsText) * 2)
    f.Close()
    out(Format("D2 binary cache write: {:.0f} ms, {} KB", Logger.Ms() - t, FileGetSize(binFile) // 1024))
    t := Logger.Ms()
    f := FileOpen(binFile, "r"), f.RawRead(magic := Buffer(4), 4), count := f.ReadUInt(), shownLen := f.ReadUInt(), dirsLen := f.ReadUInt()
    info := Buffer((count + 1) * 8), f.RawRead(info, info.Size)
    sb := Buffer(shownLen * 2), f.RawRead(sb, sb.Size), shown := StrGet(sb, shownLen, "UTF-16")
    db := Buffer(dirsLen * 2), f.RawRead(db, db.Size), dirs := StrSplit(RTrim(StrGet(db, dirsLen, "UTF-16"), "`n"), "`n")
    f.Close()
    lower := StrLower(shown)
    out(Format("D2 binary cache read: {:.0f} ms (count {}, {} folders, same text {})", Logger.Ms() - t, count, dirs.Length, lower == d.Lower))

    ; E. search: Find (InStr + scores) and TopIndexes (Format + Sort) separately, and a simple top-k
    for needle in ["re", "report", "budget_r", "udge", "1234", "xyz", "plan 12", "cs"] {
        scores := Map(), matches := []
        t := Logger.Ms()
        Loop 5
            scores := Map(), matches := [], d.Find(needle, scores, matches)
        tFind := (Logger.Ms() - t) / 5
        t := Logger.Ms()
        Loop 5
            top := FuzzyMatcher.TopIndexes(scores, 30)
        tTop := (Logger.Ms() - t) / 5
        t := Logger.Ms()
        Loop 5
            top2 := TopK(scores, 30)
        tTop2 := (Logger.Ms() - t) / 5
        same := true
        for i, v in top
            if (top2.Length < i || top2[i] != v)
                same := false
        t := Logger.Ms()
        Loop 5
            r := FileIndex.Search(needle, 30)
        tAll := (Logger.Ms() - t) / 5
        out(Format("E  [{}] {} matches: Find {:.1f} ms, TopIndexes {:.1f} ms, top-k {:.1f} ms (same order {}), Search total {:.1f} ms", needle, matches.Length, tFind, tTop, tTop2, same, tAll))
    }
    try DirDelete(root, true)
}

; the best limit indexes by score, ties by lower index (as TopIndexes): one pass keeping a short sorted list
TopK(scores, limit) {
    best := [], cutoff := -1
    for index, score in scores {
        if (best.Length >= limit && score <= cutoff)
            continue
        i := best.Length
        while (i >= 1 && best[i][1] < score)
            i--
        best.InsertAt(i + 1, [score, index])
        if (best.Length > limit)
            best.Pop()
        if (best.Length >= limit)
            cutoff := best[best.Length][1]
    }
    result := []
    for pair in best
        result.Push(pair[2])
    return result
}

class TestRunner {
    static Passed := 0, Failed := 0

    ; 8.3 short path (C:\Users\RUNNER~1\...) -> long path, as Loop Files lists them
    static LongPath(shortPath) {
        buf := Buffer(32767 * 2)
        return DllCall("GetLongPathNameW", "WStr", shortPath, "Ptr", buf, "UInt", 32767, "UInt") ? StrGet(buf) : shortPath
    }

    static Suites := [CoreTests, ProviderTests, SearchWindowTests, PreferencesTests, ExtensionTests, ProjectTests]

    ; Runs every test of every suite: a test is a static method whose name doesn't start with "_".
    ; Within a suite the tests run in name order (OwnProps lists them that way).
    static Run() {
        for suite in TestRunner.Suites {
            for name in suite.OwnProps() {
                if (SubStr(name, 1, 1) = "_" || name = "Prototype" || !suite.GetOwnPropDesc(name).HasOwnProp("Call"))
                    continue
                TestRunner.RunTest(suite, name)
            }
        }
        FileAppend("`n" TestRunner.Passed " passed, " TestRunner.Failed " failed`n", "*")
        ExitApp(TestRunner.Failed)
    }

    static RunTest(suite, name) {
        state := TestRunner._State()
        try {
            suite.%name%()
        } catch as e {
            SplitPath(e.File, &fileName)                                ; file name and call stack, so an occasional failure still shows where
            stack := RegExReplace(Trim(e.Stack, "`r`n"), "\R", " <- ")
            TestRunner.Fail(name, "exception: " e.Message " (" fileName ":" e.Line ") " SubStr(stack, 1, 600))
        }
        if ((after := TestRunner._State()) != state)                    ; each test puts back what it changed, so the order doesn't matter
            TestRunner.Fail(name, "left shared state changed (settings, language or logging): " TestRunner._Changed(state, after))
    }

    ; What a test must leave as it found it: the settings, the UI language, the log switches
    static _State() => JSON.Stringify(AppSettings.Data) "`n" I18n.Lang "`n" Logger.Enabled "|" Logger.File "|" Logger.Immediate

    ; The first difference, short enough to read in the CI annotation
    static _Changed(before, after) {
        loop Min(StrLen(before), StrLen(after))
            if (SubStr(before, A_Index, 1) != SubStr(after, A_Index, 1))
                return "... " SubStr(after, Max(1, A_Index - 40), 80) " ..."
        return "(length " StrLen(before) " -> " StrLen(after) ")"
    }

    static OnUncaught(err, mode) {
        FileAppend("UNCAUGHT " err.Message " (" err.File ":" err.Line ")`n", "*")
        ExitApp(99)
    }

    static Equal(name, actual, expected) {
        if (actual == expected)
            TestRunner.Passed++
        else
            TestRunner.Fail(name, "expected [" expected "] got [" actual "]")
    }

    static True(name, condition) {
        if condition
            TestRunner.Passed++
        else
            TestRunner.Fail(name, "condition is false")
        return condition ? true : false
    }

    static Fail(name, message) {
        TestRunner.Failed++
        FileAppend("FAIL " name ": " message "`n", "*")
    }
}
