;===============================================================================
; FileSearchProvider.ahk - 文件搜索 (Everything) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 触发方式 (和 Alfred 一样):  'report   或   open report   或   find report
;
; 搜索引擎按以下顺序自动选择:
;   1. Everything SDK (Everything64.dll / Everything32.dll), 需要 Everything 在运行
;   2. Everything 命令行 es.exe
;   3. 都没有时: 一条 "在 Everything 中搜索" (有 Everything.exe) 或
;      "用 Windows 搜索" 的结果
; DLL / es.exe / Everything.exe 放在 ALTRun 目录、Lib 目录、PATH 或 Everything 的
; 安装目录都能找到; 也可以在 Features.FileSearch.EverythingPath 里指定目录或文件。
;===============================================================================

class FileSearchProvider {
    static Id := "FileSearch"
    static _dll := "", _esExe := "", _everythingExe := "", _detected := false

    static Init() {
        FileSearchProvider._Detect()
    }

    static Search(query) {
        options := AppSettings.Feature("FileSearch")
        term := ""
        matched := (options["QuotePrefix"] && query.MatchPrefix("'", &term))
        if !matched
            matched := query.MatchKeyword(options["Keywords"], &term)
        if !matched
            return []
        if (term = "")
            return [ResultItem(I18n.T("Files.Keyword"), "'... / " options["Keywords"][1] " ...", {Icon: "folder:", Valid: false, Score: 150, Exclusive: true})]

        results := []
        for filePath in FileSearchProvider.Query(term, options["MaxResults"]) {
            isFolder := InStr(FileExist(filePath), "D") ? true : false
            SplitPath(filePath, &name)
            results.Push(ResultItem(name, filePath, {
                Kind: isFolder ? "folder" : "file", Arg: filePath, Icon: filePath,
                Uid: "file:" StrLower(filePath), Score: 150 - A_Index * 0.01, Exclusive: true
            }))
        }
        if !results.Length {
            item := FileSearchProvider.FallbackItem(term)
            item.Score := 150
            item.Exclusive := true
            results.Push(item)
        }
        return results
    }

    ; 返回匹配的完整路径数组; 没有可用的 Everything 时返回空数组
    static Query(term, maxResults := 30) {
        FileSearchProvider._Detect()
        if (FileSearchProvider._dll != "")
            return FileSearchProvider._QueryDll(term, maxResults)
        if (FileSearchProvider._esExe != "")
            return FileSearchProvider._QueryEs(term, maxResults)
        return []
    }

    static FallbackItem(term) {
        FileSearchProvider._Detect()
        if (FileSearchProvider._everythingExe != "") {
            exe := FileSearchProvider._everythingExe
            return ResultItem(I18n.T("Files.OpenEverything", term), exe, {
                Icon: exe, OnRun: (*) => Run('"' exe '" -s "' term '"')
            })
        }
        return ResultItem(I18n.T("Files.WindowsSearch", term), "search-ms:", {
            Icon: "res:imageres.dll,-8", OnRun: (*) => Run("search-ms:query=" Url.Encode(term))
        })
    }

    ;---------------------------------------------------------------------------
    ; Everything SDK
    ;---------------------------------------------------------------------------
    static _QueryDll(term, maxResults) {
        dll := FileSearchProvider._dll
        results := []
        DllCall(dll "\Everything_SetSearchW", "WStr", term)
        DllCall(dll "\Everything_SetMax", "UInt", maxResults)
        DllCall(dll "\Everything_SetRequestFlags", "UInt", 0x4)             ; EVERYTHING_REQUEST_FULL_PATH_AND_FILE_NAME
        if !DllCall(dll "\Everything_QueryW", "Int", 1)
            return results                                                  ; Everything 没有运行
        pathBuffer := Buffer(2048 * 2)
        count := DllCall(dll "\Everything_GetNumResults", "UInt")
        Loop count {
            DllCall(dll "\Everything_GetResultFullPathNameW", "UInt", A_Index - 1, "Ptr", pathBuffer, "UInt", 2048)
            results.Push(StrGet(pathBuffer, "UTF-16"))
        }
        return results
    }

    ;---------------------------------------------------------------------------
    ; es.exe
    ;---------------------------------------------------------------------------
    static _QueryEs(term, maxResults) {
        outFile := A_Temp "\ALTRun_es.txt"
        try FileDelete(outFile)
        try {
            RunWait(A_ComSpec ' /c ""' FileSearchProvider._esExe '" -n ' maxResults ' "' StrReplace(term, '"') '" > "' outFile '""', , "Hide")
        } catch {
            return []
        }
        results := []
        if FileExist(outFile) {
            for line in StrSplit(FileRead(outFile), "`n", "`r")
                if (Trim(line) != "")
                    results.Push(line)
            try FileDelete(outFile)
        }
        return results
    }

    ;---------------------------------------------------------------------------
    ; Detection
    ;---------------------------------------------------------------------------
    static _Detect() {
        if FileSearchProvider._detected
            return
        FileSearchProvider._detected := true
        dllName := (A_PtrSize = 8) ? "Everything64.dll" : "Everything32.dll"
        folders := []
        configured := Path.Resolve(AppSettings.Feature("FileSearch")["EverythingPath"])
        if (configured != "")
            folders.Push(DirExist(configured) ? configured : FileSearchProvider._DirOf(configured))
        for folder in [A_ScriptDir, A_ScriptDir "\Lib", A_ProgramFiles "\Everything", EnvGet("ProgramFiles(x86)") "\Everything", EnvGet("LocalAppData") "\Everything"]
            folders.Push(folder)

        for folder in folders {
            if (FileSearchProvider._dll = "" && FileExist(folder "\" dllName) && DllCall("LoadLibrary", "Str", folder "\" dllName, "Ptr"))
                FileSearchProvider._dll := folder "\" dllName
            if (FileSearchProvider._esExe = "" && FileExist(folder "\es.exe"))
                FileSearchProvider._esExe := folder "\es.exe"
            if (FileSearchProvider._everythingExe = "" && FileExist(folder "\Everything.exe"))
                FileSearchProvider._everythingExe := folder "\Everything.exe"
        }
        if (configured != "" && FileExist(configured) && !DirExist(configured) && RegExMatch(configured, "i)everything\.exe$"))
            FileSearchProvider._everythingExe := configured
        if (FileSearchProvider._esExe = "") {
            onPath := Path.Resolve("es.exe")
            if (onPath != "es.exe" && FileExist(onPath))
                FileSearchProvider._esExe := onPath
        }
        Logger.Debug("FileSearchProvider: dll=" FileSearchProvider._dll " es=" FileSearchProvider._esExe " everything=" FileSearchProvider._everythingExe)
    }

    static _DirOf(file) {
        SplitPath(file, , &dir)
        return dir
    }
}
