;===============================================================================
; Indexer.ahk - 开机自启/发送到/开始菜单快捷方式 + 文件索引重建 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 前三个方法在启动时直接调用一次(按配置决定创建还是删除对应的 .lnk 快捷方式),
; 不涉及动态调用, 所以不需要在 ALTRun.ahk 里留裸函数外壳, 直接
; Indexer.UpdateSendTo() 这样调用即可 - 和 CommandStore.LoadCommands()/
; Listary.Init() 等启动流程里的其它调用是同一个写法。
;
; Rebuild() 不一样: 它就是内置命令 "Func | Reindex | ..." 本身, 还登记在
; FuncList 和托盘/右键菜单里, 所以 ALTRun.ahk 里仍然留了一个裸的 Reindex(*)
; 外壳调用 Indexer.Rebuild() - 原因同其它已抽出的模块, RunCommand() 的
; %cmdPath%() 和 Menu.Add() 的回调都只认裸的全局函数名, 不认 Class.Method。
;===============================================================================

Class Indexer {

    static UpdateSendTo() {                                              ; the lnk in SendTo must point to a exe
        lnkPath := StrReplace(A_StartMenu, "\Start Menu", "\SendTo\") "ALTRun.lnk"
        if (!g_CONFIG["EnableSendTo"]) {
            try FileDelete(lnkPath)
            g_LOG.Debug("UpdateSendTo: Update SendTo shortcut...Disabled")
            return
        }

        if (A_IsCompiled)
            FileCreateShortcut(A_ScriptFullPath, lnkPath, ,"-SendTo", "Send command to ALTRun User Command list")
        else
            FileCreateShortcut(A_AhkPath, lnkPath, , A_ScriptFullPath " -SendTo", "Send command to ALTRun User Command list")

        g_LOG.Debug("UpdateSendTo: Update SendTo shortcut...OK")
    }

    static UpdateStartup() {
        lnkPath := A_Startup "\ALTRun.lnk"

        if (!g_CONFIG["AutoStartup"]) {
            try FileDelete(lnkPath)
            g_LOG.Debug("UpdateStartup: Update Startup shortcut...Disabled")
            return
        }

        FileCreateShortcut(A_ScriptFullPath, lnkPath, A_ScriptDir, "-startup", "ALTRun - An effective launcher")
        g_LOG.Debug("UpdateStartup: Update Startup shortcut...OK")
    }

    static UpdateStartMenu() {
        lnkPath := A_Programs "\ALTRun.lnk"

        if (!g_CONFIG["InStartMenu"]) {
            try FileDelete(lnkPath)
            g_LOG.Debug("UpdateStartMenu: Update StartMenu shortcut...Disabled")
            return
        }

        FileCreateShortcut(A_ScriptFullPath, lnkPath, A_ScriptDir, "-StartMenu", "ALTRun - An effective launcher")
        g_LOG.Debug("UpdateStartMenu: Update StartMenu shortcut...OK")
    }

    ; Re-create the "Index" command section from scratch by scanning g_CONFIG's
    ; IndexDir/IndexType/IndexDepth/IndexExclude, plus Windows Store apps.
    static Rebuild() {
        AppData.LoadAppData()
        indexMap := Map()

        ; Create ProgressGui at the start
        ProgressGui := Gui("-MinimizeBox +AlwaysOnTop", "Reindex")
        ProgressGui.Add("Text", , "ReIndexing...")
        ProgressGui.Add("Progress", "vMyProgress w200", 0)
        ProgressGui.Add("Text", "vMyFileName w200", "Starting...")
        ProgressGui.Show()

        ; Move repeated config queries outside loop
        maxDepth := g_CONFIG["IndexDepth"]
        shouldCheckExclude := g_CONFIG["IndexExclude"] != ""
        excludePattern := g_CONFIG["IndexExclude"]

        for dirIndex, dir in StrSplit(g_CONFIG["IndexDir"], ",") {
            searchPath := RegExReplace(Path.Resolve(Trim(dir)), "\\+$")     ; Remove trailing backslashes
            if !DirExist(searchPath)
                continue

            for extIndex, ext in StrSplit(g_CONFIG["IndexType"], ",") {
                ext := Trim(ext)
                if (ext = "")
                    continue
                Loop Files, searchPath "\" ext, "R" {                       ; Calculate path relative to searchPath and count subdir levels
                    rel := SubStr(A_LoopFileFullPath, StrLen(searchPath) + 2) ; +2 to skip the backslash
                    seps := (rel = "") ? 0 : StrLen(rel) - StrLen(StrReplace(rel, "\", "")) ; Count backslashes to determine depth

                    if (seps > maxDepth)                                    ; If file is deeper than allowed depth, skip it.
                        continue

                    if (shouldCheckExclude && RegExMatch(A_LoopFileFullPath, excludePattern))
                        continue                                            ; Skip this file and move on to the next loop.

                    indexMap["File | " . A_LoopFileFullPath] := 1   ; Collect file entry

                    ; Update ProgressGui (throttled to reduce UI overhead)
                    if (!Mod(A_Index, 20))
                        ProgressGui["MyProgress"].Value := Mod(A_Index, 100), ProgressGui["MyFileName"].Text := A_LoopFileName
                }
            }
        }

        ; Index Windows Store Apps
        if (g_CONFIG["IndexStoreApp"]) {
            try {
                ProgressGui["MyFileName"].Text  := "Indexing Store Apps..."
                tempFile := A_Temp . "\ALTRun_StoreApps.csv"
                RunWait('powershell -Command "Get-StartApps | Select-Object Name, AppID | ConvertTo-Csv -NoTypeInformation" > "' . tempFile . '"', , "Hide")
                if FileExist(tempFile) {
                    output := FileRead(tempFile)
                    FileDelete(tempFile)
                    lines := StrSplit(output, "`n", "`r")
                    for line in lines {
                        if (A_Index == 1 or Trim(line) == "")  ; Skip header
                            continue
                        fields := StrSplit(line, '","')
                        if (fields.Length >= 2) {
                            name := StrReplace(fields[1], '"', '')
                            appid := StrReplace(fields[2], '"', '')
                            indexMap["App | shell:AppsFolder\" . appid . " | " . name] := 1 ; Collect app entry
                        }
                        ProgressGui["MyProgress"].Value := A_Index
                        ProgressGui["MyFileName"].Text  := name ? name : "Unknown App"
                        Sleep 10  ; Small delay to show progress
                    }
                    g_LOG.Debug("Reindex: Indexed Windows Store apps successfully")
                } else {
                    g_LOG.Debug("Reindex: Temp file not found for Store apps")
                }
            } catch as e {
                g_LOG.Debug("Reindex: Error indexing Store apps: " . e.Message)
            }
        }

        ; Destroy ProgressGui at the end
        ProgressGui.Destroy()

        ; Keep the rank a command already earned, so reindexing does not reset SmartRank
        for cmdLine, _ in indexMap {
            if (g_CMDDATA["Index"].Has(cmdLine) && IsInteger(g_CMDDATA["Index"][cmdLine]))
                indexMap[cmdLine] := g_CMDDATA["Index"][cmdLine]
        }
        g_CMDDATA["Index"] := indexMap
        AppData.SaveAppData()

        g_LOG.Debug("Reindex: Indexing search database...OK")
        TrayTip("ReIndex database finish successfully.", g_TITLE, 8)
        CommandStore.LoadCommands()
    }
}
