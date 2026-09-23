;===============================================================================
; UpdateChecker.ahk - 检查 GitHub 上的新版本 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 用法 (ALTRun.ahk 里, 全部通过同名的裸全局函数外壳调用):
;   AutoCheckUpdate()   - 启动时按配置决定是否静默检查一次(延迟 1 秒)
;   Update()            - 托盘菜单/右键菜单"检查更新", 手动检查, 会弹提示
;   CheckUpdate(silent) - 实际检查逻辑, 上面两个的共同入口
;
; 类名不能叫 Update - 裸全局函数 Update(*) 要保留这个名字给 FuncList/
; Menu.Add() 用(不认 Class.Method, 也不能和类同名, 否则会跟 Class Update
; 的类名声明冲突), 所以类叫 UpdateChecker, 和 Kanji/Clip 那些外壳同理。
;===============================================================================

Class UpdateChecker {
    static RepoAPI     := "https://api.github.com/repos/zhugecaomao/ALTRun/releases/latest"
    static ReleasePage := "https://github.com/zhugecaomao/ALTRun/releases"

    ; 启动时调用一次: 按配置决定要不要延迟 1 秒静默检查一次新版本。
    static AutoCheck() {
        if (!g_CONFIG["AutoUpdateCheck"])
            return
        SetTimer(CheckUpdate, -1000)
    }

    static Check(silent := true) {
        try {
            tmpFile := A_Temp "\ALTRun_latest.json"
            Download(UpdateChecker.RepoAPI, tmpFile)
            json := FileRead(tmpFile, "UTF-8")

            if !RegExMatch(json, '"tag_name"\s*:\s*"([^"]+)"', &verMatch)
                throw Error("Cannot find 'tag_name' in GitHub API response.")

            latestVersion  := Trim(verMatch[1], "vV ")
            currentVersion := Trim(g_TITLE, "ALTRun - v ")

            if (UpdateChecker._CompareVersion(latestVersion, currentVersion) > 0) {
                MsgBox(g_LNG[805] latestVersion g_LNG[806], g_Title, 64)
                Run UpdateChecker.ReleasePage
            } else if (!silent) {
                ; Only show "up-to-date" message for manual checks
                MsgBox(g_LNG[807] currentVersion g_LNG[808], g_Title, 64)
            }
        } catch as e {
            if (!silent)
                MsgBox(g_LNG[809] e.Message, g_Title, 48)
            else
                g_LOG.Debug("CheckUpdate: Update check failed: " e.Message)
        }
    }

    static _CompareVersion(v1, v2) {
        v1Parts := StrSplit(v1, ".")
        v2Parts := StrSplit(v2, ".")
        Loop Max(v1Parts.Length, v2Parts.Length) {
            diff := (v1Parts[A_Index] + 0) - (v2Parts[A_Index] + 0)
            if diff
                return diff
        }
        return 0
    }
}
