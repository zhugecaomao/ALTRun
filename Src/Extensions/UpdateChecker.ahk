;===============================================================================
; UpdateChecker.ahk - 检查 GitHub 上的新版本 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 版本号是日期格式 (2026.09.23), 和 GitHub Release 的 tag 比较。
;
; 用法:
;   UpdateChecker.Check(true)     启动后静默检查 (只在有新版本时提示)
;   UpdateChecker.Check(false)    手动检查 (托盘菜单 / 系统命令), 总会给出结果
;===============================================================================

class UpdateChecker {
    static ApiUrl      := "https://api.github.com/repos/zhugecaomao/ALTRun/releases/latest"
    static ReleasePage := "https://github.com/zhugecaomao/ALTRun/releases"

    static Check(silent := true) {
        try {
            tmpFile := A_Temp "\ALTRun_latest.json"
            Download(UpdateChecker.ApiUrl, tmpFile)
            response := FileRead(tmpFile, "UTF-8")
            try FileDelete(tmpFile)
            if !RegExMatch(response, '"tag_name"\s*:\s*"([^"]+)"', &m)
                throw Error("Cannot find 'tag_name' in the GitHub response.")
            latest := Trim(m[1], "vV ")
            if (UpdateChecker.Compare(latest, App.Version) > 0) {
                if (MsgBox(I18n.T("Update.Available", latest), App.Name, "YesNo Iconi") = "Yes")
                    Run(UpdateChecker.ReleasePage)
            } else if !silent {
                MsgBox(I18n.T("Update.Latest", App.Version), App.Name, 64)
            }
        } catch as e {
            Logger.Error("UpdateChecker: " e.Message)
            if !silent
                MsgBox(I18n.T("Update.Failed", e.Message), App.Name, 48)
        }
    }

    ; 按点分隔逐段比较数字: > 0 表示 v1 更新
    static Compare(v1, v2) {
        parts1 := StrSplit(v1, "."), parts2 := StrSplit(v2, ".")
        Loop Max(parts1.Length, parts2.Length) {
            a := (A_Index <= parts1.Length && IsNumber(parts1[A_Index])) ? parts1[A_Index] + 0 : 0
            b := (A_Index <= parts2.Length && IsNumber(parts2[A_Index])) ? parts2[A_Index] + 0 : 0
            if (a != b)
                return a - b
        }
        return 0
    }
}
