;===============================================================================
; UpdateChecker.ahk - 检查 GitHub 上的新版本, 一键更新 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 版本号是日期格式 (2026.09.23), 和 GitHub Release 的 tag 比较。
;
; 用法:
;   UpdateChecker.Schedule()           启动时: 后台自动检查, 1 分钟后第一次, 之后每 6 小时一次 (见下)
;   UpdateChecker.Check()              手动检查 (托盘菜单 / 系统命令 / 偏好设置), 总会给出结果
;   UpdateChecker.Check(true)          有新版本就直接更新, 不询问 (命令行 -Update)
;   UpdateChecker.PendingItem()        后台发现的新版本 -> 搜索窗口里的一条结果 (没有时 "")
;   UpdateChecker.CleanUp()            启动时删掉上次更新留下的旧程序和临时文件
;
; 自动检查 (General.CheckForUpdates) 和 Alfred 一样不弹窗: ALTRun 常常开机启动后一直运行,
; 所以不只在启动时检查: 启动 1 分钟后, 离上次检查满 1 小时就检查 (每天开机就能发现前一天晚上发布的
; 版本; 改设置后重新载入不会反复检查), 之后每小时看一下离上次检查是否满 6 小时 (时间记在
; Data\Update.json; 检查失败时下一个小时再试)。发现新版本后, 空搜索框和搜索 "更新" 时显示一条
; "更新 ALTRun 到 x": Enter 更新, → 查看更新内容 / 跳过这个版本 (跳过的版本记在 Update.json)。
;
; 一键更新 (发现新版本时选 "立即更新"):
;   1. 下载 Release 里的 ALTRun_v<版本>.zip, 核对 GitHub 给出的 SHA256
;   2. 解压到临时文件夹, 先复制 exe 以外的文件 (Resources\ 等), 再把正在运行的
;      ALTRun.exe 改名为 ALTRun.exe.old (运行中的 exe 不能覆盖, 但可以改名), 复制新的 exe;
;      被覆盖的文件先备份, 任何一步失败都全部还原。Data\ (含 ALTRun.json)、Themes\ 不在 zip 里, 不会动
;   3. 启动新版本 (-Updated), 新版本启动时删掉 ALTRun.exe.old
; 用 Scoop / winget 安装的, 提示对应的升级命令 (搜索窗口里 Enter 复制命令); 运行源码时提示用
; git pull 或下载新版本; 程序目录不能写入、Release 没有 SHA256 时, 打开下载页面。
;===============================================================================

class UpdateChecker {
    static ApiUrl      := "https://api.github.com/repos/zhugecaomao/ALTRun/releases/latest"
    static ReleasePage := "https://github.com/zhugecaomao/ALTRun/releases"
    static TempDir     := A_Temp "\ALTRun-update"
    static OldSuffix   := ".old"
    static PackageExe  := "ALTRun.exe"                                     ; zip 里的程序名 (本机的 exe 可能改过名)
    ; 以前的版本带有、新版本已经不要的文件, 更新时删掉 (Resources\ 里其它文件可能是用户自己放的, 不动)
    static ObsoleteFiles := ["Resources\DOSBox.exe", "Resources\SDL.dll", "Resources\SDL_net.dll",
                             "Resources\SPF2M.exe", "Resources\Run.bat"]
    static UpgradeCommands := Map("scoop", "scoop update altrun", "winget", "winget upgrade zhugecaomao.ALTRun")
    static StateFile   := A_ScriptDir "\Data\Update.json"                    ; {LastCheck, Skip}
    static CheckHours  := 6                                                 ; 一直开着时: 离上次检查满几小时再查
    static StartupHours := 1                                                ; 启动时: 离上次检查满几小时就查
    static Pending     := ""                                                ; 后台发现的新版本 (ParseRelease 的结果 + Mode)
    static _timer      := ""

    static Check(install := false) {
        try {
            release := UpdateChecker.Fetch()
        } catch as e {
            Logger.Error("UpdateChecker: " e.Message)
            MsgBox(I18n.T("Update.Failed", e.Message), App.Name, 48)
            return
        }
        state := UpdateChecker.LoadState()
        state["LastCheck"] := A_Now
        UpdateChecker.SaveState(state)
        if (UpdateChecker.Compare(release.Version, App.Version) <= 0) {
            UpdateChecker.Pending := ""
            MsgBox(I18n.T("Update.Latest", App.Version), App.Name, 64)
            return
        }
        UpdateChecker.SetPending(release)                                   ; 选 "以后再说" 时搜索窗口里仍然有这一条
        switch release.Mode {
            case "install":
                choice := install ? "Install" : UpdateChecker._Ask(release)
                if (choice = "Install")
                    UpdateChecker.Install(release)
                else if (choice = "Notes")
                    Run(release.Page)
            case "source":
                if (MsgBox(I18n.T("Update.AvailableSource", release.Version), App.Name, "YesNo Iconi") = "Yes")
                    Run(release.Page)
            case "page":
                if (MsgBox(I18n.T("Update.Available", release.Version), App.Name, "YesNo Iconi") = "Yes")
                    Run(release.Page)
            default:
                if (MsgBox(I18n.T("Update.AvailableVia", release.Version, UpdateChecker.UpgradeCommands[release.Mode]), App.Name, "YesNo Iconi") = "Yes")
                    Run(release.Page)
        }
    }

    ;---------------------------------------------------------------------------
    ; 后台自动检查 + 搜索窗口里的更新提示
    ;---------------------------------------------------------------------------
    ; 启动时调用: 开机时网络可能还没连上, 1 分钟后再第一次检查
    static Schedule() {
        UpdateChecker._timer := () => UpdateChecker._Tick()
        SetTimer(() => UpdateChecker._Tick(UpdateChecker.StartupHours), -60000)
    }

    static _Tick(hours := "") {
        if UpdateChecker.IsDue(UpdateChecker.LoadState()["LastCheck"], A_Now, hours)
            UpdateChecker.CheckInBackground()
        SetTimer(UpdateChecker._timer, -3600000)
    }

    ; 离上次检查 (yyyyMMddHHmmss) 满 hours 小时 (默认 CheckHours) 了吗; 没检查过、时间不对 (改过系统时间) 时也检查
    static IsDue(lastCheck, now, hours := "") {
        if (hours = "")
            hours := UpdateChecker.CheckHours
        if (lastCheck = "")
            return true
        try {
            elapsed := DateDiff(now, lastCheck, "Hours")
            return elapsed >= hours || elapsed < 0
        }
        return true
    }

    ; 静默检查: 出错只写日志 (不记检查时间, 下一个小时再试); 有新版本 (且不是跳过的版本) 时记在 Pending
    static CheckInBackground() {
        try {
            release := UpdateChecker.Fetch()
        } catch as e {
            Logger.Error("UpdateChecker: " e.Message)
            return
        }
        state := UpdateChecker.LoadState()
        state["LastCheck"] := A_Now
        UpdateChecker.SaveState(state)
        if UpdateChecker.IsWanted(release.Version, App.Version, state["Skip"])
            UpdateChecker.SetPending(release)
        else
            UpdateChecker.Pending := ""
        Logger.Debug("UpdateChecker: latest " release.Version (IsObject(UpdateChecker.Pending) ? " (update available)" : ""))
    }

    static IsWanted(version, current, skip) {
        return UpdateChecker.Compare(version, current) > 0 && version != skip
    }

    ; 记下新版本, 并决定怎么更新 (Mode): "scoop" / "winget" / "source" (运行源码) / "install" (一键更新) / "page" (打开下载页面)
    static SetPending(release) {
        manager := UpdateChecker.InstalledBy(A_ScriptDir)
        release.Mode := (manager != "") ? manager : !A_IsCompiled ? "source" : UpdateChecker.CanInstall(release) ? "install" : "page"
        UpdateChecker.Pending := release
    }

    ; 搜索窗口里的一条 "更新 ALTRun 到 x" (空搜索框, 或搜索 "更新" 时由 SystemProvider 显示); 没有新版本时 ""
    static PendingItem() {
        release := UpdateChecker.Pending
        if !IsObject(release)
            return ""
        switch release.Mode {
            case "install": subtitle := I18n.T("Update.ItemInstall", App.Version)
            case "source" : subtitle := I18n.T("Update.ItemSource")
            case "page"   : subtitle := I18n.T("Update.ItemPage", App.Version)
            default       : subtitle := I18n.T("Update.ItemVia", UpdateChecker.UpgradeCommands[release.Mode])
        }
        actions := [ResultItem(I18n.T("Update.ReleaseNotes"), "", {Icon: "url:", OnRun: (*) => Run(release.Page)})
                  , ResultItem(I18n.T("Update.Skip"), I18n.T("Update.SkipHint"), {Icon: "res:imageres.dll,-5338", OnRun: (*) => UpdateChecker.SkipVersion(release.Version)})]
        icon := FileExist(App.IconFile) ? App.IconFile : A_ScriptFullPath
        return ResultItem(I18n.T("Update.ItemTitle", release.Version), subtitle
            , {Icon: icon, OnRun: (*) => UpdateChecker.RunPending(), Actions: actions})
    }

    ; 搜索窗口里按 Enter: 一键更新 / 打开下载页面 / 复制包管理器的升级命令
    static RunPending() {
        release := UpdateChecker.Pending
        if !IsObject(release)
            return
        switch release.Mode {
            case "install"       : UpdateChecker.Install(release)
            case "source", "page": Run(release.Page)
            default:
                command := UpdateChecker.UpgradeCommands[release.Mode]
                A_Clipboard := command
                App.Notify(I18n.T("Update.CommandCopied", command), 4000)
        }
    }

    static SkipVersion(version) {
        state := UpdateChecker.LoadState()
        state["Skip"] := version
        UpdateChecker.SaveState(state)
        UpdateChecker.Pending := ""
        App.Notify(I18n.T("Update.Skipped", version), 3000)
    }

    static LoadState() {
        state := Map("LastCheck", "", "Skip", "")
        try {
            data := JSON.Parse(FileRead(UpdateChecker.StateFile, "UTF-8"))
            for key in ["LastCheck", "Skip"]
                if (data is Map && data.Has(key))
                    state[key] := String(data[key])
        }
        return state
    }

    static SaveState(state) {
        try {
            SplitPath(UpdateChecker.StateFile, , &dir)
            DirCreate(dir)
            try FileDelete(UpdateChecker.StateFile)
            FileAppend(JSON.Stringify(state), UpdateChecker.StateFile, "UTF-8")
        } catch as e {
            Logger.Error("UpdateChecker.SaveState: " e.Message)
        }
    }

    ;---------------------------------------------------------------------------
    ; GitHub Release
    ;---------------------------------------------------------------------------
    static Fetch() {
        tmpFile := A_Temp "\ALTRun_latest.json"
        Download(UpdateChecker.ApiUrl, tmpFile)
        response := FileRead(tmpFile, "UTF-8")
        try FileDelete(tmpFile)
        return UpdateChecker.ParseRelease(response)
    }

    ; GitHub API 的 releases/latest -> {Version, Page, ZipUrl, Sha256} (没有 zip / SHA256 时为 "")
    static ParseRelease(text) {
        data := JSON.Parse(text)
        if !(data is Map) || !data.Has("tag_name") || data["tag_name"] = ""
            throw Error("Cannot find 'tag_name' in the GitHub response.")
        release := {Version: Trim(data["tag_name"], "vV "), Page: UpdateChecker.ReleasePage, ZipUrl: "", Sha256: ""}
        if data.Has("html_url") && data["html_url"] != ""
            release.Page := data["html_url"]
        if data.Has("assets") && data["assets"] is Array {
            for asset in data["assets"] {
                if !(asset is Map) || !asset.Has("name") || !RegExMatch(asset["name"], "i)^ALTRun_v?[\d.]+\.zip$")
                    continue
                release.ZipUrl := asset.Has("browser_download_url") ? asset["browser_download_url"] : ""
                if asset.Has("digest") && RegExMatch(asset["digest"], "i)^sha256:([0-9a-f]{64})$", &m)
                    release.Sha256 := StrLower(m[1])
                break
            }
        }
        return release
    }

    ; 能不能在程序里直接更新: 编译后的 exe, Release 有 zip 和 SHA256, 程序目录能写入
    static CanInstall(release) {
        return A_IsCompiled && release.ZipUrl != "" && release.Sha256 != "" && UpdateChecker.IsWritable(A_ScriptDir)
    }

    static IsWritable(dir) {
        probe := dir "\ALTRun.write-test.tmp"
        try {
            FileAppend("", probe)
            FileDelete(probe)
            return true
        }
        return false
    }

    ; 用包管理器安装的, 提示用它升级 (程序自己替换文件会让包管理器的记录对不上):
    ;   Scoop   ...\scoop\apps\altrun\current (或版本号文件夹)
    ;   winget  %LOCALAPPDATA%\Microsoft\WinGet\Packages\zhugecaomao.ALTRun_...
    static InstalledBy(dir) {
        if RegExMatch(dir, "i)\\apps\\altrun\\[^\\]+$")
            return "scoop"
        if RegExMatch(dir, "i)\\WinGet\\Packages\\zhugecaomao\.ALTRun")
            return "winget"
        return ""
    }

    ; 下载 -> 核对 SHA256 -> 解压 -> 替换 -> 启动新版本; 出错时提示打开下载页面
    static Install(release) {
        progress := UpdateChecker._Progress(I18n.T("Update.Downloading", release.Version))
        try {
            dir := UpdateChecker.TempDir
            try DirDelete(dir, true)
            DirCreate(dir)
            zip := dir "\ALTRun_v" release.Version ".zip"
            Download(release.ZipUrl, zip)
            actual := UpdateChecker.Sha256File(zip)
            if (actual != release.Sha256)
                throw Error("SHA256 does not match (expected " release.Sha256 ", got " actual ").")
            progress["Text"].Value := I18n.T("Update.Installing", release.Version)
            files := dir "\files"
            UpdateChecker.Extract(zip, files)
            UpdateChecker.Apply(files, A_ScriptDir, A_ScriptName)
        } catch as e {
            progress.Destroy()
            Logger.Error("UpdateChecker: install " release.Version " - " e.Message)
            if (MsgBox(I18n.T("Update.InstallFailed", e.Message), App.Name, "YesNo Icon!") = "Yes")
                Run(release.Page)
            return
        }
        progress.Destroy()
        Logger.Debug("UpdateChecker: updated to " release.Version)
        App.Restart("-Updated " release.Version)
    }

    ; zip -> 文件夹 (Windows PowerShell 自带 Expand-Archive)
    static Extract(zip, dest) {
        quote := (s) => "'" StrReplace(s, "'", "''") "'"
        cmd := 'powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "'
             . "$ErrorActionPreference = 'Stop'; Expand-Archive -LiteralPath " quote(zip) " -DestinationPath " quote(dest) ' -Force"'
        exitCode := RunWait(cmd, , "Hide")
        if (exitCode != 0 || !FileExist(dest "\" UpdateChecker.PackageExe))
            throw Error("Could not extract the downloaded package (exit code " exitCode ").")
    }

    ; 把 src 里的新版本复制到 dest: 先复制其它文件, 最后替换 dest\exeName (改名为 .old 后复制新的)。
    ; 被覆盖的文件先备份到 src.backup; 任何一步失败都还原 (覆盖的改回去, 新增的删掉), 再抛出错误
    static Apply(src, dest, exeName) {
        newExe := src "\" UpdateChecker.PackageExe
        if !FileExist(newExe)
            throw Error(UpdateChecker.PackageExe " is missing from the update package.")
        exe := dest "\" exeName, old := exe UpdateChecker.OldSuffix
        backup := src ".backup"
        try DirDelete(backup, true)
        changes := {Replaced: [], Created: [], Dirs: []}
        movedExe := false
        try {
            UpdateChecker._CopyTree(src, dest, backup, changes, UpdateChecker.PackageExe)
            if FileExist(old)
                FileDelete(old)                                             ; 上次更新留下的, 删不掉就不要继续
            if FileExist(exe) {
                FileMove(exe, old)
                movedExe := true
            }
            FileCopy(newExe, exe)
        } catch as e {
            if movedExe {
                try FileDelete(exe)
                try FileMove(old, exe)
            }
            UpdateChecker._Undo(changes)
            try DirDelete(backup, true)
            throw e
        }
        try DirDelete(backup, true)
        for rel in UpdateChecker.ObsoleteFiles
            if !FileExist(src "\" rel)                                    ; 新版本里还有的就留着
                try FileDelete(dest "\" rel)
    }

    ; 逐层复制 (按文件名拼路径: A_Temp 和 Loop Files 给出的路径可能一个是长文件名、一个是 8.3 短文件名,
    ; 不能按长度截取相对路径)。已有的文件覆盖前先复制到 backup, 记在 changes 里以便还原
    static _CopyTree(src, dest, backup, changes, skipFile := "") {
        if !InStr(FileExist(dest), "D") {
            DirCreate(dest)
            changes.Dirs.Push(dest)
        }
        Loop Files src "\*", "FD" {
            name := A_LoopFileName
            if InStr(A_LoopFileAttrib, "D") {
                UpdateChecker._CopyTree(src "\" name, dest "\" name, backup "\" name, changes)
                continue
            }
            if (name = skipFile)
                continue
            target := dest "\" name
            if FileExist(target) {
                DirCreate(backup)
                FileCopy(target, backup "\" name, true)
                changes.Replaced.Push([target, backup "\" name])
            } else {
                changes.Created.Push(target)
            }
            FileCopy(src "\" name, target, true)
        }
    }

    static _Undo(changes) {
        for pair in changes.Replaced
            try FileCopy(pair[2], pair[1], true)
        for created in changes.Created
            try FileDelete(created)
        Loop changes.Dirs.Length                                            ; 从里往外删新建的 (空) 文件夹
            try DirDelete(changes.Dirs[changes.Dirs.Length - A_Index + 1])
    }

    ; 新版本启动时: 删掉改了名的旧程序 (旧进程可能还没退出, 过一会儿再试) 和下载的临时文件
    static CleanUp(attempt := 1) {
        old := A_ScriptFullPath UpdateChecker.OldSuffix
        if FileExist(old) {
            try FileDelete(old)
            if FileExist(old) && (attempt < 10)
                SetTimer(() => UpdateChecker.CleanUp(attempt + 1), -1000)
        }
        if (attempt = 1 && InStr(FileExist(UpdateChecker.TempDir), "D"))
            try DirDelete(UpdateChecker.TempDir, true)
    }

    static Sha256File(path) {
        data := FileRead(path, "RAW")
        hAlg := 0
        if DllCall("bcrypt\BCryptOpenAlgorithmProvider", "Ptr*", &hAlg, "WStr", "SHA256", "Ptr", 0, "UInt", 0, "UInt")
            throw Error("SHA256 is not available.")
        digest := Buffer(32)
        status := DllCall("bcrypt\BCryptHash", "Ptr", hAlg, "Ptr", 0, "UInt", 0
                        , "Ptr", data, "UInt", data.Size, "Ptr", digest, "UInt", 32, "UInt")
        DllCall("bcrypt\BCryptCloseAlgorithmProvider", "Ptr", hAlg, "UInt", 0)
        if status
            throw Error(Format("SHA256 failed (0x{:08X}).", status))
        hex := ""
        Loop 32
            hex .= Format("{:02x}", NumGet(digest, A_Index - 1, "UChar"))
        return hex
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

    ; 发现新版本的对话框 -> "Install" / "Notes" / "" (以后再说)
    static _Ask(release) {
        choice := ""
        dlg := Gui("+AlwaysOnTop -MinimizeBox -MaximizeBox", App.Name)
        dlg.SetFont("s10")
        dlg.MarginX := 20, dlg.MarginY := 16
        dlg.AddText("w420", I18n.T("Update.Prompt", release.Version, App.Version))
        close := (value) => (choice := value, dlg.Destroy())
        dlg.AddButton("xm y+18 w130 Default", I18n.T("Update.InstallNow")).OnEvent("Click", (*) => close("Install"))
        dlg.AddButton("x+10 w130", I18n.T("Update.ReleaseNotes")).OnEvent("Click", (*) => close("Notes"))
        dlg.AddButton("x+10 w130", I18n.T("Update.Later")).OnEvent("Click", (*) => close(""))
        dlg.OnEvent("Close", (*) => close(""))
        dlg.OnEvent("Escape", (*) => close(""))
        dlg.Show()
        WinWaitClose("ahk_id " dlg.Hwnd)
        return choice
    }

    static _Progress(text) {
        dlg := Gui("+AlwaysOnTop -SysMenu +ToolWindow", App.Name)
        dlg.SetFont("s10")
        dlg.MarginX := 24, dlg.MarginY := 18
        dlg.AddText("w320 vText", text)
        dlg.Show()
        return dlg
    }
}
