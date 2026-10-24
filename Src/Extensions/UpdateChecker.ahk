;===============================================================================
; UpdateChecker.ahk - checks GitHub for a new version, one-click update (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Versions are dates (2026.09.23), compared with the GitHub Release tag.
;
; Usage:
;   UpdateChecker.Schedule()           at start-up: automatic background checks, the first after 1 minute, then every 6 hours (see below)
;   UpdateChecker.Check()              manual check (tray menu / system command / Preferences), always reports a result
;   UpdateChecker.Check(true)          updates right away when there is a new version, without asking (command line -Update)
;   UpdateChecker.PendingItem()        a new version found in the background -> a result in the search window ("" if none)
;   UpdateChecker.CleanUp()            at start-up: deletes the old program and temporary files left by the last update
;
; Automatic checks (General.CheckForUpdates) don't pop up, as in Alfred. ALTRun often starts with Windows and keeps running,
; so it doesn't only check at start-up: 1 minute after start-up it checks if the last check is at least 1 hour old (so a version released the previous evening
; is found at each day's boot, and reloading after settings changes doesn't check again and again), then every hour it looks whether the last check is 6 hours old (the time is kept in
; Data\Update.json; a failed check retries the next hour). After a new version is found, the empty search box and a search for "update" show
; "New version available: ALTRun x": Enter installs it, → release notes / skip this version (skipped versions are kept in Update.json).
;
; One-click update ("Update now" when a new version is found):
;   1. download ALTRun_v<version>.zip from the Release and verify the SHA256 GitHub gives
;   2. extract to a temporary folder, copy the files other than the exe first (Resources\ etc.), then rename the running
;      ALTRun.exe to ALTRun.exe.old (a running exe can't be overwritten but can be renamed) and copy the new exe;
;      overwritten files are backed up first, and any failing step restores everything. Data\ (with ALTRun.json) and Themes\ aren't in the zip and stay untouched
;   3. start the new version (-Updated), which deletes ALTRun.exe.old at start-up
; Installed with Scoop / winget: show the matching upgrade command (Enter in the search window copies it); running the source: suggest
; git pull or downloading the new version; when the program folder isn't writable or the Release has no SHA256, open the download page.
;===============================================================================

class UpdateChecker {
    static ApiUrl      := "https://api.github.com/repos/zhugecaomao/ALTRun/releases/latest"
    static ReleasePage := "https://github.com/zhugecaomao/ALTRun/releases"
    static RepoUrl     := "https://github.com/zhugecaomao/ALTRun/"
    static TempDir     := A_Temp "\ALTRun-update"
    static OldSuffix   := ".old"
    static PackageExe  := "ALTRun.exe"                                     ; the program name in the zip (the local exe may have been renamed)
    ; files older versions shipped that the new one no longer needs, deleted on update (other files in Resources\ may be the user's, left alone)
    static ObsoleteFiles := ["Resources\DOSBox.exe", "Resources\SDL.dll", "Resources\SDL_net.dll",
                             "Resources\SPF2M.exe", "Resources\Run.bat"]
    static UpgradeCommands := Map("scoop", "scoop update altrun", "winget", "winget upgrade zhugecaomao.ALTRun")
    static StateFile   := AppSettings.DataDir "\Update.json"                    ; {LastCheck, Skip}
    static CheckHours  := 6                                                 ; while running: check when the last check is this many hours old
    static StartupHours := 1                                                ; at start-up: check when the last check is this many hours old
    static Pending     := ""                                                ; a new version found in the background (the result of ParseRelease + Mode)
    static _timer      := "", _checking := false

    ; Downloads the release info in the background (ALTRun doesn't hang) and reports when done; clicking again during a check doesn't start another
    static Check(install := false) {
        if UpdateChecker._checking
            return
        UpdateChecker._checking := true
        UpdateChecker.FetchAsync((release) => (UpdateChecker._checking := false, UpdateChecker._Checked(release, install))
            , (message) => (UpdateChecker._checking := false, UpdateChecker._CheckFailed(message)))
    }

    static _CheckFailed(message) {
        Logger.Error("UpdateChecker: " message)
        MsgBox(I18n.T("Update.Failed", UpdateChecker.FailureText(message)), App.Name, 48)
    }

    ; What went wrong, in words people understand, with the raw error as the last line (for bug reports):
    ;   HTTP 403 / 429 = GitHub's hourly limit for requests without an account, counted per IP address, so on a shared
    ;   network (company, ISP NAT) other programs can use it up
    static FailureText(message) {
        static kinds := [["i)^HTTP (403|429)\b|rate limit", "Update.RateLimited"]
                       , ["i)^(Network error|Timed out)", "Update.Offline"]
                       , ["i)^HTTP 5\d\d\b", "Update.ServerError"]
                       , ["i)^HTTP 4\d\d\b", "Update.NotFound"]
                       , ["i)SHA256 does not match", "Update.Corrupt"]
                       , ["i)extract|missing from the update package", "Update.BadPackage"]
                       , ["i)tag_name|JSON", "Update.BadResponse"]]
        reason := I18n.T("Update.Unknown")
        for kind in kinds
            if RegExMatch(message, kind[1]) {
                reason := I18n.T(kind[2])
                break
            }
        return reason "`n`n" I18n.T("Update.Detail", message)
    }

    static _Checked(release, install) {
        state := UpdateChecker.LoadState()
        state["LastCheck"] := A_Now
        UpdateChecker.SaveState(state)
        if (UpdateChecker.Compare(release.Version, App.Version) <= 0) {
            UpdateChecker.Pending := ""
            MsgBox(I18n.T("Update.Latest", App.Version), App.Name, 64)
            return
        }
        UpdateChecker.SetPending(release)                                   ; with "Later" the result stays in the search window
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
    ; Automatic background checks + the update notice in the search window
    ;---------------------------------------------------------------------------
    ; Called at start-up: the network may not be up yet at boot, so the first check is after 1 minute
    static Schedule() {
        UpdateChecker._timer := () => UpdateChecker._Tick()
        SetTimer(() => UpdateChecker._Tick(UpdateChecker.StartupHours), -60000)
    }

    static _Tick(hours := "") {
        if UpdateChecker.IsDue(UpdateChecker.LoadState()["LastCheck"], A_Now, hours)
            UpdateChecker.CheckInBackground()
        SetTimer(UpdateChecker._timer, -3600000)
    }

    ; Whether the last check (yyyyMMddHHmmss) is at least hours old (default CheckHours); also checks when never checked or the time is invalid (system clock changed)
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

    ; Silent check: errors only go to the log (the time isn't recorded, so it retries next hour); a new version (not a skipped one) goes into Pending
    static CheckInBackground() {
        UpdateChecker.FetchAsync((release) => UpdateChecker._CheckedInBackground(release), (message) => Logger.Error("UpdateChecker: " message))
    }

    static _CheckedInBackground(release) {
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

    ; Notes the new version and decides how to update (Mode): "scoop" / "winget" / "source" (running the source) / "install" (one-click update) / "page" (open the download page)
    static SetPending(release) {
        manager := UpdateChecker.InstalledBy(A_ScriptDir)
        release.Mode := (manager != "") ? manager : !A_IsCompiled ? "source" : UpdateChecker.CanInstall(release) ? "install" : "page"
        UpdateChecker.Pending := release
    }

    ; The "New version available: ALTRun x" result in the search window (empty search box, or shown by SystemProvider when searching "update"); "" without a new version
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

    ; Enter in the search window: one-click update / open the download page / copy the package manager's upgrade command
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
            JSON.WriteFile(UpdateChecker.StateFile, state)
        } catch as e {
            Logger.Error("UpdateChecker.SaveState: " e.Message)
        }
    }

    ;---------------------------------------------------------------------------
    ; GitHub Release
    ;---------------------------------------------------------------------------
    ; Downloads the latest release info in the background: done(release) (the result of ParseRelease), fail(message) on errors
    static FetchAsync(done, fail) {
        HttpGet.Start(UpdateChecker.ApiUrl, (text) => done.Call(UpdateChecker.ParseRelease(text)), fail)
    }

    ; GitHub API releases/latest -> {Version, Page, ZipUrl, Sha256} ("" without zip / SHA256)
    static ParseRelease(text) {
        data := JSON.Parse(text)
        if !(data is Map) || !data.Has("tag_name") || data["tag_name"] = ""
            throw Error("Cannot find 'tag_name' in the GitHub response.")
        release := {Version: Trim(data["tag_name"], "vV "), Page: UpdateChecker.ReleasePage, ZipUrl: "", Sha256: ""}
        ; The page is opened with Run and the zip is downloaded: only addresses in this repository are used, whatever
        ; else the response says (the SHA256 comes from the same response, so it only catches a broken download)
        if data.Has("html_url") && UpdateChecker.InRepo(data["html_url"])
            release.Page := data["html_url"]
        if data.Has("assets") && data["assets"] is Array {
            for asset in data["assets"] {
                if !(asset is Map) || !asset.Has("name") || !RegExMatch(asset["name"], "i)^ALTRun_v?[\d.]+\.zip$")
                    continue
                release.ZipUrl := (asset.Has("browser_download_url") && UpdateChecker.InRepo(asset["browser_download_url"], "releases/download/"))
                    ? asset["browser_download_url"] : ""
                if asset.Has("digest") && RegExMatch(asset["digest"], "i)^sha256:([0-9a-f]{64})$", &m)
                    release.Sha256 := StrLower(m[1])
                break
            }
        }
        return release
    }

    ; An https address in this repository (under sub, e.g. "releases/download/"), with no tricks such as "..", spaces or quotes
    static InRepo(address, sub := "") {
        return (address is String) && InStr(address, UpdateChecker.RepoUrl sub, true) = 1 && !RegExMatch(address, '\.\.|[\s"<>^|&]')
    }

    ; Whether it can update itself: a compiled exe, a Release with zip and SHA256, and a writable program folder
    static CanInstall(release) {
        return A_IsCompiled && release.ZipUrl != "" && release.Sha256 != "" && UpdateChecker.IsWritable(A_ScriptDir)
    }

    static IsWritable(dir) => Path.IsWritable(dir)

    ; Installed with a package manager: suggest upgrading with it (replacing files itself would leave the package manager's records wrong):
    ;   Scoop   ...\scoop\apps\altrun\current (or a version folder)
    ;   winget  %LOCALAPPDATA%\Microsoft\WinGet\Packages\zhugecaomao.ALTRun_...
    static InstalledBy(dir) {
        if RegExMatch(dir, "i)\\apps\\altrun\\[^\\]+$")
            return "scoop"
        if RegExMatch(dir, "i)\\WinGet\\Packages\\zhugecaomao\.ALTRun")
            return "winget"
        return ""
    }

    ; Download -> verify SHA256 -> extract -> replace -> start the new version; on errors offer to open the download page
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
            if (MsgBox(I18n.T("Update.InstallFailed", UpdateChecker.FailureText(e.Message)), App.Name, "YesNo Icon!") = "Yes")
                Run(release.Page)
            return
        }
        progress.Destroy()
        Logger.Debug("UpdateChecker: updated to " release.Version)
        App.Restart("-Updated " release.Version)
    }

    ; zip -> folder. First tar.exe, built in since Windows 10 (1803) (fast, not limited by PowerShell policies);
    ; PowerShell's Expand-Archive if it is missing or fails
    static Extract(zip, dest) {
        DirCreate(dest)
        tar := A_WinDir "\System32\tar.exe"
        if FileExist(tar) {
            try exitCode := RunWait('"' tar '" -xf "' zip '" -C "' dest '"', , "Hide")
            if (IsSet(exitCode) && exitCode = 0 && FileExist(dest "\" UpdateChecker.PackageExe))
                return
            Logger.Debug("UpdateChecker: tar.exe could not extract the package, trying PowerShell")
        }
        quote := (s) => "'" StrReplace(s, "'", "''") "'"
        cmd := 'powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "'
             . "$ErrorActionPreference = 'Stop'; Expand-Archive -LiteralPath " quote(zip) " -DestinationPath " quote(dest) ' -Force"'
        exitCode := RunWait(cmd, , "Hide")
        if (exitCode != 0 || !FileExist(dest "\" UpdateChecker.PackageExe))
            throw Error("Could not extract the downloaded package (exit code " exitCode ").")
    }

    ; Copies the new version in src to dest: other files first, dest\exeName last (renamed to .old, then the new one copied).
    ; Overwritten files are backed up to src.backup first; any failing step restores (overwritten ones put back, new ones deleted) and then throws
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
                FileDelete(old)                                             ; left by the last update: stop if it can't be deleted
            if FileExist(exe) {
                FileMove(exe, old)
                movedExe := true
            }
            FileCopy(newExe, exe)
        } catch as e {
            if movedExe {
                try FileDelete(exe)
                try FileMove(old, exe)
                catch as restoreError                                       ; ALTRun.exe may be missing now: say where the old one is
                    Logger.Error("UpdateChecker: cannot put back " exe " from " old " - " restoreError.Message)
            }
            UpdateChecker._Undo(changes)
            try DirDelete(backup, true)
            throw e
        }
        try DirDelete(backup, true)
        for rel in UpdateChecker.ObsoleteFiles
            if !FileExist(src "\" rel)                                    ; keep what the new version still has
                try FileDelete(dest "\" rel)
    }

    ; Copies level by level (building paths from file names: A_Temp and Loop Files may give one a long name and the other an 8.3 short name,
    ; so relative paths can't be cut by length). Existing files are copied to backup before being overwritten and noted in changes so they can be restored
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
            catch as e
                Logger.Error("UpdateChecker: cannot restore " pair[1] " from " pair[2] " - " e.Message)
        for created in changes.Created
            try FileDelete(created)
        Loop changes.Dirs.Length                                            ; delete the new (empty) folders from the inside out
            try DirDelete(changes.Dirs[changes.Dirs.Length - A_Index + 1])
    }

    ; At the new version's start-up: deletes the renamed old program (the old process may not have exited yet; retried a little later) and the downloaded temporary files
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

    ; Compares the numbers between the dots one by one: > 0 means v1 is newer
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

    ; The new version dialog -> "Install" / "Notes" / "" (later)
    static _Ask(release) {
        choice := ""
        dlg := Gui("+AlwaysOnTop -MinimizeBox -MaximizeBox", I18n.T("Update.DialogTitle"))
        dlg.SetFont("s10")
        dlg.MarginX := 20, dlg.MarginY := 16
        dlg.AddText("w470", I18n.T("Update.Prompt", release.Version, App.Version))
        close := (value) => (choice := value, dlg.Destroy())
        dlg.AddButton("xm y+18 w150 Default", I18n.T("Update.InstallNow")).OnEvent("Click", (*) => close("Install"))
        dlg.AddButton("x+10 w150", I18n.T("Update.ReleaseNotes")).OnEvent("Click", (*) => close("Notes"))
        dlg.AddButton("x+10 w150", I18n.T("Update.Later")).OnEvent("Click", (*) => close(""))
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
