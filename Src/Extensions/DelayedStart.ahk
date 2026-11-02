;===============================================================================
; DelayedStart.ahk - start programs some time after signing in to Windows (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Programs that don't need to be there at once (OneDrive, sync tools, chat) can wait a few minutes, so the desktop is
; ready sooner. Set in Preferences → Delayed Start; ALTRun.json keeps the list (settings key "DelayedStart"):
;   [ { "Name": "OneDrive", "Target": "%LocalAppData%\\Microsoft\\OneDrive\\OneDrive.exe", "Arguments": "/background",
;       "Delay": 300, "Enabled": 1 } ]                                    Delay in seconds
;
; Each entry becomes a Windows Task Scheduler task (see Lib\TaskScheduler.ahk) in the folder \ALTRun\<user name>,
; triggered at this user's sign-in with the delay. Windows starts the programs, so they start even when ALTRun doesn't
; run or starts later, and the delay counts from signing in. The folder belongs to ALTRun: Sync makes it match the list,
; adding, changing and deleting tasks there (an empty list deletes the folder).
;
; When it syncs: after Preferences is saved with a changed list (problems are shown), and 30 seconds after start-up when
; the list isn't empty (problems are logged), which sets the tasks up on another PC using the same settings file and
; follows environment variables and relative paths there. A task already right isn't registered again.
;
; Usage:
;   DelayedStart.Sync()            -> "" or the problems, one per line
;   DelayedStart.Schedule()        at start-up
;   DelayedStart.EditorFields() / NewItem()   for the list in Preferences
;===============================================================================

class DelayedStart {
    static StartupSyncMs := 30000
    static Description := "Started by ALTRun (Preferences → Delayed Start). ALTRun keeps this folder in step with its list: changes made here are replaced."

    static Items => AppSettings.Data["DelayedStart"]
    static FolderPath => "\ALTRun\" DelayedStart.SafeName(A_UserName)

    static NewItem() => Map("Name", "", "Target", "", "Arguments", "", "Delay", 300, "Enabled", 1)

    static Schedule() {
        if DelayedStart.Items.Length
            WinTimer.Set(DelayedStart._StartupSync, -DelayedStart.StartupSyncMs)
    }

    static _StartupSync() {
        problems := DelayedStart.Sync()
        if (problems != "")
            Logger.Warn("DelayedStart: " StrReplace(problems, "`n", "; "))
    }

    ; Makes the tasks in FolderPath match items; returns "" or the problems, one per line (the other entries are still synced)
    static Sync(items := DelayedStart.Items, folderPath := DelayedStart.FolderPath) {
        try existing := TaskScheduler.Tasks(folderPath)
        catch as e
            return I18n.T("DelayedStart.NoScheduler", e.Message)
        plan := DelayedStart.Plan(DelayedStart.Desired(items), existing)
        problems := ""
        for name in plan.Delete {
            try TaskScheduler.Delete(folderPath, name)
            catch as e
                problems .= name ": " e.Message "`n"
        }
        for entry in plan.Register {
            try TaskScheduler.RegisterLogonTask(folderPath, entry.Name, entry.Spec)
            catch as e
                problems .= entry.Name ": " e.Message "`n"
        }
        if (!items.Length && problems = "")
            try TaskScheduler.DeleteFolder(folderPath)
        if (plan.Delete.Length || plan.Register.Length)
            Logger.Debug("DelayedStart: " plan.Register.Length " registered, " plan.Delete.Length " deleted in " folderPath)
        return RTrim(problems, "`n")
    }

    ; The tasks the list asks for: [{Name, Spec: {Path, Arguments, WorkingDirectory, Delay, Enabled, Description}}...]; entries without a
    ; program are left out. Task names come from the entry names, made safe for Task Scheduler and unique (OneDrive, OneDrive (2)...)
    static Desired(items) {
        list := [], used := Map()
        for item in items {
            target := Trim(DelayedStart._Get(item, "Target"), "`" `t")
            if (target = "")
                continue
            name := DelayedStart.SafeName(DelayedStart._Get(item, "Name"))
            if (name = "")
                name := DelayedStart.SafeName(Path.Leaf(target))
            unique := name, n := 1
            while used.Has(StrLower(unique))
                unique := name " (" (++n) ")"
            used[StrLower(unique)] := true
            program := DelayedStart.ProgramPath(target)
            SplitPath(program, , &dir)
            delay := DelayedStart._Get(item, "Delay")
            list.Push({Name: unique, Spec: {Path: InStr(program, " ") ? '"' program '"' : program   ; quoted as typed in Task Scheduler: a path with spaces works either way there
                , Arguments: Trim(DelayedStart._Get(item, "Arguments")), WorkingDirectory: dir
                , Delay: IsInteger(delay) ? Max(0, Integer(delay)) : 0, Enabled: DelayedStart._Get(item, "Enabled", 1) ? 1 : 0
                , Description: DelayedStart.Description}})
        }
        return list
    }

    ; What to change so that the tasks in the folder (existing, from TaskScheduler.Tasks) are the desired ones: {Register: [...], Delete: [names]}.
    ; A task that is already right isn't registered again; tasks no entry asks for are deleted (the folder is ALTRun's)
    static Plan(desired, existing) {
        byName := Map(), wanted := Map(), plan := {Register: [], Delete: []}
        for task in existing
            byName[StrLower(task.Name)] := task
        for entry in desired {
            wanted[StrLower(entry.Name)] := true
            key := StrLower(entry.Name)
            if !(byName.Has(key) && DelayedStart.Same(entry.Spec, byName[key]))
                plan.Register.Push(entry)
        }
        for task in existing
            if !wanted.Has(StrLower(task.Name))
                plan.Delete.Push(task.Name)
        return plan
    }

    static Same(spec, task) {
        return Trim(spec.Path, '"') = Trim(task.Path, '"') && spec.Arguments == task.Arguments && spec.WorkingDirectory = task.WorkingDirectory
            && spec.Delay = task.Delay && spec.Enabled = task.Enabled
    }

    ; The program a task runs: environment variables and A_ folders expanded, a relative path from the ALTRun folder, a bare name
    ; ("notepad.exe") found in PATH as when run from the search window. Task Scheduler runs a full path the same at every sign-in;
    ; the start-up sync follows changes (another PC, a moved ALTRun folder)
    static ProgramPath(target) {
        resolved := Path.Resolve(target)
        if (InStr(resolved, "\") && !RegExMatch(resolved, "^([A-Za-z]:|\\\\)"))
            resolved := Path.Full(resolved)
        return resolved
    }

    ; A task / folder name Task Scheduler takes: no \ / : * ? " < > | or control characters, no dots or spaces at the ends
    static SafeName(text) {
        text := RegExReplace(text, '[\\/:*?"<>|\x00-\x1F]', "_")
        return Trim(text, " .")
    }

    ; Delay for the list: "5 min", "30 s", "1 min 30 s", "No delay"
    static DelayText(seconds) {
        if !IsInteger(seconds) || seconds <= 0
            return I18n.T("DelayedStart.NoDelay")
        minutes := seconds // 60, rest := Mod(seconds, 60)
        if !minutes
            return I18n.T("DelayedStart.Seconds", rest)
        return rest ? I18n.T("DelayedStart.MinSec", minutes, rest) : I18n.T("DelayedStart.Minutes", minutes)
    }

    ;---------------------------------------------------------------------------
    ; Programs that start by themselves at sign-in (a warning when one is added here too: it would start at once, not after the delay)
    ;---------------------------------------------------------------------------
    ; The entry starting this program at sign-in ("Run: OneDrive", "Startup folder: Teams.lnk"), "" when none.
    ; Compares file names: the same program may be installed per user or for all users
    static AutoStartEntry(target) {
        leaf := StrLower(Path.Leaf(DelayedStart.ProgramPath(target)))
        if (leaf = "")
            return ""
        for entry in DelayedStart.AutoStartEntries()
            if (StrLower(Path.Leaf(entry.Program)) = leaf)
                return entry.Where ": " entry.Name
        return ""
    }

    ; The programs Windows starts at sign-in for this user: the Run keys (HKCU, HKLM, 32-bit HKLM) and the Startup folders,
    ; without those turned off in Task Manager → Startup apps (StartupApproved)
    static AutoStartEntries() {
        entries := []
        runKey := "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
        approved := "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\"
        for source in [["HKCU\" runKey, "HKCU" approved "Run"], ["HKLM\" runKey, "HKLM" approved "Run"]
                     , ["HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run", "HKLM" approved "Run32"]] {
            try {
                Loop Reg, source[1], "V" {
                    if (A_LoopRegType != "REG_SZ" && A_LoopRegType != "REG_EXPAND_SZ")
                        continue
                    name := A_LoopRegName
                    if DelayedStart._TurnedOff(source[2], name)
                        continue
                    entries.Push({Where: "Run", Name: name, Program: DelayedStart.CommandProgram(Path.ExpandEnv(RegRead()))})
                }
            }
        }
        for folder in [[A_Startup, "HKCU" approved "StartupFolder"], [A_StartupCommon, "HKLM" approved "StartupFolder"]] {
            Loop Files, folder[1] "\*.lnk" {
                if DelayedStart._TurnedOff(folder[2], A_LoopFileName)
                    continue
                target := ""
                try FileGetShortcut(A_LoopFilePath, &target)
                if (target != "")
                    entries.Push({Where: "Startup", Name: A_LoopFileName, Program: target})
            }
        }
        return entries
    }

    ; StartupApproved holds 12 bytes per entry; an odd first byte (01, 03) = turned off in Task Manager
    static _TurnedOff(key, name) {
        try data := RegRead(key, name)
        catch
            return false
        return (StrLen(data) >= 2 && Mod(Integer("0x" SubStr(data, 1, 2)), 2) = 1)
    }

    ; The program of a command line: "C:\Program Files\x.exe" /a -> C:\Program Files\x.exe; C:\x\y.exe /a -> C:\x\y.exe
    static CommandProgram(command) {
        command := Trim(command)
        if RegExMatch(command, '^"([^"]+)"', &m)
            return m[1]
        if RegExMatch(command, "i)^(.+?\.(exe|com|bat|cmd))(\s|$)", &m)
            return m[1]
        return RegExReplace(command, "\s.*$")
    }

    ;---------------------------------------------------------------------------
    ; Preferences
    ;---------------------------------------------------------------------------
    static EditorFields() {
        target := ItemEditor.Field("Target", "Prefs.Col.Program", "file", true, "", I18n.T("DelayedStart.Field.Target"))
        target.Check := (value, edited) => DelayedStart._AutoStartWarning(value)
        return [ItemEditor.Field("Name", "Prefs.Col.Name", "text", true, "", I18n.T("DelayedStart.Field.Name"))
              , target
              , ItemEditor.Field("Arguments", "Prefs.Col.Arguments", "text", false, "", I18n.T("DelayedStart.Field.Arguments"))
              , ItemEditor.Field("Delay", "Prefs.Col.DelaySeconds", "number", false, "", I18n.T("DelayedStart.Field.Delay"))
              , ItemEditor.Field("Enabled", "DelayedStart.Field.Enabled", "check")]
    }

    static _AutoStartWarning(target) {
        entry := DelayedStart.AutoStartEntry(target)
        return (entry = "") ? "" : I18n.T("DelayedStart.AlsoAutoStarts", Path.Leaf(DelayedStart.ProgramPath(target)), entry)
    }

    ; Preferences was saved: sync when the list changed; returns "" or the problems
    static AfterSave(before, after) {
        if (JSON.Stringify(before) == JSON.Stringify(after))
            return ""
        return DelayedStart.Sync(after)
    }

    static _Get(item, key, default := "") => (item is Map && item.Has(key)) ? item[key] : default
}
