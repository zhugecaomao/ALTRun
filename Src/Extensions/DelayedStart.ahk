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
;   DelayedStart.PickStartupApp(owner, items) / AfterSave(before, after)   startup apps turned into entries (see below)
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
    ; Programs that start by themselves at sign-in: the startup apps of Task Manager → Startup apps (Run keys and Startup folders).
    ; One can be turned into a delayed start entry; its own start is then turned off the way Task Manager does it (StartupApproved),
    ; nothing is deleted, and it can be turned on again there or by removing the entry here
    ;---------------------------------------------------------------------------
    static ApprovedKey := "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\"

    ; Where startup apps are: scope -> {Run: registry key or "", Folder: folder or "", Approved: its StartupApproved key, Writable: this user
    ; can turn them off (the all-users ones need an administrator)}
    static Scopes() {
        runKey := "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", approved := DelayedStart.ApprovedKey
        return Map(
            "HKCU",   {Run: "HKCU\" runKey, Folder: "", Approved: "HKCU" approved "Run", Writable: true},
            "HKLM",   {Run: "HKLM\" runKey, Folder: "", Approved: "HKLM" approved "Run", Writable: false},
            "HKLM32", {Run: "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run", Folder: "", Approved: "HKLM" approved "Run32", Writable: false},
            "User",   {Run: "", Folder: A_Startup, Approved: "HKCU" approved "StartupFolder", Writable: true},
            "Common", {Run: "", Folder: A_StartupCommon, Approved: "HKLM" approved "StartupFolder", Writable: false})
    }

    ; The startup apps: [{Source: "HKCU|OneDrive", Scope, Name, Command, Program, Arguments, Off, Writable}...];
    ; includeOff: also those turned off in Task Manager
    static AutoStartEntries(includeOff := false) {
        entries := []
        for scope, where in DelayedStart.Scopes() {
            if (where.Run != "") {
                try {
                    Loop Reg, where.Run, "V" {
                        if (A_LoopRegType != "REG_SZ" && A_LoopRegType != "REG_EXPAND_SZ")
                            continue
                        name := A_LoopRegName, command := Path.ExpandEnv(RegRead())
                        entries.Push(DelayedStart._Entry(scope, where, name, command, DelayedStart.CommandProgram(command), DelayedStart.CommandArguments(command)))
                    }
                }
            } else if (where.Folder != "") {
                Loop Files, where.Folder "\*.lnk" {
                    target := "", args := ""
                    try FileGetShortcut(A_LoopFilePath, &target, , &args)
                    if (target != "")
                        entries.Push(DelayedStart._Entry(scope, where, A_LoopFileName, Trim('"' target '" ' args), target, args))
                }
            }
        }
        if includeOff
            return entries
        on := []
        for entry in entries
            if !entry.Off
                on.Push(entry)
        return on
    }

    static _Entry(scope, where, name, command, program, args) {
        off := false
        try off := DelayedStart.IsOffData(RegRead(where.Approved, name))
        return {Source: scope "|" name, Scope: scope, Name: name, Command: command, Program: program, Arguments: args, Off: off, Writable: where.Writable}
    }

    ; The startup app a delayed start entry came from ("" when it is gone)
    static FindEntry(source) {
        for entry in DelayedStart.AutoStartEntries(true)
            if (entry.Source = source)
                return entry
        return ""
    }

    ; The startup app starting this program at sign-in (still on), "" when none. Compares file names: the same program may be installed
    ; per user or for all users
    static AutoStartEntry(target) {
        leaf := StrLower(Path.Leaf(DelayedStart.ProgramPath(target)))
        if (leaf = "")
            return ""
        for entry in DelayedStart.AutoStartEntries()
            if (StrLower(Path.Leaf(entry.Program)) = leaf)
                return entry
        return ""
    }

    ; Where a startup app is, for messages: "Run (this user): OneDrive", "Startup folder (all users): Teams.lnk"
    static SourceText(source) {
        parts := StrSplit(source, "|", , 2)
        return I18n.T("DelayedStart.Where." parts[1]) ": " (parts.Length > 1 ? parts[2] : "")
    }

    ; Turns a startup app off (or on again) as Task Manager does: its StartupApproved value. Throws when it can't be written
    static SetEntryOff(source, off) {
        parts := StrSplit(source, "|", , 2), scopes := DelayedStart.Scopes()
        if (parts.Length < 2 || !scopes.Has(parts[1]))
            throw ValueError("Unknown startup app: " source, -1)
        RegWrite(DelayedStart.ApprovedData(off), "REG_BINARY", scopes[parts[1]].Approved, parts[2])
    }

    ; A StartupApproved value: 12 bytes, the first 02 = on, 03 = off, then (when off) the time it was turned off (FILETIME)
    static ApprovedData(off, fileTime := "") {
        if !off
            return "020000000000000000000000"
        if (fileTime = "")
            DllCall("GetSystemTimeAsFileTime", "Int64*", &fileTime := 0)
        hex := "03000000"
        Loop 8
            hex .= Format("{:02X}", (fileTime >> (8 * (A_Index - 1))) & 0xFF)
        return hex
    }

    ; An odd first byte (01, 03) = turned off
    static IsOffData(data) => (StrLen(data) >= 2 && IsXDigit(SubStr(data, 1, 2)) && Mod(Integer("0x" SubStr(data, 1, 2)), 2) = 1)

    ; The program of a command line: "C:\Program Files\x.exe" /a -> C:\Program Files\x.exe; C:\x\y.exe /a -> C:\x\y.exe
    static CommandProgram(command) {
        command := Trim(command)
        if RegExMatch(command, '^"([^"]+)"', &m)
            return m[1]
        if RegExMatch(command, "i)^(.+?\.(exe|com|bat|cmd))(\s|$)", &m)
            return m[1]
        return RegExReplace(command, "\s.*$")
    }

    ; What follows the program: "C:\x\y.exe" /a /b -> /a /b
    static CommandArguments(command) {
        command := Trim(command)
        if RegExMatch(command, '^"[^"]+"\s*(.*)$', &m)
            return Trim(m[1])
        if RegExMatch(command, "i)^.+?\.(exe|com|bat|cmd)(\s+(.*))?$", &m)
            return Trim(m[3])
        return RegExMatch(command, "\s+(.*)$", &m) ? Trim(m[1]) : ""
    }

    ; A delayed start entry made from a startup app: its name (a shortcut's without .lnk), program and arguments, 5 minutes
    static ItemFromEntry(entry) {
        name := RegExReplace(entry.Name, "i)\.lnk$")
        return Map("Name", name, "Target", Path.Shorten(entry.Program), "Arguments", entry.Arguments, "Delay", 300, "Enabled", 1, "Source", entry.Source)
    }

    ; Sources added and removed between two lists (Preferences before and after): {Added: [...], Removed: [...]}
    static SourceChanges(before, after) {
        old := Map(), now := Map(), changes := {Added: [], Removed: []}
        for item in before
            if (source := DelayedStart._Get(item, "Source")) != ""
                old[source] := true
        for item in after
            if (source := DelayedStart._Get(item, "Source")) != ""
                now[source] := true
        for source in now
            if !old.Has(source)
                changes.Added.Push(source)
        for source in old
            if !now.Has(source)
                changes.Removed.Push(source)
        return changes
    }

    ;---------------------------------------------------------------------------
    ; Preferences
    ;---------------------------------------------------------------------------
    static EditorFields() {
        target := ItemEditor.Field("Target", "Prefs.Col.Program", "file", true, "", I18n.T("DelayedStart.Field.Target"))
        target.Check := (value, edited) => DelayedStart._CheckAutoStart(value, edited)
        return [ItemEditor.Field("Name", "Prefs.Col.Name", "text", true, "", I18n.T("DelayedStart.Field.Name"))
              , target
              , ItemEditor.Field("Arguments", "Prefs.Col.Arguments", "text", false, "", I18n.T("DelayedStart.Field.Arguments"))
              , ItemEditor.Field("Delay", "Prefs.Col.DelaySeconds", "number", false, "", I18n.T("DelayedStart.Field.Delay"))
              , ItemEditor.Field("Enabled", "DelayedStart.Field.Enabled", "check")]
    }

    ; OK in the edit dialog: the program also starts by itself? Offer to turn that off on Apply (Source); one for all users can't be turned off
    ; without an administrator: returns the warning ("Save anyway?")
    static _CheckAutoStart(target, edited) {
        entry := DelayedStart.AutoStartEntry(target)
        if (!IsObject(entry) || DelayedStart._Get(edited, "Source") = entry.Source)
            return ""
        program := Path.Leaf(DelayedStart.ProgramPath(target)), where := DelayedStart.SourceText(entry.Source)
        if !entry.Writable
            return I18n.T("DelayedStart.AlsoAutoStartsAdmin", program, where)
        if (MsgBox(I18n.T("DelayedStart.AlsoAutoStarts", program, where), I18n.T("Prefs.Page.DelayedStart"), "YesNo Icon? Owner" WinExist("A")) = "Yes")
            edited["Source"] := entry.Source
        return ""
    }

    ; "Startup Apps..." on the Preferences page: picks one of the startup apps not in the list yet; returns the new entry (Map), "" when none
    static PickStartupApp(owner, items) {
        used := Map(), entries := []
        for item in items
            if (source := DelayedStart._Get(item, "Source")) != ""
                used[source] := true
        for entry in DelayedStart.AutoStartEntries(true)
            if !used.Has(entry.Source)
                entries.Push(entry)
        title := I18n.T("DelayedStart.PickTitle")
        if !entries.Length {
            MsgBox(I18n.T("DelayedStart.NoStartupApps"), title, "Iconi Owner" owner.Hwnd)
            return ""
        }
        state := {Result: ""}
        g := Gui("+Owner" owner.Hwnd " -MinimizeBox", title)
        g.SetFont("s9", ThemeManager.FontName())
        g.MarginX := 14, g.MarginY := 12
        width := 640
        g.AddText("w" width, I18n.T("DelayedStart.PickDesc"))
        list := g.AddListView("xm y+10 w" width " h260 -Multi NoSort Grid"
            , [I18n.T("Prefs.Col.Name"), I18n.T("DelayedStart.Col.Command"), I18n.T("DelayedStart.Col.Where"), I18n.T("Prefs.Col.Enabled")])
        for entry in entries
            list.Add(, RegExReplace(entry.Name, "i)\.lnk$"), entry.Command, I18n.T("DelayedStart.Where." entry.Scope), I18n.T(entry.Off ? "DelayedStart.Off" : "DelayedStart.On"))
        for index, w in [130, 300, 140, 60]
            list.ModifyCol(index, w)
        list.Modify(1, "Select Focus")
        g.AddButton("xm+" (width - 170) " y+14 w80 Default", I18n.T("Prefs.OK")).OnEvent("Click", Choose)
        g.AddButton("x+10 yp w80", I18n.T("Prefs.Cancel")).OnEvent("Click", (*) => g.Destroy())
        list.OnEvent("DoubleClick", Choose)
        g.OnEvent("Escape", (*) => g.Destroy())
        g.OnEvent("Close", (*) => g.Destroy())
        owner.Opt("+Disabled")
        g.Show()
        WinWaitClose(g.Hwnd)
        owner.Opt("-Disabled")
        try WinActivate("ahk_id " owner.Hwnd)
        return state.Result

        Choose(*) {
            if !(row := list.GetNext())
                return
            state.Result := DelayedStart.ItemFromEntry(entries[row])
            g.Destroy()
        }
    }

    ; Preferences was saved: when the list changed, sync the tasks; then turn off the startup apps of new entries (only when every task was
    ; set up: otherwise a program might not start at all), and offer to turn on again those of removed entries. owner: for the messages
    static AfterSave(before, after, owner := 0) {
        if (JSON.Stringify(before) == JSON.Stringify(after))
            return
        title := I18n.T("Prefs.Page.DelayedStart"), options := owner ? " Owner" owner : ""
        problems := DelayedStart.Sync(after)
        if (problems != "") {
            MsgBox(I18n.T("DelayedStart.SyncFailed", problems), title, "Icon!" options)
            return
        }
        changes := DelayedStart.SourceChanges(before, after), manual := ""
        for source in changes.Added {
            if !IsObject(entry := DelayedStart.FindEntry(source)) || entry.Off
                continue
            if entry.Writable {
                try {
                    DelayedStart.SetEntryOff(source, true)
                    Logger.Debug("DelayedStart: turned off the startup app " source)
                    continue
                } catch as e
                    Logger.Warn("DelayedStart: cannot turn off " source " - " e.Message)
            }
            manual .= DelayedStart.SourceText(source) "`n"
        }
        if (manual != "" && MsgBox(I18n.T("DelayedStart.TurnOffManually", RTrim(manual, "`n")), title, "YesNo Iconi" options) = "Yes")
            DelayedStart.OpenStartupApps()
        for source in changes.Removed {
            if !IsObject(entry := DelayedStart.FindEntry(source)) || !entry.Off || !entry.Writable
                continue
            if (MsgBox(I18n.T("DelayedStart.TurnBackOn", DelayedStart.SourceText(source)), title, "YesNo Icon?" options) != "Yes")
                continue
            try DelayedStart.SetEntryOff(source, false)
            catch as e
                Logger.Warn("DelayedStart: cannot turn on " source " - " e.Message)
        }
    }

    ; Task Manager on its Startup apps page
    static OpenStartupApps() {
        try Run("taskmgr.exe /0 /startup")
        catch as e
            Logger.Warn("DelayedStart: Task Manager - " e.Message)
    }

    static _Get(item, key, default := "") => (item is Map && item.Has(key)) ? item[key] : default
}
