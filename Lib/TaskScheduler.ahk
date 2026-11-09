;===============================================================================
; TaskScheduler.ahk - "run at my sign-in, after a delay" tasks in Windows Task Scheduler (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Through the Task Scheduler 2.0 COM interface (Schedule.Service), not schtasks.exe: schtasks /SC ONLOGON needs an
; administrator, while a standard user may register a logon task for their own account through COM. Checked on Windows
; by .github/workflows/delayed-start.yml (as a standard user, Tests\Integration\DelayedStartCheck.ahk). Also found there:
; a task folder another user made can be used, and that user's tasks in it can't be seen, so each user's tasks stay apart.
;
; The tasks run the program in the user's session with the user's normal rights (an interactive token: no password is
; stored, nothing runs while the user is signed out). Settings that differ from Task Scheduler's defaults, which suit
; maintenance jobs rather than programs people use:
;   - starts on battery and keeps running when the power is unplugged (default: neither)
;   - no time limit (default: stopped after 3 days)
;   - normal priority, 5 (default 7: below normal CPU, I/O and memory priority for as long as the program runs)
;
; Usage:
;   TaskScheduler.Tasks("\ALTRun\me")                           [{Name, Path, Arguments, WorkingDirectory, Delay, Enabled}...], [] when no such folder
;   TaskScheduler.RegisterLogonTask("\ALTRun\me", "OneDrive", {Path, Arguments, WorkingDirectory, Delay, Enabled, Description})
;       Delay in seconds; creates the folder when missing, replaces a task of the same name
;   TaskScheduler.Delete("\ALTRun\me", "OneDrive")
;   TaskScheduler.DeleteFolder("\ALTRun\me")                    only when empty; its parent too when that is empty then
;   Each takes an optional last argument service (TaskScheduler.Service()): several calls in a row share one connection
;   TaskScheduler.Duration(330) = "PT5M30S"    TaskScheduler.Seconds("PT5M30S") = 330
; Errors (no access, Task Scheduler service not running) are thrown as they come from COM.
;===============================================================================

class TaskScheduler {
    static TASK_TRIGGER_LOGON := 9, TASK_ACTION_EXEC := 0, TASK_CREATE_OR_UPDATE := 6, TASK_LOGON_INTERACTIVE_TOKEN := 3
    static TASK_INSTANCES_IGNORE_NEW := 2, Priority := 5

    ; DOMAIN\user (or PC\user) of the user running ALTRun, as Task Scheduler names accounts
    static UserId() {
        static id := ""
        if (id = "") {
            size := 256, buf := Buffer(size * 2, 0)
            if DllCall("secur32\GetUserNameExW", "Int", 2, "Ptr", buf, "UInt*", &size)   ; NameSamCompatible
                id := StrGet(buf)
            else
                id := EnvGet("USERDOMAIN") "\" A_UserName
        }
        return id
    }

    static Service() {
        service := ComObject("Schedule.Service")
        service.Connect()
        return service
    }

    ; The folder at folderPath ("\ALTRun\me"), or "" when it doesn't exist (create: make it, with the folders above it)
    static Folder(service, folderPath, create := false) {
        try return service.GetFolder(folderPath)
        if !create
            return ""
        return service.GetFolder("\").CreateFolder(LTrim(folderPath, "\"))
    }

    static Tasks(folderPath, service := "") {
        list := []
        if !IsObject(folder := TaskScheduler.Folder(IsObject(service) ? service : TaskScheduler.Service(), folderPath))
            return list
        for task in folder.GetTasks(1) {                                    ; 1 = hidden ones too
            item := {Name: task.Name, Enabled: task.Enabled ? 1 : 0, Path: "", Arguments: "", WorkingDirectory: "", Delay: -1}
            definition := task.Definition
            if (definition.Actions.Count >= 1 && definition.Actions.Item(1).Type = TaskScheduler.TASK_ACTION_EXEC) {
                action := definition.Actions.Item(1)
                item.Path := action.Path, item.Arguments := action.Arguments, item.WorkingDirectory := action.WorkingDirectory
            }
            if (definition.Triggers.Count = 1 && definition.Triggers.Item(1).Type = TaskScheduler.TASK_TRIGGER_LOGON)
                item.Delay := TaskScheduler.Seconds(definition.Triggers.Item(1).Delay)   ; -1 = not one logon trigger: never taken as up to date
            list.Push(item)
        }
        return list
    }

    static RegisterLogonTask(folderPath, name, spec, service := "") {
        service := IsObject(service) ? service : TaskScheduler.Service()
        folder := TaskScheduler.Folder(service, folderPath, true)
        user := TaskScheduler.UserId()
        definition := service.NewTask(0)
        definition.RegistrationInfo.Author := "ALTRun"
        definition.RegistrationInfo.Description := spec.Description
        trigger := definition.Triggers.Create(TaskScheduler.TASK_TRIGGER_LOGON)
        trigger.UserId := user                                              ; only this user's sign-in (any user's needs an administrator)
        trigger.Delay := TaskScheduler.Duration(spec.Delay)
        action := definition.Actions.Create(TaskScheduler.TASK_ACTION_EXEC)
        action.Path := spec.Path
        action.Arguments := spec.Arguments
        action.WorkingDirectory := spec.WorkingDirectory
        definition.Principal.UserId := user
        definition.Principal.LogonType := TaskScheduler.TASK_LOGON_INTERACTIVE_TOKEN
        definition.Principal.RunLevel := 0                                  ; TASK_RUNLEVEL_LUA: the user's normal rights
        settings := definition.Settings
        settings.Enabled := spec.Enabled ? true : false
        settings.MultipleInstances := TaskScheduler.TASK_INSTANCES_IGNORE_NEW
        settings.DisallowStartIfOnBatteries := false
        settings.StopIfGoingOnBatteries := false
        settings.ExecutionTimeLimit := "PT0S"
        settings.Priority := TaskScheduler.Priority
        folder.RegisterTaskDefinition(name, definition, TaskScheduler.TASK_CREATE_OR_UPDATE, , , TaskScheduler.TASK_LOGON_INTERACTIVE_TOKEN)
    }

    static Delete(folderPath, name, service := "") {
        if IsObject(folder := TaskScheduler.Folder(IsObject(service) ? service : TaskScheduler.Service(), folderPath))
            folder.DeleteTask(name, 0)
    }

    ; Deletes the folder if it holds no tasks and no folders (tasks of other users can't be seen: Task Scheduler then refuses, which is fine)
    static DeleteFolder(folderPath, service := "") {
        service := IsObject(service) ? service : TaskScheduler.Service()
        Loop 2 {
            if !IsObject(folder := TaskScheduler.Folder(service, folderPath))
                return
            if (folder.GetTasks(1).Count || folder.GetFolders(0).Count)
                return
            SplitPath(folderPath, &leaf, &parent)
            if (parent = "" || parent = "\")
                parent := "\"
            try service.GetFolder(parent).DeleteFolder(leaf, 0)
            catch
                return
            if (parent = "\")
                return
            folderPath := parent
        }
    }

    ; Seconds -> an XML duration as Task Scheduler takes it ("PT5M", "PT1H30S"); 0 -> "" (no delay)
    static Duration(seconds) {
        seconds := Max(0, Integer(seconds))
        if !seconds
            return ""
        text := "PT"
        if (hours := seconds // 3600)
            text .= hours "H"
        if (minutes := Mod(seconds, 3600) // 60)
            text .= minutes "M"
        if (rest := Mod(seconds, 60))
            text .= rest "S"
        return text
    }

    ; An XML duration -> seconds ("" -> 0); days count, months and years don't occur in a delay; -1 when it can't be read
    static Seconds(duration) {
        if (duration = "")
            return 0
        if !RegExMatch(duration, "i)^P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$", &m)
            return -1
        return (m[1] = "" ? 0 : m[1]) * 86400 + (m[2] = "" ? 0 : m[2]) * 3600 + (m[3] = "" ? 0 : m[3]) * 60 + (m[4] = "" ? 0 : m[4])
    }
}
