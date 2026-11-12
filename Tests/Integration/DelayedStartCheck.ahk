;===============================================================================
; DelayedStartCheck.ahk - delayed start against the real Windows Task Scheduler (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Run by Tests\RunTests.ahk DelayedStartCheck: on Windows CI by .github/workflows/delayed-start.yml, as a new standard
; (non-administrator) user and as the administrator; also works on your own PC (it only touches its own task folder,
; \ALTRun\<user name> check, and deletes it again). Registers tasks through DelayedStart.Sync, reads them back,
; changes and removes them; turns a startup app made for the check (HKCU Run) off and on again, then deletes it. Prints one line per check (PASS / FAIL / INFO); returns the number of failed checks.
;===============================================================================
class DelayedStartCheck {
    static Failures := 0

    static Run() {
        folderPath := DelayedStart.FolderPath " check"
        DelayedStartCheck.Log("INFO user " TaskScheduler.UserId() ", administrator " (A_IsAdmin ? "yes" : "no") ", folder " folderPath)
        try DelayedStart.Sync([], folderPath)                               ; left over from an interrupted run
        dir := A_Temp "\ALTRun check dir"
        DirCreate(dir)
        script := dir "\write marker.cmd", marker := dir "\marker.txt"
        try FileDelete(marker)
        FileOpen(script, "w", "CP0").Write('@echo started> "' marker '"`r`n')
        items := [Map("Name", "Notepad", "Target", "%SystemRoot%\System32\notepad.exe", "Arguments", "/a b.txt", "Delay", 300, "Enabled", 1)
                , Map("Name", "Marker: spaces", "Target", script, "Arguments", "", "Delay", 30, "Enabled", 0)]
        try {
            started := Logger.Ms()
            problems := DelayedStart.Sync(items, folderPath)
            DelayedStartCheck.Check("register two tasks", problems = "", problems)
            DelayedStartCheck.Log("INFO registering took " Round(Logger.Ms() - started) " ms")
            tasks := DelayedStartCheck.ByName(TaskScheduler.Tasks(folderPath))
            DelayedStartCheck.Check("both there", tasks.Has("Notepad") && tasks.Has("Marker_ spaces"), DelayedStartCheck.Names(tasks))
            if tasks.Has("Notepad") {
                t := tasks["Notepad"]
                DelayedStartCheck.Check("notepad: program, arguments, folder", t.Path = EnvGet("SystemRoot") "\System32\notepad.exe" && t.Arguments == "/a b.txt"
                    && t.WorkingDirectory = EnvGet("SystemRoot") "\System32", t.Path " | " t.Arguments " | " t.WorkingDirectory)
                DelayedStartCheck.Check("notepad: 5 minutes, on", t.Delay = 300 && t.Enabled = 1, t.Delay " s, enabled " t.Enabled)
            }
            if tasks.Has("Marker_ spaces") {
                t := tasks["Marker_ spaces"]
                DelayedStartCheck.Check("path with spaces: quoted", t.Path == '"' script '"', t.Path)
                DelayedStartCheck.Check("30 s, off", t.Delay = 30 && t.Enabled = 0, t.Delay " s, enabled " t.Enabled)
            }
            DelayedStartCheck.CheckDefinition(folderPath, "Notepad")

            ; Read back, the tasks count as right: nothing registered again (what the start-up check does)
            plan := DelayedStart.Plan(DelayedStart.Desired(items), TaskScheduler.Tasks(folderPath))
            DelayedStartCheck.Check("read back: nothing to change", plan.Register.Length = 0 && plan.Delete.Length = 0, plan.Register.Length " to register, " plan.Delete.Length " to delete")
            times := []
            Loop 5 {
                started := Logger.Ms()
                DelayedStart.Sync(items, folderPath)
                times.Push(Round(Logger.Ms() - started))
            }
            DelayedStartCheck.Log("INFO start-up check with nothing to change: " DelayedStartCheck.Join(times) " ms")

            DelayedStartCheck.RunOnDemand(folderPath, marker)

            ; Changed and removed entries
            items[1]["Delay"] := 120, items.RemoveAt(2)
            problems := DelayedStart.Sync(items, folderPath)
            tasks := DelayedStartCheck.ByName(TaskScheduler.Tasks(folderPath))
            DelayedStartCheck.Check("change and remove", problems = "" && tasks.Count = 1 && tasks.Has("Notepad") && tasks["Notepad"].Delay = 120
                , problems " " DelayedStartCheck.Names(tasks))

            ; An empty list deletes the tasks and the folder
            problems := DelayedStart.Sync([], folderPath)
            DelayedStartCheck.Check("empty list: folder deleted", problems = "" && !IsObject(TaskScheduler.Folder(TaskScheduler.Service(), folderPath)), problems)
        } catch as e {
            DelayedStartCheck.Check("no error", false, e.Message " (" e.What ", line " e.Line ")")
            try DelayedStart.Sync([], folderPath)
        }
        try DirDelete(dir, true)
        DelayedStartCheck.StartupApps()
        if A_IsAdmin                                                        ; a standard user would get a UAC prompt nobody answers
            DelayedStartCheck.TaskFromXml(folderPath)
        return DelayedStartCheck.Failures
    }

    ; The Caps Lock helper's task is registered by schtasks /create /xml run as administrator (CapsLockHelper.RegisterXml), while ALTRun
    ; holds the XML file open with a handle that refuses writing and deleting: schtasks must still be able to read it, and others must not change it
    static TaskFromXml(folderPath) {
        name := "CapsLock XML"
        xml := CapsLockHelper.TaskXml(EnvGet("SystemRoot") "\System32\notepad.exe", "/a b.txt", A_Temp, TaskScheduler.UserId())
        try {
            code := CapsLockHelper.RegisterXml(folderPath "\" name, xml)
            tasks := DelayedStartCheck.ByName(TaskScheduler.Tasks(folderPath))
            DelayedStartCheck.Check("task from XML held open: schtasks read it", code = 0 && tasks.Has(name) && tasks[name].Arguments == "/a b.txt"
                , "exit code " code ", " DelayedStartCheck.Names(tasks))
            heldFile := A_Temp "\ALTRun check held.xml"
            FileAppend("<Task/>", heldFile, "UTF-16")
            held := FileOpen(heldFile, "r-wd", "UTF-16")
            blocked := 0
            try FileAppend("x", heldFile)
            catch
                blocked++
            try FileDelete(heldFile)
            catch
                blocked++
            try FileMove(heldFile, heldFile ".moved")
            catch
                blocked++
            reader := FileOpen(heldFile, "r", "UTF-16")
            DelayedStartCheck.Check("held file: no writing, deleting, renaming; reading works", blocked = 3 && reader.Read() == "<Task/>", blocked " of 3 blocked")
            reader.Close(), held.Close()
            try FileDelete(heldFile)
        } catch as e
            DelayedStartCheck.Check("task from XML: no error", false, e.Message " (" e.What ", line " e.Line ")")
        try TaskScheduler.Delete(folderPath, name)
        try TaskScheduler.DeleteFolder(folderPath)
    }

    ; A startup app of this user (a Run value made for the check): listed, turned off and on again as Task Manager does it, made into an entry
    static StartupApps() {
        runKey := "HKCU\Software\Microsoft\Windows\CurrentVersion\Run", name := "ALTRun check startup app", source := "HKCU|" name
        approved := DelayedStart.Scopes()["HKCU"].Approved
        try {
            RegWrite('"' EnvGet("SystemRoot") '\System32\notepad.exe" /a b.txt', "REG_SZ", runKey, name)
            entry := DelayedStart.FindEntry(source)
            DelayedStartCheck.Check("startup app listed, on, this user can turn it off", IsObject(entry) && !entry.Off && entry.Writable
                && entry.Program = EnvGet("SystemRoot") "\System32\notepad.exe" && entry.Arguments == "/a b.txt", IsObject(entry) ? entry.Command : "not found")
            DelayedStartCheck.Check("found for its program", IsObject(found := DelayedStart.AutoStartEntry("%SystemRoot%\System32\notepad.exe")) && found.Source = source)
            DelayedStart.SetEntryOff(source, true)
            data := RegRead(approved, name)
            DelayedStartCheck.Check("turned off as in Task Manager", SubStr(data, 1, 8) = "03000000" && StrLen(data) = 24, data)
            DelayedStartCheck.Check("listed as off, not among those that start", DelayedStart.FindEntry(source).Off && !IsObject(DelayedStart.AutoStartEntry("%SystemRoot%\System32\notepad.exe")))
            DelayedStart.SetEntryOff(source, false)
            data := RegRead(approved, name)
            DelayedStartCheck.Check("turned on again", data = "020000000000000000000000" && !DelayedStart.FindEntry(source).Off, data)
            item := DelayedStart.ItemFromEntry(DelayedStart.FindEntry(source))
            DelayedStartCheck.Check("entry made from it", item["Name"] = name && item["Arguments"] == "/a b.txt" && item["Source"] = source)
        } catch as e {
            DelayedStartCheck.Check("startup apps: no error", false, e.Message " (" e.What ", line " e.Line ")")
        }
        try RegDelete(runKey, name)
        try RegDelete(approved, name)
    }

    ; The settings ALTRun sets that differ from Task Scheduler's defaults, and the sign-in trigger for this user only
    static CheckDefinition(folderPath, name) {
        definition := TaskScheduler.Service().GetFolder(folderPath).GetTask(name).Definition
        trigger := definition.Triggers.Item(1), settings := definition.Settings, principal := definition.Principal
        DelayedStartCheck.Check("trigger: this user's sign-in", trigger.Type = 9 && trigger.UserId = TaskScheduler.UserId(), trigger.Type " " trigger.UserId)
        DelayedStartCheck.Check("runs with normal rights, no password", principal.LogonType = 3 && principal.RunLevel = 0, principal.LogonType " " principal.RunLevel)
        DelayedStartCheck.Check("battery, time limit, priority", !settings.DisallowStartIfOnBatteries && !settings.StopIfGoingOnBatteries
            && settings.ExecutionTimeLimit = "PT0S" && settings.Priority = 5 && settings.MultipleInstances = 2
            , settings.DisallowStartIfOnBatteries " " settings.StopIfGoingOnBatteries " " settings.ExecutionTimeLimit " " settings.Priority " " settings.MultipleInstances)
        DelayedStartCheck.Check("author", definition.RegistrationInfo.Author = "ALTRun", definition.RegistrationInfo.Author)
    }

    ; Starting the task with the spaces in its path by hand (Run in Task Scheduler) shows that Windows finds the program. It only runs while
    ; the user is signed in (an interactive token), which a CI runner's test user isn't: then it is only reported
    static RunOnDemand(folderPath, marker) {
        task := TaskScheduler.Service().GetFolder(folderPath).GetTask("Marker_ spaces")
        task.Enabled := true
        try task.Run("")
        catch as e {
            DelayedStartCheck.Log("INFO run on demand: " e.Message)
            return
        }
        started := A_TickCount
        while (!FileExist(marker) && A_TickCount - started < 15000)
            Sleep(250)
        DelayedStartCheck.Log("INFO run on demand (path with spaces): " (FileExist(marker) ? "started" : "not started in 15 s (user not signed in?)") ", last result " Format("0x{:08X}", task.LastTaskResult & 0xFFFFFFFF))
        task.Enabled := false
    }

    static ByName(tasks) {
        found := Map()
        for t in tasks
            found[t.Name] := t
        return found
    }

    static Names(tasks) {
        text := ""
        for name in tasks
            text .= name "; "
        return text
    }

    static Join(values) {
        text := ""
        for v in values
            text .= (text = "" ? "" : ", ") v
        return text
    }

    static Check(label, ok, detail := "") {
        DelayedStartCheck.Log((ok ? "PASS " : "FAIL ") label (detail != "" ? " (" detail ")" : ""))
        if !ok
            DelayedStartCheck.Failures++
    }

    static Log(text) => FileAppend(text "`n", "*", "UTF-8")
}
