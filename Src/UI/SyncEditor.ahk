;===============================================================================
; SyncEditor.ahk - the edit window of a sync task (Preferences → Sync) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; One window, top to bottom in the order people decide things: name, the two folders (with a button to swap them), how to
; sync (a gray line under the choice says what it does; mirror's in red, as it deletes), when (only when run, or kept in
; sync), what to leave out, then the advanced robocopy options. The robocopy command line the task makes is shown below,
; updated while typing, so people who know robocopy can check it (and copy it).
; OK checks the task (RoboSync.Check): errors must be fixed, warnings (another task's destination overlaps) can be accepted.
; Preview runs robocopy /L on what is in the window, before saving.
;
; Usage:
;   edited := SyncEditor.Edit(ownerGui, title, task, original, others)   -> the edited Map (keys not shown are kept), "" when cancelled
;       original: the task being edited ("" = a new one); others: all tasks (for the name and overlap checks; original is skipped)
;   SyncEditor.Read(task, values)    -> the task with the window's values (Map of field -> text / number), as OK saves it
;===============================================================================

class SyncEditor {
    static LabelW := 130, InputW := 500

    static Edit(owner, title, task, original := "", others := "") {
        state := {Result: ""}
        task := ItemEditor.WithDefaults(task, RoboSync.NewTask())
        g := Gui("+Owner" owner.Hwnd " -MinimizeBox", title)
        g.SetFont("s9", ThemeManager.FontName())
        g.MarginX := 14, g.MarginY := 12
        labelW := SyncEditor.LabelW, inputW := SyncEditor.InputW, inputX := "xm+" (labelW + 8)
        c := Map()

        label(key, gap := 12) => g.AddText("xm y+" gap " w" labelW " Section", I18n.T(key))
        hint(text, color := "cGray", below := "") {                         ; below: the row's tallest control, when it isn't the last one added
            y := "y+3"
            if IsObject(below) {
                below.GetPos(, &by, , &bh)
                y := "y" (by + bh + 3)
            }
            g.SetFont("s8")
            ctrl := g.AddText(inputX " " y " w" inputW " " color, text)
            g.SetFont("s9")
            return ctrl
        }
        section(key) {
            g.SetFont("bold")
            g.AddText("xm y+18 w" (labelW + 8 + inputW), I18n.T(key))
            g.SetFont("norm")
            g.AddText("xm y+3 w" (labelW + 8 + inputW) " h1 0x10")          ; SS_ETCHEDHORZ
        }
        folderRow(key, labelKey, extraW := 0) {
            label(labelKey)
            c[key] := g.AddEdit("xs+" (labelW + 8) " ys-3 w" (inputW - 34 - extraW) " r1 -Multi", task[key])
            browse := g.AddButton("x+4 yp-1 w30", I18n.T("Prefs.Browse"))
            browse.OnEvent("Click", ItemEditor._Browser(c[key], "folder", g))
        }

        label("Prefs.Col.Name", 0)
        c["Name"] := g.AddEdit("xs+" (labelW + 8) " ys-3 w240 r1 -Multi", task["Name"])
        hint(I18n.T("Sync.Field.Name.Hint", AppSettings.Feature("Sync")["Keyword"]))
        folderRow("Source", "Sync.Col.Source")
        hint(I18n.T("Sync.Field.Source.Hint"))
        folderRow("Destination", "Sync.Col.Destination", 34)
        swap := g.AddButton("x+4 yp w30", "⇅")
        Win.AddTooltip(swap.Hwnd, I18n.T("Sync.Field.Swap"))
        swap.OnEvent("Click", (*) => (value := c["Source"].Value, c["Source"].Value := c["Destination"].Value, c["Destination"].Value := value, Update()))
        label("Sync.Field.Files")
        c["Files"] := g.AddEdit("xs+" (labelW + 8) " ys-3 w240 r1 -Multi", task["Files"])
        hint(I18n.T("Sync.Field.Files.Hint"))

        label("Sync.Col.Mode")
        modeLabels := []
        for mode in RoboSync.Modes
            modeLabels.Push(I18n.T("Sync.Mode." mode))
        c["Mode"] := g.AddDropDownList("xs+" (labelW + 8) " ys-3 w240", modeLabels)
        c["Mode"].Value := Max(1, RoboSync._Index(RoboSync.Modes, task["Mode"]))
        modeHint := hint(I18n.T("Sync.ModeHint.Mirror") "`n ")             ; two lines tall: the longest hint fits

        label("Sync.Field.When")
        c["Monitor"] := g.AddDropDownList("xs+" (labelW + 8) " ys-3 w240", [I18n.T("Sync.When.Manual"), I18n.T("Sync.When.Monitor")])
        c["Monitor"].Value := task["Monitor"] ? 2 : 1
        g.AddText("x+16 ys", I18n.T("Sync.Field.MonitorChanges"))
        c["MonitorChanges"] := g.AddEdit("x+6 ys-3 w44 r1 -Multi Number", task["MonitorChanges"])
        g.AddText("x+12 ys", I18n.T("Sync.Field.MonitorMinutes"))
        c["MonitorMinutes"] := g.AddEdit("x+6 ys-3 w44 r1 -Multi Number", task["MonitorMinutes"])
        whenHint := hint(I18n.T("Sync.Field.When.Hint") "`n ")              ; two lines tall: the mirror's note fits

        label("Sync.Field.ExcludeFiles")
        c["ExcludeFiles"] := g.AddEdit("xs+" (labelW + 8) " ys-3 w" inputW " r1 -Multi", RoboSync.JoinPatterns(task["ExcludeFiles"]))
        hint(I18n.T("Sync.Field.ExcludeFiles.Hint"))
        label("Sync.Field.ExcludeDirs")
        c["ExcludeDirs"] := g.AddEdit("xs+" (labelW + 8) " ys-3 w" inputW " r1 -Multi", RoboSync.JoinPatterns(task["ExcludeDirs"]))
        hint(I18n.T("Sync.Field.ExcludeDirs.Hint"))

        section("Sync.Section.Advanced")
        label("Sync.Field.Retries", 10)
        c["Retries"] := g.AddEdit("xs+" (labelW + 8) " ys-3 w70 r1 -Multi Number", task["Retries"])
        waitLabel := g.AddText("x+16 ys", I18n.T("Sync.Field.Wait"))
        c["Wait"] := g.AddEdit("x+6 ys-3 w70 r1 -Multi Number", task["Wait"])
        threadsLabel := g.AddText("x+16 ys", I18n.T("Sync.Field.Threads"))
        c["Threads"] := g.AddEdit("x+6 ys-3 w70 r1 -Multi Number", task["Threads"])
        hint(I18n.T("Sync.Field.Retries.Hint"))
        label("Sync.Field.Copy")
        c["Copy"] := g.AddEdit("xs+" (labelW + 8) " ys-3 w70 r1 -Multi Uppercase", task["Copy"])
        dirLabel := g.AddText("x+16 ys", I18n.T("Sync.Field.DirCopy"))
        c["DirCopy"] := g.AddEdit("x+6 ys-3 w70 r1 -Multi Uppercase", task["DirCopy"])
        SyncEditor.AlignColumns(waitLabel, c["Wait"], dirLabel, c["DirCopy"], threadsLabel, c["Threads"])
        hint(I18n.T("Sync.Field.Copy.Hint"))
        label("Sync.Field.LogFile")
        c["LogFile"] := g.AddEdit("xs+" (labelW + 8) " ys-3 w" (inputW - 130) " r1 -Multi", task["LogFile"])
        c["AppendLog"] := g.AddCheckbox("x+12 ys w118", I18n.T("Sync.Field.AppendLog"))
        c["AppendLog"].Value := task["AppendLog"] ? 1 : 0
        hint(I18n.T("Sync.Field.LogFile.Hint", SyncJobs.LogDir), , c["LogFile"])
        label("Sync.Field.Extra")
        c["Extra"] := g.AddEdit("xs+" (labelW + 8) " ys-3 w" inputW " r1 -Multi", task["Extra"])
        hint(I18n.T("Sync.Field.Extra.Hint"))

        label("Sync.Field.Command")
        command := g.AddEdit("xs+" (labelW + 8) " ys-3 w" inputW " r4 +Multi ReadOnly -VScroll")   ; four lines: the log path at the end fits

        copyButton := g.AddButton("xm y+18 w120", I18n.T("Sync.Field.CopyCommand"))
        copyButton.OnEvent("Click", (*) => (A_Clipboard := command.Value, App.Notify(I18n.T("Sync.Field.CommandCopied"))))
        right := 14 + labelW + 8 + inputW
        g.AddButton("x" (right - 80 - 10 - 80 - 10 - 110) " yp w110", I18n.T("Sync.Button.Preview")).OnEvent("Click", OnPreview)
        g.AddButton("x+10 yp w80 Default", "OK").OnEvent("Click", OnOk)
        g.AddButton("x+10 yp w80", I18n.T("Prefs.Cancel")).OnEvent("Click", (*) => g.Destroy())

        for key, ctrl in c
            ctrl.OnEvent(ctrl.Type = "CheckBox" ? "Click" : "Change", (*) => Update())
        g.OnEvent("Escape", (*) => g.Destroy())
        g.OnEvent("Close", (*) => g.Destroy())
        Update()

        owner.Opt("+Disabled")
        g.Show()
        WinWaitClose(g.Hwnd)
        owner.Opt("-Disabled")
        try WinActivate("ahk_id " owner.Hwnd)
        return state.Result

        Values() {
            result := Map()
            for key, ctrl in c
                result[key] := ctrl.Value                                   ; a drop-down list: the number of the choice
            return result
        }
        Update() {                                                          ; the mode's description and the command line follow what is typed
            mode := RoboSync.Modes[c["Mode"].Value]
            modeHint.Opt(mode = "Mirror" ? "cC42B1C" : "cGray")
            modeHint.Value := I18n.T("Sync.ModeHint." mode)
            if (mode = "Mirror")                                            ; a mirror only runs when started (RoboSync.Check: MirrorMonitor)
                c["Monitor"].Value := 1
            c["Monitor"].Enabled := (mode != "Mirror")
            whenHint.Value := I18n.T(mode = "Mirror" ? "Sync.Field.When.Mirror" : "Sync.Field.When.Hint")
            monitor := c["Monitor"].Value = 2
            c["MonitorChanges"].Enabled := monitor, c["MonitorMinutes"].Enabled := monitor
            edited := SyncEditor.Read(task, Values())
            if (edited["Name"] = "")                                        ; the default log is named after the task
                edited["Name"] := RoboSync.SuggestName(edited["Destination"] != "" ? edited["Destination"] : edited["Source"])
            built := RoboSync.Build(SyncJobs.AsRun(edited))                 ; with the log it runs with (the default one when none is set)
            command.Value := "robocopy " built.Args
        }
        OnPreview(*) {
            g.Opt("+OwnDialogs")
            edited := SyncEditor.Read(task, Values())
            if (edited["Name"] = "")
                edited["Name"] := RoboSync.SuggestName(edited["Destination"] != "" ? edited["Destination"] : edited["Source"])
            SyncJobs.Preview(edited, {Run: false})                          ; not saved yet: the preview window doesn't offer to run it
        }
        OnOk(*) {
            g.Opt("+OwnDialogs")
            edited := SyncEditor.Read(task, Values())
            if (edited["Name"] = "") {
                MsgBox(I18n.T("Prefs.Required", I18n.T("Prefs.Col.Name")), title, 48)
                return c["Name"].Focus()
            }
            rest := []                                                      ; the other tasks: the one being edited is left out
            for other in (others is Array ? others : [])
                if (other != original)
                    rest.Push(other)
            for other in rest
                if (SyncJobs.Key(RoboSync._Get(other, "Name")) = SyncJobs.Key(edited["Name"])) {
                    MsgBox(I18n.T("Sync.Problem.SameName", edited["Name"]), title, 48)
                    return c["Name"].Focus()
                }
            warnings := ""
            for problem in RoboSync.Check(edited, rest) {
                if (problem.Level = "Error") {
                    MsgBox(I18n.T("Sync.Problem." problem.Code, edited["Name"]), title, 48)
                    return c[SyncEditor.FieldOf(problem.Code)].Focus()
                }
                warnings .= (warnings = "" ? "" : "`n") I18n.T("Sync.Warn." problem.Code, problem.Detail)
            }
            if (warnings != "" && MsgBox(warnings "`n`n" I18n.T("Editor.SaveAnyway"), title, "YesNo Icon! Default2") != "Yes")
                return
            state.Result := edited
            g.Destroy()
        }
    }

    ; The field to go back to for a problem code
    static FieldOf(code) {
        switch code {
            case "NoSource", "SamePath", "Nested", "SourceNotAbsolute": return "Source"
            case "NoDestination", "MirrorRoot", "DestinationNotAbsolute": return "Destination"
            case "MirrorMonitor": return "Mode"
            case "Threads": return "Threads"
        }
        return "Name"
    }

    ; Retries and File info: the second boxes in one column, after the longer of the two labels (Wait, Folder info); Threads follows
    static AlignColumns(waitLabel, wait, dirLabel, dirCopy, threadsLabel, threads) {
        waitLabel.GetPos(, , &waitW), dirLabel.GetPos(&labelX, , &dirW)
        editX := labelX + Max(waitW, dirW) + 6
        wait.Move(editX), dirCopy.Move(editX)
        wait.GetPos(, , &editW)
        threadsLabel.Move(editX + editW + 16)
        threadsLabel.GetPos(&threadsX, , &threadsW)
        threads.Move(threadsX + threadsW + 6)
    }

    ; The task with the window's values: numbers that aren't numbers become the defaults, exclude lists are split at ";"
    static Read(task, values) {
        edited := task.Clone(), defaults := RoboSync.NewTask()
        number(key) => IsInteger(Trim(values[key])) ? Integer(Trim(values[key])) : defaults[key]
        for key in ["Name", "Source", "Destination", "Files", "Copy", "DirCopy", "LogFile", "Extra"]
            edited[key] := Trim(values[key])
        if (edited["Files"] = "")
            edited["Files"] := "*.*"
        for key in ["Source", "Destination"]
            edited[key] := RoboSync.TrimSlash(edited[key])
        edited["Mode"] := RoboSync.Modes[Max(1, Min(RoboSync.Modes.Length, values["Mode"]))]
        edited["Monitor"] := (values["Monitor"] = 2 && edited["Mode"] != "Mirror") ? 1 : 0   ; a mirror doesn't keep in sync
        for key in ["MonitorChanges", "MonitorMinutes", "Retries", "Wait", "Threads"]
            edited[key] := number(key)
        edited["AppendLog"] := values["AppendLog"] ? 1 : 0
        edited["ExcludeFiles"] := RoboSync.SplitPatterns(values["ExcludeFiles"])
        edited["ExcludeDirs"] := RoboSync.SplitPatterns(values["ExcludeDirs"])
        return edited
    }
}
