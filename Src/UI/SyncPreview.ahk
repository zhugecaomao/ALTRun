;===============================================================================
; SyncPreview.ahk - the sync window: follows a task while robocopy works, and shows what a preview found (AutoHotkey v2)
;-------------------------------------------------------------------------------
; One window per task. It opens as soon as a task is previewed or run and goes through these phases:
;   Check    a preview (robocopy /L) is running: the changes it finds so far appear in the list; the bar moves, Stop ends it
;   Preview  what a run would change: a summary on top (files and size to copy, files and folders to delete), the changes below,
;            deletions first. Run Now (Run and Delete for a mirror that would delete something: Confirm, Close is the default)
;   Run      robocopy copies: the files appear as they are done; the bar shows the bytes done against what the preview found
;            (moving without a preview, e.g. Update and Copy modes, which aren't previewed so the folders aren't scanned twice)
;   Done     how it ended: copied, deleted, failed, or stopped
; The list shows the first SyncProgress.Limit changes (the log has all of them), with their size. Closing the window doesn't stop
; robocopy; the task's result in the search window or Preferences opens it again. A window opened from the background (a mirror
; asking to confirm when its preview ends) may not get to the front, as Windows doesn't let a background program take the focus:
; it flashes in the taskbar and a notification says it is waiting. A mirror previewed more than StaleMs ago is previewed (and
; confirmed) again when run, as the destination may have changed since.
;
; Usage:
;   SyncPreview.Open(task, "Check" | "Run", {Run, Confirm, Totals})     follow the task's running robocopy
;   SyncPreview.Show(task, {Summary, Changes, LogFile}, {Run, Confirm}) what a preview found
;   SyncPreview.Finish(job, code | "Stopped") -> true when a window showed it        SyncPreview.ShowJob(name)  reopen a running task's window
;   SyncPreview.Rows(changes, source, destination)   -> [[action, item, size, robocopy's label]...]: deletions first
;===============================================================================

class SyncPreview {
    static Windows := Map()                                                 ; SyncJobs.Key(name) -> the task's window (see _Create)
    static StaleMs := 5 * 60 * 1000, TickMs := 500
    static _tickTimer := () => SyncPreview._Tick()

    ;---------------------------------------------------------------------------
    ; Phases
    ;---------------------------------------------------------------------------
    ; Follows the task's running robocopy: a preview (Check) or a run (Run). options.Totals: {Files, Bytes} a preview found
    static Open(task, phase, options := "") {
        options := IsObject(options) ? options : {}
        w := SyncPreview._Window(task, phase = "Check" ? I18n.T("Sync.Preview.Title", task["Name"]) : I18n.T("Sync.Window.Title", task["Name"]))
        w.Phase := phase, w.Options := options
        w.Totals := (options.HasOwnProp("Totals") && IsObject(options.Totals)) ? options.Totals : ""
        w.Shown := 0
        w.List.Delete()
        w.Warn.Value := "", w.Confirm.Value := ""
        w.Note.Value := (phase = "Check") ? I18n.T("Sync.Preview.Note") : ""
        SyncPreview._Buttons(w, ["Log", "Stop", "Close"])
        SyncPreview._Bar(w, "", phase)
        SyncPreview._Update(w)
        SyncPreview._Present(w)
        SetTimer(SyncPreview._tickTimer, SyncPreview.TickMs)
        return w.Gui
    }

    ; What a preview found. options.Run: offers to run it; options.Confirm: a mirror that would delete something asks first
    static Show(task, result, options) {
        name := task["Name"], summary := result.Summary, changes := result.Changes
        w := SyncPreview._Window(task, I18n.T(options.Confirm ? "Sync.Preview.ConfirmTitle" : "Sync.Preview.Title", name))
        w.Phase := "Preview", w.Options := options, w.Created := A_TickCount, w.Totals := SyncJobs.Totals(result)
        deletes := IsObject(summary) ? summary.Files.Extras + summary.Dirs.Extras : 0
        if !IsObject(summary)
            text := I18n.T("Sync.Preview.NoSummary")
        else if (summary.Files.Copied = 0 && deletes = 0 && summary.Dirs.Copied = 0)
            text := I18n.T("Sync.Preview.UpToDate")
        else
            text := I18n.T("Sync.Preview.Copy", summary.Files.Copied, summary.Files.Total, summary.Dirs.Copied
                , SyncProgress.FormatSize(IsObject(summary.Bytes) ? summary.Bytes.Copied : 0))
        w.Status.Value := text
        w.Warn.Value := deletes ? I18n.T("Sync.Preview.Delete", summary.Files.Extras, summary.Dirs.Extras) : ""
        w.Confirm.Value := options.Confirm ? I18n.T("Sync.Preview.ConfirmText") : ""
        SyncPreview._Fill(w, SyncPreview.Rows(changes, w.Source, w.Destination))
        more := (changes.HasOwnProp("Total") && changes.Total > changes.Length) ? I18n.T("Sync.Preview.More", changes.Length, changes.Total) : ""
        w.Note.Value := (more != "") ? more : I18n.T("Sync.Preview.Note")
        w.Run.Text := I18n.T(options.Confirm ? "Sync.Preview.RunDelete" : "Sync.Preview.Run")
        w.Run.Enabled := !SyncJobs.IsRunning(name)
        SyncPreview._Buttons(w, (options.HasOwnProp("Run") && options.Run) ? ["Log", "Run", "Close"] : ["Log", "Close"])
        SyncPreview._Bar(w, "", "Preview")
        SyncPreview._Present(w)
        if (options.Confirm && !WinActive("ahk_id " w.Gui.Hwnd)) {          ; shown from the background: tell where the question is
            DllCall("user32\FlashWindow", "Ptr", w.Gui.Hwnd, "Int", 1)
            App.Toast(I18n.T("Sync.Preview.Waiting", name))
        }
        return w.Gui
    }

    ; robocopy ended (code, or "Stopped"): the window's last word. false when the task has no window following it
    static Finish(job, code) {
        key := SyncJobs.Key(job.Name)
        if !SyncPreview.Windows.Has(key)
            return false
        w := SyncPreview.Windows[key]
        if (w.Phase != "Check" && w.Phase != "Run")
            return false
        if IsObject(job.Progress)
            SyncPreview._Add(w, job.Progress.Read())
        if (code = "Stopped")
            w.Status.Value := I18n.T("Sync.Window.Stopped", SyncPreview.Elapsed(A_TickCount - job.StartTick))
        else if !RoboSync.Result(code).Ok
            w.Status.Value := I18n.T("Sync.Failed", job.Name, code)
        else if (w.Phase = "Run") {
            summary := ""
            try summary := SyncJobs.Summary(SyncJobs.ReadLog(job.LogFile))
            progress := job.Progress
            copied := IsObject(summary) ? summary.Files.Copied : progress.CopyFiles
            bytes := (IsObject(summary) && IsObject(summary.Bytes)) ? summary.Bytes.Copied : progress.CopyBytes
            deleted := IsObject(summary) ? summary.Files.Extras : progress.DeleteFiles
            w.Status.Value := I18n.T("Sync.Window.Done", copied, SyncProgress.FormatSize(bytes), deleted, SyncPreview.Elapsed(A_TickCount - job.StartTick))
            failed := IsObject(summary) ? summary.Files.Failed : progress.Errors
            w.Warn.Value := failed ? I18n.T("Sync.Window.FailedFiles", failed) : ""
        }
        w.Phase := "Done"
        SyncPreview._Bar(w, RoboSync.Result(code = "Stopped" ? 16 : code).Ok ? 1000 : "", "Done")
        SyncPreview._Buttons(w, ["Log", "Close"])
        return true
    }

    ; The window of a running task (from its result in the search window or Preferences); a new one when it was closed
    static ShowJob(name) {
        key := SyncJobs.Key(name)
        if SyncPreview.Windows.Has(key)
            return SyncPreview._Present(SyncPreview.Windows[key])
        for task in RoboSync.Items
            if (SyncJobs.Key(task["Name"]) = key && SyncJobs.Jobs.Has(key))
                return SyncPreview.Open(task, SyncJobs.Jobs[key].Preview ? "Check" : "Run")
    }

    static IsActive(name) => SyncPreview.Windows.Has(key := SyncJobs.Key(name)) && WinActive("ahk_id " SyncPreview.Windows[key].Gui.Hwnd)

    static Close(key) {
        if !SyncPreview.Windows.Has(key)
            return
        try SyncPreview.Windows[key].Gui.Destroy()
        SyncPreview.Windows.Delete(key)
    }

    ; Run now: what a recent preview showed is what runs, so a mirror isn't previewed again; an older mirror preview is done again
    ; (SyncJobs.Run: checks, preview, and the question when it would delete anything)
    static _Run(task, created, now := A_TickCount) {
        if SyncPreview.IsStale(task, created, now)
            return SyncJobs.Run(task)
        key := SyncJobs.Key(task["Name"])
        SyncJobs.RunNow(task, SyncPreview.Windows.Has(key) ? SyncPreview.Windows[key].Totals : "")
    }

    ; A mirror's preview older than StaleMs (other modes don't delete: their preview can't be out of date in a harmful way)
    static IsStale(task, created, now := A_TickCount) => RoboSync._Get(task, "Mode") = "Mirror" && now - created > SyncPreview.StaleMs

    ;---------------------------------------------------------------------------
    ; While robocopy works
    ;---------------------------------------------------------------------------
    ; Every TickMs while a window follows a running task: the new changes from the log, the counts and the bar
    static _Tick() {
        following := false
        for key, w in SyncPreview.Windows.Clone() {
            if (w.Phase != "Check" && w.Phase != "Run")
                continue
            following := true
            SyncPreview._Update(w)
        }
        if !following
            SetTimer(SyncPreview._tickTimer, 0)
    }

    static _Update(w) {
        if !SyncJobs.Jobs.Has(w.Key)
            return
        job := SyncJobs.Jobs[w.Key], progress := job.Progress
        if IsObject(progress)
            SyncPreview._Add(w, progress.Read())
        else
            progress := SyncProgress()
        elapsed := SyncPreview.Elapsed(A_TickCount - job.StartTick)
        if (w.Phase = "Check") {
            w.Status.Value := I18n.T("Sync.Window.Checking", progress.Total, elapsed)
            return
        }
        size := SyncProgress.FormatSize(progress.CopyBytes)
        totals := w.Totals
        if job.Monitor
            w.Status.Value := I18n.T("Sync.Window.Monitoring", progress.CopyFiles, size, elapsed)
        else if (IsObject(totals) && totals.Files > 0)
            w.Status.Value := I18n.T("Sync.Window.Running", progress.CopyFiles, totals.Files, size, SyncProgress.FormatSize(totals.Bytes), elapsed)
        else
            w.Status.Value := I18n.T("Sync.Window.RunningCount", progress.CopyFiles, size, elapsed)
        w.Warn.Value := progress.Errors ? I18n.T("Sync.Window.FailedFiles", progress.Errors) : ""
        if (!job.Monitor && IsObject(totals) && (totals.Bytes > 0 || totals.Files > 0)) {
            done := totals.Bytes > 0 ? progress.CopyBytes / totals.Bytes : progress.CopyFiles / totals.Files
            SyncPreview._Bar(w, Min(990, Round(done * 1000)), "Run")      ; full only when robocopy has ended
        }
    }

    ; New changes in the list (in the order robocopy did them)
    static _Add(w, changes) {
        if !changes.Length
            return
        w.List.Opt("-Redraw")
        for change in changes
            w.List.Add("", SyncPreview._Row(change, w.Source, w.Destination)*)
        w.List.Opt("+Redraw")
        w.Shown += changes.Length
    }

    ; "0:42", "12:05", "1:02:03"
    static Elapsed(ms) {
        seconds := ms // 1000
        return (seconds >= 3600) ? Format("{}:{:02}:{:02}", seconds // 3600, Mod(seconds, 3600) // 60, Mod(seconds, 60)) : Format("{}:{:02}", seconds // 60, Mod(seconds, 60))
    }

    ;---------------------------------------------------------------------------
    ; The window
    ;---------------------------------------------------------------------------
    ; The task's window, made when it has none: one window goes through all the phases, so it stays where the user put it
    static _Window(task, title) {
        key := SyncJobs.Key(task["Name"])
        if SyncPreview.Windows.Has(key) {
            w := SyncPreview.Windows[key]
            w.Task := task
            w.Gui.Title := title
            return w
        }
        w := SyncPreview._Create(task, title)
        SyncPreview.Windows[key] := w
        return w
    }

    static _Create(task, title) {
        key := SyncJobs.Key(task["Name"])
        g := Gui("+Resize +MinSize480x300 -MinimizeBox", title)
        g.SetFont("s9", ThemeManager.FontName())
        g.MarginX := 14, g.MarginY := 12
        width := 680
        w := {Key: key, Task: task, Gui: g, Phase: "", Options: {}, Created: A_TickCount, Totals: "", Shown: 0}
        w.Source := RoboSync.Expand(RoboSync._Get(task, "Source")), w.Destination := RoboSync.Expand(RoboSync._Get(task, "Destination"))
        g.AddText("w" width " cGray", w.Source " → " w.Destination)
        g.SetFont("s10 bold")
        w.Status := g.AddText("w" width " y+8")
        g.SetFont("s9 norm")
        w.Warn := g.AddText("w" width " y+4 cC42B1C")
        w.Confirm := g.AddText("w" width " y+4 r2")
        w.List := g.AddListView("w" width " r16 y+6 -Multi Grid", [I18n.T("Sync.Preview.Col.Action"), I18n.T("Sync.Preview.Col.Item")
            , I18n.T("Sync.Preview.Col.Size"), I18n.T("Sync.Preview.Col.Robocopy")])
        w.List.ModifyCol(1, 110), w.List.ModifyCol(2, width - 110 - 80 - 110 - 24), w.List.ModifyCol(3, "80 Right"), w.List.ModifyCol(4, 110)
        w.Bar := g.AddProgress("w" width " h10 y+6 Range0-1000")
        g.SetFont("s8")
        w.Note := g.AddText("w" width " y+4 cGray")
        g.SetFont("s9")
        w.Log := g.AddButton("xm y+10 w120", I18n.T("Sync.Action.OpenLog"))
        w.Log.OnEvent("Click", (*) => Path.OpenText(SyncPreview._LogFile(w)))
        w.Run := g.AddButton("x" (14 + width - 90 - 10 - 150) " yp w150", I18n.T("Sync.Preview.Run"))
        w.Run.OnEvent("Click", (*) => SyncPreview._Run(w.Task, w.Created))
        w.Stop := g.AddButton("xp yp wp", I18n.T("Sync.Action.Stop"))
        w.Stop.OnEvent("Click", (*) => SyncJobs.Stop(w.Task["Name"]))
        w.CloseButton := g.AddButton("x" (14 + width - 90) " yp w90 Default", I18n.T("Sync.Preview.Close"))
        w.CloseButton.OnEvent("Click", (*) => SyncPreview.Close(key))
        g.OnEvent("Escape", (*) => SyncPreview.Close(key))
        g.OnEvent("Close", (*) => SyncPreview.Close(key))
        w.Gaps := SyncPreview._Gaps(w)
        g.OnEvent("Size", (gui, minMax, cw, ch) => SyncPreview._Layout(w, minMax, cw, ch))
        return w
    }

    ; Shows the window (once laid out), Close focused: Enter closes, running (and deleting) needs a click
    static _Present(w) {
        w.Gui.Show()
        w.CloseButton.Focus()
        return w.Gui
    }

    ; The buttons a phase has: Log, Run, Stop, Close (Run and Stop share a place)
    static _Buttons(w, names) {
        for name in ["Log", "Run", "Stop"]
            w.%name%.Visible := false
        for name in names
            if (name != "Close")
                w.%name%.Visible := true
    }

    ; The bar: a value 0-1000, or "" to move back and forth (marquee) while running without a known total / hidden otherwise
    static _Bar(w, value, phase) {
        marquee := (value = "" && (phase = "Check" || phase = "Run"))
        w.Bar.Visible := (value != "" || marquee)
        if marquee {
            w.Bar.Opt("+0x8")                                               ; PBS_MARQUEE
            SendMessage(0x40A, 1, 30, w.Bar)                                ; PBM_SETMARQUEE
        } else {
            SendMessage(0x40A, 0, 0, w.Bar)
            w.Bar.Opt("-0x8")
            if (value != "")
                w.Bar.Value := value
        }
    }

    ; The log of the running or last run (a preview logs to a file of its own)
    static _LogFile(w) {
        if SyncJobs.Jobs.Has(w.Key)
            return SyncJobs.Jobs[w.Key].LogFile
        task := w.Task
        if (RoboSync._Get(task, "LogFile") != "")
            return RoboSync.Expand(task["LogFile"])
        name := SyncJobs.FileName(task["Name"])
        preview := SyncJobs.LogDir "\" name "_preview.log", today := SyncJobs.LogDir "\" name "_" FormatTime(, "yyyyMMdd") ".log"
        return (w.Phase = "Preview" || !FileExist(today)) ? preview : today
    }

    static _Fill(w, rows) {
        w.List.Opt("-Redraw")
        w.List.Delete()
        for row in rows
            w.List.Add("", row*)
        w.List.Opt("+Redraw")
        w.Shown := rows.Length
    }

    ; The spaces between the controls as first laid out, in the units GetPos / Move and the Size event use (no DPI arithmetic here)
    static _Gaps(w) {
        w.List.GetPos(&listX, &listY, , &listH)
        w.Bar.GetPos(, &barY, , &barH)
        w.Note.GetPos(, &noteY, , &noteH)
        w.Log.GetPos(, &buttonY)
        w.Run.GetPos(&runX, , &runW)
        w.CloseButton.GetPos(&closeX)
        return {Margin: listX, ListBar: barY - listY - listH, BarNote: noteY - barY - barH, NoteButtons: buttonY - noteY - noteH, Between: closeX - runX - runW}
    }

    ; Resizing: the list takes the room, the bar, the note and the buttons stay below it, Run / Stop and Close at the right edge
    static _Layout(w, minMax, cw, ch) {
        if (minMax = -1)
            return
        gaps := w.Gaps
        w.List.GetPos(, &listY)
        w.Log.GetPos(, , , &buttonH)
        w.Note.GetPos(, , , &noteH)
        w.Bar.GetPos(, , , &barH)
        buttonY := ch - gaps.Margin - buttonH
        noteY := buttonY - gaps.NoteButtons - noteH
        barY := noteY - gaps.BarNote - barH
        width := cw - 2 * gaps.Margin
        w.List.Move(, , width, Max(40, barY - gaps.ListBar - listY))
        w.Bar.Move(, barY, width)
        w.Note.Move(, noteY, width)
        for ctrl in [w.Status, w.Warn, w.Confirm]
            ctrl.Move(, , width)
        w.Log.Move(, buttonY)
        w.CloseButton.GetPos(, , &closeW)
        closeX := cw - gaps.Margin - closeW
        w.CloseButton.Move(closeX, buttonY)
        w.Run.GetPos(, , &runW)
        w.Run.Move(closeX - gaps.Between - runW, buttonY)
        w.Stop.Move(closeX - gaps.Between - runW, buttonY)
    }

    ;---------------------------------------------------------------------------
    ; The list's rows
    ;---------------------------------------------------------------------------
    ; Deletions first (what matters most before a mirror), each with the item relative to its folder
    static Rows(changes, source, destination) {
        rows := [], copies := []
        for change in changes
            (change.Kind = "Delete" ? rows : copies).Push(SyncPreview._Row(change, source, destination))
        for row in copies
            rows.Push(row)
        return rows
    }

    static _Row(change, source, destination) {
        item := change.Path
        for root in [RTrim(source, "\") "\", RTrim(destination, "\") "\"]
            if (StrLen(root) > 1 && SubStr(item, 1, StrLen(root)) = root) {
                item := SubStr(item, StrLen(root) + 1)
                break
            }
        if (item = "")
            item := "."
        action := I18n.T("Sync.Preview." change.Kind (change.Folder ? "Folder" : "File"))
        size := (!change.Folder && change.HasOwnProp("Size") && change.Size != "") ? SyncProgress.FormatSize(change.Size) : ""
        return [action, item, size, change.Label]
    }
}
