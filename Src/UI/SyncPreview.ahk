;===============================================================================
; SyncPreview.ahk - the sync window: previews and runs a task, following robocopy while it works (AutoHotkey v2)
;-------------------------------------------------------------------------------
; One window per task, titled "<task> - Folder Sync". It goes through these phases, its controls staying where they are (the
; buttons are only enabled or disabled, see Enabled):
;   Ready    opened from Preferences (Sync...): the folders and how the last run went; Preview or Sync Now starts it
;   Check    a preview (robocopy /L) is running: the changes it finds so far appear in the list; the bar moves, Stop ends it
;   Preview  what a run would change: a summary on top (files and size to copy, files and folders to delete), the changes below,
;            deletions first. Sync Now runs it (Sync and Delete when it would delete something; Close is the default button)
;   Run      robocopy copies: the files appear as they are done; the bar shows the bytes done against what the preview found
;            (moving without a preview, e.g. Update and Copy modes, which aren't previewed so the folders aren't scanned twice)
;   Done     how it ended: copied, deleted, failed, or stopped
; The list shows the first SyncProgress.Limit changes (the log has all of them), with their size; double-click shows one in
; Explorer. The bar's progress is also on the window's taskbar button (TaskbarProgress); red when files failed, yellow when
; stopped. Closing the window doesn't stop robocopy; the task's result in the search window or Preferences opens it again.
; A window opened from the background (a mirror asking to confirm when its preview ends) may not get to the front, as Windows
; doesn't let a background program take the focus: it flashes in the taskbar and a notification says it is waiting. A mirror
; previewed more than StaleMs ago is previewed (and confirmed) again when run, as the destination may have changed since.
;
; Usage:
;   SyncPreview.OpenTask(task)                                          the task's window: its progress when running, else Ready
;   SyncPreview.Open(task, "Ready" | "Check" | "Run", {Run, Confirm, Totals})   Run: false = no Sync Now (the editor's preview)
;   SyncPreview.Show(task, {Summary, Changes, LogFile}, {Run, Confirm}) what a preview found
;   SyncPreview.Finish(job, code | "Stopped") -> true when a window showed it        SyncPreview.ShowJob(name)  reopen a running task's window
;   SyncPreview.Rows(changes, source, destination)   -> [[item, action, size, full path]...]: deletions first
;===============================================================================

class SyncPreview {
    static Windows := Map()                                                 ; SyncJobs.Key(name) -> the task's window (see _Create)
    static StaleMs := 5 * 60 * 1000, TickMs := 500
    static Width := 860, ActionWidth := 150, SizeWidth := 80
    static Enabled := Map("Ready", [1, 1, 0], "Check", [0, 0, 1], "Preview", [1, 1, 0], "Run", [0, 0, 1], "Done", [1, 1, 0])   ; Preview, Sync Now, Stop
    static _tickTimer := () => SyncPreview._Tick()

    ;---------------------------------------------------------------------------
    ; Phases
    ;---------------------------------------------------------------------------
    ; From Preferences: the window of a running task, or a new one that waits for Preview or Sync Now
    static OpenTask(task) {
        if SyncJobs.IsRunning(task["Name"])
            return SyncPreview.ShowJob(task["Name"])
        return SyncPreview.Open(task, "Ready")
    }

    ; Ready, or following the task's running robocopy: a preview (Check) or a run (Run). options.Totals: {Files, Bytes} a preview found
    static Open(task, phase, options := "") {
        options := IsObject(options) ? options : {}
        w := SyncPreview._Window(task, options)
        w.Phase := phase
        w.Totals := (options.HasOwnProp("Totals") && IsObject(options.Totals)) ? options.Totals : ""
        SyncPreview._Fill(w, [])
        w.Sync.Text := I18n.T("Sync.Preview.Run")
        if (phase = "Ready") {
            w.Status.Value := I18n.T("Sync.Window.Ready")
            SyncPreview._Detail(w, SyncJobs.State(task["Name"]))
            w.Note.Value := ""
        } else {
            SyncPreview._Detail(w, "")
            w.Note.Value := (phase = "Check") ? I18n.T("Sync.Preview.Note") : ""
        }
        SyncPreview._Present(w)
        SyncPreview._Buttons(w)
        SyncPreview._Bar(w, phase = "Ready" ? 0 : "")
        if (phase != "Ready") {
            SyncPreview._Update(w)
            SetTimer(SyncPreview._tickTimer, SyncPreview.TickMs)
        }
        return w.Gui
    }

    ; What a preview found. options.Run: offers to run it; options.Confirm: a mirror that would delete something waits for the user
    static Show(task, result, options) {
        name := task["Name"], summary := result.Summary, changes := result.Changes
        w := SyncPreview._Window(task, options)
        w.Phase := "Preview", w.Created := A_TickCount, w.Totals := SyncJobs.Totals(result)
        deletes := SyncPreview._Count(changes, "DeleteFiles") + SyncPreview._Count(changes, "DeleteDirs")
        w.Status.Value := SyncPreview.Summary(summary, changes)
        if deletes
            SyncPreview._Detail(w, I18n.T("Sync.Preview.Delete", changes.DeleteFiles, changes.DeleteDirs) " " I18n.T("Sync.Preview.ConfirmText"), true)
        else
            SyncPreview._Detail(w, (kept := SyncPreview.Kept(summary, changes)) ? I18n.T("Sync.Preview.Kept", kept) : "")
        SyncPreview._Fill(w, SyncPreview.Rows(changes, w.Source, w.Destination))
        more := (changes.HasOwnProp("Total") && changes.Total > changes.Length) ? I18n.T("Sync.Preview.More", changes.Length, changes.Total) : ""
        w.Note.Value := (more != "") ? more : I18n.T("Sync.Preview.Note")
        w.Sync.Text := I18n.T(deletes ? "Sync.Preview.RunDelete" : "Sync.Preview.Run")
        SyncPreview._Present(w)
        SyncPreview._Buttons(w)
        SyncPreview._Bar(w, 0)
        if (options.HasOwnProp("Confirm") && options.Confirm && !WinActive("ahk_id " w.Gui.Hwnd)) {   ; shown from the background: tell where the question is
            DllCall("user32\FlashWindow", "Ptr", w.Gui.Hwnd, "Int", 1)
            App.Toast(I18n.T("Sync.Preview.Waiting", name))
        }
        return w.Gui
    }

    ; The preview's summary line. Folders to create and what to delete come from the log's lines (SyncJobs.Changes), not the summary:
    ; with /MT (several threads, the default) robocopy's "Dirs : Copied" is every folder, even when nothing changes, and it writes no
    ; folder lines; "Extras" also counts excluded files and folders only in the destination, which it keeps (probed on Windows)
    static Summary(summary, changes) {
        if !IsObject(summary)
            return I18n.T("Sync.Preview.NoSummary")
        files := summary.Files.Copied, newDirs := SyncPreview._Count(changes, "NewDirs")
        size := SyncProgress.FormatSize(IsObject(summary.Bytes) ? summary.Bytes.Copied : 0)
        if (files = 0 && newDirs = 0)
            return (SyncPreview._Count(changes, "DeleteFiles") + SyncPreview._Count(changes, "DeleteDirs")) ? I18n.T("Sync.Preview.NothingToCopy") : I18n.T("Sync.Preview.UpToDate", summary.Files.Total, summary.Dirs.Total)
        if (files = 0)
            return I18n.T("Sync.Preview.FoldersOnly", newDirs)
        return newDirs ? I18n.T("Sync.Preview.CopyFolders", files, summary.Files.Total, size, newDirs) : I18n.T("Sync.Preview.Copy", files, summary.Files.Total, size)
    }

    ; Excluded files and folders only in the destination: in the summary's Extras, but robocopy neither lists nor deletes them
    static Kept(summary, changes) {
        if !IsObject(summary)
            return 0
        return Max(0, summary.Files.Extras - SyncPreview._Count(changes, "DeleteFiles")) + Max(0, summary.Dirs.Extras - SyncPreview._Count(changes, "DeleteDirs"))
    }

    static _Count(changes, name) => changes.HasOwnProp(name) ? changes.%name% : 0

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
        failed := 0
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
            deleted := progress.DeleteFiles                                 ; what the log lists (Extras counts kept excluded files too)
            w.Status.Value := I18n.T("Sync.Window.Done", copied, SyncProgress.FormatSize(bytes), deleted, SyncPreview.Elapsed(A_TickCount - job.StartTick))
            failed := IsObject(summary) ? summary.Files.Failed : progress.Errors
        }
        SyncPreview._Detail(w, failed ? I18n.T("Sync.Window.FailedFiles", failed) : "", true)
        w.Phase := "Done"
        if (code = "Stopped")
            SyncPreview._Bar(w, w.Bar.Value, "Paused")
        else
            SyncPreview._Bar(w, 1000, (!RoboSync.Result(code).Ok || failed) ? "Error" : "Normal")
        SyncPreview._Buttons(w)
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
        hwnd := SyncPreview.Windows[key].Gui.Hwnd
        TaskbarProgress.Set(hwnd, "None")
        try SyncPreview.Windows[key].Gui.Destroy()
        SyncPreview.Windows.Delete(key)
    }

    ; Sync Now: after a recent preview, what it showed is what runs, so a mirror isn't previewed again; otherwise (an older
    ; preview, or none: Ready, Done) SyncJobs.Run checks, previews a mirror and asks when it would delete anything
    static _Sync(w, now := A_TickCount) {
        if (w.Phase = "Preview" && !SyncPreview.IsStale(w.Task, w.Created, now))
            return SyncJobs.RunNow(w.Task, w.Totals)
        SyncJobs.Run(w.Task)
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
        SyncPreview._Detail(w, progress.Errors ? I18n.T("Sync.Window.FailedFiles", progress.Errors) : "", true)
        w.Note.Value := (progress.Total > w.Shown) ? I18n.T("Sync.Preview.More", w.Shown, progress.Total) : ""
        if (!job.Monitor && IsObject(totals) && (totals.Bytes > 0 || totals.Files > 0)) {
            done := totals.Bytes > 0 ? progress.CopyBytes / totals.Bytes : progress.CopyFiles / totals.Files
            SyncPreview._Bar(w, Min(990, Round(done * 1000)), progress.Errors ? "Error" : "Normal")   ; full only when robocopy has ended
        }
    }

    ; New changes in the list (in the order robocopy did them)
    static _Add(w, changes) {
        if !changes.Length
            return
        rows := []
        for change in changes
            rows.Push(SyncPreview._Row(change, w.Source, w.Destination))
        SyncPreview._Fill(w, rows, true)
    }

    ; "0:42", "12:05", "1:02:03"
    static Elapsed(ms) {
        seconds := ms // 1000
        return (seconds >= 3600) ? Format("{}:{:02}:{:02}", seconds // 3600, Mod(seconds, 3600) // 60, Mod(seconds, 60)) : Format("{}:{:02}", seconds // 60, Mod(seconds, 60))
    }

    ;---------------------------------------------------------------------------
    ; The window
    ;---------------------------------------------------------------------------
    ; The task's window, made when it has none: one window goes through all the phases, so it stays where the user put it.
    ; options.Run: false when the task is still being edited (the editor's preview), so it can't be run from here
    static _Window(task, options) {
        key := SyncJobs.Key(task["Name"])
        if !SyncPreview.Windows.Has(key)
            SyncPreview.Windows[key] := SyncPreview._Create(task)
        w := SyncPreview.Windows[key]
        w.Task := task, w.Options := options
        w.CanRun := !(options.HasOwnProp("Run") && !options.Run)
        w.Source := RoboSync.Expand(RoboSync._Get(task, "Source")), w.Destination := RoboSync.Expand(RoboSync._Get(task, "Destination"))
        w.Folders.Value := w.Source " → " w.Destination
        return w
    }

    static _Create(task) {
        key := SyncJobs.Key(task["Name"])
        g := Gui("+Resize +MinSize640x380", I18n.T("Sync.Window.Title", task["Name"]))
        g.SetFont("s9", ThemeManager.FontName())
        g.MarginX := 14, g.MarginY := 12
        width := SyncPreview.Width
        w := {Key: key, Task: task, Gui: g, Phase: "", Options: {}, CanRun: true, Created: A_TickCount, Totals: "", Shown: 0, Paths: []}
        w.Folders := g.AddText("w" width " cGray")
        g.SetFont("s10 bold")
        w.Status := g.AddText("w" width " y+8")
        g.SetFont("s9 norm")
        w.Detail := g.AddText("w" width " y+4 r2")                         ; the deletions, files that failed, the last run: two lines, always there
        w.List := g.AddListView("w" width " r16 y+6 -Multi Grid NoSortHdr +LV0x4000", [I18n.T("Sync.Preview.Col.Item"), I18n.T("Sync.Preview.Col.Action")
            , I18n.T("Sync.Preview.Col.Size")])                             ; LVS_EX_LABELTIP: a cut-off path shows whole on hover
        w.List.ModifyCol(1, width - SyncPreview.ActionWidth - SyncPreview.SizeWidth - 24), w.List.ModifyCol(2, SyncPreview.ActionWidth), w.List.ModifyCol(3, SyncPreview.SizeWidth " Right")
        w.List.OnEvent("DoubleClick", (list, row) => SyncPreview._ShowItem(w, row))
        w.Bar := g.AddProgress("w" width " h10 y+6 Range0-1000")
        g.SetFont("s8")
        w.Note := g.AddText("w" width " y+4 cGray")
        g.SetFont("s9")
        w.Log := g.AddButton("xm y+10 w120", I18n.T("Sync.Action.OpenLog"))
        w.Log.OnEvent("Click", (*) => Path.OpenText(SyncPreview._LogFile(w)))
        right := 14 + width
        w.CloseButton := g.AddButton("x" (right - 90) " yp w90 Default", I18n.T("Sync.Preview.Close"))
        w.Stop := g.AddButton("x" (right - 90 - 8 - 90) " yp w90", I18n.T("Sync.Action.Stop"))
        w.Sync := g.AddButton("x" (right - 90 - 8 - 90 - 8 - 150) " yp w150", I18n.T("Sync.Preview.Run"))
        w.Preview := g.AddButton("x" (right - 90 - 8 - 90 - 8 - 150 - 8 - 110) " yp w110", I18n.T("Sync.Button.Preview"))
        w.CloseButton.OnEvent("Click", (*) => SyncPreview.Close(key))
        w.Stop.OnEvent("Click", (*) => SyncJobs.Stop(w.Task["Name"]))
        w.Sync.OnEvent("Click", (*) => SyncPreview._Sync(w))
        w.Preview.OnEvent("Click", (*) => SyncJobs.Preview(w.Task, {Run: w.CanRun}))
        g.OnEvent("Escape", (*) => SyncPreview.Close(key))
        g.OnEvent("Close", (*) => SyncPreview.Close(key))
        w.Gaps := SyncPreview._Gaps(w)
        g.OnEvent("Size", (gui, minMax, cw, ch) => SyncPreview._Layout(w, minMax, cw, ch))
        return w
    }

    ; Shows the window (once laid out), Close focused: Enter closes, running (and deleting) needs a click
    static _Present(w) {
        w.Gui.Show()
        w.Gui.GetClientPos(, , &cw, &ch)                                    ; the same size as before: no Size event
        SyncPreview._Layout(w, 0, cw, ch)
        w.CloseButton.Focus()
        return w.Gui
    }

    ; The buttons the phase allows (Enabled); the focus moves off a button that was just disabled
    static _Buttons(w) {
        focused := w.Gui.FocusedCtrl
        allowed := SyncPreview.Enabled[w.Phase]
        w.Preview.Enabled := allowed[1]
        w.Sync.Enabled := allowed[2] && w.CanRun
        w.Stop.Enabled := allowed[3]
        if (!IsObject(focused) || !focused.Enabled)
            (w.Stop.Enabled ? w.Stop : w.CloseButton).Focus()
    }

    ; The line under the summary: the deletions or failed files in red (warn), otherwise gray
    static _Detail(w, text, warn := false) {
        color := (warn && text != "") ? "cC42B1C" : "cGray"
        if (w.HasOwnProp("DetailColor") && w.DetailColor = color && w.Detail.Value = text)
            return
        w.DetailColor := color
        w.Detail.SetFont(color)
        w.Detail.Value := text
    }

    ; The bar (and the taskbar button): a value 0-1000, or "" to move back and forth while robocopy works without a known total.
    ; state: Normal, Error (red: files failed), Paused (yellow: stopped)
    static _Bar(w, value, state := "Normal") {
        if (value = "") {
            w.Bar.Opt("+0x8"), w.BarState := "Moving"                      ; PBS_MARQUEE
            SendMessage(0x40A, 1, 30, w.Bar)                                ; PBM_SETMARQUEE
            TaskbarProgress.Set(w.Gui.Hwnd, "Indeterminate")
            return
        }
        SendMessage(0x40A, 0, 0, w.Bar)
        w.Bar.Opt("-0x8")
        w.BarState := state
        SendMessage(0x410, state = "Error" ? 2 : state = "Paused" ? 3 : 1, 0, w.Bar)   ; PBM_SETSTATE: PBST_NORMAL / ERROR / PAUSED
        w.Bar.Value := value
        if (w.Phase = "Run" || w.Phase = "Done" && state != "Normal")       ; a run's progress; a finished run only when it went wrong
            TaskbarProgress.Set(w.Gui.Hwnd, state, value)
        else
            TaskbarProgress.Set(w.Gui.Hwnd, "None")
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

    ; The list's rows (append: after the ones there), each row's full path kept for double-click
    static _Fill(w, rows, append := false) {
        w.List.Opt("-Redraw")
        if !append
            w.List.Delete(), w.Paths := []
        for row in rows
            w.List.Add("", row[1], row[2], row[3]), w.Paths.Push(row[4])
        SyncPreview._Columns(w)
        w.List.Opt("+Redraw")
        w.Shown := w.Paths.Length
    }

    ; Double-click: the file or folder in Explorer (the source's for a copy, the destination's for a delete); its folder when it
    ; isn't there (any more)
    static _ShowItem(w, row) {
        if (row < 1 || row > w.Paths.Length)
            return
        target := RTrim(w.Paths[row], "\")
        while (target != "" && !FileExist(target) && InStr(target, "\"))
            target := SubStr(target, 1, InStr(target, "\", , -1) - 1)
        if (target = "" || !FileExist(target))
            return
        try Run((target = RTrim(w.Paths[row], "\")) ? 'explorer.exe /select,"' target '"' : 'explorer.exe "' target '"')
    }

    ; The spaces between the controls as first laid out, in the units GetPos / Move and the Size event use (no DPI arithmetic here)
    static _Gaps(w) {
        w.List.GetPos(&listX, &listY, , &listH)
        w.Bar.GetPos(, &barY, , &barH)
        w.Note.GetPos(, &noteY, , &noteH)
        w.Log.GetPos(, &buttonY)
        w.Stop.GetPos(&stopX, , &stopW)
        w.CloseButton.GetPos(&closeX)
        return {Margin: listX, ListY: listY, ListBar: barY - listY - listH, BarNote: noteY - barY - barH, NoteButtons: buttonY - noteY - noteH, Between: closeX - stopX - stopW}
    }

    ; Resizing: the list takes the room and its first column the width; the bar, the note and the buttons stay below it,
    ; Preview, Sync Now, Stop and Close at the right edge
    static _Layout(w, minMax, cw, ch) {
        if (minMax = -1)
            return
        gaps := w.Gaps
        w.Log.GetPos(, , , &buttonH)
        w.Note.GetPos(, , , &noteH)
        w.Bar.GetPos(, , , &barH)
        buttonY := ch - gaps.Margin - buttonH
        noteY := buttonY - gaps.NoteButtons - noteH
        barY := noteY - gaps.BarNote - barH
        width := cw - 2 * gaps.Margin
        w.List.Move(, , width, Max(40, barY - gaps.ListBar - gaps.ListY))
        w.Bar.Move(, barY, width)
        w.Note.Move(, noteY, width)
        for ctrl in [w.Folders, w.Status, w.Detail]
            ctrl.Move(, , width)
        w.Log.Move(, buttonY)
        x := cw - gaps.Margin
        for button in [w.CloseButton, w.Stop, w.Sync, w.Preview] {
            button.GetPos(, , &bw)
            x -= bw
            button.Move(x, buttonY)
            x -= gaps.Between
        }
        SyncPreview._Columns(w)
    }

    ; The first column (the path) takes the list's width that the other columns leave. In the list's own pixels (LVM_*)
    static _Columns(w) {
        rect := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", w.List.Hwnd, "Ptr", rect)
        client := NumGet(rect, 8, "Int")
        if !(WinGetStyle(w.List.Hwnd) & 0x200000)                           ; no WS_VSCROLL yet: leave room for it
            client -= DllCall("GetSystemMetrics", "Int", 2)                 ; SM_CXVSCROLL
        others := SendMessage(0x101D, 1, 0, w.List) + SendMessage(0x101D, 2, 0, w.List)   ; LVM_GETCOLUMNWIDTH
        SendMessage(0x101E, 0, Max(80, client - others), w.List)            ; LVM_SETCOLUMNWIDTH
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

    ; [item, action, size, full path]: the item relative to the source or destination; the path in the source for a copy, in the
    ; destination for a delete (robocopy lists a file to delete under the source folder it compares, without /MT)
    static _Row(change, source, destination) {
        item := change.Path
        for root in [RTrim(source, "\") "\", RTrim(destination, "\") "\"]
            if (StrLen(root) > 1 && SubStr(item, 1, StrLen(root)) = root) {
                item := SubStr(item, StrLen(root) + 1)
                break
            }
        full := (item = change.Path) ? item : RTrim(change.Kind = "Delete" ? destination : source, "\") "\" item
        if (item = "")
            item := "."
        size := (!change.Folder && change.HasOwnProp("Size") && change.Size != "") ? SyncProgress.FormatSize(change.Size) : ""
        return [item, SyncPreview.Action(change), size, full]
    }

    ; "Copy (newer)", "New folder", "Delete": why a file is copied comes from robocopy's label, in Windows' language; the English
    ; ones are translated, others shown as robocopy wrote them
    static Action(change) {
        action := I18n.T("Sync.Preview." change.Kind (change.Folder ? "Folder" : "File"))
        if (change.Kind = "Delete" || change.Folder || change.Label = "")
            return action
        static reasons := Map("new file", "New", "newer", "Newer", "older", "Older", "changed", "Changed")
        label := StrLower(Trim(change.Label))
        return I18n.T("Sync.Preview.CopyWhy", reasons.Has(label) ? I18n.T("Sync.Preview.Why." reasons[label]) : Trim(change.Label))
    }
}
