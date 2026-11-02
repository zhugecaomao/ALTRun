;===============================================================================
; SyncPreview.ahk - what a sync task would change, from a preview run (robocopy /L) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like FreeFileSync's "Compare": a summary on top (files to copy, files and folders to delete), the changes below
; (the first SyncJobs.Changes limit of them; the log has all), deletions first. Opened by:
;   - Preview (search: → Preview, Preferences → Sync → Preview, the editor's Preview button)
;   - Run of a mirror that would delete something: Confirm, the window asks before anything is deleted (Close is the default).
;     It comes from a background timer, so it may not get to the front (Windows doesn't let a background program take the
;     focus): it flashes in the taskbar and a notification says it is waiting
; One window per task: another task's preview doesn't close a confirmation still waiting. A mirror previewed more than StaleMs ago
; is previewed (and confirmed) again when run, as the destination may have changed since
;
; Usage:
;   SyncPreview.Show(task, {Summary, Changes, LogFile}, {Run: true | false, Confirm: true | false})
;   SyncPreview.Rows(changes, source, destination)   -> [[action, item, robocopy's label]...]: the list's text, deletions first
;===============================================================================

class SyncPreview {
    static Windows := Map()                                                 ; SyncJobs.Key(name) -> the task's preview window
    static StaleMs := 5 * 60 * 1000

    static Show(task, result, options) {
        key := SyncJobs.Key(task["Name"])
        SyncPreview.Close(key)                                              ; the same task's older preview
        created := A_TickCount
        name := task["Name"], summary := result.Summary, changes := result.Changes
        mirror := RoboSync._Get(task, "Mode") = "Mirror"
        title := I18n.T(options.Confirm ? "Sync.Preview.ConfirmTitle" : "Sync.Preview.Title", name)
        g := Gui("+Resize +MinSize480x300 -MinimizeBox", title)
        SyncPreview.Windows[key] := g
        g.SetFont("s9", ThemeManager.FontName())
        g.MarginX := 14, g.MarginY := 12
        width := 680

        source := RoboSync.Expand(RoboSync._Get(task, "Source")), destination := RoboSync.Expand(RoboSync._Get(task, "Destination"))
        g.AddText("w" width " cGray", source " → " destination)
        deletes := IsObject(summary) ? summary.Files.Extras + summary.Dirs.Extras : 0
        g.SetFont("s10 bold")
        if !IsObject(summary)
            text := I18n.T("Sync.Preview.NoSummary")
        else if (summary.Files.Copied = 0 && deletes = 0 && summary.Dirs.Copied = 0)
            text := I18n.T("Sync.Preview.UpToDate")
        else
            text := I18n.T("Sync.Preview.Copy", summary.Files.Copied, summary.Files.Total, summary.Dirs.Copied)
        g.AddText("w" width " y+8", text)
        if deletes {
            g.AddText("w" width " y+4 cC42B1C", I18n.T("Sync.Preview.Delete", summary.Files.Extras, summary.Dirs.Extras))
        }
        g.SetFont("s9 norm")
        if (options.Confirm)
            g.AddText("w" width " y+6", I18n.T("Sync.Preview.ConfirmText"))

        list := g.AddListView("w" width " r16 y+10 -Multi Grid", [I18n.T("Sync.Preview.Col.Action"), I18n.T("Sync.Preview.Col.Item"), I18n.T("Sync.Preview.Col.Robocopy")])
        rows := SyncPreview.Rows(changes, source, destination)
        list.Opt("-Redraw")
        for row in rows
            list.Add("", row*)
        list.ModifyCol(1, 110), list.ModifyCol(2, width - 110 - 120 - 24), list.ModifyCol(3, 120)
        list.Opt("+Redraw")
        more := (changes.HasOwnProp("Total") && changes.Total > changes.Length) ? I18n.T("Sync.Preview.More", changes.Length, changes.Total) : ""
        g.SetFont("s8")
        note := g.AddText("w" width " y+4 cGray", more != "" ? more : I18n.T("Sync.Preview.Note"))
        g.SetFont("s9")

        logButton := g.AddButton("xm y+12 w120", I18n.T("Sync.Action.OpenLog"))
        logButton.OnEvent("Click", (*) => Path.OpenText(result.LogFile))
        buttons := [logButton]
        closeX := width - 90
        if options.Run {
            runText := I18n.T(options.Confirm ? "Sync.Preview.RunDelete" : "Sync.Preview.Run")
            runButton := g.AddButton("x" (14 + width - 90 - 10 - 150) " yp w150", runText)
            runButton.OnEvent("Click", (*) => SyncPreview._Run(task, created))
            runButton.Enabled := !SyncJobs.IsRunning(name)
            buttons.Push(runButton)
        }
        closeButton := g.AddButton("x" (14 + closeX) " yp w90 Default", I18n.T("Sync.Preview.Close"))
        closeButton.OnEvent("Click", (*) => SyncPreview.Close(key))
        buttons.Push(closeButton)
        g.OnEvent("Escape", (*) => SyncPreview.Close(key))
        g.OnEvent("Close", (*) => SyncPreview.Close(key))
        gaps := SyncPreview._Gaps(list, note, buttons)
        g.OnEvent("Size", (gui, minMax, w, h) => SyncPreview._Layout(minMax, w, h, list, note, buttons, gaps))
        g.Show()
        closeButton.Focus()                                                 ; Enter closes: running (and deleting) needs a click
        if (options.Confirm && !WinActive("ahk_id " g.Hwnd)) {             ; shown from the background: tell where the question is
            DllCall("user32\FlashWindow", "Ptr", g.Hwnd, "Int", 1)
            App.Toast(I18n.T("Sync.Preview.Waiting", name))
        }
        return g
    }

    static Close(key) {
        if !SyncPreview.Windows.Has(key)
            return
        try SyncPreview.Windows[key].Destroy()
        SyncPreview.Windows.Delete(key)
    }

    ; Run now: what a recent preview showed is what runs, so a mirror isn't previewed again; an older mirror preview is done again
    ; (SyncJobs.Run: checks, preview, and the question when it would delete anything)
    static _Run(task, created, now := A_TickCount) {
        SyncPreview.Close(SyncJobs.Key(task["Name"]))
        if SyncPreview.IsStale(task, created, now)
            return SyncJobs.Run(task)
        SyncJobs._Started(SyncJobs.Start(task), task["Name"])
    }

    ; A mirror's preview older than StaleMs (other modes don't delete: their preview can't be out of date in a harmful way)
    static IsStale(task, created, now := A_TickCount) => RoboSync._Get(task, "Mode") = "Mirror" && now - created > SyncPreview.StaleMs

    ; The spaces between the controls as first laid out, in the units GetPos / Move and the Size event use (no DPI arithmetic here)
    static _Gaps(list, note, buttons) {
        list.GetPos(&listX, &listY, , &listH)
        note.GetPos(, &noteY, , &noteH)
        buttons[1].GetPos(, &buttonY)
        last := buttons[buttons.Length], lastX := 0, prevRight := 0
        last.GetPos(&lastX)
        if (buttons.Length >= 3) {
            buttons[buttons.Length - 1].GetPos(&prevX, , &prevW)
            prevRight := prevX + prevW
        }
        return {Margin: listX, ListNote: noteY - listY - listH, NoteButtons: buttonY - noteY - noteH, Between: prevRight ? lastX - prevRight : listX}
    }

    ; Resizing: the list takes the room, the note and the buttons stay below it, Run and Close at the right edge
    static _Layout(minMax, w, h, list, note, buttons, gaps) {
        if (minMax = -1)
            return
        list.GetPos(, &listY)
        buttons[1].GetPos(, , , &buttonH)
        note.GetPos(, , , &noteH)
        buttonY := h - gaps.Margin - buttonH
        noteY := buttonY - gaps.NoteButtons - noteH
        list.Move(, , w - 2 * gaps.Margin, Max(40, noteY - gaps.ListNote - listY))
        note.Move(, noteY, w - 2 * gaps.Margin)
        buttons[1].Move(, buttonY)
        right := w - gaps.Margin
        Loop buttons.Length - 1 {
            button := buttons[buttons.Length - A_Index + 1]
            button.GetPos(, , &bw)
            right -= bw
            button.Move(right, buttonY)
            right -= gaps.Between
        }
    }

    ; The list's rows: deletions first (what matters most before a mirror), each with the item relative to its folder
    static Rows(changes, source, destination) {
        rows := [], copies := []
        roots := [RTrim(source, "\") "\", RTrim(destination, "\") "\"]
        for change in changes {
            item := change.Path
            for root in roots
                if (StrLen(root) > 1 && SubStr(item, 1, StrLen(root)) = root) {
                    item := SubStr(item, StrLen(root) + 1)
                    break
                }
            if (item = "")
                item := "."
            action := I18n.T("Sync.Preview." change.Kind (change.Folder ? "Folder" : "File"))
            (change.Kind = "Delete" ? rows : copies).Push([action, item, change.Label])
        }
        for row in copies
            rows.Push(row)
        return rows
    }
}
