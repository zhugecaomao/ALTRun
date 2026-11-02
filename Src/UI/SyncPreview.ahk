;===============================================================================
; SyncPreview.ahk - what a sync task would change, from a preview run (robocopy /L) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Like FreeFileSync's "Compare": a summary on top (files to copy, files and folders to delete), the changes below
; (the first SyncJobs.Changes limit of them; the log has all), deletions first. Opened by:
;   - Preview (search: → Preview, Preferences → Sync → Preview, the editor's Preview button)
;   - Run of a mirror that would delete something: Confirm, the window asks before anything is deleted (Close is the default)
;
; Usage:
;   SyncPreview.Show(task, {Summary, Changes, LogFile}, {Run: true | false, Confirm: true | false})
;   SyncPreview.Rows(changes, source, destination)   -> [[action, item, robocopy's label]...]: the list's text, deletions first
;===============================================================================

class SyncPreview {
    static Gui := ""

    static Show(task, result, options) {
        if IsObject(SyncPreview.Gui)
            try SyncPreview.Gui.Destroy()
        name := task["Name"], summary := result.Summary, changes := result.Changes
        mirror := RoboSync._Get(task, "Mode") = "Mirror"
        title := I18n.T(options.Confirm ? "Sync.Preview.ConfirmTitle" : "Sync.Preview.Title", name)
        g := Gui("+Resize +MinSize480x300 -MinimizeBox", title)
        SyncPreview.Gui := g
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
            runButton.OnEvent("Click", (*) => SyncPreview._Run(task))
            runButton.Enabled := !SyncJobs.IsRunning(name)
            buttons.Push(runButton)
        }
        closeButton := g.AddButton("x" (14 + closeX) " yp w90 Default", I18n.T("Sync.Preview.Close"))
        closeButton.OnEvent("Click", (*) => g.Destroy())
        buttons.Push(closeButton)
        g.OnEvent("Escape", (*) => g.Destroy())
        g.OnEvent("Close", (*) => g.Destroy())
        g.OnEvent("Size", (gui, minMax, w, h) => SyncPreview._Layout(minMax, w, h, list, note, buttons))
        g.Show()
        closeButton.Focus()                                                 ; Enter closes: running (and deleting) needs a click
        return g
    }

    ; Run now: what the preview showed is what runs, so a mirror isn't previewed again
    static _Run(task) {
        if IsObject(SyncPreview.Gui)
            try SyncPreview.Gui.Destroy()
        SyncJobs._Started(SyncJobs.Start(task), task["Name"])
    }

    ; Resizing: the list takes the room, the note and the buttons stay below it
    static _Layout(minMax, w, h, list, note, buttons) {
        if (minMax = -1)
            return
        list.GetPos(&listX, &listY)
        buttons[1].GetPos(, , , &buttonH)
        note.GetPos(, , , &noteH)
        scale := A_ScreenDPI / 96
        buttonY := h - Round(12 * scale) - buttonH
        noteY := buttonY - Round(12 * scale) - noteH
        listH := noteY - Round(4 * scale) - listY
        list.Move(, , w - 2 * listX, listH)
        note.Move(, noteY, w - 2 * listX)
        buttons[1].Move(, buttonY)
        right := w - listX
        Loop buttons.Length - 1 {                                           ; Run and Close at the right edge
            button := buttons[buttons.Length - A_Index + 1]
            button.GetPos(, , &bw)
            right -= bw
            button.Move(right, buttonY)
            right -= Round(10 * scale)
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
