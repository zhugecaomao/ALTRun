;===============================================================================
; PreferencesTests.ahk - Preferences and the item editors: pages, layout and fit, list editing,
;   hotkey boxes and editor fields
;-------------------------------------------------------------------------------
; Included by Tests\RunTests.ahk. Every static method whose name doesn't start with "_" is a test;
; TestRunner runs them in name order, and each one must put back the shared state it changes.
;===============================================================================
class PreferencesTests {
    static Preferences() {
        eq := (n, a, e) => TestRunner.Equal("Preferences." n, a, e)
        data := PreferencesWindow.DeepCopy(AppSettings.Defaults())
        eq("get", PreferencesWindow.GetPath(data, "General.Hotkey"), "!Space")
        eq("get missing", PreferencesWindow.GetPath(data, "General.NoSuchKey"), "")
        PreferencesWindow.SetPath(data, "Features.Calculator.StructuralCalc", 1)
        eq("set", data["Features"]["Calculator"]["StructuralCalc"], 1)
        PreferencesWindow.SetPath(data, "New.Section.Value", "x")
        eq("set creates", data["New"]["Section"]["Value"], "x")
        original := AppSettings.Defaults()
        copy := PreferencesWindow.DeepCopy(original)
        copy["General"]["Hotkey"] := "^Space"
        eq("deep copy independent", original["General"]["Hotkey"], "!Space")
        eq("lines", TextTools.SplitLines(" a `r`n`r`nb ").Length, 2)
        eq("join lines", TextTools.Join(["a", "b"]), "a`r`nb")
        eq("csv", TextTools.SplitList("open, find ,, x")[3], "x")
        eq("join csv", TextTools.JoinList(["open", "find"]), "open, find")
        sizes := TextTools.SplitList("10, 12,16", true)             ; number lists (bar diameters) are stored as numbers
        eq("numeric csv", (sizes[1] is Number) && sizes[3] = 16 ? 1 : 0, 1)
        eq("text csv stays text", TextTools.SplitList("10")[1] is String ? 1 : 0, 1)
    }

    ; Lists edited in their own window ("Edit..."): the summary next to the button, splitting the edited text
    static EditList() {
        eq := (n, a, e) => TestRunner.Equal("EditList." n, a, e)
        eq("summary", TextTools.JoinList(["a.exe", "b.exe"]), "a.exe, b.exe")
        eq("empty summary", TextTools.JoinList([]), "")
        eq("file types: keywords", PrefsPages.TypeFilterSummary(["doc = doc docx", "bad line", "pic = png"]), "doc, pic")
        eq("defaults", PrefsPages.TypeFilterSummary(AppSettings.Defaults()["Features"]["FileSearch"]["TypeFilters"]), "doc, pic, video, audio, zip, exe, cad")
        lines := TextTools.SplitLines("  doc = doc docx `r`n`r`n pic = png`n")
        eq("split", lines.Length "|" lines[1] "|" lines[2], "2|doc = doc docx|pic = png")
        eq("help text", I18n.Strings.Has("Prefs.TypeFilters.Help"), true)
        eq("app names", PrefsPages.NameSummary(["C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Administrative Tools\Disk Cleanup.lnk"
            , "shell:AppsFolder\NVIDIACorp.NVIDIAControlPanel_56jybvy8sckqj!NVIDIACorp.NVIDIAControlPanel"]), "Disk Cleanup, NVIDIACorp.NVIDIAControlPanel")
        for key in ["Prefs.SysHidden", "Prefs.AppFolders", "Prefs.AppHidden", "Prefs.ScopeFolders", "Prefs.ClipIgnoreApps", "Prefs.QSExclude"
                  , "Prefs.QSAutoExclude", "Prefs.QSOtherDialogs", "Prefs.RenameWindows", "Prefs.AppendWindows", "Prefs.ExpandExclude"]
            eq("help: " key, I18n.Strings.Has(key ".Help"), true)
    }

    ; "Add Examples" / "Add Built-ins" on the list pages: only the defaults missing from the list are added, the rest are counted as skipped
    static ListExamples() {
        eq := (n, a, e) => TestRunner.Equal("ListExamples." n, a, e)
        pages := [["Hotkeys", (item) => PrefsPages.HotkeyKeys(item)], ["CustomCommands", (item) => PrefsPages.CommandKeys(item)]
            , ["Snippets", (item) => PrefsPages.SnippetKeys(item)], ["Features.WebSearch.Engines", (item) => PrefsPages.EngineKeys(item)]
            , ["SelectionTypes", (item) => PrefsPages.SelectionTypeKeys(item)]]
        for page in pages {
            defaults := PreferencesWindow.GetPath(AppSettings.Defaults(), page[1])
            items := []
            result := PrefsForm.AddMissing(items, defaults, page[2])
            eq(page[1] ": empty list gets all", result.Added "|" result.Skipped "|" items.Length, defaults.Length "|0|" defaults.Length)
            result := PrefsForm.AddMissing(items, PreferencesWindow.GetPath(AppSettings.Defaults(), page[1]), page[2])
            eq(page[1] ": again adds none", result.Added "|" result.Skipped "|" items.Length, "0|" defaults.Length "|" defaults.Length)
        }

        hotkeys := [Map("Key", "!^F", "Action", "ToggleWindow", "WinTitle", "")       ; Ctrl+Alt+F in another order and case: the same key
                  , Map("Key", "^r", "Action", "ToggleWindow", "WinTitle", "")         ; Ctrl+R everywhere: not the same as Ctrl+R only in ALTRun
                  , Map("Key", "~^!#v", "Action", "ToggleWindow", "WinTitle", "")]     ; ~ doesn't change the key
        result := PrefsForm.AddMissing(hotkeys, AppSettings.Defaults()["Hotkeys"], (item) => PrefsPages.HotkeyKeys(item))
        eq("hotkeys: counts", result.Added "|" result.Skipped, "4|2")
        eq("hotkeys: own kept first", hotkeys[1]["Action"] "|" hotkeys.Length, "ToggleWindow|7")
        eq("hotkeys: added in order", hotkeys[4]["Key"] "|" hotkeys[7]["Key"] "|" hotkeys[7]["WinTitle"], "^!r|^r|ALTRun")
        eq("hotkey key", PrefsPages.HotkeyKeys(Map("Key", "+^!A", "WinTitle", " ahk_exe a.exe "))[1], "^!+a|ahk_exe a.exe")
        eq("hotkey empty", PrefsPages.HotkeyKeys(Map("Key", "", "WinTitle", "x"))[1], "")
        sameWindow := [Map("Key", "^R", "WinTitle", "altrun")]                          ; window names compared regardless of case
        eq("hotkeys: window case", PrefsForm.AddMissing(sameWindow, [Map("Key", "^r", "WinTitle", "ALTRun")], (item) => PrefsPages.HotkeyKeys(item)).Added, 0)

        commands := [Map("Title", "notepad", "Keyword", "")                             ; the same title (any case)
                   , Map("Title", "My downloads", "Keyword", "x, DL")                    ; one of several keywords
                   , Map("Title", "", "Keyword", "")]                                    ; an empty title matches nothing
        result := PrefsForm.AddMissing(commands, AppSettings._DefaultCommands(), (item) => PrefsPages.CommandKeys(item))
        eq("commands: counts", result.Added "|" result.Skipped, (AppSettings._DefaultCommands().Length - 2) "|2")

        snippets := [Map("Name", "Mine", "Keyword", "THX"), Map("Name", "today's date", "Keyword", "d")]
        result := PrefsForm.AddMissing(snippets, AppSettings._DefaultSnippets(), (item) => PrefsPages.SnippetKeys(item))
        eq("snippets: counts", result.Added "|" result.Skipped, (AppSettings._DefaultSnippets().Length - 2) "|2")

        engines := [Map("Id", "", "Keyword", "g", "Title", "My search")                  ; a keyword of a built-in engine, taken by the user's own one
                  , Map("Id", "bing", "Keyword", "b", "Title", "Bing")]                  ; a built-in one with its keyword changed
        result := PrefsForm.AddMissing(engines, AppSettings._DefaultEngines(), (item) => PrefsPages.EngineKeys(item))
        eq("engines: counts", result.Added "|" result.Skipped, (AppSettings._DefaultEngines().Length - 2) "|2")
        eq("engines: no empty key", PrefsPages.EngineKeys(Map("Id", "", "Keyword", ""))[1], "")
    }

    ; Custom command editor: with the type set to "Folder" the browse button chooses a folder
    static BrowseKind() {
        eq := (n, a, e) => TestRunner.Equal("BrowseKind." n, a, e)
        fields := CustomCommandProvider.EditorFields()
        eq("target has rule", fields[3].FolderWhen[2], "Folder")
        controls := Map("Type", {Value: 2})                                  ; item 2 = Folder
        eq("folder type", ItemEditor.KindFor(fields[3].FolderWhen, fields, controls), "folder")
        controls["Type"].Value := 1
        eq("file type", ItemEditor.KindFor(fields[3].FolderWhen, fields, controls), "file")
        controls["Type"].Value := 4
        eq("url type", ItemEditor.KindFor(fields[3].FolderWhen, fields, controls), "file")
        eq("no such field", ItemEditor.KindFor(["Nope", "Folder"], fields, controls), "file")
    }

    ; A custom command's "Selection types" (#434): a box of type names with a menu that checks / unchecks each, saved only when set
    static CommandSelectionField() {
        eq := (n, a, e) => TestRunner.Equal("CommandSelectionField." n, a, e)
        shown := ""
        fields := CustomCommandProvider.EditorFields()
        field := fields[fields.Length]
        eq("last field", field.Key "|" field.Type "|" field.OmitEmpty "|" field.ButtonLabel, "Selection|tags|1|Types")
        eq("tags text", ItemEditor.TagsText(" PO,pdf ,, po , Images") "|" ItemEditor.TagsText(["IP", "Email"]), "PO, pdf, Images|IP, Email")
        eq("toggle on", ItemEditor.ToggleTag("PO", "Email"), "PO, Email")
        eq("toggle off (any case)", ItemEditor.ToggleTag("PO, email, IP", "Email"), "PO, IP")
        eq("toggle the last off", ItemEditor.ToggleTag("Email", "email"), "")
        savedTypes := AppSettings.Data["SelectionTypes"]
        AppSettings.Data["SelectionTypes"] := [Map("Name", "PO", "Kind", "Text", "Pattern", "^PO\d+$")]
        try {
            names := ""
            for each in SelectionActions.TypeNamesForEditor()
                names .= each "|"
            eq("menu names: built-in, then the user's", names, "Text|Url|File|Folder|Files|-|PO|")
            box := {Value: "po, Files"}
            tagMenu := ItemEditor._TagMenu(field, box)
            eq("menu checks what is in the box", tagMenu[1].Checked "|" tagMenu[5].Checked "|" tagMenu[7].Checked "|" tagMenu[6], "0|1|1|-")
            tagMenu[1].Click.Call("", "")
            eq("menu click adds", box.Value, "po, Files, Text")
            eq("known types: no warning", SelectionActions.UnknownTypesWarning("text, PO"), "")
            eq("unknown type warned", SelectionActions.UnknownTypesWarning("PO, Invoice, PDFs"), "Not a content type: Invoice, PDFs. Add it in Preferences → Selection, or the command won't be shown for it.")
            ; the dialog itself: what is typed is tidied, and an empty box leaves no "Selection" key
            edited := CoreFieldsDialog(Map("Title", "Ping", "Type", "Command", "Target", "cmd.exe", "Arguments", "", "Keyword", "", "Selection", "IP"), "Text ,po,")
            eq("shown in the box", shown, "IP")
            eq("saved tidied", IsObject(edited) ? edited["Selection"] : "no result", "Text, po")
            edited := CoreFieldsDialog(Map("Title", "Ping", "Type", "Command", "Target", "cmd.exe", "Arguments", "", "Keyword", "", "Selection", "IP"), "")
            eq("emptied: key removed", IsObject(edited) ? edited.Has("Selection") : "no result", false)
            edited := CoreFieldsDialog(Map("Title", "Notes", "Type", "File", "Target", "notepad.exe", "Arguments", "", "Keyword", ""), "")
            eq("never set: no key", IsObject(edited) ? edited.Has("Selection") : "no result", false)
            edited := CoreFieldsDialog(Map("Title", "Notes", "Type", "File", "Target", "notepad.exe", "Arguments", "", "Keyword", "", "Selection", ["Text", "PO"]), "keep")
            eq("an array is shown joined", shown, "Text, PO")
            eq("kept as text", IsObject(edited) ? edited["Selection"] : "no result", "Text, PO")
        } finally
            AppSettings.Data["SelectionTypes"] := savedTypes

        ; Opens the command editor with item, types into "Selection types" ("keep" leaves it as shown), clicks OK; the edited Map or "".
        ; shown: what the box showed when the dialog opened
        CoreFieldsDialog(item, typed) {
            shown := ""
            SetTimer(Fill, -200)
            return ItemEditor.Edit(ItemEditor.Owner(), "Selection Field Test", fields, item)
            Fill() {
                if !(hwnd := WinExist("Selection Field Test ahk_class AutoHotkeyGUI ahk_pid " ProcessExist())) {
                    SetTimer(Fill, -100)
                    return
                }
                selectionBox := ControlGetHwnd("Edit5", hwnd)                ; Title, Target, Arguments, Keyword, Selection (Type is a drop-down list)
                shown := ControlGetText(selectionBox)
                if (typed != "keep")
                    ControlSetText(typed, selectionBox)
                PostMessage(0xF5, 0, 0, ControlGetHwnd("OK", hwnd))          ; BM_CLICK
            }
        }
    }

    ; A command's "Triggers" line and its hotkeys (8.2): the text, writing the dialog's draft back, and the dialog itself
    static CommandTriggers() {
        eq := (n, a, e) => TestRunner.Equal("CommandTriggers." n, a, e)
        savedLang := I18n.Lang
        I18n.Init("en")
        hk(key, id := "pg7k2m4x", winTitle := "") => Map("Key", key, "Action", "custom:" id, "SearchText", "", "WinTitle", winTitle)
        eq("all three", CustomCommandProvider.TriggerText(Map("Keyword", "ping, pg", "Selection", "IP,Text files"), [hk("^!p"), hk("~#p", , "ahk_exe x.exe")])
            , "Keyword ping, pg · Selection IP, Text files · Hotkey Ctrl+Alt+P, Win+P")
        eq("hotkey only", CustomCommandProvider.TriggerText(Map("Keyword", " "), [hk("^!p")]), "Hotkey Ctrl+Alt+P")
        eq("none", CustomCommandProvider.TriggerText(Map("Title", "x"), [hk("")]), "Only found by its name in the search")

        ; the draft written back: edited in place, deleted, added at the end; the other hotkeys untouched
        other := Map("Key", "!q", "Action", "Lock", "WinTitle", "")
        first := hk("^!p"), second := hk("^!o")
        hotkeys := [first, other, second]
        draft := [{Entry: hk("^!+p"), Original: first}, {Entry: hk("^!n"), Original: ""}]
        eq("changed", CustomCommandProvider.MergeHotkeys(hotkeys, "custom:pg7k2m4x", draft), true)
        keys := ""
        for entry in hotkeys
            keys .= entry["Key"] "|"
        eq("merged", keys, "^!+p|!q|^!n|")
        eq("edited in place (the list keeps the same Map)", ObjPtr(hotkeys[1]) = ObjPtr(first) && ObjPtr(hotkeys[2]) = ObjPtr(other), true)
        eq("unchanged", CustomCommandProvider.MergeHotkeys(hotkeys, "custom:pg7k2m4x", [{Entry: hk("^!+p"), Original: hotkeys[1]}, {Entry: hk("^!n"), Original: hotkeys[3]}]), false)
        eq("all deleted", CustomCommandProvider.MergeHotkeys(hotkeys, "custom:pg7k2m4x", []) "|" hotkeys.Length, "1|1")

        ; the dialog: a new command gets its Id when it opens, the line shows its triggers and follows the keyword typed; Cancel changes nothing
        settings := {Commands: [Map("Id", "taken001", "Title", "A")], Hotkeys: [other], Value: (path) => ""}
        seen := {Text: "", Typed: "", Fits: false, Button: ""}
        try Dialogs()
        finally I18n.Init(savedLang)                                        ; a failure in a dialog doesn't leave the language changed for the tests after this one

        Dialogs() {
            result := Dialog(CustomCommandProvider.NewCommand(), true)
            eq("new: an Id", IsObject(result) && AppSettings.IsId(result.Command["Id"]) && result.Command["Id"] != "taken001", true)
            eq("new: no Triggers key saved", IsObject(result) && !result.Command.Has("Triggers"), true)
            eq("new: no hotkeys changed", IsObject(result) && !result.HotkeysChanged && settings.Hotkeys.Length = 1, true)
            eq("line when opened", seen.Text, "Only found by its name in the search")
            eq("line follows the keyword", seen.Typed, "Keyword np")
            eq("button", seen.Button, "Set Hotkey...")
            eq("fits", seen.Fits, true)
            existing := Map("Id", "pg7k2m4x", "Title", "Ping", "Type", "Command", "Target", "cmd.exe", "Arguments", "/k ping {query}", "Keyword", "ping", "Selection", "IP")
            settings.Hotkeys.Push(hk("^!p"))
            eq("cancelled", Dialog(existing, false), "")
            eq("existing: its hotkey on the line", seen.Text, "Keyword ping · Selection IP · Hotkey Ctrl+Alt+P")
            for language in I18n.Languages() {                                  ; the line and its button inside the window in every language
                I18n.Init(language[1])
                Dialog(existing, false)
                eq(language[1] " fits", seen.Fits, true)
            }
        }

        ; Opens the dialog, notes the line (and after typing a keyword), whether everything is inside the window, then OK / Cancel
        Dialog(item, ok) {
            SetTimer(Look, -200)
            SetTimer(GiveUp, -15000)                                        ; still open then: closed, so the test fails instead of hanging the run
            result := CustomCommandProvider.EditDialog(ItemEditor.Owner(), "Triggers Test", item, settings)
            SetTimer(GiveUp, 0)
            return result
            GiveUp() {
                DetectHiddenWindows(true)
                if (hwnd := WinExist("Triggers Test ahk_class AutoHotkeyGUI ahk_pid " ProcessExist())) {
                    seen.Text := "(still open after 15 s)", seen.Fits := false
                    GuiFromHwnd(hwnd).Destroy()
                }
            }
            Look() {
                DetectHiddenWindows(false)                                  ; only once shown: being built, it would only be hidden by WM_CLOSE (no Close handler yet) and shown again
                if !(hwnd := WinExist("Triggers Test ahk_class AutoHotkeyGUI ahk_pid " ProcessExist())) || !DllCall("IsWindowVisible", "Ptr", hwnd) {
                    SetTimer(Look, -100)
                    return
                }
                g := GuiFromHwnd(hwnd)
                box := "", fits := true
                WinGetClientPos(, , &clientW, &clientH, hwnd)
                for ctrl in g {
                    ctrl.GetPos(&x, &y, &w, &h)
                    fits := fits && x + w <= clientW && y + h <= clientH
                    if (ctrl.Type = "Edit" && ControlGetStyle(ctrl) & 0x800)   ; ES_READONLY: the Triggers line
                        box := ctrl
                    if (ctrl.Type = "Button" && ctrl.Text = I18n.T("Cmd.SetHotkey")) {
                        seen.Button := ctrl.Text
                        box.GetPos(&boxX, &boxY, &boxW)
                        fits := fits && x >= boxX + boxW && Abs(y - boxY) <= 4
                    }
                }
                seen.Fits := fits
                seen.Text := IsObject(box) ? box.Value : "(no line)"
                if ok {
                    ControlSetText("np", ControlGetHwnd("Edit4", hwnd))        ; Title, Target, Arguments, Keyword (Type is a drop-down list)
                    Sleep(50)
                    seen.Typed := box.Value
                    ControlSetText("Notepad", ControlGetHwnd("Edit1", hwnd))
                    ControlSetText("notepad.exe", ControlGetHwnd("Edit2", hwnd))
                    for ctrl in g                                           ; the button (the line's description mentions "OK" too)
                        if (ctrl.Type = "Button" && ctrl.Text = "OK")
                            PostMessage(0xF5, 0, 0, ctrl)                   ; BM_CLICK
                } else
                    WinClose(hwnd)
            }
        }
    }

    ; The command edit dialog asks before saving placeholders that won't work: Yes saves, No goes back to the dialog
    static CommandPlaceholderQuestion() {
        eq := (n, a, e) => TestRunner.Equal("CommandPlaceholderQuestion." n, a, e)
        savedLang := I18n.Lang
        I18n.Init("en")
        settings := {Commands: [], Hotkeys: [], Value: (path) => ""}
        command := Map("Id", "pe7k2m4x", "Title", "Echo", "Type", "Command", "Target", "cmd.exe", "Arguments", "/k echo {selection} {date}", "Keyword", "e")
        seen := {Question: "", Answer: ""}
        try {
            result := OpenAndAnswer("Yes")
            eq("asked: {selection}", InStr(seen.Question, "{selection} isn't replaced in commands") > 0, true)
            eq("asked: {date}", InStr(seen.Question, "Not a placeholder in commands: {date}.") > 0, true)
            eq("asked: save anyway", InStr(seen.Question, I18n.T("Editor.SaveAnyway")) > 0, true)
            eq("Yes: saved", IsObject(result) ? result.Command["Arguments"] : "no result", "/k echo {selection} {date}")
            result := OpenAndAnswer("No")
            eq("No: back in the dialog", seen.Answer, "still open")
            eq("No: then closed, nothing saved", result, "")
        } finally I18n.Init(savedLang)

        ; Opens the dialog, clicks OK, answers the question, then (after No) closes the dialog; the dialog's result
        OpenAndAnswer(answer) {
            seen.Question := "", seen.Answer := ""
            SetTimer(PressOk, -200)
            SetTimer(GiveUp, -15000)                                        ; still open then: closed, so the test fails instead of hanging the run
            result := CustomCommandProvider.EditDialog(ItemEditor.Owner(), "Placeholder Test", command, settings)
            SetTimer(GiveUp, 0)
            return result
            Shown(class) {
                DetectHiddenWindows(false)
                hwnd := WinExist("Placeholder Test ahk_class " class " ahk_pid " ProcessExist())
                return (hwnd && DllCall("IsWindowVisible", "Ptr", hwnd)) ? hwnd : 0
            }
            PressOk() {
                if !(hwnd := Shown("AutoHotkeyGUI"))
                    return SetTimer(PressOk, -100)
                for ctrl in GuiFromHwnd(hwnd)
                    if (ctrl.Type = "Button" && ctrl.Text = "OK")
                        PostMessage(0xF5, 0, 0, ctrl)                       ; BM_CLICK
                SetTimer(AnswerQuestion, -200)
            }
            AnswerQuestion() {
                if !(box := Shown("#32770"))
                    return SetTimer(AnswerQuestion, -100)
                seen.Question := WinGetText(box)
                PostMessage(0x111, (answer = "Yes") ? 6 : 7, 0, box)       ; WM_COMMAND IDYES / IDNO
                if (answer = "No")
                    SetTimer(CloseAfterNo, -100)
            }
            CloseAfterNo() {                                                ; only once the question has returned: the Target box got the focus back
                if !(hwnd := Shown("AutoHotkeyGUI"))
                    return
                try focused := ControlGetClassNN(ControlGetFocus(hwnd))
                catch
                    focused := ""
                if (focused != "Edit2")                                     ; Title, Target, ... (closing the window under the open question would break it)
                    return SetTimer(CloseAfterNo, -100)
                seen.Answer := "still open"
                WinClose(hwnd)
            }
            GiveUp() {
                DetectHiddenWindows(true)
                if (box := WinExist("Placeholder Test ahk_class #32770 ahk_pid " ProcessExist()))
                    WinClose(box)
                if (hwnd := WinExist("Placeholder Test ahk_class AutoHotkeyGUI ahk_pid " ProcessExist()))
                    GuiFromHwnd(hwnd).Destroy()
            }
        }
    }

    static ThemeGallery() {
        eq := (n, a, e) => TestRunner.Equal("ThemeGallery." n, a, e)
        pixel(hbm, x, y) {
            dc := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
            old := DllCall("SelectObject", "Ptr", dc, "Ptr", hbm, "Ptr")
            value := DllCall("GetPixel", "Ptr", dc, "Int", x, "Int", y, "UInt")
            DllCall("SelectObject", "Ptr", dc, "Ptr", old), DllCall("DeleteDC", "Ptr", dc)
            return value
        }
        for themeName in ThemeManager.Names() {
            hbm := ThemePreview.ForName(themeName, 112, 76)
            TestRunner.True("ThemeGallery.bitmap " themeName, hbm != 0)
            theme := ThemeManager.Resolve(themeName = "System" ? "Light" : themeName)
            if (hbm && themeName != "System") {
                eq("background " themeName, pixel(hbm, 56, 6), Win.ColorToBgr(theme["Background"]))
                eq("selected row " themeName, pixel(hbm, 104, 26), Win.ColorToBgr(theme["SelectedBackground"]))
            }
            if hbm
                DllCall("DeleteObject", "Ptr", hbm)
        }
        il := ThemePreview.ImageList(["Dark"], 112, 76)                       ; drawn through the image list: the background is opaque (GDI+ rounded corners mustn't make it transparent)
        screen := DllCall("GetDC", "Ptr", 0, "Ptr")
        dc := DllCall("CreateCompatibleDC", "Ptr", screen, "Ptr"), canvas := DllCall("CreateCompatibleBitmap", "Ptr", screen, "Int", 112, "Int", 76, "Ptr")
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", screen)
        oldCanvas := DllCall("SelectObject", "Ptr", dc, "Ptr", canvas, "Ptr")
        rect := Buffer(16), NumPut("Int", 0, "Int", 0, "Int", 112, "Int", 76, rect)
        DllCall("FillRect", "Ptr", dc, "Ptr", rect, "Ptr", DllCall("GetStockObject", "Int", 0, "Ptr"))   ; white
        DllCall("comctl32\ImageList_Draw", "Ptr", il, "Int", 0, "Ptr", dc, "Int", 0, "Int", 0, "UInt", 0)   ; ILD_NORMAL
        eq("image list background opaque", DllCall("GetPixel", "Ptr", dc, "Int", 56, "Int", 6, "UInt"), Win.ColorToBgr(ThemeManager.Resolve("Dark")["Background"]))
        DllCall("SelectObject", "Ptr", dc, "Ptr", oldCanvas), DllCall("DeleteObject", "Ptr", canvas), DllCall("DeleteDC", "Ptr", dc)
        DllCall("comctl32\ImageList_Destroy", "Ptr", il)
        hbm := ThemePreview.ForName("System", 112, 76)                      ; Light on the left half, Dark on the right
        eq("system left", pixel(hbm, 20, 6), Win.ColorToBgr(ThemeManager.Defaults()["Background"]))
        eq("system right", pixel(hbm, 100, 6), Win.ColorToBgr(ThemeManager.Resolve("Dark")["Background"]))
        DllCall("DeleteObject", "Ptr", hbm)

        PreferencesWindow.Show(1, -3000, -3000)
        PreferencesWindow.BuildAll()                                         ; pages are built when needed: build them all before checking each page
        Sleep(400)                                                          ; only changes 300 ms after opening count (see PreferencesWindow.Show)
        focused := DllCall("GetFocus", "Ptr")                               ; on opening the focus is in the page list (↑ ↓ switch pages without changing the first page's settings)
        eq("focus on page list", focused = PreferencesWindow.PageList.Hwnd, true)
        nextTab := DllCall("GetNextDlgTabItem", "Ptr", PreferencesWindow.Gui.Hwnd, "Ptr", PreferencesWindow.PageList.Hwnd, "Int", 0, "Ptr")
        onPage := false                                                     ; Tab: after the list come the current page's controls, not the bottom buttons
        for ctrl in PreferencesWindow.Pages[1].Controls
            onPage := onPage || ctrl.Hwnd = nextTab
        eq("tab from page list goes into the page", onPage, true)
        keys := ""                                                         ; PageKeys matches the pages actually built (the subpages in search rely on it)
        for page in PreferencesWindow.Pages
            keys .= page.Key " "
        expected := ""
        for pageKey in PreferencesWindow.PageKeys
            expected .= pageKey " "
        eq("page keys match the pages", keys, expected)
        SystemProvider._PageOpener("Prefs.Page.Calculator")()               ; with the window open, jump to that page
        eq("open a page from search", PreferencesWindow._page, PreferencesWindow._PageIndex("Prefs.Page.Calculator"))
        eq("page list follows", PreferencesWindow.PageList.GetNext(), PreferencesWindow._page)
        eq("page list rows", PreferencesWindow.PageList.GetCount(), PreferencesWindow.Pages.Length)
        appearancePage := PreferencesWindow._PageIndex("Prefs.Page.Appearance")
        PreferencesWindow.Show(appearancePage)                              ; opening by number while already open
        eq("open by number while open", PreferencesWindow._page "|" PreferencesWindow.PageList.GetNext(), appearancePage "|" appearancePage)
        gallery := ""                                                       ; the theme list on the Appearance page: exactly 4 thumbnails per row, no wide gap on the right
        for ctrl in PreferencesWindow.Pages[appearancePage].Controls
            if (ctrl.Type = "ListView" && !IsObject(gallery))
                gallery := ctrl
        itemPos := (index) => (point := Buffer(8, 0), SendMessage(0x1010, index, point.Ptr, gallery), [NumGet(point, 0, "Int"), NumGet(point, 4, "Int")])   ; LVM_GETITEMPOSITION
        spacing := SendMessage(0x1033, 0, 0, gallery) & 0xFFFF              ; LVM_GETITEMSPACING (large icons)
        galleryRect := Buffer(16)
        DllCall("GetClientRect", "Ptr", gallery.Hwnd, "Ptr", galleryRect)
        first := itemPos(0), fourth := itemPos(3), fifth := itemPos(4)
        eq("gallery: 4 per row", (fourth[2] = first[2] && fifth[2] > first[2]) ? 1 : 0, 1)
        TestRunner.True("ThemeGallery.gallery fills the row (left " (NumGet(galleryRect, 8, "Int") - 4 * spacing) ")"
            , NumGet(galleryRect, 8, "Int") - 4 * spacing < 8)
        PreferencesWindow.PageList.Modify(appearancePage, "-Select")        ; a click on the blank area below the list: the current page is selected again on release
        blankClick := Buffer(3 * A_PtrSize + 40, 0)                              ; NMITEMACTIVATE, iItem = -1 (no row hit)
        NumPut("Int", -1, blankClick, 3 * A_PtrSize)
        PreferencesWindow._OnPageClick(blankClick.Ptr)
        eq("blank click keeps the page selected", PreferencesWindow.PageList.GetNext(), appearancePage)
        draw := Buffer(6 * A_PtrSize + 40, 0)                               ; NMCUSTOMDRAW: no dotted focus rectangle on the selected row
        NumPut("UInt", 0x1, draw, 3 * A_PtrSize)
        eq("custom draw asks for each row and after painting", PreferencesWindow._OnPageListDraw(draw.Ptr), 0x30)
        NumPut("UInt", 0x10001, draw, 3 * A_PtrSize), NumPut("UInt", 0x11, draw, 6 * A_PtrSize + 16)   ; CDIS_SELECTED | CDIS_FOCUS
        PreferencesWindow._OnPageListDraw(draw.Ptr)
        eq("row drawn without the focus rectangle", NumGet(draw, 6 * A_PtrSize + 16, "UInt"), 0x1)
        listRect := Buffer(16)
        DllCall("GetClientRect", "Ptr", PreferencesWindow.PageList.Hwnd, "Ptr", listRect)
        rowRect := (row) => (rect := Buffer(16, 0), SendMessage(0x100E, row - 1, rect, PreferencesWindow.PageList), [NumGet(rect, 0, "Int"), NumGet(rect, 4, "Int"), NumGet(rect, 8, "Int"), NumGet(rect, 12, "Int")])   ; LVM_GETITEMRECT, LVIR_BOUNDS
        last := rowRect(PreferencesWindow.PageKeys.Length)
        style := DllCall("GetWindowLong", "Ptr", PreferencesWindow.PageList.Hwnd, "Int", -16)
        eq("every page fits in the list, no scroll bar", "bottom " (last[4] <= NumGet(listRect, 12, "Int") ? "fits" : last[4] " > " NumGet(listRect, 12, "Int"))
            . ", scroll bars " ((style & 0x300000) ? Format("{:#x}", style & 0x300000) : "none"), "bottom fits, scroll bars none")   ; WS_VSCROLL | WS_HSCROLL
        if (SendMessage(0x108F, 0, 0, PreferencesWindow.PageList) = 4) {   ; LVM_GETVIEW: tiles (Windows; Wine has no tile view and shows a plain list)
            wrong := ""                                                     ; a gap before each group's first page, none inside a group
            Loop PreferencesWindow.PageKeys.Length - 1 {
                row := A_Index                                              ; (A_Index inside the for loop below is the for loop's)
                gap := rowRect(row + 1)[2] - rowRect(row)[4]
                starts := false
                for key in PreferencesWindow.GroupStarts
                    starts := starts || PreferencesWindow.PageKeys[row + 1] = key
                if (starts ? gap < 9 : gap > 4)
                    wrong .= PreferencesWindow.PageKeys[row + 1] " (" gap ") "
            }
            eq("gaps only between the groups", wrong, "")
            first := rowRect(1)
            width := NumGet(listRect, 8, "Int")                             ; tile view keeps 4 px right of the tiles (as below the last one): a wider tile scrolls
            eq("rows as wide as the list", (first[1] <= 2 && first[3] >= width - 6) ? "yes" : first[1] "-" first[3] " in " width, "yes")
            eq("lines in the gaps", PreferencesWindow._gapLines.Length, PreferencesWindow.GroupStarts.Length)
            PreferencesWindow.PageList.Modify(0, "-Select"), PreferencesWindow.PageList.Modify(PreferencesWindow._PageIndex("Prefs.Page.Snippets"), "Select Focus")
            ControlSend("{Down}", PreferencesWindow.PageList)               ; ↓ across a gap goes to the next group's first page
            Sleep(50)
            eq("down across a gap", PreferencesWindow.PageList.GetNext(), PreferencesWindow._PageIndex("Prefs.Page.Commands"))
        } else {
            listRect := Buffer(16)                                          ; the only column fills the list width (no column separator left)
            DllCall("GetClientRect", "Ptr", PreferencesWindow.PageList.Hwnd, "Ptr", listRect)
            eq("column fills the list", SendMessage(0x101D, 0, 0, PreferencesWindow.PageList), NumGet(listRect, 8, "Int"))   ; LVM_GETCOLUMNWIDTH
        }
        PreferencesWindow.PageList.Modify(5, "Select Focus")                ; the user selected a page in the list (click or ↑ ↓)
        Sleep(50)
        eq("selecting a row shows the page", PreferencesWindow._page, 5)
        wrong := 0                                                          ; switching pages shows only that page's controls, the other pages are hidden
        for index, page in PreferencesWindow.Pages
            for ctrl in page.Controls
                wrong += (ctrl.Visible != (index = 5))
        eq("only the current page is visible", wrong, 0)
        eq("every page has an icon", NavIcons.MissingPages(PreferencesWindow.PageKeys), "")
        if (NavIcons.FontName() != "") {                                    ; Windows has the icon fonts (Wine doesn't)
            il := NavIcons.ImageList(PreferencesWindow.PageKeys)
            eq("one icon per page", DllCall("comctl32\ImageList_GetImageCount", "Ptr", il), PreferencesWindow.PageKeys.Length)
            DllCall("comctl32\ImageList_Destroy", "Ptr", il)
            eq("glyphs exist in " NavIcons.FontName(), NavIcons.MissingGlyphs(), "")
            eq("glyphs within their box", NavIcons.GlyphsOutsideBox(), "")   ; ImageList converts only the box to gray
            eq("glyphs within their box, low row", NavIcons.GlyphsOutsideBox(18), "")
        }
        PreferencesWindow.SelectPage(PreferencesWindow._PageIndex("Prefs.Page.Appearance"))
        page := PreferencesWindow.Pages[PreferencesWindow._PageIndex("Prefs.Page.Appearance")]
        gallery := ""
        for ctrl in page.Controls
            if (ctrl.Type = "ListView")
                gallery := ctrl
        themes := ThemeManager.Names()
        eq("one item per theme", gallery.GetCount(), themes.Length)
        current := AppSettings.Appearance["Theme"]
        eq("current theme selected", themes[gallery.GetNext()], current)
        target := (current = "Dark") ? "Monokai" : "Dark"
        for index, themeName in themes
            if (themeName = target)
                gallery.Modify(index, "Select Focus")
        Sleep(50)
        eq("dirty after picking", (PreferencesWindow._ready ? "ready " : "not ready ") (PreferencesWindow._dirty ? 1 : 0), "ready 1")
        picked := ""
        for bind in PreferencesWindow.Binds
            if (bind.Path = "Appearance.Theme")
                picked := bind.Read.Call()
        eq("picked theme", picked, target)
        PreferencesWindow._dirty := false
        PreferencesWindow.Close()
    }

    ; Preferences: after Apply it reopens with "-Preferences page x y"; Help opens each page's Wiki page
    static PreferencesButtons() {
        eq := (n, a, e) => TestRunner.Equal("PreferencesButtons." n, a, e)
        args := PreferencesWindow.ParseArgs(["-Preferences", "Hotkeys", "120", "-8"])
        eq("page", args.Page, "Prefs.Page.Hotkeys")
        eq("x", args.X, 120)
        eq("y (negative, second monitor)", args.Y, -8)
        args := PreferencesWindow.ParseArgs(["-Preferences", "qspanel"])
        eq("page only (any case)", args.Page "|" args.X "|" args.Y, "Prefs.Page.QSPanel||")
        eq("page by number (current order)", PreferencesWindow.ParseArgs(["-Preferences", "3"]).Page, PreferencesWindow.PageKeys[3])
        eq("unknown page", PreferencesWindow.ParseArgs(["-Preferences", "Nope"]).Page, 1)
        eq("number out of range", PreferencesWindow.ParseArgs(["-Preferences", "99"]).Page, 1)
        args := PreferencesWindow.ParseArgs(["-Preferences"])
        eq("default page", args.Page, 1)
        eq("page name", PreferencesWindow.PageName("Prefs.Page.DelayedStart"), "DelayedStart")
        layout := PreferencesWindow.PageLayout(591, 23, [6, 13, 17, 20, 22])   ; the page list at 96 DPI: 23 rows, 5 gaps
        eq("layout: row height", layout.RowH, 23)
        eq("layout: gap (the rest of the height spread over the gaps)", layout.Gap, 12)
        eq("layout: tops", layout.Tops[5] "|" layout.Tops[6] "|" layout.Tops[7] "|" layout.Tops[23], "92|127|150|566")
        eq("layout: lines in the middle of the gaps", layout.Lines[1] "|" layout.Lines.Length, "121|5")
        eq("layout: fits", layout.Tops[23] + layout.RowH <= 591, true)
        layout := PreferencesWindow.PageLayout(2000, 23, [6])
        eq("layout: tall list, rows at most 28 and the gap at most 12", layout.RowH "|" layout.Gap, "28|12")
        layout := PreferencesWindow.PageLayout(740, 23, [6, 13, 17, 20, 22], 1.25)
        eq("layout: 125%", layout.RowH "|" layout.Gap "|" (layout.Tops[23] + layout.RowH <= 740), "29|14|1")
        names := ""                                                         ; every page opens by its name, and every page has a builder
        for key in PreferencesWindow.PageKeys
            names .= (PreferencesWindow.ParseArgs(["-Preferences", PreferencesWindow.PageName(key)]).Page = key && PreferencesWindow.PageBuilders.Has(key)) ? "" : key " "
        eq("each page by name, with a builder", names, "")
        eq("builders: one per page", PreferencesWindow.PageBuilders.Count, PreferencesWindow.PageKeys.Length)
        missing := ""
        for key in PreferencesWindow.GroupStarts
            missing .= PreferencesWindow._PageIndex(key) > 1 && PreferencesWindow.PageKeys[PreferencesWindow._PageIndex(key)] = key ? "" : key " "
        eq("group starts are pages", missing, "")
        eq("wiki general", PreferencesWindow.WikiPage("Prefs.Page.General"), "Configuration")
        eq("wiki window", PreferencesWindow.WikiPage("Prefs.Page.Window"), "Usage")
        eq("wiki keyboard", PreferencesWindow.WikiPage("Prefs.Page.Keyboard"), "Hotkeys")
        eq("wiki appearance", PreferencesWindow.WikiPage("Prefs.Page.Appearance"), "Themes")
        eq("wiki features", PreferencesWindow.WikiPage("Prefs.Page.Features"), "Usage")
        eq("wiki applications", PreferencesWindow.WikiPage("Prefs.Page.Applications"), "Configuration")
        eq("wiki file search", PreferencesWindow.WikiPage("Prefs.Page.FileSearch"), "File-Search")
        eq("wiki file index", PreferencesWindow.WikiPage("Prefs.Page.FileIndex"), "File-Search")
        eq("wiki commands", PreferencesWindow.WikiPage("Prefs.Page.Commands"), "Custom-Commands")
        eq("wiki selection", PreferencesWindow.WikiPage("Prefs.Page.Selection"), "Usage")
        eq("wiki snippets", PreferencesWindow.WikiPage("Prefs.Page.Snippets"), "Snippets")
        eq("wiki clipboard", PreferencesWindow.WikiPage("Prefs.Page.Clipboard"), "Clipboard")
        eq("wiki web search", PreferencesWindow.WikiPage("Prefs.Page.WebSearch"), "Web-Search")
        eq("wiki calculator", PreferencesWindow.WikiPage("Prefs.Page.Calculator"), "Calculator")
        eq("wiki scripts", PreferencesWindow.WikiPage("Prefs.Page.Scripts"), "Scripts")
        eq("wiki hotkeys", PreferencesWindow.WikiPage("Prefs.Page.Hotkeys"), "Hotkeys")
        eq("wiki quick switch", PreferencesWindow.WikiPage("Prefs.Page.QuickSwitch"), "Quick-Switch")
        eq("wiki dialog panel", PreferencesWindow.WikiPage("Prefs.Page.QSPanel"), "Quick-Switch-Panel")
        eq("wiki date stamp", PreferencesWindow.WikiPage("Prefs.Page.DateStamp"), "Date-Stamp")
        eq("wiki usage", PreferencesWindow.WikiPage("Prefs.Page.Usage"), "Usage")
        eq("wiki advanced default", PreferencesWindow.WikiPage("Prefs.Page.Advanced"), "Configuration")
        for key in ["Prefs.OK", "Prefs.Cancel", "Prefs.Apply", "Prefs.Help", "Prefs.DiscardChanges"]
            eq("text " key, I18n.T(key) != key, true)
    }

    ; Search window position: centered and 20% from the top by default; a remembered position is per mille of the screen, so it lands in the matching place on another screen
    ; Every page's controls are above the bottom buttons (checked in Chinese and English, whose text lengths differ)
    ; Every Preferences page: vertically adjacent controls (overlapping horizontally) are at least 4 pixels apart, never crammed together.
    ; Gray descriptions (one size smaller) sit right below their setting by design and don't count
    ; Tables in Preferences: the last column fills, the columns exactly span the list (minus a vertical scroll bar if any), no horizontal scroll bar
    static ListColumnsFit() {
        eq := (n, a, e) => TestRunner.Equal("ListColumnsFit." n, a, e)
        PreferencesWindow.Show(1, -3000, -3000)
        PreferencesWindow.BuildAll()
        lists := []
        for page in PreferencesWindow.Pages
            for ctrl in page.Controls
                if (ctrl.Type = "ListView" && SendMessage(0x108F, 0, 0, ctrl.Hwnd) = 1)   ; LVM_GETVIEW = LV_VIEW_DETAILS (theme thumbnails don't count)
                    lists.Push([page.Key, ctrl])
        TestRunner.True("ListColumnsFit.found the tables (" lists.Length ")", lists.Length >= 6)
        for entry in lists {
            listView := entry[2]
            columns := SendMessage(0x1200, 0, 0, SendMessage(0x101F, 0, 0, listView.Hwnd))
            total := 0
            Loop columns
                total += SendMessage(0x101D, A_Index - 1, 0, listView.Hwnd)
            eq(entry[1] " fills the width", total, PrefsForm.ListWidth(listView))
            style := DllCall("GetWindowLong", "Ptr", listView.Hwnd, "Int", -16, "Int")
            eq(entry[1] " no horizontal scroll bar", style & 0x100000 ? 1 : 0, 0)   ; WS_HSCROLL
        }
        listView := lists[1][2]                                             ; earlier columns dragged narrower: the last column fills up
        SendMessage(0x101E, 0, 33, listView.Hwnd)
        PrefsForm.FitColumns(listView)
        columns := SendMessage(0x1200, 0, 0, SendMessage(0x101F, 0, 0, listView.Hwnd)), total := 0
        Loop columns
            total += SendMessage(0x101D, A_Index - 1, 0, listView.Hwnd)
        eq("user width kept", SendMessage(0x101D, 0, 0, listView.Hwnd) "|" total, "33|" PrefsForm.ListWidth(listView))
        PreferencesWindow.Close()
    }

    static PreferencesSpacing() {
        smallFont(ctrl) {
            if !(ctrl.Type = "Text" || ctrl.Type = "Link") || !(font := SendMessage(0x31, 0, 0, ctrl.Hwnd))   ; WM_GETFONT
                return false
            logFont := Buffer(92, 0)
            DllCall("GetObject", "Ptr", font, "Int", 92, "Ptr", logFont)
            return Abs(NumGet(logFont, 0, "Int")) <= 11 * A_ScreenDPI / 96      ; s8 is about 11 pixels, s9 about 12
        }
        savedLang := I18n.Lang
        for lang in ["en", "zh-CN", "ja"] {
            I18n.Init(lang)
            PreferencesWindow.Show(1, -3000, -3000)
            PreferencesWindow.BuildAll()
            for index, page in PreferencesWindow.Pages {
                boxes := []
                for ctrl in page.Controls {
                    ctrl.GetPos(&x, &y, &w, &h)
                    if (h > 2)                                               ; the thin line below sections doesn't count
                        boxes.Push({Ctrl: ctrl, X: x, Y: y, W: w, H: h, Small: smallFont(ctrl)})
                }
                tight := ""
                for upper in boxes {
                    below := "", gap := 9999
                    for other in boxes {
                        if (other = upper || other.Y < upper.Y + upper.H - 2 || other.X >= upper.X + upper.W || upper.X >= other.X + other.W)
                            continue
                        if (other.Y - upper.Y - upper.H < gap)
                            below := other, gap := other.Y - upper.Y - upper.H
                    }
                    if (IsObject(below) && gap < 4 && !below.Small && !upper.Small)
                        tight .= upper.Ctrl.Type " '" SubStr(upper.Ctrl.Text, 1, 20) "' -> " below.Ctrl.Type " '" SubStr(below.Ctrl.Text, 1, 20) "' (" gap "); "
                }
                TestRunner.Equal("PreferencesSpacing." lang " page " index " " page.Key, tight, "")
            }
            PreferencesWindow._dirty := false
            PreferencesWindow.Close()
        }
        I18n.Init(savedLang)
    }

    static PreferencesFit() {
        savedLang := I18n.Lang
        for language in I18n.Languages() {                                   ; sorted by code, ja before zh-TW (Wine lacks both fonts, and Japanese measured too narrow after JhengHei was used first)
            lang := language[1]
            I18n.Init(lang)
            PreferencesWindow.Show(1, -3000, -3000)
            PreferencesWindow.BuildAll()                                         ; pages are built when needed: build them all before checking each page
            limit := PrefsForm.ButtonY - 6
            rightLimit := PrefsForm.ContentX + PrefsForm.ContentW + 3                   ; the thin line below sections (SS_ETCHEDHORZ) is 2 pixels wider
            for index, page in PreferencesWindow.Pages {
                if (page.Key = "Prefs.Page.Keyboard" && !AppSettings.Keyboard["CapsLock"]) {  ; Caps Lock as usual: "also in administrator windows" is grayed out, with the reason
                    hint := 0
                    for ctrl in page.Controls
                        hint := hint || (ctrl.Type = "Text" && ctrl.Text = I18n.T("Prefs.CapsLockElevated.NeedsMode"))
                    TestRunner.True("PreferencesFit." lang " caps lock helper hint", hint)
                    shieldCheck := "", shield := ""                       ; the UAC shield sits right after the checkbox text, not on top of the checkbox
                    for ctrl in page.Controls {
                        if (ctrl.Type = "CheckBox" && ctrl.Text = I18n.T("Prefs.CapsLockElevated"))
                            shieldCheck := ctrl
                        else if (ctrl.Type = "Pic")
                            shield := ctrl
                    }
                    if IsObject(shield) && IsObject(shieldCheck) {
                        shieldCheck.GetPos(&cx, &cy, &cw, &ch), shield.GetPos(&sx, &sy, , &sh)
                        TestRunner.True("PreferencesFit." lang " shield after the text (" cx + cw " <= " sx ")", cx + cw <= sx && sx <= cx + cw + 8)
                        TestRunner.True("PreferencesFit." lang " shield centered on the row", Abs((sy + sh / 2) - (cy + ch / 2)) <= 1)
                    }
                }
                if (page.Key = "Prefs.Page.Appearance") {                 ; "Visible results" and the column below it end on one line
                    rights := ""
                    for ctrl in page.Controls
                        if (ctrl.Type = "DDL" || ctrl.Type = "Button" && ctrl.Text = I18n.T("Prefs.ResetPosition"))
                            ctrl.GetPos(&x, , &w), rights .= (rights = "" ? "" : " ") (x + w)
                    parts := StrSplit(rights, " "), same := parts.Length = 4
                    for value in parts
                        same := same && value = parts[1]
                    TestRunner.True("PreferencesFit." lang " appearance right edges (" rights ")", same)
                }
                editButtons := [], rights := Map(), previous := ""         ; every list's "Edit..." (after its read-only summary) ends where something else does
                for ctrl in page.Controls {
                    ctrl.GetPos(&x, , &w)
                    if (ctrl.Type = "Button" && IsObject(previous) && previous.Type = "Edit" && (ControlGetStyle(previous) & 0x800))   ; ES_READONLY
                        editButtons.Push(x + w)
                    else
                        rights[x + w] := true
                    previous := ctrl
                }
                for right in editButtons
                    TestRunner.True("PreferencesFit." lang " " page.Key " Edit... aligned (" right ")", rights.Has(right))
                if (page.Key = "Prefs.Page.Snippets") {                   ; Search keyword / Keyword prefix over Paste by / Delay: four boxes alike, two columns
                    boxes := []
                    for ctrl in page.Controls {
                        if !((ctrl.Type = "Edit" && !(ControlGetStyle(ctrl) & 0x804)) || ctrl.Type = "DDL")   ; not read-only, not multi-line
                            continue
                        ctrl.GetPos(&x, &y, &w)
                        boxes.Push({X: x, Y: y, W: w})
                    }
                    grid := (boxes.Length >= 4) ? [boxes[boxes.Length - 3], boxes[boxes.Length - 2], boxes[boxes.Length - 1], boxes[boxes.Length]] : []
                    TestRunner.True("PreferencesFit." lang " snippets grid", grid.Length = 4 && grid[1].W = grid[2].W && grid[2].W = grid[3].W && grid[3].W = grid[4].W
                        && grid[1].X = grid[3].X && grid[2].X = grid[4].X && grid[1].Y = grid[2].Y && grid[3].Y = grid[4].Y)
                }
                if (page.Key = "Prefs.Page.FileSearch") {                 ; File types' "Edit..." ends where the right column of the pairs ends
                    lineEnd := 0, editRights := Map()
                    for ctrl in page.Controls {
                        ctrl.GetPos(&x, , &w)
                        if (ctrl.Type = "Button" && ctrl.Text = I18n.T("Prefs.EditList"))
                            lineEnd := x + w
                        else if (ctrl.Type = "Edit")
                            editRights[x + w] := (editRights.Has(x + w) ? editRights[x + w] : 0) + 1
                    }
                    TestRunner.True("PreferencesFit." lang " " page.Key " right edge (" lineEnd ")", lineEnd && editRights.Has(lineEnd) && editRights[lineEnd] >= 2)
                }
                if (page.Key = "Prefs.Page.DateStamp") {                  ; the date format box is as wide as the hotkey box
                    widths := []
                    for ctrl in page.Controls
                        if (ctrl.Type = "Edit" && widths.Length < 2)
                            ctrl.GetPos(, , &w), widths.Push(w)
                    TestRunner.Equal("PreferencesFit." lang " date stamp widths", widths[1] " " widths[2], widths[1] " " widths[1])
                }
                if (page.Key = "Prefs.Page.Window") {                     ; "Keep the last search for (minutes)" is a sub-option: indented to the check box's text
                    checkX := "", labelX := ""
                    for ctrl in page.Controls {
                        if (ctrl.Type = "CheckBox" && ctrl.Text = I18n.T("Prefs.KeepLastQuery"))
                            ctrl.GetPos(&checkX)
                        else if (ctrl.Type = "Text" && ctrl.Text = I18n.T("Prefs.KeepLastQueryMinutes"))
                            ctrl.GetPos(&labelX)
                    }
                    TestRunner.Equal("PreferencesFit." lang " keep minutes indented", labelX, checkX = "" ? "?" : checkX + PrefsForm.CheckIndent)
                }
                if (page.Key = "Prefs.Page.Window" || page.Key = "Prefs.Page.Keyboard") {   ; the input boxes / drop-downs of these pages share one left edge
                    lefts := "", count := 0
                    for ctrl in page.Controls
                        if (ctrl.Type = "Edit" || ctrl.Type = "DDL")
                            ctrl.GetPos(&x), lefts .= (lefts = "" ? "" : " ") x, count += 1
                    first := StrSplit(lefts, " ")[1], expected := ""
                    Loop count
                        expected .= (A_Index = 1 ? "" : " ") first
                    TestRunner.Equal("PreferencesFit." lang " " page.Key " controls aligned", lefts, expected)
                    TestRunner.Equal("PreferencesFit." lang " " page.Key " control count", count, page.Key = "Prefs.Page.Window" ? 3 : 4)
                }
                suffix := I18n.T("Prefs.DefaultChoice", "")
                for ctrl in page.Controls                                   ; every drop-down marks exactly one option "(default)" and shows its longest option in full
                    if (ctrl.Type = "DDL") {
                        marked := 0
                        for text in ControlGetItems(ctrl)
                            marked += (SubStr(text, -StrLen(suffix)) = suffix)
                        ctrl.GetPos(, , &w)
                        TestRunner.Equal("PreferencesFit." lang " " page.Key " default marked once (" ControlGetItems(ctrl)[1] ")", marked, 1)
                        TestRunner.True("PreferencesFit." lang " " page.Key " drop-down fits (" ControlGetItems(ctrl)[1] ", " w " < " PrefsForm.ChoiceWidth(ctrl) ")", w >= PrefsForm.ChoiceWidth(ctrl))
                    }
                listButtons := Map("Prefs.Page.Hotkeys", "Prefs.AddExamples", "Prefs.Page.WebSearch", "Prefs.AddBuiltins"   ; the list's extra button ends before its filter box
                    , "Prefs.Page.Commands", "Prefs.More", "Prefs.Page.Snippets", "Prefs.More")
                if listButtons.Has(page.Key) {
                    buttonRight := "", filterX := "", deleteRight := "", buttonLeft := "", buttonW := "", buttonStyle := 0
                    for ctrl in page.Controls {
                        ctrl.GetPos(&x, , &w)
                        if (ctrl.Type = "Button" && ctrl.Text = I18n.T("Prefs.Delete"))
                            deleteRight := x + w
                        if (ctrl.Type = "Button" && ctrl.Text = I18n.T(listButtons[page.Key]))
                            buttonRight := x + w, buttonLeft := x, buttonW := w, buttonStyle := ControlGetStyle(ctrl)
                        else if (ctrl.Type = "Edit" && w = 160 && buttonRight != "" && filterX = "")
                            filterX := x
                    }
                    TestRunner.True("PreferencesFit." lang " " page.Key " list button before the filter (" buttonRight " + 8 <= " filterX ")"
                        , buttonRight != "" && filterX != "" && buttonRight + 8 <= filterX)
                    TestRunner.Equal("PreferencesFit." lang " " page.Key " list button spaced like Delete", buttonLeft - deleteRight, 8)
                    if (listButtons[page.Key] = "Prefs.More")                ; the drop-down button: as wide as Add / Edit / Delete, an arrow drawn by Windows
                        TestRunner.Equal("PreferencesFit." lang " " page.Key " More button", buttonW "|" (buttonStyle & 0xF), "80|12")   ; BS_SPLITBUTTON
                    else
                        TestRunner.True("PreferencesFit." lang " " page.Key " list button at least 80 (" buttonW ")", buttonW >= 80)
                }
                bottom := 0, right := 0
                for ctrl in page.Controls {
                    ctrl.GetPos(&x, &y, &w, &h)
                    bottom := Max(bottom, y + h), right := Max(right, x + w)
                }
                TestRunner.True("PreferencesFit." lang " page " index " (bottom " bottom ", limit " limit ")", bottom <= limit)
                TestRunner.True("PreferencesFit." lang " page " index " (right " right ", limit " rightLimit ")", right <= rightLimit)
            }
            PreferencesWindow.Close()
            if (lang = "en" || lang = "zh-CN") {                            ; the left column widths for Chinese and English are tuned and need no widening
                widened := ""
                for key in PrefsForm.LabelWidths
                    if (SubStr(key, 1, StrLen(lang) + 1) = lang " ")
                        widened .= key "; "
                TestRunner.Equal("PreferencesFit." lang " no widened pages", widened, "")
            }
        }
        I18n.Init(savedLang)
    }

    ; Windows Settings pages: found by name (and English name), can be turned off as a group
    ; Preferences subpages in search: one per page ("ALTRun Preferences: Appearance"), found by page name, the main entry "ALTRun Preferences" before the subpages
    static PreferencePages() {
        eq := (n, a, e) => TestRunner.Equal("PreferencePages." n, a, e)
        best(text) {
            top := ""
            for item in SystemProvider.Search(SearchQuery(text))
                if (!IsObject(top) || item.Score > top.Score)
                    top := item
            return IsObject(top) ? top.Source["Id"] : ""
        }
        count := 0
        for command in SystemProvider.Commands()
            count += (SubStr(command["Id"], 1, 12) = "Preferences.")
        eq("one per page", count, PreferencesWindow.PageKeys.Length)
        eq("appearance", best("appearance"), "Preferences.Appearance")
        eq("hotkeys page", best("preferences hotkeys"), "Preferences.Hotkeys")
        eq("main item first", best("altrun pref"), "Preferences")
        for command in SystemProvider.Commands()
            if (command["Id"] = "Preferences.Calculator")
                eq("title", command["Title"], "ALTRun Preferences: Calculator")
    }

    static SettingsPages() {
        eq := (n, a, e) => TestRunner.Equal("SettingsPages." n, a, e)
        find(text) {
            for item in SystemProvider.Search(SearchQuery(text))
                if (item.Source["Id"] = "SetBluetooth")
                    return item
            return ""
        }
        item := find("bluetooth")
        eq("found", IsObject(item) ? item.Title "|" item.Subtitle : "", "Bluetooth & devices|Windows Settings")
        pages := 0
        for command in SystemProvider.Commands()
            if command.Has("IsSetting")
                pages += 1
        TestRunner.True("SettingsPages.count " pages, pages >= 40)
        options := AppSettings.Feature("System")
        options["SettingsPages"] := 0
        eq("turned off", IsObject(find("bluetooth")), false)
        options["SettingsPages"] := 1
        eq("default on", AppSettings.Defaults()["Features"]["System"]["SettingsPages"], 1)
    }

    ; Gray descriptions in Preferences: each "<label>.Desc" has its label, and neither Chinese nor English is empty
    static PreferenceDescriptions() {
        count := 0
        for key, english in I18n.Strings {
            if !(SubStr(key, -5) = ".Desc")
                continue
            count += 1
            base := SubStr(key, 1, -5)
            TestRunner.True("PreferenceDescriptions.label " base, I18n.Strings.Has(base))
            TestRunner.True("PreferenceDescriptions.text " key, Trim(english) != "")
        }
        TestRunner.True("PreferenceDescriptions.count " count, count >= 50)
        for key in ["Prefs.RememberPosition.Desc", "Prefs.ShowOn.Desc", "Prefs.HideOnDeactivate.Desc", "Prefs.Hotkey.Desc"]
            TestRunner.True("PreferenceDescriptions.has " key, I18n.Strings.Has(key))
        ; Descriptions with Wiki links: every language has <a href="{1}">, the URL opens the Chinese / English page by UI language; other descriptions must have no links (they would show as is)
        for language in I18n.Languages() {
            strings := (language[1] = "en") ? I18n.Strings : I18n.ReadFile(language[1])
            for key, text in strings
                if (SubStr(key, -5) = ".Desc")
                    TestRunner.True("PreferenceDescriptions.link " language[1] " " key
                        , PrefsForm.DescLinks.Has(SubStr(key, 1, -5)) ? InStr(text, '<a href="{1}">') : !InStr(text, "<a "))
        }
        TestRunner.True("PreferenceDescriptions.link url", InStr(PrefsForm.DescLink("Prefs.SelectionHotkey"), "/wiki/") && PrefsForm.DescLink("Prefs.Hotkey") = "")
    }

    ; Preferences → Advanced: links show names (Home · GitHub · Release notes · Report a problem), copyright, the data folder line (long paths with an ellipsis, the full path on hover)
    static AdvancedPage() {
        eq := (n, a, e) => TestRunner.Equal("AdvancedPage." n, a, e)
        PreferencesWindow.Show("Prefs.Page.Advanced", -3000, -3000)
        try {
            texts := "", link := "", folder := ""
            for ctrl in PreferencesWindow.Pages[PreferencesWindow._page].Controls {
                if (ctrl.Type = "Link")
                    link := ctrl
                else if (ctrl.Type = "Text" && ctrl.Text = AppSettings.DataDir)
                    folder := ctrl
                if (ctrl.Type = "Text")
                    texts .= ctrl.Text "`n"
            }
            eq("four links", StrSplit(link.Text, "<a ").Length - 1, 4)
            eq("no bare website address", InStr(link.Text, ">zhugecaomao.github.io"), 0)
            eq("issues link", InStr(link.Text, App.RepoUrl "/issues/new/choose") > 0, true)
            eq("copyright", InStr(texts, "© 2013–" SubStr(App.Version, 1, 4) " zhugecaomao") > 0, true)
            eq("settings file row removed", InStr(texts, AppSettings.File), 0)
            eq("data folder: one line, path ellipsis", IsObject(folder) && (ControlGetStyle(folder) & 0x8100) = 0x8100, true)
        } finally {
            PreferencesWindow.Close()
        }
        g := Gui()
        label := g.AddText("w100 0x100", "C:\A very long path\to a folder")
        tip := Win.AddTooltip(label.Hwnd, "C:\A very long path\to a folder")
        eq("tooltip added", DllCall("SendMessageW", "Ptr", tip, "UInt", 0x40D, "Ptr", 0, "Ptr", 0), 1)   ; TTM_GETTOOLCOUNT
        g.Destroy()
    }

    ; Preferences → Clipboard: settings that don't apply are grayed out (no saved history -> where / images; no images -> how many),
    ; the unit after "Max length" on its row, and the groups History / Images / Copy in that order
    static ClipboardPage() {
        eq := (n, a, e) => TestRunner.Equal("ClipboardPage." n, a, e)
        options := AppSettings.Feature("Clipboard")
        saved := [options["Persist"], options["Images"]]
        find := (controls, type, key) => PreferencesTests._Find(controls, type, I18n.T(key))
        try {
            for combo in [[1, 1, "1 1 1"], [1, 0, "1 1 0"], [0, 1, "0 0 0"]] {
                options["Persist"] := combo[1], options["Images"] := combo[2]
                PreferencesWindow.Show("Prefs.Page.Clipboard", -3000, -3000)
                try {
                    controls := PreferencesWindow.Pages[PreferencesWindow._page].Controls
                    storage := "", maxImages := "", images := find(controls, "CheckBox", "Prefs.ClipImages")
                    for index, ctrl in controls {
                        if (ctrl.Type = "DDL")
                            storage := ctrl
                        else if (ctrl.Type = "Text" && ctrl.Text = I18n.T("Prefs.ClipMaxImages"))
                            maxImages := controls[index + 1]                        ; the box right after its label
                    }
                    eq("enabled " combo[1] " " combo[2], storage.Enabled " " images.Enabled " " maxImages.Enabled, combo[3])
                    if (A_Index = 1) {
                        box := "", unit := find(controls, "Text", "Prefs.ClipMaxLength.Unit")
                        for index, ctrl in controls
                            if (ctrl.Type = "Text" && ctrl.Text = I18n.T("Prefs.ClipMaxLength"))
                                box := controls[index + 1]
                        box.GetPos(&bx, &by, &bw), unit.GetPos(&ux, &uy)
                        eq("unit after the box", ux - (bx + bw), 6)
                        eq("unit on the box's row", uy - by, 3)
                        order := ""
                        for ctrl in controls
                            for key in ["Prefs.Group.History", "Prefs.ClipStorage", "Prefs.Group.Images", "Prefs.ClipMaxImages", "Prefs.ClipFillThumbs", "Prefs.Group.Copy"]
                                if (ctrl.Text = I18n.T(key))
                                    order .= key " "
                        eq("groups", order, "Prefs.Group.History Prefs.ClipStorage Prefs.Group.Images Prefs.ClipMaxImages Prefs.ClipFillThumbs Prefs.Group.Copy ")
                    }
                } finally {
                    PreferencesWindow.Close()
                }
            }
        } finally {
            options["Persist"] := saved[1], options["Images"] := saved[2]
        }
    }

    ; The first control of a page with this type and text ("" if none)
    static _Find(controls, type, text) {
        for ctrl in controls
            if (ctrl.Type = type && ctrl.Text = text)
                return ctrl
        return ""
    }

    ; Preferences pages are built when needed: on opening only the page shown exists; switching builds the page, with Tab order list -> page -> bottom buttons;
    ; values the program fills in while building don't count as changes, later changes are recorded as usual
    ; Preferences → Selection: the hotkey, the content types with the commands that use them, and the test box (#434)
    static PreferencesSelectionPage() {
        eq := (n, a, e) => TestRunner.Equal("PreferencesSelectionPage." n, a, e)
        savedTypes := AppSettings.Data.Has("SelectionTypes") ? AppSettings.Data["SelectionTypes"] : []
        savedCommands := AppSettings.Data["CustomCommands"]
        AppSettings.Data["SelectionTypes"] := [Map("Name", "PO", "Kind", "Text", "Pattern", "^PO\d{8}$"), Map("Name", "PDF", "Kind", "File", "Pattern", "pdf")]
        AppSettings.Data["CustomCommands"] := [Map("Title", "ERP", "Type", "Url", "Target", "https://erp.example.com/{query}", "Arguments", "", "Keyword", "", "Selection", "PO")
            , Map("Title", "WMS", "Type", "Url", "Target", "https://wms.example.com/{query}", "Arguments", "", "Keyword", "", "Selection", "po, Text")]
        PreferencesWindow.Show("Prefs.Page.Selection", -3000, -3000)
        try {
            index := PreferencesWindow._PageIndex("Prefs.Page.Selection")
            eq("page in the commands and triggers group (Commands, Scripts, Selection)", index, PreferencesWindow._PageIndex("Prefs.Page.Commands") + 2)
            ui := PrefsPages._selection
            eq("types listed", ui.List.View.GetCount(), 2)
            eq("kind shown", ui.List.View.GetText(2, 2), "File (extensions)")
            eq("commands of a type", ui.List.View.GetText(1, 4), "ERP, WMS")
            eq("unused type", ui.List.View.GetText(2, 4), "")
            hotkeyHere := false
            for ctrl in PreferencesWindow.Pages[index].Controls
                hotkeyHere := hotkeyHere || (ctrl.Type = "Text" && ctrl.Text = "Selection hotkey")
            eq("hotkey on this page", hotkeyHere, true)
            ui.Sample.Value := "PO12345678"
            PreferencesWindow.Pages[index].OnShow.Call()
            eq("test: types and commands", ui.Result.Value, "Types: Text, PO`nCommands: ERP, WMS")
            ui.Sample.Value := "C:\Docs\report.PDF"
            PreferencesWindow.Pages[index].OnShow.Call()
            eq("test: a path is a file", ui.Result.Value, "Types: File, PDF`nCommands: none")
            ui.Sample.Value := "hello"
            PreferencesWindow.Pages[index].OnShow.Call()
            eq("test: no own type", ui.Result.Value, "Types: Text (none of your types matches)`nCommands: WMS")
            ui.Sample.Value := ""
            PreferencesWindow.Pages[index].OnShow.Call()
            eq("test: empty", ui.Result.Value, "")
            ui.Sample.GetPos(&sampleX), ui.Browse.GetPos(&browseX)
            eq("test: browse button after the box", browseX > sampleX, true)
            hotkeyX := testX := hotkeyW := testW := 0
            for ctrl in PreferencesWindow.Pages[index].Controls {
                if (ctrl.Type = "Text" && ctrl.Text = "Selection hotkey")
                    ctrl.GetPos(&hotkeyX, , &hotkeyW)
                if (ctrl.Type = "Text" && ctrl.Text = "Test")
                    ctrl.GetPos(&testX, , &testW)
            }
            eq("test label under the hotkey label", testX + testW = hotkeyX + hotkeyW && hotkeyW > 0, true)   ; right-aligned labels: the same right edge
            generalHotkeyX := generalHotkeyW := 0
            eq("testing is not a change", PreferencesWindow._dirty, false)
            ; a type added on this page (not saved yet) is used by the test box at once
            ui.List.Items.Push(Map("Name", "Mail", "Kind", "Text", "Pattern", "@"))
            ui.Sample.Value := "a@b.com"
            PreferencesWindow.Pages[index].OnShow.Call()
            eq("test: unsaved type", ui.Result.Value, "Types: Text, Mail`nCommands: WMS")
            general := PreferencesWindow._PageIndex("Prefs.Page.General")
            PreferencesWindow._SelectRow(general)
            onGeneral := false
            for ctrl in PreferencesWindow.Pages[general].Controls {
                onGeneral := onGeneral || (ctrl.Type = "Text" && ctrl.Text = "Selection hotkey")
                if (ctrl.Type = "Text" && ctrl.Text = "ALTRun hotkey")
                    ctrl.GetPos(&generalHotkeyX, , &generalHotkeyW)
            }
            eq("hotkey no longer on General", onGeneral, false)
            eq("same label column as General", hotkeyX + hotkeyW, generalHotkeyX + generalHotkeyW)
            ; the page list: every page visible, and only a little space below the last row (with the icon font, i.e. on Windows)
            pageList := PreferencesWindow.PageList
            eq("page list: no scrolling", SendMessage(0x1028, 0, 0, pageList) >= pageList.GetCount(), true)   ; LVM_GETCOUNTPERPAGE
            if (NavIcons.FontName() != "") {
                rect := Buffer(16, 0), client := Buffer(16, 0)
                SendMessage(0x100E, pageList.GetCount() - 1, rect, pageList)   ; LVM_GETITEMRECT of the last row
                DllCall("GetClientRect", "Ptr", pageList.Hwnd, "Ptr", client)
                rowH := NumGet(rect, 12, "Int") - NumGet(rect, 4, "Int")
                eq("page list: little space below", NumGet(client, 12, "Int") - NumGet(rect, 12, "Int") < rowH, true)
            }
        } finally {
            AppSettings.Data["SelectionTypes"] := savedTypes
            AppSettings.Data["CustomCommands"] := savedCommands
            PreferencesWindow.Close()
        }
    }

    ; Preferences → Custom Commands: placeholders that won't work are in the Status column without checking paths; a hotkey set
    ; meanwhile (on the Hotkeys page) clears "{query} empty" when the page is shown again
    static PreferencesCommandsStatus() {
        eq := (n, a, e) => TestRunner.Equal("PreferencesCommandsStatus." n, a, e)
        savedCommands := AppSettings.Data["CustomCommands"]
        hadHotkeys := AppSettings.Data.Has("Hotkeys"), savedHotkeys := hadHotkeys ? AppSettings.Data["Hotkeys"] : ""
        AppSettings.Data["CustomCommands"] := [Map("Id", "pa7k2m4x", "Title", "Notepad", "Type", "File", "Target", "notepad.exe", "Arguments", "", "Keyword", "")
            , Map("Id", "pb7k2m4x", "Title", "Ping", "Type", "Command", "Target", "cmd.exe", "Arguments", "/k ping {query}", "Keyword", "")
            , Map("Id", "pc7k2m4x", "Title", "Echo", "Type", "Command", "Target", "cmd.exe", "Arguments", "/k echo {selection}", "Keyword", "e")
            , Map("Id", "pd7k2m4x", "Title", "Today", "Type", "Url", "Target", "https://example.com/{date}", "Arguments", "", "Keyword", "")]
        AppSettings.Data["Hotkeys"] := []
        PreferencesWindow.Show("Prefs.Page.Commands", -3000, -3000)
        try {
            list := PreferencesWindow._lists["CustomCommands"]
            statuses() {
                text := ""
                Loop list.View.GetCount()
                    text .= list.View.GetText(A_Index, 5) "|"
                return text
            }
            eq("shown without checking", statuses(), "|{query} empty|Use {query}|Unknown {date}|")
            fits := true                                                    ; the Status column is wide enough for its texts (LVM_GETSTRINGWIDTHW, plus the cell's margins)
            columnW := SendMessage(0x101D, 4, 0, list.View.Hwnd)            ; LVM_GETCOLUMNWIDTH
            for text in [I18n.T("Prefs.Status.QueryEmpty"), I18n.T("Prefs.Status.UseQuery"), I18n.T("Prefs.Status.Unknown", "{date}"), I18n.T("Prefs.Status.Missing")]
                fits := fits && SendMessage(0x1057, 0, StrPtr(text), list.View.Hwnd) + 12 <= columnW
            eq("status texts fit", fits, true)
            PreferencesWindow.GetPath(PreferencesWindow.Working, "Hotkeys").Push(Map("Key", "^!p", "Action", "custom:pb7k2m4x", "SearchText", "", "WinTitle", ""))
            PreferencesWindow.Pages[PreferencesWindow._PageIndex("Prefs.Page.Commands")].OnShow.Call()
            eq("a hotkey gives {query} its text", statuses(), "||Use {query}|Unknown {date}|")
            eq("not a change", PreferencesWindow._dirty, false)
        } finally {
            AppSettings.Data["CustomCommands"] := savedCommands
            if hadHotkeys
                AppSettings.Data["Hotkeys"] := savedHotkeys
            else
                AppSettings.Data.Delete("Hotkeys")
            PreferencesWindow.Close()
        }
    }

    ; Preferences → Hotkeys: "Other hotkeys" lists the hotkeys set on other pages, fits above the bottom buttons; double-click goes to that page
    static PreferencesOtherHotkeys() {
        eq := (n, a, e) => TestRunner.Equal("PreferencesOtherHotkeys." n, a, e)
        PreferencesWindow.Show("Prefs.Page.Hotkeys", -3000, -3000)
        try {
            ui := PrefsPages._others
            expected := PreferencesWindow.OtherHotkeys(PreferencesWindow.Working)
            eq("a row for each", ui.View.GetCount(), expected.Length)
            eq("first row: keys, use, page", ui.View.GetText(1, 1) "|" ui.View.GetText(1, 2) "|" ui.View.GetText(1, 3)
                , HotkeyBox.Label(expected[1].Key) "|" expected[1].Purpose "|" I18n.T("Prefs.Page.General"))
            ui.View.GetPos(, &y, , &h)
            eq("above the bottom buttons", y + h <= PrefsForm.ButtonY - 10, true)
            eq("room for all seven without scrolling", SendMessage(0x1028, 0, 0, ui.View) >= 7, true)   ; LVM_GETCOUNTPERPAGE
            selection := 0
            for index, each in ui.Items
                if (each.Page = "Prefs.Page.Selection")
                    selection := index
            eq("the selection hotkey is listed", selection > 0, true)
            PrefsPages.GoToOtherHotkey(selection)
            eq("double-click goes to its page", PreferencesWindow.Pages[PreferencesWindow._page].Key, "Prefs.Page.Selection")
            eq("page list follows", PreferencesWindow.PageList.GetNext(), PreferencesWindow._page)
            PrefsPages.GoToOtherHotkey(99)                                  ; no such row: nothing happens
            eq("a row that isn't there", PreferencesWindow.Pages[PreferencesWindow._page].Key, "Prefs.Page.Selection")
        } finally {
            PreferencesWindow.Close()
        }
    }

    static PreferencesLazyPages() {
        eq := (n, a, e) => TestRunner.Equal("PreferencesLazyPages." n, a, e)
        PreferencesWindow.Show("Prefs.Page.Appearance", -3000, -3000)
        try {
            built := ""
            for page in PreferencesWindow.Pages
                built .= page.Built ? page.Key " " : ""
            eq("only the shown page", Trim(built), "Prefs.Page.Appearance")
            eq("page list has every page", PreferencesWindow.PageList.GetCount(), PreferencesWindow.PageKeys.Length)
            Sleep(400)
            commands := PreferencesWindow._PageIndex("Prefs.Page.Commands")
            PreferencesWindow._SelectRow(commands)
            eq("built when selected", PreferencesWindow.Pages[commands].Built, true)
            eq("building is not a change", PreferencesWindow._dirty, false)
            eq("ready again", PreferencesWindow._ready, true)
            eq("open: window handle", PreferencesWindow.OpenHwnd() = PreferencesWindow.Gui.Hwnd, true)
            first := ""
            for ctrl in PreferencesWindow.Pages[commands].Controls
                if (!first && ctrl.Visible && (DllCall("GetWindowLong", "Ptr", ctrl.Hwnd, "Int", -16, "Int") & 0x10000))   ; WS_TABSTOP
                    first := ctrl.Hwnd
            next := DllCall("GetNextDlgTabItem", "Ptr", PreferencesWindow.Gui.Hwnd, "Ptr", PreferencesWindow.PageList.Hwnd, "Int", 0, "Ptr")
            eq("tab from the list goes into the page", next = first && next != PreferencesWindow._helpButton.Hwnd, true)
            general := PreferencesWindow._PageIndex("Prefs.Page.General")
            PreferencesWindow._SelectRow(general)
            for ctrl in PreferencesWindow.Pages[general].Controls
                if (ctrl.Type = "CheckBox") {
                    ctrl.Value := !ctrl.Value
                    ControlClick(ctrl)                                      ; the user clicks once (Click event)
                    break
                }
            Sleep(100)
            eq("user change after building is recorded", PreferencesWindow._dirty, true)
            built := 0
            for page in PreferencesWindow.Pages
                built += page.Built
            eq("unvisited pages still not built", built, 3)
            ; the window is closed while a page is being built (another thread interrupted the build): no controls added to a window that is gone, no error
            usagePage := PreferencesWindow._PageIndex("Prefs.Page.Usage")
            PreferencesWindow.Pages[usagePage].Builder := "_TestCloseWhileBuilding"
            PrefsPages.DefineProp("_TestCloseWhileBuilding", {Call: (this) => (PreferencesWindow.Close(), PrefsForm.Add("Text", "x200 w100", "x"))})
            closedOk := true
            try PreferencesWindow._SelectRow(usagePage)
            catch
                closedOk := false
            PrefsPages.DeleteProp("_TestCloseWhileBuilding")
            eq("closed while building: no error", closedOk, true)
            eq("closed while building: window gone", IsObject(PreferencesWindow.Gui), false)
            eq("closed: no window handle", PreferencesWindow.OpenHwnd(), 0)
            ; a rebuild result arriving while Close destroys the window (#290): Gui is still set but the window is gone
            destroyed := Gui()
            destroyed.Destroy()
            PreferencesWindow.Gui := destroyed
            eq("destroyed window: no handle", PreferencesWindow.OpenHwnd(), 0)
            PreferencesWindow.Gui := ""
        } finally {
            PreferencesWindow._dirty := false
            PreferencesWindow.Close()
        }
    }

    ; Hotkey display (Alt+Space) and the rules of the recording box; duplicate global hotkeys checked before Preferences saves
    static HotkeyText() {
        eq := (n, a, e) => TestRunner.Equal("HotkeyText." n, a, e)
        ok := (n, c) => TestRunner.True("HotkeyText." n, c)
        for pair in [["!Space", "Alt+Space"], ["^!c", "Ctrl+Alt+C"], ["!^c", "Ctrl+Alt+C"], ["#e", "Win+E"], ["^+F12", "Ctrl+Shift+F12"]
                   , ["~MButton", "MButton"], ["$*F5", "F5"], ["CapsLock & j", "CapsLock+J"], ["^+", "Ctrl++"], ["!r", "Alt+R"]
                   , ["<^>!a", "AltGr+A"], ["F1 up", "F1 Up"], ["", ""], ["None", ""]
                   , ["<!Space", "LAlt+Space"], [">^>+F1", "RCtrl+RShift+F1"], ["<#e", "LWin+E"], ["^>!x", "Ctrl+RAlt+X"]]
            eq("label " pair[1], Win.HotkeyLabel(pair[1]), pair[2])
        eq("names", Win.HotkeyLabel("~^MButton", Map("MButton", "Middle")), "Ctrl+Middle")
        eq("box label", HotkeyBox.Label("~MButton"), "Middle mouse button")
        eq("box label XButton1", HotkeyBox.Label("^XButton1"), "Ctrl+Mouse back button")

        eq("compose order", HotkeyBox.Compose("#+!^", "C"), "^!+#c")
        eq("compose name", HotkeyBox.Compose("!", "Space"), "!Space")
        eq("compose none", HotkeyBox.Compose("", "F5"), "F5")
        eq("compose sides", HotkeyBox.Compose("#>+<!^", "C"), "^<!>+#c")
        ; while recording, modifiers count by the downs / ups InputHook received (keys simulated by a remote desktop aren't physically down)
        savedActive := HotkeyBox._active
        HotkeyBox._active := {Hwnd: 0}
        HotkeyBox._held := Map("RControl", true, "LAlt", true)
        eq("held mods", HotkeyBox._Mods(), "^!")
        HotkeyBox._held := Map("Control", true, "LWin", true, "RShift", true)
        eq("held generic Control", HotkeyBox._Mods(), "^+#")
        HotkeyBox._held := Map()
        eq("nothing held", HotkeyBox._Mods(), "")
        ; with Sides (#303): a modifier held on one side only keeps that side; both sides, or one without a side, mean either side
        savedSides := HotkeyBox.Sides, HotkeyBox.Sides := true
        HotkeyBox._held := Map("RControl", true, "LAlt", true)
        eq("sides: held mods", HotkeyBox._Mods(), ">^<!")
        HotkeyBox._held := Map("LShift", true, "RShift", true, "Control", true, "RWin", true)
        eq("sides: both sides or none", HotkeyBox._Mods(), "^+>#")
        eq("sides: recorded", HotkeyBox.Compose(">^<!", "k"), ">^<!k")
        HotkeyBox.Sides := savedSides, HotkeyBox._held := Map()
        HotkeyBox._active := savedActive
        ok("modifier names", HotkeyBox.IsModifier("LControl") && HotkeyBox.IsModifier("RWin") && HotkeyBox.IsModifier("Shift") && !HotkeyBox.IsModifier("k"))
        for hk in ["^!c", "!Space", "#e", "F5", "^Numpad1", "+F3", "Pause", "MButton", "!Enter", "<!Space", ">^>+k"]
            ok("allowed " hk, HotkeyBox.IsAllowed(hk))
        for hk in ["a", "+a", "Space", "+Space", "Enter", "Tab", "Numpad5", "1", "Delete", "<+a", ">+Space"]
            ok("not allowed " hk, !HotkeyBox.IsAllowed(hk))
        ; boxes of closed windows must go: handles are reused, and a middle click on another control would count as recording
        g := Gui()
        box := HotkeyBox.Add(g, "w120", "^!k", true)
        boxHwnd := box.Hwnd, state := HotkeyBox.Boxes[boxHwnd]
        ok("box alive", HotkeyBox.IsAlive(state))
        g.Destroy()
        ok("box destroyed", !HotkeyBox.IsAlive(state))
        eq("mouse on destroyed box", HotkeyBox._OnMouse(0, 0x207, boxHwnd), "")
        ok("destroyed box pruned", !HotkeyBox.Boxes.Has(boxHwnd))
        g := Gui()
        HotkeyBox.Add(g, "w120", "", false)
        HotkeyBox.Boxes[boxHwnd] := state                                   ; also cleaned up on the next Add
        HotkeyBox.Add(g, "w120", "", false)
        ok("pruned on add", !HotkeyBox.Boxes.Has(boxHwnd) || HotkeyBox.Boxes[boxHwnd] != state)
        g.Destroy()
        HotkeyBox.Prune()

        ; duplicate global hotkeys in Preferences
        list := [{Key: "!Space", Label: "A"}, {Key: "!r", Label: "B"}, {Key: "~!R", Label: "C"}]
        clash := PreferencesWindow.DuplicateHotkey(list)
        eq("duplicate", IsObject(clash) ? clash.First "|" clash.Second : "", "B|C")
        eq("duplicate order", PreferencesWindow.DuplicateHotkey([{Key: "^!c", Label: "A"}, {Key: "!^c", Label: "B"}]).Second, "B")
        eq("no duplicate", PreferencesWindow.DuplicateHotkey([{Key: "!Space", Label: "A"}, {Key: "!r", Label: "B"}]), "")
        ; left / right (#303): either side clashes with one side, the left and the right one don't
        eq("either side clashes with the left", PreferencesWindow.DuplicateHotkey([{Key: "!Space", Label: "A"}, {Key: "<!Space", Label: "B"}]).Second, "B")
        eq("left and right don't clash", PreferencesWindow.DuplicateHotkey([{Key: "<!Space", Label: "A"}, {Key: ">!Space", Label: "B"}]), "")
        eq("sides in any order", PreferencesWindow.DuplicateHotkey([{Key: "<^>!k", Label: "A"}, {Key: "~>!<^K", Label: "B"}]).Second, "B")
        eq("more modifiers aren't the same press", PreferencesWindow.DuplicateHotkey([{Key: "<!Space", Label: "A"}, {Key: "^<!Space", Label: "B"}]), "")
        data := AppSettings.Defaults()
        data["Hotkeys"] := [Map("Key", "^!c", "Action", "Lock", "WinTitle", ""), Map("Key", "!Space", "Action", "Lock", "WinTitle", "ahk_exe notepad.exe")]
        keys := ""
        for entry in PreferencesWindow._GlobalHotkeys(data)
            keys .= entry.Key "|"
        eq("global hotkeys (window-limited left out)", keys, "!Space|!r|^!\|^!c|^!c|")
        ok("default clipboard hotkey clashes with a global ^!c", IsObject(PreferencesWindow.DuplicateHotkey(PreferencesWindow._GlobalHotkeys(data))))
        ok("defaults have no clash", !IsObject(PreferencesWindow.DuplicateHotkey(PreferencesWindow._GlobalHotkeys(AppSettings.Defaults()))))

        ; the Exclude boxes must be valid regular expressions before saving
        data := AppSettings.Defaults()
        eq("default patterns are valid", PreferencesWindow.BadPattern(data), "")
        data["Features"]["FileSearch"]["ScopeExclude"] := "i)\\(node_modules"
        bad := PreferencesWindow.BadPattern(data)
        eq("bad scope exclude: page", IsObject(bad) ? PreferencesWindow.PageKeys[bad.Page] : "", "Prefs.Page.FileIndex")
        ok("bad scope exclude: error given", IsObject(bad) && bad.Error != "")
        data["Features"]["Applications"]["Exclude"] := "(readme"
        eq("first bad pattern: Applications", PreferencesWindow.PageKeys[PreferencesWindow.BadPattern(data).Page], "Prefs.Page.Applications")
        eq("list cell", PrefsForm.Cell(Map("Key", "~MButton"), "Key"), "Middle mouse button")
    }

    ; After recording a hotkey: focus moves to the next control without selecting its text (with a window list below, the next typing would replace it all);
    ; while recording, the tip below the box follows the window
    static HotkeyBoxFocus() {
        eq := (n, a, e) => TestRunner.Equal("HotkeyBoxFocus." n, a, e)
        g := Gui(, "ALTRun hotkey box test")
        box := HotkeyBox.Add(g, "w160", "^!k")
        list := g.AddEdit("w300 r3 +Multi", "ahk_exe a.exe`r`nahk_exe b.exe")
        g.Show("x100 y100")
        try {
            WinActivate("ahk_id " g.Hwnd)
            if !WinWaitActive("ahk_id " g.Hwnd, , 2) {
                FileAppend("NOTE HotkeyBoxFocus: test window not active, skipped`n", "*")
                return
            }
            state := HotkeyBox.Boxes[box.Hwnd]
            box.Focus()                                                     ; clicking into the box: recording starts
            Sleep(50)
            eq("recording", HotkeyBox._active == state, true)
            eq("tip shown", state.Tip != "", true)
            before := state.TipAt
            WinMove(300, 250, , , "ahk_id " g.Hwnd)
            HotkeyBox._Watch()
            eq("tip follows the window", state.TipAt != before, true)
            HotkeyBox._Finish(state, "^d")
            Sleep(50)
            eq("value", HotkeyBox.Value(box), "^d")
            eq("hotkeys back on", A_IsSuspended, 0)
            eq("focus on the next control", DllCall("GetFocus", "Ptr") = list.Hwnd, true)
            start := 0, end := 0
            DllCall("SendMessage", "Ptr", list.Hwnd, "UInt", 0xB0, "UInt*", &start, "UInt*", &end)   ; EM_GETSEL
            eq("next control's text not selected", end - start, 0)
            eq("tip cleared", state.Tip, "")
        } finally {
            HotkeyBox.CancelActive()
            g.Destroy()
            HotkeyBox.Prune()
        }
    }

    static SnippetEditor() {
        eq := (n, a, e) => TestRunner.Equal("SnippetEditor." n, a, e)
        ok := (n, c) => TestRunner.True("SnippetEditor." n, c)
        for field in SnippetProvider.EditorFields()
            ok("hint: " field.Key, field.HasOwnProp("Hint") && field.Hint != "")
        keyword := SnippetProvider.EditorFields()[2]
        ok("keyword has check", keyword.HasOwnProp("Check"))
        ok("check as ItemEditor calls it", keyword.Check.Call(StrRepeat("a", 40), Map("AutoExpand", 1)) != "")
        eq("check short keyword", keyword.Check.Call("sig", Map("AutoExpand", 1)), "")
        ok("hint shows limit", InStr(keyword.Hint, SnippetExpander.MaxAbbreviation - StrLen(AppSettings.Feature("Snippets")["ExpandPrefix"])))
        longKeyword := StrRepeat("a", 40)
        eq("short ok", SnippetProvider.CheckKeyword("sig", Map("AutoExpand", 1), ";"), "")
        eq("empty ok", SnippetProvider.CheckKeyword("", Map("AutoExpand", 1), ";"), "")
        ok("too long", SnippetProvider.CheckKeyword(longKeyword, Map("AutoExpand", 1), ";") != "")
        eq("39 + prefix ok", SnippetProvider.CheckKeyword(StrRepeat("a", 39), Map("AutoExpand", 1), ";"), "")
        ok("space", SnippetProvider.CheckKeyword("a b", Map("AutoExpand", 1), ";") != "")
        eq("no auto-expand", SnippetProvider.CheckKeyword(longKeyword, Map("AutoExpand", 0), ";"), "")
        eq("too long not registered", SnippetExpander.Abbreviations(Map("Keyword", longKeyword, "Text", "x"), ";").Length, 0)
        eq("longest registered", TextTools.JoinList(SnippetExpander.Abbreviations(Map("Keyword", StrRepeat("a", 39), "Text", "x"), ";"), ","), ";" StrRepeat("a", 39))
        eq("several keywords ok", SnippetProvider.CheckKeyword("sig, ㄴ", Map("AutoExpand", 1), ";"), "")
        ok("several: one too long", SnippetProvider.CheckKeyword("sig, " longKeyword, Map("AutoExpand", 1), ";") != "")
        store := ResultItem("Notepad", "", {Kind: "file", Arg: "shell:AppsFolder\Microsoft.WindowsNotepad_8wekyb3d8bbwe!App"})
        list := ""
        for action in ResultActions.ListFor(store)
            list .= action.Title "|"
        ok("store app: no properties", !InStr(list, I18n.T("Action.Properties") "|"))
        list := ""
        for action in ResultActions.ListFor(ResultItem("Windows", "", {Kind: "folder", Arg: A_WinDir}))
            list .= action.Title "|"
        ok("folder: properties", InStr(list, I18n.T("Action.Properties") "|"))
    }
}
