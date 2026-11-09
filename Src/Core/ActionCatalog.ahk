;===============================================================================
; ActionCatalog.ahk - what running a result does, and the operations behind it: open / reveal / copy / paste ...
;   (AutoHotkey v2)
;-------------------------------------------------------------------------------
; As in Alfred:
;   Enter        the default action (file / program / folder / URL: open; text: copy)
;   Ctrl+Enter   file / folder: reveal in the file manager; text: paste into the foreground window
;   Alt+Enter    copy (path / URL / text)
;   Ctrl+Shift+Enter   run as administrator (programs, shortcuts, scripts, Store apps that have it); otherwise what Enter does
; The actions panel (→), F3 edit and Ctrl+Del delete are built on these in Src\UI\ResultActions.ahk.
;
; Usage:
;   ActionCatalog.RunDefault(item)
;   ActionCatalog.RunModifier(item, "ctrl" | "alt" | "admin")
;   ActionCatalog.OpenFile(path) / OpenFolder(path) / Reveal(path) / CopyText(text) / PasteText(text) ...
;   ActionCatalog.OpenMany(paths) / CopyMany(paths, folder) / RecycleMany(paths)   several marked files
;===============================================================================

class ActionCatalog {
    static RunDefault(item) {
        if IsObject(item.OnRun)
            return item.OnRun.Call(item)
        switch item.Kind {
            case "file"  : ActionCatalog.OpenFile(item.Arg, item.Arguments)
            case "folder": ActionCatalog.OpenFolder(item.Arg)
            case "url"   : ActionCatalog.OpenUrl(item.Arg)
            case "text"  : ActionCatalog.CopyText(item.Arg)
        }
    }

    static RunModifier(item, modifier) {
        if (modifier = "admin")
            return ActionCatalog.CanRunElevated(item) ? ActionCatalog.RunAsAdmin(item.Arg, item.Arguments) : ActionCatalog.RunDefault(item)
        if (modifier = "alt")
            return ActionCatalog.CopyText(item.CopyText())
        if (modifier = "ctrl") {
            switch item.Kind {
                case "file", "folder": return ActionCatalog.Reveal(item.Arg)
                case "text"          : return ActionCatalog.PasteText(item.Arg)
                case "url"           : return ActionCatalog.CopyText(item.Arg)
            }
        }
        ActionCatalog.RunDefault(item)
    }

    static OpenMany(paths) {
        for markedPath in paths
            InStr(FileExist(markedPath), "D") ? ActionCatalog.OpenFolder(markedPath) : ActionCatalog.OpenFile(markedPath)
    }

    ; Copy / move into a folder (Windows' own copy, with progress and name-conflict prompts); not into itself or its own subfolders
    static CopyMany(paths, destination, move := false) {
        folder := ComObject("Shell.Application").NameSpace(RTrim(destination, "\") (RegExMatch(destination, "^[A-Za-z]:\\?$") ? "\" : ""))
        if !IsObject(folder)
            return false
        done := 0
        for markedPath in paths {
            if (InStr(RTrim(destination, "\") "\", RTrim(markedPath, "\") "\") = 1)
                continue
            move ? folder.MoveHere(markedPath, 0) : folder.CopyHere(markedPath, 0)
            done += 1
        }
        App.Notify(I18n.T(move ? "Action.MovingMany" : "Action.CopyingMany", done, destination), 2000)
        return done > 0
    }

    ; Move to the Recycle Bin together: confirm once first (they can be restored from the Recycle Bin)
    static RecycleMany(paths) {
        if (MsgBox(I18n.T("Action.ConfirmRecycleAll", paths.Length), App.Name, "YesNo Icon? Default2") != "Yes")
            return false
        failed := 0
        for markedPath in paths
            try FileRecycle(markedPath)
            catch
                failed += 1
        if failed
            App.Notify(I18n.T("Action.RecycleFailed", failed))
        return !failed
    }

    ;---------------------------------------------------------------------------
    ; Operations
    ;---------------------------------------------------------------------------
    static OpenFile(target, arguments := "") {
        resolved := Path.Resolve(target)
        workDir := ActionCatalog.ParentDir(resolved)
        if RegExMatch(resolved, "i)^(shell:|::\{|[a-z]+:\/\/)") || !FileExist(resolved)
            return Run(Trim(resolved " " arguments))                        ; shell: paths / protocols / program names in PATH
        Run(Trim('"' resolved '" ' arguments), DirExist(workDir) ? workDir : "")
    }

    ; Store / packaged apps (shell:AppsFolder\<app id>) through their context menu's "Run as administrator", as the Start menu does
    static RunAsAdmin(target, arguments := "") {
        if ActionCatalog.IsPackagedApp(target) {
            result := ShellContextMenu([target]).InvokeVerb("runas")
            if (result < 0 && result != -2147023673)                       ; 0x800704C7: UAC canceled
                throw OSError("Run as administrator failed (0x" Format("{:08X}", result & 0xFFFFFFFF) ")", -1)
            return
        }
        Run(Trim('*RunAs "' Path.Resolve(target) '" ' arguments))
    }

    ; Whether Ctrl+Shift+Enter runs item as administrator: a file result (an app, a file found, a custom command for a program) whose
    ; target fits "Run as administrator" (CanRunAsAdmin). A result with its own action (OnRun: snippets, system commands...) runs that
    static CanRunElevated(item) => (item.Kind = "file" && !IsObject(item.OnRun) && ActionCatalog.CanRunAsAdmin(item.Arg))

    static IsPackagedApp(target) => RegExMatch(target, "i)^shell:AppsFolder\\[^\\]+$") ? true : false

    ; Whether "Run as administrator" fits the target: programs, shortcuts and scripts by their extension; a Store / packaged app when its
    ; context menu has the command, as in the Start menu (desktop apps such as Windows Terminal do, UWP apps such as Calculator don't).
    ; Asking the context menu takes some milliseconds, so the answer is kept per app (_runAs)
    static CanRunAsAdmin(target) {
        if RegExMatch(target, "i)\.(exe|lnk|bat|cmd|msc|ps1)$")
            return true
        if !ActionCatalog.IsPackagedApp(target)
            return false
        if ActionCatalog._runAs.Has(target)
            return ActionCatalog._runAs[target]
        found := false
        try {
            for verb in ActionCatalog.AppVerbs.Call(target)
                if (verb = "runas")
                    found := true
        } catch as e
            Logger.Debug("ActionCatalog: no context menu for " target " - " e.Message)
        return ActionCatalog._runAs[target] := found
    }
    static _runAs := Map()
    ; The context menu's commands of a packaged app (replaced in the tests: on a CI runner, which has no desktop shell, the Apps folder doesn't answer)
    static AppVerbs := (target) => ShellContextMenu([target]).Verbs()

    static OpenFolder(target) {
        resolved := Path.Resolve(target)
        manager := ActionCatalog.FileManagerCommand(AppSettings.General["FileManager"])
        if manager.Explorer
            return Run('explorer.exe "' resolved '"')
        Run(Trim('"' manager.Exe '" ' manager.Args) ' "' resolved '"')
    }

    static Reveal(target) {
        resolved := Path.Resolve(target)
        manager := ActionCatalog.FileManagerCommand(AppSettings.General["FileManager"])
        if manager.Explorer
            return Run('explorer.exe /select,"' resolved '"')
        command := Trim('"' manager.Exe '" ' manager.Args)
        if InStr(manager.Exe, "totalcmd")
            return Run(command ' /P "' resolved '"')                        ; Total Commander: open the containing folder and select the file
        Run(command ' "' ActionCatalog.ParentDir(resolved) '"')            ; other file managers take different arguments: open the containing folder
    }

    ; The FileManager setting -> {Exe, Args, Explorer}: the program ("quoted" when it has spaces, or up to .exe) with its
    ; %environment variables% expanded (%OneDrive%\Apps\TotalCMD64\TOTALCMD64.exe /O /T /S; Run doesn't expand them, so such a
    ; setting used to fail) and the parameters after it. Explorer when empty, explorer.exe, or a program that isn't there (logged)
    static FileManagerCommand(setting) {
        setting := Trim(setting)
        if RegExMatch(setting, '^"([^"]+)"\s*(.*)$', &m) || RegExMatch(setting, "i)^(.+?\.exe)(?:\s+(.*))?$", &m)
            program := m[1], args := m.Count >= 2 ? m[2] : ""
        else
            program := setting, args := ""
        manager := {Exe: "", Args: Trim(args), Explorer: true}
        if (program = "" || Path.Leaf(program) = "explorer.exe" || program = "explorer")
            return manager
        exe := Path.Resolve(program)                                        ; %variables%, A_ variables, a program name found on PATH
        if !FileExist(exe) {
            Logger.Warn("ActionCatalog: file manager not found (" exe "), using File Explorer")
            return manager
        }
        manager.Exe := exe, manager.Explorer := false
        return manager
    }

    static OpenUrl(address) {
        if !RegExMatch(address, "i)^[a-z][a-z0-9+.\-]*:")
            address := "https://" address
        Run(address)
    }

    ; Like "right-click -> Properties" in Explorer. Run's properties verb fails for folders on network drives etc., so SHObjectProperties is used
    static ShowProperties(target) {
        target := Path.Resolve(target)
        if !DllCall("shell32\SHObjectProperties", "Ptr", 0, "UInt", 0x2, "WStr", target, "Ptr", 0)   ; 0x2 = SHOP_FILEPATH
            Run('properties "' target '"')
    }

    static IsShellItem(target) => RegExMatch(target, "i)^shell:") ? true : false

    static CopyText(text) {
        A_Clipboard := text
        App.Notify(I18n.T("Search.Copied", ActionCatalog._Preview(text)))
    }

    ; Pastes text into the window that was active before ALTRun was shown (focusPrevious = false: into the current window).
    ; Paste mode, see Features.Snippets.PasteMode: "Clipboard" = borrow the clipboard + Ctrl+V
    ; (restored afterwards; the clipboard history doesn't record it meanwhile); "Type" = typed character by character
    ; keep: the text stays on the clipboard afterwards, so Ctrl+V pastes it again (the clipboard history, like Ditto and Win+V);
    ; then it is pasted with the clipboard even when short (no restoring, so no wait), or typed and also put on the clipboard in Type mode
    static PasteText(text, focusPrevious := true, keep := false) {
        snippetSettings := AppSettings.Feature("Snippets")
        mode  := snippetSettings["PasteMode"]
        delay := snippetSettings["PasteDelay"]

        if focusPrevious
            App.FocusPreviousWindow()
        Win.WaitModifiersUp()
        typed := keep ? mode = "Type" : ActionCatalog.TypesText(text, mode)
        if (typed && !keep) {
            SendInput("{Text}" text)
            return true
        }
        ClipboardData.Borrow(delay + 1000)
        savedClipboard := ClipboardAll()
        A_Clipboard := ""
        A_Clipboard := text
        if !ClipWait(1) {
            A_Clipboard := savedClipboard
            return false
        }
        SendInput(typed ? "{Text}" text : "^v")
        if keep
            return true
        Sleep(delay)
        A_Clipboard := savedClipboard
        return true
    }

    ; Paste mode (Features.Snippets.PasteMode) -> whether this text is typed:
    ;   Auto (default)  short (at most AutoTypeMax characters) single-line text is typed: it appears at once, leaves the clipboard alone, no wait to restore it;
    ;                   multi-line text (a typed line break is Enter, which may send a chat message), text with Tabs, very long text or emoji and other supplementary-plane characters use the clipboard
    ;   Clipboard       always the clipboard + Ctrl+V
    ;   Type            always typed character by character (for programs that don't accept pasting)
    static AutoTypeMax := 100

    static TypesText(text, mode) {
        if (mode = "Type")
            return true
        if (mode != "Auto")
            return false
        return StrLen(text) <= ActionCatalog.AutoTypeMax && !RegExMatch(text, "[\r\n\t]|[^\x{0}-\x{FFFF}]")
    }

    ; Puts files on the clipboard (the clipboard history records them as usual), to paste with Ctrl+V in TC / Explorer
    static CopyFiles(paths, cut := false) {
        if !ClipboardData.SetFiles(paths, cut)
            return false
        App.Notify(I18n.T(cut ? "Action.FileCut" : "Action.FileCopied", Path.Leaf(RTrim(paths[1], "\"))))
        return true
    }

    ; The "Edit" program Windows has registered for this file type ("Edit" in the context menu) -> {Name, Icon}; "" if none
    static EditorFor(target) {
        static ASSOCSTR_COMMAND := 1, ASSOCSTR_EXECUTABLE := 2, ASSOCSTR_FRIENDLYAPPNAME := 4
        SplitPath(target, , , &ext)
        attributes := FileExist(target)
        if (ext = "" || attributes = "" || InStr(attributes, "D"))
            return ""
        if (ActionCatalog.AssocString(ASSOCSTR_COMMAND, "." ext, "edit") = "")
            return ""
        exe := ActionCatalog.AssocString(ASSOCSTR_EXECUTABLE, "." ext, "edit")
        return {Name: ActionCatalog.AssocString(ASSOCSTR_FRIENDLYAPPNAME, "." ext, "edit")
              , Icon: (exe != "" && FileExist(exe)) ? exe : A_WinDir "\notepad.exe"}
    }

    ; AssocQueryStringW: the command, program or program name of a verb ("edit") of a file type (".ahk"); "" if none
    static AssocString(kind, ext, verb) {
        static ASSOCF_NOTRUNCATE := 0x20, ASSOCF_INIT_IGNOREUNKNOWN := 0x400   ; unregistered types must not fall back to the "Open with" settings
        size := 0
        DllCall("shlwapi\AssocQueryStringW", "UInt", ASSOCF_NOTRUNCATE | ASSOCF_INIT_IGNOREUNKNOWN, "UInt", kind, "WStr", ext, "WStr", verb
            , "Ptr", 0, "UInt*", &size)
        if (size <= 1)
            return ""
        buf := Buffer(size * 2, 0)
        if (DllCall("shlwapi\AssocQueryStringW", "UInt", ASSOCF_NOTRUNCATE | ASSOCF_INIT_IGNOREUNKNOWN, "UInt", kind, "WStr", ext, "WStr", verb
            , "Ptr", buf, "UInt*", &size) != 0)
            return ""
        return StrGet(buf)
    }

    ; Opens with the registered editor (like "Edit" in the context menu)
    static EditFile(target) {
        Run('edit "' Path.Resolve(target) '"')
    }

    static OpenWith(target) {
        Run("rundll32.exe shell32.dll,OpenAs_RunDLL " Path.Resolve(target))    ; the path isn't quoted (OpenAs_RunDLL takes the whole rest of the command line)
    }

    ; Windows' own copy / move (progress window, asks on name conflicts, can be undone), in the background so ALTRun doesn't hang
    static CopyTo(target, destination, move := false) {
        folder := ComObject("Shell.Application").NameSpace(RTrim(destination, "\") (RegExMatch(destination, "^[A-Za-z]:\\?$") ? "\" : ""))
        if !IsObject(folder)
            return false
        move ? folder.MoveHere(target, 0) : folder.CopyHere(target, 0)
        App.Notify(I18n.T(move ? "Action.Moving" : "Action.Copying", Path.Leaf(RTrim(target, "\")), destination), 2000)
        return true
    }

    ; Move to the Recycle Bin with Windows' "Delete" (confirms first; can be restored from the Recycle Bin)
    static Recycle(target) {
        SplitPath(RTrim(target, "\"), &name, &dir)
        try {
            folderItem := ComObject("Shell.Application").NameSpace(RegExMatch(dir, "^[A-Za-z]:$") ? dir "\" : dir).ParseName(name)
            folderItem.InvokeVerb("delete")
            return true
        }
        return false
    }

    static ParentDir(target) {
        SplitPath(target, , &dir)
        return dir
    }

    static _Preview(text) {
        text := RegExReplace(text, "\s+", " ")
        return (StrLen(text) > 60) ? SubStr(text, 1, 60) "..." : text
    }
}
