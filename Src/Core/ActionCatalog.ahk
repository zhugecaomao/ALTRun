;===============================================================================
; ActionCatalog.ahk - what running a result does, and the operations behind it: open / reveal / copy / paste ...
;   (AutoHotkey v2)
;-------------------------------------------------------------------------------
; As in Alfred:
;   Enter        the default action (file / program / folder / URL: open; text: copy)
;   Ctrl+Enter   file / folder: reveal in the file manager; text: paste into the foreground window
;   Alt+Enter    copy (path / URL / text)
; The actions panel (→), F3 edit and Ctrl+Del delete are built on these in Src\UI\ResultActions.ahk.
;
; Usage:
;   ActionCatalog.RunDefault(item)
;   ActionCatalog.RunModifier(item, "ctrl" | "alt")
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

    static RunAsAdmin(target, arguments := "") {
        Run(Trim('*RunAs "' Path.Resolve(target) '" ' arguments))
    }

    static OpenFolder(target) {
        resolved := Path.Resolve(target)
        fileManager := AppSettings.General["FileManager"]
        if (fileManager = "" || InStr(fileManager, "explorer"))
            return Run('explorer.exe "' resolved '"')
        Run(fileManager ' "' resolved '"')
    }

    static Reveal(target) {
        resolved := Path.Resolve(target)
        fileManager := AppSettings.General["FileManager"]
        if (fileManager = "" || InStr(fileManager, "explorer"))
            return Run('explorer.exe /select,"' resolved '"')
        if InStr(fileManager, "totalcmd")
            return Run(fileManager ' /P "' resolved '"')                   ; Total Commander: open the containing folder and select the file
        Run(fileManager ' "' ActionCatalog.ParentDir(resolved) '"')        ; other file managers take different arguments: open the containing folder
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
    static PasteText(text, focusPrevious := true) {
        snippetSettings := AppSettings.Feature("Snippets")
        mode  := snippetSettings["PasteMode"]
        delay := snippetSettings["PasteDelay"]

        if focusPrevious
            App.FocusPreviousWindow()
        Win.WaitModifiersUp()
        if ActionCatalog.TypesText(text, mode) {
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
        SendInput("^v")
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
