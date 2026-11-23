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
        item := ActionCatalog.ForWindow(item)
        switch item.Kind {
            case "file"  : ActionCatalog.OpenFile(item.Arg, item.Arguments)
            case "folder": ActionCatalog.OpenFolder(item.Arg)
            case "url"   : ActionCatalog.OpenUrl(item.Arg)
            case "text"  : ActionCatalog.CopyText(item.Arg)
        }
    }

    static RunModifier(item, modifier) {
        if (modifier = "admin")
            return ActionCatalog.CanRunElevated(item) ? ActionCatalog.RunAsAdmin((filled := ActionCatalog.ForWindow(item)).Arg, filled.Arguments) : ActionCatalog.RunDefault(item)
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

    ; {hwnd} / {pid} in a custom command's target or arguments (item.WindowArgs): the window in front before ALTRun was shown
    ; (its handle and process id, in decimal), filled in when the command runs, so a pinned or recently used entry gets the window
    ; of that moment. Returns a copy with Kind / Arg / Arguments filled in, or the item itself when it has none
    static ForWindow(item) {
        if !item.WindowArgs
            return item
        hwnd := App.PreviousWindow
        return ResultItem(item.Title, item.Subtitle, {Kind: item.Kind, Arg: ActionCatalog.FillWindow(item.Arg, hwnd)
            , Arguments: ActionCatalog.FillWindow(item.Arguments, hwnd)})
    }

    ; text with {hwnd} / {pid} (any case) replaced by the window's handle and process id; both empty when the window is gone or hwnd is 0
    static FillWindow(text, hwnd) {
        if !RegExMatch(text, "i)\{(hwnd|pid)\}")
            return text
        alive := hwnd && WinExist("ahk_id " hwnd)
        pid := ""
        if alive
            try pid := WinGetPID("ahk_id " hwnd)
        text := StrReplace(text, "{hwnd}", alive ? hwnd : "")
        return StrReplace(text, "{pid}", pid)
    }

    ; Target / Arguments contain {hwnd} or {pid}
    static UsesWindow(texts*) {
        for text in texts
            if RegExMatch(text, "i)\{(hwnd|pid)\}")
                return true
        return false
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

    ; General.FolderNewTab: File Explorer opens the folder in a new tab of a window already open (ExplorerTabs; a new window as
    ; before when there is none or it fails), Total Commander gets /O /T (see TabArgs), Arke is asked for a new tab through its
    ; interface (its command line has no switch for it; when no Arke window answers, the command line as before); other file
    ; managers use their own parameters
    static OpenFolder(target) {
        resolved := Path.Resolve(target)
        manager := ActionCatalog.FileManagerCommand(AppSettings.General["FileManager"])
        if manager.Explorer {
            if (AppSettings.General["FolderNewTab"] && ActionCatalog._InExplorerTab(resolved))
                return
            return Run('explorer.exe "' resolved '"')
        }
        if (AppSettings.General["FolderNewTab"] && ActionCatalog.IsArke(manager) && Arke.Open(resolved, true, App.PreviousWindow))
            return
        Run(ActionCatalog.FileManagerLine(manager, resolved, '"' resolved '"'))
    }

    ; Ctrl+Enter / "Reveal in File Manager": the folder that holds target, with target selected where the file manager can
    static Reveal(target) {
        resolved := ActionCatalog.RevealPath(target)
        manager := ActionCatalog.FileManagerCommand(AppSettings.General["FileManager"])
        if manager.Explorer {
            if (AppSettings.General["FolderNewTab"] && ActionCatalog._InExplorerTab(ActionCatalog.ParentDir(resolved), Path.Leaf(resolved)))
                return
            return Run('explorer.exe /select,"' resolved '"')
        }
        if (AppSettings.General["FolderNewTab"] && ActionCatalog.IsArke(manager) && Arke.Open(resolved, true, App.PreviousWindow))
            return                                                              ; a file's path: its folder, with the file selected
        Run(ActionCatalog.RevealLine(manager, resolved))
    }

    ; The file manager's command line for Reveal. Total Commander: /P before a path opens the folder that holds it and selects
    ; it (a file or a folder), also together with /L= / /R=, so it gets the item itself and /P, with or without {path}; without
    ; /P it opens the path and selects nothing. Other file managers take different arguments: the containing folder
    static RevealLine(manager, resolved) {
        if ActionCatalog.IsTotalCommander(manager)
            return ActionCatalog.FileManagerLine(manager, resolved, '/P "' resolved '"', "/P")
        folder := ActionCatalog.ParentDir(resolved)
        return ActionCatalog.FileManagerLine(manager, folder, '"' folder '"')
    }

    ; The file manager's command line. {path} in its parameters (FileManager setting) places the path: the folder to open, or
    ; what Reveal passes (see RevealLine), e.g. Total Commander's /R={path}: always the right panel. flag goes first when there
    ; is {path} and the parameters don't have it yet (Total Commander's /P for Reveal). Without {path}, legacyArgs follow the
    ; parameters, as before placeholders existed (so existing settings keep working)
    static FileManagerLine(manager, target, legacyArgs, flag := "") {
        args := ActionCatalog.TabArgs(manager)
        if !(InStr(args, "{path}") > 0)
            return Trim('"' manager.Exe '" ' Trim(args " " legacyArgs))
        if (flag != "" && !RegExMatch(args, "i)(^|\s)" flag "(\s|$)"))                ; /P=R is another switch
            args := flag " " args
        return '"' manager.Exe '" ' ActionCatalog.FillPath(args, target)
    }

    ; Every {path} (any case) replaced in one pass. A path with spaces is quoted unless that {path} is already inside quotes
    ; (an odd number of " before it, as in /R="{path}"): {path} and /R={path} both become /R="C:\My Folder". Only whitespace
    ; needs it, as Run starts the program without cmd.exe, and a Windows path can't contain ". Not RegExReplace: "$" in a path
    ; (C:\$Recycle.Bin, \\server\share$) would be read as a back-reference
    static FillPath(args, target) {
        quoted := RegExMatch(target, "\s") > 0 ? '"' target '"' : target
        out := "", pos := 1
        while ((found := InStr(args, "{path}", false, pos)) > 0) {
            StrReplace(SubStr(args, 1, found - 1), '"', '"', , &quotes)
            out .= SubStr(args, pos, found - pos) (Mod(quotes, 2) ? target : quoted)
            pos := found + StrLen("{path}")
        }
        return out SubStr(args, pos)
    }

    static IsTotalCommander(manager) => InStr(Path.Leaf(manager.Exe), "totalcmd") > 0
    static IsArke(manager) => Path.Leaf(manager.Exe) = "arke.exe"

    ; The file or folder that Reveal shows. A custom command's program given by name (notepad, chrome.exe) is looked up where Run
    ; finds it: in PATH (Path.Resolve), then under App Paths (Path.AppPath), so Ctrl+Enter opens the program's folder
    static RevealPath(target) {
        resolved := Path.Resolve(target)
        if (!InStr(resolved, "\") && (registered := Path.AppPath(resolved)) != "")
            return registered
        return resolved
    }

    ; A new tab of an Explorer window already open: the one in front before ALTRun was shown first; false = open a new window
    static _InExplorerTab(folder, file := "") {
        if ExplorerTabs.Open(folder, file, App.PreviousWindow) {
            if (file != "" && InStr(ExplorerTabs.SelectInfo, "failed") = 1)
                ActionCatalog._TabProblem("the file was not selected in the new Explorer tab (" ExplorerTabs.SelectInfo ")")
            return true
        }
        if ExplorerTabs.LastErrorExpected
            Logger.Debug("ActionCatalog: new Explorer tab not used (" ExplorerTabs.LastError "), opening a window")
        else
            ActionCatalog._TabProblem("new Explorer tab failed (" ExplorerTabs.LastError "), opened a window instead")
        return false
    }

    ; New tabs rely on Explorer commands that aren't documented, so a Windows update may break them. The first problem in a run
    ; is written to the log even when the log is off (Logger.Record), so a report of "folders open in a new window" can say why;
    ; later ones only to the debug log
    static _tabWarned := false
    static _TabProblem(text) {
        text := "ActionCatalog: " text "; Windows " A_OSVersion
        if ActionCatalog._tabWarned
            return Logger.Debug(text)
        ActionCatalog._tabWarned := true
        Logger.Record("WRN", text " (written once a run)")
    }

    ; The file manager's parameters; with General.FolderNewTab, Total Commander's gain /O (the running one) and /T (a new tab)
    ; when they aren't there yet
    static TabArgs(manager) {
        args := manager.Args
        if !(AppSettings.General["FolderNewTab"] && ActionCatalog.IsTotalCommander(manager))
            return args
        for flag in ["/O", "/T"]
            if !RegExMatch(args, "i)(^|\s)" flag "(\s|$)")
                args := Trim(args " " flag)
        return args
    }

    ; The FileManager setting -> {Exe, Args, Explorer}: the program and its parameters (see SplitCommand / ProgramPath).
    ; Explorer when empty, explorer.exe, or a program that isn't there (logged)
    static FileManagerCommand(setting) {
        parts := ActionCatalog.SplitCommand(setting)
        manager := {Exe: "", Args: parts.Args, Explorer: true}
        if (parts.Program = "" || Path.Leaf(parts.Program) = "explorer.exe" || parts.Program = "explorer")
            return manager
        if ((exe := ActionCatalog.ProgramPath(parts.Program, "file manager", "using File Explorer")) != "")
            manager.Exe := exe, manager.Explorer := false
        return manager
    }

    ; The Browser setting -> {Exe, Args}: the program that opens web pages and its parameters (see SplitCommand / ProgramPath);
    ; Exe = "" (the default browser) when empty or the program isn't there (logged)
    static BrowserCommand(setting) {
        parts := ActionCatalog.SplitCommand(setting)
        exe := (parts.Program = "") ? "" : ActionCatalog.ProgramPath(parts.Program, "browser", "using the default browser")
        return {Exe: exe, Args: parts.Args}
    }

    ; A program setting -> {Program, Args}: the program ("quoted" when it has spaces, or up to .exe) and the parameters
    ; after it, e.g. %OneDrive%\Apps\TotalCMD64\TOTALCMD64.exe /O /T /S or "C:\My Apps\Vivaldi\vivaldi.exe" --new-window
    static SplitCommand(setting) {
        setting := Trim(setting)
        if RegExMatch(setting, '^"([^"]+)"\s*(.*)$', &m) || RegExMatch(setting, "i)^(.+?\.exe)(?:\s+(.*))?$", &m)
            return {Program: m[1], Args: Trim(m.Count >= 2 ? m[2] : "")}
        return {Program: setting, Args: ""}
    }

    ; The full path of the program in a setting: %environment variables% expanded (Run doesn't expand them, so such a
    ; setting used to fail), A_ variables, a program name found on PATH; "" when it isn't there (logged)
    static ProgramPath(program, what, instead) {
        exe := Path.Resolve(program)
        if FileExist(exe)
            return exe
        Logger.Warn("ActionCatalog: " what " not found (" exe "), " instead)
        return ""
    }

    ; Web pages (http / https; an address without a scheme gets https://) open in General.Browser when it is set; other
    ; links (mailto:, ms-settings:, search-ms: ...), and web pages when it isn't, in the program Windows registered for them
    static OpenUrl(address) {
        if !RegExMatch(address, "i)^[a-z][a-z0-9+.\-]*:")
            address := "https://" address
        if (RegExMatch(address, "i)^https?://") && (browser := ActionCatalog.BrowserCommand(AppSettings.General["Browser"])).Exe != "")
            return Run(ActionCatalog.BrowserCommandLine(browser, address))
        Run(address)
    }

    ; The command line that opens address in the browser: the address is one quoted argument (a double quote in it becomes
    ; %22, and backslashes before the closing quote are doubled, so they don't escape it)
    static BrowserCommandLine(browser, address) {
        address := StrReplace(address, '"', "%22")
        if RegExMatch(address, "\\+$", &tail)
            address .= tail[0]
        return Trim('"' browser.Exe '" ' browser.Args) ' "' address '"'
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
        savedClipboard := ClipboardAll(), restore := true
        try {                                                               ; the user's clipboard comes back even when this fails half way
            A_Clipboard := ""
            A_Clipboard := text
            if !ClipWait(1)
                return false
            SendInput(typed ? "{Text}" text : "^v")
            if keep {                                                       ; Keep: the text stays on the clipboard
                restore := false
                return true
            }
            Sleep(delay)
        } finally {
            if restore
                A_Clipboard := savedClipboard
        }
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
