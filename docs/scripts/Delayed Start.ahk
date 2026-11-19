; @altrun.title          Delayed Start
; @altrun.keyword        delayedstart
; @altrun.description    Start programs such as OneDrive a few minutes after signing in to Windows, so the desktop is ready sooner (Task Scheduler, no administrator rights)
; @altrun.description_zh 登录 Windows 几分钟后再启动 OneDrive 等程序, 让桌面更快可用 (任务计划程序, 不需要管理员权限)
; @altrun.network        none
;
; Type "delayedstart" in ALTRun. Each entry is a Windows Task Scheduler task in the folder \ALTRun\<user name>, started at your
; sign-in after its delay: Windows starts the programs, so neither ALTRun nor this script has to be running. The tasks are the list
; (nothing else is saved); a change is made in Task Scheduler when you click OK. F1 opens the help
; (https://github.com/zhugecaomao/ALTRun/wiki/Delayed-Start). Up to ALTRun 2026.11.16.23 this was Preferences → Delayed Start:
; the tasks set up there are listed here as they are.
#Requires AutoHotkey v2.0

if (A_LineFile = A_ScriptFullPath)                                          ; run as a program; ALTRun's tests #Include this file instead
    DelayedStartWindow.Main()

;===============================================================================
; DelayedStart - the tasks, and the programs that start by themselves at sign-in (Task Manager → Startup apps)
;-------------------------------------------------------------------------------
;   DelayedStart.Items()           the list: [Map(Name, Target, Arguments, Delay, Enabled)...] read from Task Scheduler
;   DelayedStart.Sync(items)       makes the tasks match the list: "" or the problems, one per line
;   DelayedStart.AutoStartEntries() / AutoStartEntry(target) / OffEntry(target) / SetEntryOff(source, off)   startup apps
;===============================================================================
class DelayedStart {
    static Description := "Set up by the ALTRun script Delayed Start.ahk, which lists the tasks in this folder."
    static FolderPath => "\ALTRun\" DelayedStart.SafeName(A_UserName)
    static Lang := DelayedStart.IsChinese(A_Language) ? "zh" : "en"
    static HelpUrl => "https://github.com/zhugecaomao/ALTRun/wiki/Delayed-Start" (DelayedStart.Lang = "zh" ? "-zh" : "")

    ; Simplified Chinese (China, Singapore): the other languages get the English text
    static IsChinese(lang) => RegExMatch(lang, "i)^(0804|1004)$") > 0

    static NewItem() => Map("Name", "", "Target", "", "Arguments", "", "Delay", 300, "Enabled", 1)

    ; The tasks in the folder as entries; throws when Task Scheduler can't be used
    static Items(folderPath := DelayedStart.FolderPath) {
        items := []
        for task in TaskScheduler.Tasks(folderPath)
            items.Push(Map("Name", task.Name, "Target", DelayedStart.ShortPath(Trim(task.Path, '"')), "Arguments", task.Arguments, "Delay", task.Delay, "Enabled", task.Enabled))
        return items
    }

    ; Makes the tasks in folderPath match items; returns "" or the problems, one per line (the other entries are still set up).
    ; One connection to Task Scheduler for all of it. An empty list deletes the folder
    static Sync(items, folderPath := DelayedStart.FolderPath) {
        try {
            service := TaskScheduler.Service()
            existing := TaskScheduler.Tasks(folderPath, service)
        } catch as e
            return DelayedStart.T("NoScheduler", e.Message)
        plan := DelayedStart.Plan(DelayedStart.Desired(items), existing)
        problems := ""
        for name in plan.Delete {
            try TaskScheduler.Delete(folderPath, name, service)
            catch as e
                problems .= name ": " e.Message "`n"
        }
        for entry in plan.Register {
            try TaskScheduler.RegisterLogonTask(folderPath, entry.Name, entry.Spec, service)
            catch as e
                problems .= entry.Name ": " e.Message "`n"
        }
        if (!items.Length && problems = "")
            try TaskScheduler.DeleteFolder(folderPath, service)
        return RTrim(problems, "`n")
    }

    ; The tasks the list asks for: [{Name, Spec: {Path, Arguments, WorkingDirectory, Delay, Enabled, Description}}...]; entries without a
    ; program are left out. Task names come from the entry names, made safe for Task Scheduler and unique (OneDrive, OneDrive (2)...)
    static Desired(items) {
        list := [], used := Map()
        for item in items {
            target := Trim(DelayedStart._Get(item, "Target"), "`" `t")
            if (target = "")
                continue
            name := DelayedStart.SafeName(DelayedStart._Get(item, "Name"))
            if (name = "")
                name := DelayedStart.SafeName(DelayedStart.Leaf(target))
            unique := name, n := 1
            while used.Has(StrLower(unique))
                unique := name " (" (++n) ")"
            used[StrLower(unique)] := true
            program := DelayedStart.ProgramPath(target)
            SplitPath(program, , &dir)
            delay := DelayedStart._Get(item, "Delay")
            list.Push({Name: unique, Spec: {Path: InStr(program, " ") ? '"' program '"' : program   ; quoted as typed in Task Scheduler: a path with spaces works either way there
                , Arguments: Trim(DelayedStart._Get(item, "Arguments")), WorkingDirectory: dir
                , Delay: IsInteger(delay) ? Max(0, Integer(delay)) : 0, Enabled: DelayedStart._Get(item, "Enabled", 1) ? 1 : 0
                , Description: DelayedStart.Description}})
        }
        return list
    }

    ; What to change so that the tasks in the folder (existing, from TaskScheduler.Tasks) are the desired ones: {Register: [...], Delete: [names]}.
    ; A task that is already right isn't registered again; tasks no entry asks for are deleted
    static Plan(desired, existing) {
        byName := Map(), wanted := Map(), plan := {Register: [], Delete: []}
        for task in existing
            byName[StrLower(task.Name)] := task
        for entry in desired {
            key := StrLower(entry.Name)
            wanted[key] := true
            if !(byName.Has(key) && DelayedStart.Same(entry.Spec, byName[key]))
                plan.Register.Push(entry)
        }
        for task in existing
            if !wanted.Has(StrLower(task.Name))
                plan.Delete.Push(task.Name)
        return plan
    }

    static Same(spec, task) {
        return Trim(spec.Path, '"') = Trim(task.Path, '"') && spec.Arguments == task.Arguments && spec.WorkingDirectory = task.WorkingDirectory
            && spec.Delay = task.Delay && spec.Enabled = task.Enabled
    }

    ; The program a task runs: environment variables expanded, a bare name ("notepad.exe") found in PATH as from the Run box, a relative
    ; path taken from this script's folder. Task Scheduler gets the full path
    static ProgramPath(target) {
        resolved := DelayedStart.ExpandEnv(Trim(target))
        if (resolved = "")
            return ""
        if !InStr(resolved, "\") {
            buf := Buffer(32767 * 2)
            if DllCall("SearchPathW", "Ptr", 0, "WStr", resolved, "WStr", ".exe", "UInt", 32767, "Ptr", buf, "Ptr", 0, "UInt")
                resolved := StrGet(buf)
        } else if !RegExMatch(resolved, "^([A-Za-z]:|\\\\)") {
            buf := Buffer(32767 * 2)
            if DllCall("GetFullPathNameW", "WStr", A_ScriptDir "\" resolved, "UInt", 32767, "Ptr", buf, "Ptr", 0, "UInt")
                resolved := StrGet(buf)
        }
        return resolved
    }

    ; A path for the list and the edit window: the user's own folders as %LocalAppData%, %AppData%, %UserProfile% (expanded again when set up)
    static ShortPath(fullPath) {
        for name in ["LocalAppData", "AppData", "UserProfile"]
            if ((dir := EnvGet(name)) != "" && InStr(fullPath, dir "\") = 1)
                return "%" name "%" SubStr(fullPath, StrLen(dir) + 1)
        return fullPath
    }

    static ExpandEnv(text) {
        if !InStr(text, "%")
            return text
        size := DllCall("ExpandEnvironmentStringsW", "WStr", text, "Ptr", 0, "UInt", 0, "UInt")
        buf := Buffer(size * 2)
        DllCall("ExpandEnvironmentStringsW", "WStr", text, "Ptr", buf, "UInt", size, "UInt")
        return StrGet(buf)
    }

    static Leaf(fullPath) {
        SplitPath(Trim(fullPath, '" '), &name)
        return name
    }

    ; A task / folder name Task Scheduler takes: no \ / : * ? " < > | or control characters, no dots or spaces at the ends
    static SafeName(text) {
        text := RegExReplace(text, '[\\/:*?"<>|\x00-\x1F]', "_")
        return Trim(text, " .")
    }

    ; Delay for the list: "5 min", "30 s", "1 min 30 s", "At once"
    static DelayText(seconds) {
        if !IsInteger(seconds) || seconds <= 0
            return DelayedStart.T("NoDelay")
        minutes := seconds // 60, rest := Mod(seconds, 60)
        if !minutes
            return DelayedStart.T("Seconds", rest)
        return rest ? DelayedStart.T("MinSec", minutes, rest) : DelayedStart.T("Minutes", minutes)
    }

    ;---------------------------------------------------------------------------
    ; Programs that start by themselves at sign-in: the startup apps of Task Manager → Startup apps (Run keys and Startup folders).
    ; One can be turned into an entry; its own start is then turned off the way Task Manager does it (StartupApproved): nothing is
    ; deleted, and it can be turned on again there or when the entry is deleted here
    ;---------------------------------------------------------------------------
    static ScopeOrder := ["HKCU", "User", "HKLM", "HKLM32", "Common"]
    static ApprovedKey := "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\"

    ; Where startup apps are: scope -> {Run: registry key or "", Folder: folder or "", Approved: its StartupApproved key, Writable: this user
    ; can turn them off (the all-users ones need an administrator)}
    static Scopes() {
        runKey := "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", approved := DelayedStart.ApprovedKey
        return Map(
            "HKCU",   {Run: "HKCU\" runKey, Folder: "", Approved: "HKCU" approved "Run", Writable: true},
            "HKLM",   {Run: "HKLM\" runKey, Folder: "", Approved: "HKLM" approved "Run", Writable: false},
            "HKLM32", {Run: "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run", Folder: "", Approved: "HKLM" approved "Run32", Writable: false},
            "User",   {Run: "", Folder: A_Startup, Approved: "HKCU" approved "StartupFolder", Writable: true},
            "Common", {Run: "", Folder: A_StartupCommon, Approved: "HKLM" approved "StartupFolder", Writable: false})
    }

    ; The startup apps: [{Source: "HKCU|OneDrive", Scope, Name, Command, Program, Arguments, Off, Writable}...];
    ; includeOff: also those turned off in Task Manager
    static AutoStartEntries(includeOff := false) {
        entries := [], scopes := DelayedStart.Scopes()
        for scope in DelayedStart.ScopeOrder {                             ; this user's first, then those for all users
            where := scopes[scope]
            if (where.Run != "") {
                try {
                    Loop Reg, where.Run, "V" {
                        if (A_LoopRegType != "REG_SZ" && A_LoopRegType != "REG_EXPAND_SZ")
                            continue
                        name := A_LoopRegName, command := DelayedStart.ExpandEnv(RegRead())
                        entries.Push(DelayedStart._Entry(scope, where, name, command, DelayedStart.CommandProgram(command), DelayedStart.CommandArguments(command)))
                    }
                }
            } else if (where.Folder != "") {
                Loop Files, where.Folder "\*.lnk" {
                    target := "", args := ""
                    try FileGetShortcut(A_LoopFilePath, &target, , &args)
                    if (target != "")
                        entries.Push(DelayedStart._Entry(scope, where, A_LoopFileName, Trim('"' target '" ' args), target, args))
                }
            }
        }
        if includeOff
            return entries
        on := []
        for entry in entries
            if !entry.Off
                on.Push(entry)
        return on
    }

    static _Entry(scope, where, name, command, program, args) {
        off := false
        try off := DelayedStart.IsOffData(RegRead(where.Approved, name))
        return {Source: scope "|" name, Scope: scope, Name: name, Command: command, Program: program, Arguments: args, Off: off, Writable: where.Writable}
    }

    ; The startup app with this source ("" when it is gone)
    static FindEntry(source) {
        for entry in DelayedStart.AutoStartEntries(true)
            if (entry.Source = source)
                return entry
        return ""
    }

    ; The startup app starting this program at sign-in (still on), "" when none. Compares file names: the same program may be installed
    ; per user or for all users
    static AutoStartEntry(target) => DelayedStart._EntryFor(target, (entry) => !entry.Off)

    ; The startup app of this program whose own start is turned off and which this user can turn on again, "" when none: offered when
    ; its entry is deleted
    static OffEntry(target) => DelayedStart._EntryFor(target, (entry) => entry.Off && entry.Writable)

    static _EntryFor(target, wanted) {
        leaf := StrLower(DelayedStart.Leaf(DelayedStart.ProgramPath(target)))
        if (leaf = "")
            return ""
        for entry in DelayedStart.AutoStartEntries(true)
            if (wanted(entry) && StrLower(DelayedStart.Leaf(entry.Program)) = leaf)
                return entry
        return ""
    }

    ; Where a startup app is, for messages: "Run (you): OneDrive", "Startup folder (all users): Teams.lnk"
    static SourceText(source) {
        parts := StrSplit(source, "|", , 2)
        return DelayedStart.T("Where." parts[1]) ": " (parts.Length > 1 ? parts[2] : "")
    }

    ; Turns a startup app off (or on again) as Task Manager does: its StartupApproved value. Throws when it can't be written
    static SetEntryOff(source, off) {
        parts := StrSplit(source, "|", , 2), scopes := DelayedStart.Scopes()
        if (parts.Length < 2 || !scopes.Has(parts[1]))
            throw ValueError("Unknown startup app: " source, -1)
        RegWrite(DelayedStart.ApprovedData(off), "REG_BINARY", scopes[parts[1]].Approved, parts[2])
    }

    ; A StartupApproved value: 12 bytes, the first 02 = on, 03 = off, then (when off) the time it was turned off (FILETIME)
    static ApprovedData(off, fileTime := "") {
        if !off
            return "020000000000000000000000"
        if (fileTime = "")
            DllCall("GetSystemTimeAsFileTime", "Int64*", &fileTime := 0)
        hex := "03000000"
        Loop 8
            hex .= Format("{:02X}", (fileTime >> (8 * (A_Index - 1))) & 0xFF)
        return hex
    }

    ; An odd first byte (01, 03) = turned off
    static IsOffData(data) => (StrLen(data) >= 2 && IsXDigit(SubStr(data, 1, 2)) && Mod(Integer("0x" SubStr(data, 1, 2)), 2) = 1)

    ; The program of a command line: "C:\Program Files\x.exe" /a -> C:\Program Files\x.exe; C:\x\y.exe /a -> C:\x\y.exe
    static CommandProgram(command) {
        command := Trim(command)
        if RegExMatch(command, '^"([^"]+)"', &m)
            return m[1]
        if RegExMatch(command, "i)^(.+?\.(exe|com|bat|cmd))(\s|$)", &m)
            return m[1]
        return RegExReplace(command, "\s.*$")
    }

    ; What follows the program: "C:\x\y.exe" /a /b -> /a /b
    static CommandArguments(command) {
        command := Trim(command)
        if RegExMatch(command, '^"[^"]+"\s*(.*)$', &m)
            return Trim(m[1])
        if RegExMatch(command, "i)^.+?\.(exe|com|bat|cmd)(\s+(.*))?$", &m)
            return Trim(m[3])
        return RegExMatch(command, "\s+(.*)$", &m) ? Trim(m[1]) : ""
    }

    ; An entry made from a startup app: its name (a shortcut's without .lnk), program and arguments, 5 minutes
    static ItemFromEntry(entry) {
        name := RegExReplace(entry.Name, "i)\.lnk$")
        return Map("Name", name, "Target", entry.Program, "Arguments", entry.Arguments, "Delay", 300, "Enabled", 1, "Source", entry.Source)
    }

    static _Get(item, key, default := "") => (item is Map && item.Has(key)) ? item[key] : default

    ;---------------------------------------------------------------------------
    ; Text: English, or Simplified Chinese on a Chinese system. {1} {2}... are filled in by T
    ;---------------------------------------------------------------------------
    static T(key, args*) {
        strings := DelayedStart.Strings[DelayedStart.Lang]
        text := strings.Has(key) ? strings[key] : DelayedStart.Strings["en"][key]
        for index, arg in args
            text := StrReplace(text, "{" index "}", arg)
        return text
    }

    static Strings := Map(
        "en", Map(
            "Title", "Delayed Start",
            "Note", "Windows starts these programs after you sign in, once the delay has passed. They are tasks in {1}Task Scheduler{3} (folder ALTRun\<your user name>) and run with your normal rights: programs that need administrator rights can't be started this way. Startup Apps... turns a program that starts by itself into an entry here and turns its own start off, as in {2}Task Manager → Startup apps{3}. F1: help",
            "Col.Name", "Name", "Col.Program", "Program", "Col.Arguments", "Arguments", "Col.Delay", "Delay", "Col.Enabled", "Enabled",
            "Col.Command", "Command", "Col.Where", "From",
            "Add", "Add...", "Edit", "Edit...", "Delete", "Delete", "StartupApps", "Startup Apps...", "Close", "Close",
            "OK", "OK", "Cancel", "Cancel", "Browse", "Browse...",
            "Editor.Title", "Delayed Start Entry", "Field.DelaySeconds", "Delay (seconds)",
            "Field.Name", "Shown in the list; also the name of the task in Task Scheduler.",
            "Field.Target", "The program or script to start. Environment variables such as %LocalAppData% and program names in PATH work.",
            "Field.Arguments", "Optional, e.g. /background for OneDrive.",
            "Field.Delay", "Seconds after signing in to Windows, e.g. 300 = 5 minutes; 0 = at once.",
            "Field.Enabled", "Start it after signing in",
            "NoProgram", "Choose the program to start.",
            "Programs", "Programs",
            "On", "On", "Off", "Off", "NoDelay", "At once", "Seconds", "{1} s", "Minutes", "{1} min", "MinSec", "{1} min {2} s",
            "DeleteConfirm", "Delete {1}? Its task is deleted from Task Scheduler, and the program no longer starts after signing in.",
            "NoScheduler", "Task Scheduler can't be used: {1}",
            "SyncFailed", "Some tasks couldn't be set up in Task Scheduler:`n`n{1}",
            "AlsoAutoStarts", "{1} also starts by itself after signing in ({2}), so it would start at once, not after the delay.`n`nTurn its own start off? This is the same as turning it off in Task Manager → Startup apps; deleting this entry later offers to turn it on again.",
            "AlsoAutoStartsAdmin", "{1} also starts by itself after signing in, for all users ({2}), so it would start at once, not after the delay. Turning that off needs administrator rights: turn it off in Task Manager → Startup apps.`n`nSave anyway?",
            "PickTitle", "Startup Apps",
            "PickDesc", "Programs Windows starts at sign-in (as in Task Manager → Startup apps). Choose one to start it after a delay instead: its own start is turned off when you click OK.",
            "NoStartupApps", "No other programs start at sign-in.",
            "Where.HKCU", "Run (you)", "Where.HKLM", "Run (all users)", "Where.HKLM32", "Run (all users, 32-bit)",
            "Where.User", "Startup folder (you)", "Where.Common", "Startup folder (all users)",
            "TurnOffFailed", "The task is set up, but the program's own start couldn't be turned off:`n`n{1}",
            "TurnOffManually", "The delayed start task is set up. The program still starts by itself at sign-in, for all users, which needs administrator rights to turn off:`n`n{1}`n`nOpen Task Manager → Startup apps to turn it off?",
            "TurnBackOn", "The own start of {1} at sign-in is turned off:`n`n{2}`n`nLet it start by itself at sign-in again?"),
        "zh", Map(
            "Title", "延迟启动",
            "Note", "登录 Windows 后, 过了延迟时间由 Windows 启动这些程序。它们是{1}任务计划程序{3}里的任务 (文件夹 ALTRun\<你的用户名>), 用你的普通权限运行: 需要管理员权限的程序不能这样启动。点 启动项... 可以把自己会自启动的程序加到这里, 并像{2}任务管理器 → 启动应用{3}那样关掉它自己的启动。F1: 帮助",
            "Col.Name", "名称", "Col.Program", "程序", "Col.Arguments", "参数", "Col.Delay", "延迟", "Col.Enabled", "启用",
            "Col.Command", "命令", "Col.Where", "来源",
            "Add", "添加...", "Edit", "编辑...", "Delete", "删除", "StartupApps", "启动项...", "Close", "关闭",
            "OK", "确定", "Cancel", "取消", "Browse", "浏览...",
            "Editor.Title", "延迟启动项", "Field.DelaySeconds", "延迟 (秒)",
            "Field.Name", "显示在列表里, 也是任务计划程序里的任务名",
            "Field.Target", "要启动的程序或脚本。可以用环境变量 (如 %LocalAppData%) 和 PATH 里的程序名",
            "Field.Arguments", "可选, 例如 OneDrive 用 /background",
            "Field.Delay", "登录 Windows 后多少秒启动, 例如 300 = 5 分钟; 0 = 立即",
            "Field.Enabled", "登录后启动",
            "NoProgram", "请选择要启动的程序。",
            "Programs", "程序",
            "On", "开", "Off", "关", "NoDelay", "立即", "Seconds", "{1} 秒", "Minutes", "{1} 分钟", "MinSec", "{1} 分 {2} 秒",
            "DeleteConfirm", "删除 {1} 吗? 它在任务计划程序里的任务会被删除, 登录后不再启动这个程序。",
            "NoScheduler", "无法使用任务计划程序: {1}",
            "SyncFailed", "有些任务没能在任务计划程序里设置好:`n`n{1}",
            "AlsoAutoStarts", "{1} 自己也会在登录后启动 ({2}), 这样它会立即启动, 不会等延迟时间。`n`n关掉它自己的启动吗? 和在 任务管理器 → 启动应用 里禁用一样; 以后删除这一项时会问要不要重新打开",
            "AlsoAutoStartsAdmin", "{1} 会为所有用户在登录后自己启动 ({2}), 这样它会立即启动, 不会等延迟时间。关掉它需要管理员权限: 请在 任务管理器 → 启动应用 里禁用。`n`n仍然保存吗?",
            "PickTitle", "启动项",
            "PickDesc", "Windows 登录时启动的程序 (和 任务管理器 → 启动应用 一样)。选一个改成延迟启动: 点 确定 时会关掉它自己的启动",
            "NoStartupApps", "没有其他在登录时启动的程序",
            "Where.HKCU", "Run (当前用户)", "Where.HKLM", "Run (所有用户)", "Where.HKLM32", "Run (所有用户, 32 位)",
            "Where.User", "启动文件夹 (当前用户)", "Where.Common", "启动文件夹 (所有用户)",
            "TurnOffFailed", "任务已经设置好, 但没能关掉这个程序自己的启动:`n`n{1}",
            "TurnOffManually", "延迟启动的任务已经设置好。这个程序还会为所有用户在登录时自己启动, 关掉它需要管理员权限:`n`n{1}`n`n打开 任务管理器 → 启动应用 去禁用吗?",
            "TurnBackOn", "{1} 自己在登录时的启动已被关掉:`n`n{2}`n`n让它重新在登录时自己启动吗?"))
}

;===============================================================================
; DelayedStartWindow - the list, the edit window and the startup apps to choose from
;===============================================================================
class DelayedStartWindow {
    static Gui := "", List := "", Items := [], ExitOnClose := false

    static Main() {
        SetTitleMatchMode(3)                                                ; the window title exactly
        if WinExist(DelayedStart.T("Title") " ahk_class AutoHotkeyGUI") {   ; already open: to the front
            WinActivate()
            ExitApp()
        }
        A_IconHidden := true
        DelayedStartWindow.ExitOnClose := true
        DelayedStartWindow.Show()
    }

    static FontName => DelayedStart.Lang = "zh" ? "Microsoft YaHei UI" : "Segoe UI"

    static Show() {
        g := Gui("-MaximizeBox", DelayedStart.T("Title"))
        g.SetFont("s9", DelayedStartWindow.FontName)
        g.MarginX := 14, g.MarginY := 12
        width := 640
        g.AddLink("w" width, DelayedStart.T("Note", '<a href="taskschd.msc">', '<a href="taskmgr.exe /0 /startup">', "</a>"))   ; a link without a Click callback is run
        list := g.AddListView("xm y+10 w" width " r12 -Multi NoSortHdr Grid"
            , [DelayedStart.T("Col.Name"), DelayedStart.T("Col.Program"), DelayedStart.T("Col.Arguments"), DelayedStart.T("Col.Delay"), DelayedStart.T("Col.Enabled")])
        for index, w in [130, 250, 110, 80, 66]
            list.ModifyCol(index, w)
        list.OnEvent("DoubleClick", (ctrl, row) => row ? DelayedStartWindow.Edit(row) : DelayedStartWindow.Add())
        list.OnNotify(-155, (ctrl, lParam) => (NumGet(lParam, 3 * A_PtrSize, "UShort") = 0x2E && DelayedStartWindow.Delete()))   ; LVN_KEYDOWN, Delete
        g.AddButton("xm y+12 w90", DelayedStart.T("Add")).OnEvent("Click", (*) => DelayedStartWindow.Add())
        g.AddButton("x+8 yp w90", DelayedStart.T("Edit")).OnEvent("Click", (*) => DelayedStartWindow.Edit())
        g.AddButton("x+8 yp w90", DelayedStart.T("Delete")).OnEvent("Click", (*) => DelayedStartWindow.Delete())
        g.AddButton("x+8 yp w120", DelayedStart.T("StartupApps")).OnEvent("Click", (*) => DelayedStartWindow.AddStartupApp())
        g.AddButton("xm+" (width - 90) " yp w90", DelayedStart.T("Close")).OnEvent("Click", (*) => DelayedStartWindow.Close())
        g.OnEvent("Close", (*) => DelayedStartWindow.Close())
        g.OnEvent("Escape", (*) => DelayedStartWindow.Close())
        DelayedStartWindow.Gui := g, DelayedStartWindow.List := list
        HotIfWinActive("ahk_id " g.Hwnd)
        Hotkey("F1", (*) => Run(DelayedStart.HelpUrl))
        HotIfWinActive()
        DelayedStartWindow.Reload()
        g.Show()
    }

    static Close() {
        if DelayedStartWindow.ExitOnClose
            ExitApp()
        DelayedStartWindow.Gui.Destroy()
        DelayedStartWindow.Gui := ""
    }

    ; Lists the tasks in Task Scheduler again
    static Reload() {
        items := []
        try items := DelayedStart.Items()
        catch as e
            DelayedStartWindow.Message(DelayedStart.T("NoScheduler", e.Message), "Icon!")
        DelayedStartWindow.Items := items
        list := DelayedStartWindow.List
        list.Delete()
        for item in items
            list.Add(, item["Name"], item["Target"], item["Arguments"], DelayedStart.DelayText(item["Delay"]), DelayedStart.T(item["Enabled"] ? "On" : "Off"))
        if items.Length
            list.Modify(1, "Select Focus")
    }

    static Message(text, options := "Iconi") => MsgBox(text, DelayedStart.T("Title"), options " Owner" DelayedStartWindow.Gui.Hwnd)

    static SelectedRow() => DelayedStartWindow.List.GetNext(0, "F")

    static Add(item := "") {
        if !IsObject(edited := DelayedStartWindow.EditDialog(IsObject(item) ? item : DelayedStart.NewItem()))
            return
        items := DelayedStartWindow.Items.Clone()
        items.Push(edited.Item)
        DelayedStartWindow.Commit(items, edited.TurnOff)
    }

    static Edit(row := 0) {
        if !(row := row || DelayedStartWindow.SelectedRow())
            return
        if !IsObject(edited := DelayedStartWindow.EditDialog(DelayedStartWindow.Items[row].Clone()))
            return
        items := DelayedStartWindow.Items.Clone()
        items[row] := edited.Item
        DelayedStartWindow.Commit(items, edited.TurnOff)
    }

    ; Deletes the selected entry; when its program's own start is turned off (as Startup Apps... does), offers to turn it on again
    static Delete() {
        if !(row := DelayedStartWindow.SelectedRow())
            return
        item := DelayedStartWindow.Items[row]
        if (DelayedStartWindow.Message(DelayedStart.T("DeleteConfirm", item["Name"]), "YesNo Icon?") != "Yes")
            return
        items := DelayedStartWindow.Items.Clone()
        items.RemoveAt(row)
        if !DelayedStartWindow.Apply(items)
            return
        leaf := StrLower(DelayedStart.Leaf(item["Target"]))
        for other in items                                                  ; another entry still starts the same program
            if (StrLower(DelayedStart.Leaf(other["Target"])) = leaf)
                return
        if !IsObject(entry := DelayedStart.OffEntry(item["Target"]))
            return
        if (DelayedStartWindow.Message(DelayedStart.T("TurnBackOn", DelayedStart.Leaf(entry.Program), DelayedStart.SourceText(entry.Source)), "YesNo Icon?") != "Yes")
            return
        try DelayedStart.SetEntryOff(entry.Source, false)
        catch as e
            DelayedStartWindow.Message(e.Message, "Icon!")
    }

    ; Startup Apps...: one of the programs that start by themselves, made into an entry (the edit window opens with it filled in)
    static AddStartupApp() {
        if IsObject(item := DelayedStartWindow.PickStartupApp())
            DelayedStartWindow.Add(item)
    }

    ; Sets the tasks up for items, then turns off the startup app turnOff (only when every task was set up: otherwise the program might
    ; not start at all)
    static Commit(items, turnOff) {
        if !DelayedStartWindow.Apply(items) || !IsObject(turnOff)
            return
        where := DelayedStart.SourceText(turnOff.Source)
        if turnOff.Writable {
            try DelayedStart.SetEntryOff(turnOff.Source, true)
            catch as e
                DelayedStartWindow.Message(DelayedStart.T("TurnOffFailed", where " - " e.Message), "Icon!")
        } else if (DelayedStartWindow.Message(DelayedStart.T("TurnOffManually", where), "YesNo Iconi") = "Yes")
            try Run("taskmgr.exe /0 /startup")
    }

    ; Makes the tasks match items, shows the problems, lists the tasks again; true when every task was set up
    static Apply(items) {
        problems := DelayedStart.Sync(items)
        if (problems != "")
            DelayedStartWindow.Message(DelayedStart.T("SyncFailed", problems), "Icon!")
        DelayedStartWindow.Reload()
        return problems = ""
    }

    ; The edit window: {Item, TurnOff: the startup app to turn off afterwards or ""} on OK, "" on Cancel
    static EditDialog(item) {
        owner := DelayedStartWindow.Gui
        state := {Result: ""}
        g := Gui("+Owner" owner.Hwnd " -MinimizeBox", DelayedStart.T("Editor.Title"))
        g.SetFont("s9", DelayedStartWindow.FontName)
        g.MarginX := 14, g.MarginY := 12
        label := 110, box := 380, hintOptions := "xm+" (label + 6) " y+3 w" box " cGray"
        g.AddText("xm y+4 w" label, DelayedStart.T("Col.Name"))
        name := g.AddEdit("x+6 yp-3 w" box, DelayedStart._Get(item, "Name"))
        g.AddText(hintOptions, DelayedStart.T("Field.Name"))
        g.AddText("xm y+12 w" label, DelayedStart.T("Col.Program"))
        target := g.AddEdit("x+6 yp-3 w" (box - 90), DelayedStart._Get(item, "Target"))
        g.AddButton("x+6 yp-1 w84", DelayedStart.T("Browse")).OnEvent("Click", Browse)
        g.AddText(hintOptions, DelayedStart.T("Field.Target"))
        g.AddText("xm y+12 w" label, DelayedStart.T("Col.Arguments"))
        arguments := g.AddEdit("x+6 yp-3 w" box, DelayedStart._Get(item, "Arguments"))
        g.AddText(hintOptions, DelayedStart.T("Field.Arguments"))
        g.AddText("xm y+12 w" label, DelayedStart.T("Field.DelaySeconds"))
        delay := DelayedStart._Get(item, "Delay", 300)
        delayBox := g.AddEdit("x+6 yp-3 w90 Number", (IsInteger(delay) && delay >= 0) ? delay : 0)
        g.AddText(hintOptions, DelayedStart.T("Field.Delay"))
        enabled := g.AddCheckBox("xm+" (label + 6) " y+12" (DelayedStart._Get(item, "Enabled", 1) ? " Checked" : ""), DelayedStart.T("Field.Enabled"))
        g.AddButton("xm+" (label + 6 + box - 170) " y+16 w80 Default", DelayedStart.T("OK")).OnEvent("Click", Save)
        g.AddButton("x+10 yp w80", DelayedStart.T("Cancel")).OnEvent("Click", (*) => g.Destroy())
        g.OnEvent("Escape", (*) => g.Destroy())
        g.OnEvent("Close", (*) => g.Destroy())
        owner.Opt("+Disabled")
        g.Show()
        (DelayedStart._Get(item, "Target") = "" ? target : name).Focus()
        WinWaitClose(g.Hwnd)
        owner.Opt("-Disabled")
        try WinActivate("ahk_id " owner.Hwnd)
        return state.Result

        Browse(*) {
            g.Opt("+OwnDialogs")
            chosen := FileSelect(3, DelayedStart.ProgramPath(target.Value), , DelayedStart.T("Programs") " (*.exe; *.bat; *.cmd; *.lnk)")
            if (chosen = "")
                return
            target.Value := chosen
            if (Trim(name.Value) = "")
                name.Value := RegExReplace(DelayedStart.Leaf(chosen), "\.[^.]*$")
        }

        ; OK: a program is needed. Does it also start by itself? Turned off afterwards when chosen from the startup apps (Source), or
        ; when the user agrees; one for all users can't be turned off without an administrator: save anyway?
        Save(*) {
            g.Opt("+OwnDialogs")
            program := Trim(target.Value, "`" `t")
            if (program = "") {
                MsgBox(DelayedStart.T("NoProgram"), DelayedStart.T("Editor.Title"), "Iconi")
                return target.Focus()
            }
            edited := Map("Name", Trim(name.Value), "Target", program, "Arguments", Trim(arguments.Value)
                , "Delay", IsInteger(delayBox.Value) ? Integer(delayBox.Value) : 0, "Enabled", enabled.Value ? 1 : 0)
            turnOff := ""
            if IsObject(entry := DelayedStart.AutoStartEntry(program)) {
                leaf := DelayedStart.Leaf(DelayedStart.ProgramPath(program)), where := DelayedStart.SourceText(entry.Source)
                if (DelayedStart._Get(item, "Source") = entry.Source)
                    turnOff := entry
                else if entry.Writable {
                    if (MsgBox(DelayedStart.T("AlsoAutoStarts", leaf, where), DelayedStart.T("Editor.Title"), "YesNo Icon?") = "Yes")
                        turnOff := entry
                } else if (MsgBox(DelayedStart.T("AlsoAutoStartsAdmin", leaf, where), DelayedStart.T("Editor.Title"), "YesNo Icon!") != "Yes")
                    return
            }
            state.Result := {Item: edited, TurnOff: turnOff}
            g.Destroy()
        }
    }

    ; Startup Apps...: picks one of the startup apps whose program isn't in the list yet; returns the new entry (Map), "" when none
    static PickStartupApp() {
        owner := DelayedStartWindow.Gui
        used := Map(), entries := []
        for item in DelayedStartWindow.Items
            used[StrLower(DelayedStart.Leaf(item["Target"]))] := true
        for entry in DelayedStart.AutoStartEntries(true)
            if !used.Has(StrLower(DelayedStart.Leaf(entry.Program)))
                entries.Push(entry)
        title := DelayedStart.T("PickTitle")
        if !entries.Length {
            DelayedStartWindow.Message(DelayedStart.T("NoStartupApps"))
            return ""
        }
        state := {Result: ""}
        g := Gui("+Owner" owner.Hwnd " -MinimizeBox", title)
        g.SetFont("s9", DelayedStartWindow.FontName)
        g.MarginX := 14, g.MarginY := 12
        width := 640
        g.AddText("w" width, DelayedStart.T("PickDesc"))
        list := g.AddListView("xm y+10 w" width " h260 -Multi NoSort Grid"
            , [DelayedStart.T("Col.Name"), DelayedStart.T("Col.Command"), DelayedStart.T("Col.Where"), DelayedStart.T("Col.Enabled")])
        for entry in entries
            list.Add(, RegExReplace(entry.Name, "i)\.lnk$"), entry.Command, DelayedStart.T("Where." entry.Scope), DelayedStart.T(entry.Off ? "Off" : "On"))
        for index, w in [130, 300, 140, 66]
            list.ModifyCol(index, w)
        list.Modify(1, "Select Focus")
        g.AddButton("xm+" (width - 170) " y+14 w80 Default", DelayedStart.T("OK")).OnEvent("Click", Choose)
        g.AddButton("x+10 yp w80", DelayedStart.T("Cancel")).OnEvent("Click", (*) => g.Destroy())
        list.OnEvent("DoubleClick", Choose)
        g.OnEvent("Escape", (*) => g.Destroy())
        g.OnEvent("Close", (*) => g.Destroy())
        owner.Opt("+Disabled")
        g.Show()
        WinWaitClose(g.Hwnd)
        owner.Opt("-Disabled")
        try WinActivate("ahk_id " owner.Hwnd)
        return state.Result

        Choose(*) {
            if !(row := list.GetNext())
                return
            state.Result := DelayedStart.ItemFromEntry(entries[row])
            g.Destroy()
        }
    }
}

;===============================================================================
; TaskScheduler - "run at my sign-in, after a delay" tasks in Windows Task Scheduler
;-------------------------------------------------------------------------------
; Through the Task Scheduler 2.0 COM interface (Schedule.Service), not schtasks.exe: schtasks /SC ONLOGON needs an
; administrator, while a standard user may register a logon task for their own account through COM. Checked on Windows
; by ALTRun's .github/workflows/delayed-start.yml (as a standard user, Tests\Integration\DelayedStartCheck.ahk). Also found there:
; a task folder another user made can be used, and that user's tasks in it can't be seen, so each user's tasks stay apart.
;
; The tasks run the program in the user's session with the user's normal rights (an interactive token: no password is
; stored, nothing runs while the user is signed out). Settings that differ from Task Scheduler's defaults, which suit
; maintenance jobs rather than programs people use:
;   - starts on battery and keeps running when the power is unplugged (default: neither)
;   - no time limit (default: stopped after 3 days)
;   - normal priority, 5 (default 7: below normal CPU, I/O and memory priority for as long as the program runs)
;
;   TaskScheduler.Tasks("\ALTRun\me")                           [{Name, Path, Arguments, WorkingDirectory, Delay, Enabled}...], [] when no such folder
;   TaskScheduler.RegisterLogonTask("\ALTRun\me", "OneDrive", {Path, Arguments, WorkingDirectory, Delay, Enabled, Description})
;       Delay in seconds; creates the folder when missing, replaces a task of the same name
;   TaskScheduler.Delete("\ALTRun\me", "OneDrive")
;   TaskScheduler.DeleteFolder("\ALTRun\me")                    only when empty; its parent too when that is empty then
;   Each takes an optional last argument service (TaskScheduler.Service()): several calls in a row share one connection
;   TaskScheduler.Duration(330) = "PT5M30S"    TaskScheduler.Seconds("PT5M30S") = 330
; Errors (no access, Task Scheduler service not running) are thrown as they come from COM.
;===============================================================================
class TaskScheduler {
    static TASK_TRIGGER_LOGON := 9, TASK_ACTION_EXEC := 0, TASK_CREATE_OR_UPDATE := 6, TASK_LOGON_INTERACTIVE_TOKEN := 3
    static TASK_INSTANCES_IGNORE_NEW := 2, Priority := 5

    ; DOMAIN\user (or PC\user) of the user running the script, as Task Scheduler names accounts
    static UserId() {
        static id := ""
        if (id = "") {
            size := 256, buf := Buffer(size * 2, 0)
            if DllCall("secur32\GetUserNameExW", "Int", 2, "Ptr", buf, "UInt*", &size)   ; NameSamCompatible
                id := StrGet(buf)
            else
                id := EnvGet("USERDOMAIN") "\" A_UserName
        }
        return id
    }

    static Service() {
        service := ComObject("Schedule.Service")
        service.Connect()
        return service
    }

    ; The folder at folderPath ("\ALTRun\me"), or "" when it doesn't exist (create: make it, with the folders above it)
    static Folder(service, folderPath, create := false) {
        try return service.GetFolder(folderPath)
        if !create
            return ""
        return service.GetFolder("\").CreateFolder(LTrim(folderPath, "\"))
    }

    static Tasks(folderPath, service := "") {
        list := []
        if !IsObject(folder := TaskScheduler.Folder(IsObject(service) ? service : TaskScheduler.Service(), folderPath))
            return list
        for task in folder.GetTasks(1) {                                    ; 1 = hidden ones too
            item := {Name: task.Name, Enabled: task.Enabled ? 1 : 0, Path: "", Arguments: "", WorkingDirectory: "", Delay: -1}
            definition := task.Definition
            if (definition.Actions.Count >= 1 && definition.Actions.Item(1).Type = TaskScheduler.TASK_ACTION_EXEC) {
                action := definition.Actions.Item(1)
                item.Path := action.Path, item.Arguments := action.Arguments, item.WorkingDirectory := action.WorkingDirectory
            }
            if (definition.Triggers.Count = 1 && definition.Triggers.Item(1).Type = TaskScheduler.TASK_TRIGGER_LOGON)
                item.Delay := TaskScheduler.Seconds(definition.Triggers.Item(1).Delay)   ; -1 = not one logon trigger: never taken as up to date
            list.Push(item)
        }
        return list
    }

    static RegisterLogonTask(folderPath, name, spec, service := "") {
        service := IsObject(service) ? service : TaskScheduler.Service()
        folder := TaskScheduler.Folder(service, folderPath, true)
        user := TaskScheduler.UserId()
        definition := service.NewTask(0)
        definition.RegistrationInfo.Author := "ALTRun"
        definition.RegistrationInfo.Description := spec.Description
        trigger := definition.Triggers.Create(TaskScheduler.TASK_TRIGGER_LOGON)
        trigger.UserId := user                                              ; only this user's sign-in (any user's needs an administrator)
        trigger.Delay := TaskScheduler.Duration(spec.Delay)
        action := definition.Actions.Create(TaskScheduler.TASK_ACTION_EXEC)
        action.Path := spec.Path
        action.Arguments := spec.Arguments
        action.WorkingDirectory := spec.WorkingDirectory
        definition.Principal.UserId := user
        definition.Principal.LogonType := TaskScheduler.TASK_LOGON_INTERACTIVE_TOKEN
        definition.Principal.RunLevel := 0                                  ; TASK_RUNLEVEL_LUA: the user's normal rights
        settings := definition.Settings
        settings.Enabled := spec.Enabled ? true : false
        settings.MultipleInstances := TaskScheduler.TASK_INSTANCES_IGNORE_NEW
        settings.DisallowStartIfOnBatteries := false
        settings.StopIfGoingOnBatteries := false
        settings.ExecutionTimeLimit := "PT0S"
        settings.Priority := TaskScheduler.Priority
        folder.RegisterTaskDefinition(name, definition, TaskScheduler.TASK_CREATE_OR_UPDATE, , , TaskScheduler.TASK_LOGON_INTERACTIVE_TOKEN)
    }

    static Delete(folderPath, name, service := "") {
        if IsObject(folder := TaskScheduler.Folder(IsObject(service) ? service : TaskScheduler.Service(), folderPath))
            folder.DeleteTask(name, 0)
    }

    ; Deletes the folder if it holds no tasks and no folders (tasks of other users can't be seen: Task Scheduler then refuses, which is fine)
    static DeleteFolder(folderPath, service := "") {
        service := IsObject(service) ? service : TaskScheduler.Service()
        Loop 2 {
            if !IsObject(folder := TaskScheduler.Folder(service, folderPath))
                return
            if (folder.GetTasks(1).Count || folder.GetFolders(0).Count)
                return
            SplitPath(folderPath, &leaf, &parent)
            if (parent = "" || parent = "\")
                parent := "\"
            try service.GetFolder(parent).DeleteFolder(leaf, 0)
            catch
                return
            if (parent = "\")
                return
            folderPath := parent
        }
    }

    ; Seconds -> an XML duration as Task Scheduler takes it ("PT5M", "PT1H30S"); 0 -> "" (no delay)
    static Duration(seconds) {
        seconds := Max(0, Integer(seconds))
        if !seconds
            return ""
        text := "PT"
        if (hours := seconds // 3600)
            text .= hours "H"
        if (minutes := Mod(seconds, 3600) // 60)
            text .= minutes "M"
        if (rest := Mod(seconds, 60))
            text .= rest "S"
        return text
    }

    ; An XML duration -> seconds ("" -> 0); days count, months and years don't occur in a delay; -1 when it can't be read
    static Seconds(duration) {
        if (duration = "")
            return 0
        if !RegExMatch(duration, "i)^P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$", &m)
            return -1
        return (m[1] = "" ? 0 : m[1]) * 86400 + (m[2] = "" ? 0 : m[2]) * 3600 + (m[3] = "" ? 0 : m[3]) * 60 + (m[4] = "" ? 0 : m[4])
    }
}
