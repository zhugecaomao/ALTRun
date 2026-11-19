;===============================================================================
; DelayedStartTests.ahk - the script library's Delayed Start.ahk (docs\scripts): the tasks the list asks for, startup apps, the text
;-------------------------------------------------------------------------------
; Included by Tests\RunTests.ahk, which also #Includes the script: it opens its window only when run as a program.
; Every static method whose name doesn't start with "_" is a test; each one must put back the shared state it changes.
;===============================================================================
class DelayedStartTests {
    ; The tasks the list asks for and what to change in Task Scheduler (DelayedStart, TaskScheduler).
    ; Registering real tasks is checked on Windows as a standard user by .github/workflows/delayed-start.yml
    static Tasks() {
        eq := (n, a, e) => TestRunner.Equal("DelayedStart." n, a, e)
        lang := DelayedStart.Lang, DelayedStart.Lang := "en"
        try DelayedStartTests._Tasks(eq)
        finally DelayedStart.Lang := lang
    }

    static _Tasks(eq) {
        eq("not in ALTRun's settings any more", AppSettings.Defaults().Has("DelayedStart") "|" AppSettings.ResetData(Map("DelayedStart", [DelayedStart.NewItem()])).Has("DelayedStart"), "0|0")
        eq("not run as a program when included", DelayedStartWindow.ExitOnClose, false)
        eq("new entry", JSON.Stringify(DelayedStart.NewItem()), JSON.Stringify(Map("Name", "", "Target", "", "Arguments", "", "Delay", 300, "Enabled", 1)))

        ; durations as Task Scheduler takes and gives them
        durations := ""
        for seconds in [0, 30, 300, 330, 3600, 3630, 5400]
            durations .= TaskScheduler.Duration(seconds) " "
        eq("duration", durations, " PT30S PT5M PT5M30S PT1H PT1H30S PT1H30M ")
        if (EnvGet("GITHUB_ACTIONS") = "true") {                            ; Windows CI: what one connection to Task Scheduler costs (Sync now makes one, not one per task)
            started := A_TickCount
            Loop 5
                TaskScheduler.Service()
            FileAppend("NOTE DelayedStart: a Task Scheduler connection takes " Round((A_TickCount - started) / 5, 1) " ms`n", "*")
        }
        eq("seconds", TaskScheduler.Seconds("") "|" TaskScheduler.Seconds("PT5M") "|" TaskScheduler.Seconds("PT1H0M30S") "|" TaskScheduler.Seconds("P1DT1S") "|" TaskScheduler.Seconds("5 min"), "0|300|3630|86401|-1")
        eq("delay text", DelayedStart.DelayText(0) "|" DelayedStart.DelayText(45) "|" DelayedStart.DelayText(300) "|" DelayedStart.DelayText(90) "|" DelayedStart.DelayText(""), "At once|45 s|5 min|1 min 30 s|At once")
        eq("safe names", DelayedStart.SafeName(' One/Drive: "sync" ') "|" DelayedStart.SafeName("a<b>c|d*e?f\g.") "|" DelayedStart.SafeName("..."), "One_Drive_ _sync_|a_b_c_d_e_f_g|")

        ; the tasks the list asks for
        items := [Map("Name", "OneDrive", "Target", "C:\Users\Me\AppData\Local\Microsoft\OneDrive\OneDrive.exe", "Arguments", " /background ", "Delay", 300, "Enabled", 1)
                , Map("Name", "OneDrive", "Target", "C:\Program Files\Tool\tool.exe", "Arguments", "", "Delay", "x", "Enabled", 0)
                , Map("Name", "", "Target", "C:\Apps\Sync It.exe", "Arguments", "-q", "Delay", 60, "Enabled", 1)
                , Map("Name", "No program", "Target", "  ", "Delay", 60, "Enabled", 1)
                , Map("Name", "Old entry", "Target", "C:\Apps\old.exe")]      ; keys added later missing: the defaults
        desired := DelayedStart.Desired(items)
        eq("entries with a program", desired.Length, 4)
        if (desired.Length = 4) {
            first := desired[1].Spec, second := desired[2].Spec, third := desired[3].Spec, fourth := desired[4].Spec
            eq("names unique, from the file when empty", desired[1].Name "|" desired[2].Name "|" desired[3].Name "|" desired[4].Name, "OneDrive|OneDrive (2)|Sync It.exe|Old entry")
            eq("first", first.Path "|" first.Arguments "|" first.WorkingDirectory "|" first.Delay "|" first.Enabled
                , "C:\Users\Me\AppData\Local\Microsoft\OneDrive\OneDrive.exe|/background|C:\Users\Me\AppData\Local\Microsoft\OneDrive|300|1")
            eq("path with spaces quoted, bad delay = 0, off", second.Path "|" second.WorkingDirectory "|" second.Delay "|" second.Enabled, '"C:\Program Files\Tool\tool.exe"|C:\Program Files\Tool|0|0')
            eq("description", InStr(third.Description, "Delayed Start.ahk") > 0, true)
            eq("missing keys", fourth.Arguments "|" fourth.Delay "|" fourth.Enabled, "|0|1")
        }
        eq("environment variables", DelayedStart.ProgramPath("%SystemRoot%\notepad.exe"), EnvGet("SystemRoot") "\notepad.exe")
        eq("relative to the script's folder", DelayedStart.ProgramPath("Tools\x.exe"), Path.Full("Tools\x.exe", A_ScriptDir))
        eq("user folders shown short", DelayedStart.ShortPath(EnvGet("LocalAppData") "\Microsoft\OneDrive\OneDrive.exe") "|" DelayedStart.ShortPath("C:\Tools\x.exe")
            , "%LocalAppData%\Microsoft\OneDrive\OneDrive.exe|C:\Tools\x.exe")
        eq("bare name found in PATH", StrLower(DelayedStart.ProgramPath("cmd.exe")), StrLower(EnvGet("SystemRoot") "\System32\cmd.exe"))

        ; what to change: tasks already right are left alone, others registered, those no entry asks for deleted
        task(name, path, args, dir, delay, enabled) => {Name: name, Path: path, Arguments: args, WorkingDirectory: dir, Delay: delay, Enabled: enabled}
        existing := [task("onedrive", "C:\Users\Me\AppData\Local\Microsoft\OneDrive\OneDrive.exe", "/background", "C:\Users\Me\AppData\Local\Microsoft\OneDrive", 300, 1)   ; right (case of the name aside)
                   , task("OneDrive (2)", '"C:\Program Files\Tool\tool.exe"', "", "C:\Program Files\Tool", 0, 1)                                      ; enabled: changed
                   , task("Sync It.exe", '"C:\Apps\Sync It.exe"', "-q", "C:\Apps", 120, 1)                                                            ; delay: changed
                   , task("Removed", "C:\x.exe", "", "C:\", 60, 1)]
        plan := DelayedStart.Plan(desired, existing)
        names := ""
        for entry in plan.Register
            names .= entry.Name ";"
        eq("register", names, "OneDrive (2);Sync It.exe;Old entry;")
        eq("delete", plan.Delete.Length = 1 ? plan.Delete[1] : plan.Delete.Length, "Removed")
        eq("quotes don't count", DelayedStart.Same(desired[3].Spec, task("x", "C:\Apps\Sync It.exe", "-q", "C:\Apps", 60, 1)), true)
        eq("an unreadable trigger never counts as right", DelayedStart.Same(desired[1].Spec, task("x", desired[1].Spec.Path, "/background", desired[1].Spec.WorkingDirectory, -1, 1)), false)
        plan := DelayedStart.Plan([], existing)
        eq("empty list: delete all", plan.Register.Length "|" plan.Delete.Length, "0|4")

        ; programs that also start by themselves
        eq("command program", DelayedStart.CommandProgram('"C:\Program Files\App\app.exe" /min') "|" DelayedStart.CommandProgram("C:\Users\Me\OneDrive.exe /background")
            . "|" DelayedStart.CommandProgram("C:\Program Files\x y\z.exe -a") "|" DelayedStart.CommandProgram("rundll32 shell32.dll"), "C:\Program Files\App\app.exe|C:\Users\Me\OneDrive.exe|C:\Program Files\x y\z.exe|rundll32")
        eq("command arguments", DelayedStart.CommandArguments('"C:\Program Files\App\app.exe" /min  -x') "|" DelayedStart.CommandArguments("C:\Users\Me\OneDrive.exe /background")
            . "|" DelayedStart.CommandArguments('"C:\x\y.exe"') "|" DelayedStart.CommandArguments("rundll32 shell32.dll,Control_RunDLL"), "/min  -x|/background||shell32.dll,Control_RunDLL")

        ; startup apps turned into entries: Task Manager's on / off value, where they are, what changed in the list
        eq("approved data (as ALTRun's Lib\StartupApproved.ahk)", DelayedStart.ApprovedData(false) "|" StartupApproved.Data(true, 0x01DC4B2A3F5E6D70) "|" StartupApproved.Data(false) "|" DelayedStart.ApprovedData(true, 0x01DC4B2A3F5E6D70), "020000000000000000000000|03000000706D5E3F2A4BDC01|020000000000000000000000|03000000706D5E3F2A4BDC01")
        eq("off data", (DelayedStart.IsOffData("030000001122334455667788") ? 1 : 0) (DelayedStart.IsOffData("020000000000000000000000") ? 1 : 0)
            . (DelayedStart.IsOffData("060000000000000000000000") ? 1 : 0) (DelayedStart.IsOffData("01") ? 1 : 0) (DelayedStart.IsOffData("") ? 1 : 0) (DelayedStart.IsOffData("zz") ? 1 : 0), "100100")
        eq("today's time when off", SubStr(DelayedStart.ApprovedData(true), 1, 8) "|" StrLen(DelayedStart.ApprovedData(true)), "03000000|24")
        writable := ""
        for scope, where in DelayedStart.Scopes()
            writable .= scope "=" (where.Writable ? 1 : 0) " "
        eq("scopes, which this user can turn off", writable, "Common=0 HKCU=1 HKLM=0 HKLM32=0 User=1 ")
        order := ""
        for scope in DelayedStart.ScopeOrder
            order .= scope (DelayedStart.Scopes().Has(scope) ? "" : "?") " "
        eq("listed this user's first", order, "HKCU User HKLM HKLM32 Common ")
        eq("approved keys", DelayedStart.Scopes()["HKCU"].Approved "|" DelayedStart.Scopes()["User"].Approved
            , "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run|HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder")
        eq("where", DelayedStart.SourceText("HKCU|OneDrive") "|" DelayedStart.SourceText("Common|Teams.lnk"), "Run (you): OneDrive|Startup folder (all users): Teams.lnk")
        item := DelayedStart.ItemFromEntry({Source: "User|Sync Tool.lnk", Name: "Sync Tool.lnk", Program: "C:\Tools\sync.exe", Arguments: "--tray"})
        eq("entry from a startup app", item["Name"] "|" item["Target"] "|" item["Arguments"] "|" item["Delay"] "|" item["Enabled"] "|" item["Source"]
            , "Sync Tool|C:\Tools\sync.exe|--tray|300|1|User|Sync Tool.lnk")
        thrown := false
        try DelayedStart.SetEntryOff("Nowhere|x", true)
        catch ValueError
            thrown := true
        eq("unknown source refused", thrown, true)
    }

    ; English, or Simplified Chinese on a Chinese system (China, Singapore); every text in both; F1 opens the Chinese Wiki page in Chinese
    static Text() {
        eq := (n, a, e) => TestRunner.Equal("DelayedStartText." n, a, e)
        eq("Chinese systems", DelayedStart.IsChinese("0804") "|" DelayedStart.IsChinese("1004") "|" DelayedStart.IsChinese("0404") "|" DelayedStart.IsChinese("0409") "|" DelayedStart.IsChinese("0411"), "1|1|0|0|0")
        missing := ""
        for lang in ["en", "zh"]
            for key in DelayedStart.Strings[lang]
                if !DelayedStart.Strings[lang = "en" ? "zh" : "en"].Has(key)
                    missing .= lang ":" key " "
        eq("the same texts in both", missing, "")
        lang := DelayedStart.Lang
        try {
            DelayedStart.Lang := "zh"
            eq("Chinese", DelayedStart.T("Title") "|" DelayedStart.DelayText(90) "|" DelayedStart.SourceText("HKCU|OneDrive"), "延迟启动|1 分 30 秒|Run (当前用户): OneDrive")
            eq("help: Chinese page", DelayedStart.HelpUrl, "https://github.com/zhugecaomao/ALTRun/wiki/Delayed-Start-zh")
            DelayedStart.Lang := "en"
            eq("filled in", DelayedStart.T("DeleteConfirm", "OneDrive"), "Delete OneDrive? Its task is deleted from Task Scheduler, and the program no longer starts after signing in.")
            eq("help: English page", DelayedStart.HelpUrl, "https://github.com/zhugecaomao/ALTRun/wiki/Delayed-Start")
        } finally
            DelayedStart.Lang := lang
    }
}
