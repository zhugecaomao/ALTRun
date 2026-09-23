;===============================================================================
; CommandManager.ahk - 命令管理器窗口: 新建/编辑/删除命令 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 用法:
;   CommandManager.Open(Section, Type, Path, Desc, Rank, OriginCmd)
;       打开命令管理器窗口, 六个参数全部可省略(新建一条空的 UserCommand)。
;   CommandManager.Edit() / CommandManager.Delete() / CommandManager.UndoDelete()
;       对应主列表当前选中命令的编辑/删除/撤销删除。
;
; 注意: OpenCommandManager()/NewCommand()/EditCommand()/DelCommand()/
; UndoDelCommand() 仍以裸的全局函数外壳留在 ALTRun.ahk 里 - 它们全部登记在
; OptionsWindow 的 FuncList 数组里(可绑定自定义热键), 部分还是内置 Func 命令,
; CommandRunner.Execute() 靠 %cmdPath%() 按名字动态调用, 只认裸的全局函数名, 不认
; Class.Method, 详见那边的注释。
;===============================================================================

Class CommandManager {
    static G := ""    ; 命令管理器窗口的 Gui 对象 (原 Global g_CmdMgrGui)

    static Open(Section := "UserCommand", Type := "File", Path := "", Desc := "", Rank := 1, OriginCmd := "") {
        typeList := Array("File", "Dir", "CMD", "URL", "Func", "Clip")
        chooseIndex := Arr.IndexOf(Type, typeList)
        chooseIndex := chooseIndex ? chooseIndex : 1

        g_LOG.Debug("Starting Command Manager... Args=" Section "|" Type "|" Path "|" Desc "|" Rank)

        g := CommandManager.G := Gui(, g_LNG[700])
        g.SetFont("S9 Norm", "Microsoft Yahei")
        g.AddGroupBox("w600 h260", g_LNG[701])
        g.Add("Text", "x25 yp+30", g_LNG[702])
        g.AddDropDownList("x160 yp-5 w130 vType Choose" chooseIndex, typeList)
        g.Add("Text", "x315 yp+5", g_LNG[705])
        g.Add("Edit", "x435 yp-5 w130 Disabled vSection", Section)
        g.Add("Text", "x25 yp+60", g_LNG[703])
        g.Add("Edit", "x160 yp-5 w405 -WantReturn vPath", Path).Focus()
        g.AddButton("x575 yp w30 hp", "...").OnEvent("Click", (*) => CommandManager.PickTarget(g["Type"].Text))
        g.Add("Text", "x25 yp+80", g_LNG[704])
        g.AddEdit("x160 yp-5 w405 -WantReturn vDesc", Desc)
        g.AddText("x25 yp+60", g_LNG[706])
        g.AddEdit("x160 yp-5 w405 +Number vRank", Rank)
        g.AddButton("Default x420 w90", g_LNG[7]).OnEvent("Click", (*) => CommandManager.Save(Section, g["Type"].Text, g["Path"].Text, g["Desc"].Text, g["Rank"].Text, OriginCmd))
        g.AddButton("x521 yp w90", g_LNG[8]).OnEvent("Click", (*) => CommandManager.Close())
        g.OnEvent("Close", (*) => CommandManager.Close())
        g.OnEvent("Escape", (*) => CommandManager.Close())
        g.Show("Center")
    }

    static PickTarget(cmdType) {
        CommandManager.G.Opt("+OwnDialogs")                                 ; Make open dialog Modal

        if (cmdType = "Dir")
            cmdPath := DirSelect(, 3, 'Please select directory')
        else if (cmdType = "File")
            cmdPath := FileSelect(3, , , 'All Files (*.*)')
        else if (cmdType = "Clip")
            return Clip.EditClipText()                                      ; Clip uses a multi-line text editor instead of a file picker
        else
            return MsgBox("Path picker only supports File/Dir/Clip type.", g_LNG[700], 64)

        if (cmdPath != "")
            CommandManager.G["Path"].Value := cmdPath
    }

    static Save(section, cmdType, cmdPath, cmdDesc, cmdRank, originCmd) {
        CommandManager.G.Submit()
        validType := Map("File", 1, "Dir", 1, "CMD", 1, "URL", 1, "Func", 1, "Clip", 1)
        section := Trim(section)
        cmdType := Trim(cmdType)
        cmdPath := Trim(cmdPath)
        cmdDesc := Trim(cmdDesc)
        cmdRank := Trim(cmdRank)

        if !validType.Has(cmdType)
            return MsgBox("Invalid command type: " cmdType, g_LNG[820], 48)

        if (cmdPath = "") {
            return MsgBox(g_LNG[821], g_LNG[820], 64)
        }

        if (cmdType = "Clip") {
            ; The Path field already holds the escaped single-line form (EditClipText produced it),
            ; so only guard against stray real line breaks - never re-escape, that would double the backslashes.
            cmdPath := StrReplace(StrReplace(StrReplace(cmdPath, "`r`n", "\n"), "`n", "\n"), "`r", "\n")
            if (cmdDesc = "")
                return MsgBox("A Clip command needs a short name in the Description field, that is what you type to call it.", g_LNG[820], 48)
        }

        if (!IsInteger(cmdRank) || cmdRank <= 0)
            cmdRank := 1

        cmdLine := cmdType " | " cmdPath (cmdDesc != "" ? " | " cmdDesc : "")
        try {
            AppData.LoadAppData()
            if !g_CMDDATA.Has(section)
                section := "UserCommand"
            if (originCmd != "" && originCmd != cmdLine) {                  ; Drop the old key only when editing changed the command line
                for _, sec in ["DefaultCommand", "UserCommand", "Index"]
                    if g_CMDDATA[sec].Has(originCmd)
                        g_CMDDATA[sec].Delete(originCmd)
            }
            g_CMDDATA[section][cmdLine] := cmdRank + 0
            AppData.SaveAppData()
        } catch as e {
            MsgBox(g_LNG[822] e.Message, g_LNG[820], 64)
            return
        }
        MsgBox(g_LNG[823] section " ]`n`n" cmdLine " = " cmdRank, g_LNG[820], 64)
        CommandStore.LoadCommands()
    }

    static Close(*) {
        CommandManager.G.Destroy()
    }

    ; --- 主列表当前命令的编辑/删除/撤销 (由 ALTRun.ahk 里的裸函数外壳调用) ---

    static Edit() {
        currentCmd := g_RUNTIME["CurrentCommand"]
        if !currentCmd
            return MsgBox(g_LNG[810], g_TITLE, 64)                          ; 64 = Info icon

        AppData.LoadAppData()

        for _, section in ["DefaultCommand", "UserCommand", "Index"] {
            if !g_CMDDATA[section].Has(currentCmd)
                continue
            rank := g_CMDDATA[section][currentCmd]

            if IsInteger(rank) {
                parts := StrSplit(currentCmd, " | ")
                type := parts.Length >= 1 ? parts[1] : ""
                path := parts.Length >= 2 ? parts[2] : ""
                desc := parts.Length >= 3 ? parts[3] : ""

                g_LOG.Debug("EditCommand: Editing command=" currentCmd)
                CommandManager.Open(section, type, path, desc, rank, currentCmd)
                break
            }
        }
    }

    static Delete() {
        currentCmd := g_RUNTIME["CurrentCommand"]
        if !currentCmd
            return

        AppData.LoadAppData()

        for _, section in ["DefaultCommand", "UserCommand", "Index"] {
            if !g_CMDDATA[section].Has(currentCmd)
                continue

            result := MsgBox(g_LNG[800] section "]`n`n" currentCmd, g_LNG[801], 52) ; 52 = Yes/No + Question icon

            if result = "YES" {
                try {
                    rank := g_CMDDATA[section][currentCmd]
                    g_CMDDATA[section].Delete(currentCmd)
                    AppData.SaveAppData()
                    g_DELUNDO.Push(Map("Section", section, "CmdLine", currentCmd, "Rank", rank))  ; Ctrl+Z restores this
                    MsgBox(g_LNG[802] "`n`n" currentCmd "`n`n" g_LNG[811], g_TITLE, 64)  ; 64 = Info icon
                } catch as e {
                    MsgBox(g_LNG[803] "`n`n" currentCmd, g_TITLE, 48)       ; 48 = Error icon
                }
                break
            }
        }
        CommandStore.LoadCommands()
    }

    ; Ctrl+Z: restore the most recently deleted command (as many times in a row as things were deleted).
    ; In-memory only - once ALTRun is closed/reloaded, deleted commands can no longer be undone.
    static UndoDelete() {
        if (MainWindow.Gui.FocusedCtrl.ClassNN = "Edit1") {                       ; Typing in the input box: let the native "undo last edit" through instead
            SendInput("^z")
            return
        }

        if !g_DELUNDO.Length
            return MainWindow.SetStatus(g_LNG[812])

        entry := g_DELUNDO.Pop()
        AppData.LoadAppData()
        g_CMDDATA[entry["Section"]][entry["CmdLine"]] := entry["Rank"]
        AppData.SaveAppData()
        CommandStore.LoadCommands()
        MainWindow.SetStatus(g_LNG[813] " " entry["CmdLine"])
    }
}
