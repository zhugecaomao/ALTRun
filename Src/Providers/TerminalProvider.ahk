;===============================================================================
; TerminalProvider.ahk - 在终端里运行命令 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 输入 ">ipconfig /all" (前缀见 Features.Terminal.Prefix) 在终端里运行这条命令,
; 窗口保持打开。Shell 可选: cmd / powershell / pwsh / wt (Windows Terminal)。
;
; 用法 (其它模块里):
;   TerminalProvider.OpenAt(folder)          在指定文件夹打开终端
;   TerminalProvider.RunCommand(command)
;===============================================================================

class TerminalProvider {
    static Id := "Terminal"

    static Init() {
    }

    static Search(query) {
        options := AppSettings.Feature("Terminal")
        if !query.MatchPrefix(options["Prefix"], &command)
            return []
        icon := "res:imageres.dll,-5323"
        if (command = "")
            return [ResultItem(I18n.T("Terminal.Empty"), options["Prefix"] " ...", {Icon: icon, Valid: false, Score: 150})]
        return [ResultItem(I18n.T("Terminal.Run", command), options["Shell"], {
            Icon: icon, Score: 150, Uid: "terminal:" StrLower(command), Arg: command,
            OnRun: (*) => TerminalProvider.RunCommand(command),
            Actions: [ResultItem(I18n.T("Action.RunAsAdmin"), "", {Icon: "res:imageres.dll,-78", OnRun: (*) => TerminalProvider.RunCommand(command, true)})]
        })]
    }

    static RunCommand(command, asAdmin := false) {
        shell := AppSettings.Feature("Terminal")["Shell"]
        prefix := asAdmin ? "*RunAs " : ""
        switch shell, false {
            case "powershell": Run(prefix 'powershell.exe -NoExit -Command ' command, A_MyDocuments)
            case "pwsh"      : Run(prefix 'pwsh.exe -NoExit -Command ' command, A_MyDocuments)
            case "wt"        : Run(prefix 'wt.exe cmd /k ' command, A_MyDocuments)
            default          : Run(prefix A_ComSpec ' /k ' command, A_MyDocuments)
        }
    }

    static OpenAt(folder) {
        if (folder = "" || !DirExist(folder))
            folder := A_MyDocuments
        shell := AppSettings.Feature("Terminal")["Shell"]
        switch shell, false {
            case "powershell": Run("powershell.exe -NoExit", folder)
            case "pwsh"      : Run("pwsh.exe -NoExit", folder)
            case "wt"        : Run('wt.exe -d "' folder '"', folder)
            default          : Run(A_ComSpec, folder)
        }
    }

    ; 资源管理器 / Total Commander 当前打开的文件夹 (呼出 ALTRun 前的那个窗口)
    static OpenAtCurrentFolder() {
        folder := QuickSwitch.FolderOfWindow(App.PreviousWindow)
        if (folder = "")
            return MsgBox(I18n.T("Sys.NoFolder"), App.Name, 48)
        TerminalProvider.OpenAt(folder)
    }
}
