; @altrun.title          Environment Variables
; @altrun.keyword        envvar
; @altrun.mode           filter
; @altrun.argument       part of a name
; @altrun.description    Live results: the environment variables whose name has what you type; Enter copies the value
; @altrun.description_zh 实时结果: 名字含有输入文字的环境变量, 回车复制变量值
; @altrun.network        none
;
; Type "envvar path" in ALTRun. While you type, ALTRun runs this script with the text (A_Args[1]) and shows what it
; prints: one result per line, title<Tab>subtitle<Tab>arg<Tab>type (or JSON, see the Scripts page in ALTRun's wiki).
#Requires AutoHotkey v2.0
#NoTrayIcon

query := A_Args.Length ? A_Args[1] : ""
block := DllCall("GetEnvironmentStringsW", "Ptr"), at := block, out := ""
while ((entry := StrGet(at, "UTF-16")) != "") {
    at += (StrLen(entry) + 1) * 2
    split := InStr(entry, "=", , 2)                                  ; from the 2nd character: names like "=C:" start with one
    name := SubStr(entry, 1, split - 1), value := StrReplace(SubStr(entry, split + 1), "`t", " ")
    if (split > 1 && SubStr(name, 1, 7) != "ALTRUN_" && (query = "" || InStr(name, query)))
        out .= name "`t" value "`t" value "`ttext`n"
}
DllCall("FreeEnvironmentStringsW", "Ptr", block)
FileAppend(out, "*", "UTF-8-RAW")
