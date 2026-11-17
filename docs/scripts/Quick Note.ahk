; @altrun.title          Quick Note
; @altrun.keyword        note
; @altrun.argument       the note
; @altrun.mode           silent
; @altrun.description    Adds a line with the date and time to a Markdown notes file (Documents\Notes.md unless you choose another)
; @altrun.description_zh 把一行带日期时间的笔记追加到 Markdown 笔记文件 (默认 "文档\Notes.md", 可以改)
; @altrun.network        none
;
; Type "note call the client back tomorrow" in ALTRun and press Enter: "- 2026-11-17 16:30 call the client back tomorrow" is
; added at the end of the notes file, which is created when it doesn't exist yet. Open it like any other file (search "Notes.md").
; Another file: set NotesFile below, or the user environment variable QUICKNOTE_FILE to its full path.
#Requires AutoHotkey v2.0
#NoTrayIcon

NotesFile := ""                                     ; e.g. "D:\OneDrive\Notes.md"; empty: QUICKNOTE_FILE, else Documents\Notes.md
Chinese := RegExMatch(A_Language, "i)^(0804|0404|0C04|1004|1404)$") > 0
notes := NotesFile != "" ? NotesFile : EnvGet("QUICKNOTE_FILE") != "" ? EnvGet("QUICKNOTE_FILE") : A_MyDocuments "\Notes.md"
text := Trim(A_Args.Length ? A_Args[1] : "")
if (text = "") {
    FileAppend((Chinese ? "在 note 后面输入笔记内容" : "Type the note after 'note'") "`n", "**", "UTF-8-RAW")
    ExitApp(1)
}
try FileAppend("- " FormatTime(, "yyyy-MM-dd HH:mm") " " StrReplace(StrReplace(text, "`r"), "`n", " ") "`n", notes, "UTF-8-RAW")
catch as e {
    FileAppend((Chinese ? "不能写入 " : "Can't write to ") notes ": " e.Message "`n", "**", "UTF-8-RAW")
    ExitApp(1)
}
SplitPath(notes, &name)
FileAppend((Chinese ? "已记到 " : "Added to ") name, "*", "UTF-8-RAW")
