;===============================================================================
; NameCheck.ahk - finds values that would overwrite a method of the same name (AutoHotkey v2)
;-------------------------------------------------------------------------------
; AutoHotkey's names are case-insensitive: in a class with a method _Tick(), a value property _tick is the same property.
;   - class-level   static _tick := ...       fails while the class is set up, before OnError: AutoHotkey shows a dialog
;                                             ("Property is read-only"), so the program (or the test run) never starts
;   - in a method   this._tick := ...         fails when the line runs ("Property is read-only")
;                   ClassName._tick := ...
; /validate and #Warn don't report either, and a test can't: the first kind stops the script before any test runs.
; This reads the source files (it doesn't include them) and reports each such value. Methods inherited from another class
; in the checked files count too.
;
;   AutoHotkey64.exe /ErrorStdOut Tests\Tools\NameCheck.ahk [folder...]   ; default: Src, Lib, Tests
; Prints one line per problem; the exit code is the number of problems.
; It expects the project's style: 4-space indents, class members one level in, no continuation sections or block comments.
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut

OnError((e, *) => (FileAppend("NameCheck failed: " e.Message " (line " e.Line ")`n", "*"), ExitApp(255), 1))  ; printed, not a dialog
ExitApp(NameCheck.Main(A_Args))

class NameCheck {
    static Main(folders) {
        SplitPath(A_ScriptDir, , &tests), SplitPath(tests, , &root)       ; Tests\Tools -> the repository
        if !folders.Length
            folders := ["Src", "Lib", "Tests"]
        classes := Map(), classes.CaseSense := false
        for folder in folders {
            folderPath := InStr(folder, ":") ? folder : root "\" folder
            Loop Files folderPath "\*.ahk", "R"
                NameCheck.Scan(A_LoopFileFullPath, InStr(A_LoopFileFullPath, root "\") = 1 ? SubStr(A_LoopFileFullPath, StrLen(root) + 2) : A_LoopFileFullPath, classes)
        }
        problems := NameCheck.Problems(classes)
        for line in problems
            FileAppend(line "`n", "*")
        return problems.Length
    }

    ; Each class's methods and the values it sets: classes[name] = {Name, Parent, Methods: {Static, Instance}, Values: [...]}
    ; Methods: lowercase name -> {Name, Where}; a value: {Name, Static, Where, How}
    static Scan(filePath, shown, classes) {
        stack := []                                                         ; the classes this line is in: {Info, Indent}
        member := {Static: false}                                           ; the class member this line is in
        for lineNo, raw in StrSplit(FileRead(filePath, "UTF-8"), "`n", "`r") {
            code := NameCheck.Code(raw)
            if (Trim(code) = "")
                continue
            indent := StrLen(code) - StrLen(LTrim(code, " "))
            text := Trim(code)
            while (stack.Length && (indent < stack[-1].Indent || indent = stack[-1].Indent && SubStr(text, 1, 1) != "}"))
                stack.Pop()                                                 ; left the class (its "}" is at the class's own indent)
            where := shown ":" lineNo
            if RegExMatch(text, "i)^class\s+(\w+)(?:\s+extends\s+([\w.]+))?", &m) {
                if stack.Length && indent = stack[-1].Indent + 4            ; a nested class is a read-only static property
                    NameCheck._Method(stack[-1].Info, true, m[1], where)
                info := {Name: m[1], Parent: m[2], Where: where, Methods: {Static: NameCheck._Names(), Instance: NameCheck._Names()}, Values: []}
                classes[m[1]] := info
                stack.Push({Info: info, Indent: indent})
                continue
            }
            if !stack.Length
                continue
            info := stack[-1].Info
            if (indent = stack[-1].Indent + 4) {                            ; a class member
                isStatic := RegExMatch(text, "i)^static\s+")
                body := RegExReplace(text, "i)^static\s+")
                member := {Static: isStatic}
                if RegExMatch(body, "^(\w+)\s*\(", &m)                      ; a method: Name(...) {, Name(...) => ...
                    NameCheck._Method(info, isStatic, m[1], where)
                else if RegExMatch(body, "^(\w+)\s*(?:\[[^\]]*\])?\s*=>", &m)   ; a property with only a getter: read-only
                    NameCheck._Method(info, isStatic, m[1], where)
                else if RegExMatch(body, "^\w+\s*:=")                       ; values: name := x, other := y
                    for part in NameCheck.Split(body)
                        if RegExMatch(part, "^\s*(\w+)\s*:=", &m)
                            info.Values.Push({Name: m[1], Static: isStatic, Where: where, How: (isStatic ? "static " : "") m[1] " :="})
                continue
            }
            ; inside a member: this.name := ... (the member's own kind), ClassName.name := ... (static)
            pos := 1
            while (pos := RegExMatch(code, "(?<![\w.])(\w+)\.(\w+)\s*:=(?!=)", &m, pos)) {
                pos += m.Len
                if (m[1] = "this")
                    info.Values.Push({Name: m[2], Static: member.Static, Where: where, How: "this." m[2] " :="})
                else if (m[1] = info.Name)
                    info.Values.Push({Name: m[2], Static: true, Where: where, How: m[1] "." m[2] " :="})
            }
        }
    }

    static Problems(classes) {
        problems := []
        for , info in classes
            for value in info.Values {
                if !(found := NameCheck.Find(classes, info, value.Static, value.Name))
                    continue
                problems.Push(value.Where ": " info.Name ": " value.How " overwrites " found.Class "." found.Name
                    . " (" found.Where "): names are case-insensitive; rename the value (e.g. " value.Name "Timer, " value.Name "Fn) or the method")
            }
        return problems
    }

    ; The method a value would overwrite: the class's own, or one it inherits from a class in the checked files
    static Find(classes, info, isStatic, name) {
        seen := Map(), seen.CaseSense := false
        while (IsObject(info) && !seen.Has(info.Name)) {
            seen[info.Name] := true
            methods := isStatic ? info.Methods.Static : info.Methods.Instance
            if methods.Has(name) {
                found := methods[name]
                return {Class: info.Name, Name: found.Name, Where: found.Where}
            }
            info := (info.Parent != "" && classes.Has(info.Parent)) ? classes[info.Parent] : ""
        }
        return ""
    }

    ; The line without its comment and with its strings emptied ("..." -> ""), so neither is read as code
    static Code(line) {
        line := RegExReplace(line, "`"(?:[^`"``]|``.)*`"", "`"`"")
        line := RegExReplace(line, "'(?:[^'``]|``.)*'", "''")
        return RegExReplace(line, "(?:^|\s);.*")
    }

    ; "a := 1, b := f(x, y)" -> ["a := 1", " b := f(x, y)"]: commas inside brackets don't split
    static Split(text) {
        parts := [], depth := 0, start := 1
        Loop Parse text {
            if InStr("([{", A_LoopField)
                depth++
            else if InStr(")]}", A_LoopField)
                depth--
            else if (A_LoopField = "," && depth = 0) {
                parts.Push(SubStr(text, start, A_Index - start))
                start := A_Index + 1
            }
        }
        parts.Push(SubStr(text, start))
        return parts
    }

    static _Names() {
        names := Map(), names.CaseSense := false
        return names
    }

    static _Method(info, isStatic, name, where) {
        methods := isStatic ? info.Methods.Static : info.Methods.Instance
        if !methods.Has(name)
            methods[name] := {Name: name, Where: where}
    }
}
