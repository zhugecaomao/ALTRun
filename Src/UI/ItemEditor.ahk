;===============================================================================
; ItemEditor.ahk - 通用的 "编辑一条记录" 对话框 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 偏好设置里的自定义命令 / 片段 / 搜索引擎 / 自定义热键共用这一个对话框, 每种
; 记录只需要描述有哪些字段:
;   fields := [
;     {Key: "Title",  Label: "名称", Type: "text", Required: true},
;     {Key: "Type",   Label: "类型", Type: "choice", Choices: [["File", "文件"], ["Folder", "文件夹"]]},
;     {Key: "Target", Label: "目标", Type: "file"},        ; 文本框 + "..." 选择文件
;     {Key: "Text",   Label: "正文", Type: "multiline"},
;     {Key: "AutoExpand", Label: "自动展开", Type: "check"}
;   ]
; Type 可以是: text / multiline / choice / check / file / folder / number / hotkey
; file 字段可以加 FolderWhen: ["Type", "Folder"] = 另一个字段 (Type) 选的是 Folder 时, "浏览" 改为选择文件夹
; hotkey: 录制热键的框 (HotkeyBox, 可以录鼠标中键 / 侧键) + 复选框 "保留按键原来的功能" (写法前面的 ~)
; Hint: 输入框下面的灰色说明; Check: (value, editedMap) => 提示文字, 点 OK 时有提示就问 "仍然保存吗?"
;
; 用法:
;   result := ItemEditor.Edit(ownerGui, "标题", fields, itemMap)
;   返回编辑后的新 Map (保留 itemMap 里没有列出的键), 取消返回 ""
;   ItemEditor.WithDefaults(itemMap, defaultsMap)   缺少的键用默认值补上 (旧设置里可能没有新加的键)
;   ItemEditor.Field("Title", "Prefs.Col.Title", "text", true)   生成一个字段描述 (标签用 I18n 键)
;===============================================================================

class ItemEditor {

    static Edit(owner, title, fields, item) {
        state := {Result: ""}                                               ; 嵌套函数 OnOk 通过它把结果交回来
        g := Gui("+Owner" owner.Hwnd " -MinimizeBox", title)
        g.SetFont("s9", ThemeManager.FontName())
        g.MarginX := 14, g.MarginY := 12
        controls := Map(), passBoxes := Map()
        labelW := 110, inputW := 540                                        ; 宽一些, 长路径和命令行参数才看得全

        ; 单行输入框一定要写 r1: 不写行数时, 初始文字比框宽 (例如很长的路径) AHK 会自动
        ; 变成多行并加高, 盖住下面的控件
        for field in fields {
            g.AddText("xm w" labelW " y+10 Section", (field.Type = "check") ? "" : field.Label)   ; 复选框的文字写在框后面
            value := item.Has(field.Key) ? item[field.Key] : ""
            switch field.Type {
                case "multiline":
                    ctrl := g.AddEdit("x+8 ys-3 w" inputW " r8 +Multi +WantTab", value)
                case "check":
                    ctrl := g.AddCheckbox("x+8 ys w" inputW, field.Label)
                    ctrl.Value := value ? 1 : 0
                case "choice":
                    labels := []
                    for choice in field.Choices
                        labels.Push(choice[2])
                    ctrl := g.AddDropDownList("x+8 ys-3 w" inputW, labels)
                    ctrl.Value := Max(1, ItemEditor._ChoiceIndex(field.Choices, value))
                case "file", "folder":
                    ctrl := g.AddEdit("x+8 ys-3 w" (inputW - 34) " r1 -Multi", value)
                    browse := g.AddButton("x+4 yp-1 w30", I18n.T("Prefs.Browse"))
                    kind := field.HasOwnProp("FolderWhen") ? ItemEditor._KindGetter(field.FolderWhen, fields, controls) : field.Type
                    browse.OnEvent("Click", ItemEditor._Browser(ctrl, kind, g))
                case "number":
                    ctrl := g.AddEdit("x+8 ys-3 w100 r1 -Multi Number", value)
                case "hotkey":
                    ctrl := HotkeyBox.Add(g, "x+8 ys-3 w200", LTrim(value, "~"), true)
                    passBoxes[field.Key] := g.AddCheckbox("x+14 yp+3", I18n.T("Hotkey.PassThrough"))
                    passBoxes[field.Key].Value := (SubStr(value, 1, 1) = "~") ? 1 : 0
                default:
                    ctrl := g.AddEdit("x+8 ys-3 w" inputW " r1 -Multi", value)
            }
            if field.HasOwnProp("Hint") {                                   ; 灰色小字说明, 和偏好设置里的一样
                indent := (field.Type = "check") ? 18 : 0                   ; 复选框: 和框后面的文字对齐
                g.SetFont("s8")
                g.AddText("xs+" (labelW + 8 + indent) " y+3 w" (inputW - indent) " cGray", field.Hint)
                g.SetFont("s9")
            }
            controls[field.Key] := ctrl
        }

        g.AddButton("xm+" (labelW + inputW - 170) " y+18 w80 Default", "OK").OnEvent("Click", OnOk)
        g.AddButton("x+10 yp w80", I18n.T("Prefs.Cancel")).OnEvent("Click", (*) => g.Destroy())
        g.OnEvent("Escape", (*) => g.Destroy())
        g.OnEvent("Close", (*) => g.Destroy())

        owner.Opt("+Disabled")
        g.Show()
        WinWaitClose(g.Hwnd)
        HotkeyBox.CancelActive()                                            ; 正在录制热键时关掉了对话框
        owner.Opt("-Disabled")
        try WinActivate("ahk_id " owner.Hwnd)
        return state.Result

        OnOk(*) {
            g.Opt("+OwnDialogs")                                            ; 提示框挡住编辑窗口, 关掉提示框之前不能操作
            edited := item.Clone()
            for spec in fields {
                input := controls[spec.Key]
                switch spec.Type {
                    case "check" : newValue := input.Value
                    case "choice": newValue := spec.Choices[input.Value][1]
                    case "number": newValue := IsInteger(input.Value) ? Integer(input.Value) : 0
                    case "hotkey":
                        newValue := HotkeyBox.Value(input)
                        if (newValue != "" && passBoxes[spec.Key].Value && SubStr(newValue, 1, 1) != "~")
                            newValue := "~" newValue
                        else if !passBoxes[spec.Key].Value
                            newValue := LTrim(newValue, "~")
                    default      : newValue := input.Value
                }
                if (spec.HasOwnProp("Required") && spec.Required && Trim(newValue) = "") {
                    MsgBox(I18n.T("Prefs.Required", spec.Label), title, 48)
                    input.Focus()
                    return
                }
                edited[spec.Key] := newValue
            }
            for spec in fields {                                            ; Check: 全部字段读完后再检查, 可以看其它字段 (返回提示文字 = 有问题)
                if !spec.HasOwnProp("Check") || (warning := spec.Check.Call(edited[spec.Key], edited)) = ""
                    continue
                if (MsgBox(warning "`n`n" I18n.T("Editor.SaveAnyway"), title, "YesNo Icon! Default2") != "Yes") {
                    controls[spec.Key].Focus()
                    return
                }
            }
            state.Result := edited
            g.Destroy()
        }
    }

    static WithDefaults(item, defaults) {
        filled := item.Clone()
        for key, value in defaults
            if !filled.Has(key)
                filled[key] := value
        return filled
    }

    static Field(key, labelKey, type := "text", required := false, choices := "", hint := "") {
        spec := {Key: key, Label: I18n.T(labelKey), Type: type, Required: required}
        if IsObject(choices)
            spec.Choices := choices
        if (hint != "")
            spec.Hint := hint
        return spec
    }

    ; 编辑对话框的所属窗口: 偏好设置打开时用它, 否则用一个隐藏的窗口 (不在任务栏显示)
    static Owner() {
        static hiddenOwner := ""
        if IsObject(PreferencesWindow.Gui)
            return PreferencesWindow.Gui
        if !IsObject(hiddenOwner)
            hiddenOwner := Gui("+ToolWindow -Caption", App.Name)
        return hiddenOwner
    }

    static _ChoiceIndex(choices, value) {
        for index, choice in choices
            if (choice[1] = value)
                return index
        return 0
    }

    static _Browser(ctrl, kind, owner) {
        return (*) => ItemEditor._Browse(ctrl, kind, owner)
    }

    ; 点 "浏览" 时才看另一个字段现在选的是什么: 选的是 when[2] 就选文件夹, 否则选文件
    static _KindGetter(when, fields, controls) {
        return () => ItemEditor.KindFor(when, fields, controls)
    }

    static KindFor(when, fields, controls) {
        for spec in fields {
            if (spec.Key = when[1] && controls.Has(spec.Key) && spec.HasOwnProp("Choices")) {
                index := controls[spec.Key].Value
                return (index >= 1 && index <= spec.Choices.Length && spec.Choices[index][1] = when[2]) ? "folder" : "file"
            }
        }
        return "file"
    }

    static _Browse(ctrl, kind, owner) {
        if HasMethod(kind)
            kind := kind()
        owner.Opt("+OwnDialogs")
        current := Path.Resolve(ctrl.Value)
        selected := (kind = "folder") ? DirSelect("*" current, 3) : FileSelect(3, current)
        if (selected != "")
            ctrl.Value := selected
    }
}
