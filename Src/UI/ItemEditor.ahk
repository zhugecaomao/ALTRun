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
; Type 可以是: text / multiline / choice / check / file / folder / number
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
        controls := Map()
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
                    browse.OnEvent("Click", ItemEditor._Browser(ctrl, field.Type, g))
                case "number":
                    ctrl := g.AddEdit("x+8 ys-3 w100 r1 -Multi Number", value)
                default:
                    ctrl := g.AddEdit("x+8 ys-3 w" inputW " r1 -Multi", value)
            }
            if field.HasOwnProp("Hint") {                                   ; 灰色小字说明, 和偏好设置里的一样
                g.SetFont("s8")
                g.AddText("xs+" (labelW + 8) " y+3 w" inputW " cGray", field.Hint)
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
        owner.Opt("-Disabled")
        try WinActivate("ahk_id " owner.Hwnd)
        return state.Result

        OnOk(*) {
            edited := item.Clone()
            for spec in fields {
                input := controls[spec.Key]
                switch spec.Type {
                    case "check" : newValue := input.Value
                    case "choice": newValue := spec.Choices[input.Value][1]
                    case "number": newValue := IsInteger(input.Value) ? Integer(input.Value) : 0
                    default      : newValue := input.Value
                }
                if (spec.HasOwnProp("Required") && spec.Required && Trim(newValue) = "") {
                    MsgBox(I18n.T("Prefs.Required", spec.Label), title, 48)
                    input.Focus()
                    return
                }
                edited[spec.Key] := newValue
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

    static _Browse(ctrl, kind, owner) {
        owner.Opt("+OwnDialogs")
        current := Path.Resolve(ctrl.Value)
        selected := (kind = "folder") ? DirSelect("*" current, 3) : FileSelect(3, current)
        if (selected != "")
            ctrl.Value := selected
    }
}
