;===============================================================================
; ResultItem.ahk - 一条搜索结果 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 所有功能 (Provider) 返回的结果都是 ResultItem, 搜索窗口只认这一种格式:
;
;   Title        主标题
;   Subtitle     灰色副标题 (路径 / 网址 / 说明)
;   Icon         图标描述, 见 IconCache.Get(): 文件路径 / "res:dll,序号" / "url:" ...
;   Kind         结果类型, 决定有哪些操作: file / folder / url / text / command
;   Arg          操作的对象: 文件路径 / 网址 / 文字
;   Arguments    运行程序时附带的命令行参数 (file 类型可选)
;   Uid          稳定的唯一 Id, 用来学习 "输入什么 -> 选了什么", 留空则不学习
;   OnRun        默认操作 (Enter) 的函数 fn(item); 留空按 Kind 执行默认操作
;   Actions      额外操作 [{Title, Subtitle, Icon, Run: fn(item)}], 出现在操作面板
;   Valid        false = 不能执行 (例如提示行), Enter 时改为自动补全
;   AutoComplete Tab 自动补全成的文字
;   LargeText    Ctrl+L 大字显示的文字, 留空用 Title
;   Score        排序分数, 越大越靠前
;   Exclusive    true = 关键字模式的结果 ("clip " / "g xxx" / ">cmd"), 有这种结果时只显示它们
;
; 用法:
;   ResultItem("Notepad", "C:\Windows\notepad.exe", {Kind: "file", Arg: path, Uid: "app:" path})
;===============================================================================

class ResultItem {
    __New(title, subtitle := "", props := "") {
        this.Title        := title
        this.Subtitle     := subtitle
        this.Icon         := ""
        this.Kind         := "command"
        this.Arg          := ""
        this.Arguments    := ""
        this.Uid          := ""
        this.OnRun        := ""
        this.Actions      := []
        this.Valid        := true
        this.AutoComplete := ""
        this.LargeText    := ""
        this.Score        := 0
        this.Exclusive    := false
        this.Provider     := ""
        if IsObject(props) {
            for name, value in props.OwnProps()
                this.%name% := value
        }
    }

    ; 复制到剪贴板时用的文字
    CopyText() {
        if (this.Kind = "text" || this.Kind = "file" || this.Kind = "folder" || this.Kind = "url")
            return (this.Arg != "") ? this.Arg : this.Title
        return this.Title
    }

    ; Ctrl+L 大字显示的文字
    DisplayText() {
        return (this.LargeText != "") ? this.LargeText : this.CopyText()
    }
}
