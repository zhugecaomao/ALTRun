; BenchProbe.ahk - 只放进性能测试用的副本 (见 BenchDriver.ahk / bench.yml): 回答驱动程序
; "现在输入框里的文字的结果已经换上并画好了吗"。两个版本用同一份, 只用到两边都有的成员
class BenchProbe {
    static Text := "", Mode := "", Paints := -1, Count := -1, Calls := 0
    static Init() {
        orig := SearchWindow.GetOwnPropDesc("SetResults").Call
        SearchWindow.DefineProp("SetResults", {Call: (self, results) => BenchProbe._Set(orig, self, results)})
        OnMessage(0x8123, (w, l, *) => BenchProbe._Query(w, l))
    }
    static _Set(orig, self, results) {
        r := orig(self, results)
        BenchProbe.Calls += 1
        BenchProbe.Text := SearchWindow.Input.Value, BenchProbe.Mode := SearchWindow.Mode
        BenchProbe.Paints := SearchWindow._paints, BenchProbe.Count := results.Length
        return r
    }
    ; w = 1: 搜索结果好了 (l = 输入框里应有的字数); 3: 操作面板好了; 2: 只是回应 (启动完、空闲)
    static _Query(w, l) {
        if (w = 2)
            return 1
        if (w = 9)                                                          ; 调试: SetResults 调用了几次 / 输入框字数 / 自绘次数
            return BenchProbe.Calls
        if (w = 8)
            return StrLen(SearchWindow.Input.Value)
        if (w = 7)
            return SearchWindow._paints
        if !SearchWindow.IsVisible()
            return 0
        if (w = 3 && SearchWindow.Mode != "actions")
            return 0
        if (w = 1 && StrLen(SearchWindow.Input.Value) != l)
            return 0
        ready := BenchProbe.Text == SearchWindow.Input.Value && BenchProbe.Mode = SearchWindow.Mode
            && (BenchProbe.Count = 0 || SearchWindow._paints > BenchProbe.Paints)
        return ready ? 1 : 0
    }
}
BenchProbe.Init()
