; BenchProbe.ahk - only put in the benchmark copies (see BenchDriver.ahk / bench.yml): answers the driver
; "are the results for the text in the input box swapped in and drawn yet". Both versions use the same file, using only members both have
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
    ; w = 1: search results ready (l = the number of characters the box should have); 3: actions panel ready; 2: just a reply (started, idle)
    static _Query(w, l) {
        if (w = 2)
            return 1
        if (w = 9)                                                          ; debugging: how many times SetResults was called / characters in the box / custom-draw count
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
