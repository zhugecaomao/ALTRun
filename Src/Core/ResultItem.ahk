;===============================================================================
; ResultItem.ahk - one search result (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Every feature (provider) returns ResultItems; the search window knows only this format:
;
;   Title        main title
;   Subtitle     gray subtitle (path / URL / description)
;   Icon         icon description, see IconCache.Get(): file path / "res:dll,index" / "url:" ...
;   Kind         result type, decides the actions: file / folder / url / text / command
;   Arg          what the action works on: file path / URL / text
;   Arguments    command-line arguments when running a program (optional, file type)
;   Uid          a stable unique Id to learn "what was typed -> what was chosen"; empty = not learned
;   OnRun        the function fn(item) for the default action (Enter); empty = the default action for Kind
;   RunTitle     OnRun's name in the actions panel (e.g. "Paste image"); when set it is the first row of the panel
;   Actions      extra actions [{Title, Subtitle, Icon, Run: fn(item)}], shown in the actions panel
;   Valid        false = can't run (e.g. a hint row); Enter autocompletes instead
;   AutoComplete the text Tab completes to
;   LargeText    the text for Ctrl+L Large Type, Title if empty
;   Score        ranking score, higher comes first
;   Source       the settings entry the result comes from (the Map of a custom command / snippet / search engine), for editing and deleting
;   Exclusive    true = a keyword-mode result ("clip " / "g xxx" / ">cmd"); when there are such results only they are shown
;   Pinned       true = a pinned result (clipboard, empty search box); a pin is drawn at the icon's bottom right
;
; Usage:
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
        this.RunTitle     := ""
        this.Actions      := []
        this.Valid        := true
        this.AutoComplete := ""
        this.LargeText    := ""
        this.Score        := 0
        this.Exclusive    := false
        this.Pinned       := false
        this.Source       := ""
        this.Provider     := ""
        if IsObject(props) {
            for name, value in props.OwnProps()
                this.%name% := value
        }
    }

    ; the text copied to the clipboard
    CopyText() {
        if (this.Kind = "text" || this.Kind = "file" || this.Kind = "folder" || this.Kind = "url")
            return (this.Arg != "") ? this.Arg : this.Title
        return this.Title
    }

    ; the text for Ctrl+L Large Type
    DisplayText() {
        return (this.LargeText != "") ? this.LargeText : this.CopyText()
    }
}
