;===============================================================================
; FileDrag.ahk - drag files / folders with the mouse from a window of this program to another program (AutoHotkey v2)
;-------------------------------------------------------------------------------
; The data is the shell's own data object for the items (IShellItemArray → BHID_DataObject), the same one Explorer drags: other
; programs get CF_HDROP, file names and file contents as from Explorer. SHDoDragDrop makes its own drop source (left button held;
; Esc cancels) and shows the shell's drag image.
;
;   effect := FileDrag.Start(hwnd, ["C:\a.txt"])               while the left button is down (from LVN_BEGINDRAG); waits until it's dropped
;   effect: FileDrag.Copy / Link (/ Move when allowed), 0 when cancelled, dropped where it can't go, or the paths couldn't be read
;
; Only copy and link are offered by default: dragging a search result somewhere never moves the user's file (Explorer would move it
; within the same drive otherwise)
;===============================================================================

class FileDrag {
    static Copy := 1, Move := 2, Link := 4
    static DRAGDROP_S_DROP := 0x40100

    static Start(hwnd, paths, allowed := FileDrag.Copy | FileDrag.Link) {
        static ole := DllCall("ole32\OleInitialize", "Ptr", 0, "Int")       ; DoDragDrop needs OLE on this thread (S_FALSE when it already was)
        static BHID_DataObject := FileDrag._Guid("{B8C0BD9F-ED24-455C-83E6-D5390C4FE8C4}")
        static IID_IDataObject := FileDrag._Guid("{0000010E-0000-0000-C000-000000000046}")
        if (!IsObject(paths) || !paths.Length)
            return 0
        pidls := Buffer(A_PtrSize * paths.Length, 0), count := 0, items := 0, dataObject := 0
        try {
            for filePath in paths {
                pidl := 0
                hr := DllCall("shell32\SHParseDisplayName", "WStr", filePath, "Ptr", 0, "Ptr*", &pidl, "UInt", 0, "Ptr", 0, "Int")
                if (hr < 0 || !pidl) {
                    Logger.Warn("FileDrag: cannot read " filePath " (0x" Format("{:08X}", hr & 0xFFFFFFFF) ")")
                    return 0
                }
                NumPut("Ptr", pidl, pidls, count++ * A_PtrSize)
            }
            if (DllCall("shell32\SHCreateShellItemArrayFromIDLists", "UInt", count, "Ptr", pidls, "Ptr*", &items, "Int") < 0 || !items)
                return 0
            ComCall(3, items, "Ptr", 0, "Ptr", BHID_DataObject, "Ptr", IID_IDataObject, "Ptr*", &dataObject)   ; IShellItemArray::BindToHandler
            effect := 0
            hr := DllCall("shell32\SHDoDragDrop", "Ptr", hwnd, "Ptr", dataObject, "Ptr", 0, "UInt", allowed, "UInt*", &effect, "Int")
            Logger.Debug("FileDrag: " count " item(s), result 0x" Format("{:X}", hr & 0xFFFFFFFF) ", effect " effect)
            return (hr = FileDrag.DRAGDROP_S_DROP) ? effect : 0
        } catch as e {
            Logger.Warn("FileDrag: " e.Message)
            return 0
        } finally {
            if dataObject
                ObjRelease(dataObject)
            if items
                ObjRelease(items)
            Loop count
                DllCall("ole32\CoTaskMemFree", "Ptr", NumGet(pidls, (A_Index - 1) * A_PtrSize, "Ptr"))
        }
    }

    static _Guid(text) {
        guid := Buffer(16)
        DllCall("ole32\CLSIDFromString", "WStr", text, "Ptr", guid)
        return guid
    }
}
