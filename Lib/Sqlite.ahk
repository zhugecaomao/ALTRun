;===============================================================================
; Sqlite.ahk - read-only queries through Windows' own SQLite (winsqlite3.dll) (AutoHotkey v2)
;-------------------------------------------------------------------------------
; winsqlite3.dll is part of Windows 10 / 11 (System32), so nothing has to be shipped. Used to read the History
; database of Chromium browsers (BrowserHistoryProvider).
;
; Usage:
;   Sqlite.Available()                                  true when winsqlite3.dll can be loaded
;   rows := Sqlite.Query(dbPath, "SELECT a, b FROM t")  [[a, b], ...]; integers as numbers, text as strings, NULL as ""
;   Sqlite.Query(path, sql, maxRows, "immutable")       opens the file read-only without any locking (see Query)
;   Sqlite.Exec(dbPath, "CREATE TABLE ...; INSERT ...")  runs statements, creating the file if needed (the tests build databases with it)
; Throws an Error with SQLite's message when the file can't be opened or the query fails.
;===============================================================================

class Sqlite {
    static Dll := "winsqlite3.dll", _module := "", _tried := false
    static SQLITE_OK := 0, SQLITE_ROW := 100, SQLITE_DONE := 101
    static OPEN_READONLY := 0x1, OPEN_READWRITE := 0x2, OPEN_CREATE := 0x4, OPEN_URI := 0x40
    static TYPE_INTEGER := 1, TYPE_FLOAT := 2, TYPE_TEXT := 3, TYPE_BLOB := 4, TYPE_NULL := 5

    static Available() {
        if !Sqlite._tried {
            Sqlite._tried := true
            Sqlite._module := DllCall("LoadLibrary", "Str", Sqlite.Dll, "Ptr")   ; kept loaded: only loaded when a query is made
        }
        return Sqlite._module != 0
    }

    ; Runs one query and returns at most maxRows rows (0 = all).
    ; mode: "read" = read-only (the default); "write" = read-write, so SQLite can roll back a journal left by a copied file;
    ; "immutable" = read-only through a URI with immutable=1: no locks at all, so it can't wait for or disturb the program
    ; that has the file open, but changes in its journal / WAL aren't seen
    static Query(dbPath, sql, maxRows := 0, mode := "read") {
        if !Sqlite.Available()
            throw Error(Sqlite.Dll " is not available")
        flags := (mode = "write") ? Sqlite.OPEN_READWRITE : Sqlite.OPEN_READONLY
        name := dbPath
        if (mode = "immutable") {
            flags |= Sqlite.OPEN_URI
            name := Sqlite.FileUri(dbPath) "?immutable=1"
        }
        db := 0
        rc := DllCall("winsqlite3\sqlite3_open_v2", "Ptr", Sqlite._Utf8(name), "Ptr*", &db, "Int", flags, "Ptr", 0, "Int")
        try {
            if (rc != Sqlite.SQLITE_OK)
                throw Error("open: " Sqlite._Message(db, rc))
            stmt := 0
            rc := DllCall("winsqlite3\sqlite3_prepare_v2", "Ptr", db, "Ptr", Sqlite._Utf8(sql), "Int", -1, "Ptr*", &stmt, "Ptr", 0, "Int")
            if (rc != Sqlite.SQLITE_OK)
                throw Error("prepare: " Sqlite._Message(db, rc))
            try {
                rows := []
                columns := DllCall("winsqlite3\sqlite3_column_count", "Ptr", stmt, "Int")
                while ((rc := DllCall("winsqlite3\sqlite3_step", "Ptr", stmt, "Int")) = Sqlite.SQLITE_ROW) {
                    row := []
                    Loop columns
                        row.Push(Sqlite._Column(stmt, A_Index - 1))
                    rows.Push(row)
                    if (maxRows > 0 && rows.Length >= maxRows)
                        return rows
                }
                if (rc != Sqlite.SQLITE_DONE)
                    throw Error("step: " Sqlite._Message(db, rc))
                return rows
            } finally {
                DllCall("winsqlite3\sqlite3_finalize", "Ptr", stmt)
            }
        } finally {
            if db
                DllCall("winsqlite3\sqlite3_close_v2", "Ptr", db)
        }
    }

    static Exec(dbPath, sql) {
        if !Sqlite.Available()
            throw Error(Sqlite.Dll " is not available")
        db := 0
        rc := DllCall("winsqlite3\sqlite3_open_v2", "Ptr", Sqlite._Utf8(dbPath), "Ptr*", &db, "Int", Sqlite.OPEN_READWRITE | Sqlite.OPEN_CREATE, "Ptr", 0, "Int")
        try {
            if (rc != Sqlite.SQLITE_OK)
                throw Error("open: " Sqlite._Message(db, rc))
            rc := DllCall("winsqlite3\sqlite3_exec", "Ptr", db, "Ptr", Sqlite._Utf8(sql), "Ptr", 0, "Ptr", 0, "Ptr", 0, "Int")
            if (rc != Sqlite.SQLITE_OK)
                throw Error("exec: " Sqlite._Message(db, rc))
        } finally {
            if db
                DllCall("winsqlite3\sqlite3_close_v2", "Ptr", db)
        }
    }

    ; C:\a b\History -> file:///C:/a%20b/History (characters URIs reserve are escaped; others, Chinese too, stay as they are)
    static FileUri(filePath) {
        uri := ""
        for ch in StrSplit(StrReplace(filePath, "\", "/")) {
            code := Ord(ch)
            uri .= (code < 0x80 && !RegExMatch(ch, "[A-Za-z0-9/:._~-]")) ? Format("%{:02X}", code) : ch
        }
        return "file:///" uri                                               ; local files only (SQLite takes no server in a file URI)
    }

    static _Column(stmt, index) {
        switch DllCall("winsqlite3\sqlite3_column_type", "Ptr", stmt, "Int", index, "Int") {
            case Sqlite.TYPE_INTEGER: return DllCall("winsqlite3\sqlite3_column_int64", "Ptr", stmt, "Int", index, "Int64")
            case Sqlite.TYPE_FLOAT:   return DllCall("winsqlite3\sqlite3_column_double", "Ptr", stmt, "Int", index, "Double")
            case Sqlite.TYPE_NULL:    return ""
            default:
                textPtr := DllCall("winsqlite3\sqlite3_column_text16", "Ptr", stmt, "Int", index, "Ptr")
                return textPtr ? StrGet(textPtr, "UTF-16") : ""
        }
    }

    static _Utf8(text) {
        buf := Buffer(StrPut(text, "UTF-8"))
        StrPut(text, buf, "UTF-8")
        return buf
    }

    static _Message(db, rc) {
        message := db ? StrGet(DllCall("winsqlite3\sqlite3_errmsg16", "Ptr", db, "Ptr"), "UTF-16") : ""
        return "error " rc (message != "" ? " (" message ")" : "")
    }
}
