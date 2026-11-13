;===============================================================================
; ChildProcess.ahk - runs a program in the background and collects what it prints, without a shell (AutoHotkey v2)
;-------------------------------------------------------------------------------
; Run() with "cmd /c ... > file" goes through cmd.exe, which takes & | < > ^ % in the arguments as commands, and needs a temporary
; file; RunWait() stops the program until it ends. This starts the program itself with CreateProcessW (no cmd.exe, no window),
; reads its standard output and error through pipes on a short timer (running only while a program runs; PeekNamedPipe first,
; so ReadFile never waits) and then calls the callback. The program and everything it starts run in a job object, so Kill() and
; a timeout end the whole tree, and so does ALTRun ending (JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE).
;   - Only the two pipes and NUL are inherited (PROC_THREAD_ATTRIBUTE_HANDLE_LIST): no other handle of this process leaks to it
;   - Standard input is NUL: a program reading it gets the end at once instead of waiting
;   - Programs it starts are ended when it ends (the job is closed): not for a command meant to start something that stays
;
; Usage:
;   child := ChildProcess.Start(commandLine, done, options?)
;     commandLine  the full command line: the program's path quoted, arguments quoted with Win.QuoteArg (Lib\Util.ahk)
;     done(result) called once when the program ended, timed out or couldn't start (not after Kill):
;                  {ExitCode, Stdout, Stderr (text), TimedOut, Truncated (more than the maximum was printed), Error (why it
;                  couldn't start, "" when it ran), Ms (how long it ran), Pid}
;     options      {Dir: working folder, Env: Map(name, value) set for the program only, TimeoutMs (30 s),
;                   MaxStdout (512 KB), MaxStderr (8 KB): what is printed beyond them is read and thrown away}
;   ChildProcess.Kill(child)      ends a program that hasn't finished, with everything it started: done isn't called
;   ChildProcess.Pending          number of programs still running
;   ChildProcess.Decode(buf, size) bytes -> text: UTF-8 (BOM removed) if valid, else the OEM code page (a console's default)
;===============================================================================

class ChildProcess {
    static PollMs := 30, TimeoutMs := 30000, MaxStdout := 524288, MaxStderr := 8192
    static _active := [], _pollTimer := ""

    static Pending => ChildProcess._active.Length

    static Start(commandLine, done, options := "") {
        option := (name, default) => (IsObject(options) && options.HasOwnProp(name) && options.%name% != "") ? options.%name% : default
        timeoutMs := option("TimeoutMs", ChildProcess.TimeoutMs)
        child := {Done: done, Started: A_TickCount, Deadline: A_TickCount + timeoutMs, Pid: 0, Process: 0, Job: 0
            , OutRead: 0, Out: Buffer(4096), OutSize: 0, OutMax: option("MaxStdout", ChildProcess.MaxStdout)
            , ErrRead: 0, Err: Buffer(1024), ErrSize: 0, ErrMax: option("MaxStderr", ChildProcess.MaxStderr), Truncated: false}
        try
            ChildProcess._Create(child, commandLine, option("Dir", ""), option("Env", ""))
        catch as e {
            ChildProcess._Close(child)
            ChildProcess._Later(done, ChildProcess._Result(child, false, e.Message))
            return ""
        }
        ChildProcess._active.Push(child)
        if (ChildProcess._pollTimer = "")
            ChildProcess._pollTimer := () => ChildProcess._Poll()
        SetTimer(ChildProcess._pollTimer, ChildProcess.PollMs)
        return child
    }

    ; Ends child (from Start) and everything it started, if it is still running; done isn't called. Returns whether it was running
    static Kill(child) {
        for i, active in ChildProcess._active
            if (active == child) {
                ChildProcess._active.RemoveAt(i)
                ChildProcess._Terminate(child)
                ChildProcess._Close(child)
                ChildProcess._StopTimerIfIdle()
                return true
            }
        return false
    }

    ; bytes -> text: UTF-8 when they are valid UTF-8 (a BOM is removed), else the OEM code page, which a console program uses unless
    ; told otherwise. cut: the bytes were cut off at a limit, so a character may be split at the end: up to 3 bytes are dropped there
    static Decode(buf, size, cut := false) {
        ptr := buf.Ptr
        if (size >= 3 && NumGet(ptr, 0, "UChar") = 0xEF && NumGet(ptr, 1, "UChar") = 0xBB && NumGet(ptr, 2, "UChar") = 0xBF)
            ptr += 3, size -= 3
        if (size <= 0)
            return ""
        Loop (cut ? Min(4, size) : 1) {
            length := size - A_Index + 1
            if DllCall("MultiByteToWideChar", "UInt", 65001, "UInt", 8, "Ptr", ptr, "Int", length, "Ptr", 0, "Int", 0)   ; MB_ERR_INVALID_CHARS
                return StrGet(ptr, length, "UTF-8")
        }
        return StrGet(ptr, size, "CP" DllCall("GetOEMCP", "UInt"))
    }

    static _Create(child, commandLine, dir, env) {
        inherit := Buffer(A_PtrSize * 3, 0)                                 ; SECURITY_ATTRIBUTES with bInheritHandle
        NumPut("UInt", inherit.Size, inherit, 0), NumPut("Int", 1, inherit, A_PtrSize * 2)
        outWrite := 0, errWrite := 0, nul := 0, attributes := "", saved := Map()
        try {
            if !DllCall("CreatePipe", "Ptr*", &outRead := 0, "Ptr*", &outWrite, "Ptr", inherit, "UInt", 0)
                throw OSError()
            child.OutRead := outRead
            if !DllCall("CreatePipe", "Ptr*", &errRead := 0, "Ptr*", &errWrite, "Ptr", inherit, "UInt", 0)
                throw OSError()
            child.ErrRead := errRead
            DllCall("SetHandleInformation", "Ptr", outRead, "UInt", 1, "UInt", 0)   ; our ends aren't inherited (HANDLE_FLAG_INHERIT)
            DllCall("SetHandleInformation", "Ptr", errRead, "UInt", 1, "UInt", 0)
            nul := DllCall("CreateFileW", "Str", "NUL", "UInt", 0x80000000, "UInt", 3, "Ptr", inherit, "UInt", 3, "UInt", 0, "Ptr", 0, "Ptr")
            if (nul = -1)
                throw OSError()

            handles := Buffer(A_PtrSize * 3)                                ; exactly these three are inherited
            NumPut("Ptr", nul, "Ptr", outWrite, "Ptr", errWrite, handles)
            DllCall("InitializeProcThreadAttributeList", "Ptr", 0, "UInt", 1, "UInt", 0, "UPtr*", &listSize := 0)
            attributes := Buffer(listSize, 0)
            if !DllCall("InitializeProcThreadAttributeList", "Ptr", attributes, "UInt", 1, "UInt", 0, "UPtr*", &listSize)
                throw OSError()
            if !DllCall("UpdateProcThreadAttribute", "Ptr", attributes, "UInt", 0, "UPtr", 0x20002, "Ptr", handles, "UPtr", handles.Size, "Ptr", 0, "Ptr", 0) {   ; PROC_THREAD_ATTRIBUTE_HANDLE_LIST
                DllCall("DeleteProcThreadAttributeList", "Ptr", attributes), attributes := ""
                throw OSError()
            }

            x64 := (A_PtrSize = 8)
            startup := Buffer(x64 ? 112 : 72, 0)                            ; STARTUPINFOEXW
            NumPut("UInt", startup.Size, startup, 0)
            NumPut("UInt", 0x100, startup, x64 ? 60 : 44)                   ; STARTF_USESTDHANDLES
            NumPut("Ptr", nul, "Ptr", outWrite, "Ptr", errWrite, startup, x64 ? 80 : 56)
            NumPut("Ptr", attributes.Ptr, startup, x64 ? 104 : 68)
            info := Buffer(A_PtrSize * 2 + 8, 0)                            ; PROCESS_INFORMATION
            line := Buffer(StrPut(commandLine) * 2), StrPut(commandLine, line)   ; CreateProcessW may write to the command line

            if (env is Map)                                                 ; the program inherits this process's variables: set them just for the call
                for name, value in env
                    saved[name] := EnvGet(name), EnvSet(name, value)
            ; CREATE_SUSPENDED: into the job before it can start anything; CREATE_NO_WINDOW; EXTENDED_STARTUPINFO_PRESENT
            created := DllCall("CreateProcessW", "Ptr", 0, "Ptr", line, "Ptr", 0, "Ptr", 0, "Int", 1, "UInt", 0x4 | 0x08000000 | 0x80000
                , "Ptr", 0, dir != "" ? "Str" : "Ptr", dir != "" ? dir : 0, "Ptr", startup, "Ptr", info)
            lastError := A_LastError
            if !created
                throw OSError(lastError)
        } finally {
            for name, value in saved
                (value != "") ? EnvSet(name, value) : EnvSet(name)
            if (attributes != "")
                DllCall("DeleteProcThreadAttributeList", "Ptr", attributes)
            for handle in [outWrite, errWrite, nul]                         ; the program has its own copies now: the pipes end when it (and what it started) closes them
                if (handle && handle != -1)
                    DllCall("CloseHandle", "Ptr", handle)
        }
        child.Process := NumGet(info, 0, "Ptr"), threadHandle := NumGet(info, A_PtrSize, "Ptr"), child.Pid := NumGet(info, A_PtrSize * 2, "UInt")
        if (job := DllCall("CreateJobObjectW", "Ptr", 0, "Ptr", 0, "Ptr")) {
            limits := Buffer(x64 ? 144 : 112, 0)                            ; JOBOBJECT_EXTENDED_LIMIT_INFORMATION
            NumPut("UInt", 0x2000, limits, 16)                              ; JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
            if (DllCall("SetInformationJobObject", "Ptr", job, "Int", 9, "Ptr", limits, "UInt", limits.Size)
                && DllCall("AssignProcessToJobObject", "Ptr", job, "Ptr", child.Process))
                child.Job := job
            else
                DllCall("CloseHandle", "Ptr", job)                          ; no job (very old Windows): Kill ends the program only
        }
        DllCall("ResumeThread", "Ptr", threadHandle)
        DllCall("CloseHandle", "Ptr", threadHandle)
    }

    static _Poll() {
        Critical("On")                                                      ; a Kill from another thread mustn't change the list halfway (short: reads don't wait)
        i := ChildProcess._active.Length
        while (i >= 1) {
            child := ChildProcess._active[i]
            ChildProcess._ReadAll(child)
            if (DllCall("WaitForSingleObject", "Ptr", child.Process, "UInt", 0) = 0) {   ; ended: what it printed is all in the pipes now
                ChildProcess._active.RemoveAt(i)
                ChildProcess._ReadAll(child)
                DllCall("GetExitCodeProcess", "Ptr", child.Process, "UInt*", &code := 0)
                ChildProcess._Close(child)                                  ; closing the job ends what it started and left running
                ChildProcess._Later(child.Done, ChildProcess._Result(child, false, "", code))
            } else if (A_TickCount >= child.Deadline) {
                ChildProcess._active.RemoveAt(i)
                ChildProcess._Terminate(child)
                ChildProcess._ReadAll(child)
                ChildProcess._Close(child)
                ChildProcess._Later(child.Done, ChildProcess._Result(child, true))
            }
            i--
        }
        ChildProcess._StopTimerIfIdle()
    }

    static _ReadAll(child) {
        ChildProcess._Read(child, "Out")
        ChildProcess._Read(child, "Err")
    }

    ; Reads what is waiting in one pipe (which: "Out" / "Err") without waiting for more; beyond the maximum it is read and thrown
    ; away, so a program printing a lot doesn't stop on a full pipe
    static _Read(child, which) {
        handle := child.%which%Read
        loop {
            if !(DllCall("PeekNamedPipe", "Ptr", handle, "Ptr", 0, "UInt", 0, "Ptr", 0, "UInt*", &waiting := 0, "Ptr", 0) && waiting)
                return
            size := child.%which%Size, most := child.%which%Max, take := Min(waiting, 65536)
            if (size < most) {
                take := Min(take, most - size), buf := child.%which%
                if (buf.Size < size + take)
                    buf.Size := Min(most, Max(size + take, buf.Size * 2))
                if !DllCall("ReadFile", "Ptr", handle, "Ptr", buf.Ptr + size, "UInt", take, "UInt*", &got := 0, "Ptr", 0)
                    return
                child.%which%Size := size + got
            } else {
                scratch := Buffer(take)
                if !DllCall("ReadFile", "Ptr", handle, "Ptr", scratch, "UInt", take, "UInt*", &got := 0, "Ptr", 0)
                    return
                child.Truncated := true
            }
        }
    }

    static _Result(child, timedOut, problem := "", code := "") {
        return {ExitCode: code, TimedOut: timedOut, Error: problem, Truncated: child.Truncated, Ms: A_TickCount - child.Started, Pid: child.Pid
            , Stdout: ChildProcess.Decode(child.Out, child.OutSize, child.Truncated && child.OutSize >= child.OutMax)
            , Stderr: ChildProcess.Decode(child.Err, child.ErrSize, child.ErrSize >= child.ErrMax)}
    }

    static _Terminate(child) {
        if child.Job
            DllCall("TerminateJobObject", "Ptr", child.Job, "UInt", 1)
        else if child.Process
            DllCall("TerminateProcess", "Ptr", child.Process, "UInt", 1)
    }

    static _Close(child) {
        for name in ["OutRead", "ErrRead", "Process", "Job"]
            if child.%name% {
                DllCall("CloseHandle", "Ptr", child.%name%)
                child.%name% := 0
            }
    }

    static _StopTimerIfIdle() {
        if (!ChildProcess._active.Length && ChildProcess._pollTimer != "")
            SetTimer(ChildProcess._pollTimer, 0)
    }

    static _Later(done, result) {
        if IsObject(done)
            SetTimer(() => done.Call(result), -1)
    }
}
