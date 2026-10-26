;===============================================================================
; RunTests.ahk - ALTRun unit tests (AutoHotkey v2)
;-------------------------------------------------------------------------------
; The tests are in Tests\Suites (one class per area); this file loads the code,
; sets up in-memory settings and runs every suite. Run:
;   AutoHotkey64.exe /ErrorStdOut Tests\RunTests.ahk
; Results go to standard output; the exit code is the number of failures (0 = all passed).
;===============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#Warn All, StdOut

#Include %A_ScriptDir%\..\Lib\JSON.ahk
#Include %A_ScriptDir%\..\Lib\Logger.ahk
#Include %A_ScriptDir%\..\Lib\Util.ahk
#Include %A_ScriptDir%\..\Lib\TextTools.ahk
#Include %A_ScriptDir%\..\Lib\DelimitedText.ahk
#Include %A_ScriptDir%\..\Lib\Kanji.ahk
#Include %A_ScriptDir%\..\Lib\Everything.ahk
#Include %A_ScriptDir%\..\Lib\Units.ahk
#Include %A_ScriptDir%\..\Lib\ClipboardData.ahk
#Include %A_ScriptDir%\..\Lib\Dialogs.ahk
#Include %A_ScriptDir%\..\Lib\HttpGet.ahk
#Include %A_ScriptDir%\..\Lib\FolderWatcher.ahk
#Include %A_ScriptDir%\..\Lib\WindowsRecent.ahk
#Include %A_ScriptDir%\..\Src\Core\App.ahk
#Include %A_ScriptDir%\..\Src\Core\I18n.ahk
#Include %A_ScriptDir%\..\Src\Core\AppSettings.ahk
#Include %A_ScriptDir%\..\Src\Core\SchemaMigration.ahk
#Include %A_ScriptDir%\..\Src\Core\SearchQuery.ahk
#Include %A_ScriptDir%\..\Src\Core\ResultItem.ahk
#Include %A_ScriptDir%\..\Src\Core\FuzzyMatcher.ahk
#Include %A_ScriptDir%\..\Src\Core\BackgroundWork.ahk
#Include %A_ScriptDir%\..\Src\Core\Knowledge.ahk
#Include %A_ScriptDir%\..\Src\Core\Usage.ahk
#Include %A_ScriptDir%\..\Src\Core\ActionCatalog.ahk
#Include %A_ScriptDir%\..\Src\Core\ProviderRegistry.ahk
#Include %A_ScriptDir%\..\Src\Core\FileIndex.ahk
#Include %A_ScriptDir%\..\Src\UI\ThemeManager.ahk
#Include %A_ScriptDir%\..\Src\UI\ThemePreview.ahk
#Include %A_ScriptDir%\..\Src\UI\NavIcons.ahk
#Include %A_ScriptDir%\..\Src\UI\IconCache.ahk
#Include %A_ScriptDir%\..\Src\UI\SearchWindow.ahk
#Include %A_ScriptDir%\..\Src\UI\ResultRows.ahk
#Include %A_ScriptDir%\..\Src\UI\SearchStatusBar.ahk
#Include %A_ScriptDir%\..\Src\UI\TypeAhead.ahk
#Include %A_ScriptDir%\..\Src\UI\SearchInputLanguage.ahk
#Include %A_ScriptDir%\..\Src\UI\SearchKeys.ahk
#Include %A_ScriptDir%\..\Src\UI\ResultActions.ahk
#Include %A_ScriptDir%\..\Src\UI\WindowFade.ahk
#Include %A_ScriptDir%\..\Src\UI\LargeType.ahk
#Include %A_ScriptDir%\..\Src\UI\Hud.ahk
#Include %A_ScriptDir%\..\Src\UI\ItemEditor.ahk
#Include %A_ScriptDir%\..\Src\UI\HotkeyBox.ahk
#Include %A_ScriptDir%\..\Src\UI\PreferencesWindow.ahk
#Include %A_ScriptDir%\..\Src\UI\PrefsForm.ahk
#Include %A_ScriptDir%\..\Src\UI\PrefsPages.ahk
#Include %A_ScriptDir%\..\Src\Providers\ClipboardProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\ApplicationProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\CustomCommandProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\SnippetProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\SystemProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\CalculatorProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\WebSearchProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\BookmarkProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\FileSearchProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\TerminalProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\WindowProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\RecentProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\ScriptProvider.ahk
#Include %A_ScriptDir%\..\Src\Providers\HelpProvider.ahk
#Include %A_ScriptDir%\..\Src\Extensions\SnippetExpander.ahk
#Include %A_ScriptDir%\..\Src\Extensions\QuickSwitch.ahk
#Include %A_ScriptDir%\..\Src\Extensions\AutoDate.ahk
#Include %A_ScriptDir%\..\Src\Extensions\CapsLockSwitch.ahk
#Include %A_ScriptDir%\..\Src\Extensions\CapsLockHelper.ahk
#Include %A_ScriptDir%\..\Src\Extensions\TendonProfile.ahk
#Include %A_ScriptDir%\..\Src\Extensions\PTToolsWindow.ahk
#Include %A_ScriptDir%\..\Src\Extensions\UpdateChecker.ahk
#Include %A_ScriptDir%\..\Src\Extensions\CurrencyRates.ahk
#Include %A_ScriptDir%\..\Src\Extensions\SelectionActions.ahk

#Include %A_ScriptDir%\TestUtil.ahk
#Include %A_ScriptDir%\Suites\CoreTests.ahk
#Include %A_ScriptDir%\Suites\ProviderTests.ahk
#Include %A_ScriptDir%\Suites\SearchWindowTests.ahk
#Include %A_ScriptDir%\Suites\PreferencesTests.ahk
#Include %A_ScriptDir%\Suites\ExtensionTests.ahk
#Include %A_ScriptDir%\Suites\ProjectTests.ahk

OnError((err, mode) => TestRunner.OnUncaught(err, mode))                                             ; on a runtime error, print it and exit instead of hanging on a dialog
Logger.Enabled := false
I18n.LangDir := A_ScriptDir "\..\Resources\Lang"
I18n.Init("en")
AppSettings.Data := AppSettings.Defaults()                                  ; default settings in memory; no file is read or written
AppSettings.Feature("Clipboard")["Persist"] := 0                           ; the clipboard history tests don't write to disk
IconCache.DiskDir := ""                                                     ; icons don't use the disk cache (switched on for the IconDiskCache test only)
Pinyin.DataFile := A_ScriptDir "\..\Resources\Pinyin.txt"                   ; the scripts run from Tests\

TestRunner.Run()

class TestRunner {
    static Passed := 0, Failed := 0

    ; 8.3 short path (C:\Users\RUNNER~1\...) -> long path, as Loop Files lists them
    static LongPath(shortPath) {
        buf := Buffer(32767 * 2)
        return DllCall("GetLongPathNameW", "WStr", shortPath, "Ptr", buf, "UInt", 32767, "UInt") ? StrGet(buf) : shortPath
    }

    static Suites := [CoreTests, ProviderTests, SearchWindowTests, PreferencesTests, ExtensionTests, ProjectTests]

    ; Runs every test of every suite: a test is a static method whose name doesn't start with "_".
    ; Within a suite the tests run in name order (OwnProps lists them that way).
    static Run() {
        for suite in TestRunner.Suites {
            for name in suite.OwnProps() {
                if (SubStr(name, 1, 1) = "_" || name = "Prototype" || !suite.GetOwnPropDesc(name).HasOwnProp("Call"))
                    continue
                TestRunner.RunTest(suite, name)
            }
        }
        FileAppend("`n" TestRunner.Passed " passed, " TestRunner.Failed " failed`n", "*")
        ExitApp(TestRunner.Failed)
    }

    static RunTest(suite, name) {
        state := TestRunner._State()
        try {
            suite.%name%()
        } catch as e {
            SplitPath(e.File, &fileName)                                ; file name and call stack, so an occasional failure still shows where
            stack := RegExReplace(Trim(e.Stack, "`r`n"), "\R", " <- ")
            TestRunner.Fail(name, "exception: " e.Message " (" fileName ":" e.Line ") " SubStr(stack, 1, 600))
        }
        if ((after := TestRunner._State()) != state)                    ; each test puts back what it changed, so the order doesn't matter
            TestRunner.Fail(name, "left shared state changed (settings, language or logging): " TestRunner._Changed(state, after))
    }

    ; What a test must leave as it found it: the settings, the UI language, the log switches
    static _State() => JSON.Stringify(AppSettings.Data) "`n" I18n.Lang "`n" Logger.Enabled "|" Logger.File "|" Logger.Immediate

    ; The first difference, short enough to read in the CI annotation
    static _Changed(before, after) {
        loop Min(StrLen(before), StrLen(after))
            if (SubStr(before, A_Index, 1) != SubStr(after, A_Index, 1))
                return "... " SubStr(after, Max(1, A_Index - 40), 80) " ..."
        return "(length " StrLen(before) " -> " StrLen(after) ")"
    }

    static OnUncaught(err, mode) {
        FileAppend("UNCAUGHT " err.Message " (" err.File ":" err.Line ")`n", "*")
        ExitApp(99)
    }

    static Equal(name, actual, expected) {
        if (actual == expected)
            TestRunner.Passed++
        else
            TestRunner.Fail(name, "expected [" expected "] got [" actual "]")
    }

    static True(name, condition) {
        if condition
            TestRunner.Passed++
        else
            TestRunner.Fail(name, "condition is false")
        return condition ? true : false
    }

    static Fail(name, message) {
        TestRunner.Failed++
        FileAppend("FAIL " name ": " message "`n", "*")
    }
}
