;===============================================================================
; ALTRun - An effective launcher for Windows, in the spirit of Alfred for macOS
; https://github.com/zhugecaomao/ALTRun
;-------------------------------------------------------------------------------
; This file only lists the modules and starts the app - see Src\Core\App.ahk
; for the startup sequence.
;
; Project layout:
;   Lib\            General-purpose libraries, nothing ALTRun-specific
;   Src\Core\       App startup, settings + migration, search model, ranking, actions
;   Src\UI\         Search window, preferences window, large type, themes, icons
;   Src\Providers\  Search features: clipboard history, applications, custom commands,
;                   snippets, system commands, calculator, web search, file search, terminal
;   Src\Extensions\ Features outside the search window: snippet auto-expansion,
;                   dialog quick switch, Ctrl+D date, PT Tools, update checker
;   Resources\      Data files shipped with ALTRun (Kanji.txt, built-in themes)
;   Themes\         Optional custom themes (<Name>.json)
;   Data\           Generated at runtime: app index, learned ranking, clipboard history
;===============================================================================
;@Ahk2Exe-SetName ALTRun
;@Ahk2Exe-SetDescription ALTRun - An effective launcher for Windows
;@Ahk2Exe-SetVersion 2026.09.25
;@Ahk2Exe-SetCopyright Copyright (c) 2013-2026 zhugecaomao
;@Ahk2Exe-SetOrigFilename ALTRun.exe
; (编译: 见 .github/workflows/release.yml; SetVersion 要和 App.Version 一致, 有测试检查)

#Requires AutoHotkey v2.0
#SingleInstance Force
#NoTrayIcon
#Warn All, OutputDebug

; --- Lib ---
#Include Lib\JSON.ahk
#Include Lib\Logger.ahk
#Include Lib\Util.ahk
#Include Lib\TextTools.ahk
#Include Lib\Kanji.ahk
#Include Lib\Everything.ahk
#Include Lib\Dialogs.ahk

; --- Src\Core ---
#Include Src\Core\App.ahk
#Include Src\Core\I18n.ahk
#Include Src\Core\AppSettings.ahk
#Include Src\Core\SchemaMigration.ahk
#Include Src\Core\SearchQuery.ahk
#Include Src\Core\ResultItem.ahk
#Include Src\Core\FuzzyMatcher.ahk
#Include Src\Core\Knowledge.ahk
#Include Src\Core\ActionCatalog.ahk
#Include Src\Core\ProviderRegistry.ahk
#Include Src\Core\FileIndex.ahk

; --- Src\UI ---
#Include Src\UI\ThemeManager.ahk
#Include Src\UI\IconCache.ahk
#Include Src\UI\SearchWindow.ahk
#Include Src\UI\LargeType.ahk
#Include Src\UI\ItemEditor.ahk
#Include Src\UI\PreferencesWindow.ahk

; --- Src\Providers ---
#Include Src\Providers\ClipboardProvider.ahk
#Include Src\Providers\ApplicationProvider.ahk
#Include Src\Providers\CustomCommandProvider.ahk
#Include Src\Providers\SnippetProvider.ahk
#Include Src\Providers\SystemProvider.ahk
#Include Src\Providers\CalculatorProvider.ahk
#Include Src\Providers\WebSearchProvider.ahk
#Include Src\Providers\FileSearchProvider.ahk
#Include Src\Providers\TerminalProvider.ahk
#Include Src\Providers\HelpProvider.ahk

; --- Src\Extensions ---
#Include Src\Extensions\SnippetExpander.ahk
#Include Src\Extensions\QuickSwitch.ahk
#Include Src\Extensions\AutoDate.ahk
#Include Src\Extensions\PTToolsWindow.ahk
#Include Src\Extensions\UpdateChecker.ahk

App.Start()
