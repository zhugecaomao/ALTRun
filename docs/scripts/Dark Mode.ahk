; @altrun.title          Dark Mode
; @altrun.keyword        dark
; @altrun.mode           silent
; @altrun.description    Switches Windows between light and dark mode (apps and the taskbar together)
; @altrun.description_zh 在 Windows 的浅色和深色模式之间切换 (应用和任务栏一起)
; @altrun.network        none
;
; Type "dark" in ALTRun and press Enter: light becomes dark, dark becomes light, as in Settings → Personalization → Colors.
; With "light" or "dark" as its first argument (e.g. a custom command that runs this file) it sets that mode instead of switching.
; Changes only your own Windows settings (HKEY_CURRENT_USER); open windows follow at once, a few older programs only when restarted.
#Requires AutoHotkey v2.0
#NoTrayIcon

key := "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
Chinese := RegExMatch(A_Language, "i)^(0804|0404|0C04|1004|1404)$") > 0
wanted := StrLower(A_Args.Length ? Trim(A_Args[1]) : "")
light := (wanted = "light") ? 1 : (wanted = "dark") ? 0 : !RegRead(key, "AppsUseLightTheme", 1)
try {
    RegWrite(light, "REG_DWORD", key, "AppsUseLightTheme")
    RegWrite(light, "REG_DWORD", key, "SystemUsesLightTheme")
} catch as e {
    FileAppend((Chinese ? "不能修改设置: " : "Can't change the setting: ") e.Message "`n", "**", "UTF-8-RAW")
    ExitApp(1)
}
; tell open windows (WM_SETTINGCHANGE "ImmersiveColorSet"; at most 0.2 s for each, a window that hangs is skipped)
DllCall("SendMessageTimeoutW", "Ptr", 0xFFFF, "UInt", 0x1A, "Ptr", 0, "Str", "ImmersiveColorSet", "UInt", 0x2, "UInt", 200, "Ptr*", 0)
FileAppend(light ? (Chinese ? "浅色模式" : "Light mode") : (Chinese ? "深色模式" : "Dark mode"), "*", "UTF-8-RAW")
