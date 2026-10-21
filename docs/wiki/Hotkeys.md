**English** · [中文](Hotkeys-zh)

# Custom Hotkeys

Preferences → Hotkeys: assign a [system command](System-Commands#system-commands) to any hotkey, optionally only inside a certain window.

| Field | Example | Description |
|---|---|---|
| Hotkey (Key) | `Ctrl+Alt+T` | Click the box and press the keys (see [Setting hotkeys](#setting-hotkeys) below); mouse hotkeys: click the middle or a side button in the box. With "Keep the key's original function" ticked, the key still does what it normally does (a leading `~` in the settings file) |
| Action | `TerminalHere` | A system command Id, or `ToggleWindow` (show / hide ALTRun) |
| Only in window (WinTitle) | `ahk_class CabinetWClass` | Empty = everywhere; otherwise only in matching windows; `ALTRun` = only in ALTRun's search window |

New installs come with three examples, one for each kind:
- `Ctrl+Alt+Win+V` pastes the clipboard as plain text, everywhere (the same key as PowerToys "Paste as plain text")
- `Ctrl+Alt+T` in File Explorer (`ahk_class CabinetWClass`) opens a terminal in the current folder
- `Ctrl+R` in the ALTRun search window (`ALTRun`) rebuilds the index

Delete or change them as you like; existing settings keep their own list. `F1`–`F4` in the search window are built in, see [Shortcuts](Usage#shortcuts).

![Preferences → Hotkeys](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-hotkeys.png)

### Caps Lock switches the input method
![Preferences → Keyboard](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-keyboard.png)

Like on a Mac: in Preferences → Keyboard → Input method, set **Caps Lock** to one of:
- **Press to switch input method**: switches to the next installed input method, like `Win+Space` (for example English (US) ↔ Microsoft Pinyin)
- **Press to switch Chinese / English mode**: switches the mode of the current input method, like pressing `Shift` in Microsoft Pinyin; handy when only a Chinese input method is installed

Holding Caps Lock for 0.3 s turns caps lock on or off (without waiting for you to let go) and briefly shows "Caps Lock On" / "Caps Lock Off" (untick "Show a notice when caps lock turns on or off" to turn that off); `Shift+Caps Lock` turns caps lock on or off as usual. It's off by default (`General.CapsLock`). Custom hotkeys that use Caps Lock as a modifier (such as `CapsLock & J`) don't mix well with it. It also works over remote control such as Chrome Remote Desktop: the very short extra Caps Lock presses that the remote tool sends to sync the caps lock state are ignored.

**Also in windows running as administrator**: Windows doesn't let programs with normal rights control windows running as administrator, so by default Caps Lock doesn't switch there. Tick "Also in windows running as administrator" (`General.CapsLockElevated`) and ALTRun creates a scheduled task (confirm UAC once); from then on every start uses it to run a **small helper with administrator rights that only handles Caps Lock** (no more UAC prompts), and the helper exits with ALTRun. ALTRun itself keeps normal rights, so programs opened from it, drag and drop and network drives are not affected. Unticking it deletes the task (confirm once more). After moving the ALTRun folder, turn the option off and on again.

### Setting hotkeys
Every hotkey in Preferences (the ALTRun hotkey, clipboard history, Quick Switch, date stamp, custom hotkeys) uses the same kind of box, recording the keys directly like Alfred and PowerToys:
- The box shows `Alt+Space`, `Ctrl+Alt+C` and so on. Click the box, "Press a shortcut..." appears, and the keys you press are recorded
- `Esc` cancels (keeps the old hotkey), `Backspace` / `Delete` clears it (no hotkey)
- A single letter, digit, space, Enter and the like would get in the way of typing and need `Ctrl`, `Alt` or `Win`; `F1`–`F24`, `Pause` and similar keys work on their own
- While recording, ALTRun's own hotkeys are paused, and combinations such as `Alt+Space` or `Win+E` don't trigger Windows or other programs
- When saving, you're warned if two global hotkeys are the same (for example the ALTRun hotkey and clipboard history both `Ctrl+Alt+C`). Hotkeys that only work in some windows (Quick Switch, date stamp, custom hotkeys with a window) can repeat

The settings file still uses AutoHotkey syntax (`!` Alt, `^` Ctrl, `+` Shift, `#` Win, for example `!Space`), and you can edit it directly; special forms such as `CapsLock & J` are kept as they are until you record the hotkey again in Preferences.
