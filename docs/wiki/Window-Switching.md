**English** · [中文](Window-Switching-zh)

# Window switching

Type a window title or program name and `Enter` switches to that open window (restoring it if it is minimized), like PowerToys Run's Window Walker or Raycast's Switch Windows.

- `w keyword`: search windows only, for example `w excel`, `w quote`
- `w ` (`w` and a space): list every window, in front-to-back order (most recently used first)
- Typing without `w`: windows whose title matches well show up too, after apps and custom commands, at most 3
- Press `→` on a window: close it

![w and a space lists the open windows](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/windows.png)

Only windows you can see on the taskbar are listed: no hidden windows, tool windows, windows on other virtual desktops or ALTRun's own windows. Windows are never added to the [recent items](Recent).

## Settings
Preferences → Features: "Window switching" turns the whole feature on or off; under "Feature options", "Show in normal search" and "Keyword" map to:

| Setting | Default | Description |
|---|---|---|
| `Features.Windows.Enabled` | `1` | Turn window switching on or off |
| `Features.Windows.Keyword` | `w` | Keyword that searches windows only |
| `Features.Windows.InDefaultResults` | `1` | Also show windows whose title matches well when typing without the keyword |
