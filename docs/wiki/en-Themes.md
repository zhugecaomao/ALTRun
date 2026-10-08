**English** · [中文](Themes)

# Themes

In Preferences → Appearance, click a theme's thumbnail to pick it (each thumbnail draws a simplified search window with the theme's own colors and corner radius; custom themes have one too), then click **Apply** or **OK**.

![Preferences → Appearance](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-appearance.png)

## Built-in themes
| Theme | Name (value in settings) | Style |
|---|---|---|
| System (Light / Dark) | `System` | Follows the Windows light / dark mode (Light or Dark) and switches when Windows does |
| Light | `Light` | Default: off-white background, light blue selection |
| Dark | `Dark` | Dark gray background, blue-gray selection |
| Dark Compact | `DarkCompact` | Same colors as Dark with smaller text and rows, so more results fit |
| Classic | `Classic` | Light gray background with a bold blue full-width selection bar |
| Midnight | `Midnight` | Almost black, rounded selection |
| Midnight Compact | `MidnightCompact` | Same colors as Midnight with smaller text and rows, 24 px icons |
| Frost | `Frost` | Translucent cool white |
| Graphite | `Graphite` | macOS dark with the system blue accent |
| Ocean | `Ocean` | Blue-gray ([Nord](https://www.nordtheme.com/) palette) |
| Paper | `Paper` | Warm beige, easy on the eyes for long sessions |
| Light Compact | `LightCompact` | Same colors as Light with smaller text and rows (the counterpart of Dark Compact) |
| Dark Modern | `DarkModern` | VS Code's default Dark Modern colors: dark gray background, dark blue selection, bright blue highlights, 24 px icons |
| Light Modern | `LightModern` | VS Code's default Light Modern colors: off-white background, light blue selection, 24 px icons |
| Monokai | `Monokai` | The classic Monokai colors: dark olive background, green highlights |
| One Dark | `OneDark` | Atom / VS Code One Dark colors: blue-gray dark background, blue highlights |
| Tokyo Night | `TokyoNight` | Deep night blue with blue accents ([Tokyo Night](https://github.com/folke/tokyonight.nvim) palette) |
| Dracula | `Dracula` | Dark purple-gray with purple accents ([Dracula](https://draculatheme.com/) palette) |
| Catppuccin Mocha | `CatppuccinMocha` | Soft dark colors with lavender accents ([Catppuccin](https://catppuccin.com/) palette) |
| Gruvbox Dark | `GruvboxDark` | Warm dark brown with yellow accents, retro ([Gruvbox](https://github.com/morhetz/gruvbox) palette) |
| Solarized Light | `SolarizedLight` | Beige background with blue accents, soft contrast ([Solarized](https://ethanschoonover.com/solarized/) palette) |

| Light | Dark | Classic |
|:---:|:---:|:---:|
| ![Light](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/search.png) | ![Dark](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-dark.png) | ![Classic](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-classic.png) |
| **Midnight** | **Frost** | **Graphite** |
| ![Midnight](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-midnight.png) | ![Frost](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-frost.png) | ![Graphite](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-graphite.png) |
| **Ocean** | **Paper** | **DarkCompact** |
| ![Ocean](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-ocean.png) | ![Paper](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-paper.png) | ![DarkCompact](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-darkcompact.png) |
| **LightCompact** | **TokyoNight** | **Dracula** |
| ![LightCompact](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-lightcompact.png) | ![TokyoNight](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-tokyonight.png) | ![Dracula](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-dracula.png) |
| **CatppuccinMocha** | **GruvboxDark** | **SolarizedLight** |
| ![CatppuccinMocha](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-catppuccinmocha.png) | ![GruvboxDark](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-gruvboxdark.png) | ![SolarizedLight](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-solarizedlight.png) |
| **MidnightCompact** | **DarkModern** | **LightModern** |
| ![MidnightCompact](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-midnightcompact.png) | ![DarkModern](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-darkmodern.png) | ![LightModern](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-lightmodern.png) |
| **Monokai** | **OneDark** | |
| ![Monokai](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-monokai.png) | ![OneDark](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-onedark.png) | |

Light is built into the program; the other built-in themes are in `Resources\Themes\*.json`. Those files are replaced on upgrade, so don't edit them; use a custom theme as described below.

## Custom themes
The easiest way: Preferences → Appearance → select a theme close to what you want → **Copy as Custom Theme...** → give it a name. ALTRun creates `Themes\<name>.json` with every key, selects it and opens it in Notepad. Change the colors, save, then click **Apply** in Preferences (or tray menu → Reload) to see the result.

You can also write only the keys you want to change and use `Base` to say which theme to start from (Light by default):
```json
{
  "Base": "Dark",
  "SelectedBackground": "1D4ED8",
  "SelectedRadius": 8,
  "Opacity": 240
}
```

- A theme in `Themes\` with the same name as a built-in theme wins. For example, `Themes\Dark.json` with `"Base": "Dark"` and a few changed colors modifies the built-in Dark theme
- Theme files go in the `Themes\` folder (Preferences → Appearance → Open Themes Folder) and appear in the theme list

## Available keys
Colors are always `RRGGBB` (without `#`); font sizes are in pt; sizes are pixels at 96 DPI and scale with the display automatically.

| Key | Light default | Description |
|---|---|---|
| FontName | `auto` | Font; `auto` = Microsoft YaHei UI for a Chinese interface, Yu Gothic UI for Japanese, Segoe UI for English |
| InputFontSize | 20 | Search box font size |
| TitleFontSize | 13 | Result title font size |
| SubtitleFontSize | 9.5 | Result details font size |
| ShortcutFontSize | 9 | Font size of the Ctrl+N hints on the right |
| Padding | 14 | Inner padding of the window |
| RowHeight | 54 | Height of each row |
| IconSize | 32 | Icon size |
| SelectedRadius | 0 | Corner radius of the selected row; 0 = square full-width selection bar |
| Opacity | 255 | Window opacity 1–255, 255 = opaque |
| Background | `FAFAFA` | Background |
| Border | `C8C8C8` | Window border (Windows 11) |
| InputText | `1F1F1F` | Search box text |
| Separator | `E4E4E4` | Line between the search box and the results |
| Title | `1F1F1F` | Title |
| Subtitle | `808080` | Details |
| Shortcut | `A0A0A0` | Shortcut hints |
| SelectedBackground | `DDE7F6` | Selected row background |
| SelectedTitle | `000000` | Selected row title |
| SelectedSubtitle | `4A5568` | Selected row details |
| SelectedShortcut | `4A5568` | Selected row shortcut hint |
| Highlight | `005FE0` | Characters in the title that match your input (similar to Listary's highlighting); set it to the same value as Title to turn highlighting off |
| SelectedHighlight | `0047B8` | Matching characters in the selected row's title |

The window width, the number of visible results and the window position aren't part of a theme; set them in Preferences → Appearance (`Appearance.Width`, `Appearance.VisibleRows`, `Appearance.ShowOn`, `Appearance.RememberPosition`).
