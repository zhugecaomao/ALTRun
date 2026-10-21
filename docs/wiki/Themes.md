**English** · [中文](Themes-zh)

# Themes

In Preferences → Appearance, click a theme's thumbnail to pick it (each thumbnail draws a simplified search window with the theme's own colors and corner radius; custom themes have one too), then click **Apply** or **OK**. The tray icon's right-click menu, the status bar's ☰ menu, the right-click menu of results and the Quick Switch folder menu follow the theme: dark themes use Windows' dark menus, light themes light ones (Windows 10 version 1903 or later).

![Preferences → Appearance](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-appearance.png)

## Built-in themes
| Theme | Name (value in settings) | Style |
|---|---|---|
| System (Light / Dark) | `System` | Default: follows the Windows light / dark mode (Light or Dark) and switches when Windows does |
| Light | `Light` | Off-white background, light blue selection |
| Dark | `Dark` | Dark gray background, blue-gray selection |
| Classic | `Classic` | Light gray background with a bold blue full-width selection bar |
| Midnight | `Midnight` | Almost black, rounded selection |
| Frost | `Frost` | Translucent cool white |
| Graphite | `Graphite` | macOS dark with the system blue accent |
| Ocean | `Ocean` | Blue-gray ([Nord](https://www.nordtheme.com/) palette) |
| Paper | `Paper` | Warm beige, easy on the eyes for long sessions |
| Lavender | `Lavender` | Pale lavender background, soft purple selection and purple highlights, rounded |
| Rose | `Rose` | Pale pink background with rose accents, rounded |
| Forest | `Forest` | Deep green background with green selection and highlights |
| Sunset | `Sunset` | Warm dark brown background with orange accents |
| High Contrast | `HighContrast` | Pure black, white text and a full-width yellow selection; easier to read with low vision or in bright light |
| Dark Modern | `DarkModern` | VS Code's default Dark Modern colors: dark gray background, dark blue selection, bright blue highlights |
| Light Modern | `LightModern` | VS Code's default Light Modern colors: off-white background, light blue selection |
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
| **Ocean** | **Paper** | **Lavender** |
| ![Ocean](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-ocean.png) | ![Paper](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-paper.png) | ![Lavender](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-lavender.png) |
| **Rose** | **Forest** | **Sunset** |
| ![Rose](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-rose.png) | ![Forest](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-forest.png) | ![Sunset](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-sunset.png) |
| **HighContrast** | **Dark · Compact** | **TokyoNight** |
| ![HighContrast](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-highcontrast.png) | ![Dark · Compact](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-compact.png) | ![TokyoNight](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-tokyonight.png) |
| **Dracula** | **CatppuccinMocha** | **GruvboxDark** |
| ![Dracula](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-dracula.png) | ![CatppuccinMocha](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-catppuccinmocha.png) | ![GruvboxDark](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-gruvboxdark.png) |
| **SolarizedLight** | **DarkModern** | **LightModern** |
| ![SolarizedLight](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-solarizedlight.png) | ![DarkModern](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-darkmodern.png) | ![LightModern](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-lightmodern.png) |
| **Monokai** | **OneDark** | |
| ![Monokai](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-monokai.png) | ![OneDark](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-onedark.png) | |

Light is built into the program; the other built-in themes are in `Resources\Themes\*.json`. Those files are replaced on upgrade, so don't edit them; use a custom theme as described below.

## Density
Preferences → Appearance → Size → **Density** applies to every theme, custom ones too (`Appearance.Density` in the settings):

| Density | Value | Themes with the default sizes (Light, Dark...): input / title text, row height, icons |
|---|---|---|
| Comfortable | `Comfortable` | Each theme's own size: 20 / 13 pt, 54, 32 px |
| Normal | `Normal` | Default for new installations: 17 / 11.5 pt, 48, 28 px |
| Compact | `Compact` | 15 / 10.5 pt, 42, 24 px; the most results on screen |

Built-in themes all share the same sizes and differ only in colors; custom themes with their own sizes scale by the same ratios. Upgrades keep their previous size (Comfortable); switch to Normal for something smaller.

The former built-in themes Dark Compact (`DarkCompact`), Light Compact (`LightCompact`) and Midnight Compact (`MidnightCompact`) are gone: use Dark / Light / Midnight with Normal density instead (same sizes as before). If the settings still name one of them, it shows as "that theme + Normal"; opening Preferences → Appearance selects them for you, and **OK** saves the new form. Custom themes whose `Base` names one of them keep working.

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

The window width, the number of visible results and the window position aren't part of a theme; set them in Preferences → Appearance (`Appearance.Width`, `Appearance.VisibleRows`, `Appearance.ShowOn`, `Appearance.RememberPosition`), and so is fading in and out when showing and hiding (`Appearance.Fade`).
