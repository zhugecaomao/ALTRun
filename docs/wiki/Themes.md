# 主题

在 偏好设置 → 外观 里选择主题, 点 **应用** 或 **确定** 后生效。

![偏好设置 → 外观](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-appearance.png)

## 内置主题
| 主题 | 名称 (设置里的值) | 风格 |
|---|---|---|
| 跟随系统 | `System` | 跟随 Windows 的浅色 / 深色模式 (Light 或 Dark), 系统切换时自动更新 |
| 浅色 | `Light` | 默认, 浅灰白底 + 淡蓝选中 |
| 深色 | `Dark` | 深灰底 + 蓝灰选中 |
| 经典 | `Classic` | 浅灰底 + 醒目的蓝色整行选中条 |
| 午夜 | `Midnight` | 接近纯黑, 圆角选中 |
| 霜白 | `Frost` | 半透明的冷白色 |
| 石墨 | `Graphite` | macOS 深色 + 系统蓝强调色 |
| 海洋 | `Ocean` | 蓝灰色 ([Nord](https://www.nordtheme.com/) 配色) |
| 纸张 | `Paper` | 米黄色, 适合长时间看 |

| Light | Dark | Classic |
|:---:|:---:|:---:|
| ![Light](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/search.png) | ![Dark](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-dark.png) | ![Classic](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-classic.png) |
| **Midnight** | **Frost** | **Graphite** |
| ![Midnight](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-midnight.png) | ![Frost](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-frost.png) | ![Graphite](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-graphite.png) |
| **Ocean** | **Paper** | |
| ![Ocean](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-ocean.png) | ![Paper](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/theme-paper.png) | |

Light 写在程序里, 其它内置主题在 `Resources\Themes\*.json`。这些文件升级时会被替换, 请不要直接修改, 改用下面的自定义主题。

## 自定义主题
最简单的方法: 偏好设置 → 外观 → 选中一个接近的主题 → **复制为自定义主题...** → 起个名字。ALTRun 会在 `Themes\<名字>.json` 生成一份包含全部键的文件, 自动选中它并用记事本打开。改好颜色后保存, 再在偏好设置里点 **应用** (或托盘菜单 → 重新载入) 就能看到效果。

也可以只写要修改的键, 用 `Base` 指定从哪个主题开始 (默认 Light):
```json
{
  "Base": "Dark",
  "SelectedBackground": "1D4ED8",
  "SelectedRadius": 8,
  "Opacity": 240
}
```

- `Themes\` 里的主题和内置主题同名时, 用你的。例如 `Themes\Dark.json` 写 `"Base": "Dark"` 再改几种颜色, 就是在内置深色主题的基础上修改
- 主题文件放在 `Themes\` 文件夹 (偏好设置 → 外观 → 打开主题文件夹), 会出现在主题列表里

## 可用的键
颜色一律写 `RRGGBB` (不带 `#`); 字号单位是 pt; 尺寸是 96 DPI 下的像素, 会按屏幕缩放自动放大。

| 键 | Light 默认值 | 说明 |
|---|---|---|
| FontName | `auto` | 字体; `auto` = 中文界面用 Microsoft YaHei UI, 英文界面用 Segoe UI |
| InputFontSize | 20 | 搜索框字号 |
| TitleFontSize | 13 | 结果标题字号 |
| SubtitleFontSize | 9.5 | 结果说明字号 |
| ShortcutFontSize | 9 | 右侧 Ctrl+N 提示字号 |
| Padding | 14 | 窗口内边距 |
| RowHeight | 54 | 每行高度 |
| IconSize | 32 | 图标大小 |
| SelectedRadius | 0 | 选中行的圆角半径; 0 = 整行直角选中条 |
| Opacity | 255 | 窗口不透明度 1~255, 255 = 不透明 |
| Background | `FAFAFA` | 背景 |
| Border | `C8C8C8` | 窗口边框 (Windows 11) |
| InputText | `1F1F1F` | 搜索框文字 |
| Separator | `E4E4E4` | 搜索框和结果之间的分隔线 |
| Title | `1F1F1F` | 标题 |
| Subtitle | `808080` | 说明 |
| Shortcut | `A0A0A0` | 快捷键提示 |
| SelectedBackground | `DDE7F6` | 选中行背景 |
| SelectedTitle | `000000` | 选中行标题 |
| SelectedSubtitle | `4A5568` | 选中行说明 |
| SelectedShortcut | `4A5568` | 选中行快捷键提示 |

窗口宽度、显示几行结果和窗口位置不属于主题, 在 偏好设置 → 外观 里设置 (`Appearance.Width`、`Appearance.VisibleRows`、`Appearance.ShowOn`、`Appearance.RememberPosition`)。
