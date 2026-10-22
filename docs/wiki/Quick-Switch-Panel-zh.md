[English](Quick-Switch-Panel) · **中文**

# 对话框文件夹面板

文件夹面板是 Quick Switch 的可视操作方式。默认启用; 打开 / 保存对话框 (或 [浏览文件夹](Quick-Switch-zh#浏览文件夹对话框) 对话框) 出现时, 面板会贴在对话框下方, 下方空间不足时显示在上方。点击文件夹即可跳转, 不用记快捷键。

![打开对话框下面的文件夹面板](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/quickswitch.png)

## 面板内容与操作

- 显示每个 Total Commander 窗口的当前面板和另一侧面板、每个 Directory Opus 窗口的当前标签页 (和另一侧)、已打开的资源管理器窗口, 以及最近用过的文件夹。
- 搜索框先过滤现有列表, 再通过 Everything (或 ALTRun 内置文件索引) 搜索匹配名称的文件夹和文件; 文件夹排在文件前面。
- 选中一个文件后, 跳到它所在的文件夹, 不会替用户打开或保存该文件。
- `↑` / `↓` 选择, `Enter` 跳转, `Esc` 返回原对话框。使用中文输入法时, 第一次 `Enter` 可能只是把拼音上屏, 上屏后再按一次才会跳转。
- 对话框不在前台时, 面板会隐藏; 对话框宽度改变时, 面板会跟随调整。面板颜色跟随 ALTRun 主题。

在标准打开 / 保存对话框里, `Ctrl+Shift+G` 可临时显示面板并聚焦搜索框。即使关闭了自动显示, 仍然可以用这个热键打开。

## 搜索范围

偏好设置 → 对话框面板 → “搜索框搜索”决定搜索范围:

- **只搜文件夹** (默认): 只显示匹配的文件夹。
- **文件夹和文件**: 同时搜索文件和文件夹; 文件夹仍然单独优先显示, 不会被同名文件挤掉。

最近文件夹来自 Windows 的“最近使用的项目”。打开过文件时, 记录的是文件所在文件夹; 默认显示最近 10 个, 设为 0 则不显示最近文件夹。

## 相关设置

![偏好设置 → 对话框面板](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-dialogpanel.png)

在偏好设置 → 对话框面板中修改, 对应设置文件为 `Extensions.QuickSwitch`:

| 设置 | 默认 / 可选值 | 作用 |
|---|---|---|
| `ShowPanel` | `1` | 自动显示文件夹面板 |
| `PanelSearch` | `folders` / `all` | 搜索文件夹, 或同时搜索文件夹和文件 |
| `RecentFolders` | `10` | 最近文件夹数量; `0` = 不显示 |
| `MenuHotkey` | `Ctrl+Shift+G` | 显示面板并聚焦搜索框 |

更多快捷键和自动跳转设置见[对话框快速跳转](Quick-Switch-zh)。
