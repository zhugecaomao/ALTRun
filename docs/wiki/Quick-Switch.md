[English](en-Quick-Switch) · **中文**

# 对话框跳转

在打开 / 保存对话框里一步跳到已经打开的文件夹。做法借鉴 [Listary](https://www.listary.com/) 的 Quick Switch。

**文件夹面板** (默认打开, 和 Listary 的 Quick Switch 窗口一样): 对话框一出现, 正下方就自动贴一个和对话框一样宽的面板 (下面放不下时放在上面), 点一个文件夹对话框就跳过去, 不用记热键:

![打开对话框下面的文件夹面板](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/quickswitch.png)

- 列出每个 Total Commander 窗口的当前面板和另一侧面板、打开的资源管理器窗口、最近用过的文件夹;
- 上面的搜索框: 输入文字先过滤列表, 再用 Everything (或内置索引) 找名字匹配的文件夹和文件 (文件夹在前); `↑` `↓` 选择, `Enter` 确定, `Esc` 回到对话框;
- 选中的是文件时: 跳到它所在的文件夹 (打开和保存对话框都一样, 不替你打开或保存);
- 跳转时文件名框里原来的文件名 (例如另存为时程序预填的名字) 会保留;
- 面板的颜色跟随 ALTRun 的主题 (偏好设置 → 外观), 深色主题下面板也是深色; 对话框改变宽度时面板跟着变;
- 搜索框里用中文输入法时, 按 `Enter` 是把拼音上屏, 不会跳转; 上屏之后再按 `Enter` 才跳转;
- 搜索框搜什么: 偏好设置 → 对话框面板 → "搜索框搜索": 只搜文件夹 (默认) / 文件夹和文件 (`PanelSearch`: `folders` / `all`)。文件夹单独搜, 同名的文件很多时也不会把文件夹挤掉;
- 对话框不在前台时隐藏;
- 切到 TC 换了目录再回到对话框, 列表会刷新;
- 面板的设置 (开关、搜索范围、最近文件夹的数量、键盘操作的热键) 在 偏好设置 → 对话框面板 (`ShowPanel`)。最近的文件夹来自 Windows 的 "最近使用的项目" (打开过的文件取所在的文件夹), 默认 10 个, 0 = 不列。

键盘操作, 在标准的 "打开 / 保存文件" 对话框里:
- `Ctrl+G`: 跳到 Total Commander 当前打开的文件夹
- `Ctrl+E`: 跳到资源管理器当前打开的文件夹
- `Ctrl+Shift+G`: 光标跳到文件夹面板的搜索框, 用键盘选择: 输入文字搜索, `↑` `↓` 选择, `Enter` 跳过去, `Esc` 回到对话框。没有打开自动显示面板时, 按这个热键临时显示面板

打开 "自动跳转" (`AutoSwitch`) 后, 从 Total Commander 切换到对话框时会自动跳转。

TC 的目录是直接问 TC 要的 (TC 的 `WM_COPYDATA` 接口, TC 8.0 以上), 不经过剪贴板, Windows 的剪贴板历史 (`Win+V`) 里不会多出路径。

偏好设置 → 对话框跳转:
- **生效的窗口**: "Windows 标准对话框" (`ahk_class #32770`, 默认勾选) 和 "其它对话框", 例如 WPS 的 `ahk_class Qt5QWindowIcon`, 每行一个
- **不生效的窗口**: 在这些窗口里不跳转
- **不自动跳转的对话框**: 这些对话框里不自动跳转, 仍然可以按热键

设置文件里是 `Extensions.QuickSwitch` → `TotalCmdHotkey` / `ExplorerHotkey` / `MenuHotkey` / `RecentFolders` / `ShowPanel` / `PanelSearch` / `AutoSwitch` / `DialogWindows` / `ExcludeWindows` / `AutoSwitchExclude` (窗口条件用逗号分隔)。
