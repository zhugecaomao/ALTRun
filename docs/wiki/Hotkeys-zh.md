[English](Hotkeys) · **中文**

# 自定义热键

偏好设置 → 自定义热键: 给任意热键指定一条 [系统命令](System-Commands-zh#系统命令), 可以限定只在某个窗口里生效。

| 字段 | 例子 | 说明 |
|---|---|---|
| 热键 (Key) | `Ctrl+Alt+T` | 点一下框, 按下组合键 (见下文 [设置热键](#设置热键)); 鼠标热键: 在框里点鼠标中键或侧键。勾选 "保留按键原来的功能" 时, 按键原来的作用照常 (设置文件里是开头的 `~`) |
| 操作 (Action) | `TerminalHere` | 系统命令的 Id, 或 `ToggleWindow` (显示 / 隐藏 ALTRun) |
| 窗口 (WinTitle) | `ahk_class CabinetWClass` | 留空 = 全局; 否则只在匹配的窗口里生效; `ALTRun` = 只在 ALTRun 的搜索窗口里 |

新安装时带三个示例, 每种用法一个:
- `Ctrl+Alt+Win+V`: 在任何地方把剪贴板粘贴为纯文本 (和 PowerToys 的 "粘贴为纯文本" 同一个键)
- `Ctrl+Alt+T`: 在资源管理器 (`ahk_class CabinetWClass`) 里, 在当前文件夹打开终端
- `Ctrl+R`: 在 ALTRun 搜索窗口 (`ALTRun`) 里重建索引

可以随意删除或修改; 已有的设置保留自己的列表。搜索窗口里的 `F1` ~ `F4` 是内置的, 见 [使用方法](Usage-zh#快捷键)。

![偏好设置 → 自定义热键](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-hotkeys.png)

### CapsLock 切换输入法
![偏好设置 → 键盘](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-keyboard.png)

和 Mac 一样: 在 偏好设置 → 键盘 → 输入法 把 **CapsLock 键** 设成:
- **按一下切换输入法**: 切换到下一个已安装的输入法, 和 `Win+Space` 一样 (例如 英语(美国) ↔ 微软拼音)
- **按一下切换中 / 英模式**: 切换当前输入法的中 / 英模式, 和微软拼音里按 `Shift` 一样; 只装了一个中文输入法时用这个

按住 CapsLock 0.3 秒才开 / 关大写锁定 (不用等松开), 并短暂提示 "大写锁定: 开 / 关" (不想要提示可以取消 "显示大写锁定开 / 关的提示"); `Shift+CapsLock` 和原来一样开 / 关大写锁定。默认不开 (`General.CapsLock`)。把 CapsLock 当修饰键的自定义热键 (例如 `CapsLock & J`) 不要和它一起用。通过 Chrome 远程桌面等远程控制使用时也可以: 远程软件为了同步大写锁定补发的极短的 CapsLock 会被忽略。

**管理员窗口里也生效**: Windows 不让普通权限的程序控制以管理员身份运行的窗口, 所以默认在这些窗口里按 CapsLock 不会切换。勾选下面的 "在管理员窗口里也生效" (`General.CapsLockElevated`) 后, ALTRun 建一个计划任务 (这时确认一次 UAC), 之后每次启动都用它以管理员身份运行一个**只管 CapsLock 的小助手** (不再弹 UAC), 退出时一起退出。ALTRun 本身还是普通权限, 从它打开的程序、拖放文件、网络驱动器都不受影响。取消勾选时删掉计划任务 (再确认一次)。移动了 ALTRun 的文件夹后, 把这个选项关掉再打开一次。

### 设置热键
偏好设置里所有的热键 (呼出热键、剪贴板历史、对话框跳转、一键加日期、自定义热键) 都用同一种框, 和 Alfred、PowerToys 一样直接按键录制:
- 框里显示 `Alt+Space`、`Ctrl+Alt+C` 这样的写法。点一下框, 出现 "请按下快捷键...", 按下想要的组合键就记下来
- `Esc` 取消 (保留原来的热键), `Backspace` / `Delete` 清除 (不用这个热键)
- 单独的字母、数字、空格、回车等会影响平时打字, 要加上 `Ctrl`、`Alt` 或 `Win`; `F1` ~ `F24`、`Pause` 等可以单独用
- 录制时 ALTRun 自己的热键暂停, `Alt+Space`、`Win+E` 这类组合也不会触发系统或其它程序的功能
- 保存时, 如果两个全局热键相同 (例如呼出热键和剪贴板历史都是 `Ctrl+Alt+C`), 会先提示。只在某些窗口里生效的热键 (对话框跳转、一键加日期、填了窗口的自定义热键) 可以重复

设置文件里仍然是 AutoHotkey 的写法 (`!` Alt, `^` Ctrl, `+` Shift, `#` Win, 例如 `!Space`), 也可以直接编辑; `CapsLock & J` 这类特殊写法在偏好设置里重新录制之前会原样保留。
