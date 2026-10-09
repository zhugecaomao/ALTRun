[English](en-Extensions) · **中文**

# 扩展功能

搜索窗口以外的功能。对话框快速跳转和一键加日期各有一页设置 (偏好设置 → 对话框跳转 / 一键加日期)。

## 对话框快速跳转
做法借鉴 [Listary](https://www.listary.com/) 的 Quick Switch。

**文件夹面板** (默认打开, 和 Listary 的 Quick Switch 窗口一样): 对话框一出现, 正下方就自动贴一个和对话框一样宽的面板 (下面放不下时放在上面), 点一个文件夹对话框就跳过去, 不用记热键:
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

## 一键加日期
按热键 (默认 `Ctrl+D`, 可以改):
- **重命名文件时** (资源管理器、Total Commander、桌面、对话框): 在扩展名前加上 ` - 日期`, 已经有日期的更新为今天。例如 `Report.docx` → `Report - 23.09.2026.docx`
- **文字备注框里** (例如 Total Commander 的文件备注): 在末尾加上 ` - 日期`

偏好设置 → 一键加日期: 两种场景各有自己的热键和 "生效的窗口" 列表 (每行一个, 例如 `ahk_class CabinetWClass` 资源管理器、`ahk_class TTOTAL_CMD` Total Commander)。设置文件里是 `Extensions.AutoDate` → `RenameHotkey` / `RenameWindows` / `AppendHotkey` / `AppendWindows`。

日期格式 `DateFormat` 默认 `dd.MM.yyyy` (写法见 [AutoHotkey FormatTime](https://www.autohotkey.com/docs/v2/lib/FormatTime.htm)), 片段里的 `{date}` 也用这个格式。

## 自定义热键
偏好设置 → 自定义热键: 给任意热键指定一条 [系统命令](#系统命令), 可以限定只在某个窗口里生效。

| 字段 | 例子 | 说明 |
|---|---|---|
| 热键 (Key) | `Ctrl+Alt+P` | 点一下框, 按下组合键 (见下文 [设置热键](#设置热键)); 鼠标热键: 在框里点鼠标中键或侧键。勾选 "保留按键原来的功能" 时, 按键原来的作用照常 (设置文件里是开头的 `~`) |
| 操作 (Action) | `PTTools` | 系统命令的 Id, 或 `ToggleWindow` (显示 / 隐藏 ALTRun) |
| 窗口 (WinTitle) | `ahk_exe notepad.exe` | 留空 = 全局; 否则只在匹配的窗口里生效; `ALTRun` = 只在 ALTRun 的搜索窗口里 |

默认的一条: 在 RAPT (`RAPTW.exe`) 里按鼠标中键打开 PT 工具箱。搜索窗口里的 `F1` ~ `F4` 是内置的, 见 [使用方法](Usage#快捷键)。

### 设置热键
偏好设置里所有的热键 (呼出热键、剪贴板历史、对话框跳转、一键加日期、自定义热键) 都用同一种框, 和 Alfred、PowerToys 一样直接按键录制:
- 框里显示 `Alt+Space`、`Ctrl+Alt+C` 这样的写法。点一下框, 出现 "请按下快捷键...", 按下想要的组合键就记下来
- `Esc` 取消 (保留原来的热键), `Backspace` / `Delete` 清除 (不用这个热键)
- 单独的字母、数字、空格、回车等会影响平时打字, 要加上 `Ctrl`、`Alt` 或 `Win`; `F1` ~ `F24`、`Pause` 等可以单独用
- 录制时 ALTRun 自己的热键暂停, `Alt+Space`、`Win+E` 这类组合也不会触发系统或其它程序的功能
- 保存时, 如果两个全局热键相同 (例如呼出热键和剪贴板历史都是 `Ctrl+Alt+C`), 会先提示。只在某些窗口里生效的热键 (对话框跳转、一键加日期、填了窗口的自定义热键) 可以重复

设置文件里仍然是 AutoHotkey 的写法 (`!` Alt, `^` Ctrl, `+` Shift, `#` Win, 例如 `!Space`), 也可以直接编辑; `CapsLock & J` 这类特殊写法在偏好设置里重新录制之前会原样保留。

## 系统命令
直接在搜索框里输入名称 (中文、英文或 Id 都可以), 也可以用在自定义热键里。

![系统命令: lock](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/system.png)

| Id | 命令 |
|---|---|
| **ALTRun** | |
| `Preferences` | ALTRun 偏好设置 |
| `Preferences.Appearance` 等 | ALTRun 偏好设置: 外观 (每个设置页一条, 直接打开那一页; 输入页名就能找到, 例如 `外观`、`hotkeys`) |
| `Reload` | 重新载入 ALTRun |
| `RebuildIndex` | 重建 ALTRun 索引 |
| `CheckUpdate` | 检查更新 |
| `About` | 关于 ALTRun: 打开偏好设置的高级页 (版本、项目主页、检查更新) |
| `Log` | 打开 ALTRun 日志 |
| `Quit` | 退出 ALTRun |
| **系统** | |
| `Lock` `Sleep` `Hibernate` | 锁屏 / 睡眠 / 休眠 |
| `Shutdown` `Restart` `Logoff` | 关机 / 重启 / 注销 (执行前确认) |
| `EmptyRecycle` | 清空回收站 (执行前确认) |
| `MonitorOff` | 关闭显示器 |
| `Mute` `VolumeUp` `VolumeDown` | 静音 / 音量 +10 / 音量 -10 |
| `MediaPlayPause` `MediaNext` `MediaPrev` `MediaStop` | 播放 / 暂停、下一首、上一首、停止播放 (音乐和视频播放器) |
| `ShowIP` | 显示 IP 地址 |
| `TerminalHere` | 在当前文件夹 (资源管理器 / Total Commander) 打开终端 |
| `ListProcesses` `ListServices` | 列出运行中的进程 / 服务 |
| `PTTools` `SPF2M` | PT 工具箱 / SPF2M |
| **剪贴板文字** (处理后放回剪贴板) | |
| `TextUpper` `TextLower` `TextTitle` | 大写 / 小写 / 首字母大写 |
| `TextSortAsc` `TextSortDesc` | 按行排序 |
| `TextTrimLines` `TextRemoveBlank` `TextDedupe` | 去掉行首尾空格 / 删除空行 / 删除重复行 |
| `TextReverse` | 反转 |
| `TextToTraditional` `TextToSimplified` | 简体转繁体 / 繁体转简体 |
| `TextUrlEncode` | URL 编码 |
| **Windows 工具** | |
| | 任务管理器、控制面板、设置、设备管理器、服务、注册表编辑器、事件查看器、磁盘管理、计算机管理、任务计划程序、程序和功能、系统属性、网络连接、防火墙、资源监视器、磁盘清理、系统配置、组策略、命令提示符、PowerShell、资源管理器、回收站、此电脑、打印机、记事本、计算器、画图、Windows 版本 |

关机、重启等命令的确认可以用 `Features.System.ConfirmActions = 0` 关闭。

## 终端
输入 `>命令` 在终端里运行, 窗口保持打开, 例如 `>ipconfig /all`、`>ping 8.8.8.8`。操作面板里的 "在此处打开终端" 在文件所在文件夹打开终端。

终端程序 `Features.Terminal.Shell`: `cmd` (默认) / `powershell` / `pwsh` / `wt` (Windows Terminal); 前缀 `Prefix` 默认 `>`。

## 计算器
直接输入算式: `12*(3+4)`、`=2^10`, `Enter` 复制结果。支持 `+ - * /`、乘方 `^` (或 `**`) 和括号, 结果最多保留两位小数 (`10/3` 显示 `3.33`; 不到 0.005 的数保留到第一位有效数字后一位, 例如 `0.004`)。

在 偏好设置 → 计算器 打开 "附带结构计算" (`Features.Calculator.StructuralCalc`) 后, 结果大于 2 × 保护层时下方附带两行结构计算:
- 把结果当作 **梁宽 (mm)**: 主筋根数和间距。根数 = ⌈(梁宽 − 2 × 保护层) / 最大间距⌉ + 1, 间距 = (梁宽 − 2 × 保护层) / (根数 − 1)
- 把结果当作 **配筋面积 As (mm²)**: 每种直径需要的根数 (⌈As / 一根的面积⌉, 一根的面积 = π d² / 4)

例如 `300*2` (600) 默认显示 `Beam width 600 mm: 3 main bars @ 260 c/c` 和 `As = 600 mm²: 5H13  3H16  2H20  2H25  1H32`。参数在同一页里修改:

| 设置 | 键 | 默认 | 说明 |
|---|---|---|---|
| 保护层 (Rebar cover) | `RebarCover` | `40` | 每边, mm |
| 最大间距 | `MaxBarSpacing` | `300` | mm, 间距大于它时增加主筋 |
| 钢筋直径 | `BarSizes` | `[13, 16, 20, 25, 32]` | 配筋面积那一行列出的直径 (mm) |
| 钢筋代号 | `BarPrefix` | `H` | 钢筋代号, 例如 `T`、`Y`、`Φ` |

![偏好设置 → 计算器](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-calculator.png)

### 单位和货币换算
写法: 数值 + 单位 + `in` / `to` / `=` / `->` / `转` + 目标单位, 单位不区分大小写, `m²` 可以写成 `m2`。例如 `10 km in mi`、`5ft to cm`、`100 f to c`、`3 亩 in m2`、`1 GB to MB`、`300 kN to kip`、`20 MPa in psi`。`Enter` 复制数字。

| 类别 | 单位 |
|---|---|
| 长度 | mm cm m km in ft yd mi nmi (毫米 厘米 米 公里 英寸 英尺 码 英里 海里) |
| 质量 | mg g kg t lb oz 斤 两 |
| 面积 | mm2 cm2 m2 km2 ha acre ft2 in2 亩 |
| 体积 | ml cl l m3 cm3 ft3 gal qt pt cup floz |
| 速度 | m/s km/h mph knot ft/s |
| 时间 | ms s min h day week year |
| 数据 | bit B KB MB GB TB (按 1024) |
| 压强 / 应力 | Pa kPa MPa (N/mm2) bar psi ksi atm psf |
| 力 | N kN lbf kip kgf tf |
| 能量 / 功率 | J kJ cal kcal Wh kWh / W kW hp |
| 温度 | C F K |

**货币换算** (`100 usd to sgd`、`100 美元 to 人民币`) 默认关闭: 在 偏好设置 → 计算器 勾选 "货币换算" (`Features.Calculator.Currency`)。打开后每天从 [Frankfurter](https://frankfurter.dev) 下载一次欧洲央行等央行公布的参考汇率 (免费, 不需要注册), 保存在 `Data\Currency.json`; 这是 ALTRun 除 GitHub 之外唯一会访问的网站。结果里会显示汇率的日期。

## 脚本扩展
把自己的脚本放进 `Scripts\` 文件夹 (程序目录里; 偏好设置 → 脚本 → "打开脚本文件夹", 或者搜索 "打开脚本文件夹", 第一次打开时会建几个示例: 后台运行显示通知的、关键字后面带参数的、用文本编辑器看输出的), 就能在搜索窗口里按名称或关键字找到并运行, 和 Raycast 的 Script Commands 一样。支持 `.ahk` (用 ALTRun 自带的 AutoHotkey 运行, 不用另外安装)、`.ps1`、`.bat` / `.cmd`、`.py` (需要装 Python)。偏好设置 → 脚本 列出找到的脚本, 双击用记事本编辑:

![偏好设置 → 脚本](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-scripts.png)

脚本开头的注释里可以写这些设置 (注释符 `;` `#` `REM` `::` `//` 都行, 都可以不写):

| 设置 | 说明 |
|---|---|
| `@altrun.title 名称` | 显示的名称, 默认用文件名 |
| `@altrun.keyword 关键字` | 输入完全相同时排在最前面 |
| `@altrun.argument 提示` | 需要参数: 输入 "关键字 文字", 文字作为第一个参数传给脚本; 只搜到名称时 `Enter` 补全成 "关键字 " |
| `@altrun.mode window` | 默认: 正常运行 (有窗口) |
| `@altrun.mode silent` | 在后台运行, 结束后把输出的最后一行显示成通知 |
| `@altrun.mode output` | 在后台运行, 结束后用 `.txt` 的默认程序 (一般是记事本) 打开全部输出 |

例子 (`Scripts\Ping.ps1`):
```powershell
# @altrun.title    Ping
# @altrun.keyword  ping
# @altrun.argument 主机名或 IP
# @altrun.mode     output
ping $args[0]
```
输入 `ping 10.0.0.1` 后 `Enter`, 结果用记事本显示。

`.ahk` 脚本里用 `FileAppend("文字", "*")` 输出。增删、修改脚本后马上生效, 不用重新载入; 选中脚本按 `F3` 用记事本编辑。后台运行的脚本最多等 2 分钟。

## PT 工具箱
预应力设计用的小工具, 输入 `PTTools` / `SPF2M` 或用自定义热键打开:
- **PT Tools**: 钢筋 / BRC 面积计算器, 以及算式计算
- **SPF2M Post-Tensioning Tendon Profile Calculator (束线型计算)**: 直接在 ALTRun 里计算, 不再需要 DOSBox。选择线型 (双抛物线 / 抛物线-直线-抛物线 / 抛物线-直线 / 直线-抛物线) 和钢绞线类型, 输入起点、终点标高和水平距离, 右边的表格立即给出每个支架处的高度 (Actual、Beam 按 5mm、Slab 按 10mm 取整), 以及曲率半径和反弯点位置
  - 结果和原来的 SPF2M.EXE 完全一致: 在 DOSBox 里用原程序跑了 139 组 (4 种线型 × 6 种钢绞线, 上升 / 下降, 非整米跨度, 自定义半径 / 反弯点 / 支架间距), 1100 多个数值逐个比对, 作为单元测试的一部分
  - 可选项留空 = SPF2M 的默认值 (灰色显示): **最小曲率半径** (Slab 5000、7S 3200、12S 4200、19S 5300、22S 5700、31S 6700), **反弯点距离** (指定后反算曲率半径, 小于最小半径时提示), **支架间距** (默认最大 1000mm, 零头放在高点一侧; 也可以写 `500, 1500, 800 ...`, 加起来要等于水平距离)
  - **At C.G.**: 标高给的是束中心时勾选, 自动减去半个管道直径。管道直径留空时用默认值 (Slab 25、7S 70、12S 90、19S 100、22S 120、31S 130), 填了则每种钢绞线各自记住
  - 表格从高点开始 (和 SPF2M 一样); **Copy Table** 复制成表格, 直接粘贴到 Excel
  - SPF2M 在某些无法实现的线型上会直接退出, 这里会给出说明; 反弯点超过半跨时 SPF2M 什么也不显示, 这里也会提示
  - 原来的 SPF2M.EXE 和 DOSBox 已经不再随程序发布。从旧版本升级后, `Resources\` 里的 `DOSBox.exe`、`SDL.dll`、`SDL_net.dll`、`SPF2M.exe`、`Run.bat` 可以删掉
