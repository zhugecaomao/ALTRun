# 扩展功能

搜索窗口以外的功能, 在 偏好设置 → 扩展功能 里开关和设置。

## 对话框快速跳转
在标准的 "打开 / 保存文件" 对话框里:
- `Ctrl+G`: 跳到 Total Commander 当前打开的文件夹
- `Ctrl+E`: 跳到资源管理器当前打开的文件夹

对话框标题上会提示这两个热键。打开 "自动跳转" (`AutoSwitch`) 后, 从 Total Commander 切换到对话框时会自动跳转。

设置: `Extensions.QuickSwitch` → `TotalCmdHotkey` / `ExplorerHotkey` / `AutoSwitch` / `DialogWindows` / `ExcludeWindows`。

## Ctrl+D 加日期
- **重命名文件时** (资源管理器、Total Commander、桌面、对话框): 在扩展名前加上 ` - 日期`, 已经有日期的更新为今天。例如 `Report.docx` → `Report - 23.09.2026.docx`
- **文字备注框里** (例如 Total Commander 的文件备注): 在末尾加上 ` - 日期`

日期格式 `DateFormat` 默认 `dd.MM.yyyy` (写法见 [AutoHotkey FormatTime](https://www.autohotkey.com/docs/v2/lib/FormatTime.htm)), 片段里的 `{date}` 也用这个格式。

## 自定义热键
偏好设置 → 自定义热键: 给任意热键指定一条 [系统命令](#系统命令), 可以限定只在某个窗口里生效。

| 字段 | 例子 | 说明 |
|---|---|---|
| 热键 (Key) | `^!p` | AutoHotkey 写法: `!` Alt, `^` Ctrl, `+` Shift, `#` Win; `~` 表示不拦截原来的按键 |
| 操作 (Action) | `PTTools` | 系统命令的 Id, 或 `ToggleWindow` (显示 / 隐藏 ALTRun) |
| 窗口 (WinTitle) | `ahk_exe RAPTW.exe` | 留空 = 全局; 否则只在匹配的窗口里生效 |

默认的一条: 在 RAPT (`RAPTW.exe`) 里按鼠标中键打开 PT 工具箱。

## 系统命令
直接在搜索框里输入名称 (中文、英文或 Id 都可以), 也可以用在自定义热键里。

| Id | 命令 |
|---|---|
| **ALTRun** | |
| `Preferences` | ALTRun 偏好设置 |
| `Reload` | 重新载入 ALTRun |
| `RebuildIndex` | 重建 ALTRun 索引 |
| `CheckUpdate` | 检查更新 |
| `About` | 关于 ALTRun |
| `Log` | 打开 ALTRun 日志 |
| `Quit` | 退出 ALTRun |
| **系统** | |
| `Lock` `Sleep` `Hibernate` | 锁屏 / 睡眠 / 休眠 |
| `Shutdown` `Restart` `Logoff` | 关机 / 重启 / 注销 (执行前确认) |
| `EmptyRecycle` | 清空回收站 (执行前确认) |
| `MonitorOff` | 关闭显示器 |
| `Mute` `VolumeUp` `VolumeDown` | 静音 / 音量 +10 / 音量 -10 |
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

打开 `Features.Calculator.StructuralCalc` (偏好设置 → 功能) 后, 结果下方附带两行结构计算:
- 把结果当作 **梁宽 (mm)**: 主筋根数和间距 (保护层 40 mm, 最大间距 300 mm)
- 把结果当作 **配筋面积 As (mm²)**: H13 / H16 / H20 / H25 / H32 需要的根数

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
