[English](en-Extensions) · **中文**

# 扩展功能

搜索窗口以外的功能。下面几项各有独立的页面, 偏好设置里对应的页面按 `F1` 也会打开它们:

| 页面 | 内容 |
|---|---|
| [对话框快速跳转](Quick-Switch) | 在打开 / 保存对话框里跳到 Total Commander 或资源管理器已经打开的文件夹 (`Ctrl+G` / `Ctrl+E`), 自动跳转 |
| [对话框文件夹面板](Quick-Switch-Panel) | 对话框下面的文件夹面板: 列出打开的文件夹和最近的文件夹, 搜索框 (`Ctrl+Shift+G`) |
| [一键加日期](Date-Stamp) | 重命名文件或在文字末尾加上日期 (`Ctrl+D`) |
| [自定义热键](Hotkeys) | 给热键指定系统命令; CapsLock 切换输入法; 怎样设置热键 |
| [计算器](Calculator) | 算式、单位和货币换算、附带的结构计算 |
| [脚本扩展](Scripts) | 把自己的脚本变成搜索结果里的命令 |

下面是没有单独页面的几项。

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
