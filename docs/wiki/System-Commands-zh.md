[English](System-Commands) · **中文**

# 系统命令和终端

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
