[English](Sync-Tasks) · **中文**

# 同步任务

用 Windows 自带的 `robocopy` 让一个文件夹和另一个保持一致, 例如把服务器共享同步到 OneDrive。ALTRun 在后台运行它, 先检查危险的设置, 完成后告诉你结果。

> 偏好设置里的编辑界面在之后的版本里提供; 在那之前, 任务写在设置文件里 (见[下面](#设置文件里的写法))。

## 运行任务
- 输入任务名, 或输入 `sync ` 列出所有任务 (`sync proj` 只搜任务)。**Enter** 运行; 运行中再按 Enter 会停止。
- `→` 打开操作: 运行、预览 (不做任何改动)、打开日志、打开源 / 目标文件夹。
- 结果里显示任务做什么、现在怎样: *10:05 开始运行*、*10:05 开始持续同步*、*上次 09:30: 成功*。
- 任务结束时提示结果 (复制和删除了多少文件)。失败时可以直接打开日志。

## 安全措施
- **镜像** (`/MIR`) 让目标和源完全一致: 目标里源中没有的文件会被**删除**。镜像运行前, ALTRun 先预览 (`/L`, 不做任何改动); 如果会删除东西, 先告诉你要删多少个文件和文件夹, 确认后才运行。
- 源文件夹找不到或是空的 (例如服务器共享没连上) 时不运行镜像, 否则会把目标清空。
- 源和目标是同一个文件夹、一个在另一个里面, 或者镜像到盘符、共享、OneDrive 文件夹、用户文件夹的根目录 (会删掉那里其他所有东西) 时, 不运行。
- **持续同步** (`/MON`、`/MOT`) 的任务会一直运行, 有变化后再同步。退出 ALTRun 时会停止; ALTRun 只是重新启动 (改了设置、升级) 时会继续运行。

## 日志
每次运行都有日志: 写到任务自己的日志文件, 或 `%TEMP%\ALTRun\SyncLogs\<任务名>_<日期>.log`。**打开日志** 打开当天的日志。
- 日志不放在 Data 文件夹里: Data 在 OneDrive 里时, 不断变大的日志会一直上传。
- `%TEMP%\ALTRun\SyncLogs` 里超过 30 天的日志在运行任务时自动删除 (`Features.Sync.KeepLogDays`, `0` = 全部保留)。任务自己的日志文件不会删除。
- 每次运行在 `ALTRun.log` 里记一行: 什么时候开始、结果怎样 (复制、删除了多少文件, 或退出码)。

## 设置文件里的写法
```jsonc
"SyncTasks": [
  {
    "Name": "Projects",
    "Source": "\\\\fileserver\\Projects",
    "Destination": "%OneDrive%\\Work\\Projects",
    "Files": "*.*",
    "Mode": "Mirror",            // Mirror (/MIR 镜像)、Update (/E /XO 更新: 只复制新的和更新过的文件, 不删除)、Copy (/E 全部复制: 不同的都复制, 不删除) 或 Top (只复制顶层)
    "Copy": "DAT", "DirCopy": "DA",
    "Retries": 1, "Wait": 1, "Threads": 16,
    "Monitor": 0, "MonitorChanges": 1, "MonitorMinutes": 5,
    "ExcludeFiles": ["thumbs.db", "desktop.ini", "~$*.*"],
    "ExcludeDirs": ["System Volume Information", "$Recycle.Bin"],
    "LogFile": "%OneDrive%\\Logs\\Projects_{date}.log", "AppendLog": 1,
    "Extra": "/NP"               // 其他 robocopy 参数, 原样使用
  }
]
```
路径可以用 `%环境变量%`; 日志文件名可以用 `{date}` (yyyyMMdd) 和 `{time}` (HHmmss)。
