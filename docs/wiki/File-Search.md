# 文件搜索

## 两种用法
- **直接输入名称**: 匹配的文件和文件夹显示在应用、命令下面 (最多 6 条, 至少输入 2 个字), 只收录名称开头或单词开头匹配的
- **只搜文件**: `'报告`、`open 报告` 或 `find 报告`, 最多显示 30 条; 没有结果时可以一键在 Everything 或 Windows 搜索里继续找

`Enter` 打开, `Ctrl+Enter` 在文件管理器中显示, `Alt+Enter` 复制路径, `→` 更多操作 (在此处打开终端、属性...)。

## 数据来源
ALTRun 自动选择:

| 情况 | 搜索范围 |
|---|---|
| [Everything](https://www.voidtools.com/) 正在运行 | 通过 Everything 的 IPC 接口直接查询**全盘**, 结果即时; 不需要 `Everything64.dll` 或 `es.exe` |
| Everything 没有运行 | 内置索引: 在后台扫描 桌面、文档、下载 (深度 4 层, 最多 30000 项), 缓存在 `Data\FileIndex.json`, 每 30 分钟更新 |

偏好设置 → 文件搜索 页面顶部会显示 Everything 当前是否在运行。

推荐安装 Everything 并让它开机启动 (Everything 的 "选项 → 常规 → 开机启动"), 搜索范围和速度都更好。

## 设置
偏好设置 → 文件搜索, 或 `ALTRun.json` → `Features.FileSearch`:

| 设置 | 默认 | 说明 |
|---|---|---|
| InDefaultResults | 1 | 直接输入名称时也显示匹配的文件和文件夹 |
| DefaultResultsLimit | 6 | 普通结果里最多显示几条文件 |
| MinQueryLength | 2 | 至少输入几个字才搜文件 |
| Keywords | `open, find` | 只搜文件的关键字 |
| QuotePrefix | 1 | 以 `'` 开头只搜文件 |
| MaxResults | 30 | 只搜文件时最多显示几条 |
| UseEverything | 1 | Everything 在运行时用它搜索 |
| EverythingFilter | `!C:\Windows\ !\AppData\ !\$Recycle.Bin\` | 追加在 Everything 搜索后面的条件, `!` 表示排除 (Everything 搜索语法) |
| EverythingPath | 空 | Everything.exe 的位置, 只用于 "在 Everything 中搜索"; 留空自动查找 |
| ScopeFolders | 桌面、文档、下载 | 内置索引扫描的文件夹, 可以用 [路径变量](Commands-and-Snippets#路径里可以用的变量) |
| ScopeDepth | 4 | 内置索引的子文件夹深度 |
| ScopeExclude | `node_modules`、`.git`、`__pycache__`、回收站 | 内置索引跳过的文件夹 (正则) |
| MaxEntries | 30000 | 内置索引最多收录多少项 |
| RefreshMinutes | 30 | 内置索引多久更新一次 |

改了内置索引的范围后, 可以在偏好设置里点 "重建文件索引" 立即生效。

## 常见问题
- **Everything 装了但不起作用**: Everything 必须正在运行 (托盘里有图标)。如果 Everything 以管理员身份运行而 ALTRun 不是, Windows 可能拦截两者之间的通信, 让两者以相同的权限运行即可 (推荐 Everything 安装为服务、普通权限运行)。
- **网络驱动器上的文件**: Everything 默认不索引网络驱动器, 需要在 Everything 的 "选项 → 索引 → 文件夹" 里添加; 内置索引可以把网络文件夹加到 ScopeFolders, 但扫描会比较慢。常用的网络文件夹更适合加为 [自定义命令](Commands-and-Snippets)。
