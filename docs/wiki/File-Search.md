# 文件搜索

## 怎样搜索文件
普通搜索只显示应用、命令、片段等, **不会**混进全盘的文件 (避免 `NIRMALA.TTF` 这类无关文件)。要搜文件时:

- **空格**: 在空的搜索框里先按 `空格`, 输入框出现灰色的 "搜索文件...", 再输入名称。输入框为空时按 `Backspace` 回到普通搜索。和 Alfred 的 Quick File Search 一样
- **`'报告`**、**`open 报告`** 或 **`find 报告`**: 效果相同
- **只搜文件夹**: **`folder bk`** (关键字可以在 偏好设置 → 文件搜索 → "文件夹搜索关键字" 修改); 文件搜索模式里也可以输入 `folder bk`。有 Everything 时相当于 Everything 的 `folder:bk`

- **按类型搜索** (和 Listary 的筛选一样): **`doc 报告`** 只搜文档, 还有 `pic` 图片、`video` 视频、`audio` 音频、`zip` 压缩包、`exe` 程序、`cad` (dwg / dxf / dgn / rvt / ifc / skp...)。文件搜索模式里也可以输入 `doc 报告`。关键字和扩展名可以在 偏好设置 → 文件搜索 → "文件类型" 修改, 每行一个 `关键字 = 扩展名 扩展名 ...`。有 Everything 时相当于 `ext:doc;docx;...`
- 按修改日期找可以直接写 Everything 的语法, 例如 `空格` + `dm:today 报告`、`dm:thisweek`

关键字后面要有空格才进入文件搜索: 只输入 `folder`、`open`、`doc` 时, 名字里带这个词的命令、应用照常显示。

结果按名称的匹配程度排序: 完全相同 > 名称开头 > 单词开头 (例如 `PT2310-BK`) > 包含 (例如 `notebk`); 同样的匹配程度, 文件夹排在文件前面, 再按修改时间 (新的在前)。用 Everything 时, ALTRun 先取 Everything 的前 300 条结果再排序, 所以名称最匹配的文件夹不会因为不是最近修改的而漏掉。

Everything 的搜索语法可以直接用, 原样交给 Everything: 例如 `空格` + `folder:bk 2024`、`ext:pdf 报告`、`path:Tender 报告`。

最多显示 30 条; 没有结果时可以一键在 Everything 或 Windows 搜索里继续找。普通搜索一条结果都没有时, 最后也会给出 "搜索文件" 的兜底项。

想让普通搜索也显示文件 (旧的行为): 在 偏好设置 → 文件搜索 勾选 "普通搜索时也显示匹配的文件和文件夹", 最多显示 6 条, 只收录名称开头或单词开头匹配的。

常用的文件夹更适合加为 [自定义命令](Commands-and-Snippets): 自定义命令也按目标的文件夹名匹配, 普通搜索就能找到。

![文件搜索](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/files.png)

`Enter` 打开, `Ctrl+Enter` 在文件管理器中显示, `Alt+Enter` 复制路径, `→` 更多操作 (打开方式、复制 / 剪切文件、复制 / 移动到 TC 当前的文件夹、在此处打开终端、属性、移到回收站...), 见 [操作面板](Usage#操作面板)。

## 浏览文件夹
在搜索框里输入路径, 列出这个文件夹里的内容 (文件夹在前), 路径最后一段用来过滤:

| 输入 | 结果 |
|---|---|
| `C:\` `D:\Projects\` | 这个文件夹里的全部内容 |
| `D:\Projects\rep` | 名称和 `rep` 匹配的 (和普通搜索一样的模糊匹配) |
| `~\` | 用户文件夹 (`C:\Users\你的名字`) |
| `%OneDrive%\` | 环境变量 |
| `\\server\share\` | 网络共享 |

`Tab` 进入选中的文件夹, `Backspace` 删掉一段回到上一级; 任何结果里的文件夹 (自定义命令、文件搜索...) 也可以按 `Tab` 进入。`Enter` 打开, `→` 有文件的各种操作。隐藏和系统文件不列出。`Insert` 可以标记多个文件一起复制、移动或删除, 见 [快捷键](Usage#快捷键)。

## 数据来源
ALTRun 自动选择:

| 情况 | 搜索范围 |
|---|---|
| [Everything](https://www.voidtools.com/) 正在运行 | 通过 Everything 的 IPC 接口直接查询**全盘**, 结果即时; 不需要 `Everything64.dll` 或 `es.exe` |
| Everything 没有运行 | 内置索引: 在后台扫描 桌面、文档、下载 (深度 4 层, 最多 30000 项), 缓存在 `Data\FileIndex.json`, 超过 30 分钟启动时重新扫描。ALTRun 启动时 Everything 正在运行的话, 内置索引不读也不扫描, 等 Everything 关掉、真的要用时再读 |

偏好设置 → 搜索来源 页面顶部会显示 Everything 当前是否在运行。

推荐安装 Everything 并让它开机启动 (Everything 的 "选项 → 常规 → 开机启动"), 搜索范围和速度都更好。

## 设置
偏好设置 → 文件搜索 (怎样开始搜索、文件类型、结果数量) 和 搜索来源 (Everything、内置索引), 或 `ALTRun.json` → `Features.FileSearch`:

| 设置 | 默认 | 说明 |
|---|---|---|
| SpacePrefix | 1 | 在空的搜索框里按空格进入文件搜索模式 |
| InDefaultResults | 0 | 普通搜索时也显示匹配的文件和文件夹 |
| DefaultResultsLimit | 6 | 普通结果里最多显示几条文件 |
| MinQueryLength | 2 | 至少输入几个字才搜文件 |
| Keywords | `open, find` | 只搜文件的关键字 |
| FolderKeywords | `folder` | 只搜文件夹的关键字 |
| TypeFilters | doc / pic / video / audio / zip / exe / cad | 文件类型筛选, 每行 `关键字 = 扩展名 ...` |
| QuotePrefix | 1 | 以 `'` 开头只搜文件和文件夹 |
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
