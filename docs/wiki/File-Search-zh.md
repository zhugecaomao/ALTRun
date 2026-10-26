[English](File-Search) · **中文**

# 文件搜索

## 怎样搜索文件
普通搜索只显示应用、命令、片段等, **不会**混进全盘的文件 (避免 `NIRMALA.TTF` 这类无关文件)。要搜文件时:

- **空格**: 在空的搜索框里先按 `空格`, 输入框出现灰色的 "搜索文件...", 再输入名称。已经输入了文字时, 把光标移到最前面 (`Home`) 再按 `空格`, 文字保留并改用文件搜索。输入框为空或光标在最前面时按 `Backspace` 回到普通搜索。和 Alfred 的 Quick File Search 一样
- **`file 报告`**、**`open 报告`**、**`find 报告`** 或 **`'报告`**: 效果相同
- **只搜文件夹**: **`folder 预算`** (关键字可以在 偏好设置 → 文件搜索 → "文件夹搜索关键字" 修改); 文件搜索模式里也可以输入 `folder 预算`。有 Everything 时相当于 Everything 的 `folder:预算`

- **最近打开的文件**: **`recent 报告`** 搜最近在任何程序里打开过的文件和文件夹 (Windows 的 "最近使用的项目"), 最新的在前, 本地和网络上的都有; 只输入 **`recent `** (后面不跟文字) 列出最新的几个。直接读 Windows 的 "最近使用的项目" 文件夹, 不需要 Everything, 也不用内置索引 (Everything 默认搜不到这些: 偏好设置 → 搜索来源 → "过滤条件" 里排除了 `\AppData\`)。关键字可以在 偏好设置 → 文件搜索 → "最近文件关键字" 修改; 文件搜索模式里也可以输入 `recent 报告`。网络上的文件不去访问网络, 网络慢或断开时列表也不会卡住
- **按类型搜索** (和 Listary 的筛选一样): **`doc 报告`** 只搜文档, 还有 `pic` 图片、`video` 视频、`audio` 音频、`zip` 压缩包、`exe` 程序、`cad` (dwg / dxf / dgn / rvt / ifc / skp...)。文件搜索模式里也可以输入 `doc 报告`。关键字和扩展名在 偏好设置 → 文件搜索 → "文件类型" → **编辑...** 里修改, 每行一个 `关键字 = 扩展名 扩展名 ...` (设置页上那一行只显示关键字)。有 Everything 时相当于 `ext:doc;docx;...`
- 按修改日期找可以直接写 Everything 的语法, 例如 `空格` + `dm:today 报告`、`dm:thisweek`

关键字后面要有空格才进入文件搜索: 只输入 `folder`、`open`、`recent`、`doc` 时, 名字里带这个词的命令、应用照常显示, 另有一行这个关键字的提示。名字以这个词开头的应用和命令 (例如 `doc` 对应的 Docker) 排在提示前面, 按 `Enter` 照样打开它们; 没有的话提示排第一。在提示上按 `Enter`、`Tab` 或双击, 补全成关键字加空格 (例如 `recent `)。关键字还没输完整时 (至少 3 个字母, 例如 `fol`、`rec`、`vid`), 提示显示在列表最后。

结果按名称的匹配程度排序: 完全相同 > 名称开头 > 单词开头 (例如 `2026-Budget`) > 包含 (例如 `NewBudget`); 同样的匹配程度, 文件夹排在文件前面, 再按修改时间 (新的在前)。用 Everything 时, ALTRun 先取 Everything 的前 300 条结果再排序, 所以名称最匹配的文件夹不会因为不是最近修改的而漏掉。

Everything 的搜索语法可以直接用, 原样交给 Everything: 例如 `空格` + `folder:预算 2024`、`ext:pdf 报告`、`path:Projects 报告`。

最多显示 30 条; 没有结果时可以一键在 Everything 或 Windows 搜索里继续找。普通搜索一条结果都没有时, 最后也会给出 "搜索文件" 的兜底项。

想让普通搜索也显示文件 (旧的行为): 在 偏好设置 → 文件搜索 勾选 "普通搜索时也显示匹配的文件和文件夹", 最多显示 6 条, 只收录名称开头或单词开头匹配的。

常用的文件夹更适合加为 [自定义命令](Custom-Commands-zh): 自定义命令也按目标的文件夹名匹配, 普通搜索就能找到。

![文件搜索](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/files.png)

![按类型搜索: doc report](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/filetype.png)

`Enter` 打开, `Ctrl+Enter` 在文件管理器中显示, `Alt+Enter` 复制路径, `→` 更多操作 (打开方式、复制 / 剪切文件、复制 / 移动到 TC 当前的文件夹、在此处打开终端、属性、移到回收站...), 见 [操作面板](Usage-zh#操作面板)。

## 浏览文件夹
在搜索框里输入路径, 列出这个文件夹里的内容 (文件夹在前), 路径最后一段用来过滤:

![浏览文件夹](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/browse.png)

| 输入 | 结果 |
|---|---|
| `C:\` `D:\Projects\` | 这个文件夹里的全部内容 |
| `D:\Projects\rep` | 名称和 `rep` 匹配的 (和普通搜索一样的模糊匹配) |
| `~\` | 用户文件夹 (`C:\Users\你的名字`) |
| `%OneDrive%\` | 环境变量 |
| `\\server\share\` | 网络共享 |

`Tab` 进入选中的文件夹, `Backspace` 删掉一段回到上一级; 任何结果里的文件夹 (自定义命令、文件搜索...) 也可以按 `Tab` 进入。`Enter` 打开, `→` 有文件的各种操作。隐藏和系统文件不列出。`Insert` 可以标记多个文件一起复制、移动或删除, 见 [快捷键](Usage-zh#快捷键)。

## 数据来源
ALTRun 自动选择:

| 情况 | 搜索范围 |
|---|---|
| [Everything](https://www.voidtools.com/) 正在运行 | 通过 Everything 的 IPC 接口直接查询**全盘**, 结果即时; 不需要 `Everything64.dll` 或 `es.exe` |
| Everything 没有运行 | 内置索引: 在后台扫描 桌面、文档、下载 (深度 4 层, 最多 30000 项), 缓存在 `Data\FileIndex.json`。这些文件夹里有文件 / 文件夹新增、删除、改名时, 几秒后在后台重新扫描 (两次至少隔 5 分钟; 修改文件内容、浏览器下载中的临时文件不算); 在偏好设置里改了索引的文件夹、深度、排除或最多条数, 点 "应用" 后也会在后台重新扫描 (扫完前旧索引照常可用); 另外默认每 60 分钟兜底扫描一次。ALTRun 启动时 Everything 正在运行的话, 内置索引不读也不扫描, 等 Everything 关掉、真的要用时再读 |

偏好设置 → 搜索来源 页面顶部会显示 Everything 当前是否在运行。

推荐安装 Everything 并让它开机启动 (Everything 的 "选项 → 常规 → 开机启动"), 搜索范围和速度都更好。

## 设置

![偏好设置 → 搜索来源](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-searchsources.png)

偏好设置 → 文件搜索 (怎样开始搜索、文件类型、结果数量) 和 搜索来源 (Everything、内置索引), 或 `ALTRun.json` → `Features.FileSearch`:

| 设置 | 默认 | 说明 |
|---|---|---|
| SpacePrefix | 1 | 在空的搜索框里按空格进入文件搜索模式 |
| InDefaultResults | 0 | 普通搜索时也显示匹配的文件和文件夹 |
| DefaultResultsLimit | 6 | 普通结果里最多显示几条文件 |
| MinQueryLength | 2 | 至少输入几个字才搜文件 |
| Keywords | `open, find, file` | 只搜文件的关键字 |
| FolderKeywords | `folder` | 只搜文件夹的关键字 |
| RecentKeywords | `recent` | 搜最近打开的文件和文件夹 (Windows 的 "最近使用的项目") 的关键字 |
| TypeFilters | doc / pic / video / audio / zip / exe / cad | 文件类型筛选, 每行 `关键字 = 扩展名 ...` |
| QuotePrefix | 1 | 以 `'` 开头只搜文件和文件夹 |
| MaxResults | 30 | 只搜文件时最多显示几条 |
| UseEverything | 1 | Everything 在运行时用它搜索 |
| EverythingFilter | `!C:\Windows\ !\AppData\ !\$Recycle.Bin\` | 追加在 Everything 搜索后面的条件, `!` 表示排除 (Everything 搜索语法) |
| EverythingPath | 空 | Everything.exe 的位置, 只用于 "在 Everything 中搜索"; 留空自动查找 |
| ScopeFolders | 桌面、文档、下载 | 内置索引扫描的文件夹, 可以用 [路径变量](Custom-Commands-zh#路径里可以用的变量) |
| ScopeDepth | 4 | 内置索引的子文件夹深度 |
| ScopeExclude | `node_modules`、`.git`、`.svn`、`.hg`、`__pycache__`、`.venv`、`venv`、`.tox`、`.cache`、`.idea`、`.vs`、回收站 | 内置索引跳过的文件夹 (正则, 对完整路径)。默认跳过代码仓库、Python 虚拟环境和开发工具的缓存: 里面动辄上万个文件, 会占满 MaxEntries, 也没有要找的文档 |
| MaxEntries | 30000 | 内置索引最多收录多少项 |
| RefreshMinutes | 60 | 兜底: 内置索引的缓存超过这么久也在后台重新扫描 (最少 5 分钟; 启动时检查一次, 之后每分钟检查, 键盘鼠标空闲 30 秒、搜索窗口没开时才重建)。平时靠监视文件夹自动更新; 用 Everything 时都不扫描 |

改了内置索引的范围后, 可以在偏好设置里点 "重建文件索引" 立即生效。

## 常见问题
- **Everything 装了但不起作用**: Everything 必须正在运行 (托盘里有图标)。如果 Everything 以管理员身份运行而 ALTRun 不是, Windows 可能拦截两者之间的通信, 让两者以相同的权限运行即可 (推荐 Everything 安装为服务、普通权限运行)。
- **网络驱动器上的文件**: Everything 默认不索引网络驱动器, 需要在 Everything 的 "选项 → 索引 → 文件夹" 里添加; 内置索引可以把网络文件夹加到 ScopeFolders, 但扫描会比较慢。常用的网络文件夹更适合加为 [自定义命令](Custom-Commands-zh)。
- **用拼音搜中文文件名** (例如输入 `cs` 找到 "测试.bat", `zjsm` 找到 "证件扫描文件.zip"): 支持拼音首字母, 找到的文件同样高亮对应的字, 按首字母的匹配程度排序 (同样的匹配程度, 名称直接匹配的排在前面)
  - **没有 Everything (内置索引)**: 直接就能用。载入索引后在后台算好首字母 (1 万个中文文件名约零点几秒), 刚启动的这一小会儿里暂时不按拼音找。只按一个词找: `zj sm` 这样分开写的找不到
  - **用 Everything 时**: 文件是 Everything 找的, 要先让 Everything 支持拼音 (ALTRun 把输入原样交给 Everything, 不用另外设置):
    - **Everything 1.5 (beta)**: 工具 → 选项 → 高级, 搜索 `pinyin`, 设成 `true` (只支持首字母)。还在用 1.5 alpha 的, ALTRun 可能找不到它 (alpha 默认以 "1.5a" 命名实例运行), 请升级到 beta
    - **Everything 1.4**: 安装开源扩展 [IbEverythingExt](https://github.com/Chaoses-Ib/IbEverythingExt), 按它的说明放进 Everything 的文件夹后重启 Everything。支持简拼 (`cs`)、全拼 (`ceshi`)、双拼
    - 输入多个词时 (例如 `zj sm`), Everything 找到的文件照样高亮, 但不按拼音排序
  - 只认常用汉字 (GB2312 一级汉字), 多音字只取一个读音; 全拼只有 Everything + IbEverythingExt 能找到, ALTRun 不按全拼排序和高亮
