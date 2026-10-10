[English](en-Commands-and-Snippets) · **中文**

# 自定义命令与片段

## 自定义命令
把常用的文件、文件夹、程序或网址加进 ALTRun, 用名字或关键字打开。

![偏好设置里的自定义命令](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-commands.png)

### 添加
- 搜索结果里选中应用 / 文件 / 文件夹 / 网址, 按 `F3` (或 `→` → "添加到自定义命令...")
- 没有搜到结果时按 `F3`, 用输入的文字新建一条
- 资源管理器里右键 → 发送到 → ALTRun: 选中 **1 个** 文件 / 文件夹时弹出编辑对话框, 名称、类型、目标已经填好, 改好名称或加上关键字后按 `Enter`; 取消则不添加。选中 **多个** 时全部直接添加 (不会一个个弹对话框), 已经有的命令不重复添加。发送一个已经有的文件 / 文件夹时, 打开已有的那条命令修改。偏好设置开着时, 新命令加进 偏好设置 → 自定义命令 的列表, 按 确定 / 应用 保存
- 偏好设置 → 自定义命令 → 添加

已有的命令: 搜索到后按 `F3` 修改, `Ctrl+Del` 删除; 或在偏好设置里双击编辑。命令多的时候, 在列表右下角的 **筛选** 框里输入名称、路径或关键字里的词, 只显示匹配的命令 (文字片段、网页搜索、自定义热键的列表也有)。

### 检查路径
文件夹改名、文件移走后, 指向它的命令就失效了。偏好设置 → 自定义命令 → **检查路径** 会检查每条命令的目标是否还在, 结果显示在 "状态" 列, 并选中第一条有问题的命令, 双击即可修改路径:

| 状态 | 含义 |
|---|---|
| 正常 | 目标存在 |
| 找不到 | 文件或文件夹不在了 (改名、移动或删除), 需要修改路径或删除命令 |
| 无法访问 | 所在的驱动器或网络位置现在连不上 (例如网络盘没有连接), 无法判断, 连上后再检查 |
| - | 网址和 `shell:` 等位置不检查 |

`Command` 类型只写程序名 (例如 `cmd.exe`) 时, 在 PATH 和系统登记的程序 (App Paths) 里查找。检查结果只显示在偏好设置里, 不会修改或删除任何命令。

### 字段
编辑对话框里每个字段下面都有一行灰色的说明。

| 字段 | 说明 |
|---|---|
| 名称 (Title) | 显示的名称, 按它搜索 |
| 类型 (Type) | `File` 文件 / 程序, `Folder` 文件夹, `Command` 程序 + 参数, `Url` 网址 / 链接 (网页, 或 `ms-settings:windowsupdate`、`mailto:` 这样的链接) |
| 目标 (Target) | 路径、程序或网址 |
| 参数 (Arguments) | 运行程序时附带的命令行参数, 例如 `/k ipconfig /all` |
| 关键字 (Keyword) | 可选。输入完全相同的关键字时排在最前面 |

搜索时除了名称和关键字, `File` / `Folder` 类型还会匹配目标的文件名 (不含扩展名) 或文件夹名: Target 是 `Q:\Projects\PT1931 - 24 NIR`, 名称是 `CKR, EA, JIB`, 输入 `nir` 或 `1931` 也能找到。同样的匹配程度, 名称匹配排在前面。

文件夹用 偏好设置 → 通用 → 文件管理器 打开 (默认资源管理器, 也可以设为 Total Commander 等)。

### 带参数的命令 ({query})
目标或参数里写 `{query}`, 再设一个关键字, 输入 "关键字 文字" 时 `{query}` 换成后面的文字 (网址里自动编码), 和网页搜索的关键字一样用:

| 名称 | 类型 | 目标 | 参数 | 关键字 | 输入 |
|---|---|---|---|---|---|
| Jira | 网址 | `https://jira.example.com/browse/{query}` | | `jira` | `jira ABC-123` |
| Ping | 命令行 | `cmd.exe` | `/k ping {query}` | `ping` | `ping 10.0.0.1` |
| 项目文件夹 | 文件夹 | `D:\Projects\{query}` | | `pj` | `pj 2026-05` |
| Google Maps | 网址 | `https://www.google.com/maps/search/{query}` | | `maps` | `maps Changi Airport` |

新安装时默认命令里已经有 Ping、Google Maps 和 Documents subfolder (`docs 文件夹名` 打开 "文档" 里的子文件夹) 三个例子。只输入名称搜到这条命令时, 下面一行会说明接着输入什么 (例如 "在 'ping ' 后面输入内容, 再按 Enter"), 按 `Enter` 补全成 "关键字 ", 接着输入参数。没有关键字时 `{query}` 换成空。"检查路径" 不检查带 `{query}` 的路径。

### 路径里可以用的变量
| 写法 | 含义 |
|---|---|
| `A_Desktop` `A_DesktopCommon` | 桌面 / 公共桌面 |
| `A_MyDocuments` | 文档 |
| `A_AppData` | `%AppData%` |
| `A_Programs` `A_ProgramsCommon` | 开始菜单 "程序" (当前用户 / 所有用户) |
| `A_StartMenu` `A_StartMenuCommon` | 开始菜单 |
| `A_Startup` `A_StartupCommon` | 启动文件夹 |
| `A_ProgramFiles` `A_WinDir` `A_Temp` | Program Files / Windows / 临时文件夹 |
| `A_ScriptDir` | ALTRun 所在的文件夹 |
| `%Temp%` `%AppData%` `%UserProfile%` `%OneDrive%` | 环境变量 |

`A_` 变量要写在开头, 例如 `A_Desktop\Projects`。只写程序名 (例如 `notepad`) 时会在 PATH 里查找。

## 文字片段
常用的文字 (签名、地址、模板...) 存成片段, 搜索名称、关键字或正文里的词, `Enter` 粘贴到呼出 ALTRun 之前的窗口。输入 `snip` 列出全部片段, `snip xxx` 只在片段里搜索。

| 字段 | 说明 |
|---|---|
| 名称 (Name) | 必填, 显示在搜索结果里, 也用来搜索。写成容易想到的几个词, 例如 `PT quotation` |
| 关键字 (Keyword) | 可选, 短一点, 例如 `pq`。搜索时输入完全相同的关键字排在最前面; 自动展开时输入 `;pq`。前缀加关键字最多 40 个字符 (默认前缀 `;` 时关键字最多 39 个), 不能有空格, 太长时保存前会提示 |
| 正文 (Text) | 必填, 粘贴的内容, 可以多行, 可以用下面的占位符。正文里的词也能搜到 (至少输入 3 个字符, 不想搜正文时在偏好设置里关掉) |
| 自动展开 (AutoExpand) | 在任何程序里输入 `;关键字` 时直接替换成正文 |

### 占位符
粘贴时替换:

| 占位符 | 替换为 |
|---|---|
| `{date}` | 今天的日期, 格式同 [一键加日期](Extensions#一键加日期) 的 DateFormat (默认 `dd.MM.yyyy`) |
| `{time}` | 当前时间 `HH:mm` |
| `{datetime}` | 日期 + 时间 |
| `{clipboard}` | 剪贴板里的文字 |
| `{date:yyyy-MM-dd}` `{time:HH:mm:ss}` | 自己指定格式 ([FormatTime](https://www.autohotkey.com/docs/v2/lib/FormatTime.htm) 的写法, 例如 `dddd` 星期几) |
| `{date+7}` `{date-1:dd.MM}` | 往后 / 往前几天的日期, 也可以带格式 |
| `{clipboard:1}` `{clipboard:2}` | 剪贴板历史里往前第 1、2 条文字 (和 Alfred 一样; `{clipboard}` 是现在的剪贴板; 跳过文件和图片) |
| `{uuid}` | 随机生成的 UUID, 每个都不一样 |
| `{cursor}` | 粘贴后光标停在这里 |

### 自动展开
在任何程序里输入 `;关键字` (例如 `;sig`), 输入的文字自动删掉并换成片段正文, 和 Alfred 的 Snippet 一样。前缀 `;` 是为了避免平时打字误触发, 可以在 `Features.Snippets.ExpandPrefix` 修改。单个片段不想自动展开时, 在编辑窗口里取消 "自动展开" (`"AutoExpand": 0`), 仍然可以搜索。在 ALTRun 自己的搜索窗口、Windows 的密码输入框, 以及 偏好设置 → 片段 → "不展开的窗口" 里的程序 (默认是远程桌面、KeePass、KeePassXC; `Features.Snippets.ExpandExclude`) 里不会展开。

### 粘贴方式
`Features.Snippets.PasteMode`:
- `Clipboard` (默认): 临时借用剪贴板 + `Ctrl+V`, 之后还原剪贴板
- `Type`: 逐字输入, 适合不接受粘贴的程序

## 剪贴板历史
按 `Ctrl+Alt+C` 或输入 `clip` 列出复制过的文字、文件和图片, `clip 关键词` 过滤, `Enter` 粘贴到前台窗口, `Ctrl+Del` 从历史中删除。

![剪贴板历史](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/clipboard.png)

| 复制的是 | 显示 | `Enter` | `F3` / `→` |
|---|---|---|---|
| 文字 | 文字的开头 | 粘贴文字 | 保存为片段 / 复制、粘贴、大字显示 |
| 文件 | 文件名 (多个文件用逗号隔开) | 粘贴文件 (在资源管理器里就是复制过去) | 添加到自定义命令 / 一个文件时有打开、在文件管理器中显示、复制路径等操作 |
| 图片 (截图、网页上的图片...) | "图片 1920 × 1080" 和缩略图 | 粘贴图片 | - / 用看图程序打开、在文件管理器中显示 |

- 图片存成 PNG 放在 `Clipboard\` 文件夹 (默认在本机的 `%LOCALAPPDATA%\ALTRun\Clipboard`, Data 在同步盘里时不用每次上传), 默认最多 50 张 (`MaxImages`), 只在 "退出后保留历史" 打开时记录; 不想记录图片可以在 偏好设置 → 剪贴板历史 关掉 "也记录图片"
- 输入 `clip 图片` 或 `clip image` 只看图片
- **置顶**: 选中一条按 `→`, 选 "置顶"。置顶的条目一直排在最前面 (图标右下角有一个图钉), 超过条数时不会被挤掉, "清空剪贴板历史" 时也保留; 常用的地址、账号等可以放在这里
- **粘贴为纯文本**: 系统命令 "粘贴为纯文本" 把剪贴板里的内容去掉格式 (字体、颜色、表格...) 粘贴, 剪贴板随后还原。在 [自定义热键](Extensions#自定义热键) 里设成 `Ctrl+Shift+V` 最方便
- **连续复制合并** (默认关闭, 偏好设置 → 剪贴板历史 → "快速按两次 Ctrl+C: 接到上一条后面"): 先复制一段, 再选中下一段快速按两次 `Ctrl+C`, 两段合成一条 (中间换行), 剪贴板里也是合并后的文字, 可以直接粘贴。和 Alfred 的 Merging 一样, 适合从几个地方摘抄。注意: DeepL 等翻译软件也用 "按两次 Ctrl+C" 呼出, 打开后两边会同时响应

隐私:
- 密码管理器 (KeePass、1Password、Bitwarden...) 复制的内容不会记录; 带 "不要加入剪贴板历史" 标记的内容也不会记录
- `Features.Clipboard.IgnoreApps` 可以加上其它不想记录的程序
- `Persist = 0` 时只保存在内存里, 退出即清空 (也不记录图片); 否则保存在 `Data\ClipboardHistory.json` (很长的条目和图片单独存在 `%LOCALAPPDATA%\ALTRun\Clipboard\`; `LocalFiles = 0` 时放在 `Data\Clipboard\`)
- 偏好设置 → 剪贴板历史 里可以清空历史
