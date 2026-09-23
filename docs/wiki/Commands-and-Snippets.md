# 自定义命令与片段

## 自定义命令
把常用的文件、文件夹、程序或网址加进 ALTRun, 用名字或关键字打开。

![偏好设置里的自定义命令](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-commands.png)

### 添加
- 搜索结果里选中应用 / 文件 / 文件夹 / 网址, 按 `F3` (或 `→` → "添加到自定义命令...")
- 没有搜到结果时按 `F3`, 用输入的文字新建一条
- 资源管理器里右键 → 发送到 → ALTRun
- 偏好设置 → 自定义命令 → 添加

已有的命令: 搜索到后按 `F3` 修改, `Ctrl+Del` 删除; 或在偏好设置里双击编辑。

### 字段
| 字段 | 说明 |
|---|---|
| 名称 (Title) | 显示的名称, 按它搜索 |
| 类型 (Type) | `File` 文件 / 程序, `Folder` 文件夹, `Command` 程序 + 参数, `Url` 网址 |
| 目标 (Target) | 路径、程序或网址 |
| 参数 (Arguments) | 运行程序时附带的命令行参数, 例如 `/k ipconfig /all` |
| 关键字 (Keyword) | 可选。输入完全相同的关键字时排在最前面 |

搜索时除了名称和关键字, `File` / `Folder` 类型还会匹配目标的文件名 (不含扩展名) 或文件夹名: Target 是 `Q:\Projects\PT1931 - 24 NIR`, 名称是 `CKR, EA, JIB`, 输入 `nir` 或 `1931` 也能找到。同样的匹配程度, 名称匹配排在前面。

文件夹用 偏好设置 → 通用 → 文件管理器 打开 (默认资源管理器, 也可以设为 Total Commander 等)。

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
常用的文字 (签名、地址、模板...) 存成片段, 搜索名称或关键字, `Enter` 粘贴到呼出 ALTRun 之前的窗口。输入 `snip` 列出全部片段, `snip xxx` 只在片段里搜索。

### 占位符
粘贴时替换:

| 占位符 | 替换为 |
|---|---|
| `{date}` | 今天的日期, 格式同 [Ctrl+D 加日期](Extensions#ctrld-加日期) 的 DateFormat (默认 `dd.MM.yyyy`) |
| `{time}` | 当前时间 `HH:mm` |
| `{datetime}` | 日期 + 时间 |
| `{clipboard}` | 剪贴板里的文字 |
| `{cursor}` | 粘贴后光标停在这里 |

### 自动展开
在任何程序里输入 `;关键字` (例如 `;sig`), 输入的文字自动删掉并换成片段正文, 和 Alfred 的 Snippet 一样。前缀 `;` 是为了避免平时打字误触发, 可以在 `Features.Snippets.ExpandPrefix` 修改。单个片段不想自动展开时, 在编辑窗口里取消 "自动展开" (`"AutoExpand": 0`), 仍然可以搜索。在 ALTRun 自己的搜索窗口里不会展开。

### 粘贴方式
`Features.Snippets.PasteMode`:
- `Clipboard` (默认): 临时借用剪贴板 + `Ctrl+V`, 之后还原剪贴板
- `Type`: 逐字输入, 适合不接受粘贴的程序

## 剪贴板历史
按 `Ctrl+Alt+C` 或输入 `clip` 列出复制过的文字, `clip 关键词` 过滤, `Enter` 粘贴到前台窗口, `F3` 保存为片段, `Ctrl+Del` 从历史中删除。

![剪贴板历史](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/clipboard.png)

隐私:
- 密码管理器 (KeePass、1Password、Bitwarden...) 复制的内容不会记录; 带 "不要加入剪贴板历史" 标记的内容也不会记录
- `Features.Clipboard.IgnoreApps` 可以加上其它不想记录的程序
- `Persist = 0` 时只保存在内存里, 退出即清空; 否则保存在 `Data\ClipboardHistory.json`
- 偏好设置 → 剪贴板历史 里可以清空历史
