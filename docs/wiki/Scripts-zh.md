[English](Scripts) · **中文**

# 脚本扩展

把自己的脚本放进 `Scripts\` 文件夹 (程序目录里; 偏好设置 → 脚本 → "打开脚本文件夹", 或者搜索 "打开脚本文件夹", 第一次打开时会建几个示例: 后台运行显示通知的、关键字后面带参数的、用文本编辑器看输出的、实时结果的), 就能在搜索窗口里按名称或关键字找到并运行, 和 Raycast 的 Script Commands 一样。只想看脚本时, 输入 **`script `** (列出全部) 或 **`script 名称`** (只输入 `script` 不带空格时, 显示一行提示 (名字以 script 开头的排在它前面), 在提示上按 `Enter` 补全成 `script `; 还没输完整时, 例如 `scr`, 提示在列表最后); 关键字在 偏好设置 → 脚本 → "搜索关键字" 修改。支持 `.ahk` (用 ALTRun 自带的 AutoHotkey 运行, 不用另外安装)、`.ps1`、`.bat` / `.cmd`、`.py` (需要装 Python)。偏好设置 → 脚本 列出找到的脚本, 双击用记事本编辑:

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
| `@altrun.mode filter` | 实时结果: 输入 "关键字 文字" 时一边输入一边运行脚本, 把它的输出显示成结果 (见下面的 [实时结果](#实时结果))。需要关键字; 只支持 `.ahk`、`.py`、`.ps1` |
| `@altrun.selection Text, PDF` | [内容类型](Usage-zh#选中内容的操作) (内置的或自己的, 用逗号分开): 选中这些类型的内容按 `Ctrl+Alt+\` 时, 操作面板里出现这个脚本 (排在自定义命令后面), 选中的内容作为第一个参数传给它 (多个文件: 每个一个参数) |

例子 (`Scripts\Hash.ps1`):
```powershell
# @altrun.title    文件哈希
# @altrun.keyword  hash
# @altrun.argument 文件路径
# @altrun.mode     output
Get-FileHash -Algorithm SHA256 $args[0] | Format-List
```
输入 `hash D:\Downloads\setup.exe` 后 `Enter`, SHA256 用文本编辑器显示。关键字不要和默认命令重复 (例如 `ping` 已经是自带的 Ping 命令)。

**选中内容用的脚本**: 写了 `@altrun.selection` 的脚本, 在任何程序里选中这种类型的内容按 `Ctrl+Alt+\`, 列表里就有它。选中的文字有换行或双引号时不传, 有脚本所用 shell 的特殊字符时也不传 (`.bat` / `.cmd` 和所有 `silent` / `output` 脚本经过 cmd, `.ps1` 经过 PowerShell), 这时列表里不出现这个脚本。偏好设置 → 脚本 的 "类型" 列写出每个脚本的类型。

**给脚本设热键**: 偏好设置 → 自定义热键 里 **选择...** 也列出脚本 (存成 `"Action": "script:<文件名>"`, 例如 `script:hash.ps1`)。写了 `@altrun.argument` 或 `@altrun.selection` 的脚本把选中的内容作为参数, 检查同上; 没有选中时, 有参数又有关键字的脚本打开搜索窗口并填好关键字。改了文件名热键就失效: 显示成红字 `(脚本不存在) script:<文件名>`, 重新选一次脚本就好。

`.ahk` 脚本里用 `FileAppend("文字", "*")` 输出。增删、修改脚本后马上生效, 不用重新载入; 选中脚本按 `F3` 用记事本编辑。后台运行的脚本最多等 2 分钟。

## 实时结果

写了 `@altrun.mode filter` 的脚本自己就是一个搜索来源, 类似 Alfred 的 Script Filter。输入它的关键字和空格: ALTRun 把后面输入的文字交给脚本运行, 把脚本的输出显示成结果, 结果和其他结果一样能打开、复制、`→` 打开操作面板、`F3` 编辑脚本。

![脚本的实时结果](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/script-filter.png)

```autohotkey
; @altrun.title    找项目
; @altrun.keyword  proj
; @altrun.mode     filter
query := A_Args[1]
out := ""
Loop Files, "D:\Projects\*" query "*", "D"
    out .= A_LoopFileName "`t" A_LoopFileFullPath "`t" A_LoopFileFullPath "`tfolder`n"
FileAppend(out, "*", "UTF-8-RAW")
```
输入 `proj alt`: 列出 `D:\Projects` 下名字含 "alt" 的文件夹; `Enter` 打开。

**脚本收到什么**: 关键字后面的文字作为第一个参数 (还没输入时是 `""`), 也在环境变量 `ALTRUN_QUERY` 里 (另外有 `ALTRUN_KEYWORD` 和 `ALTRUN_VERSION`)。`.ps1` 脚本用 `$env:ALTRUN_QUERY`。脚本是直接启动的, 不经过 cmd, 所以文字里的 `&`、`|`、`>`、引号都只是文字。

**脚本输出什么**, 两种都行:
- 每行一个结果, 用 Tab 分开: `标题⇥副标题⇥目标⇥类型⇥图标` (只有标题必须写; 不写目标就是标题), 或
- JSON, 字段名和 Alfred 的 Script Filter 一样:
```json
{"items": [
  {"title": "ALTRun", "subtitle": "D:\\Projects\\ALTRun", "arg": "D:\\Projects\\ALTRun", "type": "folder", "uid": "altrun"},
  {"title": "Wiki", "arg": "https://example.com/wiki", "icon": "C:\\Icons\\wiki.ico"}
]}
```

| 字段 | 意思 |
|---|---|
| `title` | 必须有 |
| `subtitle` | 标题下面的灰色一行 |
| `arg` | `Enter` 作用的对象; 不写就是标题 |
| `type` | `file`、`folder`、`url` (打开) 或 `text` (复制)。不写时看目标的样子: 网址是 `url`, `C:\...` 或 `\\server\...` 是 `file` (以 `\` 结尾是 `folder`), 其他是 `text` |
| `icon` | 本地的 `.ico` / `.png` / `.exe` / `.dll` 文件; 不用网址。不写时文件和文件夹用它自己的图标 |
| `uid` | 让 ALTRun 像对其他结果一样记住你输入什么时选了哪个; 不写就按脚本给的顺序 |
| `valid` | `false`: 这一行不能运行 (提示行); 有 `autocomplete` 时 `Enter` 把 "关键字 autocomplete" 填进搜索框 |
| `autocomplete` | `Tab` (或在不能运行的行上按 `Enter`) 填进关键字后面的文字 |
| `text.largetype` | `Ctrl+L` 大字显示的文字 |

输出用 UTF-8: `.ahk` 脚本写 `FileAppend(text, "*", "UTF-8-RAW")`; Python 和 PowerShell 由 ALTRun 运行时已经输出 UTF-8。

**什么时候运行**: 输入关键字和空格后马上运行一次, 之后每次停下输入一会儿 (0.15 秒) 再运行。新的一次开始、离开关键字、关掉搜索窗口时, 还在运行的那次会连同它启动的程序一起结束。运行期间保留上一次的结果; 0.3 秒后还什么都没有才显示 "正在运行…"。一次最多 10 秒。结果保留到搜索窗口关闭, 退格回到搜过的文字时马上显示。

**出错时**: 显示一行原因 (脚本写到错误输出的第一行、退出码、JSON 的问题或超时)。在这一行上 `Enter` 编辑脚本, `Ctrl+L` 显示全部错误输出。什么都没输出时显示 "没有结果"。

**速度**: `.ahk` 和 `.py` 脚本启动很快; Windows PowerShell 启动明显慢, 所以 `.ps1` 脚本会感觉慢一些。给实时结果脚本设的热键会打开搜索窗口, 填好关键字和选中的那一行文字。实时结果脚本不出现在选中内容的操作面板里。
