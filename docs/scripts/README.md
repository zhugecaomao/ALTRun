# Script Library / 脚本库

Ready-made [scripts](https://github.com/zhugecaomao/ALTRun/wiki/Scripts) for ALTRun. Download one into the `Scripts\` folder (Preferences → Scripts → Open Scripts Folder) and type its keyword. Each script is a plain text file: open it to see what it does before you use it.

现成的 ALTRun [脚本](https://github.com/zhugecaomao/ALTRun/wiki/Scripts-zh)。下载到 `Scripts\` 文件夹 (偏好设置 → 脚本 → 打开脚本文件夹), 输入它的关键字就能用。每个脚本都是普通文本文件, 用之前可以打开看看它做什么。

## Adding a script / 加入脚本

Besides the `@altrun.xxx` lines ALTRun reads (title, keyword, mode...), a library script has these in its first 30 lines (ALTRun ignores them; the website's list is made from them):

除了 ALTRun 读的 `@altrun.xxx` (标题、关键字、模式等), 库里的脚本在前 30 行还要写这些 (ALTRun 不读, 网站上的列表由它们生成):

| Line / 行 | |
|---|---|
| `@altrun.description` | One sentence in English / 一句英文说明 |
| `@altrun.description_zh` | The same in Chinese / 同样的中文说明 |
| `@altrun.network` | The hosts it connects to, separated by commas, or `none` / 会连接的网站 (逗号分隔), 不联网写 `none` |
| `@altrun.requires`, `@altrun.requires_zh` | Optional: what it needs (a program, an API key, an ALTRun version) / 可选: 需要什么 (程序、API key、ALTRun 版本) |

Checked by the tests (`ProjectTests.ScriptLibrary`, and syntax in the Tests workflow): only `.ahk`, `.ps1`, `.bat` and `.cmd`; a script that makes web requests names its hosts; no API key or token is written in a script (keys go in the user's environment variables); a `.ps1` with non-ASCII text is saved with a UTF-8 byte order mark (Windows PowerShell 5 reads files without one as ANSI); a `.bat` / `.cmd` has no byte order mark and non-ASCII text only on `rem` lines (cmd reads it in the console's code page).

由测试检查 (`ProjectTests.ScriptLibrary`, 语法在 Tests workflow 里检查): 只收 `.ahk`、`.ps1`、`.bat` 和 `.cmd`; 联网的脚本要写明网站; 脚本里不写 API key 或 token (key 放用户的环境变量); 含非 ASCII 文字的 `.ps1` 存成带 BOM 的 UTF-8 (Windows PowerShell 5 把没有 BOM 的文件当 ANSI 读); `.bat` / `.cmd` 不带 BOM, 非 ASCII 文字只能在 `rem` 行 (cmd 按控制台代码页读)。
