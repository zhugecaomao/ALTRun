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
| `@altrun.author` | Optional: your GitHub user name, for a script you contribute; the website shows it / 可选: 投稿的脚本写你的 GitHub 用户名, 网站上会显示 |

Checked by the tests (`ProjectTests.ScriptLibrary`, and syntax in the Tests workflow): only `.ahk`, `.ps1`, `.bat` and `.cmd`; a script that makes web requests names its hosts; no API key or token is written in a script (keys go in the user's environment variables); a `.ps1` with non-ASCII text is saved with a UTF-8 byte order mark (Windows PowerShell 5 reads files without one as ANSI); a `.bat` / `.cmd` has no byte order mark and non-ASCII text only on `rem` lines (cmd reads it in the console's code page).

由测试检查 (`ProjectTests.ScriptLibrary`, 语法在 Tests workflow 里检查): 只收 `.ahk`、`.ps1`、`.bat` 和 `.cmd`; 联网的脚本要写明网站; 脚本里不写 API key 或 token (key 放用户的环境变量); 含非 ASCII 文字的 `.ps1` 存成带 BOM 的 UTF-8 (Windows PowerShell 5 把没有 BOM 的文件当 ANSI 读); `.bat` / `.cmd` 不带 BOM, 非 ASCII 文字只能在 `rem` 行 (cmd 按控制台代码页读)。

## Contributing a script / 投稿脚本

Scripts from other people are welcome. Send a pull request that adds the script to `docs/scripts/` (one script per pull request; Chinese or English is fine); a script submitted only as an issue attachment isn't taken in as it is. Before it goes in, the script is reviewed line by line for these, besides the checks above:

- What it does matches the description, and nothing more; it doesn't download and run other code
- Nothing you select, type or capture is written to a log or left in a temporary file (delete temporary files on every way out, errors included)
- Keys are read from the user's environment variables only and aren't passed on a command line
- What it needs besides Windows (Python, a module, a program) is in `@altrun.requires` / `requires_zh`, and no path from your own PC is written in it
- It runs from a hotkey without surprises: `@altrun.selection` / `@altrun.argument` make a hotkey copy the selection first (Ctrl+C in the program in front), so leave them out when the script doesn't use a selection
- A new file name and keyword that don't clash with the scripts already here

Put `@altrun.author` in it to be named on the website; once merged you are also listed among the [contributors](https://github.com/zhugecaomao/ALTRun/graphs/contributors). Scripts in the library are released under the [GPL-3.0](../../LICENSE) license, like the rest of ALTRun. Later fixes may be made to a contributed script, and a script that stops working and can't be fixed may be removed.

欢迎别人写的脚本。请提交 Pull Request, 把脚本加到 `docs/scripts/` (一个 PR 一个脚本; 中文英文都可以); 仅作为 Issue 附件提交的脚本不直接收录。收录前除了上面的检查, 还会逐行审查这几点:

- 做的事和说明一致, 不做别的; 不下载别的代码来运行
- 选中的、输入的、截取的内容不写进日志, 也不留在临时文件里 (所有退出路径都删临时文件, 出错时也一样)
- 密钥只从用户环境变量读, 不放在命令行参数里
- Windows 以外需要的东西 (Python、模块、程序) 写进 `@altrun.requires` / `requires_zh`, 脚本里不写死自己电脑上的路径
- 用热键运行时不出意外: `@altrun.selection` / `@altrun.argument` 会让热键先复制选中内容 (对前台程序按 Ctrl+C), 不用选中内容的脚本就不要写
- 文件名和关键字是新的, 不和库里已有的脚本重复

写上 `@altrun.author` 就会在网站上显示你的名字; 合并后你也会出现在 [Contributors](https://github.com/zhugecaomao/ALTRun/graphs/contributors) 里。库里的脚本和 ALTRun 其他部分一样以 [GPL-3.0](../../LICENSE) 许可发布。收录的脚本以后可能会被修改; 失效又修不好的脚本可能会被移除。
