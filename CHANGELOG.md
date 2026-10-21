# 更新日志

本项目的主要变化记录在这里, 格式参照 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)。
版本号使用发布日期 (`YYYY.MM.DD`; 同一天再发布时加上 `.1`、`.2`...), 与 [GitHub Releases](https://github.com/zhugecaomao/ALTRun/releases) 的标签一致。

## [未发布]

## [2026.10.20.4]

测试版。

Test build.

### 新增
- 偏好设置 → 搜索窗口 → "Tab 键" (`General.TabKey`): 默认仍是自动补全; 选 "切换到下一个结果" 后, `Tab` / `Shift+Tab` 选下一个 / 上一个结果 (最后一个之后回到第一个, 和 uTools 一样), `Ctrl+Tab` 自动补全。使用提示和 `?` 速查表跟着设置显示。#238 的建议

## [2026.10.20.3]

测试版。

Test build.

### 改进
- 使用提示和 `?` 速查表更容易看懂: 最常用的排在前面 (空格 / `folder` / `file` / `doc` + 名称、`Ctrl+Enter`、`→`、Quick Switch、`F1` ~ `F4`...), 写法统一成 "关键字 + 名称" (例如 `folder + 名称`、`file / open / ' + 名称`、`doc / zip + 名称`), 每条说明写清做什么; 新增一条 Quick Switch 的提示 (`Ctrl+G / Ctrl+E`, 在打开 / 保存对话框里)

## [2026.10.20.2]

测试版。

Test build.

### 改进
- Wiki 改成英文为默认, 和 README 一样: 英文页面去掉 `en-` 前缀 (例如 `Usage`), 中文页面加上 `-zh` 后缀 (例如 `Usage-zh`), 首页 `Home` 是英文, `Home-zh` 是中文; 侧边栏英文在前。中文界面 (简体、繁体) 的帮助 (`F1`、`?` 速查表) 打开 `-zh` 页, 其它语言打开英文页; README、官网、Bug 报告模板里的链接都改到新名字
  - 原来指向中文页的旧链接 (例如 `wiki/Usage`) 现在打开英文页, 每页开头可以切换到中文
  - 原来的英文页 (例如 `wiki/en-Usage`) 暂时保留成一行的跳转页, 指向新名字; 2026.10.20 之后发布满三个正式版后自动删掉

## [2026.10.20.1]

测试版。

Test build.

### 修复
- ALTRun 在剪贴板变化后马上去读, 会占着剪贴板, 让截图工具等写完剪贴板马上又打开它的程序 OpenClipboard 失败。现在等约 0.15 秒再读 (期间又变了就先读上一次的), 复制图片时只在复制原始数据时打开剪贴板, 转换和存 PNG 在关掉之后做。AutoHotkey 论坛用户的反馈
- 文件搜索 (用 Everything 时) 只输入 Everything 的语法 (例如 `d:`、`folder:`) 时报错 "Parameter #2 of InStr is invalid"。AutoHotkey 论坛用户的反馈
- "系统信息" (msinfo32) 等本身是对话框 (`#32770`) 的程序窗口被当成打开 / 保存对话框, 触发对话框快速跳转和文件夹面板; 现在 `#32770` 窗口要有资源管理器的文件列表才算。AutoHotkey 论坛用户的反馈

## [2026.10.20]

### Summary

New apps and files show up in search a few seconds after they appear, and background re-indexing no longer competes with your typing; find files by pinyin initials even without Everything; every theme now comes in Comfortable / Normal / Compact density, plus 5 new themes; keywords can have alternatives (e.g. `g, ㅎ`); type `g ` to search the text on the clipboard; import and export snippets from Excel / CSV; open files for editing; the Wiki is reorganized with one page per feature.

### Upgrade
- One-click upgrade: tray icon → Check for Updates → Install Update
- Installed with Scoop: `scoop update altrun`
- Installed with winget: `winget upgrade zhugecaomao.ALTRun`
- The settings file format has not changed. New options are added at startup, and upgrades keep the way things were:
  - `Appearance.Density`: upgrades keep the previous size, "Comfortable"; new installations default to "Normal" (a little smaller)
  - `Features.LargeType` (Large Type, `Ctrl+L`): stays on after upgrading; off by default for new installations
  - Built-in web search engines whose URL is still the old default switch to `{query|clipboard}` (see "New" below); engines whose URL you changed and engines you added are left alone
- **If you use "Dark Compact", "Light Compact" or "Midnight Compact"**: these three themes are gone, replaced by a density setting that works with every theme. After upgrading they show as "the matching theme + Normal", the same size as before; open Preferences → Appearance, where this is already selected, and click OK to save
- **If you use "Dark Modern" or "Light Modern"**: they are now the same size as the other themes (they used to be a size smaller). To keep them that small: Preferences → Appearance → "Density" → "Compact"
- Preferences → Applications no longer has "Refresh every" (folder watching makes it unnecessary); `Features.Applications.RefreshMinutes` in the settings file is no longer read

### New
- **New apps show up in seconds** (#189): the Applications folders are watched (by default the desktop, public desktop, Start menu and public Start menu, including subfolders); when a shortcut is added, deleted or renamed, the index is rebuilt in the background after 5 quiet seconds (at most once a minute), with no notification. It uses Windows folder change notifications and costs almost nothing while idle; the cache is also rebuilt in the background when it is older than an hour (for apps installed while ALTRun was not running)
- **The built-in file index watches folders too** (#189, used when Everything is not running): when files or folders are added, deleted or renamed in the indexed folders, they are rescanned in the background after 5 quiet seconds (at most once every 5 minutes). Edits to file contents, browser download temp files, Office lock files, and anything deeper than the scan depth or excluded do not count. The refresh interval remains as a fallback
- **Pinyin initials for file names without Everything** (#180): `cs` finds "测试.bat", `zjsm` finds "证件扫描文件.zip"; the matching characters are highlighted, ranked a little below direct name matches. Initials are computed in the background in batches after the index loads (about 0.25 s for 10,000 Chinese file names), so startup is not slowed down
- **Chinese file names that Everything finds by pinyin are ranked by how well they match** (#180; Everything 1.5 beta with the `pinyin` option, or 1.4 with IbEverythingExt), instead of coming after direct matches
- **Density**: Preferences → Appearance → Size → "Density": Comfortable / Normal / Compact, for every theme including custom ones. Comfortable is the theme's own size (row height 54, 32 px icons), Normal matches the old Dark Compact (row height 48, 28 px icons), Compact is one step smaller (row height 42, 24 px icons)
- **5 new themes**: Lavender and Rose (light), Forest and Sunset (dark), High Contrast (white on black with a yellow selection, clearer for low vision or bright light). 22 built-in themes in total
- **Several keywords**: the keyword of a web search, custom command or snippet can list alternatives separated by commas, e.g. `g, ㅎ, google`, and any of them works, so you don't have to switch back to an English layout first. Every keyword of a snippet expands; the hint and completion for a keyword alone use the one you typed. Suggested by an AutoHotkey forum user
- **Type `g ` to search the clipboard**: when the default value in the URL is `clipboard` (`{query|clipboard}`), typing the keyword and a space with nothing after it searches the text on the clipboard; if you go on typing, what you typed is searched. The built-in engines now use it, e.g. `https://www.google.com/search?q={query|clipboard}`
  - The list shows "Search Google for clipboard: '...'"; when the clipboard has no text (copied files, empty) it says so and can't be opened
  - When text and an image are copied together (e.g. Excel cells) the text is searched; content copied from password managers is not read, just like in the clipboard history
- **Web search from the clipboard history**: select a text item and press `→` to search it with each engine
- **Import / export snippets**: Preferences → Snippets → "Import/Export". Paste cells copied from Excel, or read a CSV / TSV file; columns are Name | Keyword | Text | Auto-expand (optional), a header row is skipped, and snippets with the same keyword are updated. Export to CSV (UTF-8 with BOM), edit it in Excel and import it again
- **Edit files**: the action panel of a file has "Edit" when Windows has an editor registered for that file type (like "Edit" in the right-click menu, e.g. open an .ahk script in the editor instead of running it), with the editor's name on the right; `Ctrl+E` in the search window edits the selected file. Suggested by an AutoHotkey forum user
- **Large Type can be turned off**: Preferences → Features → "Large type (Ctrl+L)". When off, `Ctrl+L` does nothing and the action panel and `?` cheat sheet don't list it

### Improvements
- **Background re-indexing no longer competes with typing** (#189):
  - The application index is scanned in the background in small batches (at most about 15 ms each), so hotkeys no longer stall during a rebuild; each folder is listed only once, about 30% faster; scanning pauses while the search window is open and resumes when it closes
  - Writing the cache after a scan (70–120 ms for the application index, about 0.6–0.8 s for a 30,000-entry file index) waits until the keyboard and mouse have been idle for 2 seconds, and happens right away on exit if still pending
  - Microsoft Store apps (PowerShell) are scanned in the background at most once an hour; a manual rebuild still scans them
  - The application and file indexes share one folder watcher: a failed read triggers a background rebuild, broken watches are retried every 30 seconds and a rebuild runs once they recover; folders whose names contain a dot are recognized
- **Changes to what is indexed take effect right away**: after changing the application index (folders, file types, exclusions, subfolder depth) or the built-in file index (folders, depth, exclusions, max entries) in Preferences, Apply rescans in the background; the old index keeps working until the scan finishes
- **Rebuild results**: "Rebuild Index" and "Rebuild File Index" show how many apps / files and folders were found when done; Rebuild File Index runs in the background with the button showing "Rebuilding..."; the result also mentions reaching the entry limit, Everything running (searches then use Everything), or unapplied changes on the page
- New installations of the built-in file index also skip `.svn`, `.hg`, `.venv`, `venv`, `.tox`, `.cache`, `.idea` and `.vs` folders by default (repositories, virtual environments and tool caches easily hold tens of thousands of files, fill up "Max entries" and push real documents out), and refresh every 60 minutes instead of 30; existing settings are unchanged
- **Pinyin initials are faster and easier to discover** (#180):
  - Each character's initial is looked up once and remembered (apps, commands and bookmarks benefit too); when Everything returns many Chinese file names, sorting takes about a quarter of the time; with several words (e.g. `pt 2024`) pinyin is no longer computed for every result
  - The Simplified Chinese usage tips (status bar, `?` cheat sheet) mention that `cs` finds apps and files by pinyin initials
  - Preferences → Search Sources: under the Everything heading, how to make Everything search Chinese file names by pinyin (with a link), and whether Everything is running below the checkbox; the built-in index section says Chinese file names can be found by pinyin initials directly
  - Preferences → Applications: the description of pinyin matching says it only applies to apps (file names always support pinyin); the README feature list and the Wiki file search FAQ mention it too
- All built-in themes are the same size at the same density, so the window no longer changes size when you switch themes
- Preferences → Appearance: theme thumbnails are spaced evenly across the list, four per row with no big gap on the right; long theme names (such as "System (Light / Dark)") are less likely to be cut off
- Web search: typing only a keyword (e.g. `g`) always shows a completion hint, and `Enter` completes it to `g `; engines with a default value (e.g. `{query|8000}`) open with it only after the keyword and a space (they used to open on the keyword alone)
- Searching the selection or a clipboard history item with a search engine uses at most 200 characters (very long text used to produce URLs the sites could not open)
- Tables in Preferences (Custom Commands, Snippets, Web Search, Scripts, Hotkeys, Usage): the last column fills the remaining width, with no horizontal scroll bar and no gap on the right; room is left for the vertical scroll bar when there are more rows than fit
- Preferences → Web Search: "URL ({query} = search term)" in the list and editor is now just "URL"; how to write `{query}`, `{query|default}` and `{query|clipboard}` is explained under the URL field, and the keyword and URL descriptions have examples (e.g. "g hello" searches Google for hello); the default-value example is `weather` instead of `8000` (which was a port number)
- Preferences → Applications: "Subfolder depth" and "Rebuild Index" share a row (like on Search Sources); the descriptions for that row and for "Max entries / Refresh every" on both pages are clearer, with a recommended refresh interval for the built-in file index (60 minutes or more)
- Preferences → File Search: more detailed descriptions for the file and folder keywords; the file types box is 4 lines tall
- Preferences → General → "Selection hotkey": a more concrete description (search text on the web, change case or Simplified / Traditional Chinese, save as a snippet; copy a file's path, open a terminal there, and more) with a "See all actions" link
- Preferences → Date Stamp: the two cases are now "File names (before the extension)" and "End of the text" (formerly "Rename files" / "Comment box"), each with an example: an existing date in a file name becomes today's, at the end of the text the date is simply added
- Examples to type in Preferences descriptions are in quotes, e.g. "doc report" searches only documents, so they stand apart from the description
- The Help button and `F1` in Preferences open the Wiki page for that feature (several pages used to open "Extensions" or "Search and shortcuts"); the links in the `?` cheat sheet point to the matching pages too
- The Wiki is reorganized with one page per feature, matching the Preferences pages: new pages (English and Chinese) for Getting started, Privacy and network access, Data files, Web search, Window switching, Clipboard history, Pinned and recent items, System commands and terminal, Quick Switch, Dialog Folder Panel, Date Stamp, Browser bookmarks, Calculator, Scripts, Custom hotkeys and more; plus how the application index updates itself, which documentation belongs to each Preferences page, the modules in the development guide, translating and the release process; more Preferences screenshots
- When the search box has selected text (e.g. the restored text of "Keep the last search when the window opens"), `→` only clears the selection and moves the cursor to the end, like any text box, instead of opening the action panel: a restored calculation (`1250+500`) can be edited; press `→` again for the action panel
- The log file stays open instead of being opened and closed for every line: showing the window during the first minute after startup is no longer slow (on old PCs antivirus software checks the file each time it is closed); the time taken by each of the two input method switching steps and Everything replies that time out (how long it waited, whether other messages arrived meanwhile) are written to the debug log
- Development: release builds are pinned to AutoHotkey v2.0.30 and Ahk2Exe 1.1.37.02a2, benchmarks to AutoHotkey v2.0.30; unit tests still use the latest stable AutoHotkey to keep checking compatibility

### Fixes
- A broken shortcut on the desktop or in the Start menu made the application scan fail with "This local variable has not been assigned a value (linkTarget)", stopping halfway, after which the application index no longer updated by itself; broken shortcuts are now indexed (without a target) without an error. The recent folders next to dialogs hit the same error with a broken shortcut in Recent Items and now skip it
- With "Keep the last search when the window opens" on and a file search as the last search, showing the window could stall for a second (waiting for Everything). Results that need no waiting now show first, the window appears right away, and the full search runs afterwards and reselects the previous row
- Showing the search window could be delayed by icons being preloaded in the background (2.3 s once on an old PC); icons are no longer read until the window is shown
- "Refresh every" of the built-in file index had no effect while running: the cache age was only checked at startup. It is now checked every minute and rebuilt silently in the background when due (waiting while the search window is open); the minimum interval is 5 minutes
- Closing Preferences while it was still building a page (e.g. ALTRun busy, the tray menu clicked or Esc pressed) showed the error "This value of type "String" has no method named "Add""
- Dialog Folder Panel: the source column on the right (e.g. "Total Commander (other panel)") now fills to the edge, so the selection color covers the whole row without a blank strip and line on the right
- "Open With..." in the action panel had no icon (the icon it used does not exist in Windows); it now uses the generic program icon
- Preferences → Quick Switch: the "Works in" checkbox was touching the "Other dialogs" box below it
- Preferences → Features: the "Keyword" boxes on the Bookmarks and Windows rows were too close together (almost touching in Chinese and Japanese); the two columns of feature checkboxes above also get a little more row spacing in Chinese and Japanese, based on the checkbox height
- Closing or restarting a folder watcher now waits for the pending asynchronous read to finish or be cancelled before freeing its memory, so memory is not released while a read is still using it

---

### 中文说明

新装的软件、新建的文件几秒后就能搜到, 后台重建索引不再和打字抢时间; 没有 Everything 时文件名也能用拼音首字母找; 所有主题都有宽松 / 常规 / 紧凑三种显示密度, 另有 5 个新主题; 关键字可以写几个 (例如 `g, ㅎ`); 输入 `g ` 就搜索剪贴板里的文字; 文字片段可以从 Excel / CSV 批量导入导出; 文件可以直接 "编辑"; Wiki 按功能重新整理成一页一个功能。

#### 升级
- 一键升级: 托盘图标 → 检查更新 → 安装更新
- 用 Scoop 安装的: `scoop update altrun`
- 用 winget 安装的: `winget upgrade zhugecaomao.ALTRun`
- 设置文件的格式没有变化。新的选项启动时自动补上, 从旧版本升级的保持原来的样子:
  - `Appearance.Density` (显示密度): 升级的保持原来的大小 "宽松"; 新安装默认 "常规" (小一些)
  - `Features.LargeType` (大字显示 `Ctrl+L`): 升级的保持打开; 新安装默认关
  - 自带的网页搜索引擎, 网址还是原来默认值的, 自动换成 `{query|clipboard}` 的写法 (见下面 "新增"); 改过网址的和自己加的不动
- **用 "紧凑深色"、"紧凑浅色"、"紧凑午夜" 主题的**: 这三个主题去掉了, 改成所有主题都能用的显示密度。升级后自动按 "对应的主题 + 常规" 显示, 大小和以前一样; 打开 偏好设置 → 外观, 已经替你选好, 点 确定 保存即可
- **用 "现代深色"、"现代浅色" 主题的**: 现在和其它主题一样大 (以前小一圈)。喜欢原来那么小的: 偏好设置 → 外观 → "显示密度" 选 "紧凑"
- 偏好设置 → 应用搜索 去掉了 "刷新间隔" (监视文件夹之后不再需要), 设置文件里的 `Features.Applications.RefreshMinutes` 不再读取

#### 新增
- **新装的软件几秒后就能搜到** (#189): 监视应用搜索的文件夹 (默认: 桌面、公共桌面、开始菜单、公共开始菜单, 包括子文件夹), 快捷方式有新增、删除、改名时, 安静 5 秒后在后台重建索引 (两次至少隔 1 分钟), 不弹通知。用的是 Windows 的文件夹变化通知, 平时几乎不占资源; 另外缓存超过 1 小时也在后台重建 (覆盖 ALTRun 没运行时装的软件)
- **内置文件索引也监视文件夹** (#189, 没有 Everything 时用): 索引的文件夹里有文件 / 文件夹新增、删除、改名, 安静 5 秒后在后台重新扫描 (两次至少隔 5 分钟); 修改文件内容、浏览器下载中的临时文件、Office 的锁文件不算, 比扫描深度深的、排除的也不算。刷新间隔留作兜底
- **没有 Everything 时文件名也能用拼音首字母找** (#180): `cs` 找到 "测试.bat", `zjsm` 找到 "证件扫描文件.zip"; 同样高亮对应的字, 比名称直接匹配的排得靠后一点。首字母在载入索引后于后台分批算好 (1 万个中文文件名约 0.25 秒), 不拖慢启动
- **Everything 按拼音找到的中文文件名按匹配程度排序** (#180; Everything 1.5 beta 打开 `pinyin` 选项, 或 1.4 装了 IbEverythingExt): 不再排在直接匹配的文件后面
- **显示密度**: 偏好设置 → 外观 → 大小 → "显示密度": 宽松 / 常规 / 紧凑, 对所有主题 (包括自定义主题) 都有效。宽松是主题本来的大小 (行高 54, 图标 32 px), 常规和以前的紧凑深色一样 (行高 48, 图标 28 px), 紧凑再小一级 (行高 42, 图标 24 px)
- **5 个新主题**: 薰衣草、玫瑰 (浅色), 森林、日落 (深色), 高对比 (黑底白字、黄色选中, 视力不好或强光下更清楚)。内置主题共 22 个
- **关键字可以写几个**: 网页搜索、自定义命令和文字片段的关键字用逗号分隔, 例如 `g, ㅎ, google`, 输入哪个都行; 经常在英文和韩文 / 中文等键盘布局之间切换时不用先切回英文。片段的每个关键字都能自动展开; 只输入关键字时的提示和补全用输入的那个。AutoHotkey 论坛用户的建议
- **输入 `g ` 就搜索剪贴板里的文字**: 网址里的默认值写成 `clipboard` (`{query|clipboard}`) 时, 输入关键字加空格、不再输入就搜索剪贴板里的文字, 继续输入了就搜输入的; 自带的搜索引擎都改成这样写, 例如 `https://www.google.com/search?q={query|clipboard}`
  - 列表里显示 "用 Google 搜索剪贴板: '...'"; 剪贴板里没有文字 (复制的是文件、空的) 时提示并且不能打开
  - 文字和图片一起复制时 (例如 Excel 的单元格) 搜索文字; 密码管理器复制的内容不读, 和剪贴板历史一样
- **剪贴板历史里的文字可以直接网页搜索**: 选中一条按 `→`, 用每个搜索引擎搜索
- **文字片段批量导入 / 导出**: 偏好设置 → 文字片段 → "导入/导出"。可以直接粘贴 Excel 里复制的单元格, 或者读取 CSV / TSV 文件; 列的顺序 名称 | 关键字 | 正文 | 自动展开 (可选), 表头自动跳过, 关键字相同的片段会更新。导出为 CSV (UTF-8 带 BOM), 用 Excel 打开改完再导入
- **编辑文件**: 文件的操作面板里加 "编辑", 这类文件在 Windows 里登记了编辑程序时才出现 (和右键菜单的 "编辑" 一样, 例如 .ahk 脚本用编辑器打开而不是运行), 右边显示编辑程序的名称; 搜索窗口里按 `Ctrl+E` 直接编辑选中的文件。AutoHotkey 论坛用户的建议
- **大字显示可以关掉**: 偏好设置 → 功能 → 启用的功能 → "大字显示 (Ctrl+L)"。关掉时 `Ctrl+L` 不起作用, 操作面板和 `?` 速查表里也没有这一项

#### 改进
- **后台重建索引不再和打字抢时间** (#189):
  - 应用索引改为在后台分批扫描 (每次最多约 15 ms), 重建时不再卡住热键; 每个文件夹只列一次目录, 扫描也快了约三成; 搜索窗口打开时暂停, 关了再继续
  - 扫完后写缓存 (应用索引一次 70–120 ms, 3 万条的文件索引约 0.6–0.8 秒) 改到键盘鼠标空闲 2 秒后再写, 退出时还没写就马上写
  - 应用商店应用 (PowerShell) 在后台最多每小时扫一次, 手动重建时照常扫
  - 应用索引和文件索引共用一套文件夹监视: 读取通知失败时触发后台重建, 每 30 秒尝试恢复失效的监视, 恢复后按 "可能有变化" 重建一次; 名称带扩展名的文件夹也认得出
- **改了索引范围马上生效**: 偏好设置里改了应用索引 (文件夹、文件类型、排除、子文件夹深度) 或内置文件索引 (文件夹、深度、排除、最多条数), 点 "应用" 后自动在后台重新扫描, 扫完前旧索引照常可用
- **重建索引有结果提示**: "重建索引" 和 "重建文件索引" 做完后弹出找到多少个应用 / 文件和文件夹; 重建文件索引在后台进行, 按钮显示 "正在重建..."; 达到最多条数、Everything 正在运行 (这时搜索用的是 Everything)、页面上还有没应用的修改时, 结果里会说明
- 新安装的内置文件索引: 默认还跳过 `.svn`、`.hg`、`.venv`、`venv`、`.tox`、`.cache`、`.idea`、`.vs` 文件夹 (代码仓库、虚拟环境和开发工具的缓存动辄上万个文件, 会占满 "最多条数", 真正的文档反而搜不到), 刷新间隔 30 → 60 分钟; 已有的设置不变
- **拼音首字母更快、更好找** (#180):
  - 每个汉字查过一次就记住 (应用、命令、书签也受益); Everything 返回大量中文文件名时, 排序的时间约为原来的四分之一; 输入多个词时 (例如 `pt 2024`) 不再给每条结果算拼音
  - 简体中文界面的使用提示 (状态栏轮换、`?` 速查表) 加一条: `cs` 按拼音首字母找应用和文件
  - 偏好设置 → 搜索来源: Everything 一节的标题下面是怎样让 Everything 用拼音搜中文文件名 (带链接), Everything 现在有没有在运行写在复选框下面; 内置索引一节写明中文文件名可以直接用拼音首字母搜
  - 偏好设置 → 应用: "中文名称按拼音首字母匹配" 的说明写明只管应用, 文件名的拼音总是可用; README 的功能列表、Wiki 文件搜索的常见问题也写上了
- 所有内置主题在同一档显示密度下一样大, 切换主题时窗口不会忽大忽小
- 偏好设置 → 外观: 主题缩略图的间距按列表宽度均匀分配, 一行 4 个刚好铺满, 右边不再空一大块; 名字长的主题 (例如 "跟随系统 (浅色 / 深色)") 也不容易被截断
- 网页搜索: 只输入关键字 (例如 `g`) 时都是提示补全, `Enter` 补全成 `g `; 写了默认值的引擎 (例如 `{query|8000}`) 在关键字加空格后才用默认值打开 (以前只输入关键字就直接打开)
- 用搜索引擎搜索选中内容或剪贴板历史时, 搜索词最多 200 个字符 (很长的文字以前会生成搜索网站打不开的超长网址)
- 偏好设置里的表格 (自定义命令、文字片段、网页搜索、脚本、自定义热键、使用统计) 最后一列补满剩下的宽度: 不再有横向滚动条, 右边也不留空; 行数多到出现竖滚动条时先让出它的宽度
- 偏好设置 → 网页搜索: 列表和编辑框里的 "网址 ({query} = 搜索词)" 改成 "网址"; `{query}`、`{query|默认值}`、`{query|clipboard}` 的写法放到网址输入框下面的说明里, 关键字和网址的说明都加了例子 (例如 "g hello" 用 Google 搜索 hello); 默认值的例子从 `8000` (其实是端口号) 换成 `天气`
- 偏好设置 → 应用搜索: "子文件夹深度" 和 "重建索引" 放在一行 (和搜索来源页一样); 两页这一行和 "最多条数 / 刷新间隔" 的说明写得更清楚, 内置文件索引的刷新间隔写上推荐值 (60 分钟以上)
- 偏好设置 → 文件搜索: 文件关键字、文件夹关键字的说明更详细; 文件类型的输入框改成 4 行
- 偏好设置 → 通用 → "选中内容的操作": 说明写得更具体 (文字可以网页搜索、转换大小写 / 简繁、存为片段; 文件可以复制路径、在此处打开终端等), 后面加 "查看全部操作" 链接
- 偏好设置 → 一键加日期: 两种场景改名为 "文件名 (加在扩展名前)" 和 "文字末尾" (以前是 "重命名文件" / "备注框"), 说明各带一个例子: 文件名里已有的日期换成今天, 文字末尾直接加上
- 偏好设置的说明里, 要输入的例子都加上引号, 例如 "doc 报告" 只搜文档, 和说明文字分得更清楚
- 偏好设置的 "帮助" 按钮和 `F1` 打开对应功能的 Wiki 页面 (以前好几页都打开 "扩展功能" 或 "搜索与快捷键"); `?` 速查表里各项的链接也改到对应的页面
- Wiki 按功能重新整理成一页一个功能, 和偏好设置的页面对应: 新增 5 分钟上手、隐私与网络访问、数据文件、网页搜索、窗口切换、剪贴板历史、置顶和最近使用、系统命令和终端、对话框快速跳转、对话框文件夹面板、一键加日期、浏览器书签、计算器、脚本扩展、自定义热键等页面 (中英文); 补充应用索引怎样自动更新、设置文件参考里每个偏好设置页面对应的文档、开发指南的模块、翻译和发布流程; 偏好设置的截图更多
- 输入框里有选中的文字时 (例如 "保留上一次的搜索" 恢复的输入是全选的), 按 `→` 和普通输入框一样只取消选择、光标到末尾, 不再直接打开操作面板: 恢复的算式 (`1250+500`) 可以接着改; 再按一次 `→` 才打开操作面板
- 日志文件一直开着, 不再每写一条都打开、关闭一次: 刚启动的一分钟里呼出窗口不再变慢 (旧电脑上杀毒软件每次关闭文件都要检查); 切换输入法的两步各用了多久、Everything 没在限时内回复 (等了多久、期间有没有别的消息) 都会写进调试日志
- 开发: 正式构建固定使用 AutoHotkey v2.0.30 和 Ahk2Exe 1.1.37.02a2, 基准测试固定使用 AutoHotkey v2.0.30; 单元测试仍用最新稳定版, 继续检测兼容性

#### 修复
- 桌面、开始菜单里有损坏的快捷方式时, 扫描应用报错 "This local variable has not been assigned a value (linkTarget)", 这次扫描停在半路, 之后应用索引不再自动更新; 现在损坏的快捷方式照样收进索引, 不报错。对话框旁的最近文件夹遇到损坏的快捷方式时也会报同样的错, 现在跳过它
- 打开了 "呼出窗口时保留上一次的搜索" 时, 如果上次搜的是文件, 呼出窗口可能卡 1 秒 (要等 Everything 回复)。现在先显示不用等的结果, 窗口马上出来, 显示后再完整搜索并选中上次的那一行
- 呼出搜索窗口时被后台预读图标耽误 (一台旧电脑上有一次多等了 2.3 秒): 窗口显示完之前不读图标
- 内置文件索引的 "刷新间隔" 在运行中不生效: 以前只在启动时看缓存是否过期, 现在每分钟检查一次, 到时间就在后台静默重建 (搜索窗口打开时等下一次); 间隔最少 5 分钟
- 偏好设置建页面的中途窗口被关掉时 (例如 ALTRun 正忙、点了托盘菜单或按了 Esc), 弹出 "This value of type "String" has no method named "Add"" 的错误
- 对话框下面的文件夹面板: 最右边的来源列 (例如 "Total Commander (other panel)") 补满到面板右边, 选中行的颜色铺满整行, 右边不再留一条空白和竖线
- 操作面板里的 "打开方式" 没有图标 (用的图标在 Windows 里不存在), 换成通用的程序图标
- 偏好设置 → 对话框快速跳转: "生效的窗口" 的复选框和下面 "其它对话框" 的输入框贴在一起
- 偏好设置 → 功能: 书签和窗口两行的 "关键字" 输入框靠得太近 (中文、日文界面几乎贴在一起); 上面功能开关的两列复选框在中文、日文界面行距也按复选框的高度加一点
- 关闭或重启文件夹监视时, 等异步读取完成或取消完成后再释放内存, 避免读取还在进行时内存被提前释放

## [2026.10.14]

### Summary

Faster: Preferences opens instantly, snippets expand right away, icons appear together with the text, and checking for updates no longer freezes ALTRun; switching the input method with Caps Lock also works in administrator windows, and everything works when ALTRun itself runs as administrator; optional fade in / out for the window and your own placeholder text; reset to defaults and copy system info; fixed Ctrl+D sometimes renaming a file to "d" and result rows that occasionally stayed blank.

### Upgrade
- One-click upgrade: tray icon → Check for Updates → Install Update
- Installed with Scoop: `scoop update altrun`
- Installed with winget: `winget upgrade zhugecaomao.ALTRun`
- 2026.09.25 and earlier: exit ALTRun and extract `ALTRun_v2026.10.14.zip` over the old folder; for upgrading from 2.x see [the 2026.09.23 notes](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)
- **The settings file is upgraded to version 5**: on first start the old settings are backed up as `Data\ALTRun.v4.backup.json` and a "Settings upgraded" notice is shown once. Only things existing users won't find surprising are changed:
  - Snippets still on the old default paste method "Clipboard + Ctrl+V" switch to the new "Automatic" (see "Improvements" below); "Typing the text" is kept
  - Defaults added later are filled in for existing users too (only what is missing; your own entries stay, emptied lists are left alone): the file search keyword `file`; Date Stamp works in Total Commander's dialogs and rename box; Quick Switch doesn't pop up in 7-Zip's extract window and Total Commander's dialogs
  - Cleanup: the "F1 → About" hotkey left by early test builds (F1 is built in now) is removed, and rebar diameters are stored as numbers
  - Visible new defaults only apply to new installations; existing settings are unchanged: the theme follows the system (was Light)
- New options are added at startup: `Appearance.Fade` (default 0, no fading), `Appearance.Placeholder` (default empty, shows "ALTRun Search"), `General.CapsLockNotify` (default 1), `General.CapsLockElevated` (default 0); `General.SwitchToEnglishInput` gains the value 2 (English mode of the current input method), 0 / 1 are unchanged
- The icon cache is stored on this PC in `%LOCALAPPDATA%\ALTRun\IconCache` (not synced with Data); you can delete that folder too when removing ALTRun completely
- When upgrading by extracting the zip, exit the running ALTRun first (tray icon → Exit) and then overwrite the files

### New
- **Caps Lock input switching also works in administrator windows** (#158; Preferences → Hotkeys → "Also in windows running as administrator", off by default): turning it on creates a scheduled task (one UAC confirmation; the option shows the Windows shield), which ALTRun then uses at every start to run a small helper that only handles Caps Lock as administrator (no more UAC prompts); it exits together with ALTRun. ALTRun itself keeps normal rights, so the programs it opens, drag and drop and network drives are not affected
- **Everything works when ALTRun runs as administrator** (#158): Everything file search, Quick Switch reading Total Commander's current folder, and the tray icon after Explorer restarts all work; starting ALTRun again hands over to the running instance; Run as administrator from the right-click menu replaces a running normal instance; the tray tooltip and version show "(Administrator)". The Wiki FAQ explains how to start ALTRun as administrator at logon with a scheduled task (no UAC prompt)
- **Fade in / out for the search window** (#158; Preferences → Appearance → Animation, off by default): fades in over about 120 ms while the window still gets focus right away, so typing is not affected; when hiding, the window disappears immediately and a shadow copy fades out over about 100 ms
- **Custom placeholder text in the empty search box** (#158; Preferences → Search Window → Empty search box → "Placeholder text"): empty shows "ALTRun Search", a single space shows nothing
- **Switch to the English mode of the current input method when shown** (#154; Preferences → Search Window → "Switch to English when shown": Off / English (US) keyboard / English mode of the current input method): no need to add an English keyboard in Windows; your usual input method (e.g. WeChat or Sogou) stays, only its Chinese mode is turned off; with "switch back when hidden" on, the previous Chinese / English mode is restored
- **The Caps Lock on / off notice can be turned off** (#158; Preferences → Hotkeys)
- **Restore defaults** (Preferences → Advanced → Reset): every option goes back to its default, while custom commands, snippets, hotkeys, web searches and pinned items are kept; the old settings are backed up as `ALTRun.json.bak`
- **Copy System Info** and **Open Log** (Preferences → Advanced → Diagnostics): paste the version, Windows version, display scaling, language, input method, theme, Everything status and so on into a bug report (no file paths or search text)

### Improvements
- **Preferences opens faster**: only the page being shown is built, others when first clicked (372 ms → 93 ms on the test machine; old PCs used to take a second or two; saving is twice as fast too)
- **Snippets expand faster**: the new paste method "Automatic" (default) types single-line text of up to 100 characters directly, so it appears right away without touching the clipboard or waiting to restore it (100 ms → 28 ms); multi-line text, tabs, longer text or emoji still go through the clipboard. Preferences → Snippets can switch back to "Clipboard + Ctrl+V"
- **Icons no longer lag behind**: about 2 seconds after startup, icons of frequent items are preloaded in the background; icons are kept in a disk cache on this PC, so after a restart they appear together with the text. They are only read while the search window is hidden, so typing and results are not slowed down; changed program icons are picked up automatically
- **Checking for updates and downloading exchange rates no longer freeze ALTRun**: the whole program used to stop until the download finished (with a slow network or a proxy timing out, hotkeys and the search window could be unresponsive for ten seconds or more); downloads now run in the background
- Pressing the hotkey while the icon cache is being saved shows the window right away (it used to wait for a slow write to finish)
- Defaults for new installations: the theme follows the system (Dark in Windows dark mode), the file search keyword `file` is added, and Quick Switch and Date Stamp know the 7-Zip and Total Commander windows
- Preferences → Advanced is rearranged: the version has its own row, and "Write debug log" moves to the "Diagnostics" section together with Open Log and Copy System Info
- The record editor (custom hotkeys / commands / snippets / search engines): a little more space between the gray descriptions and the inputs and between rows; on Preferences → Search Window the input boxes are aligned on the left
- Debug log: background tasks over 50 ms (`Perf: slow ...`), the time of each step when saving the icon cache, and readings for Ctrl+D and input method switching are recorded, so occasional problems can be traced
- Development: a new benchmark (`Tests/Bench`, "Bench" on GitHub Actions) simulates everyday use with real keystrokes on Windows and compares two versions

### Fixes
- **Ctrl+D sometimes renamed a file to "d"** (after the PC stayed on overnight, woke from sleep or was locked): Ctrl+D now only takes over in the rename box and passes the key through everywhere else
- The keyboard hook is reinstalled after waking from sleep, unlocking, or connecting / disconnecting a monitor, so hotkeys, snippet expansion, double Ctrl and Caps Lock keep working on a PC that is never shut down
- **Result rows occasionally stayed blank while hotkeys and the tray did not respond**: background icon loading no longer interrupts drawing the result rows; the first Chinese (or Japanese, Korean) result after a fresh boot no longer stalls either (the fallback font is loaded in the background after startup)
- With Caps Lock set to "switch Chinese / English mode", some input methods such as WeChat only switched once (#154)
- The text in the search box looked too high (#158)
- The last search result row was cut off at the bottom
- After recording a hotkey in Preferences, the box below was fully selected (typing a character replaced the whole window list); the recording hint did not move with the window
- The same icon was occasionally read twice
- After changing the rebar diameters in Preferences → Calculator, they were saved as text in the settings file

---

### 中文说明

更快: 偏好设置一打开就是、片段马上展开、图标和文字一起出来, 检查更新时不再卡住; CapsLock 切换输入法在管理员窗口里也能用, 以管理员身份运行 ALTRun 时一切照常; 窗口可以淡入淡出, 空搜索框的提示文字可以自定义; 恢复默认设置、复制系统信息; 修好了 Ctrl+D 偶尔把文件名换成 "d" 和结果行偶尔空白卡住的问题。

#### 升级
- 一键升级: 托盘图标 → 检查更新 → 安装更新
- 用 Scoop 安装的: `scoop update altrun`
- 用 winget 安装的: `winget upgrade zhugecaomao.ALTRun`
- 2026.09.25 及更早的版本: 退出 ALTRun, 把 `ALTRun_v2026.10.14.zip` 解压覆盖到原来的文件夹; 从 2.x 升级见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)
- **设置文件升级到版本 5**: 第一次启动时原来的设置备份为 `Data\ALTRun.v4.backup.json`, 并提示一次 "设置已升级"。只改老用户不会觉得意外的地方:
  - 片段的粘贴方式还是旧的默认值 "剪贴板 + Ctrl+V" 的, 改成新的 "自动" (见下面 "改进"); 选了 "逐字输入" 的不变
  - 后来补进默认值的内容老用户也补上 (缺什么补什么, 自己加的保留, 清空了的列表不动): 文件搜索关键字 `file`; 一键加日期认 Total Commander 的对话框和改名输入框; Quick Switch 不在 7-Zip 解压窗口和 Total Commander 的对话框里弹出
  - 清理早期测试版留下的 "F1 → 关于" 热键 (F1 早已内置)、钢筋直径存成数字
  - 看得见的新默认值只对新安装生效, 已有的设置不变: 主题跟随系统 (以前是浅色)
- 新的选项启动时自动补上: `Appearance.Fade` (默认 0, 不淡入淡出)、`Appearance.Placeholder` (默认空, 显示 "ALTRun 搜索")、`General.CapsLockNotify` (默认 1)、`General.CapsLockElevated` (默认 0); `General.SwitchToEnglishInput` 多了一个值 2 (当前输入法的英文模式), 原来的 0 / 1 不变
- 图标缓存存在本机的 `%LOCALAPPDATA%\ALTRun\IconCache` (不跟着 Data 同步); 彻底删除 ALTRun 时这个文件夹也可以删掉
- 手动解压升级时, 请先退出正在运行的 ALTRun (托盘图标 → 退出) 再覆盖文件

#### 新增
- **CapsLock 切换输入法在管理员窗口里也生效** (#158; 偏好设置 → 自定义热键 → "在管理员窗口里也生效", 默认关): 打开时建一个计划任务 (确认一次 UAC, 选项前面有 Windows 的小盾牌), 之后 ALTRun 每次启动都用它以管理员身份运行一个只管 CapsLock 的小助手 (不再弹 UAC), ALTRun 退出时一起退出。ALTRun 本身还是普通权限, 打开的程序、拖放文件、网络驱动器都不受影响
- **以管理员身份运行 ALTRun 时一切照常** (#158): Everything 文件搜索、Quick Switch 读 Total Commander 的当前文件夹、资源管理器重启后的托盘图标都正常; 再启动 ALTRun 时交给正在运行的那个; 右键 → 以管理员身份运行时替换掉正在运行的普通权限的 ALTRun; 托盘提示和版本号后面显示 "(管理员)"。Wiki 的常见问题里写了用计划任务开机以管理员身份运行 (不弹 UAC) 的步骤
- **搜索窗口淡入淡出** (#158; 偏好设置 → 外观 → 动画, 默认关): 淡入约 120 ms, 窗口照样马上获得焦点, 不影响打字; 隐藏时窗口照常马上隐藏, 留一个影子淡出约 100 ms
- **空搜索框里的提示文字可以自定义** (#158; 偏好设置 → 搜索窗口 → 空搜索框 → "提示文字"): 留空是 "ALTRun 搜索", 输入一个空格什么都不显示
- **呼出时切换到当前输入法的英文模式** (#154; 偏好设置 → 搜索窗口 → "呼出时切换到英文": 不切换 / 英语 (美国) 键盘 / 当前输入法的英文模式): 不用在 Windows 里添加英语键盘, 还用原来的输入法 (例如微信、搜狗输入法), 只是关掉它的中文; 打开 "窗口隐藏后切回" 时恢复原来的中 / 英模式
- **大写锁定的开 / 关提示可以关掉** (#158; 偏好设置 → 自定义热键)
- **恢复默认设置** (偏好设置 → 高级 → 重置): 所有选项恢复成默认值, 自定义命令、文字片段、热键、网页搜索和置顶的项目保留; 原来的设置备份为 `ALTRun.json.bak`
- **复制系统信息** 和 **打开日志** (偏好设置 → 高级 → 诊断): 报告问题时贴上版本、Windows 版本、显示缩放、语言、输入法、主题、Everything 状态等 (不含文件路径和搜索的内容)

#### 改进
- **偏好设置打开得更快**: 只建要显示的那一页, 别的页第一次点到时再建 (在测试机上 372 ms → 93 ms, 旧电脑上以前要一两秒; 保存也快了一倍)
- **片段展开更快**: 新的粘贴方式 "自动" (默认): 100 个字符以内的单行文字直接输入, 马上出来, 不动剪贴板, 也不用停下来等着还原剪贴板 (100 ms → 28 ms); 多行、带 Tab、更长或有 emoji 的仍用剪贴板。偏好设置 → 片段里可以改回 "剪贴板 + Ctrl+V"
- **图标不再晚一步出现**: 启动约 2 秒后在后台预读常用项目的图标; 读过的图标存在本机的磁盘缓存里, 重启后图标和文字一起出来。只在搜索窗口隐藏时读, 不影响打字和出结果的速度; 程序更新了图标会自动换上新的
- **检查更新、下载汇率不再让 ALTRun 卡住**: 以前在下载完之前整个程序都停着 (网络慢、代理要等超时时, 热键和搜索窗口可能十几秒没反应), 现在在后台下载
- 保存图标缓存时按热键也马上呼出窗口 (写文件慢时以前要等它写完)
- 新安装时的默认设置: 主题跟随系统 (Windows 深色模式时用深色)、文件搜索关键字多了 `file`、Quick Switch 和一键加日期补上 7-Zip 和 Total Commander 的窗口
- 偏好设置 → 高级 重新整理: 版本号单独一行, "写入调试日志" 挪到 "诊断" 一节, 和打开日志、复制系统信息放在一起
- 编辑一条记录的窗口 (自定义热键 / 命令 / 片段 / 搜索引擎): 灰色说明和输入框隔开一点, 每行之间宽松一些; 偏好设置 → 搜索窗口 几个输入框左边对齐
- 调试日志: 一次超过 50 ms 的后台任务 (`Perf: slow ...`)、保存图标缓存各步的耗时、Ctrl+D 和输入法切换的读数都会记下, 偶尔出问题时看得出原因
- 开发: 新的性能测试 (`Tests/Bench`, GitHub Actions 上的 "Bench") 在 Windows 上用真实按键模拟日常操作, 比较两个版本

#### 修复
- **Ctrl+D 有时把文件名换成 "d"** (电脑一夜不关机、睡眠唤醒或锁屏之后): Ctrl+D 现在只在改名的编辑框里接管, 其它时候按键原样交给程序
- 睡眠唤醒、解锁、接上 / 拔掉显示器之后自动重新装一次键盘钩子, 热键、片段展开、双击 Ctrl、CapsLock 不会因为一直开着不关机而失灵
- **结果行偶尔一直是空白, 同时热键和托盘也没有反应**: 画结果行时不再让后台载入图标的任务插进来; 刚开机后第一次出现中文 (或日文、韩文) 结果时也不再卡住 (启动后先在后台载入后备字体)
- CapsLock 选 "切换中 / 英模式" 时, 在微信输入法等输入法上只能切换一次 (#154)
- 搜索框里的文字看起来偏上 (#158)
- 搜索结果的最后一行下面被切掉一截
- 偏好设置里录完热键后, 下面的编辑框被全选 (接着打一个字整个窗口列表就被换掉); 录热键时的提示不跟着窗口移动
- 同一个图标偶尔读两遍
- 偏好设置 → 计算器 改了钢筋直径后, 设置文件里存成了文字

## [2026.10.12]

按热键后窗口马上出现, 马上打字也不丢字; 搜索窗口底部新增状态栏和 ☰ 菜单; 和 Mac 一样用 CapsLock 切换输入法; 网页搜索有了各自的图标, 还能写默认值; 剪贴板历史可以完全只存在本机; 弹出菜单跟着主题的深浅; 偏好设置重新整理。

The window appears the moment you press the hotkey, and nothing you type right away gets lost; a status bar with a ☰ menu at the bottom of the search window; switch the input method with Caps Lock like on a Mac; web searches get their own icons and can have a default value; clipboard history can stay entirely on this PC; pop-up menus follow the theme's light / dark look; a tidier Preferences window.

### 升级
- **从 2026.10.01 ~ 2026.10.08 升级**: 呼出搜索窗口, 空搜索框下面有一条 "发现新版本: ALTRun 2026.10.12", 按 `Enter` 安装
- **从 2026.09.30.1 ~ 2026.09.30.3 升级**: 呼出搜索窗口, 空搜索框下面有一条 "更新 ALTRun 到 2026.10.12", 按 `Enter` 更新
- **从 2026.09.26 ~ 2026.09.30 升级**: 托盘图标 → 检查更新 (2026.09.26 ~ 2026.09.28 是弹窗里选 "立即更新")。用 Scoop 安装的用 `scoop update altrun`, 用 winget 安装的用 `winget upgrade zhugecaomao.ALTRun`
- **从 2026.09.25 及更早的版本升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.10.12.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`; 从 2.x 升级见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)
- 设置和数据都保留, 设置文件的格式没有变化; 新的选项启动时自动补上: `General.CapsLock` (默认空, 不用)、`Appearance.StatusBar` (默认 1, 显示状态栏)。`Features.Clipboard.LocalFiles` 多了一个值 2 (全部只存本机), 原来的 0 / 1 不变
- **状态栏默认打开**: 使用提示从空搜索框挪到了状态栏里。想和以前一样, 在 偏好设置 → 搜索窗口 关掉 "显示状态栏"
- 手动解压升级时, 请先退出正在运行的 ALTRun (托盘图标 → 退出) 再覆盖文件

### 新增
- **搜索窗口底部的状态栏** (和 Listary、Raycast 一样, 默认打开, 偏好设置 → 搜索窗口 → "显示状态栏"): 左边是 ☰ 菜单按钮和使用提示 (每次呼出换一条; 文件搜索时是 "正在搜索文件", 操作面板里是 "Esc 返回"); 右边是选中项按 `Enter` 做什么 (打开 / 复制 / 粘贴 / 运行 / 补全) 和打开操作面板的 `→`, 画成圆角键帽。底色比窗口稍深, 跟着主题变; 只在有结果行时显示, 空搜索框又没有置顶项时只显示搜索框
- **☰ 菜单**: 点状态栏最左边, 弹出偏好设置、快捷键和语法 (`?`)、重建索引、检查更新、重新载入、退出
- **和 Mac 一样用 CapsLock 切换输入法** (偏好设置 → 自定义热键 最上面, 默认不开): 按一下切换到下一个输入法 (和 `Win+Space` 一样), 或者切换中文输入法的中 / 英模式; 按住 0.3 秒才开 / 关大写锁定, 并短暂提示 "大写锁定: 开 / 关"; `Shift+CapsLock` 照常。用 Chrome 远程桌面等远程控制时也能用 (远程软件补发来对齐大写锁定的极短按键不算)
- **网页搜索的默认值**: 网址里写 `{query|默认值}`, 只输入关键字时用默认值, 直接按 `Enter` 打开, 例如 `http://localhost:{query|8000}`
- **网页搜索的图标**: 编辑框里可以选 `.ico`、`.png` 或 `.exe` 作为图标 (相对路径从 Data 文件夹算起); "下载网站图标" 按网址找到网站自己的清晰图标, 存进 `Data\Icons` 并自动填好 (只在点按钮时访问那个网站一次, 不经过第三方服务)。自带的 Google、Bing、Baidu、GitHub、Wikipedia、YouTube、淘宝、京东、Google 翻译都有了各自的图标
- **剪贴板历史可以完全只存在本机**: 偏好设置 → 剪贴板历史 → "保存位置": Data 文件夹 / 图片和长条目只存本机 (默认, 和以前一样) / 全部只存本机 (`ClipboardHistory.json` 也放到 `%LOCALAPPDATA%\ALTRun`, 不跟着同步盘或 U 盘走)。改了之后下次启动时自动把文件搬过去 (两个方向都会搬, 两边都有时不覆盖)
- 偏好设置 → 功能 → 功能选项: 书签和窗口切换的关键字 (`bm`、`w`), 以及普通搜索里是否也显示书签 / 已打开的窗口 (以前只能在 ALTRun.json 里改)

### 改进
- **按热键后窗口出现得更快**: 窗口隐藏后就在后台把下一次的空搜索框准备好, 呼出时直接显示、拿到焦点, 不再先搜索; 也不再有系统的淡入淡出动画, 一按热键马上出现、一按 `Esc` 马上消失
- **弹出菜单跟着主题的深浅**: 托盘图标的右键菜单、☰ 菜单、搜索结果的右键菜单、Quick Switch 的文件夹菜单, 深色主题用 Windows 的深色菜单, 浅色主题用浅色菜单; 主题选 "跟随系统" 时跟着 Windows 切换 (需要 Windows 10 1903 或更新的版本)
- 按住 `→` 移动光标时, 到了末尾不再因为自动重复打开操作面板; 要打开就再按一下 `→`
- 已经输入了文字时, 把光标移到最前面再按空格, 文字保留并改用文件搜索; 文件搜索模式下光标在最前面按 `Backspace` 回到普通搜索, 文字也保留
- 带参数的命令还没输入参数时 (自定义命令的 `{query}`、带关键字和参数的脚本、网页搜索), 说明一行写清楚接着输入什么, 例如 "在 'ping ' 后面输入内容, 再按 Enter"
- 没有打开货币换算时输入 `100 usd to eur`, 提示那一行按 `Enter` 直接打开 偏好设置 → 计算器
- 第一次打开脚本文件夹时建三个示例脚本, 各演示一种写法: `Today.ahk` (后台运行, 显示成通知)、`Port.bat` (`port 8080` 列出占用这个端口的程序)、`IP Addresses.ps1` (后台运行, 用记事本打开输出)
- 日志、脚本的输出、进程 / 服务列表用这种文件的默认程序打开 (例如 `.log` 设成 Notepad4), 没有关联时仍用记事本
- 偏好设置重新整理: 搜索窗口页 "显示使用提示" 放在 "显示状态栏" 下面, 都有说明; 剪贴板历史页左列变窄, 热键和关键字放在同一行; 功能页 "隐藏的命令" 4 行高, 书签和窗口切换各占一行; 成对的设置上下对齐; 脚本页的文件夹路径太长时中间用省略号
- 启动时快捷方式 (开机启动、发送到、开始菜单) 没有变化就不重写: 开机时资源管理器很忙, 每次重写曾经让启动多花了 11 秒
- 程序的标语改为 "轻量的 Windows 启动器" / "A lightweight launcher for Windows"
- 调试日志: 呼出窗口的耗时 (`Perf: show`)、读一个图标超过 200 ms 的耗时、CapsLock 切换前后的输入法都会记下
- 文档: Wiki 补上状态栏、☰ 菜单、CapsLock、网页搜索图标和默认值、剪贴板保存位置; 常见问题加了 "剪贴板历史会不会跟着 Data 文件夹走"; README 首页的动图不再在右下角画按键

### 修复
- 按热键呼出窗口后马上打字, 前几个字会丢失或打进原来的窗口: 现在从按下热键到搜索框拿到焦点之间打的字先记下来, 显示后补进搜索框; 期间按的 `Enter` 在搜完后执行第一项 (`Ctrl` / `Alt` + `Enter` 也一样), `Esc` 关掉窗口
- 打开了 "显示窗口时切换到英文输入法" 时, 搜索窗口里按 CapsLock 切不到中文输入法
- Windows 缩放不是 100% (例如 150%) 时, 打开 / 保存对话框下面的文件夹面板比对话框宽, 下面还多出一片空白
- 偶尔出现窗口有好几行高、结果行却是空白: 换完结果后检查列表行数、窗口高度和自绘有没有运行, 对不上时自动改正; 打开 "写入调试日志" 时记下当时的情况

## [2026.10.08]

文档改为中英双语 (英文 README 和 Wiki); 偏好设置的页面列表加了图标、换页不闪, 高级页重新整理; 再次启动 ALTRun 或用 "发送到" 时交给正在运行的处理, 不再关掉重开; 剪贴板的图片和长条目只存在本机; 搜索里可以直接打开某个设置页; PT Tools 更好用; 关掉显示器后卡住的搜索窗口能自己恢复。

Bilingual documentation (English README and wiki); icons in the Preferences page list, flicker-free page switching and a tidier Advanced page; starting ALTRun again or using "Send to" hands over to the running instance instead of restarting it; clipboard images and long items stay on this PC; open a settings page straight from search; easier-to-use PT Tools; the search window recovers by itself after hanging when a monitor is turned off.

### 升级
- **从 2026.10.01 ~ 2026.10.05 (包括测试版) 升级**: 呼出搜索窗口, 空搜索框下面有一条 "发现新版本: ALTRun 2026.10.08", 按 `Enter` 安装
- **从 2026.09.30.1 ~ 2026.09.30.3 升级**: 呼出搜索窗口, 空搜索框下面有一条 "更新 ALTRun 到 2026.10.08", 按 `Enter` 更新
- **从 2026.09.26 ~ 2026.09.30 升级**: 托盘图标 → 检查更新 (2026.09.26 ~ 2026.09.28 是弹窗里选 "立即更新")。用 Scoop 安装的用 `scoop update altrun`
- **从 2026.09.25 及更早的版本升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.10.08.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`; 从 2.x 升级见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)
- 设置和数据都保留, 设置文件的格式没有变化; 新的选项 `Features.Clipboard.LocalFiles` 启动时自动补上 (默认 1): 第一次运行时把 `Data\Clipboard` 文件夹 (剪贴板的图片和长条目) 搬到本机的 `%LOCALAPPDATA%\ALTRun\Clipboard`
- 手动解压升级时, 请先退出正在运行的 ALTRun (托盘图标 → 退出) 再覆盖文件

### 改进
- 搜索结果选中行的圆角和 偏好设置 → 外观 里主题缩略图的圆角改为抗锯齿 (GDI+), 边缘平滑, 不再有台阶; 半径为 0 的主题 (例如 Classic) 不变; 缩略图改用 24 位图像 (没有透明通道), 抗锯齿的圆角不会让缩略图背景变透明
- 发布包里的 README、LICENSE、CHANGELOG 和 `Resources\` 里的文件和仓库 / 源码 ZIP 逐字节相同 (以前在 Windows 上打包时换行符被转成了 CRLF, 内容一样但文件对比显示不同)
- 搜索里可以直接打开某个设置页: 每页一条 "ALTRun 偏好设置: 外观" 这样的命令, 输入页名 (例如 `外观`、`hotkeys`) 就能找到, 回车直接打开那一页; 主项 "ALTRun 偏好设置" 仍排在最前面
- PT Tools: 用鼠标点进输入框时全选, 直接输入新数字; 只读的结果框 (面积、根数、计算结果) 不再停留 `Tab` / `Enter` 焦点, 例如 Area Expression 之后直接到 Rebar Diameter; 面积单位写成 mm²; 标签和分组名称更清楚 (Rebar Area、Bars Required、Calculator、Number of Bars、Top / Bottom Mesh、Mesh + Rebar 加上单位); 输入框和结果框都右对齐 (包括钢筋直径下拉框); Safety Factor 复选框也不停留焦点; 各组整体上移, Calculator 和 BRC Area 底边对齐; 字体改为 Segoe UI 10 号; 输入框左右留出边距, 光标不再紧贴数字。SPF2M 窗口的输入框也是点进去全选
- 已经在运行时再启动 ALTRun, 交给正在运行的处理, 不再把它关掉重开: 双击 `ALTRun.exe` 弹出搜索窗口; 开机启动的快捷方式什么都不做; 正在运行的没有回应 (卡住) 时才替换它。"应用" 设置、一键更新之后的重新启动不变
- 偏好设置 → 高级: 链接显示名称 "主页 · GitHub · 更新说明 · 报告问题" (以前显示网址), 版本后面加上版权 "© 2013–2026 zhugecaomao"; "设置文件" 一行去掉 (就是数据文件夹里的 `ALTRun.json`), 数据文件夹的路径太长时中间用省略号, 不再折成两行, 鼠标停在上面显示完整路径
- 偏好设置左边的页面列表每页前面加了图标 (Windows 自带的图标字体: Windows 11 的 Segoe Fluent Icons、Windows 10 的 Segoe MDL2 Assets, 和 Windows 设置同一套线条图标, 不另带文件); 选中和悬停效果和资源管理器一样。按页码打开偏好设置 (例如 `-Preferences 3`) 时窗口已经开着, 也会切到那一页; 切换页面时只重画右边的页面区域 (只换掉上一页和这一页的控件), 页面列表不再闪一下; 连续点 (或右键点) 列表下面的空白处也不会变成没有选中的页; 列表右边不再有一条列分隔线; 选中的页只用高亮表示, 不再有虚线焦点框 (以前第一次打开偏好设置或用 `Tab` 切回列表时会出现)

- 文档中英双语: GitHub 首页的 README 改为英文 (中文版是 `README.zh-CN.md`, 发布包里两个都有); Wiki 每页都有英文版 (`en-` 开头的页面), 侧边栏分 English / 中文 两组, 每页开头可以切换语言; `CONTRIBUTING.md` (贡献指南) 也改为英文, 中文版是 `CONTRIBUTING.zh-CN.md`; `SECURITY.md` 本来就有中英文; 更新日志每个版本的概要后面加了一句英文
- 速查表 (`?`) 和偏好设置的帮助 (`F1`) 按界面语言打开 Wiki: 中文界面打开中文页, 其它语言打开英文页

### 修复
- 偏好设置 → 剪贴板历史 里 "也记录图片" 的说明还写着图片存在 `Data\Clipboard`, 改为 "和长条目放在一起" (默认在本机的 `%LOCALAPPDATA%\ALTRun\Clipboard`)
- 搜索时偶尔弹出错误 "Item has no value" (IconCache): 后台检查文件夹图标 / 加载图标的定时器正好在两行代码之间打断上一轮时, 会删除一个已经被处理掉的项; 现在已经处理掉的直接跳过
- 偏好设置开着时用 资源管理器 → 发送到 → ALTRun 添加命令, 偏好设置会被关掉, 没保存的修改也丢了: 现在偏好设置不关, 切到 "自定义命令" 页, 新命令加进列表 (1 个时打开编辑框), 按 确定 / 应用 保存
- 打开偏好设置时焦点在左边的页面列表, `↑` `↓` 切换页面 (以前焦点落在 "常规" 页的第一个选项, 按方向键会不小心改掉双击呼出的设置); `Tab` 的顺序改为 页面列表 → 页面里的选项 → 底部按钮
- 搜索窗口偶尔弹出错误 "Invalid index" (例如从搜索结果打开 SPF2M 时): 画结果行的过程中被同时进行的新搜索换掉了结果列表, 现在先取出这一行的结果再画
- PT Tools: 改第一组的 Span Width 时, BRC Area 的 Span Width 会跟着变, 但钢筋网面积 (Top / Bottom Mesh Area) 没有重算; 现在钢筋网面积和 "+ Rebar" 一起更新
- 关掉一台显示器后, 呼出搜索窗口卡在显示的过程中 (2 分多钟, 重新打开显示器才恢复), 之后热键和托盘图标都打不开: 现在上一次显示超过 3 秒还没完成时, 再按热键会直接重新显示; 调试日志记下卡了多久、卡在哪一步, 以及显示器数量的变化
- 剪贴板历史里的图片和很长的条目 (`Clipboard` 文件夹) 改为只存在这台电脑上 (`%LOCALAPPDATA%\ALTRun\Clipboard`), 不再放进 Data 文件夹: Data 在 OneDrive 等同步盘里时, 每复制一张图片都要上传。`ClipboardHistory.json` 仍在 Data 里照常同步; 第一次运行时自动把原来的文件夹搬过去; 想和以前一样都放在 Data 里 (例如放在 U 盘随身带), 把 `Features.Clipboard.LocalFiles` 设为 0

## [2026.10.05]

选主题改为看缩略图, 新增 4 个主题和繁體中文界面; 新增切换窗口、浏览文件夹、标记多个文件、带参数的自定义命令、Windows 设置页面、脚本扩展、置顶和最近使用; 数据文件夹可以放到同步盘。

Pick themes from thumbnails, 4 new themes and a Traditional Chinese interface; new window switcher, folder browsing, marking several files, custom commands with an argument, Windows Settings pages, scripts, and pinned / recent items; the data folder can live in a synced folder.

### 升级
- **从 2026.10.01 ~ 2026.10.04 升级**: 呼出搜索窗口, 空搜索框下面有一条 "发现新版本: ALTRun 2026.10.05", 按 `Enter` 安装
- **从 2026.09.30.1 ~ 2026.09.30.3 升级**: 呼出搜索窗口, 空搜索框下面有一条 "更新 ALTRun 到 2026.10.05", 按 `Enter` 更新
- **从 2026.09.26 ~ 2026.09.30 升级**: 托盘图标 → 检查更新 (2026.09.26 ~ 2026.09.28 是弹窗里选 "立即更新")。用 Scoop 安装的用 `scoop update altrun`
- **从 2026.09.25 及更早的版本升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.10.05.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`; 从 2.x 升级见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)
- 设置和数据都保留, 设置文件的格式没有变化: 新的选项 (`Features.Windows`、`Features.Recent`、`Features.Scripts`、`Features.System.SettingsPages`、`Features.Snippets.ExpandExclude`、`Features.Calculator` 的结构计算参数) 启动时自动补上, 默认打开; 界面语言 `"zh"` 照常可用, 保存时写成 `"zh-CN"`
- 一键更新会带上新的 `Resources\Lang\` 语言文件; 手动解压升级时请把整个压缩包解压覆盖, 否则中文和日文界面会显示英文

### 新增
- 偏好设置 → 外观 的主题改为看缩略图选择 (类似 Alfred / Listary): 每个主题按自己的颜色和圆角画出一个简化的迷你搜索窗口 (背景、选中行、文字和高亮色, 结果行带彩色图标), 自定义主题也有; 列表可以滚动, 也可以用方向键选择
- 4 个新的内置主题: 现代深色 (`DarkModern`) 和现代浅色 (`LightModern`, 都是 VS Code 的默认配色)、`Monokai`、`One Dark`
- 繁體中文界面: 偏好设置 → 常规 → 界面语言 选 "繁體中文" (台湾用语, 字体用微软正黑体); 系统是繁体中文 (台湾、香港、澳门) 时 "自动" 也会选它。翻译在 `Resources\Lang\zh-TW.json`
- **带参数的自定义命令**: 目标或参数里写 `{query}`, 输入 "关键字 文字" 时换成后面的文字 (网址里自动编码), 例如 `jira ABC-123` 打开对应的页面、`ping 10.0.0.1` 运行 ping。新安装的默认命令里有 3 个例子 (`ping 10.0.0.1`、`maps 地名`、`docs 文件夹名`); 已有的设置不变, 可以照着 [说明](https://github.com/zhugecaomao/ALTRun/wiki/Commands-and-Snippets-zh) 自己加
- **搜索 Windows 设置的页面**: 输入 "蓝牙"、"显示"、"默认应用"、"Windows 更新" 等直接打开对应的设置页 (40 多个, 中英文名称都能搜); 偏好设置 → 功能 里可以关掉
- **切换窗口**: 输入窗口标题或程序名, `Enter` 切换到已打开的窗口 (最小化的先还原), `→` 可以关闭窗口; `w 关键词` 只搜窗口, `w ` 列出全部 (和 PowerToys 的 Window Walker 一样)
- **置顶和最近使用**: 呼出搜索窗口还没输入时, 列出置顶的项目 (操作面板里 "置顶到空搜索框") 和最近打开的项目 (默认不列, 在 偏好设置 → 搜索窗口 设个数, 例如 5); `Ctrl+Del` 取消置顶 / 从最近使用里去掉
- **剪贴板置顶**: 剪贴板历史里的条目可以置顶 (操作面板里 "置顶"), 一直排在最前面, 不会被新的挤掉, 清空历史时也保留; 置顶的条目图标右下角有一个图钉 (空搜索框里置顶的项目也是)
- **粘贴为纯文本**: 新的系统命令, 把剪贴板内容去掉格式粘贴; 可以在自定义热键里设成 `Ctrl+Shift+V`
- **浏览文件夹**: 输入 `D:\Projects\`、`~\`、`\\server\share\` 这样的路径, 列出文件夹里的内容并按最后一段过滤; `Tab` 进入文件夹 (任何结果里的文件夹都可以), `Backspace` 回到上一级。只高亮文件名里和最后一段匹配的字 (路径里的词不算); 文件夹不存在 (例如没有 D 盘) 时显示 "找不到这个文件夹"; 中文输入法打出的 `D：、` 也可以。偏好设置 → 文件搜索 和 `?` 速查表里有说明
- **标记多个文件**: `Insert` 标记文件或文件夹 (和 Total Commander 一样), `→` 对所有标记的一起操作: 全部打开、复制 / 剪切、复制路径、复制 / 移动到 TC 当前文件夹、添加到自定义命令、移到回收站
- **计算器: 进制和日期**: `255 in hex`、`0xff to dec`、`0b1010` 进制换算; `today + 30 days`、`2026-12-25 - 2w`、`今天 + 3 个月` 日期加减, `2026-12-25 - today` 相差几天
- **脚本扩展**: 把 `.ahk` / `.ps1` / `.bat` / `.py` 脚本放进 `Scripts\` 文件夹就能搜到并运行 (和 Raycast 的 Script Commands 一样); 开头注释里可以写名称、关键字、参数和运行方式 (正常 / 后台显示最后一行 / 后台用记事本显示输出)。偏好设置 → 脚本 里有说明、找到的脚本列表 (双击编辑) 和 "打开脚本文件夹" 按钮 (第一次会建一个示例); 也可以搜索 "打开脚本文件夹"

### 改进
- 界面语言改为语言文件: 英文编译在程序里, 简体中文、繁體中文、日本語是 `Resources\Lang\` 里的 `zh-CN.json`、`zh-TW.json`、`ja.json`。加一种语言只要放一个文件 (写上 `_name` 名称和 `_font` 字体), 偏好设置的语言列表会自动列出; 没翻译的条目显示英文 (繁体先用简体)。设置里以前的 `"zh"` 照常可用, 保存时写成 `"zh-CN"`
- **数据文件夹可以换位置**: 偏好设置 → 高级 → 数据 → 更改位置, 选同步盘 (OneDrive 等) 里的文件夹, 几台电脑共用一份设置; 现在的设置和数据会复制过去, 原来的文件夹保持不变。位置记在默认位置的 `Data\ALTRun.json` 里 (`"DataLocation"`, 可以写 `%OneDrive%\ALTRun` 这样的环境变量), 程序目录里不多出文件
- 程序放在不能写入的文件夹 (例如 `C:\Program Files`) 时, 设置和数据改为保存在 `%APPDATA%\ALTRun\`, 不再保存失败
- 片段自动展开: 可以设置不展开的窗口 (偏好设置 → 片段, 默认是远程桌面、KeePass、KeePassXC), Windows 的密码输入框里也不展开
- 一键更新改用 Windows 自带的 `tar.exe` 解压, 更快, 也不受 PowerShell 策略限制 (没有时仍用 PowerShell)
- 发布流程支持 SignPath 代码签名 (SignPath Foundation 审核通过后自动给 ALTRun.exe 签名)

### 修复
- 一键加日期: 在设置里换了日期格式 (例如 `yyyy-MM-dd`) 后, 再按 `Ctrl+D` 会更新名字末尾的日期, 不再多加一个
- 文件管理器设为 Directory Opus、XYplorer 等时, "打开所在位置" 打开文件所在的文件夹 (以前加了只有 Total Commander 认识的 `/P` 参数)
- 设置和数据文件按 JSON 标准保存: 不带 BOM, 文字里的控制字符 (例如从终端复制的内容) 写成转义形式, 别的工具也能正常读取; 应用索引、文件索引、汇率和更新状态也改为先写临时文件再替换, 写到一半中断不会损坏
- 官网分享链接时显示预览图
- ALTRun 自己的功能换成同一套风格的图标 (`Resources\Icons\`, Windows 10 和 11 上看起来一样): 剪贴板历史的文字条目、剪贴板文字工具和 "粘贴为纯文本" 用剪贴板图标 (以前是黑白的文件图标); 片段和 "存为片段" 用片段图标; 计算器结果、PT Tools 和 SPF2M 用计算器图标; `?` 速查表用帮助图标; "置顶 / 取消置顶" 操作用和结果上一样的图钉; "关于" 和 "检查更新" 用 ALTRun 的程序图标
- 计算器的结构计算可以设置参数 (以前写死在程序里): 保护层 (默认 40 mm)、最大间距 (默认 300 mm)、钢筋直径 (默认 13 / 16 / 20 / 25 / 32) 和前缀 (默认 H)。新的 偏好设置 → 计算器 页面, 货币换算的开关也移到这里; 一根钢筋的面积改为按 π d² / 4 计算
- 偏好设置 → 搜索窗口: "空搜索框" 单独一节 (使用提示、显示几个最近使用的项目), 搜索历史另起一节; 功能页的 "置顶和最近使用" 英文标签显示正确 (以前 `&` 显示成了下划线)
- 操作后的提示 (例如 "已置顶"、"已复制") 显示在搜索窗口所在的屏幕上; 没有打开搜索窗口时的提示 (片段展开等) 按 "窗口显示在" 的设置选屏幕。以前搜索窗口一隐藏, 提示就跑到鼠标所在的屏幕
- 打字很快时偶尔出现窗口有好几行高、结果却是空白的情况 (前一个字的搜索比后一个字慢, 两次搜索交叉): 过时的搜索结果不再显示, 更新列表时不会被打断, 之后再补画一次
- 启动后呼出搜索窗口时有时卡住 (托盘图标在, 但热键和点击都没反应, CPU 一直占用一个核): 窗口命令之后不再默认等 100 ms (等待期间嵌套处理消息时遇到不断的重画消息会一直等下去); 显示窗口的过程中再按热键不会重入; 写设置等数据文件时被 OneDrive 等同步盘或杀毒软件暂时占用会重试; 启动时的错误提示总在最前面
- 调试日志: 启动完成后每条马上写入 (以前攒够 60 条才写, 程序卡住被强制结束时最后的记录会丢失); 每条一行, 行尾统一 CRLF; 呼出搜索窗口和空搜索框里较慢的部分也记录耗时, 启动后 1 分钟内记下呼出窗口的每一步和加载的图标
- 偏好设置里的 "打开数据文件夹" 用设置的文件管理器 (例如 Total Commander) 打开, 不再总是用资源管理器

## [2026.10.04]

窗口隐藏后自动切回原来的输入法; 偏好设置的 "关于" 加上官网链接。

The previous input method comes back automatically when the window hides; the "About" section of Preferences links to the website.

### 升级
- **从 2026.10.01 ~ 2026.10.03 升级**: 呼出搜索窗口, 空搜索框下面有一条 "发现新版本: ALTRun 2026.10.04", 按 `Enter` 安装
- **从 2026.09.30.1 ~ 2026.09.30.3 升级**: 呼出搜索窗口, 空搜索框下面有一条 "更新 ALTRun 到 2026.10.04", 按 `Enter` 更新
- **从 2026.09.26 ~ 2026.09.30 升级**: 托盘图标 → 检查更新 (2026.09.26 ~ 2026.09.28 是弹窗里选 "立即更新")。用 Scoop 安装的用 `scoop update altrun`
- **从 2026.09.25 及更早的版本升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.10.04.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`; 从 2.x 升级见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)
- 设置和数据都保留, 设置文件的格式没有变化 (新的 `General.RestoreInput` 启动时自动补上, 默认打开)

### 改进
- "显示窗口时切换到英文输入法" 下面加了子选项 "窗口隐藏后切回原来的输入法" (默认打开): 回到别的程序时还是呼出前的输入法 (例如中文输入法), 不用再手动切换。只在所有程序共用一个输入法时 (Windows 默认) 才切回; 打开了 "允许我为每个应用窗口使用不同的输入法" 时别的程序本来就不受影响 (`General.RestoreInput`)
- 偏好设置 → 高级 的 "关于" 里加上官网链接 (https://zhugecaomao.github.io/ALTRun/); Scoop 和 winget 清单的主页也改为官网

## [2026.10.03]

内置命令可以从搜索结果里删除; 新增紧凑午夜主题; 自定义命令等列表加了筛选框; 匹配高亮的颜色跟着主题走; 修复呼出窗口后第一个字延迟、输入时字距跳动和结果区闪一下。

Built-in commands can be removed from the search results; new Midnight Compact theme; filter boxes for custom commands and other lists; the match highlight color follows the theme; fixed the delay on the first character after opening the window, jumping letter spacing while typing, and a flash in the results.

### 升级
- **从 2026.10.01 ~ 2026.10.02 升级**: 呼出搜索窗口, 空搜索框下面有一条 "发现新版本: ALTRun 2026.10.03", 按 `Enter` 安装
- **从 2026.09.30.1 ~ 2026.09.30.3 升级**: 呼出搜索窗口, 空搜索框下面有一条 "更新 ALTRun 到 2026.10.03", 按 `Enter` 更新
- **从 2026.09.26 ~ 2026.09.30 升级**: 托盘图标 → 检查更新 (2026.09.26 ~ 2026.09.28 是弹窗里选 "立即更新")。用 Scoop 安装的用 `scoop update altrun`
- **从 2026.09.25 及更早的版本升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.10.03.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`; 从 2.x 升级见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)
- 设置和数据都保留, 设置文件的格式没有变化 (新的 `Features.System.Hidden` 启动时自动补上)

### 新增
- 内置命令 (系统命令、Windows 工具等) 也能在搜索结果里按 `Ctrl+Del` 删除 (右键 → 删除 也可以), 只是不在搜索结果里显示, 热键里照样能用; 删除的命令列在 偏好设置 → 功能 → "已删除的命令", 删掉一行即可恢复
- 新的内置主题 "紧凑午夜" (`MidnightCompact`): 和午夜一样的配色, 字号、行高小一些, 图标 24 px
- 偏好设置里自定义命令、文字片段、网页搜索、自定义热键的列表右下角加了筛选框: 输入名称、路径、关键字里的词, 只显示匹配的项目 (几个词用空格分开, 都要出现)

### 改进
- 匹配高亮更醒目, 颜色跟着主题走: Dark / Dark Compact、Midnight 用更亮更饱和的蓝色, 选中行用青蓝色; Tokyo Night 用青色, Catppuccin 用紫色, Ocean 用霜蓝色 (选中行用黄色), Gruvbox 用橙色; 其它主题的高亮色也稍微加深、加饱和

### 修复
- 内置的 "设备和打印机" 没有图标: 它在控制面板下面, 以前的路径在桌面这一层找不到它。只写 `::{CLSID}` 的自定义命令也一样, 找不到时再到控制面板下面找
- 打开了 "打开窗口时保留上一次的搜索" 时, 呼出窗口偶尔结果区会白一下或闪一下: 窗口隐藏期间的重画请求会被丢掉, 以前是显示出来、拿到焦点、切换输入法之后才重画结果。现在先让窗口在屏幕上不可见, 画好整个窗口再露出来
- 输入时高亮的字和后面的字之间的间隔会跟着变 (例如 `We` → `Wei` 时 "Western" 的字距跳动): 以前每段文字分开量宽度再拼起来, 而且有高亮和没有高亮用两种画法, 都会差 1 像素左右; 现在每个字的位置都按整个标题一次量出来, 有没有高亮都用同一种画法
- 打开了 "打开窗口时保留上一次的搜索" 时, 呼出窗口后马上输入, 第一个字会延迟或丢失: 以前先把上次的搜索重新搜一遍再显示窗口, 这几十毫秒里按的键还在原来的程序里。现在输入框和结果都还是隐藏时的样子, 直接原样显示、拿到焦点并全选文字, 不再重新搜索; 输入新的文字时才搜索 (实测从按热键到能输入, 由约 30 ~ 80 ms 缩短到约 15 ms)

## [2026.10.02]

搜索结果高亮匹配的字; 新增 6 个内置主题; 紧凑主题和高 DPI 下的图标不再有锯齿。

Matched characters are highlighted in the results; 6 new built-in themes; no more jagged icons in compact themes and at high DPI.

### 升级
- **从 2026.10.01 升级**: 呼出搜索窗口, 空搜索框下面有一条 "发现新版本: ALTRun 2026.10.02", 按 `Enter` 安装
- **从 2026.09.30.1 ~ 2026.09.30.3 升级**: 呼出搜索窗口, 空搜索框下面有一条 "更新 ALTRun 到 2026.10.02", 按 `Enter` 更新
- **从 2026.09.26 ~ 2026.09.30 升级**: 托盘图标 → 检查更新 (2026.09.26 ~ 2026.09.28 是弹窗里选 "立即更新")。用 Scoop 安装的用 `scoop update altrun`
- **从 2026.09.25 及更早的版本升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.10.02.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`; 从 2.x 升级见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)
- 设置和数据都保留, 设置文件的格式没有变化。自己做的主题没写 `Highlight` / `SelectedHighlight` 时从 `Base` 主题继承

### 新增
- 6 个内置主题: Tokyo Night、Dracula、Catppuccin Mocha、Gruvbox 深色、Solarized 浅色 (常见编辑器配色), 以及和紧凑深色对应的紧凑浅色 (`LightCompact`)
- 搜索结果的标题里高亮和输入匹配的字 (类似 Listary): 连续出现的部分、单词首字母 (`vsc` → **V**isual **S**tudio **C**ode)、拼音首字母 (`jsb` → **记事本**) 都会标出。颜色由主题的 `Highlight` / `SelectedHighlight` 决定, 每个内置主题都配好了; 靠关键字、路径、片段正文匹配到的结果标题不变

### 改进
- 图标不是 Windows 自带的尺寸时 (紧凑主题的 28 px、显示缩放 125% / 175% 等) 不再有锯齿: 先取不小于需要尺寸的图标, 再用高质量的平滑缩放缩到实际大小后缓存。以前是取 32 px 的图标交给 Windows 直接压缩, 斜线和圆边会出现锯齿。正好是 16 / 32 / 48 px 时和以前一样

## [2026.10.01]

文字片段也能按正文里的词搜到, 编辑片段的窗口加了说明; 更新提示改用常见软件的写法。

Snippets can be found by words in their text, and the snippet editor explains each field; update notices use the wording common software uses.

### 升级
- **从 2026.09.30.1 ~ 2026.09.30.3 升级**: 呼出搜索窗口, 空搜索框下面有一条 "更新 ALTRun 到 2026.10.01", 按 `Enter` 更新
- **从 2026.09.26 ~ 2026.09.30 升级**: 托盘图标 → 检查更新 (2026.09.26 ~ 2026.09.28 是弹窗里选 "立即更新")。用 Scoop 安装的用 `scoop update altrun`
- **从 2026.09.25 及更早的版本升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.10.01.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`; 从 2.x 升级见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)
- 设置和数据都保留, 设置文件的格式没有变化 (新增的 `Features.Snippets.SearchText` 自动用默认值)
- 以前为了能搜到而把整句话写成关键字的片段: 现在可以把关键字改短 (例如 `pq`), 名称写成容易想到的几个词, 正文里的词照样能搜到, 也能自动展开

### 新增
- 文字片段也搜索正文: 输入 3 个字符以上时, 正文里包含每个词的片段也会列出, 排在名称、关键字匹配的后面 (不做模糊匹配, 避免命中一大堆)。偏好设置 → 文字片段 → 搜索 "也搜索片段正文" 可以关掉 (`Features.Snippets.SearchText`)
- 编辑文字片段的窗口里, 名称、关键字、正文、自动展开下面都加了灰色说明, 正文下面列出可以用的占位符

### 改进
- 更新提示的文字参照常见软件的写法: 搜索窗口里是 "发现新版本: ALTRun x", 下面一行是 "当前版本 y · Enter 安装更新 · → 更新说明、跳过此版本"; 手动检查的对话框标题改为 "ALTRun 有可用的更新", 按钮是 安装更新 / 更新说明 / 稍后
- 操作面板 (`→`) 右侧的说明按实际宽度显示, 不再只显示一截 (例如 "复制文件" 旁边的 "然后在 TC / 资源管理器里粘贴"、"跳过此版本" 旁边的说明)
- 片段关键字太长 (前缀加关键字超过 40 个字符) 或有空格、又要自动展开时, 保存前提示; 这样的关键字不再注册, 调试日志里不再每次启动都报错
- 浏览器书签改为在后台每 30 秒检查一次有没有变, 不再在打字时重新读取 (以前书签文件变了之后的第一次搜索会慢 20 ~ 30 ms)
- 调试日志里启动耗时最大的一段拆成 搜索窗口 / 托盘和热键 / 片段和扩展 / 快捷方式 四段记录

### 修复
- 操作面板的 "属性" 对网络盘上的文件夹等打不开 (提示 "Failed attempt to launch program or document"): 改用和资源管理器一样的方法打开属性窗口; 应用商店应用没有属性窗口, 不再显示这一项

## [2026.09.30.3]

关掉开机自动启动等设置时不再删掉用户自己建的快捷方式 (#113); 通知和任务管理器里显示的程序名改为 `ALTRun`。

Turning off "Launch at login" and similar settings no longer deletes shortcuts you created yourself (#113); notifications and Task Manager show the program name as `ALTRun`.

### 升级
- **从 2026.09.30.1 / 2026.09.30.2 升级**: 呼出搜索窗口, 空搜索框下面有一条 "更新 ALTRun 到 2026.09.30.3", 按 `Enter` 更新
- **从 2026.09.26 ~ 2026.09.30 升级**: 托盘图标 → 检查更新 (2026.09.26 ~ 2026.09.28 是弹窗里选 "立即更新")。用 Scoop 安装的用 `scoop update altrun`
- **从 2026.09.25 及更早的版本升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.09.30.3.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`; 从 2.x 升级见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)
- 设置和数据都保留, 设置文件的格式没有变化。Windows 会缓存程序名, 通知里的新名字可能要重新登录后才显示

### 改进
- Windows 通知顶部、任务管理器、启动应用列表里显示的名称改为 `ALTRun`: 以前 exe 的 "文件说明" 写成了 "ALTRun - An effective launcher for Windows", Windows 把它当成程序名显示 (宣传语移到 exe 的 "备注" 里)

### 修复
- 关掉 "开机自动启动" / "添加到开始菜单" / "添加到发送到菜单" 时, 以前会删掉那个位置上所有叫 `ALTRun.lnk` 的快捷方式, 包括用户自己建的; 现在只删 ALTRun 自己建的 ([#113](https://github.com/zhugecaomao/ALTRun/issues/113))

## [2026.09.30.2]

偏好设置里内容较多的列表加高到 4 行, 竖向滚动条能显示滑块。

Longer lists in Preferences are 4 rows tall, so the vertical scroll bar shows its thumb.

### 升级
- **从 2026.09.30.1 升级**: 呼出搜索窗口, 空搜索框下面有一条 "更新 ALTRun 到 2026.09.30.2", 按 `Enter` 更新
- **从 2026.09.26 ~ 2026.09.30 升级**: 托盘图标 → 检查更新 (2026.09.26 ~ 2026.09.28 是弹窗里选 "立即更新")。用 Scoop 安装的用 `scoop update altrun`
- **从 2026.09.25 及更早的版本升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.09.30.2.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`; 从 2.x 升级见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)
- 设置和数据都保留, 设置文件的格式没有变化

### 改进
- 偏好设置里原来只有 2 ~ 3 行的列表 (应用搜索 → 文件夹、隐藏的应用; 对话框跳转 → 其它对话框、不生效的窗口、不自动跳转的对话框; 一键加日期 → 备注框的生效窗口) 改为 4 行: 以前内容多时竖向滚动条只放得下上下两个箭头, 看不到滑块

## [2026.09.30.1]

自动检查更新更及时 (启动时和之后每 6 小时检查), 偏好设置 → 文件搜索 的 "文件类型" 输入框加高。

Update checks are more timely (at startup and every 6 hours after); the "File types" box in Preferences → File Search is taller.

### 升级
- **从 2026.09.26 ~ 2026.09.30 升级**: 呼出搜索窗口, 空搜索框下面有一条 "更新 ALTRun 到 2026.09.30.1", 按 `Enter` 更新; 没有看到的话, 托盘图标 → 检查更新 (2026.09.26 ~ 2026.09.28 是弹窗里选 "立即更新")。用 Scoop 安装的用 `scoop update altrun`
- **从 2026.09.25 及更早的版本升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.09.30.1.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`; 从 2.x 升级见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)
- 设置和数据都保留, 设置文件的格式没有变化。今天第二次发布, 版本号是 `2026.09.30.1` (同一天再发布时在日期后面加 `.1`、`.2`)

### 改进
- **自动检查更新更及时**: 以前每 24 小时才检查一次, 前一天晚上发布的版本, 第二天开机后可能要过大半天才提示。现在启动 1 分钟后, 离上次检查满 1 小时就检查 (改完设置重新载入不会反复检查), 一直开着时每 6 小时检查一次
- 偏好设置 → 文件搜索 → "文件类型" 的输入框加高到 7 行, 默认的几种类型都能看到; 所有多行输入框 (应用程序的文件夹、隐藏的应用、剪贴板忽略的程序等) 加上竖向滚动条, 内容多时能看出来还有更多行
- 偏好设置 → 文件搜索 的说明文字: "文件关键字" (`open 报告`) 搜的是文件和文件夹 (以前写成 "只搜文件"); "文件类型" 下面注明只搜文件夹用上面的 "文件夹关键字"

## [2026.09.30]

打开 / 保存对话框下面自动出现文件夹面板 (和 Listary 的 Quick Switch 一样), 选中内容直接调出操作 (和 Alfred 的 Universal Actions 一样), 剪贴板历史记录文件和图片, 浏览器书签、单位 / 货币换算, 读取设置和索引快了很多。

A folder panel appears below Open / Save dialogs automatically (like Listary's Quick Switch), actions on selected content (like Alfred's Universal Actions), clipboard history records files and images, browser bookmarks, unit / currency conversion, and much faster loading of settings and indexes.

### 升级
- **从 2026.09.29 升级**: 呼出搜索窗口, 空搜索框下面有一条 "更新 ALTRun 到 2026.09.30", 按 `Enter` 更新 (或托盘图标 → 检查更新)
- **从 2026.09.26 ~ 2026.09.28 升级**: 发现新版本时选 "立即更新" (或托盘图标 → 检查更新), ALTRun 会自己下载安装
- 用 Scoop 安装的用 `scoop update altrun`
- **从 2026.09.25 及更早的版本升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.09.30.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`; 从 2.x 升级见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)
- 设置、自定义命令、剪贴板历史和学习记录都保留, 设置文件的格式没有变化 (新的设置自动用默认值)。对话框文件夹面板默认打开, 不需要的话在 偏好设置 → 对话框面板 关掉

### 新增
- **选中内容的操作** (和 Alfred 的 Universal Actions 一样): 在任何程序里选中文字、文件或网址, 按 `Ctrl+Alt+\` (偏好设置 → 通用 可以改), 直接打开它们的操作面板: 文字可以复制、用各个搜索引擎搜索、存为片段、计算, 或者转成大写 / 小写 / 排序 / 简繁等之后替换回原来的程序; 文件可以打开、显示位置、复制路径、添加到自定义命令。取内容时临时借用剪贴板, 之后还原, 不记进剪贴板历史; 终端窗口里用 `Ctrl+Insert` 复制, 不会中断正在运行的程序
- **浏览器书签**: 直接输入书签名称就能找到 (排在应用和命令后面, 最多 8 条), `bm 关键词` 只搜书签, 也按网址的域名匹配。读 Chrome、Edge、Brave、Vivaldi 所有用户配置的书签 (Firefox 暂不支持), 书签有变化时自动重新读取。可以在 偏好设置 → 功能 关闭
- **单位换算**: `10 km in mi`、`100 f to c`、`20 MPa in psi`、`300 kN to kip`、`3 亩 in m2`、`1 GB to MB`... 长度、质量、面积、体积、速度、时间、数据、压强 / 应力、力、能量、功率、温度, 支持中文单位名称
- **货币换算** (默认关闭): `100 usd to sgd`、`100 美元 to 人民币`。在 偏好设置 → 功能 打开后, 每天从 frankfurter.dev 下载一次央行参考汇率 (这是 GitHub 之外唯一会访问的网站)
- **片段占位符**: `{date:yyyy-MM-dd}`、`{time:HH:mm:ss}` 自定格式; `{date+7}`、`{date-1:dd.MM}` 前后几天; `{clipboard:1}` 剪贴板历史里的前一条; `{uuid}`
- **媒体控制**: 系统命令 播放 / 暂停、下一首、上一首、停止播放 (也可以设成自定义热键)
- **剪贴板历史记录文件和图片**: 复制的文件按文件显示 (一个文件时可以打开、显示位置, 多个时显示文件名), `Enter` 粘贴回去还是文件; 截图、网页图片等存成 PNG (`Data\Clipboard\`, 默认最多 50 张, 只在保存历史时记录), 列表里显示缩略图和尺寸, `Enter` 粘贴图片。`clip 图片` 只看图片
- **连续复制合并** (默认关闭): 快速按两次 `Ctrl+C`, 把这次复制的文字接到上一条后面, 剪贴板里也是合并后的文字 (和 Alfred 的 Merging 一样)。在 偏好设置 → 剪贴板历史 打开
- **对话框文件夹面板** (和 Listary 的 Quick Switch 窗口一样): 打开 / 保存对话框一出现, 正下方自动贴一个同宽的面板, 点一个文件夹就跳过去, 不用记热键; 上面的搜索框还能搜索文件夹和文件 (Everything / 内置索引), 选中文件时跳到它所在的文件夹。跳转时文件名框里原来的文件名 (另存为时程序预填的名字) 会保留。列出每个 Total Commander 窗口的两个面板、打开的资源管理器窗口和最近用过的文件夹 (Windows 的 "最近使用的项目"); 面板不抢焦点, 回到对话框时刷新。键盘操作: 在对话框里按 `Ctrl+Shift+G`, 光标跳到面板的搜索框, `↑` `↓` 选择, `Enter` 跳转 (没有打开自动显示时临时显示面板)
- **按文件类型搜索**: `doc 报告`、`pic logo`、`cad 平面图`、`zip`、`video`、`audio`、`exe`; 类型和扩展名可以在 偏好设置 → 文件搜索 → "文件类型" 修改
- **更多文件操作** (操作面板): 打开方式、复制文件 / 剪切文件 (到 TC 或资源管理器里粘贴)、复制 / 移动到 TC (或资源管理器) 当前的文件夹 (Windows 自己的复制, 有进度、重名询问、可以撤销)、移到回收站
- **双击 Ctrl / Shift 呼出** (和 Listary 一样, 默认不用): 偏好设置 → 通用 → "双击呼出"

### 改进
- **对话框文件夹面板**:
  - 颜色跟随 ALTRun 的主题 (深色主题下面板也是深色, 选中行用主题的选中色);
  - 对话框改变宽度时面板跟着变;
  - 搜索框里用中文输入法按 `Enter` 上屏时不再直接跳转;
  - 新设置 "搜索框搜索": 只搜文件夹 (默认) / 文件夹和文件;
  - 面板的设置单独放在 偏好设置 → 对话框面板 这一页; 有了面板之后, 不再修改对话框的标题栏 (以前会在标题后面加上热键提示)。
- **偏好设置的布局**:
  - 成对的设置右边一列对齐 (例如 通用 → "界面语言" 和上面的 "第二个呼出热键" 对齐, 宽度也一样);
  - 外观页的 "主题"、"窗口显示在" 下拉框和下面的按钮一样宽;
  - 对话框跳转页的两个热键框一样宽;
  - 比较拥挤的两页分开: "对话框面板" 单独一页; Everything 的设置从 "文件搜索" 挪到 "文件索引" 页, 这一页改名为 "搜索来源"。
- **文件搜索找不到同名文件夹**: 文件和文件夹一起搜时, 如果同名的文件很多 (例如 `plans` 文件夹里有几百个 `plans_xxx.dwg`), 名字完全相同的文件夹可能排不进结果, 要写 `folder:plans` 才出来。现在用 Everything 时另外单独取一次文件夹, 同名的文件夹排在最前面。
- 对话框跳转和文件夹面板读取 TC 的目录: 以前临时借用剪贴板 (让 TC 把路径复制到剪贴板再读出来), 每次打开对话框 Windows 的剪贴板历史 (`Win+V`) 里都会多出几条路径; 现在直接用 TC 的 `WM_COPYDATA` 接口问 TC 要路径, 不再碰剪贴板 (需要 TC 8.0 以上)
- **读取设置和索引快了很多**: JSON 读取以前随文件大小平方变慢 (AutoHotkey 按值传字符串时会整段复制): 140 KB 的设置文件要 1.1 秒, 6000 条路径的文件索引要 9 秒, 3 万条要几分钟。现在是线性的: 同样的设置文件约 50 ms, 3 万条路径约 0.35 秒。启动、重新载入、剪贴板历史和学习记录的读取都受益
- 启动时 Everything 正在运行的话, 不再读取和扫描内置文件索引 (最多 3 万条, 用不上), 等真的要用时再读
- 图标缓存最多 1500 个, 超过时去掉最久没用的: ALTRun 常常连续运行几周, 图标句柄不再只增不减
- 剪贴板历史里超过 4000 字的条目单独存成 `Data\Clipboard\*.txt`, 每次复制时不用把很长的文字重新写一遍
- 调试日志 (偏好设置 → 通用 → 写入调试日志) 里记录启动各阶段和超过 30 ms 的搜索用了多少毫秒, 方便找出慢在哪里 (不记输入的文字)

### 修复
- 编辑自定义命令时, 类型选 "文件夹" 后点目标旁边的 "浏览", 以前只能选文件, 现在选择文件夹
- 对话框文件夹面板: 程序自己关掉对话框时 (光标还在面板的搜索框里), 面板不再留在屏幕上; 点一项时如果没能切回对话框, 不再把路径打到别的窗口里
- 书签: 浏览器正在写书签文件时读取失败, 以前要等下次书签有变化才重新读, 现在下次搜索时就重新读
- 热键录制框: 关掉偏好设置之后, 以前记着的框没有清理, 窗口句柄被重新使用时在别的控件上点鼠标中键可能被当成录制热键

## [2026.09.29]

热键直接按键录制 (不用再写 `!Space`), 检查更新和 Alfred 一样不再弹窗, 新版本显示在搜索窗口里。

Hotkeys are recorded by pressing the keys (no more writing `!Space`); like Alfred, update checks no longer pop up a window, and new versions appear in the search window.

### 升级
- **从 2026.09.26 ~ 2026.09.28 升级**: 发现新版本时选 "立即更新" (或托盘图标 → 检查更新), ALTRun 会自己下载安装; 用 Scoop 安装的用 `scoop update altrun`。设置、自定义命令和学习记录都保留, 设置文件的格式没有变化
- **从 2026.09.25 及更早的版本升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.09.29.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`; 从 2.x 升级见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)
- 这次仍是旧版本在启动时弹窗提示; 升级之后, 以后的新版本会显示在搜索窗口里 ("更新 ALTRun 到 x", 按 `Enter` 更新)

### 改进
- **热键直接按键录制** (和 Alfred、PowerToys 一样): 偏好设置里的热键框显示 `Alt+Space`、`Ctrl+Alt+C` 这样的写法, 不用再写 AutoHotkey 的 `!Space`。点一下框, 按下想要的组合键就记下来; `Esc` 取消, `Backspace` 清除。录制时 ALTRun 自己的热键暂停, `Alt+Space`、`Win+E` 这类组合也不会触发系统的功能。单独的字母、空格等会影响打字, 要加上 Ctrl / Alt / Win。用于呼出热键、第二个呼出热键、剪贴板历史、对话框跳转、一键加日期和自定义热键
  - 自定义热键: 在框里点鼠标中键或侧键录成鼠标热键; `~` 改成复选框 "保留按键原来的功能"; 列表里显示 "Ctrl+Alt+P"、"鼠标中键"
  - 保存时两个全局热键相同 (例如呼出热键和剪贴板历史都是 `Ctrl+Alt+C`) 会先提示
  - 设置文件里仍然是 AutoHotkey 的写法, 格式不变
  - 通过远程桌面 (Chrome Remote Desktop 等) 控制电脑时也能录制; 录制时关掉偏好设置, ALTRun 的热键照常可用
- **检查更新和 Alfred 一样, 不再弹窗**: 以前每次启动时检查, 有新版本就弹出对话框 (开机自启时很打扰)。现在在后台每天检查一次 (开机 1 分钟后, 之后每小时看一下是否满 24 小时, 一直开着也能及时发现), 有新版本时呼出搜索窗口, 空搜索框下面就有一条 "更新 ALTRun 到 x": `Enter` 立即更新, `→` 查看更新内容或跳过这个版本; 搜索 "更新" / `update` 也能找到。检查时间和跳过的版本记在 `Data\Update.json`。托盘菜单的 "检查更新" 仍然直接弹出结果
- 用 Scoop / winget 安装时, 搜索窗口里的更新提示按 `Enter` 复制升级命令; 运行源码 (`ALTRun.ahk`) 时提示用 `git pull` 更新
- 偏好设置 → 通用: "启动时检查更新" 改名为 "自动检查更新"

## [2026.09.28]

偏好设置补全 (对话框跳转、一键加日期、文件索引各有一页), 更多默认示例, `Alt+R` 第二热键, `F1` 关于; 自定义命令支持 `ms-settings:` 等链接。

Preferences completed (dialog jump, date stamp and file index each have a page), more default examples, `Alt+R` second hotkey, `F1` for About; custom commands support links such as `ms-settings:`.

### 升级
- **从 2026.09.26 / 2026.09.27 升级**: 发现新版本时选 "立即更新" (或托盘图标 → 检查更新), ALTRun 会自己下载安装; 用 Scoop 安装的用 `scoop update altrun`。设置、自定义命令和学习记录都保留, 设置文件的格式没有变化
- **从 2026.09.25 及更早的版本升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.09.28.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`; 从 2.x 升级见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)
- 新的默认示例 (自定义命令、文字片段) 和 `Alt+R` 第二热键只用于新安装; 已有的设置不会改动, 想用 `Alt+R` 可以在 偏好设置 → 通用 → "第二个呼出热键" 里设置

### 新增
- **更多默认示例** (新安装时): 自定义命令每种类型都有示例 (记事本、编辑 hosts 文件、下载文件夹、ping、Google 地图、Windows 更新...); 文字片段有日期时间、粘贴为纯文本、写邮件常用的几段 (感谢、附件、约会议、签名) 和地址
- **偏好设置里能改的设置更多了** (以前只能在 ALTRun.json 里改):
  - 新的 "文件索引" 页: Everything 没有运行时的内置索引 (文件夹、深度、排除、最多条数、刷新间隔)
  - 文件搜索: 最少字数; 剪贴板历史: 最长字数; 文字片段: 粘贴后的等待时间; 功能: 帮助 (?) 的开关
- **对话框跳转** 有了单独的一页: 生效的窗口 ("Windows 标准对话框" 复选框 + 其它对话框, 例如 WPS 的 `ahk_class Qt5QWindowIcon`)、不生效的窗口、不自动跳转的对话框 (新的 `AutoSwitchExclude`)
- **一键加日期** (原 "Ctrl+D 加日期", 热键可以改) 有了单独的一页: 重命名和备注两种场景各自的热键和生效的窗口列表
- 搜索窗口里按 `F1` 打开 "关于 ALTRun" (和 `F2` ~ `F4` 一样是内置的, 帮助速查表 `?` 里也列出了 `F1` 和 `F4`)
- 自定义热键的窗口填 `ALTRun` 表示只在 ALTRun 的搜索窗口里生效
- 自定义命令的类型 "网址" 改名为 "网址 / 链接": 也可以填 `ms-settings:windowsupdate`、`mailto:` 这样的链接, 系统设置链接显示齿轮图标
- 偏好设置的自定义命令列表里, "目标" 一列同时显示参数 (例如 `cmd.exe /k ipconfig /all`)

### 改进
- 第二个呼出热键默认是 `Alt+R` (新安装时; 已有的设置不变)
- 文件管理器的说明里加上 Total Commander 带参数的例子 (`/O /T /S`)
- 系统命令 "关于 ALTRun" 改为打开偏好设置的高级页 (版本、项目主页、检查更新), 不再是一个简单的对话框
- 自定义命令编辑窗口的提示改用通用的例子
- 使用统计里的 "Ctrl+D 加日期" 改名为 "一键加日期", 和偏好设置的页面名称一致
- README 和 Wiki 重新整理: 开头加上 "为什么选择 ALTRun", 补充一键更新、Scoop / winget、三种界面语言的说明; 安全策略写明联网只有检查更新和一键更新 (只访问 GitHub, 下载后核对 SHA256); 常见问题加上 "会联网或收集数据吗"、"怎样升级"

## [2026.09.27]

日语界面、使用统计、ALTRun 自己的程序图标; 偏好设置重新排版, 设置文件移到 `Data\`, 打字更跟手。

Japanese interface, usage statistics, ALTRun's own program icon; Preferences redesigned, the settings file moved to `Data\`, and typing feels more responsive.

### 升级
- **从 2026.09.26 升级**: 发现新版本时选 "立即更新" (或托盘图标 → 检查更新), ALTRun 会自己下载安装; 用 Scoop 安装的用 `scoop update altrun`。第一次启动时设置文件自动从程序目录移到 `Data\ALTRun.json`, 设置、自定义命令和学习记录都保留
- **从 2026.09.25 及更早的版本升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.09.27.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`; 从 2.x 升级见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)
- 以后要退回旧版本时, 请先把 `Data\ALTRun.json` 移回程序目录

### 新增
- **使用统计** (和 Alfred 的 Usage 一样): 偏好设置 → 使用统计, 显示今天 / 最近 7 天 / 最近 30 天 / 总共的使用次数、最近 30 天每天的柱状图、每个功能的次数和占比。只记录次数, 不记录输入和打开的内容, 保存在 `Data\Usage.json`
- **日语界面** (感谢 [@coolvitto](https://github.com/coolvitto), [#85](https://github.com/zhugecaomao/ALTRun/pull/85)): 偏好设置 → 通用 → 界面语言选 "日本語"; 日文 Windows 上选 "自动" 时直接使用日语, 字体用 Yu Gothic UI
- **新的内置主题 "紧凑深色" (DarkCompact)**: 和深色一样的配色, 字号和行高小一些, 一屏显示更多结果
- **ALTRun 自己的程序图标**: ALTRun.exe、托盘、窗口、开始菜单 / 开机启动 / "发送到" 快捷方式都用新的图标 (`Resources\ALTRun.ico`), 不再是 AutoHotkey 的绿色 H

### 改进
- **偏好设置界面更整齐** (和 Alfred / Listary 一样的两列表单): 左列是右对齐的标签, 同一页的输入框、下拉框、复选框和按钮都从同一条竖线开始, 每页的左列宽度按这一页的标签长短决定 (不再整体偏右); 复选框按组加标签 (启动、系统集成、更新...); 输入框统一为短 / 中 / 长三种宽度, 按钮统一尺寸; 分节标题下加一条细线
- 偏好设置的页面分节排列 (粗体标题 + 细线):
  - 搜索窗口: "窗口"、"输入"、"提示和历史"
  - 外观: "主题"、"大小"、"窗口位置"; 两个主题按钮并排, 窗口宽度和结果行数放在同一行
  - 功能: 下面加 "功能选项" 一节; 终端前缀和终端程序放在同一行
  - 应用搜索: "索引范围"、"搜索结果"
  - 文件搜索: "开始文件搜索"、"搜索结果"、"Everything"、"内置索引", 标签更短, 不再排到页面底边
  - 文字片段: 列表下面加 "选项" 一节; 网页搜索: 列表变矮, "兜底搜索" 的标签精简; 扩展功能: 两个热键放在同一行
  - 高级: 上方是程序图标、版本和项目链接 (和下面的按钮对齐), 下面是 "设置和数据" 和 "重置" 两节
- 偏好设置窗口高一些, 每页更宽松; 自定义命令列表的 "类型" 一列显示中文
- 偏好设置的左列按标签文字的实际宽度加宽 (日语等较长的标签不再换行、不和输入框重叠); 同一行放不下的两个设置自动换到下一行
- **提示信息更好看**: 操作后的提示 (已复制、已清空...) 改为跟随主题颜色和字体的圆角提示框, 固定显示在搜索窗口下方 (窗口已关闭时在屏幕中间偏下), 淡入淡出, 不再是位置不固定的黄色提示框。后台发生的事 (更新完成、导入或升级设置、索引重建完成、通过 "发送到" 添加命令) 改用 Windows 通知, 之后还能在通知中心看到 (隐藏了托盘图标时仍用提示框)
- **打字更跟手**: 启动后在后台预先算好自定义命令和应用的搜索数据 (名称、拼音首字母、目标路径)。以前启动后第一次输入要先算几百毫秒 (几百条命令时明显卡一下), 现在第一个字符和之后一样快
- 普通文件夹用系统的通用文件夹图标, 不再逐个读取 (自定义命令里文件夹很多时更流畅); 有自定义图标的文件夹 (desktop.ini) 和磁盘根目录仍显示自己的图标
- 后台加载图标每次最多占用 10 ms (以前 30 ms), 打字时更不容易被打断; 搜索窗口开着时不写学习记录和使用统计, 等窗口关闭后再保存

### 修复
- 大字显示 (Large Type) 和操作提示遇到很长的文字时报错 ("has no method named Destroy")

### 变更
- **设置文件移到 `Data\` 文件夹**: `ALTRun.json` 从程序目录移到 `Data\ALTRun.json`, 和索引、学习记录、剪贴板历史放在一起, 程序目录里只剩程序文件。升级后第一次启动时自动移动 (连同 `ALTRun.v*.backup.json` 升级备份), 并提示一次; 用 Scoop 安装的也一样, 移到 persist 里的 `Data\`。以后退回旧版本前, 请先把 `Data\ALTRun.json` 移回程序目录
- README 和 Wiki 的截图改为英文界面, 演示数据换成通用的示例内容

## [2026.09.26]

一键更新, SPF2M 束线型直接计算 (不再需要 DOSBox), 可以用 Scoop 安装。

One-key updates, SPF2M tendon profiles calculated directly (no DOSBox needed), installation with Scoop.

### 升级
- **从 2026.09.25 / 2026.09.24 / 2026.09.23 升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.09.26.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`。设置、自定义命令和学习记录都保留。`Resources\` 里旧的 `DOSBox.exe`、`SDL.dll`、`SDL_net.dll`、`SPF2M.exe`、`Run.bat` 已经不用了, 可以删掉
- **这是最后一次需要手动升级**: 从这个版本开始, 发现新版本时选 "立即更新" 即可, ALTRun 会自己下载安装 (用 Scoop 安装的用 `scoop update altrun`)
- **从 2.x (v2026.08.12 及更早) 升级**: 步骤相同, 第一次运行时自动导入旧的 `ALTRun.ini`, 见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)

### 新增
- **一键更新**: 发现新版本时可以选择 "立即更新", ALTRun 自己下载新版本、核对 SHA256、替换程序文件并重新启动, 设置和数据都保留; 出错时原来的程序保持不变。也可以用命令行 `ALTRun.exe -Update`。从这个版本开始生效 (升级到这个版本还需要手动一次)
- **SPF2M 束线型直接计算**, 不再需要 DOSBox: PT 工具箱的 SPF2M 窗口输入参数后立即显示各支架处的高度, 可以一键复制到 Excel。结果和原来的 SPF2M.EXE 完全一致 (139 组、1100 多个数值逐个比对)。另外可以指定最小曲率半径、反弯点距离和支架间距; SPF2M 会直接退出或什么都不显示的情况, 这里会给出说明
- **用 Scoop 安装**: `scoop bucket add altrun https://github.com/zhugecaomao/ALTRun`, 然后 `scoop install altrun`; 以后 `scoop update altrun` 升级, 设置、数据和自定义主题都保留。winget 清单也已准备好, 等官方仓库收录后可以 `winget install zhugecaomao.ALTRun`
- 用 Scoop / winget 安装时, "检查更新" 发现新版本会给出对应的升级命令

### 改进
- 右键 → 发送到 → ALTRun: 选中 1 个文件 / 文件夹时弹出编辑对话框 (名称、类型、目标已填好, 可以改名称、加关键字, 取消则不添加); 选中多个时全部直接添加, 已经有的命令不重复添加
- 编辑自定义命令的对话框加宽, 长路径和长命令能显示完整; 每个字段 (名称、类型、目标、参数、关键字) 下面加上灰色说明

### 变更
- 不再附带 DOSBox 和原来的 SPF2M.EXE (`DOSBox.exe`、`SDL.dll`、`SDL_net.dll`、`SPF2M.exe`、`Run.bat`), 发布包小约 4MB。升级后 `Resources\` 里这几个旧文件可以删掉
- SPF2M 窗口改名为 "SPF2M Post-Tensioning Tendon Profile Calculator"; 钢绞线类型默认 12S; 管道直径留空时显示默认值 (Slab 25、7S 70、12S 90、19S 100、22S 120、31S 130)
- 搜索记录改用 `Ctrl+↑` / `Ctrl+↓` 前后翻 (和 Alfred 一样), `↑` `↓` 只在结果里移动。以前翻出一条记录、用 `↓` 选到下面的结果后, 按 `↑` 会换成另一条记录, 而不是回到上一行

### 修复
- 同时发送多个文件到 ALTRun 时只添加了第一个

## [2026.09.25]

窗口位置可以自己定, 偏好设置更容易看懂, 以及计算器的几处修正。

Choose where the window appears, easier-to-understand Preferences, and a few calculator fixes.

### 升级
- **从 2026.09.24 或 2026.09.23 升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.09.25.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`。设置、自定义命令和学习记录都保留, 新增的选项自动使用默认值
- **从 2.x (v2026.08.12 及更早) 升级**: 步骤相同, 第一次运行时自动导入旧的 `ALTRun.ini`, 见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)

### 新增
- 搜索窗口可以拖动 (按住输入框四周的空白处)。偏好设置 → 外观 新增 "拖动窗口后记住位置" (默认关闭; 记住的是在屏幕里的相对位置, 换一块屏幕也适用) 和 "恢复默认位置"
- 偏好设置 → 外观 → "窗口显示在": 鼠标所在的屏幕 (默认, 和以前一样) / 主屏幕 / 当前窗口所在的屏幕
- 偏好设置里每个设置下面都有一行灰色的说明 (和 Alfred 一样)
- 可选的 "呼出窗口时保留上一次的搜索" (偏好设置 → 搜索窗口, 默认关闭, 和 2.x 的 "保留输入" 一样; 从 2.x 升级时沿用原来的设置): 上一次的输入、文件搜索模式和选中的行都保留, 文字全选, `Enter` 再执行一次, 直接输入开始新的搜索

### 变更
- 偏好设置窗口加高; 搜索窗口的选项 (失去焦点时隐藏、保留上一次的搜索、使用提示等) 从 "通用" 移到新的 "搜索窗口" 页

### 修复
- 计算器: 结果最多保留两位小数, 不再显示 `774.39999999999998` 这样的长尾巴 (`113.6+296.6+87.7+143.6+132.9` = `774.4`)
- 计算器: 算式里有相同的一段时结果算错 (`1*2+11*2` 得出 14, 应为 24); 大数溢出 (`100000^4`); 很小的结果算不出来。计算部分重写, PT Tools 的算式计算同样修正
- 失去焦点隐藏窗口后再呼出, 保留的搜索结果不显示 (只剩输入框里的文字)

## [2026.09.24]

输入更流畅、更容易上手, 以及几个常用的新功能。

Smoother typing, easier to get started, and a few useful new features.

### 升级
- **从 2026.09.23 升级**: 托盘图标 → 退出 ALTRun, 把 `ALTRun_v2026.09.24.zip` 里的文件解压到原来的文件夹 (覆盖), 再运行 `ALTRun.exe`。设置、自定义命令和学习记录都保留, 新增的选项自动使用默认值
- **从 2.x (v2026.08.12 及更早) 升级**: 步骤相同, 第一次运行时自动导入旧的 `ALTRun.ini`, 见 [2026.09.23 的说明](https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23)

### 新增
- **只搜文件夹**: 输入 `folder 预算`; 关键字可以在 偏好设置 → 文件搜索 → "文件夹搜索关键字" 修改。按空格进入文件搜索后也可以输入 `folder 预算`
- **使用提示**: 每次呼出窗口, 空搜索框里轮换显示一条使用提示, 开始输入就消失 (偏好设置 → 通用 可以关闭)
- **速查表**: 输入 `?` 列出所有输入语法和快捷键, `? f3` 这样可以过滤, `Enter` 打开 Wiki 对应的页面
- **检查路径**: 偏好设置 → 自定义命令 → "检查路径", 找出目标文件 / 文件夹已经改名或移走的命令, 在 "状态" 列中标出; 网络盘没有连接时显示 "无法访问", 不会误报为找不到
- **按空格执行选中项** (可选, 默认关闭): 偏好设置 → 通用; 和 2.x 的同名选项一样, 从 2.x 升级时沿用原来的设置, `Shift+空格` 输入空格 ([#21](https://github.com/zhugecaomao/ALTRun/issues/21))

### 改进
- **输入更流畅**: 自定义命令很多时 (例如从 2.x 导入的几百条), 输入第一个字母以前要 300~400 毫秒, 现在约 10~20 毫秒; 继续输入时每个字约 1~2 毫秒。常选的命令仍然靠前
- **网络盘图标**: 网络盘 (例如 `Q:\`) 和 `\\server\share` 上的文件 / 文件夹用通用图标, 不再逐个到网络上读取, 打字时不会被网络卡住
- **文件搜索排序**: 按名称匹配程度排序 (名称开头、单词开头优先, 同分时文件夹优先)。用 Everything 时先取前 300 条再排序, 以前只取最近修改的 30 条, 名称最匹配的文件夹可能被漏掉
- **偏好设置按钮** 改为 Windows 常见的 **确定 / 取消 / 应用 / 帮助**: 确定 保存并关闭, 不再重新打开窗口; 应用 保存后在同一页、同一位置重新打开, 没有修改时是灰色的; 有未保存的修改时 取消 / Esc 会先确认; 帮助 (F1) 打开当前页对应的 Wiki
- 只输入 `open` / `find` 等关键字 (后面还没有空格) 时, 不再隐藏其它结果
- 结果行数没变时不再调整窗口大小, 减少闪烁

### 其它
- README 和 Wiki 加上界面截图, 由 `Tests\Screenshots\TakeScreenshots.ahk` 在 GitHub Actions 的 Windows 机器上自动生成
- 每个 PR 自动在 Windows 上运行单元测试和语法检查; Dependabot 检查 GitHub Actions 更新; 安全策略 `SECURITY.md`

## [2026.09.23] - 3.0

3.0 参照 macOS 上的 Alfred 重新设计了搜索窗口和整体架构。

3.0 redesigns the search window and the whole architecture after Alfred for macOS.

### 从旧版本升级
1. 托盘图标 → 退出旧版本的 ALTRun
2. 下载下面的 `ALTRun_v2026.09.23.zip`, 把里面的文件全部解压到原来的 ALTRun 文件夹 (覆盖 `ALTRun.exe`)
3. 运行 `ALTRun.exe`

第一次运行时会自动从旧版本的 `ALTRun.ini` 导入设置、自定义命令和热键, 保存到新的 `ALTRun.json`;
`ALTRun.ini` 保持不变, 想退回旧版本时放回旧的 `ALTRun.exe` 即可。呼出热键、开机启动、"发送到" 菜单都会沿用。
旧版本的内置命令由新的系统命令取代; SPF2M 需要的 `DOSBox.exe` 等文件已经包含在 `Resources\` 里。
新的快捷键和用法见 [README](https://github.com/zhugecaomao/ALTRun#快捷键) 和 [Wiki](https://github.com/zhugecaomao/ALTRun/wiki)。

### 新增
- Alfred 式搜索窗口: 每行显示标题 + 路径 / 说明, 窗口高度随结果伸缩, `Ctrl+1~9` 直接执行
- 学习排序: 记住每次输入选了哪一项, 常用的自动靠前; 搜索框为空时 `↑` 调出最近的搜索
- 操作面板 (`→`) 和右键菜单: 以管理员运行、显示位置、复制路径 / 名称、在此打开终端、属性、大字显示
- 在结果里直接编辑: `F3` 修改自定义命令 / 片段 / 搜索引擎, 应用和文件一键加为自定义命令; `Ctrl+Del` 删除
- 文件和文件夹搜索: 在空的搜索框里按空格进入文件搜索模式 (和 Alfred 一样), 或用 `'` / `open` / `find`; Everything 在运行时通过 IPC 查询全盘, 否则使用内置的后台索引。普通搜索默认不再混入文件 (可在偏好设置里打开)
- 自定义命令也按目标的文件夹名 / 文件名匹配
- 应用: 不需要的应用可以从搜索结果中删除 (隐藏), 在偏好设置里恢复
- 剪贴板历史 (`Ctrl+Alt+C` / `clip`), 自动忽略密码管理器
- 文字片段和自动展开 (在任何程序里输入 `;关键字`)
- 网页搜索关键字和无结果时的兜底搜索; 终端命令 (`>命令`); 大字显示 (`Ctrl+L`)
- 系统命令: 电源、音量、Windows 工具、剪贴板文字转换 (大小写 / 排序 / 去重 / 简繁转换...)
- 图形化的偏好设置窗口, 按分类修改所有设置
- 9 套内置主题 (跟随系统、浅色、深色、经典、午夜、霜白、石墨、海洋、纸张), 支持自定义主题和 "复制为自定义主题"
- 设置自动升级: 从旧版本的 `ALTRun.ini` 导入, 之后按设置文件版本号 (`SchemaVersion`, 当前为 4) 逐版本升级

### 变更
- 代码按模块重新组织: `Lib\` 通用库, `Src\Core|UI|Providers|Extensions` 程序代码, `Resources\` 随程序发布的数据
- `Res\` 改名为 `Resources\`; 启动时自动把旧 `Res\` 里用户自己放的文件移过去
- 设置全部保存在 `ALTRun.json`, 运行时数据 (索引、学习记录、剪贴板历史) 放在 `Data\`

### 移除
- 旧的内置命令列表 (由系统命令取代)、执行历史和使用统计、旧的列表外观选项

### 修复
- 输入时的卡顿 (预先计算匹配数据, 继续输入时只在上次结果里筛选, 图标后台加载)
- 搜索框和编辑窗口里的长文字会变成多行并盖住下面的控件
- 网络驱动器断开时显示结果会卡住

## 更早的版本
见 [GitHub Releases](https://github.com/zhugecaomao/ALTRun/releases)。

[未发布]: https://github.com/zhugecaomao/ALTRun/compare/2026.10.20.4...HEAD
[2026.10.20.4]: https://github.com/zhugecaomao/ALTRun/compare/2026.10.20.3...2026.10.20.4
[2026.10.20.3]: https://github.com/zhugecaomao/ALTRun/compare/2026.10.20.2...2026.10.20.3
[2026.10.20.2]: https://github.com/zhugecaomao/ALTRun/compare/2026.10.20.1...2026.10.20.2
[2026.10.20.1]: https://github.com/zhugecaomao/ALTRun/compare/2026.10.20...2026.10.20.1
[2026.10.20]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.10.20
[2026.10.14]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.10.14
[2026.10.12]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.10.12
[2026.10.08]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.10.08
[2026.10.05]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.10.05
[2026.10.04]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.10.04
[2026.10.03]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.10.03
[2026.10.02]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.10.02
[2026.10.01]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.10.01
[2026.09.30.3]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.30.3
[2026.09.30.2]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.30.2
[2026.09.30.1]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.30.1
[2026.09.30]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.30
[2026.09.29]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.29
[2026.09.28]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.28
[2026.09.27]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.27
[2026.09.26]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.26
[2026.09.25]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.25
[2026.09.24]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.24
[2026.09.23]: https://github.com/zhugecaomao/ALTRun/releases/tag/2026.09.23
