[English](Bookmarks) · **中文**

# 浏览器书签

ALTRun 可以直接搜索浏览器书签, 不必先打开浏览器再找收藏夹。

## 搜索与打开

- 输入 `bm 关键字` 只搜索书签。只输入 `bm` 时显示一行提示, 按 `Enter` 补全成 `bm `。
- 直接输入书签名称时, 匹配的书签也会出现在普通搜索结果里, 排在应用和自定义命令后面, 最多显示 8 条。也会按网址的域名匹配, 例如输入网站域名能找到该网站的书签。
- 按 `Enter` 打开书签; 按 `Alt+Enter` 复制书签 URL。
- 支持 Chrome、Edge、Brave 和 Vivaldi 的所有用户配置文件。Firefox 使用 SQLite 数据库存储书签, 暂不支持。
- 便携版或其他基于 Chromium 的浏览器: 在 **其他浏览器** 里加上它的文件夹 (见[便携版和其他浏览器](#便携版和其他浏览器))。

ALTRun 启动时读取一次, 之后每 30 秒在后台检查书签文件是否变化, 变化后自动重新读取, 不会在打字时额外读取。

## 浏览器历史记录

- 输入 `h 关键词` 搜索浏览器历史记录: 标题或网址里包含输入的每个关键词的网页, 按最近访问时间排列。中文标题也能用拼音首字母找到。只输入 `h ` (带空格) 列出最近访问的网页; 只输入 `h` 时显示一行提示。
- 按 `Enter` 在默认浏览器里打开; 按 `Alt+Enter` 复制网址。
- 只在输入关键字时才搜, 历史记录不会出现在普通搜索结果里。
- 支持的浏览器和书签一样: Chrome、Edge、Brave 和 Vivaldi 的所有用户配置 (每个配置最近的 10000 个网页); Firefox 暂不支持。
- 只在输入 `h ` 时读取各配置的 `History` 数据库, 隔一分钟以上再搜时重新读取; 启动时和后台都不读。通过 Windows 10 / 11 自带的 SQLite (`winsqlite3.dll`) 直接只读打开, 不锁定文件, 浏览器开着也可以。ALTRun 不另存历史记录, 也不发送到任何地方; 从结果里打开的网页和其他结果一样会记进学习排序和最近使用。

## 便携版和其他浏览器

便携版浏览器 (例如 Google Chrome Portable) 或其他基于 Chromium 的浏览器把用户配置放在自己的文件夹里, ALTRun 猜不到。在偏好设置 → 功能 → 功能选项里, 点 **其他浏览器** 旁边的 **编辑...**, 每行加一个文件夹:

- 浏览器的 **User Data** 文件夹, 也就是里面有 `Default`、`Profile 1` 等的那个 (PortableApps.com 的 Google Chrome Portable 是它文件夹里的 `Data\profile`), 或者
- 单个用户配置文件夹 (里面有 `Bookmarks` 和 `History` 文件的那个)。

之后书签和浏览器历史记录都会包括它的各个配置。页面上的框显示每个文件夹在这台电脑上找到几个配置 ("没找到配置" 说明文件夹不对)。可以用环境变量 (`%OneDrive%`) 和 `A_` 变量; 相对路径从 ALTRun 所在的文件夹算起, 浏览器和 ALTRun 在同一个 U 盘上时, 换了盘符也能找到 (例如 `..\GoogleChromePortable\Data\profile`)。用通常方式安装的 Chrome、Edge、Brave、Vivaldi 不用加。

## 设置

在偏好设置 → 功能里启用或关闭“浏览器书签”和“浏览器历史记录”。在“功能选项”中可以设置:

| 设置 | 默认值 | 说明 |
|---|---|---|
| `Features.Bookmarks.Enabled` | `1` | 是否启用书签搜索 |
| `Features.Bookmarks.Keyword` | `bm` | 只搜索书签的关键字 |
| `Features.Bookmarks.InDefaultResults` | `1` | 普通搜索时是否也显示书签 |
| `Features.Bookmarks.BrowserFolders` | `[]` | 其他浏览器: 便携版或其他基于 Chromium 的浏览器的 User Data 或用户配置文件夹, 书签和历史记录都用 |
| `Features.BrowserHistory.Enabled` | `1` | 是否启用浏览器历史记录搜索 |
| `Features.BrowserHistory.Keyword` | `h` | 搜索浏览器历史记录的关键字 |

更多输入语法见[搜索与快捷键](Usage-zh#输入语法)。
