[English](Web-Search) · **中文**

# 网页搜索

输入 "关键字 搜索词" 用对应的网站搜索, 例如 `g autohotkey hotkeys` 用 Google 搜索, `Enter` 在默认浏览器里打开。

![网页搜索: g autohotkey v2 hotkeys](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/websearch.png)

## 自带的搜索引擎
| 关键字 | 网站 |
|---|---|
| `g` | Google |
| `bing` | Bing |
| `bd` | 百度 |
| `gh` | GitHub |
| `wiki` | 维基百科 (英文) |
| `yt` | YouTube |
| `tb` | 淘宝 |
| `jd` | 京东 |
| `tr` | Google 翻译 (翻成简体中文) |

- 输入关键字后面加了空格 (例如 `g `), 结果里只显示这个网站, 不和应用、文件混在一起
- 只输入关键字 (`g`) 时提示 "在 'g ' 后面输入内容", `Enter` 补全成 `g `; 输入网站名称 (例如 `google`) 也会提示
- 关键字加空格 (`g `) 后不再输入, 就搜索**剪贴板里的文字**: 自带的这几个网站的网址里都写了 `{query|clipboard}`, 列表里显示 "用 Google 搜索剪贴板: '...'", `Enter` 打开; 剪贴板里没有文字 (复制的是文件、空的) 时提示 "剪贴板里没有文字", 不能打开。继续输入了搜索词就搜输入的
- 在任何程序里选中文字后按 `Ctrl+Alt+\`, 操作面板里可以用每个网站搜索选中的文字, 见 [选中内容的操作](Usage-zh#选中内容的操作)

## 添加和修改
偏好设置 → 网页搜索: 列表里双击修改, 或者点 "添加"。每个引擎有这几项:

| 字段 | 说明 |
|---|---|
| 名称 | 显示在结果里, 例如 "Search Google for ..." |
| 关键字 | 输入这个词再加空格就用这个网站搜索。可以写几个, 用逗号分隔, 例如 `g, ㅎ, google` (经常切换键盘布局 / 输入法时, 输入哪个都行); 提示和补全用输入的那个 |
| 网址 | 搜索结果页的地址, `{query}` 换成输入的搜索词 (自动编码)。写成 `{query|默认值}` 时, 输入关键字加空格、不输入搜索词就搜索默认值, 例如 `https://www.bing.com/search?q={query|天气}` 输入 `bing ` 就搜索 "天气"; 默认值写成 `clipboard` (`https://www.google.com/search?q={query|clipboard}`) 时搜索剪贴板里的文字 |
| Id | 区分引擎用的名字 (英文字母), 兜底搜索里用它 |
| 图标 | 可选: `.ico`、`.png` 或 `.exe` 文件, 相对路径从 Data 文件夹算起。点 **下载网站图标** 会访问网址所在的网站一次, 找到的图标存到 `Data\Icons` 并填好这一项。留空 = 自带的图标 (上面的几个引擎都有) 或浏览器的图标 |

例子:

| 名称 | 关键字 | 网址 |
|---|---|---|
| 必应词典 | `dict` | `https://cn.bing.com/dict/search?q={query}` |
| Amazon | `amz` | `https://www.amazon.com/s?k={query}` |
| 本地开发服务器 | `lh` | `http://localhost:{query|8000}` (8000 是端口号: `lh ` 打开 8000, `lh 3000` 打开 3000) |

### 搜索剪贴板
- 只在输入了关键字加空格 (`g `) 时才读取剪贴板; 只输入关键字 (`g`) 或引擎名称 (例如 `google`) 的补全提示不读
- 只读文字: 空白合并成一个空格, 最多 200 个字符。文字和图片一起复制时 (例如 Excel 的单元格) 搜索文字; 复制的是文件时不搜索
- 密码管理器 (KeePass、1Password、Bitwarden...) 复制的内容不读, 和剪贴板历史一样; 回车前结果下面显示完整网址, 能看清要发出去的内容
- 从旧版本升级时, 自带引擎的网址如果还是原来的默认值 (`{query}`), 自动换成 `{query|clipboard}`; 改过网址的和自己加的引擎不动, 想要这个效果把网址里的 `{query}` 改成 `{query|clipboard}` 即可
- 剪贴板历史里选中一条文字按 `→`, 也可以用每个网站搜索这一条

## 兜底搜索
没有任何结果时, 列表里显示几个兜底项, 默认是 "用 Google 搜索 '...'"、"搜索文件"、"用 Bing 搜索 '...'"。在 偏好设置 → 网页搜索 → "兜底搜索" 里改: 写引擎的 Id, 或 `files` (文件搜索), 用逗号分隔。

## 设置文件
`Features.WebSearch` → `Engines` (引擎列表, 每一条 `Id` / `Keyword` / `Title` / `Url` / `Icon`) 和 `Fallbacks`, 见 [设置文件参考](Configuration-zh#websearch-网页搜索)。整个功能可以在 偏好设置 → 功能 里关掉。
