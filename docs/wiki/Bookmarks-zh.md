[English](Bookmarks) · **中文**

# 浏览器书签

ALTRun 可以直接搜索浏览器书签, 不必先打开浏览器再找收藏夹。

## 搜索与打开

- 输入 `bm 关键字` 只搜索书签。只输入 `bm` 时显示一行提示, 按 `Enter` 补全成 `bm `。
- 直接输入书签名称时, 匹配的书签也会出现在普通搜索结果里, 排在应用和自定义命令后面, 最多显示 8 条。也会按网址的域名匹配, 例如输入网站域名能找到该网站的书签。
- 按 `Enter` 打开书签; 按 `Alt+Enter` 复制书签 URL。
- 支持 Chrome、Edge、Brave 和 Vivaldi 的所有用户配置文件。Firefox 使用 SQLite 数据库存储书签, 暂不支持。

ALTRun 启动时读取一次, 之后每 30 秒在后台检查书签文件是否变化, 变化后自动重新读取, 不会在打字时额外读取。

## 设置

在偏好设置 → 功能里启用或关闭“浏览器书签”。在“功能选项”中可以设置:

| 设置 | 默认值 | 说明 |
|---|---|---|
| `Features.Bookmarks.Enabled` | `1` | 是否启用书签搜索 |
| `Features.Bookmarks.Keyword` | `bm` | 只搜索书签的关键字 |
| `Features.Bookmarks.InDefaultResults` | `1` | 普通搜索时是否也显示书签 |

更多输入语法见[搜索与快捷键](Usage-zh#输入语法)。
