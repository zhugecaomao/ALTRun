[English](Home) · **中文**

# ALTRun 使用文档

ALTRun 是一个参照 macOS 上的 [Alfred](https://www.alfredapp.com/) 设计的 Windows 启动器, 基于 AutoHotkey v2, 开源免费、绿色便携。按 `Alt+Space` 呼出, 输入名称, `Enter` 打开。官网: https://zhugecaomao.github.io/ALTRun/

- **小**: 约 2.5 MB, 一个 exe 加少量资源文件, 不需要 .NET、Electron 等运行库
- **快**: 输入即搜; [Everything](https://www.voidtools.com/) 在运行时毫秒级搜遍全盘文件
- **越用越懂你**: 记住每次的选择, 常用的自动排第一; 支持单词首字母 (`vsc`) 和拼音首字母 (`wx` → 微信)
- **一个窗口处理日常小事**: 计算器、网页搜索、剪贴板历史、文字片段自动展开、系统命令、终端
- **绿色便携, 数据默认保存在本机**: 不写注册表, 不需要管理员权限, 没有后台服务; 设置默认在 `Data\`, 程序目录不可写时改用 `%APPDATA%\ALTRun\Data\`, 也可在偏好设置 → 高级更改位置; 一键更新, 设置不变
- **简体中文 / 繁體中文 / English / 日本語 界面**, 22 套内置主题 (每套都有宽松 / 常规 / 紧凑三种密度), 也可以自己做主题

> 本文档适用于 **3.0** 及以后的版本。3.0 重新设计了搜索窗口和设置文件, 2.x 的设置会在第一次启动时自动转换, 见 [安装与升级](Installation-zh)。

![ALTRun: 输入即搜、操作面板、单位换算、文件搜索](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/demo.png)

## 从这里开始
| 页面 | 内容 |
|---|---|
| [5 分钟上手](Getting-Started-zh) | 第一次用: 呼出、打开程序、常用功能、加自己的命令 |
| [安装与升级](Installation-zh) | 下载、Scoop / winget、开机启动、一键更新、从 2.x 升级、卸载 |
| [搜索与快捷键](Usage-zh) | 所有输入语法、快捷键、操作面板、学习排序、使用统计 |
| [自定义命令](Custom-Commands-zh) | 命令类型、带参数的命令、路径变量 |
| [文字片段](Snippets-zh) | 占位符、自动展开、不展开的窗口、导入 / 导出 |
| [剪贴板历史](Clipboard-zh) | 文字、图片和文件的历史, 粘贴、合并、排除的程序 |
| [文件搜索](File-Search-zh) | Everything 联动、内置索引、排除规则 |
| [主题](Themes-zh) | 22 套内置主题 (也可以跟随系统浅色 / 深色)、显示密度、自定义主题、全部可用的键 |
| [浏览器书签](Bookmarks-zh) | 搜索、打开和复制书签 URL |
| [网页搜索](Web-Search-zh) | `g` / `bd` 等关键字、自己加搜索引擎、兜底搜索 |
| [窗口切换](Window-Switching-zh) | 搜索已经打开的窗口并切换过去 |
| [置顶和最近使用](Recent-zh) | 空搜索框显示的置顶项目和最近打开的项目 |
| [对话框快速跳转](Quick-Switch-zh) | 跳到资源管理器或 Total Commander 已打开的文件夹 |
| [对话框文件夹面板](Quick-Switch-Panel-zh) | 在打开 / 保存对话框附近搜索并选择文件夹 |
| [一键加日期](Date-Stamp-zh) | 重命名文件或在文本末尾添加日期 |
| [计算器](Calculator-zh) | 算式、单位和货币换算、附带的结构计算 |
| [脚本扩展](Scripts-zh) | 把自己的脚本变成搜索结果里的命令 |
| [系统命令和终端](System-Commands-zh) | 锁屏、关机等系统命令的完整列表; `>` 运行命令 |
| [自定义热键](Hotkeys-zh) | 给热键指定系统命令、CapsLock 切换输入法、怎样设置热键 |
| [扩展功能](Extensions-zh) | PT 工具箱, 以及各扩展页面的入口 |
| [设置文件参考](Configuration-zh) | ALTRun.json 每一项的含义和默认值 |
| [数据文件](Data-Files-zh) | 数据文件夹在哪里、每个文件是什么、怎样备份和换电脑 |
| [隐私与网络访问](Privacy-zh) | ALTRun 什么时候联网、访问哪里、本机保存了什么 |
| [常见问题](FAQ-zh) | 热键冲突、搜不到、杀毒软件误报、隐私... |
| [开发指南](Development-zh) | 架构、新增搜索功能、代码规范、测试 |

## 30 秒上手
| 想做什么 | 输入 |
|---|---|
| 打开程序 | `word` → `Enter` |
| 打开文件 / 文件夹 | `空格` 再输入 `报告`, 或 `'报告` |
| 计算 | `12*(3+4)` → `Enter` 复制结果 |
| 网页搜索 | `g 关键词` / `bd 关键词` |
| 剪贴板历史 | `Ctrl+Alt+C` 或 `clip` |
| 运行命令 | `>ipconfig /all` |
| 锁屏 / 关机 | `lock` / `shutdown` |
| 粘贴常用文字 | 任何程序里输入 `;sig` 自动展开签名 |
| 打开 / 保存对话框跳到某个文件夹 | 在[对话框文件夹面板](Quick-Switch-Panel-zh)里点一下, 或按 `Ctrl+G` 跳到 Total Commander 当前的文件夹 |
| 修改选中的命令 | `F3` |
| 更多操作 | `→` 或鼠标右键 |
| 忘了怎么写 | `?` 查看所有输入语法和快捷键 |

有问题或建议: [Issues](https://github.com/zhugecaomao/ALTRun/issues) · [Discussions](https://github.com/zhugecaomao/ALTRun/discussions)
