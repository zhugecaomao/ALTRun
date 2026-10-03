# ALTRun 使用文档

ALTRun 是一个参照 macOS 上的 [Alfred](https://www.alfredapp.com/) 设计的 Windows 启动器, 基于 AutoHotkey v2, 开源免费、绿色便携。按 `Alt+Space` 呼出, 输入名称, `Enter` 打开。官网: https://zhugecaomao.github.io/ALTRun/

- **小**: 下载不到 1 MB, 解压后约 2 MB, 一个 exe 加少量资源文件, 不需要 .NET、Electron 等运行库
- **快**: 输入即搜; [Everything](https://www.voidtools.com/) 在运行时毫秒级搜遍全盘文件
- **越用越懂你**: 记住每次的选择, 常用的自动排第一; 支持单词首字母 (`vsc`) 和拼音首字母 (`wx` → 微信)
- **一个窗口处理日常小事**: 计算器、网页搜索、剪贴板历史、文字片段自动展开、系统命令、终端
- **绿色便携, 数据只在本机**: 不写注册表, 不需要管理员权限, 没有后台服务; 设置都在 `Data\` 文件夹里; 一键更新, 设置不变
- **中文 / English / 日本語 界面**, 16 套内置主题, 也可以自己做主题

> 本文档适用于 **3.0** 及以后的版本。3.0 重新设计了搜索窗口和设置文件, 2.x 的设置会在第一次启动时自动转换, 见 [安装与升级](Installation)。

![ALTRun 搜索窗口](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/search.png)

## 从这里开始
| 页面 | 内容 |
|---|---|
| [安装与升级](Installation) | 下载、Scoop / winget、开机启动、一键更新、从 2.x 升级、卸载 |
| [搜索与快捷键](Usage) | 所有输入语法、快捷键、操作面板、学习排序、使用统计 |
| [自定义命令与片段](Commands-and-Snippets) | 命令类型、路径变量、片段占位符、自动展开、剪贴板历史 |
| [文件搜索](File-Search) | Everything 联动、内置索引、排除规则 |
| [主题](Themes) | 16 套内置主题 (也可以跟随系统浅色 / 深色)、自定义主题、全部可用的键 |
| [扩展功能](Extensions) | 对话框快速跳转、一键加日期、自定义热键、系统命令列表、PT 工具箱 |
| [设置文件参考](Configuration) | ALTRun.json 每一项的含义和默认值 |
| [常见问题](FAQ) | 热键冲突、搜不到、杀毒软件误报、隐私... |
| [开发指南](Development) | 架构、新增搜索功能、代码规范、测试 |

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
| 修改选中的命令 | `F3` |
| 更多操作 | `→` 或鼠标右键 |
| 忘了怎么写 | `?` 查看所有输入语法和快捷键 |

有问题或建议: [Issues](https://github.com/zhugecaomao/ALTRun/issues) · [Discussions](https://github.com/zhugecaomao/ALTRun/discussions)
