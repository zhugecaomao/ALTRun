<h1 align="Center"><img width="45" alt="ALTRun" src="https://github.com/user-attachments/assets/91f36c04-3dd3-4486-9e7e-f70c9fabd6b8" /> ALTRun </h1>

ALTRun - 基于 AutoHotkey、开源免费、轻量高效、功能强大的启动工具

> 主界面 (深色模式)

![Main GUI](https://github.com/user-attachments/assets/143643a6-9ab1-44f1-ad1b-f9cb07c8c5fd)

> 主界面 (浅色模式)

![GUI](https://github.com/user-attachments/assets/a9da445b-4386-4118-bfb7-5bd2c5972efa)

> 主界面 (简化模式)

![Dark](https://github.com/user-attachments/assets/5bd76455-4eda-42e5-9934-c915b48994df)


## 特性
- 绿色便携和轻量化, 不修改系统注册表
- 低资源占用, 单程序文件
- 支持模糊搜索, 智能匹配
- 支持搜索时匹配中文拼音首字母
- 支持自定义命令、批量管理、命令分组
- 支持优先级智能排序, 根据使用频率自动调整命令
- 支持历史记录, 删除命令可撤销 (Ctrl+Z)
- 支持通过"发送到"菜单，快速创建命令
- 支持拖拽文件/文件夹/快捷方式到主窗口快速新建命令
- 支持 Everything 搜索集成
- 支持 Total Commander 文件管理集成
- 支持多种热键自定义, 多快捷键支持
- 支持数学表达式计算, 结构计算表达式
- 内置结构工程小工具: 钢筋/BRC 面积计算器 (PT Tools) 与 SPF2M 断面自动化 (SPF2M)
- 内置常用剪贴板文本处理命令 (大小写转换、排序去重、清除空行、反转等)
- 输入 `/` 即可列出并搜索全部内置命令, 方便发现功能
- 支持中英文界面自动切换
- 支持深度自定义, 托盘菜单、右键菜单、主窗口自定义
- 暗黑模式支持，可按照需求深度自定义
- 提供右键菜单和命令管理器，操作更便捷
- 支持自动添加日期到文件名/备注
- 支持类似 Listary 快速切换目录功能, 以及在当前目录直接打开终端
- 配置与命令数据保存为 JSON, 不受传统 ini 单节 64KB 限制, 旧版 ini 用户会自动迁移


## 快速开始
1. [下载程序](https://github.com/zhugecaomao/ALTRun/releases)
2. 双击运行 `ALTRun.exe`
3. 使用 `Alt+Space` 或自定义热键呼出主界面
4. 输入关键词搜索和运行命令


## 命令类型
每条命令都有一个类型, 在命令管理器 (F3 编辑 / Ctrl+N 新建) 里选择:

| 类型 | 说明 |
|---|---|
| `File` | 打开一个文件 |
| `Dir`  | 打开一个文件夹 |
| `CMD`  | 运行任意命令行 (可带参数, 例如控制面板小程序、外部程序) |
| `URL`  | 打开网址 |
| `Func` | 调用程序内置功能, 见输入 `/` 列出的完整列表 |
| `Clip` | 粘贴一段预先保存的文本片段, 支持 `{date}` `{time}` `{datetime}` `{clipboard}` `{arg}` `{cursor}` 占位符 |

搜索框还认几个特殊前缀:

| 前缀 | 行为 |
|---|---|
| `/` | 列出全部内置 `Func` 命令(命令面板), 后面接文字继续过滤 |
| `+` | 新建命令, 后面的文字自动带入描述 |
| 空格 | 交给 Everything 搜索文件 |
| `>` | Google 搜索 |

(以上三个前缀对应的具体动作在 `ALTRun.json` 的 `FallbackCommand` 里, 可自行替换成别的命令)


## 常用热键
- `Alt+Space` / `Alt+R`：显示/隐藏主窗口
- `Esc` / `失去焦点`：隐藏主窗口
- `上下箭头`: 选择上一个或下一个命令
- `回车`: 运行命令
- `F1`：帮助关于
- `F2`：配置选项
- `F3`：编辑当前命令
- `F4`：直接修改设置文件
- `Ctrl+D`：用文件管理器定位命令所在目录
- `Ctrl+D`：重命名文件(夹) (在重命名框时自动激活)
- `Ctrl+G`：打开/保存对话框路径快速跳转到 Total Commander 目录
- `Ctrl+E`：打开/保存对话框路径快速跳转到 资源管理器 目录
- `Ctrl+I`：重建索引数据库
- `Ctrl+Q`：重载脚本
- `Ctrl+Del`：删除当前命令
- `Ctrl+Z`：撤销上一次删除的命令
- `上下箭头`：切换命令
- `鼠标滚轮`：切换命令
- `Alt+序号`：运行对应序号命令
- `Ctrl+序号`：定位到对应序号命令
- `鼠标中键`：运行命令



## 配置说明
所有配置、命令、历史记录均保存在程序目录下的 `ALTRun.json` 文件中 (纯文本, 可直接用 F4 或记事本编辑)

可通过主界面或托盘菜单进入"配置选项"进行详细设置

- 仍在使用旧版 `ALTRun.ini` 的用户, 首次启动新版会自动迁移到 `ALTRun.json`, 迁移后原 ini 会备份保留
- `ALTRun.json` 如果被手动改坏导致无法解析, 程序会把它备份为 `ALTRun.json.bad` 并用默认配置重新启动, 不会静默覆盖数据


## 项目结构
从源码运行时直接运行 `ALTRun.ahk` (需要 AutoHotkey v2), 所有模块都在其开头用 `#Include` 显式列出:

```
ALTRun.ahk          主程序: 主窗口, 搜索, 运行命令, 热键, 以及供命令/热键调用的函数入口
Lib\                通用库, 与 ALTRun 无关 (JSON, Logger, Util, Dialogs)
Src\Core\           核心: 语言表, 配置读写 (AppData), 命令存储, 旧 ini 迁移
Src\UI\             窗口: 配置选项, 命令管理器, PTTools
Src\Features\       功能: 剪贴板, 简繁转换, Listary 式快速切换, 日期自动输入, 系统操作, 更新检查, 索引
Res\                数据文件 (Kanji.txt 简繁对照表)
```

> 从旧版手动复制文件升级时, 可以删除旧 `Lib\` 下已经移走的 `.ahk` 文件和 `Lib\Kanji.txt`


## 贡献
欢迎提交 Issue 和 PR，或在 [Discussions](https://github.com/zhugecaomao/ALTRun/discussions) 交流建议

如果你喜欢这个项目，请给它一个星标 ⭐


## 其他说明
更多详细用法、FAQ、进阶技巧请见 [Wiki](https://github.com/zhugecaomao/ALTRun/wiki)

特别感谢 [ALTRun by etworker](https://github.com/etworker/ALTRun) (Delphi) 和 [RunZ by goreliu](https://github.com/goreliu/runz) (AutoHotkey)
