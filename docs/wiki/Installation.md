# 安装与升级

## 系统要求
- Windows 10 或 Windows 11 (64 位)
- 用源码运行时需要 [AutoHotkey v2](https://www.autohotkey.com/) 2.0 或更新版本; 发布的 exe 不需要

## 安装
ALTRun 是绿色软件, 不需要安装, 不写注册表。

1. 从 [Releases](https://github.com/zhugecaomao/ALTRun/releases/latest) 下载最新版本, 解压到任意文件夹 (例如 `D:\Apps\ALTRun`)
   - 或者克隆仓库, 双击 `ALTRun.ahk` 用 AutoHotkey v2 运行
2. 运行后托盘出现 ALTRun 图标, 按 `Alt+Space` 呼出搜索窗口
3. 第一次运行会在后台建立应用索引, 几秒后就能搜到开始菜单里的程序

不要放在需要管理员权限才能写入的文件夹 (例如 `C:\Program Files`), ALTRun 需要在自己的目录里保存设置。

### 用 Scoop 安装
[Scoop](https://scoop.sh/) 是 Windows 的命令行软件管理工具。ALTRun 的仓库本身就是一个 Scoop bucket:

```powershell
scoop bucket add altrun https://github.com/zhugecaomao/ALTRun
scoop install altrun
```

- 安装后开始菜单里有 "Scoop Apps → ALTRun", 程序在 `scoop\apps\altrun\current`
- 升级: 先退出 ALTRun, 再运行 `scoop update altrun`。设置 (`ALTRun.json`)、`Data\`、`Themes\` 都会保留
- 卸载: `scoop uninstall altrun`, 设置留在 `scoop\persist\altrun`, 重新安装后自动恢复; 连设置一起删除用 `scoop uninstall altrun --purge`
- 用 Scoop 安装时, "检查更新" 发现新版本会提示用 `scoop update altrun` 升级, 不要手动解压覆盖

### 用 winget 安装
winget 清单已经准备好 (`packaging\winget`), 等 [winget 官方仓库](https://github.com/microsoft/winget-pkgs) 收录后就可以:

```powershell
winget install zhugecaomao.ALTRun
winget upgrade zhugecaomao.ALTRun     # 以后升级 (先退出 ALTRun)
```

- 程序在 `%LOCALAPPDATA%\Microsoft\WinGet\Packages\zhugecaomao.ALTRun_...`, 命令行里可以直接输入 `altrun` 启动
- winget 本身不创建开始菜单快捷方式: 第一次用 `altrun` 启动后, ALTRun 会自己添加到开始菜单 (偏好设置 → 通用 → "添加到开始菜单")
- 卸载前先在 偏好设置 → 通用 里关闭 "开机自动启动"、"添加到开始菜单"、"添加到 '发送到' 菜单", 否则这些快捷方式会留下
- 第一次运行时 Windows 可能提示 "无法验证发布者" (程序是从网上下载的), 选择 "运行" 即可
- 升级和卸载都保留设置、`Data\`、`Themes\`; 卸载后想彻底删除, 手动删掉上面的文件夹

## 程序目录里的文件
| 文件 / 文件夹 | 说明 | 升级时 |
|---|---|---|
| `ALTRun.ahk` 或 `ALTRun.exe` | 主程序 | 替换 |
| `Lib\` `Src\` | 程序代码 (源码版) | 替换 |
| `Resources\` | 随程序发布的数据: 简繁对照表、内置主题 | 覆盖 |
| `ALTRun.json` | 你的全部设置和自定义命令、片段 | 保留 (自动升级格式) |
| `Data\` | 应用索引、文件索引、学习记录、剪贴板历史 | 保留 (删掉会重新生成) |
| `Themes\` | 你自己的主题 | 保留 |

## 开机启动、开始菜单、"发送到"
在 偏好设置 → 通用 里设置:
- **开机自动启动**: 在 "启动" 文件夹里创建快捷方式
- **添加到开始菜单**: 可以从开始菜单打开 ALTRun
- **添加到 "发送到" 菜单**: 在资源管理器里右键文件 → 发送到 → ALTRun, 把它加为自定义命令 (1 个时弹出编辑对话框, 多个时直接全部添加)

## 升级
把新版本复制覆盖到程序目录即可, 设置和数据不受影响。设置文件的格式有变化时, ALTRun 启动时会自动升级, 并先备份原文件。

### 从 2.x (v2026.08.12 及更早) 升级到 3.0
1. 托盘图标 → 退出旧版本
2. 把新版本的压缩包解压到原来的 ALTRun 文件夹, 覆盖 `ALTRun.exe`
3. 运行 `ALTRun.exe`

第一次运行时, 旧版本的 `ALTRun.ini` 会自动导入 (设置、自定义命令、热键), 保存为新的 `ALTRun.json`。`ALTRun.ini` 保持不变: 想退回旧版本, 放回旧的 `ALTRun.exe` 即可。开机启动、开始菜单、"发送到" 的快捷方式沿用同一个 `ALTRun.exe`, 不需要重新设置。

- **保留**: 呼出热键、开机启动等通用设置; 用户命令 (File / Dir / CMD / URL → 自定义命令, Clip → 文字片段); 索引目录、文件类型和深度; 结构计算开关; 对话框快速跳转; Ctrl+D 加日期; PT 工具箱设置; 条件热键
- **不再保留**: 内置命令列表 (由 [系统命令](Extensions#系统命令) 取代)、旧的索引 (会重新建立)、执行历史、使用统计、旧的列表外观选项

旧版本的 `Res\` 文件夹已改名为 `Resources\`, 启动时会自动把 `Res\` 里你自己放的文件移过去。

## 卸载
1. 托盘菜单 → 退出
2. 在 偏好设置 → 通用 里先关闭 "开机自动启动"、"添加到开始菜单"、"添加到 '发送到' 菜单" (或者手动删除对应的快捷方式)
3. 删除程序文件夹
