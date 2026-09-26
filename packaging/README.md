# Scoop / winget 安装包

| 文件 | 说明 |
|---|---|
| `../bucket/altrun.json` | Scoop 清单。仓库本身就是一个 bucket: `scoop bucket add altrun https://github.com/zhugecaomao/ALTRun` |
| `winget/*.yaml` | winget 清单 (zip 里的便携版 ALTRun.exe), 提交到 [microsoft/winget-pkgs](https://github.com/microsoft/winget-pkgs) 后可以 `winget install zhugecaomao.ALTRun` |
| `Update-Manifests.ps1` | 把两个清单更新到新版本 (版本号、下载地址、SHA256) |

## 发布新版本时
Release 工作流 (publish = true) 创建 GitHub Release 之后会自动:
1. 运行 `Update-Manifests.ps1`, 把 Scoop / winget 清单更新到新版本并提交到 main。Scoop 用户 `scoop update altrun` 就能拿到新版本
2. 如果仓库设置了 `WINGET_TOKEN` secret, 用 [wingetcreate](https://github.com/microsoft/winget-create) 向 winget-pkgs 提交新版本的 PR (微软自动检查后合并, 一般 1~3 天)

`WINGET_TOKEN`: GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic), 勾选 `public_repo`; 然后在本仓库 Settings → Secrets and variables → Actions 里新建 `WINGET_TOKEN`。没有设置时跳过这一步, 也可以每次手动运行 `wingetcreate update` (见下文)。

## 第一次提交到 winget
winget-pkgs 里还没有 ALTRun 时, 自动更新不起作用, 需要先手动提交一次 (在 Windows 上):

```powershell
winget install Microsoft.WingetCreate      # 装好后关闭并重新打开 PowerShell
cd $env:TEMP
Invoke-WebRequest https://github.com/zhugecaomao/ALTRun/archive/refs/heads/main.zip -OutFile ALTRun-main.zip
Expand-Archive ALTRun-main.zip -DestinationPath . -Force
wingetcreate submit .\ALTRun-main\packaging\winget
```
(不需要安装 git: 直接下载仓库的 zip, 只用到里面 `packaging\winget` 的 4 个文件。)

第一次运行会打开浏览器要求登录 GitHub 并授权 wingetcreate; 它会 fork winget-pkgs 并创建 PR。PR 里的自动检查 (清单验证、下载、杀毒扫描、安装测试) 全部通过后, 由微软的维护者合并。合并以后:
- `winget install zhugecaomao.ALTRun` 就可以安装
- 以后的版本: 设置 `WINGET_TOKEN` 让 Release 工作流自动提交; 或者手动运行
  `wingetcreate update zhugecaomao.ALTRun --version <版本> --urls <zip 下载地址> --submit`

## 设置保存在哪里
ALTRun 是便携软件, 设置和数据 (`Data\`, 设置文件是 `Data\ALTRun.json`)、`Themes\` 都在程序目录里。
- **Scoop**: 程序在 `scoop\apps\altrun\current`。`Data` (包括设置) 和 `Themes` 由 Scoop 保存在 `scoop\persist\altrun` (目录联接)。2026.09.26 及更早的版本把 `ALTRun.json` 放在程序目录, 安装时从 persist 复制进来、卸载前复制回去; 现在的清单在安装时把 persist 里旧的 `ALTRun.json` 移到 `persist\altrun\Data\`。升级、卸载后重新安装, 设置都会保留。`scoop update altrun` 之前请先退出 ALTRun
- **winget**: 程序在 `%LOCALAPPDATA%\Microsoft\WinGet\Packages\zhugecaomao.ALTRun_...`。升级只替换 zip 里的文件, 设置保留; 卸载时留下 `Data` (包括设置)、`Themes`, 重新安装后接着使用

两种方式都在 Windows 上测试过: 安装旧版本 → 修改设置 → 升级 → 设置、Data、Themes 都在。
