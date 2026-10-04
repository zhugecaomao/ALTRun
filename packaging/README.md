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

## 代码签名 (SignPath)
ALTRun.exe 没有签名时, Windows SmartScreen 会提示 "无法识别的应用 / 未知发布者", 部分杀毒软件也更容易误报。
[SignPath Foundation](https://signpath.org/) 免费为开源项目提供代码签名证书, Release 工作流已经准备好, 设置后自动签名:

1. 在 https://signpath.org/apply 申请 (要求: 开源许可、项目在维护、源码和构建都在 GitHub 上)。申请前在 README 里加上下面的 "代码签名策略" 一节
2. 通过后在 SignPath 里建项目 `ALTRun`, 签名策略 `test-signing` 和 `release-signing`, 产物配置选 "单个 PE 文件" (ALTRun.exe), 可信构建系统选 GitHub.com 并关联本仓库
3. 本仓库 Settings → Secrets and variables → Actions:
   - Secrets 新建 `SIGNPATH_API_TOKEN` (SignPath 里给 CI 用户生成的 API token)
   - Variables 新建 `SIGNPATH_ORGANIZATION_ID`
4. 之后运行 Release 时, 编译出的 ALTRun.exe 会先送到 SignPath 签名, 再做升级测试和打包。`publish = true` 用正式证书, 需要在 SignPath 网页上批准 (30 分钟内)

README 里要加的代码签名策略 (SignPath Foundation 的要求):

```markdown
## 代码签名策略 / Code signing policy
Free code signing provided by [SignPath.io](https://about.signpath.io/), certificate by [SignPath Foundation](https://signpath.org/).
- Committers and reviewers: [zhugecaomao](https://github.com/zhugecaomao)
- Approvers: [zhugecaomao](https://github.com/zhugecaomao)

隐私: ALTRun 不收集任何数据, 见 [SECURITY.md](SECURITY.md#隐私说明)。This program will not transfer any information to other networked systems unless specifically requested by the user (update checks and exchange rates, see SECURITY.md).
```

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
- **Scoop**: 程序在 `scoop\apps\altrun\current`。`Data` (包括设置) 和 `Themes` 由 Scoop 保存在 `scoop\persist\altrun` (目录联接)。2026.09.26 及更早的版本把 `ALTRun.json` 放在程序目录: 清单在安装时把它从 persist 复制进来、卸载前复制回去 (ALTRun 保存设置时替换整个文件, 不能用硬链接)。persist 的 `Data\` 里已经有 `ALTRun.json` 后就不再复制; 新版本第一次启动时自己把程序目录里的 `ALTRun.json` 移到 `Data\`, 所以同一份清单对新旧版本都适用。升级、卸载后重新安装, 设置都会保留。`scoop update altrun` 之前请先退出 ALTRun
- **winget**: 程序在 `%LOCALAPPDATA%\Microsoft\WinGet\Packages\zhugecaomao.ALTRun_...`。升级只替换 zip 里的文件, 设置保留; 卸载时留下 `Data` (包括设置)、`Themes`, 重新安装后接着使用

两种方式都在 Windows 上测试过: 安装旧版本 → 修改设置 → 升级 → 设置、Data、Themes 都在。
