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

1. 在 https://signpath.org/apply 申请 (要求: 开源许可、项目在维护、源码和构建都在 GitHub 上)。申请前在 README 里加上下面的 "代码签名策略" 一节 (已于 2026-10 通过)
2. 通过后会收到两封邮件: 新 OSS 组织的邀请, 以及 CI 用户邮箱的确认。顺序是: 先注册 SignPath 账号 → 用它接受组织邀请 → 再确认 CI 用户邮箱 (没接受邀请之前确认不了)
3. 在 SignPath 里 (https://app.signpath.io):
   - **Projects → Add**: 名称和 slug 都写 `ALTRun` (工作流里的 `project-slug`), Repository URL 写 `https://github.com/zhugecaomao/ALTRun`
   - **Artifact configuration**: 粘贴 [`signpath/artifact-configuration.xml`](signpath/artifact-configuration.xml) 并设为默认。GitHub Actions 上传的产物总是 zip, 所以配置是 "zip 里的 ALTRun.exe"
   - **Trusted build systems**: 给项目关联 `GitHub.com`
   - **Signing policies**: 先建 `test-signing` (测试证书, CI 用户可以提交, 不需要批准); 正式证书导入后再建 `release-signing` (正式证书, 需要批准人在网页上批准)
   - **CI 用户 → API token**: 生成一个 token, CI 用户要是这两个签名策略的 Submitter
   - 组织 ID 在 Organization settings 里
4. 本仓库 Settings → Secrets and variables → Actions:
   - Secrets 新建 `SIGNPATH_API_TOKEN` (上面 CI 用户的 API token)
   - Variables 新建 `SIGNPATH_ORGANIZATION_ID`
5. 之后运行 Release `publish = false`, 编译出的 ALTRun.exe 会先送到 SignPath 用测试证书签名 (自签名证书, Windows 显示为不受信任, 只用来验证流程), 再做升级测试和打包
6. 测试签名成功后告诉 SignPath Foundation, 他们检查设置后订购正式证书并导入组织。建好 `release-signing` 后, Variables 再新建 `SIGNPATH_RELEASE_SIGNING` = `true`: 之后 `publish = true` 用正式证书签名, 需要在 SignPath 网页上批准 (30 分钟内)。在这之前正式发布照常发布未签名的版本

README 的 "代码签名策略 / Code signing policy" 一节和官网下载区 (https://zhugecaomao.github.io/ALTRun/#download) 的说明已经写好 (SignPath Foundation 要求下载页写明使用他们的代码签名), 申请表里的 Download URL 填官网的这个地址。

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
