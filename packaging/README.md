# Scoop / winget / Microsoft Store 安装包

| 文件 | 说明 |
|---|---|
| `../bucket/altrun.json` | Scoop 清单。仓库本身就是一个 bucket: `scoop bucket add altrun https://github.com/zhugecaomao/ALTRun` |
| `winget/*.yaml` | winget 清单 (zip 里的便携版 ALTRun.exe)。2026-11 已收录到 [microsoft/winget-pkgs](https://github.com/microsoft/winget-pkgs): `winget install zhugecaomao.ALTRun` |
| `Update-Manifests.ps1` | 把两个清单更新到新版本 (版本号、下载地址、SHA256) |
| `Submit-Winget.ps1` | 向 winget-pkgs 提交新版本: 先把 token 所属账号 fork 的 winget-pkgs 同步到最新, 再用 wingetcreate 开 PR |

| `msix/AppxManifest.xml` | Microsoft Store 包 (MSIX) 的清单: 标识 (`zhugecaomao.ALTRun`、发布者) 来自 Partner Center 的 "产品标识" 页; 关掉了文件和注册表的写重定向 (设置照样放在 `%APPDATA%\ALTRun`, 延迟启动能改启动项), 带执行别名 `altrun` |
| `msix/Make-Msix.ps1` | 用发布目录打 MSIX 包 (makeappx): 版本号 `2026.11.03` → `2026.11.3.0`, 测试版 `2026.11.03.9` → `2026.11.3.9` |
| `msix/Assets/*.png` | 包里的图标, 由 `icon/make_store_assets.py` 生成 (和程序图标同一套图形) |

## 发布新版本时
Release 工作流 (publish = true) 创建 GitHub Release 之后会自动:
1. 运行 `Update-Manifests.ps1`, 把 Scoop / winget 清单更新到新版本并提交到 main。Scoop 用户 `scoop update altrun` 就能拿到新版本
2. 如果仓库设置了 `WINGET_TOKEN` secret, 运行 `Submit-Winget.ps1`: 先用 GitHub 的 merge-upstream 把你 fork 的 winget-pkgs 同步到最新 (winget-pkgs 每天有几百个提交, fork 放着不管会落后很多, wingetcreate 自己同步不了时会失败, 2026.11.03 就遇到过); fork 的 master 上有自己的提交、合并不了时 (2026.11.03 和 2026.11.09 都遇到过), 把它的 master 直接设成上游的最新提交, 和网页上的 Sync fork → Discard commits 一样 (这个 fork 只用来开 winget 的 PR, 每个 PR 用自己的分支, master 上的东西都不需要)。再用 [wingetcreate](https://github.com/microsoft/winget-create) 提交新版本的 PR (微软自动检查后合并, 一般 1~3 天); 下载安装包失败时 (GitHub 偶尔超时返回 504) 隔 30 / 60 / 120 / 240 秒再试, 最多 4 次。这一步失败不影响发布
3. 提交失败时 (日志里会写原因): 在 Actions 里运行 **Submit to winget** 工作流重新提交 (版本号留空 = 最新的 Release)。如果日志说 fork 同步不了 (自动丢弃 fork 的提交也失败了), 先在 GitHub 上打开你的 winget-pkgs fork → Sync fork → Discard commits, 或者直接删掉这个 fork (下次会重新创建), 再运行

`WINGET_TOKEN`: GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic), 勾选 `public_repo`; 然后在本仓库 Settings → Secrets and variables → Actions 里新建 `WINGET_TOKEN`。没有设置时跳过这一步, winget 用户就收不到新版本, 需要每次手动运行 `wingetcreate update` (见下文)。

## Microsoft Store (MSIX)
正式版的版本号 (`YYYY.MM.DD`, 三段: Release PR 上的构建和 publish = true) 运行 Release 工作流时会生成 MSIX 包, 放在运行页面底部的 Artifacts 里 (`ALTRun_v<版本>_msix`); 测试版 (四段版本号) 不打包, 构建快一些。包没有签名, 发布到 Store 时由 Store 签名; 所以不能直接双击安装。工作流会用临时证书签一份测试用的副本, 装上后通过 `altrun` 别名启动 ALTRun, 检查它在包里运行、设置写在 `%APPDATA%\ALTRun`、开机启动的快捷方式指向别名, 再跑 Windows App Certification Kit (可选项只报告, 必需项失败就算失败)。

提交新版本 (正式版发布后): Partner Center → 应用和游戏 → ALTRun → 新建提交 → 程序包, 上传正式版那次运行的 `.msix`, 保存后提交认证。只有正式版 (版本号最后一位是 0) 能提交。

商店版和其它版本的区别 (见 `App.PackageFamily`): 由 Store 更新 ("检查更新" 打开 Store 页面, 不在后台检查); 开机启动、"发送到" 和 CapsLock 辅助的计划任务用 `altrun` 别名启动 (包的安装目录每次更新都会变); 不另建开始菜单快捷方式。

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

## 第一次提交到 winget (已完成)
2026-11 已经完成, winget-pkgs 收录了 2026.10.26。下面的步骤留作记录 (在 Windows 上):

```powershell
winget install Microsoft.WingetCreate      # 装好后关闭并重新打开 PowerShell
cd $env:TEMP
Invoke-WebRequest https://github.com/zhugecaomao/ALTRun/archive/refs/heads/main.zip -OutFile ALTRun-main.zip
Expand-Archive ALTRun-main.zip -DestinationPath . -Force
wingetcreate submit .\ALTRun-main\packaging\winget
```
(不需要安装 git: 直接下载仓库的 zip, 只用到里面 `packaging\winget` 的 4 个文件。)

第一次运行会打开浏览器要求登录 GitHub 并授权 wingetcreate; 它会 fork winget-pkgs 并创建 PR。PR 里的自动检查 (清单验证、下载、杀毒扫描、安装测试) 全部通过后, 由微软的维护者合并。

以后的版本: 设置 `WINGET_TOKEN` 让 Release 工作流自动提交; 或者手动运行
  `wingetcreate update zhugecaomao.ALTRun --version <版本> --urls <zip 下载地址> --submit`

## 设置保存在哪里
ALTRun 是便携软件, 设置和数据 (`Data\`, 设置文件是 `Data\ALTRun.json`)、`Themes\` 都在程序目录里。
- **Scoop**: 程序在 `scoop\apps\altrun\current`。`Data` (包括设置) 和 `Themes` 由 Scoop 保存在 `scoop\persist\altrun` (目录联接)。2026.09.26 及更早的版本把 `ALTRun.json` 放在程序目录: 清单在安装时把它从 persist 复制进来、卸载前复制回去 (ALTRun 保存设置时替换整个文件, 不能用硬链接)。persist 的 `Data\` 里已经有 `ALTRun.json` 后就不再复制; 新版本第一次启动时自己把程序目录里的 `ALTRun.json` 移到 `Data\`, 所以同一份清单对新旧版本都适用。升级、卸载后重新安装, 设置都会保留。`scoop update altrun` 之前请先退出 ALTRun
- **winget**: 程序在 `%LOCALAPPDATA%\Microsoft\WinGet\Packages\zhugecaomao.ALTRun_...`。升级只替换 zip 里的文件, 设置保留; 卸载时留下 `Data` (包括设置)、`Themes`, 重新安装后接着使用

两种方式都在 Windows 上测试过: 安装旧版本 → 修改设置 → 升级 → 设置、Data、Themes 都在。
