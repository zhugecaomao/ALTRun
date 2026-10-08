# 参与贡献

[English](CONTRIBUTING.md) · **中文**

感谢你愿意改进 ALTRun! 反馈问题、改进文档、补充翻译、提交代码都非常欢迎。提交之前请花几分钟看一下下面的约定。

## 反馈问题
- **Bug**: 用 [Bug 报告模板](https://github.com/zhugecaomao/ALTRun/issues/new?template=bug_report.yml), 写清 Windows 版本、ALTRun 版本、复现步骤, 最好附上截图
- **建议**: 用 [功能建议模板](https://github.com/zhugecaomao/ALTRun/issues/new?template=feature_request.yml), 或先在 [Discussions](https://github.com/zhugecaomao/ALTRun/discussions) 里讨论
- 需要更多信息时, 可以在偏好设置 → 通用里打开 "写入调试日志", 复现问题后在搜索框输入 "日志" (或 log) 打开日志文件 (`%Temp%\ALTRun.log`), 附上相关部分 (注意先去掉隐私信息)

## 开发环境
1. 安装 [AutoHotkey v2](https://www.autohotkey.com/) (2.0 或更新)
2. 克隆仓库, 直接运行 `ALTRun.ahk`
3. 修改后运行单元测试:
   ```
   AutoHotkey64.exe /ErrorStdOut Tests\RunTests.ahk
   ```
   退出码为失败的数量, 0 = 全部通过
4. 语法检查 (不运行):
   ```
   AutoHotkey64.exe /ErrorStdOut /validate ALTRun.ahk
   ```

## 代码规范
- **文件与类**: 一个文件一个类, 文件名 = 类名, PascalCase (`SearchWindow.ahk` → `class SearchWindow`)
- **文件夹**: `Lib\` 放和 ALTRun 无关、可以拿到别的项目用的通用库; ALTRun 自己的代码放 `Src\` 下对应的子文件夹
- **命名**: 类、方法、属性用 PascalCase; 局部变量用 camelCase; 内部使用的成员以 `_` 开头 (`_Layout()`、`_cache`)
- **避免重名**: AutoHotkey 的名字不区分大小写, 局部变量不要和类同名 (例如不要把变量叫 `pinyin`, 会遮住 `Pinyin` 类); 同一个类里的方法和属性也不要只差大小写
- **注释**: 每个文件开头说明用途和用法; 代码里的注释说明 "为什么", 用中文
- **单行输入框**: 所有 `Edit` 控件都要写明行数 (`r1 -Multi`), 否则长文字会让它自动变成多行 (有测试检查)
- **设置**: 新的设置项加到 `AppSettings.Defaults()`; 改变已有设置的结构时 `AppSettings.CurrentVersion + 1`, 并在 `SchemaMigration` 里加一个 `_FromN()`
- **界面文字**: 英文原文写在 `I18n.ahk`, 翻译写在 `Resources\Lang\` 的每个语言文件里 (`zh-CN.json`、`zh-TW.json`、`ja.json`; 测试会检查每个文件都有、占位符一致); 偏好设置里的文字较长时, 注意各种语言都不要超出页面 (`PreferencesFit` 测试会检查)
- **翻译 / 新语言**: 复制一个语言文件, 改名为语言代码 (例如 `ko.json`、`de.json`), 改 `_name` (用这种语言写的名称) 和 `_font`, 再翻译每一条; 放进 `Resources\Lang\` 后偏好设置的语言列表会自动列出它。只翻译了一部分也可以用, 没翻译的显示英文 (或 `_fallback` 指定的语言)
- **编码**: UTF-8; `ALTRun.ahk` 保持 UTF-8 BOM + CRLF, 其它文件 UTF-8 + LF

新增一个搜索功能的步骤见 Wiki 的 [开发指南](https://github.com/zhugecaomao/ALTRun/wiki/Development)。

## 提交 Pull Request
1. 从 `main` 新建分支, 一个 PR 只做一件事
2. 提交前确认单元测试全部通过、`/validate` 没有警告
3. 新功能或修复请附带测试 (`Tests\RunTests.ahk`)
4. 涉及界面的改动请附上截图; README / Wiki 里的截图由 Screenshots workflow 生成, 用英文界面和虚构的示例数据, 不要放个人或工作相关的信息
5. 用户可见的变化请更新 README (`README.md` 英文和 `README.zh-CN.md` 中文)、`CHANGELOG.md` 的 "未发布" 部分, 必要时更新 Wiki (中文页和 `en-` 英文页都要改)

## 发布新版本
版本号用发布日期 `YYYY.MM.DD`:
1. 更新 `Src\Core\App.ahk` 的 `App.Version` 和 `ALTRun.ahk` 的 `;@Ahk2Exe-SetVersion` (有测试检查两者一致)
2. 在 `CHANGELOG.md` 把 "未发布" 改成 `## [YYYY.MM.DD]`, 这一节就是 Release 的说明; 文件末尾加上这个版本的链接, 并把 `[未发布]` 的比较链接改成从这个版本开始 (有测试检查)
3. 合并到 `main` 后, 在 Actions 里运行 **Release** (`.github/workflows/release.yml`): 先 `publish = false` 检查构建、测试和升级测试, 再 `publish = true` 创建 Release

Release workflow 在 Windows 上编译 `ALTRun.exe`, 并用 `Tests\Fixtures\ALTRun.v2026.08.12.ini` 验证从旧版本升级。

## 文档 (Wiki)
Wiki 的源文件在仓库的 [`docs/wiki/`](docs/wiki), 合并到 `main` 后由 GitHub Actions (`.github/workflows/wiki.yml`) 自动发布到 [Wiki](https://github.com/zhugecaomao/ALTRun/wiki)。请修改 `docs/wiki/` 里的文件, 直接在网页上修改的 Wiki 会在下次发布时被覆盖。页面之间的链接写页面名, 不带 `.md` (例如 `[主题](Themes)`; 英文页写 `[Themes](en-Themes)`)。每页有中文 (`Themes.md`) 和英文 (`en-Themes.md`) 两个文件, 改一种语言时请一起改另一种; 英文不熟悉的话可以只改中文, 在 PR 里说明, 维护者会补上。

提交代码即表示你同意以 [GPL-3.0](LICENSE) 许可发布你的贡献。
