# Contributing

[English](#contributing) · [中文](#参与贡献)

Thank you for wanting to improve ALTRun! Bug reports, documentation improvements, translations and code are all very welcome. Issues and pull requests are welcome in English or Chinese.

## Reporting problems
- **Bugs**: use the [bug report template](https://github.com/zhugecaomao/ALTRun/issues/new?template=bug_report.yml) and include your Windows version, ALTRun version and the steps to reproduce, ideally with a screenshot
- **Ideas**: use the [feature request template](https://github.com/zhugecaomao/ALTRun/issues/new?template=feature_request.yml), or start a discussion in [Discussions](https://github.com/zhugecaomao/ALTRun/discussions) first
- When more information is needed, turn on "Write a debug log" in Preferences → General, reproduce the problem, then type "log" in the search box to open the log file (`%Temp%\ALTRun.log`) and attach the relevant part (remove anything private first)
- Security problems: please report them privately, see [SECURITY.md](SECURITY.md)

## Pull requests
1. Create a branch from `main`; one pull request does one thing
2. Make sure all unit tests pass and the syntax check shows no warnings:
   ```
   AutoHotkey64.exe /ErrorStdOut Tests\RunTests.ahk
   AutoHotkey64.exe /ErrorStdOut /validate ALTRun.ahk
   ```
3. Include tests with new features and fixes (`Tests\RunTests.ahk`)
4. For UI changes, attach screenshots made with the English UI and made-up sample data; don't include anything personal or work-related
5. For user-visible changes, update the README (`README.md` and `README.zh-CN.md`), `CHANGELOG.md` and, where needed, the wiki (`docs/wiki/`, both the English page and the `-zh` Chinese page)

The project structure, development environment, code style, translations, documentation and release process are described in the **[Development guide](https://github.com/zhugecaomao/ALTRun/wiki/Development)** in the wiki.

By submitting code, you agree to release your contribution under the [GPL-3.0](LICENSE) license.

---

# 参与贡献

感谢你愿意改进 ALTRun! 反馈问题、改进文档、补充翻译、提交代码都非常欢迎。Issue 和 Pull Request 用中文或英文都可以。

## 反馈问题
- **Bug**: 用 [Bug 报告模板](https://github.com/zhugecaomao/ALTRun/issues/new?template=bug_report.yml), 写清 Windows 版本、ALTRun 版本、复现步骤, 最好附上截图
- **建议**: 用 [功能建议模板](https://github.com/zhugecaomao/ALTRun/issues/new?template=feature_request.yml), 或先在 [Discussions](https://github.com/zhugecaomao/ALTRun/discussions) 里讨论
- 需要更多信息时, 可以在偏好设置 → 通用里打开 "写入调试日志", 复现问题后在搜索框输入 "日志" (或 log) 打开日志文件 (`%Temp%\ALTRun.log`), 附上相关部分 (注意先去掉隐私信息)
- 安全问题请私下报告, 见 [SECURITY.md](SECURITY.md#安全策略)

## 提交 Pull Request
1. 从 `main` 新建分支, 一个 PR 只做一件事
2. 确认单元测试全部通过、语法检查没有警告:
   ```
   AutoHotkey64.exe /ErrorStdOut Tests\RunTests.ahk
   AutoHotkey64.exe /ErrorStdOut /validate ALTRun.ahk
   ```
3. 新功能或修复请附带测试 (`Tests\RunTests.ahk`)
4. 涉及界面的改动请附上截图, 用英文界面和虚构的示例数据, 不要放个人或工作相关的信息
5. 用户可见的变化请更新 README (`README.md` 和 `README.zh-CN.md`)、`CHANGELOG.md`, 必要时更新 Wiki (`docs/wiki/`, 英文页和 `-zh` 中文页都要改)

项目结构、开发环境、代码规范、界面翻译、文档和发布流程见 Wiki 的 **[开发指南](https://github.com/zhugecaomao/ALTRun/wiki/Development-zh)**。

提交代码即表示你同意以 [GPL-3.0](LICENSE) 许可发布你的贡献。
