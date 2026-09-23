# 安全策略

## 支持的版本
只有 [最新发布的版本](https://github.com/zhugecaomao/ALTRun/releases/latest) 会收到安全修复。

## 报告安全问题
如果你发现了安全问题 (例如可以让其它程序借助 ALTRun 执行命令、泄露剪贴板历史等), **请不要公开提交 Issue**, 而是通过 GitHub 的
[私下报告漏洞](https://github.com/zhugecaomao/ALTRun/security/advisories/new) 提交, 并尽量附上:

- ALTRun 版本和 Windows 版本
- 复现步骤
- 可能的影响

收到后会尽快确认, 修复后在 Release 说明里致谢 (如果你希望署名)。

## 隐私说明
ALTRun 完全在本机运行, 不收集、不上传任何数据。唯一的联网操作是启动时检查 GitHub 上的新版本 (偏好设置 → 通用 → 启动时检查更新, 可以关闭)。
剪贴板历史保存在程序目录的 `Data\ClipboardHistory.json` (可以设为只保存在内存里), 密码管理器复制的内容不会记录。
