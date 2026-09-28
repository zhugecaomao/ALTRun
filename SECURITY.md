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
ALTRun 完全在本机运行, 不收集、不上传任何数据, 没有统计或遥测。

联网只有两种情况, 都只访问 GitHub:
- **检查更新**: 每天在后台一次 (偏好设置 → 通用 → "自动检查更新", 可以关闭) 或手动检查时, 读取 GitHub 上最新 Release 的版本号
- **一键更新**: 你选择 "立即更新" 后, 从 GitHub Releases 下载新版本, 先核对 Release 给出的 SHA256, 一致才替换程序文件; 任何一步出错, 原来的程序保持不变

网页搜索、"查看更新内容" 等是用你的浏览器打开网页, ALTRun 本身不发送请求。Everything 联动通过本机的 IPC 接口, 不经过网络。

本地保存的数据都在程序目录的 `Data\` 文件夹里:
- 剪贴板历史: `Data\ClipboardHistory.json` (可以设为只保存在内存里); 密码管理器 (KeePass、1Password、Bitwarden...) 复制的内容和带 "不要加入剪贴板历史" 标记的内容不会记录
- 使用统计: `Data\Usage.json`, 只记录每个功能用了几次, 不记录输入的文字和打开的内容
- 学习排序: `Data\Knowledge.json`, 可以在 偏好设置 → 高级 里重置
