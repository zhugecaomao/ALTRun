[English](Window-Switching) · **中文**

# 窗口切换

输入窗口标题或程序名, `Enter` 切换到这个已经打开的窗口 (最小化的先还原), 和 PowerToys Run 的 Window Walker、Raycast 的 Switch Windows 一样。

- `w 关键字`: 只搜窗口, 例如 `w excel`、`w 报价`
- `w ` (`w` 加空格): 列出所有窗口, 按前后顺序 (最近用过的在前)
- 直接输入 (不加 `w`): 标题匹配得好的窗口也会显示, 排在应用和自定义命令后面, 最多 3 个
- 选中一个窗口按 `→`: 关闭这个窗口

![w 加空格列出打开的窗口](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/windows.png)

只列出任务栏上能看到的窗口: 不包括隐藏的窗口、工具窗口、别的虚拟桌面上的窗口和 ALTRun 自己的窗口。窗口不会被记进 [最近使用](Recent-zh)。

## 设置
偏好设置 → 功能: "窗口切换" 打开或关闭整个功能; 下面的 "功能选项" 里, "普通搜索里也显示" 和 "关键字" 对应:

| 设置 | 默认 | 说明 |
|---|---|---|
| `Features.Windows.Enabled` | `1` | 是否启用窗口切换 |
| `Features.Windows.Keyword` | `w` | 只搜窗口的关键字 |
| `Features.Windows.InDefaultResults` | `1` | 直接输入时也显示标题匹配得好的窗口 |
