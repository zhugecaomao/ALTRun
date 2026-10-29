[English](PT-Tools) · **中文**

# PT 工具箱

两个预应力 / 结构设计用的小工具: **PT Tools** (钢筋和网片面积、所需钢筋根数、计算器) 和 **SPF2M** (预应力束线型高度)。在搜索窗口输入 `PTTools` 或 `SPF2M` 打开, 也可以设成 [自定义热键](Hotkeys-zh) (动作 `PTTools` / `SPF2M`)。两个窗口都总在最前面, 记住窗口位置, 关闭时保存输入。

两个窗口里按 `Enter` 跳到下一个输入框 (和 `Tab` 一样), 按 `F1` 打开这个页面。

## PT Tools

![PT Tools](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/pttools.png)

所有输入框都是下拉组合框: 可以从列表里选常用值, 也可以自己输入。第一次点击输入框会全选里面的文字, 直接输入新值就行; 再点一下放光标。边输入边出结果。

| 分组 | 输入 | 结果 |
|---|---|---|
| **Rebar area** (钢筋面积) | 跨度、钢筋直径、钢筋间距 (mm) | 根数 = 跨度 ÷ 间距 (两位小数, 表示这个宽度上的根数, 不取整); 钢筋面积 = 根数 × πd²/4 (mm²) |
| **Bars required** (所需根数) | 面积算式 (数字或算式, 例如 `3500*1.2`)、钢筋直径; 可选的安全系数 (勾选后生效) | 总面积 = 算式的值; 根数 = 总面积 × 安全系数 ÷ 单根钢筋面积, 向上取整 |
| **Mesh area** (网片面积) | 跨度 (第一组的跨度改动时跟着变, 也可以单独修改)、上层和下层网片型号 (A / B / D / E + 钢丝直径, 例如 `A7`) | 这个跨度上的网片面积: A 型钢丝间距 200 mm, B 和 D 型 100 mm, E 型 150 mm; "+ bar" = 网片面积 + 第一组的钢筋面积 |
| **Calculator** (计算器) | 算式: 数字、`+ - * /`、`^` (乘方) 和括号; 列表里有示例 | 结果, 两位小数 |

列表里的常用值: 跨度 3000 / 5700 / 6000 / 7200; 钢筋直径 10 ~ 40; 钢筋间距 100 ~ 1000; 面积 3500 ~ 12000; 网片 A6 ~ E12; 安全系数 1.1 / 1.2 / 1.25 / 1.5。

只保存输入值 ([ALTRun.json](Configuration-zh) 里的 `Extensions.PTTools`), 计算结果每次打开窗口时重新算。

## SPF2M

![SPF2M](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/spf2m.png)

SPF2M 预应力束线型计算 (SPF2M Post-Tensioning Tendon Profile Calculator), 直接在 ALTRun 里计算, 不再需要 DOSBox。选择线型和钢绞线类型, 输入标高和水平距离, 右边的表格立即给出每个支架处的高度。

| 输入 | 说明 |
|---|---|
| Profile type (线型) | 双抛物线、抛物线-直线-抛物线、抛物线-直线、直线-抛物线 |
| Tendon type (钢绞线类型) | Slab、7S、12S、19S、22S、31S |
| Start level / end level (起点 / 终点标高, mm) | **At C.G.**: 标高给的是束中心时勾选, 自动减去半个管道直径 |
| Horizontal distance (水平距离, mm) | 从高点算起 |
| Min. radius of curvature (最小曲率半径, mm) | 可选 |
| Point of contraflexure (反弯点, mm) | 可选: 指定后反算曲率半径, 小于最小半径时提示 |
| Support intervals (支架间距, mm) | 可选: 默认最大 1000 mm, 零头放在高点一侧; 也可以写 `500, 1500, 800 ...`, 加起来要等于水平距离 |
| Duct dia. for At C.G. (管道直径, mm) | 可选; 填了之后每种钢绞线各自记住 |

可选项留空时用 SPF2M 的默认值 (灰色显示):

| 钢绞线类型 | Slab | 7S | 12S | 19S | 22S | 31S |
|---|---|---|---|---|---|---|
| 最小曲率半径 (mm) | 5000 | 3200 | 4200 | 5300 | 5700 | 6700 |
| 管道直径 (mm) | 25 | 70 | 90 | 100 | 120 | 130 |

**结果**: 表格从高点开始 (和 SPF2M 一样): 距离、间距、实际高度, 以及梁 (按 5 mm) 和板 (按 10 mm) 取整的高度; 下面是曲率半径和到反弯点的距离。**Copy Table** 把这些复制成表格, 直接粘贴到 Excel。

- 结果和原来的 SPF2M.EXE 完全一致: 在 DOSBox 里用原程序跑了 139 组 (4 种线型 × 6 种钢绞线, 上升 / 下降, 非整米跨度, 自定义半径 / 反弯点 / 支架间距), 1100 多个数值逐个比对, 作为单元测试的一部分
- SPF2M 在某些无法实现的线型上会直接退出, 这里会给出说明; 反弯点超过半跨时 SPF2M 什么也不显示, 这里也会提示
- 原来的 SPF2M.EXE 和 DOSBox 已经不再随程序发布。从旧版本升级后, `Resources\` 里的 `DOSBox.exe`、`SDL.dll`、`SDL_net.dll`、`SPF2M.exe`、`Run.bat` 可以删掉
