# SPF2M 参考数据生成工具

`Tests/Fixtures/SPF2M-Reference.json` 是 PT 工具箱里束线型计算 (`Src/Extensions/TendonProfile.ahk`) 的对照数据：把 139 组参数输入公司原来的 SPF2M.EXE（1993 年的 16 位 DOS 程序），读出它算出来的结果。单元测试 `TendonProfileVsSpf2m` 用 ALTRun 算同样的参数，逐个数值比对。

这里的脚本就是用来生成这份数据的。只有要增加对照用例时才需要运行，平时开发和发布都用不到。

## 做法
SPF2M 只能在 DOS 下运行，也没有文件输出，所以：
1. 在 Linux 上用 DOSBox 运行 SPF2M，窗口放在虚拟显示器 (Xvfb) 上
2. 用 xdotool 模拟键盘，一个个回答 SPF2M 的问题（线型、钢绞线、半径、标高、跨度、反弯点）
3. 每一步截图，按 8×16 的字符格把屏幕还原成文字（先让 DOSBox 显示 `CHARS.TXT` 里的全部字符，学会每个字符的点阵）
4. 从最后一屏的表格里读出各支架处的高度，汇总成 JSON

| 文件 | 作用 |
|---|---|
| `run.sh` | 一条命令完成上面全部步骤 |
| `screen.py` | 截图 → 文字 (`learn` 学习字符点阵, `decode` 识别) |
| `drive.py` | 往 SPF2M 里输入用例, 记录每一屏; SPF2M 崩溃退回 DOS 时记下来并重新启动 |
| `build_fixture.py` | 记录的屏幕 → `SPF2M-Reference.json` |
| `cases.json` | 138 组输入: 4 种线型 × 6 种钢绞线, 上升 / 下降, 非整米跨度, 自定义半径 / 反弯点, 以及 SPF2M 会报错或崩溃的输入 |
| `manual.json` | 手动输入的 1 组: 自定义支架间距 (回答 "Change Support Interval" 为 Y 后逐段输入) |
| `CHARS.TXT` | 全部可打印 ASCII 字符, 用来学习 DOSBox 的字体 |

## 运行
需要 Linux（或 WSL）和 SPF2M.EXE。SPF2M.EXE 是公司内部程序，不在仓库里。

```bash
sudo apt install dosbox xvfb xdotool imagemagick
pip install pillow
Tests/Tools/SPF2M/run.sh /path/to/SPF2M.EXE
```

大约 15 分钟，结束后直接更新 `Tests/Fixtures/SPF2M-Reference.json`，再运行单元测试比对。

只试几组时，给出自己的用例文件和输出文件（不会覆盖正式的对照数据）：
```bash
Tests/Tools/SPF2M/run.sh /path/to/SPF2M.EXE my-cases.json /tmp/result.json
```

用例格式：`{"profile": 1, "tendon": 3, "start": 600, "end": 100, "dist": 9000}`，可选 `"radius"`、`"contra"`（不写 = 直接回车，用 SPF2M 的默认值）。
- profile: 1 双抛物线 / 2 抛物线-直线-抛物线 / 3 抛物线-直线 / 4 直线-抛物线
- tendon: 1 Slab / 2 7S / 3 12S / 4 19S / 5 22S / 6 31S

## 核对过
用这套脚本重新运行 SPF2M，得到的结果和仓库里的对照数据逐项一致（包括报错和崩溃的情况）。
