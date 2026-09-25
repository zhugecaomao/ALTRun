;===============================================================================
; TendonProfile.ahk - 后张预应力束线型计算, 和 SPF2M 的结果完全一致 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; SPF2M.EXE 是公司内部 1993 年用 Turbo Basic 写的 16 位 DOS 程序, 以前要借助 DOSBox
; 运行。这里用同样的公式直接计算: 公式来自同一系列的 AutoCAD 工具 PT-Profile-VLX
; (POBLIC.lsp), 再用 SPF2M 本身跑出来的 139 组结果逐个核对 (Tests\Data\SPF2M-Reference.json,
; 1100 多个数值全部一致)。
;
; 坐标: 距离从高的一端量起 (SPF2M 的表格总是从高点开始), 标高是束的高度 (mm)。
;   1 双抛物线                  低点一侧抛物线 + 高点一侧反向抛物线, 在反弯点相切
;   2 抛物线-直线-抛物线        两端抛物线, 中间直线
;   3 抛物线-直线               高点一侧抛物线, 其余直线
;   4 直线-抛物线               高点一侧直线, 低点一侧抛物线
; 反弯点距离 (Contraflexure) 从高点量起; 默认由最小曲率半径算出, 手动指定时反过来算半径。
; 支架间距最大 1000mm; 跨度不是整米时, 零头放在高点一侧: 大于 599 单独一段, 否则和一个
; 1000 合并后分成两段 (按 100 取整)。也可以手动给出每一段。
; 两端显示输入的标高; 中间各点: Actual 截断取整, Beam 按 5mm、Slab 按 10mm 四舍五入。
;
; 用法:
;   result := TendonProfile.Calc(Map("Profile", 1, "Tendon", 3, "Start", 600, "End", 100, "Distance", 8000))
;   可选: "Radius" 最小曲率半径 (默认按钢绞线类型), "Contraflexure" 反弯点距离, "Intervals" 支架间距数组
;   result.Rows  [[距离, 间距, Actual, Beam, Slab], ...]    result.Error  "" 或错误说明
;   result.MinRadius / result.Radius (实际使用的半径) / result.Contraflexure
;===============================================================================

class TendonProfile {
    static Profiles := ["Double Parabolic", "Parabolic-Straight-Parabolic", "Parabolic-Straight", "Straight-Parabolic"]
    static Tendons  := ["Slab", "7S", "12S", "19S", "22S", "31S"]
    static Radii    := [5000, 3200, 4200, 5300, 5700, 6700]                 ; 各钢绞线类型的最小曲率半径 (mm), 和 SPF2M 一致
    static MaxInterval := 1000

    static DefaultRadius(tendon) {
        return (IsInteger(tendon) && tendon >= 1 && tendon <= TendonProfile.Radii.Length) ? TendonProfile.Radii[tendon] : TendonProfile.Radii[1]
    }

    static Calc(input) {
        result := {Rows: [], Error: "", MinRadius: 0, Radius: 0, Contraflexure: 0, RadiusChanged: false}
        get := (key) => (input.Has(key) && input[key] != "") ? input[key] : ""
        profile := get("Profile"), tendon := get("Tendon")
        start := get("Start"), finish := get("End"), l := get("Distance")
        if !(IsInteger(profile) && profile >= 1 && profile <= 4)
            return TendonProfile._Fail(result, "Choose a profile type.")
        if !(IsNumber(start) && IsNumber(finish))
            return TendonProfile._Fail(result, "Enter the start and end levels.")
        if !(IsNumber(l) && l > 0)
            return TendonProfile._Fail(result, "Enter the horizontal distance.")
        minRadius := IsNumber(get("Radius")) ? get("Radius") + 0 : TendonProfile.DefaultRadius(tendon)
        if (minRadius <= 0)
            return TendonProfile._Fail(result, "The radius of curvature must be greater than 0.")
        lo := Min(start, finish) + 0, hi := Max(start, finish) + 0, ly := hi - lo, l += 0
        result.MinRadius := minRadius

        ; 反弯点 (从高点量起) 和实际使用的半径
        rad := minRadius
        contra := TendonProfile.ContraflexureFor(profile, rad, ly, l)
        wanted := get("Contraflexure")
        if IsNumber(wanted) && (wanted + 0 != Integer(contra)) {             ; 和默认值相同 = 不改
            wanted += 0
            if (ly = 0)
                return TendonProfile._Fail(result, "Start and end levels are equal, the profile is straight - leave the point of contraflexure empty.")
            if (wanted <= 0 || !TendonProfile._ContraflexureFits(profile, wanted, l))
                return TendonProfile._Fail(result, "The point of contraflexure must lie within the span"
                    . ((profile <= 2) ? " (at most half of it)." : "."))
            rad := TendonProfile.RadiusFor(profile, wanted, ly, l)
            contra := wanted
            result.RadiusChanged := true
        }
        if (contra = "" || rad < minRadius || (profile = 1 && contra > l / 2))
            return TendonProfile._Fail(result, "MIN. RADIUS OF CURVATURE NOT ACHIEVABLE - increase the distance or reduce the level difference.")
        result.Radius := rad
        result.Contraflexure := contra

        intervals := get("Intervals")
        if (intervals = "")
            intervals := TendonProfile.DefaultIntervals(l)
        else if (msg := TendonProfile._CheckIntervals(intervals, l))
            return TendonProfile._Fail(result, msg)

        rows := [[0, 0, hi, hi, hi]]                                        ; 两端显示输入的标高
        s := 0
        for index, interval in intervals {
            s += interval
            if (index = intervals.Length) {
                rows.Push([l, interval, lo, lo, lo])
                break
            }
            y := TendonProfile.Height(profile, rad, lo, hi, l, s)
            rows.Push([s, interval, Integer(y), Integer((y + 2.5) / 5) * 5, Integer((y + 5) / 10) * 10])
        }
        result.Rows := rows
        return result
    }

    ; 默认的反弯点距离 (从高点量起); 最小半径做不到时返回 ""
    static ContraflexureFor(profile, rad, ly, l) {
        switch profile {
            case 1: return 2 * rad * ly / l
            case 2:
                disc := l * l - 4 * rad * ly
                return (disc < 0) ? "" : (l - Sqrt(disc)) / 2
            default:
                disc := l * l - 2 * rad * ly
                return (disc < 0) ? "" : l - Sqrt(disc)
        }
    }

    ; 指定反弯点距离时的曲率半径 (ContraflexureFor 的反算)
    static RadiusFor(profile, contra, ly, l) {
        switch profile {
            case 1: return contra * l / (2 * ly)
            case 2: return contra * (l - contra) / ly
            default: return contra * (2 * l - contra) / (2 * ly)
        }
    }

    static _ContraflexureFits(profile, contra, l) {
        return (profile <= 2) ? contra <= l / 2 : contra < l
    }

    ; 从高点量起 s 处的束高度
    static Height(profile, rad, lo, hi, l, s) {
        ly := hi - lo
        switch profile {
            case 1:                                                         ; 双抛物线: t 从低点量起
                t := l - s
                x := l - 2 * rad * ly / l
                if (t <= x)
                    return lo + (l - x) / (2 * rad * x) * t * t
                return lo + l * t / rad + ly - l * l / (2 * rad) - t * t / (2 * rad)
            case 2:                                                         ; 抛物线-直线-抛物线
                x1 := l / 2 - Sqrt(l * l - 4 * rad * ly) / 2
                if (s < x1)
                    return hi - s * s / (2 * rad)
                if (s > l - x1)
                    return lo + (s - l) ** 2 / (2 * rad)
                return hi + x1 * x1 / (2 * rad) - x1 * s / rad
            case 3:                                                         ; 抛物线-直线
                x := l - Sqrt(l * l - 2 * rad * ly)
                if (s <= x)
                    return hi - s * s / (2 * rad)
                return lo + l * x / rad - x * s / rad
            default:                                                        ; 直线-抛物线
                x1 := Sqrt(l * l - 2 * rad * ly)
                if (s < x1)
                    return hi + s * (x1 - l) / rad
                return lo + (s - l) ** 2 / (2 * rad)
        }
    }

    ; 最大 1000mm 的支架间距, 零头在高点一侧 (和 SPF2M 一样)
    static DefaultIntervals(l) {
        step := TendonProfile.MaxInterval
        whole := Integer(l / step)
        rest := l - whole * step
        intervals := []
        if (rest = 0) {
            Loop whole
                intervals.Push(step)
        } else if (rest > 599 || whole = 0) {
            intervals.Push(TendonProfile._Clean(rest))
            Loop whole
                intervals.Push(step)
        } else {                                                            ; 零头太短: 和一个 1000 合并后分成两段
            total := rest + step
            second := Integer(total / 200) * 100
            first := total - second
            if (Abs(second - first) >= 100) {
                second := (Integer(total / 200) + 1) * 100
                first := total - second
            }
            intervals.Push(TendonProfile._Clean(first), second)
            Loop whole - 1
                intervals.Push(step)
        }
        return intervals
    }

    ; "500, 1500, 800" 这样的文字 -> 数组; 空的返回 "" (= 自动)
    static ParseIntervals(text) {
        text := Trim(text)
        if (text = "")
            return ""
        list := []
        for part in StrSplit(RegExReplace(text, "[\s,;/]+", ","), ",") {
            if (part = "")
                continue
            if !IsNumber(part)
                return part                                                 ; 不是数字: 原样返回, _CheckIntervals 报错
            list.Push(part + 0)
        }
        return list
    }

    static _CheckIntervals(intervals, l) {
        if !(intervals is Array)
            return "Support intervals: '" intervals "' is not a number."
        total := 0
        for interval in intervals {
            if !(IsNumber(interval) && interval > 0)
                return "Support intervals must be greater than 0."
            total += interval
        }
        if (Abs(total - l) > 0.5)
            return "Support intervals add up to " TendonProfile._Clean(total) " mm, not the horizontal distance " TendonProfile._Clean(l) " mm."
        return ""
    }

    static _Clean(v) {
        return (v = Integer(v)) ? Integer(v) : Round(v, 1)
    }

    static _Fail(result, message) {
        result.Error := message
        return result
    }
}
