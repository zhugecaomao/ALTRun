;===============================================================================
; TendonProfile.ahk - post-tensioning tendon profile calculation with results identical to SPF2M (AutoHotkey v2)
;-------------------------------------------------------------------------------
; SPF2M.EXE is the company's in-house 16-bit DOS program written in Turbo Basic in 1993, which used to need DOSBox
; to run (no longer shipped with the program). This calculates directly with the same formulas: they come from PT-Profile-VLX, an AutoCAD tool of the same series
; (POBLIC.lsp), checked one by one against 139 result sets produced by SPF2M itself (Tests\Fixtures\SPF2M-Reference.json,
; over 1,100 values, all identical).
;
; Coordinates: distances are measured from the high end (SPF2M's tables always start at the high point), levels are the tendon height (mm).
;   1 double parabola              a parabola on the low side + a reverse parabola on the high side, tangent at the contraflexure point
;   2 parabola-straight-parabola   parabolas at both ends, straight in the middle
;   3 parabola-straight            a parabola on the high side, straight elsewhere
;   4 straight-parabola            straight on the high side, a parabola on the low side
; The contraflexure distance is measured from the high point; by default it follows from the minimum radius of curvature, and when given the radius is computed from it.
; Support spacing at most 1000 mm; when the span isn't whole meters, the remainder goes on the high side: over 599 it is a segment of its own, otherwise it is
; merged with one 1000 and split into two segments (rounded to 100). The segments can also be given by hand.
; The ends show the input levels; points in between: Actual truncated, Beam rounded to 5 mm, Slab to 10 mm.
;
; Usage:
;   result := TendonProfile.Calc(Map("Profile", 1, "Tendon", 3, "Start", 600, "End", 100, "Distance", 8000))
;   optional: "Radius" minimum radius of curvature (default by strand type), "Contraflexure" contraflexure distance, "Intervals" an array of support spacings
;   result.Rows  [[distance, spacing, Actual, Beam, Slab], ...]    result.Error  "" or an error description
;   result.MinRadius / result.Radius (the radius actually used) / result.Contraflexure
;===============================================================================

class TendonProfile {
    static Profiles := ["Double Parabolic", "Parabolic-Straight-Parabolic", "Parabolic-Straight", "Straight-Parabolic"]
    static Tendons  := ["Slab", "7S", "12S", "19S", "22S", "31S"]
    static Radii    := [5000, 3200, 4200, 5300, 5700, 6700]                 ; minimum radius of curvature per strand type (mm), as in SPF2M
    static DuctDiameters := [25, 70, 90, 100, 120, 130]                     ; duct diameter (mm); "At C.G." subtracts half of it from the levels
    static MaxInterval := 1000

    static DefaultRadius(tendon) {
        return (IsInteger(tendon) && tendon >= 1 && tendon <= TendonProfile.Radii.Length) ? TendonProfile.Radii[tendon] : TendonProfile.Radii[1]
    }

    static DefaultDuctDiameter(tendon) {
        return (IsInteger(tendon) && tendon >= 1 && tendon <= TendonProfile.DuctDiameters.Length) ? TendonProfile.DuctDiameters[tendon] : TendonProfile.DuctDiameters[1]
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

        ; contraflexure point (from the high point) and the radius actually used
        rad := minRadius
        contra := TendonProfile.ContraflexureFor(profile, rad, ly, l)
        wanted := get("Contraflexure")
        if IsNumber(wanted) && (wanted + 0 != Integer(contra)) {             ; the same as the default = unchanged
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

        rows := [[0, 0, hi, hi, hi]]                                        ; the ends show the input levels
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

    ; The default contraflexure distance (from the high point); "" when the minimum radius can't be met
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

    ; The radius of curvature for a given contraflexure distance (the inverse of ContraflexureFor)
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

    ; Tendon height at s from the high point
    static Height(profile, rad, lo, hi, l, s) {
        ly := hi - lo
        switch profile {
            case 1:                                                         ; double parabola: t measured from the low point
                t := l - s
                x := l - 2 * rad * ly / l
                if (t <= x)
                    return lo + (l - x) / (2 * rad * x) * t * t
                return lo + l * t / rad + ly - l * l / (2 * rad) - t * t / (2 * rad)
            case 2:                                                         ; parabola-straight-parabola
                x1 := l / 2 - Sqrt(l * l - 4 * rad * ly) / 2
                if (s < x1)
                    return hi - s * s / (2 * rad)
                if (s > l - x1)
                    return lo + (s - l) ** 2 / (2 * rad)
                return hi + x1 * x1 / (2 * rad) - x1 * s / rad
            case 3:                                                         ; parabola-straight
                x := l - Sqrt(l * l - 2 * rad * ly)
                if (s <= x)
                    return hi - s * s / (2 * rad)
                return lo + l * x / rad - x * s / rad
            default:                                                        ; straight-parabola
                x1 := Sqrt(l * l - 2 * rad * ly)
                if (s < x1)
                    return hi + s * (x1 - l) / rad
                return lo + (s - l) ** 2 / (2 * rad)
        }
    }

    ; Support spacings of at most 1000 mm, with the remainder on the high side (as in SPF2M)
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
        } else {                                                            ; the remainder is too short: merge it with one 1000 and split into two segments
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

    ; Text like "500, 1500, 800" -> an array; "" when empty (= automatic)
    static ParseIntervals(text) {
        text := Trim(text)
        if (text = "")
            return ""
        list := []
        for part in StrSplit(RegExReplace(text, "[\s,;/]+", ","), ",") {
            if (part = "")
                continue
            if !IsNumber(part)
                return part                                                 ; not numbers: returned as they are, and _CheckIntervals reports the error
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
