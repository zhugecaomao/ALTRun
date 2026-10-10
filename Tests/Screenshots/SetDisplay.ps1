# SetDisplay.ps1 - 截图前把 GitHub Actions 的 Windows 虚拟机调成更适合截图的显示设置
#   1. 分辨率 1920 x 1080 (默认 1024 x 768, 放不下 150% 的窗口)
#   2. 缩放 150%: 界面按 1.5 倍像素画, 截图在网页上缩小显示时更清晰 (高分屏上尤其明显)
#   3. 字体用灰度抗锯齿, 不用 ClearType: 子像素的彩色毛边在缩放后的截图里不好看
# 缩放用的是 Windows 的未公开接口 (SPI_SETLOGICALDPIOVERRIDE, 设置里的 "缩放" 下拉框用的就是它): 值是相对
# "推荐缩放" 的档数, 不知道推荐的是哪一档, 所以逐档试, 每次用新启动的 AutoHotkey 读 A_ScreenDPI 确认。
# 做不到时保持原样 (截图照常, 只是 100%)。实际的缩放比例写进 $env:GITHUB_ENV 的 SHOT_SCALE, 给 RoundCorners.py / MakeDemo.py 用。
#
#   pwsh Tests/Screenshots/SetDisplay.ps1 -Ahk tools/ahk/AutoHotkey64.exe [-Width 1920] [-Height 1080] [-Scale 150]
param([Parameter(Mandatory)][string]$Ahk, [int]$Width = 1920, [int]$Height = 1080, [int]$Scale = 150)

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class ShotDisplay {
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct DEVMODE {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string dmDeviceName;
        public short dmSpecVersion, dmDriverVersion, dmSize, dmDriverExtra;
        public int dmFields, dmPositionX, dmPositionY, dmDisplayOrientation, dmDisplayFixedOutput;
        public short dmColor, dmDuplex, dmYResolution, dmTTOption, dmCollate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string dmFormName;
        public short dmLogPixels;
        public int dmBitsPerPel, dmPelsWidth, dmPelsHeight, dmDisplayFlags, dmDisplayFrequency;
        public int dmICMMethod, dmICMIntent, dmMediaType, dmDitherType, dmReserved1, dmReserved2, dmPanningWidth, dmPanningHeight;
    }
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern bool EnumDisplaySettings(string device, int mode, ref DEVMODE dm);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int ChangeDisplaySettings(ref DEVMODE dm, int flags);
    [DllImport("user32.dll")] public static extern bool SystemParametersInfo(uint action, uint param, IntPtr value, uint winIni);
}
"@

# 1. 分辨率
$dm = New-Object ShotDisplay+DEVMODE
$dm.dmSize = [System.Runtime.InteropServices.Marshal]::SizeOf($dm)
[void][ShotDisplay]::EnumDisplaySettings($null, -1, [ref]$dm)                   # ENUM_CURRENT_SETTINGS
Write-Host "Display: $($dm.dmPelsWidth) x $($dm.dmPelsHeight)"
$dm.dmPelsWidth = $Width; $dm.dmPelsHeight = $Height
$dm.dmFields = 0x80000 -bor 0x100000                                            # DM_PELSWIDTH | DM_PELSHEIGHT
$result = [ShotDisplay]::ChangeDisplaySettings([ref]$dm, 0)
Write-Host "ChangeDisplaySettings $Width x $Height -> $result (0 = OK)"

# 3. 灰度字体抗锯齿 (SPI_SETFONTSMOOTHING = 1, SPI_SETFONTSMOOTHINGTYPE = FE_FONTSMOOTHINGSTANDARD)
[void][ShotDisplay]::SystemParametersInfo(0x004B, 1, [IntPtr]::Zero, 3)
[void][ShotDisplay]::SystemParametersInfo(0x200B, 0, [IntPtr]1, 3)

# 2. 缩放
$probe = Join-Path $env:TEMP "ShotDpi.ahk"
Set-Content $probe '#NoTrayIcon
FileAppend(A_ScreenDPI, "*")' -Encoding utf8
$out = Join-Path $env:TEMP "ShotDpi.txt"
function Get-Dpi {                                                                  # 新启动的程序才会用新的缩放
    Start-Process $Ahk -ArgumentList '/ErrorStdOut', $probe -Wait -NoNewWindow -RedirectStandardOutput $out
    [int]((Get-Content $out -Raw) -replace '\D', '')
}
$target = [int][Math]::Round(96 * $Scale / 100)
$dpi = Get-Dpi
Write-Host "DPI before: $dpi (target $target)"
if ($dpi -ne $target) {
    foreach ($step in 1, 2, 3, 4, -1, -2, 0) {
        [void][ShotDisplay]::SystemParametersInfo(0x009F, [uint32]($step -band 0xFFFFFFFF), [IntPtr]::Zero, 1)   # SPI_SETLOGICALDPIOVERRIDE
        Start-Sleep -Seconds 2
        $dpi = Get-Dpi
        Write-Host "  step $step -> DPI $dpi"
        if ($dpi -eq $target) { break }
    }
}
$shotScale = [Math]::Round($dpi / 96, 2)
Write-Host "Screenshots will be taken at $([int]($shotScale * 100))%"
if ($env:GITHUB_ENV) { "SHOT_SCALE=$shotScale" | Out-File -FilePath $env:GITHUB_ENV -Append -Encoding utf8 }
