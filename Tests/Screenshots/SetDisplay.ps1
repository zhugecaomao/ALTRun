# SetDisplay.ps1 - before the screenshots, gives the GitHub Actions Windows VM display settings better suited to screenshots
#   1. resolution 1920 x 1080 (default 1024 x 768, too small for 150% windows)
#   2. scaling 150%: the UI is drawn at 1.5x pixels, so screenshots shown smaller on web pages look sharper (especially on high-DPI screens)
#   3. at 150% fonts use grayscale smoothing, not ClearType: subpixel color fringes look bad in screenshots shown smaller;
#      at 100% (Preferences, Quick Switch) ClearType, to look like a normal 100% Windows (grayscale text looks bolder and blurrier)
# Scaling uses an undocumented Windows interface (SPI_SETLOGICALDPIOVERRIDE, which the "Scale" dropdown in Settings uses): the value is the number of steps relative
# to the "recommended scaling", which isn't known, so each step is tried and checked by reading A_ScreenDPI with a freshly started AutoHotkey.
# If it fails it stays as it is (screenshots as usual, just at 100%). The actual scale goes into SHOT_SCALE in $env:GITHUB_ENV, for RoundCorners.py / MakeDemo.py.
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

# 1. resolution
$dm = New-Object ShotDisplay+DEVMODE
$dm.dmSize = [System.Runtime.InteropServices.Marshal]::SizeOf($dm)
[void][ShotDisplay]::EnumDisplaySettings($null, -1, [ref]$dm)                   # ENUM_CURRENT_SETTINGS
Write-Host "Display: $($dm.dmPelsWidth) x $($dm.dmPelsHeight)"
$dm.dmPelsWidth = $Width; $dm.dmPelsHeight = $Height
$dm.dmFields = 0x80000 -bor 0x100000                                            # DM_PELSWIDTH | DM_PELSHEIGHT
$result = [ShotDisplay]::ChangeDisplaySettings([ref]$dm, 0)
Write-Host "ChangeDisplaySettings $Width x $Height -> $result (0 = OK)"

# 3. font smoothing (SPI_SETFONTSMOOTHING = 1; SPI_SETFONTSMOOTHINGTYPE: 1 = grayscale FE_FONTSMOOTHINGSTANDARD, 2 = ClearType)
$smoothing = if ($Scale -eq 100) { 2 } else { 1 }
[void][ShotDisplay]::SystemParametersInfo(0x004B, 1, [IntPtr]::Zero, 3)
[void][ShotDisplay]::SystemParametersInfo(0x200B, 0, [IntPtr]$smoothing, 3)
Write-Host "Font smoothing: $(if ($smoothing -eq 2) { 'ClearType' } else { 'grayscale' })"

# 2. scaling
$probe = Join-Path $env:TEMP "ShotDpi.ahk"
Set-Content $probe '#NoTrayIcon
FileAppend(A_ScreenDPI, "*")' -Encoding utf8
$out = Join-Path $env:TEMP "ShotDpi.txt"
function Get-Dpi {                                                                  # only newly started programs use the new scaling
    Start-Process $Ahk -ArgumentList '/ErrorStdOut', $probe -Wait -NoNewWindow -RedirectStandardOutput $out
    [int]((Get-Content $out -Raw) -replace '\D', '')
}
$target = [int][Math]::Round(96 * $Scale / 100)
$dpi = Get-Dpi
Write-Host "DPI before: $dpi (target $target)"
if ($dpi -ne $target) {
    foreach ($step in 0, 1, 2, 3, 4, -1, -2) {                                 # 0 = recommended scaling (usually what going back to 100% means)
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
