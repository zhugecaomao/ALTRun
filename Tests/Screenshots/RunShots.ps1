# RunShots.ps1 - runs TakeScreenshots.ahk and prints its output; fails when any screenshot failed (used by screenshots.yml)
#   -Only normal-scale   only these scenes
#   -Skip normal-scale   all but these scenes (passed to TakeScreenshots.ahk as "-scene")
#
#   pwsh Tests/Screenshots/RunShots.ps1 -Ahk tools/ahk/AutoHotkey64.exe [-Only scene,...] [-Skip scene,...]
param([Parameter(Mandatory)][string]$Ahk, [string[]]$Only = @(), [string[]]$Skip = @())

Write-Host "Microsoft YaHei: $(Test-Path C:\Windows\Fonts\msyh.ttc)"
Add-Type -AssemblyName System.Windows.Forms
Write-Host "Screen: $([System.Windows.Forms.Screen]::PrimaryScreen.Bounds)"
$outDir = Join-Path (Get-Location) 'docs\images\screenshots'                     # absolute path: a relative one would be resolved from the script's folder
$arguments = @('/ErrorStdOut', 'Tests\Screenshots\TakeScreenshots.ahk', $outDir) + $Only + ($Skip | ForEach-Object { "-$_" })
$proc = Start-Process $Ahk -ArgumentList $arguments -Wait -PassThru -NoNewWindow -RedirectStandardOutput shots.log -RedirectStandardError shots.err
Get-Content shots.log, shots.err
if ($proc.ExitCode -ne 0) { throw "$($proc.ExitCode) screenshot(s) failed" }
