# RunShots.ps1 - runs TakeScreenshots.ahk and prints its output; fails when any screenshot failed (used by screenshots.yml)
#   -Only normal-scale   only these scenes
#   -Skip normal-scale   all but these scenes (passed to TakeScreenshots.ahk as "-scene")
#   -Within normal-scale only the scenes that are also in this list ("=scene"), e.g. -Only prefs-general,shellmenu -Within normal-scale
#
#   pwsh Tests/Screenshots/RunShots.ps1 -Ahk tools/ahk/AutoHotkey64.exe [-Only scene,...] [-Skip scene,...] [-Within scene,...]
param([Parameter(Mandatory)][string]$Ahk, [string[]]$Only = @(), [string[]]$Skip = @(), [string[]]$Within = @())

Write-Host "Microsoft YaHei: $(Test-Path C:\Windows\Fonts\msyh.ttc)"
Add-Type -AssemblyName System.Windows.Forms
Write-Host "Screen: $([System.Windows.Forms.Screen]::PrimaryScreen.Bounds)"
$outDir = Join-Path (Get-Location) 'docs\images\screenshots'                     # absolute path: a relative one would be resolved from the script's folder
$arguments = @('/ErrorStdOut', 'Tests\Screenshots\TakeScreenshots.ahk', $outDir) + $Only + ($Skip | ForEach-Object { "-$_" }) + ($Within | ForEach-Object { "=$_" })
& (Join-Path $PSScriptRoot '..\Tools\RunAhk.ps1') -Ahk $Ahk -Arguments $arguments -Out shots.log -Err shots.err -TimeoutSeconds 1500   # 25 minutes; a full run takes about 4
$code = $LASTEXITCODE
Get-Content shots.log, shots.err
if ($code -eq 124) { throw "TakeScreenshots.ahk didn't finish in 25 minutes" }
if ($code -ne 0) { throw "$code screenshot(s) failed" }
