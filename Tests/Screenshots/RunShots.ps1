# RunShots.ps1 - 运行 TakeScreenshots.ahk, 打印它的输出; 有截图失败时以失败结束 (screenshots.yml 用)
#   -Only normal-scale   只截这些场景
#   -Skip normal-scale   除了这些场景 (传给 TakeScreenshots.ahk 时写成 "-场景名")
#
#   pwsh Tests/Screenshots/RunShots.ps1 -Ahk tools/ahk/AutoHotkey64.exe [-Only 场景,...] [-Skip 场景,...]
param([Parameter(Mandatory)][string]$Ahk, [string[]]$Only = @(), [string[]]$Skip = @())

Write-Host "Microsoft YaHei: $(Test-Path C:\Windows\Fonts\msyh.ttc)"
Add-Type -AssemblyName System.Windows.Forms
Write-Host "Screen: $([System.Windows.Forms.Screen]::PrimaryScreen.Bounds)"
$outDir = Join-Path (Get-Location) 'docs\images\screenshots'                     # 绝对路径: 相对路径会按脚本所在的文件夹算
$arguments = @('/ErrorStdOut', 'Tests\Screenshots\TakeScreenshots.ahk', $outDir) + $Only + ($Skip | ForEach-Object { "-$_" })
$proc = Start-Process $Ahk -ArgumentList $arguments -Wait -PassThru -NoNewWindow -RedirectStandardOutput shots.log -RedirectStandardError shots.err
Get-Content shots.log, shots.err
if ($proc.ExitCode -ne 0) { throw "$($proc.ExitCode) screenshot(s) failed" }
