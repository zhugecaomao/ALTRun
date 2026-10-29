# MakeDemos.ps1 - turns each folder of animation frames that TakeScreenshots.ahk saved (%TEMP%\ALTRunDemoFrames\<scene>) into
# docs\images\screenshots\<scene>.png with MakeDemo.py, then removes the folder (screenshots.yml runs it after the 150% and the 100% shots)
#
#   pwsh Tests/Screenshots/MakeDemos.ps1
$root = Join-Path $env:TEMP "ALTRunDemoFrames"
if (-not (Test-Path $root)) { Write-Host "No animation frames"; exit 0 }
foreach ($folder in Get-ChildItem $root -Directory) {
  if (-not (Test-Path (Join-Path $folder.FullName "frames.txt"))) { Write-Host "No frames.txt in $($folder.Name)"; continue }
  python Tests/Screenshots/MakeDemo.py $folder.FullName "docs/images/screenshots/$($folder.Name).png"
  if ($LASTEXITCODE -ne 0) { throw "MakeDemo.py failed for $($folder.Name)" }
  Remove-Item $folder.FullName -Recurse -Force
}
