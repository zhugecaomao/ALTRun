# 把 Scoop (bucket\altrun.json) 和 winget (packaging\winget\*.yaml) 清单更新到新版本。
# Release 工作流发布之后自动运行; 也可以手动运行:
#   pwsh packaging/Update-Manifests.ps1 -Version 2026.09.25 -Zip build/ALTRun_v2026.09.25.zip
param(
    [Parameter(Mandatory)][string]$Version,
    [Parameter(Mandatory)][string]$Zip,
    [string]$ReleaseDate = (Get-Date -Format 'yyyy-MM-dd')
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$hash = (Get-FileHash $Zip -Algorithm SHA256).Hash
$url = "https://github.com/zhugecaomao/ALTRun/releases/download/$Version/ALTRun_v$Version.zip"
$utf8 = New-Object System.Text.UTF8Encoding $false

function Update-File($path, [scriptblock]$change) {
    $text = [IO.File]::ReadAllText($path)
    $new = & $change $text
    if ($new -eq $text) { Write-Host "unchanged: $path" } else { [IO.File]::WriteAllText($path, $new, $utf8); Write-Host "updated:   $path" }
}

Update-File (Join-Path $root 'bucket/altrun.json') {
    param($t)
    $t = $t -replace '(?m)^(\s*"version":\s*")[^"]*(")', "`${1}$Version`${2}"
    $t = $t -replace '("url":\s*")https://github\.com/zhugecaomao/ALTRun/releases/download/[0-9.]+/ALTRun_v[0-9.]+\.zip(")', "`${1}$url`${2}"
    $t -replace '("hash":\s*")[0-9a-fA-F]{64}(")', "`${1}$($hash.ToLower())`${2}"
}

foreach ($file in Get-ChildItem (Join-Path $root 'packaging/winget') -Filter *.yaml) {
    Update-File $file.FullName {
        param($t)
        $t = $t -replace '(?m)^PackageVersion: .*$', "PackageVersion: $Version"
        $t = $t -replace '(?m)^ReleaseDate: .*$', "ReleaseDate: $ReleaseDate"
        $t = $t -replace '(?m)^(\s*InstallerUrl: ).*$', "`${1}$url"
        $t = $t -replace '(?m)^(\s*InstallerSha256: ).*$', "`${1}$hash"
        $t -replace '(?m)^(ReleaseNotesUrl: https://github\.com/zhugecaomao/ALTRun/releases/tag/).*$', "`${1}$Version"
    }
}
