<#
.SYNOPSIS
    Builds the Microsoft Store package (MSIX) of ALTRun from the release folder.

.DESCRIPTION
    Copies the release folder (ALTRun.exe, Resources, ...) and packaging\msix\Assets into a package folder, writes AppxManifest.xml
    with the version, and packs it with makeappx.exe from the Windows SDK. The package isn't signed: the Store signs it when it is
    published (for a test install, sign a copy with a test certificate whose subject is the manifest's Publisher).

    The MSIX version has four numbers: 2026.11.03 -> 2026.11.3.0 (a release; the Store wants the last number to be 0),
    2026.11.03.9 -> 2026.11.3.9 (a test build, never sent to the Store).

.EXAMPLE
    ./packaging/msix/Make-Msix.ps1 -Source build/ALTRun -Version 2026.11.03 -Out build/ALTRun_v2026.11.03_x64.msix
#>
param(
    [Parameter(Mandatory)] [string] $Source,
    [Parameter(Mandatory)] [string] $Version,
    [Parameter(Mandatory)] [string] $Out
)
$ErrorActionPreference = 'Stop'

function ConvertTo-MsixVersion([string] $version) {
    if ($version -notmatch '^(\d{4})\.(\d{2})\.(\d{2})(?:\.(\d+))?$') { throw "Unexpected version $version (YYYY.MM.DD or YYYY.MM.DD.N)" }
    $build = if ($Matches[4]) { [int]$Matches[4] } else { 0 }
    return "{0}.{1}.{2}.{3}" -f [int]$Matches[1], [int]$Matches[2], [int]$Matches[3], $build
}

$here = $PSScriptRoot
$msixVersion = ConvertTo-MsixVersion $Version
$layout = Join-Path ([System.IO.Path]::GetTempPath()) "ALTRun-msix-$([guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory $layout | Out-Null
try {
    Copy-Item (Join-Path $Source '*') $layout -Recurse
    Copy-Item (Join-Path $here 'Assets') (Join-Path $layout 'Assets') -Recurse
    $manifest = (Get-Content (Join-Path $here 'AppxManifest.xml') -Raw -Encoding UTF8).Replace('{VERSION}', $msixVersion)
    [System.IO.File]::WriteAllText((Join-Path $layout 'AppxManifest.xml'), $manifest, (New-Object System.Text.UTF8Encoding($false)))

    $makeappx = Get-ChildItem "${env:ProgramFiles(x86)}\Windows Kits\10\bin\*\x64\makeappx.exe" -ErrorAction SilentlyContinue |
                Sort-Object { [version]$_.Directory.Parent.Name } -Descending | Select-Object -First 1
    if (-not $makeappx) { throw "makeappx.exe (Windows SDK) not found" }
    $outFull = [System.IO.Path]::GetFullPath($Out)
    New-Item -ItemType Directory (Split-Path $outFull) -Force | Out-Null
    & $makeappx.FullName pack /d $layout /p $outFull /o /h SHA256 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "makeappx pack failed (exit code $LASTEXITCODE)" }
    Write-Host "MSIX $msixVersion -> $outFull ($([math]::Round((Get-Item $outFull).Length / 1KB)) KB)"
} finally {
    Remove-Item $layout -Recurse -Force -ErrorAction SilentlyContinue
}
