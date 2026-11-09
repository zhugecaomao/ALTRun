# Submits a released version to winget-pkgs (a pull request from the token owner's fork), with wingetcreate.
# Runs in the Release workflow after publishing, and in the "Submit to winget" workflow to submit an existing release again:
#   $env:WINGET_TOKEN = '<token with public_repo>'; pwsh packaging/Submit-Winget.ps1 -Version 2026.11.03
#
# wingetcreate opens the pull request from the token owner's fork of winget-pkgs and fails when it can't bring that fork up to
# date ("The forked repository could not be synced with the upstream commits"): winget-pkgs gets hundreds of commits a day, so
# a fork left alone falls far behind. The fork is synced first with GitHub's merge-upstream API; when that fails (the fork's
# master has commits of its own) the log says so and wingetcreate tries anyway. A missing fork is fine: wingetcreate creates it.
#
# GitHub's release download links (github.com/<repo>/releases/download/...) can answer slowly and then fail with 504 Gateway
# Timeout (seen on 2026-11-09 during a GitHub incident: 9 ~ 11 s for each redirect, some ending in 504), and wingetcreate downloads
# the package only once. A failed download is tried again a few times, waiting longer each time: wingetcreate reads the package
# before it opens anything on GitHub, so trying again can't open a second pull request. Other failures are not repeated.
param(
    [Parameter(Mandatory)][string]$Version,
    [string]$Repository = 'zhugecaomao/ALTRun',
    [string]$Package = 'zhugecaomao.ALTRun'
)
$ErrorActionPreference = 'Stop'
$token = $env:WINGET_TOKEN
if (-not $token) { throw 'WINGET_TOKEN is not set' }
$headers = @{ Authorization = "Bearer $token"; Accept = 'application/vnd.github+json'; 'X-GitHub-Api-Version' = '2022-11-28' }

# 1. Bring the fork up to date
$owner = (Invoke-RestMethod https://api.github.com/user -Headers $headers).login
try {
    $fork = Invoke-RestMethod "https://api.github.com/repos/$owner/winget-pkgs" -Headers $headers
} catch {
    $fork = $null
}
if (-not $fork) {
    Write-Host "No fork $owner/winget-pkgs yet: wingetcreate creates it"
} elseif (-not $fork.fork -or $fork.parent.full_name -ne 'microsoft/winget-pkgs') {
    Write-Host "::warning::$owner/winget-pkgs isn't a fork of microsoft/winget-pkgs; not syncing it"
} else {
    $branch = $fork.parent.default_branch
    try {
        $result = Invoke-RestMethod "https://api.github.com/repos/$owner/winget-pkgs/merge-upstream" -Method Post -Headers $headers `
            -Body (@{ branch = $branch } | ConvertTo-Json) -ContentType 'application/json'
        Write-Host "Fork $owner/winget-pkgs ($branch): $($result.message)"
    } catch {
        $status = $_.Exception.Response.StatusCode.value__
        Write-Host "::warning::Couldn't sync the fork $owner/winget-pkgs (HTTP $status): sync it on GitHub (Sync fork → Discard commits) or delete it, then run the ""Submit to winget"" workflow"
    }
}

# 2. Submit the new version
$url = "https://github.com/$Repository/releases/download/$Version/ALTRun_v$Version.zip"
Invoke-WebRequest https://aka.ms/wingetcreate/latest -OutFile wingetcreate.exe -MaximumRetryCount 4 -RetryIntervalSec 15   # also a GitHub release download
$waits = 30, 60, 120, 240                                                   # seconds before each new attempt
for ($attempt = 1; ; $attempt++) {
    $ErrorActionPreference = 'Continue'                                     # its error output is only shown, not a reason to stop
    .\wingetcreate.exe update $Package --version $Version --urls $url --submit --token $token 2>&1 | Tee-Object -Variable output | Out-Host
    $code = $LASTEXITCODE
    $ErrorActionPreference = 'Stop'
    if ($code -eq 0) { break }
    $downloadFailed = ($output | Out-String) -match 'Failed to download'
    if (-not $downloadFailed -or $attempt -gt $waits.Count) { throw "wingetcreate failed (exit code $code)" }
    $wait = $waits[$attempt - 1]
    Write-Host "::warning::Downloading $url failed (attempt $attempt); trying again in $wait s"
    Start-Sleep -Seconds $wait
}
