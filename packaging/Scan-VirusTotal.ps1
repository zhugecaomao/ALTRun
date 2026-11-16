# Scans a file with VirusTotal (API v3) before it is published: uploads it, waits for the analysis and reports how many engines
# flag it. Run by .github/workflows/release.yml before the GitHub Release is created (publish = true, or virustotal = true).
#
#   ./packaging/Scan-VirusTotal.ps1 -Path build/ALTRun/ALTRun.exe [-AllowDetections] [-NotesFile build/notes.md]
#
# The API key comes from the environment variable VT_API_KEY (the repository secret of the same name); without it the scan stops
# with an error, it is never skipped quietly. Exit code 0 = no engine reports the file as malicious (suspicious ones are only
# reported); otherwise it throws, so the release isn't created, unless -AllowDetections (the workflow input allow_detections,
# after checking the report: a false positive) lets it go on. Not finishing within -TimeoutMinutes counts as not clean too.
# Writes the result and the engines that flag the file to the job summary (GITHUB_STEP_SUMMARY), and with -NotesFile adds a line
# with the report link to the release notes.
# The public API allows 4 requests a minute: one upload, then one check every 20 s. Uploaded files are shared with VirusTotal's
# partners; the free API is for non-commercial use.
param(
    [Parameter(Mandatory)] [string] $Path,
    [switch] $AllowDetections,
    [string] $NotesFile = "",
    [int] $TimeoutMinutes = 15,
    [int] $PollSeconds = 20
)
$ErrorActionPreference = "Stop"

if (-not $env:VT_API_KEY) { throw "VT_API_KEY isn't set (repository secret): the file can't be scanned, so it isn't published" }
if (-not (Test-Path $Path)) { throw "$Path doesn't exist" }
$headers = @{ "x-apikey" = $env:VT_API_KEY }
$file = Get-Item $Path
$sha256 = (Get-FileHash $file.FullName -Algorithm SHA256).Hash.ToLower()
$report = "https://www.virustotal.com/gui/file/$sha256"
Write-Host "Uploading $($file.Name) ($($file.Length) bytes, SHA256 $sha256) to VirusTotal"

function Invoke-VT([string] $method, [string] $uri, $form = $null) {
    try {
        if ($form) { return Invoke-RestMethod -Method $method -Uri $uri -Headers $headers -Form $form }
        return Invoke-RestMethod -Method $method -Uri $uri -Headers $headers
    } catch {
        $status = $_.Exception.Response.StatusCode.value__
        throw "VirusTotal $method $uri failed (HTTP $status): $($_.ErrorDetails.Message)"   # 401: wrong key, 429: quota used up
    }
}

$upload = Invoke-VT Post "https://www.virustotal.com/api/v3/files" @{ file = $file }
$analysis = $upload.data.id
$deadline = (Get-Date).AddMinutes($TimeoutMinutes)
$status = ""
do {
    Start-Sleep -Seconds $PollSeconds
    $result = Invoke-VT Get "https://www.virustotal.com/api/v3/analyses/$analysis"
    $status = $result.data.attributes.status
    Write-Host "$(Get-Date -Format HH:mm:ss) analysis $status"
} while ($status -ne "completed" -and (Get-Date) -lt $deadline)

$summary = @("### VirusTotal: $($file.Name)", "", "SHA256 ``$sha256``, [report]($report)", "")
if ($status -ne "completed") {
    $summary += "The analysis didn't finish in $TimeoutMinutes minutes (last status: $status)."
    $clean = $false
    $line = "VirusTotal: the analysis didn't finish in $TimeoutMinutes minutes"
} else {
    $stats = $result.data.attributes.stats
    $engines = $stats.malicious + $stats.suspicious + $stats.undetected + $stats.harmless
    $flagged = foreach ($engine in $result.data.attributes.results.PSObject.Properties) {
        if ($engine.Value.category -in "malicious", "suspicious") {
            [pscustomobject]@{ Engine = $engine.Name; Category = $engine.Value.category; Result = $engine.Value.result }
        }
    }
    $line = "malicious $($stats.malicious), suspicious $($stats.suspicious), undetected $($stats.undetected), harmless $($stats.harmless) ($engines engines)"
    $summary += "**$line**"
    if ($flagged) {
        $summary += "", "| Engine | Category | Result |", "|---|---|---|"
        $summary += $flagged | Sort-Object Category, Engine | ForEach-Object { "| $($_.Engine) | $($_.Category) | $($_.Result) |" }
    }
    $clean = $stats.malicious -eq 0
    if ($NotesFile) {
        Add-Content $NotesFile -Encoding UTF8 -Value @("", "**VirusTotal** (ALTRun.exe): [$($stats.malicious) / $engines engines]($report)")
    }
}
Write-Host $line
Write-Host "Report: $report"
$flagged | Format-Table -AutoSize | Out-String | Write-Host
if ($env:GITHUB_STEP_SUMMARY) { $summary | Add-Content $env:GITHUB_STEP_SUMMARY -Encoding UTF8 }

if (-not $clean) {
    if ($AllowDetections) {
        Write-Host "::warning::Not clean on VirusTotal, published anyway (allow_detections): $report"
    } else {
        throw "Not clean on VirusTotal ($line), so the release isn't created. Check $report; if it is a false positive, run again with allow_detections"
    }
}
