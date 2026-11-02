# RunAhk.ps1 - runs an AutoHotkey script in CI with a time limit (tests.yml, release.yml, RunShots.ps1)
#
# A script that hangs, most often on a dialog nobody can close (an error AutoHotkey shows in a window instead of printing it,
# such as one thrown while a class is set up), would keep the job waiting until GitHub stops it after 6 hours. After
# -TimeoutSeconds this prints the class, title and text of the script's windows (the dialog says what went wrong), stops the
# script and exits with 124. Otherwise it exits with the script's exit code. The script's output goes to -Out / -Err.
#
#   & Tests/Tools/RunAhk.ps1 -Ahk tools/ahk/AutoHotkey64.exe -Arguments '/ErrorStdOut', 'Tests\RunTests.ahk' -Out test.log -Err test.err -TimeoutSeconds 600
#   $code = $LASTEXITCODE
# -NoAnnotation: no error annotation on timeout (tests.yml checks the time limit itself, on purpose)
param(
    [Parameter(Mandatory)][string]$Ahk,
    [Parameter(Mandatory)][string[]]$Arguments,
    [string]$Out = 'ahk.log',
    [string]$Err = 'ahk.err',
    [int]$TimeoutSeconds = 300,
    [switch]$NoAnnotation
)

$proc = Start-Process $Ahk -ArgumentList $Arguments -PassThru -NoNewWindow -RedirectStandardOutput $Out -RedirectStandardError $Err
$null = $proc.Handle                                                    # keep the handle: without it ExitCode can be empty once the process has ended
if ($proc.WaitForExit($TimeoutSeconds * 1000)) {
    $proc.WaitForExit()                                                 # the redirected output is complete after this
    exit $proc.ExitCode
}

Write-Host "$(if ($NoAnnotation) { '' } else { '::error::' })$($Arguments -join ' ') didn't finish in $TimeoutSeconds s. Its windows:"
# A second AutoHotkey reads the windows of the hung one; the script comes from standard input ("*")
$dump = "for w in WinGetList('ahk_pid $($proc.Id)')`n    FileAppend('[' WinGetClass(w) '] ' WinGetTitle(w) '``n' WinGetText(w) '``n', '*')"
$dump | & $Ahk /ErrorStdOut * | ForEach-Object { Write-Host $_ }
Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
exit 124
