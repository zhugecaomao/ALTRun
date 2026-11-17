# @altrun.title          File Hash
# @altrun.keyword        hash
# @altrun.argument       file path
# @altrun.mode           output
# @altrun.selection      File
# @altrun.description    MD5, SHA-1 and SHA-256 of a file, shown in your text editor; also in the actions for a selected file (Ctrl+Alt+\)
# @altrun.description_zh 计算文件的 MD5、SHA-1 和 SHA-256, 在文本编辑器里显示; 选中文件按 Ctrl+Alt+\ 也能用
# @altrun.network        none
#
# Type "hash C:\path\to\file" in ALTRun, or select a file and press Ctrl+Alt+\. Saved with a byte order mark (UTF-8 BOM),
# so Windows PowerShell 5 reads the Chinese line above correctly.
$path = $args[0]
if (-not $path -or -not (Test-Path -LiteralPath $path -PathType Leaf)) {
    Write-Output "Not a file: $path"
    exit 1
}
'MD5', 'SHA1', 'SHA256' | ForEach-Object { Get-FileHash -Algorithm $_ -LiteralPath $path } | Format-List Algorithm, Hash, Path
