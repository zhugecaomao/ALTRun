# @altrun.title          IP Addresses
# @altrun.mode           output
# @altrun.description    The IPv4 addresses of this PC's network adapters, shown in your text editor
# @altrun.description_zh 本机各网卡的 IPv4 地址, 在文本编辑器里显示
# @altrun.network        none
#
# Runs in the background, then opens everything it prints in your text editor. Saved with a byte order mark (UTF-8 BOM),
# so Windows PowerShell 5 reads the Chinese line above correctly.
Get-NetIPAddress -AddressFamily IPv4 | Where-Object IPAddress -ne '127.0.0.1' | Format-Table InterfaceAlias, IPAddress -AutoSize
