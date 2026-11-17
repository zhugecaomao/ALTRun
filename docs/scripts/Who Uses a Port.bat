@echo off
rem @altrun.title          Who Uses a Port
rem @altrun.keyword        port
rem @altrun.argument       a port number
rem @altrun.description    Which programs use a TCP / UDP port (netstat, with process IDs), in a console window
rem @altrun.description_zh 查哪个程序占用了某个端口 (netstat, 带进程 ID), 在命令行窗口里显示
rem @altrun.network        none
rem
rem Type "port 8080" in ALTRun: the text after the keyword is passed to the script as its first argument.
rem No mode line: it runs normally in its own window. Saved without a byte order mark (cmd would take one as part of
rem the first command); the Chinese text is only on rem lines.
netstat -ano | findstr /c:":%~1 "
pause
