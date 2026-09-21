;===============================================================================
; Language.ahk - 界面文字表 (AutoHotkey v2)
;-------------------------------------------------------------------------------
; 原来的 SetLanguage()/ReadChineseFlag() 从 ALTRun.ahk 搬到这里, 逻辑完全没变,
; 只是包了一层 class, 方便以后单独维护翻译, 不用在 3000+ 行的主文件里翻找。
;
; 用法 (ALTRun.ahk 里):
;   Global g_LNG := Lang.Load()   ; 启动时调用一次, 返回当前语言对应的 Map
;
; 编号约定 (沿用原来的分区, 新增文字请按区间续写, 不要打乱已有编号 - 现有的
; g_LNG[123] 之类的调用点到处都是, 编号一旦挪动就要满仓库搜索改调用点):
;   1~9     保留
;   10~49   主界面
;   50~99   状态栏提示 / Tips
;   100~149 选项窗口 - 常规/界面/热键/索引等复选框文字
;   150~159 选项窗口 - 文件管理器/Everything/历史记录长度
;   160~169 选项窗口 - 索引
;   170~189 选项窗口 - 界面 (字体/颜色/背景/透明度)
;   190~209 选项窗口 - 热键
;   210~249 选项窗口 - Listary (目录快速切换)
;   250~299 选项窗口 - Plugins (自动日期等)
;   300+    托盘菜单
;   400+    列表右键菜单
;   500+    状态统计窗口
;   600+    关于窗口
;   700+    命令管理器窗口
;   800+    提示/确认/错误消息
;===============================================================================

Class Lang {

    ; 构建中英文两张文字表, 并按 Lang.IsChinese() 的结果返回其中一张。
    ; 依赖 g_JSON (ALTRun.ahk 顶部已定义的全局路径常量) 和 g_LOG。
    static Load() {
        ENG     := Map()
        CHN     := Map()

        ENG[1]  := "English"                                                ; 1~9 Reserved
        ENG[2]  := "Options"
        ENG[7]  := "OK"
        ENG[8]  := "Cancel"
        ENG[9]  := "Help"
        ENG[10] := ["No.", "Type", "Command", "Description"]                ; 10~49 Main GUI
        ENG[11] := "Run"
        ENG[12] := "Options"
        ENG[13] := "Type anything here to search..."
        ENG[50] := ["Tip | F1 | Help & About", "Tip | F2 | Options and settings", "Tip | F3 | Edit current command", "Tip | F4 | Change setting file directly", "Tip | Alt+Space / Alt+R | Activate ALTRun", "Tip | Alt+Space / Esc / Lose Focus | Deactivate ALTRun", "Tip | Enter / Alt+No. | Run selected command", "Tip | Arrow Up or Down | Previous / next command", "Tip | Ctrl+D | Locate cmd's dir with File Manager"]
        ENG[51] := "Tip: "
        ENG[52] := "Activate ALTRun with a hotkey (Alt+Space)"              ; 50~99 Tips
        ENG[53] := "Smart Rank - Automatically adjusts command priority based on usage frequency."
        ENG[54] := "Arrow Up / Down = Select the previous / next command"
        ENG[55] := "Esc = Clear the input / close the window"
        ENG[56] := "Enter = Run the current command"
        ENG[57] := "Alt+Number = Run a specific command"
        ENG[58] := "Start with + = Create a new command"
        ENG[59] := "F3 = Edit current command"
        ENG[60] := "F2 = Open Options"
        ENG[61] := "Reindex = Rebuild the file index"
        ENG[62] := "F1 = ALTRun Help & About"
        ENG[63] := "Alt+Space = Show / hide the window"
        ENG[64] := "Ctrl+Q = Reload ALTRun"
        ENG[65] := "Ctrl+Number = Select a specific command"
        ENG[66] := "Alt+F4 = Close the ALTRun window"
        ENG[67] := "Ctrl+D = Open the current command's folder in File Manager"
        ENG[68] := "F4 = Edit the settings file directly (.ini)"
        ENG[69] := "Start with a space = Search files with Everything"
        ENG[70] := "Ctrl+'+' = Increase rank of current command"
        ENG[71] := "Ctrl+'-' = Decrease rank of current command"
        ENG[100] := ["General", "GUI", "Hotkey", "Index", "Listary", "Plugins", "Usage", "About"] ; 100~149 Options window (General - Check Listview)
        ENG[101] := "Launch on Windows startup"
        ENG[102] := "Enable SendTo - Create commands conveniently using Windows SendTo"
        ENG[103] := "Enable ALTRun shortcut in the Windows Start menu"
        ENG[104] := "Show tray icon in the system taskbar"
        ENG[105] := "Close window on losing focus"
        ENG[106] := "Always stay on top"
        ENG[107] := "Show the main window title bar"
        ENG[108] := "Use Windows Theme instead of Classic Theme"
        ENG[109] := "Press [ESC] to clear input, press again to close window (Untick: Close directly)"
        ENG[110] := "Keep last input and matching result on close"
        ENG[111] := "Show command icon in the result list"
        ENG[112] := "SendToGetLnk - Retrieve .lnk target on SendTo"
        ENG[113] := "Save commands execution history"
        ENG[114] := "Save application log"
        ENG[115] := "Match full path on search"
        ENG[116] := "Show Command List Grid Lines"
        ENG[117] := "Show Command List Header"
        ENG[118] := "Show Command List Serial Number"
        ENG[119] := "Show Command List Border Line"
        ENG[120] := "Smart Sorting (Auto-adjust command priority based on usage habits)"
        ENG[121] := "Smart Matching"
        ENG[122] := "Match beginning of the string (Untick: Match from any position)"
        ENG[123] := "Show Status Bar Hints/Tips"
        ENG[124] := "Show Command Executed RunCount in the Status Bar"
        ENG[125] := "Show Status Bar"
        ENG[126] := "Show [Run] Button on Main Window"
        ENG[127] := "Show [Options] Button on Main Window"
        ENG[128] := "Double-Buffering to reduce flicker (WinXP+)"
        ENG[129] := "Enable express structure calculation"
        ENG[130] := "Show shortened paths (show names instead of full paths)"
        ENG[131] := "Set language to Chinese Simplified (简体中文)"
        ENG[132] := "Match Chinese Pinyin first characters"
        ENG[133] := "Use the mouse wheel scroll to select the command"
        ENG[134] := "Click mouse middle button to run the selected command"
        ENG[135] := "Auto check for updates"
        ENG[136] := "Show large icons"
        ENG[137] := "Use rounded corners for Main Window"
        ENG[138] := "Press Space key to run command"
        ENG[139] := "Auto switch to English input method when activated"

        ENG[150] := "File Manager"                                          ; 150~159 Options window (Other than Check Listview)
        ENG[151] := "Everything"
        ENG[152] := "History length"
        ENG[160] := "Index"                                                 ; 160~169 Index
        ENG[161] := "Index location"
        ENG[162] := "Index file type"
        ENG[163] := "Index exclude"
        ENG[164] := "Index depth"
        ENG[165] := "Index Windows Store Apps"
        ENG[170] := "GUI"                                                   ; 170~189 GUI
        ENG[171] := "Search result number"
        ENG[172] := "Width of each column"
        ENG[173] := "Font (Main GUI)"
        ENG[174] := "Font (Options)"
        ENG[175] := "Font (Status Bar)"
        ENG[176] := "Window size (W x H)"
        ENG[177] := "Cmd list size (W x H)"
        ENG[178] := "Color (Command List)"
        ENG[179] := "Color (Main GUI)"
        ENG[180] := "Background picture"
        ENG[181] := "Transparency"
        ENG[182] := "Select font"
        ENG[183] := "Select color"
        ENG[184] := "Select image"
        ENG[190] := "Hotkey"                                                ; 190~209 Hotkey
        ENG[191] := "Activate Hotkey (Global)"
        ENG[192] := "Primary Hotkey"
        ENG[193] := "Secondary Hotkey"
        ENG[194] := "Two hotkeys can be set simultaneously"
        ENG[195] := "Reset hotkey"
        ENG[200] := "Actions and Hotkeys (Non-Global)"
        ENG[201] := "Hotkey"
        ENG[202] := "Trigger action"
        ENG[203] := "Hotkey"
        ENG[204] := "Hotkey 2"
        ENG[206] := "Hotkey 3"
        ENG[210] := "Listary"                                               ; 210~249 Listary
        ENG[211] := "Directory Quick Switch"
        ENG[212] := "File Manager ID"
        ENG[213] := "Open/Save Dialog ID"
        ENG[214] := "Exclude Windows ID"
        ENG[215] := "Hotkey"
        ENG[216] := "Jump to the Total Commander path"
        ENG[217] := "Jump to the Explorer path"
        ENG[218] := "Automatically switch paths in open/save dialogs"
        ENG[219] := "No Total Commander window found, please open Total Commander first!"
        ENG[220] := "No Explorer window found, please open Explorer first!"
        ENG[221] := "{1} → TC / {2} → Explorer"

        ENG[250] := "Plugins"                                               ; 250~299 Plugins
        ENG[251] := "Auto-date at end of text"
        ENG[252] := "Apply to window id"
        ENG[253] := "Hotkey"
        ENG[254] := "Date format"
        ENG[255] := "Auto-date before file extension"
        ENG[259] := "Conditional action"
        ENG[260] := "If window id/class contains"
        ENG[261] := "Hotkey (Editable)"
        ENG[262] := "Trigger action"
        ENG[300] := "Show"                                                  ; 300+ TrayMenu
        ENG[301] := "Options`tF2"
        ENG[302] := "ReIndex"
        ENG[303] := "Usage"
        ENG[304] := "About`tF1"
        ENG[305] := "Script Info"
        ENG[307] := "Reload"
        ENG[308] := "Exit"
        ENG[309] := "Update"
        ENG[310] := "Settings File`tF4"

        ENG[400] := "Run`tEnter"                                            ; 400+ LV_ContextMenu (Right-click)
        ENG[401] := "Locate`tCtrl+D"
        ENG[402] := "Copy`tCtrl+C"
        ENG[403] := "New`tCtrl+N"
        ENG[404] := "Edit`tF3"
        ENG[405] := "Delete`tCtrl+Del"
        ENG[406] := ""
        ENG[407] := "Copy statusbar text"
        ENG[408] := "Copied : "
        ENG[410] := "Show usage status"
        ENG[500] := "30 days ago"                                           ; 500+ Usage Status
        ENG[501] := "Now"
        ENG[502] := "Total number of times the command was executed"
        ENG[503] := "Number of times the program was activated today"
        ENG[600] := "About"                                                 ; 600+ About
        ENG[601] := "An open-source, lightweight, efficient and powerful launcher"
            . "`nIt provides a streamlined and efficient way to find anything on your system and launch any application in your way"
            . "`n`nSetting file:`n" g_JSON "`n`nProgram file:`n" A_ScriptFullPath
            . "`n`nCheck for Updates"
            . "`n<a href=`"https://github.com/zhugecaomao/ALTRun/releases`">https://github.com/zhugecaomao/ALTRun/releases</a>"
            . "`n`nSource code at GitHub"
            . "`n<a href=`"https://github.com/zhugecaomao/ALTRun`">https://github.com/zhugecaomao/ALTRun</a>"
            . "`n`nSee Help and Wiki page for more details"
            . "`n<a href=`"https://github.com/zhugecaomao/ALTRun/wiki`">https://github.com/zhugecaomao/ALTRun/wiki</a>"
        ENG[700] := "Command Manager"                                       ; 700+ Command Manager
        ENG[701] := "Command"
        ENG[702] := "Command type"
        ENG[703] := "Command line"
        ENG[704] := "Shortcut / Description"
        ENG[705] := "Configuration section"
        ENG[706] := "Command Rank"
        ENG[800] := "Do you really want to delete the following command?`n`n[" ; 800+ Message Info
        ENG[801] := "Confirm Want to Delete?"
        ENG[802] := "Command has been deleted successfully!"
        ENG[803] := "An error occurred while deleting the command."
        ENG[804] := "Index database is empty, please click`n`n'OK' to rebuild the index`n`n'Cancel' to enter the program without index`n`n(Please ensure the program directory is writable)"
        ENG[805] := "A new version ("
        ENG[806] := ") is available!`n`nClick OK to open the download page..."
        ENG[807] := "You are already using the latest version ("
        ENG[808] := ") !"
        ENG[809] := "Failed to check for updates! Please check your network connection.`n`nError message: "
        ENG[810] := "The current command cannot be edited."
        ENG[820] := "Command Manager"
        ENG[821] := "Command path cannot be empty. Please enter a valid path."
        ENG[822] := "An error occurred while adding the command: "
        ENG[823] := "The following command added / modified successfully!`n`n[ "

        CHN[1]  := "简体中文"                                               ; 1~9 Reserved
        CHN[2]  := "选项"
        CHN[7]  := "确定"
        CHN[8]  := "取消"
        CHN[9]  := "帮助"
        CHN[10] := ["序号", "类型", "命令", "描述"]                          ; 10~49 Main GUI
        CHN[11] := "运行"
        CHN[12] := "选项"
        CHN[13] := "在此输入搜索内容..."
        CHN[50] := ["提示 | F1 | 帮助关于", "提示 | F2 | 配置选项", "提示 | F3 | 编辑当前命令", "提示 | F4 | 直接修改设置文件", "提示 | Alt+空格 / Alt+R | 激活 ALTRun", "提示 | 失去焦点 / Esc / 快捷键 | 关闭 ALTRun", "提示 | 回车 / Alt+序号 | 运行命令", "提示 | 上下箭头键 | 选择上一个或下一个命令", "提示 | Ctrl+D | 使用文件管理器定位命令所在目录"]
        CHN[51] := "提示: "                                                 ; 50~99 Tips
        CHN[52] := "推荐使用热键激活 (ALT + 空格)"
        CHN[53] := "智能排序 - 根据使用频率自动调整命令优先级"
        CHN[54] := "上/下箭头 = 上/下一个命令"
        CHN[55] := "Esc = 清除输入 / 关闭窗口"
        CHN[56] := "回车 = 运行当前命令"
        CHN[57] := "Alt + 序号 = 运行指定的命令"
        CHN[58] := "以 + 开头 = 新建命令"
        CHN[59] := "F3 = 编辑当前命令"
        CHN[60] := "F2 = 打开选项"
        CHN[61] := "Reindex = 重建文件索引"
        CHN[62] := "F1 = ALTRun 帮助&关于"
        CHN[63] := "ALT + 空格 = 显示 / 隐藏窗口"
        CHN[64] := "Ctrl+Q = 重新加载 ALTRun"
        CHN[65] := "Ctrl + 序号 = 选择指定的命令"
        CHN[66] := "Alt + F4 = 退出"
        CHN[67] := "Ctrl+D = 使用文件管理器定位当前命令所在目录"
        CHN[68] := "F4 = 直接修改设置文件 (.ini)"
        CHN[69] := "以空格开头 = 使用 Everything 搜索文件"
        CHN[70] := "Ctrl+'+' = 增加当前命令的优先级"
        CHN[71] := "Ctrl+'-' = 减少当前命令的优先级"
        CHN[100] := ["常规", "界面", "热键", "索引", "Listary", "插件", "状态统计", "关于"] ; 100~149 Options window (Listview)
        CHN[101] := "随系统自动启动"
        CHN[102] := "添加到“发送到”菜单"
        CHN[103] := "添加到“开始”菜单"
        CHN[104] := "显示托盘图标 (系统任务栏中)"
        CHN[105] := "失去焦点时关闭窗口"
        CHN[106] := "窗口置顶"
        CHN[107] := "显示主窗口标题栏"
        CHN[108] := "使用当前系统主题 (取消勾选: 使用经典主题)"
        CHN[109] := "按下 [ESC] 清除输入, 再次按下关闭窗口 (取消勾选: 直接关闭窗口)"
        CHN[110] := "保留最近一次输入和匹配结果"
        CHN[111] := "显示命令图标"
        CHN[112] := "使用“发送到”时, 追溯 .lnk 目标文件"
        CHN[113] := "保存历史记录"
        CHN[114] := "保存运行日志"
        CHN[115] := "搜索时匹配完整路径"
        CHN[116] := "显示命令列表网格线"
        CHN[117] := "显示命令列表标题栏"
        CHN[118] := "显示命令列表序号"
        CHN[119] := "显示命令列表边框线"
        CHN[120] := "智能排序 (根据使用习惯自动调整命令优先级)"
        CHN[121] := "智能匹配"
        CHN[122] := "搜索时匹配字符串开头 (取消勾选: 匹配任意位置)"
        CHN[123] := "显示状态栏提示信息"
        CHN[124] := "显示命令执行次数 (状态栏)"
        CHN[125] := "显示状态栏"
        CHN[126] := "显示主窗口 [运行] 按钮"
        CHN[127] := "显示主窗口 [选项] 按钮"
        CHN[128] := "双缓冲绘图, 改善窗口闪烁"
        CHN[129] := "启用快速结构计算"
        CHN[130] := "显示简化路径 (仅显示文件/文件夹/应用程序名称, 而非完整路径)"
        CHN[131] := "设置语言为简体中文 (Simplified Chinese)"
        CHN[132] := "搜索时匹配拼音首字母"
        CHN[133] := "使用鼠标滚轮滚动选择命令"
        CHN[134] := "单击鼠标中键运行选定命令"
        CHN[135] := "自动检查更新"
        CHN[136] := "显示大图标"
        CHN[137] := "主窗口使用圆角"
        CHN[138] := "按空格键执行命令"
        CHN[139] := "激活时自动切换为英文输入法"

        CHN[150] := "文件管理器"                                             ; 150~159 Options window (Other than Check Listview)
        CHN[151] := "Everything"
        CHN[152] := "历史命令数量"
        CHN[160] := "索引"                                                  ; 160~169 Index
        CHN[161] := "索引位置"
        CHN[162] := "索引文件类型"
        CHN[163] := "索引排除项"
        CHN[164] := "索引目录深度"
        CHN[165] := "索引 Windows Store 应用"
        CHN[170] := "界面"                                                  ; 170~189 GUI
        CHN[171] := "搜索结果数量"
        CHN[172] := "每列宽度"
        CHN[173] := "字体 (主界面)"
        CHN[174] := "字体 (选项页)"
        CHN[175] := "字体 (状态栏)"
        CHN[176] := "主窗口尺寸 (宽 x 高)"
        CHN[177] := "命令列表尺寸 (宽 x 高)"
        CHN[178] := "颜色 (命令列表)"
        CHN[179] := "颜色 (主界面)"
        CHN[180] := "背景图片"
        CHN[181] := "透明度"
        CHN[182] := "选择字体"
        CHN[183] := "选择颜色"
        CHN[184] := "选择图片"
        CHN[190] := "热键"                                                  ; 190~209 Hotkey
        CHN[191] := "激活热键 (全局)"
        CHN[192] := "主热键"
        CHN[193] := "辅热键"
        CHN[194] := "可以同时设置两个热键"
        CHN[195] := "重置激活热键"
        CHN[200] := "快捷操作和热键 (非全局)"
        CHN[201] := "快捷键"
        CHN[202] := "触发操作"
        CHN[203] := "热键 1"
        CHN[204] := "热键 2"
        CHN[206] := "热键 3"
        CHN[210] := "Listary"                                               ; 210~249 Listary
        CHN[211] := "目录快速切换"
        CHN[212] := "文件管理器 ID"
        CHN[213] := "打开/保存对话框 ID"
        CHN[214] := "排除窗口 ID"
        CHN[215] := "热键"
        CHN[216] := "跳转到 Total Commander 路径"
        CHN[217] := "跳转到资源管理器路径"
        CHN[218] := "在打开/保存对话框中自动跳转路径"
        CHN[219] := "未找到 Total Commander 窗口，请先打开 Total Commander。"
        CHN[220] := "未找到资源管理器窗口，请先打开资源管理器。"
        CHN[221] := "{1} → TC / {2} → Explorer"

        CHN[250] := "插件"                                                  ; 250~299 Plugins
        CHN[251] := "文本末尾自动添加日期"
        CHN[252] := "应用到窗口 ID"
        CHN[253] := "热键"
        CHN[254] := "日期格式"
        CHN[255] := "扩展名前自动添加日期"
        CHN[259] := "条件触发快捷操作"
        CHN[260] := "如果窗口特征包含"
        CHN[261] := "热键 (自由定制)"
        CHN[262] := "触发操作"
        CHN[300] := "显示"                                                  ; 300+ 托盘菜单
        CHN[301] := "配置选项`tF2"
        CHN[302] := "重建索引"
        CHN[303] := "状态统计"
        CHN[304] := "关于`tF1"
        CHN[305] := "脚本信息"
        CHN[307] := "重新加载"
        CHN[308] := "退出"
        CHN[309] := "检查更新"
        CHN[310] := "配置文件`tF4"

        CHN[400] := "运行命令`tEnter"                                       ; 400+ 列表右键菜单
        CHN[401] := "定位命令`tCtrl+D"
        CHN[402] := "复制命令`tCtrl+C"
        CHN[403] := "新建命令`tCtrl+N"
        CHN[404] := "编辑命令`tF3"
        CHN[405] := "删除命令`tCtrl+Del"
        CHN[406] := ""
        CHN[407] := "复制状态栏信息"
        CHN[408] := "已复制 : "
        CHN[410] := "显示状态统计"
        CHN[500] := "30天前"
        CHN[501] := "当前"                                                  ; 500+ 状态统计
        CHN[502] := "执行命令次数统计"
        CHN[503] := "今天激活程序次数"
        CHN[600] := "关于"                                                  ; 600+ 关于
        CHN[601] := "一款开源、轻量、高效、功能强大的启动工具"
            . "`n能够快速查找系统中的内容或者启动应用程序"
            . "`n`n配置文件`n" g_JSON "`n`n程序文件`n" A_ScriptFullPath
            . "`n`n版本更新"
            . "`n<a href=`"https://github.com/zhugecaomao/ALTRun/releases`">https://github.com/zhugecaomao/ALTRun/releases</a>"
            . "`n`n源代码开源在 GitHub"
            . "`n<a href=`"https://github.com/zhugecaomao/ALTRun`">https://github.com/zhugecaomao/ALTRun</a>"
            . "`n`n有关更多详细信息，请参阅帮助和 Wiki 页面"
            . "`n<a href=`"https://github.com/zhugecaomao/ALTRun/wiki`">https://github.com/zhugecaomao/ALTRun/wiki</a>"
        CHN[700] := "命令管理器"                                             ; 700+ 命令管理器
        CHN[701] := "命令"
        CHN[702] := "命令类型"
        CHN[703] := "命令行"
        CHN[704] := "快捷方式 / 描述"
        CHN[705] := "配置节"
        CHN[706] := "命令权重"
        CHN[800] := "您确定要删除以下命令吗?`n`n["                            ; 800+ 提示消息
        CHN[801] := "确认删除?"
        CHN[802] := "命令已成功删除!"
        CHN[803] := "删除命令时发生错误!"
        CHN[804] := "检测到当前的索引数据库为空, 请点击:`n`n'确定' 重新建立索引`n`n'取消' 忽略索引进入程序`n`n(请确保程序目录有写入权限)"
        CHN[805] := "有新版本可用 ("
        CHN[806] := ") !`n`n点击确定打开下载页面..."
        CHN[807] := "您已经在使用最新版本 ("
        CHN[808] := ") 了!"
        CHN[809] := "检查更新失败! 请检查您的网络连接.`n`n错误信息: "
        CHN[810] := "当前命令无法编辑。"
        CHN[820] := "命令管理器"
        CHN[821] := "命令路径不能为空，请输入有效路径。"
        CHN[822] := "添加命令时发生错误："
        CHN[823] := "以下命令添加/修改成功!`n`n[ "

        lng := Lang.IsChinese() ? CHN : ENG
        g_LOG.Debug("Lang.Load: Set language to " lng[1] "...OK")
        return lng
    }

    ; 启动早期(LoadAppData() 尚未把配置读进内存之前)就要知道当前是中文还是
    ; 英文界面, 所以这里直接读一次 ALTRun.json, 不依赖 g_CONFIG。
    static IsChinese() {
        if FileExist(g_JSON) {
            try {
                data := JSON.parse(FileRead(g_JSON, "UTF-8"))
                if (data is Map && data.Has("Config") && data["Config"] is Map && data["Config"].Has("Chinese"))
                    return data["Config"]["Chinese"] ? 1 : 0
            }
        }
        return 0
    }
}
