# Microsoft Store listing / 商店介绍

The text for Partner Center (Apps and games → ALTRun → Store listings). Keep it in step with the README when features change;
the limits in brackets are Partner Center's. Screenshots: at least 1366 × 768 (see the last section).

商店介绍的文字 (Partner Center → 应用和游戏 → ALTRun → 商店一览)。功能变了要和 README 一起改; 方括号里是 Partner Center 的长度限制。

## Properties / 属性

| Field | Value |
|---|---|
| Category | Productivity |
| Privacy policy URL | https://github.com/zhugecaomao/ALTRun/blob/main/SECURITY.md#privacy-policy |
| Website | https://zhugecaomao.github.io/ALTRun/ |
| Support contact | https://github.com/zhugecaomao/ALTRun/issues |
| Price | Free, all markets |
| Copyright | Copyright (c) 2013-2026 zhugecaomao |
| Additional license terms | GPL-3.0: https://github.com/zhugecaomao/ALTRun/blob/main/LICENSE |
| Age rating (IARC) | No violence, no user interaction or shared content, no location, no purchases, no unrestricted web access (links open in the user's browser) |

## Restricted capabilities / 受限权限的理由 (Submission options)

> **runFullTrust**: ALTRun is a keyboard launcher for the Windows desktop. It starts the programs, files and folders the user searches for, runs system commands, registers global hotkeys (Alt+Space) with a low-level keyboard hook, reads the clipboard for its clipboard history, and moves Open / Save dialogs to folders open in file managers. This needs a full-trust desktop process.
>
> **unvirtualizedResources**: ALTRun keeps its settings in %APPDATA%\ALTRun, the same place as its other editions (portable zip, winget, Scoop), so users can move between them and back up or sync their settings. Its Delayed Start feature turns off other programs' start-up entries (HKCU\...\Explorer\StartupApproved) and starts them later through Task Scheduler; with registry write virtualization those changes would not reach Windows. Nothing is written outside the user's own profile and registry hive.

## Notes for certification / 给审核人员的说明 (Submission options)

```
ALTRun is a keyboard launcher for the Windows desktop. No account, sign-in or setup is needed.

How to test:
1. Start ALTRun from the Start menu (or run "altrun"). It runs in the notification area (tray icon).
2. Press Alt+Space to open the search window. Type a few letters, e.g. "notepad" or "calc", and press Enter to start it.
3. Type "=2*21" for the calculator, or ";" to see the snippets. The tray menu opens Preferences.

Why it needs a full-trust desktop process: it starts the programs, files and folders the user searches for, registers the global hotkey Alt+Space with a low-level keyboard hook, reads the clipboard for its clipboard history (text only stays on this PC; content from password managers is skipped), and moves Open / Save dialogs to folders open in file managers.

Optional features that change Windows settings only when the user turns them on in Preferences: start with Windows, and Delayed Start (turns off another program's start-up entry and starts it later through Task Scheduler).

ALTRun is free, open source (GPL-3.0) and has no telemetry or ads. Source code: https://github.com/zhugecaomao/ALTRun
```

---

## English (en-us)

### Short description [1000]

Press Alt+Space, type a few letters, press Enter: apps, files, web search, calculations, clipboard history and system commands, all from one small, fast, keyboard-first window. Free and open source.

### Description [10000]

ALTRun is a lightweight, open-source launcher for Windows. Press Alt+Space, type a few letters and press Enter to open an app, a file or a website, calculate, convert units, paste from your clipboard history or run a system command, without reaching for the mouse.

Fast and small
• Results appear as you type. With Everything (voidtools) running, file searches cover all disks in milliseconds; otherwise ALTRun uses its own index.
• About 2.5 MB, no runtime to install, low memory use.

Smart search
• Learns from what you pick and moves frequent choices up.
• Matches word initials (vsc → Visual Studio Code) and pinyin initials (wx → 微信), and highlights what matched.
• Apps from the Start menu, the desktop and the Microsoft Store are indexed automatically.

More than a launcher
• Calculator with unit, number base, date and optional currency conversion.
• Clipboard history for text, files and images, with pinning; content from password managers is never recorded.
• Text snippets that expand anywhere when you type ;keyword.
• Web search with customizable engines, browser bookmarks, window switching, more than 40 Windows Settings pages and system commands.
• Quick Switch: jump Open / Save dialogs to the folder open in Explorer, Total Commander or Directory Opus.
• Actions on selected text, files or URLs in any program.
• Folder sync with a live progress window, delayed start of programs after sign-in, and scripts (.ahk, .ps1, .bat, .py).

Made for the keyboard
• → opens the action panel (run as administrator, open file location, copy path, terminal here, Windows context menu).
• F3 edits in place; Ctrl+1 to Ctrl+9 open a result directly; the status bar shows what Enter will do.

Yours to adjust
• English, Simplified Chinese, Traditional Chinese and Japanese.
• 22 themes in three densities, following Windows light and dark mode.
• Settings in plain JSON files.

Private
• No account, no telemetry, no ads. Your data stays on your computer.

ALTRun is free software under the GPL-3.0 license. Documentation, source code and issue tracker: https://github.com/zhugecaomao/ALTRun

### What's new in this version [1500]

First release in the Microsoft Store. The Store keeps ALTRun up to date; your settings stay in %APPDATA%\ALTRun.

### Product features [20 × 200]

1. Open apps, files, folders and websites from one window: Alt+Space, a few letters, Enter
2. Results as you type; learns from what you pick
3. Word-initial and pinyin-initial matching with highlighting
4. Instant file search across all disks with Everything, or the built-in index
5. Calculator with unit, number base, date and currency conversion
6. Clipboard history for text, files and images, with pinning
7. Text snippets that expand anywhere with ;keyword
8. Web search with customizable engines and browser bookmarks
9. Window switcher and more than 40 Windows Settings pages
10. Quick Switch: jump Open / Save dialogs to folders open in file managers
11. Actions on selected text, files or URLs in any program
12. Action panel: run as administrator, open file location, copy path, terminal here
13. Folder sync with live progress; delayed start of programs after sign-in
14. English, Simplified Chinese, Traditional Chinese and Japanese
15. 22 themes; follows Windows light and dark mode
16. No account, no telemetry, no ads; free and open source (GPL-3.0)

### Search terms [7 × 30, 21 words in all]

launcher, app launcher, file search, clipboard history, keyboard, productivity, quick search

---

## 简体中文 (zh-cn)

### 简短说明 [1000]

按 Alt+Space, 输入几个字母, 回车: 应用、文件、网页搜索、计算、剪贴板历史和系统命令, 都在一个小巧、快速、以键盘为主的窗口里完成。免费开源。

### 说明 [10000]

ALTRun 是一款轻量、开源的 Windows 启动器。按 Alt+Space, 输入几个字母再按回车, 就能打开应用、文件或网站, 计算、换算单位, 从剪贴板历史粘贴, 或者运行系统命令, 手不用离开键盘。

又快又小
• 边输入边出结果。运行 Everything (voidtools) 时, 几毫秒就能搜遍所有磁盘的文件; 没有它时使用 ALTRun 自己的索引。
• 约 2.5 MB, 不用安装任何运行库, 占用内存少。

聪明的搜索
• 记住你选了什么, 常用的排在前面。
• 支持单词首字母 (vsc → Visual Studio Code) 和拼音首字母 (wx → 微信), 并高亮匹配的部分。
• 自动收录开始菜单、桌面和 Microsoft Store 里的应用。

不只是启动器
• 计算器, 支持单位、进制、日期换算和可选的汇率换算。
• 剪贴板历史, 记录文字、文件和图片, 可以置顶; 密码管理器复制的内容不会记录。
• 文字片段, 在任何程序里输入 ;关键字 自动展开。
• 可自定义的网页搜索、浏览器书签、窗口切换、40 多个 Windows 设置页和系统命令。
• Quick Switch: 在打开 / 保存对话框里一键跳到资源管理器、Total Commander 或 Directory Opus 当前的文件夹。
• 对任何程序里选中的文字、文件或网址执行操作。
• 带实时进度窗口的文件夹同步、登录后延迟启动程序, 以及脚本 (.ahk、.ps1、.bat、.py)。

为键盘而设计
• → 打开操作面板 (以管理员身份运行、打开所在位置、复制路径、在此处打开终端、Windows 右键菜单)。
• F3 直接编辑; Ctrl+1 到 Ctrl+9 直接打开结果; 状态栏显示 Enter 会做什么。

随你调整
• 英文、简体中文、繁体中文和日文界面。
• 22 款主题, 三种密度, 跟随 Windows 浅色 / 深色模式。
• 设置是普通的 JSON 文件。

注重隐私
• 不需要账号, 没有遥测, 没有广告。数据都留在你的电脑上。

ALTRun 是 GPL-3.0 许可的自由软件。文档、源代码和问题反馈: https://github.com/zhugecaomao/ALTRun

### 此版本的新增内容 [1500]

首次在 Microsoft Store 发布。由 Store 负责更新; 设置仍保存在 %APPDATA%\ALTRun。

### 产品功能 [20 × 200]

1. 在一个窗口里打开应用、文件、文件夹和网站: Alt+Space, 几个字母, 回车
2. 边输入边出结果; 记住你的选择
3. 单词首字母和拼音首字母匹配, 并高亮
4. 配合 Everything 瞬间搜遍所有磁盘, 或使用内置索引
5. 计算器: 单位、进制、日期和汇率换算
6. 剪贴板历史: 文字、文件和图片, 可以置顶
7. 文字片段: 在任何程序里输入 ;关键字 展开
8. 可自定义的网页搜索和浏览器书签
9. 窗口切换和 40 多个 Windows 设置页
10. Quick Switch: 打开 / 保存对话框一键跳到文件管理器当前的文件夹
11. 对任何程序里选中的文字、文件或网址执行操作
12. 操作面板: 以管理员身份运行、打开所在位置、复制路径、在此处打开终端
13. 带实时进度的文件夹同步; 登录后延迟启动程序
14. 英文、简体中文、繁体中文和日文界面
15. 22 款主题, 跟随 Windows 浅色和深色模式
16. 不需要账号、没有遥测和广告; 免费开源 (GPL-3.0)

### 搜索词 [7 × 30, 共 21 个词]

启动器, 快速启动, 文件搜索, 剪贴板历史, 效率工具, 拼音搜索, launcher

---

## 中文（繁體）(zh-tw)

The package declares zh-TW and ja-JP (`AppxManifest.xml` Resources), so Partner Center asks for a listing in each. Taiwan terms.

套件宣告了 zh-TW 和 ja-JP (`AppxManifest.xml` 的 Resources), 所以 Partner Center 要求每種語言都有商店頁面。用台灣的用語。

### 簡短說明 [1000]

按 Alt+Space, 輸入幾個字母, 再按 Enter: 應用程式、檔案、網頁搜尋、計算、剪貼簿歷程和系統指令, 都在一個小巧、快速、以鍵盤為主的視窗裡完成。免費開源。

### 說明 [10000]

ALTRun 是一款輕巧、開源的 Windows 啟動器。按 Alt+Space, 輸入幾個字母再按 Enter, 就能開啟應用程式、檔案或網站, 計算、換算單位, 從剪貼簿歷程貼上, 或執行系統指令, 雙手不用離開鍵盤。

又快又小
• 邊輸入邊出結果。執行 Everything (voidtools) 時, 幾毫秒就能搜尋所有磁碟的檔案; 沒有它時使用 ALTRun 自己的索引。
• 約 2.5 MB, 不需安裝任何執行階段, 記憶體佔用少。

聰明的搜尋
• 記住你選了什麼, 常用的排在前面。
• 支援英文單字字首 (vsc → Visual Studio Code) 和拼音字首 (wx → 微信), 並醒目標示符合的部分。
• 自動收錄開始功能表、桌面和 Microsoft Store 裡的應用程式。

不只是啟動器
• 計算機, 支援單位、進位制、日期換算和可選的匯率換算。
• 剪貼簿歷程, 記錄文字、檔案和圖片, 可以釘選; 密碼管理工具複製的內容不會記錄。
• 文字片段, 在任何程式裡輸入 ;關鍵字 自動展開。
• 可自訂的網頁搜尋、瀏覽器書籤、視窗切換、40 多個 Windows 設定頁面和系統指令。
• Quick Switch: 在開啟 / 儲存對話方塊裡一鍵跳到檔案總管、Total Commander 或 Directory Opus 目前的資料夾。
• 對任何程式裡選取的文字、檔案或網址執行動作。
• 附即時進度視窗的資料夾同步、登入後延遲啟動程式, 以及指令碼 (.ahk、.ps1、.bat、.py)。

為鍵盤而設計
• → 開啟動作面板 (以系統管理員身分執行、開啟檔案位置、複製路徑、在此開啟終端機、Windows 右鍵選單)。
• F3 直接編輯; Ctrl+1 到 Ctrl+9 直接開啟結果; 狀態列顯示 Enter 會做什麼。

隨你調整
• 英文、簡體中文、繁體中文和日文介面。
• 22 款主題, 三種密度, 跟隨 Windows 淺色 / 深色模式。
• 設定是一般的 JSON 檔案。

重視隱私
• 不需要帳號, 沒有遙測, 沒有廣告。資料都留在你的電腦上。

ALTRun 是採用 GPL-3.0 授權的自由軟體。文件、原始碼和問題回報: https://github.com/zhugecaomao/ALTRun

### 此版本的新功能 [1500]

首次在 Microsoft Store 發佈。由 Store 負責更新; 設定仍儲存在 %APPDATA%\ALTRun。

### 產品功能 [20 × 200]

1. 在一個視窗裡開啟應用程式、檔案、資料夾和網站: Alt+Space, 幾個字母, Enter
2. 邊輸入邊出結果; 記住你的選擇
3. 英文單字字首和拼音字首比對, 並醒目標示
4. 搭配 Everything 瞬間搜尋所有磁碟, 或使用內建索引
5. 計算機: 單位、進位制、日期和匯率換算
6. 剪貼簿歷程: 文字、檔案和圖片, 可以釘選
7. 文字片段: 在任何程式裡輸入 ;關鍵字 展開
8. 可自訂的網頁搜尋和瀏覽器書籤
9. 視窗切換和 40 多個 Windows 設定頁面
10. Quick Switch: 開啟 / 儲存對話方塊一鍵跳到檔案管理工具目前的資料夾
11. 對任何程式裡選取的文字、檔案或網址執行動作
12. 動作面板: 以系統管理員身分執行、開啟檔案位置、複製路徑、在此開啟終端機
13. 附即時進度的資料夾同步; 登入後延遲啟動程式
14. 英文、簡體中文、繁體中文和日文介面
15. 22 款主題, 跟隨 Windows 淺色和深色模式
16. 不需要帳號、沒有遙測和廣告; 免費開源 (GPL-3.0)

### 搜尋詞 [7 × 30]

啟動器, 快速啟動, 檔案搜尋, 剪貼簿歷程, 效率工具, 鍵盤, launcher

---

## 日本語 (ja-jp)

### 簡単な説明 [1000]

Alt+Space を押して数文字入力し Enter: アプリ、ファイル、Web 検索、計算、クリップボード履歴、システムコマンドを、小さく高速なキーボード中心のウィンドウ 1 つで。無料のオープンソース。

### 説明 [10000]

ALTRun は軽量でオープンソースの Windows 用ランチャーです。Alt+Space を押して数文字入力し Enter を押すだけで、アプリ、ファイル、Web サイトを開いたり、計算や単位変換をしたり、クリップボード履歴から貼り付けたり、システムコマンドを実行したりできます。マウスに手を伸ばす必要はありません。

高速で軽量
• 入力と同時に結果が表示されます。Everything (voidtools) が動いていれば、すべてのドライブのファイルを数ミリ秒で検索。なければ ALTRun 独自のインデックスを使います。
• 約 2.5 MB。ランタイムのインストールは不要で、メモリ使用量も少なめです。

賢い検索
• 選んだ項目を学習し、よく使うものを上位に表示します。
• 単語の頭文字 (vsc → Visual Studio Code) や中国語ピンインの頭文字に対応し、一致した部分を強調表示します。
• スタートメニュー、デスクトップ、Microsoft Store のアプリを自動で登録します。

ランチャー以上の機能
• 単位・進数・日付の変換、オプションで為替換算もできる電卓。
• テキスト、ファイル、画像のクリップボード履歴 (ピン留め可能)。パスワードマネージャーからコピーした内容は記録しません。
• どのアプリでも ;キーワード と入力すると展開されるテキストスニペット。
• カスタマイズ可能な Web 検索、ブラウザーのブックマーク、ウィンドウ切り替え、40 以上の Windows 設定ページとシステムコマンド。
• Quick Switch: [開く] / [保存] ダイアログを、エクスプローラー、Total Commander、Directory Opus で開いているフォルダーへ移動。
• どのアプリでも、選択したテキスト、ファイル、URL に対して操作を実行。
• リアルタイムの進捗ウィンドウ付きフォルダー同期、サインイン後のプログラム遅延起動、スクリプト (.ahk、.ps1、.bat、.py)。

キーボード中心の設計
• → でアクションパネルを開く (管理者として実行、ファイルの場所を開く、パスのコピー、ここでターミナルを開く、Windows のコンテキストメニュー)。
• F3 でその場で編集、Ctrl+1 ～ Ctrl+9 で結果を直接開く。ステータスバーに Enter で何が起きるかを表示します。

自分好みに
• 英語、簡体字中国語、繁体字中国語、日本語に対応。
• 22 種類のテーマと 3 段階の表示密度。Windows のライト / ダークモードに追従します。
• 設定はシンプルな JSON ファイル。

プライバシー
• アカウント不要、テレメトリなし、広告なし。データはあなたの PC に保存されます。

ALTRun は GPL-3.0 ライセンスのフリーソフトウェアです。ドキュメント、ソースコード、問題の報告: https://github.com/zhugecaomao/ALTRun

### このバージョンの新機能 [1500]

Microsoft Store での初めてのリリースです。更新は Store が行います。設定は引き続き %APPDATA%\ALTRun に保存されます。

### 製品の機能 [20 × 200]

1. アプリ、ファイル、フォルダー、Web サイトを 1 つのウィンドウから: Alt+Space、数文字、Enter
2. 入力と同時に結果を表示、選んだ項目を学習
3. 単語の頭文字とピンインの頭文字で検索、一致部分を強調表示
4. Everything ですべてのドライブを瞬時に検索、または内蔵インデックス
5. 電卓: 単位・進数・日付・為替の変換
6. クリップボード履歴: テキスト、ファイル、画像、ピン留め可能
7. テキストスニペット: どのアプリでも ;キーワード で展開
8. カスタマイズ可能な Web 検索とブラウザーのブックマーク
9. ウィンドウ切り替えと 40 以上の Windows 設定ページ
10. Quick Switch: [開く] / [保存] ダイアログをファイルマネージャーのフォルダーへ移動
11. どのアプリでも選択したテキスト、ファイル、URL を操作
12. アクションパネル: 管理者として実行、ファイルの場所を開く、パスのコピー、ターミナル
13. リアルタイム進捗付きのフォルダー同期、サインイン後の遅延起動
14. 英語、簡体字中国語、繁体字中国語、日本語
15. 22 種類のテーマ、Windows のライト / ダークモードに追従
16. アカウント不要、テレメトリ・広告なし、無料のオープンソース (GPL-3.0)

### 検索語句 [7 × 30]

ランチャー, アプリ起動, ファイル検索, クリップボード履歴, キーボード, 効率化, launcher

---

## Screenshots / 截图

Partner Center wants desktop screenshots of at least 1366 × 768 (PNG, up to 10, the first ones show first). The README's
screenshots are window-sized (about 700–1000 px wide), so the Store set is made separately: `packaging/msix/Screenshots/*.png`,
the windows on a 1920 × 1080 background, English interface, made-up data (as all of ALTRun's screenshots). Made by
`Tests/Screenshots/MakeStore.py`; the Screenshots workflow makes them again whenever it updates the screenshots. Order:

1. Search window with results (`search.png` / `demo.png`)
2. Action panel (`actions.png`)
3. Calculator and unit conversion (`demo-calculator.png`)
4. Clipboard history (`demo-clipboard.png`)
5. File search with Everything (`files.png`)
6. Quick Switch (`demo-quickswitch.png`)
7. Themes (`theme-dark.png` and others)
8. Preferences (`prefs-general.png`)

商店截图至少 1366 × 768。README 里的截图是窗口大小, 所以另外做一套 (`packaging/msix/Screenshots/*.png`): 窗口放在 1920 × 1080 的背景上, 英文界面、虚构数据。由 `Tests/Screenshots/MakeStore.py` 生成, 截图工作流更新截图时会一起重新生成。

### Captions / 截图说明 (each screenshot, optional)

| File | en-us | zh-cn | zh-tw | ja-jp |
|---|---|---|---|---|
| 01-search.png | Alt+Space, a few letters, Enter | Alt+Space, 几个字母, 回车 | Alt+Space, 幾個字母, Enter | Alt+Space、数文字、Enter |
| 02-actions.png | The action panel: run as administrator, open file location, copy path | 操作面板: 以管理员身份运行、打开所在位置、复制路径 | 動作面板: 以系統管理員身分執行、開啟檔案位置、複製路徑 | アクションパネル: 管理者として実行、ファイルの場所を開く、パスのコピー |
| 03-calculator.png | Calculator with unit and currency conversion | 计算器, 支持单位和汇率换算 | 計算機, 支援單位和匯率換算 | 単位と為替の換算ができる電卓 |
| 04-clipboard.png | Clipboard history for text, files and images | 剪贴板历史: 文字、文件和图片 | 剪貼簿歷程: 文字、檔案和圖片 | クリップボード履歴: テキスト、ファイル、画像 |
| 05-files.png | Instant file search with Everything | 配合 Everything 瞬间搜索文件 | 搭配 Everything 瞬間搜尋檔案 | Everything で瞬時にファイル検索 |
| 06-quickswitch.png | Quick Switch: jump Open / Save dialogs to your open folders | Quick Switch: 打开 / 保存对话框跳到当前文件夹 | Quick Switch: 開啟 / 儲存對話方塊跳到目前的資料夾 | Quick Switch: [開く] / [保存] ダイアログを開いているフォルダーへ |
| 07-theme-dark.png | 22 themes, light and dark | 22 款主题, 浅色和深色 | 22 款主題, 淺色和深色 | 22 種類のテーマ、ライトとダーク |
| 08-prefs-appearance.png | Preferences | 偏好设置 | 偏好設定 | 設定 |

The same English screenshots go into every language's listing (each needs at least one). / 每种语言的商店页面都用同一套英文截图 (每种至少要一张)。
