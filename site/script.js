// ALTRun 官网: 中英文切换、截图切换、复制命令
(function () {
  var root = document.documentElement;
  var titles = {
    zh: "ALTRun — 轻量的 Windows 启动器",
    en: "ALTRun — A lightweight launcher for Windows"
  };
  var captions = {                                                   // screenshots and animations (demo-*.png): [Chinese caption, English caption, file name]
    quickswitch: ["在打开 / 保存对话框里按 Ctrl+G 跳到文件管理器当前的文件夹; 下面的文件夹面板列出已经打开和最近用过的文件夹, 输入几个字回车就跳过去", "In an Open / Save dialog, Ctrl+G jumps to the folder open in your file manager; the folder panel below lists the folders you have open or used recently: type a few letters and press Enter to jump", "demo-quickswitch.png"],
    actions: ["选中结果按 → 打开操作面板: 以管理员身份运行、打开所在位置、复制路径…", "Press → on a result for actions: run as administrator, open location, copy path…", "actions.png"],
    files: ["空白搜索框先按空格再输入名称, 搜索文件和文件夹 (Everything 或内置索引)", "Press Space first to search files and folders (Everything or the built-in index)", "files.png"],
    filetype: ["doc 报告 只搜文档, 还有 pic、video、audio、zip、exe、cad, 类型可以自定义", "doc report searches documents only; pic, video, audio, zip, exe and cad too, and you can add your own types", "filetype.png"],
    clipboard: ["Ctrl+Alt+C 或输入 clip: 复制过的文字、文件和图片; 输入几个字筛选, 回车粘贴到当前的程序", "Ctrl+Alt+C or type clip: text, files and images you copied; type to filter, Enter pastes into the program you were in", "demo-clipboard.png"],
    calculator: ["直接输入算式, 边输入边出结果, 也支持单位换算 (10 km in mi)", "Type a formula and see the result as you type; unit conversion works too (10 km in mi)", "demo-calculator.png"],
    pinyin: ["拼音首字母: jsb → 记事本", "Pinyin initials: jsb → 记事本 (Notepad)", "pinyin.png"],
    websearch: ["g 关键词 用 Google 搜索, 搜索引擎可以自定义", "g keywords searches Google; engines are customizable", "websearch.png"],
    empty: ["还没输入时列出置顶的项目 (带图钉) 和最近打开的项目", "Before you type: pinned items (with a pin) and recently opened items", "empty.png"],
    windows: ["w 加空格列出所有打开的窗口, 再输入标题或程序名筛选, Enter 切换过去, → 可以关闭窗口", "w and a space lists your open windows; type a title or program name to filter, Enter switches to it, → closes it", "windows.png"],
    sendto: ["资源管理器里右键 → 发送到 → ALTRun, 把文件和文件夹加成自定义命令 (可以多选)", "In Explorer, right-click → Send to → ALTRun adds files and folders as custom commands (several at once works too)", "sendto.png"],
    browse: ["输入路径浏览文件夹, Tab 进入, 输入几个字筛选, Backspace 返回上一级", "Type a path to browse a folder; Tab goes in, typing filters, Backspace goes up", "demo-browse.png"],
    "prefs-general": ["偏好设置: 每一项都有说明", "Preferences: every option is explained", "prefs-general.png"],
    snippets: ["文字片段: 在任何程序里输入 ;关键字 自动展开 (这里是 ;thx), 光标停在要接着写的地方; 也可以在 ALTRun 里搜索名称或正文粘贴", "Snippets: type ;keyword in any program to expand it (here ;thx), with the caret where you go on typing; or search them by name or text in ALTRun", "demo-snippets.png"]
  };

  function current() { return root.getAttribute("data-lang"); }

  function setLang(lang) {
    root.setAttribute("data-lang", lang);
    root.lang = lang === "zh" ? "zh-CN" : "en";
    document.title = titles[lang];
    document.getElementById("lang-toggle").textContent = lang === "zh" ? "EN" : "中文";
    updateCaption();
    try { localStorage.setItem("altrun-lang", lang); } catch (e) { /* 隐私模式等: 不记住也能用 */ }
  }

  var saved = null;
  try { saved = localStorage.getItem("altrun-lang"); } catch (e) { saved = null; }
  var preferred = saved || ((navigator.language || "").toLowerCase().indexOf("zh") === 0 ? "zh" : "en");

  document.getElementById("lang-toggle").addEventListener("click", function () {
    setLang(current() === "zh" ? "en" : "zh");
  });

  // 截图切换
  var shot = document.getElementById("shot");
  var tabs = document.querySelectorAll(".tabs button");
  var selected = "actions";
  function updateCaption() {
    var text = captions[selected];
    document.getElementById("shot-caption").textContent = text ? text[current() === "zh" ? 0 : 1] : "";
  }
  tabs.forEach(function (tab) {
    tab.addEventListener("click", function () {
      var key = tab.getAttribute("data-shot");
      if (!Object.prototype.hasOwnProperty.call(captions, key)) return;   // 只认上面列出的截图
      selected = key;
      shot.src = "images/" + captions[key][2];
      tabs.forEach(function (t) { t.setAttribute("aria-selected", t === tab ? "true" : "false"); });
      updateCaption();
    });
  });

  // 复制命令
  document.querySelectorAll(".copy").forEach(function (button) {
    button.addEventListener("click", function () {
      var text = button.getAttribute("data-copy");
      var done = function () {
        var label = Array.prototype.slice.call(button.childNodes);           // 原来的中英文标签, 之后原样放回
        button.textContent = current() === "zh" ? "已复制" : "Copied";
        setTimeout(function () { button.replaceChildren.apply(button, label); }, 1500);
      };
      if (navigator.clipboard) navigator.clipboard.writeText(text).then(done, function () {});
    });
  });

  setLang(preferred);
})();
