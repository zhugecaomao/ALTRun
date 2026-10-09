// ALTRun 官网: 中英文切换、截图切换、复制命令
(function () {
  var root = document.documentElement;
  var titles = {
    zh: "ALTRun — 轻量的 Windows 启动器",
    en: "ALTRun — A lightweight launcher for Windows"
  };
  var captions = {                                                   // 截图: [中文说明, 英文说明, 文件名]
    quickswitch: ["在打开 / 保存对话框下面列出已经打开和最近用过的文件夹, 点一下就跳过去 (Ctrl+G 跳到 Total Commander 当前的文件夹)", "Below an Open / Save dialog: the folders you have open or used recently; click one to jump there (Ctrl+G jumps to the current Total Commander folder)", "quickswitch.png"],
    actions: ["选中结果按 → 打开操作面板: 以管理员身份运行、打开所在位置、复制路径…", "Press → on a result for actions: run as administrator, open location, copy path…", "actions.png"],
    files: ["空白搜索框先按空格再输入名称, 搜索文件和文件夹 (Everything 或内置索引)", "Press Space first to search files and folders (Everything or the built-in index)", "files.png"],
    clipboard: ["Ctrl+Alt+C 或输入 clip: 复制过的文字、文件和图片", "Ctrl+Alt+C or type clip: text, files and images you copied", "clipboard.png"],
    calculator: ["直接输入算式, 也支持单位换算 (10 km in mi)", "Type a formula; unit conversion works too (10 km in mi)", "calculator.png"],
    pinyin: ["拼音首字母: jsb → 记事本", "Pinyin initials: jsb → 记事本 (Notepad)", "pinyin.png"],
    websearch: ["g 关键词 用 Google 搜索, 搜索引擎可以自定义", "g keywords searches Google; engines are customizable", "websearch.png"],
    empty: ["还没输入时列出置顶的项目 (带图钉) 和最近打开的项目", "Before you type: pinned items (with a pin) and recently opened items", "empty.png"],
    browse: ["输入路径浏览文件夹, Tab 进入, Backspace 返回上一级", "Type a path to browse a folder; Tab goes in, Backspace goes up", "browse.png"],
    "prefs-general": ["偏好设置: 每一项都有说明", "Preferences: every option is explained", "prefs-general.png"]
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
