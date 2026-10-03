// ALTRun 官网: 中英文切换、截图切换、复制命令
(function () {
  var root = document.documentElement;
  var titles = {
    zh: "ALTRun — 轻量的 Windows 启动器",
    en: "ALTRun — A lightweight launcher for Windows"
  };
  var captions = {
    actions: ["选中结果按 → 打开操作面板: 以管理员身份运行、打开所在位置、复制路径…", "Press → on a result for actions: run as administrator, open location, copy path…"],
    files: ["空白搜索框先按空格再输入名称, 搜索文件和文件夹 (Everything 或内置索引)", "Press Space first to search files and folders (Everything or the built-in index)"],
    clipboard: ["Ctrl+Alt+C 或输入 clip: 复制过的文字、文件和图片", "Ctrl+Alt+C or type clip: text, files and images you copied"],
    calculator: ["直接输入算式, 也支持单位换算 (10 km in mi)", "Type a formula; unit conversion works too (10 km in mi)"],
    pinyin: ["拼音首字母: jsb → 记事本", "Pinyin initials: jsb → 记事本 (Notepad)"],
    websearch: ["g 关键词 用 Google 搜索, 搜索引擎可以自定义", "g keywords searches Google; engines are customizable"],
    "prefs-general": ["偏好设置: 每一项都有说明", "Preferences: every option is explained"]
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
      selected = tab.getAttribute("data-shot");
      shot.src = "images/" + selected + ".png";
      tabs.forEach(function (t) { t.setAttribute("aria-selected", t === tab ? "true" : "false"); });
      updateCaption();
    });
  });

  // 复制命令
  document.querySelectorAll(".copy").forEach(function (button) {
    button.addEventListener("click", function () {
      var text = button.getAttribute("data-copy");
      var done = function () {
        var old = button.innerHTML;
        button.textContent = current() === "zh" ? "已复制" : "Copied";
        setTimeout(function () { button.innerHTML = old; }, 1500);
      };
      if (navigator.clipboard) navigator.clipboard.writeText(text).then(done, function () {});
    });
  });

  setLang(preferred);
})();
