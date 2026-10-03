#!/usr/bin/env python3
"""生成 ALTRun 官网 (GitHub Pages): site/index.html 模板 + 仓库里的数据 -> _site/

自动更新的内容:
  - 最新版本号、发布日期、下载链接和 zip 大小、总下载次数 (GitHub API; 取不到时用 App.ahk 里的版本号)
  - 主题截图 (docs/images/screenshots/theme-*.png, 顺序和名称取自 ThemeManager.ahk / I18n.ahk)
  - 最新版本的更新内容 (CHANGELOG.md 里第一个已发布的版本)

只用 Python 标准库。本地预览:
  python3 site/build.py && python3 -m http.server -d _site
环境变量: GITHUB_REPOSITORY (默认 zhugecaomao/ALTRun), GITHUB_TOKEN (可选, 避免 API 限流), SITE_OUT (默认 _site)
"""
import html
import json
import os
import re
import shutil
import struct
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SITE = os.path.join(ROOT, "site")
SHOTS = os.path.join(ROOT, "docs", "images", "screenshots")
OUT = os.path.join(ROOT, os.environ.get("SITE_OUT", "_site"))
REPO = os.environ.get("GITHUB_REPOSITORY", "zhugecaomao/ALTRun")


def read(*parts):
    with open(os.path.join(ROOT, *parts), encoding="utf-8-sig") as f:
        return f.read()


def api(path):
    request = urllib.request.Request("https://api.github.com/repos/" + REPO + path,
                                     headers={"Accept": "application/vnd.github+json", "User-Agent": "altrun-site"})
    token = os.environ.get("GITHUB_TOKEN")
    if token:
        request.add_header("Authorization", "Bearer " + token)
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.load(response)


def release_info():
    """最新版本; 取不到 (本地没有网络等) 时用源码里的版本号, 链接指向 Releases 页"""
    version = re.search(r'static Version := "([^"]+)"', read("Src", "Core", "App.ahk")).group(1)
    info = {"version": version, "date": "", "url": f"https://github.com/{REPO}/releases/latest",
            "size": "", "downloads": ""}
    try:
        latest = api("/releases/latest")
        info["version"] = latest["tag_name"]
        info["date"] = latest["published_at"][:10]
        zips = [a for a in latest["assets"] if a["name"].lower().endswith(".zip")]
        if zips:
            info["url"] = zips[0]["browser_download_url"]
            info["size"] = f"{zips[0]['size'] / 1024:.0f} KB"
        total, page = 0, 1
        while True:
            releases = api(f"/releases?per_page=100&page={page}")
            total += sum(a["download_count"] for r in releases for a in r["assets"])
            if len(releases) < 100:
                break
            page += 1
        info["downloads"] = f"{total:,}"
    except Exception as e:                                   # noqa: BLE001 - 网站照样生成, 只是少几个数字
        print("GitHub API unavailable, using fallback values:", e)
    return info


def inline(text):
    """CHANGELOG 里的行内格式: `代码`、**粗体**、[链接](url)"""
    text = html.escape(text, quote=False)
    text = re.sub(r"`([^`]+)`", r"<code>\1</code>", text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", text)
    return re.sub(r"\[([^\]]+)\]\((https?://[^)\s]+)\)", r'<a href="\2">\1</a>', text)


def latest_changes():
    """CHANGELOG.md 里第一个已发布的版本 -> (版本号, HTML)"""
    match = re.search(r"^## \[(\d[^\]]*)\]\s*\n(.*?)(?=^## \[|\Z)", read("CHANGELOG.md"), re.M | re.S)
    if not match:
        return "", ""
    out, in_list = [], False
    for line in match.group(2).splitlines():
        line = line.rstrip()
        if line.startswith("- "):
            if not in_list:
                out.append("<ul>")
                in_list = True
            out.append(f"<li>{inline(line[2:])}</li>")
            continue
        if in_list:
            out.append("</ul>")
            in_list = False
        if line.startswith("### "):
            out.append(f"<h4>{inline(line[4:])}</h4>")
        elif line:
            out.append(f"<p>{inline(line)}</p>")
    if in_list:
        out.append("</ul>")
    return match.group(1), "\n".join(out)


def png_size(path):
    """PNG 的宽和高 (IHDR), 写进 <img> 让图片载入前就按正确的比例占位"""
    with open(path, "rb") as f:
        return struct.unpack(">II", f.read(24)[16:24])


def theme_gallery():
    """theme-*.png (Light 用 search.png), 按主题列表里的顺序, 显示中英文名称"""
    names = {m.group(1).lower(): (m.group(2), m.group(3))
             for m in re.finditer(r's\["Theme\.(\w+)"\]\s*:=\s*\["([^"]*)",\s*"([^"]*)"', read("Src", "Core", "I18n.ahk"))}
    order = re.findall(r'"(\w+)"', re.search(r"BuiltinOrder\s*:=\s*\[(.*?)\]", read("Src", "UI", "ThemeManager.ahk"), re.S).group(1))
    files = {"light": "search.png"}
    for name in os.listdir(SHOTS):
        if name.startswith("theme-") and name.endswith(".png"):
            files[name[6:-4]] = name
    keys = ["light"] + [k.lower() for k in order]
    keys += sorted(k for k in files if k not in keys)                    # 有截图但不在列表里的也显示
    cards = []
    for key in keys:
        if key not in files:
            continue
        en, zh = names.get(key, (key, key))
        width, height = png_size(os.path.join(SHOTS, files[key]))
        cards.append(
            f'<figure class="theme"><img src="images/{files[key]}" alt="{html.escape(en)}" loading="lazy" width="{width}" height="{height}">'
            f'<figcaption><span lang="zh">{html.escape(zh)}</span><span lang="en">{html.escape(en)}</span></figcaption></figure>')
    return len(cards), "\n".join(cards)


def main():
    info = release_info()
    changes_version, changes = latest_changes()
    theme_count, themes = theme_gallery()
    values = {
        "VERSION": info["version"],
        "DATE": info["date"],
        "DOWNLOAD_URL": info["url"],
        "SIZE": info["size"] or "< 1 MB",
        "DOWNLOADS": info["downloads"] or "—",
        "THEME_COUNT": str(theme_count),
        "THEMES": themes,
        "CHANGES_VERSION": changes_version,
        "CHANGES": changes,
        "REPO": REPO,
    }
    page = read("site", "index.html")
    page = re.sub(r"\{\{(\w+)\}\}", lambda m: values[m.group(1)], page)

    shutil.rmtree(OUT, ignore_errors=True)
    os.makedirs(os.path.join(OUT, "images"))
    with open(os.path.join(OUT, "index.html"), "w", encoding="utf-8") as f:
        f.write(page)
    for name in ("style.css", "script.js"):
        shutil.copy(os.path.join(SITE, name), OUT)
    for name in os.listdir(SHOTS):
        if name.endswith(".png"):
            shutil.copy(os.path.join(SHOTS, name), os.path.join(OUT, "images", name))
    shutil.copy(os.path.join(ROOT, "docs", "images", "logo.png"), os.path.join(OUT, "images", "logo.png"))
    shutil.copy(os.path.join(ROOT, "Resources", "ALTRun.ico"), os.path.join(OUT, "favicon.ico"))
    print(f"Built {OUT}: {info['version']} ({info['size'] or 'size unknown'}), {theme_count} themes, changes {changes_version}")


if __name__ == "__main__":
    main()
