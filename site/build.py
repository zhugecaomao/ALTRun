#!/usr/bin/env python3
"""Builds the ALTRun website (GitHub Pages): the site/index.html template + data from the repository -> _site/

Filled in automatically:
  - the latest version, release date, download link and zip size, total downloads (GitHub API; the version in App.ahk if it can't be read)
  - the theme screenshots (docs/images/screenshots/theme-*.png; order and names from ThemeManager.ahk / I18n.ahk / Resources/Lang/zh-CN.json)
  - the latest version's changes (the section of CHANGELOG.md for the version being downloaded; the English page shows its English
    summary, with the English and Chinese details in a section that can be expanded)
  - the width and height of the front page animation demo.png (APNG)

Python standard library only. Local preview:
  python3 site/build.py && python3 -m http.server -d _site
Environment variables: GITHUB_REPOSITORY (default zhugecaomao/ALTRun), GITHUB_TOKEN (optional, avoids the API rate limit), SITE_OUT (default _site),
                       SITE_URL (the site's address; link previews need the full URL; default https://zhugecaomao.github.io/ALTRun/)
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
SITE_URL = os.environ.get("SITE_URL", "https://zhugecaomao.github.io/ALTRun/").rstrip("/") + "/"


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
    except Exception as e:                                   # noqa: BLE001 - the site is built anyway, just missing a few numbers
        print("GitHub API unavailable, using fallback values:", e)
    return info


def inline(text):
    """CHANGELOG 里的行内格式: `代码`、**粗体**、[链接](url)"""
    text = html.escape(text, quote=False)
    text = re.sub(r"`([^`]+)`", r"<code>\1</code>", text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", text)
    return re.sub(r"\[([^\]]+)\]\((https?://[^)\s]+)\)", r'<a href="\2">\1</a>', text)


CJK = re.compile(r"[\u3400-\u9fff]")


def changes_html(lines):
    """CHANGELOG 的几行 -> HTML (段落、### 小标题、- 列表, 缩进两格的 "  - " 是上一项下面的子列表)"""
    out, in_list, in_sub = [], False, False
    for line in lines:
        if line.startswith("  - ") and in_list:
            if not in_sub:
                out[-1] = out[-1].removesuffix("</li>") + "<ul>"            # the sublist goes into the previous <li>
                in_sub = True
            out.append(f"<li>{inline(line[4:])}</li>")
            continue
        if in_sub:
            out.append("</ul></li>")
            in_sub = False
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
    if in_sub:
        out.append("</ul></li>")
    if in_list:
        out.append("</ul>")
    return "\n".join(out)


def latest_changes(version):
    """The section of CHANGELOG.md for the version being downloaded (without one, the first three-part release, never a test build
    YYYY.MM.DD.N) -> (version, HTML). The Chinese page shows all of it; the English page shows the English summary (the paragraphs
    without Chinese) and puts the English and Chinese details in a <details> that can be expanded"""
    changelog = read("CHANGELOG.md")
    match = re.search(r"^## \[(" + re.escape(version) + r")\]\s*\n(.*?)(?=^## \[|\Z)", changelog, re.M | re.S) \
        or re.search(r"^## \[(\d{4}\.\d{2}\.\d{2})\]\s*\n(.*?)(?=^## \[|\Z)", changelog, re.M | re.S)
    if not match:
        return "", ""
    lines = [line.rstrip() for line in match.group(2).splitlines()]
    english = [line for line in lines if line and not line.startswith(("- ", "### ", "  ")) and not CJK.search(line)]
    details = [line for line in lines if line not in english]
    en = changes_html(english) or "<p>Release notes (in Chinese):</p>"
    en += ('\n<details><summary>Details (in English and Chinese)</summary>\n<div lang="zh-CN">\n'   # lang="zh-CN": not hidden by the [lang="zh"] rules
           + changes_html(details) + "\n</div>\n</details>")
    return match.group(1), f'<div lang="zh">\n{changes_html(lines)}\n</div>\n<div lang="en">\n{en}\n</div>'


def png_size(path):
    """PNG 的宽和高 (IHDR), 写进 <img> 让图片载入前就按正确的比例占位"""
    with open(path, "rb") as f:
        return struct.unpack(">II", f.read(24)[16:24])


def theme_gallery():
    """theme-*.png (Light 用 search.png), 按主题列表里的顺序, 显示中英文名称"""
    english = {m.group(1).lower(): m.group(2)
               for m in re.finditer(r's\["Theme\.(\w+)"\]\s*:=\s*"([^"]*)"', read("Src", "Core", "I18n.ahk"))}
    chinese = {key[6:].lower(): text for key, text in json.loads(read("Resources", "Lang", "zh-CN.json")).items()
               if key.startswith("Theme.")}
    names = {key: (en, chinese.get(key, en)) for key, en in english.items()}
    order = re.findall(r'"(\w+)"', re.search(r"BuiltinOrder\s*:=\s*\[(.*?)\]", read("Src", "UI", "ThemeManager.ahk"), re.S).group(1))
    files = {"light": "search.png"}
    for name in os.listdir(SHOTS):
        # theme-compact.png shows Dark in Compact density, not a separate theme, so it is not counted or shown
        if name.startswith("theme-") and name.endswith(".png") and name != "theme-compact.png":
            files[name[6:-4]] = name
    keys = ["light"] + [k.lower() for k in order]
    keys += sorted(k for k in files if k not in keys)                    # screenshots that exist but aren't in the list are shown too
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
    changes_version, changes = latest_changes(info["version"])
    theme_count, themes = theme_gallery()
    demo_w, demo_h = png_size(os.path.join(SHOTS, "demo.png"))
    values = {
        "DEMO_W": str(demo_w),
        "DEMO_H": str(demo_h),
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
        "SITE_URL": SITE_URL,
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
        if name.endswith((".png", ".gif")):
            shutil.copy(os.path.join(SHOTS, name), os.path.join(OUT, "images", name))
    shutil.copy(os.path.join(ROOT, "docs", "images", "logo.png"), os.path.join(OUT, "images", "logo.png"))
    shutil.copy(os.path.join(ROOT, "Resources", "ALTRun.ico"), os.path.join(OUT, "favicon.ico"))
    print(f"Built {OUT}: {info['version']} ({info['size'] or 'size unknown'}), {theme_count} themes, changes {changes_version}")


if __name__ == "__main__":
    main()
