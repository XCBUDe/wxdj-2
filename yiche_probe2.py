#!/usr/bin/env python3
"""
易车探针 v2 — 摸清"全部车系 + 每款车本店报价"的页面结构。
经销商：广州品华 UID=100107636
在本地 Windows 跑，把全部输出文字发给 agent。
"""
import re
import requests

UID = "100107636"
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0.0.0 Safari/537.36"),
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Referer": "https://www.yiche.com/",
}


def get(url):
    r = requests.get(url, headers=HEADERS, timeout=15)
    r.encoding = "utf-8"
    return r.status_code, r.text


def clean(html):
    t = re.sub(r"<script[^>]*>.*?</script>", "", html, flags=re.S)
    t = re.sub(r"<style[^>]*>.*?</style>", "", t, flags=re.S)
    t = re.sub(r"<[^>]+>", " ", t)
    return re.sub(r"\s+", " ", t).strip()


print("=" * 70)
print("  易车探针 v2  UID=" + UID)
print("=" * 70)

# ── 1. 首页：提取所有车系链接（cars_{id}.html）和 PC 车型报价页 ──────────
home_url = f"https://dealer.yiche.com/{UID}/"
code, home = get(home_url)
print(f"\n[1] 首页 {home_url}  HTTP {code}  长度 {len(home)}")

# 找所有 cars_数字.html 链接
series_links = sorted(set(re.findall(r'cars_(\d+)\.html', home)))
print(f"\n  首页出现的车系ID（cars_<id>.html）: {series_links}")

# 找所有指向 dealer.yiche.com 子页的链接 + 锚文本
print("\n  首页所有内部链接（href + 文本，过滤含 car/报价/price）:")
for m in re.finditer(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', home, flags=re.S):
    href, txt = m.group(1), re.sub(r"<[^>]+>", "", m.group(2)).strip()
    if any(k in href.lower() for k in ["car", "price", "报价"]) or "报价" in txt:
        if txt:
            print(f"    {txt[:20]:20s} -> {href}")

# ── 2. PC 全部车型报价页候选 ─────────────────────────────────────────────
for path in ["car.html", "cars.html", "car/", "price.html"]:
    u = f"https://dealer.yiche.com/{UID}/{path}"
    try:
        code, html = get(u)
        print(f"\n[2] 试 {u}  HTTP {code}  长度 {len(html)}")
        if code == 200 and len(html) > 2000:
            ids = sorted(set(re.findall(r'cars_(\d+)\.html', html)))
            print(f"    该页车系ID: {ids}")
            print(f"    文本前1500字: {clean(html)[:1500]}")
    except Exception as e:
        print(f"    [ERR] {u}: {e}")

# ── 3. 取 1 个车系详情页，完整 dump 每款车结构 ───────────────────────────
if series_links:
    sid = series_links[0]
    for base in [f"https://dealer.yiche.com/{UID}/cars_{sid}.html",
                 f"https://dealer.m.yiche.com/d{UID}/cars_{sid}.html"]:
        try:
            code, html = get(base)
            print(f"\n[3] 车系详情 {base}  HTTP {code}  长度 {len(html)}")
            if code == 200:
                print(f"    --- 文本（前4000字）---")
                print(clean(html)[:4000])
        except Exception as e:
            print(f"    [ERR] {base}: {e}")

# ── 4. 软文页详情，确认最新日期取法 ──────────────────────────────────────
news_url = f"https://dealer.m.yiche.com/d{UID}/news.html"
code, news = get(news_url)
print(f"\n[4] 软文页 {news_url}  HTTP {code}")
dates = re.findall(r"20\d{2}-\d{2}-\d{2}", news)
print(f"    页面出现的日期: {dates[:10]}")

print("\n\n探针v2完成。请把全部输出文字复制发给 agent。")
