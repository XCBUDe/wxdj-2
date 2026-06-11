#!/usr/bin/env python3
"""
易车经销商页面结构探针 — 在本地 Windows 机器运行，把输出发给 agent。
测试经销商：广州品华，UID=100107636
"""
import sys, textwrap, re
import requests

UID = "100107636"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Referer": "https://www.yiche.com/",
}
URLS = [
    ("PC主页",       f"https://dealer.yiche.com/{UID}/"),
    ("移动主页",     f"https://dealer.m.yiche.com/d{UID}/"),
    ("移动-车型",    f"https://dealer.m.yiche.com/d{UID}/cars.html"),
    ("移动-软文",    f"https://dealer.m.yiche.com/d{UID}/news.html"),
]

def fetch(label, url):
    print(f"\n{'='*60}")
    print(f"  {label}")
    print(f"  URL: {url}")
    print(f"{'='*60}")
    try:
        r = requests.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
        print(f"  HTTP: {r.status_code}  编码: {r.encoding}  长度: {len(r.text)} chars")
        print(f"  最终URL: {r.url}")
        # 打印前 3000 字符的可读文本
        text = r.text
        # 粗略去掉 script/style 标签内容
        text_clean = re.sub(r'<script[^>]*>.*?</script>', '', text, flags=re.S)
        text_clean = re.sub(r'<style[^>]*>.*?</style>', '', text_clean, flags=re.S)
        text_clean = re.sub(r'<[^>]+>', ' ', text_clean)
        text_clean = re.sub(r'\s+', ' ', text_clean).strip()
        print(f"\n--- 页面文本（前3000字）---")
        print(text_clean[:3000])
        print(f"\n--- 原始HTML（前1000字）---")
        print(text[:1000])
    except Exception as e:
        print(f"  [ERROR] {e}")

print("=" * 60)
print("  易车页面结构探针  UID=" + UID)
print("=" * 60)

for label, url in URLS:
    fetch(label, url)

# 如果 requests 都能访问，再用 Playwright 抓一次 PC 主页，截图+全文
print("\n\n" + "="*60)
print("  尝试 Playwright 访问（需要 playwright 已安装）")
print("="*60)
try:
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        browser = None
        for method, kwargs in [
            ("自带Chromium", {}),
            ("系统Edge",     {"channel": "msedge"}),
            ("系统Chrome",   {"channel": "chrome"}),
        ]:
            try:
                browser = pw.chromium.launch(
                    headless=True,
                    args=["--no-sandbox", "--ignore-certificate-errors",
                          "--disable-blink-features=AutomationControlled"],
                    **kwargs,
                )
                print(f"  [浏览器] 使用 {method}")
                break
            except Exception as e:
                print(f"  [跳过] {method}: {e}")

        if browser is None:
            print("  [ERROR] 无可用浏览器，跳过 Playwright 阶段")
        else:
            ctx = browser.new_context(
                viewport={"width": 1440, "height": 900},
                user_agent=HEADERS["User-Agent"],
                ignore_https_errors=True,
                locale="zh-CN",
            )
            ctx.add_init_script(
                "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"
            )
            page = ctx.new_page()
            pc_url = f"https://dealer.yiche.com/{UID}/"
            print(f"  访问: {pc_url}")
            try:
                page.goto(pc_url, wait_until="networkidle", timeout=30000)
            except Exception:
                page.goto(pc_url, wait_until="domcontentloaded", timeout=20000)
            print(f"  Title: {page.title()}")
            # 展开所有「展开」按钮
            page.evaluate("""
                document.querySelectorAll('button,a').forEach(el => {
                    if (el.textContent.includes('展开') || el.textContent.includes('更多'))
                        el.click();
                });
            """)
            body = page.inner_text("body")
            print(f"  Body 长度: {len(body)}")
            print(f"\n--- Playwright body 文本（前4000字）---")
            print(body[:4000])
            page.screenshot(path="yiche_probe_screenshot.png", full_page=False)
            print(f"\n  截图已保存: yiche_probe_screenshot.png")
            browser.close()
except ImportError:
    print("  playwright 未安装，跳过")
except Exception as e:
    print(f"  [ERROR] {e}")

print("\n\n探针完成。请把以上全部输出文字复制发给 agent。")
print("如有截图 yiche_probe_screenshot.png，一并上传。")
