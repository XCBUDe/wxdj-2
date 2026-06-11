"""
Step 2 — 单店全字段采集（Playwright 渲染）

输出 dict:
  province, city, dealer_id, name,
  pricing: [{series, trim, price_bare, price_msrp}],
  article_date: "YYYY-MM-DD",
  screenshot_path: "output/screenshots/{id}.png"

采集逻辑：
  1. 首页  → 店头名称 / 最新软文日期 / 头图截图（Playwright 真实截图）
  2. 报价页→ 车型报价（滚动加载所有车系 → 逐车系点「展开」→ 逐行提取）
"""
import re
import time
from pathlib import Path
from playwright.sync_api import sync_playwright, Page, TimeoutError as PWTimeout

from .config import (
    CHROMIUM_PATH, CHROMIUM_ARGS, VIEWPORT,
    SCREENSHOT_SELECTOR, SCREENSHOT_CLIP_HEIGHT,
    DEALER_PAGE_URL, SCREENSHOT_DIR, REQUEST_DELAY,
)

# ── 页面 URL ──────────────────────────────────────────────────────────────
# 报价单页：经销商主页 + /price.html  （已由用户确认）
PRICE_PAGE_SUFFIX = "/price.html"   # 例: dealer.autohome.com.cn/2045891/price.html

# ── 首页选择器（兼容旧版 class 与新版 Tailwind 结构）────────────────────
SEL_BREADCRUMB  = ".bread-crumbs a, .breadcrumb a, .crumb-list a, .bread a"
SEL_DEALER_NAME = ".dealer-name, .info-name, .shopName, .dname, .shop-name"
SEL_SCREENSHOT  = (
    ".dealer-main-hd, .detail-top, .index-banner, "
    ".dealer-index-banner, .banner-wrap, .main-hd, .dealer-home-top"
)
SEL_DATE        = ".news-time, .art-time, [class*='art-date'], [class*='news-date']"

# ── 报价页（新 Tailwind 版页面用文本解析；旧版选择器保留作 fallback）──────
_PRICE_RE      = re.compile(r"[\d]+\.[\d]+\s*万")
_TRIM_RE       = re.compile(r"\d{4}款")
# 页面标题中提取经销商名称：【城市4S店】{name}4S店电话_汽车之家
_TITLE_NAME_RE = re.compile(r"】(.+?)(?=4S店|电话|_汽车)")
# 面包屑文本模式：当前位置：\n城市\n{name}\n首页
_BREADCRUMB_NAME_RE = re.compile(r"当前位置[：:]\s*\n\S+\n(.+?)\n首页")


# ─────────────────────────────────────────────────────────────────────────
# 首页：店头名称
# ─────────────────────────────────────────────────────────────────────────
def _extract_name(page: Page, fallback: str) -> str:
    # 1. Page title: 【城市4S店】{dealer_name}4S店电话_汽车之家
    try:
        m = _TITLE_NAME_RE.search(page.title())
        if m:
            return m.group(1).strip()
    except Exception:
        pass
    # 2. Breadcrumb text (当前位置：\n城市\n经销商名\n首页)
    try:
        body = page.inner_text("body", timeout=3000)
        m = _BREADCRUMB_NAME_RE.search(body)
        if m:
            return m.group(1).strip()
    except Exception:
        pass
    # 3. Legacy DOM selectors
    try:
        crumbs = page.locator(SEL_BREADCRUMB).all_text_contents()
        crumbs = [c.strip() for c in crumbs if c.strip() and c.strip() not in ("首页",)]
        if crumbs:
            return crumbs[-1]
    except Exception:
        pass
    for sel in SEL_DEALER_NAME.split(","):
        try:
            t = page.locator(sel.strip()).first.inner_text(timeout=2000).strip()
            if t:
                return t
        except Exception:
            pass
    return fallback


# ─────────────────────────────────────────────────────────────────────────
# 首页：最新软文日期
# ─────────────────────────────────────────────────────────────────────────
def _extract_article_date(page: Page) -> str:
    date_re = re.compile(r"\d{4}-\d{2}-\d{2}")
    dates = []
    for sel in SEL_DATE.split(","):
        try:
            for el in page.locator(sel.strip()).all():
                m = date_re.search(el.inner_text(timeout=800))
                if m:
                    dates.append(m.group())
        except Exception:
            pass
    if not dates:
        try:
            dates = date_re.findall(page.inner_text("body", timeout=3000))
        except Exception:
            pass
    return max(dates) if dates else ""


# ─────────────────────────────────────────────────────────────────────────
# 首页：头图截图（Playwright 真实截图，无 PIL 合成）
# ─────────────────────────────────────────────────────────────────────────
def _is_blocked_page(page: Page) -> str:
    """
    判断当前页是否为真实经销商页。
    返回非空字符串=被拦截/异常的原因；空字符串=正常。
    """
    BLOCK_MARKERS = [
        "Host not in allowlist", "allowlist", "403 Forbidden",
        "访问被拒绝", "您的访问出错了", "page not found", "404",
    ]
    try:
        body = page.inner_text("body", timeout=3000).strip()
    except Exception:
        body = ""
    # 拦截关键字
    for mk in BLOCK_MARKERS:
        if mk.lower() in body.lower():
            return mk
    # 内容过短（真实店铺页有大量文字）
    if len(body) < 200:
        return f"内容过短({len(body)}字)"
    # 必须能找到经销商标志性元素之一
    has_anchor = False
    for sel in (".bread-crumbs", ".breadcrumb", ".dealer-name", ".info-name", ".shop-name"):
        try:
            if page.locator(sel).count() > 0:
                has_anchor = True
                break
        except Exception:
            pass
    if not has_anchor and "经销商" not in body and "报价" not in body:
        return "未找到经销商页面特征元素"
    return ""


def _take_screenshot(page: Page, dealer_id: str) -> str:
    Path(SCREENSHOT_DIR).mkdir(parents=True, exist_ok=True)
    save_path = str(Path(SCREENSHOT_DIR) / f"{dealer_id}.png")

    # 1. 优先截取指定容器元素
    for sel in [s.strip() for s in SEL_SCREENSHOT.split(",")]:
        try:
            el = page.locator(sel).first
            if el.count() and el.is_visible(timeout=2000):
                el.screenshot(path=save_path)
                print(f"    [截图] 元素 {sel} -> {save_path}")
                return save_path
        except Exception:
            continue

    # 2. fallback：视口顶部固定高度裁剪
    try:
        page.screenshot(
            path=save_path,
            clip={"x": 0, "y": 0, "width": VIEWPORT["width"], "height": SCREENSHOT_CLIP_HEIGHT},
        )
        print(f"    [截图] viewport clip -> {save_path}")
        return save_path
    except Exception as e:
        print(f"    [WARN] 截图失败: {e}")
        return ""


# ─────────────────────────────────────────────────────────────────────────
# 报价页：车型报价（滚动加载 + 展开所有车系）
# ─────────────────────────────────────────────────────────────────────────
def _scroll_to_bottom(page: Page, pause: float = 0.8, max_rounds: int = 15):
    """分段滚动直到页面底部，触发懒加载"""
    prev_h = 0
    for _ in range(max_rounds):
        page.evaluate("window.scrollBy(0, window.innerHeight)")
        time.sleep(pause)
        cur_h = page.evaluate("document.body.scrollHeight")
        if cur_h == prev_h:
            break
        prev_h = cur_h
    page.evaluate("window.scrollTo(0, 0)")


def _expand_all_series(page: Page):
    """点击所有「展开」按钮，使折叠车型全部可见。
    用 JS 直接 click 绕过可见性/遮挡检测，确保全部车系展开。
    """
    for _ in range(3):
        try:
            clicked = page.evaluate(
                """() => {
                    let n = 0;
                    document.querySelectorAll('button').forEach(btn => {
                        if (btn.textContent.includes('展开')) {
                            btn.click(); n++;
                        }
                    });
                    return n;
                }"""
            )
            if clicked == 0:
                break
            time.sleep(0.8)
        except Exception:
            break


# ─────────────────────────────────────────────────────────────────────────
# 文本解析：新版 Tailwind 页面的「车型报价」区段
# 展开后文本格式（每个 trim）：
#   {year}款 {trim_name}         ← 含年份款的车型行
#   {指导价}万\t{裸车价}万        ← tab 分隔的价格行
#   查报价单预约试驾               ← 按钮文本，忽略
# ─────────────────────────────────────────────────────────────────────────
_SERIES_START_RE = re.compile(
    r"^(艾瑞泽|瑞虎|风云|奇瑞|捷途|星途|大蚂蚁|小蚂蚁|QQ|iCAR)"
)
_PRICE_PAIR_RE = re.compile(r"^(\d+\.?\d*万)\s+(\d+\.?\d*万)")


def _parse_pricing_from_text(body: str) -> list[dict]:
    """从 inner_text 中解析「车型报价」区段，返回 trim 级记录列表。"""
    idx = body.find("车型报价")
    if idx < 0:
        return []
    text = body[idx:]

    lines = [l.strip() for l in text.split("\n") if l.strip()]
    results = []
    current_series = ""
    pending_trim = ""

    i = 0
    while i < len(lines):
        line = lines[i]

        if "关于我们" in line or "版权所有" in line:
            break

        # 车系名：以已知车系前缀开头，短行，不含「万」「款」
        if (
            _SERIES_START_RE.match(line)
            and "万" not in line
            and "款" not in line
            and len(line) < 25
        ):
            current_series = line
            pending_trim = ""
            i += 1
            continue

        # 配置款行：以年份款开头，无价格数字，无「马力」
        if _TRIM_RE.match(line) and "万" not in line and "马力" not in line:
            pending_trim = line
            # 如果下一行就是价格行，立即采集
            if i + 1 < len(lines):
                m = _PRICE_PAIR_RE.match(lines[i + 1])
                if m:
                    results.append({
                        "series": current_series,
                        "trim": pending_trim,
                        "price_msrp": m.group(1),
                        "price_bare": m.group(2),
                    })
                    i += 2
                    pending_trim = ""
                    continue
            i += 1
            continue

        # 价格行（有 pending_trim 时才采集）
        if pending_trim and current_series:
            m = _PRICE_PAIR_RE.match(line)
            if m:
                results.append({
                    "series": current_series,
                    "trim": pending_trim,
                    "price_msrp": m.group(1),
                    "price_bare": m.group(2),
                })
                pending_trim = ""

        i += 1

    print(f"    [报价] 文本解析 → {len(results)} 条")
    return results


def _extract_pricing_from_page(page: Page) -> list[dict]:
    """
    在报价页提取所有车系 + 配置款价格。
    新版 Tailwind 页面：滚动 + 点击所有「展开」按钮 + 文本解析。
    """
    _scroll_to_bottom(page)
    _expand_all_series(page)
    time.sleep(1.0)

    try:
        body = page.inner_text("body", timeout=5000)
    except Exception as e:
        print(f"    [WARN] inner_text 失败: {e}")
        return []

    return _parse_pricing_from_text(body)


def _launch_browser(pw_ctx):
    """按优先级启动浏览器：
    1. 配置的 CHROMIUM_PATH（如手动指定）
    2. Playwright 自带 Chromium（playwright install 下载的）
    3. 系统已安装的 Edge / Chrome（无需下载，Windows 自带 Edge）
    """
    base = dict(headless=True, args=CHROMIUM_ARGS)
    attempts = []
    if CHROMIUM_PATH:
        attempts.append(("指定路径", dict(base, executable_path=CHROMIUM_PATH)))
    attempts.append(("自带Chromium", dict(base)))
    attempts.append(("系统Edge", dict(base, channel="msedge")))
    attempts.append(("系统Chrome", dict(base, channel="chrome")))

    last_err = None
    for label, kwargs in attempts:
        try:
            browser = pw_ctx.chromium.launch(**kwargs)
            print(f"    [浏览器] 使用 {label}")
            return browser
        except Exception as e:
            last_err = e
    raise RuntimeError(
        "无法启动任何浏览器（已尝试 自带Chromium / 系统Edge / 系统Chrome）。"
        f" 最后错误: {last_err}"
    )


# ─────────────────────────────────────────────────────────────────────────
# 主采集函数
# ─────────────────────────────────────────────────────────────────────────
def fetch_dealer_detail(dealer: dict, features: dict, browser=None) -> dict:
    """
    采集单个经销商全字段。
    dealer: {province, city, dealer_id, name}
    features: {pricing, screenshot, article_date}
    browser: 复用的 Playwright Browser 实例；None 时内部创建。
    """
    base_url  = DEALER_PAGE_URL.format(dealer_id=dealer["dealer_id"])
    price_url = base_url.rstrip("/") + PRICE_PAGE_SUFFIX

    result = {
        "province":        dealer.get("province", ""),
        "city":            dealer.get("city", ""),
        "dealer_id":       dealer["dealer_id"],
        "name":            dealer.get("name", ""),
        "pricing":         [],
        "article_date":    "",
        "screenshot_path": "",
        "error":           "",
    }

    _own_browser = browser is None
    pw_ctx = None

    try:
        if _own_browser:
            pw_ctx = sync_playwright().start()
            browser = _launch_browser(pw_ctx)

        ctx = browser.new_context(
            viewport=VIEWPORT,
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            ignore_https_errors=True,
        )

        # ── 首页：店头名称 + 软文日期 + 截图 ────────────────────────────
        page_home = ctx.new_page()
        print(f"  [INFO] 首页 {base_url}")
        page_home.goto(base_url, timeout=40000, wait_until="domcontentloaded")
        try:
            page_home.wait_for_selector(
                ".bread-crumbs, .breadcrumb, .dealer-name, .info-name",
                timeout=15000,
            )
        except PWTimeout:
            pass
        time.sleep(3.0)

        # 拦截/异常页检测：若不是真实店铺页，直接判失败，不截图不提取
        blocked = _is_blocked_page(page_home)
        if blocked:
            result["error"] = f"页面不可达或被拦截: {blocked}"
            print(f"  [BLOCKED] {result['error']}")
            page_home.close()
            ctx.close()
            return result

        result["name"] = _extract_name(page_home, result["name"])

        if features.get("article_date"):
            result["article_date"] = _extract_article_date(page_home)

        if features.get("screenshot"):
            result["screenshot_path"] = _take_screenshot(page_home, dealer["dealer_id"])

        page_home.close()

        # ── 报价页：车型报价（滚动 + 展开 + 提取）───────────────────────
        if features.get("pricing"):
            page_price = ctx.new_page()
            print(f"  [INFO] 报价页 {price_url}")
            page_price.goto(price_url, timeout=40000, wait_until="domcontentloaded")
            try:
                page_price.wait_for_selector(
                    "text=车型报价, text=展开, .dealer-table, table",
                    timeout=15000,
                )
            except PWTimeout:
                pass
            time.sleep(3.0)
            result["pricing"] = _extract_pricing_from_page(page_price)
            print(f"    -> {len(result['pricing'])} 条车型报价")
            page_price.close()

        ctx.close()

    except Exception as e:
        result["error"] = str(e)
        print(f"  [ERROR] {dealer['dealer_id']}: {e}")
    finally:
        if _own_browser and pw_ctx:
            try:
                browser.close()
            except Exception:
                pass
            pw_ctx.stop()

    return result


# ─────────────────────────────────────────────────────────────────────────
# 样本经销商（--sample 模式）
# 清远冠荣 = 2045891；车型报价由 data/standards.json 真实数据动态加载
# ─────────────────────────────────────────────────────────────────────────
def build_sample_detail() -> dict:
    """构建清远冠荣样本记录，报价取自 standards.json 全量真实数据。"""
    from .price_source import load_pricing
    return {
        "province": "广东", "city": "清远", "dealer_id": "2045891",
        "name": "奇瑞清远冠荣体验中心",
        "pricing": load_pricing(),       # 186 个车型，真实价格
        "article_date": "2026-04-04",    # 软文日期需联网采集；此处为占位
        "screenshot_path": "",           # 头图截图需联网采集，sample 模式留空
        "error": "",
    }


SAMPLE_DETAILS = {"2045891": build_sample_detail()}
