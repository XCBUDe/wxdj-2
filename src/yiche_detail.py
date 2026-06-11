"""
易车单店采集 — 纯 requests，无需浏览器。

采集项：
  - 全部车系 × 全部车款的 指导价 + 本店报价（裸车价）
  - 最新软文日期
  - 头图 Banner（下载 PNG）

URL 体系：
  主页   https://dealer.yiche.com/{uid}/
  车型   https://dealer.yiche.com/{uid}/cars.html        ← 全部车系+报价，静态HTML
  软文   https://dealer.m.yiche.com/d{uid}/news.html     ← 日期在文本中，regex可取
"""
import re
import time
import requests
from pathlib import Path
from bs4 import BeautifulSoup

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Referer":         "https://www.yiche.com/",
}
REQUEST_DELAY  = 1.5   # 两次请求间隔（秒）
REQUEST_TIMEOUT = 20


def _get(url: str) -> str:
    r = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
    r.encoding = "utf-8"
    return r.text


# ─────────────────────────────────────────────────────────────────────────────
# 解析 cars.html → [{series, trim, msrp_str, bare_str}]
# ─────────────────────────────────────────────────────────────────────────────

def _parse_cars_html(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")

    # ── Step 1: 从侧边栏/导航取 series_id → series_name ────────────────────
    series_map: dict[str, str] = {}
    for a in soup.select("a[href*='cars_']"):
        m = re.search(r"cars_(\d+)", a.get("href", ""))
        txt = a.get_text(strip=True)
        if m and txt and len(txt) <= 20 and not txt.isdigit() and not "查看" in txt:
            series_map.setdefault(m.group(1), txt)

    # ── Step 2: 找所有含报价数据的 <tr>，关联到车系 ────────────────────────
    pricing: list[dict] = []
    current_series: str | None = None
    current_sid:    str | None = None

    def _update_series_from(tag):
        """扫描 tag 内的 cars_{id} 链接，返回 (sid, sname) 或 (None,None)"""
        for a in tag.find_all("a", href=re.compile(r"cars_\d+")):
            m = re.search(r"cars_(\d+)", a["href"])
            if m and m.group(1) in series_map:
                return m.group(1), series_map[m.group(1)]
        return None, None

    # 遍历所有直接父 tr 的容器（tbody / table），按顺序处理
    for table in soup.find_all("table"):
        # 表头上方的车系标题（通常在 table 前一个兄弟节点里）
        for sibling in table.previous_siblings:
            if not hasattr(sibling, "find_all"):
                continue
            sid, sname = _update_series_from(sibling)
            if sname:
                current_series, current_sid = sname, sid
                break
            # 尝试：兄弟节点本身就是车系标题 div
            txt = sibling.get_text(" ", strip=True)
            # 匹配 "瑞虎8 8.99-11.99 万 (厂商指导价" 这类格式
            m = re.match(r"^([^\d（(]{2,15})\s+[\d.]+", txt)
            if m and "厂商指导价" in txt:
                current_series = m.group(1).strip()
                break

        # 如果上面没找到，从 table 的父容器找 cars_{id} 链接
        if not current_series:
            for parent in table.parents:
                sid, sname = _update_series_from(parent)
                if sname:
                    current_series, current_sid = sname, sid
                    break
                if parent.name in ("body", "html"):
                    break

        # 解析每个报价行
        for tr in table.find_all("tr"):
            cells = [td.get_text(" ", strip=True) for td in tr.find_all(["td", "th"])]
            row_text = " ".join(cells)
            if "询价" not in row_text:
                continue
            prices = re.findall(r"([\d]+\.[\d]+)\s*万", row_text)
            if len(prices) < 2:
                continue

            # 跳过表头行
            if cells and cells[0] in ("车款", "车型", "厂商指导价"):
                continue

            # 车款名：第一个单元格，去掉噪音词
            trim_name = re.sub(r"\s*(有现车|外观颜色|↓|直降|降价)\s*", " ",
                               cells[0] if cells else "").strip()
            if not trim_name or len(trim_name) < 3:
                continue

            msrp = prices[0]
            bare = prices[-1]

            # 过滤掉全是小数无效行（如 1.00 1.00）
            if float(bare) < 1:
                continue

            if current_series:
                pricing.append({
                    "series":    current_series,
                    "trim":      trim_name,
                    "msrp_str":  f"{msrp}万",
                    "bare_str":  f"{bare}万",
                })

    # ── Step 3: 如果 BS4 表格解析为空，回退到文本状态机 ────────────────────
    if not pricing:
        pricing = _parse_cars_text_fallback(html, series_map)

    return pricing


def _parse_cars_text_fallback(html: str, series_map: dict) -> list[dict]:
    """
    回退：用正则从清理文本中提取价格。
    当 BS4 没找到 <table> 时使用。
    """
    # 清理 HTML
    text = re.sub(r"<script[^>]*>.*?</script>", "", html, flags=re.S)
    text = re.sub(r"<style[^>]*>.*?</style>",  "", text, flags=re.S)
    text = re.sub(r"<[^>]+>",                  " ", text)
    text = re.sub(r"\s+",                       " ", text).strip()

    # 反转 series_map 以便按名字匹配
    all_series = list(series_map.values())

    pricing: list[dict] = []
    current_series = None

    # 状态机：遍历 token（按空格切分）
    # 车系标题格式："{名称} {X.XX}-{Y.YY} 万 （厂商指导价"
    # 报价行格式："{年款 车型名} {msrp}万 [降/本店价 {disc}万] {bare}万 询价"
    series_re = re.compile(
        r"([^\d（(]+?)\s+([\d.]+)-([\d.]+)\s*万\s*[（(]厂商指导价"
    )
    for m in series_re.finditer(text):
        current_series = m.group(1).strip()

    # 重新按车系分段
    # 先找每个车系标题的位置
    segments: list[tuple[int, str]] = []
    for m in series_re.finditer(text):
        segments.append((m.start(), m.group(1).strip()))

    for i, (pos, sname) in enumerate(segments):
        end = segments[i + 1][0] if i + 1 < len(segments) else len(text)
        chunk = text[pos:end]

        # 找报价行：包含 X.XX万 ... X.XX万 询价
        # 格式1：{msrp}万 降 {disc}万 {bare}万 询价
        # 格式2：{msrp}万 本店价 {bare}万 询价
        row_re = re.compile(
            r"(\d{4}款[^询]{3,80}?)"           # 车型名（年款开头）
            r"\s+([\d.]+)\s*万"                  # MSRP
            r"(?:\s+(?:降|直降|本店价|优惠)[^万]*?([\d.]+)\s*万)?"  # 可选优惠
            r"\s+([\d.]+)\s*万\s+询价"           # 本店报价
        )
        for rm in row_re.finditer(chunk):
            trim = rm.group(1).strip()
            msrp = rm.group(2)
            bare = rm.group(4)
            if float(bare) >= 1:
                pricing.append({
                    "series":   sname,
                    "trim":     trim,
                    "msrp_str": f"{msrp}万",
                    "bare_str": f"{bare}万",
                })

    return pricing


# ─────────────────────────────────────────────────────────────────────────────
# 软文最新日期
# ─────────────────────────────────────────────────────────────────────────────

def _fetch_article_date(uid: str) -> str:
    try:
        html = _get(f"https://dealer.m.yiche.com/d{uid}/news.html")
        dates = re.findall(r"20\d{2}-\d{2}-\d{2}", html)
        return dates[0] if dates else ""
    except Exception:
        return ""


# ─────────────────────────────────────────────────────────────────────────────
# 头图 Banner 下载
# ─────────────────────────────────────────────────────────────────────────────

def _download_banner(uid: str, save_dir: str) -> str:
    """下载经销商头图，返回本地路径；失败返回空串。"""
    Path(save_dir).mkdir(parents=True, exist_ok=True)
    save_path = str(Path(save_dir) / f"{uid}.png")
    if Path(save_path).exists():
        return save_path
    try:
        html = _get(f"https://dealer.yiche.com/{uid}/")
        soup = BeautifulSoup(html, "html.parser")

        # 找 banner 区域的图片
        img_url = None
        for selector in [
            "div.dealer-banner img", "div.banner img",
            "div.shop-header img", ".index-banner img",
            "div.head-img img", ".dealer-head img",
        ]:
            el = soup.select_one(selector)
            if el and el.get("src"):
                img_url = el["src"]
                break

        # 回退：找页面顶部区域最大的图
        if not img_url:
            for img in soup.find_all("img")[:20]:
                src = img.get("src", "")
                # 过滤掉 logo、图标等小图
                w = int(img.get("width", 0) or 0)
                h = int(img.get("height", 0) or 0)
                if src and (w > 200 or h > 80 or
                            any(k in src for k in ["banner", "head", "cover", "adsimg"])):
                    img_url = src
                    break

        if not img_url:
            return ""

        if img_url.startswith("//"):
            img_url = "https:" + img_url
        elif img_url.startswith("/"):
            img_url = "https://dealer.yiche.com" + img_url

        r = requests.get(img_url, headers=HEADERS, timeout=15)
        if r.status_code == 200:
            with open(save_path, "wb") as f:
                f.write(r.content)
            return save_path
    except Exception:
        pass
    return ""


# ─────────────────────────────────────────────────────────────────────────────
# 主采集函数（供 yiche_task.py 调用）
# ─────────────────────────────────────────────────────────────────────────────

def fetch_dealer_yiche(dealer: dict, features: dict,
                       screenshot_dir: str = "output/screenshots_yiche") -> dict:
    """
    采集单家经销商。
    dealer: {province, city, dealer_id, name}
    features: {pricing, article_date, screenshot}
    """
    uid = str(dealer["dealer_id"])
    result = {
        "province":    dealer.get("province", ""),
        "city":        dealer.get("city", ""),
        "dealer_id":   uid,
        "name":        dealer.get("name", ""),
        "pricing":     [],
        "article_date": "",
        "screenshot_path": "",
        "error":       "",
    }

    try:
        # 1. 车型报价（同时顺便拿店名）
        if features.get("pricing", True):
            html = _get(f"https://dealer.yiche.com/{uid}/cars.html")
            result["pricing"] = _parse_cars_html(html)
            # 从 <title> 补全店名
            if not result["name"] or result["name"] == dealer.get("name", ""):
                m = re.search(r"【(.+?)4[sS]店】", html)
                if m:
                    result["name"] = m.group(1)
            time.sleep(REQUEST_DELAY)

        # 2. 软文日期
        if features.get("article_date", True):
            result["article_date"] = _fetch_article_date(uid)
            time.sleep(REQUEST_DELAY)

        # 3. 头图截图
        if features.get("screenshot", True):
            result["screenshot_path"] = _download_banner(uid, screenshot_dir)
            time.sleep(REQUEST_DELAY)

    except Exception as e:
        result["error"] = str(e)

    return result
