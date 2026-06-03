"""
报价数据源：从 data/standards.json（奇瑞全量车型标准库）读取车型报价。

JSON 中每个车型：
  {series 车系, mediaName 车型, guidePrice 指导价(元), discount 优惠(元), remark? 备注}

换算：
  指导价 = guidePrice
  裸车价 = guidePrice - discount
标记「报价单无此车型，建议下架」的条目会被过滤（不在经销商报价单内）。
"""
import json
from pathlib import Path

STANDARDS_PATH = Path(__file__).resolve().parent.parent / "data" / "standards.json"


def _fmt_wan(yuan: int) -> str:
    """元 → "X.XX万"（去掉多余的0）"""
    if yuan is None:
        return ""
    v = yuan / 10000
    s = f"{v:.2f}".rstrip("0").rstrip(".")
    return f"{s}万"


def load_pricing(path: Path = STANDARDS_PATH, include_offshelf: bool = False) -> list[dict]:
    """
    读取标准库，返回 [{series, trim, price_bare, price_msrp, discount}]
    include_offshelf=False 时过滤掉「报价单无此车型」的条目。
    """
    p = Path(path)
    if not p.exists():
        return []
    data = json.loads(p.read_text(encoding="utf-8"))
    models = data.get("models", {})
    rows = []

    for key, m in models.items():
        remark = m.get("remark", "")
        if not include_offshelf and "报价单无此车型" in remark:
            continue
        guide = m.get("guidePrice")
        disc = m.get("discount", 0) or 0
        bare = (guide - disc) if guide is not None else None
        rows.append({
            "series":     m.get("series", ""),
            "trim":       m.get("mediaName", ""),
            "price_bare": _fmt_wan(bare),
            "price_msrp": _fmt_wan(guide),
            "discount":   _fmt_wan(disc) if disc else "",
        })

    # 按车系分组保持顺序（同车系车型相邻）
    rows.sort(key=lambda r: (r["series"],))
    return rows


def load_article_date(path: Path = STANDARDS_PATH) -> str:
    """从 currentCampaign 推断软文活动（无真实日期时返回空）。"""
    return ""


if __name__ == "__main__":
    rows = load_pricing()
    print(f"共 {len(rows)} 个车型")
    for r in rows[:8]:
        print(r)
