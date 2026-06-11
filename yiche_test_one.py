#!/usr/bin/env python3
"""
测试易车采集：跑广州品华（UID=100107636）一家，打印结果。
运行：python yiche_test_one.py
"""
import sys, json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from src.yiche_detail import fetch_dealer_yiche

dealer = {
    "province":  "广东",
    "city":      "广州",
    "dealer_id": "100107636",
    "name":      "广州品华",
}
features = {"pricing": True, "article_date": True, "screenshot": True}

print("开始采集广州品华（易车 UID=100107636）...")
result = fetch_dealer_yiche(dealer, features, screenshot_dir="output_yiche/screenshots_yiche")

print(f"\n店名：{result['name']}")
print(f"最新软文日期：{result['article_date']}")
print(f"头图路径：{result['screenshot_path']}")
print(f"错误：{result['error'] or '无'}")
print(f"\n采集到 {len(result['pricing'])} 款车型：")
for i, p in enumerate(result['pricing'], 1):
    print(f"  {i:3d}. [{p['series']}] {p['trim']}  指导价={p['msrp_str']}  本店报价={p['bare_str']}")

print("\n--- JSON 原始结果 ---")
r2 = {k: v for k, v in result.items() if k != "pricing"}
r2["pricing_count"] = len(result["pricing"])
print(json.dumps(r2, ensure_ascii=False, indent=2))
