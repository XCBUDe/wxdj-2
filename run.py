#!/usr/bin/env python3
"""
汽车之家 奇瑞经销商信息采集 — 脚本启动器

用法示例：
  # 全功能，测试单店（联网）
  python run.py --test 2207406

  # 样本数据演示（不联网，验证 Excel 格式）
  python run.py --sample

  # 采集广东全省（联网）
  python run.py --province 广东

  # 采集广东东莞，关闭截图
  python run.py --province 广东 --city 东莞 --no-screenshot

  # 三省全量，2并发，续跑
  python run.py --province 全部 --concurrency 2 --checkpoint output/cp.jsonl

  # 只要店头名称和软文日期，不要报价和截图
  python run.py --province 福建 --no-pricing --no-screenshot
"""
import argparse
import sys


def parse_args():
    p = argparse.ArgumentParser(
        description="汽车之家 奇瑞经销商信息采集",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )

    # ── 区域 ──────────────────────────────────────────────────────────────
    region = p.add_argument_group("区域参数")
    region.add_argument(
        "--province", default="全部",
        choices=["全部", "广东", "福建", "海南"],
        help="省份（默认：全部）",
    )
    region.add_argument(
        "--city", default="全部",
        metavar="CITY",
        help="城市名称，如 东莞（默认：全部）",
    )

    # ── 功能开关 ──────────────────────────────────────────────────────────
    feat = p.add_argument_group("功能开关（默认全开）")
    feat.add_argument(
        "--pricing", default=True,
        action=argparse.BooleanOptionalAction,
        help="采集车型报价",
    )
    feat.add_argument(
        "--screenshot", default=True,
        action=argparse.BooleanOptionalAction,
        help="采集头图截图（需 Playwright + Chromium）",
    )
    feat.add_argument(
        "--article-date", dest="article_date", default=True,
        action=argparse.BooleanOptionalAction,
        help="采集最新软文发布日期",
    )

    # ── 输出 ──────────────────────────────────────────────────────────────
    out = p.add_argument_group("输出参数")
    out.add_argument(
        "--output", default="output/result.xlsx",
        metavar="PATH",
        help="输出 xlsx 路径（默认：output/result.xlsx）",
    )
    out.add_argument(
        "--checkpoint", default="output/checkpoint.jsonl",
        metavar="PATH",
        help="断点续跑文件（默认：output/checkpoint.jsonl）",
    )

    # ── 运行模式 ──────────────────────────────────────────────────────────
    mode = p.add_argument_group("运行模式")
    mode.add_argument(
        "--test", metavar="DEALER_ID",
        help="只测单个经销商ID（如：2207406），快速验证",
    )
    mode.add_argument(
        "--sample", action="store_true",
        help="使用内置样本数据演示输出格式，不联网",
    )
    mode.add_argument(
        "--concurrency", type=int, default=2, metavar="N",
        help="并发浏览器数（默认：2，建议不超过 4）",
    )

    return p.parse_args()


def main():
    args = parse_args()

    features = {
        "pricing":      args.pricing,
        "screenshot":   args.screenshot,
        "article_date": args.article_date,
    }

    print("=" * 56)
    print("  汽车之家 奇瑞经销商信息采集")
    print("=" * 56)
    print(f"  省份:     {args.province}")
    print(f"  城市:     {args.city}")
    print(f"  车型报价: {'✓' if features['pricing']      else '✗'}")
    print(f"  头图截图: {'✓' if features['screenshot']   else '✗'}")
    print(f"  软文日期: {'✓' if features['article_date'] else '✗'}")
    print(f"  样本模式: {'✓' if args.sample              else '✗'}")
    print(f"  输出:     {args.output}")
    print("=" * 56)

    from src.crawler import run_crawl
    run_crawl(
        province=args.province,
        city=args.city,
        features=features,
        output=args.output,
        test_dealer_id=args.test,
        concurrency=args.concurrency,
        checkpoint_file=args.checkpoint,
        sample_mode=args.sample,
    )


if __name__ == "__main__":
    main()
